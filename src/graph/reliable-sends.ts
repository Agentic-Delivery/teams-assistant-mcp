import { GraphError } from './graph-client.js';
import type {
  AttachmentPayload,
  ChatSummary,
  OutboundFile,
  OutboundImage,
  TeamsChatsPort,
} from './teams-chats.js';
import type { ChatMessage, ReadResult } from '../messages.js';

export interface ReliableSendOptions {
  /** Total attempts per send, first try included. */
  attempts?: number;
  sleepFn?: (ms: number) => Promise<void>;
  nowFn?: () => Date;
}

/**
 * How far back a readback looks for "our" copy. Generous on purpose: clock skew between this
 * machine and Graph plus the time a slow request spends in flight — but short enough that a
 * genuinely identical broadcast from an earlier session is not mistaken for this attempt's.
 */
const ATTEMPT_WINDOW_SKEW_MS = 3 * 60 * 1000;
const READBACK_LIMIT = 20;

function normalized(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

/**
 * Decorator over any TeamsChatsPort that makes sends safe to retry.
 *
 * Born from the 2026-08-24 incident where one broadcast landed eleven times: every POST
 * succeeded, every response was misread as a failure, and every "retry" was really a duplicate.
 * The rule this class encodes: a failure report describes the response path, not the chat.
 * Before ANY re-send, read the chat back — if our copy is standing, that copy IS the success.
 */
export class ReliableTeamsChats implements TeamsChatsPort {
  private readonly attempts: number;
  private readonly sleepFn: (ms: number) => Promise<void>;
  private readonly nowFn: () => Date;

  constructor(
    private readonly inner: TeamsChatsPort,
    options: ReliableSendOptions = {},
  ) {
    this.attempts = options.attempts ?? 3;
    this.sleepFn = options.sleepFn ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));
    this.nowFn = options.nowFn ?? (() => new Date());
  }

  listChats(): Promise<ChatSummary[]> {
    return this.inner.listChats();
  }

  readMessages(chatId: string, since?: string, limit?: number): Promise<ReadResult> {
    return this.inner.readMessages(chatId, since, limit);
  }

  sendMessage(chatId: string, text: string): Promise<ChatMessage> {
    return this.sendGuarded(chatId, text, () => this.inner.sendMessage(chatId, text));
  }

  sendImage(chatId: string, image: OutboundImage, text?: string): Promise<ChatMessage> {
    // No readback key exists for an image (the text may be empty), so no blind retry either:
    // one attempt, honest error. Callers who need certainty read the chat themselves.
    return this.inner.sendImage(chatId, image, text);
  }

  sendFile(chatId: string, file: OutboundFile, text?: string): Promise<ChatMessage> {
    return this.inner.sendFile(chatId, file, text);
  }

  replyToMessage(chatId: string, replyToMessageId: string, text: string): Promise<ChatMessage> {
    return this.sendGuarded(chatId, text, () =>
      this.inner.replyToMessage(chatId, replyToMessageId, text),
    );
  }

  editMessage(chatId: string, messageId: string, newText: string): Promise<void> {
    return this.inner.editMessage(chatId, messageId, newText);
  }

  deleteMessage(chatId: string, messageId: string): Promise<void> {
    return this.inner.deleteMessage(chatId, messageId);
  }

  setReaction(chatId: string, messageId: string, reactionType: string): Promise<void> {
    // Idempotent by Graph's own semantics (setting the same reaction twice is one reaction),
    // so it needs no readback guard.
    return this.inner.setReaction(chatId, messageId, reactionType);
  }

  getAttachment(
    chatId: string,
    messageId: string,
    attachmentId?: string,
  ): Promise<AttachmentPayload> {
    return this.inner.getAttachment(chatId, messageId, attachmentId);
  }

  private async sendGuarded(
    chatId: string,
    text: string,
    doSend: () => Promise<ChatMessage>,
  ): Promise<ChatMessage> {
    const windowStart = new Date(this.nowFn().getTime() - ATTEMPT_WINDOW_SKEW_MS).toISOString();

    for (let attempt = 1; ; attempt += 1) {
      let failure: unknown;
      try {
        return await doSend();
      } catch (caught) {
        failure = caught;
      }

      // The send REPORTED failure — that is a claim about the response path, not the chat.
      // Only the chat itself can say whether the write landed.
      const landed = await this.findLandedCopy(chatId, text, windowStart);
      if (landed) {
        return landed;
      }
      if (attempt >= this.attempts) {
        throw failure;
      }
      await this.sleepFn(this.backoffMs(attempt, failure));
    }
  }

  private async findLandedCopy(
    chatId: string,
    text: string,
    windowStart: string,
  ): Promise<ChatMessage | undefined> {
    const wanted = normalized(text);
    try {
      const { messages } = await this.inner.readMessages(chatId, undefined, READBACK_LIMIT);
      return messages.find(
        (candidate) =>
          !candidate.isDeleted &&
          candidate.createdDateTime >= windowStart &&
          normalized(candidate.text).includes(wanted),
      );
    } catch {
      // A readback that cannot run proves nothing either way; the retry loop carries on.
      return undefined;
    }
  }

  private backoffMs(attempt: number, failure: unknown): number {
    const retryAfter =
      failure instanceof GraphError && failure.retryAfterSeconds
        ? failure.retryAfterSeconds
        : undefined;
    return retryAfter ? retryAfter * 1000 : 2 ** attempt * 2500;
  }
}
