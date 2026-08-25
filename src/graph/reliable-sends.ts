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
  /**
   * The assistant's own display name, as Graph reports it on messages this account sends.
   * Required: without it a readback could claim a colleague's message quoting our text as
   * "our copy" — a genuinely failed send reported as success with a foreign message id.
   */
  selfDisplayName: string;
  /** Total attempts per send, first try included. */
  attempts?: number;
  sleepFn?: (ms: number) => Promise<void>;
  nowFn?: () => Date;
}

/**
 * Clock-skew allowance between this machine and Graph's timestamps. Kept tight on purpose:
 * the wider this window, the likelier an EARLIER identical message of our own (a repeated
 * status line, say) gets mistaken for this attempt's copy.
 */
const ATTEMPT_WINDOW_SKEW_MS = 60 * 1000;
const READBACK_LIMIT = 20;

function normalized(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

type MatchShape = 'whole-message' | 'reply-tail';

/**
 * Decorator over any TeamsChatsPort that makes sends safe to retry.
 *
 * Born from the 2026-08-24 incident where one broadcast landed eleven times: every POST
 * succeeded, every response was misread as a failure, and every "retry" was really a duplicate.
 * The rule this class encodes: a failure report describes the response path, not the chat.
 * Before ANY re-send, read the chat back — if our copy is standing, that copy IS the success.
 *
 * "Our copy" is deliberately strict: sent by this account's own display name, created within
 * this attempt's window, not deleted, and matching the sent text exactly (whole message for a
 * send; the tail after the quote card for a reply). Newest match wins.
 */
export class ReliableTeamsChats implements TeamsChatsPort {
  private readonly selfDisplayName: string;
  private readonly attempts: number;
  private readonly sleepFn: (ms: number) => Promise<void>;
  private readonly nowFn: () => Date;

  constructor(
    private readonly inner: TeamsChatsPort,
    options: ReliableSendOptions,
  ) {
    this.selfDisplayName = options.selfDisplayName;
    this.attempts = options.attempts ?? 2;
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
    return this.sendGuarded(chatId, text, 'whole-message', () =>
      this.inner.sendMessage(chatId, text),
    );
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
    // A landed reply's text carries the quoted original BEFORE our own words, so the match is
    // "ends with what we sent" — never containment, which the quoted original itself could
    // satisfy when someone replies with words the original already contains.
    return this.sendGuarded(chatId, text, 'reply-tail', () =>
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
    shape: MatchShape,
    doSend: () => Promise<ChatMessage>,
  ): Promise<ChatMessage> {
    const windowStart = this.nowFn().getTime() - ATTEMPT_WINDOW_SKEW_MS;

    for (let attempt = 1; ; attempt += 1) {
      if (attempt > 1) {
        // A retry only ever runs against a chat PROVEN not to hold our copy — including the
        // case where the previous readback itself failed (say, on the same dead connection
        // the send died on) and the backoff gave both a chance to recover.
        const landedLate = await this.findLandedCopy(chatId, text, shape, windowStart);
        if (landedLate) {
          return landedLate;
        }
      }

      let failure: unknown;
      try {
        return await doSend();
      } catch (caught) {
        failure = caught;
      }

      // A 429 is not an unknown outcome — Graph refused the write outright, nothing landed,
      // and reading the chat back now would only be another throttled request feeding the
      // penalty window. Wait the named window out FIRST; the readback then doubles as the
      // proof the gate has reopened.
      if (failure instanceof GraphError && failure.status === 429) {
        if (attempt >= this.attempts) {
          throw failure;
        }
        await this.sleepFn(this.backoffMs(attempt, failure));
        continue; // the loop's pre-send readback runs next, then the single retry
      }

      // Any other failure IS an unknown outcome — a claim about the response path, not the
      // chat. Only the chat itself can say whether the write landed.
      const landed = await this.findLandedCopy(chatId, text, shape, windowStart);
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
    shape: MatchShape,
    windowStartMs: number,
  ): Promise<ChatMessage | undefined> {
    const wanted = normalized(text);
    try {
      const { messages } = await this.inner.readMessages(chatId, undefined, READBACK_LIMIT);
      // Messages arrive oldest-first (applyWatermark sorts ascending); the newest match is
      // the one this attempt could have produced.
      return messages.findLast((candidate) => {
        if (candidate.isDeleted || candidate.from !== this.selfDisplayName) {
          return false;
        }
        const createdMs = Date.parse(candidate.createdDateTime);
        if (!Number.isFinite(createdMs) || createdMs < windowStartMs) {
          return false;
        }
        const candidateText = normalized(candidate.text);
        return shape === 'whole-message'
          ? candidateText === wanted
          : candidateText.endsWith(wanted);
      });
    } catch {
      // A readback that cannot run proves nothing either way; the retry loop re-checks before
      // any re-send.
      return undefined;
    }
  }

  private backoffMs(attempt: number, failure: unknown): number {
    const retryAfter =
      failure instanceof GraphError && failure.retryAfterSeconds
        ? failure.retryAfterSeconds
        : undefined;
    // Capped: an aggressive Retry-After must not park the caller for minutes — better to give
    // up honestly after a bounded wait than to hang a tool call nobody can see into.
    return retryAfter ? Math.min(retryAfter, 60) * 1000 : 2 ** attempt * 2500;
  }
}
