import { DEFAULT_THROTTLE_WINDOW_MS, GraphError, MAX_RETRY_SLEEP_MS } from './graph-client.js';
import type {
  AttachmentPayload,
  ChatSummary,
  OutboundFile,
  OutboundImage,
  TeamsChatsPort,
} from './teams-chats.js';
import { htmlToText, type ChatMessage, type ReadResult } from '../messages.js';

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

function describe(failure: unknown): string {
  return failure instanceof Error ? `${failure.name}: ${failure.message}` : String(failure);
}

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

  sendHtmlMessage(chatId: string, html: string): Promise<ChatMessage> {
    // sendGuarded's second argument is the MATCH text, not necessarily what gets posted: a
    // landed copy's readback always comes back as htmlToText(body) (toChatMessage runs every
    // html body through it — see messages.ts), so comparing raw markup against that readback
    // would never match and every retry would re-post a duplicate. Reducing the caller's html
    // through the SAME converter the reader uses is what makes the comparison meaningful.
    return this.sendGuarded(chatId, htmlToText(html), 'whole-message', () =>
      this.inner.sendHtmlMessage(chatId, html),
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

  editHtmlMessage(chatId: string, messageId: string, html: string): Promise<void> {
    // A PATCH has no send/duplicate hazard to guard against — it targets an existing message
    // id, so there is nothing here for sendGuarded's readback dance to do. Same passthrough as
    // editMessage.
    return this.inner.editHtmlMessage(chatId, messageId, html);
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

  /**
   * matchText is what the readback is compared against, not necessarily what doSend actually
   * posts: for a plain send it IS the sent text, but sendHtmlMessage passes htmlToText(html)
   * here because a landed copy's readback always comes back through that same converter (see
   * sendHtmlMessage's own comment). Keeping the two separate is what lets one guard mechanism
   * serve both formats honestly.
   */
  private async sendGuarded(
    chatId: string,
    matchText: string,
    shape: MatchShape,
    doSend: () => Promise<ChatMessage>,
  ): Promise<ChatMessage> {
    const windowStart = this.nowFn().getTime() - ATTEMPT_WINDOW_SKEW_MS;
    let lastFailure: unknown;

    for (let attempt = 1; ; attempt += 1) {
      if (attempt > 1) {
        // A retry only ever runs against a chat PROVEN not to hold our copy — including the
        // case where the previous readback itself failed (say, on the same dead connection
        // the send died on) and the backoff gave both a chance to recover.
        const late = await this.findLandedCopy(chatId, matchText, shape, windowStart);
        if (late.landed) {
          return late.landed;
        }
        if (late.blocked) {
          if (lastFailure instanceof GraphError && lastFailure.status === 429) {
            // The previous send was REFUSED — nothing landed, the outcome is known. Say so,
            // with Graph's own Retry-After, not a fictional "unknown".
            throw lastFailure;
          }
          throw new GraphError(
            `Send outcome UNKNOWN: the send failed (${describe(lastFailure)}) and the retry could ` +
              'not read the chat back because the client is throttled. Do not re-send blindly — ' +
              'read the chat once the throttle clears.',
            0,
            'UnknownOutcome',
          );
        }
      }

      let failure: unknown;
      try {
        return await doSend();
      } catch (caught) {
        failure = caught;
        lastFailure = caught;
      }

      if (failure instanceof GraphError && failure.code === 'MessageFetchThrottled') {
        // The reply never got as far as a send — its original could not be fetched. Retrying
        // the whole reply (readback included) would only re-hit the family that is throttled.
        throw failure;
      }
      // A 429 is not an unknown outcome — Graph refused the write outright, nothing landed,
      // and reading the chat back now would only be another throttled request feeding the
      // penalty window. Wait the named window out FIRST; the readback then doubles as the
      // proof the gate has reopened.
      if (failure instanceof GraphError && failure.status === 429) {
        const waitMs = this.throttleWaitMs(failure);
        // Same rule as the client: a window that cannot clear inside one call's sleep cap
        // means there is no honest retry — fail NOW with the 429 and its Retry-After, rather
        // than sleep a minute and then fail locally against a still-closed gate.
        if (attempt >= this.attempts || waitMs > MAX_RETRY_SLEEP_MS) {
          throw failure;
        }
        await this.sleepFn(waitMs);
        continue; // the loop's pre-send readback runs next, then the single retry
      }

      // Any other failure IS an unknown outcome — a claim about the response path, not the
      // chat. Only the chat itself can say whether the write landed.
      const readback = await this.findLandedCopy(chatId, matchText, shape, windowStart);
      if (readback.landed) {
        return readback.landed;
      }
      if (readback.blocked) {
        // The readback could not run (the client is gate-closed by a concurrent 429). Nothing
        // is known: the write may be standing in the chat. Saying "not sent" here is the
        // 2026-08-24 lie through a new door — say the truth and name what actually happened.
        throw new GraphError(
          `Send outcome UNKNOWN: the send failed (${describe(failure)}) and the chat could not be ` +
            'read back because the client is throttled. Do not re-send blindly — read the chat ' +
            'once the throttle clears.',
          0,
          'UnknownOutcome',
        );
      }
      if (attempt >= this.attempts) {
        throw failure;
      }
      await this.sleepFn(this.backoffMs(attempt));
    }
  }

  private async findLandedCopy(
    chatId: string,
    matchText: string,
    shape: MatchShape,
    windowStartMs: number,
  ): Promise<{ landed?: ChatMessage; blocked: boolean }> {
    const wanted = normalized(matchText);
    try {
      const { messages } = await this.inner.readMessages(chatId, undefined, READBACK_LIMIT);
      // Messages arrive oldest-first (applyWatermark sorts ascending); the newest match is
      // the one this attempt could have produced.
      const landed = messages.findLast((candidate) => {
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
      return { ...(landed ? { landed } : {}), blocked: false };
    } catch (caught) {
      // A readback refused by a 429 — our own gate's or Graph's — never proved anything and
      // cannot until the throttle clears: that is "unknown", never "not landed", and the
      // caller must hear it instead of a retry. Any OTHER readback failure (a dead link, say)
      // proves nothing either way; the loop simply re-checks before any re-send.
      const blocked = caught instanceof GraphError && caught.status === 429;
      return { blocked };
    }
  }

  /** Backoff for a non-throttle failure (a 429 never reaches here — its branch uses throttleWaitMs). */
  private backoffMs(attempt: number): number {
    return 2 ** attempt * 2500;
  }

  /**
   * Reconciled with the client's gate: a 429 that names no Retry-After closes the gate for
   * DEFAULT_THROTTLE_WINDOW_MS, so any shorter wait guarantees the retry fails locally. Not
   * capped here — the caller compares it with MAX_RETRY_SLEEP_MS and fails honestly instead.
   */
  private throttleWaitMs(failure: GraphError): number {
    return failure.retryAfterSeconds ? failure.retryAfterSeconds * 1000 : DEFAULT_THROTTLE_WINDOW_MS;
  }
}
