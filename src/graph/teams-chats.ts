import { GRAPH_BASE_URL, GraphClient, GraphError } from './graph-client.js';
import {
  buildGraphMentionsPayload,
  renderHtmlWithMentions,
  renderTextWithMentions,
  resolveMentionTargets,
  type ChatMember,
  type MentionTarget,
} from './mentions.js';
import { NullSelfIdCache, type SelfIdCachePort } from './self-id-cache.js';
import { textToHtml } from '../formatting.js';
import {
  type ChatAttachmentRef,
  type ChatMessage,
  type ReadResult,
  applyWatermark,
  htmlToText,
  toChatMessage,
} from '../messages.js';

export interface ChatSummary {
  id: string;
  topic: string;
  chatType: string;
  lastUpdatedDateTime?: string;
  members: ChatMember[];
}

export interface AttachmentPayload {
  bytes: Uint8Array;
  contentType: string;
  name: string;
}

/** One pinned-message entry, as returned by listPinnedMessages/pinMessage. */
export interface PinnedMessage {
  /** The PIN resource's own id — what unpinMessage needs internally, never the chat message id. */
  id: string;
  /** The underlying chat message's id, when Graph reported one. */
  messageId?: string;
  /** Plain-text preview of the pinned message, via the same html-to-text conversion as reads. */
  preview: string;
}

export interface OutboundImage {
  bytes: Uint8Array;
  /** Only PNG and JPEG are worth supporting; Teams renders both inline. */
  contentType: 'image/png' | 'image/jpeg';
}

export interface OutboundFile {
  bytes: Uint8Array;
  name: string;
  contentType?: string;
}

/**
 * Secondary port for Teams chats. Everything above it (the MCP tools) speaks ChatSummary and
 * ChatMessage, never Graph JSON, so a different backing API or a test double swaps in here.
 */
export interface TeamsChatsPort {
  listChats(): Promise<ChatSummary[]>;
  readMessages(chatId: string, since?: string, limit?: number): Promise<ReadResult>;
  /**
   * Resolves each of `names` against the chat's CURRENT member list — case-insensitive,
   * unambiguous-substring ("Shiv" matches "Garg, Shivankit"; a name matching zero or more than
   * one member throws) — into the id/displayName pairs the send/reply/edit methods below need to
   * build a mention that actually notifies. Call this before any of them when the caller supplied
   * mention names; pass the result straight through as `mentions`.
   */
  resolveMentions(chatId: string, names: readonly string[]): Promise<MentionTarget[]>;
  sendMessage(chatId: string, text: string, mentions?: readonly MentionTarget[]): Promise<ChatMessage>;
  /**
   * Opt-in raw path: html is posted as the Graph body content VERBATIM — no textToHtml, no
   * escaping. The caller owns entity-escaping `<`, `>`, `&` inside their own content; this is
   * what lets Teams' HTML subset (tables, headings, colour — see the teams-styling skill's
   * verified vocabulary) actually render, since textToHtml only ever produces plain
   * paragraphs/breaks/links. `mentions` (from resolveMentions) are placed wherever the html
   * carries a matching `@{Name}` token — see renderHtmlWithMentions's doc comment for the
   * contract.
   */
  sendHtmlMessage(chatId: string, html: string, mentions?: readonly MentionTarget[]): Promise<ChatMessage>;
  sendImage(chatId: string, image: OutboundImage, text?: string): Promise<ChatMessage>;
  sendFile(chatId: string, file: OutboundFile, text?: string): Promise<ChatMessage>;
  replyToMessage(
    chatId: string,
    replyToMessageId: string,
    text: string,
    mentions?: readonly MentionTarget[],
  ): Promise<ChatMessage>;
  editMessage(chatId: string, messageId: string, newText: string, mentions?: readonly MentionTarget[]): Promise<void>;
  /** Same verbatim contract as sendHtmlMessage, applied to an edit. */
  editHtmlMessage(chatId: string, messageId: string, html: string, mentions?: readonly MentionTarget[]): Promise<void>;
  deleteMessage(chatId: string, messageId: string): Promise<void>;
  setReaction(chatId: string, messageId: string, reactionType: string): Promise<void>;
  getAttachment(chatId: string, messageId: string, attachmentId?: string): Promise<AttachmentPayload>;
  /**
   * The attachment metadata of one message — name, contentType, id — WITHOUT downloading
   * anything. Includes the quoted-reply card (contentType "messageReference") when there is
   * one, so a caller sees the message the way Graph reports it; the download methods are the
   * ones that skip the card.
   */
  listAttachments(chatId: string, messageId: string): Promise<ChatAttachmentRef[]>;
  /**
   * Downloads EVERY downloadable attachment on one message (quoted-reply cards are never
   * downloadable), optionally narrowed by a case-insensitive substring match on the attachment
   * name. The message is fetched once, not once per attachment — the single-message endpoint
   * has its own throttle budget and per-attachment refetching is how a three-file message
   * becomes four throttled calls. Zero matches is an error naming why (no attachments at all,
   * only a quote card, or a filter that matched nothing), never a silent empty success.
   */
  getAttachments(chatId: string, messageId: string, nameFilter?: string): Promise<AttachmentPayload[]>;
  /**
   * Pins a message, replacing whatever was pinned before it: Graph reports success on a SECOND
   * pin while silently dropping the first — a chat effectively holds exactly one pin (verified
   * live 2026-08-25). Returns the pinned-list state AFTER the call so callers can see what
   * actually happened rather than trusting the POST.
   */
  pinMessage(chatId: string, messageId: string): Promise<PinnedMessage[]>;
  unpinMessage(chatId: string, messageId: string): Promise<void>;
  listPinnedMessages(chatId: string): Promise<PinnedMessage[]>;
}

export type { ChatMember, MentionTarget };

/**
 * Graph's sharing-URL encoding: base64 of the absolute URL, made URL-safe, prefixed "u!".
 * A file shared into a chat lives in SharePoint/OneDrive and its contentUrl wants browser
 * cookies, not a Graph bearer token — fetching it directly answers 401. The /shares facade
 * resolves the same URL into a driveItem that the Graph token *can* download.
 *
 * Exported for its tests: the encoding must be byte-exact (padding stripped, `+` → `-`,
 * `/` → `_`) or Graph answers an unhelpful 400/404, and a mistake here is far easier to read
 * off the string itself than through a mocked download.
 */
export function shareIdFor(url: string): string {
  const base64 = Buffer.from(url, 'utf8').toString('base64');
  return `u!${base64.replace(/=+$/, '').replace(/\+/g, '-').replace(/\//g, '_')}`;
}

interface GraphDriveItem {
  id?: string;
  name?: string;
  eTag?: string;
  webUrl?: string;
}

/**
 * The shape of a successful `POST /me/drive/items/{id}/invite` response: a collection of
 * Permission resources, one per grant Graph actually created. Live-verified 2026-09-02 (see
 * KNOWN-ISSUES.md's wire-shape snapshots): each landed grant's AAD user id lives under
 * `grantedToV2.user.id` for a direct per-recipient grant — NOT proven by the HTTP status alone,
 * which can be 200 with an empty or partial `value` array when a recipient's grant did not
 * actually land. `grantedToIdentitiesV2` (an array, sibling to `grantedToV2` on the SAME
 * permission entry) is a documented Graph variant some permission kinds use instead — several
 * recipients' grants folded under one permission resource rather than one entry each — tolerated
 * as an equivalent echo (2026-09-02 re-review MAJOR 2) so that shape drift degrades to acceptance,
 * not a false outage.
 */
interface GraphInviteResult {
  value?: Array<
    {
      grantedToV2?: { user?: { id?: string } | null } | null;
      grantedToIdentitiesV2?: ReadonlyArray<{ user?: { id?: string } | null } | null> | null;
    } | null
  > | null;
}

/**
 * The narrow shape resolveMentions actually needs from a members cache — deliberately an
 * interface, not the concrete `MembersCache` class, so a test that has no reason to touch mention
 * resolution can wire a trivial in-memory double instead of a real disk-backed cache.
 */
export interface MembersCachePort {
  get(chatId: string): ChatMember[] | undefined;
  set(chatId: string, members: ChatMember[]): void;
}

export interface GraphTeamsChatsOptions {
  /** OneDrive folder (under the account's drive root) where outbound files are uploaded. */
  uploadDir?: string;
  /**
   * Disk-persisted per-chat membership cache — see members-cache.ts's doc comment for why this
   * exists (0.4.1: a shared Graph throttle budget on `/members` starved mention resolution).
   * REQUIRED, not optional: an optional cache let production silently fall back to the throttled
   * endpoint with no test able to catch a dropped wire (0.4.1 review round 1) — every caller,
   * production or test, must say explicitly what it wants here.
   */
  membersCache: MembersCachePort;
  /**
   * Disk-persisted cache for the signed-in account's own AAD id — see self-id-cache.ts and
   * resolveSelfId's own doc comment for the incident this exists to fix (0.4.3, live-diagnosed
   * 2026-09-03: a throttled `/me` failing on every fresh CLI process, because the in-memory memo
   * this cache backs up dies with each process). Optional, unlike membersCache above: omitting
   * it degrades to the pre-0.4.3 behaviour (memo-only, gone at process exit) rather than
   * reintroducing a correctness hazard — there is no throttled endpoint a missing wire falls
   * back to here, only a missed optimisation, so the 0.4.1-review reasoning that made
   * membersCache required does not apply. Defaults to NullSelfIdCache.
   */
  selfIdCache?: SelfIdCachePort;
  /**
   * TEAMS_MCP_SELF_ID — the last-resort operator seed for a throttle bad enough that even the
   * persisted cache never got its first write. Wins over BOTH the persisted cache and a live
   * `/me` call; never itself written to the persisted cache (an operator removing the override
   * later must not find a stale copy of it surviving there). Validated as GUID-shaped before it
   * reaches here — see config.ts's selfIdOverrideFrom.
   */
  selfIdOverride?: string;
  /**
   * Low-volume operational line, same shape as InboxPoller's own `log` option (inbox.ts).
   * Defaults to a no-op so no test needs to wire one. Used only to make resolveSelfId's
   * cache-vs-live-/me choice visible to an operator — never carries the resolved id or the
   * override value itself beyond what error messages elsewhere already surface.
   */
  log?: (line: string) => void;
}

export const DEFAULT_UPLOAD_DIR = 'ai-test';

interface GraphMember {
  displayName?: string | null;
  email?: string | null;
  /** The AAD user id — the field a mention actually needs; distinct from the member's own `id`
   *  (a composite membership identifier Graph rejects if used as a mention target). */
  userId?: string | null;
}

interface GraphChat {
  id?: string;
  topic?: string | null;
  chatType?: string;
  lastUpdatedDateTime?: string;
  members?: GraphMember[] | null;
}

function toChatMember(member: GraphMember): ChatMember | undefined {
  return member.displayName
    ? { displayName: member.displayName, ...(member.userId ? { id: member.userId } : {}) }
    : undefined;
}

interface GraphPinnedMessage {
  id?: string;
  message?: {
    id?: string;
    body?: { contentType?: string; content?: string } | null;
  } | null;
}

function toPinnedMessage(entry: GraphPinnedMessage): PinnedMessage | undefined {
  if (!entry.id) {
    return undefined;
  }
  const body = entry.message?.body;
  const content = body?.content ?? '';
  const preview = body?.contentType === 'html' ? htmlToText(content) : content.trim();
  return { id: entry.id, ...(entry.message?.id ? { messageId: entry.message.id } : {}), preview };
}

export class GraphTeamsChats implements TeamsChatsPort {
  private readonly uploadDir: string;
  private readonly membersCache: MembersCachePort;
  private readonly selfIdCache: SelfIdCachePort;
  private readonly selfIdOverride: string | undefined;
  private readonly log: (line: string) => void;
  /** Memoized by resolveSelfId, but ONLY on success — see its doc comment for why a failed
   *  lookup must NOT stick for the instance's lifetime. */
  private selfId: string | undefined;

  constructor(
    private readonly graph: GraphClient,
    options: GraphTeamsChatsOptions,
  ) {
    this.uploadDir = options.uploadDir?.trim() || DEFAULT_UPLOAD_DIR;
    this.membersCache = options.membersCache;
    this.selfIdCache = options.selfIdCache ?? new NullSelfIdCache();
    this.selfIdOverride = options.selfIdOverride;
    this.log = options.log ?? (() => {});
  }

  async listChats(): Promise<ChatSummary[]> {
    const chats = await this.graph.getAll<GraphChat>('/me/chats?$expand=members&$top=50');
    return chats.flatMap((chat) => {
      if (!chat.id) {
        return [];
      }
      const members = (chat.members ?? []).flatMap((member) => {
        const mapped = toChatMember(member);
        return mapped ? [mapped] : [];
      });
      const names = members.map((member) => member.displayName);
      return [
        {
          id: chat.id,
          // Group chats often have no topic; the member list is the only way to recognise them.
          topic: chat.topic ?? names.join(', ') ?? '',
          chatType: chat.chatType ?? 'unknown',
          ...(chat.lastUpdatedDateTime ? { lastUpdatedDateTime: chat.lastUpdatedDateTime } : {}),
          members,
        },
      ];
    });
  }

  /**
   * The chat's current members, id included — used only to resolve @mentions. A dedicated fetch
   * (not listChats, which lists and expands members for EVERY allowlisted chat) because a send
   * with mentions only ever needs one chat's roster. getAll, not a single get: a large chat's
   * roster pages, and a single-page fetch used to silently truncate it — a member past page 1
   * would come back "No chat member matches", indistinguishable from them genuinely not being in
   * the chat (review round 2, 2026-08-26).
   */
  private async membersOf(chatId: string): Promise<ChatMember[]> {
    const raw = await this.graph.getAll<GraphMember>(`/chats/${encodeURIComponent(chatId)}/members`);
    return raw.flatMap((member) => {
      const mapped = toChatMember(member);
      return mapped ? [mapped] : [];
    });
  }

  /**
   * A miss deserves a refresh; a definitively wrong CALL does not. Only "not found" (the cached
   * roster genuinely lacks this name) and "no AAD id" (the cached member has one, but Graph never
   * reported an id for them) are staleness shapes a fresh roster might fix. "Ambiguous" (matches
   * more than one member) and "empty name" are caller errors a refresh cannot fix — refreshing
   * anyway would spend a Graph call and delay the real error behind a THROTTLED one if the
   * endpoint happens to be down (0.4.1 review round 1: the original catch was unconditional).
   */
  private static isRefreshWorthy(error: unknown): boolean {
    return (
      error instanceof Error &&
      (error.message.startsWith('No chat member matches') || /has no AAD user id on record/.test(error.message))
    );
  }

  /**
   * A single, never-retried `/members` refresh, with the 429→THROTTLED translation shared by
   * BOTH resolveMentions and membersForInvite — those two used to carry nearly identical but
   * independently-worded catch blocks (2026-09-02 review MINOR: one owner now). `reason` is the
   * only thing that differs between callers, folded into one message; `extraOnThrottle`, when
   * given, is appended verbatim — resolveMentions uses it to keep its own specific reassurance
   * ("this does not mean the name does not exist") alive through the shared helper instead of
   * silently dropping it (2026-09-02 re-review MINOR: the first extraction lost this exact clause,
   * the same shape of regression as 0.4.1 losing Retry-After off the MCP tool path — pinned by a
   * test this time, not just a doc comment promise). Same rule either way: the wait itself is NOT
   * stated in the message text — retryAfterSeconds (the 4th constructor argument) is the one place
   * that number lives; retryAfterSuffix (graph-client.ts) is the only renderer, shared by the CLI
   * and the MCP tool path (0.4.1 review round 2).
   */
  private async refreshMembers(
    chatId: string,
    reason: string,
    extraOnThrottle = '',
  ): Promise<ChatMember[]> {
    return this.membersOf(chatId).catch((caught: unknown) => {
      if (caught instanceof GraphError && caught.status === 429) {
        throw new GraphError(
          `THROTTLED: the member list refresh ${reason} was throttled; nothing was done. Wait ` +
            `and try again.${extraOnThrottle}`,
          429,
          'MembersRefreshThrottled',
          caught.retryAfterSeconds,
        );
      }
      throw caught;
    });
  }

  /**
   * A fresh roster is worth caching only when it is non-empty (2026-09-02 review NIT, generalised
   * from membersForInvite's own reasoning to the one place both callers share): an empty answer —
   * whether from a genuinely memberless chat or some transient gap — is a non-answer, not a
   * confirmed fact worth trusting until the TTL expires. Applying this in ONE place means
   * resolveMentions and membersForInvite can never drift into inconsistent caching policies again.
   */
  private cacheIfNonEmpty(chatId: string, members: ChatMember[]): void {
    if (members.length > 0) {
      this.membersCache.set(chatId, members);
    }
  }

  /**
   * Cache-first: a hit resolves every name against the on-disk roster with ZERO Graph calls. A
   * miss — no cache, an expired entry, or a name the cached roster does not have — refreshes ONCE
   * (a single `/members` call, never a retry loop) and re-checks against the fresh roster; a name
   * still unresolved after that gets resolveMentionTargets's own clear error. A 429 on that
   * refresh is never swallowed into "no match" — it surfaces as its own THROTTLED error naming
   * Retry-After, so a caller sees a throttle, not a false "no such member" (0.4.1, live-diagnosed:
   * this endpoint shares a throttle budget with another daemon on the same first-party client id).
   */
  async resolveMentions(chatId: string, names: readonly string[]): Promise<MentionTarget[]> {
    const cache = this.membersCache;
    const cached = cache.get(chatId);
    if (cached) {
      try {
        return resolveMentionTargets(names, cached);
      } catch (error) {
        if (!GraphTeamsChats.isRefreshWorthy(error)) {
          throw error; // ambiguous / empty name: a refresh cannot fix a caller error
        }
        // Fall through to a single refresh — see doc comment above.
      }
    }
    const fresh = await this.refreshMembers(
      chatId,
      'for mention resolution',
      ' This does not mean the name does not exist.',
    );
    this.cacheIfNonEmpty(chatId, fresh);
    return resolveMentionTargets(names, fresh);
  }

  /**
   * The chat's current member roster for sendFile's permission grant — the SAME cache-first,
   * refresh-once-on-miss path resolveMentions uses (see its doc comment for the full rationale):
   * a cache hit costs zero Graph calls, a miss refreshes ONCE via membersOf and persists it (via
   * cacheIfNonEmpty — see its own doc comment for why an empty result is never persisted). Never
   * calls `/chats/{id}/members` directly on the send path — that endpoint shares a Graph throttle
   * budget across every process signed in with the same client id, and the 0.4.1 incident that
   * starved mention resolution on it applies identically to a file share.
   *
   * Deliberately does NOT treat "empty" as "nobody to invite": an empty roster — whether the
   * cache never had one, or a fresh call to Graph itself reports nobody — is read as "we do not
   * reliably know who is in this chat" and left for sendFile to fail loudly on, never silently
   * skipped as if the file were being shared into an empty room.
   */
  private async membersForInvite(chatId: string): Promise<ChatMember[]> {
    const cached = this.membersCache.get(chatId);
    if (cached && cached.length > 0) {
      return cached;
    }
    const fresh = await this.refreshMembers(chatId, 'needed to grant file access');
    this.cacheIfNonEmpty(chatId, fresh);
    return fresh;
  }

  /**
   * The assistant's own AAD id — used only so sendFile can exclude the assistant from its own
   * permission grant (it already owns the uploaded item as the uploader; granting itself `read`
   * on top would be harmless but noisy, not wrong). Returning `undefined` here does NOT itself
   * throw — the caller (sendFile) decides what an undetermined self id means: a hard PRE-UPLOAD
   * refusal, never a soft "assume it's fine" fallback — see sendFile's own doc comment for the
   * fuller history of why.
   *
   * Resolution order (0.4.3, live-diagnosed 2026-09-03 — eight consecutive `teams-send-file` CLI
   * attempts over 20 minutes each failed this lookup, because every CLI invocation is a fresh
   * process and the in-memory memo below dies with it, paying and losing a throttled `/me` call
   * every single time; the account's own id never changes, so it is now worth persisting):
   *  1. `selfIdOverride` (TEAMS_MCP_SELF_ID, GUID-validated in config.ts) — a last-resort operator
   *     seed, wins outright: no persisted-cache read, no `/me` call, never written back to the
   *     cache (an operator who unsets the override later must not find a stale copy surviving on
   *     disk).
   *  2. The in-memory memo (`this.selfId`) from an earlier successful resolution THIS instance
   *     already made.
   *  3. The persisted cache (`selfIdCache`, self-id-cache.ts) — no TTL, since the account's own
   *     id does not change. A hit here is used AS IS, with ZERO `/me` calls: the whole point is
   *     that a fresh CLI process never has to pay the throttled lookup again once any process has
   *     resolved it once. Known limitation, out of scope for 0.4.3: if a later Graph call ever
   *     proved this cached id wrong (e.g. a 403 naming a different principal after an account
   *     swap), nothing here invalidates it — the cache is trusted until the file is deleted by
   *     hand.
   *  4. `GET /me?$select=id` (`readRetries: 0` — see below) — only reached when steps 1–3 all
   *     miss. On success, both the in-memory memo and the persisted cache are written before
   *     returning.
   * Only when step 4 ALSO fails (readRetries: 0's 429, or any other error) does this return
   * `undefined` — memoized nowhere, so the NEXT call (the next sendFile/delete, or a fresh
   * process with a still-cold cache) retries it rather than being stuck on one transient failure
   * for the rest of an instance's lifetime (BLOCKER, 2026-09-02 review, live-probed).
   *
   * `readRetries: 0`: a 429 here used to cost a real Retry-After sleep before failing anyway
   * (2026-09-02 review MAJOR 1 follow-up) — now that an undetermined self id refuses the whole
   * send regardless of why, there is no value in sleeping through the throttle only to still
   * refuse; fail fast and let the caller retry the whole `sendFile` call once the throttle
   * clears, same as `fetchMessage`'s `readRetries: 0` (a different rationale there — a cheaper
   * fallback exists — but the same "don't sleep for nothing" conclusion). The persisted cache is
   * exactly what makes that retry cheap now instead of paying another throttled `/me`.
   */
  private async resolveSelfId(): Promise<string | undefined> {
    if (this.selfIdOverride !== undefined) {
      return this.selfIdOverride;
    }
    if (this.selfId !== undefined) {
      return this.selfId;
    }
    const cached = this.selfIdCache.read();
    if (cached) {
      this.selfId = cached.id;
      this.log('self id served from the persisted cache; /me not called.');
      return this.selfId;
    }
    try {
      const me = await this.graph.get<{ id?: string }>('/me?$select=id', { readRetries: 0 });
      if (me.id) {
        this.selfId = me.id;
        this.selfIdCache.write({ id: me.id, resolvedAt: Date.now() });
      }
    } catch {
      return undefined; // best-effort, deliberately NOT memoized — see doc comment above
    }
    return this.selfId;
  }

  async readMessages(chatId: string, since?: string, limit = 50): Promise<ReadResult> {
    const raw = await this.graph.getAll<unknown>(
      `/chats/${encodeURIComponent(chatId)}/messages?$top=${Math.min(limit, 50)}`,
      limit,
    );
    return applyWatermark(
      raw.map((message) => toChatMessage(message, chatId)),
      since,
    );
  }

  async sendMessage(
    chatId: string,
    text: string,
    mentions: readonly MentionTarget[] = [],
  ): Promise<ChatMessage> {
    // Always HTML: a 'text' body renders in Teams as one unbroken blob — no line breaks, no
    // clickable links. textToHtml escapes everything, so plain text stays plain. With mentions,
    // renderTextWithMentions does the same escaping plus <at> tags at every occurrence of each
    // mention's name — see its doc comment in mentions.ts.
    const created = await this.graph.post<unknown>(
      `/chats/${encodeURIComponent(chatId)}/messages`,
      {
        body: { contentType: 'html', content: renderTextWithMentions(text, mentions) },
        ...(mentions.length > 0 ? { mentions: buildGraphMentionsPayload(mentions) } : {}),
      },
    );
    return toChatMessage(created, chatId);
  }

  async sendHtmlMessage(
    chatId: string,
    html: string,
    mentions: readonly MentionTarget[] = [],
  ): Promise<ChatMessage> {
    // No textToHtml here — html IS the body content, posted exactly as given, except that any
    // `@{Name}` placeholder the caller wrote gets swapped for the matching <at> tag — see
    // TeamsChatsPort.sendHtmlMessage's doc comment for why this exists alongside sendMessage.
    const created = await this.graph.post<unknown>(
      `/chats/${encodeURIComponent(chatId)}/messages`,
      {
        body: { contentType: 'html', content: renderHtmlWithMentions(html, mentions) },
        ...(mentions.length > 0 ? { mentions: buildGraphMentionsPayload(mentions) } : {}),
      },
    );
    return toChatMessage(created, chatId);
  }

  async sendImage(chatId: string, image: OutboundImage, text?: string): Promise<ChatMessage> {
    // The temporary-id pattern: the img src refers to hostedContents entry "1" by its
    // temporaryId, and Graph rewrites the src to the real hosted content URL on creation.
    const created = await this.graph.post<unknown>(
      `/chats/${encodeURIComponent(chatId)}/messages`,
      {
        body: {
          contentType: 'html',
          content:
            (text ? textToHtml(text) : '') +
            '<img src="../hostedContents/1/$value" alt="image">',
        },
        hostedContents: [
          {
            '@microsoft.graph.temporaryId': '1',
            contentBytes: Buffer.from(image.bytes).toString('base64'),
            contentType: image.contentType,
          },
        ],
      },
    );
    return toChatMessage(created, chatId);
  }

  /**
   * Uploads `file` to the account's OneDrive and shares it into the chat as a reference
   * attachment.
   *
   * BUG FIXED (0.4.2, live-verified 2026-09-02): the upload alone leaves the driveItem readable
   * by nobody but the assistant — every chat member's card answered "can't be viewed or
   * downloaded" until a human ran a manual `POST /me/drive/items/{id}/invite` by hand. This method
   * now grants each OTHER chat member read access on the uploaded item (that same invite call,
   * `requireSignIn: true`, `sendInvitation: false`, `roles: ['read']`) BEFORE the chat message is
   * posted — Teams' own native upload does this implicitly; we have to do it explicitly. See
   * KNOWN-ISSUES.md for the live-verified wire-shape snapshot this contract is anchored to. The
   * roster comes from membersForInvite, the SAME cache-backed path resolveMentions uses — never a
   * direct call to the throttled `/chats/{id}/members` endpoint on the send path (see that
   * method's doc comment for why: the 0.4.1 mention-429 incident is the same shared throttle
   * budget).
   *
   * Self-exclusion: the assistant's own id is excluded from the grant when resolveSelfId can
   * determine it — it already owns the item as the uploader, so granting itself `read` on top
   * would be harmless but noisy. When it CANNOT be determined (resolveSelfId returns undefined —
   * a `/me` outage, a transient 403, …), this REFUSES the send before ever touching OneDrive
   * (2026-09-02 review MAJOR 1) — the earlier "include everyone, self included" soft fallback
   * looked harmless in isolation but meant a `/me` outage orphaned one upload attempt per
   * `sendFile` call for its entire duration, 100% failure with no clean pre-upload signal. Before
   * that, an even earlier BLOCKER version silently excluded whichever member merely happened to
   * have no `id` (mistaking an id-less REAL member for "self" because `undefined !== undefined`
   * is false), letting their dead card go out while the assistant granted itself access instead —
   * live-probed 2026-09-02. This version's structure makes that class of bug impossible again: the
   * pre-upload refusal below narrows `selfId` to a definite `string` for everything that follows,
   * so there is no code path left where "self" and "undetermined" can be confused. See
   * resolveSelfId's own doc comment for the retry/no-memoize-on-failure behaviour this refusal
   * pairs with.
   *
   * Failure contract — no dead cards, ever:
   *  - an unresolvable or empty member roster (cache empty/expired AND the refresh fails or comes
   *    back empty) throws BEFORE the OneDrive upload — nothing is wasted;
   *  - an undetermined self id (the `/me` lookup itself failed) ALSO throws BEFORE the upload,
   *    for the reason above;
   *  - ANY other member (after self-exclusion) with no AAD id Graph reported also throws BEFORE
   *    the upload — a MIXED roster (some resolvable, some not) is refused whole, never granted
   *    only to the resolvable ones while posting anyway (2026-09-02 review MAJOR: silently
   *    partial grants are exactly the "no dead cards" contract this method exists to keep);
   *  - a successful upload followed by a FAILED `/invite` call throws AFTER the upload (which
   *    cannot be undone from here, so the uploaded item is left orphaned in OneDrive) and BEFORE
   *    the chat message post — a loud, diagnosable failure is preferred over a card recipients
   *    cannot open;
   *  - a `/invite` call that answers HTTP success is NOT trusted by status alone: Graph can
   *    report 200 with an empty or partial `value` array when a grant did not actually land
   *    (2026-09-02 review MAJOR — the response's own `grantedToV2.user.id` entries, OR the
   *    equivalent `grantedToIdentitiesV2[].user.id` shape some permission kinds use instead
   *    (2026-09-02 re-review MAJOR 2), are the source of truth — see KNOWN-ISSUES.md for both
   *    live-verified wire shapes this reads against). Any recipient missing from those grants
   *    throws, same as an outright /invite failure, same orphan-honesty;
   *  - a chat whose only member is the assistant itself (no OTHER member at all) skips the invite
   *    call entirely and sends normally — there is nobody who could see a dead card.
   */
  async sendFile(chatId: string, file: OutboundFile, text?: string): Promise<ChatMessage> {
    const members = await this.membersForInvite(chatId);
    if (members.length === 0) {
      throw new Error(
        `Cannot share ${file.name} into chat ${chatId}: the chat's member list resolved to ` +
          'empty, so no recipient permission grant could be attempted. Nothing was uploaded.',
      );
    }
    const selfId = await this.resolveSelfId();
    if (selfId === undefined) {
      throw new Error(
        `Cannot share ${file.name} into chat ${chatId}: the assistant's own account id could ` +
          "not be determined (the /me lookup failed), so members could not be safely excluded " +
          'from the read permission grant. Nothing was uploaded — try again once /me is reachable.',
      );
    }
    // selfId is a definite string from here on — see the method doc comment for why that
    // structurally rules out the earlier self/id-less-member confusion.
    const others = members.filter((member) => member.id !== selfId);
    const unresolvable = others.filter((member) => !member.id);
    if (unresolvable.length > 0) {
      throw new Error(
        `Cannot share ${file.name} into chat ${chatId}: ${unresolvable.length} of this chat's ` +
          `${others.length} other member(s) — ${unresolvable.map((member) => member.displayName).join(', ')} ` +
          '— has no AAD id Graph reported, so a read permission grant could not be attempted for ' +
          'them. Nothing was uploaded.',
      );
    }
    const recipientIds = others.map((member) => member.id as string);

    // A chat cannot host a real file; it has to live in the sender's OneDrive first, and the
    // message then carries a reference attachment pointing at it.
    const folder = this.uploadDir.split('/').map(encodeURIComponent).join('/');
    const item = await this.graph.putBinary<GraphDriveItem>(
      `/me/drive/root:/${folder}/${encodeURIComponent(file.name)}:/content`,
      file.bytes,
      file.contentType ?? 'application/octet-stream',
    );

    // The attachment id must be the GUID inside the driveItem's eTag ("{GUID},n") — the
    // driveItem id itself is a different identifier and Teams rejects it as the attachment id
    // (it IS, however, the right id for the /invite call below).
    const guid = /\{([0-9a-fA-F-]+)\}/.exec(item.eTag ?? '')?.[1];
    if (!guid || !item.webUrl || !item.id) {
      throw new Error(
        `OneDrive upload of ${file.name} returned no usable id/eTag/webUrl; cannot build the ` +
          'reference attachment or grant access.',
      );
    }

    if (recipientIds.length > 0) {
      let inviteResult: GraphInviteResult;
      try {
        inviteResult = await this.graph.post<GraphInviteResult>(
          `/me/drive/items/${encodeURIComponent(item.id)}/invite`,
          {
            recipients: recipientIds.map((id) => ({ objectId: id })),
            requireSignIn: true,
            sendInvitation: false,
            roles: ['read'],
          },
        );
      } catch (caught) {
        throw new Error(
          `Uploaded ${file.name} to OneDrive but could not grant chat members read access on it ` +
            `(${caught instanceof Error ? caught.message : String(caught)}) — nothing was posted ` +
            'to the chat, since a card the recipients cannot open is worse than none at all. The ' +
            `uploaded item (id ${item.id}) is orphaned in OneDrive and was not cleaned up.`,
        );
      }
      // HTTP success alone is not proof of a grant — see the method doc comment and
      // KNOWN-ISSUES.md's wire-shape snapshot for the live evidence this reads against.
      const grantedIds = new Set(
        (inviteResult.value ?? []).flatMap((permission) => {
          const direct = permission?.grantedToV2?.user?.id;
          // grantedToIdentitiesV2 is a documented Graph variant — see GraphInviteResult's doc
          // comment — some permission kinds use instead of one grantedToV2 entry per recipient.
          const viaIdentities = (permission?.grantedToIdentitiesV2 ?? []).flatMap((identity) => {
            const id = identity?.user?.id;
            return id ? [id] : [];
          });
          return [...(direct ? [direct] : []), ...viaIdentities];
        }),
      );
      const missing = recipientIds.filter((id) => !grantedIds.has(id));
      if (missing.length > 0) {
        throw new Error(
          `Uploaded ${file.name} to OneDrive and called /invite (HTTP success), but its response ` +
            `lists no grant for ${missing.length} of ${recipientIds.length} recipient(s) — ` +
            'nothing was posted to the chat, since a card those recipients cannot open is worse ' +
            `than none at all. The uploaded item (id ${item.id}) is orphaned in OneDrive and was ` +
            'not cleaned up.',
        );
      }
    }

    const created = await this.graph.post<unknown>(
      `/chats/${encodeURIComponent(chatId)}/messages`,
      {
        body: {
          contentType: 'html',
          content:
            (text ? textToHtml(text) : '') +
            `<attachment id="${guid}"></attachment>`,
        },
        attachments: [
          {
            id: guid,
            contentType: 'reference',
            contentUrl: item.webUrl,
            name: item.name ?? file.name,
          },
        ],
      },
    );
    return toChatMessage(created, chatId);
  }

  /**
   * One message by id. The single-message endpoint is throttled on its own budget — on
   * 2026-08-25 it answered 429 for hours while the list endpoint for the same chat stayed
   * healthy, which silently broke every quoted reply and every attachment download. So a 429
   * there falls back to scanning the recent list for the id: a page costs one healthy request
   * and the message we want is almost always among the last fifty.
   */
  private async fetchMessage(chatId: string, messageId: string): Promise<ChatMessage> {
    try {
      return toChatMessage(
        await this.graph.get<unknown>(
          `/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}`,
          { readRetries: 0 }, // the list fallback below is cheaper than sleeping and re-hitting a throttled family
        ),
        chatId,
      );
    } catch (caught) {
      if (!(caught instanceof GraphError) || caught.status !== 429) {
        throw caught;
      }
      const recent = await this.graph.get<{ value?: unknown[] }>(
        `/chats/${encodeURIComponent(chatId)}/messages?$top=50`,
      );
      const found = (recent.value ?? [])
        .map((raw) => toChatMessage(raw, chatId))
        .find((message) => message.id === messageId);
      if (!found) {
        // Same rule as MembersRefreshThrottled above: the wait itself is not stated in the
        // message text — retryAfterSeconds is the one place it lives, rendered once by
        // retryAfterSuffix (0.4.1 review round 2).
        throw new GraphError(
          `Message ${messageId} could not be fetched (that endpoint is throttled) and is not ` +
            'among the chat\'s last 50 messages — nothing was posted.',
          429,
          'MessageFetchThrottled',
          caught.retryAfterSeconds,
        );
      }
      return found;
    }
  }

  async replyToMessage(
    chatId: string,
    replyToMessageId: string,
    text: string,
    mentions: readonly MentionTarget[] = [],
  ): Promise<ChatMessage> {
    // Chats have no reply threads (that is channels-only). What the Teams UI calls a reply in a
    // chat is a new message carrying a messageReference attachment - a quote card built from the
    // original message - so that is what gets posted here.
    const original = await this.fetchMessage(chatId, replyToMessageId);

    const reference = JSON.stringify({
      messageId: original.id,
      messagePreview: original.text.slice(0, 150),
      messageSender: {
        application: null,
        device: null,
        user: {
          userIdentityType: 'aadUser',
          id: original.fromId ?? null,
          displayName: original.from,
        },
      },
    });

    const created = await this.graph.post<unknown>(
      `/chats/${encodeURIComponent(chatId)}/messages`,
      {
        body: {
          contentType: 'html',
          content: `<attachment id="${original.id}"></attachment>${renderTextWithMentions(text, mentions)}`,
        },
        attachments: [
          { id: original.id, contentType: 'messageReference', content: reference },
        ],
        ...(mentions.length > 0 ? { mentions: buildGraphMentionsPayload(mentions) } : {}),
      },
    );
    return toChatMessage(created, chatId);
  }

  async editMessage(
    chatId: string,
    messageId: string,
    newText: string,
    mentions: readonly MentionTarget[] = [],
  ): Promise<void> {
    // Graph only lets the delegated user edit messages they sent themselves; anything else
    // comes back as a Graph error, which the caller sees verbatim.
    await this.graph.patch(
      `/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}`,
      {
        body: { contentType: 'html', content: renderTextWithMentions(newText, mentions) },
        ...(mentions.length > 0 ? { mentions: buildGraphMentionsPayload(mentions) } : {}),
      },
    );
  }

  async editHtmlMessage(
    chatId: string,
    messageId: string,
    html: string,
    mentions: readonly MentionTarget[] = [],
  ): Promise<void> {
    // Same verbatim contract as sendHtmlMessage — no textToHtml, no escaping.
    await this.graph.patch(
      `/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}`,
      {
        body: { contentType: 'html', content: renderHtmlWithMentions(html, mentions) },
        ...(mentions.length > 0 ? { mentions: buildGraphMentionsPayload(mentions) } : {}),
      },
    );
  }

  async deleteMessage(chatId: string, messageId: string): Promise<void> {
    // The reversible soft delete — the message becomes "This message was deleted" in Teams and
    // can be restored. Hard delete is deliberately not offered here.
    await this.graph.postAction(
      `/me/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}/softDelete`,
    );
  }

  async setReaction(chatId: string, messageId: string, reactionType: string): Promise<void> {
    // The payload is {reactionType} at the top level — NOT wrapped in {body: …} like a message
    // post. Graph answers 204 No Content, hence postNoContent. Both halves of this comment are
    // paid-for knowledge: the 2026-08-24 ad-hoc script got each of them wrong in turn.
    await this.graph.postNoContent(
      `/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}/setReaction`,
      { reactionType },
    );
  }

  async getAttachment(
    chatId: string,
    messageId: string,
    attachmentId?: string,
  ): Promise<AttachmentPayload> {
    const parsed = await this.fetchMessage(chatId, messageId);
    // Without an explicit id, "the attachment" means the first one that IS a file — a quoted
    // reply carries a messageReference card first, and a caller asking for "the attachment" on
    // "Logo: <file>" never means the quote (2026-08-25: that pick answered 404 from hostedContents).
    const attachment = attachmentId
      ? parsed.attachments.find((candidate) => candidate.id === attachmentId)
      : parsed.attachments.find((candidate) => candidate.contentType !== 'messageReference');

    if (!attachment) {
      const onlyQuote = parsed.attachments.length > 0;
      throw new Error(
        attachmentId
          ? `Message ${messageId} has no attachment ${attachmentId}.`
          : onlyQuote
            ? `Message ${messageId} has no downloadable attachment — only a quoted-message card.`
            : `Message ${messageId} has no attachments.`,
      );
    }

    return this.downloadAttachment(chatId, messageId, attachment);
  }

  async listAttachments(chatId: string, messageId: string): Promise<ChatAttachmentRef[]> {
    // fetchMessage's throttle fallback (list scan on a 429) applies here too — a metadata
    // listing must not be more fragile than the download it usually precedes.
    return (await this.fetchMessage(chatId, messageId)).attachments;
  }

  async getAttachments(
    chatId: string,
    messageId: string,
    nameFilter?: string,
  ): Promise<AttachmentPayload[]> {
    const parsed = await this.fetchMessage(chatId, messageId);
    const downloadable = parsed.attachments.filter(
      (candidate) => candidate.contentType !== 'messageReference',
    );
    const wanted = nameFilter
      ? downloadable.filter((candidate) =>
          (candidate.name ?? '').toLowerCase().includes(nameFilter.toLowerCase()),
        )
      : downloadable;

    if (wanted.length === 0) {
      throw new Error(
        parsed.attachments.length === 0
          ? `Message ${messageId} has no attachments.`
          : downloadable.length === 0
            ? `Message ${messageId} has no downloadable attachment — only a quoted-message card.`
            : `Message ${messageId} has no attachment whose name contains "${nameFilter}" — ` +
              `it has: ${downloadable.map((candidate) => candidate.name ?? candidate.id).join(', ')}.`,
      );
    }

    // Sequential on purpose: these downloads all land on the same throttle families (see the
    // gate doc in graph-client.ts), and firing them in parallel is how one message's worth of
    // files turns a shared per-mailbox budget into a closed gate for everyone.
    const payloads: AttachmentPayload[] = [];
    for (const attachment of wanted) {
      payloads.push(await this.downloadAttachment(chatId, messageId, attachment));
    }
    return payloads;
  }

  /**
   * One attachment's bytes. Two shapes exist. A pasted image is hosted content on the message
   * itself; a shared file is a driveItem in SharePoint/OneDrive whose contentUrl must be
   * resolved through the /shares facade — see shareIdFor.
   */
  private async downloadAttachment(
    chatId: string,
    messageId: string,
    attachment: ChatAttachmentRef,
  ): Promise<AttachmentPayload> {
    if (!attachment.contentUrl) {
      const hosted = await this.graph.getBinary(
        `/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}` +
          `/hostedContents/${encodeURIComponent(attachment.id)}/$value`,
      );
      return { ...hosted, name: attachment.name ?? `${attachment.id}.bin` };
    }

    let downloaded: { bytes: Uint8Array; contentType: string };
    try {
      downloaded = await this.graph.getBinary(
        `/shares/${shareIdFor(attachment.contentUrl)}/driveItem/content`,
      );
    } catch (caught) {
      // A 403 here is a permission story, not a code bug, and the caller deserves the whole
      // story in one message: with the default setup (a Microsoft first-party client id and the
      // `.default` scope) this download works out of the box — live-verified 2026-09-02. The
      // 403 shows up when TEAMS_MCP_CLIENT_ID points at a custom app registration whose token
      // never asked for (or was never consented) the file-read permission.
      if (caught instanceof GraphError && caught.status === 403 && !caught.isLicenceProblem) {
        throw new GraphError(
          `SharePoint refused the download of "${attachment.name ?? attachment.id}" (403` +
            `${caught.code ? ` ${caught.code}` : ''}): the signed-in token cannot read the shared ` +
            'file. With the default Microsoft first-party client id this works without any ' +
            'setup; if TEAMS_MCP_CLIENT_ID is a custom app registration, that registration ' +
            'needs the delegated Files.Read.All (or Sites.Read.All) Graph permission, and ' +
            'granting it may require admin consent. Graph said: ' +
            caught.message,
          caught.status,
          caught.code,
          caught.retryAfterSeconds,
        );
      }
      throw caught;
    }
    // attachment.contentType is usually the literal "reference", not a media type; the download
    // response knows better.
    const declared = attachment.contentType?.includes('/') ? attachment.contentType : undefined;
    return {
      bytes: downloaded.bytes,
      contentType: declared ?? downloaded.contentType,
      name: attachment.name ?? `${attachment.id}.bin`,
    };
  }

  async listPinnedMessages(chatId: string): Promise<PinnedMessage[]> {
    let raw: { value?: GraphPinnedMessage[] };
    try {
      raw = await this.graph.get<{ value?: GraphPinnedMessage[] }>(
        `/chats/${encodeURIComponent(chatId)}/pinnedMessages?$expand=message`,
      );
    } catch (caught) {
      // Empirical, not documented: GET .../pinnedMessages on a chat with ZERO pins answers a
      // bare 404 "NotFound" rather than 200 with an empty value array (verified live
      // 2026-08-26 — pinning, then unpinning the chat's only pin, then listing again 404s where
      // every other call on the same chatId keeps succeeding). Any OTHER caller of this method
      // already went through the allowlist, so a 404 here is read as "nothing pinned", not "chat
      // does not exist" — treating it as a real failure would make list_pinned_messages error on
      // the single most common case: a chat with nothing pinned.
      if (caught instanceof GraphError && caught.status === 404) {
        return [];
      }
      throw caught;
    }
    return (raw.value ?? []).flatMap((entry) => {
      const mapped = toPinnedMessage(entry);
      return mapped ? [mapped] : [];
    });
  }

  async pinMessage(chatId: string, messageId: string): Promise<PinnedMessage[]> {
    // A chat effectively holds ONE pin: pinning a second message silently REPLACES the first
    // while this POST reports success either way (verified live 2026-08-25) — see
    // TeamsChatsPort.pinMessage's doc comment. Re-listing after the POST, rather than trusting
    // its own response, is what lets the caller see that replacement happen instead of believing
    // both messages are pinned.
    await this.graph.post<unknown>(`/chats/${encodeURIComponent(chatId)}/pinnedMessages`, {
      'message@odata.bind': `${GRAPH_BASE_URL}/chats/${encodeURIComponent(chatId)}/messages/${encodeURIComponent(messageId)}`,
    });
    return this.listPinnedMessages(chatId);
  }

  async unpinMessage(chatId: string, messageId: string): Promise<void> {
    // DELETE takes the PIN's own id, not the chat message id, so the message id has to be
    // resolved through the current pinned list first.
    const pinned = await this.listPinnedMessages(chatId);
    const match = pinned.find((entry) => entry.messageId === messageId);
    if (!match) {
      throw new Error(`Message ${messageId} is not currently pinned in chat ${chatId} — nothing to unpin.`);
    }
    await this.graph.del(
      `/chats/${encodeURIComponent(chatId)}/pinnedMessages/${encodeURIComponent(match.id)}`,
    );
  }
}

export type { ChatAttachmentRef, ChatMessage, ReadResult };
