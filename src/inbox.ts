import { appendFile, mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import type { ChatAllowlist } from './allowlist.js';
import type { TeamsChatsPort, WarmMembersResult } from './graph/teams-chats.js';
import { readYield } from './inbox-yield.js';
import type { ChatMessage } from './messages.js';

/**
 * Background inbox poller. Runs inside the MCP server process and appends every new message from
 * the allowlisted chats to a JSONL file at a stable path, so an orchestrator can arm a file
 * watcher once and stop reinventing a polling daemon per session.
 *
 * One line per message: {"chat","id","from","at","text","attachments"} — the shape the old
 * scratchpad daemon wrote, kept so existing consumers keep parsing. `text` is the FULL message
 * body, verbatim, up to MAX_INBOX_TEXT_BYTES UTF-8 bytes (a pathological-size guard, not a
 * preview budget); a message at or over that many bytes also carries `truncated: true` and an
 * explicit "…[truncated, N chars / M bytes total]" suffix on `text` — never a silent cut (0.6.2;
 * see KNOWN-ISSUES.md). A failed poll appends {"error","at","consecutiveFailures"} instead,
 * because a watcher must be able to tell auth death from a quiet chat.
 *
 * Since 0.5.4 the poller also proves it is alive: after EVERY poll (including a yielded one — see
 * writeHealth) it rewrites a small health file beside the inbox. That carries the poller-
 * supervision work from PR #8: a host reboot once killed one project's daemon while a second
 * project's daemon — same dist/index.js path, different env — survived, and a pgrep-based check
 * matched the survivor: the dead pipeline read as "up" and an allowlisted chat sat undelivered
 * for 1.5 hours. The health file's AGE is the liveness signal: it answers "is THIS inbox's
 * pipeline polling", independent of what other processes look like. The single-instance lock
 * (poller-lock.ts) closes the other half of that incident: two daemons racing on one shared
 * account fed each other's throttle penalty.
 */

export interface SignedInAccount {
  id?: string;
  displayName?: string;
}

/**
 * Narrow sink the poller hands harvested (id, displayName) pairs to — deliberately just the
 * `merge` slice of MembersCache (members-cache.ts), not the full cache, so a test double never
 * needs to implement get()/set()/getStale() to exercise the poller. Mitigation 2
 * (docs/throttling-mitigation.md §4, stage 1 item 2): every message this poller already reads
 * carries its sender's AAD id and display name at zero MARGINAL Graph cost — merging those into
 * the persisted roster is what lets a chat's roster fill from traffic alone, taking `/members` off
 * the mention-resolution send path for the common case (see GraphTeamsChats.resolveMentions's
 * stale-serve fallback, teams-chats.ts, for what still happens on the rarer miss). What this
 * builds is always a PARTIAL roster (MembersCache.merge's own doc comment) — it never drives
 * `send_chat_file`'s permission grant, which requires a COMPLETE roster from a real `/members`
 * fetch (0.5.2 BLOCKER 1 fix, see membersForInvite in teams-chats.ts).
 */
export interface RosterHarvestPort {
  merge(chatId: string, members: ReadonlyArray<{ id: string; displayName: string }>): void;
}

export interface InboxPollerDeps {
  chats: Pick<TeamsChatsPort, 'readMessages' | 'warmMembers'>;
  allowlist: ChatAllowlist;
  /**
   * Resolves who the server is signed in as. The assistant's own posts must not come back as
   * inbox events, or every send would wake the orchestrator to read its own words.
   *
   * `build-inbox-poller.ts`'s real implementation resolves this through the SAME operator-seed
   * (TEAMS_MCP_SELF_ID) -> persisted-cache -> live `/me` chain `GraphTeamsChats.resolveSelfId`
   * uses for `sendFile` (`resolveSelfIdStatus`, teams-chats.ts) — before 0.6.4 this called a raw,
   * seed/cache-blind `/me` of its own, so a tenant-wide `/me` 429 failed every poll and, after
   * three cycles, forced a token re-authentication for a condition that was never auth-shaped
   * (issue #28, live 2026-09-09). A throttled resolution is expected to REJECT with a 429-shaped
   * error (`status: 429`, optionally `retryAfterSeconds`) so `pollAllowlistedChats` classifies it
   * as a THROTTLED cycle rather than an "unrecognised shape" failure — see that method's own doc
   * comment. Any other rejection, or a resolution reporting no `id` at all, fails the whole poll
   * (nothing is delivered) without being memoized, so the NEXT poll retries it.
   */
  self: () => Promise<SignedInAccount>;
  inboxPath: string;
  statePath: string;
  /**
   * Where ad-hoc Graph readers ask this poller to go quiet — see inbox-yield.ts for the
   * measured starvation (2026-09-02) that makes the coordination necessary. Optional: without
   * it the poller never yields, which is the pre-0.5.0 behaviour.
   */
  yieldPath?: string;
  /**
   * Where harvested (senderId, senderDisplayName) pairs go — see RosterHarvestPort's own doc
   * comment (mitigation 2). Optional: a poller with no roster wired simply does not harvest,
   * exactly the pre-mitigation-2 behaviour — there is no throttled endpoint a missing wire falls
   * back to here, only a missed optimisation (same posture as GraphTeamsChatsOptions.selfIdCache,
   * self-id-cache.ts).
   */
  roster?: RosterHarvestPort;
  /** Where the health snapshot lands; defaults to poller-health.json beside the inbox. */
  healthPath?: string;
  /** The pid recorded in the health file; defaults to process.pid. Injectable for tests. */
  pid?: number;
  pollMs?: number;
  /** Injectable clock for the 403 park, the health timestamps and the escalation lines; defaults
   *  to Date.now. */
  nowFn?: () => number;
  maxBackoffMs?: number;
  /**
   * How long a chat sits out further warm-up attempts after one comes back throttled — issue (a),
   * live-diagnosed 2026-09-08: the SAME chat's roster warm-up was throttled on two consecutive
   * poll cycles just 4 minutes apart, so the pre-0.6.3 "warm once, ever" gate (a chat marked
   * attempted the moment warmMembers was first CALLED, whatever the outcome) was not actually
   * closing the gap this incident showed. This is the DEFAULT window only — a throttled attempt
   * that carries Graph's own Retry-After (WarmMembersResult.retryAfterSeconds) honours that
   * instead, same posture as noteRetryAfter's read-loop floor. TEAMS_INBOX_WARMUP_BACKOFF_SECONDS
   * is the env knob (see index.ts); defaults to DEFAULT_WARMUP_BACKOFF_MS (15 minutes).
   */
  warmupBackoffMs?: number;
  log?: (line: string) => void;
  /**
   * Called once a poll cycle ends with an auth-shaped failure (401, or a message naming
   * invalid_grant/an AADSTS code/token expiry) for the `authFailureThreshold`th CONSECUTIVE time
   * — see the 0.4.1 doc comment on trackAuthHealth for the live-diagnosed stuck-auth mode this
   * exists to break out of. Typically wired to the token provider's own `invalidate()`.
   */
  onAuthStuck?: () => void;
  /** Consecutive poll failures (of any shape — see the 0.4.1 last-resort-tier doc comment on
   *  trackAuthHealth) before onAuthStuck fires. Default 3. */
  authFailureThreshold?: number;
  /** Injectable write primitives for saveState — tests use these to prove the sidecar's
   *  destination path is reached ONLY via rename, never a direct write (0.4.1 review round 1). */
  writeFileFn?: typeof writeFile;
  renameFn?: typeof rename;
}

/**
 * Auth-shaped: the kind of failure a dead-but-not-yet-locally-expired cached token produces.
 * Matched broadly on purpose — a 401 status, this server's own AuthenticationError, or Entra's
 * usual vocabulary (invalid_grant, an AADSTS code, "token" + "expir…") — because the actual shape
 * Graph answers with when a token goes bad server-side was not pinned down from code inspection
 * alone; see trackAuthHealth's doc comment for what IS known and what remains a hypothesis.
 *
 * NOT exhaustive by design: the live incident's own log line — `inbox poll failed: <chat>: fetch
 * failed`, repeated, no status code, while parallel curl probes got 200 from Graph — matches
 * NONE of these shapes (see KNOWN-ISSUES.md's incident entry for the verbatim evidence). This
 * function is used for LOG WORDING (was this recognisably auth-related, or not) — it no longer
 * gates whether the forced re-mint fires at all; trackAuthHealth's last-resort tier covers the
 * shapeless case this function cannot name.
 */
/**
 * Pairs the two halves of "this cycle is throttled" so a call site cannot set one without the
 * other (issue c, 2026-09-08 review: the pairing used to be convention only — a throttled cycle
 * with no matching `failures` entry would slip past `pollAllowlistedChats`'s own
 * `failures.length === 0` early return and be misreported as clean; see that method's own comment
 * on the return statement). Both call sites today (the warm-up 429 and the readMessages 429) go
 * through this; a third can never forget the push half. Exported for its own direct test.
 */
export function markThrottled(failures: string[], message: string): true {
  failures.push(message);
  return true;
}

function isAuthShaped(error: unknown): boolean {
  if (!(error instanceof Error)) {
    return false;
  }
  if ((error as { status?: number }).status === 401) {
    return true;
  }
  if (error.name === 'AuthenticationError') {
    return true;
  }
  return /invalid_grant|AADSTS\d+|token[^a-z]{0,10}expir/i.test(error.message);
}

interface ChatState {
  /**
   * ISO-8601 createdDateTime of the newest delivered message; exclusive, like read watermarks.
   * Optional: a chat that has been polled at least once but never had a message yet has an
   * entry with no watermark — see the bootstrap comment in pollOnce for why this differs from
   * having no entry at all.
   */
  watermark?: string;
  /** Id of that newest message, so an exact replay is dropped even if the port re-serves it. */
  lastId?: string;
}

/**
 * What an external watcher reads to judge the pipeline. The contract: lastAttemptAt older than
 * a couple of backoffMs means the poller is dead or wedged, whatever pgrep says; ok=false with a
 * fresh lastAttemptAt means the process is alive but its polls are failing (the {error} lines in
 * the inbox say why). A cycle skipped for the quota yield (inbox-yield.ts) is neither of those —
 * it is deliberate politeness, not liveness or failure — so it is marked `yielded: true` with
 * `ok: true` and an UNCHANGED `lastSuccessAt`/`consecutiveFailures`: a watcher must not mistake a
 * yield for a fresh clean poll, and must not mistake it for the pipeline going quiet either.
 */
export interface PollerHealth {
  pid: number;
  inboxPath: string;
  lastAttemptAt: string;
  /** Stamp of the last CLEAN, non-yielded poll; absent until the first one. Survives failures
   *  (and yields), so a watcher can say how long an outage has been running. */
  lastSuccessAt?: string;
  ok: boolean;
  /** True only on a cycle skipped for the quota yield — see this interface's own doc comment. */
  yielded?: boolean;
  consecutiveFailures: number;
  /** The delay before the next attempt — how stale lastAttemptAt may legitimately get. */
  backoffMs: number;
}

export const DEFAULT_POLL_MS = 30_000;
/** A chat that answers 403 (not a member) is re-tried this often instead of every cycle. */
export const PARK_FORBIDDEN_CHAT_MS = 10 * 60_000;
export const DEFAULT_MAX_BACKOFF_MS = 600_000;
/** Twice was already enough to be worth fixing (0.4.1's live diagnosis); one more than that
 *  guards against firing on an ordinary transient blip that would have cleared on its own. */
export const DEFAULT_AUTH_FAILURE_THRESHOLD = 3;
export const DEFAULT_HEALTH_FILENAME = 'poller-health.json';
/** Default per-chat warm-up back-off window (issue a, live 2026-09-08) — 15 minutes, comfortably
 *  past the 4-minute gap the live incident's two throttled attempts actually showed. Overridden
 *  by `warmupBackoffMs` / `TEAMS_INBOX_WARMUP_BACKOFF_SECONDS`; a throttled attempt that names its
 *  own Retry-After uses that instead — see InboxPollerDeps.warmupBackoffMs's own doc comment. */
export const DEFAULT_WARMUP_BACKOFF_MS = 15 * 60_000;

/**
 * Pathological-size guard only — 64 KiB, not a "keep the line short" budget. Before 0.6.2 this
 * was a silent `text.slice(0, 2000)` present since the initial public release, with no rationale
 * ever recorded; a live incident (2026-09-09, one deployment: the inbox record of a 2,847-char
 * message was exactly 2,000 characters while the read tool returned the full text, content
 * withheld — customer material) showed a downstream reader has no way to tell a cut record from a
 * complete one. Measured in UTF-8 BYTES, not JS string length — the purpose is a disk/memory
 * bound, and a CJK/emoji-heavy message can carry many fewer than 65,536 UTF-16 code units while
 * still exceeding 65,536 bytes on the wire. Every message strictly UNDER this many bytes is
 * written verbatim; a message AT or over it gets an explicit
 * `…[truncated, N chars / M bytes total]` suffix and `truncated: true` alongside `text`, never a
 * silent cut. See KNOWN-ISSUES.md for the incident and root cause.
 */
export const MAX_INBOX_TEXT_BYTES = 64 * 1024;

/**
 * Truncates `text` to at most `maxBytes` UTF-8 bytes without splitting a multi-byte code point
 * (which would otherwise corrupt the last character into a replacement glyph). Iterates by code
 * point (not UTF-16 code unit) so a surrogate pair is never split either.
 */
function truncateToUtf8Bytes(text: string, maxBytes: number): string {
  let bytes = 0;
  let result = '';
  for (const codePoint of text) {
    const codePointBytes = Buffer.byteLength(codePoint, 'utf8');
    if (bytes + codePointBytes > maxBytes) {
      break;
    }
    bytes += codePointBytes;
    result += codePoint;
  }
  return result;
}

/** Builds the `text`/`truncated` fields of one inbox record — see MAX_INBOX_TEXT_BYTES. */
function inboxTextFields(text: string): { text: string; truncated?: true } {
  const byteLength = Buffer.byteLength(text, 'utf8');
  if (byteLength < MAX_INBOX_TEXT_BYTES) {
    return { text };
  }
  const marker = `…[truncated, ${text.length} chars / ${byteLength} bytes total]`;
  return { text: truncateToUtf8Bytes(text, MAX_INBOX_TEXT_BYTES) + marker, truncated: true };
}
/**
 * Consecutive whole-poll failures at which a standing outage is re-surfaced in the inbox. The
 * forever-dedupe wrote ONE {error} line for a five-hour outage — correct against flooding,
 * useless for a watcher asking "is this still going on?". Crossing a threshold changes the
 * dedupe signature, so the same failure gets a fresh line with a running count; recovery past
 * the first threshold is announced too. Bounded on purpose: an outage writes at most
 * 1 + thresholds lines, however long it runs.
 */
export const FAILURE_ESCALATION_THRESHOLDS = [10, 50] as const;
/** A Retry-After naming more than this is treated as this — same sanity cap as the Graph gate
 *  (RETRY_AFTER_CAP_MS in graph-client.ts). */
const RETRY_AFTER_FLOOR_CAP_MS = 60 * 60_000;

export class InboxPoller {
  private readonly pollMs: number;
  private readonly maxBackoffMs: number;
  private readonly healthPath: string;
  private readonly pid: number;
  private readonly log: (line: string) => void;
  private state: Record<string, ChatState> | undefined;
  private me: SignedInAccount | undefined;
  private timer: NodeJS.Timeout | undefined;
  private backoffMs: number;
  /** The delay start() waits before the next pollOnce — the normal interval, the doubled
   *  backoff, or a Retry-After floor; see computeNextDelay. Also what writeHealth reports as
   *  backoffMs, since that IS the number a watcher needs to judge staleness by. */
  private nextDelayMs: number;
  /**
   * When a failed cycle carried a Retry-After (the twin-daemon race carried over from PR #8: two
   * servers on one shared account, Graph naming waits the blind doubling ignored), the next delay
   * is floored at what Graph asked for instead of coming back sooner and feeding the penalty
   * window.
   */
  private retryAfterFloorMs = 0;
  /** Consecutive whole-poll failures — drives both the escalating error lines and the health
   *  file's own consecutiveFailures. Distinct from authFailureStreak below: this counts every
   *  failure shape, that one gates the auth-stuck remedy specifically. */
  private consecutiveFailures = 0;
  private lastSuccessAt: string | undefined;
  private lastErrorSignature: string | undefined;
  private readonly parked = new Map<string, number>();
  private readonly now: () => number;
  private readonly onAuthStuck: (() => void) | undefined;
  private readonly authFailureThreshold: number;
  private authFailureStreak = 0;
  /** True once onAuthStuck has fired for the CURRENT failing streak — prevents firing again on
   *  every subsequent poll while still failing; see trackAuthHealth's doc comment. */
  private authRemedyFired = false;
  /** The `until` of the yield last logged, so one yield episode logs once, not once per cycle. */
  private yieldLoggedUntil: number | undefined;
  /** Chat ids whose warm-up has been ATTEMPTED WITHOUT BEING THROTTLED — behaviour 4 (0.6.0, live
   *  2026-09-08): once a chat lands here there is nothing left to warm, ever, for this process's
   *  lifetime. This is NOT the same claim as "succeeded": GraphTeamsChats.warmMembers's own
   *  contract (unchanged from 0.6.0) reports `{ throttled: false }` alike for a real success
   *  (fetched, or the cache was already warm) and for a non-throttle failure (network error,
   *  licence problem, a misbehaving port throwing, caught defensively by the `.catch` below) —
   *  none of those say trying again right now would help, which is why both land here the same
   *  way. A THROTTLED attempt does NOT land here — it goes into warmBackoffUntil below instead, so
   *  a later cycle can retry once the window elapses rather than being permanently skipped (0.6.3,
   *  issue a). */
  private readonly warmComplete = new Set<string>();
  /** Chat id -> the clock time (this.now()) before which this chat's warm-up is skipped, set
   *  after a throttled attempt (issue a, live 2026-09-08: the SAME chat's warm-up was throttled on
   *  two consecutive poll cycles 4 minutes apart). See warmupBackoffFor for how the window is
   *  computed and InboxPollerDeps.warmupBackoffMs's doc comment for the knobs. */
  private readonly warmBackoffUntil = new Map<string, number>();
  private readonly warmupBackoffMs: number;
  private readonly writeFileFn: typeof writeFile;
  private readonly renameFn: typeof rename;

  constructor(private readonly deps: InboxPollerDeps) {
    this.pollMs = deps.pollMs ?? DEFAULT_POLL_MS;
    this.maxBackoffMs = deps.maxBackoffMs ?? DEFAULT_MAX_BACKOFF_MS;
    this.healthPath = deps.healthPath ?? join(dirname(deps.inboxPath), DEFAULT_HEALTH_FILENAME);
    this.pid = deps.pid ?? process.pid;
    this.log = deps.log ?? (() => {});
    this.backoffMs = this.pollMs;
    this.nextDelayMs = this.pollMs;
    this.now = deps.nowFn ?? (() => Date.now());
    this.onAuthStuck = deps.onAuthStuck;
    this.authFailureThreshold = deps.authFailureThreshold ?? DEFAULT_AUTH_FAILURE_THRESHOLD;
    this.writeFileFn = deps.writeFileFn ?? writeFile;
    this.renameFn = deps.renameFn ?? rename;
    this.warmupBackoffMs = deps.warmupBackoffMs ?? DEFAULT_WARMUP_BACKOFF_MS;
  }

  /**
   * How long THIS throttled warm-up attempt backs its chat off for. MAJOR 1 (review round 1,
   * fresh-context re-review of this branch): Graph's own Retry-After FLOORS the configured/default
   * window, it does not REPLACE it — a short named wait (e.g. GraphClient's LocallyThrottled gate
   * forwards `retryAfterSeconds = ceil(remaining/1000)`, which can be a handful of seconds) used
   * to re-open the chat on the very next poll cycle, reproduced live as 5 warm-up attempts in 5
   * consecutive 30s cycles. `noteRetryAfter` (below) takes the same `max` posture for the whole
   * poll's own next delay — this mirrors it. Still capped at RETRY_AFTER_FLOOR_CAP_MS so a
   * pathological header value can't wedge a chat cold for a day.
   */
  private warmupBackoffFor(retryAfterSeconds: number | undefined): number {
    const named =
      typeof retryAfterSeconds === 'number' && Number.isFinite(retryAfterSeconds) && retryAfterSeconds > 0
        ? retryAfterSeconds * 1000
        : 0;
    return Math.min(Math.max(named, this.warmupBackoffMs), RETRY_AFTER_FLOOR_CAP_MS);
  }

  /**
   * The timers are unref'd: the poller must never be the thing keeping the process alive after
   * the MCP client hangs up, or every session leaks a server. The wait between polls is whatever
   * pollOnce decided (normal interval, doubled backoff, or a Retry-After floor — see
   * computeNextDelay).
   */
  start(): void {
    if (this.timer) {
      return;
    }
    const run = (): void => {
      void this.pollOnce().then(() => {
        this.timer = setTimeout(run, this.nextDelayMs);
        this.timer.unref();
      });
    };
    this.timer = setTimeout(run, 0);
    this.timer.unref();
  }

  stop(): void {
    if (this.timer) {
      clearTimeout(this.timer);
      this.timer = undefined;
    }
  }

  /**
   * One poll over every allowlisted chat, followed by the bookkeeping every outcome shares:
   * failure/recovery lines, the auth-stuck check, the next delay, and the health snapshot —
   * written after EVERY poll, clean, failed, or yielded, because a health file that freezes on
   * failure is exactly the ambiguous silence it exists to remove. Never throws — a poller that
   * can take down the MCP server is worse than no poller.
   *
   * Returns false only when the whole poll failed (auth death, network gone), which is what
   * start() backs off on. A single failing chat — typically one the account is not a member of
   * yet — must not slow down the healthy ones, so it is surfaced but does not count. A yielded
   * cycle (see writeHealth) returns true — being polite is not failing — but is reported to the
   * health file as `yielded`, not as a fresh success.
   */
  async pollOnce(): Promise<boolean> {
    // Quota yield: an attachment download (or any ad-hoc reader) holding the yield file gets
    // the mailbox's Graph budget to itself — this poller polling on regardless is exactly what
    // starved such readers for 20+ minutes on 2026-09-02 (continuous 429, retry-after 62, on
    // every attempt while the poller ran). A skipped cycle counts as CLEAN: the poller is
    // being polite, not failing, so the backoff must not double. Checked once per cycle; the
    // yield's own deadline (capped in readYield) bounds how long a crashed reader can silence
    // the inbox.
    if (this.deps.yieldPath) {
      const standing = await readYield(this.deps.yieldPath, this.now());
      if (standing) {
        if (this.yieldLoggedUntil !== standing.until) {
          this.yieldLoggedUntil = standing.until;
          const remainS = Math.ceil((standing.until - this.now()) / 1000);
          this.log(
            `inbox poller: yielding the Graph read quota to ${standing.reason} ` +
              `(pid ${standing.pid}, up to ${remainS}s more)`,
          );
        }
        await this.writeHealth({ ok: true, yielded: true });
        return true;
      }
      this.yieldLoggedUntil = undefined;
    }

    let clean: boolean;
    let failureMessage: string | undefined;
    let authShapedFailure = false;
    let selfIdThrottled = false;
    try {
      const cycle = await this.pollAllowlistedChats();
      clean = cycle.clean;
      failureMessage = cycle.failureMessage;
      authShapedFailure = cycle.authShapedFailure;
      selfIdThrottled = cycle.selfIdThrottled ?? false;
    } catch (error) {
      this.noteRetryAfter(error);
      failureMessage = error instanceof Error ? error.message : String(error);
      authShapedFailure = isAuthShaped(error);
      clean = false;
    }

    if (clean) {
      if (this.consecutiveFailures >= FAILURE_ESCALATION_THRESHOLDS[0]) {
        await this.reportRecovery(this.consecutiveFailures);
      }
      this.consecutiveFailures = 0;
      this.lastSuccessAt = new Date(this.now()).toISOString();
    } else {
      this.consecutiveFailures += 1;
    }

    if (failureMessage !== undefined) {
      await this.reportFailure(failureMessage);
    } else {
      this.lastErrorSignature = undefined;
    }

    this.trackAuthHealth(clean, authShapedFailure, selfIdThrottled);
    this.nextDelayMs = this.computeNextDelay(clean);
    await this.writeHealth({ ok: clean });
    return clean;
  }

  /**
   * The actual per-chat work: bootstrap, watermark tracking, roster harvest, park/throttle
   * handling. Split out of pollOnce so the shared bookkeeping (escalation, health, auth
   * tracking, backoff) runs exactly once regardless of whether this throws or returns.
   */
  private async pollAllowlistedChats(): Promise<{
    clean: boolean;
    failureMessage?: string;
    authShapedFailure: boolean;
    /** True only when the self-id resolution below was the SOLE, THROTTLED reason this cycle
     *  failed — see trackAuthHealth's own doc comment for why that is exempted from the auth-stuck
     *  streak (issue #28, live 2026-09-09: a tenant-wide `/me` 429 with no operator seed and no
     *  persisted cache used to fail every poll as an "unrecognised shape" failure and, after three
     *  cycles, force a token re-authentication that could never clear a Graph rate limit). */
    selfIdThrottled?: boolean;
  }> {
    await mkdir(dirname(this.deps.inboxPath), { recursive: true });
    this.state ??= await this.loadState();

    const failures: string[] = [];
    const lines: string[] = [];
    let throttled = false;
    let authShapedFailure = false;

    // Without knowing who "self" is, delivered messages could include the assistant's own posts,
    // so an unresolved self id fails the whole poll rather than delivering wrongly. `deps.self`
    // (build-inbox-poller.ts) resolves through the SAME operator-seed -> persisted-cache -> live
    // `/me` chain GraphTeamsChats.resolveSelfId uses for sendFile (issue #28) — a THROTTLED live
    // call is reported below exactly like any other 429 (markThrottled/noteRetryAfter), never
    // folded into the generic "unrecognised shape" failure the last-resort auth-stuck tier would
    // otherwise eventually act on.
    if (this.me === undefined) {
      try {
        this.me = await this.deps.self();
      } catch (error) {
        const message = `self id resolution: ${error instanceof Error ? error.message : String(error)}`;
        const status = (error as { status?: number }).status;
        if (status === 429) {
          throttled = markThrottled(failures, message);
          this.noteRetryAfter(error);
        } else {
          failures.push(message);
        }
        if (isAuthShaped(error)) {
          authShapedFailure = true;
        }
      }
      // Violating-double guard: a `self` implementation that RESOLVES (rather than throws) with
      // no id at all must not be trusted any more than one that threw — it is never memoized, so
      // the next poll retries it, and this cycle still delivers nothing (see the "with no known
      // self id, nothing is delivered" rule below).
      if (this.me?.id === undefined) {
        this.me = undefined;
      }
    }

    if (this.me === undefined) {
      if (failures.length === 0) {
        failures.push('self id resolution: no id reported');
      }
      return {
        clean: false,
        failureMessage: failures.join(' | '),
        authShapedFailure,
        selfIdThrottled: throttled,
      };
    }

    let attempted = 0;
    for (const entry of this.deps.allowlist.entries()) {
      if (throttled) {
        break; // one 429 ends the cycle — every further request would feed the penalty window
      }
      const parkedUntil = this.parked.get(entry.id) ?? 0;
      if (parkedUntil > this.now()) {
        continue; // a chat this account can't read (403) is re-tried on a slow cadence, not every cycle
      }
      attempted += 1;
      // Behaviour 4 (0.6.0, live 2026-09-08), per-chat back-off since 0.6.3 (issue a): warm this
      // chat's roster cache — best-effort, GraphTeamsChats.warmMembers itself never throws, but a
      // `.catch` here is cheap insurance against a differently-behaved TeamsChatsPort
      // implementation (a test double, a future decorator) doing so instead, matching this
      // poller's own "never take the server down" contract (class doc comment above).
      // `readRetries: 0` inside warmMembers means this never sleeps through a Retry-After itself
      // (review round 1 MAJOR 4) — a `throttled: true` result (this single, non-retried attempt
      // WAS throttled) still ends the cycle here, same "one 429 ends the cycle" rule
      // readMessages's own 429 handling follows below: skip THIS chat's read and stop before any
      // LATER chat's warm-up or read, rather than each cold chat in the same cycle issuing its own
      // live 429 against a gate the first one already closed. A chat already attempted WITHOUT
      // being throttled (warmed successfully, cache already warm, or a non-throttle failure — see
      // warmComplete's own doc comment) is skipped forever, unchanged from 0.6.0; a chat still
      // inside its back-off window from an EARLIER throttled attempt is skipped for
      // THIS cycle only, falling straight through to the ordinary readMessages call below — live-
      // diagnosed 2026-09-08: the SAME chat's warm-up was throttled on two consecutive poll cycles
      // just 4 minutes apart, which the old "attempted once, ever" gate (marking a chat as done the
      // moment warmMembers was first CALLED, whatever the outcome) did not actually prevent within
      // one process lifetime.
      if (this.deps.chats.warmMembers && !this.warmComplete.has(entry.id)) {
        const backoffUntil = this.warmBackoffUntil.get(entry.id) ?? 0;
        if (backoffUntil <= this.now()) {
          const warmResult = await this.deps.chats
            .warmMembers(entry.id)
            .catch((): WarmMembersResult => ({ throttled: false }));
          if (warmResult.throttled) {
            throttled = markThrottled(failures, `${entry.label}: THROTTLED (member roster warm-up)`);
            this.warmBackoffUntil.set(
              entry.id,
              this.now() + this.warmupBackoffFor(warmResult.retryAfterSeconds),
            );
            // MINOR (review round 1): a warm-up 429 is still a 429 against the SAME shared
            // /members gate readMessages's own 429 handling floors the whole poll's next delay
            // for (see noteRetryAfter's doc comment) — Graph's named wait should govern the next
            // POLL just as much as it governs this one chat's own back-off window above.
            this.noteRetryAfter(warmResult);
            continue; // the `if (throttled) break;` at the top of the next iteration ends the cycle
          }
          // Memory hygiene only: warmComplete below is checked FIRST on every later iteration for
          // this chat, so a stale warmBackoffUntil entry left behind here is never read again —
          // this delete has no effect on behaviour, it just stops the map growing unboundedly for
          // a chat that eventually succeeds after several throttled attempts.
          this.warmBackoffUntil.delete(entry.id);
          // "COMPLETE" here means "attempted without being throttled", not "succeeded" — the
          // underlying GraphTeamsChats.warmMembers contract (and, defensively, the `.catch` above
          // for a differently-behaved port) already reports `{ throttled: false }` for a REAL
          // success (fetched, or cache already warm) and for a non-throttle failure alike (network
          // error, licence problem, a misbehaving port throwing) — unchanged from 0.6.0. Telling
          // those two apart would be a real behaviour change (a network-error chat would need its
          // own retry story, not "never try again"), which is out of scope here.
          this.warmComplete.add(entry.id);
        }
        // else: still inside this chat's back-off window from an earlier throttled attempt — skip
        // the warm-up silently this cycle and fall through to the ordinary readMessages path
        // below; the window elapsing is what makes a LATER cycle attempt it again.
      }
      const known = this.state[entry.id];
      // 0.4.1 (live-diagnosed: a restart was observed replaying ~40 old messages): no entry on
      // record for this chat means "history unknown to THIS process" — not "chat is empty" —
      // whether that is because the sidecar was lost, corrupted, or this is a genuinely fresh
      // install. Backfilling whatever already existed at that moment used to be the "safe"
      // fallback; it is what actually produced the backfill. The fix: a chat with no known
      // watermark gets exactly one settling poll that ESTABLISHES the watermark and delivers
      // nothing — every poll after that behaves exactly as before.
      const isBootstrap = known === undefined;
      try {
        const result = await this.deps.chats.readMessages(entry.id, known?.watermark);
        // Mitigation 2 (docs/throttling-mitigation.md §4, stage 1 item 2): every message this
        // poll already fetched carries its sender's AAD id and display name at zero MARGINAL
        // Graph cost, whether or not that particular message goes on to be DELIVERED as an
        // inbox event below — a bootstrap/settling poll's messages are never delivered (0.4.1
        // contract, unchanged) but their senders are still real chat members worth learning.
        // Harvested BEFORE the isBootstrap/isSelf/isDeleted filters below, which govern inbox
        // delivery only, not roster membership.
        if (this.deps.roster) {
          const harvested = result.messages.flatMap((message) => this.harvestable(message));
          if (harvested.length > 0) {
            this.deps.roster.merge(entry.id, harvested);
          }
        }
        for (const message of result.messages) {
          if (isBootstrap) {
            continue;
          }
          if (message.id === known?.lastId) {
            continue;
          }
          if (this.isSelf(message.fromId, message.from)) {
            continue;
          }
          if (message.isDeleted || (!message.text.trim() && message.attachments.length === 0)) {
            continue;
          }
          lines.push(
            JSON.stringify({
              chat: entry.id,
              id: message.id,
              from: message.from,
              at: message.createdDateTime,
              ...inboxTextFields(message.text),
              attachments: message.attachments.length,
            }),
          );
        }
        const newest = result.messages.at(-1);
        // Always record an entry — even an empty one — once a chat has been successfully
        // polled: that is what ends bootstrap mode for it. A cycle with no messages must not
        // erase a watermark/lastId a PRIOR cycle already established.
        const watermark = result.watermark ?? known?.watermark;
        const lastId = newest?.id ?? known?.lastId;
        this.state[entry.id] = {
          ...(watermark ? { watermark } : {}),
          ...(lastId ? { lastId } : {}),
        };
      } catch (error) {
        // One unreachable chat must not blank out the poll for the others.
        const message = `${entry.label}: ` + (error instanceof Error ? error.message : String(error));
        const status = (error as { status?: number }).status;
        if (status === 429) {
          // 2026-08-25: polling on through a throttle kept an account throttled for hours.
          // markThrottled (issue c, 2026-09-08 review) keeps the failures entry and the throttled
          // flag paired — see that function's own doc comment for why this used to be convention
          // only.
          throttled = markThrottled(failures, message);
          this.noteRetryAfter(error);
        } else {
          failures.push(message);
          if (status === 403) {
            this.parked.set(entry.id, this.now() + PARK_FORBIDDEN_CHAT_MS);
          }
        }
        if (isAuthShaped(error)) {
          authShapedFailure = true;
        }
      }
    }

    if (lines.length > 0) {
      await appendFile(this.deps.inboxPath, lines.map((line) => `${line}\n`).join(''));
    }
    await this.saveState();

    if (failures.length === 0) {
      return { clean: true, authShapedFailure: false };
    }
    // A throttled cycle is never "clean": start() doubles the wait, which is the only
    // thing that ends a throttle. A single dead chat still doesn't slow the healthy ones —
    // but the denominator is the chats this cycle actually ASKED: with the 403 chats
    // parked, "one of three failed" must not hide "the only chat asked failed" (auth death).
    return {
      clean: !throttled && failures.length < attempted,
      failureMessage: failures.join(' | '),
      authShapedFailure,
    };
  }

  /** Remembers a 429's Retry-After (when the Graph layer named one) as a floor for the next
   *  delay. Carried over from PR #8: when two daemons raced on one shared account, Graph named
   *  waits the blind doubling ignored, and the poller came back sooner than asked. */
  private noteRetryAfter(error: unknown): void {
    const seconds = (error as { retryAfterSeconds?: number }).retryAfterSeconds;
    if (typeof seconds === 'number' && Number.isFinite(seconds) && seconds > 0) {
      this.retryAfterFloorMs = Math.max(
        this.retryAfterFloorMs,
        Math.min(seconds * 1000, RETRY_AFTER_FLOOR_CAP_MS),
      );
    }
  }

  /**
   * Same backoff the standalone daemon used — double the wait while failing, capped, and snap
   * back to the normal interval on the first clean poll — except that a Retry-After named by a
   * 429 floors the next delay: Graph's asked-for wait beats our own guess, even past the
   * doubling cap, because coming back early is what escalates a penalty window.
   */
  private computeNextDelay(clean: boolean): number {
    if (clean) {
      this.backoffMs = this.pollMs;
      this.retryAfterFloorMs = 0;
      return this.pollMs;
    }
    const delay = Math.max(this.backoffMs, this.retryAfterFloorMs);
    this.backoffMs = Math.min(delay * 2, this.maxBackoffMs);
    this.retryAfterFloorMs = 0;
    return delay;
  }

  /**
   * Live-diagnosed stuck-auth mode (0.4.1): repeated poll failures while Graph itself was
   * reachable, recovered only by a process restart. Review round 1 supplied the incident's own
   * log evidence — `inbox poll failed: <chat>: fetch failed`, repeated, NO status code, Graph
   * answering 200 to parallel curl probes the whole time (recorded verbatim in KNOWN-ISSUES.md).
   * That shape matches nothing isAuthShaped recognises, so a detector gated on auth vocabulary
   * would never have fired on the actual incident.
   *
   * Two tiers, same threshold (N=3, DEFAULT_AUTH_FAILURE_THRESHOLD): auth-shaped failures are the
   * fast, well-understood path (RopcTokenProvider trusting a dead cached token purely by local
   * clock is a plausible mechanism, though unconfirmed); ANY OTHER consecutive failure shape is
   * the LAST-RESORT tier — a spurious forced re-mint during a genuine network outage costs one
   * extra password grant on the next successful call and nothing else, which is cheap next to
   * staying stuck until a human restarts the process.
   *
   * The remedy fires ONCE per failing streak, not once per poll: firing only FLAGS the cached
   * token (invalidate() is synchronous, no network call) — the actual re-mint attempt happens on
   * the next getAccessToken() call, whose outcome shows up as the NEXT poll's own clean/failing
   * result. Resetting the streak the instant onAuthStuck is called would therefore claim a
   * recovery the code cannot know yet; the streak (and authRemedyFired) resets ONLY when a
   * subsequent poll actually comes back clean.
   *
   * `selfIdThrottled` (0.6.4, issue #28, live 2026-09-09): a cycle whose ONLY failure was a
   * THROTTLED self-id resolution (no operator seed, no persisted cache, a live `/me` 429) is
   * exempted from this streak entirely — neither progressed nor reset. Graph rate-limiting `/me`
   * is a known, non-auth condition; forcing a token re-mint would not clear it, so counting it
   * toward the last-resort tier only spends a password grant for nothing while the actual
   * unrelated throttle keeps running its own course (markThrottled/noteRetryAfter, tracked
   * separately by pollAllowlistedChats). A poll that ALSO fails for some other reason still
   * progresses the streak normally — this exemption applies only when the self-id throttle was
   * the cycle's sole failure.
   */
  private trackAuthHealth(clean: boolean, authShaped: boolean, selfIdThrottled: boolean): void {
    if (clean) {
      if (this.authRemedyFired) {
        this.log('inbox poller: recovered after a forced token re-authentication');
      }
      this.authFailureStreak = 0;
      this.authRemedyFired = false;
      return;
    }
    if (selfIdThrottled) {
      return;
    }
    this.authFailureStreak += 1;
    if (this.authFailureStreak < this.authFailureThreshold) {
      return;
    }
    if (this.authRemedyFired) {
      this.log('inbox poller: still failing after a forced token re-authentication');
      return;
    }
    this.authRemedyFired = true;
    const shape = authShaped
      ? 'consecutive auth-shaped poll failures'
      : 'consecutive poll failures of unrecognised shape (last-resort: Graph is not clearing on its own)';
    this.log(`inbox poller: ${this.authFailureThreshold} ${shape} — requested a forced token re-authentication`);
    this.onAuthStuck?.();
  }

  private isSelf(fromId: string | undefined, from: string): boolean {
    if (this.me?.id !== undefined && fromId === this.me.id) {
      return true;
    }
    return this.me?.displayName !== undefined && from === this.me.displayName;
  }

  /**
   * Violating-double guard (mitigation 2): `message.from` degrades to the literal string
   * 'unknown' (or 'system' for a system event) when toChatMessage (messages.ts) could not map a
   * real sender — a shape Graph itself can produce (an id with no displayName reported). Neither
   * sentinel is a real person's name, and merging one into the roster would poison mention
   * resolution with junk. Real display names are never expected to collide with these two exact
   * strings, so a straight equality check is enough — no heuristics needed. The self account IS
   * harvested (it is a real chat member); only isSelf-DELIVERY to the inbox skips it, above.
   */
  private harvestable(message: ChatMessage): Array<{ id: string; displayName: string }> {
    return message.fromId && message.from !== 'unknown' && message.from !== 'system'
      ? [{ id: message.fromId, displayName: message.from }]
      : [];
  }

  private async loadState(): Promise<Record<string, ChatState>> {
    try {
      const parsed: unknown = JSON.parse(await readFile(this.deps.statePath, 'utf8'));
      if (typeof parsed === 'object' && parsed !== null && !Array.isArray(parsed)) {
        return parsed as Record<string, ChatState>;
      }
    } catch {
      // Missing or corrupt state is a fresh install as far as this process is concerned: every
      // chat starts in bootstrap mode (see pollOnce) and the next cycle settles from NOW, never
      // replaying old history — see the 0.4.1 comment on isBootstrap for why "safe" used to mean
      // the opposite of that.
    }
    return {};
  }

  /** Write-to-temp-then-rename: a crash mid-write must never leave a half-written sidecar that
   *  the next startup's loadState() reads as corrupt and treats as a fresh install. */
  private async saveState(): Promise<void> {
    const path = this.deps.statePath;
    await mkdir(dirname(path), { recursive: true });
    const tmpPath = `${path}.tmp-${process.pid}-${this.now()}`;
    await this.writeFileFn(tmpPath, JSON.stringify(this.state ?? {}, null, 2));
    await this.renameFn(tmpPath, path);
  }

  /**
   * The liveness proof, rewritten after every poll — including a yielded one, so the age of
   * this file alone tells a watcher the pipeline is still ticking (or is not). Atomic (tmp +
   * rename) so a watcher never reads a torn snapshot, and a write failure never throws past
   * pollOnce — the same discipline as reportFailure: supervision must not be able to kill the
   * thing it supervises.
   */
  private async writeHealth(state: { ok: boolean; yielded?: boolean }): Promise<void> {
    const snapshot: PollerHealth = {
      pid: this.pid,
      inboxPath: this.deps.inboxPath,
      lastAttemptAt: new Date(this.now()).toISOString(),
      ...(this.lastSuccessAt !== undefined ? { lastSuccessAt: this.lastSuccessAt } : {}),
      ok: state.ok,
      ...(state.yielded ? { yielded: true as const } : {}),
      consecutiveFailures: this.consecutiveFailures,
      backoffMs: this.nextDelayMs,
    };
    const tmpPath = `${this.healthPath}.tmp-${this.pid}`;
    try {
      await writeFile(tmpPath, JSON.stringify(snapshot, null, 2));
      await rename(tmpPath, this.healthPath);
    } catch (writeError) {
      this.log(`inbox poller cannot write ${this.healthPath}: ${String(writeError)}`);
    }
  }

  /**
   * Silence must not be ambiguous: a watcher on the inbox file has to see auth death too. But a
   * chat that fails identically poll after poll (an allowlisted chat the account has not been
   * added to yet) would flood the file, so the same failure is only written once — until it
   * changes, clears, or crosses an escalation threshold (see FAILURE_ESCALATION_THRESHOLDS),
   * which re-surfaces a standing outage with a running count instead of one line for five hours.
   */
  private async reportFailure(message: string): Promise<void> {
    this.log(`inbox poll failed: ${message}`);
    const bucket = FAILURE_ESCALATION_THRESHOLDS.filter(
      (threshold) => this.consecutiveFailures >= threshold,
    ).length;
    const signature = `${bucket}:${message}`;
    if (signature === this.lastErrorSignature) {
      return;
    }
    this.lastErrorSignature = signature;
    const text =
      bucket > 0 ? `still failing after ${this.consecutiveFailures} polls: ${message}` : message;
    const line = JSON.stringify({
      error: text,
      at: new Date(this.now()).toISOString(),
      consecutiveFailures: this.consecutiveFailures,
    });
    await this.appendToInbox(line);
  }

  /**
   * The other half of escalation: after an outage long enough to have been re-surfaced, the
   * watcher must also see the pipeline come BACK — otherwise the last thing in the file is
   * "still failing" forever, and quiet chats become indistinguishable from a dead recovery.
   * Short blips stay quiet; only outages past the first threshold get the line.
   */
  private async reportRecovery(afterFailures: number): Promise<void> {
    this.log(`inbox poll recovered after ${afterFailures} failed polls`);
    const line = JSON.stringify({
      recovered: true,
      at: new Date(this.now()).toISOString(),
      afterFailures,
    });
    await this.appendToInbox(line);
  }

  private async appendToInbox(line: string): Promise<void> {
    try {
      await appendFile(this.deps.inboxPath, `${line}\n`);
    } catch (writeError) {
      // Even the error/recovery line failing must not throw past pollOnce.
      this.log(`inbox poller cannot write ${this.deps.inboxPath}: ${String(writeError)}`);
    }
  }
}
