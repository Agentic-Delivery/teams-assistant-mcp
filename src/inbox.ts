import { appendFile, mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import { dirname } from 'node:path';
import type { ChatAllowlist } from './allowlist.js';
import type { TeamsChatsPort } from './graph/teams-chats.js';
import { readYield } from './inbox-yield.js';
import type { ChatMessage } from './messages.js';

/**
 * Background inbox poller. Runs inside the MCP server process and appends every new message from
 * the allowlisted chats to a JSONL file at a stable path, so an orchestrator can arm a file
 * watcher once and stop reinventing a polling daemon per session.
 *
 * One line per message: {"chat","id","from","at","text","attachments"} — the shape the old
 * scratchpad daemon wrote, kept so existing consumers keep parsing. A failed poll appends
 * {"error","at"} instead, because a watcher must be able to tell auth death from a quiet chat.
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
 * the persisted roster is what lets a chat's roster fill and refresh from traffic alone, taking
 * `/members` off the mention-resolution send path for the common case (see
 * GraphTeamsChats.resolveMentions's stale-serve fallback, teams-chats.ts, for what still happens
 * on the rarer miss).
 */
export interface RosterHarvestPort {
  merge(chatId: string, members: ReadonlyArray<{ id: string; displayName: string }>): void;
}

export interface InboxPollerDeps {
  chats: Pick<TeamsChatsPort, 'readMessages'>;
  allowlist: ChatAllowlist;
  /**
   * Resolves who the server is signed in as. The assistant's own posts must not come back as
   * inbox events, or every send would wake the orchestrator to read its own words.
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
  pollMs?: number;
  /** Injectable clock for the 403 park; defaults to Date.now. */
  nowFn?: () => number;
  maxBackoffMs?: number;
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

export const DEFAULT_POLL_MS = 30_000;
/** A chat that answers 403 (not a member) is re-tried this often instead of every cycle. */
export const PARK_FORBIDDEN_CHAT_MS = 10 * 60_000;
export const DEFAULT_MAX_BACKOFF_MS = 600_000;
/** Twice was already enough to be worth fixing (0.4.1's live diagnosis); one more than that
 *  guards against firing on an ordinary transient blip that would have cleared on its own. */
export const DEFAULT_AUTH_FAILURE_THRESHOLD = 3;

export class InboxPoller {
  private readonly pollMs: number;
  private readonly maxBackoffMs: number;
  private readonly log: (line: string) => void;
  private state: Record<string, ChatState> | undefined;
  private me: SignedInAccount | undefined;
  private timer: NodeJS.Timeout | undefined;
  private backoffMs: number;
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
  private readonly writeFileFn: typeof writeFile;
  private readonly renameFn: typeof rename;

  constructor(private readonly deps: InboxPollerDeps) {
    this.pollMs = deps.pollMs ?? DEFAULT_POLL_MS;
    this.maxBackoffMs = deps.maxBackoffMs ?? DEFAULT_MAX_BACKOFF_MS;
    this.log = deps.log ?? (() => {});
    this.backoffMs = this.pollMs;
    this.now = deps.nowFn ?? (() => Date.now());
    this.onAuthStuck = deps.onAuthStuck;
    this.authFailureThreshold = deps.authFailureThreshold ?? DEFAULT_AUTH_FAILURE_THRESHOLD;
    this.writeFileFn = deps.writeFileFn ?? writeFile;
    this.renameFn = deps.renameFn ?? rename;
  }

  /**
   * The timers are unref'd: the poller must never be the thing keeping the process alive after
   * the MCP client hangs up, or every session leaks a server.
   */
  start(): void {
    if (this.timer) {
      return;
    }
    const run = (): void => {
      void this.pollOnce().then((clean) => {
        // Same backoff the standalone daemon used: double the wait while failing, capped, and
        // snap back to the normal interval on the first clean poll.
        const delay = clean ? this.pollMs : this.backoffMs;
        this.backoffMs = clean ? this.pollMs : Math.min(this.backoffMs * 2, this.maxBackoffMs);
        this.timer = setTimeout(run, delay);
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
   * One poll over every allowlisted chat. Never throws — a poller that can take down the MCP
   * server is worse than no poller.
   *
   * Returns false only when the whole poll failed (auth death, network gone), which is what
   * start() backs off on. A single failing chat — typically one the account is not a member of
   * yet — must not slow down the healthy ones, so it is surfaced but does not count.
   */
  async pollOnce(): Promise<boolean> {
    try {
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
          return true;
        }
        this.yieldLoggedUntil = undefined;
      }
      await mkdir(dirname(this.deps.inboxPath), { recursive: true });
      this.state ??= await this.loadState();
      // Without knowing who "self" is, delivered messages could include the assistant's own
      // posts, so a failed /me fails the whole poll rather than delivering wrongly.
      this.me ??= await this.deps.self();

      const failures: string[] = [];
      const lines: string[] = [];

      let throttled = false;
      let attempted = 0;
      let authShapedFailure = false;
      for (const entry of this.deps.allowlist.entries()) {
        if (throttled) {
          break; // one 429 ends the cycle — every further request would feed the penalty window
        }
        const parkedUntil = this.parked.get(entry.id) ?? 0;
        if (parkedUntil > this.now()) {
          continue; // a chat this account can't read (403) is re-tried on a slow cadence, not every cycle
        }
        attempted += 1;
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
                text: message.text.slice(0, 2000),
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
          failures.push(
            `${entry.label}: ` + (error instanceof Error ? error.message : String(error)),
          );
          const status = (error as { status?: number }).status;
          if (status === 429) {
            throttled = true; // 2026-08-25: polling on through a throttle kept an account throttled for hours
          } else if (status === 403) {
            this.parked.set(entry.id, this.now() + PARK_FORBIDDEN_CHAT_MS);
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

      if (failures.length > 0) {
        await this.reportFailure(failures.join(' | '));
        // A throttled cycle is never "clean": start() doubles the wait, which is the only
        // thing that ends a throttle. A single dead chat still doesn't slow the healthy ones —
        // but the denominator is the chats this cycle actually ASKED: with the 403 chats
        // parked, "one of three failed" must not hide "the only chat asked failed" (auth death).
        const clean = !throttled && failures.length < attempted;
        this.trackAuthHealth(clean, authShapedFailure);
        return clean;
      }
      this.lastErrorSignature = undefined;
      this.trackAuthHealth(true, false);
      return true;
    } catch (error) {
      await this.reportFailure(error instanceof Error ? error.message : String(error));
      this.trackAuthHealth(false, isAuthShaped(error));
      return false;
    }
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
   */
  private trackAuthHealth(clean: boolean, authShaped: boolean): void {
    if (clean) {
      if (this.authRemedyFired) {
        this.log('inbox poller: recovered after a forced token re-authentication');
      }
      this.authFailureStreak = 0;
      this.authRemedyFired = false;
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
   * Silence must not be ambiguous: a watcher on the inbox file has to see auth death too. But a
   * chat that fails identically poll after poll (an allowlisted chat the account has not been
   * added to yet) would flood the file, so the same failure is only written once until it
   * changes or clears.
   */
  private async reportFailure(message: string): Promise<void> {
    this.log(`inbox poll failed: ${message}`);
    if (message === this.lastErrorSignature) {
      return;
    }
    this.lastErrorSignature = message;
    const line = JSON.stringify({ error: message, at: new Date().toISOString() });
    try {
      await appendFile(this.deps.inboxPath, `${line}\n`);
    } catch (writeError) {
      // Even the error line failing must not throw past pollOnce.
      this.log(`inbox poller cannot write ${this.deps.inboxPath}: ${String(writeError)}`);
    }
  }
}
