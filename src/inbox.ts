import { appendFile, mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import type { ChatAllowlist } from './allowlist.js';
import type { TeamsChatsPort } from './graph/teams-chats.js';

/**
 * Background inbox poller. Runs inside the MCP server process and appends every new message from
 * the allowlisted chats to a JSONL file at a stable path, so an orchestrator can arm a file
 * watcher once and stop reinventing a polling daemon per session.
 *
 * One line per message: {"chat","id","from","at","text","attachments"} — the shape the old
 * scratchpad daemon wrote, kept so existing consumers keep parsing. A failed poll appends
 * {"error","at","consecutiveFailures"} instead, because a watcher must be able to tell auth
 * death from a quiet chat.
 *
 * Since 0.5.0 the poller also proves it is alive: after EVERY poll it rewrites a small health
 * file beside the inbox. That exists because of 2026-09-01, when a host reboot killed one
 * project's daemon while a second project's daemon — same dist/index.js path, different env —
 * survived, and a pgrep-based check matched the survivor: the dead pipeline read as "up" and an
 * allowlisted chat sat undelivered for 1.5 hours. The health file's AGE is the liveness signal:
 * it answers "is THIS inbox's pipeline polling", independent of what other processes look like.
 */

export interface SignedInAccount {
  id?: string;
  displayName?: string;
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
  /** Where the health snapshot lands; defaults to poller-health.json beside the inbox. */
  healthPath?: string;
  /** The pid recorded in the health file; defaults to process.pid. Injectable for tests. */
  pid?: number;
  pollMs?: number;
  /** Injectable clock for the 403 park and the health timestamps; defaults to Date.now. */
  nowFn?: () => number;
  maxBackoffMs?: number;
  log?: (line: string) => void;
}

interface ChatState {
  /** ISO-8601 createdDateTime of the newest delivered message; exclusive, like read watermarks. */
  watermark: string;
  /** Id of that newest message, so an exact replay is dropped even if the port re-serves it. */
  lastId?: string;
}

/**
 * What an external watcher reads to judge the pipeline. The contract: lastAttemptAt older than
 * a couple of backoffMs means the poller is dead or wedged, whatever pgrep says; ok=false with a
 * fresh lastAttemptAt means the process is alive but its polls are failing (the {error} lines in
 * the inbox say why).
 */
export interface PollerHealth {
  pid: number;
  inboxPath: string;
  lastAttemptAt: string;
  /** Stamp of the last CLEAN poll; absent until the first one. Survives failures, so a watcher
   *  can say how long an outage has been running. */
  lastSuccessAt?: string;
  ok: boolean;
  consecutiveFailures: number;
  /** The delay before the next attempt — how stale lastAttemptAt may legitimately get. */
  backoffMs: number;
}

export const DEFAULT_POLL_MS = 30_000;
/** A chat that answers 403 (not a member) is re-tried this often instead of every cycle. */
export const PARK_FORBIDDEN_CHAT_MS = 10 * 60_000;
export const DEFAULT_MAX_BACKOFF_MS = 600_000;
export const DEFAULT_HEALTH_FILENAME = 'poller-health.json';
/**
 * Consecutive whole-poll failures at which a standing outage is re-surfaced in the inbox. The
 * forever-dedupe wrote ONE {error} line for a five-hour outage — correct against flooding,
 * useless for a watcher asking "is this still going on?". Crossing a threshold changes the
 * dedupe signature, so the same failure gets a fresh line with a running count; recovery past
 * the first threshold is announced too. Bounded on purpose: an outage writes at most
 * 1 + thresholds lines, however long it runs.
 */
export const FAILURE_ESCALATION_THRESHOLDS = [10, 50] as const;
/** A Retry-After naming more than this is treated as this — same sanity cap as the Graph gate. */
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
  private nextDelayMs: number;
  /**
   * When a failed cycle carried a Retry-After (the twin-daemon race of 2026-09-01: two servers
   * on one shared account, Graph naming waits the blind doubling ignored), the next delay is
   * floored at what Graph asked for instead of coming back sooner and feeding the penalty window.
   */
  private retryAfterFloorMs = 0;
  private consecutiveFailures = 0;
  private lastSuccessAt: string | undefined;
  private lastErrorSignature: string | undefined;
  private readonly parked = new Map<string, number>();
  private readonly now: () => number;

  constructor(private readonly deps: InboxPollerDeps) {
    this.pollMs = deps.pollMs ?? DEFAULT_POLL_MS;
    this.maxBackoffMs = deps.maxBackoffMs ?? DEFAULT_MAX_BACKOFF_MS;
    this.healthPath = deps.healthPath ?? join(dirname(deps.inboxPath), DEFAULT_HEALTH_FILENAME);
    this.pid = deps.pid ?? process.pid;
    this.log = deps.log ?? (() => {});
    this.backoffMs = this.pollMs;
    this.nextDelayMs = this.pollMs;
    this.now = deps.nowFn ?? (() => Date.now());
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
   * failure/recovery lines, the next delay, and the health snapshot — written after EVERY poll,
   * clean or failed, because a health file that freezes on failure is exactly the ambiguous
   * silence it exists to remove. Never throws — a poller that can take down the MCP server is
   * worse than no poller.
   *
   * Returns false only when the whole poll failed (auth death, network gone), which is what
   * start() backs off on. A single failing chat — typically one the account is not a member of
   * yet — must not slow down the healthy ones, so it is surfaced but does not count.
   */
  async pollOnce(): Promise<boolean> {
    let clean: boolean;
    let failureMessage: string | undefined;
    try {
      const cycle = await this.pollAllowlistedChats();
      clean = cycle.clean;
      failureMessage = cycle.failureMessage;
    } catch (error) {
      this.noteRetryAfter(error);
      failureMessage = error instanceof Error ? error.message : String(error);
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

    this.nextDelayMs = this.computeNextDelay(clean);
    await this.writeHealth(clean);
    return clean;
  }

  private async pollAllowlistedChats(): Promise<{ clean: boolean; failureMessage?: string }> {
    await mkdir(dirname(this.deps.inboxPath), { recursive: true });
    this.state ??= await this.loadState();
    // Without knowing who "self" is, delivered messages could include the assistant's own
    // posts, so a failed /me fails the whole poll rather than delivering wrongly.
    this.me ??= await this.deps.self();

    const failures: string[] = [];
    const lines: string[] = [];

    let throttled = false;
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
      const known = this.state[entry.id];
      try {
        const result = await this.deps.chats.readMessages(entry.id, known?.watermark);
        for (const message of result.messages) {
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
        if (result.watermark) {
          this.state[entry.id] = {
            watermark: result.watermark,
            ...(newest?.id ? { lastId: newest.id } : {}),
          };
        }
      } catch (error) {
        // One unreachable chat must not blank out the poll for the others.
        failures.push(
          `${entry.label}: ` + (error instanceof Error ? error.message : String(error)),
        );
        const status = (error as { status?: number }).status;
        if (status === 429) {
          throttled = true; // 2026-08-25: polling on through a throttle kept an account throttled for hours
          this.noteRetryAfter(error);
        } else if (status === 403) {
          this.parked.set(entry.id, this.now() + PARK_FORBIDDEN_CHAT_MS);
        }
      }
    }

    if (lines.length > 0) {
      await appendFile(this.deps.inboxPath, lines.map((line) => `${line}\n`).join(''));
    }
    await this.saveState();

    if (failures.length === 0) {
      return { clean: true };
    }
    // A throttled cycle is never "clean": start() doubles the wait, which is the only
    // thing that ends a throttle. A single dead chat still doesn't slow the healthy ones —
    // but the denominator is the chats this cycle actually ASKED: with the 403 chats
    // parked, "one of three failed" must not hide "the only chat asked failed" (auth death).
    return {
      clean: !throttled && failures.length < attempted,
      failureMessage: failures.join(' | '),
    };
  }

  /** Remembers a 429's Retry-After (when the Graph layer named one) as a floor for the next delay. */
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

  private isSelf(fromId: string | undefined, from: string): boolean {
    if (this.me?.id !== undefined && fromId === this.me.id) {
      return true;
    }
    return this.me?.displayName !== undefined && from === this.me.displayName;
  }

  private async loadState(): Promise<Record<string, ChatState>> {
    try {
      const parsed: unknown = JSON.parse(await readFile(this.deps.statePath, 'utf8'));
      if (typeof parsed === 'object' && parsed !== null && !Array.isArray(parsed)) {
        return parsed as Record<string, ChatState>;
      }
    } catch {
      // Missing or corrupt state means starting from "now-ish": the first poll re-reads the
      // recent window once. Better a rare duplicate after losing the file than never starting.
    }
    return {};
  }

  private async saveState(): Promise<void> {
    await writeFile(this.deps.statePath, JSON.stringify(this.state ?? {}, null, 2));
  }

  /**
   * The liveness proof, rewritten after every poll. Atomic (tmp + rename) so a watcher never
   * reads a torn snapshot, and a write failure never throws past pollOnce — the same discipline
   * as reportFailure: supervision must not be able to kill the thing it supervises.
   */
  private async writeHealth(ok: boolean): Promise<void> {
    const snapshot: PollerHealth = {
      pid: this.pid,
      inboxPath: this.deps.inboxPath,
      lastAttemptAt: new Date(this.now()).toISOString(),
      ...(this.lastSuccessAt !== undefined ? { lastSuccessAt: this.lastSuccessAt } : {}),
      ok,
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
      // Even the error line failing must not throw past pollOnce.
      this.log(`inbox poller cannot write ${this.deps.inboxPath}: ${String(writeError)}`);
    }
  }
}
