import { appendFile, mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import { dirname } from 'node:path';
import type { ChatAllowlist } from './allowlist.js';
import type { TeamsChatsPort } from './graph/teams-chats.js';

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
  pollMs?: number;
  /** Injectable clock for the 403 park; defaults to Date.now. */
  nowFn?: () => number;
  maxBackoffMs?: number;
  log?: (line: string) => void;
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

  constructor(private readonly deps: InboxPollerDeps) {
    this.pollMs = deps.pollMs ?? DEFAULT_POLL_MS;
    this.maxBackoffMs = deps.maxBackoffMs ?? DEFAULT_MAX_BACKOFF_MS;
    this.log = deps.log ?? (() => {});
    this.backoffMs = this.pollMs;
    this.now = deps.nowFn ?? (() => Date.now());
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
        return !throttled && failures.length < attempted;
      }
      this.lastErrorSignature = undefined;
      return true;
    } catch (error) {
      await this.reportFailure(error instanceof Error ? error.message : String(error));
      return false;
    }
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
    await writeFile(tmpPath, JSON.stringify(this.state ?? {}, null, 2));
    await rename(tmpPath, path);
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
