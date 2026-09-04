import { mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import { dirname } from 'node:path';

/**
 * Single-instance lock for the inbox poller, keyed to the inbox path (poller.lock beside it) —
 * carries the poller-supervision work from PR #8 onto 0.5.2's inboxPathFor(env) (inbox-yield.ts):
 * index.ts derives the lock path from the SAME helper the CLIs and the poller itself use for the
 * inbox path, so a TEAMS_INBOX_PATH override moves the lock with everything else it already moves.
 *
 * Why it exists: two servers for different projects run from the same dist/index.js, and for a
 * while two pollers raced on one shared account — every interval carried double the Graph
 * traffic, and under throttling each daemon's polls fed the other's penalty window. One poller
 * per inbox has to be a checked invariant, not an assumption about how many processes happen to
 * be running.
 *
 * The lock is keyed PER INBOX PATH, not per config directory — two projects with distinct
 * TEAMS_INBOX_PATH values correctly keep two independent pollers. The trap this creates: an
 * instance that sets TEAMS_MCP_CONFIG/TEAMS_MCP_TOKEN_CACHE to its own directory but leaves
 * TEAMS_INBOX_PATH unset falls back to the SAME default inbox path as any other instance that
 * also left it unset, and therefore silently contends for the SAME lock — see index.ts's own
 * comment at the acquirePollerLock call site, and the "Supervising the daemon" section of the
 * README, for the operator-facing warning this requires (env.example says it too).
 *
 * The lock is advisory, pid-based, and deliberately simple: a JSON file naming the holder. A
 * holder whose pid is dead (the reboot case) is taken over; a live holder wins, and — since
 * 0.5.4 — the loser logs why and exits non-zero rather than lingering as a tools-only server: a
 * server that silently never polls is exactly the invisible failure mode this lock exists to
 * prevent, and a process manager needs a non-zero exit to notice and alert on. The read-then-
 * write window is not atomic across processes, but the failure mode this guards against is a
 * long-lived daemon meeting a newcomer, not two processes starting in the same millisecond.
 */

export interface PollerLockDeps {
  lockPath: string;
  /** The pid to record as holder; defaults to process.pid. Injectable for tests. */
  pid?: number;
  /**
   * Whether a pid is alive; defaults to signal 0 via process.kill. EPERM counts as alive —
   * it means the process exists but belongs to someone else, which is still a live holder.
   */
  isPidAlive?: (pid: number) => boolean;
  /** Injectable clock for the recorded start time; defaults to Date.now. */
  nowFn?: () => number;
}

export type PollerLockResult =
  | { acquired: true }
  | { acquired: false; holderPid: number; holderStartedAt?: string };

function defaultIsPidAlive(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch (error) {
    return (error as { code?: string }).code === 'EPERM';
  }
}

/**
 * Takes the lock at deps.lockPath, or reports the live holder that already has it. Only a
 * holder file that parses AND names a live pid other than our own blocks acquisition —
 * corruption or a dead pid means the holder is gone, and refusing to take over would deadlock
 * every restart after a crash. The write is tmp + rename so a rival reading mid-write sees the
 * old holder or the new one, never a torn file.
 */
export async function acquirePollerLock(deps: PollerLockDeps): Promise<PollerLockResult> {
  const pid = deps.pid ?? process.pid;
  const isPidAlive = deps.isPidAlive ?? defaultIsPidAlive;
  const now = deps.nowFn ?? (() => Date.now());

  const holder = await readHolder(deps.lockPath);
  if (holder !== undefined && holder.pid !== pid && isPidAlive(holder.pid)) {
    return {
      acquired: false,
      holderPid: holder.pid,
      ...(holder.startedAt !== undefined ? { holderStartedAt: holder.startedAt } : {}),
    };
  }

  await mkdir(dirname(deps.lockPath), { recursive: true });
  const tmpPath = `${deps.lockPath}.tmp-${pid}`;
  await writeFile(tmpPath, JSON.stringify({ pid, startedAt: new Date(now()).toISOString() }, null, 2));
  await rename(tmpPath, deps.lockPath);
  return { acquired: true };
}

async function readHolder(
  lockPath: string,
): Promise<{ pid: number; startedAt?: string } | undefined> {
  try {
    const parsed: unknown = JSON.parse(await readFile(lockPath, 'utf8'));
    if (typeof parsed === 'object' && parsed !== null) {
      const { pid, startedAt } = parsed as { pid?: unknown; startedAt?: unknown };
      if (typeof pid === 'number' && Number.isInteger(pid) && pid > 0) {
        return { pid, ...(typeof startedAt === 'string' ? { startedAt } : {}) };
      }
    }
  } catch {
    // Missing or unreadable lock = no holder. A corrupt file must not deadlock every restart.
  }
  return undefined;
}
