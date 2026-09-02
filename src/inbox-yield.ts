import { mkdir, readFile, unlink, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { dirname, join, resolve } from 'node:path';

/**
 * Cross-process quota coordination between the inbox poller and ad-hoc Graph readers.
 *
 * Why this exists (measured live 2026-09-02): Graph's read quota is per mailbox, and the inbox
 * poller consumes it CONTINUOUSLY — with a poller running, ad-hoc single-message GETs answered
 * 429 with `retry-after: 62` on every attempt across 20+ minutes of patient backoff. Honouring
 * Retry-After is necessary but not sufficient: the waiting caller and the poller share one
 * budget, and the poller spends it again the moment the window reopens. The only thing that
 * makes room is the POLLER going quiet — so an attachment download asks it to.
 *
 * The mechanism is a file, because files in the inbox directory are already how this package
 * coordinates across processes (inbox.jsonl, the inbox-state.json sidecar): a reader that needs
 * quota writes `inbox-yield.json` next to the inbox with a deadline, the poller checks it at the
 * top of every cycle and skips polling while it stands, and the reader removes it when done. It
 * works identically whether the reader is an MCP tool in the SAME process as a poller or a
 * standalone CLI next to a running daemon — the poller never cares who wrote the file.
 *
 * Deliberately rough edges, chosen over a lock manager: the deadline caps the pause (a crashed
 * reader silences the inbox for at most MAX_YIELD_MS), a second reader extends rather than
 * shortens a standing yield, and release leaves another process's still-fresh yield alone. Two
 * overlapping readers in ONE process can still release each other early — the cost is one poll
 * cycle's worth of contention, not a correctness problem, and not worth reference counting.
 */
export interface QuotaYield {
  pid: number;
  reason: string;
  /** Epoch ms after which the yield no longer holds, whether or not it was released. */
  until: number;
}

/** Long enough for a message fetch, one live-measured 62 s Retry-After wait, and several file
 *  downloads; short enough that a crashed reader costs the inbox one skipped poll or two. */
export const DEFAULT_YIELD_HOLD_MS = 180_000;
/** No reader may silence the inbox longer than this, whatever its file claims. */
export const MAX_YIELD_MS = 10 * 60_000;

/** THE inbox path, exactly as index.ts resolves it — one derivation, shared with the CLIs so a
 *  standalone process and the server can never disagree about where the coordination files live. */
export function inboxPathFor(env: NodeJS.ProcessEnv = process.env): string {
  return resolve(
    env['TEAMS_INBOX_PATH']?.trim() || join(homedir(), '.teams-assistant', 'inbox.jsonl'),
  );
}

/** The yield file lives next to the inbox, same convention as the inbox-state.json sidecar. */
export function inboxYieldPathFor(env: NodeJS.ProcessEnv = process.env): string {
  return join(dirname(inboxPathFor(env)), 'inbox-yield.json');
}

/** The currently-standing yield, or undefined when there is none, it expired, or the file is
 *  unreadable/corrupt (a broken yield file must degrade to "keep polling", never to a crash).
 *  An absurd deadline — a bad clock, a hand-edited file — is clamped to MAX_YIELD_MS from now. */
export async function readYield(path: string, now = Date.now()): Promise<QuotaYield | undefined> {
  try {
    const parsed = JSON.parse(await readFile(path, 'utf8')) as Partial<QuotaYield>;
    if (typeof parsed.until !== 'number' || parsed.until <= now) {
      return undefined;
    }
    return {
      pid: typeof parsed.pid === 'number' ? parsed.pid : -1,
      reason: typeof parsed.reason === 'string' ? parsed.reason : 'unknown',
      until: Math.min(parsed.until, now + MAX_YIELD_MS),
    };
  } catch {
    return undefined;
  }
}

export async function requestYield(
  path: string,
  options: { reason: string; holdMs?: number; now?: number },
): Promise<void> {
  const now = options.now ?? Date.now();
  const holdMs = Math.min(options.holdMs ?? DEFAULT_YIELD_HOLD_MS, MAX_YIELD_MS);
  // Extend, never shorten: a standing yield some other reader still needs keeps its deadline.
  const standing = await readYield(path, now);
  const until = Math.max(now + holdMs, standing?.until ?? 0);
  await mkdir(dirname(path), { recursive: true });
  await writeFile(path, JSON.stringify({ pid: process.pid, reason: options.reason, until } satisfies QuotaYield));
}

export async function releaseYield(path: string, now = Date.now()): Promise<void> {
  const standing = await readYield(path, now);
  if (standing && standing.pid !== process.pid) {
    return; // another process's still-fresh yield is not ours to remove
  }
  await unlink(path).catch(() => undefined);
}

/**
 * Runs `fn` under a quota yield: requested before, released after — on failure too, because a
 * download that died must not leave the inbox silenced until the deadline. With no path (a
 * caller that opted out, a test with no daemon in play) it just runs `fn`. The yield write
 * itself is best-effort: an unwritable inbox directory must degrade to "download without
 * coordination", never block the download it exists to serve.
 */
export async function withQuotaYield<T>(
  path: string | undefined,
  reason: string,
  fn: () => Promise<T>,
): Promise<T> {
  if (!path) {
    return fn();
  }
  await requestYield(path, { reason }).catch(() => undefined);
  try {
    return await fn();
  } finally {
    await releaseYield(path).catch(() => undefined);
  }
}
