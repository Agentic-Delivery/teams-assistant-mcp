import { dirname, join } from 'node:path';
import { buildInboxPoller, type BuildInboxPollerOptions } from './build-inbox-poller.js';
import type { InboxPoller } from './inbox.js';
import { decideInboxRole, isInboxDisabled, type InboxRole } from './inbox-role.js';
import { acquirePollerLock } from './poller-lock.js';

export interface StartInboxSupervisionOptions {
  env: NodeJS.ProcessEnv;
  inboxPath: string;
  /** Everything buildInboxPoller needs except inboxPath and log, which this function supplies. */
  pollerOptions: Omit<BuildInboxPollerOptions, 'inboxPath' | 'log'>;
  log?: (line: string) => void;
  /** Injectable for tests only; production always wires the real buildInboxPoller. */
  buildPoller?: (options: BuildInboxPollerOptions) => InboxPoller;
}

export interface InboxSupervisionResult {
  role: InboxRole;
  poller?: InboxPoller;
}

/**
 * The inbox poller's startup wiring (extracted from index.ts, PR #20 blocker fix, 2026-09-04
 * review): acquires the per-inbox lock, decides the role via decideInboxRole (inbox-role.ts),
 * and acts on it. Crucially, the losing side of a contended lock NEVER calls process.exit — see
 * inbox-role.ts's InboxRole doc comment for why. This function's own return value (rather than a
 * process-wide side effect) is what makes that "no exit" behaviour something a test can assert
 * directly, instead of having to spawn a subprocess to observe it.
 */
export async function startInboxSupervision(
  options: StartInboxSupervisionOptions,
): Promise<InboxSupervisionResult> {
  const log = options.log ?? ((line: string) => process.stderr.write(`${line}\n`));
  const buildPoller = options.buildPoller ?? buildInboxPoller;

  if (isInboxDisabled(options.env)) {
    log('inbox poller off (TEAMS_INBOX_DISABLED)');
    return { role: 'disabled' };
  }

  const lockPath = join(dirname(options.inboxPath), 'poller.lock');
  const lock = await acquirePollerLock({ lockPath });
  const role = decideInboxRole({ env: options.env, lock });

  if (role === 'serve-only') {
    log(`inbox poller: another instance holds ${lockPath}; serving tools without polling`);
    return { role };
  }

  const poller = buildPoller({ ...options.pollerOptions, inboxPath: options.inboxPath, log });
  poller.start();
  log(`inbox poller on: ${options.inboxPath}`);
  return { role, poller };
}
