import type { PollerLockResult } from './poller-lock.js';

/**
 * The inbox poller's startup role, decided once at process start (index.ts) and driving what
 * happens after `acquirePollerLock` — extracted as its own pure function (PR #20 blocker fix,
 * 2026-09-04 review) so the decision has a test independent of index.ts's top-level `main()`,
 * which was previously the ONE behaviour in the poller-supervision change with zero coverage.
 *
 * - 'disabled': TEAMS_INBOX_DISABLED opts this process out entirely — no lock touched, no poller.
 * - 'poll': this process holds the per-inbox lock (poller-lock.ts) and runs the poller.
 * - 'serve-only': another live process already holds the lock. Before this fix, index.ts called
 *   `process.exit(1)` here — correct for a real second daemon, but wrong for the (verified live
 *   on this host, see the PR #20 review comment) MCP-tools registration that runs the SAME
 *   `dist/index.js` against the SAME daemon `.env`, and therefore resolves the SAME default
 *   inbox path and lock: exiting there silently drops every Teams MCP tool for that session, for
 *   a reason no MCP client surfaces. The lock's job is "one poller per inbox", and that job is
 *   fully satisfied by this process simply not starting a second poller — it still serves every
 *   non-poller MCP tool normally.
 */
export type InboxRole = 'poll' | 'serve-only' | 'disabled';

/**
 * TEAMS_INBOX_DISABLED is opt-OUT: absent, empty, or "0" means the poller runs; anything else
 * (including plain whitespace-trimmed truthy strings like "1" or "true") disables it. This is
 * the one place that reads the var — index.ts and decideInboxRole both call this instead of each
 * keeping their own copy of the parsing rule.
 */
export function isInboxDisabled(env: NodeJS.ProcessEnv): boolean {
  const disabled = env['TEAMS_INBOX_DISABLED']?.trim();
  return disabled !== undefined && disabled !== '' && disabled !== '0';
}

export interface DecideInboxRoleOptions {
  env: NodeJS.ProcessEnv;
  lock: PollerLockResult;
}

/**
 * Pure: no I/O, no process.exit, no logging. Callers act on the returned role — see
 * inbox-startup.ts's startInboxSupervision for the one production caller.
 */
export function decideInboxRole({ env, lock }: DecideInboxRoleOptions): InboxRole {
  if (isInboxDisabled(env)) {
    return 'disabled';
  }
  return lock.acquired ? 'poll' : 'serve-only';
}
