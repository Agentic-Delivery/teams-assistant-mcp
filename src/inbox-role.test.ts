import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { decideInboxRole, isInboxDisabled } from './inbox-role.js';
import { acquirePollerLock } from './poller-lock.js';

/**
 * PR #20 blocker (2026-09-04 review): index.ts's startup decision — poll vs. serve-only vs.
 * disabled — was the one behaviour in the poller-supervision change with zero test coverage.
 * These drive decideInboxRole directly, the pure function index.ts (via inbox-startup.ts) now
 * calls instead of deciding inline.
 */
describe('decideInboxRole', () => {
  it('lock free (acquired) with the poller enabled: poll', () => {
    expect(decideInboxRole({ env: {}, lock: { acquired: true } })).toBe('poll');
  });

  it('lock held by a live instance: serve-only, not disabled and not poll', () => {
    const lock = { acquired: false as const, holderPid: 4242, holderStartedAt: '2026-09-04T10:00:00.000Z' };
    expect(decideInboxRole({ env: {}, lock })).toBe('serve-only');
  });

  it('TEAMS_INBOX_DISABLED overrides an acquired lock: disabled, not poll', () => {
    expect(decideInboxRole({ env: { TEAMS_INBOX_DISABLED: '1' }, lock: { acquired: true } })).toBe(
      'disabled',
    );
  });

  it('TEAMS_INBOX_DISABLED overrides a contended lock too: disabled, not serve-only', () => {
    const lock = { acquired: false as const, holderPid: 4242 };
    expect(decideInboxRole({ env: { TEAMS_INBOX_DISABLED: '1' }, lock })).toBe('disabled');
  });

  it.each([
    ['unset', undefined],
    ['empty string', ''],
    ['the literal "0"', '0'],
  ])('TEAMS_INBOX_DISABLED %s: not disabled', (_label, value) => {
    expect(isInboxDisabled(value === undefined ? {} : { TEAMS_INBOX_DISABLED: value })).toBe(false);
  });

  it.each([
    ['1', '1'],
    ['true', 'true'],
    ['padded with whitespace', '  1  '],
  ])('TEAMS_INBOX_DISABLED = %s: disabled', (_label, value) => {
    expect(isInboxDisabled({ TEAMS_INBOX_DISABLED: value })).toBe(true);
  });
});

describe('decideInboxRole against a REAL acquirePollerLock (stale-lock takeover)', () => {
  let dir: string;
  let lockPath: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'inbox-role-'));
    lockPath = join(dir, 'poller.lock');
  });
  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it('a stale (dead-pid) holder is taken over: poll, not serve-only', async () => {
    // A prior instance died without cleaning up its lock file — the reboot case.
    await acquirePollerLock({ lockPath, pid: 1111, isPidAlive: () => true, nowFn: () => 1_000 });

    const lock = await acquirePollerLock({
      lockPath,
      pid: 2222,
      isPidAlive: () => false,
      nowFn: () => 2_000,
    });

    expect(decideInboxRole({ env: {}, lock })).toBe('poll');
  });

  it('a live holder is NOT taken over: serve-only, not poll', async () => {
    await acquirePollerLock({ lockPath, pid: 1111, isPidAlive: () => true, nowFn: () => 1_000 });

    const lock = await acquirePollerLock({
      lockPath,
      pid: 2222,
      isPidAlive: () => true,
      nowFn: () => 2_000,
    });

    expect(decideInboxRole({ env: {}, lock })).toBe('serve-only');
  });
});
