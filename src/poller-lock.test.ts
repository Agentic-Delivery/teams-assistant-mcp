import { mkdtemp, readdir, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { acquirePollerLock } from './poller-lock.js';

/**
 * The 2026-09-01 twin-daemon race: two servers on one shared account meant every poll interval
 * carried double the Graph traffic, and under throttling each daemon's retries fed the other's
 * penalty window. The lock makes "one poller per inbox" a checked invariant instead of an
 * assumption about how many processes happen to be running.
 */
describe('poller lock', () => {
  let dir: string;
  let lockPath: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'poller-lock-'));
    lockPath = join(dir, 'nested', 'poller.lock');
  });
  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  async function lockContent(): Promise<Record<string, unknown>> {
    return JSON.parse(await readFile(lockPath, 'utf8')) as Record<string, unknown>;
  }

  it('acquires when no lock exists, recording pid and start time', async () => {
    const result = await acquirePollerLock({
      lockPath,
      pid: 4242,
      isPidAlive: () => true,
      nowFn: () => 1_756_700_000_000,
    });

    expect(result).toEqual({ acquired: true });
    expect(await lockContent()).toEqual({
      pid: 4242,
      startedAt: new Date(1_756_700_000_000).toISOString(),
    });
  });

  it('refuses when the holder pid is alive, naming the holder', async () => {
    const checked: number[] = [];
    await acquirePollerLock({ lockPath, pid: 1111, isPidAlive: () => true, nowFn: () => 1_000 });

    const result = await acquirePollerLock({
      lockPath,
      pid: 2222,
      isPidAlive: (pid) => {
        checked.push(pid);
        return true;
      },
      nowFn: () => 2_000,
    });

    expect(result).toEqual({
      acquired: false,
      holderPid: 1111,
      holderStartedAt: new Date(1_000).toISOString(),
    });
    expect(checked).toEqual([1111]);
    // The refused caller must not have clobbered the standing lock.
    expect((await lockContent())['pid']).toBe(1111);
  });

  it('takes over a dead holder\'s lock — the reboot case', async () => {
    await acquirePollerLock({ lockPath, pid: 1111, isPidAlive: () => true, nowFn: () => 1_000 });

    const result = await acquirePollerLock({
      lockPath,
      pid: 2222,
      isPidAlive: () => false,
      nowFn: () => 2_000,
    });

    expect(result).toEqual({ acquired: true });
    expect((await lockContent())['pid']).toBe(2222);
  });

  it('re-acquires its own lock — a restart that got the same pid back is not a twin', async () => {
    await acquirePollerLock({ lockPath, pid: 4242, isPidAlive: () => true, nowFn: () => 1_000 });

    const result = await acquirePollerLock({
      lockPath,
      pid: 4242,
      // Deliberately claims alive: our own pid being "alive" is us, not a rival.
      isPidAlive: () => true,
      nowFn: () => 2_000,
    });

    expect(result).toEqual({ acquired: true });
  });

  it('takes over a corrupt lock file rather than deadlocking on garbage', async () => {
    const { mkdir } = await import('node:fs/promises');
    await mkdir(join(dir, 'nested'), { recursive: true });
    await writeFile(lockPath, 'not json at all');

    const result = await acquirePollerLock({
      lockPath,
      pid: 2222,
      isPidAlive: () => true,
      nowFn: () => 2_000,
    });

    expect(result).toEqual({ acquired: true });
    expect((await lockContent())['pid']).toBe(2222);
  });

  it('writes the lock atomically — no .tmp leftovers', async () => {
    await acquirePollerLock({ lockPath, pid: 4242, isPidAlive: () => true, nowFn: () => 1_000 });

    const leftovers = (await readdir(join(dir, 'nested'))).filter((name) => name.includes('.tmp'));
    expect(leftovers).toEqual([]);
  });
});
