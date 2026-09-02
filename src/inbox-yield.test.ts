import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  DEFAULT_YIELD_HOLD_MS,
  MAX_YIELD_MS,
  inboxPathFor,
  inboxYieldPathFor,
  readYield,
  releaseYield,
  requestYield,
  withQuotaYield,
} from './inbox-yield.js';

describe('quota yield — the file the poller and ad-hoc readers coordinate through', () => {
  let dir: string;
  let path: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'yield-test-'));
    path = join(dir, 'nested', 'inbox-yield.json');
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it('request → active with the default hold; release → gone', async () => {
    const now = 1_000_000;
    await requestYield(path, { reason: 'teams-attachments', now });

    const standing = await readYield(path, now);
    expect(standing?.reason).toBe('teams-attachments');
    expect(standing?.pid).toBe(process.pid);
    expect(standing?.until).toBe(now + DEFAULT_YIELD_HOLD_MS);

    await releaseYield(path, now);
    expect(await readYield(path, now)).toBeUndefined();
    expect(existsSync(path)).toBe(false);
  });

  it('an expired yield is no yield — a crashed reader silences the inbox only until the deadline', async () => {
    await requestYield(path, { reason: 'crashy', holdMs: 5_000, now: 1_000_000 });

    expect(await readYield(path, 1_004_999)).toBeDefined();
    expect(await readYield(path, 1_005_000)).toBeUndefined();
  });

  it('a missing or corrupt file degrades to "keep polling", never a crash', async () => {
    expect(await readYield(path)).toBeUndefined();
    await requestYield(path, { reason: 'x' }); // ensures the directory exists
    await writeFile(path, 'not json at all');
    expect(await readYield(path)).toBeUndefined();
  });

  it('an absurd deadline is clamped to MAX_YIELD_MS from now, whatever the file claims', async () => {
    await requestYield(path, { reason: 'x' });
    await writeFile(path, JSON.stringify({ pid: 1, reason: 'clock-broken', until: Number.MAX_SAFE_INTEGER }));

    const standing = await readYield(path, 1_000_000);
    expect(standing?.until).toBe(1_000_000 + MAX_YIELD_MS);
  });

  it('a second request extends a standing yield, never shortens it', async () => {
    await requestYield(path, { reason: 'long job', holdMs: 300_000, now: 1_000_000 });
    await requestYield(path, { reason: 'quick job', holdMs: 10_000, now: 1_001_000 });

    const standing = await readYield(path, 1_001_000);
    expect(standing?.until).toBe(1_000_000 + 300_000); // the longer deadline survived
  });

  it('release leaves another process\'s still-fresh yield alone', async () => {
    await requestYield(path, { reason: 'x' }); // ensures the directory exists
    const foreign = { pid: process.pid + 1, reason: 'their download', until: Date.now() + 60_000 };
    await writeFile(path, JSON.stringify(foreign));

    await releaseYield(path);

    expect(JSON.parse(await readFile(path, 'utf8'))).toEqual(foreign);
  });

  it('withQuotaYield holds the file across fn and releases it after — on failure too', async () => {
    let duringSuccess = false;
    await withQuotaYield(path, 'happy', async () => {
      duringSuccess = (await readYield(path)) !== undefined;
    });
    expect(duringSuccess).toBe(true);
    expect(await readYield(path)).toBeUndefined();

    let duringFailure = false;
    await expect(
      withQuotaYield(path, 'sad', async () => {
        duringFailure = (await readYield(path)) !== undefined;
        throw new Error('download died');
      }),
    ).rejects.toThrow('download died');
    expect(duringFailure).toBe(true);
    // A dead download must not leave the inbox silenced until the deadline.
    expect(await readYield(path)).toBeUndefined();
  });

  it('withQuotaYield without a path just runs fn — no daemon in play, nothing to coordinate', async () => {
    expect(await withQuotaYield(undefined, 'n/a', async () => 42)).toBe(42);
  });
});

describe('path derivation — one derivation shared by the server and every CLI', () => {
  it('the yield file lives next to the inbox, following a TEAMS_INBOX_PATH override', () => {
    const env = { TEAMS_INBOX_PATH: '/data/teams/inbox.jsonl' } as NodeJS.ProcessEnv;
    expect(inboxPathFor(env)).toBe('/data/teams/inbox.jsonl');
    expect(inboxYieldPathFor(env)).toBe('/data/teams/inbox-yield.json');
  });

  it('defaults under ~/.teams-assistant, same as the daemon', () => {
    const env = {} as NodeJS.ProcessEnv;
    expect(inboxPathFor(env).endsWith('/.teams-assistant/inbox.jsonl')).toBe(true);
    expect(inboxYieldPathFor(env).endsWith('/.teams-assistant/inbox-yield.json')).toBe(true);
  });
});
