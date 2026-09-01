import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { DEFAULT_MEMBERS_TTL_MS, MembersCache } from './members-cache.js';

const CHAT = '19:pilot@thread.v2';
const members = [
  { id: 'aad-1', displayName: 'Garg, Shivankit' },
  { id: 'aad-2', displayName: 'Spännare, Johan' },
];

describe('MembersCache — disk-persisted, per-chat, TTL-bounded (0.4.1)', () => {
  let dir: string;
  let path: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'members-cache-test-'));
    path = join(dir, 'members-cache.json');
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it('returns undefined for a chat never cached', () => {
    const cache = new MembersCache({ path });

    expect(cache.get(CHAT)).toBeUndefined();
  });

  it('returns what was set, for the same chat only', () => {
    const cache = new MembersCache({ path });

    cache.set(CHAT, members);

    expect(cache.get(CHAT)).toEqual(members);
    expect(cache.get('19:other@thread.v2')).toBeUndefined();
  });

  it('survives a fresh instance over the same file — the point of disk persistence', () => {
    new MembersCache({ path }).set(CHAT, members);

    const restarted = new MembersCache({ path });

    expect(restarted.get(CHAT)).toEqual(members);
  });

  it('defaults the TTL to 24 hours', () => {
    let clock = 1_000_000;
    const cache = new MembersCache({ path, now: () => clock });
    cache.set(CHAT, members);

    clock += DEFAULT_MEMBERS_TTL_MS - 1;
    expect(cache.get(CHAT)).toEqual(members);

    clock += 2; // now past the TTL
    expect(cache.get(CHAT)).toBeUndefined();
  });

  it('honours a shorter injected TTL, expiring the entry once it elapses', () => {
    let clock = 0;
    const cache = new MembersCache({ path, ttlMs: 1000, now: () => clock });
    cache.set(CHAT, members);

    clock = 999;
    expect(cache.get(CHAT)).toEqual(members);
    clock = 1001;
    expect(cache.get(CHAT)).toBeUndefined();
  });

  it('a corrupt cache file degrades to "never cached" rather than throwing', async () => {
    const { mkdir, writeFile } = await import('node:fs/promises');
    await mkdir(dir, { recursive: true });
    await writeFile(path, 'not json at all');
    const cache = new MembersCache({ path });

    expect(cache.get(CHAT)).toBeUndefined();
  });

  it('writing never leaves a partially-written file behind — set is atomic', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, members);

    const { readdir } = await import('node:fs/promises');
    const entries = await readdir(dir);

    // Only the real file remains; no leftover .tmp-* from the rename-based write.
    expect(entries).toEqual(['members-cache.json']);
  });
});
