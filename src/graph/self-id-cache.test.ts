import { renameSync, statSync, writeFileSync } from 'node:fs';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { FileSelfIdCache, NullSelfIdCache } from './self-id-cache.js';

// Disk mechanics that back scenario (e) from the 0.5.1 dispatch ("cache written on first
// success, file mode 0600") — modeled directly on members-cache.test.ts (0.4.1), the sibling
// this class is deliberately built to match.
describe('FileSelfIdCache — disk-persisted, no TTL (0.5.1, live-diagnosed /me throttle)', () => {
  let dir: string;
  let path: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'self-id-cache-test-'));
    path = join(dir, '.self-id-cache.json');
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it('returns undefined when never written', () => {
    const cache = new FileSelfIdCache({ path });

    expect(cache.read()).toBeUndefined();
  });

  it('returns what was written, written mode 0600 (scenario e)', () => {
    const cache = new FileSelfIdCache({ path, now: () => 1_000 });

    cache.write({ id: 'aad-self-1', resolvedAt: 1_000 });

    expect(cache.read()).toEqual({ id: 'aad-self-1', resolvedAt: 1_000 });
    expect(statSync(path).mode & 0o777).toBe(0o600);
  });

  it('survives a fresh instance over the same file — the point of disk persistence', () => {
    new FileSelfIdCache({ path }).write({ id: 'aad-self-1', resolvedAt: 1_000 });

    const restarted = new FileSelfIdCache({ path });

    expect(restarted.read()).toEqual({ id: 'aad-self-1', resolvedAt: 1_000 });
  });

  it('a later write replaces the earlier one — one value, not an append log', () => {
    const cache = new FileSelfIdCache({ path });
    cache.write({ id: 'aad-old', resolvedAt: 1 });

    cache.write({ id: 'aad-new', resolvedAt: 2 });

    expect(cache.read()).toEqual({ id: 'aad-new', resolvedAt: 2 });
  });

  it('a corrupt cache file degrades to "never cached" rather than throwing', async () => {
    await writeFile(path, 'not json at all');
    const cache = new FileSelfIdCache({ path });

    expect(cache.read()).toBeUndefined();
  });

  it('an entry missing a string id degrades to "never cached" rather than trusting garbage', async () => {
    await writeFile(path, JSON.stringify({ id: 42, resolvedAt: 1 }));
    const cache = new FileSelfIdCache({ path });

    expect(cache.read()).toBeUndefined();
  });

  it('an entry with a non-numeric resolvedAt degrades to "never cached"', async () => {
    await writeFile(path, JSON.stringify({ id: 'aad-1', resolvedAt: 'nope' }));
    const cache = new FileSelfIdCache({ path });

    expect(cache.read()).toBeUndefined();
  });

  it('tightens permissions on a file that already existed too open', () => {
    writeFileSync(path, JSON.stringify({ id: 'aad-1', resolvedAt: 1 }), { mode: 0o644 });

    new FileSelfIdCache({ path }).write({ id: 'aad-2', resolvedAt: 2 });

    expect(statSync(path).mode & 0o777).toBe(0o600);
  });

  // MINOR (review round 2): FileSelfIdCache's constructor has carried writeFileFn/renameFn
  // injection seams since it was first written, but nothing ever drove them — mutation-verified
  // by the reviewer: replacing the rename call in write() with a direct write left the whole
  // suite green. The mechanism itself is now shared with MembersCache (atomic-cache-write.ts,
  // see its own doc comment), and MembersCache's own atomicity test already proves that shared
  // implementation — but that alone does not prove THIS class actually calls it rather than some
  // other write path, which is what this test pins, mirroring members-cache.test.ts's identically
  // named test for its sibling.
  it('write() reaches the destination path ONLY via rename, never a direct write (atomicity)', () => {
    const directWrites: string[] = [];
    const renames: Array<{ from: string; to: string }> = [];
    const cache = new FileSelfIdCache({
      path,
      writeFileFn: (target, data, options) => {
        directWrites.push(String(target));
        writeFileSync(target, data, options);
      },
      renameFn: (from, to) => {
        renames.push({ from: String(from), to: String(to) });
        renameSync(from, to);
      },
    });

    cache.write({ id: 'aad-1', resolvedAt: 1 });

    expect(directWrites).toHaveLength(1);
    expect(directWrites[0]).not.toBe(path); // never written to the real destination directly
    expect(renames).toEqual([{ from: directWrites[0], to: path }]); // the ONLY route to `path` is a rename FROM that exact temp file
    expect(cache.read()).toEqual({ id: 'aad-1', resolvedAt: 1 }); // and the data really did land
  });
});

describe('NullSelfIdCache — the safe default when no persisted cache is wired', () => {
  it('always misses on read and silently drops writes', () => {
    const cache = new NullSelfIdCache();

    cache.write({ id: 'aad-1', resolvedAt: 1 });

    expect(cache.read()).toBeUndefined();
  });
});
