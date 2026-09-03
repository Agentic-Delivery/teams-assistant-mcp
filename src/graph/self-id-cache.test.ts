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

  // NIT (review round 2): a distinct predicate from the one above — `typeof parsed.id !==
  // 'string'` alone would happily accept an EMPTY string as a valid id, which resolveSelfId would
  // then memoize and use as "self" (silently disabling self-exclusion: `member.id !== ''` is true
  // for every real member, so nobody is ever excluded from the grant). The `.trim() === ''` half
  // of the same `if` is what actually rejects this shape.
  it('an entry with an EMPTY string id degrades to "never cached" (distinct from a non-string id)', async () => {
    await writeFile(path, JSON.stringify({ id: '', resolvedAt: 1 }));
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

// [MINOR] review round 2: a stale cache entry surviving TEAMS_MCP_USERNAME being repointed at
// the same instance dir (a consuming project switching which service account it uses, without
// switching TEAMS_MCP_TOKEN_CACHE/selfIdCachePath to a fresh directory) used to be trusted
// verbatim — the OLD account's AAD id, read back as if it were the NEW account's. Stamping the
// resolving account's username into the entry, and treating a mismatch as a plain miss (not a
// throw — same posture as every other "garbage in, miss out" shape check this cache already
// does), closes that specific case.
describe('FileSelfIdCache — username stamp (review round 2, closes the "username repointed at the same instance dir" case)', () => {
  let dir: string;
  let path: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'self-id-cache-username-'));
    path = join(dir, '.self-id-cache.json');
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it('write() stamps expectedUsername into the persisted entry', async () => {
    const cache = new FileSelfIdCache({ path, expectedUsername: 'assistant@example.com' });

    cache.write({ id: 'aad-1', resolvedAt: 1 });

    const raw = JSON.parse(await (await import('node:fs/promises')).readFile(path, 'utf8'));
    expect(raw).toEqual({ id: 'aad-1', resolvedAt: 1, username: 'assistant@example.com' });
  });

  it('read() treats an entry stamped with a DIFFERENT username as a miss', () => {
    new FileSelfIdCache({ path, expectedUsername: 'old-account@example.com' }).write({
      id: 'aad-old-account',
      resolvedAt: 1,
    });

    const repointed = new FileSelfIdCache({ path, expectedUsername: 'new-account@example.com' });

    expect(repointed.read()).toBeUndefined();
  });

  it('read() treats a legacy entry with no username field as a miss once expectedUsername is set', async () => {
    const { writeFile } = await import('node:fs/promises');
    await writeFile(path, JSON.stringify({ id: 'aad-legacy', resolvedAt: 1 }));

    const cache = new FileSelfIdCache({ path, expectedUsername: 'assistant@example.com' });

    expect(cache.read()).toBeUndefined();
  });

  it('read() accepts an entry stamped with the SAME username', () => {
    new FileSelfIdCache({ path, expectedUsername: 'assistant@example.com' }).write({
      id: 'aad-1',
      resolvedAt: 1,
    });

    const cache = new FileSelfIdCache({ path, expectedUsername: 'assistant@example.com' });

    expect(cache.read()).toEqual({ id: 'aad-1', resolvedAt: 1 });
  });

  it('omitting expectedUsername skips the check entirely — backward compatible with every other test in this file', () => {
    new FileSelfIdCache({ path, expectedUsername: 'assistant@example.com' }).write({
      id: 'aad-1',
      resolvedAt: 1,
    });

    const noUsernameCheck = new FileSelfIdCache({ path });

    expect(noUsernameCheck.read()).toEqual({ id: 'aad-1', resolvedAt: 1 });
  });
});

describe('NullSelfIdCache — the safe default when no persisted cache is wired', () => {
  it('always misses on read and silently drops writes', () => {
    const cache = new NullSelfIdCache();

    cache.write({ id: 'aad-1', resolvedAt: 1 });

    expect(cache.read()).toBeUndefined();
  });
});
