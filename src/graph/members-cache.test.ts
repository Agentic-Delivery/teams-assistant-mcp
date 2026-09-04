import { mkdirSync, renameSync, statSync, writeFileSync } from 'node:fs';
import { mkdtemp, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
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

  // 0.4.1 review round 1, MINOR: the top-level shape check (object, not array) let a per-entry
  // garbage shape straight through — `fetchedAt: "nope"` makes `now() - fetchedAt` produce NaN,
  // and `NaN > ttlMs` is false, so a garbage entry used to read as "still fresh" and get served
  // as real members. Each entry's own shape is now validated too.
  it('an entry whose members is not an array degrades to "never cached" for that chat', async () => {
    const { writeFile } = await import('node:fs/promises');
    await writeFile(path, JSON.stringify({ [CHAT]: { members: 'not-an-array', fetchedAt: Date.now() } }));
    const cache = new MembersCache({ path });

    expect(cache.get(CHAT)).toBeUndefined();
  });

  it('an entry whose fetchedAt is not a number degrades to "never cached" rather than reading as eternally fresh', async () => {
    const { writeFile } = await import('node:fs/promises');
    await writeFile(path, JSON.stringify({ [CHAT]: { members, fetchedAt: 'not-a-number' } }));
    const cache = new MembersCache({ path });

    expect(cache.get(CHAT)).toBeUndefined();
  });

  it('one chat with a garbage entry does not poison a sibling chat with a valid one', async () => {
    const { writeFile } = await import('node:fs/promises');
    await writeFile(
      path,
      JSON.stringify({
        [CHAT]: { members: null, fetchedAt: Date.now() },
        '19:other@thread.v2': { members, fetchedAt: Date.now() },
      }),
    );
    const cache = new MembersCache({ path });

    expect(cache.get(CHAT)).toBeUndefined();
    expect(cache.get('19:other@thread.v2')).toEqual(members);
  });

  // Mitigation 1 (docs/throttling-mitigation.md §4, stage 1 item 2): a cache hit must never
  // become a hard dependency on the throttled /members endpoint. get() answers undefined past
  // TTL by design (existing behaviour above) — getStale() is the escape hatch a caller reaches
  // for ONLY when a live refresh itself failed, so a throttled refresh can still be served from
  // what's on disk instead of failing the whole mention resolution.
  describe('getStale — the TTL-ignoring read used only when a live refresh fails (mitigation 1)', () => {
    it('returns the entry past its TTL, unlike get()', () => {
      let clock = 0;
      const cache = new MembersCache({ path, ttlMs: 1000, now: () => clock });
      cache.set(CHAT, members);
      clock = 5000; // well past the TTL

      expect(cache.get(CHAT)).toBeUndefined(); // still the old, TTL-bound contract
      expect(cache.getStale(CHAT)).toEqual({ members, fetchedAt: 0 });
    });

    it('returns undefined for a chat never cached, same as get()', () => {
      const cache = new MembersCache({ path });

      expect(cache.getStale(CHAT)).toBeUndefined();
    });

    it('a corrupt cache file degrades to undefined rather than throwing', async () => {
      const { writeFile } = await import('node:fs/promises');
      await writeFile(path, 'not json at all');
      const cache = new MembersCache({ path });

      expect(cache.getStale(CHAT)).toBeUndefined();
    });
  });

  // Mitigation 2 (docs/throttling-mitigation.md §4, stage 1 item 2): the inbox poller hands every
  // polled message's (senderId, senderDisplayName) here at zero Graph cost — merge() is how that
  // traffic actually fills/refreshes the roster the throttled /members endpoint used to be the
  // only source of.
  //
  // BLOCKER-1 (2026-09-04 review): merge() used to stamp `fetchedAt` — the SAME field a real
  // /members fetch stamps — on every call, so a harvested-only roster never expired and
  // `membersForInvite` (teams-chats.ts) trusted it verbatim as sendFile's permission-grant list.
  // These tests now pin the PROVENANCE distinction that fixes it: get()/getComplete() answer
  // differently for a PARTIAL (harvest-only) vs COMPLETE (real-fetch) entry, and merge() never
  // extends a COMPLETE entry's own fetchedAt.
  describe('merge — harvesting a roster from message traffic at zero Graph cost (mitigation 2; PARTIAL/COMPLETE provenance, BLOCKER-1 fix)', () => {
    it('BLOCKER-1: a merge into an empty cache creates a PARTIAL entry — get() serves it (mention resolution may), getComplete() refuses it (sendFile\'s grant path must not)', () => {
      let clock = 123;
      const cache = new MembersCache({ path, now: () => clock });

      cache.merge(CHAT, [{ id: 'aad-1', displayName: 'Garg, Shivankit' }]);

      expect(cache.get(CHAT)).toEqual([{ id: 'aad-1', displayName: 'Garg, Shivankit' }]);
      expect(cache.getComplete(CHAT)).toBeUndefined();
      expect(cache.getStaleComplete(CHAT)).toBeUndefined();
      // fetchedAt: 0, not omitted — the 0.5.1-compatible "instantly expired, but shape-valid"
      // marker (MembersCacheEntry.fetchedAt's own doc comment); harvestedAt carries the real time.
      expect(cache.getStale(CHAT)).toEqual({
        members: [{ id: 'aad-1', displayName: 'Garg, Shivankit' }],
        fetchedAt: 0,
      });
    });

    it('adds a newly-seen sender to an existing roster without dropping the members already cached', () => {
      const cache = new MembersCache({ path });
      cache.set(CHAT, [{ id: 'aad-1', displayName: 'Garg, Shivankit' }]);

      cache.merge(CHAT, [{ id: 'aad-2', displayName: 'Spännare, Johan' }]);

      expect(cache.get(CHAT)).toEqual(
        expect.arrayContaining([
          { id: 'aad-1', displayName: 'Garg, Shivankit' },
          { id: 'aad-2', displayName: 'Spännare, Johan' },
        ]),
      );
    });

    it('BLOCKER-1: a merge into an already-COMPLETE (set()) roster keeps it COMPLETE and keeps its ORIGINAL fetchedAt — traffic does not extend the grant-path freshness window', () => {
      let clock = 0;
      const cache = new MembersCache({ path, ttlMs: 1000, now: () => clock });
      cache.set(CHAT, [{ id: 'aad-1', displayName: 'Garg, Shivankit' }]); // fetchedAt stamped at clock=0
      clock = 999; // just inside the TTL

      cache.merge(CHAT, [{ id: 'aad-1', displayName: 'Garg, Shivankit' }]);
      clock = 1500; // past the TTL measured from the ORIGINAL fetchedAt (0) — merge must not have moved it

      // Before the fix this stayed fresh forever (merge() restamped fetchedAt to 999). Now the
      // COMPLETE roster expires on its OWN real-fetch schedule, unaffected by the reconfirming merge.
      expect(cache.getComplete(CHAT)).toBeUndefined();
      // get() (mention resolution) DOES still count the merge's harvestedAt as fresh evidence.
      expect(cache.get(CHAT)).toEqual([{ id: 'aad-1', displayName: 'Garg, Shivankit' }]);
    });

    it('a departed member is not kept in a PARTIAL roster\'s grant path forever: getComplete() never serves a partial entry regardless of how recently traffic confirmed it', () => {
      let clock = 0;
      const cache = new MembersCache({ path, now: () => clock });
      cache.merge(CHAT, [{ id: 'aad-1', displayName: 'Garg, Shivankit' }]);
      clock = 1; // freshly harvested a moment ago — get() would happily serve this

      expect(cache.get(CHAT)).toEqual([{ id: 'aad-1', displayName: 'Garg, Shivankit' }]);
      expect(cache.getComplete(CHAT)).toBeUndefined();
    });

    // Violating-double coverage: a message sender Graph (or our own mapping) failed to name
    // properly must never corrupt the persisted roster with a junk entry — the poller's harvest
    // is untrusted input reaching this cache directly, with no /members validation in between.
    it('ignores a harvested pair with no id or no displayName rather than persisting junk', () => {
      const cache = new MembersCache({ path });

      cache.merge(CHAT, [
        { id: '', displayName: 'Nobody' } as { id: string; displayName: string },
        { id: 'aad-1', displayName: '' } as { id: string; displayName: string },
      ]);

      expect(cache.get(CHAT)).toBeUndefined();
    });

    it('a harvest with nothing valid never touches the disk (no write, no rename)', () => {
      const writeFileFn = vi.fn(writeFileSync);
      const renameFn = vi.fn(renameSync);
      const cache = new MembersCache({ path, writeFileFn, renameFn });

      cache.merge(CHAT, []);

      expect(writeFileFn).not.toHaveBeenCalled();
      expect(renameFn).not.toHaveBeenCalled();
    });
  });

  // getComplete/getStaleComplete: the reads membersForInvite (sendFile's permission-grant roster,
  // teams-chats.ts) uses instead of get()/getStale() — BLOCKER-1 fix, 2026-09-04 review.
  describe('getComplete / getStaleComplete — COMPLETE-only reads for the permission-grant path (BLOCKER-1 fix)', () => {
    it('a fresh set() entry is served by getComplete()', () => {
      const cache = new MembersCache({ path });
      cache.set(CHAT, members);

      expect(cache.getComplete(CHAT)).toEqual(members);
    });

    it('a legacy entry with no complete flag at all (pre-0.5.2 on-disk shape) reads as COMPLETE', async () => {
      const { writeFile } = await import('node:fs/promises');
      await writeFile(path, JSON.stringify({ [CHAT]: { members, fetchedAt: Date.now() } }));
      const cache = new MembersCache({ path });

      expect(cache.getComplete(CHAT)).toEqual(members);
    });

    it('getComplete() honours the same TTL as get() against fetchedAt', () => {
      let clock = 0;
      const cache = new MembersCache({ path, ttlMs: 1000, now: () => clock });
      cache.set(CHAT, members);
      clock = 1001;

      expect(cache.getComplete(CHAT)).toBeUndefined();
    });

    it('getStaleComplete() serves a COMPLETE entry past its TTL, unlike getComplete()', () => {
      let clock = 0;
      const cache = new MembersCache({ path, ttlMs: 1000, now: () => clock });
      cache.set(CHAT, members);
      clock = 5000;

      expect(cache.getComplete(CHAT)).toBeUndefined();
      expect(cache.getStaleComplete(CHAT)).toEqual({ members, fetchedAt: 0 });
    });

    it('getStaleComplete() refuses a PARTIAL (harvest-only) entry even though getStale() would serve it', () => {
      const cache = new MembersCache({ path });
      cache.merge(CHAT, [{ id: 'aad-1', displayName: 'Garg, Shivankit' }]);

      expect(cache.getStale(CHAT)).toBeDefined();
      expect(cache.getStaleComplete(CHAT)).toBeUndefined();
    });

    it('returns undefined for a chat never cached', () => {
      const cache = new MembersCache({ path });

      expect(cache.getComplete(CHAT)).toBeUndefined();
      expect(cache.getStaleComplete(CHAT)).toBeUndefined();
    });
  });

  // MAJOR 4 (review round 1): the previous version of this test only asserted no leftover
  // .tmp-* file was left behind, which passes identically whether set() writes via temp+rename
  // OR writes the destination directly (mutation-verified: deleting the rename call and writing
  // straight to `path` left this exact assertion green). Injecting the write primitives is what
  // makes the destination-reached-only-via-rename claim an honest, mutation-proof observable.
  it('set() reaches the destination path ONLY via rename, never a direct write (atomicity), and the tmp file is already 0600 before that rename', () => {
    const directWrites: string[] = [];
    const renames: Array<{ from: string; to: string }> = [];
    const tmpModesAtRenameTime: number[] = [];
    const cache = new MembersCache({
      path,
      writeFileFn: (target, data, options) => {
        directWrites.push(String(target));
        writeFileSync(target, data, options);
      },
      renameFn: (from, to) => {
        // review round 2, follow-on: mutation M4b showed the tmp file's OWN mode can be dropped
        // unnoticed when only the final path's mode is asserted (a chmodSync on the final path
        // alone still leaves this test green). Capturing the tmp file's mode HERE — after
        // writeFile+chmodSync have run, before the rename moves it away — is the only point it
        // can still be observed at all.
        tmpModesAtRenameTime.push(statSync(from).mode & 0o777);
        renames.push({ from: String(from), to: String(to) });
        renameSync(from, to);
      },
    });

    cache.set(CHAT, members);

    expect(directWrites).toHaveLength(1);
    expect(directWrites[0]).not.toBe(path); // never written to the real destination directly
    expect(renames).toEqual([{ from: directWrites[0], to: path }]); // the ONLY route to `path` is a rename FROM that exact temp file
    expect(tmpModesAtRenameTime).toEqual([0o600]);
    expect(cache.get(CHAT)).toEqual(members); // and the data really did land
  });

  // MAJOR 2 (review round 1): mirrors the token-cache sibling's permission posture exactly —
  // 0600, chmodSync as belt-and-braces (writeFileSync's mode option only applies on create),
  // applied to the tmp file before the rename AND to the final path after it.
  it('writes the cache file readable only by the owner (0600), matching the token cache sibling', () => {
    const cache = new MembersCache({ path });

    cache.set(CHAT, members);

    expect(statSync(path).mode & 0o777).toBe(0o600);
  });

  it('tightens permissions on a cache file that already existed too open', () => {
    mkdirSync(dir, { recursive: true });
    writeFileSync(path, '{}', { mode: 0o644 });

    new MembersCache({ path }).set(CHAT, members);

    expect(statSync(path).mode & 0o777).toBe(0o600);
  });

  it('leaves no tmp file behind after a successful write', async () => {
    const cache = new MembersCache({ path });

    cache.set(CHAT, members);

    const entries = await readdir(dir);
    expect(entries).toEqual(['members-cache.json']);
  });
});
