import { readFileSync, type renameSync, type writeFileSync } from 'node:fs';
import { writeAtomicCacheFile } from './atomic-cache-write.js';
import type { ChatMember } from './mentions.js';

/** TEAMS_MCP_MEMBERS_TTL_SECONDS overrides this. Chat membership in these allowlisted chats
 *  effectively never changes, so a day-long TTL trades staleness risk (a departed member still
 *  mentionable for up to this long) for taking the throttled `/members` endpoint off the send
 *  path entirely — see resolveMentions in teams-chats.ts for the refresh-on-miss fallback that
 *  bounds that risk to "one bad mention attempt", not "permanently wrong". Bounds mention
 *  resolution's own `get()`/`getComplete()` freshness window for BOTH a real `/members` fetch
 *  (`fetchedAt`) and traffic-harvested evidence (`harvestedAt`) — but `membersForInvite`
 *  (sendFile's permission-grant roster, teams-chats.ts) never trusts a PARTIAL (harvested-only)
 *  roster at all, regardless of TTL: see `getComplete`'s own doc comment (0.5.2 BLOCKER 1 fix,
 *  2026-09-04 review) for why a departed member merely NOT re-confirmed by traffic must not keep a
 *  file grant past this window either. THE single source of truth for the 24h default: config.ts
 *  derives its own (seconds-flavoured) default from this, rather than hardcoding a second "24h"
 *  that could drift from this one (0.4.1 review round 1). */
export const DEFAULT_MEMBERS_TTL_MS = 24 * 60 * 60 * 1000;

interface MembersCacheEntry {
  members: ChatMember[];
  /** Epoch ms of the last real `GET /chats/{id}/members` fetch (`set()`). `0` for an entry that
   *  has NEVER been backed by a real fetch — a PARTIAL, traffic-harvested-only roster (see
   *  `complete` below). `0` (not omitted) is deliberate: 0.5.1's own `isValidEntry` requires this
   *  field to be a `number` to accept the entry at all, so omitting it would make a 0.5.1 daemon
   *  silently DROP a partial entry on its next read-modify-write cycle; `0` keeps the entry
   *  shape-valid for 0.5.1 (which then reads it as instantly expired via its own `now() - 0 >
   *  ttlMs` check — i.e. "never cached", never "trust this verbatim") while surviving round-trips
   *  through an old daemon untouched. See merge()'s own doc comment for the fuller reasoning
   *  (0.5.2 BLOCKER 1 fix, 2026-09-04 review). */
  fetchedAt: number;
  /** Epoch ms of the most recent traffic harvest (`merge()`) that touched this entry — present
   *  once ANY merge has landed, on a COMPLETE roster too (traffic reconfirming an authoritative
   *  roster is still useful signal for `get()`, just never for `getComplete()`). Absent on an
   *  entry only ever written by `set()`. */
  harvestedAt?: number;
  /** `true` = an authoritative roster confirmed by a real `/members` fetch (`set()`).
   *  `false` = a PARTIAL roster assembled ONLY from traffic (`merge()`) — never confirmed against
   *  `/members`, and therefore never authoritative enough for a permission grant (`getComplete`/
   *  `getStaleComplete` refuse it). Absent reads as `true` — every entry from before this field
   *  existed (0.5.1 and earlier) was written by `set()` alone, since `merge()` did not exist yet,
   *  so a legacy entry with no opinion here IS a real `/members` roster. */
  complete?: boolean;
}

type MembersCacheFile = Record<string, MembersCacheEntry>;

/** True only for a shape `get()`/callers can trust: `fetchedAt` a number (a non-number would make
 *  `now() - fetchedAt` a NaN comparison, which silently reads as "never expires" — 0.4.1 review),
 *  `members` an array, and `harvestedAt`/`complete` — if present at all — honestly typed. Doesn't
 *  validate each member's own shape (id/displayName): a chat member the roster couldn't resolve is
 *  already handled downstream in mentions.ts, and this cache's job ends at "did the JSON honestly
 *  describe an entry, not garbage". */
function isValidEntry(value: unknown): value is MembersCacheEntry {
  if (
    typeof value !== 'object' ||
    value === null ||
    !Array.isArray((value as { members?: unknown }).members) ||
    typeof (value as { fetchedAt?: unknown }).fetchedAt !== 'number'
  ) {
    return false;
  }
  const harvestedAt = (value as { harvestedAt?: unknown }).harvestedAt;
  const complete = (value as { complete?: unknown }).complete;
  return (
    (harvestedAt === undefined || typeof harvestedAt === 'number') &&
    (complete === undefined || typeof complete === 'boolean')
  );
}

export interface MembersCacheOptions {
  /** Absolute path to the cache file — same instance dir as the token cache, so one server
   *  install has one members cache, not one per whoever happens to construct this class. */
  path: string;
  ttlMs?: number;
  /** Injectable clock for tests; defaults to Date.now. */
  now?: () => number;
  /** Injectable write primitives — tests use these to prove the destination path is reached
   *  ONLY via rename, never a direct write (0.4.1 review round 1: an assertion that merely
   *  checked "no leftover tmp file" passed identically with atomicity deleted entirely). */
  writeFileFn?: typeof writeFileSync;
  renameFn?: typeof renameSync;
}

/**
 * Disk-persisted per-chat membership cache. Born from a live diagnosis (0.4.1): `GET
 * /chats/{id}/members` shares a Graph throttle budget with another daemon on the same first-party
 * client id, and a 429's Retry-After never actually clears under continuous consumption — mention
 * resolution was hostage to an endpoint this server does not even need most of the time, since
 * chat membership in a fixed allowlist of pilot chats effectively never changes.
 *
 * A miss (never cached, past its TTL, or a corrupt/invalid entry) reads identically to "no
 * entry" — the caller decides what a miss means (teams-chats.ts refreshes once and re-checks);
 * this class only ever answers "what do we have on disk, and is it still within TTL and honestly
 * shaped", never makes a network call itself.
 */
export class MembersCache {
  private readonly path: string;
  private readonly ttlMs: number;
  private readonly now: () => number;
  /** Undefined means "use writeAtomicCacheFile's own default" — this class no longer needs the
   *  real writeFileSync/renameSync values itself now that the write mechanism lives in
   *  atomic-cache-write.ts (review round 2). */
  private readonly writeFileFn: typeof writeFileSync | undefined;
  private readonly renameFn: typeof renameSync | undefined;

  constructor(options: MembersCacheOptions) {
    this.path = options.path;
    this.ttlMs = options.ttlMs ?? DEFAULT_MEMBERS_TTL_MS;
    this.now = options.now ?? Date.now;
    this.writeFileFn = options.writeFileFn;
    this.renameFn = options.renameFn;
  }

  /**
   * Fresh-enough read for MENTION resolution — the only caller with any business reading a
   * PARTIAL (traffic-harvested-only) roster (0.5.2 BLOCKER 1 fix, 2026-09-04 review): a name
   * absent from a partial roster still falls through to a live refresh (teams-chats.ts), so
   * serving one here costs nothing but a possibly-incomplete match set, never a wrong grant.
   * Freshness is judged against whichever evidence is more recent — a real `/members` fetch
   * (`fetchedAt`) OR the last traffic harvest (`harvestedAt`) — so a busy chat's partial roster
   * stays usable without a live call for as long as traffic keeps confirming it, same as a
   * COMPLETE roster always has. `membersForInvite`'s permission-grant path must NEVER call this —
   * see `getComplete` below, the one it uses instead.
   */
  get(chatId: string): ChatMember[] | undefined {
    const entry = this.readFile()[chatId];
    if (!entry) {
      return undefined;
    }
    const freshAt = Math.max(entry.fetchedAt, entry.harvestedAt ?? 0);
    if (this.now() - freshAt > this.ttlMs) {
      return undefined;
    }
    return entry.members;
  }

  /**
   * The COMPLETE-only twin of get() — `membersForInvite`'s (sendFile's permission-grant roster,
   * teams-chats.ts) fresh read (0.5.2 BLOCKER 1 fix, 2026-09-04 review). Refuses a PARTIAL
   * (traffic-harvested-only, `complete: false`) entry outright, returning undefined exactly as if
   * nothing were cached — the caller then has no choice but a real `/members` fetch, which is the
   * whole point: a chat member who has never spoken must still be found and granted access, and a
   * departed member who stops speaking must still fall out of the grant once this roster's OWN
   * `fetchedAt` (never extended by traffic — see merge()'s own doc comment) goes past the TTL.
   * Note: `merge()` into a COMPLETE entry adds a demonstrated speaker's id to `members` while
   * keeping `complete: true` — COMPLETE therefore means "at least everything `/members` returned"; it
   * can gain a speaker, never lose a silent member, so the no-dead-cards contract holds.
   * Freshness is judged against `fetchedAt` ONLY, deliberately ignoring `harvestedAt`: traffic
   * reconfirming who is still around is good enough evidence for a mention, but not for a file
   * permission grant, which re-verifies against the authoritative endpoint on its own schedule.
   */
  getComplete(chatId: string): ChatMember[] | undefined {
    const entry = this.readFile()[chatId];
    if (!entry || entry.complete === false) {
      return undefined;
    }
    if (this.now() - entry.fetchedAt > this.ttlMs) {
      return undefined;
    }
    return entry.members;
  }

  /**
   * The TTL-ignoring twin of get() — reads whatever is on disk for `chatId`, expired or not.
   * Mitigation 1 (docs/throttling-mitigation.md §4, stage 1 item 2, live-diagnosed 2026-09-04):
   * a `/members` refresh 429ing on an expired cache used to mean the cache could never be
   * rewritten, converting a working cached path into a PERMANENT hard dependency on the
   * throttled endpoint. This method is the caller's ONLY escape hatch, and it is deliberately
   * not what get() itself does — a normal cache hit must still honour the TTL; only a caller
   * that has ALREADY tried and failed a live refresh has any business reaching for stale data.
   * Serves ANY roster (complete or partial) — resolveMentions's stale-serve fallback may safely
   * use a partial roster on a throttled refresh, same reasoning as get() above.
   */
  getStale(chatId: string): { members: ChatMember[]; fetchedAt: number } | undefined {
    const entry = this.readFile()[chatId];
    return entry ? { members: entry.members, fetchedAt: entry.fetchedAt } : undefined;
  }

  /**
   * The COMPLETE-only twin of getStale() — `membersForInvite`'s stale-serve fallback on a
   * throttled/transient `/members` refresh (0.5.2 BLOCKER 1 fix, 2026-09-04 review). A PARTIAL
   * entry returns undefined here even though getStale() would happily serve it: a harvested-only
   * roster has never been confirmed complete, so there is nothing honest to "serve stale" — the
   * caller must fail the send instead (no partial grants, no dead cards).
   */
  getStaleComplete(chatId: string): { members: ChatMember[]; fetchedAt: number } | undefined {
    const entry = this.readFile()[chatId];
    if (!entry || entry.complete === false) {
      return undefined;
    }
    return { members: entry.members, fetchedAt: entry.fetchedAt };
  }

  /**
   * Merges harvested (id, displayName) pairs into `chatId`'s roster at zero Graph cost — the
   * inbox poller's own doc comment names the incident this exists for (mitigation 2,
   * docs/throttling-mitigation.md §4, stage 1 item 2): every polled message already carries its
   * sender's AAD id and display name, and that traffic is exactly the population worth
   * @mentioning. A harvested pair missing either field is dropped rather than persisted — this
   * is untrusted data reaching the cache directly, with no `/members` validation in between, and
   * a junk entry (an id with no name, or vice versa) would otherwise corrupt real mention
   * resolution later.
   *
   * PROVENANCE (0.5.2 BLOCKER 1 fix, 2026-09-04 review — a harvested roster used verbatim as
   * `sendFile`'s permission-grant list silently omitted real, silent chat members, and a departed
   * member's file grant never expired because every merge used to restamp `fetchedAt`, the SAME
   * field a real `/members` fetch stamps): a merge into an entry that is already COMPLETE (a real
   * `/members` fetch, or no `complete` flag at all — a legacy 0.5.1 entry, which can only ever
   * have come from `set()`) keeps it complete and — critically — keeps its ORIGINAL `fetchedAt`
   * untouched; traffic augments a complete roster's member list (a genuinely new member Graph
   * hasn't been asked about yet, say) but never re-certifies it as freshly fetched. A merge into
   * no entry, or an already-PARTIAL one, produces/extends a PARTIAL entry: `fetchedAt: 0` (the
   * 0.5.1-compatible "instantly expired" marker — see `MembersCacheEntry.fetchedAt`'s own doc
   * comment for why 0, not omitted), `complete: false`. Either way `harvestedAt` is stamped with
   * now() — that field alone is what "refreshes from traffic alone" now means, and get() (never
   * getComplete()) is the only reader that honours it. A merge with nothing valid to add never
   * touches the disk at all.
   */
  merge(chatId: string, harvested: ReadonlyArray<{ id: string; displayName: string }>): void {
    const valid = harvested.filter((entry) => entry.id.trim() !== '' && entry.displayName.trim() !== '');
    if (valid.length === 0) {
      return;
    }
    const file = this.readFile();
    const existing = file[chatId];
    const wasComplete = existing !== undefined && existing.complete !== false;
    const byId = new Map((existing?.members ?? []).flatMap((member) => (member.id ? [[member.id, member] as const] : [])));
    for (const { id, displayName } of valid) {
      byId.set(id, { id, displayName });
    }
    file[chatId] = {
      members: [...byId.values()],
      fetchedAt: wasComplete ? (existing as MembersCacheEntry).fetchedAt : 0,
      harvestedAt: this.now(),
      complete: wasComplete,
    };
    this.writeFile(file);
  }

  /**
   * Concurrent writers (two processes on the same instance dir, e.g. a stray second server —
   * see KNOWN-ISSUES.md's "two server instances race" entry) are NOT coordinated here: this is a
   * read-modify-write over the whole file with no lock, so two near-simultaneous set() calls for
   * DIFFERENT chats can race and one's update can be lost. Acceptable for this cache specifically
   * — a lost update degrades to "cache miss, refresh once" on the next resolveMentions call, not
   * a wrong answer — but real cross-process coordination (a lock file, as KNOWN-ISSUES already
   * names as the fix direction for the inbox) is not something this method attempts.
   */
  set(chatId: string, members: ChatMember[]): void {
    const file = this.readFile();
    // A real /members fetch is always fully authoritative — complete: true (and no harvestedAt:
    // any prior traffic evidence is superseded, not merged with, since this IS the ground truth
    // that evidence was only ever a substitute for). See merge()'s own doc comment for the
    // PARTIAL/COMPLETE distinction this entry now carries (0.5.2 BLOCKER 1 fix).
    file[chatId] = { members, fetchedAt: this.now(), complete: true };
    this.writeFile(file);
  }

  private readFile(): MembersCacheFile {
    try {
      const parsed: unknown = JSON.parse(readFileSync(this.path, 'utf8'));
      if (typeof parsed === 'object' && parsed !== null && !Array.isArray(parsed)) {
        const file: MembersCacheFile = {};
        for (const [chatId, entry] of Object.entries(parsed as Record<string, unknown>)) {
          if (isValidEntry(entry)) {
            file[chatId] = entry;
          }
          // An invalid entry for one chat is dropped silently and read as "never cached" for
          // that chat only — see isValidEntry's doc comment; a sibling chat's valid entry is
          // unaffected.
        }
        return file;
      }
    } catch {
      // Missing or corrupt cache reads the same as "never cached" — a refresh, not a crash.
    }
    return {};
  }

  /**
   * The write-to-temp-then-rename/0600 mechanism itself lives in atomic-cache-write.ts (review
   * round 2: FileSelfIdCache, self-id-cache.ts, used to carry its own copy of this exact code with
   * no test ever exercising its rename seam — sharing one implementation means the atomicity test
   * below covers both). This method's own job is only building the JSON payload.
   */
  private writeFile(file: MembersCacheFile): void {
    writeAtomicCacheFile({
      path: this.path,
      data: JSON.stringify(file, null, 2),
      now: this.now,
      ...(this.writeFileFn !== undefined ? { writeFileFn: this.writeFileFn } : {}),
      ...(this.renameFn !== undefined ? { renameFn: this.renameFn } : {}),
    });
  }
}
