import { readFileSync, type renameSync, type writeFileSync } from 'node:fs';
import { writeAtomicCacheFile } from './atomic-cache-write.js';
import type { ChatMember } from './mentions.js';

/** TEAMS_MCP_MEMBERS_TTL_SECONDS overrides this. Chat membership in these allowlisted chats
 *  effectively never changes, so a day-long TTL trades staleness risk (a departed member still
 *  mentionable for up to this long) for taking the throttled `/members` endpoint off the send
 *  path entirely — see resolveMentions in teams-chats.ts for the refresh-on-miss fallback that
 *  bounds that risk to "one bad mention attempt", not "permanently wrong". THE single source of
 *  truth for the 24h default: config.ts derives its own (seconds-flavoured) default from this,
 *  rather than hardcoding a second "24h" that could drift from this one (0.4.1 review round 1). */
export const DEFAULT_MEMBERS_TTL_MS = 24 * 60 * 60 * 1000;

interface MembersCacheEntry {
  members: ChatMember[];
  fetchedAt: number;
}

type MembersCacheFile = Record<string, MembersCacheEntry>;

/** True only for a shape `get()`/callers can trust: `fetchedAt` a number (a non-number would make
 *  `now() - fetchedAt` a NaN comparison, which silently reads as "never expires" — 0.4.1 review),
 *  `members` an array. Doesn't validate each member's own shape (id/displayName): a chat member
 *  the roster couldn't resolve is already handled downstream in mentions.ts, and this cache's job
 *  ends at "did the JSON honestly describe an entry, not garbage". */
function isValidEntry(value: unknown): value is MembersCacheEntry {
  return (
    typeof value === 'object' &&
    value !== null &&
    Array.isArray((value as { members?: unknown }).members) &&
    typeof (value as { fetchedAt?: unknown }).fetchedAt === 'number'
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

  get(chatId: string): ChatMember[] | undefined {
    const entry = this.readFile()[chatId];
    if (!entry) {
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
   */
  getStale(chatId: string): { members: ChatMember[]; fetchedAt: number } | undefined {
    const entry = this.readFile()[chatId];
    return entry ? { members: entry.members, fetchedAt: entry.fetchedAt } : undefined;
  }

  /**
   * Merges harvested (id, displayName) pairs into `chatId`'s roster at zero Graph cost — the
   * inbox poller's own doc comment names the incident this exists for (mitigation 2,
   * docs/throttling-mitigation.md §4, stage 1 item 2): every polled message already carries its
   * sender's AAD id and display name, and that traffic is exactly the population worth
   * @mentioning. A harvested pair missing either field is dropped rather than persisted — this
   * is untrusted data reaching the cache directly, with no `/members` validation in between, and
   * a junk entry (an id with no name, or vice versa) would otherwise corrupt real mention
   * resolution later. `fetchedAt` is stamped with now() on every merge that persists anything, not
   * just a brand-new sender: traffic from an ALREADY-known member is still live evidence the
   * roster is current, which is the whole point of "refreshes from traffic alone" — a chat that
   * stays busy never needs a live `/members` call again. A merge with nothing valid to add never
   * touches the disk at all.
   */
  merge(chatId: string, harvested: ReadonlyArray<{ id: string; displayName: string }>): void {
    const valid = harvested.filter((entry) => entry.id.trim() !== '' && entry.displayName.trim() !== '');
    if (valid.length === 0) {
      return;
    }
    const file = this.readFile();
    const existing = file[chatId];
    const byId = new Map((existing?.members ?? []).flatMap((member) => (member.id ? [[member.id, member] as const] : [])));
    for (const { id, displayName } of valid) {
      byId.set(id, { id, displayName });
    }
    file[chatId] = { members: [...byId.values()], fetchedAt: this.now() };
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
    file[chatId] = { members, fetchedAt: this.now() };
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
