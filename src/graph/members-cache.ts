import { mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs';
import { dirname } from 'node:path';
import type { ChatMember } from './mentions.js';

/** TEAMS_MCP_MEMBERS_TTL_SECONDS overrides this. Chat membership in these allowlisted chats
 *  effectively never changes, so a day-long TTL trades staleness risk (a departed member still
 *  mentionable for up to this long) for taking the throttled `/members` endpoint off the send
 *  path entirely — see resolveMentions in teams-chats.ts for the refresh-on-miss fallback that
 *  bounds that risk to "one bad mention attempt", not "permanently wrong". */
export const DEFAULT_MEMBERS_TTL_MS = 24 * 60 * 60 * 1000;

interface MembersCacheEntry {
  members: ChatMember[];
  fetchedAt: number;
}

type MembersCacheFile = Record<string, MembersCacheEntry>;

export interface MembersCacheOptions {
  /** Absolute path to the cache file — same instance dir as the token cache, so one server
   *  install has one members cache, not one per whoever happens to construct this class. */
  path: string;
  ttlMs?: number;
  /** Injectable clock for tests; defaults to Date.now. */
  now?: () => number;
}

/**
 * Disk-persisted per-chat membership cache. Born from a live diagnosis (0.4.1): `GET
 * /chats/{id}/members` shares a Graph throttle budget with another daemon on the same first-party
 * client id, and a 429's Retry-After never actually clears under continuous consumption — mention
 * resolution was hostage to an endpoint this server does not even need most of the time, since
 * chat membership in a fixed allowlist of pilot chats effectively never changes.
 *
 * A miss (never cached, or past its TTL) reads identically to "no entry" — the caller decides
 * what a miss means (teams-chats.ts refreshes once and re-checks); this class only ever answers
 * "what do we have on disk, and is it still within TTL", never makes a network call itself.
 */
export class MembersCache {
  private readonly path: string;
  private readonly ttlMs: number;
  private readonly now: () => number;

  constructor(options: MembersCacheOptions) {
    this.path = options.path;
    this.ttlMs = options.ttlMs ?? DEFAULT_MEMBERS_TTL_MS;
    this.now = options.now ?? Date.now;
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

  set(chatId: string, members: ChatMember[]): void {
    const file = this.readFile();
    file[chatId] = { members, fetchedAt: this.now() };
    this.writeFile(file);
  }

  private readFile(): MembersCacheFile {
    try {
      const parsed: unknown = JSON.parse(readFileSync(this.path, 'utf8'));
      if (typeof parsed === 'object' && parsed !== null && !Array.isArray(parsed)) {
        return parsed as MembersCacheFile;
      }
    } catch {
      // Missing or corrupt cache reads the same as "never cached" — a refresh, not a crash.
    }
    return {};
  }

  /** Write-to-temp-then-rename: a crash mid-write must never leave a half-written cache file that
   *  the next read chokes on — same failure mode the watermark sidecar's plain writeFile had. */
  private writeFile(file: MembersCacheFile): void {
    mkdirSync(dirname(this.path), { recursive: true });
    const tmpPath = `${this.path}.tmp-${process.pid}-${this.now()}`;
    writeFileSync(tmpPath, JSON.stringify(file, null, 2), 'utf8');
    renameSync(tmpPath, this.path);
  }
}
