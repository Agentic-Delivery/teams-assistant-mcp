import { readFileSync, type renameSync, type writeFileSync } from 'node:fs';
import { writeAtomicCacheFile } from './atomic-cache-write.js';

/**
 * The signed-in account's own AAD id, plus when it was last confirmed against `/me`.
 * `resolvedAt` carries no TTL policy here — the account's own id does not change — it exists for
 * operator visibility only (when was this last confirmed live). See resolveSelfId in
 * teams-chats.ts for the full read order this cache participates in, including the
 * username-stamp check (FileSelfIdCacheOptions.expectedUsername, this module) and sendFile's
 * roster-membership re-check — two defences against a wrong cached id, and the one residual
 * limitation neither catches — that resolveSelfId's own doc comment names in full (review round
 * 2, out of scope for 0.5.1).
 */
export interface SelfIdCacheEntry {
  id: string;
  /** Epoch milliseconds. */
  resolvedAt: number;
}

export interface SelfIdCachePort {
  read(): SelfIdCacheEntry | undefined;
  write(entry: SelfIdCacheEntry): void;
}

/**
 * The safe default when a caller does not wire a real persisted cache: read() always misses,
 * write() does nothing. Unlike MembersCache (required on GraphTeamsChatsOptions since 0.4.1,
 * because an unwired members cache silently fell back to a THROTTLED live endpoint), omitting a
 * self id cache carries no such hazard — it degrades to exactly the pre-0.5.1 behaviour, an
 * in-memory-only memo gone the moment the process exits, never a wrong answer. So this one stays
 * optional; see GraphTeamsChatsOptions.selfIdCache's own doc comment.
 */
export class NullSelfIdCache implements SelfIdCachePort {
  read(): SelfIdCacheEntry | undefined {
    return undefined;
  }

  write(): void {
    // deliberately a no-op — see class doc comment
  }
}

export interface FileSelfIdCacheOptions {
  /** Absolute path — see config.ts's selfIdCachePath (derived next to the token cache, same
   *  reasoning as membersCachePath). */
  path: string;
  /**
   * The resolving account's username (config.username / TEAMS_MCP_USERNAME) — review round 2:
   * stamped into every written entry and checked on every read. A mismatch (including a legacy
   * entry with no username field at all) reads as a plain miss, same posture as every other
   * "garbage in, miss out" shape check this class already applies — never a throw. Closes the
   * realistic case of a consuming project repointing TEAMS_MCP_USERNAME at the same instance dir
   * (a different service account, same token-cache directory) and silently inheriting the OLD
   * account's AAD id as if it were the new one's. Optional and unchecked when omitted — every
   * production call site (build-chats.ts) always supplies it; tests that do not care about this
   * specific concern are free to omit it, same posture as selfIdCache's own optionality on
   * GraphTeamsChatsOptions.
   */
  expectedUsername?: string;
  /** Injectable clock for tests; defaults to Date.now. */
  now?: () => number;
  writeFileFn?: typeof writeFileSync;
  renameFn?: typeof renameSync;
}

/**
 * Persists the signed-in account's own AAD id between process runs. Born from a live incident
 * (2026-09-03): `resolveSelfId()`'s in-memory-only memo (teams-chats.ts) protects nothing across
 * process boundaries, and every standalone CLI invocation is a fresh process — eight consecutive
 * `teams-send-file` attempts over 20 minutes each paid, and each lost, the same throttled `/me`
 * call. The account's own id never changes, so once resolved it is worth keeping past the
 * process that resolved it.
 *
 * Modeled directly on MembersCache (members-cache.ts): 0600, write-to-temp-then-rename so a
 * crash mid-write never leaves a half-written file the next read chokes on, a corrupt or
 * garbage-shaped file degrades to "never cached" rather than throwing.
 */
export class FileSelfIdCache implements SelfIdCachePort {
  private readonly path: string;
  private readonly expectedUsername: string | undefined;
  private readonly now: () => number;
  /** Undefined means "use writeAtomicCacheFile's own default" — same reasoning as MembersCache's
   *  identical fields (members-cache.ts), review round 2. */
  private readonly writeFileFn: typeof writeFileSync | undefined;
  private readonly renameFn: typeof renameSync | undefined;

  constructor(options: FileSelfIdCacheOptions) {
    this.path = options.path;
    this.expectedUsername = options.expectedUsername;
    this.now = options.now ?? Date.now;
    this.writeFileFn = options.writeFileFn;
    this.renameFn = options.renameFn;
  }

  read(): SelfIdCacheEntry | undefined {
    let raw: string;
    try {
      raw = readFileSync(this.path, 'utf8');
    } catch {
      return undefined;
    }
    try {
      const parsed = JSON.parse(raw) as Partial<SelfIdCacheEntry & { username?: string }>;
      if (
        typeof parsed.id !== 'string' ||
        parsed.id.trim() === '' ||
        typeof parsed.resolvedAt !== 'number'
      ) {
        return undefined;
      }
      // A mismatch (or a legacy entry with no username field at all, once expectedUsername is
      // set) reads as a plain miss — see FileSelfIdCacheOptions.expectedUsername's own doc
      // comment for why. Skipped entirely when this instance was not given one to check against.
      if (this.expectedUsername !== undefined && parsed.username !== this.expectedUsername) {
        return undefined;
      }
      return { id: parsed.id, resolvedAt: parsed.resolvedAt };
    } catch {
      return undefined;
    }
  }

  /**
   * The write-to-temp-then-rename/0600 mechanism itself lives in atomic-cache-write.ts, shared
   * with MembersCache (review round 2) — see that module's doc comment and members-cache.test.ts's
   * atomicity test, which this class's write path now runs through unchanged rather than a second,
   * separately-untested copy of the same nine lines.
   */
  write(entry: SelfIdCacheEntry): void {
    const stamped = {
      id: entry.id,
      resolvedAt: entry.resolvedAt,
      ...(this.expectedUsername !== undefined ? { username: this.expectedUsername } : {}),
    };
    writeAtomicCacheFile({
      path: this.path,
      data: JSON.stringify(stamped, null, 2),
      now: this.now,
      ...(this.writeFileFn !== undefined ? { writeFileFn: this.writeFileFn } : {}),
      ...(this.renameFn !== undefined ? { renameFn: this.renameFn } : {}),
    });
  }
}
