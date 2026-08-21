import { chmodSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname } from 'node:path';

export interface CachedToken {
  accessToken: string;
  /** Epoch milliseconds. */
  expiresAt: number;
  refreshToken?: string;
}

export interface TokenCache {
  read(): CachedToken | undefined;
  write(token: CachedToken): void;
}

/** In-memory only. Used by tests and by anyone who does not want a token on disk. */
export class MemoryTokenCache implements TokenCache {
  private current: CachedToken | undefined;

  read(): CachedToken | undefined {
    return this.current;
  }

  write(token: CachedToken): void {
    this.current = token;
  }
}

/**
 * Persists the token between server restarts so a restart does not mean another password grant.
 * The file holds a live refresh token, so it is written 0600 and its name is gitignored.
 */
export class FileTokenCache implements TokenCache {
  constructor(private readonly path: string) {}

  read(): CachedToken | undefined {
    let raw: string;
    try {
      raw = readFileSync(this.path, 'utf8');
    } catch {
      return undefined;
    }
    try {
      const parsed = JSON.parse(raw) as Partial<CachedToken>;
      if (typeof parsed.accessToken !== 'string' || typeof parsed.expiresAt !== 'number') {
        return undefined;
      }
      return {
        accessToken: parsed.accessToken,
        expiresAt: parsed.expiresAt,
        ...(typeof parsed.refreshToken === 'string' ? { refreshToken: parsed.refreshToken } : {}),
      };
    } catch {
      return undefined;
    }
  }

  write(token: CachedToken): void {
    mkdirSync(dirname(this.path), { recursive: true });
    writeFileSync(this.path, JSON.stringify(token, null, 2), { encoding: 'utf8', mode: 0o600 });
    // writeFileSync's mode only applies on create; an existing file keeps its old permissions.
    chmodSync(this.path, 0o600);
  }
}
