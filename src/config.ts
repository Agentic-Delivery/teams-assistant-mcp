import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { type AllowedChat, ChatAllowlist } from './allowlist.js';

export interface TeamsMcpConfig {
  tenantId: string;
  clientId: string;
  username: string;
  password: string;
  tokenCachePath: string;
  /**
   * What the service account's Teams profile is expected to be called. The server compares this
   * against the signed-in account's displayName at startup and warns on a mismatch; it does not
   * refuse to run, because renaming an account is IT's job, not the server's.
   */
  assistantDisplayName: string;
  allowlist: ChatAllowlist;
  /** Members cache lives next to the token cache — one instance dir, one cache, same reasoning
   *  as the inbox watermark sidecar living next to the inbox. */
  membersCachePath: string;
  /** TEAMS_MCP_MEMBERS_TTL_SECONDS overrides the 24h default (see members-cache.ts). */
  membersTtlMs: number;
}

export const DEFAULT_MEMBERS_TTL_SECONDS = 24 * 60 * 60;

export interface AllowlistFile {
  assistantDisplayName?: string;
  allowedChats?: AllowedChat[];
}

/**
 * Microsoft's first-party Microsoft Teams client id. Using it is what lets this server sign in
 * with no app registration and no admin consent: the tenant already trusts Microsoft's own
 * client. It is a published public identifier, not a secret.
 */
export const TEAMS_FIRST_PARTY_CLIENT_ID = '1fec8e78-bce4-4aaf-ab1b-5451cc387264';

export const DEFAULT_ASSISTANT_DISPLAY_NAME = 'Assistant (AI)';

function pick(env: NodeJS.ProcessEnv, names: readonly string[]): string | undefined {
  for (const name of names) {
    const value = env[name];
    if (value !== undefined && value.trim() !== '') {
      return value.trim();
    }
  }
  return undefined;
}

function requireEnv(env: NodeJS.ProcessEnv, names: readonly string[]): string {
  const value = pick(env, names);
  if (value === undefined) {
    throw new Error(
      `Missing required environment variable ${names[0]}` +
        (names.length > 1 ? ` (or ${names.slice(1).join(', ')})` : '') +
        '. Credentials are never read from anywhere but the environment.',
    );
  }
  return value;
}

export function parseAllowlistFile(raw: string, source: string): AllowlistFile {
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch (error) {
    throw new Error(`${source} is not valid JSON: ${(error as Error).message}`);
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error(`${source} must contain a JSON object.`);
  }
  const file = parsed as Record<string, unknown>;
  const chats = file['allowedChats'];
  if (chats !== undefined && !Array.isArray(chats)) {
    throw new Error(`${source}: allowedChats must be an array.`);
  }
  const allowedChats = ((chats ?? []) as unknown[]).map((entry, index) => {
    if (typeof entry !== 'object' || entry === null) {
      throw new Error(`${source}: allowedChats[${index}] must be an object.`);
    }
    const record = entry as Record<string, unknown>;
    if (typeof record['id'] !== 'string' || record['id'].trim() === '') {
      throw new Error(`${source}: allowedChats[${index}].id must be a non-empty string.`);
    }
    // canPost defaults to false. A chat added to the list without saying so gets read access
    // only, which is the safer way round to be wrong.
    return {
      id: record['id'],
      label: typeof record['label'] === 'string' ? record['label'] : record['id'],
      canPost: record['canPost'] === true,
    } satisfies AllowedChat;
  });

  const displayName = file['assistantDisplayName'];
  return {
    allowedChats,
    ...(typeof displayName === 'string' ? { assistantDisplayName: displayName } : {}),
  };
}

/** A malformed or non-positive override must never crash the whole server over a typo in one
 *  optional env var — it falls back to the default instead, same posture as every other
 *  best-effort env knob in this file. */
function membersTtlSecondsFrom(env: NodeJS.ProcessEnv): number {
  const raw = pick(env, ['TEAMS_MCP_MEMBERS_TTL_SECONDS']);
  if (raw === undefined) {
    return DEFAULT_MEMBERS_TTL_SECONDS;
  }
  const parsed = Number(raw);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : DEFAULT_MEMBERS_TTL_SECONDS;
}

export function loadConfig(env: NodeJS.ProcessEnv = process.env): TeamsMcpConfig {
  const configPath = resolve(pick(env, ['TEAMS_MCP_CONFIG']) ?? 'teams-mcp.config.json');

  let file: AllowlistFile;
  try {
    file = parseAllowlistFile(readFileSync(configPath, 'utf8'), configPath);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') {
      throw new Error(
        `Allowlist config not found at ${configPath}. Copy teams-mcp.config.example.json and ` +
          'set TEAMS_MCP_CONFIG to its path. The server does not run without an allowlist.',
      );
    }
    throw error;
  }

  const tokenCachePath = resolve(pick(env, ['TEAMS_MCP_TOKEN_CACHE']) ?? '.token-cache.json');

  return {
    tenantId: requireEnv(env, ['TEAMS_MCP_TENANT_ID']),
    clientId: pick(env, ['TEAMS_MCP_CLIENT_ID']) ?? TEAMS_FIRST_PARTY_CLIENT_ID,
    username: requireEnv(env, ['TEAMS_MCP_USERNAME']),
    password: requireEnv(env, ['TEAMS_MCP_PASSWORD']),
    tokenCachePath,
    assistantDisplayName: file.assistantDisplayName ?? DEFAULT_ASSISTANT_DISPLAY_NAME,
    allowlist: new ChatAllowlist(file.allowedChats ?? []),
    membersCachePath: join(dirname(tokenCachePath), 'members-cache.json'),
    membersTtlMs: membersTtlSecondsFrom(env) * 1000,
  };
}
