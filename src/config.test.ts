import { describe, expect, it } from 'vitest';
import { mkdtempSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { DEFAULT_ASSISTANT_DISPLAY_NAME, loadConfig, parseAllowlistFile } from './config.js';
import { FileTokenCache } from './auth/token-cache.js';

function withConfigFile(contents: unknown): string {
  const dir = mkdtempSync(join(tmpdir(), 'teams-mcp-config-'));
  const path = join(dir, 'teams-mcp.config.json');
  writeFileSync(path, JSON.stringify(contents));
  return path;
}

const CREDS = {
  TEAMS_MCP_TENANT_ID: '11111111-2222-3333-4444-555555555555',
  TEAMS_MCP_USERNAME: 'assistant@example.com',
  TEAMS_MCP_PASSWORD: 'secret',
};

describe('allowlist file parsing', () => {
  it('reads chat ids, labels and post permission', () => {
    const file = parseAllowlistFile(
      JSON.stringify({
        assistantDisplayName: 'Assistant (AI)',
        allowedChats: [{ id: '19:a@thread.v2', label: 'Pilot', canPost: true }],
      }),
      'test',
    );

    expect(file.assistantDisplayName).toBe('Assistant (AI)');
    expect(file.allowedChats).toEqual([{ id: '19:a@thread.v2', label: 'Pilot', canPost: true }]);
  });

  it('defaults a chat to read-only when canPost is not spelled out', () => {
    const file = parseAllowlistFile(
      JSON.stringify({ allowedChats: [{ id: '19:a@thread.v2' }] }),
      'test',
    );

    expect(file.allowedChats?.[0]).toEqual({
      id: '19:a@thread.v2',
      label: '19:a@thread.v2',
      canPost: false,
    });
  });

  it('rejects a malformed file rather than starting with a half-read allowlist', () => {
    expect(() => parseAllowlistFile('{ nope', 'cfg')).toThrow(/not valid JSON/);
    expect(() => parseAllowlistFile('[]', 'cfg')).toThrow(/JSON object/);
    expect(() => parseAllowlistFile('{"allowedChats": {}}', 'cfg')).toThrow(/must be an array/);
    expect(() => parseAllowlistFile('{"allowedChats": [{"label": "x"}]}', 'cfg')).toThrow(
      /non-empty string/,
    );
  });
});

describe('config loading', () => {
  it('builds a usable config from env plus the allowlist file', () => {
    const path = withConfigFile({
      assistantDisplayName: 'Assistant (AI)',
      allowedChats: [{ id: '19:a@thread.v2', label: 'Pilot', canPost: true }],
    });

    const config = loadConfig({ ...CREDS, TEAMS_MCP_CONFIG: path });

    expect(config.username).toBe('assistant@example.com');
    expect(config.allowlist.isAllowed('19:a@thread.v2')).toBe(true);
    expect(config.assistantDisplayName).toBe('Assistant (AI)');
    // The first-party Teams client id is the default; no app registration is involved.
    expect(config.clientId).toBe('1fec8e78-bce4-4aaf-ab1b-5451cc387264');
  });

  it('falls back to the default display name when the file does not set one', () => {
    const path = withConfigFile({ allowedChats: [{ id: '19:a@thread.v2' }] });

    const config = loadConfig({ ...CREDS, TEAMS_MCP_CONFIG: path });

    expect(config.assistantDisplayName).toBe(DEFAULT_ASSISTANT_DISPLAY_NAME);
  });

  it('refuses to start when a credential is missing', () => {
    const path = withConfigFile({ allowedChats: [{ id: '19:a@thread.v2' }] });

    expect(() => loadConfig({ TEAMS_MCP_TENANT_ID: 't', TEAMS_MCP_CONFIG: path })).toThrow(
      /TEAMS_MCP_USERNAME/,
    );
  });

  it('refuses to start when the allowlist is empty', () => {
    const path = withConfigFile({ allowedChats: [] });

    expect(() => loadConfig({ ...CREDS, TEAMS_MCP_CONFIG: path })).toThrow(/empty/i);
  });

  it('says what to do when the config file is missing entirely', () => {
    expect(() =>
      loadConfig({ ...CREDS, TEAMS_MCP_CONFIG: join(tmpdir(), 'nope-does-not-exist.json') }),
    ).toThrow(/Allowlist config not found/);
  });
});

describe('member cache config (0.4.1)', () => {
  it('places the members cache next to the token cache, defaulting the TTL to 24 hours', () => {
    const path = withConfigFile({ allowedChats: [{ id: '19:a@thread.v2' }] });
    const tokenCache = join(tmpdir(), 'some-instance-dir', '.token-cache.json');

    const config = loadConfig({ ...CREDS, TEAMS_MCP_CONFIG: path, TEAMS_MCP_TOKEN_CACHE: tokenCache });

    expect(config.membersCachePath).toBe(join(tmpdir(), 'some-instance-dir', '.members-cache.json'));
    expect(config.membersTtlMs).toBe(24 * 60 * 60 * 1000);
  });

  it('TEAMS_MCP_MEMBERS_TTL_SECONDS overrides the default TTL', () => {
    const path = withConfigFile({ allowedChats: [{ id: '19:a@thread.v2' }] });

    const config = loadConfig({ ...CREDS, TEAMS_MCP_CONFIG: path, TEAMS_MCP_MEMBERS_TTL_SECONDS: '60' });

    expect(config.membersTtlMs).toBe(60_000);
  });

  it('ignores a non-numeric or non-positive override rather than crashing the whole server on it', () => {
    const path = withConfigFile({ allowedChats: [{ id: '19:a@thread.v2' }] });

    const config = loadConfig({ ...CREDS, TEAMS_MCP_CONFIG: path, TEAMS_MCP_MEMBERS_TTL_SECONDS: 'nope' });

    expect(config.membersTtlMs).toBe(24 * 60 * 60 * 1000);
  });
});

// 0.5.1 — live-diagnosed 2026-09-03: eight consecutive teams-send-file CLI attempts over 20
// minutes each failed the /me lookup, because every CLI invocation is a fresh process. The
// account's own id never changes, so it is now persisted next to the token cache, same
// reasoning and same derivation as the members cache above.
describe('self id cache config (0.5.1)', () => {
  it('places the self id cache next to the token cache', () => {
    const path = withConfigFile({ allowedChats: [{ id: '19:a@thread.v2' }] });
    const tokenCache = join(tmpdir(), 'some-instance-dir', '.token-cache.json');

    const config = loadConfig({ ...CREDS, TEAMS_MCP_CONFIG: path, TEAMS_MCP_TOKEN_CACHE: tokenCache });

    expect(config.selfIdCachePath).toBe(join(tmpdir(), 'some-instance-dir', '.self-id-cache.json'));
  });

  it('has no selfIdOverride when TEAMS_MCP_SELF_ID is unset', () => {
    const path = withConfigFile({ allowedChats: [{ id: '19:a@thread.v2' }] });

    const config = loadConfig({ ...CREDS, TEAMS_MCP_CONFIG: path });

    expect(config.selfIdOverride).toBeUndefined();
  });

  it('reads a GUID-shaped TEAMS_MCP_SELF_ID as selfIdOverride', () => {
    const path = withConfigFile({ allowedChats: [{ id: '19:a@thread.v2' }] });

    const config = loadConfig({
      ...CREDS,
      TEAMS_MCP_CONFIG: path,
      TEAMS_MCP_SELF_ID: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
    });

    expect(config.selfIdOverride).toBe('aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee');
  });

  it('ignores a TEAMS_MCP_SELF_ID that is not GUID-shaped rather than crashing the whole server on a typo', () => {
    const path = withConfigFile({ allowedChats: [{ id: '19:a@thread.v2' }] });

    const config = loadConfig({
      ...CREDS,
      TEAMS_MCP_CONFIG: path,
      TEAMS_MCP_SELF_ID: 'not-a-guid-at-all',
    });

    expect(config.selfIdOverride).toBeUndefined();
  });
});

describe('token cache on disk', () => {
  it('writes the cache readable only by the owner', () => {
    const path = join(mkdtempSync(join(tmpdir(), 'teams-mcp-cache-')), 'nested', '.token-cache.json');
    const cache = new FileTokenCache(path);

    cache.write({ accessToken: 'tok', expiresAt: 42, refreshToken: 'rt' });

    expect(statSync(path).mode & 0o777).toBe(0o600);
    expect(cache.read()).toEqual({ accessToken: 'tok', expiresAt: 42, refreshToken: 'rt' });
  });

  it('tightens permissions on a file that already existed too open', () => {
    const dir = mkdtempSync(join(tmpdir(), 'teams-mcp-cache-'));
    const path = join(dir, '.token-cache.json');
    writeFileSync(path, '{}', { mode: 0o644 });

    new FileTokenCache(path).write({ accessToken: 'tok', expiresAt: 42 });

    expect(statSync(path).mode & 0o777).toBe(0o600);
  });

  it('treats a missing or corrupt cache as no cache', () => {
    const dir = mkdtempSync(join(tmpdir(), 'teams-mcp-cache-'));

    expect(new FileTokenCache(join(dir, 'absent.json')).read()).toBeUndefined();

    const corrupt = join(dir, 'corrupt.json');
    writeFileSync(corrupt, 'not json');
    expect(new FileTokenCache(corrupt).read()).toBeUndefined();
  });
});
