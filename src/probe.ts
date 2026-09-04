/**
 * Read-only auth probe. Run it by hand (`npm run probe`) to answer one question: does this
 * account get a Graph token, and how far does that token actually get?
 *
 * It never posts anything. It does not need the allowlist config, because its second job is
 * discovering the chat ids you then put in that config — the MCP tools deliberately cannot list
 * chats outside the allowlist, so discovery lives here, in a thing a human runs.
 */
import { MemoryTokenCache } from './auth/token-cache.js';
import { RopcTokenProvider } from './auth/ropc-token-provider.js';
import { GraphClient, GraphError } from './graph/graph-client.js';
import { GraphTeamsChats } from './graph/teams-chats.js';
import { DEFAULT_ASSISTANT_DISPLAY_NAME, TEAMS_FIRST_PARTY_CLIENT_ID } from './config.js';

function claims(token: string): Record<string, unknown> {
  const payload = token.split('.')[1];
  if (!payload) {
    return {};
  }
  return JSON.parse(Buffer.from(payload, 'base64url').toString('utf8')) as Record<string, unknown>;
}

function env(name: string): string {
  const value = process.env[name];
  if (value && value.trim() !== '') {
    return value.trim();
  }
  throw new Error(`Set ${name} before running the probe.`);
}

async function main(): Promise<void> {
  const tokenProvider = new RopcTokenProvider({
    tenantId: env('TEAMS_MCP_TENANT_ID'),
    clientId: process.env['TEAMS_MCP_CLIENT_ID']?.trim() || TEAMS_FIRST_PARTY_CLIENT_ID,
    username: env('TEAMS_MCP_USERNAME'),
    password: env('TEAMS_MCP_PASSWORD'),
    // Deliberately in memory: a probe should not leave a refresh token behind on disk.
    cache: new MemoryTokenCache(),
  });

  const token = await tokenProvider.getAccessToken();
  const parsed = claims(token);
  const scopes = String(parsed['scp'] ?? '').split(' ').filter(Boolean);

  console.log('token acquired');
  console.log(`  aud    ${String(parsed['aud'] ?? '?')}`);
  console.log(`  tid    ${String(parsed['tid'] ?? '?')}`);
  console.log(`  upn    ${String(parsed['upn'] ?? parsed['unique_name'] ?? '?')}`);
  console.log(`  scopes ${scopes.length} (${scopes.slice(0, 8).join(', ')}${scopes.length > 8 ? ', …' : ''})`);

  const graph = new GraphClient({ tokenProvider });

  try {
    const me = await graph.get<{ displayName?: string; userPrincipalName?: string; id?: string }>(
      '/me?$select=id,displayName,userPrincipalName',
    );
    console.log('GET /me ok');
    console.log(`  displayName ${me.displayName ?? '?'}`);
    console.log(`  upn         ${me.userPrincipalName ?? '?'}`);
    const expected = process.env['TEAMS_MCP_DISPLAY_NAME'] ?? DEFAULT_ASSISTANT_DISPLAY_NAME;
    if (me.displayName !== expected) {
      console.log(
        `  WARNING display name is "${me.displayName ?? ''}", expected "${expected}". ` +
          'People in the chat must be able to see they are talking to an AI assistant.',
      );
    }
  } catch (error) {
    console.log(`GET /me failed: ${(error as Error).message}`);
  }

  try {
    // The probe never resolves mentions, so a real disk-backed cache would be pure overhead —
    // membersCache is a required GraphTeamsChats dependency, so a trivial no-op stands in.
    const chats = await new GraphTeamsChats(graph, {
      membersCache: { get: () => undefined, set: () => {}, getStale: () => undefined },
    }).listChats();
    console.log(`GET /me/chats ok — ${chats.length} chat(s)`);
    for (const chat of chats) {
      console.log(`  ${chat.id}  [${chat.chatType}]  ${chat.topic}`);
    }
    console.log('Copy the ids you want into teams-mcp.config.json (allowedChats).');
  } catch (error) {
    if (error instanceof GraphError && error.isLicenceProblem) {
      console.log(
        `GET /me/chats -> ${error.status} ${error.code ?? ''}: ${error.message}\n` +
          '  Expected for an account with no Teams-enabled Office 365 licence. Auth is fine; ' +
          'the licence is the blocker. Ask IT to assign one to the account.',
      );
      return;
    }
    console.log(`GET /me/chats failed: ${(error as Error).message}`);
  }
}

main().catch((error: unknown) => {
  console.error(`probe failed: ${(error as Error).message}`);
  process.exit(1);
});
