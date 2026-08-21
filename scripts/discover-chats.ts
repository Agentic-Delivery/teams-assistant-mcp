/**
 * Lists every chat the assistant account can see, with the id to paste into
 * teams-mcp.config.json. A human runs this (`npm run discover-chats`); the MCP tools themselves
 * refuse to look outside the allowlist, so discovery has to live in a separate script.
 *
 * Credentials come from .env in the repo root. Variables already exported in the shell win over
 * the file, same as everywhere else.
 */
import { existsSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseEnv } from 'node:util';
import { MemoryTokenCache } from '../src/auth/token-cache.js';
import { RopcTokenProvider } from '../src/auth/ropc-token-provider.js';
import { GraphClient, GraphError } from '../src/graph/graph-client.js';
import { GraphTeamsChats } from '../src/graph/teams-chats.js';
import { TEAMS_FIRST_PARTY_CLIENT_ID } from '../src/config.js';
import { formatChatList } from '../src/discover.js';

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..');

function loadDotEnv(): void {
  const path = join(repoRoot, '.env');
  if (!existsSync(path)) {
    return;
  }
  for (const [key, value] of Object.entries(parseEnv(readFileSync(path, 'utf8')))) {
    process.env[key] ??= value;
  }
}

function env(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) {
    throw new Error(`Set ${name} in ${join(repoRoot, '.env')} (start from env.example).`);
  }
  return value;
}

async function main(): Promise<void> {
  loadDotEnv();

  const username = env('TEAMS_MCP_USERNAME');
  const tokenProvider = new RopcTokenProvider({
    tenantId: env('TEAMS_MCP_TENANT_ID'),
    clientId: process.env['TEAMS_MCP_CLIENT_ID']?.trim() || TEAMS_FIRST_PARTY_CLIENT_ID,
    username,
    password: env('TEAMS_MCP_PASSWORD'),
    // In memory on purpose: a discovery run should not leave a refresh token on disk.
    cache: new MemoryTokenCache(),
  });

  const chats = await new GraphTeamsChats(new GraphClient({ tokenProvider })).listChats();
  console.log(`Signed in as ${username}. ${chats.length} chat(s) visible.\n`);
  console.log(formatChatList(chats));
  if (chats.length > 0) {
    console.log('\nCopy the ids you want into teams-mcp.config.json (allowedChats).');
  }
}

main().catch((error: unknown) => {
  if (error instanceof GraphError && error.isLicenceProblem) {
    console.error(
      `Graph answered ${error.status}: ${error.message}\n` +
        'That is the no-Teams-licence failure. Auth worked; ask the licensing owner to assign ' +
        'an Office 365 licence with the Teams service plan to the account.',
    );
  } else {
    console.error(`discover-chats failed: ${(error as Error).message}`);
  }
  process.exit(1);
});
