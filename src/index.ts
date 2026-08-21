#!/usr/bin/env node
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { homedir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { FileTokenCache } from './auth/token-cache.js';
import { RopcTokenProvider } from './auth/ropc-token-provider.js';
import type { TokenProvider } from './auth/token-provider.js';
import { GraphClient } from './graph/graph-client.js';
import { GraphTeamsChats } from './graph/teams-chats.js';
import { loadConfig } from './config.js';
import { InboxPoller, type SignedInAccount } from './inbox.js';
import { buildServer } from './server.js';

async function main(): Promise<void> {
  const config = loadConfig();

  // The one line to change when ROPC dies. Everything downstream takes a TokenProvider.
  const tokenProvider: TokenProvider = new RopcTokenProvider({
    tenantId: config.tenantId,
    clientId: config.clientId,
    username: config.username,
    password: config.password,
    cache: new FileTokenCache(config.tokenCachePath),
  });

  const uploadDir = process.env['TEAMS_MCP_UPLOAD_DIR'];
  const graph = new GraphClient({ tokenProvider });
  const chats = new GraphTeamsChats(graph, uploadDir ? { uploadDir } : {});

  const server = buildServer({
    chats,
    allowlist: config.allowlist,
    assistantDisplayName: config.assistantDisplayName,
    ...(process.env['TEAMS_MCP_DOWNLOAD_DIR']
      ? { downloadDir: process.env['TEAMS_MCP_DOWNLOAD_DIR'] }
      : {}),
  });

  await server.connect(new StdioServerTransport());
  // stdout is the MCP transport. Anything human-readable goes to stderr or it corrupts the protocol.
  process.stderr.write(
    `teams-assistant-mcp ready: ${config.allowlist.entries().length} allowed chat(s), auth=${tokenProvider.kind}\n`,
  );

  // The background inbox: every new message in the allowlisted chats lands as one JSON line in a
  // session-independent file, so an orchestrator watches the file instead of running its own
  // polling daemon. Set TEAMS_INBOX_DISABLED=1 to skip it; consumers that only post pay nothing.
  const disabled = process.env['TEAMS_INBOX_DISABLED']?.trim();
  if (disabled !== undefined && disabled !== '' && disabled !== '0') {
    process.stderr.write('inbox poller off (TEAMS_INBOX_DISABLED)\n');
    return;
  }
  const inboxPath = resolve(
    process.env['TEAMS_INBOX_PATH']?.trim() || join(homedir(), '.teams-assistant', 'inbox.jsonl'),
  );
  const poller = new InboxPoller({
    chats,
    allowlist: config.allowlist,
    self: () => graph.get<SignedInAccount>('/me?$select=id,displayName'),
    inboxPath,
    // The state sidecar follows the inbox file, so a TEAMS_INBOX_PATH override moves both.
    statePath: join(dirname(inboxPath), 'inbox-state.json'),
    log: (line) => process.stderr.write(`${line}\n`),
  });
  poller.start();
  process.stderr.write(`inbox poller on: ${inboxPath}\n`);
}

main().catch((error: unknown) => {
  process.stderr.write(`teams-assistant-mcp failed to start: ${(error as Error).message}\n`);
  process.exit(1);
});
