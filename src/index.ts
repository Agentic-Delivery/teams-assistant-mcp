#!/usr/bin/env node
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { homedir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { buildChats } from './build-chats.js';
import { loadConfig } from './config.js';
import { InboxPoller, type SignedInAccount } from './inbox.js';
import { acquirePollerLock } from './poller-lock.js';
import { buildServer } from './server.js';

async function main(): Promise<void> {
  const config = loadConfig();

  const uploadDir = process.env['TEAMS_MCP_UPLOAD_DIR'];
  const { chats, graph } = buildChats(config, uploadDir ? { uploadDir } : {});

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
    `teams-assistant-mcp ready: ${config.allowlist.entries().length} allowed chat(s), auth=ropc\n`,
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

  // One poller per inbox, enforced: on 2026-09-01 two daemons for different projects raced on
  // one shared account and fed each other's throttle penalty. A live holder wins; this server
  // keeps serving its tools without polling, which is loudly said rather than silently done.
  // A dead holder's lock (the reboot case) is taken over.
  const lock = await acquirePollerLock({ lockPath: join(dirname(inboxPath), 'poller.lock') });
  if (!lock.acquired) {
    process.stderr.write(
      `inbox poller NOT started: pid ${lock.holderPid} already polls ${inboxPath} ` +
        `(holds ${join(dirname(inboxPath), 'poller.lock')}` +
        `${lock.holderStartedAt ? `, since ${lock.holderStartedAt}` : ''}); ` +
        'this server keeps serving MCP tools without polling\n',
    );
    return;
  }

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
