#!/usr/bin/env node
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { buildChats } from './build-chats.js';
import { buildInboxPoller } from './build-inbox-poller.js';
import { loadConfig } from './config.js';
import { inboxPathFor, inboxYieldPathFor } from './inbox-yield.js';
import { buildServer } from './server.js';

async function main(): Promise<void> {
  const config = loadConfig();

  const uploadDir = process.env['TEAMS_MCP_UPLOAD_DIR'];
  const { chats, graph, tokenProvider, membersCache } = buildChats(config, uploadDir ? { uploadDir } : {});

  // The yield file is passed even when THIS process runs no poller: the download tools write it
  // to quiet whichever poller shares the mailbox — possibly a daemon in another process — and a
  // yield file nobody reads costs nothing. See inbox-yield.ts for the measured starvation story.
  const inboxYieldPath = inboxYieldPathFor(process.env);

  const server = buildServer({
    chats,
    allowlist: config.allowlist,
    assistantDisplayName: config.assistantDisplayName,
    inboxYieldPath,
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
  const inboxPath = inboxPathFor(process.env);
  // The 30s default across N allowlisted chats consumes a real share of the per-mailbox Graph
  // read budget (measured 2026-09-02 — see inbox-yield.ts). This knob lets a deployment that
  // also does ad-hoc reads slow the poller down; garbage or non-positive values fall back to
  // the default, same posture as every other best-effort env knob.
  const pollSeconds = Number(process.env['TEAMS_INBOX_POLL_SECONDS']);
  // buildInboxPoller (build-inbox-poller.ts) is the ONE place the roster harvest wire
  // (`roster: membersCache`) is set — extracted so a composition test can drive the real
  // InboxPoller against a real MembersCache through the exact same wiring this call uses,
  // rather than a wire that could silently drop out of THIS file alone (MAJOR 3, 2026-09-04
  // review).
  const poller = buildInboxPoller({
    chats,
    graph,
    tokenProvider,
    membersCache,
    allowlist: config.allowlist,
    inboxPath,
    inboxYieldPath,
    ...(Number.isFinite(pollSeconds) && pollSeconds > 0 ? { pollMs: pollSeconds * 1000 } : {}),
    log: (line) => process.stderr.write(`${line}\n`),
  });
  poller.start();
  process.stderr.write(`inbox poller on: ${inboxPath}\n`);
}

main().catch((error: unknown) => {
  process.stderr.write(`teams-assistant-mcp failed to start: ${(error as Error).message}\n`);
  process.exit(1);
});
