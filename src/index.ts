#!/usr/bin/env node
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { dirname, join } from 'node:path';
import { buildChats } from './build-chats.js';
import { buildInboxPoller } from './build-inbox-poller.js';
import { loadConfig } from './config.js';
import { inboxPathFor, inboxYieldPathFor } from './inbox-yield.js';
import { acquirePollerLock } from './poller-lock.js';
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

  // One poller per inbox, enforced (carries the poller-supervision work from PR #8 onto 0.5.2):
  // two daemons for different projects once raced on one shared account and fed each other's
  // throttle penalty. The lock path is derived from the SAME inboxPathFor(env) helper the poller
  // and the CLIs already use — not a second copy of the resolve logic — so it moves with
  // TEAMS_INBOX_PATH exactly like the inbox and its sidecars do. Keyed per inbox PATH, not per
  // config directory: an instance that sets its own TEAMS_MCP_CONFIG/TEAMS_MCP_TOKEN_CACHE but
  // leaves TEAMS_INBOX_PATH unset still falls back to the same default inbox path as any other
  // such instance, and therefore silently contends for the same lock — see env.example and the
  // README's "Supervising the daemon" section for the operator-facing warning this requires. A
  // live holder wins; the loser here logs why and exits non-zero (0.5.4) rather than lingering as
  // a tools-only server, because a server that silently never polls is exactly the invisible
  // failure this lock exists to prevent, and only a non-zero exit lets a process manager notice.
  const lockPath = join(dirname(inboxPath), 'poller.lock');
  const lock = await acquirePollerLock({ lockPath });
  if (!lock.acquired) {
    process.stderr.write(
      `inbox poller NOT started: pid ${lock.holderPid} already polls ${inboxPath} ` +
        `(holds ${lockPath}${lock.holderStartedAt ? `, since ${lock.holderStartedAt}` : ''}); ` +
        'exiting rather than serving as a silent tools-only twin. If this is a second account ' +
        'that legitimately needs its own poller, set TEAMS_INBOX_PATH to a distinct path for ' +
        'this instance.\n',
    );
    process.exit(1);
  }

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
