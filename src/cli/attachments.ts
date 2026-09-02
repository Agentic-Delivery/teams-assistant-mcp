#!/usr/bin/env node
// Usage: teams-attachments <chatId> <messageId> [--list] [--name <filter>] [--out <dir>]
// Default: downloads EVERY downloadable attachment on the message (quote cards skipped) into
// --out, else TEAMS_MCP_DOWNLOAD_DIR, else a tmpdir — one JSON success line {ok, files: [...]}
// with the absolute paths. --list prints the attachment metadata instead and downloads nothing.
// --name (download mode only) narrows to attachments whose name contains the filter, case-
// insensitive; --list always shows everything. Downloads never overwrite: collisions get -1/-2/….
import { inboxYieldPathFor } from '../inbox-yield.js';
import { buildContext, doDownloadAttachments, doListAttachments, parseAttachmentFlags, run, succeed, usage } from './common.js';

await run(async () => {
  const chatId = process.argv[2];
  const messageId = process.argv[3];
  if (!chatId || !messageId) {
    usage('usage: teams-attachments <chatId> <messageId> [--list] [--name <filter>] [--out <dir>]');
  }
  const { list, name, out } = parseAttachmentFlags(process.argv.slice(4));
  if (list && (name !== undefined || out !== undefined)) {
    usage('--list takes no --name/--out: it prints every attachment\'s metadata and downloads nothing');
  }

  const context = buildContext();
  // The yield file quiets any inbox poller sharing this mailbox's Graph budget — without it, a
  // running daemon starves this CLI outright (measured 2026-09-02; see inbox-yield.ts).
  const yieldPath = inboxYieldPathFor(process.env);
  succeed(
    list
      ? await doListAttachments(context, chatId, messageId, { yieldPath })
      : await doDownloadAttachments(context, chatId, messageId, {
          yieldPath,
          ...(name !== undefined ? { name } : {}),
          ...(out !== undefined ? { out } : {}),
        }),
  );
});
