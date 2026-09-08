#!/usr/bin/env node
// Usage: teams-send-file <chatId> <path> [more paths...] [--caption "text"]
//        [--grant-to <id>[,<id>...] | --no-grant]
// --caption "text" (optional, anywhere in argv): shown above the FIRST file's card only — see
// doSendFile's doc comment in common.ts for why a caption is not repeated on every file.
// --grant-to <id>[,<id>...] (0.6.0, live 2026-09-08): skips the roster lookup entirely and grants
// read access to exactly the given AAD ids — no /members or /me call. --no-grant uploads and
// posts the card with NO permission grant at all (only the sender can open it); mutually
// exclusive with --grant-to. Both exist for when the roster is throttled/unavailable and the
// caller already knows who should (or need not) see the file — see GraphTeamsChats.sendFile's
// own doc comment for the full reasoning.
// Streams one JSON success line PER FILE to stdout as EACH one lands, not buffered until the
// whole batch finishes — a later file's failure must never swallow the visible proof that
// earlier files in this same invocation already landed in the chat (see doSendFile's doc
// comment for the 2026-08-24-incident-shaped reasoning this avoids). Exit 0 only once every
// file has sent and its line has fully drained; a failure partway through exits non-zero with
// stderr prose — the caller can tell from the lines ALREADY on stdout which files landed and
// must NOT blindly re-send the whole batch. See common.ts for the rest of the exit-code contract.
import { buildContext, doSendFile, parseSendFileFlags, run, usage, writeLine } from './common.js';

const USAGE =
  'usage: teams-send-file <chatId> <path> [more paths...] [--caption "text"] ' +
  '[--grant-to <id>[,<id>...] | --no-grant]';

await run(async () => {
  const chatId = process.argv[2];
  if (!chatId) usage(USAGE);
  const { caption, paths, grantTo, noGrant } = parseSendFileFlags(process.argv.slice(3));
  if (paths.length === 0) {
    usage(USAGE);
  }

  await doSendFile(buildContext(), chatId, paths, caption, writeLine, {
    ...(grantTo !== undefined ? { grantTo } : {}),
    ...(noGrant !== undefined ? { noGrant } : {}),
  });
  process.exit(0);
});
