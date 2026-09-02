#!/usr/bin/env node
// Usage: teams-send-file <chatId> <path> [more paths...] [--caption "text"]
// --caption "text" (optional, anywhere in argv): shown above the FIRST file's card only — see
// doSendFile's doc comment in common.ts for why a caption is not repeated on every file.
// Streams one JSON success line PER FILE to stdout as EACH one lands, not buffered until the
// whole batch finishes — a later file's failure must never swallow the visible proof that
// earlier files in this same invocation already landed in the chat (see doSendFile's doc
// comment for the 2026-08-24-incident-shaped reasoning this avoids). Exit 0 only once every
// file has sent and its line has fully drained; a failure partway through exits non-zero with
// stderr prose — the caller can tell from the lines ALREADY on stdout which files landed and
// must NOT blindly re-send the whole batch. See common.ts for the rest of the exit-code contract.
import { buildContext, doSendFile, parseSendFileFlags, run, usage, writeLine } from './common.js';

await run(async () => {
  const chatId = process.argv[2];
  if (!chatId) usage('usage: teams-send-file <chatId> <path> [more paths...] [--caption "text"]');
  const { caption, paths } = parseSendFileFlags(process.argv.slice(3));
  if (paths.length === 0) {
    usage('usage: teams-send-file <chatId> <path> [more paths...] [--caption "text"]');
  }

  await doSendFile(buildContext(), chatId, paths, caption, writeLine);
  process.exit(0);
});
