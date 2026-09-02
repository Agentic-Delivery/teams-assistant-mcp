#!/usr/bin/env node
// Usage: teams-send-file <chatId> <path> [more paths...] [--caption "text"]
// --caption "text" (optional, anywhere in argv): shown above the FIRST file's card only — see
// doSendFile's doc comment in common.ts for why a caption is not repeated on every file.
// Success: one JSON line PER SENT FILE on stdout ({ok, action, id, chat, name, bytes}), exit 0.
// See common.ts for the exit-code/no-text-branching contract.
import { buildContext, doSendFile, parseSendFileFlags, run, succeedMany, usage } from './common.js';

await run(async () => {
  const chatId = process.argv[2];
  if (!chatId) usage('usage: teams-send-file <chatId> <path> [more paths...] [--caption "text"]');
  const { caption, paths } = parseSendFileFlags(process.argv.slice(3));
  if (paths.length === 0) {
    usage('usage: teams-send-file <chatId> <path> [more paths...] [--caption "text"]');
  }

  succeedMany(await doSendFile(buildContext(), chatId, paths, caption));
});
