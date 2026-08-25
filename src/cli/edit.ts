#!/usr/bin/env node
// Usage: teams-edit <chatId> <messageId> [--html] < new-text.txt
// --html: stdin is raw Teams-subset HTML, posted verbatim — the caller escapes their own <,>,&.
// Success: one JSON line {ok, id, chat} on stdout, exit 0. See common.ts for the contract.
import { buildContext, doEdit, readStdin, run, succeed, usage } from './common.js';

await run(async () => {
  const chatId = process.argv[2];
  const messageId = process.argv[3];
  if (!chatId || !messageId) usage('usage: teams-edit <chatId> <messageId> [--html]  (new text on stdin)');
  const html = process.argv.slice(4).includes('--html');
  const text = await readStdin();
  if (!text) usage('empty message on stdin');

  succeed(await doEdit(buildContext(), chatId, messageId, text, html));
});
