#!/usr/bin/env node
// Usage: teams-post <chatId> [--html] < message.txt
// --html: stdin is raw Teams-subset HTML, posted verbatim — the caller escapes their own <,>,&.
// Success: one JSON line {ok, id, chat} on stdout, exit 0. See common.ts for the contract.
import { buildContext, doPost, readStdin, run, succeed, usage } from './common.js';

await run(async () => {
  const chatId = process.argv[2];
  if (!chatId) usage('usage: teams-post <chatId> [--html]  (message text on stdin)');
  const html = process.argv.slice(3).includes('--html');
  const text = await readStdin();
  if (!text) usage('empty message on stdin');

  succeed(await doPost(buildContext(), chatId, text, html));
});
