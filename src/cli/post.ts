#!/usr/bin/env node
// Usage: teams-post <chatId> [--html] [--mention "Name"]... < message.txt
// --html: stdin is raw Teams-subset HTML, posted verbatim — the caller escapes their own <,>,&.
// --mention "Name" (repeatable): @mention that person — see server.ts's mentionsSchema for the
// full contract (name must occur in the text / as an @{Name} token with --html).
// Success: one JSON line {ok, id, chat} on stdout, exit 0. See common.ts for the contract.
import { buildContext, doPost, parseSendFlags, readStdin, run, succeed, usage } from './common.js';

await run(async () => {
  const chatId = process.argv[2];
  if (!chatId) usage('usage: teams-post <chatId> [--html] [--mention "Name"]...  (message text on stdin)');
  const { html, mentions } = parseSendFlags(process.argv.slice(3));
  const text = await readStdin();
  if (!text) usage('empty message on stdin');

  succeed(await doPost(buildContext(), chatId, text, html, mentions));
});
