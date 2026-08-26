#!/usr/bin/env node
// Usage: teams-edit <chatId> <messageId> [--html] [--mention "Name"]... < new-text.txt
// --html: stdin is raw Teams-subset HTML, posted verbatim — the caller escapes their own <,>,&.
// --mention "Name" (repeatable): @mention that person — see server.ts's mentionsSchema for the
// full contract (name must occur in the text / as an @{Name} token with --html).
// Success: one JSON line {ok, id, chat} on stdout, exit 0. See common.ts for the contract.
import { buildContext, doEdit, parseSendFlags, readStdin, run, succeed, usage } from './common.js';

await run(async () => {
  const chatId = process.argv[2];
  const messageId = process.argv[3];
  if (!chatId || !messageId) {
    usage('usage: teams-edit <chatId> <messageId> [--html] [--mention "Name"]...  (new text on stdin)');
  }
  const { html, mentions } = parseSendFlags(process.argv.slice(4));
  const text = await readStdin();
  if (!text) usage('empty message on stdin');

  succeed(await doEdit(buildContext(), chatId, messageId, text, html, mentions));
});
