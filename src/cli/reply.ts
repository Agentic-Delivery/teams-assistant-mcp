#!/usr/bin/env node
// Usage: teams-reply <chatId> <messageId> [--html] [--mention "Name"]... < reply.txt
// --html: stdin is raw Teams-subset HTML, posted verbatim after the quote card — same contract
// as teams-post's --html (the caller escapes their own <,>,&; mentions via the @{Name}
// placeholder — see server.ts's mentionsSchema and mentions.ts's renderHtmlWithMentions).
// --mention "Name" (repeatable): @mention that person — see server.ts's mentionsSchema for the
// full contract (the name must occur in the reply text, or as an @{Name} token with --html).
import { buildContext, doReply, parseSendFlags, readStdin, run, succeed, usage } from './common.js';

await run(async () => {
  const chatId = process.argv[2];
  const messageId = process.argv[3];
  if (!chatId || !messageId) {
    usage('usage: teams-reply <chatId> <messageId> [--html] [--mention "Name"]...  (text on stdin)');
  }
  const { html, mentions, rest } = parseSendFlags(process.argv.slice(4));
  if (rest.length > 0) usage(`teams-reply: unrecognised argument(s): ${rest.join(' ')}`);
  const text = await readStdin();
  if (!text) usage('empty reply on stdin');

  succeed(await doReply(buildContext(), chatId, messageId, text, html, mentions));
});
