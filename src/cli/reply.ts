#!/usr/bin/env node
// Usage: teams-reply <chatId> <messageId> [--mention "Name"]... < reply.txt
// --mention "Name" (repeatable): @mention that person — see server.ts's mentionsSchema for the
// full contract (the name must occur in the reply text).
import { buildContext, doReply, parseSendFlags, readStdin, run, succeed, usage } from './common.js';

await run(async () => {
  const chatId = process.argv[2];
  const messageId = process.argv[3];
  if (!chatId || !messageId) {
    usage('usage: teams-reply <chatId> <messageId> [--mention "Name"]...  (text on stdin)');
  }
  const { html, mentions, rest } = parseSendFlags(process.argv.slice(4));
  // parseSendFlags also understands --html (shared with post/edit), but teams-reply has no raw-
  // HTML path — silently accepting and ignoring it would post the reply as plain text with no
  // error at all, which is exactly the kind of quiet, wrong-thing-happened failure this CLI
  // suite exists to refuse. Any other leftover argument gets the same treatment.
  if (html) usage('teams-reply does not support --html');
  if (rest.length > 0) usage(`teams-reply: unrecognised argument(s): ${rest.join(' ')}`);
  const text = await readStdin();
  if (!text) usage('empty reply on stdin');

  succeed(await doReply(buildContext(), chatId, messageId, text, mentions));
});
