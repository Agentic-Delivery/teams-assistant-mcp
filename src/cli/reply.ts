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
  const { mentions } = parseSendFlags(process.argv.slice(4));
  const text = await readStdin();
  if (!text) usage('empty reply on stdin');

  succeed(await doReply(buildContext(), chatId, messageId, text, mentions));
});
