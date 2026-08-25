#!/usr/bin/env node
// Usage: teams-reply <chatId> <messageId> < reply.txt
import { buildContext, readStdin, run, succeed, usage } from './common.js';

await run(async () => {
  const chatId = process.argv[2];
  const messageId = process.argv[3];
  if (!chatId || !messageId) usage('usage: teams-reply <chatId> <messageId>  (text on stdin)');
  const text = await readStdin();
  if (!text) usage('empty reply on stdin');

  const { chats, allowlist } = buildContext();
  const entry = allowlist.assertPostable(chatId);
  const sent = await chats.replyToMessage(chatId, messageId, text);
  succeed({ action: 'reply', id: sent.id, inReplyTo: messageId, chat: entry.label });
});
