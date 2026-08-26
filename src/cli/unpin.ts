#!/usr/bin/env node
// Usage: teams-unpin <chatId> <messageId>
// Refuses with a clear error if messageId is not the currently-pinned message.
import { buildContext, run, succeed, usage } from './common.js';

await run(async () => {
  const chatId = process.argv[2];
  const messageId = process.argv[3];
  if (!chatId || !messageId) usage('usage: teams-unpin <chatId> <messageId>');

  const { chats, allowlist } = buildContext();
  const entry = allowlist.assertPostable(chatId);
  await chats.unpinMessage(chatId, messageId);
  succeed({ action: 'unpin', messageId, chat: entry.label });
});
