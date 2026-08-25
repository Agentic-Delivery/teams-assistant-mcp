#!/usr/bin/env node
// Usage: teams-react <chatId> <messageId> <emoji>
import { buildContext, run, succeed, usage } from './common.js';

await run(async () => {
  const chatId = process.argv[2];
  const messageId = process.argv[3];
  const emoji = process.argv[4];
  if (!chatId || !messageId || !emoji) usage('usage: teams-react <chatId> <messageId> <emoji>');

  const { chats, allowlist } = buildContext();
  const entry = allowlist.assertPostable(chatId);
  await chats.setReaction(chatId, messageId, emoji);
  succeed({ action: 'react', messageId, emoji, chat: entry.label });
});
