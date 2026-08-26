#!/usr/bin/env node
// Usage: teams-pin <chatId> <messageId>
// A chat effectively holds ONE pin: this REPLACES whatever was pinned before it, even though
// Graph reports success either way (verified live 2026-08-25) — see server.ts's
// pin_chat_message description for the full note.
import { buildContext, run, succeed, usage } from './common.js';

await run(async () => {
  const chatId = process.argv[2];
  const messageId = process.argv[3];
  if (!chatId || !messageId) usage('usage: teams-pin <chatId> <messageId>');

  const { chats, allowlist } = buildContext();
  const entry = allowlist.assertPostable(chatId);
  const pinned = await chats.pinMessage(chatId, messageId);
  succeed({ action: 'pin', messageId, chat: entry.label, pinnedMessages: pinned });
});
