#!/usr/bin/env node
// Usage: teams-post <chatId> < message.txt
// Success: one JSON line {ok, id, chat} on stdout, exit 0. See common.ts for the contract.
import { buildContext, readStdin, run, succeed, usage } from './common.js';

await run(async () => {
  const chatId = process.argv[2];
  if (!chatId) usage('usage: teams-post <chatId>  (message text on stdin)');
  const text = await readStdin();
  if (!text) usage('empty message on stdin');

  const { chats, allowlist } = buildContext();
  const entry = allowlist.assertPostable(chatId);
  const sent = await chats.sendMessage(chatId, text);
  succeed({ action: 'post', id: sent.id, chat: entry.label });
});
