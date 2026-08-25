#!/usr/bin/env node
// Usage: teams-read <chatId> [--limit N] [--since ISO]
// Success: one JSON line {ok, messages: [...]} on stdout — ids, timestamps, sender, text.
import { buildContext, run, succeed, usage } from './common.js';

await run(async () => {
  const chatId = process.argv[2];
  if (!chatId) usage('usage: teams-read <chatId> [--limit N] [--since ISO]');
  const args = process.argv.slice(3);
  const limitIndex = args.indexOf('--limit');
  const sinceIndex = args.indexOf('--since');
  const limit = limitIndex >= 0 ? Number(args[limitIndex + 1]) : 20;
  const since = sinceIndex >= 0 ? args[sinceIndex + 1] : undefined;
  if (!Number.isFinite(limit) || limit < 1) usage('--limit must be a positive number');

  const { chats, allowlist } = buildContext();
  allowlist.assertReadable(chatId);
  const { messages } = await chats.readMessages(chatId, since, limit);
  succeed({
    action: 'read',
    count: messages.length,
    messages: messages.map((m) => ({
      id: m.id,
      at: m.createdDateTime,
      from: m.from,
      deleted: m.isDeleted,
      text: m.text,
    })),
  });
});
