#!/usr/bin/env node
// Usage: teams-pin <chatId> <messageId>
// A chat effectively holds ONE pin: this REPLACES whatever was pinned before it, even though
// Graph reports success either way (verified live 2026-08-25) — see server.ts's
// pin_chat_message description for the full note. Refuses (never succeeds) if the post-pin
// re-list does not actually show messageId pinned — see doPin in common.ts.
import { buildContext, doPin, run, succeed, usage } from './common.js';

await run(async () => {
  const chatId = process.argv[2];
  const messageId = process.argv[3];
  if (!chatId || !messageId) usage('usage: teams-pin <chatId> <messageId>');

  succeed(await doPin(buildContext(), chatId, messageId));
});
