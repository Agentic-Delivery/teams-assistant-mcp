#!/usr/bin/env node
// Usage: teams-delete <chatId> <messageId> [--undo] [--force]
// Soft-deletes a message this account sent (Teams shows "This message was deleted"; the
// original can be put back with --undo). Own messages only: the message is fetched and its
// author checked before anything is sent — somebody else's message is refused unless --force.
// Success: one JSON line {ok, action, messageId, chat} on stdout, exit 0. See common.ts for
// the contract.
import { buildContext, doDelete, parseDeleteFlags, run, succeed, usage } from './common.js';

await run(async () => {
  const chatId = process.argv[2];
  const messageId = process.argv[3];
  if (!chatId || !messageId) {
    usage('usage: teams-delete <chatId> <messageId> [--undo] [--force]');
  }
  const flags = parseDeleteFlags(process.argv.slice(4));

  succeed(await doDelete(buildContext(), chatId, messageId, flags));
});
