#!/usr/bin/env node
// Usage: teams-read <chatId> [--limit N] [--since ISO]
// Success: one JSON line {ok, messages: [...]} on stdout — ids, timestamps, sender, text, and
// (when a message carries any) attachment metadata. Metadata only: teams-attachments downloads.
import { buildContext, doRead, run, succeed, usage } from './common.js';

await run(async () => {
  const chatId = process.argv[2];
  if (!chatId) usage('usage: teams-read <chatId> [--limit N] [--since ISO]');
  const args = process.argv.slice(3);
  const limitIndex = args.indexOf('--limit');
  const sinceIndex = args.indexOf('--since');
  const limit = limitIndex >= 0 ? Number(args[limitIndex + 1]) : 20;
  const since = sinceIndex >= 0 ? args[sinceIndex + 1] : undefined;
  if (!Number.isFinite(limit) || limit < 1) usage('--limit must be a positive number');
  if (sinceIndex >= 0 && !since) usage('--since needs an ISO-8601 value');

  succeed(await doRead(buildContext(), chatId, { ...(since !== undefined ? { since } : {}), limit }));
});
