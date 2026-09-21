#!/usr/bin/env node
// Usage: teams-edit <chatId> <messageId> [--html | --text] [--mention "Name"]... < new-text.txt
// --html: stdin is raw Teams-subset HTML, posted verbatim — the caller escapes their own <,>,&.
// --text: sends stdin as plain text even when it looks structured (see the plain-text guard in
// common.ts's doEdit) — the deliberate override. --html and --text together is an error.
// --mention "Name" (repeatable): @mention that person — see server.ts's mentionsSchema for the
// full contract (name must occur in the text / as an @{Name} token with --html).
// Flags parse anywhere in argv — same order-independence as teams-post (C2, audit fix
// 2026-09-21); a resolved chat id that still starts with `--` is refused rather than silently
// misparsed.
// Success: one JSON line {ok, id, chat} on stdout, exit 0. See common.ts for the contract.
import { buildContext, doEdit, parseSendFlags, readStdin, run, succeed, usage } from './common.js';

await run(async () => {
  const { html, text, mentions, rest } = parseSendFlags(process.argv.slice(2));
  const [chatId, messageId] = rest;
  if (!chatId || !messageId) {
    usage('usage: teams-edit <chatId> <messageId> [--html | --text] [--mention "Name"]...  (new text on stdin)');
  }
  if (chatId.startsWith('--')) usage('that looks like a flag; the chat id comes first');
  const body = await readStdin();
  if (!body) usage('empty message on stdin');

  succeed(await doEdit(buildContext(), chatId, messageId, body, html, mentions, { plainTextOverride: text }));
});
