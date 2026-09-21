#!/usr/bin/env node
// Usage: teams-post <chatId> [--html | --text] [--mention "Name"]... < message.txt
// --html: stdin is raw Teams-subset HTML, posted verbatim — the caller escapes their own <,>,&.
// --text: sends stdin as plain text even when it looks structured (see the plain-text guard in
// common.ts's doPost) — the deliberate override. --html and --text together is an error.
// --mention "Name" (repeatable): @mention that person — see server.ts's mentionsSchema for the
// full contract (name must occur in the text / as an @{Name} token with --html).
// Flags parse anywhere in argv — `teams-post --html <chatId>` and `teams-post <chatId> --html`
// are equivalent (C2, audit fix 2026-09-21); a resolved chat id that still starts with `--` is
// refused rather than silently misparsed.
// Success: one JSON line {ok, id, chat} on stdout, exit 0. See common.ts for the contract.
import { buildContext, doPost, parseSendFlags, readStdin, run, succeed, usage } from './common.js';

await run(async () => {
  const { html, text, mentions, rest } = parseSendFlags(process.argv.slice(2));
  const chatId = rest[0];
  if (!chatId) {
    usage('usage: teams-post <chatId> [--html | --text] [--mention "Name"]...  (message text on stdin)');
  }
  if (chatId.startsWith('--')) usage('that looks like a flag; the chat id comes first');
  const body = await readStdin();
  if (!body) usage('empty message on stdin');

  succeed(await doPost(buildContext(), chatId, body, html, mentions, { plainTextOverride: text }));
});
