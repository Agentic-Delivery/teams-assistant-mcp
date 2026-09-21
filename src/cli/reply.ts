#!/usr/bin/env node
// Usage: teams-reply <chatId> <messageId> [--html | --text] [--mention "Name"]... < reply.txt
// --html: stdin is raw Teams-subset HTML, posted verbatim after the quote card — same contract
// as teams-post's --html (the caller escapes their own <,>,&; mentions via the @{Name}
// placeholder — see server.ts's mentionsSchema and mentions.ts's renderHtmlWithMentions).
// --text: sends stdin as plain text even when it looks structured (see the plain-text guard in
// common.ts's doReply) — the deliberate override. --html and --text together is an error.
// --mention "Name" (repeatable): @mention that person — see server.ts's mentionsSchema for the
// full contract (the name must occur in the reply text, or as an @{Name} token with --html).
// Flags parse anywhere in argv — same order-independence as teams-post (C2, audit fix
// 2026-09-21); a resolved chat id that still starts with `--` is refused rather than silently
// misparsed.
import { buildContext, doReply, parseSendFlags, readStdin, run, succeed, usage } from './common.js';

await run(async () => {
  const { html, text, mentions, rest } = parseSendFlags(process.argv.slice(2));
  const [chatId, messageId, ...extra] = rest;
  if (!chatId || !messageId) {
    usage('usage: teams-reply <chatId> <messageId> [--html | --text] [--mention "Name"]...  (text on stdin)');
  }
  if (chatId.startsWith('--')) usage('that looks like a flag; the chat id comes first');
  if (extra.length > 0) usage(`teams-reply: unrecognised argument(s): ${extra.join(' ')}`);
  const body = await readStdin();
  if (!body) usage('empty reply on stdin');

  succeed(await doReply(buildContext(), chatId, messageId, body, html, mentions, { plainTextOverride: text }));
});
