import { buildChats } from '../build-chats.js';
import { ChatNotAllowedError, type ChatAllowlist } from '../allowlist.js';
import { loadConfig } from '../config.js';
import { retryAfterSuffix } from '../graph/graph-client.js';
import type { ReliableTeamsChats } from '../graph/reliable-sends.js';
import type { MentionTarget, PinnedMessage } from '../graph/teams-chats.js';

/**
 * Shared plumbing for the standalone CLIs (teams-post, teams-reply, teams-edit, teams-react,
 * teams-read).
 *
 * The output contract is the whole point, learned the hard way on 2026-08-24 when a caller
 * grepped for a success token the old ad-hoc script never printed and re-posted a broadcast
 * ten extra times: SUCCESS is exactly one JSON line on stdout and exit 0 — nothing else ever
 * reaches stdout. Failure is prose on stderr and a non-zero exit (2 usage, 3 allowlist,
 * 1 everything else). Callers branch on the exit code, never on output text.
 */
export interface CliContext {
  chats: ReliableTeamsChats;
  allowlist: ChatAllowlist;
}

export function buildContext(): CliContext {
  const config = loadConfig();
  return { chats: buildChats(config).chats, allowlist: config.allowlist };
}

/**
 * Parses the flags teams-post and teams-edit share off their trailing argv: `--html` (a bare
 * flag) and repeatable `--mention <name>` (one name per occurrence, in order). Anything else is
 * left in `rest` untouched — callers that take positional args after the chat/message ids
 * (none currently do) are free to inspect it.
 */
export function parseSendFlags(args: readonly string[]): { html: boolean; mentions: string[]; rest: string[] } {
  let html = false;
  const mentions: string[] = [];
  const rest: string[] = [];
  for (let i = 0; i < args.length; i += 1) {
    const arg = args[i];
    if (arg === '--html') {
      html = true;
    } else if (arg === '--mention') {
      const value = args[i + 1];
      // A missing value AND a flag-like value ("--mention --html", the next flag swallowed as
      // the name) both fail loudly here — the alternative, silently taking "--html" as a mention
      // name, only surfaces later as a confusing "No chat member matches "--html"" from
      // resolveMentions, which gives no hint the real problem was a missing --mention argument.
      if (value === undefined || value.startsWith('--')) {
        usage(
          value === undefined
            ? '--mention needs a name'
            : `--mention needs a name, got "${value}" which looks like a flag`,
        );
      }
      mentions.push(value);
      i += 1;
    } else {
      rest.push(arg as string);
    }
  }
  return { html, mentions, rest };
}

/** Resolves --mention names against the chat's member list, or [] when none were given — same
 *  as the MCP tools' resolveMentions/mentions plumbing in server.ts. */
async function resolveMentions(
  chats: CliContext['chats'],
  chatId: string,
  mentions: readonly string[],
): Promise<MentionTarget[]> {
  return mentions.length > 0 ? chats.resolveMentions(chatId, mentions) : [];
}

/**
 * teams-post's --html routing, pulled out of post.ts so it is unit-testable without a live
 * send: a subprocess test cannot distinguish "the --html branch runs" from "the flag was
 * ignored", because assertPostable always throws (or not) before either send path is ever
 * reached — the allowlist gate looks identical either way. Calling this directly with a fake
 * TeamsChatsPort proves which method actually gets called.
 */
export async function doPost(
  { chats, allowlist }: CliContext,
  chatId: string,
  text: string,
  html: boolean,
  mentions: readonly string[] = [],
): Promise<{ action: 'post'; id: string; chat: string }> {
  const entry = allowlist.assertPostable(chatId);
  const resolved = await resolveMentions(chats, chatId, mentions);
  const sent = html
    ? await chats.sendHtmlMessage(chatId, text, resolved)
    : await chats.sendMessage(chatId, text, resolved);
  return { action: 'post', id: sent.id, chat: entry.label };
}

/** teams-edit's --html routing — same rationale as doPost above. */
export async function doEdit(
  { chats, allowlist }: CliContext,
  chatId: string,
  messageId: string,
  newText: string,
  html: boolean,
  mentions: readonly string[] = [],
): Promise<{ action: 'edit'; id: string; chat: string }> {
  const entry = allowlist.assertPostable(chatId);
  const resolved = await resolveMentions(chats, chatId, mentions);
  if (html) {
    await chats.editHtmlMessage(chatId, messageId, newText, resolved);
  } else {
    await chats.editMessage(chatId, messageId, newText, resolved);
  }
  return { action: 'edit', id: messageId, chat: entry.label };
}

/** teams-reply's --mention plumbing — same rationale as doPost/doEdit above (testable without a
 *  live send; a subprocess test can't distinguish "mentions were resolved and forwarded" from
 *  "the flag was silently ignored"). */
export async function doReply(
  { chats, allowlist }: CliContext,
  chatId: string,
  replyToMessageId: string,
  text: string,
  mentions: readonly string[] = [],
): Promise<{ action: 'reply'; id: string; inReplyTo: string; chat: string }> {
  const entry = allowlist.assertPostable(chatId);
  const resolved = await resolveMentions(chats, chatId, mentions);
  const sent = await chats.replyToMessage(chatId, replyToMessageId, text, resolved);
  return { action: 'reply', id: sent.id, inReplyTo: replyToMessageId, chat: entry.label };
}

/**
 * teams-pin's confirm-before-claiming-success check — same reasoning as pin_chat_message in
 * server.ts: Graph reporting the pinMessage POST as a success is not proof the pin landed, only
 * the re-list pinMessage itself returns is. Duplicated here (not shared with server.ts) because
 * the two live in separate module graphs with no existing shared "tool logic" layer — same
 * wording, kept in sync by hand, same as the rest of the CLI/MCP-tool boundary.
 */
export async function doPin(
  { chats, allowlist }: CliContext,
  chatId: string,
  messageId: string,
): Promise<{ action: 'pin'; messageId: string; chat: string; pinnedMessages: readonly PinnedMessage[] }> {
  const entry = allowlist.assertPostable(chatId);
  const pinned = await chats.pinMessage(chatId, messageId);
  if (!pinned.some((entry2) => entry2.messageId === messageId)) {
    throw new Error(
      `Pin request for message ${messageId} was accepted, but the post-pin list does not show it ` +
        `pinned (currently pinned: ${pinned.map((entry2) => entry2.messageId).join(', ') || '(nothing)'}) ` +
        '— the outcome is not confirmed; do not assume the pin landed.',
    );
  }
  return { action: 'pin', messageId, chat: entry.label, pinnedMessages: pinned };
}

export function readStdin(): Promise<string> {
  return new Promise((resolve, reject) => {
    let buffer = '';
    process.stdin.setEncoding('utf8'); // Buffer concat would corrupt a multibyte char split across chunks.
    process.stdin.on('data', (chunk) => (buffer += chunk));
    process.stdin.on('end', () => resolve(buffer.trim()));
    process.stdin.on('error', reject);
  });
}

/** Writes the single JSON success line, WAITS for the pipe to drain, then exits 0 — a payload
 *  bigger than the 64 KiB pipe buffer would otherwise be truncated mid-JSON with exit 0. */
export function succeed(payload: Record<string, unknown>): void {
  process.stdout.write(`${JSON.stringify({ ok: true, ...payload })}\n`, () => process.exit(0));
}

export function usage(text: string): never {
  process.stderr.write(`${text}\n`);
  process.exit(2);
}

/**
 * A Graph 429 anywhere in the send/reply/edit flow must never look like an ordinary failure an
 * operator might reflexively retry by hand: writes never auto-retry past a 429 (that rule already
 * lives in GraphClient/ReliableTeamsChats), so the CLI's own error text is the only place left to
 * say "wait, don't just run it again" — named in seconds, straight from Graph's own Retry-After,
 * when Graph actually sent one. Nothing is invented when it didn't (a 429 with no named window
 * gets no "retry after" claim at all — a wrong number is worse than none). The phrasing itself
 * comes from retryAfterSuffix (graph-client.ts) — the MCP tool path's guard() (server.ts) renders
 * the same suffix on the same errors, so the two consumer-facing surfaces never disagree.
 */
function formatCliError(caught: unknown): string {
  const base = caught instanceof Error ? caught.message : String(caught);
  return `${base}${retryAfterSuffix(caught)}`;
}

export async function run(main: () => Promise<void>): Promise<void> {
  try {
    await main();
  } catch (caught) {
    process.stderr.write(`${formatCliError(caught)}\n`);
    process.exit(caught instanceof ChatNotAllowedError ? 3 : 1);
  }
}
