import { buildChats } from '../build-chats.js';
import { ChatNotAllowedError, type ChatAllowlist } from '../allowlist.js';
import { loadConfig } from '../config.js';
import type { ReliableTeamsChats } from '../graph/reliable-sends.js';
import type { MentionTarget } from '../graph/teams-chats.js';

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
      if (value === undefined) usage('--mention needs a name');
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

export async function run(main: () => Promise<void>): Promise<void> {
  try {
    await main();
  } catch (caught) {
    process.stderr.write(`${caught instanceof Error ? caught.message : String(caught)}\n`);
    process.exit(caught instanceof ChatNotAllowedError ? 3 : 1);
  }
}
