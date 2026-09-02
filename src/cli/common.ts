import { readFile } from 'node:fs/promises';
import { basename } from 'node:path';
import { buildChats } from '../build-chats.js';
import { ChatNotAllowedError, type ChatAllowlist } from '../allowlist.js';
import { defaultDownloadDir, sanitizeFileName, writeDownload } from '../downloads.js';
import { loadConfig } from '../config.js';
import { retryAfterSuffix } from '../graph/graph-client.js';
import type { ReliableTeamsChats } from '../graph/reliable-sends.js';
import type { MentionTarget, PinnedMessage } from '../graph/teams-chats.js';

/**
 * Shared plumbing for the standalone CLIs (teams-post, teams-reply, teams-edit, teams-react,
 * teams-read, teams-send-file).
 *
 * The output contract is the whole point, learned the hard way on 2026-08-24 when a caller
 * grepped for a success token the old ad-hoc script never printed and re-posted a broadcast
 * ten extra times: SUCCESS is exactly one JSON line on stdout and exit 0 — nothing else ever
 * reaches stdout. Failure is prose on stderr and a non-zero exit (2 usage, 3 allowlist,
 * 1 everything else). Callers branch on the exit code, never on output text.
 *
 * ONE exception, documented here rather than only in README/SETUP (2026-09-02 re-review MINOR —
 * a contract stated once in the code it governs, not just in the docs describing it): teams-
 * send-file, given several paths, STREAMS one such JSON success line per file as EACH one lands,
 * rather than buffering until the whole batch finishes — see doSendFile's own doc comment below
 * for why (a later file's failure must never swallow the visible proof that earlier files in the
 * same invocation already landed). The per-line shape and the exit-code rule are otherwise
 * unchanged: each line is still exactly one JSON object, and the run still ends in exactly one
 * exit code.
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

/**
 * Parses teams-send-file's trailing argv: an optional `--caption <text>` (anywhere among the
 * positionals, same flag-anywhere convention as --mention above) and one or more positional file
 * paths, in order. Any OTHER argument starting with `--` is refused (2026-09-02 review MINOR:
 * aligned with teams-reply's doctrine of refusing a stray `--html`/unrecognised leftover instead
 * of silently accepting it) — a typo'd flag must fail loudly, not get quietly uploaded as a
 * literal filename.
 */
export function parseSendFileFlags(args: readonly string[]): { caption?: string; paths: string[] } {
  let caption: string | undefined;
  const paths: string[] = [];
  for (let i = 0; i < args.length; i += 1) {
    const arg = args[i] as string;
    if (arg === '--caption') {
      const value = args[i + 1];
      // Same reasoning as --mention's identical guard above: a missing or flag-like value fails
      // loudly here instead of silently uploading "--caption" (or nothing) as a caption/path.
      if (value === undefined || value.startsWith('--')) {
        usage(
          value === undefined
            ? '--caption needs a value'
            : `--caption needs a value, got "${value}" which looks like a flag`,
        );
      }
      caption = value;
      i += 1;
    } else if (arg.startsWith('--')) {
      usage(`teams-send-file: unrecognised flag ${arg}`);
    } else {
      paths.push(arg);
    }
  }
  return { ...(caption !== undefined ? { caption } : {}), paths };
}

/**
 * Parses teams-attachments' trailing argv: a bare `--list` (metadata only, no download), an
 * optional `--name <filter>` and an optional `--out <dir>`. Any other `--flag` is refused loudly
 * — same doctrine as parseSendFileFlags above — and so is a leftover positional, since this CLI
 * takes none after the chat and message ids.
 */
export function parseAttachmentFlags(args: readonly string[]): { list: boolean; name?: string; out?: string } {
  let list = false;
  let name: string | undefined;
  let out: string | undefined;
  for (let i = 0; i < args.length; i += 1) {
    const arg = args[i] as string;
    if (arg === '--list') {
      list = true;
    } else if (arg === '--name' || arg === '--out') {
      const value = args[i + 1];
      // Same guard as --mention/--caption above: a missing or flag-like value fails loudly
      // instead of being swallowed as the next flag's name.
      if (value === undefined || value.startsWith('--')) {
        usage(
          value === undefined
            ? `${arg} needs a value`
            : `${arg} needs a value, got "${value}" which looks like a flag`,
        );
      }
      if (arg === '--name') {
        name = value;
      } else {
        out = value;
      }
      i += 1;
    } else {
      usage(`teams-attachments: unrecognised argument ${arg}`);
    }
  }
  return { list, ...(name !== undefined ? { name } : {}), ...(out !== undefined ? { out } : {}) };
}

/**
 * teams-read's read-and-map plumbing — same direct-call testability rationale as doPost above,
 * pulled out of read.ts when the output gained attachment metadata (0.5.0): a message carrying a
 * file used to be indistinguishable in this output from one without, so no reader of teams-read
 * ever knew there was anything to download. Attachment metadata (id, name, contentType — never
 * the bytes) is included exactly when a message has any; the inbox daemon's `attachments` count
 * (inbox.ts) stays the coarse signal, this is the detailed one.
 */
export async function doRead(
  { chats, allowlist }: CliContext,
  chatId: string,
  options: { since?: string; limit?: number } = {},
): Promise<{
  action: 'read';
  count: number;
  messages: Array<{
    id: string;
    at: string;
    from: string;
    deleted: boolean;
    text: string;
    attachments?: Array<{ id: string; name?: string; contentType?: string }>;
  }>;
}> {
  allowlist.assertReadable(chatId);
  const { messages } = await chats.readMessages(chatId, options.since, options.limit ?? 20);
  return {
    action: 'read',
    count: messages.length,
    messages: messages.map((m) => ({
      id: m.id,
      at: m.createdDateTime,
      from: m.from,
      deleted: m.isDeleted,
      text: m.text,
      ...(m.attachments.length > 0
        ? {
            attachments: m.attachments.map((a) => ({
              id: a.id,
              ...(a.name ? { name: a.name } : {}),
              ...(a.contentType ? { contentType: a.contentType } : {}),
            })),
          }
        : {}),
    })),
  };
}

/** One downloaded-attachment success entry — what doDownloadAttachments returns per file. */
export interface DownloadedAttachment {
  path: string;
  name: string;
  contentType: string;
  bytes: number;
}

/**
 * teams-attachments' download plumbing — same direct-call testability rationale as doPost above.
 * The allowlist gate runs before any Graph call; the port fetches the message ONCE and downloads
 * every downloadable attachment (optionally narrowed by --name); each file is written through
 * writeDownload, so names are sanitized and an existing file gets a -1/-2/… suffix rather than
 * being silently overwritten.
 */
export async function doDownloadAttachments(
  { chats, allowlist }: CliContext,
  chatId: string,
  messageId: string,
  options: { name?: string; out?: string } = {},
): Promise<{ action: 'attachments'; chat: string; messageId: string; count: number; files: DownloadedAttachment[] }> {
  const entry = allowlist.assertReadable(chatId);
  const payloads = await chats.getAttachments(chatId, messageId, options.name);
  const dir = options.out ?? defaultDownloadDir();
  const files: DownloadedAttachment[] = [];
  for (const payload of payloads) {
    // Sanitized BEFORE the messageId prefix goes on, same as the server tools: a hostile
    // "../../x" name must lose its path components without basename() also eating the prefix.
    const path = await writeDownload(dir, `${messageId}-${sanitizeFileName(payload.name)}`, payload.bytes);
    files.push({ path, name: payload.name, contentType: payload.contentType, bytes: payload.bytes.byteLength });
  }
  return { action: 'attachments', chat: entry.label, messageId, count: files.length, files };
}

/** teams-attachments --list: the metadata, nothing downloaded. Mirrors list_chat_attachments. */
export async function doListAttachments(
  { chats, allowlist }: CliContext,
  chatId: string,
  messageId: string,
): Promise<{
  action: 'attachments-list';
  chat: string;
  messageId: string;
  count: number;
  attachments: Array<{ id: string; name?: string; contentType?: string; downloadable: boolean }>;
}> {
  const entry = allowlist.assertReadable(chatId);
  const attachments = await chats.listAttachments(chatId, messageId);
  return {
    action: 'attachments-list',
    chat: entry.label,
    messageId,
    count: attachments.length,
    attachments: attachments.map((attachment) => ({
      id: attachment.id,
      ...(attachment.name ? { name: attachment.name } : {}),
      ...(attachment.contentType ? { contentType: attachment.contentType } : {}),
      downloadable: attachment.contentType !== 'messageReference',
    })),
  };
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

/** One send-file success payload — the shape `onSent` (doSendFile below) delivers, and
 *  `writeLine` (below) turns into one JSON stdout line. */
export interface SendFileResult {
  action: 'send-file';
  id: string;
  chat: string;
  name: string;
  bytes: number;
}

/**
 * teams-send-file's per-path plumbing. One `chats.sendFile` call per path, in order; `caption`
 * (from --caption) is applied to the FIRST file only — a caption on every card in a multi-file
 * send would repeat the same text under each one, which is never what a caller wants when they
 * pass several paths in one invocation. The allowlist gate runs BEFORE any file is read from
 * disk, same ordering as doPost/doReply/doEdit above: a chat without canPost is refused without
 * even touching the filesystem.
 *
 * STREAMS via `onSent` rather than collecting a return array (2026-09-02 review MAJOR): a
 * multi-file send that fails partway through used to discard the JSON lines for files already
 * posted — exit 1, empty stdout, file 1 irreversibly in the chat, and a caller with no way to
 * tell it had already landed re-runs the whole batch and duplicates it (the exact 2026-08-24
 * incident class the CLI output contract exists to prevent). `onSent` is awaited for EACH file
 * before moving to the next, so — wired to `writeLine` by the real CLI — every earlier success is
 * already flushed to stdout by the time a later file's failure propagates out of this function.
 */
export async function doSendFile(
  { chats, allowlist }: CliContext,
  chatId: string,
  paths: readonly string[],
  caption: string | undefined,
  onSent: (payload: SendFileResult) => void | Promise<void>,
): Promise<void> {
  const entry = allowlist.assertPostable(chatId);
  for (const [index, filePath] of paths.entries()) {
    const buffer = await readFile(filePath);
    const name = basename(filePath);
    const sent = await chats.sendFile(
      chatId,
      { bytes: new Uint8Array(buffer), name },
      index === 0 ? caption : undefined,
    );
    await onSent({ action: 'send-file', id: sent.id, chat: entry.label, name, bytes: buffer.byteLength });
  }
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

/**
 * teams-send-file's per-file streaming primitive: writes ONE JSON success line and resolves only
 * once it is fully drained — same 64 KiB pipe-truncation guard as succeed() above, but awaitable
 * so a caller (doSendFile, via its `onSent` parameter) can write several lines in a row, each
 * confirmed landed on stdout before either sending the next file or letting a later failure
 * propagate. Does NOT itself exit — teams-send-file only exits once every file in the batch has
 * been streamed this way (see send-file.ts).
 */
export function writeLine<T extends object>(payload: T): Promise<void> {
  return new Promise((resolve) => {
    process.stdout.write(`${JSON.stringify({ ok: true, ...payload })}\n`, () => resolve());
  });
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
