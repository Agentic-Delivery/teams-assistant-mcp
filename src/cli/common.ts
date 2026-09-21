import { readFile } from 'node:fs/promises';
import { basename } from 'node:path';
import { buildChats } from '../build-chats.js';
import { ChatNotAllowedError, type ChatAllowlist } from '../allowlist.js';
import { defaultDownloadDir, sanitizeFileName, writeDownload } from '../downloads.js';
import { withQuotaYield } from '../inbox-yield.js';
import { loadConfig } from '../config.js';
import { retryAfterSuffix } from '../graph/graph-client.js';
import type { ReliableTeamsChats } from '../graph/reliable-sends.js';
import { MessageOwnershipError, type MentionTarget, type PinnedMessage } from '../graph/teams-chats.js';

/**
 * Shared plumbing for the standalone CLIs (teams-post, teams-reply, teams-edit, teams-react,
 * teams-read, teams-send-file, teams-attachments, teams-delete).
 *
 * The output contract is the whole point, learned the hard way on 2026-08-24 when a caller
 * grepped for a success token the old ad-hoc script never printed and re-posted a broadcast
 * ten extra times: SUCCESS is exactly one JSON line on stdout and exit 0 — nothing else ever
 * reaches stdout. Failure is prose on stderr and a non-zero exit (2 usage — including the C1
 * plain-text guard below, PlainTextRefusedError, a caller-input problem the same shape as a bad
 * flag; 3 allowlist; 4 ownership refusal — teams-delete's own-message gate, MessageOwnershipError,
 * FINDING 4 of the message-withdrawal review — 1 everything else). Callers branch on the exit
 * code, never on output text.
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
 * Parses the flags teams-post, teams-reply and teams-edit share, ANYWHERE in their argv (C2,
 * audit fix 2026-09-21 — the caller-facing entry points feed this the WHOLE trailing argv,
 * chat/message id included, and take their positionals off `rest`; see post.ts/reply.ts/edit.ts):
 * `--html` (a bare flag), `--text` (a bare flag, the deliberate plain-text override for the
 * guard in doPost/doReply/doEdit below — see common.ts's CliContext doc comment for the exit
 * code), and repeatable `--mention <name>` (one name per occurrence, in order). `--html` and
 * `--text` together is refused — the caller must pick one. Anything else (including the
 * chat/message id) is left in `rest`, in order, untouched.
 */
export function parseSendFlags(
  args: readonly string[],
): { html: boolean; text: boolean; mentions: string[]; rest: string[] } {
  let html = false;
  let text = false;
  const mentions: string[] = [];
  const rest: string[] = [];
  for (let i = 0; i < args.length; i += 1) {
    const arg = args[i];
    if (arg === '--html') {
      html = true;
    } else if (arg === '--text') {
      text = true;
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
  if (html && text) {
    usage(
      '--html and --text are mutually exclusive: --html sends styled markup, --text forces ' +
        'plain text regardless of length. Pick one.',
    );
  }
  return { html, text, mentions, rest };
}

/**
 * Parses teams-send-file's trailing argv: an optional `--caption <text>` (anywhere among the
 * positionals, same flag-anywhere convention as --mention above), an optional `--grant-to
 * <id>[,<id>…]` (comma-separated, at least one non-blank id) OR a bare `--no-grant` — mutually
 * exclusive, both skip the roster lookup entirely (0.6.0, live 2026-09-08: the caller-facing
 * escape hatch around a throttled/unavailable roster, see GraphTeamsChats.sendFile's own doc
 * comment) — and one or more positional file paths, in order. Any OTHER argument starting with
 * `--` is refused (2026-09-02 review MINOR: aligned with teams-reply's doctrine of refusing a
 * stray `--html`/unrecognised leftover instead of silently accepting it) — a typo'd flag must
 * fail loudly, not get quietly uploaded as a literal filename.
 */
export function parseSendFileFlags(
  args: readonly string[],
): { caption?: string; paths: string[]; grantTo?: string[]; noGrant?: boolean } {
  let caption: string | undefined;
  let grantTo: string[] | undefined;
  let noGrant = false;
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
    } else if (arg === '--grant-to') {
      const value = args[i + 1];
      if (value === undefined || value.startsWith('--')) {
        usage(
          value === undefined
            ? '--grant-to needs a value'
            : `--grant-to needs a value, got "${value}" which looks like a flag`,
        );
      }
      // Blank entries (a stray comma, leading/trailing whitespace) are dropped, not passed
      // through as a literal empty-string id Graph would reject with a confusing error far from
      // the actual mistake; an ALL-blank list collapses to the same "no one to grant" ambiguity
      // an explicit empty array would be, refused loudly here rather than reaching sendFile at
      // all — mirrors the empty-mention-name refusal in parseSendFlags above.
      const ids = value.split(',').map((id) => id.trim()).filter((id) => id !== '');
      if (ids.length === 0) {
        usage(`--grant-to needs at least one non-blank id, got "${value}"`);
      }
      grantTo = ids;
      i += 1;
    } else if (arg === '--no-grant') {
      noGrant = true;
    } else if (arg.startsWith('--')) {
      usage(`teams-send-file: unrecognised flag ${arg}`);
    } else {
      paths.push(arg);
    }
  }
  if (grantTo !== undefined && noGrant) {
    usage('--grant-to and --no-grant are mutually exclusive.');
  }
  return {
    ...(caption !== undefined ? { caption } : {}),
    paths,
    ...(grantTo !== undefined ? { grantTo } : {}),
    ...(noGrant ? { noGrant } : {}),
  };
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
 * Parses teams-delete's trailing argv: a bare `--undo` (restore instead of delete) and a bare
 * `--force` (skip the own-message check). No values, no positionals — anything else is refused
 * loudly, same doctrine as parseAttachmentFlags above.
 */
export function parseDeleteFlags(args: readonly string[]): { undo: boolean; force: boolean } {
  let undo = false;
  let force = false;
  for (const arg of args) {
    if (arg === '--undo') {
      undo = true;
    } else if (arg === '--force') {
      force = true;
    } else {
      usage(`teams-delete: unrecognised argument ${arg}`);
    }
  }
  return { undo, force };
}

/**
 * teams-delete's routing — same direct-call testability rationale as doPost below: a subprocess
 * test stops at the allowlist gate either way, so which port method runs (delete or undo) and
 * whether --force reached it can only be proven by calling this with a fake port. The
 * own-message check itself lives in the port (GraphTeamsChats.assertOwnMessage), not here, so
 * the MCP tool and this CLI cannot drift apart on it.
 */
export async function doDelete(
  { chats, allowlist }: CliContext,
  chatId: string,
  messageId: string,
  options: { undo: boolean; force: boolean },
): Promise<{ action: 'delete' | 'undo-delete'; messageId: string; chat: string }> {
  const entry = allowlist.assertPostable(chatId);
  if (options.undo) {
    await chats.undoDeleteMessage(chatId, messageId, { force: options.force });
    return { action: 'undo-delete', messageId, chat: entry.label };
  }
  await chats.deleteMessage(chatId, messageId, { force: options.force });
  return { action: 'delete', messageId, chat: entry.label };
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
    format: 'html' | 'text';
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
      // C3 (audit fix, 2026-09-21): additive only — `text` above is still the exact same
      // flattened string as before; `format` is the new field, derived by toChatMessage
      // (messages.ts) from Graph's body.contentType. `m.format` is only optional on ChatMessage
      // for hand-built fixtures elsewhere that never reach this path; anything actually read
      // from Graph always has it set, but the ternary mirrors toChatMessage's own "not literally
      // html ⇒ text" rule rather than assuming that.
      format: m.format === 'html' ? 'html' : 'text',
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
  options: { name?: string; out?: string; yieldPath?: string } = {},
): Promise<{ action: 'attachments'; chat: string; messageId: string; count: number; files: DownloadedAttachment[] }> {
  const entry = allowlist.assertReadable(chatId);
  // yieldPath (attachments.ts passes the real one) asks any running inbox poller — usually the
  // daemon in another process — to go quiet while this CLI spends the shared per-mailbox Graph
  // read budget; without the yield, the poller starves ad-hoc reads outright (measured
  // 2026-09-02, see inbox-yield.ts). Only the Graph work is held under it; local file writes
  // need no quota.
  const payloads = await withQuotaYield(options.yieldPath, 'teams-attachments', () =>
    chats.getAttachments(chatId, messageId, options.name),
  );
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
  options: { yieldPath?: string } = {},
): Promise<{
  action: 'attachments-list';
  chat: string;
  messageId: string;
  count: number;
  attachments: Array<{ id: string; name?: string; contentType?: string; downloadable: boolean }>;
}> {
  const entry = allowlist.assertReadable(chatId);
  // Same quota yield as doDownloadAttachments above — the metadata read hits the same
  // throttle-prone family and usually runs right before a download.
  const attachments = await withQuotaYield(options.yieldPath, 'teams-attachments --list', () =>
    chats.listAttachments(chatId, messageId),
  );
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
 * Thrown by the C1 plain-text guard below (doPost/doReply/doEdit) when a plain (html === false)
 * body looks structured and no override was given. `run()` maps it to exit 2 — the same "bad
 * caller input" bucket as every other usage() refusal — NOT process.exit itself: doPost/doReply/
 * doEdit are called directly, in-process, by this file's own unit tests (no subprocess), so a
 * process.exit() here would kill the test runner instead of failing one assertion.
 */
export class PlainTextRefusedError extends Error {}

// A REAL markdown/pipe-table row: the line both STARTS and ENDS with a pipe (ignoring leading/
// trailing whitespace), e.g. "| a | b |" or "|---|---|" — narrower than "any line with two
// pipes" (fix round, 2026-09-22, BLOCKING MAJOR review round 1: a shell pipeline like
// "cat x | grep y | head" has two pipes on one line but is not a table row — it starts with
// "cat" and ends with "head", neither a pipe).
const TABLE_LIKE_LINE = /^[ \t]*\|.*\|[ \t]*$/m;
// A run of two or more newlines — Teams renders a blank-line-separated body as multiple
// paragraphs, itself a styling decision the skill says needs <div>&nbsp;</div> air, not a
// plain-text double linebreak.
const BLANK_LINE = /\n[ \t]*\r?\n/;
// A . ! ? only counts as a candidate sentence terminator when it is followed by whitespace or
// the end of the string. Fix round, 2026-09-22 (BLOCKING MAJOR, review round 1): counting every
// literal . ! ? character refused ordinary one-sentence status posts containing a version number
// ("v0.7.2"), a filename ("findings.md"), a URL, a decimal ("30.9 percent"), or a path — a
// decimal/semver/extension dot is by construction never followed by whitespace (the next
// character is a digit or letter), so this lookahead excludes those for free, with no maintained
// list. It does NOT exclude a genuine sentence-ending abbreviation ("e.g. ", "kl. ") — those ARE
// followed by whitespace, so they still match here; ABBREVIATION_TERMINATOR_CI/CS below (fix
// round, 2026-09-22, review round 2, MINOR 2) subtract the short fixed list of those out of the
// count separately, rather than trying to make this one regex do both jobs.
const SENTENCE_TERMINATOR = /[.!?](?=\s|$)/g;
// A short, deliberately non-exhaustive list of abbreviations whose final dot reads as a sentence
// terminator by the rule above (it IS followed by whitespace) but isn't one. Case-insensitive for
// "e.g."/"i.e."/Swedish "kl." (klockan, a time-of-day marker, e.g. "kl. 14.30"); "No." (as in
// "item No. 42") is deliberately case-SENSITIVE (capital N only) so the common word "no." ending
// an ordinary sentence ("I said no. Then I left.") is not silently exempted. This is a known,
// bounded gap, not a general abbreviation detector — MINOR 2, review round 2.
const ABBREVIATION_TERMINATOR_CI = /\b(?:e\.g|i\.e|kl)\.(?=\s|$)/gi;
const ABBREVIATION_TERMINATOR_CS = /\bNo\.(?=\s|$)/g;

function countAbbreviationTerminators(text: string): number {
  const ci = text.match(ABBREVIATION_TERMINATOR_CI)?.length ?? 0;
  const cs = text.match(ABBREVIATION_TERMINATOR_CS)?.length ?? 0;
  return ci + cs;
}
// Recall gap (fix round, 2026-09-22, review item 1): a body can be structured with NO
// terminators and NO blank line at all — a bulleted list (one short line per item, nothing to
// end a "sentence") or a very long single paragraph. Thresholds picked well clear of the
// reviewer's six one-liner repro bodies (max 59 characters, 1 line each) and well under a
// realistic long paragraph (a 200-word paragraph runs ~1500 characters) — see the corpus test
// for both boundaries exercised together.
const MANY_LINES_THRESHOLD = 4;
const LONG_BODY_THRESHOLD = 400;

/**
 * Classifies a plain-text body against the teams-styling skill's own threshold ("more than two
 * sentences ⇒ not plain text"): three or more sentence terminators (. ! ?, each followed by
 * whitespace or end-of-string, MINUS the short fixed list of abbreviation-final dots that shape
 * also matches — see ABBREVIATION_TERMINATOR_CI/CS above), OR any blank line, OR a real
 * pipe-table-looking line, OR the body is long/many-lined enough that it is obviously not a short
 * conversational reply even with none of the above. Returns the reason phrase to quote in the
 * refusal, or undefined when the body is short/plain enough to send as-is. C1 (audit fix,
 * 2026-09-21; classifier corrected in the 2026-09-22 fix rounds after review found the naive
 * terminator count over-triggered on ordinary operational text, and on ordinary abbreviations) —
 * measured 31% styling compliance, traced to plain text being the tool's silent default for
 * bodies exactly this shape.
 */
export function structuredTextReason(text: string): string | undefined {
  const rawTerminators = text.match(SENTENCE_TERMINATOR)?.length ?? 0;
  const terminators = Math.max(0, rawTerminators - countAbbreviationTerminators(text));
  if (terminators >= 3) {
    return `it reads as ${terminators} sentences — more than two is not plain text, per the teams-styling skill`;
  }
  if (BLANK_LINE.test(text)) {
    return 'it has a blank line';
  }
  if (TABLE_LIKE_LINE.test(text)) {
    return 'it has a pipe-table-looking line';
  }
  const nonEmptyLines = text.split('\n').filter((line) => line.trim() !== '').length;
  if (nonEmptyLines >= MANY_LINES_THRESHOLD) {
    return `it has ${nonEmptyLines} lines — that reads as a list, not a short reply`;
  }
  if (text.length > LONG_BODY_THRESHOLD) {
    return `it is ${text.length} characters long — too long for a short reply`;
  }
  return undefined;
}

/**
 * The C1 guard itself, shared by doPost/doReply/doEdit below. Only applies on the plain-text
 * path (html === false) with no --text override — --html content is never plain text in the
 * first place, and --text is the caller's deliberate "send it anyway" (audit fix C1,
 * 2026-09-21: CLI entry points only, per the brief — server.ts's MCP send tools have their own,
 * separate implementation and are untouched). `usageStyled`/`usageOverride` are the exact
 * corrected commands to show the caller: (b) chat id FIRST, then --html — the C2 trap this skips
 * around entirely by never taking a fixed argv position — and (c) the deliberate override.
 */
function assertPlainTextAllowed(
  cli: string,
  usageStyled: string,
  usageOverride: string,
  text: string,
  html: boolean,
  override: boolean,
): void {
  if (html || override) {
    return;
  }
  const reason = structuredTextReason(text);
  if (!reason) {
    return;
  }
  throw new PlainTextRefusedError(
    `${cli}: refusing to send this as plain text — ${reason}. Read the teams-styling skill before ` +
      `composing a message like this. Resend styled: "${usageStyled}" (the chat id comes first, ` +
      `then --html). To send this exact text as plain anyway: "${usageOverride}".`,
  );
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
  options: { plainTextOverride?: boolean } = {},
): Promise<{ action: 'post'; id: string; chat: string }> {
  // Allowlist gate FIRST (fix round, 2026-09-22, review item 2): the README's exit-code contract
  // documents 3 (allowlist) as a distinct, prior failure class from 2 (usage, including this
  // guard) — a chat that fails the allowlist gate must exit 3 regardless of what its body looks
  // like, not be shadowed by a plain-text refusal that never even reached a chat Graph will
  // refuse anyway.
  const entry = allowlist.assertPostable(chatId);
  assertPlainTextAllowed(
    'teams-post',
    `teams-post ${chatId} --html`,
    `teams-post ${chatId} --text`,
    text,
    html,
    options.plainTextOverride ?? false,
  );
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
  options: { plainTextOverride?: boolean } = {},
): Promise<{ action: 'edit'; id: string; chat: string }> {
  // Allowlist gate FIRST — same ordering rationale as doPost above.
  const entry = allowlist.assertPostable(chatId);
  assertPlainTextAllowed(
    'teams-edit',
    `teams-edit ${chatId} ${messageId} --html`,
    `teams-edit ${chatId} ${messageId} --text`,
    newText,
    html,
    options.plainTextOverride ?? false,
  );
  const resolved = await resolveMentions(chats, chatId, mentions);
  if (html) {
    await chats.editHtmlMessage(chatId, messageId, newText, resolved);
  } else {
    await chats.editMessage(chatId, messageId, newText, resolved);
  }
  return { action: 'edit', id: messageId, chat: entry.label };
}

/** teams-reply's --html/--mention plumbing — same rationale as doPost/doEdit above (testable
 *  without a live send; a subprocess test can't distinguish "the html branch runs" from "the
 *  flag was ignored", because assertPostable always throws (or not) before either reply path is
 *  ever reached). `html` gained reply --html parity, 0.6.x — same verbatim raw-HTML contract as
 *  doPost's html branch, applied to replyToHtmlMessage instead of sendHtmlMessage. */
export async function doReply(
  { chats, allowlist }: CliContext,
  chatId: string,
  replyToMessageId: string,
  text: string,
  html: boolean,
  mentions: readonly string[] = [],
  options: { plainTextOverride?: boolean } = {},
): Promise<{ action: 'reply'; id: string; inReplyTo: string; chat: string }> {
  // Allowlist gate FIRST — same ordering rationale as doPost above.
  const entry = allowlist.assertPostable(chatId);
  assertPlainTextAllowed(
    'teams-reply',
    `teams-reply ${chatId} ${replyToMessageId} --html`,
    `teams-reply ${chatId} ${replyToMessageId} --text`,
    text,
    html,
    options.plainTextOverride ?? false,
  );
  const resolved = await resolveMentions(chats, chatId, mentions);
  const sent = html
    ? await chats.replyToHtmlMessage(chatId, replyToMessageId, text, resolved)
    : await chats.replyToMessage(chatId, replyToMessageId, text, resolved);
  return { action: 'reply', id: sent.id, inReplyTo: replyToMessageId, chat: entry.label };
}

/** One send-file success payload — the shape `onSent` (doSendFile below) delivers, and
 *  `writeLine` (below) turns into one JSON stdout line. `granted`/`note` (0.6.0, live
 *  2026-09-08): the CLI's own "says so on stdout" contract for `--no-grant` — this JSON line IS
 *  the only output surface doSendFile has, so a caller reading it (human or orchestrator) sees
 *  `granted: false` and the reason without having to know which flag produced it. `granted` is
 *  `true` for both the default roster-derived grant and an explicit `--grant-to` — the caller who
 *  gave `--grant-to` already knows who they told sendFile to grant to. */
export interface SendFileResult {
  action: 'send-file';
  id: string;
  chat: string;
  name: string;
  bytes: number;
  granted: boolean;
  note?: string;
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
  sendOptions: { grantTo?: readonly string[]; noGrant?: boolean } = {},
): Promise<void> {
  const entry = allowlist.assertPostable(chatId);
  const options = sendOptions.noGrant
    ? { noGrant: true as const }
    : sendOptions.grantTo !== undefined
      ? { grantTo: sendOptions.grantTo }
      : undefined;
  for (const [index, filePath] of paths.entries()) {
    const buffer = await readFile(filePath);
    const name = basename(filePath);
    const sent = await chats.sendFile(
      chatId,
      { bytes: new Uint8Array(buffer), name },
      index === 0 ? caption : undefined,
      options,
    );
    await onSent({
      action: 'send-file',
      id: sent.id,
      chat: entry.label,
      name,
      bytes: buffer.byteLength,
      granted: !sendOptions.noGrant,
      ...(sendOptions.noGrant
        ? { note: 'uploaded without a permission grant (--no-grant) — only the sender can open this file' }
        : {}),
    });
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
    // FINDING 4 (message-withdrawal review): a MessageOwnershipError — teams-delete's own-message
    // gate refusing somebody else's (or an unverifiable) message — is its own exit code, distinct
    // from both the allowlist refusal (3) and the generic 1 catch-all, so a caller can branch on
    // "the ownership check refused this" without parsing stderr text. A throttled /me during that
    // same gate is NOT this: it is the package's ordinary 429-shaped GraphError and exits 1, same
    // as any other Graph failure — see assertOwnMessage's own doc comment (teams-chats.ts) for why
    // the two are deliberately different exception types. PlainTextRefusedError (C1, 2026-09-21)
    // shares exit 2 with usage() — it is the same class of problem (bad caller input, no network
    // reached), not a new code.
    process.exit(
      caught instanceof ChatNotAllowedError
        ? 3
        : caught instanceof MessageOwnershipError
          ? 4
          : caught instanceof PlainTextRefusedError
            ? 2
            : 1,
    );
  }
}
