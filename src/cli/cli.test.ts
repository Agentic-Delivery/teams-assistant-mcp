import { execFile } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { buildChats } from '../build-chats.js';
import { ChatAllowlist, ChatNotAllowedError } from '../allowlist.js';
import { GraphClient, GraphError } from '../graph/graph-client.js';
import { MembersCache } from '../graph/members-cache.js';
import { ReliableTeamsChats } from '../graph/reliable-sends.js';
import { GraphTeamsChats, MessageOwnershipError, type TeamsChatsPort } from '../graph/teams-chats.js';
import type { ChatMessage, ReadResult } from '../messages.js';
import { loadConfig } from '../config.js';
import {
  doDelete,
  doDownloadAttachments,
  doEdit,
  doListAttachments,
  doPin,
  doPost,
  doRead,
  doReply,
  doSendFile,
  parseAttachmentFlags,
  parseDeleteFlags,
  parseSendFileFlags,
  parseSendFlags,
  PlainTextRefusedError,
  run,
  structuredTextReason,
  succeed,
  writeLine,
} from './common.js';

const repoRoot = join(import.meta.dirname, '..', '..');
const tsx = join(repoRoot, 'node_modules', '.bin', 'tsx');

interface CliRun {
  code: number;
  stdout: string;
  stderr: string;
}

// `stdin` defaults to a short, guard-safe body ("the message text" — zero sentence terminators,
// no blank line, no pipe-table line) so every existing caller keeps getting the plain send it
// always got; the C1 plain-text-guard tests below pass a deliberately structured body instead.
function runCli(
  script: string,
  args: string[],
  env: Record<string, string>,
  stdin = 'the message text',
): Promise<CliRun> {
  return new Promise((resolve) => {
    const child = execFile(
      tsx,
      [join(repoRoot, 'src', 'cli', script), ...args],
      { env: { ...process.env, ...env }, cwd: repoRoot },
      (error, stdout, stderr) => {
        resolve({
          code: (error as { code?: number } | null)?.code ?? 0,
          stdout: String(stdout),
          stderr: String(stderr),
        });
      },
    );
    child.stdin?.end(stdin);
  });
}

function fixtureEnv(): Record<string, string> {
  const dir = mkdtempSync(join(tmpdir(), 'teams-cli-test-'));
  const configPath = join(dir, 'config.json');
  writeFileSync(
    configPath,
    JSON.stringify({
      assistantDisplayName: 'Assistant',
      allowedChats: [
        { id: '19:readonly@thread.v2', label: 'read-only chat', canPost: false },
        // Postable (2026-09-22 fix round, review item 2): the C1 guard now runs AFTER
        // allowlist.assertPostable, per the README's exit-code contract (3 before 2) — a
        // subprocess test proving the GUARD fires (exit 2) needs a chat that clears the
        // allowlist gate first, or it would only ever prove the allowlist gate fires (exit 3).
        { id: '19:postable@thread.v2', label: 'postable chat', canPost: true },
      ],
    }),
  );
  return {
    TEAMS_MCP_CONFIG: configPath,
    TEAMS_MCP_TENANT_ID: 'tenant',
    TEAMS_MCP_USERNAME: 'user@example.test',
    TEAMS_MCP_PASSWORD: 'not-a-real-password',
    TEAMS_MCP_TOKEN_CACHE: join(dir, 'token-cache.json'),
  };
}

describe('the composition both entry points share', () => {
  it('buildChats puts ReliableTeamsChats in front of every consumer', () => {
    // README: "same code paths as the server tools — including the send reliability."
    // This is the test that fails if a refactor ever drops the readback decorator.
    const env = fixtureEnv();
    const config = loadConfig(env as NodeJS.ProcessEnv);

    expect(buildChats(config).chats).toBeInstanceOf(ReliableTeamsChats);
  });
});

describe('the CLI contract — exit codes, and nothing but the JSON line on stdout', () => {
  // These paths need no network: usage fails before config, allowlist fails before any token.
  it('missing arguments: exit 2, stdout empty', async () => {
    const result = await runCli('post.ts', [], {});

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/usage/);
  });

  it('a chat outside the allowlist: exit 3, stdout empty', async () => {
    const result = await runCli('post.ts', ['19:never-heard-of@thread.v2'], fixtureEnv());

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
  });

  it('an allowlisted chat without canPost: exit 3, stdout empty', async () => {
    const result = await runCli('post.ts', ['19:readonly@thread.v2'], fixtureEnv());

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
  });

  it('teams-read refuses --since without a value: exit 2', async () => {
    const result = await runCli('read.ts', ['19:readonly@thread.v2', '--since'], fixtureEnv());

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
  });

  // Review round 2, MINOR 7 (2026-08-26): `--mention --html` used to consume "--html" as the
  // mention NAME (silently, no error here) and only fail later, confusingly, when resolveMentions
  // couldn't find a chat member called "--html". A flag-like value is refused loudly at parse time
  // instead.
  it('teams-post rejects a flag-like value for --mention instead of silently accepting it as a name: exit 2', async () => {
    const result = await runCli(
      'post.ts',
      ['19:readonly@thread.v2', '--mention', '--html'],
      fixtureEnv(),
    );

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/--mention/);
  });

  it('teams-post --mention at the end of argv with nothing after it: exit 2, same as any missing value', async () => {
    const result = await runCli('post.ts', ['19:readonly@thread.v2', '--mention'], fixtureEnv());

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/--mention/);
  });

  it('teams-post --html: still gated by the allowlist exactly like plain text (exit 3, no network reached)', async () => {
    // This proves --html does not get swallowed as an extra positional (which would misparse
    // chatId or trip a usage error) — it does NOT prove the flag routes to sendHtmlMessage,
    // since assertPostable throws before either send path is ever reached either way. That
    // routing is proven in-process, without a subprocess, in the describe block below.
    const result = await runCli(
      'post.ts',
      ['19:readonly@thread.v2', '--html'],
      fixtureEnv(),
    );

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
  });

  it('teams-edit missing arguments: exit 2, stdout empty', async () => {
    const result = await runCli('edit.ts', [], {});

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/usage/);
  });

  it('teams-edit missing messageId: exit 2, stdout empty', async () => {
    const result = await runCli('edit.ts', ['19:readonly@thread.v2'], {});

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
  });

  it('teams-edit a chat outside the allowlist: exit 3, stdout empty', async () => {
    const result = await runCli(
      'edit.ts',
      ['19:never-heard-of@thread.v2', 'msg-1'],
      fixtureEnv(),
    );

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
  });

  it('teams-edit an allowlisted chat without canPost: exit 3, stdout empty (same gate as teams-post)', async () => {
    const result = await runCli('edit.ts', ['19:readonly@thread.v2', 'msg-1'], fixtureEnv());

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
  });

  it('teams-edit --html: reaches the same allowlist gate as plain text (exit 3, no network reached) — routing itself is proven in-process below', async () => {
    const result = await runCli(
      'edit.ts',
      ['19:readonly@thread.v2', 'msg-1', '--html'],
      fixtureEnv(),
    );

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
  });

  it('teams-reply --mention: still gated by the allowlist exactly like a plain reply (exit 3, no network reached)', async () => {
    // Same shape as the --html test above: proves --mention/"Mika" are parsed as a flag+value
    // pair, not misread as positionals, without needing a live resolveMentions call. The
    // resolution/forwarding itself is proven in-process below (doReply).
    const result = await runCli(
      'reply.ts',
      ['19:readonly@thread.v2', 'msg-1', '--mention', 'Mika'],
      fixtureEnv(),
    );

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
  });

  // reply --html parity (0.6.x): teams-reply gained --html (see the "reply --html" describe
  // block below for the routing proof and the doReply tests above for the wiring) — same
  // subprocess-can't-prove-routing reasoning as teams-edit --html above, so this only proves the
  // flag is parsed as a bare flag (not swallowing the next positional) and reaches the same gate.
  it('teams-reply --html: reaches the same allowlist gate as plain text (exit 3, no network reached) — routing itself is proven in-process below', async () => {
    const result = await runCli(
      'reply.ts',
      ['19:readonly@thread.v2', 'msg-1', '--html'],
      fixtureEnv(),
    );

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
  });

  it('teams-reply rejects an unrecognised leftover argument instead of silently ignoring it: exit 2', async () => {
    const result = await runCli(
      'reply.ts',
      ['19:readonly@thread.v2', 'msg-1', '--bogus', 'junk'],
      fixtureEnv(),
    );

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
  });
});

// C2 (audit fix, 2026-09-21): chatId used to be taken as the fixed positional argv[2], BEFORE
// any flag parsing ran — `post.mjs --html <chat>` consumed "--html" itself as the chat id and
// the real chat id was silently dropped, producing a misleading allowlist error that named
// "--html", not the chat the caller meant. Flags now parse across the whole argv first; the
// first surviving positional is the chat id, in whichever position the caller put it.
describe('C2 — flags parse anywhere in argv; a flag-like resolved chat id is refused', () => {
  it('C2a: teams-post --html <chatId> resolves the REAL chat id, not "--html" (stderr names the real chat)', async () => {
    const result = await runCli('post.ts', ['--html', '19:readonly@thread.v2'], fixtureEnv());

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
    expect(result.stderr).toContain('19:readonly@thread.v2');
    expect(result.stderr).not.toContain('Chat --html');
  });

  it('C2b: teams-reply --html <chatId> <messageId> resolves the REAL chat id the same way', async () => {
    const result = await runCli(
      'reply.ts',
      ['--html', '19:readonly@thread.v2', 'msg-1'],
      fixtureEnv(),
    );

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
    expect(result.stderr).toContain('19:readonly@thread.v2');
    expect(result.stderr).not.toContain('Chat --html');
  });

  it('C2c: teams-edit --html <chatId> <messageId> resolves the REAL chat id the same way', async () => {
    const result = await runCli(
      'edit.ts',
      ['--html', '19:readonly@thread.v2', 'msg-1'],
      fixtureEnv(),
    );

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
    expect(result.stderr).toContain('19:readonly@thread.v2');
    expect(result.stderr).not.toContain('Chat --html');
  });

  it('C2d: teams-send-file --caption "hi" <chatId> <path> resolves the REAL chat id the same way', async () => {
    const result = await runCli(
      'send-file.ts',
      ['--caption', 'hi', '19:readonly@thread.v2', '/tmp/does-not-matter.txt'],
      fixtureEnv(),
    );

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
    expect(result.stderr).toContain('19:readonly@thread.v2');
  });

  // Non-triggering side: ordinary "chat id first" invocations are completely unaffected.
  it('C2e: teams-post <chatId> --html still works exactly as before (order does not matter)', async () => {
    const result = await runCli('post.ts', ['19:readonly@thread.v2', '--html'], fixtureEnv());

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
    expect(result.stderr).toContain('19:readonly@thread.v2');
  });

  it('C2f: teams-post --bogus (an unrecognised flag left with nothing after it) refuses naming the flag-first mistake', async () => {
    const result = await runCli('post.ts', ['--bogus'], fixtureEnv());

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toBe('that looks like a flag; the chat id comes first\n');
  });

  it('C2g: teams-reply --bogus msg-1 — same refusal on the reply CLI', async () => {
    const result = await runCli('reply.ts', ['--bogus', 'msg-1'], fixtureEnv());

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toBe('that looks like a flag; the chat id comes first\n');
  });

  // No send-file equivalent of C2f/C2g: parseSendFileFlags (unlike parseSendFlags) already
  // refuses any UNRECOGNISED --flag outright, immediately, regardless of position (see its own
  // doc comment) — so an unrecognised flag can never reach send-file.ts's positional-resolution
  // step in the first place to be mistaken for the chat id. A RECOGNISED flag before the chat id
  // (--caption, --grant-to, --no-grant) is exactly the C2d shape above, already covered.
});

// C1 (audit fix, 2026-09-21, the audit's highest-leverage fix): the highest-measured failure was
// plain text going out for bodies that should have been styled — 31% overall compliance, and a
// 4.8x gap between sessions that had loaded the teams-styling skill and ones that hadn't. This
// guard makes the refusal happen at the one place guaranteed to run regardless of whether the
// skill was ever read: the CLI send path itself.
describe('C1 — plain-text guard (subprocess: exit codes and the refusal message)', () => {
  const threeSentences = 'First sentence. Second sentence. Third sentence.';
  const blankLineBody = 'First line.\n\nSecond line.';
  const tableLikeBody = '| a | b |\n| 1 | 2 |';
  // Postable (see fixtureEnv) — the fix-round reorder (review item 2) means the guard only
  // fires after assertPostable passes, so every test that means to PROVE THE GUARD fires needs
  // a chat that clears the allowlist gate first.
  const postable = '19:postable@thread.v2';

  it('C1a: a plain 3-sentence body with no --html/--text is refused, exit 2', async () => {
    const result = await runCli('post.ts', [postable], fixtureEnv(), threeSentences);

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/teams-styling/);
    // (b) the correct invocation shape — chat id first, then --html — not the trap order.
    expect(result.stderr).toMatch(/teams-post 19:postable@thread\.v2 --html/);
    expect(result.stderr).not.toMatch(/--html 19:postable@thread\.v2/);
    // (c) the deliberate override is named.
    expect(result.stderr).toMatch(/teams-post 19:postable@thread\.v2 --text/);
  });

  // C1b and C1c were DELETED here (review round 2, BLOCKER 2): both ran against the read-only
  // chat, which exits 3 at allowlist.assertPostable REGARDLESS of what the guard/override do —
  // a classifier that refuses EVERYTHING, or an override that does nothing at all, would have
  // left both green (reviewer-verified: mutation M-B, "refuse everything", leaves them green).
  // The "2-sentence body is not refused" claim is carried by the structuredTextReason corpus
  // test below (the exact 'One sentence. Two sentences.' row); the "--text overrides the guard"
  // claim is carried by the in-process 'doPost: html=false, plainTextOverride=true...' test
  // above, which uses the POSTABLE chat and an unwrapped 3-sentence body — a real trigger the
  // override must actually suppress, not a chat that would have exited 3 either way.

  it('C1d: --html and --text together is an error, exit 2', async () => {
    const result = await runCli(
      'post.ts',
      ['19:readonly@thread.v2', '--html', '--text'],
      fixtureEnv(),
    );

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/mutually exclusive/);
  });

  // C1e was DELETED here (review round 2, BLOCKER 2): same "unconditional allowlist exit" flaw
  // as C1b/C1c above, for the "--html bypasses the guard" claim — that claim is now carried by
  // the in-process 'doPost/doReply/doEdit: html=true, body whose terminators ARE followed by
  // whitespace...' tests (see the describe block above this one), which use a body that would
  // genuinely trip the guard were html not bypassing it, and assert the html send method itself
  // ran — reviewer-verified against mutation M-G (`if (html || override)` -> `if (override)`).

  it('C1f: a body with a blank line is refused even with fewer than 3 sentence terminators', async () => {
    const result = await runCli('post.ts', [postable], fixtureEnv(), blankLineBody);

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/blank line/);
  });

  it('C1g: a body with a REAL pipe-table-looking line is refused', async () => {
    const result = await runCli('post.ts', [postable], fixtureEnv(), tableLikeBody);

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/pipe-table/);
  });

  it('C1h: teams-reply applies the same guard, naming teams-reply\'s own invocation shape', async () => {
    const result = await runCli('reply.ts', [postable, 'msg-1'], fixtureEnv(), threeSentences);

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/teams-styling/);
    expect(result.stderr).toMatch(/teams-reply 19:postable@thread\.v2 msg-1 --html/);
    expect(result.stderr).toMatch(/teams-reply 19:postable@thread\.v2 msg-1 --text/);
  });

  it('C1i: teams-edit applies the same guard, naming teams-edit\'s own invocation shape', async () => {
    const result = await runCli('edit.ts', [postable, 'msg-1'], fixtureEnv(), threeSentences);

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/teams-styling/);
    expect(result.stderr).toMatch(/teams-edit 19:postable@thread\.v2 msg-1 --html/);
    expect(result.stderr).toMatch(/teams-edit 19:postable@thread\.v2 msg-1 --text/);
  });

  // C1j was DELETED here (review round 2, BLOCKER 2, reviewer evidence): all 6 rows ran the
  // reviewer's bodies against the READ-ONLY chat and asserted exit 3 — which
  // allowlist.assertPostable produces UNCONDITIONALLY for that chat, before the guard ever runs.
  // Reverting to the naive round-1 terminator count (the exact MAJOR this was meant to catch)
  // left every C1j row green; only the structuredTextReason unit corpus rows below went red. The
  // six reviewer bodies are pinned there instead — as `notStructured` rows in the
  // 'structuredTextReason — the plain-text guard's classifier' describe block further down —
  // where the classifier's actual return value is asserted directly, not shadowed by an
  // allowlist gate that would produce the same exit code however the classifier behaved. (A
  // postable-chat + "code !== 2 && no /refusing to send/" version was considered and rejected:
  // it requires a real Graph/token-acquisition attempt with fake credentials past the guard,
  // which is slow and non-deterministic in a test environment — see this file's own
  // "no network reached" doctrine throughout the rest of this describe block.)

  // C1k (fix round, review item 2): the guard now runs AFTER allowlist.assertPostable, per the
  // README's documented exit-code contract (3 before 2) — a chat that fails the allowlist gate
  // must exit 3 even when its body would ALSO have tripped the plain-text guard; the guard never
  // gets the chance to run.
  it('C1k: a non-postable chat with a structured body exits 3 (allowlist), not 2 (the guard never runs)', async () => {
    const result = await runCli('post.ts', ['19:readonly@thread.v2'], fixtureEnv(), threeSentences);

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
  });
});

describe('structuredTextReason — the plain-text guard\'s classifier (C1, fix round 2026-09-22)', () => {
  // Table-driven corpus (review item 3, root cause of the MAJOR): a plain assertion per case
  // would not have caught the naive-terminator-count bug, because none of the ORIGINAL tests
  // exercised a realistic body containing a version number/URL/path/decimal/filename — every
  // case here is a REAL message shape, not a synthetic "One. Two. Three." Each row states the
  // reason a human would give for the expected verdict, so a future classifier tweak that
  // breaks one of these fails with that reason visible in the test name.
  const notStructured: Array<[string, string]> = [
    ['no terminator here', 'no terminator, no blank line, no table line'],
    ['One sentence.', 'a single real sentence'],
    ['One sentence. Two sentences.', 'exactly two real sentences — the skill\'s own threshold'],
    ['Shipped teams-assistant-mcp v0.7.2 to the server.', 'a version number\'s dots are not sentence ends'],
    ['See findings.md and notes.md in the workspace.', 'filenames\' dots are not sentence ends'],
    ['The build is at https://dev.azure.com/if/CTP/_build?id=42.', 'a URL\'s dots are not sentence ends'],
    ['Compliance moved from 30.9 percent to 48.0 percent.', 'decimal points are not sentence ends'],
    ['Config lives at /home/johan/.claude/settings.json and .env.', 'a path\'s dots are not sentence ends'],
    ['Run: cat x | grep y | head', 'a shell pipeline is not a table row (no leading/trailing pipe)'],
    ['a | b', 'a single mid-line pipe is not a table row'],
    ['See e.g. the notes.', 'a lone abbreviation does not by itself reach the 3-terminator threshold'],
    ['It costs $3.50 per unit.', 'a currency decimal is not a sentence end'],
    ['Node v20.11.0 shipped.', 'a semver\'s dots are not sentence ends'],
    // MINOR 2 (review round 2): abbreviation-final dots ARE followed by whitespace, so the base
    // lookahead alone counts them as real terminators — these would each be 3 raw terminators
    // without the explicit exemption list below, and 2 with it.
    ['Send it e.g. tomorrow morning. Thanks.', '"e.g." is on the abbreviation exemption list'],
    ['Note i.e. this matters a lot. OK.', '"i.e." is on the abbreviation exemption list'],
    ['Meet kl. 14.30 today. See you.', 'Swedish "kl." (klockan) is on the abbreviation exemption list'],
    ['See No. 42 in the list. Thanks.', '"No." (capital N) is on the abbreviation exemption list'],
  ];
  it.each(notStructured)('NOT structured: %j (%s)', (text) => {
    expect(structuredTextReason(text)).toBeUndefined();
  });

  const structured: Array<[string, string]> = [
    ['One. Two. Three.', 'three real sentence terminators, each followed by whitespace'],
    ['Really? Yes! Confirmed.', 'a mix of . ! ? still counts toward the 3-terminator threshold'],
    ['Done! Ready for review? Yes.', 'three genuinely separate real sentences'],
    ['first line\n\nsecond line', 'a blank line, even with zero sentence terminators'],
    ['| a | b |', 'a real single-row table line (starts AND ends with a pipe)'],
    ['| a | b |\n| 1 | 2 |', 'a real two-row markdown table'],
    [
      Array.from({ length: 10 }, (_, i) => `- item ${i}`).join('\n'),
      'a 10-line bulleted list with no terminators and no blank lines at all (recall gap, review item 1)',
    ],
    [
      Array.from({ length: 200 }, (_, i) => `word${i}`).join(' ') + '.',
      'a 200-word single paragraph with only one terminator (recall gap, review item 1)',
    ],
    // MINOR 2 (review round 2): the exemption list must not defeat real detection — a body with
    // an exempted abbreviation PLUS enough genuine sentences to still cross the threshold must
    // still refuse.
    [
      'See e.g. the file. It works well. Great job. Thanks.',
      'one exempted "e.g." plus four real sentences — still well over the threshold',
    ],
    // The deliberate case-sensitivity boundary on "No." (capital N only): the common lowercase
    // word "no." ending an ordinary sentence must NOT be silently exempted.
    [
      'I said no. Then I left. It happened fast.',
      'lowercase "no." is an ordinary word, not the "No." abbreviation — three real sentences',
    ],
  ];
  it.each(structured)('structured: %j (%s)', (text) => {
    expect(structuredTextReason(text)).toBeDefined();
  });
});

describe('doPost / doReply / doEdit — the plain-text guard (in-process, no subprocess, no network — C1)', () => {
  function fakePort(overrides: Partial<TeamsChatsPort>): TeamsChatsPort {
    const reject = () => Promise.reject(new Error('not part of this test'));
    return {
      listChats: reject,
      readMessages: async () => ({ messages: [] }) as unknown as ReadResult,
      resolveMentions: reject,
      sendMessage: reject,
      sendHtmlMessage: reject,
      sendImage: reject,
      sendFile: reject,
      replyToMessage: reject,
      replyToHtmlMessage: reject,
      editMessage: reject,
      editHtmlMessage: reject,
      deleteMessage: reject,
      undoDeleteMessage: reject,
      setReaction: reject,
      getAttachment: reject,
      listAttachments: reject,
      getAttachments: reject,
      pinMessage: reject,
      unpinMessage: reject,
      listPinnedMessages: reject,
      ...overrides,
    } as TeamsChatsPort;
  }

  const allowlist = new ChatAllowlist([{ id: '19:a@thread.v2', label: 'chat A', canPost: true }]);
  const structured = 'First sentence. Second sentence. Third sentence.';
  // BLOCKER 1 (review round 2): `<p>${structured}</p>` (used below) is a WEAK html=true repro —
  // its third terminator is followed by "</p>" (the closing tag), not whitespace, so it never
  // reaches the 3-terminator threshold in the first place; mutating `if (html || override)` to
  // `if (override)` (--html no longer bypasses the guard) left that test green by accident. This
  // body's three terminators are each followed by a real space, so it WOULD trip the guard were
  // html not bypassing it — the mutation must turn this test red.
  const structuredWithWhitespaceTerminators = '<div>One. Two. Three. </div>';

  it('doPost: html=false, no override, structured text — rejects with PlainTextRefusedError BEFORE any send', async () => {
    const sendMessage = vi.fn();
    const sendHtmlMessage = vi.fn();
    const chats = new ReliableTeamsChats(fakePort({ sendMessage, sendHtmlMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    await expect(doPost({ chats, allowlist }, '19:a@thread.v2', structured, false)).rejects.toThrow(
      PlainTextRefusedError,
    );
    expect(sendMessage).not.toHaveBeenCalled();
    expect(sendHtmlMessage).not.toHaveBeenCalled();
  });

  it('doPost: html=true, structured text — the guard does not apply, sendHtmlMessage still runs', async () => {
    const sendHtmlMessage = vi.fn(async () => ({
      id: 'm1',
      chatId: '19:a@thread.v2',
      createdDateTime: '2026-09-21T10:00:00Z',
      from: 'Assistant',
      text: '',
      isDeleted: false,
      attachments: [],
    }));
    const chats = new ReliableTeamsChats(fakePort({ sendHtmlMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    const result = await doPost({ chats, allowlist }, '19:a@thread.v2', `<p>${structured}</p>`, true);

    expect(sendHtmlMessage).toHaveBeenCalledWith('19:a@thread.v2', `<p>${structured}</p>`, []);
    expect(result).toEqual({ action: 'post', id: 'm1', chat: 'chat A' });
  });

  // BLOCKER 1 fix (review round 2): the load-bearing case the test above accidentally could not
  // catch — see structuredWithWhitespaceTerminators' own comment. Kill verified against mutation
  // M-G (`if (html || override)` -> `if (override)`) below.
  it('doPost: html=true, body whose terminators ARE followed by whitespace — the guard still does not apply, sendHtmlMessage still runs', async () => {
    const sendHtmlMessage = vi.fn(async () => ({
      id: 'm1b',
      chatId: '19:a@thread.v2',
      createdDateTime: '2026-09-22T10:00:00Z',
      from: 'Assistant',
      text: '',
      isDeleted: false,
      attachments: [],
    }));
    const chats = new ReliableTeamsChats(fakePort({ sendHtmlMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    const result = await doPost(
      { chats, allowlist },
      '19:a@thread.v2',
      structuredWithWhitespaceTerminators,
      true,
    );

    expect(sendHtmlMessage).toHaveBeenCalledWith(
      '19:a@thread.v2',
      structuredWithWhitespaceTerminators,
      [],
    );
    expect(result).toEqual({ action: 'post', id: 'm1b', chat: 'chat A' });
  });

  it('doPost: html=false, plainTextOverride=true, structured text — the override bypasses the guard, sendMessage still runs', async () => {
    const sendMessage = vi.fn(async () => ({
      id: 'm2',
      chatId: '19:a@thread.v2',
      createdDateTime: '2026-09-21T10:00:00Z',
      from: 'Assistant',
      text: '',
      isDeleted: false,
      attachments: [],
    }));
    const chats = new ReliableTeamsChats(fakePort({ sendMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    const result = await doPost({ chats, allowlist }, '19:a@thread.v2', structured, false, [], {
      plainTextOverride: true,
    });

    expect(sendMessage).toHaveBeenCalledWith('19:a@thread.v2', structured, []);
    expect(result).toEqual({ action: 'post', id: 'm2', chat: 'chat A' });
  });

  it('doReply: html=false, no override, structured text — rejects with PlainTextRefusedError BEFORE any send', async () => {
    const replyToMessage = vi.fn();
    const chats = new ReliableTeamsChats(fakePort({ replyToMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    await expect(
      doReply({ chats, allowlist }, '19:a@thread.v2', 'orig-1', structured, false),
    ).rejects.toThrow(PlainTextRefusedError);
    expect(replyToMessage).not.toHaveBeenCalled();
  });

  // BLOCKER 1 fix (review round 2): same whitespace-terminator repro as doPost above, on the
  // reply path.
  it('doReply: html=true, body whose terminators ARE followed by whitespace — the guard still does not apply, replyToHtmlMessage still runs', async () => {
    const replyToHtmlMessage = vi.fn(async () => ({
      id: 'r1b',
      chatId: '19:a@thread.v2',
      createdDateTime: '2026-09-22T10:00:00Z',
      from: 'Assistant',
      text: '',
      isDeleted: false,
      attachments: [],
    }));
    const chats = new ReliableTeamsChats(fakePort({ replyToHtmlMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    const result = await doReply(
      { chats, allowlist },
      '19:a@thread.v2',
      'orig-1',
      structuredWithWhitespaceTerminators,
      true,
    );

    expect(replyToHtmlMessage).toHaveBeenCalledWith(
      '19:a@thread.v2',
      'orig-1',
      structuredWithWhitespaceTerminators,
      [],
    );
    expect(result).toEqual({ action: 'reply', id: 'r1b', inReplyTo: 'orig-1', chat: 'chat A' });
  });

  it('doEdit: html=false, no override, structured text — rejects with PlainTextRefusedError BEFORE any send', async () => {
    const editMessage = vi.fn();
    const chats = new ReliableTeamsChats(fakePort({ editMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    await expect(
      doEdit({ chats, allowlist }, '19:a@thread.v2', 'msg-1', structured, false),
    ).rejects.toThrow(PlainTextRefusedError);
    expect(editMessage).not.toHaveBeenCalled();
  });

  // BLOCKER 1 fix (review round 2): same whitespace-terminator repro as doPost above, on the
  // edit path.
  it('doEdit: html=true, body whose terminators ARE followed by whitespace — the guard still does not apply, editHtmlMessage still runs', async () => {
    const editHtmlMessage = vi.fn(async () => undefined);
    const chats = new ReliableTeamsChats(fakePort({ editHtmlMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    const result = await doEdit(
      { chats, allowlist },
      '19:a@thread.v2',
      'msg-1',
      structuredWithWhitespaceTerminators,
      true,
    );

    expect(editHtmlMessage).toHaveBeenCalledWith(
      '19:a@thread.v2',
      'msg-1',
      structuredWithWhitespaceTerminators,
      [],
    );
    expect(result).toEqual({ action: 'edit', id: 'msg-1', chat: 'chat A' });
  });
});

describe('teams-post / teams-edit — the --html routing decision (in-process, no subprocess, no network)', () => {
  // A subprocess test cannot prove this: allowlist.assertPostable always runs first and throws
  // (or not) identically whether or not --html was passed, so mutating away the ternary/if-else
  // that picks sendHtmlMessage vs sendMessage would leave every exit-code test above green.
  // doPost/doEdit are called here directly with a fake TeamsChatsPort, so the assertion is on
  // which METHOD actually got called — a mutation to that branch fails these immediately.
  function fakePort(overrides: Partial<TeamsChatsPort>): TeamsChatsPort {
    const reject = () => Promise.reject(new Error('not part of this test'));
    return {
      listChats: reject,
      readMessages: async () => ({ messages: [] }) as unknown as ReadResult,
      resolveMentions: reject,
      sendMessage: reject,
      sendHtmlMessage: reject,
      sendImage: reject,
      sendFile: reject,
      replyToMessage: reject,
      replyToHtmlMessage: reject,
      editMessage: reject,
      editHtmlMessage: reject,
      deleteMessage: reject,
      undoDeleteMessage: reject,
      setReaction: reject,
      getAttachment: reject,
      listAttachments: reject,
      getAttachments: reject,
      pinMessage: reject,
      unpinMessage: reject,
      listPinnedMessages: reject,
      ...overrides,
    } as TeamsChatsPort;
  }

  const allowlist = new ChatAllowlist([{ id: '19:a@thread.v2', label: 'chat A', canPost: true }]);
  const stubMessage = (id: string) =>
    ({ id, chatId: '19:a@thread.v2', createdDateTime: '2026-08-25T10:00:00Z', from: 'Assistant', text: '', isDeleted: false, attachments: [] }) as ChatMessage;

  it('doPost without --html calls sendMessage — never sendHtmlMessage', async () => {
    const sendMessage = vi.fn(async () => stubMessage('m1'));
    const sendHtmlMessage = vi.fn();
    const chats = new ReliableTeamsChats(fakePort({ sendMessage, sendHtmlMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    const result = await doPost({ chats, allowlist }, '19:a@thread.v2', 'hello', false);

    expect(sendMessage).toHaveBeenCalledWith('19:a@thread.v2', 'hello', []);
    expect(sendHtmlMessage).not.toHaveBeenCalled();
    expect(result).toEqual({ action: 'post', id: 'm1', chat: 'chat A' });
  });

  it('doPost with --html calls sendHtmlMessage — never sendMessage', async () => {
    const sendMessage = vi.fn();
    const sendHtmlMessage = vi.fn(async () => stubMessage('m2'));
    const chats = new ReliableTeamsChats(fakePort({ sendMessage, sendHtmlMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    const result = await doPost({ chats, allowlist }, '19:a@thread.v2', '<b>hi</b>', true);

    expect(sendHtmlMessage).toHaveBeenCalledWith('19:a@thread.v2', '<b>hi</b>', []);
    expect(sendMessage).not.toHaveBeenCalled();
    expect(result).toEqual({ action: 'post', id: 'm2', chat: 'chat A' });
  });

  it('doEdit without --html calls editMessage — never editHtmlMessage', async () => {
    const editMessage = vi.fn(async () => undefined);
    const editHtmlMessage = vi.fn();
    const chats = new ReliableTeamsChats(fakePort({ editMessage, editHtmlMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    const result = await doEdit({ chats, allowlist }, '19:a@thread.v2', 'msg-1', 'corrected', false);

    expect(editMessage).toHaveBeenCalledWith('19:a@thread.v2', 'msg-1', 'corrected', []);
    expect(editHtmlMessage).not.toHaveBeenCalled();
    expect(result).toEqual({ action: 'edit', id: 'msg-1', chat: 'chat A' });
  });

  it('doEdit with --html calls editHtmlMessage — never editMessage', async () => {
    const editMessage = vi.fn();
    const editHtmlMessage = vi.fn(async () => undefined);
    const chats = new ReliableTeamsChats(fakePort({ editMessage, editHtmlMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    const result = await doEdit(
      { chats, allowlist },
      '19:a@thread.v2',
      'msg-1',
      '<b>corrected</b>',
      true,
    );

    expect(editHtmlMessage).toHaveBeenCalledWith('19:a@thread.v2', 'msg-1', '<b>corrected</b>', []);
    expect(editMessage).not.toHaveBeenCalled();
    expect(result).toEqual({ action: 'edit', id: 'msg-1', chat: 'chat A' });
  });

  it('doPost with --mention resolves the name and forwards the resolved mention to sendMessage', async () => {
    const resolveMentions = vi.fn(async () => [
      { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' },
    ]);
    const sendMessage = vi.fn(async () => stubMessage('m3'));
    const chats = new ReliableTeamsChats(fakePort({ resolveMentions, sendMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    await doPost({ chats, allowlist }, '19:a@thread.v2', 'Mika please review', false, ['Mika']);

    expect(resolveMentions).toHaveBeenCalledWith('19:a@thread.v2', ['Mika']);
    expect(sendMessage).toHaveBeenCalledWith('19:a@thread.v2', 'Mika please review', [
      { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' },
    ]);
  });

  it('doPost with no --mention never calls resolveMentions', async () => {
    const resolveMentions = vi.fn();
    const sendMessage = vi.fn(async () => stubMessage('m4'));
    const chats = new ReliableTeamsChats(fakePort({ resolveMentions, sendMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    await doPost({ chats, allowlist }, '19:a@thread.v2', 'hello', false);

    expect(resolveMentions).not.toHaveBeenCalled();
  });

  it('doEdit with --mention resolves the name and forwards it to editMessage', async () => {
    const resolveMentions = vi.fn(async () => [
      { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' },
    ]);
    const editMessage = vi.fn(async () => undefined);
    const chats = new ReliableTeamsChats(fakePort({ resolveMentions, editMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    await doEdit({ chats, allowlist }, '19:a@thread.v2', 'msg-1', 'Mika see above', false, ['Mika']);

    expect(editMessage).toHaveBeenCalledWith('19:a@thread.v2', 'msg-1', 'Mika see above', [
      { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' },
    ]);
  });

  // Review round 2, MINOR 5 (2026-08-26): teams-reply had no --mention even though the MCP tool
  // and the port both support it on replies — a half-shipped surface.
  it('doReply with --mention resolves the name and forwards it to replyToMessage', async () => {
    const resolveMentions = vi.fn(async () => [
      { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' },
    ]);
    const replyToMessage = vi.fn(async () => stubMessage('r1'));
    const chats = new ReliableTeamsChats(fakePort({ resolveMentions, replyToMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    const result = await doReply(
      { chats, allowlist },
      '19:a@thread.v2',
      'orig-1',
      'Mika can you confirm?',
      false,
      ['Mika'],
    );

    expect(resolveMentions).toHaveBeenCalledWith('19:a@thread.v2', ['Mika']);
    expect(replyToMessage).toHaveBeenCalledWith('19:a@thread.v2', 'orig-1', 'Mika can you confirm?', [
      { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' },
    ]);
    expect(result).toEqual({ action: 'reply', id: 'r1', inReplyTo: 'orig-1', chat: 'chat A' });
  });

  it('doReply with no --mention never calls resolveMentions', async () => {
    const resolveMentions = vi.fn();
    const replyToMessage = vi.fn(async () => stubMessage('r2'));
    const chats = new ReliableTeamsChats(fakePort({ resolveMentions, replyToMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    await doReply({ chats, allowlist }, '19:a@thread.v2', 'orig-1', 'plain reply', false, []);

    expect(resolveMentions).not.toHaveBeenCalled();
  });

  // reply --html parity (0.6.x): a subprocess test cannot prove this — assertPostable always
  // throws (or not) identically whether or not --html was passed, same reasoning as doPost's own
  // --html routing tests above. doReply is called here directly with a fake TeamsChatsPort, so
  // the assertion is on which METHOD actually got called.
  it('doReply without --html calls replyToMessage — never replyToHtmlMessage', async () => {
    const replyToMessage = vi.fn(async () => stubMessage('r3'));
    const replyToHtmlMessage = vi.fn();
    const chats = new ReliableTeamsChats(fakePort({ replyToMessage, replyToHtmlMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    const result = await doReply({ chats, allowlist }, '19:a@thread.v2', 'orig-1', 'plain reply', false);

    expect(replyToMessage).toHaveBeenCalledWith('19:a@thread.v2', 'orig-1', 'plain reply', []);
    expect(replyToHtmlMessage).not.toHaveBeenCalled();
    expect(result).toEqual({ action: 'reply', id: 'r3', inReplyTo: 'orig-1', chat: 'chat A' });
  });

  it('doReply with --html calls replyToHtmlMessage — never replyToMessage', async () => {
    const replyToMessage = vi.fn();
    const replyToHtmlMessage = vi.fn(async () => stubMessage('r4'));
    const chats = new ReliableTeamsChats(fakePort({ replyToMessage, replyToHtmlMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    const result = await doReply({ chats, allowlist }, '19:a@thread.v2', 'orig-1', '<b>done</b>', true);

    expect(replyToHtmlMessage).toHaveBeenCalledWith('19:a@thread.v2', 'orig-1', '<b>done</b>', []);
    expect(replyToMessage).not.toHaveBeenCalled();
    expect(result).toEqual({ action: 'reply', id: 'r4', inReplyTo: 'orig-1', chat: 'chat A' });
  });

  it('doReply --html with --mention resolves the name and forwards it to replyToHtmlMessage', async () => {
    const resolveMentions = vi.fn(async () => [
      { name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' },
    ]);
    const replyToHtmlMessage = vi.fn(async () => stubMessage('r5'));
    const chats = new ReliableTeamsChats(fakePort({ resolveMentions, replyToHtmlMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    const result = await doReply(
      { chats, allowlist },
      '19:a@thread.v2',
      'orig-1',
      '<p>@{Mika} can you confirm?</p>',
      true,
      ['Mika'],
    );

    expect(resolveMentions).toHaveBeenCalledWith('19:a@thread.v2', ['Mika']);
    expect(replyToHtmlMessage).toHaveBeenCalledWith(
      '19:a@thread.v2',
      'orig-1',
      '<p>@{Mika} can you confirm?</p>',
      [{ name: 'Mika', id: 'aad-mika', displayName: 'Berggren, Mikael' }],
    );
    expect(result).toEqual({ action: 'reply', id: 'r5', inReplyTo: 'orig-1', chat: 'chat A' });
  });
});

describe('doPin — confirms the target message actually landed before claiming success (review round 2, MINOR 4)', () => {
  function fakePinPort(overrides: Partial<TeamsChatsPort>): TeamsChatsPort {
    const reject = () => Promise.reject(new Error('not part of this test'));
    return {
      listChats: reject,
      readMessages: () => Promise.resolve({ messages: [] } as unknown as ReadResult),
      resolveMentions: reject,
      sendMessage: reject,
      sendHtmlMessage: reject,
      sendImage: reject,
      sendFile: reject,
      replyToMessage: reject,
      replyToHtmlMessage: reject,
      editMessage: reject,
      editHtmlMessage: reject,
      deleteMessage: reject,
      undoDeleteMessage: reject,
      setReaction: reject,
      getAttachment: reject,
      listAttachments: reject,
      getAttachments: reject,
      pinMessage: reject,
      unpinMessage: reject,
      listPinnedMessages: reject,
      ...overrides,
    } as TeamsChatsPort;
  }

  const allowlist = new ChatAllowlist([{ id: '19:a@thread.v2', label: 'chat A', canPost: true }]);

  it('resolves normally when the re-list shows the target message pinned', async () => {
    const pinMessage = vi.fn(async () => [{ id: 'pin-m1', messageId: 'm1', preview: 'x' }]);
    const chats = new ReliableTeamsChats(fakePinPort({ pinMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    const result = await doPin({ chats, allowlist }, '19:a@thread.v2', 'm1');

    expect(result).toEqual({
      action: 'pin',
      messageId: 'm1',
      chat: 'chat A',
      pinnedMessages: [{ id: 'pin-m1', messageId: 'm1', preview: 'x' }],
    });
  });

  it('throws — never succeed()s — when the re-list does not show the target message pinned', async () => {
    const pinMessage = vi.fn(async () => [] as never[]);
    const chats = new ReliableTeamsChats(fakePinPort({ pinMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    await expect(doPin({ chats, allowlist }, '19:a@thread.v2', 'm1')).rejects.toThrow(
      /not (confirmed|show|pinned)/i,
    );
  });
});

describe('doSendFile — one sendFile call per positional path, --caption applied to the FIRST file only (new capability, 0.4.2)', () => {
  function fakeFilePort(overrides: Partial<TeamsChatsPort>): TeamsChatsPort {
    const reject = () => Promise.reject(new Error('not part of this test'));
    return {
      listChats: reject,
      readMessages: () => Promise.resolve({ messages: [] } as unknown as ReadResult),
      resolveMentions: reject,
      sendMessage: reject,
      sendHtmlMessage: reject,
      sendImage: reject,
      sendFile: reject,
      replyToMessage: reject,
      replyToHtmlMessage: reject,
      editMessage: reject,
      editHtmlMessage: reject,
      deleteMessage: reject,
      undoDeleteMessage: reject,
      setReaction: reject,
      getAttachment: reject,
      listAttachments: reject,
      getAttachments: reject,
      pinMessage: reject,
      unpinMessage: reject,
      listPinnedMessages: reject,
      ...overrides,
    } as TeamsChatsPort;
  }

  const allowlist = new ChatAllowlist([{ id: '19:a@thread.v2', label: 'chat A', canPost: true }]);
  const stubMessage = (id: string) =>
    ({
      id,
      chatId: '19:a@thread.v2',
      createdDateTime: '2026-09-02T10:00:00Z',
      from: 'Assistant',
      text: '',
      isDeleted: false,
      attachments: [],
    }) as ChatMessage;

  it('sends a single file with no caption, streamed via onSent', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'teams-send-file-'));
    const filePath = join(dir, 'a.txt');
    writeFileSync(filePath, 'hello');
    const sendFile = vi.fn(async () => stubMessage('f1'));
    const chats = new ReliableTeamsChats(fakeFilePort({ sendFile }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });
    const sent: unknown[] = [];

    await doSendFile({ chats, allowlist }, '19:a@thread.v2', [filePath], undefined, (payload) =>
      sent.push(payload),
    );

    expect(sendFile).toHaveBeenCalledTimes(1);
    expect(sendFile).toHaveBeenCalledWith(
      '19:a@thread.v2',
      { bytes: expect.any(Uint8Array), name: 'a.txt' },
      undefined,
      undefined,
    );
    expect(sent).toEqual([
      { action: 'send-file', id: 'f1', chat: 'chat A', name: 'a.txt', bytes: 5, granted: true },
    ]);
  });

  it('applies --caption to the FIRST file only, when several paths are given', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'teams-send-file-'));
    const path1 = join(dir, 'first.txt');
    const path2 = join(dir, 'second.txt');
    writeFileSync(path1, 'one');
    writeFileSync(path2, 'two-two');
    const sendFile = vi.fn(async (_chatId: string, file: { name: string }, text?: string) =>
      stubMessage(text ? 'f-with-caption' : `f-${file.name}`),
    );
    const chats = new ReliableTeamsChats(fakeFilePort({ sendFile }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });
    const sent: unknown[] = [];

    await doSendFile({ chats, allowlist }, '19:a@thread.v2', [path1, path2], 'see attached', (payload) =>
      sent.push(payload),
    );

    expect(sendFile).toHaveBeenCalledTimes(2);
    expect(sendFile).toHaveBeenNthCalledWith(
      1,
      '19:a@thread.v2',
      { bytes: expect.any(Uint8Array), name: 'first.txt' },
      'see attached',
      undefined,
    );
    expect(sendFile).toHaveBeenNthCalledWith(
      2,
      '19:a@thread.v2',
      { bytes: expect.any(Uint8Array), name: 'second.txt' },
      undefined, // the second (and every later) file gets no caption
      undefined,
    );
    expect(sent).toEqual([
      { action: 'send-file', id: 'f-with-caption', chat: 'chat A', name: 'first.txt', bytes: 3, granted: true },
      { action: 'send-file', id: 'f-second.txt', chat: 'chat A', name: 'second.txt', bytes: 7, granted: true },
    ]);
  });

  // MAJOR fix (2026-09-02 review): a multi-file send that fails partway through used to discard
  // the JSON success lines for files already posted -- exit 1, empty stdout, file 1 irreversibly
  // in the chat, and a caller with no way to tell it had already landed re-runs the whole batch
  // and duplicates it (the exact 2026-08-24 incident class the CLI output contract exists to
  // prevent). onSent must fire for each landed file BEFORE a later failure propagates.
  it('a mid-list failure still delivers the earlier successes via onSent before rejecting', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'teams-send-file-'));
    const path1 = join(dir, 'first.txt');
    const path2 = join(dir, 'second.txt');
    writeFileSync(path1, 'one');
    writeFileSync(path2, 'two');
    const sendFile = vi.fn(async (_chatId: string, file: { name: string }) => {
      if (file.name === 'second.txt') {
        throw new Error('Graph unavailable');
      }
      return stubMessage(`f-${file.name}`);
    });
    const chats = new ReliableTeamsChats(fakeFilePort({ sendFile }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });
    const sent: unknown[] = [];

    await expect(
      doSendFile({ chats, allowlist }, '19:a@thread.v2', [path1, path2], undefined, (payload) =>
        sent.push(payload),
      ),
    ).rejects.toThrow(/Graph unavailable/);

    // The first file's success is visible even though the batch overall failed on the second.
    expect(sent).toEqual([
      { action: 'send-file', id: 'f-first.txt', chat: 'chat A', name: 'first.txt', bytes: 3, granted: true },
    ]);
    expect(sendFile).toHaveBeenCalledTimes(2); // both were attempted; only the second failed
  });

  // Review round 1 MAJOR 1 (fresh-context re-review of PR #24): doSendFile's 6th (sendOptions)
  // parameter reaching chats.sendFile as its 4th argument was untested — deleting the `options,`
  // forwarding argument left the full 571-test suite green. These two pin the exact shape.
  it('MAJOR 1: forwards sendOptions.grantTo to sendFile as its 4th argument, untouched', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'teams-send-file-'));
    const filePath = join(dir, 'a.txt');
    writeFileSync(filePath, 'hello');
    const sendFile = vi.fn(async () => stubMessage('f1'));
    const chats = new ReliableTeamsChats(fakeFilePort({ sendFile }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    await doSendFile({ chats, allowlist }, '19:a@thread.v2', [filePath], undefined, () => {}, {
      grantTo: ['aad-x', 'aad-y'],
    });

    expect(sendFile).toHaveBeenCalledWith(
      '19:a@thread.v2',
      { bytes: expect.any(Uint8Array), name: 'a.txt' },
      undefined,
      { grantTo: ['aad-x', 'aad-y'] },
    );
  });

  it('MAJOR 1: forwards sendOptions.noGrant to sendFile as its 4th argument, untouched', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'teams-send-file-'));
    const filePath = join(dir, 'a.txt');
    writeFileSync(filePath, 'hello');
    const sendFile = vi.fn(async () => stubMessage('f1'));
    const chats = new ReliableTeamsChats(fakeFilePort({ sendFile }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    await doSendFile({ chats, allowlist }, '19:a@thread.v2', [filePath], undefined, () => {}, {
      noGrant: true,
    });

    expect(sendFile).toHaveBeenCalledWith(
      '19:a@thread.v2',
      { bytes: expect.any(Uint8Array), name: 'a.txt' },
      undefined,
      { noGrant: true },
    );
  });

  // Review round 1 MAJOR 2: the --no-grant stdout disclosure (granted:false + the "only the
  // sender" note) had no test at the doSendFile level — `granted: !sendOptions.noGrant` was
  // replaceable by a hardcoded `granted: true` with the suite still green.
  it('MAJOR 2: --no-grant reports granted:false and the "only the sender" note on the streamed payload', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'teams-send-file-'));
    const filePath = join(dir, 'a.txt');
    writeFileSync(filePath, 'hello');
    const sendFile = vi.fn(async () => stubMessage('f1'));
    const chats = new ReliableTeamsChats(fakeFilePort({ sendFile }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });
    const sent: unknown[] = [];

    await doSendFile(
      { chats, allowlist },
      '19:a@thread.v2',
      [filePath],
      undefined,
      (payload) => sent.push(payload),
      { noGrant: true },
    );

    expect(sent).toEqual([
      {
        action: 'send-file',
        id: 'f1',
        chat: 'chat A',
        name: 'a.txt',
        bytes: 5,
        granted: false,
        note: expect.stringMatching(/only the sender can open/),
      },
    ]);
  });

  // MINOR fix (2026-09-02 review, mutation-verified gap): a bare `.rejects.toThrow()` with no
  // matcher and a genuinely nonexistent path passes on ANY rejection, including readFile's own
  // ENOENT -- it does not actually prove the allowlist check ran BEFORE the filesystem read. A
  // real file plus a specific error match closes that gap: if the allowlist check were ever
  // skipped, this would proceed to read the (real) file and reach sendFile instead of throwing.
  it('a chat without canPost is refused before any file is even read from disk', async () => {
    const readonly = new ChatAllowlist([{ id: '19:ro@thread.v2', label: 'read-only', canPost: false }]);
    const dir = mkdtempSync(join(tmpdir(), 'teams-send-file-'));
    const realFile = join(dir, 'exists.txt');
    writeFileSync(realFile, 'hello');
    const sendFile = vi.fn();
    const onSent = vi.fn();
    const chats = new ReliableTeamsChats(fakeFilePort({ sendFile }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    await expect(
      doSendFile({ chats, allowlist: readonly }, '19:ro@thread.v2', [realFile], undefined, onSent),
    ).rejects.toThrow(/not on the allowlist/);
    expect(sendFile).not.toHaveBeenCalled();
    expect(onSent).not.toHaveBeenCalled();
  });
});

describe('writeLine() — the per-file streaming primitive teams-send-file uses (replaces succeedMany, 2026-09-02 review MINOR: untested happy path / dead empty-branch)', () => {
  afterEach(() => vi.restoreAllMocks());

  it('resolves only after the write callback fires — a full pipe cannot truncate the line', async () => {
    let flushCallback: (() => void) | undefined;
    const write = vi
      .spyOn(process.stdout, 'write')
      .mockImplementation((_chunk: unknown, cb?: unknown) => {
        flushCallback = cb as () => void;
        return false; // signal a full pipe: the data is NOT yet delivered
      });

    let resolved = false;
    const promise = writeLine({ action: 'send-file', id: 'f1' }).then(() => {
      resolved = true;
    });

    expect(write).toHaveBeenCalledWith(
      `${JSON.stringify({ ok: true, action: 'send-file', id: 'f1' })}\n`,
      expect.any(Function),
    );
    expect(resolved).toBe(false); // the old bug this guards against: resolving before the flush
    flushCallback?.();
    await promise;
    expect(resolved).toBe(true);
  });
});

describe('parseSendFlags — --html, --text and repeatable --mention', () => {
  // `text: false` added to every pre-existing toEqual below (2026-09-21, C1): the return shape
  // grew a `text` field for the new --text override flag — the schema change breaks the exact
  // toEqual match, not any behaviour these particular cases exercise.
  it('finds no flags in an empty argv', () => {
    expect(parseSendFlags([])).toEqual({ html: false, text: false, mentions: [], rest: [] });
  });

  it('finds a bare --html', () => {
    expect(parseSendFlags(['--html'])).toEqual({ html: true, text: false, mentions: [], rest: [] });
  });

  // C1: --text is the plain-text guard's deliberate override, same bare-flag shape as --html.
  it('finds a bare --text', () => {
    expect(parseSendFlags(['--text'])).toEqual({ html: false, text: true, mentions: [], rest: [] });
  });

  it('collects one --mention', () => {
    expect(parseSendFlags(['--mention', 'Mika'])).toEqual({ html: false, text: false, mentions: ['Mika'], rest: [] });
  });

  it('collects several repeated --mention flags, in order', () => {
    expect(parseSendFlags(['--mention', 'Mika', '--mention', 'Johan'])).toEqual({
      html: false,
      text: false,
      mentions: ['Mika', 'Johan'],
      rest: [],
    });
  });

  it('mixes --html and --mention in any order', () => {
    expect(parseSendFlags(['--mention', 'Mika', '--html'])).toEqual({
      html: true,
      text: false,
      mentions: ['Mika'],
      rest: [],
    });
  });

  it('leaves unrecognised arguments in rest, untouched', () => {
    expect(parseSendFlags(['--weird', 'value'])).toEqual({
      html: false,
      text: false,
      mentions: [],
      rest: ['--weird', 'value'],
    });
  });
});

describe('parseSendFileFlags — positional paths plus an optional --caption', () => {
  it('collects one path, no caption', () => {
    expect(parseSendFileFlags(['a.txt'])).toEqual({ caption: undefined, paths: ['a.txt'] });
  });

  it('collects several paths, in order', () => {
    expect(parseSendFileFlags(['a.txt', 'b.txt', 'c.txt'])).toEqual({
      caption: undefined,
      paths: ['a.txt', 'b.txt', 'c.txt'],
    });
  });

  it('extracts --caption regardless of where it appears among the paths, leaving it out of paths', () => {
    expect(parseSendFileFlags(['a.txt', '--caption', 'see attached', 'b.txt'])).toEqual({
      caption: 'see attached',
      paths: ['a.txt', 'b.txt'],
    });
  });

  // 0.6.0, live 2026-09-08: --grant-to/--no-grant, the caller-facing escape hatch around a
  // throttled/unavailable roster (see GraphTeamsChats.sendFile's own doc comment).
  it('--grant-to <ids> splits a comma-separated list into grantTo, leaving it out of paths', () => {
    expect(parseSendFileFlags(['a.txt', '--grant-to', 'aad-1,aad-2'])).toEqual({
      caption: undefined,
      paths: ['a.txt'],
      grantTo: ['aad-1', 'aad-2'],
    });
  });

  it('--grant-to with a single id (no comma) still produces a one-element array', () => {
    expect(parseSendFileFlags(['a.txt', '--grant-to', 'aad-1'])).toEqual({
      caption: undefined,
      paths: ['a.txt'],
      grantTo: ['aad-1'],
    });
  });

  it('--no-grant is a bare flag, leaving it out of paths', () => {
    expect(parseSendFileFlags(['a.txt', '--no-grant'])).toEqual({
      caption: undefined,
      paths: ['a.txt'],
      noGrant: true,
    });
  });

  // Invalid-input cases (usage()/process.exit(2)) are exercised via subprocess only, same
  // convention as --caption's own error paths below — see "teams-send-file — exit codes
  // (subprocess)": usage() really calls process.exit, which a direct unit call would either kill
  // the test worker or (mocked) fall through past TypeScript's `never`-typed control-flow
  // assumption, unlike a real subprocess exit.
});

describe('teams-pin / teams-unpin — exit codes (subprocess)', () => {
  it('teams-pin missing arguments: exit 2, stdout empty', async () => {
    const result = await runCli('pin.ts', [], {});

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/usage/);
  });

  it('teams-pin a chat outside the allowlist: exit 3, stdout empty', async () => {
    const result = await runCli('pin.ts', ['19:never-heard-of@thread.v2', 'm1'], fixtureEnv());

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
  });

  it('teams-pin an allowlisted chat without canPost: exit 3, stdout empty', async () => {
    const result = await runCli('pin.ts', ['19:readonly@thread.v2', 'm1'], fixtureEnv());

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
  });

  it('teams-unpin missing arguments: exit 2, stdout empty', async () => {
    const result = await runCli('unpin.ts', [], {});

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/usage/);
  });

  it('teams-unpin a chat outside the allowlist: exit 3, stdout empty', async () => {
    const result = await runCli('unpin.ts', ['19:never-heard-of@thread.v2', 'm1'], fixtureEnv());

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
  });

  it('teams-unpin an allowlisted chat without canPost: exit 3, stdout empty', async () => {
    const result = await runCli('unpin.ts', ['19:readonly@thread.v2', 'm1'], fixtureEnv());

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
  });
});

describe('teams-send-file — exit codes (subprocess)', () => {
  it('missing arguments: exit 2, stdout empty', async () => {
    const result = await runCli('send-file.ts', [], {});

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/usage/);
  });

  it('a chatId with no path at all: exit 2, stdout empty', async () => {
    const result = await runCli('send-file.ts', ['19:readonly@thread.v2'], fixtureEnv());

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
  });

  it('a chat outside the allowlist: exit 3, stdout empty', async () => {
    const result = await runCli(
      'send-file.ts',
      ['19:never-heard-of@thread.v2', '/tmp/does-not-matter.txt'],
      fixtureEnv(),
    );

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
  });

  it('an allowlisted chat without canPost: exit 3, stdout empty', async () => {
    const result = await runCli(
      'send-file.ts',
      ['19:readonly@thread.v2', '/tmp/does-not-matter.txt'],
      fixtureEnv(),
    );

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
  });

  it('--caption with no value: exit 2', async () => {
    const result = await runCli(
      'send-file.ts',
      ['19:readonly@thread.v2', '/tmp/does-not-matter.txt', '--caption'],
      fixtureEnv(),
    );

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/--caption/);
  });

  // MINOR fix (2026-09-02 review): align with teams-reply's doctrine of refusing a stray leftover
  // flag instead of silently treating it as a literal path.
  it('an unrecognised leftover flag: exit 2, stdout empty', async () => {
    const result = await runCli(
      'send-file.ts',
      ['19:readonly@thread.v2', '/tmp/does-not-matter.txt', '--verbose'],
      fixtureEnv(),
    );

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/--verbose/);
  });

  // 0.6.0, live 2026-09-08: --grant-to/--no-grant's error paths.
  it('--grant-to and --no-grant together: exit 2, stdout empty, mutually exclusive', async () => {
    const result = await runCli(
      'send-file.ts',
      ['19:readonly@thread.v2', '/tmp/does-not-matter.txt', '--grant-to', 'aad-1', '--no-grant'],
      fixtureEnv(),
    );

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/mutually exclusive/);
  });

  it('--grant-to with no value: exit 2', async () => {
    const result = await runCli(
      'send-file.ts',
      ['19:readonly@thread.v2', '/tmp/does-not-matter.txt', '--grant-to'],
      fixtureEnv(),
    );

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/--grant-to needs a value/);
  });

  it('--grant-to with an empty/all-blank list: exit 2 — same "no one to grant" ambiguity as an empty array', async () => {
    const result = await runCli(
      'send-file.ts',
      ['19:readonly@thread.v2', '/tmp/does-not-matter.txt', '--grant-to', ' , ,'],
      fixtureEnv(),
    );

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/--grant-to/);
  });
});

describe('succeed() drains stdout before exiting', () => {
  afterEach(() => vi.restoreAllMocks());

  it('exit(0) happens only after the write callback fires — a full pipe cannot truncate the line', () => {
    let flushCallback: (() => void) | undefined;
    const write = vi
      .spyOn(process.stdout, 'write')
      .mockImplementation((_chunk: unknown, cb?: unknown) => {
        flushCallback = cb as () => void;
        return false; // signal a full pipe: the data is NOT yet delivered
      });
    const exit = vi.spyOn(process, 'exit').mockImplementation(() => undefined as never);

    succeed({ action: 'test' });

    expect(write).toHaveBeenCalledTimes(1);
    expect(exit).not.toHaveBeenCalled(); // the old bug: exiting here truncates at 64 KiB
    flushCallback?.();
    expect(exit).toHaveBeenCalledWith(0);
  });
});

describe('run() — Retry-After discipline on the send path (0.4.1)', () => {
  afterEach(() => vi.restoreAllMocks());

  function captured() {
    const lines: string[] = [];
    const write = vi.spyOn(process.stderr, 'write').mockImplementation((chunk: unknown) => {
      lines.push(String(chunk));
      return true;
    });
    const exit = vi.spyOn(process, 'exit').mockImplementation(() => undefined as never);
    return { write, exit, text: () => lines.join('') };
  }

  it('a Graph 429 with a named Retry-After: the CLI error output names it, operator-readable', async () => {
    const { exit, text } = captured();

    await run(async () => {
      throw new GraphError('Too many requests', 429, 'TooManyRequests', 17);
    });

    expect(text()).toMatch(/throttled, retry after 17s/);
    expect(exit).toHaveBeenCalledWith(1);
  });

  it('a Graph 429 with NO named Retry-After: no invented number is printed', async () => {
    const { text } = captured();

    await run(async () => {
      throw new GraphError('Too many requests', 429, 'TooManyRequests');
    });

    expect(text()).not.toMatch(/throttled, retry after/);
    expect(text()).toMatch(/Too many requests/);
  });

  // MINOR (review round 1): a REAL GraphClient-thrown LocallyThrottled error used to state the
  // wait itself ("12s remain") inside its own message, and the CLI then ALSO appended
  // "(throttled, retry after 12s)" — two different renderings of the same number in one line.
  // This drives the real GraphClient throttle path (not a hand-built GraphError) end to end
  // through run(), so it catches the duplication at its actual source, not just in
  // formatCliError's own unit tests.
  it('a LOCALLY throttled 429 (the client-side gate, not a live Graph response) prints the wait exactly once, not twice', async () => {
    const { text } = captured();
    const { GraphClient } = await import('../graph/graph-client.js');
    const stubToken = { kind: 'stub', getAccessToken: async () => 't' };
    const fetchFn = async () =>
      new Response(JSON.stringify({ error: { code: 'TooManyRequests', message: 'Too many requests' } }), {
        status: 429,
        headers: { 'content-type': 'application/json', 'retry-after': '12' },
      });
    const client = new GraphClient({ tokenProvider: stubToken as never, fetchFn: fetchFn as never });

    await run(async () => {
      await client.get('/chats/x/members', { readRetries: 0 }).catch(() => {}); // 1st: a real 429, closes the local gate
      await client.get('/chats/x/members', { readRetries: 0 }); // 2nd: refused LOCALLY — this is the one that must reach run()
    });

    const output = text();
    expect((output.match(/12s/g) ?? []).length).toBeLessThanOrEqual(1); // the number appears at most once
    expect(output).not.toMatch(/\d+s remain/); // the old, now-redundant phrasing is gone
    // retryAfterSuffix (graph-client.ts) is the shared renderer both formatCliError (here) and
    // guard() (server.ts, see server.test.ts's own "guard()" describe block) call — one owner.
    expect(output).toMatch(/throttled, retry after \d+s/);
  });

  it('a non-429 failure: no throttle phrasing is added at all', async () => {
    const { text } = captured();

    await run(async () => {
      throw new Error('network unreachable');
    });

    expect(text()).not.toMatch(/throttled, retry after/);
    expect(text()).toMatch(/network unreachable/);
  });

  it('a 429 raised from deep in the send/reply/edit flow (ReliableTeamsChats rethrow) still surfaces Retry-After', async () => {
    const { text } = captured();
    const chats = new ReliableTeamsChats(
      {
        listChats: () => Promise.reject(new Error('n/a')),
        readMessages: () => Promise.resolve({ messages: [] }),
        resolveMentions: () => Promise.reject(new Error('n/a')),
        sendMessage: () => Promise.reject(new GraphError('Too many requests', 429, 'TooManyRequests', 9)),
        sendHtmlMessage: () => Promise.reject(new Error('n/a')),
        sendImage: () => Promise.reject(new Error('n/a')),
        sendFile: () => Promise.reject(new Error('n/a')),
        replyToMessage: () => Promise.reject(new Error('n/a')),
        editMessage: () => Promise.reject(new Error('n/a')),
        editHtmlMessage: () => Promise.reject(new Error('n/a')),
        deleteMessage: () => Promise.reject(new Error('n/a')),
        undoDeleteMessage: () => Promise.reject(new Error('n/a')),
        setReaction: () => Promise.reject(new Error('n/a')),
        getAttachment: () => Promise.reject(new Error('n/a')),
        listAttachments: () => Promise.reject(new Error('n/a')),
        getAttachments: () => Promise.reject(new Error('n/a')),
        pinMessage: () => Promise.reject(new Error('n/a')),
        unpinMessage: () => Promise.reject(new Error('n/a')),
        listPinnedMessages: () => Promise.reject(new Error('n/a')),
      },
      { selfDisplayName: 'Assistant', sleepFn: async () => {} },
    );
    const allowlist = new ChatAllowlist([{ id: '19:a@thread.v2', label: 'chat A', canPost: true }]);

    await run(async () => {
      await doPost({ chats, allowlist }, '19:a@thread.v2', 'hello', false);
    });

    expect(text()).toMatch(/throttled, retry after 9s/);
  });
});

// FINDING 4 (message-withdrawal review): a MessageOwnershipError from teams-delete's own-message
// gate must exit 4, distinct from ChatNotAllowedError's 3 and the generic 1 everything-else —
// same in-process run()-with-mocked-exit style as the Retry-After tests above, since reaching
// this deep into the delete flow needs a real (mocked) Graph call the subprocess exit-code tests
// (2/3, argv/allowlist only) deliberately never make.
describe('run() — exit code mapping (FINDING 4: MessageOwnershipError → 4, ChatNotAllowedError stays 3)', () => {
  afterEach(() => vi.restoreAllMocks());

  function captured() {
    const lines: string[] = [];
    const write = vi.spyOn(process.stderr, 'write').mockImplementation((chunk: unknown) => {
      lines.push(String(chunk));
      return true;
    });
    const exit = vi.spyOn(process, 'exit').mockImplementation(() => undefined as never);
    return { write, exit, text: () => lines.join('') };
  }

  it('a MessageOwnershipError (the ownership gate refusing somebody else\'s message) exits 4', async () => {
    const { exit, text } = captured();

    await run(async () => {
      throw new MessageOwnershipError(
        'Refusing to delete message m1 in chat 19:a@thread.v2: it was written by Alice, not by this account.',
      );
    });

    expect(text()).toMatch(/Refusing to delete message m1/);
    expect(exit).toHaveBeenCalledWith(4);
  });

  it('a ChatNotAllowedError still exits 3, not 4 — the two refusals stay distinguishable', async () => {
    const { exit } = captured();

    await run(async () => {
      throw new ChatNotAllowedError('19:a@thread.v2', 'post');
    });

    expect(exit).toHaveBeenCalledWith(3);
  });

  it('every other error still exits 1', async () => {
    const { exit } = captured();

    await run(async () => {
      throw new Error('some other failure');
    });

    expect(exit).toHaveBeenCalledWith(1);
  });
});

describe('teams-attachments — flag parsing (valid shapes in-process; refusals are subprocess tests below)', () => {
  it('defaults to download-everything mode with no flags', () => {
    expect(parseAttachmentFlags([])).toEqual({ list: false });
  });

  it('--list flips to metadata-only mode', () => {
    expect(parseAttachmentFlags(['--list'])).toEqual({ list: true });
  });

  it('--name and --out take their values wherever they appear', () => {
    expect(parseAttachmentFlags(['--out', '/data/dl', '--name', 'plan'])).toEqual({
      list: false,
      name: 'plan',
      out: '/data/dl',
    });
  });
});

describe('teams-attachments — the do* routing (in-process, fake port, real tmpdir writes)', () => {
  function fakeAttachmentPort(overrides: Partial<TeamsChatsPort>): TeamsChatsPort {
    const reject = () => Promise.reject(new Error('not part of this test'));
    return {
      listChats: reject,
      readMessages: async () => ({ messages: [] }) as unknown as ReadResult,
      resolveMentions: reject,
      sendMessage: reject,
      sendHtmlMessage: reject,
      sendImage: reject,
      sendFile: reject,
      replyToMessage: reject,
      replyToHtmlMessage: reject,
      editMessage: reject,
      editHtmlMessage: reject,
      deleteMessage: reject,
      undoDeleteMessage: reject,
      setReaction: reject,
      getAttachment: reject,
      listAttachments: reject,
      getAttachments: reject,
      pinMessage: reject,
      unpinMessage: reject,
      listPinnedMessages: reject,
      ...overrides,
    } as TeamsChatsPort;
  }

  // canPost false on purpose: downloading needs a READABLE chat, nothing more.
  const allowlist = new ChatAllowlist([{ id: '19:r@thread.v2', label: 'watched chat', canPost: false }]);

  function reliable(overrides: Partial<TeamsChatsPort>): ReliableTeamsChats {
    return new ReliableTeamsChats(fakeAttachmentPort(overrides), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });
  }

  it('doDownloadAttachments writes each payload with a sanitized, message-prefixed name and reports the paths', async () => {
    const getAttachments = vi.fn(async () => [
      { bytes: new Uint8Array([1]), contentType: 'application/pdf', name: '../../plan.pdf' },
      { bytes: new Uint8Array([2, 2]), contentType: 'image/png', name: 'logo.png' },
    ]);
    const out = mkdtempSync(join(tmpdir(), 'teams-attachments-test-'));

    const result = await doDownloadAttachments(
      { chats: reliable({ getAttachments }), allowlist },
      '19:r@thread.v2',
      'msg-7',
      { out },
    );

    expect(getAttachments).toHaveBeenCalledWith('19:r@thread.v2', 'msg-7', undefined);
    expect(result.action).toBe('attachments');
    expect(result.chat).toBe('watched chat');
    expect(result.count).toBe(2);
    // The hostile ../../ is gone, the messageId prefix survives, and the files really exist.
    expect(result.files.map((file) => file.path)).toEqual([
      join(out, 'msg-7-plan.pdf'),
      join(out, 'msg-7-logo.png'),
    ]);
    expect([...readFileSync(result.files[0]!.path)]).toEqual([1]);
    expect([...readFileSync(result.files[1]!.path)]).toEqual([2, 2]);
  });

  // GH-13 (https://github.com/Agentic-Delivery/teams-assistant-mcp/issues/13): the teams-attachments
  // CLI's result JSON reported `bytes: null` for completed downloads in the field an operator uses
  // to verify a download actually completed. On HEAD the field is already populated from the bytes
  // ACTUALLY WRITTEN (writeDownload receives this exact array), but nothing asserted the value —
  // this is the permanent regression test closing that gap, per differently-sized payloads so a
  // stub returning a fixed length could never pass it by accident.
  it('GH-13: doDownloadAttachments reports each file\'s actual downloaded byte count, never null', async () => {
    const getAttachments = vi.fn(async () => [
      { bytes: new Uint8Array(5), contentType: 'application/pdf', name: 'five.pdf' },
      { bytes: new Uint8Array(12), contentType: 'image/png', name: 'twelve.png' },
    ]);
    const out = mkdtempSync(join(tmpdir(), 'teams-attachments-test-'));

    const result = await doDownloadAttachments(
      { chats: reliable({ getAttachments }), allowlist },
      '19:r@thread.v2',
      'msg-9',
      { out },
    );

    expect(result.files.map((file) => file.bytes)).toEqual([5, 12]);
    // The reported count matches what actually landed on disk, not just the in-memory claim.
    expect(statSync(result.files[0]!.path).size).toBe(5);
    expect(statSync(result.files[1]!.path).size).toBe(12);
  });

  it('doDownloadAttachments never overwrites: the same message downloaded twice suffixes the second copy', async () => {
    const getAttachments = vi.fn(async () => [
      { bytes: new Uint8Array([9]), contentType: 'application/pdf', name: 'plan.pdf' },
    ]);
    const out = mkdtempSync(join(tmpdir(), 'teams-attachments-test-'));
    const context = { chats: reliable({ getAttachments }), allowlist };

    const first = await doDownloadAttachments(context, '19:r@thread.v2', 'msg-7', { out });
    const second = await doDownloadAttachments(context, '19:r@thread.v2', 'msg-7', { out });

    expect(first.files[0]!.path).toBe(join(out, 'msg-7-plan.pdf'));
    expect(second.files[0]!.path).toBe(join(out, 'msg-7-plan-1.pdf'));
  });

  it('doDownloadAttachments forwards --name as the port-level filter', async () => {
    const getAttachments = vi.fn(async () => [
      { bytes: new Uint8Array([1]), contentType: 'image/png', name: 'logo.png' },
    ]);
    const out = mkdtempSync(join(tmpdir(), 'teams-attachments-test-'));

    await doDownloadAttachments(
      { chats: reliable({ getAttachments }), allowlist },
      '19:r@thread.v2',
      'msg-7',
      { name: 'logo', out },
    );

    expect(getAttachments).toHaveBeenCalledWith('19:r@thread.v2', 'msg-7', 'logo');
  });

  it('doListAttachments returns metadata with downloadable flags and downloads nothing', async () => {
    const listAttachments = vi.fn(async () => [
      { id: 'quote-1', contentType: 'messageReference', content: '{}' },
      { id: 'file-1', name: 'plan.pdf', contentType: 'reference', contentUrl: 'https://x/p' },
    ]);

    const result = await doListAttachments(
      { chats: reliable({ listAttachments }), allowlist },
      '19:r@thread.v2',
      'msg-7',
    );

    expect(result).toEqual({
      action: 'attachments-list',
      chat: 'watched chat',
      messageId: 'msg-7',
      count: 2,
      attachments: [
        { id: 'quote-1', contentType: 'messageReference', downloadable: false },
        { id: 'file-1', name: 'plan.pdf', contentType: 'reference', downloadable: true },
      ],
    });
  });

  it('both refuse a chat outside the allowlist before any port call', async () => {
    const getAttachments = vi.fn();
    const listAttachments = vi.fn();
    const context = { chats: reliable({ getAttachments, listAttachments }), allowlist };

    await expect(doDownloadAttachments(context, '19:other@thread.v2', 'msg-7')).rejects.toThrow();
    await expect(doListAttachments(context, '19:other@thread.v2', 'msg-7')).rejects.toThrow();
    expect(getAttachments).not.toHaveBeenCalled();
    expect(listAttachments).not.toHaveBeenCalled();
  });
});

describe('teams-read — attachment metadata in the output (0.5.0: a file used to be invisible here)', () => {
  const allowlist = new ChatAllowlist([{ id: '19:r@thread.v2', label: 'watched chat', canPost: false }]);

  it('doRead includes id/name/contentType when a message carries attachments, and omits the field when not', async () => {
    const messages: ChatMessage[] = [
      { id: 'm1', chatId: '19:r@thread.v2', createdDateTime: '2026-09-02T08:00:00Z', from: 'Maja', text: 'plain', isDeleted: false, attachments: [], format: 'text' },
      { id: 'm2', chatId: '19:r@thread.v2', createdDateTime: '2026-09-02T08:01:00Z', from: 'Maja', text: 'file attached', isDeleted: false, format: 'html',
        attachments: [{ id: 'att-1', name: 'plan.xlsx', contentType: 'reference', contentUrl: 'https://x/p' }] },
    ];
    const readMessages = vi.fn(async () => ({ messages }) as ReadResult);
    const chats = new ReliableTeamsChats(
      {
        readMessages,
      } as unknown as TeamsChatsPort,
      { selfDisplayName: 'Assistant', sleepFn: async () => {} },
    );

    const result = await doRead({ chats, allowlist }, '19:r@thread.v2', { limit: 5 });

    expect(readMessages).toHaveBeenCalledWith('19:r@thread.v2', undefined, 5);
    expect(result.messages[0]).not.toHaveProperty('attachments');
    // Metadata only — never contentUrl, never bytes.
    expect(result.messages[1]?.attachments).toEqual([
      { id: 'att-1', name: 'plan.xlsx', contentType: 'reference' },
    ]);
  });

  // C3 (audit fix, 2026-09-21): messages.ts:184 used to flatten HTML and discard contentType,
  // so a caller reading back its own sends could never tell whether the plain-text guard (C1)
  // was actually being honoured — this is the field the harvest ritual now measures compliance
  // from. Additive only: `text` keeps carrying the exact same flattened string as before either
  // way (asserted below), `format` is the new field.
  it('C3a: an html-contentType message reads back format: "html", text still flattened the same way', async () => {
    const messages: ChatMessage[] = [
      {
        id: 'm3',
        chatId: '19:r@thread.v2',
        createdDateTime: '2026-09-21T08:00:00Z',
        from: 'Assistant',
        text: 'Styled update\nSecond line',
        isDeleted: false,
        attachments: [],
        format: 'html',
      },
    ];
    const readMessages = vi.fn(async () => ({ messages }) as ReadResult);
    const chats = new ReliableTeamsChats({ readMessages } as unknown as TeamsChatsPort, {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    const result = await doRead({ chats, allowlist }, '19:r@thread.v2', { limit: 5 });

    expect(result.messages[0]).toMatchObject({ format: 'html', text: 'Styled update\nSecond line' });
  });

  // C3b: the non-triggering side — a plain-contentType message reads back format: "text".
  it('C3b: a text-contentType message reads back format: "text"', async () => {
    const messages: ChatMessage[] = [
      {
        id: 'm4',
        chatId: '19:r@thread.v2',
        createdDateTime: '2026-09-21T08:00:00Z',
        from: 'Assistant',
        text: 'plain update',
        isDeleted: false,
        attachments: [],
        format: 'text',
      },
    ];
    const readMessages = vi.fn(async () => ({ messages }) as ReadResult);
    const chats = new ReliableTeamsChats({ readMessages } as unknown as TeamsChatsPort, {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    const result = await doRead({ chats, allowlist }, '19:r@thread.v2', { limit: 5 });

    expect(result.messages[0]).toMatchObject({ format: 'text', text: 'plain update' });
  });
});

describe('teams-attachments — usage refusals (subprocess: the argv contract, exit 2/3)', () => {
  it('missing ids: exit 2 with usage', async () => {
    const result = await runCli('attachments.ts', ['19:readonly@thread.v2'], fixtureEnv());

    expect(result.code).toBe(2);
    expect(result.stderr).toContain('usage: teams-attachments');
    expect(result.stdout).toBe('');
  });

  it('an unrecognised flag: exit 2, refused loudly rather than silently ignored', async () => {
    const result = await runCli('attachments.ts', ['19:readonly@thread.v2', 'msg-1', '--nmae', 'x'], fixtureEnv());

    expect(result.code).toBe(2);
    expect(result.stderr).toContain('--nmae');
  });

  it('--list mixed with --name/--out: exit 2 — the combination has no meaning', async () => {
    const result = await runCli('attachments.ts', ['19:readonly@thread.v2', 'msg-1', '--list', '--name', 'x'], fixtureEnv());

    expect(result.code).toBe(2);
  });

  it('--name without a value: exit 2 naming the flag', async () => {
    const result = await runCli('attachments.ts', ['19:readonly@thread.v2', 'msg-1', '--name'], fixtureEnv());

    expect(result.code).toBe(2);
    expect(result.stderr).toContain('--name');
  });

  it('a chat outside the allowlist: exit 3 before any network call', async () => {
    const result = await runCli('attachments.ts', ['19:never-heard-of@thread.v2', 'msg-1'], fixtureEnv());

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
  });
});

describe('teams-attachments — the quota yield (0.5.0: a running daemon starved ad-hoc reads, measured 2026-09-02)', () => {
  const allowlist = new ChatAllowlist([{ id: '19:r@thread.v2', label: 'watched chat', canPost: false }]);

  function reliableWith(overrides: Partial<TeamsChatsPort>): ReliableTeamsChats {
    return new ReliableTeamsChats(overrides as TeamsChatsPort, {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });
  }

  it('doDownloadAttachments holds the yield file across the Graph work and releases it after', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'teams-attachments-yield-'));
    const yieldPath = join(dir, 'inbox-yield.json');
    let stoodDuringRead = false;
    const getAttachments = vi.fn(async () => {
      stoodDuringRead = existsSync(yieldPath);
      return [{ bytes: new Uint8Array([1]), contentType: 'application/pdf', name: 'plan.pdf' }];
    });

    await doDownloadAttachments(
      { chats: reliableWith({ getAttachments }), allowlist },
      '19:r@thread.v2',
      'msg-7',
      { out: dir, yieldPath },
    );

    expect(stoodDuringRead).toBe(true); // the poller sees this and sits the cycle out
    expect(existsSync(yieldPath)).toBe(false); // released the moment the Graph work is done
  });

  it('a failed download releases the yield too — a dead CLI must not silence the inbox until the deadline', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'teams-attachments-yield-'));
    const yieldPath = join(dir, 'inbox-yield.json');
    const getAttachments = vi.fn(async () => {
      throw new GraphError('throttled', 429, 'TooManyRequests', 62);
    });

    await expect(
      doDownloadAttachments(
        { chats: reliableWith({ getAttachments }), allowlist },
        '19:r@thread.v2',
        'msg-7',
        { out: dir, yieldPath },
      ),
    ).rejects.toThrow('throttled');

    expect(existsSync(yieldPath)).toBe(false);
  });

  it('doListAttachments yields the same way', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'teams-attachments-yield-'));
    const yieldPath = join(dir, 'inbox-yield.json');
    let stoodDuringRead = false;
    const listAttachments = vi.fn(async () => {
      stoodDuringRead = existsSync(yieldPath);
      return [];
    });

    await doListAttachments(
      { chats: reliableWith({ listAttachments }), allowlist },
      '19:r@thread.v2',
      'msg-7',
      { yieldPath },
    );

    expect(stoodDuringRead).toBe(true);
    expect(existsSync(yieldPath)).toBe(false);
  });
});

// GH-14 (https://github.com/Agentic-Delivery/teams-assistant-mcp/issues/14): the reported
// incident was against `teams-post`. These tests exercise doPost/doEdit through the REAL
// production chain (ReliableTeamsChats wrapping GraphTeamsChats, same composition buildChats
// wires for every CLI — build-chats.ts), mocking only the Graph HTTP transport, per the
// pragmatic-tdd mock-discipline boundary — never the whole TeamsChatsPort, which is where the
// existing doPost/doEdit routing tests above mock and so never actually exercise this guard.
describe('teams-post / teams-edit --html --mention — the orphaned-mention refusal through the real send chain (GH-14)', () => {
  const stubToken = { kind: 'stub' as const, getAccessToken: async () => 'tok' };
  const allowlist = new ChatAllowlist([{ id: '19:a@thread.v2', label: 'chat A', canPost: true }]);

  function realChats(fetchFn: typeof fetch) {
    const graph = new GraphClient({ tokenProvider: stubToken, fetchFn });
    const dir = mkdtempSync(join(tmpdir(), 'teams-gh14-members-'));
    const cache = new MembersCache({ path: join(dir, 'members.json') });
    cache.set('19:a@thread.v2', [{ id: 'aad-maja', displayName: 'Nordqvist, Maja' }]);
    return new ReliableTeamsChats(new GraphTeamsChats(graph, { membersCache: cache }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });
  }

  it('GH-14d: doPost --html --mention "Nordqvist, Maja" (name written plainly, no @{} token) refuses BEFORE any send', async () => {
    const fetchFn = vi.fn(async () => {
      throw new Error('must never be called — the refusal must happen before any Graph request');
    });
    const chats = realChats(fetchFn as unknown as typeof fetch);

    await expect(
      doPost(
        { chats, allowlist },
        '19:a@thread.v2',
        '<p>Please review Nordqvist, Maja</p>',
        true,
        ['Nordqvist, Maja'],
      ),
    ).rejects.toThrow(/no @\{Name\}-style placeholder/);
    expect(fetchFn).not.toHaveBeenCalled();
  });

  it('GH-14e: doEdit --html --mention "Nordqvist, Maja" (name written plainly, no @{} token) refuses BEFORE any send — same guard on the edit path', async () => {
    const fetchFn = vi.fn(async () => {
      throw new Error('must never be called — the refusal must happen before any Graph request');
    });
    const chats = realChats(fetchFn as unknown as typeof fetch);

    await expect(
      doEdit(
        { chats, allowlist },
        '19:a@thread.v2',
        'msg-1',
        '<p>Please review Nordqvist, Maja</p>',
        true,
        ['Nordqvist, Maja'],
      ),
    ).rejects.toThrow(/no @\{Name\}-style placeholder/);
    expect(fetchFn).not.toHaveBeenCalled();
  });

  it('GH-14f: doPost --html --mention with the @{Name} token present sends normally (the other decision side)', async () => {
    const fetchFn = vi.fn(async (url: string, init?: RequestInit) => {
      expect(init?.method).toBe('POST');
      return new Response(
        JSON.stringify({ id: 'sent-1', chatId: '19:a@thread.v2', createdDateTime: '2026-09-04T10:00:00Z', body: { contentType: 'html', content: '' } }),
        { status: 200, headers: { 'content-type': 'application/json' } },
      );
    });
    const chats = realChats(fetchFn as unknown as typeof fetch);

    const result = await doPost(
      { chats, allowlist },
      '19:a@thread.v2',
      '<p>Please review @{Nordqvist, Maja}</p>',
      true,
      ['Nordqvist, Maja'],
    );

    expect(result).toEqual({ action: 'post', id: 'sent-1', chat: 'chat A' });
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });
});

describe('teams-delete — flag parsing (0.7.1)', () => {
  it('defaults to delete, no force', () => {
    expect(parseDeleteFlags([])).toEqual({ undo: false, force: false });
  });

  it('--undo and --force are bare flags, any order', () => {
    expect(parseDeleteFlags(['--undo'])).toEqual({ undo: true, force: false });
    expect(parseDeleteFlags(['--force', '--undo'])).toEqual({ undo: true, force: true });
  });
});

describe('teams-delete — the do* routing (in-process, fake port, no network)', () => {
  // Same reasoning as the --html routing block above: the allowlist gate makes a subprocess
  // blind to WHICH port method ran and whether --force reached it, so doDelete is called
  // directly with a fake port. The own-message check itself is the port's job — proven in
  // graph-client.test.ts — which is exactly why this layer only has to prove the forwarding.
  function fakePort(overrides: Partial<TeamsChatsPort>): TeamsChatsPort {
    const reject = () => Promise.reject(new Error('not part of this test'));
    return {
      listChats: reject,
      readMessages: async () => ({ messages: [] }) as unknown as ReadResult,
      resolveMentions: reject,
      sendMessage: reject,
      sendHtmlMessage: reject,
      sendImage: reject,
      sendFile: reject,
      replyToMessage: reject,
      editMessage: reject,
      editHtmlMessage: reject,
      deleteMessage: reject,
      undoDeleteMessage: reject,
      setReaction: reject,
      getAttachment: reject,
      listAttachments: reject,
      getAttachments: reject,
      pinMessage: reject,
      unpinMessage: reject,
      listPinnedMessages: reject,
      ...overrides,
    } as TeamsChatsPort;
  }
  const allowlist = new ChatAllowlist([
    { id: '19:a@thread.v2', label: 'chat A', canPost: true },
    { id: '19:r@thread.v2', label: 'read-only', canPost: false },
  ]);
  function reliable(overrides: Partial<TeamsChatsPort>): ReliableTeamsChats {
    return new ReliableTeamsChats(fakePort(overrides), { selfDisplayName: 'Assistant', sleepFn: async () => {} });
  }

  it('default: deleteMessage with force false — never undoDeleteMessage', async () => {
    const deleteMessage = vi.fn(async () => undefined);
    const undoDeleteMessage = vi.fn();

    const result = await doDelete({ chats: reliable({ deleteMessage, undoDeleteMessage }), allowlist }, '19:a@thread.v2', 'm1', { undo: false, force: false });

    expect(deleteMessage).toHaveBeenCalledWith('19:a@thread.v2', 'm1', { force: false });
    expect(undoDeleteMessage).not.toHaveBeenCalled();
    expect(result).toEqual({ action: 'delete', messageId: 'm1', chat: 'chat A' });
  });

  it('--force reaches the port as force: true', async () => {
    const deleteMessage = vi.fn(async () => undefined);

    await doDelete({ chats: reliable({ deleteMessage }), allowlist }, '19:a@thread.v2', 'm1', { undo: false, force: true });

    expect(deleteMessage).toHaveBeenCalledWith('19:a@thread.v2', 'm1', { force: true });
  });

  it('--undo: undoDeleteMessage — never deleteMessage — and the action says so', async () => {
    const deleteMessage = vi.fn();
    const undoDeleteMessage = vi.fn(async () => undefined);

    const result = await doDelete({ chats: reliable({ deleteMessage, undoDeleteMessage }), allowlist }, '19:a@thread.v2', 'm1', { undo: true, force: false });

    expect(undoDeleteMessage).toHaveBeenCalledWith('19:a@thread.v2', 'm1', { force: false });
    expect(deleteMessage).not.toHaveBeenCalled();
    expect(result).toEqual({ action: 'undo-delete', messageId: 'm1', chat: 'chat A' });
  });

  it('the port\'s ownership refusal propagates as itself — the CLI exits 1 with its text, never 0', async () => {
    const deleteMessage = vi.fn(async () => {
      throw new MessageOwnershipError('Refusing to delete message m1 in chat 19:a@thread.v2: it was written by Alice');
    });

    await expect(
      doDelete({ chats: reliable({ deleteMessage }), allowlist }, '19:a@thread.v2', 'm1', { undo: false, force: false }),
    ).rejects.toBeInstanceOf(MessageOwnershipError);
  });

  it('a read-only chat is refused before the port is touched, delete and undo alike', async () => {
    const deleteMessage = vi.fn();
    const undoDeleteMessage = vi.fn();
    const chats = reliable({ deleteMessage, undoDeleteMessage });

    await expect(doDelete({ chats, allowlist }, '19:r@thread.v2', 'm1', { undo: false, force: true })).rejects.toThrow(/19:r@thread.v2/);
    await expect(doDelete({ chats, allowlist }, '19:r@thread.v2', 'm1', { undo: true, force: true })).rejects.toThrow(/19:r@thread.v2/);
    expect(deleteMessage).not.toHaveBeenCalled();
    expect(undoDeleteMessage).not.toHaveBeenCalled();
  });
});

describe('teams-delete — exit codes (subprocess: the argv contract)', () => {
  it('missing arguments: exit 2 with usage, stdout empty', async () => {
    const result = await runCli('delete.ts', [], {});

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/usage: teams-delete <chatId> <messageId> \[--undo\] \[--force\]/);
  });

  it('missing messageId: exit 2, stdout empty', async () => {
    const result = await runCli('delete.ts', ['19:readonly@thread.v2'], {});

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
  });

  it('an unrecognised flag: exit 2, refused loudly rather than silently ignored', async () => {
    const result = await runCli('delete.ts', ['19:readonly@thread.v2', 'm1', '--hard'], fixtureEnv());

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/--hard/);
  });

  it('a chat outside the allowlist: exit 3 before any network call', async () => {
    const result = await runCli('delete.ts', ['19:never-heard-of@thread.v2', 'm1'], fixtureEnv());

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
  });

  it('an allowlisted chat without canPost: exit 3 — --force does not get past the allowlist', async () => {
    const result = await runCli('delete.ts', ['19:readonly@thread.v2', 'm1', '--force'], fixtureEnv());

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
  });
});
