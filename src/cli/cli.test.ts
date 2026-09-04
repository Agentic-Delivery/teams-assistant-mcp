import { execFile } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { buildChats } from '../build-chats.js';
import { ChatAllowlist } from '../allowlist.js';
import { GraphClient, GraphError } from '../graph/graph-client.js';
import { MembersCache } from '../graph/members-cache.js';
import { ReliableTeamsChats } from '../graph/reliable-sends.js';
import { GraphTeamsChats, type TeamsChatsPort } from '../graph/teams-chats.js';
import type { ChatMessage, ReadResult } from '../messages.js';
import { loadConfig } from '../config.js';
import {
  doDownloadAttachments,
  doEdit,
  doListAttachments,
  doPin,
  doPost,
  doRead,
  doReply,
  doSendFile,
  parseAttachmentFlags,
  parseSendFileFlags,
  parseSendFlags,
  run,
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

function runCli(script: string, args: string[], env: Record<string, string>): Promise<CliRun> {
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
    child.stdin?.end('the message text');
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
    // Same shape as the --html test above: proves --mention/"Shiv" are parsed as a flag+value
    // pair, not misread as positionals, without needing a live resolveMentions call. The
    // resolution/forwarding itself is proven in-process below (doReply).
    const result = await runCli(
      'reply.ts',
      ['19:readonly@thread.v2', 'msg-1', '--mention', 'Shiv'],
      fixtureEnv(),
    );

    expect(result.code).toBe(3);
    expect(result.stdout).toBe('');
  });

  // Review round 2 follow-up NIT (2026-08-26): parseSendFlags also understands --html, but
  // reply.ts only ever destructured { mentions } — a stray --html (or any other leftover
  // argument) used to be silently accepted and ignored, and the reply still went out as plain
  // text with no error at all. teams-reply does not support raw HTML; it must say so.
  it('teams-reply rejects --html instead of silently posting plain text: exit 2', async () => {
    const result = await runCli(
      'reply.ts',
      ['19:readonly@thread.v2', 'msg-1', '--html'],
      fixtureEnv(),
    );

    expect(result.code).toBe(2);
    expect(result.stdout).toBe('');
    expect(result.stderr).toMatch(/--html/);
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
      editMessage: reject,
      editHtmlMessage: reject,
      deleteMessage: reject,
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
      { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' },
    ]);
    const sendMessage = vi.fn(async () => stubMessage('m3'));
    const chats = new ReliableTeamsChats(fakePort({ resolveMentions, sendMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    await doPost({ chats, allowlist }, '19:a@thread.v2', 'Shiv please review', false, ['Shiv']);

    expect(resolveMentions).toHaveBeenCalledWith('19:a@thread.v2', ['Shiv']);
    expect(sendMessage).toHaveBeenCalledWith('19:a@thread.v2', 'Shiv please review', [
      { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' },
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
      { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' },
    ]);
    const editMessage = vi.fn(async () => undefined);
    const chats = new ReliableTeamsChats(fakePort({ resolveMentions, editMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    await doEdit({ chats, allowlist }, '19:a@thread.v2', 'msg-1', 'Shiv see above', false, ['Shiv']);

    expect(editMessage).toHaveBeenCalledWith('19:a@thread.v2', 'msg-1', 'Shiv see above', [
      { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' },
    ]);
  });

  // Review round 2, MINOR 5 (2026-08-26): teams-reply had no --mention even though the MCP tool
  // and the port both support it on replies — a half-shipped surface.
  it('doReply with --mention resolves the name and forwards it to replyToMessage', async () => {
    const resolveMentions = vi.fn(async () => [
      { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' },
    ]);
    const replyToMessage = vi.fn(async () => stubMessage('r1'));
    const chats = new ReliableTeamsChats(fakePort({ resolveMentions, replyToMessage }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });

    const result = await doReply({ chats, allowlist }, '19:a@thread.v2', 'orig-1', 'Shiv can you confirm?', ['Shiv']);

    expect(resolveMentions).toHaveBeenCalledWith('19:a@thread.v2', ['Shiv']);
    expect(replyToMessage).toHaveBeenCalledWith('19:a@thread.v2', 'orig-1', 'Shiv can you confirm?', [
      { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' },
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

    await doReply({ chats, allowlist }, '19:a@thread.v2', 'orig-1', 'plain reply', []);

    expect(resolveMentions).not.toHaveBeenCalled();
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
      editMessage: reject,
      editHtmlMessage: reject,
      deleteMessage: reject,
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
      editMessage: reject,
      editHtmlMessage: reject,
      deleteMessage: reject,
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
    );
    expect(sent).toEqual([{ action: 'send-file', id: 'f1', chat: 'chat A', name: 'a.txt', bytes: 5 }]);
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
    );
    expect(sendFile).toHaveBeenNthCalledWith(
      2,
      '19:a@thread.v2',
      { bytes: expect.any(Uint8Array), name: 'second.txt' },
      undefined, // the second (and every later) file gets no caption
    );
    expect(sent).toEqual([
      { action: 'send-file', id: 'f-with-caption', chat: 'chat A', name: 'first.txt', bytes: 3 },
      { action: 'send-file', id: 'f-second.txt', chat: 'chat A', name: 'second.txt', bytes: 7 },
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
      { action: 'send-file', id: 'f-first.txt', chat: 'chat A', name: 'first.txt', bytes: 3 },
    ]);
    expect(sendFile).toHaveBeenCalledTimes(2); // both were attempted; only the second failed
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

describe('parseSendFlags — --html and repeatable --mention', () => {
  it('finds no flags in an empty argv', () => {
    expect(parseSendFlags([])).toEqual({ html: false, mentions: [], rest: [] });
  });

  it('finds a bare --html', () => {
    expect(parseSendFlags(['--html'])).toEqual({ html: true, mentions: [], rest: [] });
  });

  it('collects one --mention', () => {
    expect(parseSendFlags(['--mention', 'Shiv'])).toEqual({ html: false, mentions: ['Shiv'], rest: [] });
  });

  it('collects several repeated --mention flags, in order', () => {
    expect(parseSendFlags(['--mention', 'Shiv', '--mention', 'Johan'])).toEqual({
      html: false,
      mentions: ['Shiv', 'Johan'],
      rest: [],
    });
  });

  it('mixes --html and --mention in any order', () => {
    expect(parseSendFlags(['--mention', 'Shiv', '--html'])).toEqual({
      html: true,
      mentions: ['Shiv'],
      rest: [],
    });
  });

  it('leaves unrecognised arguments in rest, untouched', () => {
    expect(parseSendFlags(['--weird', 'value'])).toEqual({ html: false, mentions: [], rest: ['--weird', 'value'] });
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
      editMessage: reject,
      editHtmlMessage: reject,
      deleteMessage: reject,
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
      { id: 'm1', chatId: '19:r@thread.v2', createdDateTime: '2026-09-02T08:00:00Z', from: 'Celine', text: 'plain', isDeleted: false, attachments: [] },
      { id: 'm2', chatId: '19:r@thread.v2', createdDateTime: '2026-09-02T08:01:00Z', from: 'Celine', text: 'file attached', isDeleted: false,
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
    cache.set('19:a@thread.v2', [{ id: 'aad-celine', displayName: 'Kleivdal, Celine' }]);
    return new ReliableTeamsChats(new GraphTeamsChats(graph, { membersCache: cache }), {
      selfDisplayName: 'Assistant',
      sleepFn: async () => {},
    });
  }

  it('GH-14d: doPost --html --mention "Kleivdal, Celine" (name written plainly, no @{} token) refuses BEFORE any send', async () => {
    const fetchFn = vi.fn(async () => {
      throw new Error('must never be called — the refusal must happen before any Graph request');
    });
    const chats = realChats(fetchFn as unknown as typeof fetch);

    await expect(
      doPost(
        { chats, allowlist },
        '19:a@thread.v2',
        '<p>Please review Kleivdal, Celine</p>',
        true,
        ['Kleivdal, Celine'],
      ),
    ).rejects.toThrow(/no @\{Name\}-style placeholder/);
    expect(fetchFn).not.toHaveBeenCalled();
  });

  it('GH-14e: doEdit --html --mention "Kleivdal, Celine" (name written plainly, no @{} token) refuses BEFORE any send — same guard on the edit path', async () => {
    const fetchFn = vi.fn(async () => {
      throw new Error('must never be called — the refusal must happen before any Graph request');
    });
    const chats = realChats(fetchFn as unknown as typeof fetch);

    await expect(
      doEdit(
        { chats, allowlist },
        '19:a@thread.v2',
        'msg-1',
        '<p>Please review Kleivdal, Celine</p>',
        true,
        ['Kleivdal, Celine'],
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
      '<p>Please review @{Kleivdal, Celine}</p>',
      true,
      ['Kleivdal, Celine'],
    );

    expect(result).toEqual({ action: 'post', id: 'sent-1', chat: 'chat A' });
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });
});
