import { execFile } from 'node:child_process';
import { mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { buildChats } from '../build-chats.js';
import { ChatAllowlist } from '../allowlist.js';
import { ReliableTeamsChats } from '../graph/reliable-sends.js';
import type { TeamsChatsPort } from '../graph/teams-chats.js';
import type { ChatMessage, ReadResult } from '../messages.js';
import { loadConfig } from '../config.js';
import { doEdit, doPost, succeed } from './common.js';

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

    expect(sendMessage).toHaveBeenCalledWith('19:a@thread.v2', 'hello');
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

    expect(sendHtmlMessage).toHaveBeenCalledWith('19:a@thread.v2', '<b>hi</b>');
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

    expect(editMessage).toHaveBeenCalledWith('19:a@thread.v2', 'msg-1', 'corrected');
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

    expect(editHtmlMessage).toHaveBeenCalledWith('19:a@thread.v2', 'msg-1', '<b>corrected</b>');
    expect(editMessage).not.toHaveBeenCalled();
    expect(result).toEqual({ action: 'edit', id: 'msg-1', chat: 'chat A' });
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
