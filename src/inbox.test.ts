import { mkdtemp, readdir, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { ChatAllowlist } from './allowlist.js';
import { DEFAULT_POLL_MS, InboxPoller, PARK_FORBIDDEN_CHAT_MS } from './inbox.js';
import { type ChatMessage, applyWatermark } from './messages.js';

const CHAT = '19:pilot@thread.v2';
const OTHER = '19:leadership@thread.v2';

const me = { id: 'me-id', displayName: 'Assistant (AI)' };

function message(overrides: Partial<ChatMessage>): ChatMessage {
  return {
    id: 'm1',
    chatId: CHAT,
    createdDateTime: '2026-08-21T10:00:00Z',
    from: 'Alice',
    fromId: 'alice-id',
    text: 'hello',
    isDeleted: false,
    attachments: [],
    ...overrides,
  };
}

/**
 * A port double backed by a mutable message store, applying the real watermark logic so the
 * poller sees exactly what GraphTeamsChats.readMessages would deliver.
 */
function chatStore(initial: Record<string, ChatMessage[]> = {}) {
  const byChat = new Map(Object.entries(initial));
  return {
    add(chatId: string, ...messages: ChatMessage[]): void {
      byChat.set(chatId, [...(byChat.get(chatId) ?? []), ...messages]);
    },
    readMessages: (chatId: string, since?: string) =>
      Promise.resolve(applyWatermark(byChat.get(chatId) ?? [], since)),
  };
}

describe('inbox poller', () => {
  let dir: string;
  let inboxPath: string;
  let statePath: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'inbox-test-'));
    inboxPath = join(dir, 'nested', 'inbox.jsonl');
    statePath = join(dir, 'nested', 'inbox-state.json');
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  function poller(chats: Pick<ReturnType<typeof chatStore>, 'readMessages'>, chatIds = [CHAT]) {
    return new InboxPoller({
      chats,
      allowlist: new ChatAllowlist(chatIds.map((id) => ({ id, label: id, canPost: true }))),
      self: () => Promise.resolve(me),
      inboxPath,
      statePath,
    });
  }

  async function inboxLines(): Promise<Array<Record<string, unknown>>> {
    const raw = await readFile(inboxPath, 'utf8').catch(() => '');
    return raw
      .split('\n')
      .filter((line) => line.trim() !== '')
      .map((line) => JSON.parse(line) as Record<string, unknown>);
  }

  it('appends one JSON line per new message, in the daemon-compatible shape', async () => {
    const store = chatStore({
      [CHAT]: [
        message({ id: 'a', text: 'first', createdDateTime: '2026-08-21T10:00:00Z' }),
        message({
          id: 'b',
          text: 'with a file',
          createdDateTime: '2026-08-21T10:01:00Z',
          attachments: [{ id: 'att-1', name: 'plan.xlsx' }],
        }),
      ],
    });

    await poller(store).pollOnce();

    expect(await inboxLines()).toEqual([
      { chat: CHAT, id: 'a', from: 'Alice', at: '2026-08-21T10:00:00Z', text: 'first', attachments: 0 },
      {
        chat: CHAT,
        id: 'b',
        from: 'Alice',
        at: '2026-08-21T10:01:00Z',
        text: 'with a file',
        attachments: 1,
      },
    ]);
  });

  it('never echoes the signed-in account back as an inbox event', async () => {
    const store = chatStore({
      [CHAT]: [
        message({ id: 'own', fromId: me.id, from: me.displayName, text: 'my own post' }),
        message({ id: 'theirs', text: 'a question', createdDateTime: '2026-08-21T10:01:00Z' }),
      ],
    });

    await poller(store).pollOnce();

    expect((await inboxLines()).map((line) => line['id'])).toEqual(['theirs']);
  });

  it('drops deleted stubs and messages with no text and no attachments', async () => {
    const store = chatStore({
      [CHAT]: [
        message({ id: 'gone', isDeleted: true, text: '' }),
        message({ id: 'event', text: '   ', createdDateTime: '2026-08-21T10:01:00Z' }),
        message({ id: 'real', text: 'kept', createdDateTime: '2026-08-21T10:02:00Z' }),
      ],
    });

    await poller(store).pollOnce();

    expect((await inboxLines()).map((line) => line['id'])).toEqual(['real']);
  });

  it('does not re-emit delivered messages after a restart, thanks to the state sidecar', async () => {
    const store = chatStore({ [CHAT]: [message({ id: 'a' })] });
    await poller(store).pollOnce();

    // Fresh instance over the same files = a server restart.
    const restarted = poller(store);
    store.add(CHAT, message({ id: 'b', text: 'later', createdDateTime: '2026-08-21T10:05:00Z' }));
    await restarted.pollOnce();

    expect((await inboxLines()).map((line) => line['id'])).toEqual(['a', 'b']);
  });

  it('drops an exact replay of the newest message by id, not just by timestamp', async () => {
    // A port that ignores the watermark and always re-serves the same message.
    const replaying = {
      readMessages: () =>
        Promise.resolve({ messages: [message({ id: 'a' })], watermark: '2026-08-21T10:00:00Z' }),
    };
    const inbox = poller(replaying);

    await inbox.pollOnce();
    await inbox.pollOnce();

    expect((await inboxLines()).map((line) => line['id'])).toEqual(['a']);
  });

  it('truncates very long messages to 2000 characters', async () => {
    const store = chatStore({ [CHAT]: [message({ id: 'long', text: 'x'.repeat(5000) })] });

    await poller(store).pollOnce();

    const [line] = await inboxLines();
    expect(line?.['text']).toHaveLength(2000);
  });

  it('surfaces a dead poll as an error line instead of silence, and does not throw', async () => {
    const dead = {
      readMessages: () => Promise.reject(new Error('AADSTS50055: password expired')),
    };

    const clean = await poller(dead).pollOnce();

    // false = back off: everything failed, so this is auth/network death, not one bad chat.
    expect(clean).toBe(false);
    const [line] = await inboxLines();
    expect(line?.['error']).toMatch(/AADSTS50055/);
    expect(line?.['at']).toBeTruthy();
  });

  it('keeps delivering the healthy chats, at full speed, when one chat fails', async () => {
    const store = chatStore({ [OTHER]: [message({ id: 'ok', chatId: OTHER })] });
    const flaky = {
      readMessages: (chatId: string, since?: string) =>
        chatId === CHAT
          ? Promise.reject(new Error('chat gone'))
          : store.readMessages(chatId, since),
    };

    const clean = await poller(flaky, [CHAT, OTHER]).pollOnce();

    // true = no backoff: one chat the account cannot read must not slow the healthy ones down.
    expect(clean).toBe(true);
    const lines = await inboxLines();
    expect(lines.map((line) => line['id'] ?? 'error')).toEqual(['ok', 'error']);
  });

  it('writes an unchanged failure once, not once per poll, and again after recovery', async () => {
    let alive = false;
    const store = chatStore({});
    const port = {
      readMessages: (chatId: string, since?: string) =>
        alive ? store.readMessages(chatId, since) : Promise.reject(new Error('token expired')),
    };
    const inbox = poller(port);

    await inbox.pollOnce();
    await inbox.pollOnce();
    alive = true;
    await inbox.pollOnce();
    alive = false;
    await inbox.pollOnce();

    const errors = (await inboxLines()).filter((line) => line['error'] !== undefined);
    expect(errors).toHaveLength(2);
  });

  it('starts from scratch when the state file is corrupt rather than refusing to poll', async () => {
    const store = chatStore({ [CHAT]: [message({ id: 'a' })] });
    await poller(store).pollOnce();

    const { writeFile } = await import('node:fs/promises');
    await writeFile(statePath, 'not json at all');
    const clean = await poller(store).pollOnce();

    expect(clean).toBe(true);
    // The one acceptable duplicate: losing the state file re-reads the recent window once.
    expect((await inboxLines()).map((line) => line['id'])).toEqual(['a', 'a']);
  });
});

describe('inbox poller — behaviour under throttle (2026-08-25)', () => {
  let dir: string;
  let inboxPath: string;
  let statePath: string;
  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'inbox-throttle-'));
    inboxPath = join(dir, 'inbox.jsonl');
    statePath = join(dir, 'inbox-state.json');
  });
  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  let clock = 1_000_000;
  function pollerOver(readMessages: (chatId: string, since?: string) => Promise<ReturnType<typeof applyWatermark>>, chatIds: string[]) {
    return new InboxPoller({
      chats: { readMessages },
      allowlist: new ChatAllowlist(chatIds.map((id) => ({ id, label: id, canPost: true }))),
      self: () => Promise.resolve(me),
      inboxPath,
      statePath,
      nowFn: () => clock,
    });
  }

  it('the first 429 ends the cycle — later chats are not asked — and the cycle counts as unclean so start() backs off', async () => {
    const asked: string[] = [];
    const readMessages = async (chatId: string) => {
      asked.push(chatId);
      const err = Object.assign(new Error('Too many requests'), { status: 429 });
      throw err;
    };
    const p = pollerOver(readMessages, ['19:one@thread.v2', '19:two@thread.v2', '19:three@thread.v2']);

    const clean = await p.pollOnce();

    expect(asked).toEqual(['19:one@thread.v2']);
    expect(clean).toBe(false);
  });

  it('a healthy chat followed by a throttled one: the cycle is NOT clean even though one of two succeeded', async () => {
    const readMessages = async (chatId: string, since?: string) => {
      if (chatId.includes('two')) throw Object.assign(new Error('Too many requests'), { status: 429 });
      return applyWatermark([], since);
    };
    const p = pollerOver(readMessages, ['19:one@thread.v2', '19:two@thread.v2']);

    expect(await p.pollOnce()).toBe(false);
  });

  it('a chat answering 403 is parked: not asked again on the next cycle, while healthy chats still are', async () => {
    const asked: string[] = [];
    const readMessages = async (chatId: string, since?: string) => {
      asked.push(chatId);
      if (chatId === '19:forbidden@thread.v2') {
        throw Object.assign(new Error('InsufficientPrivileges'), { status: 403 });
      }
      return applyWatermark([], since);
    };
    const p = pollerOver(readMessages, ['19:forbidden@thread.v2', '19:ok@thread.v2']);

    await p.pollOnce();
    await p.pollOnce();

    expect(asked).toEqual(['19:forbidden@thread.v2', '19:ok@thread.v2', '19:ok@thread.v2']);
  });
});

describe('inbox poller — the health file (2026-09-01)', () => {
  // The incident this exists for: after a host reboot one project's poller was dead while a
  // second project's daemon (same dist/index.js, different env) survived, and a pgrep-based
  // check matched the survivor — the dead pipeline read as "up" for 1.5 hours. The health file
  // is the liveness contract that replaces process-pattern guessing: its age answers "is THIS
  // inbox's pipeline alive", whoever else happens to be running.
  let dir: string;
  let inboxPath: string;
  let statePath: string;
  let clock: number;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'inbox-health-'));
    inboxPath = join(dir, 'inbox.jsonl');
    statePath = join(dir, 'inbox-state.json');
    clock = 1_756_700_000_000; // 2026-09-01-ish, so the ISO timestamps read like the incident
  });
  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  function pollerWith(
    readMessages: (chatId: string, since?: string) => Promise<ReturnType<typeof applyWatermark>>,
    overrides: { healthPath?: string } = {},
  ) {
    return new InboxPoller({
      chats: { readMessages },
      allowlist: new ChatAllowlist([{ id: CHAT, label: 'pilot', canPost: true }]),
      self: () => Promise.resolve(me),
      inboxPath,
      statePath,
      pid: 4242,
      nowFn: () => clock,
      ...overrides,
    });
  }

  async function health(): Promise<Record<string, unknown>> {
    return JSON.parse(await readFile(join(dir, 'poller-health.json'), 'utf8')) as Record<
      string,
      unknown
    >;
  }

  it('writes poller-health.json beside the inbox after a clean poll', async () => {
    const p = pollerWith((chatId, since) => Promise.resolve(applyWatermark([], since)));

    await p.pollOnce();

    expect(await health()).toEqual({
      pid: 4242,
      inboxPath,
      lastAttemptAt: new Date(clock).toISOString(),
      lastSuccessAt: new Date(clock).toISOString(),
      ok: true,
      consecutiveFailures: 0,
      backoffMs: DEFAULT_POLL_MS,
    });
  });

  it('keeps writing after a FAILED poll — a dead pipeline must age visibly, not freeze the file', async () => {
    let alive = true;
    const p = pollerWith((chatId, since) =>
      alive ? Promise.resolve(applyWatermark([], since)) : Promise.reject(new Error('token expired')),
    );

    await p.pollOnce();
    const successAt = new Date(clock).toISOString();
    alive = false;
    clock += 30_000;
    await p.pollOnce();
    clock += 30_000;
    await p.pollOnce();

    const snapshot = await health();
    expect(snapshot['ok']).toBe(false);
    expect(snapshot['consecutiveFailures']).toBe(2);
    expect(snapshot['lastAttemptAt']).toBe(new Date(clock).toISOString());
    // The last GOOD poll's stamp survives failures, so a watcher can say how long the outage is.
    expect(snapshot['lastSuccessAt']).toBe(successAt);
    // First failure waited the base interval; the second doubled it. backoffMs is the delay
    // until the next attempt, so a watcher knows how stale the file may legitimately get.
    expect(snapshot['backoffMs']).toBe(DEFAULT_POLL_MS * 2);
  });

  it('leaves no .tmp file behind — the write is tmp + rename, never a partial health file', async () => {
    const p = pollerWith((chatId, since) => Promise.resolve(applyWatermark([], since)));

    await p.pollOnce();

    const leftovers = (await readdir(dir)).filter((name) => name.includes('.tmp'));
    expect(leftovers).toEqual([]);
  });

  it('a health write failure never throws past pollOnce, and the inbox still gets its lines', async () => {
    const store = chatStore({ [CHAT]: [message({ id: 'a' })] });
    // A health path in a directory that does not exist: every write attempt fails.
    const p = pollerWith((chatId, since) => store.readMessages(chatId, since), {
      healthPath: join(dir, 'no-such-dir', 'poller-health.json'),
    });

    await expect(p.pollOnce()).resolves.toBe(true);

    const raw = await readFile(inboxPath, 'utf8');
    expect(raw).toContain('"id":"a"');
  });
});

describe('inbox poller — escalating error lines and recovery (2026-09-01)', () => {
  // The forever-dedupe wrote ONE line for a five-hour outage. Correct against flooding, useless
  // for a watcher asking "is this still going on?" — so the dedupe signature now carries an
  // escalation bucket and the thresholds re-surface the failure with a running count.
  let dir: string;
  let inboxPath: string;
  let statePath: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'inbox-escalate-'));
    inboxPath = join(dir, 'inbox.jsonl');
    statePath = join(dir, 'inbox-state.json');
  });
  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  function pollerOver(readMessages: (chatId: string, since?: string) => Promise<ReturnType<typeof applyWatermark>>) {
    return new InboxPoller({
      chats: { readMessages },
      allowlist: new ChatAllowlist([{ id: CHAT, label: 'pilot', canPost: true }]),
      self: () => Promise.resolve(me),
      inboxPath,
      statePath,
    });
  }

  async function lines(): Promise<Array<Record<string, unknown>>> {
    const raw = await readFile(inboxPath, 'utf8').catch(() => '');
    return raw
      .split('\n')
      .filter((line) => line.trim() !== '')
      .map((line) => JSON.parse(line) as Record<string, unknown>);
  }

  it('an unchanged outage re-surfaces at the thresholds: first failure, then 10, then 50', async () => {
    const p = pollerOver(() => Promise.reject(new Error('token expired')));

    for (let i = 0; i < 60; i += 1) {
      await p.pollOnce();
    }

    const errors = (await lines()).filter((line) => line['error'] !== undefined);
    expect(errors).toHaveLength(3);
    expect(errors.map((line) => line['consecutiveFailures'])).toEqual([1, 10, 50]);
    expect(errors[1]?.['error']).toMatch(/still failing after 10 polls/);
    expect(errors[2]?.['error']).toMatch(/still failing after 50 polls/);
    expect(errors.every((line) => typeof line['at'] === 'string')).toBe(true);
  });

  it('recovery after a long outage is announced, so the watcher sees the pipeline come back', async () => {
    let alive = false;
    const p = pollerOver((chatId, since) =>
      alive ? Promise.resolve(applyWatermark([], since)) : Promise.reject(new Error('token expired')),
    );

    for (let i = 0; i < 12; i += 1) {
      await p.pollOnce();
    }
    alive = true;
    await p.pollOnce();

    const recovered = (await lines()).filter((line) => line['recovered'] !== undefined);
    expect(recovered).toEqual([
      { recovered: true, at: expect.stringMatching(/T/) as unknown, afterFailures: 12 },
    ]);
  });

  it('a short blip stays quiet: recovery after fewer than 10 failures writes no recovered line', async () => {
    let alive = false;
    const p = pollerOver((chatId, since) =>
      alive ? Promise.resolve(applyWatermark([], since)) : Promise.reject(new Error('token expired')),
    );

    await p.pollOnce();
    await p.pollOnce();
    alive = true;
    await p.pollOnce();

    expect((await lines()).filter((line) => line['recovered'] !== undefined)).toEqual([]);
  });
});

describe('inbox poller — Retry-After as the backoff floor (2026-09-01)', () => {
  // When two daemons raced on one shared account, Graph named waits the blind doubling ignored:
  // the poller came back sooner than asked and fed the penalty window. GraphError already
  // carries retryAfterSeconds, so the next delay honours it as a floor over the doubled backoff.
  let dir: string;
  let inboxPath: string;
  let statePath: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'inbox-retry-after-'));
    inboxPath = join(dir, 'inbox.jsonl');
    statePath = join(dir, 'inbox-state.json');
  });
  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  async function backoffMs(): Promise<number> {
    const raw = JSON.parse(await readFile(join(dir, 'poller-health.json'), 'utf8')) as {
      backoffMs: number;
    };
    return raw.backoffMs;
  }

  it('a 429 naming Retry-After lifts the next delay to at least that wait', async () => {
    const p = new InboxPoller({
      chats: {
        readMessages: () =>
          Promise.reject(
            Object.assign(new Error('Too many requests'), { status: 429, retryAfterSeconds: 120 }),
          ),
      },
      allowlist: new ChatAllowlist([{ id: CHAT, label: 'pilot', canPost: true }]),
      self: () => Promise.resolve(me),
      inboxPath,
      statePath,
    });

    await p.pollOnce();

    // Blind doubling alone would wait DEFAULT_POLL_MS (30s); Graph asked for 120s.
    expect(await backoffMs()).toBe(120_000);
  });

  it('a 429 without Retry-After keeps the plain doubling', async () => {
    const p = new InboxPoller({
      chats: {
        readMessages: () =>
          Promise.reject(Object.assign(new Error('Too many requests'), { status: 429 })),
      },
      allowlist: new ChatAllowlist([{ id: CHAT, label: 'pilot', canPost: true }]),
      self: () => Promise.resolve(me),
      inboxPath,
      statePath,
    });

    await p.pollOnce();

    expect(await backoffMs()).toBe(DEFAULT_POLL_MS);
  });
});

describe('inbox poller — the clean verdict counts only the chats actually asked (review round 1)', () => {
  let dir: string; let inboxPath: string; let statePath: string; let clock = 5_000_000;
  beforeEach(async () => { dir = await mkdtemp(join(tmpdir(), 'inbox-denom-')); inboxPath = join(dir, 'inbox.jsonl'); statePath = join(dir, 'inbox-state.json'); });
  afterEach(async () => { await rm(dir, { recursive: true, force: true }); });

  it('two parked 403 chats plus the only live chat dying with 401: NOT clean (auth death must back off)', async () => {
    const readMessages = async (chatId: string, since?: string) => {
      if (chatId.includes('forbidden')) throw Object.assign(new Error('InsufficientPrivileges'), { status: 403 });
      if (clock > 5_000_000) throw Object.assign(new Error('token expired'), { status: 401 });
      return applyWatermark([], since);
    };
    const p = new InboxPoller({
      chats: { readMessages },
      allowlist: new ChatAllowlist(['19:forbidden-1@thread.v2', '19:forbidden-2@thread.v2', '19:live@thread.v2'].map((id) => ({ id, label: id, canPost: true }))),
      self: () => Promise.resolve(me), inboxPath, statePath, nowFn: () => clock,
    });

    await p.pollOnce();          // parks both 403 chats; live chat fine
    clock += 1_000;
    const clean = await p.pollOnce(); // only the live chat is asked, and it fails

    expect(clean).toBe(false);
  });

  it('a parked chat is asked again once the park expires', async () => {
    const asked: string[] = [];
    const readMessages = async (chatId: string, _since?: string) => {
      asked.push(chatId);
      throw Object.assign(new Error('InsufficientPrivileges'), { status: 403 });
    };
    const p = new InboxPoller({
      chats: { readMessages },
      allowlist: new ChatAllowlist([{ id: '19:forbidden@thread.v2', label: 'f', canPost: true }]),
      self: () => Promise.resolve(me), inboxPath, statePath, nowFn: () => clock,
    });

    await p.pollOnce();
    clock += 60_000;  await p.pollOnce();             // inside the park: skipped
    clock += PARK_FORBIDDEN_CHAT_MS; await p.pollOnce(); // park expired: asked again

    expect(asked).toHaveLength(2);
  });
});
