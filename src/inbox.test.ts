import { mkdtemp, readFile, readdir, rename, rm, writeFile } from 'node:fs/promises';
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

  // 0.4.1: a chat with no watermark on record — because this is the very first poll of a fresh
  // install, or the state was lost/corrupted — is treated as "history unknown", not "empty".
  // That first poll only ESTABLISHES the watermark; nothing already sitting in the chat at that
  // moment is ever delivered. So every test below that wants to see delivery runs an initial
  // "settling" poll first (harmless: nothing is emitted), exactly the way a real fresh install's
  // first cycle behaves, then asserts on what a SECOND poll delivers.
  async function settle(store: ReturnType<typeof chatStore>, chatIds = [CHAT]): Promise<void> {
    // Poll once against whatever the store holds RIGHT NOW, before the messages under test are
    // added, so the settle itself proves nothing about the chat's true starting content.
    await poller(store, chatIds).pollOnce();
  }

  it('a brand-new chat (or a lost/corrupt sidecar) never backfills what already existed — starts from NOW (0.4.1)', async () => {
    const store = chatStore({
      [CHAT]: [message({ id: 'ancient', text: 'predates this process', createdDateTime: '2026-08-01T00:00:00Z' })],
    });

    await poller(store).pollOnce(); // the very first poll: settles the watermark, delivers nothing

    expect(await inboxLines()).toEqual([]);
  });

  it('appends one JSON line per new message, in the daemon-compatible shape', async () => {
    const store = chatStore({});
    await settle(store);
    store.add(
      CHAT,
      message({ id: 'a', text: 'first', createdDateTime: '2026-08-21T10:00:00Z' }),
      message({
        id: 'b',
        text: 'with a file',
        createdDateTime: '2026-08-21T10:01:00Z',
        attachments: [{ id: 'att-1', name: 'plan.xlsx' }],
      }),
    );

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
    const store = chatStore({});
    await settle(store);
    store.add(
      CHAT,
      message({ id: 'own', fromId: me.id, from: me.displayName, text: 'my own post' }),
      message({ id: 'theirs', text: 'a question', createdDateTime: '2026-08-21T10:01:00Z' }),
    );

    await poller(store).pollOnce();

    expect((await inboxLines()).map((line) => line['id'])).toEqual(['theirs']);
  });

  it('drops deleted stubs and messages with no text and no attachments', async () => {
    const store = chatStore({});
    await settle(store);
    store.add(
      CHAT,
      message({ id: 'gone', isDeleted: true, text: '' }),
      message({ id: 'event', text: '   ', createdDateTime: '2026-08-21T10:01:00Z' }),
      message({ id: 'real', text: 'kept', createdDateTime: '2026-08-21T10:02:00Z' }),
    );

    await poller(store).pollOnce();

    expect((await inboxLines()).map((line) => line['id'])).toEqual(['real']);
  });

  it('does not re-emit delivered messages after a restart, thanks to the state sidecar', async () => {
    const store = chatStore({});
    await settle(store);
    store.add(CHAT, message({ id: 'a' }));
    await poller(store).pollOnce(); // 'a' genuinely delivered, pre-restart

    // Fresh instance over the same files = a server restart.
    const restarted = poller(store);
    store.add(CHAT, message({ id: 'b', text: 'later', createdDateTime: '2026-08-21T10:05:00Z' }));
    await restarted.pollOnce();

    expect((await inboxLines()).map((line) => line['id'])).toEqual(['a', 'b']);
  });

  it('drops an exact replay of the newest message by id, not just by timestamp', async () => {
    // A port that ignores the watermark and always re-serves the same message — simulates Graph
    // returning a message exactly at the watermark boundary again, which timestamp comparison
    // alone cannot distinguish from a genuinely new message sharing that timestamp.
    const replaying = {
      readMessages: () =>
        Promise.resolve({ messages: [message({ id: 'a' })], watermark: '2026-08-21T10:00:00Z' }),
    };
    const inbox = poller(replaying);

    await inbox.pollOnce(); // settles: 'a' establishes the watermark, delivered to no one
    await inbox.pollOnce(); // same 'a' served again — the id check must still drop it

    expect(await inboxLines()).toEqual([]);
  });

  it('truncates very long messages to 2000 characters', async () => {
    const store = chatStore({});
    await settle(store);
    store.add(CHAT, message({ id: 'long', text: 'x'.repeat(5000) }));

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
    const store = chatStore({});
    await settle(store, [CHAT, OTHER]);
    store.add(OTHER, message({ id: 'ok', chatId: OTHER }));
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

  it('a corrupt state file degrades to starting from now, never replaying full history (0.4.1)', async () => {
    const store = chatStore({});
    await settle(store);
    store.add(CHAT, message({ id: 'a' }));
    await poller(store).pollOnce(); // 'a' genuinely delivered

    const { writeFile } = await import('node:fs/promises');
    await writeFile(statePath, 'not json at all'); // sidecar lost
    store.add(CHAT, message({ id: 'b', text: 'arrives after the corruption', createdDateTime: '2026-08-21T10:10:00Z' }));
    const clean = await poller(store).pollOnce(); // treated as a fresh install: settles, delivers nothing

    expect(clean).toBe(true);
    expect((await inboxLines()).map((line) => line['id'])).toEqual(['a']); // 'a' never replayed, 'b' not yet either

    store.add(CHAT, message({ id: 'c', text: 'genuinely new, after the re-settle', createdDateTime: '2026-08-21T10:20:00Z' }));
    await poller(store).pollOnce(); // a normal poll now — 'c' postdates the re-settled watermark

    // 'b' was itself the newest message present AT the re-settle poll, so it became the new
    // watermark baseline rather than being delivered — same "settles, does not deliver" contract
    // as the very first bootstrap. Only 'c', which arrived strictly after that, is new.
    expect((await inboxLines()).map((line) => line['id'])).toEqual(['a', 'c']);
  });

  it('leaves no tmp file behind after a successful poll', async () => {
    const store = chatStore({});
    await settle(store);
    store.add(CHAT, message({ id: 'a' }));
    await poller(store).pollOnce();

    const { readdir } = await import('node:fs/promises');
    const entries = await readdir(join(dir, 'nested'));

    expect(entries.filter((name) => name.includes('.tmp-'))).toEqual([]);
    expect(entries).toContain('inbox-state.json');
  });

  // MAJOR 4 (review round 1): the test above only proves no leftover tmp file — that passes
  // identically whether saveState writes via temp+rename OR writes the destination directly
  // (mutation-verified: deleting the rename and writing straight to statePath left it green).
  // Injecting the write primitives is what makes "the destination is reached ONLY via rename" an
  // honest, mutation-proof observable.
  it('the state sidecar is reached ONLY via rename, never a direct write (atomicity, 0.4.1)', async () => {
    const directWrites: string[] = [];
    const renames: Array<{ from: string; to: string }> = [];
    const store = chatStore({});
    const p = new InboxPoller({
      chats: store,
      allowlist: new ChatAllowlist([{ id: CHAT, label: CHAT, canPost: true }]),
      self: () => Promise.resolve(me),
      inboxPath,
      statePath,
      writeFileFn: (async (target: Parameters<typeof writeFile>[0], ...rest: unknown[]) => {
        directWrites.push(String(target));
        return (writeFile as (...args: unknown[]) => Promise<void>)(target, ...rest);
      }) as typeof writeFile,
      renameFn: (async (from: Parameters<typeof rename>[0], to: Parameters<typeof rename>[1]) => {
        renames.push({ from: String(from), to: String(to) });
        return rename(from, to);
      }) as typeof rename,
    });

    await p.pollOnce();

    expect(directWrites).toHaveLength(1);
    expect(directWrites[0]).not.toBe(statePath); // never written to the real destination directly
    expect(renames).toEqual([{ from: directWrites[0], to: statePath }]);
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

describe('inbox poller — the health file (carried over from PR #8 onto 0.5.2)', () => {
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
    overrides: { healthPath?: string; yieldPath?: string } = {},
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
    const store = chatStore();
    // A health path in a directory that does not exist: every write attempt fails.
    const p = pollerWith((chatId, since) => store.readMessages(chatId, since), {
      healthPath: join(dir, 'no-such-dir', 'poller-health.json'),
    });
    // 0.4.1: the first poll on an unknown chat only settles the watermark and delivers nothing —
    // so the message this test cares about is added only AFTER that settling poll.
    await p.pollOnce();
    store.add(CHAT, message({ id: 'a' }));

    await expect(p.pollOnce()).resolves.toBe(true);

    const raw = await readFile(inboxPath, 'utf8');
    expect(raw).toContain('"id":"a"');
  });

  it('a cycle skipped for the quota yield is marked distinctly — neither a fresh success nor a failure (0.5.4: PR #8 predates the yield)', async () => {
    const yieldPath = join(dir, 'inbox-yield.json');
    const p = pollerWith((chatId, since) => Promise.resolve(applyWatermark([], since)), { yieldPath });

    await p.pollOnce(); // clean, no yield yet — establishes lastSuccessAt
    const successAt = new Date(clock).toISOString();

    clock += 30_000;
    await writeFile(
      yieldPath,
      JSON.stringify({ pid: 999, reason: 'teams-attachments', until: clock + 60_000 }),
    );
    await p.pollOnce(); // yielded

    const snapshot = await health();
    expect(snapshot['ok']).toBe(true);
    expect(snapshot['yielded']).toBe(true);
    expect(snapshot['lastAttemptAt']).toBe(new Date(clock).toISOString());
    // A fresh success would have bumped lastSuccessAt to NOW; a yielded cycle must not — it is
    // politeness, not a poll that actually happened.
    expect(snapshot['lastSuccessAt']).toBe(successAt);
    // Nor may it look like a failure: no error line, no failure count.
    expect(snapshot['consecutiveFailures']).toBe(0);
  });
});

describe('inbox poller — escalating error lines and recovery (carried over from PR #8 onto 0.5.2)', () => {
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

describe('inbox poller — Retry-After as the backoff floor (carried over from PR #8 onto 0.5.2)', () => {
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

describe('inbox poller — stuck-auth self-healing (0.4.1, live-diagnosed: only a process restart recovered)', () => {
  let dir: string; let inboxPath: string; let statePath: string;
  beforeEach(async () => { dir = await mkdtemp(join(tmpdir(), 'inbox-authstuck-')); inboxPath = join(dir, 'inbox.jsonl'); statePath = join(dir, 'inbox-state.json'); });
  afterEach(async () => { await rm(dir, { recursive: true, force: true }); });

  function authError(): Error {
    return Object.assign(new Error('token expired'), { status: 401 });
  }

  it('N consecutive auth-shaped poll failures force a token re-mint exactly once', async () => {
    let onAuthStuckCalls = 0;
    const readMessages = () => Promise.reject(authError());
    const p = new InboxPoller({
      chats: { readMessages },
      allowlist: new ChatAllowlist([{ id: CHAT, label: 'pilot', canPost: true }]),
      self: () => Promise.resolve(me),
      inboxPath,
      statePath,
      authFailureThreshold: 3,
      onAuthStuck: () => {
        onAuthStuckCalls += 1;
      },
    });

    await p.pollOnce();
    expect(onAuthStuckCalls).toBe(0);
    await p.pollOnce();
    expect(onAuthStuckCalls).toBe(0);
    await p.pollOnce(); // the 3rd consecutive auth-shaped failure
    expect(onAuthStuckCalls).toBe(1);
  });

  it('a successful poll resets the streak — three failures, one success, two more failures never fires', async () => {
    let alive = false;
    let onAuthStuckCalls = 0;
    const store = chatStore({});
    const readMessages = (chatId: string, since?: string) =>
      alive ? store.readMessages(chatId, since) : Promise.reject(authError());
    const p = new InboxPoller({
      chats: { readMessages },
      allowlist: new ChatAllowlist([{ id: CHAT, label: 'pilot', canPost: true }]),
      self: () => Promise.resolve(me),
      inboxPath,
      statePath,
      authFailureThreshold: 3,
      onAuthStuck: () => {
        onAuthStuckCalls += 1;
      },
    });

    await p.pollOnce();
    await p.pollOnce();
    alive = true;
    await p.pollOnce(); // recovers — resets the streak
    alive = false;
    await p.pollOnce();
    await p.pollOnce();

    expect(onAuthStuckCalls).toBe(0);
  });

  // MAJOR 5 (review round 1, evidence from the orchestrator's incident logs): the live symptom
  // was literally `inbox poll failed: <chat>: fetch failed`, repeatedly, with NO status code
  // visible in the log, while Graph answered 200 to parallel curl probes — recovered only by a
  // process restart. A detector that only fires on 401/invalid_grant/AADSTS/token-expiry text
  // would never have fired on the real incident. This is now a LAST-RESORT tier: N consecutive
  // poll failures of ANY shape also force a re-mint — a spurious re-mint during a genuine network
  // outage is cheap and harmless (the very next successful getAccessToken() call just does one
  // extra password grant instead of reusing a cache); staying stuck is not. This test used to
  // assert the OPPOSITE (never fires) — that assertion was locking the incident's exact shape out
  // of the detector.
  it('N consecutive failures of ANY shape (no status, no auth vocabulary) — the last-resort tier still forces a re-mint', async () => {
    let onAuthStuckCalls = 0;
    const readMessages = () => Promise.reject(new Error('fetch failed'));
    const p = new InboxPoller({
      chats: { readMessages },
      allowlist: new ChatAllowlist([{ id: CHAT, label: 'pilot', canPost: true }]),
      self: () => Promise.resolve(me),
      inboxPath,
      statePath,
      authFailureThreshold: 3,
      onAuthStuck: () => {
        onAuthStuckCalls += 1;
      },
    });

    await p.pollOnce();
    expect(onAuthStuckCalls).toBe(0);
    await p.pollOnce();
    expect(onAuthStuckCalls).toBe(0);
    await p.pollOnce(); // 3rd consecutive failure, shapeless — the last-resort tier
    expect(onAuthStuckCalls).toBe(1);
  });

  it('a poller with no onAuthStuck configured never throws, even past the threshold', async () => {
    const readMessages = () => Promise.reject(authError());
    const p = new InboxPoller({
      chats: { readMessages },
      allowlist: new ChatAllowlist([{ id: CHAT, label: 'pilot', canPost: true }]),
      self: () => Promise.resolve(me),
      inboxPath,
      statePath,
      authFailureThreshold: 2,
    });

    await expect(p.pollOnce()).resolves.toBe(false);
    await expect(p.pollOnce()).resolves.toBe(false);
    await expect(p.pollOnce()).resolves.toBe(false);
  });

  // MAJOR 3 (review round 1): resetting the streak the instant onAuthStuck is CALLED claims an
  // outcome (recovery) the code cannot actually know yet — invalidate() only flags the token for
  // re-auth; the real HTTP re-mint happens on the NEXT getAccessToken() call, whose result shows
  // up as THIS poller's next poll. The streak must not reset until a poll actually comes back
  // clean, and firing must not repeat on every single poll once past the threshold.
  it('the remedy fires exactly once per failing streak — not once per poll past the threshold — and only resets on an actual clean poll', async () => {
    let alive = false;
    let onAuthStuckCalls = 0;
    const store = chatStore({});
    const readMessages = (chatId: string, since?: string) =>
      alive ? store.readMessages(chatId, since) : Promise.reject(authError());
    const p = new InboxPoller({
      chats: { readMessages },
      allowlist: new ChatAllowlist([{ id: CHAT, label: 'pilot', canPost: true }]),
      self: () => Promise.resolve(me),
      inboxPath,
      statePath,
      authFailureThreshold: 3,
      onAuthStuck: () => {
        onAuthStuckCalls += 1;
      },
    });

    await p.pollOnce(); // 1
    await p.pollOnce(); // 2
    await p.pollOnce(); // 3 — fires
    expect(onAuthStuckCalls).toBe(1);
    // Keep failing well past a second threshold's worth with NO recovery in between. A design
    // that reset the streak the instant it fires would count another 3 straight to 6 and fire
    // AGAIN here — this is exactly the case that distinguishes "reset on fire" from "reset only
    // on an actual clean poll".
    await p.pollOnce(); // 4
    await p.pollOnce(); // 5
    await p.pollOnce(); // 6 — would be a 2nd trigger under "reset on fire"
    await p.pollOnce(); // 7
    expect(onAuthStuckCalls).toBe(1); // still just once — no clean poll has happened yet

    alive = true;
    await p.pollOnce(); // recovers — THIS is what may reset the streak, not the earlier firing
    alive = false;
    await p.pollOnce(); // 1 of a NEW streak
    await p.pollOnce(); // 2
    expect(onAuthStuckCalls).toBe(1);
    await p.pollOnce(); // 3 of the new streak — fires again
    expect(onAuthStuckCalls).toBe(2);
  });

  // MAJOR 3 / MINOR (log fires, injected dep): the log line must describe what is actually known
  // at each point — a request was made, the remedy has not yet cleared it, or it recovered — not
  // a single "forcing…" claim asserted once and never revisited.
  it('the log line reports outcome, not just intent: requested, then still-failing, then recovered', async () => {
    let alive = false;
    const store = chatStore({});
    const readMessages = (chatId: string, since?: string) =>
      alive ? store.readMessages(chatId, since) : Promise.reject(authError());
    const lines: string[] = [];
    const p = new InboxPoller({
      chats: { readMessages },
      allowlist: new ChatAllowlist([{ id: CHAT, label: 'pilot', canPost: true }]),
      self: () => Promise.resolve(me),
      inboxPath,
      statePath,
      authFailureThreshold: 3,
      onAuthStuck: () => {},
      log: (line) => lines.push(line),
    });

    await p.pollOnce();
    await p.pollOnce();
    await p.pollOnce(); // fires
    expect(lines.some((l) => /requested a forced token re-authentication/.test(l))).toBe(true);

    await p.pollOnce(); // still failing — the remedy has not visibly worked yet
    expect(lines.some((l) => /still failing after a forced token re-authentication/.test(l))).toBe(true);

    alive = true;
    await p.pollOnce();
    expect(lines.some((l) => /recovered after a forced token re-authentication/.test(l))).toBe(true);
  });
});

describe('inbox poller — the quota yield (0.5.0: the poller starved ad-hoc readers, measured 2026-09-02)', () => {
  let dir: string;
  let inboxPath: string;
  let statePath: string;
  let yieldPath: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'inbox-yield-poll-test-'));
    inboxPath = join(dir, 'inbox.jsonl');
    statePath = join(dir, 'inbox-state.json');
    yieldPath = join(dir, 'inbox-yield.json');
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  function yieldingPoller(chats: { readMessages: (chatId: string, since?: string) => Promise<ReturnType<typeof applyWatermark>> }, log: (line: string) => void = () => {}) {
    return new InboxPoller({
      chats,
      allowlist: new ChatAllowlist([{ id: CHAT, label: CHAT, canPost: true }]),
      self: () => Promise.resolve(me),
      inboxPath,
      statePath,
      yieldPath,
      log,
    });
  }

  it('a standing yield skips the whole cycle — not one Graph read — and counts as CLEAN, so backoff never doubles', async () => {
    let reads = 0;
    const chats = {
      readMessages: () => {
        reads += 1;
        return Promise.resolve(applyWatermark([]));
      },
    };
    await writeFile(yieldPath, JSON.stringify({ pid: 4321, reason: 'teams-attachments', until: Date.now() + 60_000 }));

    const clean = await yieldingPoller(chats).pollOnce();

    expect(clean).toBe(true); // being polite is not failing: a doubled backoff would outlast the yield
    expect(reads).toBe(0);
  });

  it('one yield episode logs once, not once per cycle', async () => {
    const lines: string[] = [];
    const chats = { readMessages: () => Promise.resolve(applyWatermark([])) };
    await writeFile(yieldPath, JSON.stringify({ pid: 4321, reason: 'teams-attachments', until: Date.now() + 60_000 }));

    const poller = yieldingPoller(chats, (line) => lines.push(line));
    await poller.pollOnce();
    await poller.pollOnce();
    await poller.pollOnce();

    const yieldLines = lines.filter((line) => line.includes('yielding the Graph read quota'));
    expect(yieldLines).toHaveLength(1);
    expect(yieldLines[0]).toContain('teams-attachments');
    expect(yieldLines[0]).toContain('4321');
  });

  it('an expired or absent yield polls normally — a crashed reader cannot silence the inbox past its deadline', async () => {
    let reads = 0;
    const chats = {
      readMessages: () => {
        reads += 1;
        return Promise.resolve(applyWatermark([]));
      },
    };
    await writeFile(yieldPath, JSON.stringify({ pid: 4321, reason: 'crashed', until: Date.now() - 1 }));

    await yieldingPoller(chats).pollOnce();
    expect(reads).toBe(1);

    await rm(yieldPath);
    await yieldingPoller(chats).pollOnce();
    expect(reads).toBe(2);
  });
});

// Mitigation 2 (docs/throttling-mitigation.md §4, stage 1 item 2): every message the poller
// already reads carries its sender's AAD id and display name at zero MARGINAL Graph cost — this
// is exactly the population worth @mentioning, and harvesting it is what lets a chat's roster
// fill and refresh from traffic alone, taking `/members` off the mention-resolution send path
// for good (see teams-chats.test.ts's stale-serve tests for the other half of this mitigation).
describe('inbox poller — roster harvest from poll results (mitigation 2)', () => {
  let dir: string;
  let inboxPath: string;
  let statePath: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'inbox-roster-test-'));
    inboxPath = join(dir, 'inbox.jsonl');
    statePath = join(dir, 'inbox-state.json');
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  async function inboxLines(): Promise<Array<Record<string, unknown>>> {
    const raw = await readFile(inboxPath, 'utf8').catch(() => '');
    return raw
      .split('\n')
      .filter((line) => line.trim() !== '')
      .map((line) => JSON.parse(line) as Record<string, unknown>);
  }

  function poller(chats: Pick<ReturnType<typeof chatStore>, 'readMessages'>, chatIds = [CHAT]) {
    return new InboxPoller({
      chats,
      allowlist: new ChatAllowlist(chatIds.map((id) => ({ id, label: id, canPost: true }))),
      self: () => Promise.resolve(me),
      inboxPath,
      statePath,
    });
  }

  function rosterSink() {
    const merged: Array<{ chatId: string; members: ReadonlyArray<{ id: string; displayName: string }> }> = [];
    return {
      merged,
      merge: (chatId: string, members: ReadonlyArray<{ id: string; displayName: string }>) => {
        merged.push({ chatId, members: [...members] });
      },
    };
  }

  function pollerWithRoster(
    chats: Pick<ReturnType<typeof chatStore>, 'readMessages'>,
    roster: { merge: (chatId: string, members: ReadonlyArray<{ id: string; displayName: string }>) => void },
    chatIds = [CHAT],
  ) {
    return new InboxPoller({
      chats,
      allowlist: new ChatAllowlist(chatIds.map((id) => ({ id, label: id, canPost: true }))),
      self: () => Promise.resolve(me),
      inboxPath,
      statePath,
      roster,
    });
  }

  it('TRIGGERING: harvests (fromId, from) from every new message into the roster sink for that chat, at zero extra Graph calls', async () => {
    const store = chatStore({});
    const roster = rosterSink();
    await pollerWithRoster(store, roster).pollOnce(); // settle
    store.add(
      CHAT,
      message({ id: 'a', from: 'Alice', fromId: 'alice-id', createdDateTime: '2026-08-21T10:00:00Z' }),
      message({ id: 'b', from: 'Bob', fromId: 'bob-id', createdDateTime: '2026-08-21T10:01:00Z' }),
    );

    await pollerWithRoster(store, roster).pollOnce();

    const forChat = roster.merged.filter((entry) => entry.chatId === CHAT);
    expect(forChat).toHaveLength(1);
    expect(forChat[0]?.members).toEqual(
      expect.arrayContaining([
        { id: 'alice-id', displayName: 'Alice' },
        { id: 'bob-id', displayName: 'Bob' },
      ]),
    );
  });

  it('NON-TRIGGERING: an inbox poller with no roster sink configured never throws — harvest is optional, not a hard dependency', async () => {
    const store = chatStore({});
    await poller(store).pollOnce(); // settle, no `roster` in deps at all
    store.add(CHAT, message({ id: 'a', from: 'Alice', fromId: 'alice-id' }));

    const clean = await poller(store).pollOnce();

    expect(clean).toBe(true);
  });

  it('harvests from a BOOTSTRAP (settling) poll too — the very message the settling poll intentionally does not deliver to the inbox is still worth learning the sender of', async () => {
    const roster = rosterSink();
    const store = chatStore({
      [CHAT]: [message({ id: 'ancient', from: 'Ancient Sender', fromId: 'ancient-id' })],
    });

    await pollerWithRoster(store, roster).pollOnce(); // the settling poll itself

    expect(await inboxLines()).toEqual([]); // still never backfills (0.4.1 contract, unchanged)
    const forChat = roster.merged.filter((entry) => entry.chatId === CHAT);
    expect(forChat.some((entry) => entry.members.some((m) => m.id === 'ancient-id'))).toBe(true);
  });

  // Violating-double coverage: a message whose sender Graph could not fully identify (an id with
  // no real display name — toChatMessage's own 'unknown'/'system' fallback, see messages.ts) must
  // never be merged into the roster as if 'unknown' were a real person's name. This is untrusted
  // data (Graph's own response shape) reaching the roster cache directly, with no `/members`
  // validation in between.
  it('VIOLATING DOUBLE: a message with no resolvable sender name (fromId present, from is the "unknown" fallback) is never harvested', async () => {
    const roster = rosterSink();
    const store = chatStore({});
    await pollerWithRoster(store, roster).pollOnce(); // settle
    store.add(
      CHAT,
      // Simulates toChatMessage's own fallback shape for a sender Graph reported an id for but no
      // displayName — see messages.ts's `from: sender?.displayName ?? (... : 'unknown')`.
      { ...message({ id: 'weird' }), from: 'unknown', fromId: 'has-id-no-name' },
    );

    await pollerWithRoster(store, roster).pollOnce();

    const harvestedIds = roster.merged.flatMap((entry) => entry.members.map((m) => m.id));
    expect(harvestedIds).not.toContain('has-id-no-name');
  });

  it('a system message (no fromId at all) is never harvested', async () => {
    const roster = rosterSink();
    const store = chatStore({});
    await pollerWithRoster(store, roster).pollOnce(); // settle
    store.add(CHAT, { ...message({ id: 'sys' }), from: 'system', fromId: undefined });

    await pollerWithRoster(store, roster).pollOnce();

    expect(roster.merged.filter((entry) => entry.chatId === CHAT)).toHaveLength(0);
  });

  it("the assistant's own posted messages ARE still harvested (isSelf only filters what gets DELIVERED to the inbox, not the roster) — self is a real chat member", async () => {
    const roster = rosterSink();
    const store = chatStore({});
    await pollerWithRoster(store, roster).pollOnce(); // settle
    store.add(CHAT, message({ id: 'self-msg', from: me.displayName, fromId: me.id }));

    await pollerWithRoster(store, roster).pollOnce();

    expect(await inboxLines()).toEqual([]); // never delivered as an inbox event (unchanged 0.4.1 rule)
    const harvestedIds = roster.merged.flatMap((entry) => entry.members.map((m) => m.id));
    expect(harvestedIds).toContain(me.id);
  });
});
