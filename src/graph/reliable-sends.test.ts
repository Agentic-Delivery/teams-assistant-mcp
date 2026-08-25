import { describe, expect, it, vi } from 'vitest';
import { GraphError } from './graph-client.js';
import { ReliableTeamsChats } from './reliable-sends.js';
import type { TeamsChatsPort } from './teams-chats.js';
import type { ChatMessage, ReadResult } from '../messages.js';

function message(overrides: Partial<ChatMessage>): ChatMessage {
  return {
    id: 'msg-1',
    chatId: '19:a@thread.v2',
    createdDateTime: '2026-08-25T06:00:05Z',
    from: 'Assistant',
    text: 'hello channel',
    isDeleted: false,
    attachments: [],
    ...overrides,
  };
}

function portWith(overrides: Partial<TeamsChatsPort>): TeamsChatsPort {
  const reject = () => Promise.reject(new Error('not part of this test'));
  return {
    listChats: reject,
    readMessages: () => Promise.resolve({ messages: [] } as unknown as ReadResult),
    sendMessage: reject,
    sendImage: reject,
    sendFile: reject,
    replyToMessage: reject,
    editMessage: reject,
    deleteMessage: reject,
    setReaction: reject,
    getAttachment: reject,
    ...overrides,
  } as TeamsChatsPort;
}

const fixedNow = () => new Date('2026-08-25T06:00:00Z');

describe('reliable sends — readback before any retry', () => {
  it('passes a clean send straight through', async () => {
    const sent = message({ id: 'fresh' });
    const inner = portWith({ sendMessage: vi.fn(async () => sent) });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    expect(await chats.sendMessage('19:a@thread.v2', 'hello channel')).toBe(sent);
    expect(inner.sendMessage).toHaveBeenCalledTimes(1);
  });

  it('when the send "fails" but the message landed, the readback finds it and NOTHING is re-sent', async () => {
    // The exact 2026-08-24 shape: the POST succeeded, the response path lied.
    const landed = message({ text: 'Nytt i stage nå 🚀' });
    const sendMessage = vi.fn(async () => {
      throw new GraphError('socket hang up mid-response', 0);
    });
    const inner = portWith({
      sendMessage,
      readMessages: vi.fn(async () => ({ messages: [landed] }) as unknown as ReadResult),
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    const result = await chats.sendMessage('19:a@thread.v2', 'Nytt i stage nå 🚀');

    expect(result).toBe(landed);
    expect(sendMessage).toHaveBeenCalledTimes(1);
  });

  it('a message older than this attempt does not count as the landed copy', async () => {
    const stale = message({ text: 'hello channel', createdDateTime: '2026-08-25T05:00:00Z' });
    const sendMessage = vi
      .fn()
      .mockRejectedValueOnce(new GraphError('throttled', 429, 'TooManyRequests', 1))
      .mockResolvedValueOnce(message({ id: 'second-try' }));
    const inner = portWith({
      sendMessage,
      readMessages: async () => ({ messages: [stale] }) as unknown as ReadResult,
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    const result = await chats.sendMessage('19:a@thread.v2', 'hello channel');

    expect(result.id).toBe('second-try');
    expect(sendMessage).toHaveBeenCalledTimes(2);
  });

  it('a deleted copy does not count as landed either', async () => {
    const deleted = message({ text: 'hello channel', isDeleted: true });
    const sendMessage = vi
      .fn()
      .mockRejectedValueOnce(new GraphError('throttled', 429))
      .mockResolvedValueOnce(message({ id: 'second-try' }));
    const inner = portWith({
      sendMessage,
      readMessages: async () => ({ messages: [deleted] }) as unknown as ReadResult,
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    expect((await chats.sendMessage('19:a@thread.v2', 'hello channel')).id).toBe('second-try');
  });

  it('honours the Retry-After the throttle names before the second attempt', async () => {
    const waits: number[] = [];
    const sendMessage = vi
      .fn()
      .mockRejectedValueOnce(new GraphError('throttled', 429, 'TooManyRequests', 11))
      .mockResolvedValueOnce(message({ id: 'second-try' }));
    const inner = portWith({ sendMessage });
    const chats = new ReliableTeamsChats(inner, {
      selfDisplayName: 'Assistant',
      sleepFn: async (ms) => void waits.push(ms),
      nowFn: fixedNow,
    });

    await chats.sendMessage('19:a@thread.v2', 'hello channel');

    expect(waits).toEqual([11000]);
  });

  it('gives up honestly after the attempt budget, surfacing the last error', async () => {
    const sendMessage = vi.fn(async () => {
      throw new GraphError('still throttled', 429, 'TooManyRequests', 1);
    });
    const inner = portWith({ sendMessage });
    const chats = new ReliableTeamsChats(inner, {
      selfDisplayName: 'Assistant',
      attempts: 3,
      sleepFn: async () => {},
      nowFn: fixedNow,
    });

    await expect(chats.sendMessage('19:a@thread.v2', 'hello channel')).rejects.toThrow(
      /still throttled/,
    );
    expect(sendMessage).toHaveBeenCalledTimes(3);
  });

  it('a readback that fails for a non-throttle reason never masks the send outcome — it just means retry', async () => {
    // (A readback refused by a 429 is different: that is an UNKNOWN outcome, pinned in the
    // real-assembly suite below — reading the chat is impossible until the throttle clears.)
    const sendMessage = vi
      .fn()
      .mockRejectedValueOnce(new GraphError('hang up', 0))
      .mockResolvedValueOnce(message({ id: 'second-try' }));
    const inner = portWith({
      sendMessage,
      readMessages: async () => {
        throw new GraphError('read died on the same dead link', 0);
      },
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    expect((await chats.sendMessage('19:a@thread.v2', 'hello channel')).id).toBe('second-try');
  });

  it('replies get the same protection, matched by containment (a reply carries the quote too)', async () => {
    const landedReply = message({
      id: 'landed-reply',
      text: 'Quoted original text\nmy actual answer here',
    });
    const replyToMessage = vi.fn(async () => {
      throw new GraphError('hang up', 0);
    });
    const inner = portWith({
      replyToMessage,
      readMessages: async () => ({ messages: [landedReply] }) as unknown as ReadResult,
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    const result = await chats.replyToMessage('19:a@thread.v2', 'orig-1', 'my actual answer here');

    expect(result).toBe(landedReply);
    expect(replyToMessage).toHaveBeenCalledTimes(1);
  });

  it('whitespace differences from the HTML round-trip do not defeat the match', async () => {
    const landed = message({ text: 'two  words\n\n  spread   out' });
    const sendMessage = vi.fn(async () => {
      throw new GraphError('hang up', 0);
    });
    const inner = portWith({
      sendMessage,
      readMessages: async () => ({ messages: [landed] }) as unknown as ReadResult,
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    expect(await chats.sendMessage('19:a@thread.v2', 'two words spread out')).toBe(landed);
    expect(sendMessage).toHaveBeenCalledTimes(1);
  });

  it('reads and other operations delegate untouched', async () => {
    const inner = portWith({
      editMessage: vi.fn(async () => undefined),
      deleteMessage: vi.fn(async () => undefined),
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    await chats.editMessage('19:a@thread.v2', 'm1', 'new');
    await chats.deleteMessage('19:a@thread.v2', 'm1');

    expect(inner.editMessage).toHaveBeenCalledWith('19:a@thread.v2', 'm1', 'new');
    expect(inner.deleteMessage).toHaveBeenCalledWith('19:a@thread.v2', 'm1');
  });
});

describe('reliable sends — the match must be OUR copy, nothing else (review round 1 MAJORs)', () => {
  it("a colleague's message quoting our text is never claimed as the landed copy", async () => {
    const foreign = message({ id: 'johans', from: 'Johan', text: '"Deploy klar" — ja, jeg ser den' });
    const sendMessage = vi
      .fn()
      .mockRejectedValueOnce(new GraphError('hang up', 0))
      .mockResolvedValueOnce(message({ id: 'second-try', text: 'Deploy klar' }));
    const inner = portWith({
      sendMessage,
      readMessages: async () => ({ messages: [foreign] }) as unknown as ReadResult,
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    const result = await chats.sendMessage('19:a@thread.v2', 'Deploy klar');

    expect(result.id).toBe('second-try');
    expect(sendMessage).toHaveBeenCalledTimes(2);
  });

  it('a short reply is not satisfied by the quoted ORIGINAL containing the same words', async () => {
    // Replying "ja" to "Ska jag deploya nu? ja/nej" — the original contains "ja" but is not
    // our reply. The tail match plus the sender pin both have to hold.
    const original = message({ id: 'orig', from: 'Johan', text: 'Ska jag deploya nu? ja/nej' });
    const replyToMessage = vi
      .fn()
      .mockRejectedValueOnce(new GraphError('hang up', 0))
      .mockResolvedValueOnce(message({ id: 'real-reply', text: 'Ska jag deploya nu? ja/nej\nja' }));
    const inner = portWith({
      replyToMessage,
      readMessages: async () => ({ messages: [original] }) as unknown as ReadResult,
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    const result = await chats.replyToMessage('19:a@thread.v2', 'orig', 'ja');

    expect(result.id).toBe('real-reply');
    expect(replyToMessage).toHaveBeenCalledTimes(2);
  });

  it('a whole-message send is matched by equality, not containment', async () => {
    const superset = message({ id: 'older-superset', text: 'status ok — men les videre' });
    const sendMessage = vi
      .fn()
      .mockRejectedValueOnce(new GraphError('hang up', 0))
      .mockResolvedValueOnce(message({ id: 'second-try', text: 'status ok' }));
    const inner = portWith({
      sendMessage,
      readMessages: async () => ({ messages: [superset] }) as unknown as ReadResult,
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    expect((await chats.sendMessage('19:a@thread.v2', 'status ok')).id).toBe('second-try');
  });

  it('when two own copies stand in the window, the NEWEST one wins', async () => {
    const older = message({ id: 'older', text: 'status ok', createdDateTime: '2026-08-25T05:59:30Z' });
    const newer = message({ id: 'newer', text: 'status ok', createdDateTime: '2026-08-25T06:00:10Z' });
    const sendMessage = vi.fn(async () => {
      throw new GraphError('hang up', 0);
    });
    const inner = portWith({
      sendMessage,
      readMessages: async () => ({ messages: [older, newer] }) as unknown as ReadResult,
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    expect((await chats.sendMessage('19:a@thread.v2', 'status ok')).id).toBe('newer');
  });

  it('an offset-form timestamp still matches (no string comparison of ISO forms)', async () => {
    const landed = message({ text: 'hello channel', createdDateTime: '2026-08-25T06:00:05+00:00' });
    const sendMessage = vi.fn(async () => {
      throw new GraphError('hang up', 0);
    });
    const inner = portWith({
      sendMessage,
      readMessages: async () => ({ messages: [landed] }) as unknown as ReadResult,
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    expect(await chats.sendMessage('19:a@thread.v2', 'hello channel')).toBe(landed);
    expect(sendMessage).toHaveBeenCalledTimes(1);
  });

  it('an absurd Retry-After is capped instead of parking the caller', async () => {
    const waits: number[] = [];
    const sendMessage = vi
      .fn()
      .mockRejectedValueOnce(new GraphError('throttled', 429, 'TooManyRequests', 300))
      .mockResolvedValueOnce(message({ id: 'second-try' }));
    const inner = portWith({ sendMessage });
    const chats = new ReliableTeamsChats(inner, {
      selfDisplayName: 'Assistant',
      sleepFn: async (ms) => void waits.push(ms),
      nowFn: fixedNow,
    });

    await chats.sendMessage('19:a@thread.v2', 'hello channel');

    expect(waits).toEqual([60000]);
  });

  it('a retry re-checks the chat right before re-sending, catching a late-landing copy', async () => {
    // First readback dies with the same connection the send died on; by the retry, both work.
    const landed = message({ text: 'hello channel' });
    const sendMessage = vi.fn(async () => {
      throw new GraphError('hang up', 0);
    });
    const readMessages = vi
      .fn()
      .mockRejectedValueOnce(new GraphError('same dead link', 0))
      .mockResolvedValueOnce({ messages: [landed] } as unknown as ReadResult);
    const inner = portWith({ sendMessage, readMessages });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    expect(await chats.sendMessage('19:a@thread.v2', 'hello channel')).toBe(landed);
    expect(sendMessage).toHaveBeenCalledTimes(1);
  });
});

describe('reliable sends — under a 429 the chat is NOT read back before the wait (2026-08-25)', () => {
  it('a throttled send waits the named window first, reads back once, then retries once — bounded', async () => {
    const calls: string[] = [];
    const sendMessage = vi
      .fn()
      .mockImplementationOnce(async () => { calls.push('send'); throw new GraphError('throttled', 429, 'TooManyRequests', 20); })
      .mockImplementationOnce(async () => { calls.push('send'); return message({ id: 'second-try' }); });
    const readMessages = vi.fn(async () => { calls.push('read'); return { messages: [] } as unknown as ReadResult; });
    const inner = portWith({ sendMessage, readMessages });
    const waits: number[] = [];
    const chats = new ReliableTeamsChats(inner, {
      selfDisplayName: 'Assistant',
      sleepFn: async (ms) => void waits.push(ms),
      nowFn: fixedNow,
    });

    expect((await chats.sendMessage('19:a@thread.v2', 'hello channel')).id).toBe('second-try');
    expect(calls).toEqual(['send', 'read', 'send']); // no readback BEFORE the wait
    expect(waits).toEqual([20000]);
  });

  it('two 429s in a row: give up after the second, three requests total, never a storm', async () => {
    const sendMessage = vi.fn(async () => { throw new GraphError('throttled', 429, 'TooManyRequests', 5); });
    const readMessages = vi.fn(async () => ({ messages: [] }) as unknown as ReadResult);
    const inner = portWith({ sendMessage, readMessages });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    await expect(chats.sendMessage('19:a@thread.v2', 'hello channel')).rejects.toMatchObject({ status: 429 });
    expect(sendMessage).toHaveBeenCalledTimes(2);
    expect(readMessages).toHaveBeenCalledTimes(1);
  });
});


import { GraphClient } from './graph-client.js';
import { GraphTeamsChats } from './teams-chats.js';
import type { TokenProvider } from '../auth/token-provider.js';

describe('the real assembly — GraphClient + GraphTeamsChats + ReliableTeamsChats over a fake fetch (review round 1)', () => {
  const stubToken: TokenProvider = { kind: 'stub', getAccessToken: async () => 't' };
  const jsonResponse = (body: unknown, status = 200) =>
    new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
  const created = (id: string, text: string) => ({
    id, createdDateTime: '2026-08-25T06:00:05Z', from: { user: { displayName: 'Assistant', id: 'me' } },
    body: { contentType: 'html', content: text }, attachments: [],
  });

  function stack(fetchFn: ReturnType<typeof vi.fn>, sleeps: number[]) {
    let now = Date.parse('2026-08-25T06:00:00Z');
    const client = new GraphClient({
      tokenProvider: stubToken, fetchFn: fetchFn as never,
      sleepFn: async (ms) => { sleeps.push(ms); now += ms; },
      nowFn: () => now,
    });
    const chats = new ReliableTeamsChats(new GraphTeamsChats(client), {
      selfDisplayName: 'Assistant',
      sleepFn: async (ms) => { sleeps.push(ms); now += ms; },
      nowFn: () => new Date(now),
    });
    return { client, chats };
  }

  it('a 429 without Retry-After: the decorator waits the gate default, readback and retry then reach Graph and succeed', async () => {
    const sleeps: number[] = [];
    const fetchFn = vi.fn()
      .mockResolvedValueOnce(jsonResponse({ error: { code: 'TooManyRequests', message: 'throttled' } }, 429)) // POST #1
      .mockResolvedValueOnce(jsonResponse({ value: [] }))                                                    // readback
      .mockResolvedValueOnce(jsonResponse(created('m2', 'hello')));                                          // POST #2
    const { chats } = stack(fetchFn, sleeps);

    const sent = await chats.sendMessage('19:a@thread.v2', 'hello');

    expect(sent.id).toBe('m2');
    expect(fetchFn).toHaveBeenCalledTimes(3);
    expect(sleeps[0]).toBeGreaterThanOrEqual(30_000); // reconciled with the client's default gate
  });

  it('a send whose response path dies while a concurrent 429 has closed the gate: UNKNOWN outcome, naming the hang-up — never "not sent"', async () => {
    // The POST reaches Graph and dies on the way back (the write may be standing); every GET —
    // the readbacks — meets a 429, which closes the gate. Nothing can be known until it reopens.
    const fetchFn = vi.fn(async (_url: string, init: RequestInit) => {
      if (init.method === 'POST') { throw new TypeError('socket hang up'); }
      return jsonResponse({ error: { code: 'TooManyRequests', message: 'poller throttled' } }, 429);
    });
    let now = Date.parse('2026-08-25T06:00:00Z');
    const client = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never, sleepFn: async (ms) => { now += ms; }, nowFn: () => now, readRetries: 0 });
    const chats = new ReliableTeamsChats(new GraphTeamsChats(client), { selfDisplayName: 'Assistant', sleepFn: async () => { /* no clock advance: the gate stays closed */ }, nowFn: () => new Date(now) });

    const error = (await chats.sendMessage('19:a@thread.v2', 'hello').catch((c: unknown) => c)) as GraphError;

    expect(error).toBeInstanceOf(GraphError);
    expect(error.code).toBe('UnknownOutcome');
    expect(error.message).toMatch(/socket hang up/);
    expect(error.message).not.toMatch(/not sent/);
  });
});
