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
    const chats = new ReliableTeamsChats(inner, { sleepFn: async () => {}, nowFn: fixedNow });

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
    const chats = new ReliableTeamsChats(inner, { sleepFn: async () => {}, nowFn: fixedNow });

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
    const chats = new ReliableTeamsChats(inner, { sleepFn: async () => {}, nowFn: fixedNow });

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
    const chats = new ReliableTeamsChats(inner, { sleepFn: async () => {}, nowFn: fixedNow });

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
      attempts: 3,
      sleepFn: async () => {},
      nowFn: fixedNow,
    });

    await expect(chats.sendMessage('19:a@thread.v2', 'hello channel')).rejects.toThrow(
      /still throttled/,
    );
    expect(sendMessage).toHaveBeenCalledTimes(3);
  });

  it('a failed readback never masks the send outcome — it just means retry', async () => {
    const sendMessage = vi
      .fn()
      .mockRejectedValueOnce(new GraphError('hang up', 0))
      .mockResolvedValueOnce(message({ id: 'second-try' }));
    const inner = portWith({
      sendMessage,
      readMessages: async () => {
        throw new GraphError('reads throttled too', 429);
      },
    });
    const chats = new ReliableTeamsChats(inner, { sleepFn: async () => {}, nowFn: fixedNow });

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
    const chats = new ReliableTeamsChats(inner, { sleepFn: async () => {}, nowFn: fixedNow });

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
    const chats = new ReliableTeamsChats(inner, { sleepFn: async () => {}, nowFn: fixedNow });

    expect(await chats.sendMessage('19:a@thread.v2', 'two words spread out')).toBe(landed);
    expect(sendMessage).toHaveBeenCalledTimes(1);
  });

  it('reads and other operations delegate untouched', async () => {
    const inner = portWith({
      editMessage: vi.fn(async () => undefined),
      deleteMessage: vi.fn(async () => undefined),
    });
    const chats = new ReliableTeamsChats(inner, { sleepFn: async () => {}, nowFn: fixedNow });

    await chats.editMessage('19:a@thread.v2', 'm1', 'new');
    await chats.deleteMessage('19:a@thread.v2', 'm1');

    expect(inner.editMessage).toHaveBeenCalledWith('19:a@thread.v2', 'm1', 'new');
    expect(inner.deleteMessage).toHaveBeenCalledWith('19:a@thread.v2', 'm1');
  });
});
