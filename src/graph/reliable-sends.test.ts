import { describe, expect, it, vi } from 'vitest';
import { GraphClient, GraphError } from './graph-client.js';
import { ReliableTeamsChats } from './reliable-sends.js';
import { GraphTeamsChats, type MembersCachePort, type MentionTarget, type TeamsChatsPort } from './teams-chats.js';
import type { TokenProvider } from '../auth/token-provider.js';
import { toChatMessage, type ChatMessage, type ReadResult } from '../messages.js';

// None of this file's "real assembly" describe blocks touch mention resolution; membersCache is
// a required GraphTeamsChats dependency (0.4.1 review round 1), so a trivial no-op double stands
// in everywhere below, module-scoped so every describe block can see it.
const noMembersCache: MembersCachePort = { get: () => undefined, set: () => {}, getStale: () => undefined };

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
    pinMessage: reject,
    unpinMessage: reject,
    listPinnedMessages: reject,
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

    expect(inner.editMessage).toHaveBeenCalledWith('19:a@thread.v2', 'm1', 'new', []);
    expect(inner.deleteMessage).toHaveBeenCalledWith('19:a@thread.v2', 'm1');
  });
});

describe('reliable sends — html format: readback dedup compares TEXT, not raw markup', () => {
  it('passes a clean html send straight through, untouched', async () => {
    const sent = message({ id: 'fresh-html' });
    const inner = portWith({ sendHtmlMessage: vi.fn(async () => sent) });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    expect(await chats.sendHtmlMessage('19:a@thread.v2', '<b>Deploy</b> done')).toBe(sent);
    expect(inner.sendHtmlMessage).toHaveBeenCalledWith('19:a@thread.v2', '<b>Deploy</b> done', []);
    expect(inner.sendHtmlMessage).toHaveBeenCalledTimes(1);
  });

  it('an html send that "fails" but landed is found by its TEXT rendering — a raw-markup compare would never match and would duplicate', async () => {
    // The chat readback always comes back as text (toChatMessage runs every html body through
    // htmlToText — see messages.ts): a landed "<b>Deploy</b> done" reads back as "Deploy done".
    // If the guard compared the raw markup instead, this never matches and the retry re-posts.
    const landed = message({ text: 'Deploy done' });
    const sendHtmlMessage = vi.fn(async () => {
      throw new GraphError('socket hang up mid-response', 0);
    });
    const inner = portWith({
      sendHtmlMessage,
      readMessages: vi.fn(async () => ({ messages: [landed] }) as unknown as ReadResult),
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    const result = await chats.sendHtmlMessage('19:a@thread.v2', '<b>Deploy</b> done');

    expect(result).toBe(landed);
    expect(sendHtmlMessage).toHaveBeenCalledTimes(1); // found the standing copy — never duplicated
  });

  it('an html send whose landed copy is not found (genuinely unsent) still retries and posts the same raw html again', async () => {
    const sendHtmlMessage = vi
      .fn()
      .mockRejectedValueOnce(new GraphError('hang up', 0))
      .mockResolvedValueOnce(message({ id: 'second-try' }));
    const inner = portWith({
      sendHtmlMessage,
      readMessages: async () => ({ messages: [] }) as unknown as ReadResult,
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    const result = await chats.sendHtmlMessage('19:a@thread.v2', '<b>Deploy</b> done');

    expect(result.id).toBe('second-try');
    expect(sendHtmlMessage).toHaveBeenNthCalledWith(2, '19:a@thread.v2', '<b>Deploy</b> done', []);
  });

  it('editHtmlMessage delegates untouched — a PATCH targets an existing id, nothing to guard', async () => {
    const inner = portWith({ editHtmlMessage: vi.fn(async () => undefined) });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    await chats.editHtmlMessage('19:a@thread.v2', 'm1', '<b>corrected</b>');

    expect(inner.editHtmlMessage).toHaveBeenCalledWith('19:a@thread.v2', 'm1', '<b>corrected</b>', []);
  });

  it('an html body that reduces to no text at all (image/hr-only) gets no guard: one attempt, the send failure itself, no readback', async () => {
    // The hazard this closes: an EMPTY match key equality-matches ANY earlier own message whose
    // text also reduces to empty (an unrelated image sent minutes ago, say) — findLandedCopy
    // would report THAT as "this attempt's landed copy" and swallow a genuine failure as success.
    const earlierEmpty = message({
      id: 'earlier-empty',
      text: '',
      createdDateTime: '2026-08-25T05:59:50Z', // inside the attempt window
    });
    const sendHtmlMessage = vi.fn(async () => {
      throw new GraphError('hang up', 0);
    });
    const readMessages = vi.fn(async () => ({ messages: [earlierEmpty] }) as unknown as ReadResult);
    const inner = portWith({ sendHtmlMessage, readMessages });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    await expect(
      chats.sendHtmlMessage('19:a@thread.v2', '<img src="https://example.test/x.png">'),
    ).rejects.toThrow(/hang up/);

    expect(readMessages).not.toHaveBeenCalled(); // no blind readback against an empty match key
    expect(sendHtmlMessage).toHaveBeenCalledTimes(1); // one attempt, never a retry
  });

  it('a clean send of empty-reducing html still succeeds — one attempt is enough when nothing fails', async () => {
    const sent = message({ id: 'hr-1', text: '' });
    const sendHtmlMessage = vi.fn(async () => sent);
    const inner = portWith({ sendHtmlMessage });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    expect(await chats.sendHtmlMessage('19:a@thread.v2', '<hr>')).toBe(sent);
    expect(sendHtmlMessage).toHaveBeenCalledTimes(1);
  });
});

describe('reliable sends — html match key vs a REAL captured Teams readback (fixture, 2026-08-25)', () => {
  // Capture provenance (re-runnable): chat 19:8af48977045e48c9b9ed5049ba8f94ad@thread.v2
  // ("MCP dev test", allowlisted canPost), message id 1787666759027, captured 2026-08-25.
  // Sent this MINIFIED table (no whitespace between adjacent tags) via a real sendHtmlMessage
  // call, then GET the message back off Graph and read RAW body.content — bypassing
  // toChatMessage/htmlToText entirely, so this is exactly what Teams stored, not what our own
  // code thinks it stored.
  const SENT_HTML =
    '<table border="1"><tr><th>Item</th><th>State</th></tr><tr><td>build</td><td>ok</td></tr></table>';
  const CAPTURED_RAW_READBACK =
    '<table border="1">\n<tbody>\n<tr>\n<th>Item</th>\n<th>State</th>\n</tr>\n<tr>\n' +
    '<td>build</td>\n<td>ok</td>\n</tr>\n</tbody>\n</table>';

  it('toChatMessage reduces the captured readback to "Item State\\nbuild ok" (pins the capture itself)', () => {
    const landed = toChatMessage(
      {
        id: 'captured-1',
        createdDateTime: '2026-08-25T06:00:05Z',
        from: { user: { id: 'me', displayName: 'Assistant' } },
        body: { contentType: 'html', content: CAPTURED_RAW_READBACK },
      },
      '19:a@thread.v2',
    );

    expect(landed.text).toBe('Item State\nbuild ok');
  });

  it('the sent (minified) html matches its own REAL captured readback — table cell boundaries included', async () => {
    // Without htmlMatchText's tag-boundary fix this reduces to "ItemState\nbuildok" locally,
    // which never equals "Item State\nbuild ok" — this is the exact case that used to duplicate.
    const landed = toChatMessage(
      {
        id: 'captured-1',
        createdDateTime: '2026-08-25T06:00:05Z',
        from: { user: { id: 'me', displayName: 'Assistant' } },
        body: { contentType: 'html', content: CAPTURED_RAW_READBACK },
      },
      '19:a@thread.v2',
    );
    const sendHtmlMessage = vi.fn(async () => {
      throw new GraphError('socket hang up mid-response', 0);
    });
    const inner = portWith({
      sendHtmlMessage,
      readMessages: vi.fn(async () => ({ messages: [landed] }) as unknown as ReadResult),
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    const result = await chats.sendHtmlMessage('19:a@thread.v2', SENT_HTML);

    expect(result).toBe(landed);
    expect(sendHtmlMessage).toHaveBeenCalledTimes(1); // found the REAL captured landed copy — no duplicate
  });
});

describe('reliable sends — html match key vs a REAL captured Teams readback, INLINE adjacency (fixture, 2026-08-25)', () => {
  // Capture provenance (re-runnable): chat 19:8af48977045e48c9b9ed5049ba8f94ad@thread.v2
  // ("MCP dev test", allowlisted canPost), message id 1787667664136, captured 2026-08-25.
  // Sent `<p><b>a</b><i>b</i></p>` — zero whitespace at three bare boundaries (p→b, b→i, i→p) —
  // via a real sendHtmlMessage call, then GET the message back and read RAW body.content.
  //
  // This is the capture that DISPROVES the first fix attempt ("insert a space at every bare
  // tag-to-tag boundary"): the raw readback came back COMPLETELY UNCHANGED, byte-identical to
  // what was sent — Teams did not touch even the p→b boundary, despite p being a block tag. Only
  // the table capture above showed rewriting. See REWRITTEN_TAG_BOUNDARY's doc comment in
  // reliable-sends.ts for what this and the table capture together prove.
  const SENT_HTML = '<p><b>a</b><i>b</i></p>';
  const CAPTURED_RAW_READBACK = '<p><b>a</b><i>b</i></p>'; // verbatim — Teams did not rewrite this

  it('toChatMessage reduces the captured readback to "ab" (pins the capture itself)', () => {
    const landed = toChatMessage(
      {
        id: 'captured-2',
        createdDateTime: '2026-08-25T06:00:05Z',
        from: { user: { id: 'me', displayName: 'Assistant' } },
        body: { contentType: 'html', content: CAPTURED_RAW_READBACK },
      },
      '19:a@thread.v2',
    );

    expect(landed.text).toBe('ab');
  });

  it('the sent html matches its own REAL captured readback — inline tags get NO inserted space', async () => {
    // The dangerous direction: if htmlMatchText over-eagerly inserted a space at the b→i
    // boundary (or the p→b one), the computed key would be "a b" while the real landed copy's
    // text is "ab" — a permanent false mismatch, retrying a genuinely landed html send into a
    // real duplicate. This is the case a blanket "every boundary" rule gets wrong.
    const landed = toChatMessage(
      {
        id: 'captured-2',
        createdDateTime: '2026-08-25T06:00:05Z',
        from: { user: { id: 'me', displayName: 'Assistant' } },
        body: { contentType: 'html', content: CAPTURED_RAW_READBACK },
      },
      '19:a@thread.v2',
    );
    const sendHtmlMessage = vi.fn(async () => {
      throw new GraphError('socket hang up mid-response', 0);
    });
    const inner = portWith({
      sendHtmlMessage,
      readMessages: vi.fn(async () => ({ messages: [landed] }) as unknown as ReadResult),
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    const result = await chats.sendHtmlMessage('19:a@thread.v2', SENT_HTML);

    expect(result).toBe(landed);
    expect(sendHtmlMessage).toHaveBeenCalledTimes(1); // found the REAL captured landed copy — no duplicate
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

  it('an absurd Retry-After on a refused send: fail NOW with the 429, no sleep, no retry', async () => {
    const waits: number[] = [];
    const sendMessage = vi.fn(async () => { throw new GraphError('throttled', 429, 'TooManyRequests', 300); });
    const inner = portWith({ sendMessage });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async (ms) => void waits.push(ms), nowFn: fixedNow });

    await expect(chats.sendMessage('19:a@thread.v2', 'hello channel')).rejects.toMatchObject({ status: 429, retryAfterSeconds: 300 });
    expect(waits).toEqual([]);
    expect(sendMessage).toHaveBeenCalledTimes(1);
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
    const chats = new ReliableTeamsChats(new GraphTeamsChats(client, { membersCache: noMembersCache }), {
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
    const chats = new ReliableTeamsChats(new GraphTeamsChats(client, { membersCache: noMembersCache }), { selfDisplayName: 'Assistant', sleepFn: async () => { /* no clock advance: the gate stays closed */ }, nowFn: () => new Date(now) });

    const error = (await chats.sendMessage('19:a@thread.v2', 'hello').catch((c: unknown) => c)) as GraphError;

    expect(error).toBeInstanceOf(GraphError);
    expect(error.code).toBe('UnknownOutcome');
    expect(error.message).toMatch(/socket hang up/);
    expect(error.message).not.toMatch(/not sent/);
  });

  it('sendHtmlMessage posts the raw html verbatim — contentType html, content untouched by textToHtml', async () => {
    const sleeps: number[] = [];
    const fetchFn = vi.fn().mockResolvedValueOnce(jsonResponse(created('m1', 'irrelevant')));
    const { chats } = stack(fetchFn, sleeps);
    const html = '<b>Deploy</b> &amp; <span style="color:#c00">done</span>';

    await chats.sendHtmlMessage('19:a@thread.v2', html);

    const [, init] = fetchFn.mock.calls[0] as unknown as [string, RequestInit];
    const posted = JSON.parse(init.body as string) as { body: { contentType: string; content: string } };
    expect(posted.body).toEqual({ contentType: 'html', content: html }); // no escaping, no textToHtml
  });

  it('editHtmlMessage PATCHes the raw html verbatim', async () => {
    const sleeps: number[] = [];
    const fetchFn = vi.fn().mockResolvedValueOnce(new Response(null, { status: 204 }));
    const { chats } = stack(fetchFn, sleeps);
    const html = '<table border="1"><tr><td>ok</td></tr></table>';

    await chats.editHtmlMessage('19:a@thread.v2', 'm1', html);

    const [, init] = fetchFn.mock.calls[0] as unknown as [string, RequestInit];
    expect(init.method).toBe('PATCH');
    const patched = JSON.parse(init.body as string) as { body: { contentType: string; content: string } };
    expect(patched.body).toEqual({ contentType: 'html', content: html });
  });
});

describe('the real assembly — round 2', () => {
  const stubToken: TokenProvider = { kind: 'stub', getAccessToken: async () => 't' };
  const throttled = (retryAfter?: string) =>
    new Response(JSON.stringify({ error: { code: 'ApplicationThrottled', message: 'Rate limit is exceeded.' } }), {
      status: 429, headers: { 'content-type': 'application/json', ...(retryAfter ? { 'retry-after': retryAfter } : {}) },
    });
  const jsonResponse = (body: unknown, status = 200) =>
    new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });

  function stack(fetchFn: (url: string, init: RequestInit) => Promise<Response>, opts: { advance: boolean } = { advance: true }) {
    let now = Date.parse('2026-08-25T06:00:00Z');
    const sleeps: number[] = [];
    const sleepFn = async (ms: number) => { sleeps.push(ms); if (opts.advance) now += ms; };
    const client = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never, sleepFn, nowFn: () => now, readRetries: 0 });
    const chats = new ReliableTeamsChats(new GraphTeamsChats(client, { membersCache: noMembersCache }), { selfDisplayName: 'Assistant', sleepFn, nowFn: () => new Date(now) });
    return { chats, sleeps };
  }

  it('a send Graph REFUSED with a long Retry-After: reported as the 429 with that Retry-After, nothing slept, never "unknown"', async () => {
    const posts: number[] = [];
    const { chats, sleeps } = stack(async (_u, init) => { if (init.method === 'POST') { posts.push(1); return throttled('300'); } return jsonResponse({ value: [] }); });

    const error = (await chats.sendMessage('19:a@thread.v2', 'hello').catch((c: unknown) => c)) as GraphError;

    expect(error.status).toBe(429);
    expect(error.code).toBe('ApplicationThrottled');
    expect(error.retryAfterSeconds).toBe(300);
    expect(error.message).not.toMatch(/UNKNOWN/);
    expect(posts).toHaveLength(1);
    expect(sleeps).toEqual([]); // no minute of theatre before the honest answer
  });

  it('a send refused with a SHORT Retry-After, then the gate still closed at retry time: still the 429, still not "unknown"', async () => {
    const { chats } = stack(async (_u, init) => (init.method === 'POST' ? throttled('20') : throttled('20')), { advance: false });

    const error = (await chats.sendMessage('19:a@thread.v2', 'hello').catch((c: unknown) => c)) as GraphError;

    expect(error.status).toBe(429);
    expect(error.code).toBe('ApplicationThrottled'); // Graph's refusal, not our gate's
    expect(error.retryAfterSeconds).toBe(20);
  });

  it('first attempt: response path dies and the readback is blocked at once → UnknownOutcome immediately, no sleep, no second send', async () => {
    let posts = 0;
    const { chats, sleeps } = stack(async (_u, init) => { if (init.method === 'POST') { posts += 1; throw new TypeError('socket hang up'); } return throttled('1'); }, { advance: false });

    const error = (await chats.sendMessage('19:a@thread.v2', 'hello').catch((c: unknown) => c)) as GraphError;

    expect(error.code).toBe('UnknownOutcome');
    expect(posts).toBe(1);
    expect(sleeps).toEqual([]);
  });

  it('response path dies, then the readback meets a REAL Graph 429 (not our gate): UnknownOutcome naming the hang-up', async () => {
    const { chats } = stack(async (_u, init) => { if (init.method === 'POST') { throw new TypeError('socket hang up'); } return throttled('1'); });

    const error = (await chats.sendMessage('19:a@thread.v2', 'hello').catch((c: unknown) => c)) as GraphError;

    expect(error.code).toBe('UnknownOutcome');
    expect(error.message).toMatch(/socket hang up/);
  });
});

describe('the real assembly — MessageFetchThrottled through the decorator', () => {
  const stubToken: TokenProvider = { kind: 'stub', getAccessToken: async () => 't' };
  const jsonResponse = (body: unknown, status = 200) => new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });

  it('a reply whose original cannot be fetched fails at once with MessageFetchThrottled — no send, no readback, no sleep', async () => {
    const posts: number[] = [];
    const sleeps: number[] = [];
    const fetchFn = vi.fn(async (url: string, init: RequestInit) => {
      if (init.method === 'POST') { posts.push(1); return jsonResponse({ id: 'never' }, 201); }
      if (url.endsWith('/messages/old-id')) return new Response(JSON.stringify({ error: { code: 'TooManyRequests', message: 't' } }), { status: 429, headers: { 'retry-after': '60', 'content-type': 'application/json' } });
      return jsonResponse({ value: [] });
    });
    let now = 0;
    const client = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never, sleepFn: async (ms) => { sleeps.push(ms); now += ms; }, nowFn: () => now });
    const chats = new ReliableTeamsChats(new GraphTeamsChats(client, { membersCache: noMembersCache }), { selfDisplayName: 'Assistant', sleepFn: async (ms) => { sleeps.push(ms); now += ms; }, nowFn: () => new Date(now) });

    const error = (await chats.replyToMessage('19:a@thread.v2', 'old-id', 'hi').catch((c: unknown) => c)) as GraphError;

    expect(error.code).toBe('MessageFetchThrottled');
    expect(error.message).toMatch(/nothing was posted/);
    expect(posts).toHaveLength(0);
    expect(sleeps).toEqual([]);
  });
});

describe('reliable sends — @mentions rewrite the match key, not just the sent text', () => {
  const MENTION: MentionTarget = { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' };

  it('sendMessage with a mention matches a landed copy by the RESOLVED displayName, not the caller\'s search string', async () => {
    // "Shiv" was typed, but the <at> tag (and so the readback) carries "Garg, Shivankit" — a
    // match key built from the raw input text ("Shiv please review") would never equal a landed
    // copy's readback ("Garg, Shivankit please review") and would retry a genuinely landed send
    // into a real duplicate.
    const landed = message({ text: 'Garg, Shivankit please review' });
    const sendMessage = vi.fn(async () => {
      throw new GraphError('socket hang up mid-response', 0);
    });
    const inner = portWith({
      sendMessage,
      readMessages: vi.fn(async () => ({ messages: [landed] }) as unknown as ReadResult),
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    const result = await chats.sendMessage('19:a@thread.v2', 'Shiv please review', [MENTION]);

    expect(result).toBe(landed);
    expect(sendMessage).toHaveBeenCalledTimes(1); // found the landed copy — no duplicate
  });

  it('sendMessage without mentions still matches on the raw text unchanged (no regression)', async () => {
    const landed = message({ text: 'hello channel' });
    const sendMessage = vi.fn(async () => {
      throw new GraphError('socket hang up mid-response', 0);
    });
    const inner = portWith({
      sendMessage,
      readMessages: vi.fn(async () => ({ messages: [landed] }) as unknown as ReadResult),
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    expect(await chats.sendMessage('19:a@thread.v2', 'hello channel')).toBe(landed);
  });

  it('replyToMessage with a mention matches on the resolved displayName in the reply tail', async () => {
    const landed = message({ text: 'quoted original text\nGarg, Shivankit can you confirm?' });
    const replyToMessage = vi.fn(async () => {
      throw new GraphError('socket hang up mid-response', 0);
    });
    const inner = portWith({
      replyToMessage,
      readMessages: vi.fn(async () => ({ messages: [landed] }) as unknown as ReadResult),
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    const result = await chats.replyToMessage('19:a@thread.v2', 'orig-1', 'Shiv can you confirm?', [MENTION]);

    expect(result).toBe(landed);
    expect(replyToMessage).toHaveBeenCalledTimes(1);
  });

  it('sendHtmlMessage with a mention placeholder matches on the resolved displayName', async () => {
    const landed = message({ text: 'Garg, Shivankit see the table below' });
    const sendHtmlMessage = vi.fn(async () => {
      throw new GraphError('socket hang up mid-response', 0);
    });
    const inner = portWith({
      sendHtmlMessage,
      readMessages: vi.fn(async () => ({ messages: [landed] }) as unknown as ReadResult),
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    const result = await chats.sendHtmlMessage('19:a@thread.v2', '<p>@{Shiv} see the table below</p>', [MENTION]);

    expect(result).toBe(landed);
    expect(sendHtmlMessage).toHaveBeenCalledTimes(1);
  });
});

describe('reliable sends — @mention match key vs a REAL captured Teams readback (fixture, 2026-08-26)', () => {
  // Capture provenance (re-runnable): chat 19:8af48977045e48c9b9ed5049ba8f94ad@thread.v2
  // ("MCP dev test", allowlisted canPost), message id 1787722849960, captured 2026-08-26. Sent
  // via a real sendMessage call mentioning "Johan" (resolved to "Spännare, Johan"), then GET the
  // message back off Graph and read RAW body.content — bypassing toChatMessage/htmlToText
  // entirely, so this is exactly what Teams stored, not what our own code thinks it stored.
  //
  // Unlike the table case in the format:'html' fixtures above, Teams did NOT rewrite anything
  // here: the raw readback is byte-identical to what renderTextWithMentions produced — an <at>
  // tag behaves like the other inline tags (<b>, <i>) proven untouched by the earlier capture,
  // not like the table tags that get a <tbody> and boundary newlines inserted. So <at> needs no
  // entry in REWRITTEN_TAG_BOUNDARY, and plain htmlToText reduction is exactly right for it.
  const MENTION: MentionTarget = {
    name: 'Johan',
    id: '47588e53-8923-487c-a5db-6f5157a9d97b',
    displayName: 'Spännare, Johan',
  };
  const SENT_TEXT =
    'Mention capture test (0.4.0 feature dev) — Johan please ignore, verifying @mention notify + readback shape.';
  const CAPTURED_RAW_READBACK =
    '<p>Mention capture test (0.4.0 feature dev) — <at id="0">Spännare, Johan</at> please ignore, ' +
    'verifying @mention notify + readback shape.</p>';

  it('toChatMessage reduces the captured readback to the resolved displayName in place of the mention (pins the capture itself)', () => {
    const landed = toChatMessage(
      {
        id: 'captured-mention-1',
        createdDateTime: '2026-08-26T06:00:05Z',
        from: { user: { id: 'me', displayName: 'Assistant' } },
        body: { contentType: 'html', content: CAPTURED_RAW_READBACK },
      },
      '19:a@thread.v2',
    );

    expect(landed.text).toBe(
      'Mention capture test (0.4.0 feature dev) — Spännare, Johan please ignore, verifying @mention notify + readback shape.',
    );
  });

  it('a mention-bearing send\'s match key equals its own REAL captured readback — no duplicate on retry', async () => {
    const landed = toChatMessage(
      {
        id: 'captured-mention-1',
        createdDateTime: '2026-08-26T06:00:05Z',
        from: { user: { id: 'me', displayName: 'Assistant' } },
        body: { contentType: 'html', content: CAPTURED_RAW_READBACK },
      },
      '19:a@thread.v2',
    );
    const sendMessage = vi.fn(async () => {
      throw new GraphError('socket hang up mid-response', 0);
    });
    const inner = portWith({
      sendMessage,
      readMessages: vi.fn(async () => ({ messages: [landed] }) as unknown as ReadResult),
    });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {}, nowFn: fixedNow });

    const result = await chats.sendMessage('19:a@thread.v2', SENT_TEXT, [MENTION]);

    expect(result).toBe(landed);
    expect(sendMessage).toHaveBeenCalledTimes(1); // found the REAL captured landed copy — no duplicate
  });
});

describe('reliable sends — resolveMentions and pin/unpin/list are pure passthroughs', () => {
  it('resolveMentions delegates untouched — a read against the member list, nothing to guard', async () => {
    const resolveMentions = vi.fn(async () => [{ name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' }]);
    const inner = portWith({ resolveMentions });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {} });

    const result = await chats.resolveMentions('19:a@thread.v2', ['Shiv']);

    expect(resolveMentions).toHaveBeenCalledWith('19:a@thread.v2', ['Shiv']);
    expect(result).toEqual([{ name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' }]);
  });

  it('pinMessage, unpinMessage and listPinnedMessages all delegate untouched', async () => {
    const pinMessage = vi.fn(async () => [{ id: 'pin-1', messageId: 'm1', preview: 'x' }]);
    const unpinMessage = vi.fn(async () => undefined);
    const listPinnedMessages = vi.fn(async () => [{ id: 'pin-1', messageId: 'm1', preview: 'x' }]);
    const inner = portWith({ pinMessage, unpinMessage, listPinnedMessages });
    const chats = new ReliableTeamsChats(inner, { selfDisplayName: 'Assistant', sleepFn: async () => {} });

    await chats.pinMessage('19:a@thread.v2', 'm1');
    await chats.unpinMessage('19:a@thread.v2', 'm1');
    await chats.listPinnedMessages('19:a@thread.v2');

    expect(pinMessage).toHaveBeenCalledWith('19:a@thread.v2', 'm1');
    expect(unpinMessage).toHaveBeenCalledWith('19:a@thread.v2', 'm1');
    expect(listPinnedMessages).toHaveBeenCalledWith('19:a@thread.v2');
  });
});
