import { describe, expect, it, vi } from 'vitest';
import { GraphClient, GraphError } from './graph-client.js';
import { GraphTeamsChats } from './teams-chats.js';
import type { TokenProvider } from '../auth/token-provider.js';

const stubToken: TokenProvider = { kind: 'stub', getAccessToken: async () => 'the-token' };

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

describe('graph client', () => {
  it('puts the provider token on every request', async () => {
    const fetchFn = vi.fn(async () => json({ value: [] }));
    const client = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never });

    await client.get('/me');

    const [, init] = fetchFn.mock.calls[0] as unknown as [string, RequestInit];
    expect((init.headers as Record<string, string>)['authorization']).toBe('Bearer the-token');
  });

  it('recognises the missing-licence 403 for what it is', async () => {
    const fetchFn = vi.fn(async () =>
      json(
        { error: { code: 'UnknownError', message: 'Failed to get license information for the user' } },
        403,
      ),
    );
    const client = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never });

    const error = await client.get('/me/chats').catch((caught: unknown) => caught);

    expect(error).toBeInstanceOf(GraphError);
    expect((error as GraphError).status).toBe(403);
    expect((error as GraphError).isLicenceProblem).toBe(true);
  });

  it('does not mistake an ordinary 403 for a licence problem', async () => {
    const fetchFn = vi.fn(async () =>
      json({ error: { code: 'Forbidden', message: 'Insufficient privileges' } }, 403),
    );
    const client = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never });

    const error = (await client.get('/me/chats').catch((caught: unknown) => caught)) as GraphError;

    expect(error.isLicenceProblem).toBe(false);
  });

  it('follows nextLink until the pages run out', async () => {
    const fetchFn = vi
      .fn()
      .mockResolvedValueOnce(
        json({ value: [{ id: 'a' }], '@odata.nextLink': 'https://graph.example/page2' }),
      )
      .mockResolvedValueOnce(json({ value: [{ id: 'b' }] }));
    const client = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never });

    expect(await client.getAll<{ id: string }>('/me/chats')).toEqual([{ id: 'a' }, { id: 'b' }]);
  });
});

describe('teams chats over graph', () => {
  it('sends the body as rendered HTML so line breaks and links survive Teams', async () => {
    const fetchFn = vi.fn(async () =>
      json({ id: 'sent', createdDateTime: '2026-08-19T10:00:00Z', body: { content: 'Hi' } }),
    );
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }),
    );

    await chats.sendMessage('19:a@thread.v2', 'Hi\nsee https://example.com/plan');

    const [url, init] = fetchFn.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toBe('https://graph.microsoft.com/v1.0/chats/19%3Aa%40thread.v2/messages');
    expect(JSON.parse(init.body as string)).toEqual({
      body: {
        contentType: 'html',
        content: '<p>Hi<br>see <a href="https://example.com/plan">https://example.com/plan</a></p>',
      },
    });
  });

  it('downloads a shared file through the /shares facade, never the raw contentUrl', async () => {
    const contentUrl = 'https://contoso.sharepoint.com/personal/x/Documents/ai-test/Test file.docx';
    const fetchFn = vi
      .fn()
      .mockResolvedValueOnce(
        json({
          id: 'm1',
          createdDateTime: '2026-08-19T08:00:00Z',
          body: { contentType: 'text', content: 'file' },
          attachments: [{ id: 'att-1', name: 'Test file.docx', contentType: 'reference', contentUrl }],
        }),
      )
      .mockResolvedValueOnce(
        new Response(new Uint8Array([80, 75, 3, 4]), {
          status: 200,
          headers: {
            'content-type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          },
        }),
      );
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }),
    );

    const payload = await chats.getAttachment('19:a@thread.v2', 'm1');

    const [downloadUrl, init] = fetchFn.mock.calls[1] as unknown as [string, RequestInit];
    const expectedShareId = `u!${Buffer.from(contentUrl, 'utf8')
      .toString('base64')
      .replace(/=+$/, '')
      .replace(/\+/g, '-')
      .replace(/\//g, '_')}`;
    expect(downloadUrl).toBe(
      `https://graph.microsoft.com/v1.0/shares/${expectedShareId}/driveItem/content`,
    );
    // The Graph token goes on the /shares call; fetching contentUrl directly answers 401.
    expect((init.headers as Record<string, string>)['authorization']).toBe('Bearer the-token');
    expect(payload.name).toBe('Test file.docx');
    expect(payload.contentType).toContain('wordprocessingml');
    expect([...payload.bytes]).toEqual([80, 75, 3, 4]);
  });

  it('downloads a pasted inline image from the message hostedContents', async () => {
    const fetchFn = vi
      .fn()
      .mockResolvedValueOnce(
        json({
          id: 'm2',
          createdDateTime: '2026-08-19T08:00:00Z',
          body: {
            contentType: 'html',
            content: '<img src="https://g/chats/c/messages/m2/hostedContents/hc-1/$value">',
          },
        }),
      )
      .mockResolvedValueOnce(
        new Response(new Uint8Array([137, 80, 78, 71]), {
          status: 200,
          headers: { 'content-type': 'image/png' },
        }),
      );
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }),
    );

    const payload = await chats.getAttachment('19:a@thread.v2', 'm2');

    const [downloadUrl] = fetchFn.mock.calls[1] as unknown as [string];
    expect(downloadUrl).toBe(
      'https://graph.microsoft.com/v1.0/chats/19%3Aa%40thread.v2/messages/m2/hostedContents/hc-1/$value',
    );
    expect(payload.name).toBe('inline-image-1');
    expect(payload.contentType).toBe('image/png');
  });

  it('posts an image as hosted content using the temporary-id pattern', async () => {
    const fetchFn = vi.fn(async () =>
      json({ id: 'sent', createdDateTime: '2026-08-19T10:00:00Z', body: { content: '<img …>' } }),
    );
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }),
    );

    await chats.sendImage(
      '19:a@thread.v2',
      { bytes: new Uint8Array([1, 2, 3]), contentType: 'image/png' },
      'a <diagram>',
    );

    const [url, init] = fetchFn.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toBe('https://graph.microsoft.com/v1.0/chats/19%3Aa%40thread.v2/messages');
    const body = JSON.parse(init.body as string) as {
      body: { contentType: string; content: string };
      hostedContents: Array<Record<string, string>>;
    };
    expect(body.body.contentType).toBe('html');
    // The caption is escaped, and the img src refers to the hosted content by temporary id.
    expect(body.body.content).toContain('a &lt;diagram&gt;');
    expect(body.body.content).toContain('src="../hostedContents/1/$value"');
    expect(body.hostedContents).toEqual([
      {
        '@microsoft.graph.temporaryId': '1',
        contentBytes: Buffer.from([1, 2, 3]).toString('base64'),
        contentType: 'image/png',
      },
    ]);
  });

  it('shares a file by uploading to OneDrive and attaching the driveItem by its eTag GUID', async () => {
    const fetchFn = vi
      .fn()
      .mockResolvedValueOnce(
        json({
          id: 'ITEM-ID',
          name: 'notes.txt',
          eTag: '"{ABC123DE-0000-1111-2222-333344445555},2"',
          webUrl: 'https://contoso-my.sharepoint.com/personal/x/Documents/ai-test/notes.txt',
        }),
      )
      .mockResolvedValueOnce(
        json({ id: 'sent', createdDateTime: '2026-08-19T10:00:00Z', body: { content: 'file' } }),
      );
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }),
      { uploadDir: 'ai-test' },
    );

    await chats.sendFile(
      '19:a@thread.v2',
      { bytes: new Uint8Array([104, 105]), name: 'notes.txt' },
      'here',
    );

    const [uploadUrl, uploadInit] = fetchFn.mock.calls[0] as unknown as [string, RequestInit];
    expect(uploadUrl).toBe(
      'https://graph.microsoft.com/v1.0/me/drive/root:/ai-test/notes.txt:/content',
    );
    expect(uploadInit.method).toBe('PUT');

    const [, messageInit] = fetchFn.mock.calls[1] as unknown as [string, RequestInit];
    const body = JSON.parse(messageInit.body as string) as {
      body: { content: string };
      attachments: Array<Record<string, string>>;
    };
    // The attachment id is the GUID inside the eTag, not the driveItem id.
    expect(body.attachments).toEqual([
      {
        id: 'ABC123DE-0000-1111-2222-333344445555',
        contentType: 'reference',
        contentUrl: 'https://contoso-my.sharepoint.com/personal/x/Documents/ai-test/notes.txt',
        name: 'notes.txt',
      },
    ]);
    expect(body.body.content).toContain('<attachment id="ABC123DE-0000-1111-2222-333344445555">');
  });

  it('replies by quoting the original message as a messageReference attachment', async () => {
    const fetchFn = vi
      .fn()
      .mockResolvedValueOnce(
        json({
          id: 'orig-1',
          createdDateTime: '2026-08-19T08:00:00Z',
          from: { user: { id: 'oid-9', displayName: 'Alice Anderson' } },
          body: { contentType: 'html', content: '<p>Where does the report end up?</p>' },
        }),
      )
      .mockResolvedValueOnce(
        json({ id: 'reply-1', createdDateTime: '2026-08-19T10:00:00Z', body: { content: 'Here' } }),
      );
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }),
    );

    await chats.replyToMessage('19:a@thread.v2', 'orig-1', 'In the ai-test folder');

    const [postUrl, init] = fetchFn.mock.calls[1] as unknown as [string, RequestInit];
    expect(postUrl).toBe('https://graph.microsoft.com/v1.0/chats/19%3Aa%40thread.v2/messages');
    const body = JSON.parse(init.body as string) as {
      body: { contentType: string; content: string };
      attachments: Array<{ id: string; contentType: string; content: string }>;
    };
    expect(body.body.content).toContain('<attachment id="orig-1">');
    expect(body.body.content).toContain('In the ai-test folder');
    expect(body.attachments[0]?.id).toBe('orig-1');
    expect(body.attachments[0]?.contentType).toBe('messageReference');
    expect(JSON.parse(body.attachments[0]!.content)).toEqual({
      messageId: 'orig-1',
      messagePreview: 'Where does the report end up?',
      messageSender: {
        application: null,
        device: null,
        user: { userIdentityType: 'aadUser', id: 'oid-9', displayName: 'Alice Anderson' },
      },
    });
  });

  it('edits a message with a PATCH carrying the replacement body', async () => {
    const fetchFn = vi.fn(async () => new Response(null, { status: 204 }));
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }),
    );

    await chats.editMessage('19:a@thread.v2', 'm7', 'corrected text');

    const [url, init] = fetchFn.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toBe('https://graph.microsoft.com/v1.0/chats/19%3Aa%40thread.v2/messages/m7');
    expect(init.method).toBe('PATCH');
    expect(JSON.parse(init.body as string)).toEqual({
      body: { contentType: 'html', content: '<p>corrected text</p>' },
    });
  });

  it('soft-deletes a message through the /me softDelete action, with no body', async () => {
    const fetchFn = vi.fn(async () => new Response(null, { status: 204 }));
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }),
    );

    await chats.deleteMessage('19:a@thread.v2', 'm7');

    const [url, init] = fetchFn.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toBe(
      'https://graph.microsoft.com/v1.0/me/chats/19%3Aa%40thread.v2/messages/m7/softDelete',
    );
    expect(init.method).toBe('POST');
    expect(init.body).toBeUndefined();
  });

  it('surfaces Graph\'s own refusal when editing someone else\'s message', async () => {
    const fetchFn = vi.fn(async () =>
      json({ error: { code: 'Forbidden', message: 'User is not the sender of the message' } }, 403),
    );
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }),
    );

    await expect(chats.editMessage('19:a@thread.v2', 'not-mine', 'x')).rejects.toThrow(
      /not the sender/,
    );
  });

  it('refuses to share a file when OneDrive returns no usable eTag', async () => {
    const fetchFn = vi.fn(async () => json({ id: 'ITEM-ID', name: 'notes.txt' }));
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }),
    );

    await expect(
      chats.sendFile('19:a@thread.v2', { bytes: new Uint8Array([1]), name: 'notes.txt' }),
    ).rejects.toThrow(/eTag/);
    // The message post never happened.
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });

  it('falls back to the member list when a group chat has no topic', async () => {
    const fetchFn = vi.fn(async () =>
      json({
        value: [
          {
            id: '19:a@thread.v2',
            chatType: 'group',
            topic: null,
            members: [{ displayName: 'Alice' }, { displayName: 'Bob' }],
          },
        ],
      }),
    );
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }),
    );

    expect((await chats.listChats())[0]?.topic).toBe('Alice, Bob');
  });
});

describe('graph client — empty-success bodies (the 11-copy broadcast incident, 2026-08-24)', () => {
  it('treats a 204 answer to postNoContent() as success, not a parse error', async () => {
    // setReaction and friends answer 204 No Content; the old unconditional response.json()
    // threw SyntaxError on that success, which is how a landed write got reported as failed.
    const fetchFn = vi.fn(async () => new Response(null, { status: 204 }));
    const client = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never });

    await expect(client.postNoContent('/chats/x/messages/1/setReaction', { reactionType: '👍' }))
      .resolves.toBeUndefined();
  });

  it('an empty success to a CREATE post() is an unknown outcome said plainly, never a phantom resource', async () => {
    // A create must hand back the resource; defaulting an empty body would give callers a
    // message with no id. The error text says the write may have landed — the readback layer
    // is what turns that honesty into a safe retry.
    const fetchFn = vi.fn(async () => new Response('', { status: 200 }));
    const client = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never });

    const error = (await client.post('/chats/x/messages', {}).catch((c: unknown) => c)) as GraphError;
    expect(error).toBeInstanceOf(GraphError);
    expect(error.code).toBe('EmptyCreateResponse');
    expect(error.message).toMatch(/unknown/);
  });
});

describe('graph client — throttle handling on reads', () => {
  it('a Retry-After longer than one call may sleep: no retry, honest 429 now, gate closed for the FULL window', async () => {
    let now = 0;
    const waits: number[] = [];
    const fetchFn = vi.fn().mockResolvedValueOnce(
      new Response(JSON.stringify({ error: { code: 'TooManyRequests', message: 'throttled' } }), {
        status: 429,
        headers: { 'retry-after': '300', 'content-type': 'application/json' },
      }),
    );
    const client = new GraphClient({
      tokenProvider: stubToken,
      fetchFn: fetchFn as never,
      sleepFn: async (ms) => { waits.push(ms); now += ms; },
      nowFn: () => now,
    });

    await expect(client.get('/me')).rejects.toMatchObject({ status: 429, retryAfterSeconds: 300 });
    expect(fetchFn).toHaveBeenCalledTimes(1); // sleeping a minute to fail locally would be theatre
    expect(waits).toEqual([]);
    expect(client.throttledForMs()).toBe(300_000); // the window Graph named, not a 60 s stand-in
  });

  it('a 429 with NO Retry-After still waits the gate out before the single retry (not a 1 s exponential)', async () => {
    let now = 0;
    const waits: number[] = [];
    const fetchFn = vi
      .fn()
      .mockResolvedValueOnce(new Response(JSON.stringify({ error: { code: 'TooManyRequests', message: 'throttled' } }), { status: 429 }))
      .mockResolvedValueOnce(json({ id: 'fine' }));
    const client = new GraphClient({
      tokenProvider: stubToken,
      fetchFn: fetchFn as never,
      sleepFn: async (ms) => { waits.push(ms); now += ms; },
      nowFn: () => now,
    });

    expect(await client.get('/me')).toEqual({ id: 'fine' });
    expect(waits).toEqual([30_000]); // DEFAULT_THROTTLE_WINDOW_MS — a shorter sleep retries into a closed gate
  });

  it('retries a 429 GET after the Retry-After the server names, then succeeds', async () => {
    const waits: number[] = [];
    const fetchFn = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(JSON.stringify({ error: { code: 'TooManyRequests', message: 'throttled' } }), {
          status: 429,
          headers: { 'retry-after': '7', 'content-type': 'application/json' },
        }),
      )
      .mockResolvedValueOnce(json({ id: 'fine' }));
    let now = 0;
    const client = new GraphClient({
      tokenProvider: stubToken,
      fetchFn: fetchFn as never,
      sleepFn: async (ms) => { waits.push(ms); now += ms; },
      nowFn: () => now,
    });

    expect(await client.get('/me')).toEqual({ id: 'fine' });
    expect(waits).toEqual([7000]);
  });

  it('gives up after the retry budget and surfaces the throttle with its Retry-After', async () => {
    const fetchFn = vi.fn(async () =>
      new Response(JSON.stringify({ error: { code: 'TooManyRequests', message: 'throttled' } }), {
        status: 429,
        headers: { 'retry-after': '3', 'content-type': 'application/json' },
      }),
    );
    let now = 0;
    const client = new GraphClient({
      tokenProvider: stubToken,
      fetchFn: fetchFn as never,
      sleepFn: async (ms) => { now += ms; },
      nowFn: () => now,
    });

    const error = (await client.get('/me').catch((c: unknown) => c)) as GraphError;
    expect(error).toBeInstanceOf(GraphError);
    expect(error.status).toBe(429);
    expect(error.retryAfterSeconds).toBe(3);
    expect(fetchFn.mock.calls.length).toBe(2); // 1 try + ONE retry, then honesty
  });

  it('never auto-retries a POST — a write is not idempotent', async () => {
    const fetchFn = vi.fn(async () =>
      new Response(JSON.stringify({ error: { code: 'TooManyRequests', message: 'throttled' } }), {
        status: 429,
        headers: { 'retry-after': '1', 'content-type': 'application/json' },
      }),
    );
    const client = new GraphClient({
      tokenProvider: stubToken,
      fetchFn: fetchFn as never,
      sleepFn: async () => {},
    });

    await expect(client.post('/chats/x/messages', {})).rejects.toThrow();
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });
});

describe('reactions', () => {
  it('setReaction posts the bare payload and accepts the 204 answer', async () => {
    const fetchFn = vi.fn(async () => new Response(null, { status: 204 }));
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }),
    );

    await expect(chats.setReaction('19:a@thread.v2', 'msg-1', '👍')).resolves.toBeUndefined();

    const [url, init] = fetchFn.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toContain('/messages/msg-1/setReaction');
    // The payload is {reactionType} directly — wrapping it in {body: …} answers
    // "ReactionType cannot be null", the 2026-08-24 react.mjs bug.
    expect(JSON.parse(init.body as string)).toEqual({ reactionType: '👍' });
  });
});

describe('graph client — the throttle gate (2026-08-25: retries under 429 amplified an escalating throttle)', () => {
  const throttled = (retryAfter: string) =>
    new Response(JSON.stringify({ error: { code: 'TooManyRequests', message: 'throttled' } }), {
      status: 429,
      headers: { 'retry-after': retryAfter, 'content-type': 'application/json' },
    });

  it('after one 429, every request from this client fails fast until Retry-After has passed — no network hit', async () => {
    let now = 1_000_000;
    const fetchFn = vi.fn().mockResolvedValueOnce(throttled('30'));
    const client = new GraphClient({
      tokenProvider: stubToken,
      fetchFn: fetchFn as never,
      sleepFn: async () => {},
      readRetries: 0,
      nowFn: () => now,
    });

    await expect(client.get('/chats/a/messages')).rejects.toMatchObject({ status: 429 });
    // A second request inside the window: refused locally, fetch NOT called again.
    await expect(client.get('/chats/b/messages')).rejects.toMatchObject({ status: 429, code: 'LocallyThrottled' });
    await expect(client.post('/chats/b/messages', {})).rejects.toMatchObject({ status: 429, code: 'LocallyThrottled' });
    expect(fetchFn).toHaveBeenCalledTimes(1);

    now += 31_000;
    fetchFn.mockResolvedValueOnce(json({ value: [] }));
    expect(await client.get('/chats/b/messages')).toEqual({ value: [] });
    expect(fetchFn).toHaveBeenCalledTimes(2);
  });

  it('a GET retries at most ONCE after a 429, and never while the gate is closed', async () => {
    let now = 0;
    const waits: number[] = [];
    const fetchFn = vi.fn().mockResolvedValue(throttled('5'));
    const client = new GraphClient({
      tokenProvider: stubToken,
      fetchFn: fetchFn as never,
      sleepFn: async (ms) => { waits.push(ms); now += ms; },
      nowFn: () => now,
    });

    await expect(client.get('/me')).rejects.toMatchObject({ status: 429 });
    expect(fetchFn).toHaveBeenCalledTimes(2); // first try + exactly one retry
    expect(waits).toEqual([5000]);
  });

  it('a 429 without Retry-After closes the gate for a default window, not zero', async () => {
    let now = 0;
    const fetchFn = vi.fn().mockResolvedValueOnce(
      new Response(JSON.stringify({ error: { code: 'TooManyRequests', message: 'throttled' } }), { status: 429 }),
    );
    const client = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never, sleepFn: async () => {}, readRetries: 0, nowFn: () => now });

    await expect(client.get('/me')).rejects.toMatchObject({ status: 429 });
    await expect(client.get('/me')).rejects.toMatchObject({ code: 'LocallyThrottled' });
    expect(fetchFn).toHaveBeenCalledTimes(1);
    expect(client.throttledForMs()).toBeGreaterThan(0);
  });
});

describe('graph client — Retry-After is honoured on 503/504 as well (round 2)', () => {
  it('a 503 naming Retry-After 30 waits those 30 s, not a 1 s exponential', async () => {
    let now = 0;
    const waits: number[] = [];
    const fetchFn = vi
      .fn()
      .mockResolvedValueOnce(new Response('', { status: 503, headers: { 'retry-after': '30' } }))
      .mockResolvedValueOnce(json({ id: 'fine' }));
    const client = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never, sleepFn: async (ms) => { waits.push(ms); now += ms; }, nowFn: () => now });

    expect(await client.get('/me')).toEqual({ id: 'fine' });
    expect(waits).toEqual([30_000]);
  });

  it('the fail-now path keeps Graph\'s own message and code', async () => {
    const fetchFn = vi.fn(async () =>
      new Response(JSON.stringify({ error: { code: 'ApplicationThrottled', message: 'Rate limit is exceeded. Try again in 300 seconds.' } }), {
        status: 429, headers: { 'retry-after': '300', 'content-type': 'application/json' },
      }));
    const client = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never, sleepFn: async () => {}, nowFn: () => 0 });

    const error = (await client.get('/me').catch((c: unknown) => c)) as GraphError;
    expect(error.code).toBe('ApplicationThrottled');
    expect(error.message).toMatch(/300 seconds/);
  });
});
