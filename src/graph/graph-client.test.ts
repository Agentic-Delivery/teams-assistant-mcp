import { describe, expect, it, vi } from 'vitest';
import { GraphClient, GraphError } from './graph-client.js';
import type { MembersCachePort } from './teams-chats.js';
import { GraphTeamsChats } from './teams-chats.js';
import type { TokenProvider } from '../auth/token-provider.js';

const stubToken: TokenProvider = { kind: 'stub', getAccessToken: async () => 'the-token' };

// membersCache is a required GraphTeamsChats dependency (0.4.1 review: an optional cache let
// production silently fall back to the throttled /members endpoint with no test noticing). None
// of the behaviour in THIS file touches mention resolution, so every construction below wires a
// trivial no-op double — the member-cache behaviour itself is covered in teams-chats.test.ts.
const noMembersCache: MembersCachePort = { get: () => undefined, set: () => {} };

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
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

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
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

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
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

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
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

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
      { membersCache: noMembersCache, uploadDir: 'ai-test' });

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
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

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
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

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
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

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
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

    await expect(chats.editMessage('19:a@thread.v2', 'not-mine', 'x')).rejects.toThrow(
      /not the sender/,
    );
  });

  it('refuses to share a file when OneDrive returns no usable eTag', async () => {
    const fetchFn = vi.fn(async () => json({ id: 'ITEM-ID', name: 'notes.txt' }));
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

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
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

    expect((await chats.listChats())[0]?.topic).toBe('Alice, Bob');
  });

  it('keeps each member\'s AAD userId, not the membership id, alongside displayName', async () => {
    // $expand=members carries BOTH a membership `id` (composite, useless for a mention) and a
    // `userId` (the actual AAD id a mention needs) — see the live capture this fixture is based
    // on (2026-08-25, /chats/{id}/members). Dropping userId here is exactly the bug a mention
    // silently posting an <at> tag that never notifies would trace back to.
    const fetchFn = vi.fn(async () =>
      json({
        value: [
          {
            id: '19:a@thread.v2',
            chatType: 'group',
            topic: 'Pilot',
            members: [
              { id: 'MEMBERSHIP-COMPOSITE-ID', displayName: 'Garg, Shivankit', userId: 'aad-shiv' },
            ],
          },
        ],
      }),
    );
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

    expect((await chats.listChats())[0]?.members).toEqual([
      { id: 'aad-shiv', displayName: 'Garg, Shivankit' },
    ]);
  });
});

describe('graph client — del()', () => {
  it('DELETEs with no body and treats 204 as success', async () => {
    const fetchFn = vi.fn(async () => new Response(null, { status: 204 }));
    const client = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never });

    await expect(client.del('/chats/x/pinnedMessages/pin-1')).resolves.toBeUndefined();

    const [url, init] = fetchFn.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toBe('https://graph.microsoft.com/v1.0/chats/x/pinnedMessages/pin-1');
    expect(init.method).toBe('DELETE');
    expect(init.body).toBeUndefined();
  });

  it('surfaces a non-2xx DELETE as a GraphError, same as any other verb', async () => {
    const fetchFn = vi.fn(async () => json({ error: { code: 'NotFound', message: 'not found' } }, 404));
    const client = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never });

    await expect(client.del('/chats/x/pinnedMessages/pin-1')).rejects.toThrow(/not found/);
  });
});

describe('teams chats — @mentions', () => {
  it('resolveMentions fetches /chats/{id}/members and resolves against it', async () => {
    const fetchFn = vi.fn(async () =>
      json({
        value: [
          { id: 'membership-1', displayName: 'Garg, Shivankit', userId: 'aad-shiv' },
          { id: 'membership-2', displayName: 'Spännare, Johan', userId: 'aad-johan' },
        ],
      }),
    );
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

    const resolved = await chats.resolveMentions('19:a@thread.v2', ['Shiv']);

    const [url] = fetchFn.mock.calls[0] as unknown as [string];
    expect(url).toBe('https://graph.microsoft.com/v1.0/chats/19%3Aa%40thread.v2/members');
    expect(resolved).toEqual([{ name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' }]);
  });

  it('resolveMentions follows pagination on /members — a member on page 2 still resolves (review round 2, MINOR 3)', async () => {
    // A single-page get() used to silently truncate the roster to whatever fit on page 1, and a
    // member past that page would come back "No chat member matches" — a false negative that
    // looks exactly like the member genuinely not being in the chat.
    const fetchFn = vi
      .fn()
      .mockResolvedValueOnce(
        json({
          value: [{ id: 'membership-1', displayName: 'Alice Anderson', userId: 'aad-alice' }],
          '@odata.nextLink':
            'https://graph.microsoft.com/v1.0/chats/19%3Aa%40thread.v2/members?$skiptoken=x',
        }),
      )
      .mockResolvedValueOnce(
        json({ value: [{ id: 'membership-2', displayName: 'Garg, Shivankit', userId: 'aad-shiv' }] }),
      );
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

    const resolved = await chats.resolveMentions('19:a@thread.v2', ['Shiv']);

    expect(fetchFn).toHaveBeenCalledTimes(2);
    expect(resolved).toEqual([{ name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' }]);
  });

  it('sendMessage with mentions posts an <at> tag AND the parallel Graph mentions array', async () => {
    const fetchFn = vi.fn(async () =>
      json({ id: 'sent', createdDateTime: '2026-08-19T10:00:00Z', body: { content: 'x' } }),
    );
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });
    const mention = { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' };

    await chats.sendMessage('19:a@thread.v2', 'Shiv please review', [mention]);

    const [, init] = fetchFn.mock.calls[0] as unknown as [string, RequestInit];
    const body = JSON.parse(init.body as string) as {
      body: { content: string };
      mentions: Array<{ id: number; mentionText: string; mentioned: { user: { id: string; displayName: string } } }>;
    };
    expect(body.body.content).toBe('<p><at id="0">Garg, Shivankit</at> please review</p>');
    expect(body.mentions).toEqual([
      { id: 0, mentionText: 'Garg, Shivankit', mentioned: { user: { id: 'aad-shiv', displayName: 'Garg, Shivankit' } } },
    ]);
  });

  it('sendMessage without mentions posts no mentions array at all', async () => {
    const fetchFn = vi.fn(async () =>
      json({ id: 'sent', createdDateTime: '2026-08-19T10:00:00Z', body: { content: 'x' } }),
    );
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

    await chats.sendMessage('19:a@thread.v2', 'no mentions here');

    const [, init] = fetchFn.mock.calls[0] as unknown as [string, RequestInit];
    const body = JSON.parse(init.body as string) as Record<string, unknown>;
    expect('mentions' in body).toBe(false);
  });

  it('sendHtmlMessage replaces @{Name} placeholders and posts the mentions array', async () => {
    const fetchFn = vi.fn(async () =>
      json({ id: 'sent', createdDateTime: '2026-08-19T10:00:00Z', body: { content: 'x' } }),
    );
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });
    const mention = { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' };

    await chats.sendHtmlMessage('19:a@thread.v2', '<table><tr><td>@{Shiv}</td></tr></table>', [mention]);

    const [, init] = fetchFn.mock.calls[0] as unknown as [string, RequestInit];
    const body = JSON.parse(init.body as string) as { body: { content: string }; mentions: unknown[] };
    expect(body.body.content).toBe('<table><tr><td><at id="0">Garg, Shivankit</at></td></tr></table>');
    expect(body.mentions).toHaveLength(1);
  });

  it('replyToMessage carries mentions on the reply text, not the quoted original', async () => {
    const fetchFn = vi
      .fn()
      .mockResolvedValueOnce(
        json({
          id: 'orig-1',
          createdDateTime: '2026-08-19T08:00:00Z',
          from: { user: { id: 'oid-9', displayName: 'Alice Anderson' } },
          body: { contentType: 'html', content: '<p>Shiv already knows about this</p>' },
        }),
      )
      .mockResolvedValueOnce(
        json({ id: 'reply-1', createdDateTime: '2026-08-19T10:00:00Z', body: { content: 'x' } }),
      );
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });
    const mention = { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' };

    await chats.replyToMessage('19:a@thread.v2', 'orig-1', 'Shiv can you confirm?', [mention]);

    const [, init] = fetchFn.mock.calls[1] as unknown as [string, RequestInit];
    const body = JSON.parse(init.body as string) as { body: { content: string } };
    // The quote card is untouched html; only OUR text after it carries the <at> tag.
    expect(body.body.content).toBe(
      '<attachment id="orig-1"></attachment><p><at id="0">Garg, Shivankit</at> can you confirm?</p>',
    );
  });

  it('editMessage and editHtmlMessage carry mentions through the PATCH body', async () => {
    const fetchFn = vi.fn(async () => new Response(null, { status: 204 }));
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });
    const mention = { name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' };

    await chats.editMessage('19:a@thread.v2', 'm1', 'Shiv see above', [mention]);

    const [, init] = fetchFn.mock.calls[0] as unknown as [string, RequestInit];
    const body = JSON.parse(init.body as string) as { body: { content: string }; mentions: unknown[] };
    expect(body.body.content).toBe('<p><at id="0">Garg, Shivankit</at> see above</p>');
    expect(body.mentions).toHaveLength(1);
  });
});

describe('teams chats — pinned messages', () => {
  it('listPinnedMessages GETs $expand=message and previews the body through html-to-text', async () => {
    const fetchFn = vi.fn(async () =>
      json({
        value: [
          {
            id: 'pin-1',
            message: { id: 'm1', body: { contentType: 'html', content: '<p>Deploy plan</p>' } },
          },
        ],
      }),
    );
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

    const pinned = await chats.listPinnedMessages('19:a@thread.v2');

    const [url] = fetchFn.mock.calls[0] as unknown as [string];
    expect(url).toBe(
      'https://graph.microsoft.com/v1.0/chats/19%3Aa%40thread.v2/pinnedMessages?$expand=message',
    );
    expect(pinned).toEqual([{ id: 'pin-1', messageId: 'm1', preview: 'Deploy plan' }]);
  });

  it('pinMessage POSTs message@odata.bind, then RE-LISTS rather than trusting the POST response', async () => {
    // The single-pin-slot replace behaviour (verified live 2026-08-25) means the POST's own
    // response is not proof of the resulting state — only a fresh list is.
    const fetchFn = vi
      .fn()
      .mockResolvedValueOnce(json({ id: 'pin-2' }, 201))
      .mockResolvedValueOnce(
        json({
          value: [{ id: 'pin-2', message: { id: 'm2', body: { contentType: 'text', content: 'later' } } }],
        }),
      );
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

    const result = await chats.pinMessage('19:a@thread.v2', 'm2');

    const [postUrl, postInit] = fetchFn.mock.calls[0] as unknown as [string, RequestInit];
    expect(postUrl).toBe('https://graph.microsoft.com/v1.0/chats/19%3Aa%40thread.v2/pinnedMessages');
    expect(JSON.parse(postInit.body as string)).toEqual({
      'message@odata.bind':
        'https://graph.microsoft.com/v1.0/chats/19%3Aa%40thread.v2/messages/m2',
    });
    const [listUrl] = fetchFn.mock.calls[1] as unknown as [string];
    expect(listUrl).toContain('/pinnedMessages?$expand=message');
    expect(result).toEqual([{ id: 'pin-2', messageId: 'm2', preview: 'later' }]);
  });

  it('unpinMessage resolves the message id to its PIN id via a list, then DELETEs the pin id', async () => {
    const fetchFn = vi
      .fn()
      .mockResolvedValueOnce(
        json({
          value: [{ id: 'pin-9', message: { id: 'm9', body: { contentType: 'text', content: 'x' } } }],
        }),
      )
      .mockResolvedValueOnce(new Response(null, { status: 204 }));
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

    await chats.unpinMessage('19:a@thread.v2', 'm9');

    const [deleteUrl, deleteInit] = fetchFn.mock.calls[1] as unknown as [string, RequestInit];
    // The DELETE targets the PIN's id ("pin-9"), never the chat message id ("m9").
    expect(deleteUrl).toBe(
      'https://graph.microsoft.com/v1.0/chats/19%3Aa%40thread.v2/pinnedMessages/pin-9',
    );
    expect(deleteInit.method).toBe('DELETE');
  });

  it('unpinMessage refuses when the given message id is not the one currently pinned', async () => {
    const fetchFn = vi.fn(async () =>
      json({
        value: [{ id: 'pin-9', message: { id: 'm9', body: { contentType: 'text', content: 'x' } } }],
      }),
    );
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

    await expect(chats.unpinMessage('19:a@thread.v2', 'not-pinned')).rejects.toThrow(
      /not currently pinned/,
    );
    expect(fetchFn).toHaveBeenCalledTimes(1); // no DELETE was ever attempted
  });

  it('listPinnedMessages reads a 404 as "nothing pinned", not as a failure (verified live 2026-08-26)', async () => {
    // Empirical, undocumented Graph behaviour: GET .../pinnedMessages on a chat with zero pins
    // answers a bare 404 "NotFound" rather than 200 with an empty value array — caught live by
    // pinning, unpinning the chat's only pin, then listing again. Every other call on the same
    // chatId (including the unpin that just ran) keeps succeeding, so the 404 here means "no
    // pins", not "chat not found".
    const fetchFn = vi.fn(async () => json({ error: { code: 'NotFound', message: 'NotFound' } }, 404));
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

    await expect(chats.listPinnedMessages('19:a@thread.v2')).resolves.toEqual([]);
  });

  it('unpinMessage still refuses cleanly when the chat has nothing pinned at all (the 404-as-empty path)', async () => {
    const fetchFn = vi.fn(async () => json({ error: { code: 'NotFound', message: 'NotFound' } }, 404));
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

    await expect(chats.unpinMessage('19:a@thread.v2', 'm1')).rejects.toThrow(/not currently pinned/);
  });

  it('a non-404 listPinnedMessages failure is NOT swallowed — only 404 means "empty"', async () => {
    const fetchFn = vi.fn(async () => json({ error: { code: 'Forbidden', message: 'nope' } }, 403));
    const chats = new GraphTeamsChats(
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

    await expect(chats.listPinnedMessages('19:a@thread.v2')).rejects.toThrow(/nope/);
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
    expect(client.throttledForMs('/me')).toBe(300_000); // the window Graph named, not a 60 s stand-in
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
      new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never }), { membersCache: noMembersCache });

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
    const fetchFn = vi.fn(async () => throttled('5')); // a fresh Response per call — bodies are single-use
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
    expect(client.throttledForMs('/me')).toBeGreaterThan(0);
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

describe('graph client — 503 WITHOUT Retry-After (the non-triggering side)', () => {
  it('waits the short exponential pause, not the throttle default', async () => {
    let now = 0;
    const waits: number[] = [];
    const fetchFn = vi.fn().mockResolvedValueOnce(new Response('', { status: 503 })).mockResolvedValueOnce(json({ id: 'fine' }));
    const client = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never, sleepFn: async (ms) => { waits.push(ms); now += ms; }, nowFn: () => now });

    expect(await client.get('/me')).toEqual({ id: 'fine' });
    expect(waits).toEqual([1000]);
  });
});

describe('graph client — the gate is per resource family (2026-08-25: single-message GETs throttled, lists healthy)', () => {
  it('a 429 on /chats/{id}/messages/{id} does not close /chats/{id}/messages', async () => {
    const fetchFn = vi
      .fn()
      .mockResolvedValueOnce(new Response(JSON.stringify({ error: { code: 'TooManyRequests', message: 't' } }), { status: 429, headers: { 'retry-after': '60', 'content-type': 'application/json' } }))
      .mockResolvedValueOnce(json({ value: [{ id: '1' }] }));
    const client = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never, sleepFn: async () => {}, readRetries: 0, nowFn: () => 0 });

    await expect(client.get('/chats/19%3Aabc%40thread.v2/messages/1787644105434')).rejects.toMatchObject({ status: 429 });
    expect(await client.get('/chats/19%3Aabc%40thread.v2/messages?$top=50')).toEqual({ value: [{ id: '1' }] });
    expect(client.throttledForMs('/chats/19%3Aabc%40thread.v2/messages/999')).toBe(60_000);
    expect(client.throttledForMs('/chats/19%3Aabc%40thread.v2/messages')).toBe(0);
  });

  it('gateKeyFor blanks ids but keeps the resource shape', () => {
    expect(GraphClient.gateKeyFor('/chats/19%3Aabc%40thread.v2/messages/1787644105434')).toBe('/chats/*/messages/*');
    expect(GraphClient.gateKeyFor('/chats/19%3Aabc%40thread.v2/messages?$top=50')).toBe('/chats/*/messages');
    expect(GraphClient.gateKeyFor('/me/chats')).toBe('/me/chats');
    expect(GraphClient.gateKeyFor('https://graph.microsoft.com/v1.0/me/chats?$top=1')).toBe('/me/chats');
  });
});

describe('teams chats — single-message fetch falls back to the list under throttle', () => {
  const throttled = () => new Response(JSON.stringify({ error: { code: 'TooManyRequests', message: 't' } }), { status: 429, headers: { 'retry-after': '60', 'content-type': 'application/json' } });
  const raw = (id: string, text: string) => ({ id, createdDateTime: '2026-08-25T07:48:25Z', from: { user: { displayName: 'Kleivdal, Celine', id: 'u1' } }, body: { contentType: 'text', content: text }, attachments: [] });

  it('replyToMessage: the original is found in the recent list and the quoted reply is posted', async () => {
    const fetchFn = vi.fn(async (url: string, init: RequestInit) => {
      if (init.method === 'POST') return json({ ...raw('new-1', 'reply'), from: { user: { displayName: 'Assistant' } } }, 201);
      if (url.endsWith('/messages/1787644105434')) return throttled();
      return json({ value: [raw('other', 'x'), raw('1787644105434', 'Logo:')] });
    });
    const sleeps: number[] = [];
    const chats = new GraphTeamsChats(new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never, sleepFn: async (ms) => void sleeps.push(ms), nowFn: () => 0 }), { membersCache: noMembersCache });

    const sent = await chats.replyToMessage('19:a@thread.v2', '1787644105434', 'Takk!');

    expect(sent.id).toBe('new-1');
    // The throttled endpoint was hit exactly ONCE, and nobody slept: the fallback is cheaper than
    // the client's own retry, so fetchMessage opts out of it.
    expect(fetchFn.mock.calls.filter(([url]) => String(url).endsWith('/messages/1787644105434'))).toHaveLength(1);
    expect(sleeps).toEqual([]);
    const post = fetchFn.mock.calls.find(([, init]) => (init as RequestInit).method === 'POST')!;
    expect(JSON.parse((post[1] as RequestInit).body as string).attachments[0].content).toMatch(/"messagePreview":"Logo:"/);
  });

  it('replyToMessage: not in the last 50 either → a named MessageFetchThrottled error, nothing posted', async () => {
    const fetchFn = vi.fn(async (url: string) => (url.endsWith('/messages/old-id') ? throttled() : json({ value: [raw('other', 'x')] })));
    const chats = new GraphTeamsChats(new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never, sleepFn: async () => {}, nowFn: () => 0 }), { membersCache: noMembersCache });

    await expect(chats.replyToMessage('19:a@thread.v2', 'old-id', 'hi')).rejects.toMatchObject({ code: 'MessageFetchThrottled', status: 429, retryAfterSeconds: 60 });
    expect(fetchFn.mock.calls.some(([, init]) => (init as RequestInit)?.method === 'POST')).toBe(false);
  });
});

describe('graph client — global vs family gates, drive paths (review round 1 of the per-family change)', () => {
  const appThrottled = () => new Response(JSON.stringify({ error: { code: 'ApplicationThrottled', message: 'Rate limit is exceeded. Try again in 300 seconds.' } }), { status: 429, headers: { 'retry-after': '300', 'content-type': 'application/json' } });

  it('an APPLICATION-wide 429 (by error code) closes every family, not just the one that answered', async () => {
    const fetchFn = vi.fn().mockResolvedValueOnce(appThrottled());
    const client = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never, sleepFn: async () => {}, readRetries: 0, nowFn: () => 0 });

    await expect(client.get('/me/chats')).rejects.toMatchObject({ status: 429, code: 'ApplicationThrottled' });
    await expect(client.post('/chats/x/messages', {})).rejects.toMatchObject({ code: 'LocallyThrottled' });
    await expect(client.get('/chats/x/messages/1')).rejects.toMatchObject({ code: 'LocallyThrottled' });
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });

  it('a plain per-resource 429 (TooManyRequests) closes only its own family — the non-triggering side', async () => {
    const fetchFn = vi
      .fn()
      .mockResolvedValueOnce(new Response(JSON.stringify({ error: { code: 'TooManyRequests', message: 't' } }), { status: 429, headers: { 'retry-after': '60', 'content-type': 'application/json' } }))
      .mockResolvedValueOnce(json({ id: 'posted' }));
    const client = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never, sleepFn: async () => {}, readRetries: 0, nowFn: () => 0 });

    await expect(client.get('/chats/x/messages/1')).rejects.toMatchObject({ status: 429 });
    expect(await client.post('/chats/x/messages', {})).toEqual({ id: 'posted' });
  });

  it('OneDrive path syntax: every file under the drive root is ONE family', () => {
    expect(GraphClient.gateKeyFor('/me/drive/root:/ai-test/report.pdf:/content')).toBe('/me/drive/root:*');
    expect(GraphClient.gateKeyFor('/me/drive/root:/other-dir/x.png:/content')).toBe('/me/drive/root:*');
    expect(GraphClient.gateKeyFor('/shares/u!abc/driveItem/content')).toBe('/shares/*/driveItem/content');
  });
});

describe('teams chats — getAttachment under the single-message throttle', () => {
  it('falls back to the list, finds the attachment there, and downloads it', async () => {
    const throttled = () => new Response(JSON.stringify({ error: { code: 'TooManyRequests', message: 't' } }), { status: 429, headers: { 'retry-after': '60', 'content-type': 'application/json' } });
    const listed = { id: 'm-1', createdDateTime: '2026-08-25T07:48:25Z', from: { user: { displayName: 'Celine', id: 'u1' } }, body: { contentType: 'html', content: 'Logo:' }, attachments: [{ id: 'att-1', contentType: 'image/png', name: 'logo.png', contentUrl: null }] };
    const fetchFn = vi.fn(async (url: string) => {
      if (url.endsWith('/messages/m-1')) return throttled();
      if (url.includes('/hostedContents/')) return new Response(new Uint8Array([1, 2, 3]), { status: 200, headers: { 'content-type': 'image/png' } });
      return json({ value: [listed] });
    });
    const chats = new GraphTeamsChats(new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never, sleepFn: async () => {}, nowFn: () => 0 }), { membersCache: noMembersCache });

    const payload = await chats.getAttachment('19:a@thread.v2', 'm-1', 'att-1');

    expect(payload.contentType).toBe('image/png');
    expect(payload.bytes.length).toBe(3);
  });
});

describe('graph client — the peek must never break the gate (round 2 of the per-family change)', () => {
  it('a 429 with an HTML (non-JSON) body still closes the family gate for the named window', async () => {
    const fetchFn = vi.fn().mockResolvedValueOnce(new Response('<html>Too many requests</html>', { status: 429, headers: { 'retry-after': '45', 'content-type': 'text/html' } }));
    const client = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never, sleepFn: async () => {}, readRetries: 0, nowFn: () => 0 });

    await expect(client.get('/chats/x/messages/1')).rejects.toMatchObject({ status: 429 });
    expect(client.throttledForMs('/chats/x/messages/1')).toBe(45_000);
    expect(client.throttledForMs('/me/chats')).toBe(0); // the FAMILY gate, never the global one, for a body with no code
    await expect(client.get('/chats/x/messages/2')).rejects.toMatchObject({ code: 'LocallyThrottled' });
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });
});

describe('teams chats — the fallback is for throttles only (the non-triggering side)', () => {
  it('a 404 on the single-message GET propagates as-is; the list is never scanned', async () => {
    const fetchFn = vi.fn(async (url: string) => {
      if (url.endsWith('/messages/gone')) return new Response(JSON.stringify({ error: { code: 'NotFound', message: 'Message not found' } }), { status: 404, headers: { 'content-type': 'application/json' } });
      return json({ value: [] });
    });
    const chats = new GraphTeamsChats(new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never, sleepFn: async () => {}, nowFn: () => 0 }), { membersCache: noMembersCache });

    await expect(chats.replyToMessage('19:a@thread.v2', 'gone', 'hi')).rejects.toMatchObject({ status: 404, code: 'NotFound' });
    expect(fetchFn).toHaveBeenCalledTimes(1);
  });
});

describe('teams chats — getAttachment without an id skips the quote card', () => {
  it('on a quoted reply carrying a file, "the attachment" is the file, never the messageReference', async () => {
    const msg = { id: 'm-2', createdDateTime: '2026-08-25T07:48:25Z', from: { user: { displayName: 'Celine', id: 'u1' } }, body: { contentType: 'html', content: 'Logo:' },
      attachments: [{ id: 'quote-1', contentType: 'messageReference', name: null, contentUrl: null }, { id: 'file-1', contentType: 'reference', name: 'logo.png', contentUrl: 'https://tenant-my.sharepoint.com/personal/x/Documents/logo.png' }] };
    const fetchFn = vi.fn(async (url: string) => {
      if (url.endsWith('/messages/m-2')) return json(msg);
      if (url.includes('/shares/')) return new Response(new Uint8Array([7, 7]), { status: 200, headers: { 'content-type': 'image/png' } });
      return json({ value: [msg] });
    });
    const chats = new GraphTeamsChats(new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never, sleepFn: async () => {}, nowFn: () => 0 }), { membersCache: noMembersCache });

    const payload = await chats.getAttachment('19:a@thread.v2', 'm-2');

    expect(payload.name).toBe('logo.png');
    expect(payload.bytes.length).toBe(2);
    expect(fetchFn.mock.calls.some(([url]) => String(url).includes('/hostedContents/'))).toBe(false);
  });
});

describe('teams chats — a quoted reply with no file says so honestly', () => {
  it('reports "only a quoted-message card", never "no attachments"', async () => {
    const msg = { id: 'm-3', createdDateTime: '2026-08-25T07:48:25Z', from: { user: { displayName: 'Celine', id: 'u1' } }, body: { contentType: 'html', content: 'see above' },
      attachments: [{ id: 'quote-1', contentType: 'messageReference', name: null, contentUrl: null }] };
    const fetchFn = vi.fn(async () => json(msg));
    const chats = new GraphTeamsChats(new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as never, sleepFn: async () => {}, nowFn: () => 0 }), { membersCache: noMembersCache });

    await expect(chats.getAttachment('19:a@thread.v2', 'm-3')).rejects.toThrow(/only a quoted-message card/);
    expect(fetchFn.mock.calls.some(([url]) => String(url).includes('/hostedContents/'))).toBe(false);
  });
});
