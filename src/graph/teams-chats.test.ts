import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { GraphClient, GraphError } from './graph-client.js';
import { MembersCache } from './members-cache.js';
import { FileSelfIdCache, NullSelfIdCache, type SelfIdCacheEntry } from './self-id-cache.js';
import { GraphTeamsChats } from './teams-chats.js';
import type { TokenProvider } from '../auth/token-provider.js';

const stubToken: TokenProvider = { kind: 'stub', getAccessToken: async () => 'the-token' };
const CHAT = '19:pilot@thread.v2';

function json(body: unknown, status = 200, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json', ...headers },
  });
}

function membersPage(): Response {
  return json({
    value: [
      { userId: 'aad-shiv', displayName: 'Garg, Shivankit' },
      { userId: 'aad-johan', displayName: 'Spännare, Johan' },
    ],
  });
}

/** Counts only requests that actually hit the `/members` collection — the endpoint the live
 *  throttle diagnosis (0.4.1) is about; other Graph calls in the same test must not inflate it. */
function countingMembersFetch(respond: () => Response) {
  const calls: string[] = [];
  const fetchFn = vi.fn(async (url: string) => {
    if (String(url).includes('/members')) {
      calls.push(String(url));
      return respond();
    }
    throw new Error(`unexpected call in this test: ${String(url)}`);
  });
  return { fetchFn, calls };
}

describe('GraphTeamsChats.resolveMentions — member cache (0.4.1, live-diagnosed 429 on /members)', () => {
  let dir: string;
  let path: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'teams-chats-members-'));
    path = join(dir, 'members-cache.json');
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  function subject(fetchFn: typeof fetch, cache: MembersCache) {
    const graph = new GraphClient({ tokenProvider: stubToken, fetchFn });
    return new GraphTeamsChats(graph, { membersCache: cache });
  }

  it('resolves a mention straight from a pre-warmed cache — ZERO calls to /members', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [
      { id: 'aad-shiv', displayName: 'Garg, Shivankit' },
      { id: 'aad-johan', displayName: 'Spännare, Johan' },
    ]);
    const { fetchFn, calls } = countingMembersFetch(membersPage);
    const chats = subject(fetchFn as unknown as typeof fetch, cache);

    const resolved = await chats.resolveMentions(CHAT, ['Shiv']);

    expect(resolved).toEqual([{ name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' }]);
    expect(calls).toHaveLength(0);
  });

  it('a cache miss (never warmed) triggers exactly one refresh call, then resolves and persists it', async () => {
    const cache = new MembersCache({ path });
    const { fetchFn, calls } = countingMembersFetch(membersPage);
    const chats = subject(fetchFn as unknown as typeof fetch, cache);

    const resolved = await chats.resolveMentions(CHAT, ['Shiv']);

    expect(resolved).toEqual([{ name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' }]);
    expect(calls).toHaveLength(1);
    // The refresh's result is now on disk — a second resolution needs no further call.
    expect(cache.get(CHAT)).toEqual([
      { id: 'aad-shiv', displayName: 'Garg, Shivankit' },
      { id: 'aad-johan', displayName: 'Spännare, Johan' },
    ]);
  });

  // 2026-09-02 review NIT: cacheIfNonEmpty applies the "never persist an empty answer" principle
  // (originally sendFile-only) to resolveMentions too, in the one shared place both callers use.
  it('an empty refresh (Graph genuinely reports nobody) is never persisted to the cache either', async () => {
    const cache = new MembersCache({ path });
    const { fetchFn } = countingMembersFetch(() => json({ value: [] }));
    const chats = subject(fetchFn as unknown as typeof fetch, cache);

    await expect(chats.resolveMentions(CHAT, ['Shiv'])).rejects.toThrow(
      /No chat member matches mention "Shiv"/,
    );
    expect(cache.get(CHAT)).toBeUndefined();
  });

  it('a name the cache does not have refreshes once, then still fails with the existing clear error', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [{ id: 'aad-shiv', displayName: 'Garg, Shivankit' }]);
    const { fetchFn, calls } = countingMembersFetch(membersPage); // fresh list still has no "Nobody"
    const chats = subject(fetchFn as unknown as typeof fetch, cache);

    await expect(chats.resolveMentions(CHAT, ['Nobody'])).rejects.toThrow(
      /No chat member matches mention "Nobody"/,
    );
    expect(calls).toHaveLength(1); // one refresh, not a retry loop
  });

  it('a 429 on the refresh surfaces as a distinct THROTTLED error carrying Retry-After structurally, not a silent fallback', async () => {
    const cache = new MembersCache({ path }); // empty — refresh is unavoidable
    const { fetchFn, calls } = countingMembersFetch(() =>
      json({ error: { code: 'TooManyRequests', message: 'Too many requests' } }, 429, {
        // Past GraphClient's MAX_RETRY_SLEEP_MS (90s since 0.5.0 — the live-measured window is
        // 62s and must now be SLEPT, not refused), so this fails fast locally instead of
        // actually sleeping during the test run.
        'retry-after': '100',
      }),
    );
    const chats = subject(fetchFn as unknown as typeof fetch, cache);

    const error = await chats.resolveMentions(CHAT, ['Shiv']).catch((caught: unknown) => caught);

    expect(error).toBeInstanceOf(GraphError);
    expect((error as GraphError).status).toBe(429);
    expect((error as GraphError).code).toBe('MembersRefreshThrottled');
    // 0.4.1 review round 2: retryAfterSeconds is the ONE place the wait lives — the message text
    // itself must NOT also state it (retryAfterSuffix, shared by the CLI and guard(), is the
    // sole renderer, so it is never stated twice on the surfaces that actually show it to someone).
    expect((error as GraphError).retryAfterSeconds).toBe(100);
    expect((error as GraphError).message).not.toMatch(/retry after|\d+s/);
    // 2026-09-02 re-review MINOR: the refreshMembers extraction (shared with membersForInvite)
    // dropped this exact reassurance clause once already — pinned here so it cannot regress a
    // second time silently (same shape of loss as 0.4.1's Retry-After dropping off the MCP path).
    expect((error as GraphError).message).toMatch(/does not mean the name does not exist/);
    expect(calls).toHaveLength(1);
  });

  // 0.4.1 review round 1: the catch around a cache hit used to be unconditional, so an AMBIGUOUS
  // or EMPTY-NAME caller error (which a refresh cannot fix) silently spent a Graph call and only
  // then surfaced the same error — or, worse, a different one if the refresh happened to fail.
  it('an ambiguous name against the cached roster rethrows immediately — no refresh call spent on a caller error', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [
      { id: 'aad-shiv', displayName: 'Garg, Shivankit' },
      { id: 'aad-shiv2', displayName: 'Shivam, Kumar' },
    ]);
    const { fetchFn, calls } = countingMembersFetch(membersPage);
    const chats = subject(fetchFn as unknown as typeof fetch, cache);

    await expect(chats.resolveMentions(CHAT, ['Shiv'])).rejects.toThrow(/ambiguous/);
    expect(calls).toHaveLength(0);
  });

  it('an empty mention name rethrows immediately — no refresh call spent on a caller error', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [{ id: 'aad-shiv', displayName: 'Garg, Shivankit' }]);
    const { fetchFn, calls } = countingMembersFetch(membersPage);
    const chats = subject(fetchFn as unknown as typeof fetch, cache);

    await expect(chats.resolveMentions(CHAT, ['   '])).rejects.toThrow(/cannot be empty/);
    expect(calls).toHaveLength(0);
  });
});

describe('buildChats — the composition actually wires the members cache (0.4.1 review round 1)', () => {
  // MAJOR 1: an optional membersCache let the wiring in build-chats.ts be silently dropped with
  // no test noticing (mutation-verified: deleting the wiring left the full suite green). This
  // test drives a mention resolution through the REAL composition buildChats returns — not a
  // hand-built GraphTeamsChats — so a future regression here can only pass by actually reaching
  // Graph's /members endpoint, which the counting fetchFn below would then catch.
  let dir: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'build-chats-members-'));
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it('a pre-warmed cache at config.membersCachePath resolves a mention through buildChats() with ZERO /members calls', async () => {
    const { loadConfig } = await import('../config.js');
    const { buildChats } = await import('../build-chats.js');
    const configPath = join(dir, 'teams-mcp.config.json');
    const { writeFile } = await import('node:fs/promises');
    await writeFile(
      configPath,
      JSON.stringify({ allowedChats: [{ id: CHAT, label: 'pilot', canPost: true }] }),
    );
    const config = loadConfig({
      TEAMS_MCP_CONFIG: configPath,
      TEAMS_MCP_TENANT_ID: 'tenant',
      TEAMS_MCP_USERNAME: 'assistant@example.com',
      TEAMS_MCP_PASSWORD: 'secret',
      TEAMS_MCP_TOKEN_CACHE: join(dir, '.token-cache.json'),
    });

    // Pre-warm exactly where buildChats will look for it.
    new MembersCache({ path: config.membersCachePath }).set(CHAT, [
      { id: 'aad-shiv', displayName: 'Garg, Shivankit' },
    ]);

    const { calls } = countingMembersFetch(membersPage);
    const failEverythingElse = vi.fn(async () => {
      throw new Error('this test only exercises resolveMentions — nothing else should be called');
    });
    // buildChats wires its own GraphClient from RopcTokenProvider internals we cannot inject a
    // fetchFn into directly, so this proves the cache wiring the only honest way available at
    // this seam: resolveMentions must never reach the network at all on a cache hit, regardless
    // of what fetchFn would have answered — swap the global fetch for the test's duration.
    const originalFetch = globalThis.fetch;
    globalThis.fetch = (async (url: string, init?: RequestInit) => {
      if (String(url).includes('/members')) {
        calls.push(String(url));
        return membersPage();
      }
      return failEverythingElse(url, init);
    }) as typeof fetch;
    try {
      const { chats, tokenProvider } = buildChats(config);
      vi.spyOn(tokenProvider, 'getAccessToken').mockResolvedValue('fake-token');

      const resolved = await chats.resolveMentions(CHAT, ['Shiv']);

      expect(resolved).toEqual([{ name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' }]);
      expect(calls).toHaveLength(0);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

// 0.5.1 — same composition-wiring hazard the members-cache test above guards against (0.4.1
// review MAJOR 1): an optional selfIdCache wired inconsistently (or not at all) in build-chats.ts
// would leave production exactly where the 2026-09-03 incident found it, with every test still
// green, since GraphTeamsChats itself degrades safely to NullSelfIdCache. This test drives a real
// sendFile through the REAL composition buildChats returns, so a dropped wire here can only pass
// by actually reaching Graph's throttled /me endpoint, which the guard below would then catch.
describe('buildChats — the composition wires the self id cache (0.5.1)', () => {
  let dir: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'build-chats-self-id-'));
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it('a pre-warmed cache at config.selfIdCachePath resolves self through buildChats() with ZERO /me calls', async () => {
    const { loadConfig } = await import('../config.js');
    const { buildChats } = await import('../build-chats.js');
    const configPath = join(dir, 'teams-mcp.config.json');
    const { writeFile } = await import('node:fs/promises');
    await writeFile(
      configPath,
      JSON.stringify({ allowedChats: [{ id: CHAT, label: 'pilot', canPost: true }] }),
    );
    const config = loadConfig({
      TEAMS_MCP_CONFIG: configPath,
      TEAMS_MCP_TENANT_ID: 'tenant',
      TEAMS_MCP_USERNAME: 'assistant@example.com',
      TEAMS_MCP_PASSWORD: 'secret',
      TEAMS_MCP_TOKEN_CACHE: join(dir, '.token-cache.json'),
    });

    new MembersCache({ path: config.membersCachePath }).set(CHAT, [
      { id: 'aad-self', displayName: 'Assistant (AI)' },
      { id: 'aad-bob', displayName: 'Bob Brown' },
    ]);
    // Pre-warm exactly where buildChats will look for it.
    new FileSelfIdCache({ path: config.selfIdCachePath }).write({ id: 'aad-self', resolvedAt: 1 });

    const originalFetch = globalThis.fetch;
    globalThis.fetch = (async (url: string, init?: RequestInit) => {
      const u = String(url);
      const method = init?.method ?? 'GET';
      if (u.includes('/me?') && u.includes('select=id')) {
        throw new Error('must never call /me — the self id cache is warm');
      }
      if (u.includes('/root:') && method === 'PUT') {
        return json({
          id: 'drive-item-1',
          eTag: '"{ABCDEF12-3456-7890-ABCD-EF1234567890},1"',
          webUrl: 'https://contoso.sharepoint.com/personal/assistant/report.pdf',
          name: 'report.pdf',
        });
      }
      if (u.includes('/invite')) {
        return json({ value: [{ grantedToV2: { user: { id: 'aad-bob' } } }] });
      }
      if (u.includes('/messages') && method === 'POST') {
        return json({
          id: 'msg-composition',
          chatId: CHAT,
          createdDateTime: '2026-09-02T10:00:00Z',
          body: { contentType: 'html', content: '' },
        });
      }
      throw new Error(`this test only exercises sendFile — unexpected call: ${method} ${u}`);
    }) as typeof fetch;
    try {
      const { chats, tokenProvider } = buildChats(config);
      vi.spyOn(tokenProvider, 'getAccessToken').mockResolvedValue('fake-token');

      const sent = await chats.sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' });

      expect(sent.id).toBe('msg-composition');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  // MAJOR 1 (review round 2): the TEAMS_MCP_SELF_ID override wire from loadConfig into buildChats
  // (build-chats.ts's `...(config.selfIdOverride !== undefined ? { selfIdOverride: ... } : {})`)
  // had no test — deleting that one line left the full suite green, because the GraphTeamsChats
  // unit tests for scenario (d) construct selfIdOverride directly, never through config+buildChats
  // together. This drives the REAL composition end to end: a config built from TEAMS_MCP_SELF_ID,
  // a COLD self id cache (never written), and a live /me that would 429 if it were ever called.
  it('TEAMS_MCP_SELF_ID reaches buildChats() end to end — a cold cache and a 429ing /me are both bypassed', async () => {
    const { loadConfig } = await import('../config.js');
    const { buildChats } = await import('../build-chats.js');
    const configPath = join(dir, 'teams-mcp.config.json');
    const { writeFile } = await import('node:fs/promises');
    await writeFile(
      configPath,
      JSON.stringify({ allowedChats: [{ id: CHAT, label: 'pilot', canPost: true }] }),
    );
    const config = loadConfig({
      TEAMS_MCP_CONFIG: configPath,
      TEAMS_MCP_TENANT_ID: 'tenant',
      TEAMS_MCP_USERNAME: 'assistant@example.com',
      TEAMS_MCP_PASSWORD: 'secret',
      TEAMS_MCP_TOKEN_CACHE: join(dir, '.token-cache.json'),
      TEAMS_MCP_SELF_ID: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
    });
    expect(config.selfIdOverride).toBe('aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee');

    new MembersCache({ path: config.membersCachePath }).set(CHAT, [
      { id: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee', displayName: 'Assistant (AI)' },
      { id: 'aad-bob', displayName: 'Bob Brown' },
    ]);
    // Deliberately COLD: config.selfIdCachePath is never written in this test — the override
    // must win without ever touching either the cache or /me.
    expect(new FileSelfIdCache({ path: config.selfIdCachePath }).read()).toBeUndefined();

    const originalFetch = globalThis.fetch;
    globalThis.fetch = (async (url: string, init?: RequestInit) => {
      const u = String(url);
      const method = init?.method ?? 'GET';
      if (u.includes('/me?') && u.includes('select=id')) {
        // A live /me that WOULD 429 if reached — the override must never let this be called.
        return json({ error: { code: 'TooManyRequests', message: 'Too many requests' } }, 429, {
          'retry-after': '5',
        });
      }
      if (u.includes('/root:') && method === 'PUT') {
        return json({
          id: 'drive-item-override',
          eTag: '"{ABCDEF12-3456-7890-ABCD-EF1234567890},1"',
          webUrl: 'https://contoso.sharepoint.com/personal/assistant/report.pdf',
          name: 'report.pdf',
        });
      }
      if (u.includes('/invite')) {
        return json({ value: [{ grantedToV2: { user: { id: 'aad-bob' } } }] });
      }
      if (u.includes('/messages') && method === 'POST') {
        return json({
          id: 'msg-override-composition',
          chatId: CHAT,
          createdDateTime: '2026-09-02T10:00:00Z',
          body: { contentType: 'html', content: '' },
        });
      }
      throw new Error(`this test only exercises sendFile — unexpected call: ${method} ${u}`);
    }) as typeof fetch;
    try {
      const { chats, tokenProvider } = buildChats(config);
      vi.spyOn(tokenProvider, 'getAccessToken').mockResolvedValue('fake-token');

      const sent = await chats.sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' });

      expect(sent.id).toBe('msg-override-composition');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('GraphTeamsChats.sendFile — grants chat members read access on the uploaded item (bug fix 0.4.2, live-verified 2026-09-02: dead "can\'t be viewed" cards, only fixed by a manual /invite)', () => {
  let dir: string;
  let path: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'teams-chats-sendfile-'));
    path = join(dir, 'members-cache.json');
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  function subject(fetchFn: typeof fetch, cache: MembersCache) {
    const graph = new GraphClient({ tokenProvider: stubToken, fetchFn });
    return new GraphTeamsChats(graph, { membersCache: cache });
  }

  const uploadResponse = (itemId = 'drive-item-1') =>
    json({
      id: itemId,
      eTag: '"{ABCDEF12-3456-7890-ABCD-EF1234567890},1"',
      webUrl: 'https://contoso.sharepoint.com/personal/assistant/report.pdf',
      name: 'report.pdf',
    });

  function selfIdResponse() {
    return json({ id: 'aad-self' });
  }

  /** A realistic /invite success body: one permission entry per invited AAD user, each carrying
   *  its grant under grantedToV2.user.id — the exact field the 2026-09-02 live verification
   *  showed real grants land under (see KNOWN-ISSUES.md's wire-shape snapshot). Real Graph invite
   *  responses are read for their actual grants, not trusted by HTTP status alone (MAJOR fix). */
  function grantsFor(objectIds: readonly string[]) {
    return json({ value: objectIds.map((id) => ({ grantedToV2: { user: { id } } })) });
  }

  it('issues the /invite with the exact body contract, BEFORE the chat message post, and reads real grants back before trusting it', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [
      { id: 'aad-self', displayName: 'Assistant (AI)' },
      { id: 'aad-shiv', displayName: 'Garg, Shivankit' },
      { id: 'aad-johan', displayName: 'Spännare, Johan' },
    ]);
    const callOrder: string[] = [];
    let inviteBody: unknown;
    const fetchFn = vi.fn(async (url: string, init?: RequestInit) => {
      const u = String(url);
      const method = init?.method ?? 'GET';
      if (u.includes('/me?') && u.includes('select=id')) {
        return selfIdResponse();
      }
      if (u.includes('/root:') && method === 'PUT') {
        callOrder.push('upload');
        return uploadResponse();
      }
      if (u.includes('/drive/items/drive-item-1/invite')) {
        callOrder.push('invite');
        inviteBody = JSON.parse(String(init?.body));
        return grantsFor(['aad-shiv', 'aad-johan']);
      }
      if (u.includes('/messages') && method === 'POST') {
        callOrder.push('message');
        return json({
          id: 'msg-1',
          chatId: CHAT,
          createdDateTime: '2026-09-02T10:00:00Z',
          body: { contentType: 'html', content: '' },
        });
      }
      throw new Error(`unexpected call in this test: ${method} ${u}`);
    });
    const chats = subject(fetchFn as unknown as typeof fetch, cache);

    const sent = await chats.sendFile(CHAT, { bytes: new Uint8Array([1, 2, 3]), name: 'report.pdf' });

    expect(callOrder).toEqual(['upload', 'invite', 'message']);
    expect(inviteBody).toEqual({
      // Self excluded: it already owns the item — see resolveSelfId's doc comment.
      recipients: [{ objectId: 'aad-shiv' }, { objectId: 'aad-johan' }],
      requireSignIn: true,
      sendInvitation: false,
      roles: ['read'],
    });
    expect(sent.id).toBe('msg-1');
  });

  it('an /invite failure fails the send loudly — no chat message is ever posted (no dead card)', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [
      { id: 'aad-self', displayName: 'Assistant (AI)' },
      { id: 'aad-shiv', displayName: 'Garg, Shivankit' },
    ]);
    const fetchFn = vi.fn(async (url: string, init?: RequestInit) => {
      const u = String(url);
      const method = init?.method ?? 'GET';
      if (u.includes('/me?') && u.includes('select=id')) return selfIdResponse();
      if (u.includes('/root:') && method === 'PUT') return uploadResponse();
      if (u.includes('/invite')) {
        return json({ error: { code: 'AccessDenied', message: 'insufficient privileges' } }, 403);
      }
      if (u.includes('/messages') && method === 'POST') {
        // Deliberately no "grant"/"permission"/"invite" wording here — if the code under test
        // still reaches this branch pre-fix, the assertion below must fail on a REAL mismatch,
        // not accidentally pass because this guard's own wording happened to match the regex.
        throw new Error('TEST-ONLY GUARD: sendFile must not reach the chat message post on this path');
      }
      throw new Error(`unexpected call in this test: ${method} ${u}`);
    });
    const chats = subject(fetchFn as unknown as typeof fetch, cache);

    await expect(
      chats.sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' }),
    ).rejects.toThrow(/grant|permission|invite/i);
  });

  // MAJOR fix (2026-09-02 review): a 200 from /invite is not proof of a grant. Graph's own
  // invite action can answer HTTP success with an empty (or partial) value array when nothing
  // was actually granted — the previous fixtures ABOVE encoded exactly this zero-grant body as
  // the success case, which would have hidden this exact failure mode.
  it('a 200 /invite response with NO actual grants for the recipients fails loudly — no dead card despite HTTP success', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [
      { id: 'aad-self', displayName: 'Assistant (AI)' },
      { id: 'aad-bob', displayName: 'Bob Brown' },
    ]);
    const fetchFn = vi.fn(async (url: string, init?: RequestInit) => {
      const u = String(url);
      const method = init?.method ?? 'GET';
      if (u.includes('/me?') && u.includes('select=id')) return selfIdResponse();
      if (u.includes('/root:') && method === 'PUT') return uploadResponse();
      if (u.includes('/invite')) {
        return json({ value: [] }); // HTTP 200, but nothing was actually granted
      }
      if (u.includes('/messages') && method === 'POST') {
        throw new Error('TEST-ONLY GUARD: sendFile must not reach the chat message post on this path');
      }
      throw new Error(`unexpected call in this test: ${method} ${u}`);
    });
    const chats = subject(fetchFn as unknown as typeof fetch, cache);

    await expect(
      chats.sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' }),
    ).rejects.toThrow(/grant|invite/i);
  });

  // MAJOR fix (2026-09-02 review): the SAME as the test above, but with a PARTIAL grant — one
  // of two recipients actually got one, the other did not. Both must be checked; the whole send
  // must still fail loudly (nobody gets a dead card left unaddressed).
  it('a 200 /invite response with a PARTIAL grant (one of two recipients missing) fails loudly too', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [
      { id: 'aad-self', displayName: 'Assistant (AI)' },
      { id: 'aad-shiv', displayName: 'Garg, Shivankit' },
      { id: 'aad-johan', displayName: 'Spännare, Johan' },
    ]);
    const fetchFn = vi.fn(async (url: string, init?: RequestInit) => {
      const u = String(url);
      const method = init?.method ?? 'GET';
      if (u.includes('/me?') && u.includes('select=id')) return selfIdResponse();
      if (u.includes('/root:') && method === 'PUT') return uploadResponse();
      if (u.includes('/invite')) {
        return grantsFor(['aad-shiv']); // aad-johan never got a grant
      }
      if (u.includes('/messages') && method === 'POST') {
        throw new Error('TEST-ONLY GUARD: sendFile must not reach the chat message post on this path');
      }
      throw new Error(`unexpected call in this test: ${method} ${u}`);
    });
    const chats = subject(fetchFn as unknown as typeof fetch, cache);

    await expect(
      chats.sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' }),
    ).rejects.toThrow(/grant|invite/i);
  });

  // MAJOR 2 (2026-09-02 re-review): grantedToIdentitiesV2 is a documented Graph variant some
  // permission kinds use — several recipients' grants folded under ONE permission entry, an
  // array, instead of one grantedToV2 entry per recipient. Both are live-verified wire shapes
  // (see KNOWN-ISSUES.md); tolerating this one too means a shape drift degrades to acceptance,
  // not a false outage.
  it('accepts grantedToIdentitiesV2[].user.id as an equivalent grant echo — a documented Graph variant, not a false failure', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [
      { id: 'aad-self', displayName: 'Assistant (AI)' },
      { id: 'aad-shiv', displayName: 'Garg, Shivankit' },
      { id: 'aad-johan', displayName: 'Spännare, Johan' },
    ]);
    const fetchFn = vi.fn(async (url: string, init?: RequestInit) => {
      const u = String(url);
      const method = init?.method ?? 'GET';
      if (u.includes('/me?') && u.includes('select=id')) return selfIdResponse();
      if (u.includes('/root:') && method === 'PUT') return uploadResponse();
      if (u.includes('/invite')) {
        return json({
          value: [
            {
              grantedToIdentitiesV2: [{ user: { id: 'aad-shiv' } }, { user: { id: 'aad-johan' } }],
            },
          ],
        });
      }
      if (u.includes('/messages') && method === 'POST') {
        return json({
          id: 'msg-identities-variant',
          chatId: CHAT,
          createdDateTime: '2026-09-02T10:00:00Z',
          body: { contentType: 'html', content: '' },
        });
      }
      throw new Error(`unexpected call in this test: ${method} ${u}`);
    });
    const chats = subject(fetchFn as unknown as typeof fetch, cache);

    const sent = await chats.sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' });

    expect(sent.id).toBe('msg-identities-variant');
  });

  it('an empty/unresolvable member roster fails BEFORE the upload — no orphaned OneDrive item, and nothing useless is cached', async () => {
    const cache = new MembersCache({ path }); // never warmed -> triggers a refresh
    const fetchFn = vi.fn(async (url: string) => {
      const u = String(url);
      if (u.includes('/members')) {
        return json({ value: [] }); // Graph genuinely reports nobody
      }
      throw new Error(`unexpected call — upload/invite/message must never be reached: ${u}`);
    });
    const chats = subject(fetchFn as unknown as typeof fetch, cache);

    await expect(
      chats.sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' }),
    ).rejects.toThrow(/member|roster|resolve/i);
    // NIT fix: an empty roster is never worth persisting — the NEXT call should refresh again,
    // not trust a stale "confirmed empty" that was really "we could not tell".
    expect(cache.get(CHAT)).toBeUndefined();
  });

  it('a throttled member-list refresh fails the send loudly, naming the throttle — no upload attempted', async () => {
    const cache = new MembersCache({ path });
    const fetchFn = vi.fn(async (url: string) => {
      const u = String(url);
      if (u.includes('/members')) {
        return json({ error: { code: 'TooManyRequests', message: 'Too many requests' } }, 429, {
          // Past GraphClient's MAX_RETRY_SLEEP_MS (90s since 0.5.0: the live-measured 62s
          // window is now slept, not refused), so it fails fast locally instead of actually
          // sleeping during the test run (same reasoning as the resolveMentions 429 test above).
          'retry-after': '100',
        });
      }
      throw new Error(`unexpected call — upload must never be reached: ${u}`);
    });
    const chats = subject(fetchFn as unknown as typeof fetch, cache);

    const error = await chats
      .sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' })
      .catch((caught: unknown) => caught);

    expect(error).toBeInstanceOf(GraphError);
    expect((error as GraphError).status).toBe(429);
    expect((error as GraphError).retryAfterSeconds).toBe(100);
  });

  // Violating-double: Graph (an external system) reports a real chat member with NO AAD id —
  // the exact shape resolveMentions already treats as "cannot mention them"; sendFile must treat
  // it as "cannot grant them access", not silently post a card only the assistant itself can open.
  it('the only OTHER member present has no AAD id on record — fails loudly BEFORE upload rather than post a card they cannot open', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [
      { id: 'aad-self', displayName: 'Assistant (AI)' },
      { displayName: 'No Id Person' }, // Graph reported no id for this member
    ]);
    const fetchFn = vi.fn(async (url: string, init?: RequestInit) => {
      const u = String(url);
      const method = init?.method ?? 'GET';
      if (u.includes('/me?') && u.includes('select=id')) return selfIdResponse();
      throw new Error(`unexpected call — upload must never be reached: ${method} ${u}`);
    });
    const chats = subject(fetchFn as unknown as typeof fetch, cache);

    await expect(
      chats.sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' }),
    ).rejects.toThrow(/AAD id|grant/i);
  });

  // MAJOR fix (2026-09-02 review): a MIXED roster — one other member IS resolvable, one is NOT —
  // used to silently grant only the resolvable one and post anyway, leaving the unresolvable
  // member with a dead card. Now the whole send fails loudly instead, consistent with the
  // no-dead-cards contract (there is no partial grant this method will quietly accept).
  it('a MIXED roster (one other member resolvable, one not) fails loudly rather than silently granting only the resolvable one', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [
      { id: 'aad-self', displayName: 'Assistant (AI)' },
      { id: 'aad-bob', displayName: 'Bob Brown' }, // resolvable
      { displayName: 'No Id Person' }, // NOT resolvable
    ]);
    const fetchFn = vi.fn(async (url: string, init?: RequestInit) => {
      const u = String(url);
      const method = init?.method ?? 'GET';
      if (u.includes('/me?') && u.includes('select=id')) return selfIdResponse();
      throw new Error(`unexpected call — upload/invite must never be reached: ${method} ${u}`);
    });
    const chats = subject(fetchFn as unknown as typeof fetch, cache);

    await expect(
      chats.sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' }),
    ).rejects.toThrow(/AAD id|grant/i);
  });

  // BLOCKER fix (2026-09-02 review, live-probed): resolveSelfId() previously memoized a FAILURE
  // (undefined) just like a real id, and `member.id !== selfId` with selfId === undefined is
  // TRUE for every member who has an id and FALSE for every member who does NOT — i.e. it kept
  // the real assistant "as an other" and silently treated the id-less member as if THEY were
  // self, dropping them out of the exclusion set and out of the "who is unresolvable" check
  // entirely. MAJOR 1 (2026-09-02 re-review) went further: the FIRST fix's fallback — include
  // everyone, self included, as a "harmless but noisy" grant when self is undetermined — still
  // meant a /me outage orphaned one upload per sendFile attempt for its entire duration, with no
  // clean pre-upload signal. This version refuses BEFORE the upload instead, for any reason /me
  // could not be resolved — no orphan, no partial roster analysis, no dependency on what else is
  // in the roster (an id-less member or not).
  it('a failed self-id lookup refuses BEFORE the upload — no orphan, regardless of roster shape', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [
      { id: 'aad-self', displayName: 'Assistant (AI)' },
      { displayName: 'No Id Person' }, // Graph reported no id for this REAL member — irrelevant
      // here: the /me failure itself is refused on before this roster shape is ever examined.
    ]);
    const fetchFn = vi.fn(async (url: string, init?: RequestInit) => {
      const u = String(url);
      const method = init?.method ?? 'GET';
      if (u.includes('/me?') && u.includes('select=id')) {
        return json({ error: { code: 'Forbidden', message: 'insufficient privileges' } }, 403);
      }
      throw new Error(`unexpected call — upload/invite/message must never be reached: ${method} ${u}`);
    });
    const chats = subject(fetchFn as unknown as typeof fetch, cache);

    // 2026-09-02 re-review NIT: /determined|\/me/i used to also match this test's OWN "unexpected
    // call" guard text (which names the /me URL as part of the upload path it refuses) — deleting
    // the production refusal made the code fall through to the upload, tripping the test's guard
    // instead of the intended error, and the loose matcher could not tell the difference. Matching
    // the exact production wording closes that gap.
    await expect(
      chats.sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' }),
    ).rejects.toThrow(/own account id could not be determined/i);
  });

  // MAJOR 1 follow-up (2026-09-02 re-review): a 429 on the OPTIONAL-looking /me lookup used to
  // cost a real Retry-After sleep before failing anyway — now that an undetermined self id
  // refuses the whole send regardless of the reason, sleeping through the throttle only to still
  // refuse serves no purpose. readRetries: 0 on that GET means GraphClient never sleeps for it;
  // proven here with a REAL GraphClient (not a stub) and a spied sleepFn a mutation to that
  // option would cause to actually be invoked.
  it('a /me 429 is never slept through — sendFile refuses immediately rather than waiting out Retry-After', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [
      { id: 'aad-self', displayName: 'Assistant (AI)' },
      { id: 'aad-bob', displayName: 'Bob Brown' },
    ]);
    const sleepFn = vi.fn(async () => {});
    const fetchFn = vi.fn(async (url: string) => {
      const u = String(url);
      if (u.includes('/me?') && u.includes('select=id')) {
        return json({ error: { code: 'TooManyRequests', message: 'Too many requests' } }, 429, {
          'retry-after': '5', // short enough that a real sleep would NOT fast-fail on its own
        });
      }
      throw new Error(`unexpected call — upload/invite/message must never be reached: ${u}`);
    });
    const graph = new GraphClient({
      tokenProvider: stubToken,
      fetchFn: fetchFn as unknown as typeof fetch,
      sleepFn,
    });
    const chats = new GraphTeamsChats(graph, { membersCache: cache });

    // Same matcher-precision fix as the test above — the loose /determined|\/me/i also matched
    // this test's own "unexpected call" guard text.
    await expect(
      chats.sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' }),
    ).rejects.toThrow(/own account id could not be determined/i);
    expect(sleepFn).not.toHaveBeenCalled();
  });

  // BLOCKER fix, second half: a failed self-id lookup is NOT memoized for the process/instance
  // lifetime — the next sendFile call on the SAME GraphTeamsChats retries it. First call: /me
  // fails, self undetermined -> refused before touching upload/invite/message at all (MAJOR 1).
  // Second call: /me now succeeds, and self is properly excluded from the grant, send lands.
  it('a failed self-id lookup is retried on the NEXT sendFile call — first refuses, second recovers', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [
      { id: 'aad-self', displayName: 'Assistant (AI)' },
      { id: 'aad-bob', displayName: 'Bob Brown' },
    ]);
    let meCalls = 0;
    const invitedObjectIds: string[][] = [];
    const fetchFn = vi.fn(async (url: string, init?: RequestInit) => {
      const u = String(url);
      const method = init?.method ?? 'GET';
      if (u.includes('/me?') && u.includes('select=id')) {
        meCalls += 1;
        return meCalls === 1
          ? json({ error: { code: 'Forbidden', message: 'insufficient privileges' } }, 403)
          : selfIdResponse();
      }
      if (u.includes('/root:') && method === 'PUT') {
        return uploadResponse(`drive-item-${meCalls}`);
      }
      if (u.includes('/invite')) {
        const body = JSON.parse(String(init?.body)) as { recipients: Array<{ objectId: string }> };
        const objectIds = body.recipients.map((recipient) => recipient.objectId);
        invitedObjectIds.push(objectIds);
        return grantsFor(objectIds);
      }
      if (u.includes('/messages') && method === 'POST') {
        return json({
          id: `msg-${meCalls}`,
          chatId: CHAT,
          createdDateTime: '2026-09-02T10:00:00Z',
          body: { contentType: 'html', content: '' },
        });
      }
      throw new Error(`unexpected call in this test: ${method} ${u}`);
    });
    const chats = subject(fetchFn as unknown as typeof fetch, cache);

    // First call: /me fails -> refused before ever touching upload/invite/message.
    await expect(
      chats.sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' }),
    ).rejects.toThrow(/determined|\/me/i);
    expect(invitedObjectIds).toHaveLength(0);

    // Second call: /me now succeeds -> self is properly excluded and the send lands.
    await chats.sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' });
    expect(invitedObjectIds).toEqual([['aad-bob']]);
  });

  it('a chat with no OTHER members (assistant-only roster) skips the invite entirely and still sends', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [{ id: 'aad-self', displayName: 'Assistant (AI)' }]);
    const fetchFn = vi.fn(async (url: string, init?: RequestInit) => {
      const u = String(url);
      const method = init?.method ?? 'GET';
      if (u.includes('/me?') && u.includes('select=id')) return selfIdResponse();
      if (u.includes('/root:') && method === 'PUT') return uploadResponse();
      if (u.includes('/invite')) throw new Error('must never invite when there is nobody else to grant to');
      if (u.includes('/messages') && method === 'POST') {
        return json({
          id: 'msg-solo',
          chatId: CHAT,
          createdDateTime: '2026-09-02T10:00:00Z',
          body: { contentType: 'html', content: '' },
        });
      }
      throw new Error(`unexpected call in this test: ${method} ${u}`);
    });
    const chats = subject(fetchFn as unknown as typeof fetch, cache);

    const sent = await chats.sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' });

    expect(sent.id).toBe('msg-solo');
  });

  // 0.5.1 — live incident 2026-09-03: eight consecutive teams-send-file CLI attempts over 20
  // minutes each failed the pre-upload refusal above, because every CLI invocation is a fresh
  // process and resolveSelfId's in-memory memo dies with it, paying (and losing) a throttled
  // `/me` call every single time. The account's own id never changes, so it is now persisted to
  // disk and read back BEFORE any `/me` call is attempted.

  // Scenario (a): a warm persisted cache answers resolveSelfId with ZERO /me calls.
  it('(0.5.1 scenario a) a warm persisted self-id cache resolves self with ZERO /me calls', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [
      { id: 'aad-self', displayName: 'Assistant (AI)' },
      { id: 'aad-bob', displayName: 'Bob Brown' },
    ]);
    const warmSelfIdCache = {
      read: (): SelfIdCacheEntry | undefined => ({ id: 'aad-self', resolvedAt: 1_000 }),
      write: () => {
        throw new Error('must not write — this test only exercises a warm read');
      },
    };
    const fetchFn = vi.fn(async (url: string, init?: RequestInit) => {
      const u = String(url);
      const method = init?.method ?? 'GET';
      if (u.includes('/me?') && u.includes('select=id')) {
        throw new Error('must never call /me when the persisted self-id cache is warm');
      }
      if (u.includes('/root:') && method === 'PUT') return uploadResponse();
      if (u.includes('/invite')) return grantsFor(['aad-bob']);
      if (u.includes('/messages') && method === 'POST') {
        return json({
          id: 'msg-cache-hit',
          chatId: CHAT,
          createdDateTime: '2026-09-02T10:00:00Z',
          body: { contentType: 'html', content: '' },
        });
      }
      throw new Error(`unexpected call in this test: ${method} ${u}`);
    });
    const graph = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as unknown as typeof fetch });
    const chats = new GraphTeamsChats(graph, { membersCache: cache, selfIdCache: warmSelfIdCache });

    const sent = await chats.sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' });

    expect(sent.id).toBe('msg-cache-hit');
  });

  // Scenario (b): /me itself fails, but a persisted cache is present — the send proceeds and the
  // self-exclusion grant math is unchanged from the healthy-/me case (aad-bob invited, aad-self
  // excluded), because the cached id is used as the source of truth for self-exclusion.
  it('(0.5.1 scenario b) a failed /me does not block the send when a persisted self-id cache is present', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [
      { id: 'aad-self', displayName: 'Assistant (AI)' },
      { id: 'aad-bob', displayName: 'Bob Brown' },
    ]);
    const warmSelfIdCache = {
      read: (): SelfIdCacheEntry | undefined => ({ id: 'aad-self', resolvedAt: 1_000 }),
      write: () => {
        throw new Error('must not write — a cache hit never re-resolves against /me');
      },
    };
    let inviteBody: unknown;
    const fetchFn = vi.fn(async (url: string, init?: RequestInit) => {
      const u = String(url);
      const method = init?.method ?? 'GET';
      if (u.includes('/me?') && u.includes('select=id')) {
        // Present to prove the cache is what is actually used — reaching this at all would
        // itself be a scenario-(a) failure, but scenario (b) is deliberately tolerant of /me
        // being unreachable (a network failure, not just a 429) and asserts the OUTCOME.
        return json({ error: { code: 'ServiceUnavailable', message: 'down' } }, 503);
      }
      if (u.includes('/root:') && method === 'PUT') return uploadResponse();
      if (u.includes('/invite')) {
        inviteBody = JSON.parse(String(init?.body));
        return grantsFor(['aad-bob']);
      }
      if (u.includes('/messages') && method === 'POST') {
        return json({
          id: 'msg-cache-fallback',
          chatId: CHAT,
          createdDateTime: '2026-09-02T10:00:00Z',
          body: { contentType: 'html', content: '' },
        });
      }
      throw new Error(`unexpected call in this test: ${method} ${u}`);
    });
    const graph = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as unknown as typeof fetch });
    const chats = new GraphTeamsChats(graph, { membersCache: cache, selfIdCache: warmSelfIdCache });

    const sent = await chats.sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' });

    expect(sent.id).toBe('msg-cache-fallback');
    expect(inviteBody).toEqual({
      recipients: [{ objectId: 'aad-bob' }],
      requireSignIn: true,
      sendInvitation: false,
      roles: ['read'],
    });
  });

  // Scenario (d): TEAMS_MCP_SELF_ID (validated GUID-shaped by config.ts before it reaches here)
  // wins over both the persisted cache and a live /me — the last-resort operator seed for a
  // throttle bad enough that even the cache never got its first write.
  it('(0.5.1 scenario d) selfIdOverride wins over the persisted cache and /me — zero of either called', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [
      { id: 'aad-override-self', displayName: 'Assistant (AI)' },
      { id: 'aad-bob', displayName: 'Bob Brown' },
    ]);
    const selfIdCache = {
      read: (): SelfIdCacheEntry | undefined => {
        throw new Error('must not read the persisted cache — the override wins outright');
      },
      write: () => {
        throw new Error('must not write the persisted cache — the override is never persisted');
      },
    };
    const fetchFn = vi.fn(async (url: string, init?: RequestInit) => {
      const u = String(url);
      const method = init?.method ?? 'GET';
      if (u.includes('/me?') && u.includes('select=id')) {
        throw new Error('must never call /me when an override is set');
      }
      if (u.includes('/root:') && method === 'PUT') return uploadResponse();
      if (u.includes('/invite')) return grantsFor(['aad-bob']);
      if (u.includes('/messages') && method === 'POST') {
        return json({
          id: 'msg-override',
          chatId: CHAT,
          createdDateTime: '2026-09-02T10:00:00Z',
          body: { contentType: 'html', content: '' },
        });
      }
      throw new Error(`unexpected call in this test: ${method} ${u}`);
    });
    const graph = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as unknown as typeof fetch });
    const chats = new GraphTeamsChats(graph, {
      membersCache: cache,
      selfIdCache,
      selfIdOverride: 'aad-override-self',
    });

    const sent = await chats.sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' });

    expect(sent.id).toBe('msg-override');
  });

  // Scenario (e): the FIRST successful /me call, with a cold (never-written) persisted cache,
  // writes the resolved id to disk — proven with a real FileSelfIdCache, not a double, so the
  // write is exercised end to end.
  it('(0.5.1 scenario e) a cold persisted cache is written on the first successful /me call', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [
      { id: 'aad-self', displayName: 'Assistant (AI)' },
      { id: 'aad-bob', displayName: 'Bob Brown' },
    ]);
    const selfIdCachePath = join(dir, '.self-id-cache.json');
    const selfIdCache = new FileSelfIdCache({ path: selfIdCachePath });
    expect(selfIdCache.read()).toBeUndefined(); // cold: never written before this test

    const fetchFn = vi.fn(async (url: string, init?: RequestInit) => {
      const u = String(url);
      const method = init?.method ?? 'GET';
      if (u.includes('/me?') && u.includes('select=id')) return selfIdResponse();
      if (u.includes('/root:') && method === 'PUT') return uploadResponse();
      if (u.includes('/invite')) return grantsFor(['aad-bob']);
      if (u.includes('/messages') && method === 'POST') {
        return json({
          id: 'msg-first-write',
          chatId: CHAT,
          createdDateTime: '2026-09-02T10:00:00Z',
          body: { contentType: 'html', content: '' },
        });
      }
      throw new Error(`unexpected call in this test: ${method} ${u}`);
    });
    const graph = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as unknown as typeof fetch });
    const chats = new GraphTeamsChats(graph, { membersCache: cache, selfIdCache });

    await chats.sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' });

    expect(selfIdCache.read()).toEqual({ id: 'aad-self', resolvedAt: expect.any(Number) });
  });

  // Scenario (c) — the existing pre-0.5.1 refusal test above already covers this: omitting
  // selfIdCache must refuse the send exactly like a real NullSelfIdCache would (no cache, no
  // override, /me fails). Pinned explicitly here so the equivalence itself — not just the
  // outcome — cannot silently drift (e.g. a future default that happens to swallow errors
  // differently from NullSelfIdCache would still pass the outcome-only test above).
  it('(0.5.1 scenario c) omitting selfIdCache behaves exactly like an explicit NullSelfIdCache', async () => {
    const cache = new MembersCache({ path });
    cache.set(CHAT, [{ id: 'aad-self', displayName: 'Assistant (AI)' }, { id: 'aad-bob', displayName: 'Bob Brown' }]);
    const failingMe = async (url: string) => {
      const u = String(url);
      if (u.includes('/me?') && u.includes('select=id')) {
        return json({ error: { code: 'Forbidden', message: 'insufficient privileges' } }, 403);
      }
      throw new Error(`unexpected call — upload/invite/message must never be reached: ${u}`);
    };
    const graphDefault = new GraphClient({ tokenProvider: stubToken, fetchFn: vi.fn(failingMe) as unknown as typeof fetch });
    const graphExplicit = new GraphClient({ tokenProvider: stubToken, fetchFn: vi.fn(failingMe) as unknown as typeof fetch });
    const chatsDefault = new GraphTeamsChats(graphDefault, { membersCache: cache });
    const chatsExplicit = new GraphTeamsChats(graphExplicit, { membersCache: cache, selfIdCache: new NullSelfIdCache() });

    const defaultResult = await chatsDefault
      .sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' })
      .catch((error: unknown) => error);
    const explicitResult = await chatsExplicit
      .sendFile(CHAT, { bytes: new Uint8Array([1]), name: 'report.pdf' })
      .catch((error: unknown) => error);

    expect(defaultResult).toBeInstanceOf(Error);
    expect((defaultResult as Error).message).toBe((explicitResult as Error).message);
  });
});

describe('GraphTeamsChats.sendImage — hosted content, no OneDrive item, no grant applies here', () => {
  it('never touches OneDrive or /invite — only the plain message-with-hostedContents POST', async () => {
    const fetchFn = vi.fn(async (url: string, init?: RequestInit) => {
      const u = String(url);
      const method = init?.method ?? 'GET';
      if (u.includes('/messages') && method === 'POST') {
        return json({
          id: 'img-1',
          chatId: CHAT,
          createdDateTime: '2026-09-02T10:00:00Z',
          body: { contentType: 'html', content: '' },
        });
      }
      throw new Error(`unexpected call in this test (sendImage must not touch OneDrive/invite): ${method} ${u}`);
    });
    const graph = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as unknown as typeof fetch });
    const cache = new MembersCache({ path: join(await mkdtemp(join(tmpdir(), 'teams-chats-sendimage-')), 'm.json') });
    const chats = new GraphTeamsChats(graph, { membersCache: cache });

    const sent = await chats.sendImage(CHAT, { bytes: new Uint8Array([1, 2, 3]), contentType: 'image/png' });

    expect(sent.id).toBe('img-1');
  });
});
