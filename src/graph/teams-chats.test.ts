import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { GraphClient, GraphError } from './graph-client.js';
import { MembersCache } from './members-cache.js';
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
        'retry-after': '62',
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
    expect((error as GraphError).retryAfterSeconds).toBe(62);
    expect((error as GraphError).message).not.toMatch(/retry after|\d+s/);
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
