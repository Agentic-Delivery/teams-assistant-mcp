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

  it('a 429 on the refresh surfaces as a distinct THROTTLED error naming Retry-After, not a silent fallback', async () => {
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
    expect((error as GraphError).retryAfterSeconds).toBe(62);
    expect((error as GraphError).message).toMatch(/retry after 62s/);
    expect(calls).toHaveLength(1);
  });

  it('without a members cache configured, resolveMentions falls back to the plain uncached call (back-compat)', async () => {
    const { fetchFn, calls } = countingMembersFetch(membersPage);
    const graph = new GraphClient({ tokenProvider: stubToken, fetchFn: fetchFn as unknown as typeof fetch });
    const chats = new GraphTeamsChats(graph);

    const resolved = await chats.resolveMentions(CHAT, ['Shiv']);

    expect(resolved).toEqual([{ name: 'Shiv', id: 'aad-shiv', displayName: 'Garg, Shivankit' }]);
    expect(calls).toHaveLength(1);
  });
});
