import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const CHAT = '19:pilot@thread.v2';

function json(body: unknown, status = 200, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json', ...headers },
  });
}

// MAJOR 3 (2026-09-04 review): deleting `roster: membersCache` (previously an inline literal in
// index.ts) left the full 485-test suite green — every inbox.test.ts test drives a hand-built
// `rosterSink` double, never the real InboxPoller against a real MembersCache, and index.ts itself
// is not unit-testable (a top-level async main()). Fixed by extracting the wire into
// build-inbox-poller.ts's `buildInboxPoller` — index.ts and THIS test are now the only two
// callers, so a dropped `roster:` line there can only pass by ALSO being dropped here, and this
// test would then fail (see the file's own doc comment for the removal proof, done once and
// reverted, referenced in the commit message).
describe('buildInboxPoller — the real composition, driven end to end (MAJOR 3, 2026-09-04 review)', () => {
  let dir: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'build-inbox-poller-'));
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  it('polling one message through the REAL stack (buildChats + buildInboxPoller) harvests the sender into the REAL on-disk MembersCache as a PARTIAL entry', async () => {
    const { loadConfig } = await import('./config.js');
    const { buildChats } = await import('./build-chats.js');
    const { buildInboxPoller } = await import('./build-inbox-poller.js');
    const { MembersCache } = await import('./graph/members-cache.js');

    const configPath = join(dir, 'teams-mcp.config.json');
    await writeFile(configPath, JSON.stringify({ allowedChats: [{ id: CHAT, label: 'pilot', canPost: true }] }));
    const config = loadConfig({
      TEAMS_MCP_CONFIG: configPath,
      TEAMS_MCP_TENANT_ID: 'tenant',
      TEAMS_MCP_USERNAME: 'assistant@example.com',
      TEAMS_MCP_PASSWORD: 'secret',
      TEAMS_MCP_TOKEN_CACHE: join(dir, '.token-cache.json'),
    });

    const originalFetch = globalThis.fetch;
    globalThis.fetch = (async (url: string) => {
      const u = String(url);
      if (u.includes('/me?') && u.includes('select=id')) {
        return json({ id: 'me-id', displayName: 'Assistant (AI)' });
      }
      // 0.6.0's behaviour-4 warm-up (inbox.ts) now calls warmMembers on the FIRST poll of every
      // allowlisted chat — modelled here as throttled (a realistic "the endpoint is unreachable
      // this cycle" shape, matching this test's own harvest-only scenario) rather than left
      // unhandled, so the roster genuinely stays PARTIAL for the reason the assertion below
      // names, not because of an unrelated mock gap. See the SECOND test in this file for the
      // warm-up-SUCCEEDS case through this same real composition.
      if (u.includes(`/chats/${encodeURIComponent(CHAT)}/members`)) {
        // retry-after past GraphClient's MAX_RETRY_SLEEP_MS (90s) so every retry fails fast
        // locally rather than actually sleeping in this test (same convention used throughout
        // teams-chats.test.ts/graph-client.test.ts).
        return new Response(
          JSON.stringify({ error: { code: 'TooManyRequests', message: 'Too many requests' } }),
          { status: 429, headers: { 'content-type': 'application/json', 'retry-after': '100' } },
        );
      }
      if (u.includes(`/chats/${encodeURIComponent(CHAT)}/messages`)) {
        return json({
          value: [
            {
              id: 'msg-harvest-1',
              chatId: CHAT,
              createdDateTime: '2026-09-04T10:00:00Z',
              from: { user: { id: 'aad-bob', displayName: 'Bob Brown' } },
              body: { contentType: 'text', content: 'hello from Bob' },
            },
          ],
        });
      }
      throw new Error(`unexpected call in this test: ${u}`);
    }) as typeof fetch;
    try {
      const { chats, tokenProvider, membersCache } = buildChats(config);
      vi.spyOn(tokenProvider, 'getAccessToken').mockResolvedValue('fake-token');

      const poller = buildInboxPoller({
        chats,
        tokenProvider,
        membersCache,
        allowlist: config.allowlist,
        inboxPath: join(dir, 'inbox.jsonl'),
      });

      // Cycle 1: behaviour-4 warm-up (0.6.0) hits the throttled /members mock ONCE — readRetries:0
      // means no sleep, but review round 1 MAJOR 4 also made a throttled warm-up end the CYCLE
      // (same "one 429 ends the cycle" rule readMessages's own 429 handling follows), so this
      // cycle never reaches readMessages/harvest at all. Cycle 2 runs back to back with no real
      // wait, so it is still well inside this chat's per-chat warm-up back-off window (0.6.3,
      // issue a: 15 minutes by default) opened by cycle 1's throttled attempt — the warm-up is
      // skipped again and the cycle proceeds straight to the bootstrap poll (0.4.1: the first poll
      // on a chat with no known watermark only settles it, delivering nothing — but harvest runs
      // BEFORE that filter, so even this settling poll harvests Bob).
      await poller.pollOnce();
      await poller.pollOnce();

      const freshCache = new MembersCache({ path: config.membersCachePath });
      expect(freshCache.get(CHAT)).toEqual(
        expect.arrayContaining([{ id: 'aad-bob', displayName: 'Bob Brown' }]),
      );
      // PARTIAL, not COMPLETE: this entry has only ever been touched by merge(), never a real
      // /members fetch — getComplete() must refuse it exactly as membersForInvite (sendFile's
      // permission-grant path) needs it to (0.5.2 BLOCKER 1 fix).
      expect(freshCache.getComplete(CHAT)).toBeUndefined();

      // Carried over from PR #8 (0.5.4): the health file is written by InboxPoller itself, not
      // by a wire buildInboxPoller could drop — but it is still worth proving through the REAL
      // composition, same reasoning as the roster assertion above: a hand-built double in
      // inbox.test.ts proves the class writes it, not that the real wiring this process actually
      // uses produces a file a watcher could read.
      const health = JSON.parse(
        await readFile(join(dir, 'poller-health.json'), 'utf8'),
      ) as Record<string, unknown>;
      expect(health['ok']).toBe(true);
      expect(health['inboxPath']).toBe(join(dir, 'inbox.jsonl'));
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  // Behaviour 4 (0.6.0, live 2026-09-08) — the wire proof for warmMembers, same shape as MAJOR
  // 3's `roster: membersCache` proof above: `BuildInboxPollerOptions.chats` is typed
  // `Pick<TeamsChatsPort, 'readMessages' | 'warmMembers'>`, so a dropped `warmMembers` forward
  // anywhere in the real chain (ReliableTeamsChats → GraphTeamsChats) would leave this the only
  // test able to catch it — inbox.test.ts's own warm-up tests all drive hand-built doubles.
  it('the FIRST poll of a newly-allowlisted chat warms its roster to COMPLETE through the REAL stack, when /members is reachable', async () => {
    const { loadConfig } = await import('./config.js');
    const { buildChats } = await import('./build-chats.js');
    const { buildInboxPoller } = await import('./build-inbox-poller.js');
    const { MembersCache } = await import('./graph/members-cache.js');

    const configPath = join(dir, 'teams-mcp.config.json');
    await writeFile(configPath, JSON.stringify({ allowedChats: [{ id: CHAT, label: 'pilot', canPost: true }] }));
    const config = loadConfig({
      TEAMS_MCP_CONFIG: configPath,
      TEAMS_MCP_TENANT_ID: 'tenant',
      TEAMS_MCP_USERNAME: 'assistant@example.com',
      TEAMS_MCP_PASSWORD: 'secret',
      TEAMS_MCP_TOKEN_CACHE: join(dir, '.token-cache.json'),
    });

    const originalFetch = globalThis.fetch;
    let membersCalls = 0;
    globalThis.fetch = (async (url: string) => {
      const u = String(url);
      if (u.includes('/me?') && u.includes('select=id')) {
        return json({ id: 'me-id', displayName: 'Assistant (AI)' });
      }
      if (u.includes(`/chats/${encodeURIComponent(CHAT)}/members`)) {
        membersCalls += 1;
        return json({ value: [{ userId: 'aad-carol', displayName: 'Carol Chen' }] });
      }
      if (u.includes(`/chats/${encodeURIComponent(CHAT)}/messages`)) {
        return json({ value: [] }); // no traffic yet — warm-up is the ONLY source of this roster
      }
      throw new Error(`unexpected call in this test: ${u}`);
    }) as typeof fetch;
    try {
      const { chats, tokenProvider, membersCache } = buildChats(config);
      vi.spyOn(tokenProvider, 'getAccessToken').mockResolvedValue('fake-token');

      const poller = buildInboxPoller({
        chats,
        tokenProvider,
        membersCache,
        allowlist: config.allowlist,
        inboxPath: join(dir, 'inbox.jsonl'),
      });

      await poller.pollOnce();

      const freshCache = new MembersCache({ path: config.membersCachePath });
      // COMPLETE, not just present: a real /members fetch backs this entry — sendFile's
      // permission grant (getComplete/getStaleComplete) can use it with no further live call.
      expect(freshCache.getComplete(CHAT)).toEqual([{ id: 'aad-carol', displayName: 'Carol Chen' }]);
      expect(membersCalls).toBe(1);

      await poller.pollOnce();
      expect(membersCalls).toBe(1); // still once — not once per poll cycle
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  // MAJOR 3 (review round 1, fresh-context re-review): the TEAMS_INBOX_WARMUP_BACKOFF_SECONDS ->
  // warmupBackoffMs wire (index.ts -> buildInboxPoller -> InboxPoller) had no test — deleting the
  // `...(options.warmupBackoffMs !== undefined ? { warmupBackoffMs: options.warmupBackoffMs } : {})`
  // line in build-inbox-poller.ts left the full suite green. This drives the REAL composition with
  // an explicit warmupBackoffMs and a controlled system clock (vi.setSystemTime), proving the
  // configured window — not just the 15-minute default — actually governs the real InboxPoller a
  // real buildInboxPoller call produces.
  it('warmupBackoffMs passed into buildInboxPoller actually governs the real InboxPoller: no re-warm before the configured window, one after it', async () => {
    const { loadConfig } = await import('./config.js');
    const { buildChats } = await import('./build-chats.js');
    const { buildInboxPoller } = await import('./build-inbox-poller.js');

    const configPath = join(dir, 'teams-mcp.config.json');
    await writeFile(configPath, JSON.stringify({ allowedChats: [{ id: CHAT, label: 'pilot', canPost: true }] }));
    const config = loadConfig({
      TEAMS_MCP_CONFIG: configPath,
      TEAMS_MCP_TENANT_ID: 'tenant',
      TEAMS_MCP_USERNAME: 'assistant@example.com',
      TEAMS_MCP_PASSWORD: 'secret',
      TEAMS_MCP_TOKEN_CACHE: join(dir, '.token-cache.json'),
    });

    const originalFetch = globalThis.fetch;
    let membersCalls = 0;
    globalThis.fetch = (async (url: string) => {
      const u = String(url);
      if (u.includes('/me?') && u.includes('select=id')) {
        return json({ id: 'me-id', displayName: 'Assistant (AI)' });
      }
      if (u.includes(`/chats/${encodeURIComponent(CHAT)}/members`)) {
        membersCalls += 1;
        // No retry-after header at all: GraphError.retryAfterSeconds is undefined, so the
        // configured warmupBackoffMs (not Graph's own wait) is what this test is actually
        // proving — see warmupBackoffFor's floor-vs-override doc comment (MAJOR 1) for why that
        // distinction matters.
        return new Response(
          JSON.stringify({ error: { code: 'TooManyRequests', message: 'Too many requests' } }),
          { status: 429, headers: { 'content-type': 'application/json' } },
        );
      }
      if (u.includes(`/chats/${encodeURIComponent(CHAT)}/messages`)) {
        return json({ value: [] });
      }
      throw new Error(`unexpected call in this test: ${u}`);
    }) as typeof fetch;

    vi.useFakeTimers({ toFake: ['Date'] });
    vi.setSystemTime(new Date('2026-09-08T14:42:00.000Z'));
    try {
      const { chats, tokenProvider, membersCache } = buildChats(config);
      vi.spyOn(tokenProvider, 'getAccessToken').mockResolvedValue('fake-token');

      const poller = buildInboxPoller({
        chats,
        tokenProvider,
        membersCache,
        allowlist: config.allowlist,
        inboxPath: join(dir, 'inbox.jsonl'),
        warmupBackoffMs: 60_000, // 60s, not the 15-minute default
      });

      await poller.pollOnce(); // opens the 60s window
      expect(membersCalls).toBe(1);

      vi.setSystemTime(new Date('2026-09-08T14:42:59.000Z')); // 59s later — still inside the window
      await poller.pollOnce();
      expect(membersCalls).toBe(1); // NOT re-warmed yet

      vi.setSystemTime(new Date('2026-09-08T14:43:01.000Z')); // 61s later — past the configured window
      await poller.pollOnce();
      expect(membersCalls).toBe(2); // re-warmed once the CONFIGURED window elapsed
    } finally {
      globalThis.fetch = originalFetch;
      vi.useRealTimers();
    }
  });
});

// Issue #28 (live 2026-09-09): before 0.6.4 the poller's `self` dependency called a raw, unwired
// `/me?$select=id,displayName` on the graph client — bypassing TEAMS_MCP_SELF_ID and the persisted
// self-id cache entirely. Under a tenant-wide `/me` 429 with neither available, every poll failed
// and, after three cycles, forced a token re-authentication for a condition that was never
// auth-shaped. These tests drive the REAL composition (buildChats + buildInboxPoller), mirroring
// the sendFile-side proofs in teams-chats.test.ts's "buildChats — the composition wires the self
// id cache" describe block, but for the poller's own delivery/filtering path.
describe('buildInboxPoller — self id resolution reuses the seed/cache/live chain (0.6.4, issue #28)', () => {
  let dir: string;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'build-inbox-poller-selfid-'));
  });

  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  /** A `/messages` responder that delivers nothing on its first call (the bootstrap/settling poll
   *  — see inbox.ts's own isBootstrap comment) and the given messages on every call after that. */
  function messagesAfterSettle(messages: unknown[]) {
    let calls = 0;
    return () => {
      calls += 1;
      return json({ value: calls === 1 ? [] : messages });
    };
  }

  it('TEAMS_MCP_SELF_ID reaches the real poller end to end: a 429ing /me is never even called, and the seeded id filters the assistant\'s own post', async () => {
    const { loadConfig } = await import('./config.js');
    const { buildChats } = await import('./build-chats.js');
    const { buildInboxPoller } = await import('./build-inbox-poller.js');

    const SELF_ID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
    const configPath = join(dir, 'teams-mcp.config.json');
    await writeFile(configPath, JSON.stringify({ allowedChats: [{ id: CHAT, label: 'pilot', canPost: true }] }));
    const config = loadConfig({
      TEAMS_MCP_CONFIG: configPath,
      TEAMS_MCP_TENANT_ID: 'tenant',
      TEAMS_MCP_USERNAME: 'assistant@example.com',
      TEAMS_MCP_PASSWORD: 'secret',
      TEAMS_MCP_TOKEN_CACHE: join(dir, '.token-cache.json'),
      TEAMS_MCP_SELF_ID: SELF_ID,
    });
    expect(config.selfIdOverride).toBe(SELF_ID);

    const respondMessages = messagesAfterSettle([
      {
        id: 'msg-own',
        chatId: CHAT,
        createdDateTime: '2026-09-09T10:00:00Z',
        from: { user: { id: SELF_ID, displayName: 'Assistant (AI)' } },
        body: { contentType: 'text', content: 'my own earlier post' },
      },
      {
        id: 'msg-bob',
        chatId: CHAT,
        createdDateTime: '2026-09-09T10:01:00Z',
        from: { user: { id: 'aad-bob', displayName: 'Bob Brown' } },
        body: { contentType: 'text', content: 'hello from Bob' },
      },
    ]);

    const originalFetch = globalThis.fetch;
    globalThis.fetch = (async (url: string) => {
      const u = String(url);
      if (u.includes('/me?') && u.includes('select=id')) {
        throw new Error('must never call /me — TEAMS_MCP_SELF_ID must win outright');
      }
      if (u.includes(`/chats/${encodeURIComponent(CHAT)}/members`)) {
        return json({ value: [] }); // warm-up succeeds trivially — not this test's concern
      }
      if (u.includes(`/chats/${encodeURIComponent(CHAT)}/messages`)) {
        return respondMessages();
      }
      throw new Error(`unexpected call in this test: ${u}`);
    }) as typeof fetch;
    try {
      const { chats, tokenProvider, membersCache } = buildChats(config);
      vi.spyOn(tokenProvider, 'getAccessToken').mockResolvedValue('fake-token');

      const poller = buildInboxPoller({
        chats,
        tokenProvider,
        membersCache,
        allowlist: config.allowlist,
        inboxPath: join(dir, 'inbox.jsonl'),
      });

      await poller.pollOnce(); // settles the watermark, delivers nothing
      const clean = await poller.pollOnce(); // the two messages above are now "new"

      expect(clean).toBe(true);
      const inboxRaw = await readFile(join(dir, 'inbox.jsonl'), 'utf8');
      const delivered = inboxRaw
        .split('\n')
        .filter((line) => line.trim() !== '')
        .map((line) => JSON.parse(line) as Record<string, unknown>);
      expect(delivered).toHaveLength(1);
      expect(delivered[0]?.['from']).toBe('Bob Brown'); // the seeded id's OWN post was filtered
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('a persisted self-id cache (no operator seed) reaches the real poller end to end: a 429ing /me is never called, same filtering', async () => {
    const { loadConfig } = await import('./config.js');
    const { buildChats } = await import('./build-chats.js');
    const { buildInboxPoller } = await import('./build-inbox-poller.js');
    const { FileSelfIdCache } = await import('./graph/self-id-cache.js');

    const configPath = join(dir, 'teams-mcp.config.json');
    await writeFile(configPath, JSON.stringify({ allowedChats: [{ id: CHAT, label: 'pilot', canPost: true }] }));
    const config = loadConfig({
      TEAMS_MCP_CONFIG: configPath,
      TEAMS_MCP_TENANT_ID: 'tenant',
      TEAMS_MCP_USERNAME: 'assistant@example.com',
      TEAMS_MCP_PASSWORD: 'secret',
      TEAMS_MCP_TOKEN_CACHE: join(dir, '.token-cache.json'),
    });
    expect(config.selfIdOverride).toBeUndefined();
    new FileSelfIdCache({ path: config.selfIdCachePath, expectedUsername: config.username }).write({
      id: 'aad-cached-self',
      resolvedAt: 1,
    });

    const respondMessages = messagesAfterSettle([
      {
        id: 'msg-own',
        chatId: CHAT,
        createdDateTime: '2026-09-09T10:00:00Z',
        from: { user: { id: 'aad-cached-self', displayName: 'Assistant (AI)' } },
        body: { contentType: 'text', content: 'my own earlier post' },
      },
      {
        id: 'msg-bob',
        chatId: CHAT,
        createdDateTime: '2026-09-09T10:01:00Z',
        from: { user: { id: 'aad-bob', displayName: 'Bob Brown' } },
        body: { contentType: 'text', content: 'hello from Bob' },
      },
    ]);

    const originalFetch = globalThis.fetch;
    globalThis.fetch = (async (url: string) => {
      const u = String(url);
      if (u.includes('/me?') && u.includes('select=id')) {
        throw new Error('must never call /me — the persisted self-id cache is warm');
      }
      if (u.includes(`/chats/${encodeURIComponent(CHAT)}/members`)) {
        return json({ value: [] });
      }
      if (u.includes(`/chats/${encodeURIComponent(CHAT)}/messages`)) {
        return respondMessages();
      }
      throw new Error(`unexpected call in this test: ${u}`);
    }) as typeof fetch;
    try {
      const { chats, tokenProvider, membersCache } = buildChats(config);
      vi.spyOn(tokenProvider, 'getAccessToken').mockResolvedValue('fake-token');

      const poller = buildInboxPoller({
        chats,
        tokenProvider,
        membersCache,
        allowlist: config.allowlist,
        inboxPath: join(dir, 'inbox.jsonl'),
      });

      await poller.pollOnce();
      const clean = await poller.pollOnce();

      expect(clean).toBe(true);
      const inboxRaw = await readFile(join(dir, 'inbox.jsonl'), 'utf8');
      const delivered = inboxRaw
        .split('\n')
        .filter((line) => line.trim() !== '')
        .map((line) => JSON.parse(line) as Record<string, unknown>);
      expect(delivered).toHaveLength(1);
      expect(delivered[0]?.['from']).toBe('Bob Brown');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('no operator seed and no persisted cache: a persistently-THROTTLED /me delivers nothing and never forces a token re-authentication through the REAL stack', async () => {
    const { loadConfig } = await import('./config.js');
    const { buildChats } = await import('./build-chats.js');
    const { buildInboxPoller } = await import('./build-inbox-poller.js');
    const { FileSelfIdCache } = await import('./graph/self-id-cache.js');

    const configPath = join(dir, 'teams-mcp.config.json');
    await writeFile(configPath, JSON.stringify({ allowedChats: [{ id: CHAT, label: 'pilot', canPost: true }] }));
    const config = loadConfig({
      TEAMS_MCP_CONFIG: configPath,
      TEAMS_MCP_TENANT_ID: 'tenant',
      TEAMS_MCP_USERNAME: 'assistant@example.com',
      TEAMS_MCP_PASSWORD: 'secret',
      TEAMS_MCP_TOKEN_CACHE: join(dir, '.token-cache.json'),
    });
    expect(config.selfIdOverride).toBeUndefined();
    expect(new FileSelfIdCache({ path: config.selfIdCachePath }).read()).toBeUndefined(); // cold

    let meCalls = 0;
    let messagesCalls = 0;
    const originalFetch = globalThis.fetch;
    globalThis.fetch = (async (url: string) => {
      const u = String(url);
      if (u.includes('/me?') && u.includes('select=id')) {
        meCalls += 1;
        return json({ error: { code: 'TooManyRequests', message: 'Too many requests' } }, 429, {
          'retry-after': '5',
        });
      }
      if (u.includes(`/chats/${encodeURIComponent(CHAT)}/messages`)) {
        messagesCalls += 1;
        return json({
          value: [
            {
              id: 'msg-bob',
              chatId: CHAT,
              createdDateTime: '2026-09-09T10:01:00Z',
              from: { user: { id: 'aad-bob', displayName: 'Bob Brown' } },
              body: { contentType: 'text', content: 'hello from Bob' },
            },
          ],
        });
      }
      // /members is never expected to be reached: with no known self id the whole cycle fails
      // before the per-chat loop (warm-up included) ever runs.
      throw new Error(`unexpected call in this test: ${u}`);
    }) as typeof fetch;
    try {
      const { chats, tokenProvider, membersCache } = buildChats(config);
      vi.spyOn(tokenProvider, 'getAccessToken').mockResolvedValue('fake-token');
      const invalidateSpy = vi.spyOn(tokenProvider, 'invalidate');

      const poller = buildInboxPoller({
        chats,
        tokenProvider,
        membersCache,
        allowlist: config.allowlist,
        inboxPath: join(dir, 'inbox.jsonl'),
      });

      for (let i = 0; i < 6; i += 1) {
        // eslint-disable-next-line no-await-in-loop
        const clean = await poller.pollOnce();
        expect(clean).toBe(false);
      }

      expect(messagesCalls).toBe(0); // no known self id — the chat is never even asked
      // Exactly ONE real network call: GraphClient's own local throttle gate (a live 429 closes
      // the `/me` family for its Retry-After window) answers the next 5 attempts with its own
      // LocallyThrottled 429 — still `status: 429`, so it is classified identically. Either way,
      // /me is never memoized on a throttled attempt (retried every cycle it stays unresolved).
      expect(meCalls).toBe(1);
      expect(invalidateSpy).not.toHaveBeenCalled(); // the forced re-auth path never trips

      const inboxRaw = await readFile(join(dir, 'inbox.jsonl'), 'utf8');
      const errorLines = inboxRaw
        .split('\n')
        .filter((line) => line.trim() !== '')
        .map((line) => JSON.parse(line) as Record<string, unknown>);
      expect(errorLines.every((line) => !('from' in line))).toBe(true); // never a delivered message
      expect(String(errorLines[0]?.['error'])).toMatch(/self id resolution/i);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  // Point 4 of issue #28: the old raw `/me?$select=id,displayName` call also fetched displayName
  // for isSelf's fallback (a message reporting no fromId at all). That live fetch is dropped;
  // `assistantDisplayName` — already known from config, unrelated to `/me` — is threaded through
  // instead. This is the wire proof for that: BuildInboxPollerOptions.assistantDisplayName ->
  // InboxPollerDeps.self()'s returned `displayName` -> isSelf's fallback match, driven through the
  // REAL composition so a dropped wire here (not just an inline literal) would fail this test,
  // same shape as the `roster: membersCache`/`warmupBackoffMs` wire proofs above.
  it('assistantDisplayName reaches the real InboxPoller and is used by isSelf\'s no-fromId fallback', async () => {
    const { loadConfig } = await import('./config.js');
    const { buildChats } = await import('./build-chats.js');
    const { buildInboxPoller } = await import('./build-inbox-poller.js');

    const SELF_ID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';
    const configPath = join(dir, 'teams-mcp.config.json');
    await writeFile(configPath, JSON.stringify({ allowedChats: [{ id: CHAT, label: 'pilot', canPost: true }] }));
    const config = loadConfig({
      TEAMS_MCP_CONFIG: configPath,
      TEAMS_MCP_TENANT_ID: 'tenant',
      TEAMS_MCP_USERNAME: 'assistant@example.com',
      TEAMS_MCP_PASSWORD: 'secret',
      TEAMS_MCP_TOKEN_CACHE: join(dir, '.token-cache.json'),
      TEAMS_MCP_SELF_ID: SELF_ID,
    });

    // Graph itself reports no `user.id` for this message (an id-less sender shape Graph can
    // produce) but the SAME displayName the assistant is configured under — only the displayName
    // fallback, not the id check, can catch this one.
    const respondMessages = messagesAfterSettle([
      {
        id: 'msg-own-no-id',
        chatId: CHAT,
        createdDateTime: '2026-09-09T10:00:00Z',
        from: { user: { displayName: config.assistantDisplayName } },
        body: { contentType: 'text', content: 'my own post, no fromId reported' },
      },
      {
        id: 'msg-bob',
        chatId: CHAT,
        createdDateTime: '2026-09-09T10:01:00Z',
        from: { user: { id: 'aad-bob', displayName: 'Bob Brown' } },
        body: { contentType: 'text', content: 'hello from Bob' },
      },
    ]);

    const originalFetch = globalThis.fetch;
    globalThis.fetch = (async (url: string) => {
      const u = String(url);
      if (u.includes('/me?') && u.includes('select=id')) {
        throw new Error('must never call /me — TEAMS_MCP_SELF_ID must win outright');
      }
      if (u.includes(`/chats/${encodeURIComponent(CHAT)}/members`)) {
        return json({ value: [] });
      }
      if (u.includes(`/chats/${encodeURIComponent(CHAT)}/messages`)) {
        return respondMessages();
      }
      throw new Error(`unexpected call in this test: ${u}`);
    }) as typeof fetch;
    try {
      const { chats, tokenProvider, membersCache } = buildChats(config);
      vi.spyOn(tokenProvider, 'getAccessToken').mockResolvedValue('fake-token');

      const poller = buildInboxPoller({
        chats,
        tokenProvider,
        membersCache,
        allowlist: config.allowlist,
        inboxPath: join(dir, 'inbox.jsonl'),
        assistantDisplayName: config.assistantDisplayName,
      });

      await poller.pollOnce(); // settles the watermark
      await poller.pollOnce(); // the two messages above are now "new"

      const inboxRaw = await readFile(join(dir, 'inbox.jsonl'), 'utf8');
      const delivered = inboxRaw
        .split('\n')
        .filter((line) => line.trim() !== '')
        .map((line) => JSON.parse(line) as Record<string, unknown>);
      expect(delivered).toHaveLength(1);
      expect(delivered[0]?.['from']).toBe('Bob Brown'); // the id-less "own" post was still filtered
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
