import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const CHAT = '19:pilot@thread.v2';

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
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
      if (u.includes('/me?') && u.includes('select=id,displayName')) {
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
      const { chats, graph, tokenProvider, membersCache } = buildChats(config);
      vi.spyOn(tokenProvider, 'getAccessToken').mockResolvedValue('fake-token');

      const poller = buildInboxPoller({
        chats,
        graph,
        tokenProvider,
        membersCache,
        allowlist: config.allowlist,
        inboxPath: join(dir, 'inbox.jsonl'),
      });

      // Cycle 1: behaviour-4 warm-up (0.6.0) hits the throttled /members mock ONCE — readRetries:0
      // means no sleep, but review round 1 MAJOR 4 also made a throttled warm-up end the CYCLE
      // (same "one 429 ends the cycle" rule readMessages's own 429 handling follows), so this
      // cycle never reaches readMessages/harvest at all. warmedChats already records CHAT as
      // attempted, so cycle 2 skips warm-up and proceeds straight to the bootstrap poll (0.4.1:
      // the first poll on a chat with no known watermark only settles it, delivering nothing —
      // but harvest runs BEFORE that filter, so even this settling poll harvests Bob).
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
      if (u.includes('/me?') && u.includes('select=id,displayName')) {
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
      const { chats, graph, tokenProvider, membersCache } = buildChats(config);
      vi.spyOn(tokenProvider, 'getAccessToken').mockResolvedValue('fake-token');

      const poller = buildInboxPoller({
        chats,
        graph,
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
});
