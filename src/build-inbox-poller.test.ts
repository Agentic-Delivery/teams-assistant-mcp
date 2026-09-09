import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const CHAT = '19:pilot@thread.v2';

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } });
}

/** A 429 with a Retry-After long enough that a retry inside this test's own process would sleep
 *  for real if anything foolishly retried it — same convention used throughout
 *  teams-chats.test.ts/graph-client.test.ts. Every route below is set up to prove a specific
 *  endpoint is NEVER called at all, so answering it (rather than throwing) would only mask a
 *  regression as a slow pass instead of a loud failure. */
function throttled(): Response {
  return new Response(
    JSON.stringify({ error: { code: 'TooManyRequests', message: 'Too many requests' } }),
    { status: 429, headers: { 'content-type': 'application/json', 'retry-after': '100' } },
  );
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

  async function configFor(dir: string) {
    const { loadConfig } = await import('./config.js');
    const configPath = join(dir, 'teams-mcp.config.json');
    await writeFile(configPath, JSON.stringify({ allowedChats: [{ id: CHAT, label: 'pilot', canPost: true }] }));
    return loadConfig({
      TEAMS_MCP_CONFIG: configPath,
      TEAMS_MCP_TENANT_ID: 'tenant',
      TEAMS_MCP_USERNAME: 'assistant@example.com',
      TEAMS_MCP_PASSWORD: 'secret',
      TEAMS_MCP_TOKEN_CACHE: join(dir, '.token-cache.json'),
    });
  }

  it('polling one message through the REAL stack (buildChats + buildInboxPoller) harvests the sender into the REAL on-disk MembersCache as a PARTIAL entry — and NEVER calls /members (2026-09-09, poll-path throttle fix: roster warm-up no longer runs on the poll path at all)', async () => {
    const { buildChats } = await import('./build-chats.js');
    const { buildInboxPoller } = await import('./build-inbox-poller.js');
    const { MembersCache } = await import('./graph/members-cache.js');

    const config = await configFor(dir);

    const originalFetch = globalThis.fetch;
    globalThis.fetch = (async (url: string) => {
      const u = String(url);
      if (u.includes('/me?')) {
        return json({ id: 'me-id' });
      }
      // Throttled rather than left unhandled: this test's point is that this route is never
      // reached by a poll cycle at all — an unhandled route throwing "unexpected call" would
      // prove the same thing less clearly than a throttle response the poll would visibly choke
      // on if it were still reached.
      if (u.includes(`/chats/${encodeURIComponent(CHAT)}/members`)) {
        return throttled();
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

      // Cycle 1 is the bootstrap/settling poll (0.4.1: delivers nothing, but harvest runs before
      // that filter). Cycle 2 proves the SAME holds on an ordinary cycle too — neither ever
      // touches /me or /members, which would 429 immediately above if they did.
      const clean1 = await poller.pollOnce();
      const clean2 = await poller.pollOnce();
      expect(clean1).toBe(true);
      expect(clean2).toBe(true);

      const freshCache = new MembersCache({ path: config.membersCachePath });
      expect(freshCache.get(CHAT)).toEqual(
        expect.arrayContaining([{ id: 'aad-bob', displayName: 'Bob Brown' }]),
      );
      // PARTIAL, not COMPLETE: this entry has only ever been touched by merge(), never a real
      // /members fetch — getComplete() must refuse it exactly as membersForInvite (sendFile's
      // permission-grant path) needs it to (0.5.2 BLOCKER 1 fix). Also proves the entry did not
      // somehow get upgraded by a poll-path warm-up, which no longer exists.
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

  // 2026-09-09 (poll-path throttle fix): self id resolves through GraphTeamsChats.resolveSelfId —
  // the operator-seed -> persisted-cache -> live /me chain sendFile already uses — instead of a
  // raw, unconditional /me call. This drives the REAL composition with a WARM persisted self-id
  // cache and proves /me is never called at all.
  it('self id resolves from the REAL persisted cache through the REAL stack, with zero /me calls', async () => {
    const { buildChats } = await import('./build-chats.js');
    const { buildInboxPoller } = await import('./build-inbox-poller.js');
    const { FileSelfIdCache } = await import('./graph/self-id-cache.js');

    const config = await configFor(dir);
    new FileSelfIdCache({ path: config.selfIdCachePath, expectedUsername: config.username }).write({
      id: 'me-id',
      resolvedAt: Date.now(),
    });

    const originalFetch = globalThis.fetch;
    globalThis.fetch = (async (url: string) => {
      const u = String(url);
      if (u.includes('/me?')) {
        throw new Error('unexpected /me call: the persisted self-id cache should have answered');
      }
      if (u.includes(`/chats/${encodeURIComponent(CHAT)}/members`)) {
        return throttled();
      }
      if (u.includes(`/chats/${encodeURIComponent(CHAT)}/messages`)) {
        return json({
          value: [
            {
              id: 'msg-self-1',
              chatId: CHAT,
              createdDateTime: '2026-09-04T10:00:00Z',
              from: { user: { id: 'me-id', displayName: 'Assistant (AI)' } },
              body: { contentType: 'text', content: 'my own post' },
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

      await poller.pollOnce(); // settle
      const clean = await poller.pollOnce();
      expect(clean).toBe(true);

      // The self-filtered message must not have been delivered — proof the persisted id actually
      // reached isSelf, not just that resolution didn't throw.
      const inboxRaw = await readFile(join(dir, 'inbox.jsonl'), 'utf8').catch(() => '');
      expect(inboxRaw.trim()).toBe('');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  // The direct evidence for the incident this whole fix closes (docs/throttling-mitigation.md's
  // dated section, 2026-09-09): /me AND /members both 429 throughout, with no seed and no
  // persisted cache — the poll must still read and deliver every allowed chat's messages.
  it('/me and /members throttled on EVERY call, no seed or persisted cache — messages still get delivered and the health file reports ok', async () => {
    const { buildChats } = await import('./build-chats.js');
    const { buildInboxPoller } = await import('./build-inbox-poller.js');

    const config = await configFor(dir);

    let meCalls = 0;
    let membersCalls = 0;
    let messagesCalls = 0;
    const originalFetch = globalThis.fetch;
    globalThis.fetch = (async (url: string) => {
      const u = String(url);
      if (u.includes('/me?')) {
        meCalls += 1;
        return throttled();
      }
      if (u.includes(`/chats/${encodeURIComponent(CHAT)}/members`)) {
        membersCalls += 1;
        return throttled();
      }
      if (u.includes(`/chats/${encodeURIComponent(CHAT)}/messages`)) {
        messagesCalls += 1;
        // Empty on the FIRST call (the bootstrap/settling poll — 0.4.1: establishes the
        // watermark, delivers nothing) so the second call's message is genuinely NEW against
        // that watermark, rather than the same static message being filtered out as stale.
        if (messagesCalls === 1) {
          return json({ value: [] });
        }
        return json({
          value: [
            {
              id: 'msg-1',
              chatId: CHAT,
              createdDateTime: '2026-09-09T10:00:00Z',
              from: { user: { id: 'aad-dana', displayName: 'Dana Duffy' } },
              body: { contentType: 'text', content: 'are you there?' },
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

      await poller.pollOnce(); // bootstrap/settling poll — establishes the watermark, delivers nothing
      const clean = await poller.pollOnce();

      expect(clean).toBe(true);
      const lines = (await readFile(join(dir, 'inbox.jsonl'), 'utf8')).trim().split('\n');
      expect(lines).toHaveLength(1);
      expect(JSON.parse(lines[0] as string)).toMatchObject({ chat: CHAT, id: 'msg-1', from: 'Dana Duffy' });

      const health = JSON.parse(
        await readFile(join(dir, 'poller-health.json'), 'utf8'),
      ) as Record<string, unknown>;
      expect(health['ok']).toBe(true);

      // /me is retried at most once per process (see InboxPoller's self-resolution doc comment) —
      // NOT once per poll cycle, which would keep feeding the same throttled budget.
      expect(meCalls).toBe(1);
      // /members is never called at all — roster warm-up no longer runs on the poll path.
      expect(membersCalls).toBe(0);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
