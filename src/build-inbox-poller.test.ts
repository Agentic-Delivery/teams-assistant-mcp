import { mkdtemp, rm, writeFile } from 'node:fs/promises';
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

      // Bootstrap poll (0.4.1: the first poll on a chat with no known watermark only settles it,
      // delivering nothing) — but harvest runs BEFORE that filter (inbox.ts's own doc comment), so
      // even this settling poll must have harvested Bob already.
      await poller.pollOnce();

      const freshCache = new MembersCache({ path: config.membersCachePath });
      expect(freshCache.get(CHAT)).toEqual(
        expect.arrayContaining([{ id: 'aad-bob', displayName: 'Bob Brown' }]),
      );
      // PARTIAL, not COMPLETE: this entry has only ever been touched by merge(), never a real
      // /members fetch — getComplete() must refuse it exactly as membersForInvite (sendFile's
      // permission-grant path) needs it to (0.5.2 BLOCKER 1 fix).
      expect(freshCache.getComplete(CHAT)).toBeUndefined();
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
