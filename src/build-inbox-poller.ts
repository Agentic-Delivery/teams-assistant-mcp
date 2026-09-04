import { dirname, join } from 'node:path';
import type { ChatAllowlist } from './allowlist.js';
import type { GraphClient } from './graph/graph-client.js';
import type { MembersCache } from './graph/members-cache.js';
import { InboxPoller, type SignedInAccount } from './inbox.js';
import type { TeamsChatsPort } from './graph/teams-chats.js';
import type { TokenProvider } from './auth/token-provider.js';

export interface BuildInboxPollerOptions {
  chats: Pick<TeamsChatsPort, 'readMessages'>;
  /** Only `/me` is used here (to learn who "self" is) — the raw client, not the chats wrapper. */
  graph: GraphClient;
  tokenProvider: TokenProvider;
  /**
   * The SAME MembersCache instance `buildChats` wires into `GraphTeamsChats` — deliberately the
   * concrete class, not `RosterHarvestPort` (inbox.ts): this module's whole reason to exist is to
   * be the ONE place that wires `roster:` at all, so the wire itself is reachable by a composition
   * test (see build-inbox-poller.test.ts) instead of living only as an inline literal in index.ts
   * (MAJOR 3, 2026-09-04 review: deleting `roster: membersCache` there left the full suite green,
   * because every inbox.ts test drives a hand-built roster double, never the real MembersCache).
   */
  membersCache: MembersCache;
  allowlist: ChatAllowlist;
  inboxPath: string;
  inboxYieldPath?: string;
  pollMs?: number;
  log?: (line: string) => void;
}

/**
 * The one composition of the background inbox poller, mirroring build-chats.ts's own role for the
 * send stack: index.ts (the server entrypoint) and this module's own test are the only two callers,
 * so a dropped `roster:` wire here is structurally impossible to miss — both read from the exact
 * same function. See BuildInboxPollerOptions.membersCache's own doc comment for the incident this
 * extraction closes (MAJOR 3, 2026-09-04 review).
 */
export function buildInboxPoller(options: BuildInboxPollerOptions): InboxPoller {
  return new InboxPoller({
    chats: options.chats,
    allowlist: options.allowlist,
    self: () => options.graph.get<SignedInAccount>('/me?$select=id,displayName'),
    inboxPath: options.inboxPath,
    // The state sidecar follows the inbox file, so a TEAMS_INBOX_PATH override moves both — and
    // the yield file with them (inboxYieldPathFor derives from the same inbox path, index.ts).
    statePath: join(dirname(options.inboxPath), 'inbox-state.json'),
    ...(options.inboxYieldPath !== undefined ? { yieldPath: options.inboxYieldPath } : {}),
    // Mitigation 2 (docs/throttling-mitigation.md §4, stage 1 item 2): the SAME MembersCache
    // instance GraphTeamsChats resolves mentions against — every message this poller reads
    // harvests its sender into it, at zero Graph cost, as a PARTIAL entry (MembersCache.merge's
    // own doc comment) — sendFile's permission grant never trusts this without a live /members
    // re-check; see teams-chats.ts's membersForInvite for the COMPLETE-only read that enforces it.
    roster: options.membersCache,
    ...(options.pollMs !== undefined ? { pollMs: options.pollMs } : {}),
    log: options.log ?? (() => {}),
    // 0.4.1 stuck-auth self-healing: a token that goes bad without the local cache's own expiry
    // catching up needs an external nudge to drop it — see InboxPoller.trackAuthHealth's doc
    // comment for the live diagnosis this closes.
    onAuthStuck: () => options.tokenProvider.invalidate?.(),
  });
}
