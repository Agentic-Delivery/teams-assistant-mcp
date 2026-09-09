import { dirname, join } from 'node:path';
import type { ChatAllowlist } from './allowlist.js';
import type { MembersCache } from './graph/members-cache.js';
import { InboxPoller, type SignedInAccount } from './inbox.js';
import type { TeamsChatsPort } from './graph/teams-chats.js';
import type { TokenProvider } from './auth/token-provider.js';

export interface BuildInboxPollerOptions {
  // 2026-09-09 (poll-path throttle fix): 'warmMembers' dropped from this Pick along with it —
  // roster warm-up no longer runs on the poll path at all, so this composition has no reason to
  // depend on it any more. 'resolveSelfIdStatus' (0.6.4) is not picked either: the poller resolves
  // self through the plain, never-throwing 'resolveSelfId' below and never distinguishes WHY a
  // resolution is missing — see the `self` wiring's own comment.
  chats: Pick<TeamsChatsPort, 'readMessages' | 'resolveSelfId'>;
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
  // warmupBackoffMs (0.6.3) dropped in the same fix that removed 'warmMembers' above — the poll
  // path no longer backs a chat's roster warm-up off, because it no longer warms it at all.
  /**
   * What the isSelf displayName fallback (inbox.ts) compares a message's `from` against when
   * Graph reports no `fromId` on it at all, OR when self-id resolution has not (yet, or ever)
   * produced an id — see config.assistantDisplayName's own doc comment. Sourced from config, not
   * a live call: unlike the id (below), this value is already known at startup, so passing it
   * through costs nothing and can never block a poll (0.6.4, issue #28, point 4 — the poller's old
   * raw `/me?$select=id,displayName` fetched this live; dropping the live half and sourcing it
   * from config instead is strictly cheaper and still best-effort). 2026-09-09 (poll-path throttle
   * fix): also what lets self-message filtering keep working by NAME even on a process whose
   * self-id resolution never succeeds at all (a persistent /me throttle with no seed or cache) —
   * see the `self` wiring's own comment below. Optional: a caller with no assistantDisplayName
   * configured just loses that one fallback, never the id-based check the loop otherwise relies on.
   */
  assistantDisplayName?: string;
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
    // 2026-09-09 (poll-path throttle fix): resolved through GraphTeamsChats.resolveSelfId — the
    // SAME operator-seed (TEAMS_MCP_SELF_ID) -> persisted-cache -> live `/me` chain sendFile
    // already uses, which never throws — instead of a raw `graph.get('/me?$select=id,displayName')`
    // of this module's own. That raw call is what used to fail the WHOLE poll cycle on a throttled
    // `/me`: see InboxPollerDeps.self's own doc comment (inbox.ts) for the incident. Deliberately
    // NOT `resolveSelfIdStatus` (0.6.4): that richer variant exists so a caller can distinguish a
    // THROTTLED miss from any other one, which is exactly the distinction the poll path (unlike
    // sendFile) has no use for any more — self-id resolution never fails or delays a cycle here
    // either way, so there is nothing for a caller to react to differently.
    //
    // `assistantDisplayName` (0.6.4, issue #28, point 4) rides along regardless of whether the id
    // itself resolved: it is known from config at startup, zero Graph cost, so it can never be the
    // reason a poll blocks. This is what lets isSelf's displayName fallback keep working even on a
    // process whose self-id resolution never succeeds at all (a persistent /me throttle with no
    // seed or cache) — strictly more robust than depending on the id alone.
    self: async (): Promise<SignedInAccount> => {
      const id = await options.chats.resolveSelfId?.();
      return {
        ...(id !== undefined ? { id } : {}),
        ...(options.assistantDisplayName !== undefined
          ? { displayName: options.assistantDisplayName }
          : {}),
      };
    },
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
