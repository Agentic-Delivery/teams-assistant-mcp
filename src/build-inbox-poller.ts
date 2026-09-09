import { dirname, join } from 'node:path';
import type { ChatAllowlist } from './allowlist.js';
import type { MembersCache } from './graph/members-cache.js';
import { InboxPoller, type SignedInAccount } from './inbox.js';
import type { TeamsChatsPort } from './graph/teams-chats.js';
import type { TokenProvider } from './auth/token-provider.js';

export interface BuildInboxPollerOptions {
  chats: Pick<TeamsChatsPort, 'readMessages' | 'warmMembers' | 'resolveSelfIdStatus'>;
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
  /** Per-chat warm-up back-off window override — see InboxPollerDeps.warmupBackoffMs (inbox.ts)
   *  and TEAMS_INBOX_WARMUP_BACKOFF_SECONDS (index.ts) for where this comes from. */
  warmupBackoffMs?: number;
  /**
   * What the isSelf displayName fallback (inbox.ts) compares a message's `from` against when
   * Graph reports no `fromId` on it at all — see config.assistantDisplayName's own doc comment.
   * Sourced from config, not a live call: unlike the id (below), this value is already known at
   * startup, so passing it through costs nothing and can never block a poll (0.6.4, issue #28,
   * point 4 — the poller's old raw `/me?$select=id,displayName` fetched this live; dropping the
   * live half and sourcing it from config instead is strictly cheaper and still best-effort:
   * omitting it simply leaves that one fallback unavailable, same posture as every other optional
   * field on this interface). Optional: a caller with no assistantDisplayName configured yet just
   * loses that one fallback, never the id-based check the loop otherwise relies on.
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
    // Reuses the SAME operator-seed (TEAMS_MCP_SELF_ID) -> persisted-cache -> live `/me` chain
    // GraphTeamsChats.resolveSelfId already exposes to sendFile (teams-chats.ts's
    // resolveSelfIdStatus) instead of a raw, seed/cache-blind `/me` call of this module's own.
    // Before 0.6.4 this was `() => options.graph.get('/me?$select=id,displayName')` on the raw
    // client — a tenant-wide `/me` 429 then failed every poll (nothing in the seed/cache chain
    // was ever consulted) and, after three consecutive cycles, forced a token re-authentication
    // for a condition that was never auth-shaped (issue #28, live 2026-09-09 — see
    // KNOWN-ISSUES.md). A throttled resolution is surfaced here as a thrown, 429-shaped error so
    // InboxPoller.pollOnce classifies it as a THROTTLED cycle (markThrottled/noteRetryAfter), the
    // same treatment any other 429 already gets, rather than an "unrecognised shape" failure that
    // would eventually trip the forced re-auth path on its own.
    self: async (): Promise<SignedInAccount> => {
      const resolved = (await options.chats.resolveSelfIdStatus?.()) ?? {};
      if (resolved.id !== undefined) {
        return {
          id: resolved.id,
          ...(options.assistantDisplayName !== undefined
            ? { displayName: options.assistantDisplayName }
            : {}),
        };
      }
      throw Object.assign(
        new Error(
          resolved.throttled
            ? 'self id resolution: Too many requests (throttled /me, no operator seed or persisted cache available)'
            : 'self id resolution: could not be determined',
        ),
        {
          ...(resolved.throttled ? { status: 429 } : {}),
          ...(resolved.retryAfterSeconds !== undefined
            ? { retryAfterSeconds: resolved.retryAfterSeconds }
            : {}),
        },
      );
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
    ...(options.warmupBackoffMs !== undefined ? { warmupBackoffMs: options.warmupBackoffMs } : {}),
    log: options.log ?? (() => {}),
    // 0.4.1 stuck-auth self-healing: a token that goes bad without the local cache's own expiry
    // catching up needs an external nudge to drop it — see InboxPoller.trackAuthHealth's doc
    // comment for the live diagnosis this closes.
    onAuthStuck: () => options.tokenProvider.invalidate?.(),
  });
}
