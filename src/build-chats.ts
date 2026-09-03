import { FileTokenCache } from './auth/token-cache.js';
import { RopcTokenProvider } from './auth/ropc-token-provider.js';
import type { TokenProvider } from './auth/token-provider.js';
import { GraphClient } from './graph/graph-client.js';
import { MembersCache } from './graph/members-cache.js';
import { FileSelfIdCache } from './graph/self-id-cache.js';
import { GraphTeamsChats, type GraphTeamsChatsOptions } from './graph/teams-chats.js';

/** The options buildChats itself may override; membersCache and selfIdCache are always this
 *  module's own wiring (see below) — a caller cannot silently unwire either, since both are
 *  excluded from this Partial (membersCache is a required field on the underlying type;
 *  selfIdCache is optional there but excluded here for the same reason: production always wants
 *  the disk-persisted one, per config.selfIdCachePath, not whatever default GraphTeamsChats
 *  would otherwise pick). */
export type BuildChatsOptions = Partial<Omit<GraphTeamsChatsOptions, 'membersCache' | 'selfIdCache'>>;
import { ReliableTeamsChats } from './graph/reliable-sends.js';
import type { TeamsMcpConfig } from './config.js';

/**
 * The one composition of the send stack, shared by the server (index.ts) and every CLI. README
 * promises "same code paths as the server tools — including the send reliability"; this function
 * is what makes that promise structural instead of a matter of two files staying in sync.
 */
export interface ChatsStack {
  chats: ReliableTeamsChats;
  /** The raw client, for the few callers that need Graph beyond chats (the inbox poller's /me). */
  graph: GraphClient;
  /** Exposed so a caller (the inbox poller's stuck-auth recovery, 0.4.1) can force a re-mint
   *  without reaching back into this module's private wiring. */
  tokenProvider: TokenProvider;
}

export function buildChats(
  config: TeamsMcpConfig,
  options: BuildChatsOptions = {},
): ChatsStack {
  const tokenProvider = new RopcTokenProvider({
    tenantId: config.tenantId,
    clientId: config.clientId,
    username: config.username,
    password: config.password,
    cache: new FileTokenCache(config.tokenCachePath),
  });
  const graph = new GraphClient({ tokenProvider });
  // Always THIS module's own cache — membersCache is deliberately excluded from
  // BuildChatsOptions so nothing calling buildChats can accidentally unwire it (0.4.1 review).
  const membersCache = new MembersCache({ path: config.membersCachePath, ttlMs: config.membersTtlMs });
  // Same reasoning, 0.4.3: the whole point of persisting the self id is that every fresh CLI
  // process (buildChats's other caller, cli/common.ts) benefits from a resolution some earlier
  // process already paid for — an unwired selfIdCache here would leave that CLI exactly where
  // the 2026-09-03 incident found it, with the class of unit test at the GraphTeamsChats level
  // unable to catch a dropped composition wire (see this module's own buildChats test in
  // teams-chats.test.ts for the one that does).
  const selfIdCache = new FileSelfIdCache({ path: config.selfIdCachePath });
  // Readback-before-retry on every send: a failure report is a claim about the response path,
  // not the chat, and re-sending without checking is how one broadcast became eleven.
  const chats = new ReliableTeamsChats(
    new GraphTeamsChats(graph, {
      log: (line) => process.stderr.write(`[teams-assistant-mcp] ${line}\n`),
      ...(config.selfIdOverride !== undefined ? { selfIdOverride: config.selfIdOverride } : {}),
      ...options,
      membersCache,
      selfIdCache,
    }),
    { selfDisplayName: config.assistantDisplayName },
  );
  return { chats, graph, tokenProvider };
}
