import { FileTokenCache } from './auth/token-cache.js';
import { RopcTokenProvider } from './auth/ropc-token-provider.js';
import type { TokenProvider } from './auth/token-provider.js';
import { GraphClient } from './graph/graph-client.js';
import { MembersCache } from './graph/members-cache.js';
import { GraphTeamsChats, type GraphTeamsChatsOptions } from './graph/teams-chats.js';
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
  options: GraphTeamsChatsOptions = {},
): ChatsStack {
  const tokenProvider = new RopcTokenProvider({
    tenantId: config.tenantId,
    clientId: config.clientId,
    username: config.username,
    password: config.password,
    cache: new FileTokenCache(config.tokenCachePath),
  });
  const graph = new GraphClient({ tokenProvider });
  const membersCache = new MembersCache({ path: config.membersCachePath, ttlMs: config.membersTtlMs });
  // Readback-before-retry on every send: a failure report is a claim about the response path,
  // not the chat, and re-sending without checking is how one broadcast became eleven.
  const chats = new ReliableTeamsChats(
    new GraphTeamsChats(graph, { ...options, membersCache: options.membersCache ?? membersCache }),
    { selfDisplayName: config.assistantDisplayName },
  );
  return { chats, graph, tokenProvider };
}
