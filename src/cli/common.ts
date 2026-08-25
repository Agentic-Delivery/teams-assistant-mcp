import { FileTokenCache } from '../auth/token-cache.js';
import { RopcTokenProvider } from '../auth/ropc-token-provider.js';
import { GraphClient } from '../graph/graph-client.js';
import { GraphTeamsChats } from '../graph/teams-chats.js';
import { ReliableTeamsChats } from '../graph/reliable-sends.js';
import { ChatNotAllowedError, type ChatAllowlist } from '../allowlist.js';
import { loadConfig } from '../config.js';

/**
 * Shared plumbing for the standalone CLIs (teams-post, teams-reply, teams-react, teams-read).
 *
 * The output contract is the whole point, learned the hard way on 2026-08-24 when a caller
 * grepped for a success token the old ad-hoc script never printed and re-posted a broadcast
 * ten extra times: SUCCESS is exactly one JSON line on stdout and exit 0 — nothing else ever
 * reaches stdout. Failure is prose on stderr and a non-zero exit (2 usage, 3 allowlist,
 * 1 everything else). Callers branch on the exit code, never on output text.
 */
export interface CliContext {
  chats: ReliableTeamsChats;
  allowlist: ChatAllowlist;
}

export function buildContext(): CliContext {
  const config = loadConfig();
  const tokenProvider = new RopcTokenProvider({
    tenantId: config.tenantId,
    clientId: config.clientId,
    username: config.username,
    password: config.password,
    cache: new FileTokenCache(config.tokenCachePath),
  });
  const graph = new GraphClient({ tokenProvider });
  return {
    chats: new ReliableTeamsChats(new GraphTeamsChats(graph)),
    allowlist: config.allowlist,
  };
}

export function readStdin(): Promise<string> {
  return new Promise((resolve) => {
    let buffer = '';
    process.stdin.on('data', (chunk) => (buffer += chunk));
    process.stdin.on('end', () => resolve(buffer.trim()));
  });
}

export function succeed(payload: Record<string, unknown>): never {
  process.stdout.write(`${JSON.stringify({ ok: true, ...payload })}\n`);
  process.exit(0);
}

export function usage(text: string): never {
  process.stderr.write(`${text}\n`);
  process.exit(2);
}

export async function run(main: () => Promise<never>): Promise<void> {
  try {
    await main();
  } catch (caught) {
    process.stderr.write(`${caught instanceof Error ? caught.message : String(caught)}\n`);
    process.exit(caught instanceof ChatNotAllowedError ? 3 : 1);
  }
}
