/**
 * Prints the registration command for this checkout. `npm run install-local` builds first, so by
 * the time anything is printed the dist/ path in the command exists.
 *
 * The server itself never opens .env — dist/index.js reads process.env and nothing else. Node's
 * --env-file flag is what loads the file, which is why it is part of the printed command.
 */
import { existsSync, readFileSync } from 'node:fs';
import { dirname, isAbsolute, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseEnv } from 'node:util';

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..');
const entry = join(repoRoot, 'dist', 'index.js');
const envFile = join(repoRoot, '.env');

if (!existsSync(entry)) {
  console.error(`${entry} is missing even though the build just ran. Check the build output.`);
  process.exit(1);
}

const warnings: string[] = [];

if (!existsSync(envFile)) {
  warnings.push(`There is no .env yet. Copy env.example to .env and fill it in (see SETUP.md).`);
} else {
  const env = parseEnv(readFileSync(envFile, 'utf8'));
  const config = env['TEAMS_MCP_CONFIG'];
  if (!config) {
    warnings.push('TEAMS_MCP_CONFIG is not set in .env; the server will not start without it.');
  } else if (!isAbsolute(config)) {
    warnings.push(
      `TEAMS_MCP_CONFIG is relative (${config}). It resolves against the directory the agent ` +
        `starts the server in, not this repo. Use ${join(repoRoot, 'teams-mcp.config.json')}.`,
    );
  }
  const cache = env['TEAMS_MCP_TOKEN_CACHE'];
  if (!cache) {
    warnings.push(
      'TEAMS_MCP_TOKEN_CACHE is not set. The default drops a live refresh token as ' +
        `.token-cache.json in the agent's working directory. Set it to ` +
        `${join(repoRoot, '.token-cache.json')} to keep it in this repo, where it is gitignored.`,
    );
  } else if (!isAbsolute(cache)) {
    warnings.push(`TEAMS_MCP_TOKEN_CACHE is relative (${cache}). Same problem as above; use an absolute path.`);
  }
}

for (const warning of warnings) {
  console.log(`WARNING: ${warning}\n`);
}

console.log('Run this inside the project that should get the Teams tools:\n');
console.log(`  claude mcp add teams-assistant -- node --env-file=${envFile} ${entry}\n`);
console.log('Or commit the equivalent into that project\'s .mcp.json (paths are machine-specific,');
console.log('so each teammate runs install-local in their own checkout):\n');
console.log(
  JSON.stringify(
    {
      mcpServers: {
        'teams-assistant': { command: 'node', args: [`--env-file=${envFile}`, entry] },
      },
    },
    null,
    2,
  ),
);
