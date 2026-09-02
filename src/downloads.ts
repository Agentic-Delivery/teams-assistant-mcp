import { mkdir, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, extname, join, resolve } from 'node:path';

/**
 * Where attachment downloads land when the caller names nothing: TEAMS_MCP_DOWNLOAD_DIR from the
 * environment, else a directory under the OS tmpdir (which the OS may clean between sessions —
 * env.example says so next to the variable). One resolver for every surface that writes
 * downloads — the MCP server's default (src/index.ts) and the teams-attachments CLI — so the two
 * can never disagree about where "the download dir" is.
 */
export function defaultDownloadDir(env: NodeJS.ProcessEnv = process.env): string {
  return env['TEAMS_MCP_DOWNLOAD_DIR'] ?? join(tmpdir(), 'teams-assistant-mcp');
}

/** Windows' reserved set plus control chars — the strictest common denominator, applied on every
 *  platform so a download made on Linux still syncs to a Windows checkout without a rename. */
// oxlint-disable-next-line no-control-regex -- matching the control range is the whole point
const UNSAFE_CHARS = /[<>:"/\\|?*\u0000-\u001f]/g;
const MAX_NAME_LENGTH = 150;

/**
 * An attachment's name is SENDER-CONTROLLED data being turned into a local file name, so it gets
 * the full treatment: path components stripped (both separator flavours — basename alone leaves
 * `..\evil` intact on Linux), reserved/control characters replaced, trailing dots and spaces
 * trimmed (Windows silently drops them, which would make the returned path a lie), and a length
 * cap that keeps the extension. A name with nothing usable left becomes "attachment.bin" rather
 * than an error — the bytes are already downloaded and the caller wants them.
 */
export function sanitizeFileName(name: string): string {
  const lastComponent = basename(name.replace(/\\/g, '/'));
  const cleaned = lastComponent.replace(UNSAFE_CHARS, '_').replace(/[. ]+$/, '').trim();
  if (cleaned === '' || cleaned === '.' || cleaned === '..' || /^_+$/.test(cleaned)) {
    return 'attachment.bin';
  }
  if (cleaned.length <= MAX_NAME_LENGTH) {
    return cleaned;
  }
  const extension = extname(cleaned);
  return cleaned.slice(0, MAX_NAME_LENGTH - extension.length) + extension;
}

/** How many `name-N.ext` variants writeDownload tries before giving up. Hitting this means the
 *  directory already holds a thousand same-named downloads — that is a caller bug, not bad luck. */
const MAX_COLLISION_SUFFIX = 1000;

/**
 * Writes one downloaded attachment into `dir` and returns the absolute path it landed at.
 * NEVER overwrites: an existing file makes the write land at `name-1.ext`, `name-2.ext`, … —
 * two messages carrying files with the same name must yield two files, not one silently
 * replacing the other. The exclusive-create flag (`wx`) is what makes that race-safe: existence
 * is checked and the file claimed in one syscall, so two concurrent downloads cannot both
 * "see nothing there" and clobber each other.
 */
export async function writeDownload(dir: string, name: string, bytes: Uint8Array): Promise<string> {
  await mkdir(dir, { recursive: true });
  const safe = sanitizeFileName(name);
  const extension = extname(safe);
  const stem = safe.slice(0, safe.length - extension.length);
  for (let attempt = 0; attempt <= MAX_COLLISION_SUFFIX; attempt += 1) {
    const candidate = attempt === 0 ? safe : `${stem}-${attempt}${extension}`;
    const path = resolve(join(dir, candidate));
    try {
      await writeFile(path, bytes, { flag: 'wx' });
      return path;
    } catch (caught) {
      if ((caught as NodeJS.ErrnoException).code !== 'EEXIST') {
        throw caught;
      }
    }
  }
  throw new Error(
    `Could not find a free file name for "${safe}" in ${dir} after ${MAX_COLLISION_SUFFIX} suffixes.`,
  );
}
