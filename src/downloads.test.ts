import { describe, expect, it } from 'vitest';
import { mkdtempSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join } from 'node:path';
import { defaultDownloadDir, sanitizeFileName, writeDownload } from './downloads.js';

describe('sanitizeFileName — sender-controlled names becoming local files', () => {
  it('strips path components of both separator flavours', () => {
    // basename() alone leaves the backslash form intact on Linux — that is the case that
    // matters, since a Teams sender types the name and the server runs wherever it runs.
    expect(sanitizeFileName('../../etc/passwd')).toBe('passwd');
    expect(sanitizeFileName('..\\..\\windows\\system32\\evil.dll')).toBe('evil.dll');
    expect(sanitizeFileName('/absolute/path/report.pdf')).toBe('report.pdf');
  });

  it('replaces the characters Windows refuses and every control character', () => {
    expect(sanitizeFileName('a<b>c:d"e|f?g*h.txt')).toBe('a_b_c_d_e_f_g_h.txt');
    expect(sanitizeFileName('bell\u0007name\u001f.txt')).toBe('bell_name_.txt');
  });

  it('keeps ordinary names — spaces, dashes, unicode — untouched', () => {
    expect(sanitizeFileName('Q3 sales plan — final (2).xlsx')).toBe('Q3 sales plan — final (2).xlsx');
  });

  it('trims trailing dots and spaces, which Windows silently drops', () => {
    expect(sanitizeFileName('report.pdf. . .')).toBe('report.pdf');
  });

  it('falls back to attachment.bin when nothing usable is left', () => {
    expect(sanitizeFileName('')).toBe('attachment.bin');
    expect(sanitizeFileName('..')).toBe('attachment.bin');
    expect(sanitizeFileName('???')).toBe('attachment.bin');
    expect(sanitizeFileName('   ')).toBe('attachment.bin');
  });

  it('caps absurd lengths while keeping the extension', () => {
    const sanitized = sanitizeFileName(`${'x'.repeat(400)}.docx`);
    expect(sanitized.length).toBeLessThanOrEqual(150);
    expect(sanitized.endsWith('.docx')).toBe(true);
  });
});

describe('writeDownload — never overwrites, never escapes the directory', () => {
  it('writes the bytes and returns an absolute path inside the directory', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'downloads-test-'));

    const path = await writeDownload(dir, 'notes.txt', new Uint8Array([104, 105]));

    expect(path).toBe(join(dir, 'notes.txt'));
    expect(readFileSync(path, 'utf8')).toBe('hi');
  });

  it('suffixes on collision instead of silently replacing the earlier download', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'downloads-test-'));

    const first = await writeDownload(dir, 'plan.xlsx', new Uint8Array([1]));
    const second = await writeDownload(dir, 'plan.xlsx', new Uint8Array([2]));
    const third = await writeDownload(dir, 'plan.xlsx', new Uint8Array([3]));

    expect(basename(first)).toBe('plan.xlsx');
    expect(basename(second)).toBe('plan-1.xlsx');
    expect(basename(third)).toBe('plan-2.xlsx');
    // The first download's bytes are untouched by the later ones.
    expect([...readFileSync(first)]).toEqual([1]);
  });

  it('a hostile name lands inside the directory, sanitized', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'downloads-test-'));

    const path = await writeDownload(dir, '../../escape.pdf', new Uint8Array([9]));

    expect(path).toBe(join(dir, 'escape.pdf'));
  });

  it('creates the directory when it does not exist yet', async () => {
    const dir = join(mkdtempSync(join(tmpdir(), 'downloads-test-')), 'deeper', 'still');

    const path = await writeDownload(dir, 'a.bin', new Uint8Array([0]));

    expect(path.startsWith(dir)).toBe(true);
  });
});

describe('defaultDownloadDir', () => {
  it('prefers TEAMS_MCP_DOWNLOAD_DIR, else falls back under the OS tmpdir', () => {
    expect(defaultDownloadDir({ TEAMS_MCP_DOWNLOAD_DIR: '/data/teams' } as NodeJS.ProcessEnv)).toBe('/data/teams');
    expect(defaultDownloadDir({} as NodeJS.ProcessEnv)).toBe(join(tmpdir(), 'teams-assistant-mcp'));
  });
});
