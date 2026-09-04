import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { InboxPoller } from './inbox.js';
import { startInboxSupervision } from './inbox-startup.js';

/**
 * PR #20 blocker (2026-09-04 review): src/index.ts:41-71 exited the ENTIRE process (after
 * server.connect() had already handed an MCP client a live session) whenever the per-inbox lock
 * was held by another live process. That is correct for a real second daemon, but the same
 * dist/index.js entrypoint is what an MCP-tools registration runs (scripts/install-local.ts:54),
 * against the SAME .env as a running daemon — so the tools process would exit at startup and the
 * client would silently lose every Teams tool. These tests drive startInboxSupervision (the real
 * lock + real decideInboxRole, only the poller construction and process.exit are stood in for)
 * to prove the losing side now serves tools with no poller and never exits.
 */
describe('startInboxSupervision — composition through the real lock + real decideInboxRole', () => {
  let dir: string;
  let inboxPath: string;
  let lockPath: string;
  let exitSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'inbox-startup-'));
    inboxPath = join(dir, 'inbox.jsonl');
    lockPath = join(dir, 'poller.lock');
    // Mocked, not just spied: a real process.exit would kill the test runner itself.
    exitSpy = vi.spyOn(process, 'exit').mockImplementation(() => undefined as never);
  });

  afterEach(async () => {
    exitSpy.mockRestore();
    await rm(dir, { recursive: true, force: true });
  });

  const pollerOptions = {} as never; // buildPoller is stubbed below; the real shape is build-inbox-poller.test.ts's job.

  it('lock already held by a live OTHER process: serves tools, builds NO poller, logs the exact line, never exits', async () => {
    // pid 1 (init/systemd) always exists and is never this test process — a live holder that is
    // reliably NOT us, without injecting isPidAlive (this is the REAL acquirePollerLock).
    await writeFile(
      lockPath,
      JSON.stringify({ pid: 1, startedAt: '2026-09-04T09:00:00.000Z' }, null, 2),
    );
    const lines: string[] = [];
    const buildPoller = vi.fn();

    const result = await startInboxSupervision({
      env: {},
      inboxPath,
      pollerOptions,
      log: (line) => lines.push(line),
      buildPoller,
    });

    expect(result.role).toBe('serve-only');
    expect(result.poller).toBeUndefined();
    expect(buildPoller).not.toHaveBeenCalled();
    expect(exitSpy).not.toHaveBeenCalled();
    expect(lines).toEqual([`inbox poller: another instance holds ${lockPath}; serving tools without polling`]);
    // The live holder's lock file must be left exactly as it was — the loser must not clobber it.
    expect(JSON.parse(await readFile(lockPath, 'utf8'))).toEqual({ pid: 1, startedAt: '2026-09-04T09:00:00.000Z' });
  });

  it('lock free: builds and starts the poller, logs it is on, never exits', async () => {
    const lines: string[] = [];
    const fakePoller = { start: vi.fn() } as unknown as InboxPoller;
    const buildPoller = vi.fn().mockReturnValue(fakePoller);

    const result = await startInboxSupervision({
      env: {},
      inboxPath,
      pollerOptions,
      log: (line) => lines.push(line),
      buildPoller,
    });

    expect(result.role).toBe('poll');
    expect(result.poller).toBe(fakePoller);
    expect(buildPoller).toHaveBeenCalledTimes(1);
    expect(buildPoller).toHaveBeenCalledWith(expect.objectContaining({ inboxPath }));
    expect(fakePoller.start).toHaveBeenCalledTimes(1);
    expect(exitSpy).not.toHaveBeenCalled();
    expect(lines).toEqual([`inbox poller on: ${inboxPath}`]);
    // We took the lock for real: the file now names OUR pid, not a placeholder.
    expect(JSON.parse(await readFile(lockPath, 'utf8'))).toMatchObject({ pid: process.pid });
  });

  it('TEAMS_INBOX_DISABLED: no lock file is even written, no poller, never exits', async () => {
    const lines: string[] = [];
    const buildPoller = vi.fn();

    const result = await startInboxSupervision({
      env: { TEAMS_INBOX_DISABLED: '1' },
      inboxPath,
      pollerOptions,
      log: (line) => lines.push(line),
      buildPoller,
    });

    expect(result.role).toBe('disabled');
    expect(buildPoller).not.toHaveBeenCalled();
    expect(exitSpy).not.toHaveBeenCalled();
    expect(lines).toEqual(['inbox poller off (TEAMS_INBOX_DISABLED)']);
    await expect(readFile(lockPath, 'utf8')).rejects.toThrow();
  });
});
