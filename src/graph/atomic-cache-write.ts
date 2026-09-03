import { chmodSync, mkdirSync, renameSync, writeFileSync } from 'node:fs';
import { dirname } from 'node:path';

/**
 * Write-to-temp-then-rename, 0600, shared by every disk-persisted cache in this package
 * (MembersCache, FileSelfIdCache) — extracted (review round 2) so the atomicity mechanism is
 * proven exactly once instead of once per cache class. Before this extraction, MembersCache and
 * FileSelfIdCache each carried their own copy of this exact same nine lines; only MembersCache's
 * copy had a test that actually exercised the injected write/rename seams (members-cache.test.ts's
 * "reaches the destination path ONLY via rename" test) — FileSelfIdCache's copy had the same
 * seams on its constructor but nothing ever drove them, so a mutation replacing its rename with a
 * direct write left the whole suite green. One implementation, shared, means MembersCache's
 * existing atomicity test now structurally covers FileSelfIdCache too — both classes call this
 * function and nothing else for their writes.
 *
 * A crash mid-write must never leave a half-written cache file the next read chokes on — same
 * failure mode the watermark sidecar's plain writeFile used to have (see members-cache.ts's own
 * doc comment for the fuller history). 0600 mirrors the token cache sibling's posture
 * (FileTokenCache): set on the tmp file at create time, chmodSync as belt-and-braces
 * (writeFileSync's mode option only applies on create, not on an existing path), and once more on
 * the final path after the rename, since an exotic umask/platform edge case is cheaper to guard
 * against here than to debug later.
 */
export interface AtomicCacheWriteOptions {
  /** Destination path — the ONLY thing ever reached via a rename, never a direct write. */
  path: string;
  data: string;
  /** Injectable clock, so the tmp filename is deterministic in tests. */
  now: () => number;
  /** Injectable write primitives — tests use these to prove the destination path is reached
   *  ONLY via rename (0.4.1 review round 1: an assertion that merely checked "no leftover tmp
   *  file" passed identically with atomicity deleted entirely). */
  writeFileFn?: typeof writeFileSync;
  renameFn?: typeof renameSync;
}

export function writeAtomicCacheFile(options: AtomicCacheWriteOptions): void {
  const writeFileFn = options.writeFileFn ?? writeFileSync;
  const renameFn = options.renameFn ?? renameSync;
  mkdirSync(dirname(options.path), { recursive: true });
  const tmpPath = `${options.path}.tmp-${process.pid}-${options.now()}`;
  writeFileFn(tmpPath, options.data, { encoding: 'utf8', mode: 0o600 });
  chmodSync(tmpPath, 0o600);
  renameFn(tmpPath, options.path);
  chmodSync(options.path, 0o600);
}
