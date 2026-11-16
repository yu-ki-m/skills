// Temporary folders of the tests (browser profiles, built workbooks, downloads), all named excel-html-test-*.
// A folder is removed when its test is done. On Windows a just-closed browser can keep files locked for a moment;
// such a folder is tried again when the process exits, and folders left by an interrupted run are removed by the
// next run (removeStale), so that the system temp folder does not fill up.
import {mkdtempSync, readdirSync, rmSync, statSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';

const PREFIX = 'excel-html-test-';
const pending = new Set();
process.on('exit', () => {
  for (const dir of pending) { try { rmSync(dir, {recursive: true, force: true}); } catch {} }
});

/** A new empty folder under the system temp folder. */
export function makeTempDir(kind = '') { return mkdtempSync(join(tmpdir(), PREFIX + kind)); }

export function removeTempDir(dir) {
  if (!dir) return;
  try {
    rmSync(dir, {recursive: true, force: true, maxRetries: 20, retryDelay: 100});
    pending.delete(dir);
  } catch {
    pending.add(dir);
  }
}

/** Removes test folders that were last changed more than maxAge ago (left behind when a run was interrupted). */
export function removeStale(maxAge = 6 * 60 * 60 * 1000) {
  let names = [];
  try { names = readdirSync(tmpdir()); } catch { return; }
  for (const name of names) {
    if (!name.startsWith(PREFIX)) continue;
    const dir = join(tmpdir(), name);
    try { if (Date.now() - statSync(dir).mtimeMs > maxAge) rmSync(dir, {recursive: true, force: true}); } catch {}
  }
}
