/**
 * Point os.homedir() at a throwaway directory so nothing under test can read
 * or delete the real ~/.config/m365-users (MSAL session, token file).
 *
 * Import this BEFORE any src/ module: auth.js and token-store.js resolve their
 * config directory at import time.
 *
 * Only works because ava runs with `workerThreads: false` (see package.json).
 * In a worker thread process.env is a JS-side copy and os.homedir() keeps
 * returning the real home, so we fail loudly rather than touch real files.
 */
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir, homedir } from 'node:os';
import { join } from 'node:path';

export const tmpHome = mkdtempSync(join(tmpdir(), 'm365-users-test-'));
process.env.HOME = tmpHome;
process.env.USERPROFILE = tmpHome;

if (homedir() !== tmpHome) {
  rmSync(tmpHome, { recursive: true, force: true });
  throw new Error(
    'Test home isolation failed: os.homedir() still points at the real home. ' +
    'Keep "workerThreads": false in the ava config.',
  );
}

export function cleanupHome() {
  rmSync(tmpHome, { recursive: true, force: true });
}
