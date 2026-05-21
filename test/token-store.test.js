import test from 'ava';
import { join } from 'node:path';
import { mkdtemp, rm, stat } from 'node:fs/promises';
import { tmpdir } from 'node:os';

// Set HOME to a temp dir BEFORE the module is imported so CONFIG_DIR
// points to our isolated directory instead of ~/.config/m365-users.
const tmpHome = await mkdtemp(join(tmpdir(), 'm365-users-test-'));
process.env.HOME = tmpHome;

const {
  saveToken,
  loadToken,
  clearToken,
  isTokenExpired,
  _resetKeyCache,
  TOKEN_FILE,
} = await import('../src/utils/token-store.js');

// Run serially so tests don't race on the same token file.
// Clean up between each test.
test.beforeEach(async () => {
  _resetKeyCache();
  await clearToken();
});

test.after.always(async () => {
  await rm(tmpHome, { recursive: true, force: true });
});

// ---------------------------------------------------------------------------
// isTokenExpired (pure, no I/O)
// ---------------------------------------------------------------------------

test.serial('isTokenExpired: null tokenData → expired', (t) => {
  t.true(isTokenExpired(null));
});

test.serial('isTokenExpired: missing expiresAt → expired', (t) => {
  t.true(isTokenExpired({ accessToken: 'abc' }));
});

test.serial('isTokenExpired: expiry in the past → expired', (t) => {
  t.true(isTokenExpired({ expiresAt: Date.now() - 1000 }));
});

test.serial('isTokenExpired: expiry within 5-minute buffer → expired', (t) => {
  t.true(isTokenExpired({ expiresAt: Date.now() + 4 * 60 * 1000 }));
});

test.serial('isTokenExpired: expiry beyond 5-minute buffer → not expired', (t) => {
  t.false(isTokenExpired({ expiresAt: Date.now() + 10 * 60 * 1000 }));
});

// ---------------------------------------------------------------------------
// saveToken + loadToken round-trip
// ---------------------------------------------------------------------------

test.serial('saveToken then loadToken returns the original data', async (t) => {
  const token = { accessToken: 'tok_abc123', expiresAt: Date.now() + 3600_000 };
  await saveToken(token);
  const loaded = await loadToken();
  t.deepEqual(loaded, token);
});

test.serial('saveToken writes a file with restricted permissions (0o600)', async (t) => {
  await saveToken({ accessToken: 'tok', expiresAt: Date.now() + 3600_000 });
  const s = await stat(TOKEN_FILE);
  t.is(s.mode & 0o777, 0o600);
});

test.serial('loadToken returns null when no file exists', async (t) => {
  const result = await loadToken();
  t.is(result, null);
});

test.serial('loadToken returns null for corrupt/truncated file', async (t) => {
  const { writeFile, mkdir } = await import('node:fs/promises');
  const { dirname } = await import('node:path');
  await mkdir(dirname(TOKEN_FILE), { recursive: true });
  await writeFile(TOKEN_FILE, 'NOT_JSON', { mode: 0o600 });
  const result = await loadToken();
  t.is(result, null);
});

// ---------------------------------------------------------------------------
// clearToken
// ---------------------------------------------------------------------------

test.serial('clearToken removes the token file', async (t) => {
  await saveToken({ accessToken: 'tok', expiresAt: Date.now() + 3600_000 });
  await clearToken();
  const result = await loadToken();
  t.is(result, null);
});

test.serial('clearToken is idempotent — no error if called twice', async (t) => {
  await clearToken();
  await t.notThrowsAsync(() => clearToken());
});

// ---------------------------------------------------------------------------
// Overwrite behaviour
// ---------------------------------------------------------------------------

test.serial('second saveToken overwrites first — loadToken returns latest', async (t) => {
  const first  = { accessToken: 'first',  expiresAt: Date.now() + 1000 };
  const second = { accessToken: 'second', expiresAt: Date.now() + 9999 };
  await saveToken(first);
  await saveToken(second);
  const loaded = await loadToken();
  t.is(loaded.accessToken, 'second');
});
