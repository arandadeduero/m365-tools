import test from 'ava';
import { mkdir, writeFile, stat } from 'node:fs/promises';
import { join } from 'node:path';
// Must come before the src import: logout() deletes the MSAL cache under HOME.
import { tmpHome, cleanupHome } from './helpers/home.js';

// auth.js tests must run serially — they share module-level state (_config, _msalApp)
// and touch the filesystem (MSAL cache file).

const {
  initAuth,
  getDomain,
  logout,
} = await import('../src/auth.js');

const FAKE_CONFIG = {
  tenantId: 'test-tenant-id',
  clientId: 'test-client-id',
  domain:   'contoso.com',
  scopes:   ['User.ReadWrite.All'],
};

const CACHE_FILE = join(tmpHome, '.config', 'm365-users', 'msal-cache.json');

test.beforeEach(() => {
  initAuth(FAKE_CONFIG);
});

test.after.always(() => cleanupHome());

// ---------------------------------------------------------------------------
// initAuth / getDomain
// ---------------------------------------------------------------------------

test.serial('getDomain: returns empty string when initAuth called with null', (t) => {
  initAuth(null);
  t.is(getDomain(), '');
});

test.serial('getDomain: returns domain from config after initAuth', (t) => {
  initAuth(FAKE_CONFIG);
  t.is(getDomain(), 'contoso.com');
});

test.serial('initAuth: config without domain → getDomain returns empty string', (t) => {
  initAuth({ tenantId: 'x', clientId: 'y' });
  t.is(getDomain(), '');
});

// ---------------------------------------------------------------------------
// logout
// ---------------------------------------------------------------------------

test.serial('logout: removes the MSAL cache file if it exists', async (t) => {
  await mkdir(join(tmpHome, '.config', 'm365-users'), { recursive: true });
  await writeFile(CACHE_FILE, 'stale');

  await t.notThrowsAsync(() => logout());

  const exists = await stat(CACHE_FILE).then(() => true).catch(() => false);
  t.false(exists);
});

test.serial('logout: is idempotent — calling twice does not throw', async (t) => {
  await t.notThrowsAsync(() => logout());
  await t.notThrowsAsync(() => logout());
});
