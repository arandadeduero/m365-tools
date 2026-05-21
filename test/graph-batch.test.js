import test from 'ava';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

// This file runs in its own AVA worker. We set HOME before any module import
// so token-store uses our isolated directory, preventing cross-test contamination.
const tmpHome = await mkdtemp(join(tmpdir(), 'm365-batch-test-'));
process.env.HOME = tmpHome;

const { saveToken, clearToken, _resetKeyCache } = await import('../src/utils/token-store.js');
const { initAuth } = await import('../src/auth.js');

test.beforeEach(async () => {
  _resetKeyCache();
  await clearToken();
  initAuth({ tenantId: 'test-t', clientId: 'test-c', scopes: [] });
});

test.after.always(async () => {
  await rm(tmpHome, { recursive: true, force: true });
});

test.serial('fetchManagerMap: throws with correct message for non-OK HTTP responses', async (t) => {
  const originalFetch = globalThis.fetch;

  // --- 401 scenario ---
  await saveToken({ accessToken: 'fake-token', expiresAt: Date.now() + 3600_000 });
  globalThis.fetch = async () => ({
    ok:         false,
    status:     401,
    statusText: 'Unauthorized',
    json:       async () => ({ error: { message: 'Access token is missing or invalid.' } }),
  });
  try {
    const { fetchManagerMap } = await import('../src/graph.js');
    const err401 = await t.throwsAsync(() => fetchManagerMap([{ id: 'user-1' }]));
    t.true(err401.message.includes('401'), `expected 401 in: ${err401.message}`);
    t.true(err401.message.includes('Access token is missing or invalid.'), err401.message);
  } finally {
    globalThis.fetch = originalFetch;
  }

  // --- 503 scenario ---
  _resetKeyCache();
  await saveToken({ accessToken: 'fake-token-2', expiresAt: Date.now() + 3600_000 });
  globalThis.fetch = async () => ({
    ok:         false,
    status:     503,
    statusText: 'Service Unavailable',
    json:       async () => ({}),
  });
  try {
    const { fetchManagerMap } = await import('../src/graph.js');
    const err503 = await t.throwsAsync(() => fetchManagerMap([{ id: 'user-1' }]));
    t.true(err503.message.includes('503'), `expected 503 in: ${err503.message}`);
    t.true(err503.message.includes('Service Unavailable'), err503.message);
  } finally {
    globalThis.fetch = originalFetch;
  }
});
