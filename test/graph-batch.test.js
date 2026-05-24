import test from 'ava';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const tmpHome = await mkdtemp(join(tmpdir(), 'm365-batch-test-'));
process.env.HOME = tmpHome;

const { initAuth, _setFakeToken, _clearFakeToken } = await import('../src/auth.js');

test.before(() => {
  initAuth({ tenantId: 'test-t', clientId: 'test-c', scopes: [] });
});

test.after.always(async () => {
  _clearFakeToken();
  await rm(tmpHome, { recursive: true, force: true });
});

test.serial('fetchManagerMap: throws with correct message for non-OK HTTP responses', async (t) => {
  const originalFetch = globalThis.fetch;

  // --- 401 scenario ---
  _setFakeToken('fake-token');
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
    _clearFakeToken();
  }

  // --- 503 scenario ---
  _setFakeToken('fake-token-2');
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
    _clearFakeToken();
  }
});
