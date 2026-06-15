/**
 * Authentication module.
 *
 * Uses @azure/msal-node directly so we can provide a custom cachePlugin that
 * persists MSAL's full token cache — including the refresh token — to an
 * encrypted file on disk.  This means silent token renewal works across
 * process restarts for the lifetime of the refresh token (~90 days rolling).
 *
 * Flow on every getAccessToken() call:
 *   1. Load the serialized MSAL cache from disk into the PublicClientApp.
 *   2. Try acquireTokenSilent() using the stored account.
 *   3. If that fails (first run, or refresh token truly expired), fall back
 *      to the Device Code flow which writes a new cache to disk.
 */

import { PublicClientApplication } from '@azure/msal-node';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { join } from 'node:path';
import os from 'node:os';
import chalk from 'chalk';

const CONFIG_DIR  = join(os.homedir(), '.config', 'm365-users');
const CACHE_FILE  = join(CONFIG_DIR, 'msal-cache.json');

// Machine-specific AES-256-GCM key (same derivation as token-store.js)
const _pbkdf2Password = `${os.hostname()}:${os.userInfo().username}`;
const _pbkdf2Salt     = new TextEncoder().encode('m365-users-v1');
let   _cachedKey      = null;

async function _deriveKey() {
  if (_cachedKey) return _cachedKey;
  const enc  = new TextEncoder();
  const km   = await crypto.subtle.importKey('raw', enc.encode(_pbkdf2Password), 'PBKDF2', false, ['deriveKey']);
  _cachedKey = await crypto.subtle.deriveKey(
    { name: 'PBKDF2', salt: _pbkdf2Salt, iterations: 100_000, hash: 'SHA-256' },
    km,
    { name: 'AES-GCM', length: 256 },
    false,
    ['encrypt', 'decrypt'],
  );
  return _cachedKey;
}

async function _encryptCache(plaintext) {
  const key        = await _deriveKey();
  const iv         = crypto.getRandomValues(new Uint8Array(12));
  const encoded    = new TextEncoder().encode(plaintext);
  const ciphertext = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, encoded);
  return JSON.stringify({
    iv:   Buffer.from(iv).toString('hex'),
    data: Buffer.from(ciphertext).toString('hex'),
  });
}

async function _decryptCache(json) {
  const stored     = JSON.parse(json);
  const key        = await _deriveKey();
  const iv         = Buffer.from(stored.iv, 'hex');
  const ciphertext = Buffer.from(stored.data, 'hex');
  const plaintext  = await crypto.subtle.decrypt({ name: 'AES-GCM', iv }, key, ciphertext);
  return new TextDecoder().decode(plaintext);
}

async function _readCacheFile() {
  try {
    const raw = await readFile(CACHE_FILE, 'utf8');
    return await _decryptCache(raw);
  } catch {
    return null; // file missing or corrupt → start fresh
  }
}

async function _writeCacheFile(serialized) {
  await mkdir(CONFIG_DIR, { recursive: true });
  const encrypted = await _encryptCache(serialized);
  await writeFile(CACHE_FILE, encrypted, { mode: 0o600 });
}

// ---------------------------------------------------------------------------
// MSAL cachePlugin — called by MSAL before/after every token cache access
// ---------------------------------------------------------------------------
const _msalCachePlugin = {
  async beforeCacheAccess(ctx) {
    const data = await _readCacheFile();
    if (data) ctx.tokenCache.deserialize(data);
  },
  async afterCacheAccess(ctx) {
    if (ctx.cacheHasChanged) {
      await _writeCacheFile(ctx.tokenCache.serialize());
    }
  },
};

// ---------------------------------------------------------------------------
// Module state
// ---------------------------------------------------------------------------
let _config = null;
let _msalApp = null;   // single PublicClientApplication instance per process

export function initAuth(config) {
  _config = config;
  _msalApp = null; // reset if config changes
}

export function getDomain() {
  return _config?.domain || '';
}

function _getMsalApp() {
  if (_msalApp) return _msalApp;
  if (!_config) throw new Error('Auth not initialized. Call initAuth(config) first.');

  _msalApp = new PublicClientApplication({
    auth: {
      clientId: _config.clientId,
      authority: `https://login.microsoftonline.com/${_config.tenantId}`,
    },
    cache: {
      cachePlugin: _msalCachePlugin,
    },
  });
  return _msalApp;
}

function _normalizeScopes(scopes) {
  const defaults = [
    'https://graph.microsoft.com/User.ReadWrite.All',
    'https://graph.microsoft.com/Directory.ReadWrite.All',
  ];
  return (scopes || defaults).map((s) =>
    s.startsWith('https://') ? s : `https://graph.microsoft.com/${s}`
  );
}

/**
 * Returns a valid access token for Microsoft Graph.
 * Silently refreshes using the persisted MSAL cache when possible.
 * Only shows the Device Code prompt on first run or after refresh token expiry.
 */
export async function getAccessToken() {
  if (_fakeToken) return _fakeToken;

  const app    = _getMsalApp();
  const scopes = _normalizeScopes(_config?.scopes);

  // Find the cached account (loaded from disk by beforeCacheAccess)
  const accounts = await app.getTokenCache().getAllAccounts();
  const account  = accounts[0] ?? null;

  if (account) {
    try {
      const result = await app.acquireTokenSilent({ scopes, account });
      return result.accessToken;
    } catch {
      // Silent failed — refresh token expired or account gone; fall through to device code
    }
  }

  return _deviceCodeFlow(app, scopes);
}

async function _deviceCodeFlow(app, scopes) {
  console.log(chalk.yellow('\nNo valid session found. Starting authentication...\n'));

  const result = await app.acquireTokenByDeviceCode({
    scopes,
    deviceCodeCallback: (info) => {
      console.log(chalk.cyan('\n========================================'));
      console.log(chalk.bold('  Microsoft 365 Authentication'));
      console.log(chalk.cyan('========================================'));
      const fullUrl = `${info.verificationUri}?usercode=${info.userCode}`;
      console.log(`\n  1. Open your browser and go to (code pre-filled):`);
      console.log(chalk.underline.blue(`     ${fullUrl}`));
      console.log(`\n  2. Enter the code (if not already filled):`);
      console.log(chalk.bold.yellow(`     ${info.userCode}`));
      console.log(`\n  3. Sign in with your Microsoft 365 account`);
      console.log(chalk.cyan('\n========================================\n'));
    },
  });

  // afterCacheAccess will have already saved the cache (including refresh token)
  console.log(chalk.green('\nAuthentication successful. Session saved.\n'));
  return result.accessToken;
}

/**
 * Force re-authentication: clear the persisted cache and run device code.
 */
export async function login() {
  await _clearCache();
  _msalApp = null;
  return getAccessToken();
}

/**
 * Clear the persisted MSAL cache (logout).
 */
export async function logout() {
  await _clearCache();
  _msalApp = null;
  console.log(chalk.green('Logged out. Session token cleared.'));
}

async function _clearCache() {
  const { unlink } = await import('node:fs/promises');
  try { await unlink(CACHE_FILE); } catch { /* already gone */ }
}

// ---------------------------------------------------------------------------
// Test helpers (not for production use)
// ---------------------------------------------------------------------------

let _fakeToken = null;

/** Inject a fake access token so tests can bypass the full MSAL flow. */
export function _setFakeToken(token) { _fakeToken = token; }
export function _clearFakeToken()    { _fakeToken = null; }
