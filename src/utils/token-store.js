import { readFile, writeFile, mkdir, unlink } from 'node:fs/promises';
import { join } from 'node:path';
import os from 'node:os';

const CONFIG_DIR = join(os.homedir(), '.config', 'm365-users');
export const TOKEN_FILE = join(CONFIG_DIR, 'token.json');

// Machine-specific password for PBKDF2 key derivation.
// Salt is a separate fixed string — password and salt must not be the same.
const _pbkdf2Password = `${os.hostname()}:${os.userInfo().username}`;
const _pbkdf2Salt = new TextEncoder().encode('m365-users-v1');

// Cache the derived key — PBKDF2 with 100k iterations is intentionally slow;
// re-derive only once per process lifetime (key is deterministic for this machine/user).
let _cachedKey = null;

async function deriveKey() {
  if (_cachedKey) return _cachedKey;
  const enc = new TextEncoder();
  const keyMaterial = await crypto.subtle.importKey(
    'raw',
    enc.encode(_pbkdf2Password),
    'PBKDF2',
    false,
    ['deriveKey']
  );
  _cachedKey = await crypto.subtle.deriveKey(
    { name: 'PBKDF2', salt: _pbkdf2Salt, iterations: 100_000, hash: 'SHA-256' },
    keyMaterial,
    { name: 'AES-GCM', length: 256 },
    false,
    ['encrypt', 'decrypt']
  );
  return _cachedKey;
}

async function encrypt(data) {
  const key = await deriveKey();
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const encoded = new TextEncoder().encode(JSON.stringify(data));
  const ciphertext = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, encoded);
  return {
    iv: Buffer.from(iv).toString('hex'),
    data: Buffer.from(ciphertext).toString('hex'),
  };
}

async function decrypt(stored) {
  const key = await deriveKey();
  const iv = Buffer.from(stored.iv, 'hex');
  const ciphertext = Buffer.from(stored.data, 'hex');
  const plaintext = await crypto.subtle.decrypt({ name: 'AES-GCM', iv }, key, ciphertext);
  return JSON.parse(new TextDecoder().decode(plaintext));
}

export async function saveToken(tokenData) {
  await mkdir(CONFIG_DIR, { recursive: true });
  const encrypted = await encrypt(tokenData);
  await writeFile(TOKEN_FILE, JSON.stringify(encrypted), { mode: 0o600 });
}

export async function loadToken() {
  try {
    const stored = JSON.parse(await readFile(TOKEN_FILE, 'utf8'));
    return await decrypt(stored);
  } catch {
    return null;
  }
}

/** Remove the token file entirely. Safe to call even if no token exists. */
export async function clearToken() {
  try {
    await unlink(TOKEN_FILE);
  } catch {
    // file may not exist — nothing to clear
  }
}

// Exported for testing: reset the in-process key cache
export function _resetKeyCache() {
  _cachedKey = null;
}

export function isTokenExpired(tokenData) {
  if (!tokenData || !tokenData.expiresAt) return true;
  // Consider expired 5 minutes before actual expiry
  return Date.now() >= tokenData.expiresAt - 5 * 60 * 1000;
}
