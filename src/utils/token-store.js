import { createCipheriv, createDecipheriv, randomBytes, scryptSync } from 'crypto';
import { readFileSync, writeFileSync, mkdirSync, existsSync } from 'fs';
import { join } from 'path';
import os from 'os';

const CONFIG_DIR = join(os.homedir(), '.config', 'm365-users');
const TOKEN_FILE = join(CONFIG_DIR, 'token.json');

// Derive a machine-specific key using hostname + username as salt
function deriveKey() {
  const salt = `${os.hostname()}:${os.userInfo().username}:m365-users`;
  return scryptSync(salt, salt, 32);
}

function encrypt(data) {
  const key = deriveKey();
  const iv = randomBytes(16);
  const cipher = createCipheriv('aes-256-gcm', key, iv);
  const encrypted = Buffer.concat([
    cipher.update(JSON.stringify(data), 'utf8'),
    cipher.final(),
  ]);
  const authTag = cipher.getAuthTag();
  return {
    iv: iv.toString('hex'),
    authTag: authTag.toString('hex'),
    data: encrypted.toString('hex'),
  };
}

function decrypt(stored) {
  const key = deriveKey();
  const iv = Buffer.from(stored.iv, 'hex');
  const authTag = Buffer.from(stored.authTag, 'hex');
  const encryptedData = Buffer.from(stored.data, 'hex');
  const decipher = createDecipheriv('aes-256-gcm', key, iv);
  decipher.setAuthTag(authTag);
  const decrypted = Buffer.concat([decipher.update(encryptedData), decipher.final()]);
  return JSON.parse(decrypted.toString('utf8'));
}

export function saveToken(tokenData) {
  if (!existsSync(CONFIG_DIR)) {
    mkdirSync(CONFIG_DIR, { recursive: true });
  }
  const encrypted = encrypt(tokenData);
  writeFileSync(TOKEN_FILE, JSON.stringify(encrypted), { mode: 0o600 });
}

export function loadToken() {
  if (!existsSync(TOKEN_FILE)) return null;
  try {
    const stored = JSON.parse(readFileSync(TOKEN_FILE, 'utf8'));
    return decrypt(stored);
  } catch {
    return null;
  }
}

export function clearToken() {
  if (existsSync(TOKEN_FILE)) {
    writeFileSync(TOKEN_FILE, JSON.stringify({}));
  }
}

export function isTokenExpired(tokenData) {
  if (!tokenData || !tokenData.expiresAt) return true;
  // Consider expired 5 minutes before actual expiry
  return Date.now() >= tokenData.expiresAt - 5 * 60 * 1000;
}
