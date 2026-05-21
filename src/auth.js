import { DeviceCodeCredential } from '@azure/identity';
import { saveToken, loadToken, isTokenExpired, clearToken } from './utils/token-store.js';
import chalk from 'chalk';

let _credential = null;
let _config = null;

export function initAuth(config) {
  _config = config;
}

/** Returns the configured tenant domain (e.g. "contoso.com") */
export function getDomain() {
  return _config?.domain || '';
}

/**
 * Returns a valid access token for Microsoft Graph.
 * Performs Device Code Flow on first run or if token is expired.
 */
export async function getAccessToken() {
  const cached = await loadToken();

  if (cached && !isTokenExpired(cached)) {
    return cached.accessToken;
  }

  // Need to authenticate
  const credential = getCredential();
  const scopes = _config.scopes || [
    'https://graph.microsoft.com/User.ReadWrite.All',
    'https://graph.microsoft.com/Directory.ReadWrite.All',
  ];

  // Normalize scopes to full URIs if needed
  const normalizedScopes = scopes.map((s) =>
    s.startsWith('https://') ? s : `https://graph.microsoft.com/${s}`
  );

  console.log(chalk.yellow('\nNo valid session found. Starting authentication...\n'));

  const tokenResponse = await credential.getToken(normalizedScopes);

  await saveToken({
    accessToken: tokenResponse.token,
    expiresAt: tokenResponse.expiresOnTimestamp,
  });

  console.log(chalk.green('\nAuthentication successful. Session saved.\n'));

  return tokenResponse.token;
}

function getCredential() {
  if (_credential) return _credential;

  if (!_config) {
    throw new Error('Auth not initialized. Call initAuth(config) first.');
  }

  _credential = new DeviceCodeCredential({
    tenantId: _config.tenantId,
    clientId: _config.clientId,
    userPromptCallback: (info) => {
      console.log(chalk.cyan('\n========================================'));
      console.log(chalk.bold('  Microsoft 365 Authentication'));
      console.log(chalk.cyan('========================================'));
      console.log(`\n  1. Open your browser and go to:`);
      console.log(chalk.underline.blue(`     ${info.verificationUri}`));
      console.log(`\n  2. Enter the code:`);
      console.log(chalk.bold.yellow(`     ${info.userCode}`));
      console.log(`\n  3. Sign in with your Microsoft 365 account`);
      console.log(chalk.cyan('\n========================================\n'));
    },
  });

  return _credential;
}

/**
 * Force re-authentication by clearing cached token and getting a new one.
 */
export async function login() {
  await clearToken();
  _credential = null; // Reset credential to force new device code prompt
  return getAccessToken();
}

/**
 * Clear saved session token.
 */
export async function logout() {
  await clearToken();
  _credential = null;
  console.log(chalk.green('Logged out. Session token cleared.'));
}
