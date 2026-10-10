import test from 'ava';
import { cleanupHome } from './helpers/home.js';
import { mockGraph, mockPrompts, runCommand, writes } from './helpers/graph-mock.js';

const { initAuth, _setFakeToken, _clearFakeToken } = await import('../src/auth.js');
const { generatePassword, resetPasswordCommand } = await import('../src/commands/reset-password.js');

test.before(() => {
  initAuth({ tenantId: 't', clientId: 'c', domain: 'contoso.com' });
  _setFakeToken('fake-token');
});

test.after.always(() => {
  _clearFakeToken();
  cleanupHome();
});

// ---------------------------------------------------------------------------
// generatePassword
// ---------------------------------------------------------------------------

test('generatePassword: has the Word-NNNN-Word shape', (t) => {
  for (let i = 0; i < 200; i++) {
    t.regex(generatePassword(), /^\p{Lu}\p{Ll}+-\d{4}-\p{Lu}\p{Ll}+$/u);
  }
});

test('generatePassword: meets Entra ID complexity (upper, lower, digit, 8+ chars)', (t) => {
  for (let i = 0; i < 200; i++) {
    const pw = generatePassword();
    t.true(pw.length >= 8, pw);
    t.regex(pw, /[A-Z]/);
    t.regex(pw, /[a-z]/);
    t.regex(pw, /\d/);
  }
});

test('generatePassword: does not return the same value every time', (t) => {
  const seen = new Set(Array.from({ length: 50 }, () => generatePassword()));
  t.true(seen.size > 40);
});

// ---------------------------------------------------------------------------
// resetPasswordCommand
// ---------------------------------------------------------------------------

const USER = {
  id: 'id-ana',
  userPrincipalName: 'ana@contoso.com',
  displayName: 'Ana Pérez',
  jobTitle: 'Técnico',
};

function graphWithUser(t, upn = USER.userPrincipalName, user = USER) {
  return mockGraph(t, ({ method, path }) => {
    if (method === 'GET' && path.startsWith(`/users/${upn}`)) return { body: user };
    if (method === 'PATCH' && path === `/users/${upn}`) return { status: 204 };
    return undefined;
  });
}

test.serial('reset-password: confirmed → PATCHes a temporary password that must be changed', async (t) => {
  const calls = graphWithUser(t);
  mockPrompts(t, [{ confirmed: true }]);

  const { exitCode, output } = await runCommand(() => resetPasswordCommand('ana@contoso.com'));

  t.is(exitCode, null);
  const patches = writes(calls);
  t.is(patches.length, 1);
  t.is(patches[0].method, 'PATCH');
  const { passwordProfile } = patches[0].body;
  t.true(passwordProfile.forceChangePasswordNextSignIn);
  t.regex(passwordProfile.password, /^\p{L}+-\d{4}-\p{L}+$/u);

  // The password shown to the operator is the one that was sent
  t.true(output.includes(passwordProfile.password));
  t.true(output.includes(`mailto:${encodeURIComponent('ana@contoso.com')}`));
  t.true(output.includes(encodeURIComponent(`password: ${passwordProfile.password}`)));
});

test.serial('reset-password: declined → nothing is written', async (t) => {
  const upn = 'declined@contoso.com';
  const calls = graphWithUser(t, upn, { ...USER, id: 'id-declined', userPrincipalName: upn });
  mockPrompts(t, [{ confirmed: false }]);

  const { exitCode, output } = await runCommand(() => resetPasswordCommand(upn));

  t.is(exitCode, null);
  t.deepEqual(writes(calls), []);
  t.true(output.includes('No changes made'));
});

test.serial('reset-password: unknown user → exit 1 without prompting', async (t) => {
  const calls = mockGraph(t, () => undefined); // everything 404s
  mockPrompts(t, []); // any prompt would throw

  const { exitCode, output } = await runCommand(() => resetPasswordCommand('ghost@contoso.com'));

  t.is(exitCode, 1);
  t.deepEqual(writes(calls), []);
  t.true(output.includes('User not found: ghost@contoso.com'));
});

test.serial('reset-password: Graph rejects the change → exit 1', async (t) => {
  const upn = 'denied@contoso.com';
  const calls = mockGraph(t, ({ method, path }) => {
    if (method === 'GET' && path.startsWith(`/users/${upn}`)) {
      return { body: { ...USER, id: 'id-denied', userPrincipalName: upn } };
    }
    if (method === 'PATCH') {
      return { status: 403, body: { error: { code: 'Authorization_RequestDenied', message: 'Insufficient privileges' } } };
    }
    return undefined;
  });
  mockPrompts(t, [{ confirmed: true }]);

  const { exitCode, output } = await runCommand(() => resetPasswordCommand(upn));

  t.is(exitCode, 1);
  t.is(writes(calls).length, 1);
  t.true(output.includes('Insufficient privileges'));
  t.false(output.includes('mailto:'), 'no credentials link when the reset failed');
});
