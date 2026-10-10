import test from 'ava';
import { cleanupHome } from './helpers/home.js';
import { mockGraph, batchResponse, runCommand, writes } from './helpers/graph-mock.js';

const { initAuth, _setFakeToken, _clearFakeToken } = await import('../src/auth.js');
const { listCommand } = await import('../src/commands/list.js');

test.before(() => {
  initAuth({ tenantId: 't', clientId: 'c', domain: 'contoso.com' });
  _setFakeToken('fake-token');
});

test.after.always(() => {
  _clearFakeToken();
  cleanupHome();
});

const USERS = [
  {
    id: 'ls-full', userPrincipalName: 'full@contoso.com', displayName: 'Full Profile',
    jobTitle: 'Arquitecta', department: 'Urbanismo',
    signInActivity: { lastSignInDateTime: '2026-01-10T08:00:00Z' },
  },
  {
    id: 'ls-nomgr', userPrincipalName: 'nomgr@contoso.com', displayName: 'No Manager',
    jobTitle: 'Técnico', department: 'Cultura',
    signInActivity: { lastSignInDateTime: '2026-02-01T08:00:00Z' },
  },
  {
    id: 'ls-bare', userPrincipalName: 'bare@contoso.com', displayName: 'Bare Account',
    jobTitle: '', department: '  ',
  },
];

const LICENSED = new Set(['ls-full', 'ls-bare', 'full@contoso.com', 'bare@contoso.com']);

function tenant(t) {
  return mockGraph(t, (call) => {
    const { method, path } = call;
    if (path.endsWith('/$batch')) {
      return batchResponse(call, (r) => {
        const id = r.url.split('/')[2];
        if (r.url.includes('/memberOf/')) {
          return {
            body: {
              value: id === 'ls-full'
                ? [{ id: 'g1', displayName: 'Urbanismo-Equipo' }, { id: 'g2', displayName: 'Todos los usuarios' }]
                : [],
            },
          };
        }
        return { body: { value: LICENSED.has(id) ? [{ skuId: 's', skuPartNumber: 'O365_BUSINESS_PREMIUM' }] : [] } };
      });
    }
    if (method === 'GET' && path.startsWith('/users?')) return { body: { value: USERS } };
    const mgr = path.match(/^\/users\/([^/]+)\/manager/);
    if (mgr) {
      return mgr[1] === 'ls-full'
        ? { body: { userPrincipalName: 'boss@contoso.com', displayName: 'The Boss' } }
        : undefined;
    }
    const lic = path.match(/^\/users\/([^/]+)\/licenseDetails/);
    if (lic) return { body: { value: LICENSED.has(lic[1]) ? [{ skuId: 's' }] : [] } };
    return undefined;
  });
}

const listed = (output) => USERS.filter((u) => output.includes(u.userPrincipalName)).map((u) => u.id);

test.serial('list: shows every user, hides auto-assigned groups, never writes', async (t) => {
  const calls = tenant(t);
  const { output } = await runCommand(() => listCommand());

  t.deepEqual(listed(output), ['ls-full', 'ls-nomgr', 'ls-bare']);
  t.true(output.includes('Urbanismo-Equipo'));
  t.false(output.includes('Todos los usuarios'));
  t.true(output.includes('Found 3 user(s) out of 3 active total'));
  t.deepEqual(writes(calls), []);
});

for (const [flag, expected] of [
  ['noManager',     ['ls-nomgr', 'ls-bare']],
  ['noDepartment',  ['ls-bare']],
  ['noJobTitle',    ['ls-bare']],
  ['neverSignedIn', ['ls-bare']],
]) {
  test.serial(`list --${flag}: keeps only matching users`, async (t) => {
    tenant(t);
    const { output } = await runCommand(() => listCommand({ [flag]: true }));
    t.deepEqual(listed(output), expected);
  });
}

test.serial('list: combined filters are ANDed', async (t) => {
  tenant(t);
  const { output } = await runCommand(() => listCommand({ noManager: true, noJobTitle: true }));
  t.deepEqual(listed(output), ['ls-bare']);
  t.true(output.includes('[filters: no manager, no job title]'));
});

test.serial('list --disabled: asks Graph for disabled accounts', async (t) => {
  const calls = tenant(t);
  await runCommand(() => listCommand({ disabled: true }));
  t.true(calls[0].url.includes('accountEnabled eq false'));
});

test.serial('list --for-mailing: prints licensed UPNs joined by ";"', async (t) => {
  tenant(t);
  const { output } = await runCommand(() => listCommand({ forMailing: true }));
  t.true(output.includes('full@contoso.com;bare@contoso.com'));
  t.false(output.includes('nomgr@contoso.com'));
});

test.serial('list: no user matches the filters → message instead of an empty table', async (t) => {
  mockGraph(t, (call) => {
    if (call.path.endsWith('/$batch')) return batchResponse(call, () => ({ body: { value: [] } }));
    return { body: { value: [] } };
  });
  const { output } = await runCommand(() => listCommand({ noManager: true }));
  t.true(output.includes('No users match the specified filters'));
});
