import test from 'ava';
import { cleanupHome } from './helpers/home.js';
import { mockGraph, mockPrompts, runCommand, writes } from './helpers/graph-mock.js';
import { writeWorkbook, dmy } from './helpers/excel.js';

const { initAuth, _setFakeToken, _clearFakeToken } = await import('../src/auth.js');
const { syncCheckCommand } = await import('../src/commands/sync-check.js');
const { DOMAIN } = await import('../src/constants.js');

test.before(() => {
  initAuth({ tenantId: 't', clientId: 'c', domain: DOMAIN });
  _setFakeToken('fake-token');
});

test.after.always(() => {
  _clearFakeToken();
  cleanupHome();
});

const HEADERS = ['Trabajador', 'e_mail', 'Fecha de Baja'];
const at = (name) => `${name}@${DOMAIN}`;
const cloudUser = (name, extra = {}) => ({
  id: `id-${name}`,
  userPrincipalName: at(name),
  displayName: name,
  accountEnabled: true,
  ...extra,
});

/**
 * Fake tenant: `users` is what GET /users returns; `licenses` maps UPN → SKU list.
 * `failOn(call)` can make individual writes fail.
 */
function tenant(t, { users, licenses = {}, failOn = () => false }) {
  return mockGraph(t, (call) => {
    const { method, path } = call;
    if (failOn(call)) {
      return { status: 403, body: { error: { code: 'Authorization_RequestDenied', message: 'Insufficient privileges' } } };
    }
    if (method === 'GET' && path.startsWith('/users?')) return { body: { value: users } };
    const lic = path.match(/^\/users\/([^/]+)\/licenseDetails/);
    if (method === 'GET' && lic) return { body: { value: licenses[lic[1]] ?? [] } };
    if (method === 'POST' && path.endsWith('/assignLicense')) return { body: {} };
    if (method === 'PATCH' && path.startsWith('/users/')) return { status: 204 };
    return undefined;
  });
}

/** Text of one report section: from its "● Check …" heading to the next one. */
function section(output, heading) {
  const start = output.indexOf(heading);
  if (start === -1) return '';
  const next = output.indexOf('●', start + heading.length);
  return output.slice(start, next === -1 ? undefined : next);
}

// ---------------------------------------------------------------------------
// Read-only report
// ---------------------------------------------------------------------------

test.serial('sync-check: Excel and tenant agree → exit 0, no writes', async (t) => {
  const file = await writeWorkbook({
    ACTIVOS: [HEADERS, ['ANA', at('ana'), ''], ['LUIS', at('luis'), '']],
  });
  const calls = tenant(t, { users: [cloudUser('ana'), cloudUser('luis')] });

  const { exitCode, output } = await runCommand(() => syncCheckCommand(file));

  t.is(exitCode, 0);
  t.true(output.includes('Everything is in sync'));
  t.deepEqual(writes(calls), []);
});

test.serial('sync-check: each kind of mismatch lands in its own check', async (t) => {
  const file = await writeWorkbook({
    ACTIVOS: [
      HEADERS,
      ['OK ACTIVE',        at('ok'),          ''],
      ['MISSING ACTIVE',   at('missing'),     ''],
      ['LEFT STILL ON',    at('left'),        dmy(-30)],
      ['LEAVES TODAY',     at('today'),       dmy(0)],
      ['LEAVES NEXT WEEK', at('future'),      dmy(7)],
      ['GMAIL PERSON',     'someone@gmail.com', ''],
      ['NO EMAIL PERSON',  '',                ''],
      ['UPPERCASE MAIL',   at('MAYUS').toUpperCase(), ''],
    ],
    CONCEJALES: [['nombre', 'email'], ['CONCEJAL UNO', at('concejal')]],
  });
  const calls = tenant(t, {
    users: [
      cloudUser('ok'),
      cloudUser('left'),
      cloudUser('today'),
      cloudUser('future'),
      cloudUser('mayus'),
      cloudUser('concejal'),
      cloudUser('stranger'),
      cloudUser('sala-plenos'),
      cloudUser('salapz1'),
      { id: 'id-other', userPrincipalName: 'someone@otherdomain.com', displayName: 'other domain', accountEnabled: true },
    ],
  });

  const { exitCode, output } = await runCommand(() => syncCheckCommand(file));

  t.is(exitCode, 1);
  t.deepEqual(writes(calls), [], 'the report never writes without --disable-left-workers');

  const check1a = section(output, '● Check 1a');
  t.true(check1a.includes('MISSING ACTIVE'));
  t.false(check1a.includes('OK ACTIVE'));
  t.false(check1a.includes('UPPERCASE MAIL'), 'email comparison is case-insensitive');
  t.false(check1a.includes('LEAVES NEXT WEEK'), 'a future Fecha de Baja still counts as active');

  const check1b = section(output, '● Check 1b');
  t.true(check1b.includes('GMAIL PERSON'));
  t.true(check1b.includes('NO EMAIL PERSON'));
  t.true(check1b.includes('(empty)'));

  const check2 = section(output, '● Check 2');
  t.true(check2.includes('LEFT STILL ON'));
  t.true(check2.includes('LEAVES TODAY'), 'a Fecha de Baja of today is already effective');
  t.false(check2.includes('LEAVES NEXT WEEK'));

  const check3 = section(output, '● Check 3');
  t.true(check3.includes(at('stranger')));
  t.false(check3.includes(at('concejal')), 'CONCEJALES sheet counts as "present in Excel"');
  t.false(check3.includes('sala-plenos'), 'room accounts are ignored');
  t.false(check3.includes('salapz1'), 'room accounts are ignored');
  t.false(check3.includes('otherdomain.com'), 'other domains are out of scope');

  // 1 missing + 2 no-email + 2 not disabled + 1 cloud-only
  t.true(output.includes('6 sync issue(s) found'));
});

test.serial('sync-check: a left employee with no cloud account is not an issue', async (t) => {
  const file = await writeWorkbook({
    ACTIVOS: [HEADERS, ['GONE AND DELETED', at('gone'), dmy(-400)]],
  });
  tenant(t, { users: [] });

  const { exitCode } = await runCommand(() => syncCheckCommand(file));
  t.is(exitCode, 0);
});

test.serial('sync-check: works without a CONCEJALES sheet', async (t) => {
  const file = await writeWorkbook({ ACTIVOS: [HEADERS, ['ANA', at('ana'), '']] });
  tenant(t, { users: [cloudUser('ana')] });

  const { exitCode, output } = await runCommand(() => syncCheckCommand(file));
  t.is(exitCode, 0);
  t.true(output.includes('0 concejales found'));
});

test.serial('sync-check: workbook without an ACTIVOS sheet → exit 1 before calling Graph', async (t) => {
  const file = await writeWorkbook({ Hoja1: [HEADERS] });
  const calls = tenant(t, { users: [] });

  const { exitCode, output } = await runCommand(() => syncCheckCommand(file));

  t.is(exitCode, 1);
  t.true(output.includes('Worksheet "ACTIVOS" not found'));
  t.is(calls.length, 0);
});

test.serial('sync-check: Graph failure while listing users → exit 1', async (t) => {
  const file = await writeWorkbook({ ACTIVOS: [HEADERS, ['ANA', at('ana'), '']] });
  tenant(t, { users: [], failOn: () => true });

  const { exitCode, output } = await runCommand(() => syncCheckCommand(file));
  t.is(exitCode, 1);
  t.true(output.includes('Failed to fetch users from M365'));
});

// ---------------------------------------------------------------------------
// --disable-left-workers
// ---------------------------------------------------------------------------

async function leftWorkersFixture() {
  return writeWorkbook({
    ACTIVOS: [
      HEADERS,
      ['STAYS',           at('stays'),    ''],
      ['LEFT LICENSED',   at('licensed'), dmy(-10)],
      ['LEFT UNLICENSED', at('bare'),     dmy(-10)],
      ['LEAVES LATER',    at('later'),    dmy(10)],
    ],
  });
}

const LEFT_USERS = [cloudUser('stays'), cloudUser('licensed'), cloudUser('bare'), cloudUser('later'), cloudUser('stranger')];
const LEFT_LICENSES = {
  [at('licensed')]: [
    { skuId: 'sku-1', skuPartNumber: 'O365_BUSINESS_PREMIUM' },
    { skuId: 'sku-2', skuPartNumber: 'EXCHANGESTANDARD' },
  ],
};

test.serial('sync-check --disable-left-workers: both confirmations → unlicense, then disable, only Check 2 accounts', async (t) => {
  const file = await leftWorkersFixture();
  const calls = tenant(t, { users: LEFT_USERS, licenses: LEFT_LICENSES });
  mockPrompts(t, [{ confirm1: true }, { confirm2: true }]);

  const { exitCode, output } = await runCommand(() => syncCheckCommand(file, { fix: true }));

  t.is(exitCode, 0);
  t.deepEqual(writes(calls).map((c) => [c.method, c.path, c.body]), [
    // Licences go first: some tenants refuse licence changes on a disabled account
    ['POST',  `/users/${at('licensed')}/assignLicense`, { addLicenses: [], removeLicenses: ['sku-1', 'sku-2'] }],
    ['PATCH', `/users/${at('licensed')}`, { accountEnabled: false }],
    ['PATCH', `/users/${at('bare')}`, { accountEnabled: false }],
  ]);
  t.true(output.includes('2 account(s) disabled and unlicensed successfully'));
  t.true(output.includes('Total: 2 licence assignment(s) across 2 account(s)'));
});

test.serial('sync-check --disable-left-workers: declining the first confirmation changes nothing', async (t) => {
  const file = await leftWorkersFixture();
  const calls = tenant(t, { users: LEFT_USERS, licenses: LEFT_LICENSES });
  mockPrompts(t, [{ confirm1: false }]);

  const { exitCode, output } = await runCommand(() => syncCheckCommand(file, { fix: true }));

  t.is(exitCode, 1);
  t.deepEqual(writes(calls), []);
  t.true(output.includes('Aborted — no changes made'));
});

test.serial('sync-check --disable-left-workers: declining the second confirmation changes nothing', async (t) => {
  const file = await leftWorkersFixture();
  const calls = tenant(t, { users: LEFT_USERS, licenses: LEFT_LICENSES });
  mockPrompts(t, [{ confirm1: true }, { confirm2: false }]);

  const { exitCode } = await runCommand(() => syncCheckCommand(file, { fix: true }));

  t.is(exitCode, 1);
  t.deepEqual(writes(calls), []);
});

test.serial('sync-check --disable-left-workers: a failed licence removal leaves that account enabled and exits 1', async (t) => {
  const file = await leftWorkersFixture();
  const calls = tenant(t, {
    users: LEFT_USERS,
    licenses: LEFT_LICENSES,
    failOn: ({ method, path }) => method === 'POST' && path.endsWith('/assignLicense'),
  });
  mockPrompts(t, [{ confirm1: true }, { confirm2: true }]);

  const { exitCode, output } = await runCommand(() => syncCheckCommand(file, { fix: true }));

  t.is(exitCode, 1);
  const patched = writes(calls).filter((c) => c.method === 'PATCH').map((c) => c.path);
  t.deepEqual(patched, [`/users/${at('bare')}`], 'the account whose licences could not be removed is not disabled');
  t.true(output.includes('1 succeeded'));
  t.true(output.includes('1 failed'));
});

test.serial('sync-check --disable-left-workers: nothing to disable → no prompts, report only', async (t) => {
  const file = await writeWorkbook({ ACTIVOS: [HEADERS, ['ANA', at('ana'), '']] });
  const calls = tenant(t, { users: [cloudUser('ana'), cloudUser('stranger')] });
  mockPrompts(t, []); // any prompt would throw

  const { exitCode } = await runCommand(() => syncCheckCommand(file, { fix: true }));

  t.is(exitCode, 1, 'the cloud-only account is still an issue');
  t.deepEqual(writes(calls), []);
});
