import test from 'ava';
import { cleanupHome } from './helpers/home.js';
import { mockGraph, runCommand, writes } from './helpers/graph-mock.js';
import { writeWorkbook } from './helpers/excel.js';

const { initAuth, _setFakeToken, _clearFakeToken } = await import('../src/auth.js');
const { importUsers } = await import('../src/commands/import.js');

const DOMAIN = 'contoso.com';

test.before(() => {
  initAuth({ tenantId: 't', clientId: 'c', domain: DOMAIN });
  _setFakeToken('fake-token');
});

test.after.always(() => {
  _clearFakeToken();
  cleanupHome();
});

// The real employee file: Spanish headers on the ACTIVOS sheet.
const HEADERS = ['id_empleado', 'Trabajador', 'e_mail', 'Departamento', 'Descripción', 'ID responsable'];

/**
 * Fake tenant for the import flow.
 * @param {object[]} users      - existing accounts (returned by GET /users)
 * @param {object}   managers   - user id → current manager object
 * @param {Function} failOn     - make matching requests fail with 400
 */
function tenant(t, { users = [], managers = {}, failOn = () => false } = {}) {
  let created = 0;
  return mockGraph(t, (call) => {
    const { method, path } = call;
    if (failOn(call)) {
      return { status: 400, body: { error: { code: 'Request_BadRequest', message: 'Property is invalid' } } };
    }
    if (method === 'GET' && path.startsWith('/users?')) {
      const byUpn = path.match(/userPrincipalName eq '([^']+)'/);
      if (!byUpn) return { body: { value: users } };
      return { body: { value: users.filter((u) => u.userPrincipalName === byUpn[1]) } };
    }
    if (method === 'POST' && path === '/users') {
      return { status: 201, body: { id: `id-created-${++created}`, ...call.body } };
    }
    const mgr = path.match(/^\/users\/([^/]+)\/manager/);
    if (method === 'GET' && mgr) return managers[mgr[1]] ? { body: managers[mgr[1]] } : undefined;
    if (method === 'PUT' && mgr) return { status: 204 };
    if (method === 'GET' && path.includes('/licenseDetails')) return { body: { value: [{ skuId: 's', skuPartNumber: 'P' }] } };
    if (method === 'PATCH' && path.startsWith('/users/')) return { status: 204 };
    return undefined;
  });
}

const boss = (n) => ({
  id: `id-boss-${n}`,
  userPrincipalName: `boss${n}@${DOMAIN}`,
  displayName: 'The Boss',
  employeeId: String(9000 + n),
});

// ---------------------------------------------------------------------------
// Create
// ---------------------------------------------------------------------------

test.serial('import: new employee → created enabled, with a generated one-time password and normalised name', async (t) => {
  const file = await writeWorkbook({
    ACTIVOS: [HEADERS, [101, 'GARCÍA LÓPEZ, MARÍA JOSÉ', `mgarcia@${DOMAIN}`, 'URBANISMO', 'Arquitecta', 9001]],
  });
  const calls = tenant(t, { users: [boss(1)] });

  const { output } = await runCommand(() => importUsers(file));

  const [create, setManager, ...rest] = writes(calls);
  t.deepEqual(rest, []);

  t.deepEqual([create.method, create.path], ['POST', '/users']);
  const { passwordProfile, ...profile } = create.body;
  t.deepEqual(profile, {
    userPrincipalName: `mgarcia@${DOMAIN}`,
    displayName: 'María José García López',
    givenName: 'María José',
    surname: 'García López',
    mailNickname: 'mgarcia',
    jobTitle: 'Arquitecta',
    department: 'URBANISMO',
    employeeId: '101',
    accountEnabled: true,
  });
  t.regex(passwordProfile.password, /^\p{L}+-\d{4}-\p{L}+$/u);
  t.true(passwordProfile.forceChangePasswordNextSignIn);

  // "ID responsable" is the manager's employee number, resolved to their object ID
  t.deepEqual([setManager.method, setManager.path, setManager.body], [
    'PUT',
    '/users/id-created-1/manager/$ref',
    { '@odata.id': 'https://graph.microsoft.com/v1.0/users/id-boss-1' },
  ]);

  t.true(output.includes(`[CREATE] mgarcia@${DOMAIN}`));
  t.true(output.includes('Created: 1'));
});

test.serial('import: a short username in e_mail gets the tenant domain appended', async (t) => {
  const file = await writeWorkbook({
    ACTIVOS: [HEADERS, [102, 'RUIZ, PEDRO', 'pruiz', 'OBRAS', 'Peón', '']],
  });
  const calls = tenant(t);

  await runCommand(() => importUsers(file));

  const [create] = writes(calls);
  t.is(create.body.userPrincipalName, `pruiz@${DOMAIN}`);
  t.is(create.body.mailNickname, 'pruiz');
});

// ---------------------------------------------------------------------------
// Update / no-op
// ---------------------------------------------------------------------------

test.serial('import: existing employee with identical data → no writes at all', async (t) => {
  const existing = {
    id: 'id-same',
    userPrincipalName: `same@${DOMAIN}`,
    displayName: 'Ana Pérez',
    givenName: 'Ana',
    surname: 'Pérez',
    jobTitle: 'Técnico',
    department: 'CULTURA',
    employeeId: '201',
  };
  const file = await writeWorkbook({
    ACTIVOS: [HEADERS, [201, 'PÉREZ, ANA', existing.userPrincipalName, 'CULTURA', 'Técnico', 9002]],
  });
  const calls = tenant(t, {
    users: [existing, boss(2)],
    managers: { 'id-same': { id: 'id-boss-2', userPrincipalName: `boss2@${DOMAIN}` } },
  });

  const { output } = await runCommand(() => importUsers(file));

  t.deepEqual(writes(calls), []);
  t.true(output.includes(`[IS-READY] same@${DOMAIN}`));
  t.true(output.includes('Updated: 0'));
});

test.serial('import: existing employee with a changed field → PATCH without password or mailNickname', async (t) => {
  const existing = {
    id: 'id-moved',
    userPrincipalName: `moved@${DOMAIN}`,
    displayName: 'Luis Sanz',
    givenName: 'Luis',
    surname: 'Sanz',
    jobTitle: 'Administrativo',
    department: 'CULTURA',
    employeeId: '202',
  };
  const file = await writeWorkbook({
    ACTIVOS: [HEADERS, [202, 'SANZ, LUIS', existing.userPrincipalName, 'DEPORTES', 'Administrativo', '']],
  });
  const calls = tenant(t, { users: [existing] });

  const { output } = await runCommand(() => importUsers(file));

  const [patch, ...rest] = writes(calls);
  t.deepEqual(rest, []);
  t.deepEqual([patch.method, patch.path], ['PATCH', `/users/moved@${DOMAIN}`]);
  t.is(patch.body.department, 'DEPORTES');
  t.false('passwordProfile' in patch.body, 'an import must never reset an existing password');
  t.false('mailNickname' in patch.body);

  t.true(output.includes('Change department: CULTURA -> DEPORTES'));
  t.true(output.includes('Updated: 1'));
});

test.serial('import: employeeId stored as a number in the cloud is not a change', async (t) => {
  const existing = {
    id: 'id-numeric',
    userPrincipalName: `numeric@${DOMAIN}`,
    displayName: 'Eva Gil',
    givenName: 'Eva',
    surname: 'Gil',
    jobTitle: 'Técnico',
    department: 'CULTURA',
    employeeId: 203,
  };
  const file = await writeWorkbook({
    ACTIVOS: [HEADERS, [203, 'GIL, EVA', existing.userPrincipalName, 'CULTURA', 'Técnico', '']],
  });
  const calls = tenant(t, { users: [existing] });

  await runCommand(() => importUsers(file));
  t.deepEqual(writes(calls), []);
});

// "Fecha de Alta" → employeeHireDate
const HIRE_HEADERS = [...HEADERS, 'Fecha de Alta'];
const hired = (n, employeeHireDate) => ({
  id: `id-hired-${n}`,
  userPrincipalName: `hired${n}@${DOMAIN}`,
  displayName: 'Sara Rey',
  givenName: 'Sara',
  surname: 'Rey',
  jobTitle: 'Técnico',
  department: 'CULTURA',
  employeeId: String(250 + n),
  ...(employeeHireDate === undefined ? {} : { employeeHireDate }),
});
const hiredRow = (n, fecha) => [250 + n, 'REY, SARA', `hired${n}@${DOMAIN}`, 'CULTURA', 'Técnico', '', fecha];

test.serial('import: hire date already matching the cloud value → no write', async (t) => {
  const file = await writeWorkbook({ ACTIVOS: [HIRE_HEADERS, hiredRow(1, '05/03/2020')] });
  // Graph serialises without milliseconds
  const calls = tenant(t, { users: [hired(1, '2020-03-05T00:00:00Z')] });

  const { output } = await runCommand(() => importUsers(file));

  t.deepEqual(writes(calls), []);
  t.true(output.includes(`[IS-READY] hired1@${DOMAIN}`));
});

test.serial('import: hire date missing in the cloud → PATCH with the Excel date', async (t) => {
  const file = await writeWorkbook({ ACTIVOS: [HIRE_HEADERS, hiredRow(2, '05/03/2020')] });
  const calls = tenant(t, { users: [hired(2, null)] });

  const { output } = await runCommand(() => importUsers(file));

  const [patch, ...rest] = writes(calls);
  t.deepEqual(rest, []);
  t.is(patch.method, 'PATCH');
  t.is(patch.body.employeeHireDate, '2020-03-05T00:00:00.000Z');
  t.true(output.includes('Change employeeHireDate: (null) -> 2020-03-05T00:00:00.000Z'));
});

test.serial('import: a day/month-swapped hire date left by an older import is corrected', async (t) => {
  const file = await writeWorkbook({ ACTIVOS: [HIRE_HEADERS, hiredRow(3, '05/03/2020')] });
  const calls = tenant(t, { users: [hired(3, '2020-05-02T22:00:00Z')] });

  await runCommand(() => importUsers(file));

  const [patch] = writes(calls);
  t.is(patch.body.employeeHireDate, '2020-03-05T00:00:00.000Z');
});

test.serial('import: an empty or invalid Fecha de Alta never clears the cloud hire date', async (t) => {
  const file = await writeWorkbook({
    ACTIVOS: [HIRE_HEADERS, hiredRow(4, ''), hiredRow(5, '31/02/2020')],
  });
  const calls = tenant(t, { users: [hired(4, '2019-01-01T00:00:00Z'), hired(5, '2019-01-01T00:00:00Z')] });

  await runCommand(() => importUsers(file));
  t.deepEqual(writes(calls), []);
});

test.serial('import: asks Graph for employeeHireDate when loading existing users', async (t) => {
  const file = await writeWorkbook({ ACTIVOS: [HIRE_HEADERS, hiredRow(6, '05/03/2020')] });
  const calls = tenant(t, { users: [hired(6, '2020-03-05T00:00:00Z')] });

  await runCommand(() => importUsers(file));
  t.true(calls[0].url.includes('employeeHireDate'));
});

// ---------------------------------------------------------------------------
// Manager resolution
// ---------------------------------------------------------------------------

test.serial('import: manager already correct → no manager write', async (t) => {
  const file = await writeWorkbook({
    ACTIVOS: [HEADERS, [301, 'MORA, INÉS', `imora@${DOMAIN}`, 'OBRAS', 'Jefa', 9003]],
  });
  const calls = tenant(t, {
    users: [boss(3)],
    managers: { 'id-created-1': { id: 'id-boss-3', userPrincipalName: `boss3@${DOMAIN}` } },
  });

  await runCommand(() => importUsers(file));
  t.deepEqual(writes(calls).map((c) => c.method), ['POST']);
});

test.serial('import: unknown manager employee number → reported, user still imported, no manager write', async (t) => {
  const file = await writeWorkbook({
    ACTIVOS: [HEADERS, [302, 'VEGA, JUAN', `jvega@${DOMAIN}`, 'OBRAS', 'Oficial', 7777]],
  });
  const calls = tenant(t);

  const { output } = await runCommand(() => importUsers(file));

  t.deepEqual(writes(calls).map((c) => c.method), ['POST']);
  t.true(output.includes('Could not find manager with EmployeeID 7777'));
  t.true(output.includes('Manager errors: 1'));
});

test.serial('import: a non-numeric manager is treated as a username in the tenant domain', async (t) => {
  const mgr = { id: 'id-named-boss', userPrincipalName: `jefa@${DOMAIN}`, displayName: 'Jefa' };
  const file = await writeWorkbook({
    ACTIVOS: [HEADERS, [303, 'LEAL, RAÚL', `rleal@${DOMAIN}`, 'OBRAS', 'Oficial', 'Jefa']],
  });
  const calls = tenant(t, { users: [mgr] });

  await runCommand(() => importUsers(file));

  const put = writes(calls).find((c) => c.method === 'PUT');
  t.is(put.body['@odata.id'], 'https://graph.microsoft.com/v1.0/users/id-named-boss');
});

// ---------------------------------------------------------------------------
// Validation, limits and failures
// ---------------------------------------------------------------------------

test.serial('import: rows failing validation are reported and skipped; valid rows still go through', async (t) => {
  const file = await writeWorkbook({
    ACTIVOS: [
      HEADERS,
      [401, 'SIN CORREO, PEPE', '', 'OBRAS', 'Peón', ''],
      [402, 'OTRO DOMINIO, PEPA', 'pepa@gmail.com', 'OBRAS', 'Peón', ''],
      [403, 'BUENO, PACO', `pbueno@${DOMAIN}`, 'OBRAS', 'Peón', ''],
    ],
  });
  const calls = tenant(t);

  const { output } = await runCommand(() => importUsers(file));

  t.deepEqual(writes(calls).map((c) => c.body.userPrincipalName), [`pbueno@${DOMAIN}`]);
  t.true(output.includes('Found 2 row(s) with validation errors'));
  t.true(output.includes('Row 2 ((unknown)): Missing required field: userPrincipalName'));
  t.true(output.includes(`Row 3 (pepa@gmail.com): UPN must end with @${DOMAIN}`));
});

test.serial('import: no valid rows → nothing is written', async (t) => {
  const file = await writeWorkbook({ ACTIVOS: [HEADERS, [501, 'SIN CORREO, PEPE', '', '', '', '']] });
  const calls = tenant(t);

  const { output } = await runCommand(() => importUsers(file));

  t.deepEqual(writes(calls), []);
  t.true(output.includes('No valid rows to process'));
});

test.serial('import --limit: stops after N rows', async (t) => {
  const file = await writeWorkbook({
    ACTIVOS: [
      HEADERS,
      [601, 'UNO, A', `uno@${DOMAIN}`, 'X', 'Y', ''],
      [602, 'DOS, B', `dos@${DOMAIN}`, 'X', 'Y', ''],
      [603, 'TRES, C', `tres@${DOMAIN}`, 'X', 'Y', ''],
    ],
  });
  const calls = tenant(t);

  const { output } = await runCommand(() => importUsers(file, { limit: 2 }));

  t.deepEqual(writes(calls).map((c) => c.body.userPrincipalName), [`uno@${DOMAIN}`, `dos@${DOMAIN}`]);
  t.true(output.includes('Reached limit of 2 users'));
});

test.serial('import: one failing user does not stop the rest', async (t) => {
  const file = await writeWorkbook({
    ACTIVOS: [
      HEADERS,
      [701, 'MAL, A', `bad@${DOMAIN}`, 'X', 'Y', ''],
      [702, 'BIEN, B', `good@${DOMAIN}`, 'X', 'Y', ''],
    ],
  });
  const calls = tenant(t, {
    failOn: ({ method, body }) => method === 'POST' && body?.userPrincipalName === `bad@${DOMAIN}`,
  });

  const { output } = await runCommand(() => importUsers(file));

  t.is(writes(calls).length, 2);
  t.true(output.includes(`[FAILED] bad@${DOMAIN}: Property is invalid`));
  t.true(output.includes(`[CREATE] good@${DOMAIN}`));
  t.true(output.includes('Created: 1'));
  t.true(output.includes('Failed:  1'));
});

test.serial('import: an account without licences triggers a warning', async (t) => {
  const file = await writeWorkbook({ ACTIVOS: [HEADERS, [801, 'LIBRE, A', `nolic@${DOMAIN}`, 'X', 'Y', '']] });
  mockGraph(t, ({ method, path, body }) => {
    if (method === 'GET' && path.startsWith('/users?')) return { body: { value: [] } };
    if (method === 'POST') return { status: 201, body: { id: 'id-nolic', ...body } };
    if (path.includes('/licenseDetails')) return { body: { value: [] } };
    return undefined;
  });

  const { output } = await runCommand(() => importUsers(file));
  t.true(output.includes(`Warning: User nolic@${DOMAIN} has no active licenses`));
});
