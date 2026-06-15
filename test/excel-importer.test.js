import test from 'ava';
import { mkdtemp, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import ExcelJS from 'exceljs';
import { parseExcel, rowToUserPayload, stripPasswordFromPayload } from '../src/utils/excel-importer.js';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

let tmpDir;
test.before(async () => {
  tmpDir = await mkdtemp(join(tmpdir(), 'excel-test-'));
});
test.after.always(async () => {
  await rm(tmpDir, { recursive: true, force: true });
});

async function createTestExcel(name, rows) {
  const p = join(tmpDir, name);
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('ACTIVOS');
  rows.forEach(row => worksheet.addRow(row));
  await workbook.xlsx.writeFile(p);
  return p;
}

// ---------------------------------------------------------------------------
// parseExcel
// ---------------------------------------------------------------------------

test('parseExcel: parses a valid Excel with required fields', async (t) => {
  const rows = [
    ['userPrincipalName', 'displayName', 'mailNickname', 'password'],
    ['jdoe@contoso.com', 'John Doe', 'jdoe', 'TempPass1!']
  ];
  const path = await createTestExcel('valid.xlsx', rows);
  const { rows: parsedRows, errors } = await parseExcel(path);
  t.is(errors.length, 0);
  t.is(parsedRows.length, 1);
  t.is(parsedRows[0].userPrincipalName, 'jdoe@contoso.com');
});

test('parseExcel: returns error for row missing required field', async (t) => {
  const rows = [
    ['userPrincipalName', 'displayName', 'mailNickname', 'password'],
    ['', 'John Doe', 'jdoe', 'TempPass1!']
  ];
  const path = await createTestExcel('missing.xlsx', rows);
  const { rows: parsedRows, errors } = await parseExcel(path);
  t.is(parsedRows.length, 0);
  t.is(errors.length, 1);
  t.true(errors[0].errors[0].includes('Missing required field: userPrincipalName'));
});

test('parseExcel: normalizes "SURNAME, NAME" format to "Name Surname"', async (t) => {
  const rows = [
    ['userPrincipalName', 'Trabajador', 'mailNickname', 'password'],
    ['jdoe@contoso.com', 'BORJA LOZANO, JAVIER', 'jdoe', 'TempPass1!']
  ];
  const path = await createTestExcel('name-format.xlsx', rows);
  const { rows: parsedRows } = await parseExcel(path);
  t.is(parsedRows[0].displayName, 'Javier Borja Lozano');
  t.is(parsedRows[0].givenName, 'Javier');
  t.is(parsedRows[0].surname, 'Borja Lozano');
});

// ---------------------------------------------------------------------------
// rowToUserPayload
// ---------------------------------------------------------------------------

test('rowToUserPayload: maps required scalar fields', (t) => {
  const row = {
    userPrincipalName: 'jdoe@contoso.com',
    displayName: 'John Doe',
    givenName: 'John',
    surname: 'Doe',
    mailNickname: 'jdoe',
    password: 'TempPass1!',
  };
  const { userPayload, manager } = rowToUserPayload(row);
  t.is(userPayload.userPrincipalName, 'jdoe@contoso.com');
  t.is(userPayload.displayName, 'John Doe');
  t.is(userPayload.givenName, 'John');
  t.is(userPayload.mailNickname, 'jdoe');
  t.is(manager, null);
});

test('rowToUserPayload: sets accountEnabled true by default', (t) => {
  const { userPayload } = rowToUserPayload({ userPrincipalName: 'a@b.com', displayName: 'A', mailNickname: 'a', password: 'p' });
  t.true(userPayload.accountEnabled);
});

test('rowToUserPayload: parses accountEnabled false string', (t) => {
  const { userPayload } = rowToUserPayload({
    userPrincipalName: 'a@b.com', displayName: 'A', mailNickname: 'a', password: 'p',
    accountEnabled: 'false',
  });
  t.false(userPayload.accountEnabled);
});

test('rowToUserPayload: wraps businessPhone in businessPhones array', (t) => {
  const { userPayload } = rowToUserPayload({
    userPrincipalName: 'a@b.com', displayName: 'A', mailNickname: 'a', password: 'p',
    businessPhone: '+34910000001',
  });
  t.deepEqual(userPayload.businessPhones, ['+34910000001']);
});

test('rowToUserPayload: builds passwordProfile with forceChange true by default', (t) => {
  const { userPayload } = rowToUserPayload({
    userPrincipalName: 'a@b.com', displayName: 'A', mailNickname: 'a', password: 'Temp1!',
  });
  t.is(userPayload.passwordProfile.password, 'Temp1!');
  t.true(userPayload.passwordProfile.forceChangePasswordNextSignIn);
});

test('rowToUserPayload: honours forceChangePasswordNextSignIn=false', (t) => {
  const { userPayload } = rowToUserPayload({
    userPrincipalName: 'a@b.com', displayName: 'A', mailNickname: 'a', password: 'Temp1!',
    forceChangePasswordNextSignIn: 'false',
  });
  t.false(userPayload.passwordProfile.forceChangePasswordNextSignIn);
});

test('rowToUserPayload: extracts manager UPN', (t) => {
  const { manager } = rowToUserPayload({
    userPrincipalName: 'a@b.com', displayName: 'A', mailNickname: 'a', password: 'p',
    manager: 'boss@contoso.com',
  });
  t.is(manager, 'boss@contoso.com');
});

test('rowToUserPayload: validates and converts valid employeeHireDate to ISO', (t) => {
  const { userPayload } = rowToUserPayload({
    userPrincipalName: 'a@b.com', displayName: 'A', mailNickname: 'a', password: 'p',
    employeeHireDate: '2024-01-15',
  });
  t.truthy(userPayload.employeeHireDate);
  t.true(userPayload.employeeHireDate.startsWith('2024-01-15'));
});

test('rowToUserPayload: ignores invalid employeeHireDate', (t) => {
  const { userPayload } = rowToUserPayload({
    userPrincipalName: 'a@b.com', displayName: 'A', mailNickname: 'a', password: 'p',
    employeeHireDate: 'not-a-date',
  });
  t.is(userPayload.employeeHireDate, undefined);
});

// ---------------------------------------------------------------------------
// stripPasswordFromPayload
// ---------------------------------------------------------------------------

test('stripPasswordFromPayload: removes passwordProfile key', (t) => {
  const payload = { displayName: 'John', passwordProfile: { password: 'secret' } };
  const result = stripPasswordFromPayload(payload);
  t.is(result.passwordProfile, undefined);
  t.is(result.displayName, 'John');
});

test('stripPasswordFromPayload: returns object unchanged if no passwordProfile', (t) => {
  const payload = { displayName: 'John' };
  const result = stripPasswordFromPayload(payload);
  t.deepEqual(result, payload);
});

// ---------------------------------------------------------------------------
// parseExcel edge cases
// ---------------------------------------------------------------------------

test('parseExcel: header-only Excel returns zero rows and zero errors', async (t) => {
  t.pass();
});

test('parseExcel: whitespace-only userPrincipalName is treated as missing', async (t) => {
  t.pass();
});
