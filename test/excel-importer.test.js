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
  const path = await createTestExcel('header-only.xlsx', [
    ['userPrincipalName', 'displayName', 'mailNickname'],
  ]);
  const { rows, errors } = await parseExcel(path, 'contoso.com');
  t.deepEqual(rows, []);
  t.deepEqual(errors, []);
});

test('parseExcel: a row without userPrincipalName is rejected with its sheet row number', async (t) => {
  const path = await createTestExcel('missing-upn.xlsx', [
    ['userPrincipalName', 'displayName', 'mailNickname'],
    ['ok@contoso.com', 'Ok User', 'ok'],
    ['', 'No Upn', 'noupn'],
  ]);
  const { rows, errors } = await parseExcel(path, 'contoso.com');
  t.is(rows.length, 1);
  t.deepEqual(errors, [
    { row: 3, upn: '(unknown)', errors: ['Missing required field: userPrincipalName'] },
  ]);
});

// ---------------------------------------------------------------------------
// employeeHireDate — the employee file uses dd/mm/yyyy
// ---------------------------------------------------------------------------

const hireDate = (value) => rowToUserPayload({
  userPrincipalName: 'jdoe@contoso.com',
  displayName: 'John Doe',
  mailNickname: 'jdoe',
  employeeHireDate: value,
}).userPayload.employeeHireDate;

test('rowToUserPayload: dd/mm/yyyy hire date keeps its day and month', (t) => {
  // new Date('05/03/2020') would read this as 3 May
  t.is(hireDate('05/03/2020'), '2020-03-05T00:00:00.000Z');
});

test('rowToUserPayload: dd/mm/yyyy hire date with day > 12 is not dropped', (t) => {
  t.is(hireDate('25/03/2020'), '2020-03-25T00:00:00.000Z');
});

test('rowToUserPayload: hire date accepts single-digit day and month', (t) => {
  t.is(hireDate('5/3/2020'), '2020-03-05T00:00:00.000Z');
});

test('rowToUserPayload: hire date tolerates surrounding whitespace', (t) => {
  t.is(hireDate('  31/12/1999 '), '1999-12-31T00:00:00.000Z');
});

test('rowToUserPayload: ISO 8601 hire dates still work', (t) => {
  t.is(hireDate('2020-03-05'), '2020-03-05T00:00:00.000Z');
  t.is(hireDate('2020-03-05T08:30:00Z'), '2020-03-05T08:30:00.000Z');
});

test('rowToUserPayload: impossible or unparseable hire dates are omitted', (t) => {
  t.is(hireDate('31/02/2020'), undefined);
  t.is(hireDate('00/01/2020'), undefined);
  t.is(hireDate('15/13/2020'), undefined);
  t.is(hireDate('pendiente'), undefined);
  t.is(hireDate('   '), undefined);
});

test('parseExcel + rowToUserPayload: a real Date cell under "Fecha de Alta" keeps its calendar day', async (t) => {
  const path = await createTestExcel('hire-date-cell.xlsx', [
    ['Trabajador', 'e_mail', 'Fecha de Alta'],
    ['DOE, JOHN', 'jdoe@contoso.com', new Date(2021, 0, 7)],
    ['ROE, JANE', 'jroe@contoso.com', '19/11/2018'],
  ]);
  const { rows, errors } = await parseExcel(path, 'contoso.com');
  t.deepEqual(errors, []);
  t.deepEqual(rows.map((r) => rowToUserPayload(r).userPayload.employeeHireDate), [
    '2021-01-07T00:00:00.000Z',
    '2018-11-19T00:00:00.000Z',
  ]);
});
