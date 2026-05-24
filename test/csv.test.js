import test from 'ava';
import { writeFile, mkdtemp, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { parseCSV, rowToUserPayload, stripPasswordFromPayload } from '../src/utils/csv.js';

// ---------------------------------------------------------------------------
// parseCSV
// ---------------------------------------------------------------------------

const VALID_HEADER = 'userPrincipalName,displayName,mailNickname,password';
const VALID_ROW    = 'jdoe@contoso.com,John Doe,jdoe,TempPass1!';

let tmpDir;
test.before(async () => {
  tmpDir = await mkdtemp(join(tmpdir(), 'csv-test-'));
});
test.after.always(async () => {
  await rm(tmpDir, { recursive: true, force: true });
});

async function writeCsv(name, content) {
  const p = join(tmpDir, name);
  await writeFile(p, content, 'utf8');
  return p;
}

test('parseCSV: parses a valid CSV with required fields', async (t) => {
  const path = await writeCsv('valid.csv', `${VALID_HEADER}\n${VALID_ROW}\n`);
  const { rows, errors } = await parseCSV(path);
  t.is(errors.length, 0);
  t.is(rows.length, 1);
  t.is(rows[0].userPrincipalName, 'jdoe@contoso.com');
});

test('parseCSV: returns error for row missing required field', async (t) => {
  const path = await writeCsv('missing.csv', `${VALID_HEADER}\n,John Doe,jdoe,TempPass1!\n`);
  const { rows, errors } = await parseCSV(path);
  t.is(rows.length, 0);
  t.is(errors.length, 1);
  t.true(errors[0].errors[0].includes('userPrincipalName'));
});

test('parseCSV: handles Excel BOM prefix', async (t) => {
  // UTF-8 BOM is \uFEFF
  const bom = '\uFEFF';
  const path = await writeCsv('bom.csv', `${bom}${VALID_HEADER}\n${VALID_ROW}\n`);
  const { rows, errors } = await parseCSV(path);
  t.is(errors.length, 0);
  t.is(rows.length, 1);
});

test('parseCSV: skips empty lines', async (t) => {
  const path = await writeCsv('empty-lines.csv', `${VALID_HEADER}\n\n${VALID_ROW}\n\n`);
  const { rows } = await parseCSV(path);
  t.is(rows.length, 1);
});

test('parseCSV: reports correct row numbers (1-indexed + header)', async (t) => {
  // Row 2 in the file is header, data starts at row 2 → error on row 2
  const path = await writeCsv('rownums.csv', `${VALID_HEADER}\n,bad,missing,row\n`);
  const { errors } = await parseCSV(path);
  t.is(errors[0].row, 2);
});

test('parseCSV: parses multiple rows and accumulates errors independently', async (t) => {
  const path = await writeCsv('multi.csv',
    `${VALID_HEADER}\n${VALID_ROW}\n,No UPN,alias,pass\n`
  );
  const { rows, errors } = await parseCSV(path);
  t.is(rows.length, 1);
  t.is(errors.length, 1);
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
// parseCSV edge cases
// ---------------------------------------------------------------------------

test('parseCSV: header-only CSV returns zero rows and zero errors', async (t) => {
  const path = await writeCsv('header-only.csv', `${VALID_HEADER}\n`);
  const { rows, errors } = await parseCSV(path);
  t.is(rows.length, 0);
  t.is(errors.length, 0);
});

test('parseCSV: whitespace-only userPrincipalName is treated as missing', async (t) => {
  const path = await writeCsv('whitespace-upn.csv', `${VALID_HEADER}\n   ,John Doe,jdoe,TempPass1!\n`);
  const { rows, errors } = await parseCSV(path);
  t.is(rows.length, 0);
  t.is(errors.length, 1);
  t.true(errors[0].errors[0].includes('userPrincipalName'));
});
