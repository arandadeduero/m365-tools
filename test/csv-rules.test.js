import test from 'ava';
import { RULES } from '../src/validators/csv-rules.js';

// Helper to find a rule by id
const rule = (id) => RULES.find((r) => r.id === id);

// ---------------------------------------------------------------------------
// Rule: id_empleado-required-unique
// ---------------------------------------------------------------------------

const ruleIdEmpleado = rule('id_empleado-required-unique');

test('id_empleado: no issues for valid unique ids', (t) => {
  const rows = [
    { id_empleado: 'E001' },
    { id_empleado: 'E002' },
    { id_empleado: 'E003' },
  ];
  t.deepEqual(ruleIdEmpleado.validate(rows), []);
});

test('id_empleado: reports empty id_empleado as issue', (t) => {
  const rows = [{ id_empleado: '' }];
  const issues = ruleIdEmpleado.validate(rows);
  t.is(issues.length, 1);
  t.is(issues[0].row, 1);
  t.is(issues[0].column, 'id_empleado');
  t.regex(issues[0].message, /required/i);
});

test('id_empleado: reports missing id_empleado field (undefined) as issue', (t) => {
  const rows = [{}];
  const issues = ruleIdEmpleado.validate(rows);
  t.is(issues.length, 1);
  t.is(issues[0].column, 'id_empleado');
});

test('id_empleado: reports duplicate id_empleado', (t) => {
  const rows = [
    { id_empleado: 'E001' },
    { id_empleado: 'E001' },
  ];
  const issues = ruleIdEmpleado.validate(rows);
  t.is(issues.length, 1);
  t.is(issues[0].row, 2);
  t.regex(issues[0].message, /duplicate/i);
  t.regex(issues[0].message, /row 1/i);
});

test('id_empleado: multiple empty rows produce multiple issues', (t) => {
  const rows = [{ id_empleado: '' }, { id_empleado: '' }];
  const issues = ruleIdEmpleado.validate(rows);
  t.is(issues.length, 2);
});

test('id_empleado: empty rows are not counted as duplicates of each other', (t) => {
  const rows = [{ id_empleado: '' }, { id_empleado: '' }];
  const issues = ruleIdEmpleado.validate(rows);
  // Each empty row should produce a "required" issue, not a "duplicate" issue
  t.true(issues.every((i) => /required/i.test(i.message)));
});

test('id_empleado: no issues for empty rows array', (t) => {
  t.deepEqual(ruleIdEmpleado.validate([]), []);
});

// ---------------------------------------------------------------------------
// Rule: email-required-domain
// ---------------------------------------------------------------------------

const ruleEmail = rule('email-required-domain');

test('email-required-domain: valid email passes', (t) => {
  const rows = [{ e_mail: 'jdoe@arandadeduero.es' }];
  t.deepEqual(ruleEmail.validate(rows), []);
});

test('email-required-domain: empty email is an issue', (t) => {
  const rows = [{ e_mail: '' }];
  const issues = ruleEmail.validate(rows);
  t.is(issues.length, 1);
  t.is(issues[0].column, 'e_mail');
  t.regex(issues[0].message, /required/i);
});

test('email-required-domain: wrong domain is an issue', (t) => {
  const rows = [{ e_mail: 'jdoe@gmail.com' }];
  const issues = ruleEmail.validate(rows);
  t.is(issues.length, 1);
  t.regex(issues[0].message, /arandadeduero\.es/);
});

test('email-required-domain: domain check is case-insensitive', (t) => {
  const rows = [{ e_mail: 'JDOE@ARANDADEDUERO.ES' }];
  t.deepEqual(ruleEmail.validate(rows), []);
});

test('email-required-domain: undefined e_mail treated as empty', (t) => {
  const rows = [{}];
  const issues = ruleEmail.validate(rows);
  t.is(issues.length, 1);
  t.regex(issues[0].message, /required/i);
});

// ---------------------------------------------------------------------------
// Rule: email-unique
// ---------------------------------------------------------------------------

const ruleEmailUnique = rule('email-unique');

test('email-unique: unique emails pass', (t) => {
  const rows = [
    { e_mail: 'a@arandadeduero.es' },
    { e_mail: 'b@arandadeduero.es' },
  ];
  t.deepEqual(ruleEmailUnique.validate(rows), []);
});

test('email-unique: duplicate email is an issue', (t) => {
  const rows = [
    { e_mail: 'a@arandadeduero.es' },
    { e_mail: 'a@arandadeduero.es' },
  ];
  const issues = ruleEmailUnique.validate(rows);
  t.is(issues.length, 1);
  t.is(issues[0].row, 2);
  t.regex(issues[0].message, /duplicate/i);
  t.regex(issues[0].message, /row 1/i);
});

test('email-unique: comparison is case-insensitive', (t) => {
  const rows = [
    { e_mail: 'A@arandadeduero.es' },
    { e_mail: 'a@arandadeduero.es' },
  ];
  const issues = ruleEmailUnique.validate(rows);
  t.is(issues.length, 1);
});

test('email-unique: empty emails are skipped (no false positives)', (t) => {
  const rows = [{ e_mail: '' }, { e_mail: '' }];
  t.deepEqual(ruleEmailUnique.validate(rows), []);
});

// ---------------------------------------------------------------------------
// Rule: id_responsable-exists
// ---------------------------------------------------------------------------

const ruleIdResponsable = rule('id_responsable-exists');

test('id_responsable-exists: valid reference passes', (t) => {
  const rows = [
    { id_empleado: 'E001', 'ID Responsable': '' },
    { id_empleado: 'E002', 'ID Responsable': 'E001' },
  ];
  t.deepEqual(ruleIdResponsable.validate(rows), []);
});

test('id_responsable-exists: empty ID Responsable is allowed (top-level)', (t) => {
  const rows = [{ id_empleado: 'E001', 'ID Responsable': '' }];
  t.deepEqual(ruleIdResponsable.validate(rows), []);
});

test('id_responsable-exists: unknown reference is an issue', (t) => {
  const rows = [
    { id_empleado: 'E001', 'ID Responsable': 'E999' },
  ];
  const issues = ruleIdResponsable.validate(rows);
  t.is(issues.length, 1);
  t.is(issues[0].column, 'ID Responsable');
  t.is(issues[0].value, 'E999');
  t.regex(issues[0].message, /E999/);
});

test('id_responsable-exists: undefined ID Responsable treated as empty (ok)', (t) => {
  const rows = [{ id_empleado: 'E001' }];
  t.deepEqual(ruleIdResponsable.validate(rows), []);
});

// ---------------------------------------------------------------------------
// Rule: fecha-de-baja-format
// ---------------------------------------------------------------------------

const ruleFormat = rule('fecha-de-baja-format');

test('fecha-de-baja-format: valid dd/mm/yyyy passes', (t) => {
  const rows = [{ 'Fecha de Baja': '31/12/2025' }];
  t.deepEqual(ruleFormat.validate(rows), []);
});

test('fecha-de-baja-format: empty value passes (optional field)', (t) => {
  const rows = [{ 'Fecha de Baja': '' }];
  t.deepEqual(ruleFormat.validate(rows), []);
});

test('fecha-de-baja-format: wrong format (yyyy-mm-dd) is an issue', (t) => {
  const rows = [{ 'Fecha de Baja': '2025-12-31' }];
  const issues = ruleFormat.validate(rows);
  t.is(issues.length, 1);
  t.is(issues[0].column, 'Fecha de Baja');
});

test('fecha-de-baja-format: single-digit day is an issue', (t) => {
  const rows = [{ 'Fecha de Baja': '1/12/2025' }];
  const issues = ruleFormat.validate(rows);
  t.is(issues.length, 1);
});

test('fecha-de-baja-format: text value is an issue', (t) => {
  const rows = [{ 'Fecha de Baja': 'not-a-date' }];
  const issues = ruleFormat.validate(rows);
  t.is(issues.length, 1);
});

test('fecha-de-baja-format: undefined value treated as empty (ok)', (t) => {
  const rows = [{}];
  t.deepEqual(ruleFormat.validate(rows), []);
});

// ---------------------------------------------------------------------------
// Rule: fecha-de-baja-not-future
// ---------------------------------------------------------------------------

const ruleNotFuture = rule('fecha-de-baja-not-future');

test('fecha-de-baja-not-future: past date passes', (t) => {
  const rows = [{ 'Fecha de Baja': '01/01/2000' }];
  t.deepEqual(ruleNotFuture.validate(rows), []);
});

test('fecha-de-baja-not-future: future date is an issue', (t) => {
  const rows = [{ 'Fecha de Baja': '01/01/2099' }];
  const issues = ruleNotFuture.validate(rows);
  t.is(issues.length, 1);
  t.is(issues[0].column, 'Fecha de Baja');
  t.regex(issues[0].message, /future/i);
});

test('fecha-de-baja-not-future: today is allowed', (t) => {
  const now = new Date();
  const dd = String(now.getDate()).padStart(2, '0');
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const yyyy = now.getFullYear();
  const rows = [{ 'Fecha de Baja': `${dd}/${mm}/${yyyy}` }];
  t.deepEqual(ruleNotFuture.validate(rows), []);
});

test('fecha-de-baja-not-future: empty value is skipped', (t) => {
  const rows = [{ 'Fecha de Baja': '' }];
  t.deepEqual(ruleNotFuture.validate(rows), []);
});

test('fecha-de-baja-not-future: invalid format is skipped (other rule handles it)', (t) => {
  const rows = [{ 'Fecha de Baja': 'bad-format' }];
  t.deepEqual(ruleNotFuture.validate(rows), []);
});

// ---------------------------------------------------------------------------
// RULES array export
// ---------------------------------------------------------------------------

test('RULES array contains all 6 rules', (t) => {
  const expectedIds = [
    'id_empleado-required-unique',
    'email-required-domain',
    'email-unique',
    'id_responsable-exists',
    'fecha-de-baja-format',
    'fecha-de-baja-not-future',
  ];
  t.is(RULES.length, 6);
  t.deepEqual(RULES.map((r) => r.id), expectedIds);
});
