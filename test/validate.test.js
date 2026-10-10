import test from 'ava';
import { runCommand } from './helpers/graph-mock.js';
import { writeWorkbook } from './helpers/excel.js';
import { validateCommand } from '../src/commands/validate.js';
import { DOMAIN } from '../src/constants.js';

const HEADERS = ['id_empleado', 'Trabajador', 'e_mail', 'ID responsable', 'Fecha de Baja'];

test.serial('validate: clean file → exit 0', async (t) => {
  const file = await writeWorkbook({
    ACTIVOS: [
      HEADERS,
      [1, 'JEFA', `jefa@${DOMAIN}`, '', ''],
      [2, 'ANA', `ana@${DOMAIN}`, 1, ''],
    ],
  });

  const { exitCode, output } = await runCommand(() => validateCommand(file));
  t.is(exitCode, 0);
  t.true(output.includes('No issues found'));
});

test.serial('validate: issues are grouped by rule, name the worker, and exit 1', async (t) => {
  const file = await writeWorkbook({
    ACTIVOS: [
      HEADERS,
      [1, 'JEFA', `jefa@${DOMAIN}`, '', ''],
      [1, 'DUPLICADA', 'dup@gmail.com', 99, '2024-13-45'],
    ],
  });

  const { exitCode, output } = await runCommand(() => validateCommand(file));

  t.is(exitCode, 1);
  for (const rule of ['id_empleado-required-unique', 'email-required-domain', 'id_responsable-exists', 'fecha-de-baja-format']) {
    t.true(output.includes(`● Rule: ${rule}`), rule);
  }
  t.true(output.includes('DUPLICADA'));
  t.true(output.includes('4 issue(s) found across 1 row(s)'));
});

test.serial('validate: header-only sheet → exit 0 with a notice', async (t) => {
  const file = await writeWorkbook({ ACTIVOS: [HEADERS] });
  const { exitCode, output } = await runCommand(() => validateCommand(file));
  t.is(exitCode, 0);
  t.true(output.includes('no data rows'));
});

test.serial('validate: unreadable file → exit 1', async (t) => {
  const { exitCode, output } = await runCommand(() => validateCommand('/nonexistent/file.xlsx'));
  t.is(exitCode, 1);
  t.true(output.includes('Failed to read Excel file'));
});
