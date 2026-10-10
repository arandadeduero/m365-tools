import test from 'ava';
import { cleanupHome } from './helpers/home.js';
import { mockGraph, mockPrompts, runCommand, writes } from './helpers/graph-mock.js';

const { initAuth, _setFakeToken, _clearFakeToken } = await import('../src/auth.js');
const { fixJobTitles } = await import('../src/commands/fix.js');
const { fixTimezone } = await import('../src/commands/fix-timezone.js');

test.before(() => {
  initAuth({ tenantId: 't', clientId: 'c', domain: 'contoso.com' });
  _setFakeToken('fake-token');
});

test.after.always(() => {
  _clearFakeToken();
  cleanupHome();
});

const user = (id, extra = {}) => ({
  id,
  userPrincipalName: `${id}@contoso.com`,
  displayName: id,
  accountEnabled: true,
  ...extra,
});

// ---------------------------------------------------------------------------
// fix job-titles
// ---------------------------------------------------------------------------

const TITLED_USERS = [
  user('jt-upper', { jobTitle: 'AGENTE DE POLICÍA' }),
  user('jt-mixed', { jobTitle: 'jefe De Servicio' }),
  user('jt-ok',    { jobTitle: 'Arquitecto' }),
  user('jt-none',  { jobTitle: null }),
  user('jt-blank', { jobTitle: '   ' }),
];

function jobTitleTenant(t, failFor = null) {
  return mockGraph(t, ({ method, path }) => {
    if (method === 'GET' && path.startsWith('/users?')) return { body: { value: TITLED_USERS } };
    if (method === 'PATCH' && path === `/users/${failFor}`) {
      return { status: 403, body: { error: { code: 'Authorization_RequestDenied', message: 'Insufficient privileges' } } };
    }
    if (method === 'PATCH') return { status: 204 };
    return undefined;
  });
}

test.serial('fix job-titles: confirmed twice → only titles that change are PATCHed, in Sentence case', async (t) => {
  const calls = jobTitleTenant(t);
  mockPrompts(t, [{ selected: ['jt-upper', 'jt-mixed'] }, { confirm1: true }, { confirm2: true }]);

  const { output } = await runCommand(() => fixJobTitles());

  const patches = writes(calls).map((c) => [c.path, c.body]).sort();
  t.deepEqual(patches, [
    ['/users/jt-mixed', { jobTitle: 'Jefe de servicio' }],
    ['/users/jt-upper', { jobTitle: 'Agente de policía' }],
  ]);
  t.true(output.includes('Updated: 2'));
});

test.serial('fix job-titles: only the users ticked in the checklist are touched', async (t) => {
  const calls = jobTitleTenant(t);
  mockPrompts(t, [{ selected: ['jt-mixed'] }, { confirm1: true }, { confirm2: true }]);

  await runCommand(() => fixJobTitles());
  t.deepEqual(writes(calls).map((c) => c.path), ['/users/jt-mixed']);
});

for (const [label, answers] of [
  ['nothing selected', [{ selected: [] }]],
  ['first confirmation declined', [{ selected: ['jt-upper'] }, { confirm1: false }]],
  ['second confirmation declined', [{ selected: ['jt-upper'] }, { confirm1: true }, { confirm2: false }]],
]) {
  test.serial(`fix job-titles: ${label} → no writes`, async (t) => {
    const calls = jobTitleTenant(t);
    mockPrompts(t, answers);
    await runCommand(() => fixJobTitles());
    t.deepEqual(writes(calls), []);
  });
}

test.serial('fix job-titles: everything already in Sentence case → no prompt, no writes', async (t) => {
  const calls = mockGraph(t, () => ({ body: { value: [user('jt-fine', { jobTitle: 'Arquitecto' })] } }));
  mockPrompts(t, []);

  const { output } = await runCommand(() => fixJobTitles());

  t.deepEqual(writes(calls), []);
  t.true(output.includes('Nothing to do'));
});

test.serial('fix job-titles: a failed update is listed and does not block the others', async (t) => {
  const calls = jobTitleTenant(t, 'jt-upper');
  mockPrompts(t, [{ selected: ['jt-upper', 'jt-mixed'] }, { confirm1: true }, { confirm2: true }]);

  const { output } = await runCommand(() => fixJobTitles());

  t.is(writes(calls).length, 2);
  t.true(output.includes('Updated: 1'));
  t.true(output.includes('Failed:  1'));
  t.true(output.includes('jt-upper <jt-upper@contoso.com>: Insufficient privileges'));
});

// ---------------------------------------------------------------------------
// fix timezone
// ---------------------------------------------------------------------------

const TARGET = 'Romance Standard Time';
const hours = (tz) => ({
  daysOfWeek: ['monday', 'friday'],
  startTime: '07:30:00.0000000',
  endTime: '14:30:00.0000000',
  timeZone: { name: tz },
});

const MAILBOXES = {
  'tz-ok':        { body: { timeZone: TARGET, workingHours: hours(TARGET) } },
  'tz-ok-nohrs':  { body: { timeZone: TARGET } },
  'tz-wrong':     { body: { timeZone: 'UTC', workingHours: hours('UTC') } },
  // Top-level zone is right but Teams/free-busy still follows workingHours
  'tz-hours':     { body: { timeZone: TARGET, workingHours: hours('Mountain Standard Time') } },
  'tz-nohours':   { body: { timeZone: 'Pacific Standard Time' } },
  'tz-nomailbox': { status: 404, body: { error: { code: 'MailboxNotEnabledForRESTAPI', message: 'no mailbox' } } },
};

function timezoneTenant(t, mailboxes = MAILBOXES) {
  return mockGraph(t, ({ method, path }) => {
    if (method === 'GET' && path.startsWith('/users?')) {
      return { body: { value: Object.keys(mailboxes).map((id) => user(id)) } };
    }
    const m = path.match(/^\/users\/([^/]+)\/mailboxSettings$/);
    if (m && method === 'GET') return mailboxes[m[1]];
    if (m && method === 'PATCH') return { body: {} };
    return undefined;
  });
}

test.serial('fix timezone --dry-run: reports what would change and writes nothing', async (t) => {
  const calls = timezoneTenant(t);
  mockPrompts(t, []); // dry-run must not even ask

  const { output } = await runCommand(() => fixTimezone({ dryRun: true }));

  t.deepEqual(writes(calls), []);
  t.true(output.includes('[dry-run] 3 user(s) would be updated'));
  t.regex(output, /Already correct \(Romance Standard Time\): 2/);
  t.regex(output, /No mailbox \/ no license:\s+1/);
});

test.serial('fix timezone: confirmed twice → patches only mailboxes that differ, keeping their schedule', async (t) => {
  const calls = timezoneTenant(t);
  mockPrompts(t, [{ confirm1: true }, { confirm2: true }]);

  const { output } = await runCommand(() => fixTimezone());

  const patches = Object.fromEntries(writes(calls).map((c) => [c.path, c.body]));
  t.deepEqual(Object.keys(patches).sort(), [
    '/users/tz-hours/mailboxSettings',
    '/users/tz-nohours/mailboxSettings',
    '/users/tz-wrong/mailboxSettings',
  ]);
  t.deepEqual(patches['/users/tz-wrong/mailboxSettings'], { timeZone: TARGET, workingHours: hours(TARGET) });
  t.deepEqual(patches['/users/tz-hours/mailboxSettings'], { timeZone: TARGET, workingHours: hours(TARGET) });
  t.deepEqual(patches['/users/tz-nohours/mailboxSettings'], { timeZone: TARGET });
  t.regex(output, /Updated:\s+3/);
});

for (const [label, answers] of [
  ['first confirmation declined', [{ confirm1: false }]],
  ['second confirmation declined', [{ confirm1: true }, { confirm2: false }]],
]) {
  test.serial(`fix timezone: ${label} → no writes`, async (t) => {
    const calls = timezoneTenant(t);
    mockPrompts(t, answers);
    await runCommand(() => fixTimezone());
    t.deepEqual(writes(calls), []);
  });
}

test.serial('fix timezone: all mailboxes already correct → no prompt, no writes', async (t) => {
  const calls = timezoneTenant(t, { 'tz-fine': MAILBOXES['tz-ok'] });
  mockPrompts(t, []);

  const { output } = await runCommand(() => fixTimezone());

  t.deepEqual(writes(calls), []);
  t.true(output.includes('Nothing to do'));
});

test.serial('fix timezone: access denied is explained (Exchange Administrator role) and nothing is written', async (t) => {
  const calls = timezoneTenant(t, {
    'tz-denied': { status: 403, body: { error: { code: 'ErrorAccessDenied', message: 'Access is denied' } } },
  });
  mockPrompts(t, []);

  const { output } = await runCommand(() => fixTimezone());

  t.deepEqual(writes(calls), []);
  t.true(output.includes('ACCESS DENIED for 1 user(s)'));
  t.true(output.includes('Exchange Administrator'));
});
