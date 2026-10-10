import test from 'ava';
import { cleanupHome } from './helpers/home.js';
import { mockGraph, batchResponse, writes } from './helpers/graph-mock.js';

const { initAuth, _setFakeToken, _clearFakeToken } = await import('../src/auth.js');
const graph = await import('../src/graph.js');

// graph.js keeps a process-lifetime cache, so every test uses its own user IDs/UPNs.

test.before(() => {
  initAuth({ tenantId: 't', clientId: 'c', domain: 'contoso.com' });
  _setFakeToken('fake-token');
});

test.after.always(() => {
  _clearFakeToken();
  cleanupHome();
});

const graphError = (status, code, message = code) => ({ status, body: { error: { code, message } } });

// ---------------------------------------------------------------------------
// getUser / updateUser
// ---------------------------------------------------------------------------

test.serial('getUser: returns the user and serves the second lookup from cache', async (t) => {
  const user = { id: 'id-get-1', userPrincipalName: 'get1@contoso.com', displayName: 'Get One' };
  const calls = mockGraph(t, () => ({ body: user }));

  t.like(await graph.getUser('get1@contoso.com'), user);
  t.like(await graph.getUser('id-get-1'), user, 'object ID hits the same cache entry');
  t.is(calls.length, 1);
});

test.serial('getUser: 404 → null', async (t) => {
  mockGraph(t, () => graphError(404, 'Request_ResourceNotFound'));
  t.is(await graph.getUser('missing@contoso.com'), null);
});

test.serial('getUser: other Graph errors propagate with their status code', async (t) => {
  mockGraph(t, () => graphError(403, 'Authorization_RequestDenied'));
  const err = await t.throwsAsync(() => graph.getUser('forbidden@contoso.com'));
  t.is(err.statusCode, 403);
});

test.serial('updateUser: PATCHes the user and invalidates the cached copy', async (t) => {
  const upn = 'upd1@contoso.com';
  let department = 'Old';
  const calls = mockGraph(t, ({ method, body }) => {
    if (method === 'PATCH') { department = body.department; return { status: 204 }; }
    return { body: { id: 'id-upd-1', userPrincipalName: upn, department } };
  });

  t.is((await graph.getUser(upn)).department, 'Old');
  await graph.updateUser(upn, { department: 'New' });
  t.is((await graph.getUser(upn)).department, 'New', 'stale cache entry was dropped');

  t.deepEqual(writes(calls).map((c) => [c.method, c.path, c.body]), [
    ['PATCH', `/users/${upn}`, { department: 'New' }],
  ]);
});

// ---------------------------------------------------------------------------
// searchUsers
// ---------------------------------------------------------------------------

test.serial('searchUsers: uses $search with the eventual-consistency header path', async (t) => {
  const calls = mockGraph(t, () => ({ body: { value: [{ id: 's1' }] } }));
  const result = await graph.searchUsers('Juan', 5);

  t.deepEqual(result, [{ id: 's1' }]);
  t.is(calls.length, 1);
  t.true(calls[0].url.includes('$search='));
  t.true(calls[0].url.includes('"displayName:Juan"'));
  t.true(calls[0].url.includes('$top=5'));
});

test.serial('searchUsers: falls back to a startsWith $filter when $search fails', async (t) => {
  const calls = mockGraph(t, ({ url }) => (
    url.includes('$search=')
      ? graphError(400, 'Request_UnsupportedQuery')
      : { body: { value: [{ id: 's2' }] } }
  ));

  t.deepEqual(await graph.searchUsers("O'Brien"), [{ id: 's2' }]);
  t.is(calls.length, 2);
  // OData string literals escape a single quote by doubling it
  t.true(calls[1].url.includes("startsWith(displayName,'O''Brien')"));
  t.true(calls[1].url.includes("startsWith(mail,'O''Brien')"));
  t.true(calls[1].url.includes("startsWith(userPrincipalName,'O''Brien')"));
});

test.serial('searchUsers: missing value array → empty list', async (t) => {
  mockGraph(t, () => ({ body: {} }));
  t.deepEqual(await graph.searchUsers('nobody'), []);
});

// ---------------------------------------------------------------------------
// listAllUsers
// ---------------------------------------------------------------------------

test.serial('listAllUsers: follows @odata.nextLink and drops #EXT# guests', async (t) => {
  const page2 = 'https://graph.microsoft.com/v1.0/users?$skiptoken=PAGE2';
  const calls = mockGraph(t, ({ url }) => {
    if (url === page2) {
      return { body: { value: [{ id: 'l3', userPrincipalName: 'l3@contoso.com' }] } };
    }
    return {
      body: {
        value: [
          { id: 'l1', userPrincipalName: 'l1@contoso.com' },
          { id: 'l2', userPrincipalName: 'guest_gmail.com#EXT#@contoso.onmicrosoft.com' },
        ],
        '@odata.nextLink': page2,
      },
    };
  });
  const progress = [];

  const users = await graph.listAllUsers({ onProgress: (n) => progress.push(n) });

  t.deepEqual(users.map((u) => u.id), ['l1', 'l3']);
  t.deepEqual(progress, [1, 2]);
  t.is(calls.length, 2);
  t.true(calls[0].url.includes('$filter=accountEnabled eq true'));
});

test.serial('listAllUsers: onlyDisabled switches the server-side filter', async (t) => {
  const calls = mockGraph(t, () => ({ body: { value: [] } }));
  t.deepEqual(await graph.listAllUsers({ onlyDisabled: true }), []);
  t.true(calls[0].url.includes('$filter=accountEnabled eq false'));
});

test.serial('listAllUsers: seeds the user cache so later getUser calls are free', async (t) => {
  const calls = mockGraph(t, () => ({
    body: { value: [{ id: 'id-seeded', userPrincipalName: 'seeded@contoso.com' }] },
  }));
  await graph.listAllUsers();
  t.is((await graph.getUser('seeded@contoso.com')).id, 'id-seeded');
  t.is(calls.length, 1);
});

test.serial('listAllUsers: checkManager attaches { upn, name } or null per user', async (t) => {
  mockGraph(t, ({ path }) => {
    if (path.startsWith('/users/cm-1/manager')) {
      return { body: { userPrincipalName: 'boss@contoso.com', displayName: 'The Boss' } };
    }
    if (path.includes('/manager')) return graphError(404, 'Request_ResourceNotFound');
    return {
      body: {
        value: [
          { id: 'cm-1', userPrincipalName: 'cm1@contoso.com' },
          { id: 'cm-2', userPrincipalName: 'cm2@contoso.com' },
        ],
      },
    };
  });

  const users = await graph.listAllUsers({ checkManager: true });
  t.deepEqual(users[0].manager, { upn: 'boss@contoso.com', name: 'The Boss' });
  t.is(users[1].manager, null);
});

// ---------------------------------------------------------------------------
// Managers
// ---------------------------------------------------------------------------

test.serial('getManager: 404 means "no manager" and is cached as such', async (t) => {
  const calls = mockGraph(t, () => graphError(404, 'Request_ResourceNotFound'));
  t.is(await graph.getManager('id-nomgr'), null);
  t.is(await graph.getManager('id-nomgr'), null);
  t.is(calls.length, 1);
});

test.serial('setManager: resolves the manager UPN to an object ID and PUTs the reference', async (t) => {
  const calls = mockGraph(t, ({ method, path }) => {
    if (method === 'PUT') return { status: 204 };
    if (path.startsWith('/users/mgr-a@contoso.com')) {
      return { body: { id: 'id-mgr-a', userPrincipalName: 'mgr-a@contoso.com' } };
    }
    return undefined;
  });

  await graph.setManager('id-report-a', 'mgr-a@contoso.com');

  t.deepEqual(writes(calls).map((c) => [c.method, c.path, c.body]), [[
    'PUT',
    '/users/id-report-a/manager/$ref',
    { '@odata.id': 'https://graph.microsoft.com/v1.0/users/id-mgr-a' },
  ]]);
});

test.serial('setManager: a pre-resolved manager ID skips the lookup', async (t) => {
  const calls = mockGraph(t, ({ method }) => (method === 'PUT' ? { status: 204 } : undefined));
  await graph.setManager('id-report-b', 'ignored@contoso.com', 'id-mgr-b');

  t.is(calls.length, 1);
  t.is(calls[0].body['@odata.id'], 'https://graph.microsoft.com/v1.0/users/id-mgr-b');
});

test.serial('setManager: unknown manager → error, nothing written', async (t) => {
  const calls = mockGraph(t, () => graphError(404, 'Request_ResourceNotFound'));
  const err = await t.throwsAsync(() => graph.setManager('id-report-c', 'ghost@contoso.com'));
  t.is(err.message, 'Manager not found: ghost@contoso.com');
  t.deepEqual(writes(calls), []);
});

test.serial('setManager: invalidates the cached manager of that user', async (t) => {
  let boss = { id: 'id-old-boss', displayName: 'Old' };
  mockGraph(t, ({ method }) => {
    if (method === 'PUT') { boss = { id: 'id-new-boss', displayName: 'New' }; return { status: 204 }; }
    return { body: boss };
  });

  t.is((await graph.getManager('id-report-d')).id, 'id-old-boss');
  await graph.setManager('id-report-d', 'new@contoso.com', 'id-new-boss');
  t.is((await graph.getManager('id-report-d')).id, 'id-new-boss');
});

// ---------------------------------------------------------------------------
// $batch helpers
// ---------------------------------------------------------------------------

test.serial('fetchManagerMap: splits into batches of 20 and maps missing managers to null', async (t) => {
  const users = Array.from({ length: 25 }, (_, i) => ({ id: `bm-${i}` }));
  const calls = mockGraph(t, (call) => batchResponse(call, (r) => (
    r.url.startsWith('/users/bm-3/')
      ? { status: 404, body: { error: { code: 'Request_ResourceNotFound' } } }
      : { body: { id: 'the-boss' } }
  )));
  const progress = [];

  const map = await graph.fetchManagerMap(users, (done, total) => progress.push([done, total]));

  t.deepEqual(calls.map((c) => c.body.requests.length), [20, 5]);
  t.is(map.size, 25);
  t.is(map.get('bm-0'), 'the-boss');
  t.is(map.get('bm-3'), null);
  t.is(map.get('bm-24'), 'the-boss');
  t.deepEqual(progress, [[20, 25], [25, 25]]);
});

test.serial('fetchManagerMap: sends the bearer token', async (t) => {
  const original = globalThis.fetch;
  let auth;
  globalThis.fetch = async (_url, init) => {
    auth = init.headers.Authorization;
    return new Response(JSON.stringify({ responses: [] }), { status: 200 });
  };
  t.teardown(() => { globalThis.fetch = original; });

  await graph.fetchManagerMap([{ id: 'bm-auth' }]);
  t.is(auth, 'Bearer fake-token');
});

test.serial('fetchManagerMap: empty user list makes no request', async (t) => {
  const calls = mockGraph(t, () => ({ body: {} }));
  t.is((await graph.fetchManagerMap([])).size, 0);
  t.is(calls.length, 0);
});

test.serial('fetchGroupsMap: failed sub-requests become empty lists; results seed getUserGroups', async (t) => {
  const calls = mockGraph(t, (call) => batchResponse(call, (r) => (
    r.url.startsWith('/users/bg-2/')
      ? { status: 404, body: {} }
      : { body: { value: [{ id: 'grp-1', displayName: 'Urbanismo' }] } }
  )));

  const map = await graph.fetchGroupsMap([{ id: 'bg-1' }, { id: 'bg-2' }]);

  t.deepEqual(map.get('bg-1'), [{ id: 'grp-1', displayName: 'Urbanismo' }]);
  t.deepEqual(map.get('bg-2'), []);

  t.deepEqual(await graph.getUserGroups('bg-1'), [{ id: 'grp-1', displayName: 'Urbanismo' }]);
  t.is(calls.length, 1, 'getUserGroups was answered from the seeded cache');
});

test.serial('fetchGroupsMap: non-OK batch response throws with status and Graph message', async (t) => {
  mockGraph(t, () => graphError(401, 'InvalidAuthenticationToken', 'Token expired'));
  const err = await t.throwsAsync(() => graph.fetchGroupsMap([{ id: 'bg-err' }]));
  t.is(err.message, 'Graph $batch failed (401): Token expired');
});

test.serial('fetchLicensesMap: returns licence details per user', async (t) => {
  mockGraph(t, (call) => batchResponse(call, (r) => (
    r.url.startsWith('/users/bl-1/')
      ? { body: { value: [{ skuId: 'sku-1', skuPartNumber: 'O365_BUSINESS_PREMIUM' }] } }
      : { body: { value: [] } }
  )));

  const map = await graph.fetchLicensesMap([{ id: 'bl-1' }, { id: 'bl-2' }]);
  t.is(map.get('bl-1')[0].skuPartNumber, 'O365_BUSINESS_PREMIUM');
  t.deepEqual(map.get('bl-2'), []);
});

// ---------------------------------------------------------------------------
// Licences
// ---------------------------------------------------------------------------

test.serial('removeAllLicenses: removes every assigned SKU in one call', async (t) => {
  const calls = mockGraph(t, ({ method }) => (
    method === 'POST'
      ? { body: {} }
      : { body: { value: [{ skuId: 'sku-a' }, { skuId: 'sku-b' }] } }
  ));

  await graph.removeAllLicenses('lic1@contoso.com');

  t.deepEqual(writes(calls).map((c) => [c.path, c.body]), [[
    '/users/lic1@contoso.com/assignLicense',
    { addLicenses: [], removeLicenses: ['sku-a', 'sku-b'] },
  ]]);
});

test.serial('removeAllLicenses: user without licences → no write', async (t) => {
  const calls = mockGraph(t, () => ({ body: { value: [] } }));
  await graph.removeAllLicenses('lic2@contoso.com');
  t.deepEqual(writes(calls), []);
});

test.serial('removeAllLicenses: a pre-fetched list skips the lookup', async (t) => {
  const calls = mockGraph(t, () => ({ body: {} }));
  await graph.removeAllLicenses('lic3@contoso.com', [{ skuId: 'sku-c' }]);
  t.is(calls.length, 1);
  t.deepEqual(calls[0].body.removeLicenses, ['sku-c']);
});

// ---------------------------------------------------------------------------
// Groups
// ---------------------------------------------------------------------------

test.serial('addMemberToGroup: POSTs a directoryObject reference', async (t) => {
  const calls = mockGraph(t, () => ({ status: 204 }));
  await graph.addMemberToGroup('grp-x', 'id-member-1');
  t.deepEqual([calls[0].method, calls[0].path, calls[0].body], [
    'POST',
    '/groups/grp-x/members/$ref',
    { '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/id-member-1' },
  ]);
});

test.serial('addMemberToGroup: 409 "already a member" is treated as success', async (t) => {
  mockGraph(t, () => graphError(409, 'Request_BadRequest', 'already exist'));
  await t.notThrowsAsync(() => graph.addMemberToGroup('grp-x', 'id-member-2'));
});

test.serial('addMemberToGroup: other errors propagate', async (t) => {
  mockGraph(t, () => graphError(403, 'Authorization_RequestDenied'));
  const err = await t.throwsAsync(() => graph.addMemberToGroup('grp-x', 'id-member-3'));
  t.is(err.statusCode, 403);
});

test.serial('removeMemberFromGroup: 404 "not a member" is treated as success', async (t) => {
  const calls = mockGraph(t, () => graphError(404, 'Request_ResourceNotFound'));
  await t.notThrowsAsync(() => graph.removeMemberFromGroup('grp-x', 'id-member-4'));
  t.deepEqual([calls[0].method, calls[0].path], ['DELETE', '/groups/grp-x/members/id-member-4/$ref']);
});

test.serial('searchGroups: escapes single quotes in the name prefix', async (t) => {
  const calls = mockGraph(t, () => ({ body: { value: [] } }));
  await graph.searchGroups("L'Hospitalet");
  t.true(calls[0].url.includes("startsWith(displayName,'L''Hospitalet')"));
});

test.serial('listAllGroups: paginates, sorts by name (es) and caches the result', async (t) => {
  const page2 = 'https://graph.microsoft.com/v1.0/groups?$skiptoken=G2';
  const calls = mockGraph(t, ({ url }) => (
    url === page2
      ? { body: { value: [{ id: 'g-a', displayName: 'Alcaldía' }] } }
      : { body: { value: [{ id: 'g-z', displayName: 'Zona azul' }], '@odata.nextLink': page2 } }
  ));

  const groups = await graph.listAllGroups();
  t.deepEqual(groups.map((g) => g.displayName), ['Alcaldía', 'Zona azul']);

  await graph.listAllGroups();
  t.is(calls.length, 2, 'second call served from cache');
});

// ---------------------------------------------------------------------------
// Department / job-title queries
// ---------------------------------------------------------------------------

test.serial('getUsersByDepartment: escapes quotes and excludes guests', async (t) => {
  const calls = mockGraph(t, () => ({
    body: {
      value: [
        { id: 'd1', userPrincipalName: 'd1@contoso.com' },
        { id: 'd2', userPrincipalName: 'x#EXT#@contoso.onmicrosoft.com' },
      ],
    },
  }));

  const users = await graph.getUsersByDepartment("Cap d'Estudis");
  t.deepEqual(users.map((u) => u.id), ['d1']);
  t.true(calls[0].url.includes("department eq 'Cap d''Estudis'"));
  t.true(calls[0].url.includes('accountEnabled eq true'));
});

test.serial('getUsersByJobTitle: escapes quotes in the title', async (t) => {
  const calls = mockGraph(t, () => ({ body: { value: [] } }));
  await graph.getUsersByJobTitle("Cap d'Unitat");
  t.true(calls[0].url.includes("jobTitle eq 'Cap d''Unitat'"));
});

test.serial('listAllDepartments: trims, de-duplicates, sorts and caches', async (t) => {
  const calls = mockGraph(t, () => ({
    body: {
      value: [
        { userPrincipalName: 'a@contoso.com', department: ' Urbanismo ' },
        { userPrincipalName: 'b@contoso.com', department: 'Alcaldía' },
        { userPrincipalName: 'c@contoso.com', department: 'Urbanismo' },
        { userPrincipalName: 'd@contoso.com', department: '   ' },
        { userPrincipalName: 'e#EXT#@contoso.onmicrosoft.com', department: 'Externos' },
      ],
    },
  }));

  t.deepEqual(await graph.listAllDepartments(), ['Alcaldía', 'Urbanismo']);
  await graph.listAllDepartments();
  t.is(calls.length, 1);
});

// ---------------------------------------------------------------------------
// Mailbox timezone
// ---------------------------------------------------------------------------

const WORKING_HOURS = {
  daysOfWeek: ['monday', 'tuesday'],
  startTime: '08:00:00.0000000',
  endTime: '15:00:00.0000000',
  timeZone: { name: 'Mountain Standard Time' },
};

test.serial('getMailboxTzSettings: exposes both timezones and never sends $select', async (t) => {
  const calls = mockGraph(t, () => ({ body: { timeZone: 'UTC', workingHours: WORKING_HOURS } }));

  const s = await graph.getMailboxTzSettings('tz-1');

  t.is(s.timeZone, 'UTC');
  t.is(s.workingHoursTimeZone, 'Mountain Standard Time');
  t.deepEqual(s.workingHours, WORKING_HOURS);
  t.is(s._error, null);
  // $select on the mailboxSettings singleton makes Exchange answer 400
  t.false(calls[0].url.includes('$select'));
});

test.serial('getMailboxTzSettings: mailbox without workingHours → null fields', async (t) => {
  mockGraph(t, () => ({ body: { timeZone: 'Romance Standard Time' } }));
  const s = await graph.getMailboxTzSettings('tz-2');
  t.is(s.workingHoursTimeZone, null);
  t.is(s.workingHours, null);
});

for (const [label, response, reason] of [
  ['404', graphError(404, 'ResourceNotFound'), 'no-mailbox'],
  ['MailboxNotEnabledForRESTAPI', graphError(401, 'MailboxNotEnabledForRESTAPI'), 'no-mailbox'],
  ['403', graphError(403, 'ErrorAccessDenied'), 'access-denied'],
  ['400', graphError(400, 'BadRequest'), 'bad-request'],
]) {
  test.serial(`getMailboxTzSettings: ${label} is reported as "${reason}" instead of throwing`, async (t) => {
    mockGraph(t, () => response);
    const s = await graph.getMailboxTzSettings(`tz-err-${label}`);
    t.is(s._reason, reason);
    t.truthy(s._error);
    t.is(s._raw, null);
  });
}

test.serial('getMailboxTzSettings: unexpected errors propagate', async (t) => {
  mockGraph(t, () => graphError(500, 'InternalServerError'));
  const err = await t.throwsAsync(() => graph.getMailboxTzSettings('tz-500'));
  t.is(err.statusCode, 500);
});

test.serial('setMailboxTzSettings: changes only the timezone inside workingHours', async (t) => {
  const calls = mockGraph(t, () => ({ body: {} }));

  await graph.setMailboxTzSettings('tz-set-1', 'Romance Standard Time', WORKING_HOURS);

  t.deepEqual([calls[0].method, calls[0].path], ['PATCH', '/users/tz-set-1/mailboxSettings']);
  t.deepEqual(calls[0].body, {
    timeZone: 'Romance Standard Time',
    workingHours: {
      daysOfWeek: ['monday', 'tuesday'],
      startTime: '08:00:00.0000000',
      endTime: '15:00:00.0000000',
      timeZone: { name: 'Romance Standard Time' },
    },
  });
});

test.serial('setMailboxTzSettings: without workingHours only the top-level timezone is sent', async (t) => {
  const calls = mockGraph(t, () => ({ body: {} }));
  await graph.setMailboxTzSettings('tz-set-2', 'Romance Standard Time', null);
  t.deepEqual(calls[0].body, { timeZone: 'Romance Standard Time' });
});
