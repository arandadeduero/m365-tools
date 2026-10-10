import test from 'ava';
import {
  filterAutoGroups,
  getCachedUser, setCachedUser, invalidateUser, seedUsers,
  getCachedManager, setCachedManager, invalidateManager,
  getCachedUserGroups, setCachedUserGroups, invalidateUserGroups, seedGroupsMap,
  getCachedAllGroups, setCachedAllGroups,
  getGroupName, setGroupName,
  getCachedDepartments, setCachedDepartments,
  getCachedJobTitles, setCachedJobTitles,
} from '../src/utils/cache.js';

// The cache is module-level state with no reset, so every test uses its own keys.

// ---------------------------------------------------------------------------
// filterAutoGroups
// ---------------------------------------------------------------------------

test('filterAutoGroups: drops the tenant-wide auto-assigned groups', (t) => {
  const groups = [
    { id: '1', displayName: 'Todo Ayuntamiento' },
    { id: '2', displayName: 'Urbanismo' },
    { id: '3', displayName: 'Todos los usuarios' },
    { id: '4', displayName: 'Expertos 365' },
  ];
  t.deepEqual(filterAutoGroups(groups), [{ id: '2', displayName: 'Urbanismo' }]);
});

test('filterAutoGroups: also accepts plain group-name strings', (t) => {
  t.deepEqual(filterAutoGroups(['Todos los usuarios', 'Policía']), ['Policía']);
});

test('filterAutoGroups: match is exact and case-sensitive', (t) => {
  t.deepEqual(filterAutoGroups(['todos los usuarios']), ['todos los usuarios']);
});

// ---------------------------------------------------------------------------
// User cache
// ---------------------------------------------------------------------------

test('user cache: a user is reachable by both object ID and UPN', (t) => {
  const user = { id: 'u-both', userPrincipalName: 'both@contoso.com' };
  setCachedUser(user);
  t.is(getCachedUser('u-both'), user);
  t.is(getCachedUser('both@contoso.com'), user);
});

test('user cache: miss returns null', (t) => {
  t.is(getCachedUser('u-never-set'), null);
});

test('user cache: setCachedUser ignores null', (t) => {
  t.notThrows(() => setCachedUser(null));
});

test('user cache: invalidating by UPN also drops the object-ID entry', (t) => {
  setCachedUser({ id: 'u-inv-1', userPrincipalName: 'inv1@contoso.com' });
  invalidateUser('inv1@contoso.com');
  t.is(getCachedUser('u-inv-1'), null);
  t.is(getCachedUser('inv1@contoso.com'), null);
});

test('user cache: invalidating by object ID also drops the UPN entry', (t) => {
  setCachedUser({ id: 'u-inv-2', userPrincipalName: 'inv2@contoso.com' });
  invalidateUser('u-inv-2');
  t.is(getCachedUser('u-inv-2'), null);
  t.is(getCachedUser('inv2@contoso.com'), null);
});

test('user cache: invalidating an unknown key is a no-op', (t) => {
  t.notThrows(() => invalidateUser('u-unknown'));
});

test('seedUsers: caches every user in the list', (t) => {
  seedUsers([
    { id: 'u-seed-1', userPrincipalName: 'seed1@contoso.com' },
    { id: 'u-seed-2', userPrincipalName: 'seed2@contoso.com' },
  ]);
  t.is(getCachedUser('seed1@contoso.com').id, 'u-seed-1');
  t.is(getCachedUser('u-seed-2').userPrincipalName, 'seed2@contoso.com');
});

// ---------------------------------------------------------------------------
// Manager cache — "no manager" must be distinguishable from "not cached"
// ---------------------------------------------------------------------------

test('manager cache: miss reports hit=false', (t) => {
  t.deepEqual(getCachedManager('m-miss'), { hit: false, value: null });
});

test('manager cache: a cached null manager is a hit', (t) => {
  setCachedManager('m-null', null);
  t.deepEqual(getCachedManager('m-null'), { hit: true, value: null });
});

test('manager cache: stores and invalidates a manager', (t) => {
  const boss = { id: 'boss' };
  setCachedManager('m-set', boss);
  t.deepEqual(getCachedManager('m-set'), { hit: true, value: boss });
  invalidateManager('m-set');
  t.false(getCachedManager('m-set').hit);
});

// ---------------------------------------------------------------------------
// Group caches
// ---------------------------------------------------------------------------

test('user-groups cache: miss returns null, an empty membership is a hit', (t) => {
  t.is(getCachedUserGroups('g-miss'), null);
  setCachedUserGroups('g-empty', []);
  t.deepEqual(getCachedUserGroups('g-empty'), []);
});

test('user-groups cache: storing groups also fills the group-name lookup', (t) => {
  setCachedUserGroups('g-names', [{ id: 'grp-a', displayName: 'Grupo A' }]);
  t.is(getGroupName('grp-a'), 'Grupo A');
});

test('user-groups cache: invalidateUserGroups forces a re-fetch', (t) => {
  setCachedUserGroups('g-inv', [{ id: 'grp-b', displayName: 'Grupo B' }]);
  invalidateUserGroups('g-inv');
  t.is(getCachedUserGroups('g-inv'), null);
});

test('seedGroupsMap: caches memberships for every user in the map', (t) => {
  seedGroupsMap(new Map([
    ['g-seed-1', [{ id: 'grp-c', displayName: 'Grupo C' }]],
    ['g-seed-2', []],
  ]));
  t.is(getCachedUserGroups('g-seed-1').length, 1);
  t.deepEqual(getCachedUserGroups('g-seed-2'), []);
  t.is(getGroupName('grp-c'), 'Grupo C');
});

test('group names: unknown ID returns null, setGroupName overrides', (t) => {
  t.is(getGroupName('grp-unknown'), null);
  setGroupName('grp-d', 'Grupo D');
  t.is(getGroupName('grp-d'), 'Grupo D');
});

// Whole-tenant lists start out null ("not fetched yet"); kept in one serial
// test because they are single slots rather than keyed maps.
test.serial('tenant-wide lists: null until set, then returned as stored', (t) => {
  t.is(getCachedAllGroups(), null);
  t.is(getCachedDepartments(), null);
  t.is(getCachedJobTitles(), null);

  setCachedAllGroups([{ id: 'grp-e', displayName: 'Grupo E' }]);
  setCachedDepartments(['Urbanismo']);
  setCachedJobTitles(['Arquitecto']);

  t.is(getCachedAllGroups().length, 1);
  t.is(getGroupName('grp-e'), 'Grupo E');
  t.deepEqual(getCachedDepartments(), ['Urbanismo']);
  t.deepEqual(getCachedJobTitles(), ['Arquitecto']);
});
