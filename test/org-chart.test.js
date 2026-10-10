import test from 'ava';
import { cleanupHome } from './helpers/home.js';

const { analyseOrgTree } = await import('../src/commands/fix.js');

test.after.always(() => cleanupHome());

/** Build the inputs analyseOrgTree expects from a { userId: managerId|null } spec. */
function tree(spec) {
  const users = Object.keys(spec).map((id) => ({ id, displayName: id }));
  return {
    users,
    managerMap: new Map(Object.entries(spec)),
    byId: new Map(users.map((u) => [u.id, u])),
  };
}

const ids = (list) => list.map((u) => u.id).sort();

function analyse(spec, rootId = 'root') {
  const { users, managerMap, byId } = tree(spec);
  return analyseOrgTree(users, managerMap, rootId, byId);
}

test('analyseOrgTree: a healthy tree has no issues', (t) => {
  const issues = analyse({ root: null, a: 'root', b: 'a', c: 'b' });
  t.deepEqual(issues, { noManager: [], wrongRoot: [], cycles: [] });
});

test('analyseOrgTree: the root is never reported for lacking a manager', (t) => {
  t.deepEqual(analyse({ root: null }).noManager, []);
});

test('analyseOrgTree: users without a manager are reported as noManager', (t) => {
  const issues = analyse({ root: null, a: 'root', orphan: null });
  t.deepEqual(ids(issues.noManager), ['orphan']);
  t.deepEqual(issues.wrongRoot, []);
});

test('analyseOrgTree: a chain ending at someone other than the root is wrongRoot', (t) => {
  // "other" is a second top-level person; b hangs off it
  const issues = analyse({ root: null, other: null, b: 'other', c: 'b' });
  t.deepEqual(ids(issues.noManager), ['other']);
  t.deepEqual(ids(issues.wrongRoot), ['b', 'c']);
});

test('analyseOrgTree: a manager outside the user list breaks the chain', (t) => {
  const issues = analyse({ root: null, a: 'disabled-boss' });
  t.deepEqual(ids(issues.wrongRoot), ['a']);
});

test('analyseOrgTree: a two-person cycle is reported once, with both members', (t) => {
  const issues = analyse({ root: null, a: 'b', b: 'a' });
  t.is(issues.cycles.length, 1);
  t.deepEqual(ids(issues.cycles[0]), ['a', 'b']);
  t.deepEqual(issues.wrongRoot, []);
});

test('analyseOrgTree: a user who is their own manager is a cycle of one', (t) => {
  const issues = analyse({ root: null, a: 'a' });
  t.is(issues.cycles.length, 1);
  t.deepEqual(ids(issues.cycles[0]), ['a']);
});

test('analyseOrgTree: the root being managed by someone does not create a cycle', (t) => {
  // Chains stop as soon as they reach the root, whatever the root reports to
  const issues = analyse({ root: 'a', a: 'root' });
  t.deepEqual(issues, { noManager: [], wrongRoot: [], cycles: [] });
});

test('analyseOrgTree: chains deeper than 50 hops are flagged rather than looping forever', (t) => {
  const spec = { root: null };
  let parent = 'root';
  for (let i = 0; i < 60; i++) {
    spec[`n${i}`] = parent;
    parent = `n${i}`;
  }
  const issues = analyse(spec);
  t.is(issues.cycles.length, 0);
  // n0..n49 reach the root within the hop limit; the 10 deepest do not
  t.is(issues.wrongRoot.length, 10);
});
