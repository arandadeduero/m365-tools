import test from 'ava';
import { normalizeUpn } from '../src/utils/upn.js';

// ---------------------------------------------------------------------------
// normalizeUpn
// ---------------------------------------------------------------------------

test('normalizeUpn: appends domain when no @ present', (t) => {
  t.is(normalizeUpn('jdoe', 'contoso.com'), 'jdoe@contoso.com');
});

test('normalizeUpn: leaves value untouched when @ already present', (t) => {
  t.is(normalizeUpn('jdoe@other.com', 'contoso.com'), 'jdoe@other.com');
});

test('normalizeUpn: trims surrounding whitespace before processing', (t) => {
  t.is(normalizeUpn('  jdoe  ', 'contoso.com'), 'jdoe@contoso.com');
});

test('normalizeUpn: returns value as-is when domain is empty string', (t) => {
  t.is(normalizeUpn('jdoe', ''), 'jdoe');
});

test('normalizeUpn: returns value as-is when domain is null/undefined', (t) => {
  t.is(normalizeUpn('jdoe', null), 'jdoe');
  t.is(normalizeUpn('jdoe', undefined), 'jdoe');
});

test('normalizeUpn: returns value when input is null', (t) => {
  t.is(normalizeUpn(null, 'contoso.com'), null);
});

test('normalizeUpn: handles already-uppercased UPN', (t) => {
  t.is(normalizeUpn('JDOE', 'contoso.com'), 'JDOE@contoso.com');
});

test('normalizeUpn: handles UPN with multiple @ signs', (t) => {
  // Edge case: if already has @, returns as is (even if malformed)
  t.is(normalizeUpn('jdoe@@contoso.com', 'contoso.com'), 'jdoe@@contoso.com');
});
