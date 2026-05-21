import test from 'ava';
import { normalizeUpn, parseDomain } from '../src/utils/upn.js';

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

test('normalizeUpn: returns empty string for blank-only input', (t) => {
  t.is(normalizeUpn('   ', 'contoso.com'), '');
});

// ---------------------------------------------------------------------------
// parseDomain
// ---------------------------------------------------------------------------

test('parseDomain: extracts domain from a full UPN', (t) => {
  t.is(parseDomain('jdoe@contoso.com'), 'contoso.com');
});

test('parseDomain: returns bare domain unchanged', (t) => {
  t.is(parseDomain('contoso.com'), 'contoso.com');
});

test('parseDomain: returns empty string for falsy input', (t) => {
  t.is(parseDomain(null), '');
  t.is(parseDomain(''), '');
  t.is(parseDomain(undefined), '');
});

test('parseDomain: handles subdomain UPN correctly', (t) => {
  t.is(parseDomain('user@mail.contoso.com'), 'mail.contoso.com');
});
