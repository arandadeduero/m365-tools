import test from 'ava';
import { cellToString } from '../src/utils/excel.js';

// ---------------------------------------------------------------------------
// cellToString
// ---------------------------------------------------------------------------

test('cellToString: null returns empty string', (t) => {
  t.is(cellToString(null), '');
});

test('cellToString: undefined returns empty string', (t) => {
  t.is(cellToString(undefined), '');
});

test('cellToString: plain string is trimmed and returned', (t) => {
  t.is(cellToString('  hello  '), 'hello');
});

test('cellToString: number is converted to string', (t) => {
  t.is(cellToString(42), '42');
});

test('cellToString: boolean true is converted to string', (t) => {
  t.is(cellToString(true), 'true');
});

test('cellToString: Date is formatted as dd/mm/yyyy', (t) => {
  // Use a fixed date: 5 March 2024
  const d = new Date(2024, 2, 5); // month is 0-indexed
  t.is(cellToString(d), '05/03/2024');
});

test('cellToString: Date pads day and month with leading zeros', (t) => {
  const d = new Date(2000, 0, 1); // 1 Jan 2000
  t.is(cellToString(d), '01/01/2000');
});

test('cellToString: hyperlink object uses text property', (t) => {
  t.is(cellToString({ text: 'example', hyperlink: 'https://example.com' }), 'example');
});

test('cellToString: hyperlink object with nested rich-text in text', (t) => {
  const val = { text: { richText: [{ text: 'foo' }, { text: ' bar' }] }, hyperlink: 'x' };
  t.is(cellToString(val), 'foo bar');
});

test('cellToString: rich text object concatenates all parts', (t) => {
  const val = { richText: [{ text: 'Hello' }, { text: ' World' }] };
  t.is(cellToString(val), 'Hello World');
});

test('cellToString: rich text object trims result', (t) => {
  const val = { richText: [{ text: '  trim me  ' }] };
  t.is(cellToString(val), 'trim me');
});

test('cellToString: rich text with missing text fields uses empty string', (t) => {
  const val = { richText: [{ text: 'A' }, {}] };
  t.is(cellToString(val), 'A');
});

test('cellToString: formula result object recurses into result', (t) => {
  t.is(cellToString({ formula: '=A1', result: 'calculated' }), 'calculated');
});

test('cellToString: formula result with numeric result', (t) => {
  t.is(cellToString({ formula: '=1+1', result: 2 }), '2');
});

test('cellToString: formula result with null result returns empty string', (t) => {
  t.is(cellToString({ formula: '=A1', result: null }), '');
});

test('cellToString: empty string returns empty string', (t) => {
  t.is(cellToString(''), '');
});
