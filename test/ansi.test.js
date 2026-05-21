import test from 'ava';
import { stripAnsi } from '../src/utils/ansi.js';

test('stripAnsi: removes SGR color codes', (t) => {
  t.is(stripAnsi('\x1B[32mhello\x1B[0m'), 'hello');
});

test('stripAnsi: removes bold/reset sequences', (t) => {
  t.is(stripAnsi('\x1B[1mBold\x1B[22m'), 'Bold');
});

test('stripAnsi: removes cursor movement sequences (CSI A)', (t) => {
  t.is(stripAnsi('foo\x1B[1Abar'), 'foobar');
});

test('stripAnsi: leaves plain strings untouched', (t) => {
  t.is(stripAnsi('hello world'), 'hello world');
});

test('stripAnsi: handles empty string', (t) => {
  t.is(stripAnsi(''), '');
});

test('stripAnsi: coerces non-string to string', (t) => {
  t.is(stripAnsi(42), '42');
});

test('stripAnsi: removes multiple sequences in one string', (t) => {
  t.is(stripAnsi('\x1B[31mred\x1B[0m and \x1B[34mblue\x1B[0m'), 'red and blue');
});
