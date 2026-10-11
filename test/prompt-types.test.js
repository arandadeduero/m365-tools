import test from 'ava';
import { readdir, readFile } from 'node:fs/promises';
import { join } from 'node:path';
import inquirer from 'inquirer';

const COMMANDS_DIR = join(import.meta.dirname, '..', 'src', 'commands');

// Command tests replace inquirer.prompt, so a prompt type that the installed
// inquirer no longer registers (e.g. the old 'list') would only fail in real use.
test('every prompt type used by the commands is registered in inquirer', async (t) => {
  const registered = new Set(Object.keys(inquirer.prompt.prompts));
  const used = [];

  for (const file of await readdir(COMMANDS_DIR)) {
    const source = await readFile(join(COMMANDS_DIR, file), 'utf8');
    for (const [, type] of source.matchAll(/\btype:\s*'([a-z]+)'/g)) {
      used.push({ file, type });
    }
  }

  t.true(used.length > 0);
  t.deepEqual(used.filter(({ type }) => !registered.has(type)), []);
});
