import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const currentFile = fileURLToPath(import.meta.url);
const currentDir = dirname(currentFile);
console.log('Current dir:', currentDir);
const root = resolve(currentDir, '..'); // Since test-path is at root, '..' is workspace
console.log('Resolved root:', root);
