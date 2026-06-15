import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const currentFile = fileURLToPath(import.meta.url);
console.log('Current file:', currentFile);
const root = resolve(currentFile, '../../');
console.log('Resolved root:', root);
