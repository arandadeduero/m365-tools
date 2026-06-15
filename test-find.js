import { findExcelFile } from './src/utils/excel.js';
try {
    const file = await findExcelFile();
    console.log('Found:', file);
} catch (e) {
    console.error(e);
}
