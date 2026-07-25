/**
 * Dist surface smoke test — verifies published CJS entry exports.
 */
import * as path from 'path';

const pkg = require(path.join(__dirname, '..', 'dist', 'cjs', 'index.js'));

const required = [
    'XlsbWriter',
    'XlsxWriter',
    'XlsbReader',
    'XlsxReader',
    'ReaderFactory',
    'BigBuffer',
    'F',
    'sanitizeSheetName',
    'StreamingSheetState',
];

let failed = 0;
for (const name of required) {
    if (typeof pkg[name] === 'undefined') {
        console.error(`❌ missing export: ${name}`);
        failed++;
    } else {
        console.log(`✅ export: ${name}`);
    }
}

if (failed > 0) {
    process.exit(1);
}

console.log('Dist surface OK');
