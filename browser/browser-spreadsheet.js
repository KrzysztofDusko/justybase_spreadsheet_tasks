/**
 * Browser Spreadsheet - Entry point for browser-based Excel file generation.
 * Features: autofit columns, frozen+bold header row, auto-filter.
 *
 * Usage:
 *   import { downloadXlsx, downloadXlsb } from './browser-spreadsheet.js';
 *   downloadXlsx('report.xlsx', [['Alice',30],['Bob',25]], ['Name','Age']);
 *   downloadXlsb('report.xlsb', [['Alice',30],['Bob',25]], ['Name','Age']);
 */
import { BrowserXlsxWriter } from './BrowserXlsxWriter.js';
import { BrowserXlsbWriter } from './BrowserXlsbWriter.js';

export { BrowserXlsxWriter } from './BrowserXlsxWriter.js';
export { BrowserXlsbWriter } from './BrowserXlsbWriter.js';
export { BrowserBigBuffer } from './BrowserBigBuffer.js';
export { BrowserZip } from './BrowserZip.js';

/** Trigger browser download of a Blob */
function triggerDownload(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = fileName;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 5000);
}

/**
 * Download data as XLSX. Headers = bold + frozen + auto-filter. Columns = autofit.
 * @param {string} fileName
 * @param {any[][]} rows
 * @param {string[]} headers
 * @param {string} [sheetName='Sheet1']
 */
export function downloadXlsx(fileName, rows, headers, sheetName = 'Sheet1') {
    const w = new BrowserXlsxWriter();
    w.addSheet(sheetName);
    w.writeSheet(rows, headers, true);
    triggerDownload(w.finalize(), fileName);
}

/**
 * Download multiple sheets as XLSX.
 * @param {string} fileName
 * @param {{ name: string, rows: any[][], headers: string[] }[]} sheets
 */
export function downloadXlsxMultiSheet(fileName, sheets) {
    const w = new BrowserXlsxWriter();
    for (const s of sheets) { w.addSheet(s.name); w.writeSheet(s.rows, s.headers, true); }
    triggerDownload(w.finalize(), fileName);
}

/**
 * Download data as XLSB. Headers = bold + frozen + auto-filter. Columns = autofit.
 * @param {string} fileName
 * @param {any[][]} rows
 * @param {string[]} headers
 * @param {string} [sheetName='Sheet1']
 */
export function downloadXlsb(fileName, rows, headers, sheetName = 'Sheet1') {
    const w = new BrowserXlsbWriter();
    w.addSheet(sheetName);
    w.writeSheet(rows, headers, true);
    triggerDownload(w.finalize(), fileName);
}

/**
 * Download multiple sheets as XLSB.
 * @param {string} fileName
 * @param {{ name: string, rows: any[][], headers: string[] }[]} sheets
 */
export function downloadXlsbMultiSheet(fileName, sheets) {
    const w = new BrowserXlsbWriter();
    for (const s of sheets) { w.addSheet(s.name); w.writeSheet(s.rows, s.headers, true); }
    triggerDownload(w.finalize(), fileName);
}
