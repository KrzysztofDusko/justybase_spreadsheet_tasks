/**
 * XlsbUpdater Integration Tests
 *
 * Verifies in-place replacement of worksheet data inside an existing XLSB
 * workbook (pivot tables, other sheets and shared strings must be preserved).
 */

import * as fs from 'fs';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { XlsbUpdater } from '../../src/XlsbUpdater';
import { XlsbWriter } from '../../src/XlsbWriter';
import { BiffReaderWriter } from '../../src/BiffReaderWriter';
import { parseSharedStringsBin, readRecord, readUtf16 } from '../../src/biff12Utils';

const repoRoot = path.join(__dirname, '..', '..');
const outputDir = path.join(repoRoot, 'test-output', 'integration', 'xlsb-updater');

interface TestResult {
    name: string;
    passed: boolean;
    message?: string;
}

const results: TestResult[] = [];

function assert(condition: boolean, testName: string, message?: string): void {
    if (condition) {
        results.push({ name: testName, passed: true });
        console.log(`✅ ${testName}`);
    } else {
        results.push({ name: testName, passed: false, message });
        console.log(`❌ ${testName}${message ? ': ' + message : ''}`);
    }
}

function assertEqual(actual: any, expected: any, testName: string): void {
    const isEqual = JSON.stringify(actual) === JSON.stringify(expected);
    if (isEqual) {
        results.push({ name: testName, passed: true });
        console.log(`✅ ${testName}`);
    } else {
        const msg = `Expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`;
        results.push({ name: testName, passed: false, message: msg });
        console.log(`❌ ${testName}: ${msg}`);
    }
}

function copyFixture(name: string, targetName?: string): string {
    const target = path.join(outputDir, targetName ?? name);
    fs.copyFileSync(path.join(repoRoot, name), target);
    return target;
}

function getEntry(filePath: string, name: string): Buffer | null {
    const zip = new AdmZip(filePath);
    const e = zip.getEntry(name);
    return e ? e.getData() : null;
}

/** Decode worksheet cells into rows of [col, xf, cellType, value]. */
function decodeSheet(filePath: string, sheetPath: string): Array<Array<[number, number, number, any]>> {
    const zip = new AdmZip(filePath);
    const sst = parseSharedStringsBin(zip.getEntry('xl/sharedStrings.bin')!.getData());
    const buf = zip.getEntry(sheetPath)!.getData();

    const rows: Array<Array<[number, number, number, any]>> = [];
    const bw = new BiffReaderWriter(buf);
    let row = -1;
    while (bw.readWorksheet()) {
        if (bw._rowIndex !== -1) row = bw._rowIndex;
        if (!bw._readCell) continue;
        let val: any = null;
        switch (bw._cellType) {
            case 2: val = sst.values[bw._intValue]; break;
            case 3: val = bw._doubleVal; break;
            case 4: val = bw._boolValue; break;
            case 5: val = bw._stringValue; break;
            default: val = null;
        }
        if (!rows[row]) rows[row] = [];
        rows[row][bw._columnNum] = [bw._columnNum, bw._xfIndex, bw._cellType, val];
    }
    return rows;
}

/** Read the pivot cache source record (0xbb) range for a cache definition. */
function readPivotSource(filePath: string, cachePath: string): { name: string; rwLast: number; colFirst: number; colLast: number } | null {
    const buf = getEntry(filePath, cachePath);
    if (!buf) return null;
    const rec = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
    let pos = 0;
    while (readRecord(buf, pos, rec)) {
        pos = rec.dataEnd;
        if (rec.id === 0x00bb && rec.len >= 23) {
            const cch = buf.readUInt32LE(rec.dataStart + 3);
            const nameStart = rec.dataStart + 7;
            const refStart = nameStart + cch * 2;
            return {
                name: readUtf16(buf, nameStart, cch),
                rwLast: buf.readInt32LE(refStart + 4),
                colFirst: buf.readInt32LE(refStart + 8),
                colLast: buf.readInt32LE(refStart + 12),
            };
        }
    }
    return null;
}

function refreshOnLoadSet(filePath: string, cachePath: string): boolean {
    const buf = getEntry(filePath, cachePath);
    if (!buf) return false;
    const rec = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
    let pos = 0;
    while (readRecord(buf, pos, rec)) {
        pos = rec.dataEnd;
        if (rec.id === 0x00b3 && rec.len >= 4) {
            return (buf[rec.dataStart + 3] & 0x04) !== 0;
        }
    }
    return false;
}

function countRows(filePath: string, sheetPath: string): number {
    const buf = getEntry(filePath, sheetPath);
    if (!buf) return -1;
    const rec = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
    let pos = 0;
    let n = 0;
    while (readRecord(buf, pos, rec)) {
        if (rec.id === 0x00) n++;
        pos = rec.dataEnd;
    }
    return n;
}

/** Read the sheet dimension record (0x98): rwLast @ payload+24, colLast @ payload+32. */
function readDimension(filePath: string, sheetPath: string): { rwLast: number; colLast: number } | null {
    const buf = getEntry(filePath, sheetPath);
    if (!buf) return null;
    const rec = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
    let pos = 0;
    while (readRecord(buf, pos, rec)) {
        pos = rec.dataEnd;
        if (rec.id === 0x0098 && rec.len >= 36) {
            return {
                rwLast: buf.readInt32LE(rec.dataStart + 24),
                colLast: buf.readInt32LE(rec.dataStart + 32),
            };
        }
    }
    return null;
}

/** True when the row-block terminator records (0x92/0x217/0x1dd/0x1dc) are present after the rows. */
function hasRowBlockTerminator(filePath: string, sheetPath: string): boolean {
    const buf = getEntry(filePath, sheetPath);
    if (!buf) return false;
    const rec = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
    let pos = 0;
    const seen: boolean[] = [false, false, false, false];
    while (readRecord(buf, pos, rec)) {
        if (rec.id === 0x0092) seen[0] = true;
        else if (rec.id === 0x0217) seen[1] = true;
        else if (rec.id === 0x01dd) seen[2] = true;
        else if (rec.id === 0x01dc) seen[3] = true;
        pos = rec.dataEnd;
    }
    return seen.every(Boolean);
}

// ==================== BASIC REPLACEMENT ====================

async function testBasicReplacement(): Promise<void> {
    console.log('\n--- Basic Replacement (report.xlsb data1) ---\n');

    const filePath = copyFixture('report.xlsb', 'basic.xlsb');
    const original = new AdmZip('report.xlsb');

    const rows: any[][] = [
        [1, 'Alice', 1500.5, new Date(2026, 0, 15), true, null],
        [2, 'Bob', 2400.25, new Date(2026, 1, 20), false, undefined],
        [3, 'Carol <&> "quoted"', null, new Date(2026, 2, 25), true, 'tail'],
    ];

    const updater = new XlsbUpdater(filePath);
    assertEqual(updater.getSheetNames(), ['report', 'data1', 'sql1'], 'sheet names preserved');
    updater.replaceSheetData('data1', rows, { headers: ['ID', 'NAME', 'AMOUNT', 'WHEN', 'FLAG', 'EXTRA'] });
    updater.save();

    // --- data values ---
    const dataRows = decodeSheet(filePath, 'xl/worksheets/sheet2.bin');
    assertEqual(countRows(filePath, 'xl/worksheets/sheet2.bin'), 4, 'exactly 4 rows (header + 3 data)');
    assertEqual(dataRows[0][0], [0, 1, 2, 'ID'], 'header A1 = ID with xf 1');
    assertEqual(dataRows[0][5], [5, 1, 2, 'EXTRA'], 'header F1 = EXTRA with xf 1');
    assertEqual(dataRows[1][0], [0, 0, 3, 1], 'A2 = 1 (number)');
    assertEqual(dataRows[1][1], [1, 0, 2, 'Alice'], 'B2 = Alice (string)');
    assertEqual(dataRows[1][2], [2, 0, 3, 1500.5], 'C2 = 1500.5 (real)');
    assertEqual(dataRows[1][4], [4, 0, 4, true], 'E2 = true (bool)');
    assertEqual(dataRows[2][0], [0, 0, 3, 2], 'A3 = 2');
    assertEqual(dataRows[3][1], [1, 0, 2, 'Carol <&> "quoted"'], 'B4 = escaped string');
    assertEqual(dataRows[3][3], [3, 0, 3, 46105.958333333336], 'D4 = date serial');
    assertEqual(dataRows[3][5], [5, 0, 2, 'tail'], 'F4 = tail');
    assertEqual(dataRows[1][5], undefined, 'F2 null skipped');
    assertEqual(dataRows[2][5], undefined, 'F3 undefined skipped');
    assertEqual(dataRows[3][2], undefined, 'C4 null skipped');

    // --- shared strings ---
    const sstNew = parseSharedStringsBin(getEntry(filePath, 'xl/sharedStrings.bin')!);
    const sstOld = parseSharedStringsBin(original.getEntry('xl/sharedStrings.bin')!.getData());
    assertEqual(sstNew.total, sstOld.total + 10, 'SST total incremented by 10 cells');
    assertEqual(sstNew.unique, sstOld.unique + 10, 'SST uniqueCount incremented by 10');
    assertEqual(sstNew.values[0], sstOld.values[0], 'first shared string untouched');
    assert(sstNew.values.includes('Alice'), 'Alice appended');
    assert(sstNew.values.includes('Carol <&> "quoted"'), 'escaped string appended');

    // --- sheet dimension + row-block terminator ---
    assertEqual(readDimension(filePath, 'xl/worksheets/sheet2.bin'),
        { rwLast: 3, colLast: 5 }, 'sheet dimension record updated (last row + last col)');
    assert(hasRowBlockTerminator(filePath, 'xl/worksheets/sheet2.bin'),
        'row-block terminator records preserved');

    // --- pivot integration ---
    const src = readPivotSource(filePath, 'xl/pivotCache/pivotCacheDefinition1.bin');
    assertEqual(src, { name: 'data1', rwLast: 3, colFirst: 0, colLast: 5 }, 'pivot cache source range updated');
    assert(refreshOnLoadSet(filePath, 'xl/pivotCache/pivotCacheDefinition1.bin'), 'refreshOnLoad flag set');

    // --- other parts byte-identical ---
    for (const part of [
        'xl/worksheets/sheet1.bin',
        'xl/worksheets/sheet3.bin',
        'xl/workbook.bin',
        'xl/styles.bin',
        'xl/theme/theme1.xml',
    ]) {
        const a = original.getEntry(part)!.getData();
        const b = getEntry(filePath, part)!;
        assert(a.equals(b), `byte-identical: ${part}`);
    }

    // --- file size shrank ---
    assert(fs.statSync(filePath).size < 300000, 'replaced sheet is much smaller than original');
}

// ==================== LARGER THAN ORIGINAL ====================

async function testLargerDataset(): Promise<void> {
    console.log('\n--- Larger Dataset (4000 rows × 20 cols) ---\n');

    const filePath = copyFixture('report.xlsb', 'larger.xlsb');
    const headers = Array.from({ length: 20 }, (_, i) => `COL${i + 1}`);
    const rows: any[][] = [];
    for (let i = 0; i < 4000; i++) {
        rows.push(Array.from({ length: 20 }, (_, c) => (c === 0 ? i + 1 : `${i}_${c}`)));
    }

    const updater = new XlsbUpdater(filePath);
    updater.replaceSheetData('data1', rows, { headers });
    updater.save();

    assertEqual(countRows(filePath, 'xl/worksheets/sheet2.bin'), 4001, '4001 rows total');
    const dataRows = decodeSheet(filePath, 'xl/worksheets/sheet2.bin');
    assertEqual(dataRows[1][0], [0, 0, 3, 1], 'row 1 A = 1');
    assertEqual(dataRows[4000][19], [19, 0, 2, '3999_19'], 'last cell = 3999_19');
    const src = readPivotSource(filePath, 'xl/pivotCache/pivotCacheDefinition1.bin');
    assertEqual(src!.rwLast, 4000, 'pivot rwLast = 4000');
    assertEqual(src!.colLast, 19, 'pivot colLast = 19');
}

// ==================== EMPTY DATASET ====================

async function testEmptyDataset(): Promise<void> {
    console.log('\n--- Empty Dataset ---\n');

    const filePath = copyFixture('report.xlsb', 'empty.xlsb');
    const updater = new XlsbUpdater(filePath);
    updater.replaceSheetData('data1', []);
    updater.save();

    assertEqual(countRows(filePath, 'xl/worksheets/sheet2.bin'), 0, 'no rows after clearing');
    const src = readPivotSource(filePath, 'xl/pivotCache/pivotCacheDefinition1.bin');
    assertEqual(src!.rwLast, 0, 'pivot rwLast = 0');
    assertEqual(src!.colLast, 0, 'pivot colLast = 0');
}

// ==================== MULTI-SHEET ====================

async function testMultipleSheets(): Promise<void> {
    console.log('\n--- Multiple Sheets Updated ---\n');

    const filePath = copyFixture('report.xlsb', 'multi.xlsb');
    const updater = new XlsbUpdater(filePath);
    updater.replaceSheetData('data1', [[1, 'one']], { headers: ['N', 'W'] });
    updater.replaceSheetData('sql1', [[2, 'two']], { headers: ['N', 'W'] });
    updater.save();

    const d1 = decodeSheet(filePath, 'xl/worksheets/sheet2.bin');
    const s1 = decodeSheet(filePath, 'xl/worksheets/sheet3.bin');
    assertEqual(d1[1][0], [0, 0, 3, 1], 'data1 updated');
    assertEqual(s1[1][1], [1, 0, 2, 'two'], 'sql1 updated');
}

// ==================== STYLE INHERITANCE ====================

async function testStyleInheritance(): Promise<void> {
    console.log('\n--- Style Inheritance (dates / numbers) ---\n');

    const filePath = path.join(outputDir, 'styles.xlsb');
    const writer = new XlsbWriter(filePath);
    writer.addSheet('data1');
    writer.writeSheet([
        ['DATA', 'AMOUNT'],
        [new Date(2024, 0, 1), 100.5],
        [new Date(2024, 5, 15), 200.25],
    ]);
    await writer.finalize();

    const updater = new XlsbUpdater(filePath);
    updater.replaceSheetData('data1', [
        [new Date(2025, 11, 24), 42.5],
        [new Date(2026, 3, 1), 7],
    ], { headers: ['DATA', 'AMOUNT'] });
    updater.save();

    const dataRows = decodeSheet(filePath, 'xl/worksheets/sheet1.bin');
    assertEqual(dataRows[1][0][1], 1, 'date cell keeps the date xf (1)');
    assertEqual(dataRows[1][0][2], 3, 'date cell stored as a real number');
    assert(dataRows[1][0][3] > 45000 && dataRows[1][0][3] < 50000, 'date serial in 2025/2026 range');
    assertEqual(dataRows[2][1], [1, 0, 3, 7], 'number cell uses column style');
}

// ==================== AUTO-FILTER ====================

async function testAutoFilter(): Promise<void> {
    console.log('\n--- Auto-Filter Range Updated ---\n');

    if (!fs.existsSync(fixture('autofilter.xlsb'))) {
        console.log('⚠️  SKIPPED: fixture autofilter.xlsb not present');
        return;
    }
    const filePath = copyFixture('autofilter.xlsb', 'autofilter.bin');
    const updater = new XlsbUpdater(filePath);
    updater.replaceSheetData('data1', [[1, 'x'], [2, 'y'], [3, 'z']], { headers: ['A', 'B'] });
    updater.save();

    const buf = getEntry(filePath, 'xl/worksheets/sheet2.bin')!;
    const rec = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
    let pos = 0;
    let af: { startRow: number; endRow: number; startCol: number; endCol: number } | null = null;
    while (readRecord(buf, pos, rec)) {
        pos = rec.dataEnd;
        if (rec.id === 0x00a1 && rec.len >= 16) {
            af = {
                startRow: buf.readInt32LE(rec.dataStart),
                endRow: buf.readInt32LE(rec.dataStart + 4),
                startCol: buf.readInt32LE(rec.dataStart + 8),
                endCol: buf.readInt32LE(rec.dataStart + 12),
            };
        }
    }
    assertEqual(af, { startRow: 0, endRow: 3, startCol: 0, endCol: 1 }, 'autofilter range follows new data');
}

// ==================== MULTIPLE PIVOT CACHES ====================

async function testMultiplePivotCaches(): Promise<void> {
    console.log('\n--- Multiple Pivot Caches ---\n');

    if (!fs.existsSync(fixture('pivot_multi.xlsb'))) {
        console.log('⚠️  SKIPPED: fixture pivot_multi.xlsb not present');
        return;
    }
    const filePath = copyFixture('pivot_multi.xlsb', 'pivotmulti.xlsb');
    const updater = new XlsbUpdater(filePath);
    updater.replaceSheetData('data1', Array.from({ length: 20 }, (_, i) =>
        Array.from({ length: 20 }, (_, c) => (c === 0 ? i + 1 : `${i}_${c}`))
    ));
    updater.save();

    assertEqual(readPivotSource(filePath, 'xl/pivotCache/pivotCacheDefinition1.bin'),
        { name: 'data1', rwLast: 19, colFirst: 0, colLast: 19 }, 'cache 1 range updated (no headers → rwLast = rows-1)');
    assertEqual(readPivotSource(filePath, 'xl/pivotCache/pivotCacheDefinition2.bin'),
        { name: 'data1', rwLast: 19, colFirst: 11, colLast: 19 }, 'cache 2 range updated (colFirst preserved)');
    assert(refreshOnLoadSet(filePath, 'xl/pivotCache/pivotCacheDefinition1.bin'), 'cache 1 refreshOnLoad');
    assert(refreshOnLoadSet(filePath, 'xl/pivotCache/pivotCacheDefinition2.bin'), 'cache 2 refreshOnLoad');
}

// ==================== TRAILING EMPTY ROWS ====================

async function testTrailingEmptyRows(): Promise<void> {
    console.log('\n--- Trailing Empty Rows (no "(blank)" in pivot) ---\n');

    const filePath = copyFixture('report.xlsb', 'trailing.xlsb');
    const rows: any[][] = [
        [1, 'a'],
        [2, 'b'],
        [3, 'c'],
        [null, null],
        [undefined, null],
    ];

    const updater = new XlsbUpdater(filePath);
    updater.replaceSheetData('data1', rows);
    updater.save();

    assertEqual(countRows(filePath, 'xl/worksheets/sheet2.bin'), 3, 'trailing empty rows are not written');
    const dataRows = decodeSheet(filePath, 'xl/worksheets/sheet2.bin');
    assertEqual(dataRows[2][1], [1, 0, 2, 'c'], 'last real row kept');
    const src = readPivotSource(filePath, 'xl/pivotCache/pivotCacheDefinition1.bin');
    assertEqual(src!.rwLast, 2, 'pivot range ends at the last real row (rows-1, no headers)');

    // middle empty rows are preserved
    const filePath2 = copyFixture('report.xlsb', 'trailing2.xlsb');
    const updater2 = new XlsbUpdater(filePath2);
    updater2.replaceSheetData('data1', [[1, 'x'], [], [2, 'y']], { headers: ['A', 'B'] });
    updater2.save();
    assertEqual(countRows(filePath2, 'xl/worksheets/sheet2.bin'), 4, 'middle empty row preserved (header + 3 rows)');
}

// ==================== ERRORS ====================

async function testErrors(): Promise<void> {
    console.log('\n--- Error Handling ---\n');

    const filePath = copyFixture('report.xlsb', 'errors.xlsb');
    const updater = new XlsbUpdater(filePath);

    let threw = false;
    try {
        updater.replaceSheetData('missing', []);
    } catch (e: any) {
        threw = true;
        assert(String(e.message).includes('not found'), 'missing sheet throws descriptive error');
    }
    assert(threw, 'missing sheet throws');

    // .xlsx rejected
    const xlsxPath = copyFixture('report.xlsx', 'report.xlsx.bin');
    threw = false;
    try {
        new XlsbUpdater(xlsxPath);
    } catch (e: any) {
        threw = true;
        assert(String(e.message).includes('XLSB'), 'xlsx file rejected with XLSB error');
    }
    assert(threw, 'xlsx file rejected');

    // missing file
    threw = false;
    try {
        new XlsbUpdater(path.join(outputDir, 'does-not-exist.xlsb'));
    } catch (e: any) {
        threw = true;
    }
    assert(threw, 'missing input file throws');
}

// ==================== MAIN ====================

function fixture(name: string): string {
    return path.join(repoRoot, name);
}

async function runXlsbUpdaterTests(): Promise<void> {
    if (!fs.existsSync(fixture('report.xlsb'))) {
        console.log('⚠️  SKIPPED: fixture report.xlsb not present (gitignored, generated in Excel).');
        console.log('To run: create an XLSB with sheets report/data1/sql1 and a pivot table on "report".');
        return;
    }
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }

    try {
        await testBasicReplacement();
        await testLargerDataset();
        await testEmptyDataset();
        await testMultipleSheets();
        await testStyleInheritance();
        await testAutoFilter();
        await testMultiplePivotCaches();
        await testTrailingEmptyRows();
        await testErrors();
    } catch (err) {
        console.error('Test error:', err);
    }

    console.log('\n' + '='.repeat(60));
    const passed = results.filter(r => r.passed).length;
    const failed = results.filter(r => !r.passed).length;
    console.log(`XlsbUpdater Tests: ${passed} passed, ${failed} failed`);
    console.log('='.repeat(60));

    if (failed > 0) {
        for (const r of results.filter(r => !r.passed)) {
            console.error(`  FAILED: ${r.name} — ${r.message ?? ''}`);
        }
        process.exit(1);
    }
}

runXlsbUpdaterTests().catch(console.error);
