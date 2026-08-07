/**
 * XlsxUpdater Integration Tests
 *
 * Verifies in-place replacement of worksheet data inside an existing XLSX
 * workbook (pivot tables, other sheets and shared strings must be preserved).
 */

import * as fs from 'fs';
import * as path from 'path';
import AdmZip from 'adm-zip';
import ExcelJS from 'exceljs';
import { XlsxUpdater } from '../../src/XlsxUpdater';

const repoRoot = path.join(__dirname, '..', '..');
const fixturePath = path.join(repoRoot, 'report.xlsx');
const outputDir = path.join(repoRoot, 'test-output', 'integration', 'xlsx-updater');

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

function assertContains(haystack: string, needle: string, testName: string): void {
    if (haystack.includes(needle)) {
        results.push({ name: testName, passed: true });
        console.log(`✅ ${testName}`);
    } else {
        const msg = `Expected to find ${JSON.stringify(needle)}`;
        results.push({ name: testName, passed: false, message: msg });
        console.log(`❌ ${testName}: ${msg}`);
    }
}

function copyFixture(name: string): string {
    const target = path.join(outputDir, name);
    fs.copyFileSync(fixturePath, target);
    return target;
}

function getZipEntryText(filePath: string, entryName: string): string | null {
    const zip = new AdmZip(filePath);
    const entry = zip.getEntry(entryName);
    if (!entry) return null;
    return entry.getData().toString('utf8');
}

// ==================== BASIC REPLACEMENT ====================

async function testBasicReplacement(): Promise<void> {
    console.log('\n--- Basic Replacement (report.xlsx data1) ---\n');

    const filePath = copyFixture('basic.xlsx');

    const rows: any[][] = [
        [1, 'Alice', 1500.5, new Date(2026, 0, 15), true, null],
        [2, 'Bob', 2400.25, new Date(2026, 1, 20), false, undefined],
        [3, 'Carol <&> "quoted"', null, new Date(2026, 2, 25), true, 'tail'],
    ];

    const updater = new XlsxUpdater(filePath);
    assertEqual(updater.getSheetNames(), ['report', 'data1', 'sql1'], 'sheet names preserved');

    updater.replaceSheetData('data1', rows, { headers: ['ID', 'NAME', 'AMOUNT', 'WHEN', 'FLAG', 'EXTRA'] });
    updater.save();

    // --- sheet XML assertions ---
    const sheetXml = getZipEntryText(filePath, 'xl/worksheets/sheet2.xml')!;
    assertContains(sheetXml, '<dimension ref="A1:F4"/>', 'dimension updated to A1:F4');

    const sdStart = sheetXml.indexOf('<sheetData>');
    const sdEnd = sheetXml.indexOf('</sheetData>');
    const sd = sheetXml.substring(sdStart, sdEnd);
    assertContains(sd, '<row r="1">', 'header row written at r=1');
    assertContains(sd, '<c r="D2"', 'date cell D2 written');
    assertContains(sd, '<c r="E2" t="b"', 'boolean cell E2 written');
    assert(!sd.includes('<c r="F2"'), 'null cell F2 skipped');
    assert(!sd.includes('<c r="F3"'), 'undefined cell F3 skipped');
    assert(!sd.includes('<c r="C4"'), 'null cell C4 skipped');
    assertContains(sd, '<c r="B4" t="s"', 'string with special chars written as shared string');

    // header row inherits the original header style (s="1" from report.xlsx)
    assertContains(sd, '<c r="A1" t="s" s="1">', 'header cell inherits original header style');

    // --- sharedStrings ---
    const sst = getZipEntryText(filePath, 'xl/sharedStrings.xml')!;
    const sstTag = /<sst\b[^>]*>/.exec(sst)![0];
    assertContains(sstTag, 'count="73086"', 'shared string count incremented (73076 + 10 cells)');
    assertContains(sstTag, 'uniqueCount="7777"', 'shared string uniqueCount incremented (7767 + 10)');
    const firstSi = sst.substring(sst.indexOf('<si>'), sst.indexOf('</si>') + 5);
    assertEqual(firstSi, '<si><t>SELECT CURRENT_TIMESTAMP,* FROM DIMDATE</t></si>', 'original shared strings untouched');
    assertContains(sst, '<si><t>Alice</t></si>', 'new string appended');
    assertContains(sst, '<si><t>Carol &lt;&amp;&gt; &quot;quoted&quot;</t></si>', 'special chars escaped in shared strings');

    // --- pivot integration ---
    const pc = getZipEntryText(filePath, 'xl/pivotCache/pivotCacheDefinition1.xml')!;
    assertContains(pc, '<worksheetSource ref="A1:F4" sheet="data1"/>', 'pivot cache source ref updated');
    assertContains(pc, 'recordCount="3"', 'pivot cache recordCount updated');

    const pt = getZipEntryText(filePath, 'xl/pivotTables/pivotTable1.xml')!;
    assertContains(pt, 'refreshOnLoad="1"', 'refreshOnLoad added to pivot table');

    // --- other parts untouched ---
    const repBefore = getZipEntryText(fixturePath, 'xl/worksheets/sheet1.xml')!;
    const repAfter = getZipEntryText(filePath, 'xl/worksheets/sheet1.xml')!;
    assertEqual(repAfter, repBefore, 'report sheet XML byte-identical');
    const sqlBefore = getZipEntryText(fixturePath, 'xl/worksheets/sheet3.xml')!;
    const sqlAfter = getZipEntryText(filePath, 'xl/worksheets/sheet3.xml')!;
    assertEqual(sqlAfter, sqlBefore, 'sql1 sheet XML byte-identical');

    // --- semantic verification with exceljs ---
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    assertEqual(wb.worksheets.map((w: any) => w.name), ['report', 'data1', 'sql1'], 'exceljs: sheet names preserved');

    const ws = wb.getWorksheet('data1')!;
    assertEqual(ws.getCell('A1').value, 'ID', 'exceljs: header A1');
    assertEqual(ws.getCell('A2').value, 1, 'exceljs: number A2');
    assertEqual(ws.getCell('B2').value, 'Alice', 'exceljs: string B2');
    assertEqual(ws.getCell('C2').value, 1500.5, 'exceljs: float C2');
    assertEqual(ws.getCell('B3').value, 'Bob', 'exceljs: string B3');
    assertEqual(ws.getCell('C4').value, null, 'exceljs: null cell stays empty');
    assertEqual(ws.getCell('F4').value, 'tail', 'exceljs: last column value');
    assertEqual(ws.getCell('B4').value, 'Carol <&> "quoted"', 'exceljs: escaped chars read back');
    assertEqual(ws.actualRowCount, 4, 'exceljs: row count 4');
}

// ==================== LARGER THAN ORIGINAL ====================

async function testLargerThanOriginal(): Promise<void> {
    console.log('\n--- Larger Dataset (4000 rows) ---\n');

    const filePath = copyFixture('larger.xlsx');
    const rows: any[][] = [];
    for (let i = 0; i < 4000; i++) {
        rows.push([i, `Name${i}`, i * 1.5, `KEY${i % 100}`]);
    }

    const updater = new XlsxUpdater(filePath);
    updater.replaceSheetData('data1', rows);
    updater.save();

    const sheetXml = getZipEntryText(filePath, 'xl/worksheets/sheet2.xml')!;
    assertContains(sheetXml, '<dimension ref="A1:D4000"/>', 'dimension grows to A1:D4000');

    const pc = getZipEntryText(filePath, 'xl/pivotCache/pivotCacheDefinition1.xml')!;
    assertContains(pc, '<worksheetSource ref="A1:D4000" sheet="data1"/>', 'pivot cache ref grows');
    assertContains(pc, 'recordCount="4000"', 'pivot cache recordCount = 4000');

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet('data1')!;
    assertEqual(ws.actualRowCount, 4000, 'exceljs: 4000 rows');
    assertEqual(ws.getCell('D4000').value, 'KEY99', 'exceljs: last cell value');
}

// ==================== EMPTY REPLACEMENT ====================

async function testEmptyReplacement(): Promise<void> {
    console.log('\n--- Empty Dataset ---\n');

    const filePath = copyFixture('empty.xlsx');
    const updater = new XlsxUpdater(filePath);
    updater.replaceSheetData('data1', []);
    updater.save();

    const sheetXml = getZipEntryText(filePath, 'xl/worksheets/sheet2.xml')!;
    assertContains(sheetXml, '<dimension ref="A1"/>', 'dimension reset to A1');

    const pc = getZipEntryText(filePath, 'xl/pivotCache/pivotCacheDefinition1.xml')!;
    assertContains(pc, '<worksheetSource ref="A1" sheet="data1"/>', 'pivot cache ref reset to A1');
    assertContains(pc, 'recordCount="0"', 'pivot cache recordCount = 0');

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    assertEqual(wb.getWorksheet('data1')!.actualRowCount, 0, 'exceljs: no rows');
}

// ==================== MULTI-SHEET ====================

async function testMultipleSheets(): Promise<void> {
    console.log('\n--- Multiple Sheets Updated ---\n');

    const filePath = copyFixture('multi.xlsx');
    const updater = new XlsxUpdater(filePath);
    updater.replaceSheetData('data1', [[1, 'one']], { headers: ['N', 'W'] });
    updater.replaceSheetData('sql1', [[2, 'two']], { headers: ['N', 'W'] });
    updater.save();

    const sheet2 = getZipEntryText(filePath, 'xl/worksheets/sheet2.xml')!;
    const sheet3 = getZipEntryText(filePath, 'xl/worksheets/sheet3.xml')!;
    assertContains(sheet2, '<dimension ref="A1:B2"/>', 'data1 dimension updated');
    assertContains(sheet3, '<dimension ref="A1:B2"/>', 'sql1 dimension updated');
    assertContains(sheet2, '<c r="A2"><v>1</v></c>', 'data1 value written');
    assertContains(sheet3, '<c r="A2"><v>2</v></c>', 'sql1 value written');

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    assertEqual(wb.getWorksheet('data1')!.getCell('B2').value, 'one', 'exceljs: data1 updated');
    assertEqual(wb.getWorksheet('sql1')!.getCell('B2').value, 'two', 'exceljs: sql1 updated');
}

// ==================== STYLE INHERITANCE ====================

async function createStyledFixture(filePath: string): Promise<void> {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('data1');
    ws.getCell('A1').value = 'DATA';
    ws.getCell('B1').value = 'AMOUNT';
    ws.getCell('A2').value = new Date(2024, 0, 1);
    ws.getCell('A2').numFmt = 'dd.mm.yyyy';
    ws.getCell('A3').value = new Date(2024, 5, 15);
    ws.getCell('A3').numFmt = 'dd.mm.yyyy';
    ws.getCell('B2').value = 100.5;
    ws.getCell('B2').numFmt = '#,##0.00';
    ws.getCell('B3').value = 200.25;
    ws.getCell('B3').numFmt = '#,##0.00';
    ws.getCell('C2').value = 'plain';
    await wb.xlsx.writeFile(filePath);
}

async function testStyleInheritance(): Promise<void> {
    console.log('\n--- Style Inheritance (dates / numbers) ---\n');

    const filePath = path.join(outputDir, 'styles.xlsx');
    await createStyledFixture(filePath);

    const updater = new XlsxUpdater(filePath);
    updater.replaceSheetData('data1', [
        [new Date(2025, 11, 24), 42.5, 'new'],
        [new Date(2026, 3, 1), 7, null],
    ], { headers: ['DATA', 'AMOUNT', 'EXTRA'] });
    updater.save();

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet('data1')!;

    const dateVal = ws.getCell('A2').value;
    const numVal = ws.getCell('B2').value;
    assert(dateVal instanceof Date, 'inherited column style keeps A2 a date');
    if (dateVal instanceof Date) {
        assertEqual(dateVal.getTime(), new Date(2025, 11, 24).getTime(), 'A2 date value correct');
    }
    assertEqual(typeof numVal, 'number', 'inherited column style keeps B2 numeric');
    assertEqual(ws.getCell('A3').value instanceof Date, true, 'inherited column style keeps A3 a date');
    assertEqual(ws.getCell('C2').value, 'new', 'plain string column still readable');

    // styleFallback 'general' produces no s attributes on new data cells
    const filePath2 = path.join(outputDir, 'styles-general.xlsx');
    await createStyledFixture(filePath2);
    const updater2 = new XlsxUpdater(filePath2);
    updater2.replaceSheetData('data1', [[new Date(2025, 11, 24), 42.5, 'new']], { styleFallback: 'general' });
    updater2.save();

    const sheetXml = getZipEntryText(filePath2, 'xl/worksheets/sheet1.xml')!;
    const sd = sheetXml.substring(sheetXml.indexOf('<sheetData>'), sheetXml.indexOf('</sheetData>'));
    assert(!/<c r="[A-Z]+2"[^>]*s=/.test(sd), 'styleFallback general: no s attribute on new cells');
}

// ==================== INLINE STRINGS FALLBACK ====================

async function createInlineStrFixture(filePath: string): Promise<void> {
    const zip = new AdmZip();
    zip.addFile('[Content_Types].xml', Buffer.from(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`
    ));
    zip.addFile('_rels/.rels', Buffer.from(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
    ));
    zip.addFile('xl/workbook.xml', Buffer.from(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="data1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`
    ));
    zip.addFile('xl/_rels/workbook.xml.rels', Buffer.from(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
    ));
    zip.addFile('xl/styles.xml', Buffer.from(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>`
    ));
    zip.addFile('xl/worksheets/sheet1.xml', Buffer.from(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1:B2"/>
<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>OLD</t></is></c><c r="B1" t="inlineStr"><is><t>DATA</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>old value</t></is></c></row></sheetData>
</worksheet>`
    ));
    zip.writeZip(filePath);
}

async function testInlineStrFallback(): Promise<void> {
    console.log('\n--- Inline Strings Fallback (no sharedStrings.xml) ---\n');

    const filePath = path.join(outputDir, 'inlinestr.xlsx');
    await createInlineStrFixture(filePath);

    const updater = new XlsxUpdater(filePath);
    updater.replaceSheetData('data1', [['new value 1', 7], ['new value 2', 8]], { headers: ['A', 'B'] });
    updater.save();

    const sheetXml = getZipEntryText(filePath, 'xl/worksheets/sheet1.xml')!;
    assertContains(sheetXml, '<c r="A2" t="inlineStr"><is><t>new value 1</t></is></c>', 'new string as inlineStr');
    assertContains(sheetXml, '<c r="B2"><v>7</v></c>', 'number cell written');
    assertEqual(getZipEntryText(filePath, 'xl/sharedStrings.xml'), null, 'no sharedStrings.xml created');

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet('data1')!;
    assertEqual(ws.getCell('A1').value, 'A', 'exceljs: header A1');
    assertEqual(ws.getCell('A2').value, 'new value 1', 'exceljs: inline string read back');
    assertEqual(ws.getCell('B2').value, 7, 'exceljs: number read back');
}

// ==================== ERROR HANDLING ====================

async function testErrors(): Promise<void> {
    console.log('\n--- Error Handling ---\n');

    const filePath = copyFixture('errors.xlsx');
    const updater = new XlsxUpdater(filePath);

    let threw = false;
    try {
        updater.replaceSheetData('missing', []);
    } catch (e: any) {
        threw = true;
        assertContains(String(e.message), 'not found', 'missing sheet throws descriptive error');
    }
    assert(threw, 'missing sheet throws');

    // XLSB rejection: fake xlsb package
    const xlsbPath = path.join(outputDir, 'fake.xlsb');
    const zip = new AdmZip();
    zip.addFile('xl/workbook.bin', Buffer.from([1, 2, 3]));
    zip.writeZip(xlsbPath);

    threw = false;
    try {
        new XlsxUpdater(xlsbPath);
    } catch (e: any) {
        threw = true;
        assertContains(String(e.message), 'XLSB', 'xlsb file throws descriptive error');
    }
    assert(threw, 'xlsb file throws');

    // missing file
    threw = false;
    try {
        new XlsxUpdater(path.join(outputDir, 'does-not-exist.xlsx'));
    } catch (e: any) {
        threw = true;
    }
    assert(threw, 'missing input file throws');
}

// ==================== SELF-CLOSING SHEETDATA ====================

async function testSelfClosingSheetData(): Promise<void> {
    console.log('\n--- Self-Closing <sheetData/> ---\n');

    const filePath = path.join(outputDir, 'selfclosing.xlsx');
    const zip = new AdmZip();
    zip.addFile('[Content_Types].xml', Buffer.from(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`
    ));
    zip.addFile('_rels/.rels', Buffer.from(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
    ));
    zip.addFile('xl/workbook.xml', Buffer.from(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="data1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`
    ));
    zip.addFile('xl/_rels/workbook.xml.rels', Buffer.from(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`
    ));
    zip.addFile('xl/worksheets/sheet1.xml', Buffer.from(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1"/>
<sheetData/>
</worksheet>`
    ));
    zip.writeZip(filePath);

    const updater = new XlsxUpdater(filePath);
    updater.replaceSheetData('data1', [[1, 'a'], [2, 'b']]);
    updater.save();

    const sheetXml = getZipEntryText(filePath, 'xl/worksheets/sheet1.xml')!;
    assertContains(sheetXml, '<dimension ref="A1:B2"/>', 'dimension updated');
    assertContains(sheetXml, '<sheetData><row r="1">', 'rows written into sheetData');
    assertContains(sheetXml, '<c r="B2" t="inlineStr"><is><t>b</t></is></c>', 'inline string written');

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    const ws = wb.getWorksheet('data1')!;
    assertEqual(ws.getCell('A2').value, 2, 'exceljs: second row read');
    assertEqual(ws.getCell('B1').value, 'a', 'exceljs: first row string read');
}

// ==================== MAIN ====================

async function runXlsxUpdaterTests(): Promise<void> {
    if (!fs.existsSync(fixturePath)) {
        console.log(`⚠️  SKIPPED: fixture ${fixturePath} not present (gitignored, generated in Excel)`);
        console.log('To run: create an XLSX with sheets report/data1/sql1 and a pivot table on "report".');
        return;
    }
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }

    try {
        await testBasicReplacement();
        await testLargerThanOriginal();
        await testEmptyReplacement();
        await testMultipleSheets();
        await testStyleInheritance();
        await testInlineStrFallback();
        await testSelfClosingSheetData();
        await testErrors();
    } catch (err) {
        console.error('Test error:', err);
    }

    console.log('\n' + '='.repeat(60));
    const passed = results.filter(r => r.passed).length;
    const failed = results.filter(r => !r.passed).length;
    console.log(`XlsxUpdater Tests: ${passed} passed, ${failed} failed`);
    console.log('='.repeat(60));

    if (failed > 0) {
        for (const r of results.filter(r => !r.passed)) {
            console.error(`  FAILED: ${r.name} — ${r.message ?? ''}`);
        }
        process.exit(1);
    }
}

runXlsxUpdaterTests().catch(console.error);
