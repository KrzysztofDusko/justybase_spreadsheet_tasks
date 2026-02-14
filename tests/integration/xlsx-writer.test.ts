/**
 * XLSX Writer Integration Tests
 */

import { XlsxWriter } from '../../dist/XlsxWriter';
import { XlsxReader } from '../../dist/XlsxReader';
import * as fs from 'fs';
import * as path from 'path';

const outputDir = path.join(__dirname, '..', '..', 'test-output', 'integration', 'xlsx-writer');

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

function assertApprox(actual: number, expected: number, tolerance: number, testName: string): void {
    const diff = Math.abs(actual - expected);
    if (diff <= tolerance) {
        results.push({ name: testName, passed: true });
        console.log(`✅ ${testName}`);
    } else {
        const msg = `Expected ~${expected}, got ${actual} (diff: ${diff})`;
        results.push({ name: testName, passed: false, message: msg });
        console.log(`❌ ${testName}: ${msg}`);
    }
}

// ==================== DATA TYPE TESTS ====================

async function testIntegerValues(): Promise<void> {
    console.log('\n--- Integer Values Test ---\n');

    const filePath = path.join(outputDir, 'integers.xlsx');
    const data = [
        [0, 'zero'],
        [1, 'one'],
        [-1, 'minus one'],
        [1000, 'thousand'],
        [-9999, 'negative'],
        [2147483647, 'max int32'],
        [-2147483648, 'min int32'],
    ];

    const writer = new XlsxWriter(filePath);
    writer.addSheet('Integers');
    writer.writeSheet(data, ['Value', 'Description'], false);
    await writer.finalize();

    assert(fs.existsSync(filePath), 'Integer file created');

    const reader = new XlsxReader();
    await reader.open(filePath);

    const rows: any[][] = [];
    while (await reader.read()) {
        const row = [reader.getValue(0), reader.getValue(1)];
        rows.push(row);
    }

    const dataRows = rows.slice(1);
    assertEqual(dataRows.length, data.length, 'Integer row count');

    for (let i = 0; i < data.length; i++) {
        assertEqual(dataRows[i][0], data[i][0], `Integer value ${i}: ${data[i][0]}`);
        assertEqual(dataRows[i][1], data[i][1], `Integer desc ${i}`);
    }

    await reader.close();
}

async function testFloatValues(): Promise<void> {
    console.log('\n--- Float Values Test ---\n');

    const filePath = path.join(outputDir, 'floats.xlsx');
    const data = [
        [0.0, 'zero'],
        [1.5, 'one point five'],
        [-3.14159, 'negative pi approx'],
        [1234.5678, 'large decimal'],
        [0.000001, 'very small'],
        [999999.999999, 'large with decimals'],
    ];

    const writer = new XlsxWriter(filePath);
    writer.addSheet('Floats');
    writer.writeSheet(data, ['Value', 'Description'], false);
    await writer.finalize();

    assert(fs.existsSync(filePath), 'Float file created');

    const reader = new XlsxReader();
    await reader.open(filePath);

    const rows: any[][] = [];
    while (await reader.read()) {
        const row = [reader.getValue(0), reader.getValue(1)];
        rows.push(row);
    }

    const dataRows = rows.slice(1);
    assertEqual(dataRows.length, data.length, 'Float row count');

    for (let i = 0; i < data.length; i++) {
        const actualVal = dataRows[i][0];
        const expectedVal = data[i][0] as number;
        if (typeof actualVal === 'number' && typeof expectedVal === 'number') {
            const tolerance = Math.abs(expectedVal) * 1e-10 + 1e-15;
            assertApprox(actualVal, expectedVal, tolerance, `Float value ${i}: ${expectedVal}`);
        }
        assertEqual(dataRows[i][1], data[i][1], `Float desc ${i}`);
    }

    await reader.close();
}

async function testStringValues(): Promise<void> {
    console.log('\n--- String Values Test ---\n');

    const filePath = path.join(outputDir, 'strings.xlsx');
    const data = [
        ['Hello World', 'basic string'],
        ['', 'empty string'],
        ['   ', 'whitespace only'],
        ['ąęćśńźółĄĘĆŚŃŹÓŁ', 'Polish chars'],
        ['日本語テスト', 'Japanese'],
        ['中文测试', 'Chinese'],
        ['Ελληνικά', 'Greek'],
        ['🎉🔥💻', 'Emojis'],
        ['Line1\nLine2', 'Newline in string'],
        ['Tab\there', 'Tab in string'],
        ['Quote"test', 'Double quote'],
        ["Single'quote", 'Single quote'],
        ['A'.repeat(1000), 'Long string (1000 chars)'],
    ];

    const writer = new XlsxWriter(filePath);
    writer.addSheet('Strings');
    writer.writeSheet(data, ['Value', 'Description'], false);
    await writer.finalize();

    assert(fs.existsSync(filePath), 'String file created');

    const reader = new XlsxReader();
    await reader.open(filePath);

    const rows: any[][] = [];
    while (await reader.read()) {
        const row = [reader.getValue(0), reader.getValue(1)];
        rows.push(row);
    }

    const dataRows = rows.slice(1);
    assertEqual(dataRows.length, data.length, 'String row count');

    for (let i = 0; i < data.length; i++) {
        assertEqual(dataRows[i][0], data[i][0], `String value ${i}: ${data[i][1]}`);
    }

    await reader.close();
}

async function testBooleanValues(): Promise<void> {
    console.log('\n--- Boolean Values Test ---\n');

    const filePath = path.join(outputDir, 'booleans.xlsx');
    const data = [
        [true, 'true value'],
        [false, 'false value'],
        [true, true],
        [false, false],
    ];

    const writer = new XlsxWriter(filePath);
    writer.addSheet('Booleans');
    writer.writeSheet(data, ['Value', 'Description'], false);
    await writer.finalize();

    assert(fs.existsSync(filePath), 'Boolean file created');

    const reader = new XlsxReader();
    await reader.open(filePath);

    const rows: any[][] = [];
    while (await reader.read()) {
        const row = [reader.getValue(0), reader.getValue(1)];
        rows.push(row);
    }

    const dataRows = rows.slice(1);
    assertEqual(dataRows.length, data.length, 'Boolean row count');

    assertEqual(dataRows[0][0], true, 'Boolean true value');
    assertEqual(dataRows[1][0], false, 'Boolean false value');

    await reader.close();
}

async function testDateValues(): Promise<void> {
    console.log('\n--- Date Values Test ---\n');

    const filePath = path.join(outputDir, 'dates.xlsx');
    const data = [
        [new Date('2024-01-01'), 'New Year 2024'],
        [new Date('2024-12-31'), 'New Year Eve 2024'],
        [new Date('2000-01-01'), 'Y2K'],
        [new Date('2024-06-15T12:30:45'), 'With time'],
        [new Date('2024-02-29'), 'Leap year'],
    ];

    const writer = new XlsxWriter(filePath);
    writer.addSheet('Dates');
    writer.writeSheet(data, ['Value', 'Description'], false);
    await writer.finalize();

    assert(fs.existsSync(filePath), 'Date file created');

    const reader = new XlsxReader();
    await reader.open(filePath);

    const rows: any[][] = [];
    while (await reader.read()) {
        const row = [reader.getValue(0), reader.getValue(1)];
        rows.push(row);
    }

    const dataRows = rows.slice(1);
    assertEqual(dataRows.length, data.length, 'Date row count');

    for (let i = 0; i < data.length; i++) {
        const actualVal = dataRows[i][0];
        const expectedVal = data[i][0];
        if (actualVal instanceof Date && expectedVal instanceof Date) {
            const diff = Math.abs(actualVal.getTime() - expectedVal.getTime());
            assert(diff < 86400000, `Date value ${i}: ${expectedVal.toISOString()}`, `Diff: ${diff}ms`);
        }
    }

    await reader.close();
}

// ==================== SHEET TESTS ====================

async function testMultipleSheets(): Promise<void> {
    console.log('\n--- Multiple Sheets Test ---\n');

    const filePath = path.join(outputDir, 'multi_sheet.xlsx');

    const writer = new XlsxWriter(filePath);
    
    writer.addSheet('Sheet1');
    writer.writeSheet([[1, 'A']], ['Col1', 'Col2'], false);
    
    writer.addSheet('Sheet2');
    writer.writeSheet([[2, 'B']], ['Col1', 'Col2'], false);
    
    writer.addSheet('Sheet3');
    writer.writeSheet([[3, 'C']], ['Col1', 'Col2'], false);
    
    await writer.finalize();

    assert(fs.existsSync(filePath), 'Multi-sheet file created');

    const reader = new XlsxReader();
    await reader.open(filePath);

    const sheetNames = reader.getSheetNames();
    assertEqual(sheetNames.length, 3, 'Multi-sheet count');
    assertEqual(sheetNames, ['Sheet1', 'Sheet2', 'Sheet3'], 'Multi-sheet names');

    await reader.close();
}

async function testSheetNameSanitization(): Promise<void> {
    console.log('\n--- Sheet Name Sanitization Test ---\n');

    const filePath = path.join(outputDir, 'sheet_names.xlsx');

    const writer = new XlsxWriter(filePath);
    
    writer.addSheet('Sheet/With/Slashes');
    writer.writeSheet([[1]], ['A'], false);
    
    writer.addSheet('Sheet*With*Stars');
    writer.writeSheet([[2]], ['A'], false);
    
    writer.addSheet('Sheet?With?Questions');
    writer.writeSheet([[3]], ['A'], false);
    
    writer.addSheet('Sheet[With]Brackets');
    writer.writeSheet([[4]], ['A'], false);
    
    writer.addSheet('Sheet:With:Colons');
    writer.writeSheet([[5]], ['A'], false);
    
    writer.addSheet('Sheet\\With\\Backslashes');
    writer.writeSheet([[6]], ['A'], false);
    
    await writer.finalize();

    assert(fs.existsSync(filePath), 'Sheet names file created');

    const reader = new XlsxReader();
    await reader.open(filePath);

    const sheetNames = reader.getSheetNames();
    assertEqual(sheetNames.length, 6, 'Sanitized sheet count');

    for (const name of sheetNames) {
        const hasInvalid = /[\\/*?[\]:]/.test(name);
        assert(!hasInvalid, `Sheet name "${name}" has no invalid chars`);
    }

    await reader.close();
}

async function testEmptySheet(): Promise<void> {
    console.log('\n--- Empty Sheet Test ---\n');

    const filePath = path.join(outputDir, 'empty_sheet.xlsx');

    const writer = new XlsxWriter(filePath);
    writer.addSheet('Empty');
    writer.writeSheet([], ['Header1', 'Header2'], false);
    await writer.finalize();

    assert(fs.existsSync(filePath), 'Empty sheet file created');

    const reader = new XlsxReader();
    await reader.open(filePath);

    const rows: any[][] = [];
    while (await reader.read()) {
        const row = [reader.getValue(0), reader.getValue(1)];
        rows.push(row);
    }

    assertEqual(rows.length, 1, 'Empty sheet has only header');

    await reader.close();
}

// ==================== LARGE DATA TESTS ====================

async function testLargeDataset(): Promise<void> {
    console.log('\n--- Large Dataset Test ---\n');

    const filePath = path.join(outputDir, 'large.xlsx');
    const rowCount = 10000;
    const colCount = 10;

    const data: any[][] = [];
    for (let r = 0; r < rowCount; r++) {
        const row: any[] = [];
        for (let c = 0; c < colCount; c++) {
            row.push(`R${r}C${c}`);
        }
        data.push(row);
    }

    const headers = Array.from({ length: colCount }, (_, i) => `Col${i}`);

    const writer = new XlsxWriter(filePath);
    writer.addSheet('Large');
    writer.writeSheet(data, headers, false);
    await writer.finalize();

    assert(fs.existsSync(filePath), 'Large file created');

    const stats = fs.statSync(filePath);
    assert(stats.size > 0, 'Large file has content');

    const reader = new XlsxReader();
    await reader.open(filePath);

    const rows: any[][] = [];
    while (await reader.read()) {
        const row: any[] = [];
        for (let i = 0; i < reader.fieldCount; i++) {
            row.push(reader.getValue(i));
        }
        rows.push(row);
    }

    assertEqual(rows.length, rowCount + 1, 'Large dataset row count');

    await reader.close();
}

async function testManyColumns(): Promise<void> {
    console.log('\n--- Many Columns Test ---\n');

    const filePath = path.join(outputDir, 'many_columns.xlsx');
    const colCount = 100;

    const headers = Array.from({ length: colCount }, (_, i) => `Col${i}`);
    const data = [Array.from({ length: colCount }, (_, i) => i)];

    const writer = new XlsxWriter(filePath);
    writer.addSheet('ManyCols');
    writer.writeSheet(data, headers, false);
    await writer.finalize();

    assert(fs.existsSync(filePath), 'Many columns file created');

    const reader = new XlsxReader();
    await reader.open(filePath);

    const rows: any[][] = [];
    while (await reader.read()) {
        const row: any[] = [];
        for (let i = 0; i < reader.fieldCount; i++) {
            row.push(reader.getValue(i));
        }
        rows.push(row);
    }

    assertEqual(rows[0].length, colCount, 'Column count matches');

    await reader.close();
}

// ==================== SPECIAL CASES ====================

async function testSharedStringsDeduplication(): Promise<void> {
    console.log('\n--- Shared Strings Deduplication Test ---\n');

    const filePath = path.join(outputDir, 'shared_strings.xlsx');
    const repeatedString = 'RepeatedValue';

    const data = [
        [repeatedString, 1],
        [repeatedString, 2],
        [repeatedString, 3],
        [repeatedString, 4],
        [repeatedString, 5],
    ];

    const writer = new XlsxWriter(filePath);
    writer.addSheet('SharedStrings');
    writer.writeSheet(data, ['Text', 'Id'], false);
    await writer.finalize();

    assert(fs.existsSync(filePath), 'Shared strings file created');

    const reader = new XlsxReader();
    await reader.open(filePath);

    const rows: any[][] = [];
    while (await reader.read()) {
        const row = [reader.getValue(0), reader.getValue(1)];
        rows.push(row);
    }

    const dataRows = rows.slice(1);
    for (let i = 0; i < dataRows.length; i++) {
        assertEqual(dataRows[i][0], repeatedString, `Shared string ${i} matches`);
    }

    await reader.close();
}

async function testMixedTypesInColumn(): Promise<void> {
    console.log('\n--- Mixed Types in Column Test ---\n');

    const filePath = path.join(outputDir, 'mixed_types.xlsx');
    const data = [
        [1, 'number'],
        ['text', 'string'],
        [true, 'boolean'],
        [null, 'null'],
        [1.5, 'float'],
        [new Date('2024-01-01'), 'date'],
    ];

    const writer = new XlsxWriter(filePath);
    writer.addSheet('Mixed');
    writer.writeSheet(data, ['Value', 'Type'], false);
    await writer.finalize();

    assert(fs.existsSync(filePath), 'Mixed types file created');

    const reader = new XlsxReader();
    await reader.open(filePath);

    const rows: any[][] = [];
    while (await reader.read()) {
        const row = [reader.getValue(0), reader.getValue(1)];
        rows.push(row);
    }

    const dataRows = rows.slice(1);
    assertEqual(dataRows.length, data.length, 'Mixed types row count');

    await reader.close();
}

async function testXmlEscaping(): Promise<void> {
    console.log('\n--- XML Escaping Test ---\n');

    const filePath = path.join(outputDir, 'xml_escape.xlsx');
    // Note: XML entities are used internally by XLSX format, but the actual stored/read values
    // are the decoded characters. We test that special characters are preserved correctly.
    const data = [
        ['<script>alert("xss")</script>', 'HTML-like tags'],
        ['&', 'Ampersand'],
        ['<', 'Less than'],
        ['>', 'Greater than'],
        ['"', 'Double quote'],
        ["'", 'Single quote'],
        ['A & B < C > D', 'Combined special chars'],
    ];

    const writer = new XlsxWriter(filePath);
    writer.addSheet('XmlEscape');
    writer.writeSheet(data, ['Value', 'Description'], false);
    await writer.finalize();

    assert(fs.existsSync(filePath), 'XML escape file created');

    const reader = new XlsxReader();
    await reader.open(filePath);

    const rows: any[][] = [];
    while (await reader.read()) {
        const row = [reader.getValue(0), reader.getValue(1)];
        rows.push(row);
    }

    const dataRows = rows.slice(1);
    for (let i = 0; i < data.length; i++) {
        assertEqual(dataRows[i][0], data[i][0], `XML escape ${i}: ${data[i][1]}`);
    }

    await reader.close();
}

async function runXlsxWriterTests(): Promise<void> {
    console.log('='.repeat(60));
    console.log('XLSX Writer Integration Tests');
    console.log('='.repeat(60));

    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }

    try {
        // Data types
        await testIntegerValues();
        await testFloatValues();
        await testStringValues();
        await testBooleanValues();
        await testDateValues();

        // Sheets
        await testMultipleSheets();
        await testSheetNameSanitization();
        await testEmptySheet();

        // Large data
        await testLargeDataset();
        await testManyColumns();

        // Special cases
        await testSharedStringsDeduplication();
        await testMixedTypesInColumn();
        await testXmlEscaping();
    } catch (err) {
        console.error('Test error:', err);
    }

    console.log('\n' + '='.repeat(60));
    const passed = results.filter(r => r.passed).length;
    const failed = results.filter(r => !r.passed).length;
    console.log(`XLSX Writer Tests: ${passed} passed, ${failed} failed`);
    console.log('='.repeat(60));

    if (failed > 0) {
        process.exit(1);
    }
}

runXlsxWriterTests().catch(console.error);
