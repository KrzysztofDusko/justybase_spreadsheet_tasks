/**
 * Smoke Tests - Quick verification that basic functionality works
 * These tests run fast and verify core write/read operations
 */

import { XlsbWriter } from '../src/XlsbWriter';
import { XlsxWriter } from '../src/XlsxWriter';
import { XlsbReader } from '../src/XlsbReader';
import { XlsxReader } from '../src/XlsxReader';
import * as fs from 'fs';
import * as path from 'path';

const outputDir = path.join(__dirname, '..', 'test-output', 'smoke');

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

// Test data
const testHeaders = ['ID', 'Name', 'Value'];
const testData = [
    [1, 'Test One', 100.5],
    [2, 'Test Two', 200.75],
];

async function testXlsbWriteRead(): Promise<void> {
    console.log('\n--- XLSB Smoke Test ---\n');

    const filePath = path.join(outputDir, 'smoke_xlsb.xlsb');

    const writer = new XlsbWriter(filePath);
    writer.addSheet('TestSheet');
    writer.writeSheet(testData, testHeaders, false);
    await writer.finalize();

    assert(fs.existsSync(filePath), 'XLSB file created');

    const reader = new XlsbReader();
    await reader.open(filePath);
    assertEqual(reader.getSheetNames(), ['TestSheet'], 'XLSB sheet name');

    const rows: any[][] = [];
    while (await reader.read()) {
        const row: any[] = [];
        for (let i = 0; i < reader.fieldCount; i++) {
            row.push(reader.getValue(i));
        }
        rows.push(row);
    }

    assertEqual(rows.length, 3, 'XLSB row count (header + data)');
}

async function testXlsxWriteRead(): Promise<void> {
    console.log('\n--- XLSX Smoke Test ---\n');

    const filePath = path.join(outputDir, 'smoke_xlsx.xlsx');

    const writer = new XlsxWriter(filePath);
    writer.addSheet('TestSheet');
    writer.writeSheet(testData, testHeaders, false);
    await writer.finalize();

    assert(fs.existsSync(filePath), 'XLSX file created');

    const reader = new XlsxReader();
    await reader.open(filePath);
    assertEqual(reader.getSheetNames(), ['TestSheet'], 'XLSX sheet name');

    const rows: any[][] = [];
    while (await reader.read()) {
        const row: any[] = [];
        for (let i = 0; i < reader.fieldCount; i++) {
            row.push(reader.getValue(i));
        }
        rows.push(row);
    }

    assertEqual(rows.length, 3, 'XLSX row count (header + data)');
    await reader.close();
}

async function runSmokeTests(): Promise<void> {
    console.log('='.repeat(50));
    console.log('Smoke Tests');
    console.log('='.repeat(50));

    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }

    try {
        await testXlsbWriteRead();
        await testXlsxWriteRead();
    } catch (err) {
        console.error('Smoke test error:', err);
    }

    console.log('\n' + '='.repeat(50));
    const passed = results.filter(r => r.passed).length;
    const failed = results.filter(r => !r.passed).length;
    console.log(`Smoke Tests: ${passed} passed, ${failed} failed`);
    console.log('='.repeat(50));

    if (failed > 0) {
        process.exit(1);
    }
}

runSmokeTests().catch(console.error);
