/**
 * Large Dataset Example
 * 
 * This example demonstrates how to efficiently handle large datasets
 * using spreadsheet-tasks. It shows the performance benefits of XLSB format.
 * 
 * Run with: npx ts-node examples/large-dataset.ts
 */

import { XlsbWriter, XlsxWriter, XlsbReader, XlsxReader } from '../dist/cjs';
import * as path from 'path';
import * as fs from 'fs';

// Ensure output directory exists
const outputDir = path.join(__dirname, 'output');
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
}

// Configuration
const ROW_COUNT = 100000; // 100K rows

/**
 * Generate test data with various data types
 */
function generateTestData(rowCount: number): any[][] {
    const data: any[][] = [];

    // Headers
    data.push(['ID', 'Name', 'Email', 'Amount', 'Date', 'Active', 'Score']);

    const names = ['Alice', 'Bob', 'Charlie', 'Diana', 'Eve', 'Frank', 'Grace', 'Henry'];
    const domains = ['gmail.com', 'yahoo.com', 'outlook.com', 'company.com'];

    for (let i = 0; i < rowCount; i++) {
        const name = names[i % names.length];
        const domain = domains[i % domains.length];

        data.push([
            i + 1,                                          // ID (number)
            `${name} ${i}`,                                 // Name (string)
            `${name.toLowerCase()}${i}@${domain}`,          // Email (string)
            Math.round(Math.random() * 10000 * 100) / 100,  // Amount (decimal)
            new Date(2020, i % 12, (i % 28) + 1),          // Date
            i % 3 !== 0,                                    // Active (boolean)
            Math.floor(Math.random() * 100),                // Score (integer)
        ]);
    }

    return data;
}

/**
 * Write large dataset to XLSB file
 */
async function writeLargeXlsb(data: any[][]): Promise<{ time: number; size: number }> {
    const filePath = path.join(outputDir, 'large-dataset.xlsb');

    const startTime = performance.now();

    const writer = new XlsbWriter(filePath);
    writer.addSheet('Data');
    writer.writeSheet(data);
    await writer.finalize();

    const endTime = performance.now();
    const stats = fs.statSync(filePath);

    return {
        time: endTime - startTime,
        size: stats.size,
    };
}

/**
 * Write large dataset to XLSX file
 */
async function writeLargeXlsx(data: any[][]): Promise<{ time: number; size: number }> {
    const filePath = path.join(outputDir, 'large-dataset.xlsx');

    const startTime = performance.now();

    const writer = new XlsxWriter(filePath);
    writer.addSheet('Data');
    writer.writeSheet(data);
    await writer.finalize();

    const endTime = performance.now();
    const stats = fs.statSync(filePath);

    return {
        time: endTime - startTime,
        size: stats.size,
    };
}

/**
 * Read large XLSB file
 */
async function readLargeXlsb(): Promise<{ time: number; rowCount: number }> {
    const filePath = path.join(outputDir, 'large-dataset.xlsb');

    const startTime = performance.now();

    const reader = new XlsbReader();
    await reader.open(filePath);

    let rowCount = 0;
    while (await reader.read()) {
        rowCount++;
        // Access values to simulate real-world usage
        for (let i = 0; i < reader.fieldCount; i++) {
            const _ = reader.getValue(i);
        }
    }

    const endTime = performance.now();

    return {
        time: endTime - startTime,
        rowCount,
    };
}

/**
 * Read large XLSX file
 */
async function readLargeXlsx(): Promise<{ time: number; rowCount: number }> {
    const filePath = path.join(outputDir, 'large-dataset.xlsx');

    const startTime = performance.now();

    const reader = new XlsxReader();
    await reader.open(filePath);

    let rowCount = 0;
    while (await reader.read()) {
        rowCount++;
        // Access values to simulate real-world usage
        for (let i = 0; i < reader.fieldCount; i++) {
            const _ = reader.getValue(i);
        }
    }

    await reader.close();
    const endTime = performance.now();

    return {
        time: endTime - startTime,
        rowCount,
    };
}

function formatBytes(bytes: number): string {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(2)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

async function main() {
    console.log('='.repeat(60));
    console.log('Large Dataset Performance Test');
    console.log('='.repeat(60));
    console.log(`Row count: ${ROW_COUNT.toLocaleString()}`);
    console.log('');

    // Generate test data
    console.log('Generating test data...');
    const data = generateTestData(ROW_COUNT);
    console.log(`Data generated: ${data.length.toLocaleString()} rows\n`);

    // Write tests
    console.log('--- WRITE TESTS ---\n');

    console.log('Writing XLSB...');
    const xlsbWrite = await writeLargeXlsb(data);
    console.log(`  Time: ${xlsbWrite.time.toFixed(2)} ms`);
    console.log(`  Size: ${formatBytes(xlsbWrite.size)}\n`);

    console.log('Writing XLSX...');
    const xlsxWrite = await writeLargeXlsx(data);
    console.log(`  Time: ${xlsxWrite.time.toFixed(2)} ms`);
    console.log(`  Size: ${formatBytes(xlsxWrite.size)}\n`);

    // Read tests
    console.log('--- READ TESTS ---\n');

    console.log('Reading XLSB...');
    const xlsbRead = await readLargeXlsb();
    console.log(`  Time: ${xlsbRead.time.toFixed(2)} ms`);
    console.log(`  Rows: ${xlsbRead.rowCount.toLocaleString()}\n`);

    console.log('Reading XLSX...');
    const xlsxRead = await readLargeXlsx();
    console.log(`  Time: ${xlsxRead.time.toFixed(2)} ms`);
    console.log(`  Rows: ${xlsxRead.rowCount.toLocaleString()}\n`);

    // Summary
    console.log('='.repeat(60));
    console.log('SUMMARY');
    console.log('='.repeat(60));
    console.log('');
    console.log('Write Performance:');
    console.log(`  XLSB: ${xlsbWrite.time.toFixed(2)} ms (${formatBytes(xlsbWrite.size)})`);
    console.log(`  XLSX: ${xlsxWrite.time.toFixed(2)} ms (${formatBytes(xlsxWrite.size)})`);
    console.log(`  XLSB is ${(xlsxWrite.time / xlsbWrite.time).toFixed(2)}x faster`);
    console.log(`  XLSB is ${((1 - xlsbWrite.size / xlsxWrite.size) * 100).toFixed(0)}% smaller`);
    console.log('');
    console.log('Read Performance:');
    console.log(`  XLSB: ${xlsbRead.time.toFixed(2)} ms`);
    console.log(`  XLSX: ${xlsxRead.time.toFixed(2)} ms`);
    console.log(`  XLSB is ${(xlsxRead.time / xlsbRead.time).toFixed(2)}x faster`);
    console.log('');
}

main().catch(console.error);
