/**
 * Datetime Column Widths Example
 * 
 * Demonstrates auto-sizing columns for datetime data in both batch and streaming modes.
 * Columns with datetime values are now correctly sized to show full datetime strings
 * without Excel truncating to ########.
 * 
 * Run with: npx ts-node examples/datetime-widths.ts
 */

import { XlsbWriter, XlsxWriter } from '../dist/cjs';
import * as path from 'path';
import * as fs from 'fs';

const outputDir = path.join(__dirname, 'output');
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
}

async function batchModeExample() {
    console.log('=== Batch Mode: DateTime Column Widths ===\n');

    const data = [
        [1, new Date('2024-06-15T10:30:00'), 'Meeting with team'],
        [2, new Date('2024-01-01'), 'New Year'],
        [3, new Date('2023-12-25T14:00:00'), 'Christmas'],
        [4, new Date('2024-03-15T08:45:30'), 'Project deadline'],
        [5, new Date('2024-07-04'), 'Independence Day'],
    ];

    const headers = ['ID', 'DateTime', 'Event'];

    // XLSB batch
    const xlsbPath = path.join(outputDir, 'datetime-batch.xlsb');
    const xlsbWriter = new XlsbWriter(xlsbPath);
    xlsbWriter.addSheet('Events');
    xlsbWriter.writeSheet(data, headers);
    await xlsbWriter.finalize();
    console.log(`XLSB batch: ${xlsbPath}`);

    // XLSX batch
    const xlsxPath = path.join(outputDir, 'datetime-batch.xlsx');
    const xlsxWriter = new XlsxWriter(xlsxPath);
    xlsxWriter.addSheet('Events');
    xlsxWriter.writeSheet(data, headers);
    await xlsxWriter.finalize();
    console.log(`XLSX batch: ${xlsxPath}\n`);
}

async function streamingModeExample() {
    console.log('=== Streaming Mode: DateTime Column Widths ===\n');

    // Generate sample rows for width calculation
    const sampleRows = [
        [new Date('2024-06-15T10:30:00'), 'Meeting with team'],
        [new Date('2024-01-01'), 'New Year'],
    ];

    // XLSB streaming
    const xlsbPath = path.join(outputDir, 'datetime-streaming.xlsb');
    const xlsbWriter = new XlsbWriter(xlsbPath);
    xlsbWriter.startSheet('Events', 2, ['DateTime', 'Event'], {
        sampleRows: sampleRows
    });

    for (let i = 0; i < 1000; i++) {
        xlsbWriter.writeRow([
            new Date(2024, i % 12, (i % 28) + 1, i % 24, i % 60),
            `Event ${i + 1}`
        ]);
    }
    xlsbWriter.endSheet();
    await xlsbWriter.finalize();
    console.log(`XLSB streaming: ${xlsbPath}`);

    // XLSX streaming
    const xlsxPath = path.join(outputDir, 'datetime-streaming.xlsx');
    const xlsxWriter = new XlsxWriter(xlsxPath);
    xlsxWriter.startSheet('Events', 2, ['DateTime', 'Event'], {
        sampleRows: sampleRows
    });

    for (let i = 0; i < 1000; i++) {
        xlsxWriter.writeRow([
            new Date(2024, i % 12, (i % 28) + 1, i % 24, i % 60),
            `Event ${i + 1}`
        ]);
    }
    xlsxWriter.endSheet();
    await xlsxWriter.finalize();
    console.log(`XLSX streaming: ${xlsxPath}\n`);
}

async function main() {
    try {
        await batchModeExample();
        await streamingModeExample();
        console.log('=== All datetime width examples completed ===');
        console.log('\nOpen the generated files in Excel to verify:');
        console.log('  - DateTime columns show full values (not ########)');
        console.log('  - Column widths auto-fit to accommodate datetime text');
    } catch (error) {
        console.error('Error:', error);
    }
}

main();
