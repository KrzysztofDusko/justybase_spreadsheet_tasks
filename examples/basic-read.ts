/**
 * Basic Read Example
 * 
 * This example demonstrates how to read data from XLSB and XLSX files
 * using spreadsheet-tasks.
 * 
 * Run with: npx ts-node examples/basic-read.ts
 */

import { XlsbReader, XlsxReader, ReaderFactory } from '../dist';
import * as path from 'path';

async function readXlsbExample() {
    console.log('=== Reading XLSB file ===\n');

    const filePath = path.join(__dirname, 'output', 'basic.xlsb');

    const reader = new XlsbReader();
    await reader.open(filePath);

    // Get sheet names
    const sheetNames = reader.getSheetNames();
    console.log('Sheet names:', sheetNames);
    console.log('Number of sheets:', reader.resultsCount);
    console.log('');

    // Read all rows
    let rowCount = 0;
    while (reader.read()) {
        const row: any[] = [];
        for (let i = 0; i < reader.fieldCount; i++) {
            row.push(reader.getValue(i));
        }

        if (rowCount === 0) {
            console.log('Headers:', row);
        } else {
            console.log(`Row ${rowCount}:`, row);
        }
        rowCount++;
    }

    console.log(`\nTotal rows read: ${rowCount}\n`);
}

async function readXlsxExample() {
    console.log('=== Reading XLSX file ===\n');

    const filePath = path.join(__dirname, 'output', 'basic.xlsx');

    const reader = new XlsxReader();
    await reader.open(filePath);

    // Get sheet names
    const sheetNames = reader.getSheetNames();
    console.log('Sheet names:', sheetNames);
    console.log('');

    // Read all rows
    let rowCount = 0;
    while (await reader.read()) {
        const row: any[] = [];
        for (let i = 0; i < reader.fieldCount; i++) {
            row.push(reader.getValue(i));
        }

        if (rowCount === 0) {
            console.log('Headers:', row);
        } else {
            console.log(`Row ${rowCount}:`, row);
        }
        rowCount++;
    }

    await reader.close();
    console.log(`\nTotal rows read: ${rowCount}\n`);
}

async function useReaderFactory() {
    console.log('=== Using ReaderFactory ===\n');

    // ReaderFactory automatically selects the right reader based on file extension
    const xlsbPath = path.join(__dirname, 'output', 'basic.xlsb');
    const xlsxPath = path.join(__dirname, 'output', 'basic.xlsx');

    // Read XLSB using factory
    const xlsbReader = ReaderFactory.create(xlsbPath);
    console.log('Reader for .xlsb:', xlsbReader.constructor.name);

    // Read XLSX using factory
    const xlsxReader = ReaderFactory.create(xlsxPath);
    console.log('Reader for .xlsx:', xlsxReader.constructor.name);
}

async function main() {
    try {
        // Make sure to run basic-write.ts first to create the files
        await readXlsbExample();
        await readXlsxExample();
        useReaderFactory();

        console.log('\nAll files read successfully!');
    } catch (error) {
        console.error('Error:', error);
        console.log('\nNote: Make sure to run basic-write.ts first to create the sample files.');
    }
}

main();
