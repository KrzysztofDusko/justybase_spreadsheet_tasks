/**
 * Basic Write Example
 * 
 * This example demonstrates how to write data to XLSB and XLSX files
 * using spreadsheet-tasks.
 * 
 * Run with: npx ts-node examples/basic-write.ts
 */

import { XlsbWriter, XlsxWriter } from '../dist';
import * as path from 'path';
import * as fs from 'fs';

// Ensure output directory exists
const outputDir = path.join(__dirname, 'output');
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
}

async function writeXlsbExample() {
    console.log('Writing XLSB file...');

    const filePath = path.join(outputDir, 'basic.xlsb');
    const writer = new XlsbWriter(filePath);

    // Add a sheet
    writer.addSheet('Employees');

    // Write data with headers
    const data = [
        ['Name', 'Department', 'Salary', 'Start Date', 'Active'],
        ['Alice Johnson', 'Engineering', 85000, new Date('2020-03-15'), true],
        ['Bob Smith', 'Marketing', 65000, new Date('2019-07-01'), true],
        ['Charlie Brown', 'Sales', 72000, new Date('2021-01-10'), false],
        ['Diana Ross', 'Engineering', 95000, new Date('2018-11-20'), true],
        ['Eve Williams', 'HR', 55000, new Date('2022-05-05'), true],
    ];

    writer.writeSheet(data);

    // Finalize the file
    await writer.finalize();

    console.log(`XLSB file created: ${filePath}`);
}

async function writeXlsxExample() {
    console.log('Writing XLSX file...');

    const filePath = path.join(outputDir, 'basic.xlsx');
    const writer = new XlsxWriter(filePath);

    // Add a sheet
    writer.addSheet('Products');

    // Write data - using separate headers parameter
    const headers = ['Product ID', 'Name', 'Price', 'Quantity', 'In Stock'];
    const data = [
        ['P001', 'Widget A', 19.99, 150, true],
        ['P002', 'Widget B', 29.99, 75, true],
        ['P003', 'Gadget X', 49.99, 0, false],
        ['P004', 'Gadget Y', 39.99, 200, true],
        ['P005', 'Tool Z', 9.99, 500, true],
    ];

    writer.writeSheet(data, headers);

    // Finalize the file
    await writer.finalize();

    console.log(`XLSX file created: ${filePath}`);
}

async function main() {
    try {
        await writeXlsbExample();
        await writeXlsxExample();
        console.log('\nAll files created successfully!');
    } catch (error) {
        console.error('Error:', error);
    }
}

main();
