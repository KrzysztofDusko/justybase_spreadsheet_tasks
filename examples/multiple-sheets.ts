/**
 * Multiple Sheets Example
 * 
 * This example demonstrates how to work with multiple worksheets
 * in a single workbook.
 * 
 * Run with: npx ts-node examples/multiple-sheets.ts
 */

import { XlsbWriter, XlsxWriter } from '../dist/cjs';
import * as path from 'path';
import * as fs from 'fs';

// Ensure output directory exists
const outputDir = path.join(__dirname, 'output');
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
}

async function createMultiSheetXlsb() {
    console.log('Creating multi-sheet XLSB file...');

    const filePath = path.join(outputDir, 'multi-sheet.xlsb');
    const writer = new XlsbWriter(filePath);

    // Sheet 1: Sales Summary
    writer.addSheet('Sales Summary');
    writer.writeSheet([
        ['Region', 'Q1', 'Q2', 'Q3', 'Q4', 'Total'],
        ['North', 125000, 142000, 138000, 165000, 570000],
        ['South', 98000, 105000, 112000, 128000, 443000],
        ['East', 156000, 167000, 178000, 195000, 696000],
        ['West', 189000, 198000, 205000, 225000, 817000],
    ]);

    // Sheet 2: Monthly Details
    writer.addSheet('Monthly Details');
    writer.writeSheet([
        ['Month', 'Revenue', 'Expenses', 'Profit'],
        ['January', 45000, 32000, 13000],
        ['February', 52000, 35000, 17000],
        ['March', 48000, 33000, 15000],
        ['April', 55000, 36000, 19000],
        ['May', 61000, 38000, 23000],
        ['June', 58000, 37000, 21000],
    ]);

    // Sheet 3: Hidden configuration sheet
    writer.addSheet('Config', true); // hidden = true
    writer.writeSheet([
        ['Setting', 'Value'],
        ['Version', '1.0.0'],
        ['LastUpdated', new Date()],
        ['AutoRefresh', true],
    ]);

    // Sheet 4: Products
    writer.addSheet('Products');
    writer.writeSheet([
        ['SKU', 'Product Name', 'Category', 'Price', 'Stock'],
        ['SKU-001', 'Laptop Pro 15', 'Electronics', 1299.99, 45],
        ['SKU-002', 'Wireless Mouse', 'Accessories', 29.99, 500],
        ['SKU-003', 'USB-C Hub', 'Accessories', 49.99, 200],
        ['SKU-004', 'Monitor 27"', 'Electronics', 399.99, 75],
        ['SKU-005', 'Mechanical Keyboard', 'Accessories', 129.99, 150],
    ]);

    await writer.finalize();
    console.log(`Created: ${filePath}`);
    console.log('Sheets: Sales Summary, Monthly Details, Config (hidden), Products\n');
}

async function createMultiSheetXlsx() {
    console.log('Creating multi-sheet XLSX file...');

    const filePath = path.join(outputDir, 'multi-sheet.xlsx');
    const writer = new XlsxWriter(filePath);

    // Sheet 1: Employees
    writer.addSheet('Employees');
    writer.writeSheet([
        ['ID', 'Name', 'Department', 'Hire Date', 'Salary'],
        [1, 'John Doe', 'Engineering', new Date('2020-01-15'), 85000],
        [2, 'Jane Smith', 'Marketing', new Date('2019-06-01'), 72000],
        [3, 'Bob Johnson', 'Sales', new Date('2021-03-10'), 68000],
        [4, 'Alice Brown', 'HR', new Date('2018-09-22'), 62000],
        [5, 'Charlie Wilson', 'Engineering', new Date('2022-02-28'), 90000],
    ]);

    // Sheet 2: Departments
    writer.addSheet('Departments');
    writer.writeSheet([
        ['Department', 'Manager', 'Budget', 'Headcount'],
        ['Engineering', 'Sarah Connor', 2500000, 45],
        ['Marketing', 'Mike Ross', 1200000, 22],
        ['Sales', 'Lisa Chen', 1800000, 35],
        ['HR', 'Tom Hardy', 800000, 12],
        ['Finance', 'Emma Watson', 600000, 8],
    ]);

    // Sheet 3: Projects
    writer.addSheet('Projects');
    writer.writeSheet([
        ['Project Name', 'Status', 'Start Date', 'End Date', 'Budget'],
        ['Website Redesign', 'In Progress', new Date('2024-01-01'), new Date('2024-06-30'), 150000],
        ['Mobile App v2', 'Planning', new Date('2024-04-01'), new Date('2024-12-31'), 300000],
        ['CRM Integration', 'Completed', new Date('2023-06-01'), new Date('2024-01-31'), 200000],
        ['Data Migration', 'On Hold', new Date('2024-03-01'), new Date('2024-09-30'), 100000],
    ]);

    await writer.finalize();
    console.log(`Created: ${filePath}`);
    console.log('Sheets: Employees, Departments, Projects\n');
}

async function main() {
    try {
        await createMultiSheetXlsb();
        await createMultiSheetXlsx();
        console.log('All multi-sheet files created successfully!');
    } catch (error) {
        console.error('Error:', error);
    }
}

main();
