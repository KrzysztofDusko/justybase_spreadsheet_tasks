/**
 * Update Existing Workbook Example
 *
 * This example demonstrates how to replace the data of a worksheet inside an
 * existing Excel file (e.g. a report whose pivot tables are wired to a "data1"
 * sheet) without rebuilding the workbook. The same API works for XLSX and XLSB.
 *
 * Run with: npx ts-node examples/update-existing.ts
 */

import { XlsxUpdater, XlsbUpdater } from '../dist/cjs';
import * as path from 'path';
import * as fs from 'fs';

const outputDir = path.join(__dirname, 'output');
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
}

async function updateExistingWorkbook() {
    // New data for the "data1" sheet (same columns as before)
    const rows = [
        [new Date('2026-01-01'), 'PL', 'Warsaw', 1200.5, true],
        [new Date('2026-01-02'), 'DE', 'Berlin', 980.25, true],
        [new Date('2026-01-03'), 'PL', 'Krakow', 1530.0, false],
    ];
    const headers = ['DATEKEY', 'COUNTRY', 'CITY', 'AMOUNT', 'ACTIVE'];

    for (const ext of ['xlsx', 'xlsb']) {
        const input = path.join(__dirname, '..', `report.${ext}`);
        const output = path.join(outputDir, `report_updated.${ext}`);

        if (!fs.existsSync(input)) {
            console.log(`Template workbook not found: ${input} (skipping)`);
            continue;
        }

        // Open the existing workbook (everything outside the replaced sheet
        // — pivot tables, other sheets, styles — is preserved)
        const updater = ext === 'xlsb' ? new XlsbUpdater(input) : new XlsxUpdater(input);
        console.log(`[${ext}] Sheets:`, updater.getSheetNames());

        updater.replaceSheetData('data1', rows, { headers });

        // Write the updated workbook to a new file (or omit the path to
        // overwrite the source in place)
        updater.save(output);
        console.log(`[${ext}] Workbook updated: ${output}`);
    }
}

updateExistingWorkbook();

