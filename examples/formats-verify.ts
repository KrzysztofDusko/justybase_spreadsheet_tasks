import { XlsbReader } from '../src/XlsbReader';
import { XlsxReader } from '../src/XlsxReader';
import * as fs from 'fs';
import * as path from 'path';

const OUTPUT_DIR = './test-output';

async function verifyFile(filePath: string, label: string) {
    console.log(`\n--- Verifying: ${label} ---`);
    const ext = path.extname(filePath).toLowerCase();

    let reader: XlsbReader | XlsxReader;
    if (ext === '.xlsb') {
        reader = new XlsbReader();
    } else {
        reader = new XlsxReader();
    }

    await reader.open(filePath);

    const sheetNames = reader.getSheetNames();
    console.log(`  Sheets: ${sheetNames.join(', ')}`);
    console.log(`  Rows: ${reader.rowCount}, Cols: ${reader.fieldCount}`);

    const allRows: any[][] = [];
    const isAsync = ext !== '.xlsb';
    while (isAsync ? await (reader as XlsxReader).read() : (reader as XlsbReader).read()) {
        const row: any[] = [];
        for (let i = 0; i < reader.fieldCount; i++) {
            row.push(reader.getValue(i));
        }
        allRows.push(row);
    }

    const printCount = Math.min(allRows.length, 15);
    for (let r = 0; r < printCount; r++) {
        const row = allRows[r];
        const vals = row.map((v: any) => {
            if (v === null || v === undefined) return '';
            if (v instanceof Date) return v.toISOString();
            if (typeof v === 'number') return v.toString();
            return String(v).substring(0, 40);
        });
        console.log(`    [${vals.join(', ')}]`);
    }
    if (allRows.length > printCount) {
        console.log(`    ... (${allRows.length - printCount} more rows)`);
    }

    reader.close();
    console.log(`  ✅ ${label} - ${allRows.length} rows read`);
}

async function main() {
    const files = [
        { file: `${OUTPUT_DIR}/formats_demo_xlsb.xlsb`, label: 'XLSB Format Demo' },
        { file: `${OUTPUT_DIR}/formats_demo_xlsx.xlsx`, label: 'XLSX Format Demo' },
        { file: `${OUTPUT_DIR}/formats_streaming_demo.xlsb`, label: 'XLSB Streaming Demo' },
        { file: `${OUTPUT_DIR}/sales_report_demo.xlsx`, label: 'Sales Report Demo' },
        { file: `${OUTPUT_DIR}/comprehensive_formats_demo.xlsb`, label: 'Comprehensive Demo' },
    ];

    for (const f of files) {
        if (fs.existsSync(f.file)) {
            await verifyFile(f.file, f.label);
        } else {
            console.log(`\n--- Skipping: ${f.label} (file not found) ---`);
        }
    }

    console.log('\n✅ All verifications complete!');
}

main().catch(console.error);
