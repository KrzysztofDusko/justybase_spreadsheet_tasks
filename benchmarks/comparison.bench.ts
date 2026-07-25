/**
 * Benchmark comparing XlsbWriterNode with ExcelJS
 * Run with: npx ts-node benchmarks/comparison.bench.ts
 */
import { XlsbWriter } from '../dist/XlsbWriter';
import { XlsxWriter } from '../dist/XlsxWriter';
import { XlsbReader } from '../dist/XlsbReader';
import { XlsxReader } from '../dist/XlsxReader';
import ExcelJS from 'exceljs';
import * as fs from 'fs';
import * as path from 'path';

const ROWS = 50000;
const outFileXlsb = path.join(__dirname, '..', 'output', 'benchmark_output.xlsb');
const outFileXlsx = path.join(__dirname, '..', 'output', 'benchmark_output.xlsx');
const outFileOurXlsx = path.join(__dirname, '..', 'output', 'benchmark_output_our.xlsx');

const headers = ['ID', 'Name', 'Count', 'Score', 'Date', 'Active', 'Description'];

function getDataRow(i: number): any[] {
    return [
        i,
        `Produkt ${i} żółć`,
        Math.floor(Math.random() * 10000),
        Math.random() * 100,
        new Date(),
        i % 2 === 0,
        `Opis produktu ${i} z polskimi znakami: ąęśćńźółĄĘŚĆŃŹÓŁ oraz dłuższy tekst testowy.`
    ];
}

async function benchmark(): Promise<void> {
    console.log(`Generating data for ${ROWS} rows...`);
    const data: any[][] = [];
    for (let i = 0; i < ROWS; i++) {
        data.push(getDataRow(i));
    }

    // Ensure output directory exists
    const outputDir = path.dirname(outFileXlsb);
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }

    // --- XlsbWriter ---
    if (fs.existsSync(outFileXlsb)) fs.unlinkSync(outFileXlsb);
    if (typeof global.gc === 'function') global.gc();

    console.log(`\n--- XlsbWriter ---`);
    const startXlsb = performance.now();

    const xlsb = new XlsbWriter(outFileXlsb);
    xlsb.addSheet('Benchmark');
    xlsb.writeSheet(data, headers);
    await xlsb.finalize();

    const endXlsb = performance.now();
    const timeXlsb = (endXlsb - startXlsb).toFixed(2);
    const sizeXlsb = (fs.statSync(outFileXlsb).size / 1024 / 1024).toFixed(2);
    console.log(`Time: ${timeXlsb} ms`);
    console.log(`Size: ${sizeXlsb} MB`);

    // --- XlsxWriter (Our Implementation) ---
    if (fs.existsSync(outFileOurXlsx)) fs.unlinkSync(outFileOurXlsx);
    if (typeof global.gc === 'function') global.gc();

    console.log(`\n--- XlsxWriter (Our Implementation) ---`);
    const startOurXlsx = performance.now();

    const ourXlsx = new XlsxWriter(outFileOurXlsx);
    ourXlsx.addSheet('Benchmark');
    ourXlsx.writeSheet(data, headers);
    await ourXlsx.finalize();

    const endOurXlsx = performance.now();
    const timeOurXlsx = (endOurXlsx - startOurXlsx).toFixed(2);
    const sizeOurXlsx = (fs.statSync(outFileOurXlsx).size / 1024 / 1024).toFixed(2);
    console.log(`Time: ${timeOurXlsx} ms`);
    console.log(`Size: ${sizeOurXlsx} MB`);

    // --- ExcelJS ---
    if (fs.existsSync(outFileXlsx)) fs.unlinkSync(outFileXlsx);
    if (typeof global.gc === 'function') global.gc();

    console.log(`\n--- ExcelJS (Stream) ---`);
    const startExcel = performance.now();

    const options = {
        filename: outFileXlsx,
        useStyles: true,
        useSharedStrings: true
    };
    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);
    const worksheet = workbook.addWorksheet('Benchmark');
    worksheet.columns = [
        { header: 'ID', key: 'id' },
        { header: 'Name', key: 'name' },
        { header: 'Count', key: 'count' },
        { header: 'Score', key: 'score' },
        { header: 'Date', key: 'date' },
        { header: 'Active', key: 'active' },
        { header: 'Description', key: 'desc' }
    ];

    for (let i = 0; i < ROWS; i++) {
        worksheet.addRow(data[i]).commit();
    }
    worksheet.commit();
    await workbook.commit();

    const endExcel = performance.now();
    const timeExcel = (endExcel - startExcel).toFixed(2);
    const sizeExcel = (fs.statSync(outFileXlsx).size / 1024 / 1024).toFixed(2);

    console.log(`Time: ${timeExcel} ms`);
    console.log(`Size: ${sizeExcel} MB`);

    console.log(`\n--- Comparison ---`);
    console.log(`Rows: ${ROWS}`);
    console.log(`XlsbWriter: ${timeXlsb} ms, ${sizeXlsb} MB`);
    console.log(`XlsxWriter: ${timeOurXlsx} ms, ${sizeOurXlsx} MB`);
    console.log(`ExcelJS:    ${timeExcel} ms, ${sizeExcel} MB`);

    console.log(`\nRelative Speed (vs ExcelJS):`);
    console.log(`XlsbWriter: ${(parseFloat(timeExcel) / parseFloat(timeXlsb)).toFixed(2)}x faster`);
    console.log(`XlsxWriter: ${(parseFloat(timeExcel) / parseFloat(timeOurXlsx)).toFixed(2)}x faster`);

    console.log(`\nRelative Size (vs ExcelJS):`);
    console.log(`XlsbWriter: ${(parseFloat(sizeExcel) / parseFloat(sizeXlsb)).toFixed(2)}x smaller`);
    console.log(`XlsxWriter: ${(parseFloat(sizeExcel) / parseFloat(sizeOurXlsx)).toFixed(2)}x smaller`);

    // --- READ BENCHMARKS ---
    console.log(`\n--- Reading Benchmarks ---`);

    // 1. XlsbReader
    if (typeof global.gc === 'function') global.gc();
    console.log(`\n--- Reading XlsbReader ---`);
    const startReadXlsb = performance.now();
    let rowsReadXlsb = 0;

    const xlsbReader = new XlsbReader();
    await xlsbReader.open(outFileXlsb);
    while (await xlsbReader.read()) {
        rowsReadXlsb++;
        xlsbReader.getValue(1);
    }

    const endReadXlsb = performance.now();
    const timeReadXlsb = (endReadXlsb - startReadXlsb).toFixed(2);
    console.log(`Time: ${timeReadXlsb} ms`);
    console.log(`Rows: ${rowsReadXlsb}`);

    // 2. XlsxReader
    if (typeof global.gc === 'function') global.gc();
    console.log(`\n--- Reading XlsxReader ---`);
    const startReadXlsx = performance.now();
    let rowsReadXlsx = 0;

    const xlsxReader = new XlsxReader();
    await xlsxReader.open(outFileOurXlsx);
    while (await xlsxReader.read()) {
        rowsReadXlsx++;
        xlsxReader.getValue(1);
    }

    const endReadXlsx = performance.now();
    const timeReadXlsx = (endReadXlsx - startReadXlsx).toFixed(2);
    console.log(`Time: ${timeReadXlsx} ms`);
    console.log(`Rows: ${rowsReadXlsx}`);

    // 3. ExcelJS Read
    if (typeof global.gc === 'function') global.gc();
    console.log(`\n--- Reading ExcelJS (readFile) ---`);
    const startReadExcel = performance.now();
    let rowsReadExcel = 0;

    const workbookRead = new ExcelJS.Workbook();
    await workbookRead.xlsx.readFile(outFileOurXlsx);
    const sheet = workbookRead.getWorksheet(1);
    if (sheet) {
        sheet.eachRow(() => {
            rowsReadExcel++;
        });
    }

    const endReadExcel = performance.now();
    const timeReadExcel = (endReadExcel - startReadExcel).toFixed(2);
    console.log(`Time: ${timeReadExcel} ms`);
    console.log(`Rows: ${rowsReadExcel}`);

    console.log(`\n--- Read Comparison ---`);
    console.log(`XlsbReader: ${timeReadXlsb} ms`);
    console.log(`XlsxReader: ${timeReadXlsx} ms`);
    console.log(`ExcelJS:    ${timeReadExcel} ms`);
    console.log(`Relative Speed (vs ExcelJS):`);
    console.log(`XlsbReader: ${(parseFloat(timeReadExcel) / parseFloat(timeReadXlsb)).toFixed(2)}x faster`);
    console.log(`XlsxReader: ${(parseFloat(timeReadExcel) / parseFloat(timeReadXlsx)).toFixed(2)}x faster`);
}

benchmark().catch(console.error);
