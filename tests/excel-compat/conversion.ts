import * as fs from 'fs';
import * as path from 'path';
import { spawnSync } from 'child_process';
import { XlsbWriter } from '../../src/XlsbWriter';
import { XlsxWriter } from '../../src/XlsxWriter';
import { F } from '../../src/Formats';
import { convertXlsbToXlsx, convertXlsxToXlsb } from '../../src/ExcelConverter';

const OUTPUT_DIR = path.join(__dirname, '..', '..', 'test-output', 'excel-conversion');
const VALIDATOR = path.join(__dirname, 'validate-conversion.ps1');

function cleanOutput(): void {
    if (fs.existsSync(OUTPUT_DIR)) {
        fs.rmSync(OUTPUT_DIR, { recursive: true, force: true });
    }
    fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

function excelPreflight(): { available: boolean; details: string } {
    if (process.platform !== 'win32') {
        return { available: false, details: 'Windows is required' };
    }

    const command = [
        "$ErrorActionPreference = 'Stop'",
        "$excel = New-Object -ComObject Excel.Application",
        "$version = $excel.Version",
        "$excel.Quit()",
        '[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null',
        'Write-Output "READY|$version"',
    ].join('; ');

    const result = spawnSync(
        'powershell.exe',
        ['-NoLogo', '-NoProfile', '-NonInteractive', '-STA', '-ExecutionPolicy', 'Bypass', '-Command', command],
        { encoding: 'utf8', timeout: 300000, windowsHide: true }
    );

    if (result.status !== 0) {
        const details = (result.stderr || result.stdout || 'COM unavailable').replace(/\s+/g, ' ').trim();
        return { available: false, details };
    }

    return { available: true, details: (result.stdout || '').trim() };
}

async function writeFixture(
    Writer: typeof XlsbWriter | typeof XlsxWriter,
    filePath: string
): Promise<void> {
    const writer = new Writer(filePath);
    writer.addSheet('Data');
    writer.writeSheet([
        [1, 'Ala', { value: 1234.5, format: F.CURRENCY_PLN }, { value: new Date(2024, 0, 2, 15, 30), format: F.DATETIME_ISO }, true, null],
        [2, 'Zażółć gęślą jaźń 🚀', { value: -42.25, format: F.TWO_DECIMALS }, { value: new Date(2024, 5, 30), format: F.DATE_ISO }, false, 'tekst'],
        [3, '日本語', 0.125, new Date(2025, 11, 31), true, undefined],
    ], ['ID', 'Name', 'Amount', 'Date', 'Active', 'Optional'], false);

    writer.addSheet('Second Sheet');
    writer.writeSheet([
        ['Warszawa', 52.2297, 21.0122],
        ['Łódź', 51.7592, 19.4560],
    ], ['City', 'Latitude', 'Longitude'], false);

    await writer.finalize();
}

async function main(): Promise<void> {
    cleanOutput();

    const preflight = excelPreflight();
    if (!preflight.available) {
        console.log(`SKIP|Excel COM unavailable|${preflight.details}`);
        return;
    }

    console.log(`Excel COM preflight: ${preflight.details}`);

    const sourceXlsb = path.join(OUTPUT_DIR, 'source.xlsb');
    const sourceXlsx = path.join(OUTPUT_DIR, 'source.xlsx');
    const convertedXlsx = path.join(OUTPUT_DIR, 'converted-from-xlsb.xlsx');
    const convertedXlsb = path.join(OUTPUT_DIR, 'converted-from-xlsx.xlsb');

    await writeFixture(XlsbWriter, sourceXlsb);
    await writeFixture(XlsxWriter, sourceXlsx);
    await convertXlsbToXlsx(sourceXlsb, convertedXlsx);
    await convertXlsxToXlsb(sourceXlsx, convertedXlsb);

    const validation = spawnSync(
        'powershell.exe',
        [
            '-NoLogo',
            '-NoProfile',
            '-NonInteractive',
            '-STA',
            '-ExecutionPolicy',
            'Bypass',
            '-File',
            VALIDATOR,
            '-TestDir',
            OUTPUT_DIR,
        ],
        { encoding: 'utf8', timeout: 300000, windowsHide: true }
    );

    if (validation.stdout) process.stdout.write(validation.stdout);
    if (validation.stderr) process.stderr.write(validation.stderr);
    if (validation.status !== 0) {
        throw new Error(`Excel conversion validation failed with exit code ${validation.status ?? 'unknown'}.`);
    }

    console.log('Excel COM conversion validation passed.');
}

main().catch((error: unknown) => {
    console.error(error);
    process.exit(1);
});
