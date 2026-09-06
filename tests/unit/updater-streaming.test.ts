import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { XlsbReader } from '../../src/XlsbReader';
import { XlsbUpdater } from '../../src/XlsbUpdater';
import { XlsbWriter } from '../../src/XlsbWriter';
import { XlsxUpdater, XlsmUpdater } from '../../src/XlsxUpdater';
import { parseSharedStringsBin } from '../../src/biff12Utils';
import { writeBufferAtomically } from '../../src/atomicFile';

const outputDir = fs.mkdtempSync(path.join(os.tmpdir(), 'spreadsheet-updater-tests-'));

function assert(condition: boolean, message: string): void {
    if (!condition) throw new Error(message);
}

function addText(zip: AdmZip, name: string, text: string): void {
    zip.addFile(name, Buffer.from(text, 'utf8'));
}

function createMinimalXlsx(filePath: string, withMacro = false): void {
    const zip = new AdmZip();
    addText(zip, '[Content_Types].xml', '<?xml version="1.0"?><Types/>');
    addText(zip, 'xl/workbook.xml',
        '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="data1" sheetId="1" r:id="rId1"/></sheets></workbook>');
    addText(zip, 'xl/_rels/workbook.xml.rels',
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
    addText(zip, 'xl/worksheets/sheet1.xml',
        '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B1"/><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row></sheetData></worksheet>');
    addText(zip, 'xl/sharedStrings.xml',
        '<?xml version="1.0"?><x:sst xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><x:si><x:t>old-a</x:t></x:si><x:si><x:t>old-b</x:t></x:si></x:sst>');
    if (withMacro) zip.addFile('xl/vbaProject.bin', Buffer.from('macro-payload'));
    zip.writeZip(filePath);
}

function zipText(filePath: string, name: string): string {
    const zip = new AdmZip(filePath);
    const entry = zip.getEntry(name);
    assert(!!entry, `missing ZIP member ${name}`);
    return entry!.getData().toString('utf8');
}

async function testXlsxStreamingAndRollback(): Promise<void> {
    const source = path.join(outputDir, 'source.xlsx');
    const target = path.join(outputDir, 'streamed.xlsx');
    createMinimalXlsx(source);

    const updater = new XlsxUpdater(source);
    await updater.replaceSheetDataStream('data1', (async function* () {
        yield ['new-a'];
        yield ['new-a'];
        yield [null, 'new-b'];
    })(), { headers: ['header'] });
    await updater.saveStreaming(target);

    const sheet = zipText(target, 'xl/worksheets/sheet1.xml');
    assert(sheet.includes('<dimension ref="A1:B4"/>'), 'streaming XLSX updates dimension');
    const sst = zipText(target, 'xl/sharedStrings.xml');
    assert(sst.includes('count="4"'), 'streaming XLSX subtracts replaced string references');
    assert(sst.includes('uniqueCount="5"'), 'streaming XLSX appends each new unique string once');
    assert(sst.includes('<x:t>new-a</x:t>') && sst.includes('<x:t>new-b</x:t>'), 'streaming XLSX writes new strings');

    const repeat = new XlsxUpdater(target);
    await repeat.replaceSheetDataStream('data1', [['replacement']]);
    const repeated = path.join(outputDir, 'repeated.xlsx');
    await repeat.saveStreaming(repeated);
    const repeatedSst = zipText(repeated, 'xl/sharedStrings.xml');
    assert(repeatedSst.includes('count="1"'), 'repeated replacement recalculates shared string count');
    assert(repeatedSst.includes('uniqueCount="6"'), 'repeated replacement does not duplicate pending strings');

    const failed = new XlsxUpdater(source);
    let rejected = false;
    try {
        await failed.replaceSheetDataStream('data1', (async function* () {
            yield ['temporary'];
            throw new Error('source failed');
        })());
    } catch (error) {
        rejected = String(error).includes('source failed');
    }
    assert(rejected, 'streaming XLSX propagates source errors');
    await failed.replaceSheetDataStream('data1', [['recovered']]);
    const recovered = path.join(outputDir, 'recovered.xlsx');
    failed.save(recovered);
    const recoveredSst = zipText(recovered, 'xl/sharedStrings.xml');
    assert(!recoveredSst.includes('temporary'), 'failed streaming XLSX rolls back temporary strings');
    assert(recoveredSst.includes('recovered'), 'updater remains reusable after a failed stream');
    assert(zipText(recovered, 'xl/worksheets/sheet1.xml').includes('<dimension ref="A1:A1"/>'), 'save materialises staged worksheet data');
}

async function testXlsmAlias(): Promise<void> {
    const source = path.join(outputDir, 'macro.xlsm');
    const target = path.join(outputDir, 'macro-updated.xlsm');
    createMinimalXlsx(source, true);
    const updater = new XlsmUpdater(source);
    await updater.replaceSheetDataStream('data1', [['macro-safe']]);
    await updater.saveStreaming(target);
    const zip = new AdmZip(target);
    assert(zip.getEntry('xl/vbaProject.bin')?.getData().toString() === 'macro-payload', 'XLSM alias preserves VBA parts');
}

async function testXlsbStreaming(): Promise<void> {
    const source = path.join(outputDir, 'source.xlsb');
    const target = path.join(outputDir, 'streamed.xlsb');
    const writer = new XlsbWriter(source);
    writer.addSheet('data1');
    writer.writeSheet([[1, 'old'], [2, 'value']], ['ID', 'TEXT'], false);
    await writer.finalize();

    const updater = new XlsbUpdater(source);
    await updater.replaceSheetDataStream('data1', (async function* () {
        yield [3, 'new'];
        yield [4, 'new'];
    })(), { headers: ['ID', 'TEXT'] });
    await updater.saveStreaming(target);

    const reader = new XlsbReader();
    await reader.open(target);
    const rows: unknown[][] = [];
    while (await reader.read()) rows.push([reader.getValue(0), reader.getValue(1)]);
    assert(JSON.stringify(rows) === JSON.stringify([
        ['ID', 'TEXT'],
        [3, 'new'],
        [4, 'new'],
    ]), 'streaming XLSB preserves rows and shared strings');
    const zip = new AdmZip(target);
    const sst = parseSharedStringsBin(zip.getEntry('xl/sharedStrings.bin')!.getData());
    assert(sst.total === 4, 'streaming XLSB recalculates shared string count');
}

function testAtomicBufferSave(): void {
    const target = path.join(outputDir, 'atomic.bin');
    fs.writeFileSync(target, 'old');
    writeBufferAtomically(Buffer.from('new'), target);
    assert(fs.readFileSync(target, 'utf8') === 'new', 'buffer save replaces destination atomically');
    assert(fs.readdirSync(outputDir).every(name => !name.endsWith('.tmp')), 'atomic save cleans temporary files');
}

async function main(): Promise<void> {
    testAtomicBufferSave();
    await testXlsxStreamingAndRollback();
    await testXlsmAlias();
    await testXlsbStreaming();
    console.log('Updater streaming regression tests passed');
}

main().catch(error => {
    console.error(error);
    process.exitCode = 1;
});
