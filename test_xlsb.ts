import { XlsbWriter, XlsxWriter, F } from './dist';
import * as path from 'path';

const dir = __dirname;

async function main() {
    // Very simple XLSB - no dates, no formats
    const xlsb = new XlsbWriter(path.join(dir, 'test_minimal.xlsb'));
    xlsb.addSheet('Test');
    xlsb.writeSheet([
        [1, 'Hello', 3.14],
        [2, 'World', 2.71],
    ], ['Number', 'Text', 'Value']);
    await xlsb.finalize();
    console.log('test_minimal.xlsb OK');

    // XLSB with Date
    const xlsb2 = new XlsbWriter(path.join(dir, 'test_date.xlsb'));
    xlsb2.addSheet('Dates');
    xlsb2.writeSheet([
        [new Date(2025, 5, 1), 'Today'],
        [new Date(2025, 0, 1), 'New Year'],
    ], ['Date', 'Name']);
    await xlsb2.finalize();
    console.log('test_date.xlsb OK');

    // XLSB with formats
    const xlsb3 = new XlsbWriter(path.join(dir, 'test_formats.xlsb'));
    xlsb3.addSheet('Formats');
    xlsb3.writeSheet([
        [{ value: 15000, format: F.CURRENCY_PLN }, { value: 0.15, format: F.PERCENTAGE }],
        [{ value: 3200.5, format: F.CURRENCY_PLN }, null],
    ], ['Amount', 'Rate']);
    await xlsb3.finalize();
    console.log('test_formats.xlsb OK');
}

main().catch(e => { console.error('Error:', e); process.exit(1); });