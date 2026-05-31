import { XlsbWriter } from '../src/XlsbWriter';
import { XlsxWriter } from '../src/XlsxWriter';
import { F, FormattedCell } from '../src/Formats';
import * as fs from 'fs';

const OUTPUT_DIR = './test-output';

if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

const demoDate = new Date(2026, 5, 1, 14, 34, 20); // June 1, 2026 14:34:20
const demoValue = 100000;

// ========== XLSB Format Demo ==========
async function createXlsbFormatDemo() {
    const wb = new XlsbWriter(`${OUTPUT_DIR}/formats_demo_xlsb.xlsb`);

    // Sheet 1: Number formats
    wb.addSheet('Number Formats');
    wb.writeSheet([
        ['Format Name', 'Value'],
        ['No formatting', demoValue],
        ['Thousands separator', { value: demoValue, format: F.THOUSANDS_SEP }],
        ['Currency PLN', { value: demoValue, format: F.CURRENCY_PLN }],
        ['Currency EUR', { value: demoValue, format: F.CURRENCY_EUR }],
        ['Percentage', { value: 1000, format: F.PERCENTAGE }],
        ['Scientific', { value: demoValue, format: F.SCIENTIFIC }],
        ['Two decimals', { value: demoValue, format: F.TWO_DECIMALS }],
        ['Text', { value: demoValue, format: F.TEXT }],
        ['Leading zeros', { value: demoValue, format: F.LEADING_ZEROS }],
    ], ['Format Name', 'Value']);

    // Sheet 2: Date & DateTime formats
    const dateRows: any[] = [
        ['Format Name', 'Value'],
        [],
        ['--- Date Formats ---', ''],
        ['No formatting (number)', demoDate.getTime()],
        ['Short (dd.mm.yyyy)', { value: demoDate, format: F.DATE_SHORT }],
        ['Long (d mmmm yyyy)', { value: demoDate, format: F.DATE_LONG }],
        ['Day-Month-Year (dd-mm-yyyy)', { value: demoDate, format: F.DATE_DAY_MONTH_YEAR }],
        ['ISO (yyyy-mm-dd)', { value: demoDate, format: F.DATE_ISO }],
        ['Month Year (mmmm yyyy)', { value: demoDate, format: F.DATE_MONTH_YEAR }],
        ['Weekday', { value: demoDate, format: F.DATE_WEEKDAY }],
        ['Day Month (d mmmm)', { value: demoDate, format: F.DATE_DAY_MONTH }],
        ['Year only (yyyy)', { value: demoDate, format: F.DATE_YEAR_ONLY }],
        [],
        ['--- DateTime Formats ---', ''],
        ['Short DT (dd.mm.yyyy hh:mm)', { value: demoDate, format: F.DATETIME_SHORT }],
        ['Long DT (d mmmm yyyy hh:mm:ss)', { value: demoDate, format: F.DATETIME_LONG }],
        ['Time only (hh:mm)', { value: demoDate, format: F.TIME_HH_MM }],
        ['Time only (hh:mm:ss)', { value: demoDate, format: F.TIME_HH_MM_SS }],
        ['Time 12h', { value: demoDate, format: F.TIME_12H }],
        ['24h DT (dd.mm.yyyy hh:mm:ss)', { value: demoDate, format: F.DATETIME_24H }],
        ['ISO DT', { value: demoDate, format: F.DATETIME_ISO }],
        ['Milliseconds (hh:mm:ss.000)', { value: demoDate, format: F.TIME_MS }],
    ];

    wb.addSheet('Date Formats');
    wb.writeSheet(dateRows, ['Format Name', 'Value']);

    await wb.finalize();
    console.log(`✅ XLSB format demo: ${OUTPUT_DIR}/formats_demo_xlsb.xlsb`);
}

// ========== XLSX Format Demo ==========
async function createXlsxFormatDemo() {
    const wb = new XlsxWriter(`${OUTPUT_DIR}/formats_demo_xlsx.xlsx`);

    wb.addSheet('Number Formats');
    wb.writeSheet([
        ['Format Name', 'Value'],
        ['No formatting', demoValue],
        ['Thousands separator', { value: demoValue, format: F.THOUSANDS_SEP }],
        ['Currency PLN', { value: demoValue, format: F.CURRENCY_PLN }],
        ['Currency EUR', { value: demoValue, format: F.CURRENCY_EUR }],
        ['Percentage', { value: 1000, format: F.PERCENTAGE }],
        ['Scientific', { value: demoValue, format: F.SCIENTIFIC }],
        ['Two decimals', { value: demoValue, format: F.TWO_DECIMALS }],
        ['Text', { value: demoValue, format: F.TEXT }],
        ['Leading zeros', { value: demoValue, format: F.LEADING_ZEROS }],
    ], ['Format Name', 'Value']);

    wb.addSheet('Date Formats');
    wb.writeSheet([
        ['Format Name', 'Value'],
        [],
        ['--- Date Formats ---', ''],
        ['Short (dd.mm.yyyy)', { value: demoDate, format: F.DATE_SHORT }],
        ['Long (d mmmm yyyy)', { value: demoDate, format: F.DATE_LONG }],
        ['ISO (yyyy-mm-dd)', { value: demoDate, format: F.DATE_ISO }],
        ['Month Year (mmmm yyyy)', { value: demoDate, format: F.DATE_MONTH_YEAR }],
        ['Weekday', { value: demoDate, format: F.DATE_WEEKDAY }],
        ['Year only (yyyy)', { value: demoDate, format: F.DATE_YEAR_ONLY }],
        [],
        ['--- DateTime Formats ---', ''],
        ['Short DT (dd.mm.yyyy hh:mm)', { value: demoDate, format: F.DATETIME_SHORT }],
        ['Long DT (d mmmm yyyy hh:mm:ss)', { value: demoDate, format: F.DATETIME_LONG }],
        ['Time only (hh:mm)', { value: demoDate, format: F.TIME_HH_MM }],
        ['Time only (hh:mm:ss)', { value: demoDate, format: F.TIME_HH_MM_SS }],
        ['Time 12h', { value: demoDate, format: F.TIME_12H }],
        ['24h DT (dd.mm.yyyy hh:mm:ss)', { value: demoDate, format: F.DATETIME_24H }],
        ['ISO DT', { value: demoDate, format: F.DATETIME_ISO }],
        ['Milliseconds (hh:mm:ss.000)', { value: demoDate, format: F.TIME_MS }],
    ], ['Format Name', 'Value']);

    await wb.finalize();
    console.log(`✅ XLSX format demo: ${OUTPUT_DIR}/formats_demo_xlsx.xlsx`);
}

// ========== Streaming API Demo ==========
async function createStreamingFormatDemo() {
    const wb = new XlsbWriter(`${OUTPUT_DIR}/formats_streaming_demo.xlsb`);

    wb.startSheet('Number Formats', 2, ['Format', 'Value']);
    wb.writeRow(['No formatting', 123456.78]);
    wb.writeRow(['Thousands separator', { value: 123456.78, format: F.THOUSANDS_SEP }]);
    wb.writeRow(['Currency PLN', { value: 123456.78, format: F.CURRENCY_PLN }]);
    wb.writeRow(['Currency EUR', { value: 123456.78, format: F.CURRENCY_EUR }]);
    wb.writeRow(['Percentage', { value: 0.25, format: F.PERCENTAGE }]);
    wb.writeRow(['Scientific', { value: 123456.78, format: F.SCIENTIFIC }]);
    wb.writeRow(['Two decimals', { value: 123456.78, format: F.TWO_DECIMALS }]);
    wb.writeRow(['Text', { value: '123456.78', format: F.TEXT }]);
    wb.writeRow(['Leading zeros', { value: 42, format: F.LEADING_ZEROS }]);
    wb.endSheet();

    wb.startSheet('Date Formats', 2, ['Format', 'Value']);
    wb.writeRow(['Short date', { value: demoDate, format: F.DATE_SHORT }]);
    wb.writeRow(['Long date', { value: demoDate, format: F.DATE_LONG }]);
    wb.writeRow(['Short datetime', { value: demoDate, format: F.DATETIME_SHORT }]);
    wb.writeRow(['Long datetime', { value: demoDate, format: F.DATETIME_LONG }]);
    wb.writeRow(['Time only', { value: demoDate, format: F.TIME_HH_MM_SS }]);
    wb.writeRow(['Time 12h', { value: demoDate, format: F.TIME_12H }]);
    wb.writeRow(['ISO', { value: demoDate, format: F.DATETIME_ISO }]);
    wb.writeRow(['Milliseconds', { value: demoDate, format: F.TIME_MS }]);
    wb.endSheet();

    await wb.finalize();
    console.log(`✅ XLSB streaming format demo: ${OUTPUT_DIR}/formats_streaming_demo.xlsb`);
}

// ========== Real-world Scenario Demo ==========
async function createSalesReportDemo() {
    const wb = new XlsxWriter(`${OUTPUT_DIR}/sales_report_demo.xlsx`);

    wb.addSheet('Sales');
    wb.writeSheet([
        ['Product', 'Price', 'Cost', 'Margin', 'Sale Date', 'Qty', 'Total'],
        ['Laptop Dell XPS 15', { value: 5499.99, format: F.CURRENCY_PLN }, { value: 4200.00, format: F.CURRENCY_PLN }, { value: 0.30, format: F.PERCENTAGE }, { value: new Date(2026, 5, 1), format: F.DATE_SHORT }, { value: 5, format: F.THOUSANDS_SEP }, { value: 27499.95, format: F.CURRENCY_PLN }],
        ['Monitor LG 27"', { value: 1299.00, format: F.CURRENCY_PLN }, { value: 950.00, format: F.CURRENCY_PLN }, { value: 0.37, format: F.PERCENTAGE }, { value: new Date(2026, 5, 2), format: F.DATE_SHORT }, 3, { value: 3897.00, format: F.CURRENCY_PLN }],
        ['Mechanical keyboard', { value: 349.99, format: F.CURRENCY_PLN }, { value: 200.00, format: F.CURRENCY_PLN }, { value: 0.75, format: F.PERCENTAGE }, { value: new Date(2026, 5, 3), format: F.DATE_SHORT }, 10, { value: 3499.90, format: F.CURRENCY_PLN }],
        ['Wireless mouse', { value: 199.00, format: F.CURRENCY_PLN }, { value: 120.00, format: F.CURRENCY_PLN }, { value: 0.66, format: F.PERCENTAGE }, { value: new Date(2026, 5, 4), format: F.DATE_SHORT }, 8, { value: 1592.00, format: F.CURRENCY_PLN }],
        ['SSD Drive 1TB', { value: 599.00, format: F.CURRENCY_PLN }, { value: 380.00, format: F.CURRENCY_PLN }, { value: 0.58, format: F.PERCENTAGE }, { value: new Date(2026, 5, 5), format: F.DATE_SHORT }, 2, { value: 1198.00, format: F.CURRENCY_PLN }],
    ], ['Product', 'Price', 'Cost', 'Margin', 'Sale Date', 'Qty', 'Total']);

    wb.addSheet('Summary');
    wb.writeSheet([
        ['Total Sales', { value: 37486.85, format: F.CURRENCY_PLN }],
    ], ['Description', 'Value']);

    await wb.finalize();
    console.log(`✅ Sales report demo: ${OUTPUT_DIR}/sales_report_demo.xlsx`);
}

// ========== All-in-one comprehensive demo ==========
async function createComprehensiveDemo() {
    const wb = new XlsbWriter(`${OUTPUT_DIR}/comprehensive_formats_demo.xlsb`);

    wb.addSheet('All Formats');
    const rows: any[] = [['Category', 'Format Name', 'Value', 'Format Code']];

    // Number formats
    const numVal = 1234567.89;
    rows.push(['Numeric', 'No formatting', numVal, 'Raw number']);
    rows.push(['Numeric', 'Thousands separator', { value: numVal, format: F.THOUSANDS_SEP }, '#,##0']);
    rows.push(['Numeric', 'Currency PLN', { value: numVal, format: F.CURRENCY_PLN }, '#,##0.00 "zł"']);
    rows.push(['Numeric', 'Currency EUR', { value: numVal, format: F.CURRENCY_EUR }, '#,##0.00 €']);
    rows.push(['Numeric', 'Percentage', { value: 0.125, format: F.PERCENTAGE }, '0%']);
    rows.push(['Numeric', 'Scientific', { value: numVal, format: F.SCIENTIFIC }, '0.00E+00']);
    rows.push(['Numeric', 'Two decimals', { value: numVal, format: F.TWO_DECIMALS }, '#,##0.00']);
    rows.push(['Numeric', 'Text', { value: numVal, format: F.TEXT }, '@']);
    rows.push(['Numeric', 'Leading zeros', { value: 42, format: F.LEADING_ZEROS }, '000000000']);

    // Date formats
    const dt = new Date(2026, 11, 24, 15, 45, 30); // Dec 24, 2026 15:45:30
    rows.push([]);
    rows.push(['Date', 'Short date', { value: dt, format: F.DATE_SHORT }, 'dd.mm.yyyy']);
    rows.push(['Date', 'Long date', { value: dt, format: F.DATE_LONG }, 'd mmmm yyyy']);
    rows.push(['Date', 'Day-Month-Year', { value: dt, format: F.DATE_DAY_MONTH_YEAR }, 'dd-mm-yyyy']);
    rows.push(['Date', 'ISO date', { value: dt, format: F.DATE_ISO }, 'yyyy-mm-dd']);
    rows.push(['Date', 'Month Year', { value: dt, format: F.DATE_MONTH_YEAR }, 'mmmm yyyy']);
    rows.push(['Date', 'Weekday', { value: dt, format: F.DATE_WEEKDAY }, 'dddd, d mmmm yyyy']);
    rows.push(['Date', 'Day Month', { value: dt, format: F.DATE_DAY_MONTH }, 'd mmmm']);
    rows.push(['Date', 'Year only', { value: dt, format: F.DATE_YEAR_ONLY }, 'yyyy']);

    // DateTime formats
    rows.push([]);
    rows.push(['DateTime', 'Short DT', { value: dt, format: F.DATETIME_SHORT }, 'dd.mm.yyyy hh:mm']);
    rows.push(['DateTime', 'Long DT', { value: dt, format: F.DATETIME_LONG }, 'd mmmm yyyy hh:mm:ss']);
    rows.push(['DateTime', 'Time hh:mm', { value: dt, format: F.TIME_HH_MM }, 'hh:mm']);
    rows.push(['DateTime', 'Time hh:mm:ss', { value: dt, format: F.TIME_HH_MM_SS }, 'hh:mm:ss']);
    rows.push(['DateTime', 'Time 12h', { value: dt, format: F.TIME_12H }, 'h:mm AM/PM']);
    rows.push(['DateTime', '24h DT', { value: dt, format: F.DATETIME_24H }, 'dd.mm.yyyy hh:mm:ss']);
    rows.push(['DateTime', 'ISO DT', { value: dt, format: F.DATETIME_ISO }, 'yyyy-mm-dd"T"hh:mm:ss']);
    rows.push(['DateTime', 'Milliseconds', { value: dt, format: F.TIME_MS }, 'hh:mm:ss.000']);

    wb.writeSheet(rows);
    await wb.finalize();
    console.log(`✅ Comprehensive format demo: ${OUTPUT_DIR}/comprehensive_formats_demo.xlsb`);
}

async function main() {
    console.log('=== SpreadSheet Tasks Format Demo ===\n');
    await createXlsbFormatDemo();
    await createXlsxFormatDemo();
    await createStreamingFormatDemo();
    await createSalesReportDemo();
    await createComprehensiveDemo();
    console.log('\n✅ All format demos generated!');
    console.log('📁 Open these files in Excel to verify formatting:');
    console.log(`   ${OUTPUT_DIR}/`);
}

main().catch(console.error);
