import * as fs from 'fs';
import * as path from 'path';
import { XlsbWriter } from '../../src/XlsbWriter';
import { XlsxWriter } from '../../src/XlsxWriter';
import { F } from '../../src/Formats';

const ROOT = path.resolve(__dirname, '../..');

interface SuiteDef {
  id: string;
  name: string;
  sheets: number;
  filter: boolean;
  special: string;
  generate: (w: any, formats: any) => void;
}

const SUITES: SuiteDef[] = [
  {
    id: '01',
    name: 'basic-single',
    sheets: 1, filter: false, special: 'Minimal smoke test',
    generate: (w) => {
      w.addSheet('Basic');
      w.writeSheet([[1, 'Hello', true, null]], ['Number', 'Text', 'Bool', 'Empty']);
    }
  },
  {
    id: '02',
    name: 'basic-multi-2',
    sheets: 2, filter: true, special: 'Multi-sheet baseline',
    generate: (w) => {
      w.addSheet('Sheet1');
      w.writeSheet([['A1', 'B1'], ['A2', 'B2']], ['ColA', 'ColB']);
      w.addSheet('Sheet2');
      w.writeSheet([[10, 20], [30, 40]], ['X', 'Y']);
    }
  },
  {
    id: '03',
    name: 'basic-multi-4',
    sheets: 4, filter: true, special: 'Our "simple working" case',
    generate: (w) => {
      for (let i = 1; i <= 4; i++) {
        w.addSheet(`Sheet${i}`);
        w.writeSheet([[i, i * 10]], ['ID', 'Value']);
      }
    }
  },
  {
    id: '04',
    name: 'boundary-10',
    sheets: 10, filter: true, special: 'Exactly 10 filtered sheets',
    generate: (w) => {
      for (let i = 1; i <= 10; i++) {
        w.addSheet(`Sheet${i}`);
        w.writeSheet([[i, `data${i}`]], ['No', 'Data']);
      }
    }
  },
  {
    id: '05',
    name: 'bug-regression-15',
    sheets: 15, filter: true, special: 'The exact bug: >10 filters + formats',
    generate: (w) => {
      // Sheet 1: Basic Types
      w.addSheet('Basic Types');
      w.writeSheet([[42, 3.14159, 'Hello World', new Date(2025, 5, 1), true, null]],
        ['Integer', 'Floating Point', 'Text', 'Date', 'Boolean', 'Empty']);
      // Sheet 2: Finances
      w.addSheet('Finances');
      w.writeSheet([
        ['Deposit', { value: 15000.0, format: F.CURRENCY_PLN }, null, null, { value: 0.15, format: F.PERCENTAGE }]
      ], ['Description', 'Income', 'Expense', 'Balance', 'Percent']);
      // Sheet 3: Dates
      w.addSheet('Dates and Time');
      w.writeSheet([
        [{ value: new Date(2025, 0, 1), format: F.DATE_ISO }, '08:00:00', { value: new Date(2025, 0, 1, 8, 0), format: F.DATETIME_ISO }, 'Wednesday']
      ], ['Date', 'Time', 'Date and Time', 'Weekday']);
      // Sheet 4: Scientific
      w.addSheet('Scientific');
      w.writeSheet([
        ['Planck Constant', { value: 6.62607015e-34, format: F.SCIENTIFIC }, '6.626×10⁻³⁴', 'J·s']
      ], ['Parameter', 'Value', 'Notation', 'Unit']);
      // Sheet 5: Geographic
      w.addSheet('Geographic');
      w.writeSheet([
        ['Warsaw', 'Poland', 52.2297, 21.0122, 'UTC+1', 1.79]
      ], ['City', 'Country', 'Latitude', 'Longitude', 'Zone', 'Population']);
      // Sheet 6: Warehouse
      w.addSheet('Warehouse');
      w.writeSheet([
        ['P-001', 'Dell Laptop', 'Electronics', { value: 4500, format: F.CURRENCY_PLN }, '23%', 15]
      ], ['ID', 'Name', 'Category', 'Net Price', 'VAT', 'Stock']);
      // Sheet 7: Grades
      w.addSheet('Grades');
      w.writeSheet([
        ['Anna Nowak', 5, 4, 6, 5, 5]
      ], ['Student', 'Math', 'Physics', 'CS', 'Polish', 'English']);
      // Sheet 8: Various Types
      w.addSheet('Various Types');
      w.writeSheet([
        ['Integer', 1, 2, 3, 'integers'],
        ['Float', 0.1, 0.01, 0.001, 'floating point'],
        ['String', 'Ala', 'ma', 'kota', 'strings'],
        ['Date', new Date(2025, 0, 1), new Date(2025, 5, 1), new Date(2025, 11, 31), 'dates'],
        ['Boolean', true, false, true, 'booleans'],
        ['Null', null, null, null, 'empty cells']
      ], ['Type', 'Value 1', 'Value 2', 'Value 3', 'Description']);
      // Sheet 9: Employees
      w.addSheet('Employees');
      w.writeSheet([
        [1, 'Adam', 'Mickiewicz', 'IT', 'Developer', new Date(2020, 2, 1), { value: 12000, format: F.CURRENCY_PLN }]
      ], ['ID', 'First Name', 'Last Name', 'Department', 'Position', 'Hire Date', 'Salary']);
      // Sheet 10: Sales
      w.addSheet('Sales');
      w.writeSheet([
        ['Q1 2025', 45000, 32000, 28000, 120000]
      ], ['Quarter', 'Product A', 'Product B', 'Product C', 'Target']);
      // Sheet 11: Template
      w.addSheet('Template');
      const tpl = [];
      for (let i = 1; i <= 10; i++) tpl.push([i, '', 0, 0.0]);
      w.writeSheet(tpl, ['No.', 'Name', 'Quantity', 'Price']);
      // Sheet 12: Notes
      w.addSheet('Notes');
      w.writeSheet([
        [1, 'Important Meeting', 'Meeting with client regarding new project.', 'AD', new Date(2025, 5, 1), 'High']
      ], ['ID', 'Title', 'Content', 'Author', 'Date', 'Priority']);
      // Sheet 13: Percentages
      w.addSheet('Percentages');
      const total = 450000 + 320000 + 180000;
      w.writeSheet([
        ['Online Sales', 450000, { value: 450000 / total, format: F.PERCENTAGE }],
        ['Retail Sales', 320000, { value: 320000 / total, format: F.PERCENTAGE }],
        ['Export', 180000, { value: 180000 / total, format: F.PERCENTAGE }]
      ], ['Category', 'Value', 'Share %']);
      // Sheet 14: Flags
      w.addSheet('Flags');
      w.writeSheet([
        [1, 'user1', true, false, false],
        [2, 'user2', true, false, true],
        [3, 'user3', false, false, false],
        [4, 'admin1', true, true, false],
        [5, 'mod1', true, false, true]
      ], ['ID', 'Username', 'Active', 'Admin', 'Blocked']);
      // Sheet 15: TOC
      w.addSheet('Table of Contents');
      w.writeSheet([
        [1, 'Basic Types', 'int, float, string, date, bool, null', 2, 6],
        [2, 'Finances', 'currencies and percentages', 3, 5],
        [3, 'Dates and Time', 'date, time, datetime ISO', 3, 4],
        [4, 'Scientific', 'scientific notation', 4, 4],
        [5, 'Geographic', 'city coordinates', 5, 6],
        [6, 'Warehouse', 'prices, stock levels', 5, 6],
        [7, 'Grades', 'school grades 1-6', 5, 6],
        [8, 'Various Types', 'mix: int, float, string, date, bool, null', 6, 5],
        [9, 'Employees', 'personal data, dates, salaries', 5, 7],
        [10, 'Sales', 'quarterly sales', 4, 5],
        [11, 'Template', 'empty template 10 rows', 10, 4],
        [12, 'Notes', 'long texts, dates, priorities', 4, 6],
        [13, 'Percentages', 'percentage shares', 3, 3],
        [14, 'Flags', 'boolean flags', 5, 5],
        [15, 'Table of Contents', 'this sheet', 15, 5]
      ], ['No.', 'Sheet Name', 'Description', 'Rows', 'Columns']);
    }
  },
  {
    id: '06',
    name: 'boundary-20',
    sheets: 20, filter: true, special: 'Exactly 20 filtered sheets',
    generate: (w) => {
      for (let i = 1; i <= 20; i++) {
        w.addSheet(`Sheet${i}`);
        w.writeSheet([[i, `row${i}`]], ['ID', 'Data']);
      }
    }
  },
  {
    id: '07',
    name: 'over-boundary-25',
    sheets: 25, filter: true, special: '>20 filters (tests 0x80 branch)',
    generate: (w) => {
      for (let i = 1; i <= 25; i++) {
        w.addSheet(`Sheet${i}`);
        w.writeSheet([[i, `row${i}`]], ['ID', 'Data']);
      }
    }
  },
  {
    id: '08',
    name: 'formats-all',
    sheets: 1, filter: false, special: 'All 23 format constants',
    generate: (w) => {
      const d = new Date(2026, 5, 1, 14, 34, 20);
      w.addSheet('Formats');
      w.writeSheet([
        ['Plain', 1234567.89],
        ['Thousands', { value: 1234567.89, format: F.THOUSANDS_SEP }],
        ['Currency PLN', { value: 1234567.89, format: F.CURRENCY_PLN }],
        ['Currency EUR', { value: 1234567.89, format: F.CURRENCY_EUR }],
        ['Percent', { value: 0.125, format: F.PERCENTAGE }],
        ['Scientific', { value: 1234567.89, format: F.SCIENTIFIC }],
        ['Two Decimals', { value: 1234567.89, format: F.TWO_DECIMALS }],
        ['Text', { value: 1234567.89, format: F.TEXT }],
        ['Leading Zeros', { value: 42, format: F.LEADING_ZEROS }],
        ['Date Short', { value: d, format: F.DATE_SHORT }],
        ['Date Long', { value: d, format: F.DATE_LONG }],
        ['Date ISO', { value: d, format: F.DATE_ISO }],
        ['Date Month+Year', { value: d, format: F.DATE_MONTH_YEAR }],
        ['Date Weekday', { value: d, format: F.DATE_WEEKDAY }],
        ['Date Year Only', { value: d, format: F.DATE_YEAR_ONLY }],
        ['DateTime Short', { value: d, format: F.DATETIME_SHORT }],
        ['DateTime Long', { value: d, format: F.DATETIME_LONG }],
        ['Time HH:MM', { value: d, format: F.TIME_HH_MM }],
        ['Time HH:MM:SS', { value: d, format: F.TIME_HH_MM_SS }],
        ['Time 12h', { value: d, format: F.TIME_12H }],
        ['DateTime 24h', { value: d, format: F.DATETIME_24H }],
        ['DateTime ISO', { value: d, format: F.DATETIME_ISO }],
        ['Time MS', { value: d, format: F.TIME_MS }],
      ], ['Format', 'Value']);
    }
  },
  {
    id: '09',
    name: 'unicode-emoji',
    sheets: 1, filter: false, special: 'Unicode + emoji strings',
    generate: (w) => {
      w.addSheet('Unicode');
      w.writeSheet([
        ['Hello', 'こんにちは', 'Zażółć gęślą jaźń', '🚀🌟✨', '👨‍👩‍👧‍👦'],
        ['Test', '中文测试', 'العربية', '𝕳𝖊𝖑𝖑𝖔', '♠♥♦♣'],
      ], ['English', 'Japanese', 'Polish', 'Emoji', 'Mixed']);
    }
  },
  {
    id: '10',
    name: 'long-strings',
    sheets: 1, filter: false, special: '32k+ char strings',
    generate: (w) => {
      w.addSheet('Long Strings');
      const long = 'A'.repeat(32767);
      const medium = 'B'.repeat(1000);
      w.writeSheet([[long, medium, 'short']], ['Very Long', 'Medium', 'Short']);
    }
  },
  {
    id: '11',
    name: 'many-cols',
    sheets: 1, filter: true, special: '100 columns',
    generate: (w) => {
      w.addSheet('Wide');
      const headers = Array.from({ length: 100 }, (_, i) => `Col${i}`);
      const row = Array.from({ length: 100 }, (_, i) => i);
      w.writeSheet([row], headers);
    }
  },
  {
    id: '12',
    name: 'many-rows',
    sheets: 1, filter: true, special: '10,000 rows',
    generate: (w) => {
      w.addSheet('Tall');
      const rows = [];
      for (let i = 0; i < 10000; i++) rows.push([i, `User_${i}`, i % 100]);
      w.writeSheet(rows, ['ID', 'Name', 'Score']);
    }
  },
  {
    id: '13',
    name: 'null-mixed',
    sheets: 1, filter: false, special: 'null, undefined, bool mix',
    generate: (w) => {
      w.addSheet('Nulls');
      w.writeSheet([
        [1, 'A', true, null, undefined],
        [null, null, false, null, undefined],
        [2, 'B', true, 42, 'text'],
      ], ['Col1', 'Col2', 'Col3', 'Col4', 'Col5']);
    }
  },
  {
    id: '14',
    name: 'empty-sheet',
    sheets: 1, filter: false, special: 'Zero-row sheet',
    generate: (w) => {
      w.addSheet('Empty');
      w.writeSheet([], ['H1', 'H2']);
    }
  },
  {
    id: '15',
    name: 'scientific',
    sheets: 1, filter: false, special: 'Scientific notation values',
    generate: (w) => {
      w.addSheet('Sci');
      w.writeSheet([
        [{ value: 1.23e-10, format: F.SCIENTIFIC }],
        [{ value: 9.99e23, format: F.SCIENTIFIC }],
        [{ value: 6.626e-34, format: F.SCIENTIFIC }],
      ], ['Value']);
    }
  },
  {
    id: '16',
    name: 'streaming-api',
    sheets: 1, filter: true, special: 'startSheet/writeRow mode',
    generate: (w) => {
      w.startSheet('Stream', 3, ['A', 'B', 'C']);
      for (let i = 0; i < 5; i++) w.writeRow([i, i * 2, `row${i}`]);
      w.endSheet();
    }
  },
];

function ensureDir(dir: string) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

async function generateNodeSuites(outDir: string) {
  ensureDir(outDir);
  for (const suite of SUITES) {
    // XLSX
    const xlsxPath = path.join(outDir, `${suite.id}-${suite.name}.xlsx`);
    const xw = new XlsxWriter(xlsxPath);
    suite.generate(xw, F);
    await xw.finalize();

    // XLSB
    const xlsbPath = path.join(outDir, `${suite.id}-${suite.name}.xlsb`);
    const bw = new XlsbWriter(xlsbPath);
    suite.generate(bw, F);
    await bw.finalize();
  }
}

async function generateBrowserSuites(outDir: string) {
  ensureDir(outDir);
  const {
    BrowserXlsbWriter,
    BrowserXlsxWriter,
    F: BF,
  } = await import('../../browser/browser-spreadsheet.js');

  const browserSuites = [
    { id: '17', name: 'browser-equiv-15', sheets: 15, filter: true, special: 'Browser API: same as #05' },
    { id: '18', name: 'browser-formats', sheets: 1, filter: false, special: 'Browser API: all formats' },
  ];

  for (const meta of browserSuites) {
    const suite = SUITES.find(s => s.id === (meta.id === '17' ? '05' : '08'))!;

    // Browser XLSB
    const bw = new BrowserXlsbWriter();
    suite.generate(bw, BF);
    const bBlob = bw.finalize();
    const bBuf = Buffer.from(await bBlob.arrayBuffer());
    fs.writeFileSync(path.join(outDir, `${meta.id}-${meta.name}.xlsb`), bBuf);

    // Browser XLSX
    const xw = new BrowserXlsxWriter();
    suite.generate(xw, BF);
    const xBlob = xw.finalize();
    const xBuf = Buffer.from(await xBlob.arrayBuffer());
    fs.writeFileSync(path.join(outDir, `${meta.id}-${meta.name}.xlsx`), xBuf);
  }
}

export async function generateAllSuites(baseDir: string) {
  const nodeDir = path.join(baseDir, 'node');
  const browserDir = path.join(baseDir, 'browser');

  console.log('Generating Node API suites...');
  await generateNodeSuites(nodeDir);

  console.log('Generating Browser API suites...');
  await generateBrowserSuites(browserDir);
}

// CLI entrypoint
if (require.main === module) {
  const outDir = process.argv[2] || path.join(__dirname, 'test-output');
  generateAllSuites(outDir).then(() => {
    console.log('All suites generated.');
  }).catch(err => {
    console.error('Generation failed:', err);
    process.exit(1);
  });
}
