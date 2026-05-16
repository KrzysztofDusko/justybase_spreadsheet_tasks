# SpreadsheetTasks

A high-performance TypeScript library for reading and writing Excel files in XLSB and XLSX formats.

[![npm version](https://badge.fury.io/js/@justybase%2Fspreadsheet-tasks.svg)](https://www.npmjs.com/package/@justybase/spreadsheet-tasks)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## 🚀 Features

- **High Performance** - XLSB format is 3.3x faster to write and 2.1x faster to read than XLSX
- **Small File Size** - XLSB files are ~47% smaller than equivalent XLSX files
- **TypeScript First** - Full TypeScript support with type definitions
- **Dual Format Support** - Read and write both XLSB and XLSX formats
- **Zero External Dependencies** - Uses only Node.js built-in modules and minimal zip libraries
- **Streaming Support** - Efficient memory usage for large files
- **Multiple Sheets** - Support for multiple worksheets per workbook
- **Auto-filter** - Automatic filter headers support

## 📊 Benchmark Results

| Operation | XLSB | XLSX | Performance Gain |
|-----------|------|------|------------------|
| Write (50K rows) | 140 ms | 467 ms | **3.3x faster** |
| Read (50K rows) | 118 ms | 276 ms | **2.3x faster** |
| File Size | 1.49 MB | 2.83 MB | **47% smaller** |

## 📦 Installation

```bash
npm install @justybase/spreadsheet-tasks
```

## 🌐 Browser Support & CDN (Zero Dependencies)

This library includes browser-compatible modules to generate XLSB and XLSX files directly on the client side without relying on Node.js APIs or external dependencies.

To bundle the browser version into a single minified file (e.g. for deployment to a CDN), run the following command in the project root:

```bash
npx esbuild browser/browser-spreadsheet.js --bundle --minify --format=esm --outfile=browser/justybase-spreadsheet.min.js
```

You can then import this single file via a script tag in any web application:

```html
<script type="module">
  import { downloadXlsb, downloadXlsx } from 'https://your-cdn.com/justybase-spreadsheet.min.js';
  
  const headers = ["ID", "Name", "Score"];
  const rows = [
    [1, "Alice", 99.5],
    [2, "Bob", 88.0]
  ];
  
  // Generates the file and triggers the browser download automatically
  document.getElementById('downloadBtn').addEventListener('click', () => {
    downloadXlsb('report.xlsb', rows, headers, { headerStyle: 'bold+fill' });
  });
</script>
```

**Features included in the browser build:**
- Built-in Zip archiving (no external libraries needed)
- Memory-efficient `Uint8Array` buffer handling
- Auto-fitted column widths
- Frozen and styled headers (styles: `bold`, `fill`, `bold+fill`)
- Auto-filters enabled by default

## 🔧 Quick Start

### Writing Excel Files

```typescript
import { XlsbWriter, XlsxWriter } from '@justybase/spreadsheet-tasks';

// Create XLSB file (recommended for performance)
const xlsbWriter = new XlsbWriter('output.xlsb');
xlsbWriter.addSheet('Sheet1');
xlsbWriter.writeSheet([
    ['Name', 'Age', 'City'],
    ['Alice', 30, 'New York'],
    ['Bob', 25, 'Los Angeles'],
    ['Charlie', 35, 'Chicago']
]);
await xlsbWriter.finalize();

// Or create XLSX file for compatibility
const xlsxWriter = new XlsxWriter('output.xlsx');
xlsxWriter.addSheet('Sheet1');
xlsxWriter.writeSheet([
    ['Name', 'Age', 'City'],
    ['Alice', 30, 'New York'],
    ['Bob', 25, 'Los Angeles'],
    ['Charlie', 35, 'Chicago']
]);
await xlsxWriter.finalize();
```

### Streaming API (for Large Datasets)

For large datasets that don't fit in memory, use the streaming API to write rows one at a time:

```typescript
import { XlsbWriter } from '@justybase/spreadsheet-tasks';

const writer = new XlsbWriter('large-output.xlsb');

// Start a sheet with column count and optional headers
writer.startSheet('Data', 5, ['ID', 'Name', 'Value', 'Date', 'Active']);

// Write rows one at a time - no need to load all data in memory
for (let i = 0; i < 1_000_000; i++) {
    writer.writeRow([
        i + 1,
        `User_${i + 1}`,
        Math.random() * 10000,
        new Date(),
        i % 2 === 0
    ]);
}

// Finalize the sheet
writer.endSheet();

// You can create multiple sheets
writer.startSheet('MoreData', 3, ['Col1', 'Col2', 'Col3']);
// ... write more rows
writer.endSheet();

await writer.finalize();
```

**Benefits of Streaming API:**
- ✅ Constant memory usage regardless of dataset size
- ✅ Write millions of rows without loading all data into RAM
- ✅ Generate data on-the-fly from databases, APIs, or other sources
- ✅ Same performance as batch mode

### Reading Excel Files


```typescript
import { XlsbReader, XlsxReader, ReaderFactory } from '@justybase/spreadsheet-tasks';

// Using ReaderFactory (auto-detects format)
const reader = ReaderFactory.create('data.xlsb');
await reader.open('data.xlsb');

console.log('Sheet names:', reader.getSheetNames());

while (reader.read()) {
    const row = [];
    for (let i = 0; i < reader.fieldCount; i++) {
        row.push(reader.getValue(i));
    }
    console.log(row);
}

// Or use specific reader
const xlsbReader = new XlsbReader();
await xlsbReader.open('data.xlsb');
// ... same API as above
```

## 📖 API Documentation

See [API Documentation](./docs/API.md) for detailed API reference.

## 💡 Examples

Check out the [examples](./examples) folder for more usage examples:

- [Basic Write](./examples/basic-write.ts) - Writing data to Excel files
- [Basic Read](./examples/basic-read.ts) - Reading data from Excel files
- [Multiple Sheets](./examples/multiple-sheets.ts) - Working with multiple worksheets
- [Large Dataset](./examples/large-dataset.ts) - Handling large datasets efficiently
- [Streaming Example](./examples/streaming-example.ts) - Write millions of rows without loading all data in memory

Run examples with:
```bash
npm run build
npx ts-node examples/basic-write.ts
```


## 🔍 When to Use XLSB vs XLSX

| Use XLSB When | Use XLSX When |
|---------------|---------------|
| Performance is critical | Need maximum compatibility |
| Working with large datasets | Sharing with older Excel versions |
| Internal/backend processing | Human-readable XML is needed |
| Storage space is limited | Third-party integrations require it |

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Inspired by and reimplemented from [SpreadSheetTasks](https://github.com/KrzysztofDusko/SpreadSheetTasks)
