# API Documentation

## Table of Contents

- [XlsbWriter](#xlsbwriter)
- [XlsxWriter](#xlsxwriter)
- [XlsbReader](#xlsbreader)
- [XlsxReader](#xlsxreader)
- [ReaderFactory](#readerfactory)
- [Types](#types)

---

## XlsbWriter

High-performance writer for Excel Binary Workbook (.xlsb) format.

### Constructor

```typescript
new XlsbWriter(filePath: string)
```

**Parameters:**
- `filePath` - Path where the XLSB file will be created

**Example:**
```typescript
const writer = new XlsbWriter('output.xlsb');
```

### Methods

#### `addSheet(sheetName: string, hidden?: boolean): void`

Adds a new worksheet to the workbook.

**Parameters:**
- `sheetName` - Name of the worksheet (max 31 characters, invalid characters will be sanitized)
- `hidden` - Optional. If `true`, the sheet will be hidden (default: `false`)

**Example:**
```typescript
writer.addSheet('Sales Data');
writer.addSheet('Hidden Sheet', true);
```

#### `writeSheet(rows: any[][], headers?: string[] | null, doAutofilter?: boolean): void`

Writes data to the current sheet.

**Parameters:**
- `rows` - 2D array of data to write
- `headers` - Optional. Array of header strings (if not provided, first row is used as headers)
- `doAutofilter` - Optional. If `true`, adds autofilter to headers (default: `true`)

**Supported Data Types:**
- `string` - Written as shared string
- `number` - Written as number (integer or double)
- `boolean` - Written as boolean
- `Date` - Written as Excel date serial number
- `null` / `undefined` - Written as empty cell

**Example:**
```typescript
writer.writeSheet([
    ['Product', 'Price', 'Quantity'],
    ['Widget A', 19.99, 100],
    ['Widget B', 29.99, 50]
]);

// With separate headers
writer.writeSheet(
    [['Widget A', 19.99, 100], ['Widget B', 29.99, 50]],
    ['Product', 'Price', 'Quantity']
);
```

#### `finalize(): Promise<void>`

Finalizes and closes the workbook. Must be called after all data is written.

**Example:**
```typescript
await writer.finalize();
```

---

## XlsxWriter

Writer for Excel Open XML Workbook (.xlsx) format. Has the same API as XlsbWriter.

### Constructor

```typescript
new XlsxWriter(filePath: string)
```

### Methods

All methods are identical to [XlsbWriter](#xlsbwriter):
- `addSheet(sheetName: string, hidden?: boolean): void`
- `writeSheet(rows: any[][], headers?: string[] | null, doAutofilter?: boolean): void`
- `finalize(): Promise<void>`

---

## XlsbReader

High-performance reader for Excel Binary Workbook (.xlsb) format.

### Constructor

```typescript
new XlsbReader()
```

### Properties

#### `fieldCount: number`

Number of columns in the current row.

#### `resultsCount: number`

Number of sheets in the workbook.

### Methods

#### `open(path: string, readSharedStrings?: boolean): Promise<void>`

Opens an XLSB file for reading.

**Parameters:**
- `path` - Path to the XLSB file
- `readSharedStrings` - Optional. If `true`, reads shared strings table (default: `true`)

**Example:**
```typescript
const reader = new XlsbReader();
await reader.open('data.xlsb');
```

#### `getSheetNames(): string[]`

Returns array of worksheet names.

**Example:**
```typescript
const sheets = reader.getSheetNames();
console.log(sheets); // ['Sheet1', 'Sheet2']
```

#### `read(): boolean`

Reads the next row. Returns `true` if a row was read, `false` if end of sheet reached.

**Example:**
```typescript
while (reader.read()) {
    // Process row
}
```

#### `getValue(columnIndex: number): any`

Gets value at the specified column index in the current row.

**Parameters:**
- `columnIndex` - Zero-based column index

**Returns:** Value at the column (string, number, boolean, Date, or null)

**Example:**
```typescript
while (reader.read()) {
    const name = reader.getValue(0);
    const age = reader.getValue(1);
    console.log(`${name}: ${age}`);
}
```

---

## XlsxReader

Reader for Excel Open XML Workbook (.xlsx) format.

### Constructor

```typescript
new XlsxReader()
```

### Properties

Same as [XlsbReader](#xlsbreader):
- `fieldCount: number`
- `resultsCount: number`

### Methods

#### `open(path: string, readSharedStrings?: boolean): Promise<void>`

Opens an XLSX file for reading.

#### `close(): Promise<void>`

Closes the reader and releases resources.

#### `getSheetNames(): string[]`

Returns array of worksheet names.

#### `read(): Promise<boolean>`

Reads the next row. Returns `true` if a row was read.

> **Note:** Unlike XlsbReader, this method is async due to the XML parsing nature.

#### `getValue(columnIndex: number): any`

Gets value at the specified column index.

---

## ReaderFactory

Factory class for creating appropriate reader based on file extension.

### Methods

#### `static create(filePath: string): XlsbReader | XlsxReader`

Creates a reader instance based on file extension.

**Parameters:**
- `filePath` - Path to the Excel file

**Returns:** `XlsbReader` for .xlsb files, `XlsxReader` for .xlsx files

**Example:**
```typescript
const reader = ReaderFactory.create('data.xlsb');
await reader.open('data.xlsb');
```

---

## Types

### Supported Cell Types

| TypeScript Type | Excel Cell Type |
|-----------------|-----------------|
| `string` | Shared String |
| `number` (integer) | RK Number |
| `number` (float) | Double |
| `boolean` | Boolean |
| `Date` | Date Serial Number |
| `null` / `undefined` | Empty Cell |

### Sheet Name Restrictions

Sheet names have the following restrictions:
- Maximum 31 characters
- Cannot contain: `\ / * ? [ ] :`
- Cannot be empty

Invalid characters are automatically removed when adding sheets.
