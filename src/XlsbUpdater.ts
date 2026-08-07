import * as fs from 'fs';
import AdmZip from 'adm-zip';
import { CellValue, getFormat, unwrapCell } from './Formats';
import {
    readRecord,
    buildRecord,
    readUtf16,
    parseSharedStringsBin,
    Biff12Record,
} from './biff12Utils';
import { trimTrailingEmptyRows } from './xmlUtils';

/**
 * Options for {@link XlsbUpdater.replaceSheetData}.
 */
export interface ReplaceSheetDataOptions {
    /**
     * Optional header row written into row 1 of the target sheet.
     * When omitted, row 1 is filled with the data rows.
     */
    headers?: string[];
    /**
     * Style strategy for the new cells:
     * - `'inherit'` (default): reuse the dominant cell style (xf) of each column from the
     *   existing data (and the original header style for row 1), so date/number
     *   formatting of the sheet is preserved. Falls back to General (0).
     * - `'general'`: all new cells use xf 0 (General).
     */
    styleFallback?: 'inherit' | 'general';
}

/** Worksheet cell records (BrtCellBlank … BrtFmlaError). */
const CELL_RECORDS = new Set<number>([0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0a, 0x0b]);

const RK_INT_LOWER = -1 << 29;
const RK_INT_UPPER = (1 << 29) - 1;

/**
 * Update the data of a worksheet inside an existing XLSB file without rebuilding
 * the workbook. Everything outside the target sheet's rows — pivot tables and
 * their caches, other sheets, styles, column widths, views — is preserved.
 * This makes it possible to refresh a data sheet that pivot tables are wired to.
 *
 * ```ts
 * const updater = new XlsbUpdater('report.xlsb');
 * updater.replaceSheetData('data1', rows, { headers: ['ID', 'NAME', 'AMOUNT'] });
 * updater.save('report_new.xlsb'); // or updater.save() to overwrite in place
 * ```
 */
export class XlsbUpdater {
    private readonly zip: AdmZip;
    private readonly sourcePath: string;

    private sheetNameToPath: Map<string, string> = new Map();
    private pivotCacheDefPaths: string[] = [];

    private sharedStringsBuf: Buffer | null = null;
    private sharedStringsValues: string[] = [];
    private sharedStringsTotal: number = 0;
    private sharedStringsUnique: number = 0;
    private sharedStringsEndSst: number = 0;
    private stringIndexMap: Map<string, number> = new Map();

    private readonly _oaEpoch: number;

    /**
     * @param path Path to an existing `.xlsb` file.
     */
    constructor(path: string) {
        this.sourcePath = path;
        this._oaEpoch = Date.UTC(1899, 11, 30);

        if (!fs.existsSync(path)) {
            throw new Error(`XlsbUpdater: file not found: ${path}`);
        }

        this.zip = new AdmZip(path);

        if (!this.zip.getEntry('xl/workbook.bin')) {
            throw new Error(
                'XlsbUpdater: not an XLSB workbook (xl/workbook.bin missing). Convert the file to XLSB first.'
            );
        }

        this._loadWorkbookStructure();
    }

    /** Worksheet names in workbook order. */
    getSheetNames(): string[] {
        return [...this.sheetNameToPath.keys()];
    }

    /**
     * Clear the entire cell data of the target sheet and write new rows in its place.
     * All other parts of the workbook (pivot tables, other sheets, styles, …) are preserved.
     *
     * @param sheetName Name of the worksheet to replace.
     * @param rows New data rows (`CellValue` cells; null/undefined = blank).
     * @param options See {@link ReplaceSheetDataOptions}.
     */
    replaceSheetData(
        sheetName: string,
        rows: CellValue[][],
        options: ReplaceSheetDataOptions = {}
    ): void {
        const { headers, styleFallback = 'inherit' } = options;
        rows = trimTrailingEmptyRows(rows);

        const sheetPath = this.sheetNameToPath.get(sheetName);
        if (!sheetPath) {
            throw new Error(`XlsbUpdater: sheet "${sheetName}" not found in the workbook.`);
        }

        const entry = this.zip.getEntry(sheetPath);
        if (!entry) {
            throw new Error(`XlsbUpdater: worksheet part missing for sheet "${sheetName}".`);
        }

        const sheetBuf = entry.getData();
        this._ensureSharedStringsLoaded();

        let dataStyles: Map<number, number> = new Map();
        let headerStyles: Map<number, number> = new Map();
        let dateXf: number | null = null;

        if (styleFallback !== 'general') {
            const collected = this._collectStyles(sheetBuf);
            dataStyles = collected.dataStyles;
            headerStyles = collected.headerStyles;
            dateXf = this._findDateXf();
        }

        const lastCol = Math.max(0, Math.max(
            rows.reduce((m, r) => Math.max(m, r.length), 0),
            headers?.length ?? 0
        ) - 1);

        const newRows = this._buildRows(rows, headers, dataStyles, headerStyles, dateXf);
        const { rowsStart, rowsEnd, autoFilterRanges } = this._findRowsRegion(sheetBuf);

        const newSheet = Buffer.alloc(sheetBuf.length - (rowsEnd - rowsStart) + newRows.length);
        sheetBuf.copy(newSheet, 0, 0, rowsStart);
        newRows.copy(newSheet, rowsStart);
        sheetBuf.copy(newSheet, rowsStart + newRows.length, rowsEnd);

        const lastRowIdx = rows.length + (headers ? 1 : 0) - 1;
        const delta = newRows.length - (rowsEnd - rowsStart);
        for (const af of autoFilterRanges) {
            const newAf = af + delta;
            newSheet.writeInt32LE(lastCol, newAf + 12);
            newSheet.writeInt32LE(lastRowIdx, newAf + 4);
        }

        // Keep the sheet dimension record (0x98) in sync with the new data. Excel uses
        // its rwLast/colLast fields to read the full row range; stale values make the
        // last written row invisible (and drop it from pivot refreshes).
        this._patchDimension(newSheet, lastRowIdx, lastCol);

        this.zip.updateFile(sheetPath, newSheet);

        if (this.sharedStringsBuf !== null && this._appendCount > 0) {
            this._commitSharedStrings();
        }

        const rwLast = Math.max(0, rows.length + (headers ? 1 : 0) - 1);
        const colLast = lastCol;
        this._patchPivotCaches(sheetName, rwLast, colLast);
    }

    /**
     * Write the updated workbook to disk.
     * @param outputPath Destination path; defaults to the source file (overwrites in place).
     */
    save(outputPath?: string): void {
        const target = outputPath ?? this.sourcePath;
        fs.writeFileSync(target, this.zip.toBuffer());
    }

    /** The updated workbook as an in-memory ZIP buffer. */
    toBuffer(): Buffer {
        return this.zip.toBuffer();
    }

    // ------------------------------------------------------------------
    // Workbook structure
    // ------------------------------------------------------------------

    private _loadWorkbookStructure(): void {
        const relsEntry = this.zip.getEntry('xl/_rels/workbook.bin.rels');
        const rIdToTarget: Record<string, string> = {};
        if (relsEntry) {
            const rels = relsEntry.getData().toString('utf8');
            const relRegex = /<Relationship[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"/g;
            let match;
            while ((match = relRegex.exec(rels)) !== null) {
                rIdToTarget[match[1]] = match[2];
            }
        }

        const wbEntry = this.zip.getEntry('xl/workbook.bin');
        if (wbEntry) {
            const wbBuf = wbEntry.getData();
            const rec: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
            let pos = 0;
            while (readRecord(wbBuf, pos, rec)) {
                pos = rec.dataEnd;
                if (rec.id !== 0x009c || rec.len < 8) continue;

                // BrtBundleSh: hidden(4) + sheetId(4) + rId(string) + name(string)
                let o = rec.dataStart + 8;
                const rIdLen = wbBuf.readUInt32LE(o);
                o += 4;
                const rId = readUtf16(wbBuf, o, rIdLen);
                o += rIdLen * 2;
                const nameLen = wbBuf.readUInt32LE(o);
                o += 4;
                const name = readUtf16(wbBuf, o, nameLen);

                let target = rIdToTarget[rId] ?? '';
                if (target.startsWith('/')) target = target.substring(1);
                else if (!target.startsWith('xl/')) target = 'xl/' + target;
                if (target) {
                    this.sheetNameToPath.set(name, target);
                }
            }
        }

        for (const e of this.zip.getEntries()) {
            if (/^xl\/pivotCache\/pivotCacheDefinition\d*\.bin$/.test(e.entryName)) {
                this.pivotCacheDefPaths.push(e.entryName);
            }
        }
    }

    // ------------------------------------------------------------------
    // Shared strings
    // ------------------------------------------------------------------

    private _appendCount: number = 0;

    private _ensureSharedStringsLoaded(): void {
        if (this.stringIndexMap.size > 0 || this.sharedStringsBuf !== null) {
            return;
        }
        const entry = this.zip.getEntry('xl/sharedStrings.bin');
        if (!entry) {
            this.sharedStringsBuf = null;
            return;
        }
        this.sharedStringsBuf = entry.getData();
        const parsed = parseSharedStringsBin(this.sharedStringsBuf);
        this.sharedStringsValues = parsed.values;
        this.sharedStringsTotal = parsed.total;
        this.sharedStringsUnique = parsed.unique;
        this.sharedStringsEndSst = parsed.endSstOffset;

        this.sharedStringsValues.forEach((value, index) => {
            if (!this.stringIndexMap.has(value)) {
                this.stringIndexMap.set(value, index);
            }
        });
        this._appendCount = 0;
    }

    private _commitSharedStrings(): void {
        if (this.sharedStringsBuf === null || this._appendCount === 0) return;

        const originalValuesLength = this.sharedStringsValues.length - this._appendCount;

        let newItems = Buffer.alloc(0);
        for (let i = originalValuesLength; i < this.sharedStringsValues.length; i++) {
            const text = this.sharedStringsValues[i];
            const payload = Buffer.alloc(1 + 4 + text.length * 2);
            payload[0] = 0x00;
            payload.writeUInt32LE(text.length, 1);
            payload.write(text, 5, text.length * 2, 'utf16le');
            newItems = Buffer.concat([newItems, buildRecord(0x0013, payload)]);
        }

        const out = Buffer.alloc(this.sharedStringsBuf.length + newItems.length);
        this.sharedStringsBuf.copy(out, 0, 0, this.sharedStringsEndSst);
        newItems.copy(out, this.sharedStringsEndSst);
        this.sharedStringsBuf.copy(out, this.sharedStringsEndSst + newItems.length, this.sharedStringsEndSst);

        // Patch cstTotal / cstUnique inside BrtBeginSst.
        const rec: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
        let pos = 0;
        while (readRecord(out, pos, rec)) {
            if (rec.id === 0x009f) {
                out.writeUInt32LE(this.sharedStringsTotal, rec.dataStart);
                out.writeUInt32LE(this.sharedStringsUnique, rec.dataStart + 4);
                break;
            }
            pos = rec.dataEnd;
        }

        this.zip.updateFile('xl/sharedStrings.bin', out);
    }

    // ------------------------------------------------------------------
    // Styles
    // ------------------------------------------------------------------

    private _collectStyles(sheetBuf: Buffer): { dataStyles: Map<number, number>; headerStyles: Map<number, number> } {
        const dataCounts = new Map<number, Map<number, number>>();
        const headerStyles = new Map<number, number>();

        const rec: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
        let pos = 0;
        let row = -1;
        while (readRecord(sheetBuf, pos, rec)) {
            pos = rec.dataEnd;
            if (rec.id === 0x00) {
                row = sheetBuf.readInt32LE(rec.dataStart);
                continue;
            }
            if (!CELL_RECORDS.has(rec.id) || rec.len < 8) continue;
            const col = sheetBuf.readUInt32LE(rec.dataStart);
            const xf = sheetBuf.readUInt32LE(rec.dataStart + 4) & 0xffffff;
            if (xf <= 0) continue;

            if (row === 0) {
                if (!headerStyles.has(col)) headerStyles.set(col, xf);
                continue;
            }
            let colCounts = dataCounts.get(col);
            if (!colCounts) {
                colCounts = new Map();
                dataCounts.set(col, colCounts);
            }
            colCounts.set(xf, (colCounts.get(xf) ?? 0) + 1);
        }

        const dataStyles = new Map<number, number>();
        for (const [col, colCounts] of dataCounts) {
            let best = 0;
            let bestCount = -1;
            for (const [xf, count] of colCounts) {
                if (count > bestCount) {
                    best = xf;
                    bestCount = count;
                }
            }
            if (bestCount > 0) dataStyles.set(col, best);
        }

        return { dataStyles, headerStyles };
    }

    private _findDateXf(): number | null {
        const entry = this.zip.getEntry('xl/styles.bin');
        if (!entry) return null;
        const stylesBuf = entry.getData();

        const rec: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
        let pos = 0;
        let inCellXfs = false;
        let xfIndex = 0;
        while (readRecord(stylesBuf, pos, rec)) {
            pos = rec.dataEnd;
            if (rec.id === 0x0269) {
                inCellXfs = true;
                xfIndex = 0;
                continue;
            }
            if (rec.id === 0x026a) {
                inCellXfs = false;
                continue;
            }
            if (inCellXfs && rec.id === 0x002f && rec.len >= 4) {
                const numFmtId = stylesBuf.readUInt16LE(rec.dataStart + 2);
                const isDate =
                    (numFmtId >= 14 && numFmtId <= 22) ||
                    (numFmtId >= 45 && numFmtId <= 47);
                if (isDate) return xfIndex;
                xfIndex++;
            }
        }
        return null;
    }

    // ------------------------------------------------------------------
    // Row / cell serialization
    // ------------------------------------------------------------------

    private _buildRows(
        rows: CellValue[][],
        headers: string[] | undefined,
        dataStyles: Map<number, number>,
        headerStyles: Map<number, number>,
        dateXf: number | null
    ): Buffer {
        const chunks: Buffer[] = [];
        const maxCol = Math.max(rows.reduce((m, r) => Math.max(m, r.length), 0), headers?.length ?? 0, 0);
        const colLast = Math.max(0, maxCol - 1);

        const pushRow = (rowNum: number, cellBytes: Buffer[]): void => {
            // BrtRowHeader (0x00, len 25): rw(4) + zero(4) + flags(9) + colFirst(4) + colLast(4)
            const header = Buffer.alloc(27);
            header[0] = 0x00;
            header[1] = 25;
            header.writeInt32LE(rowNum, 2);
            header[10] = 0x2c;
            header[11] = 0x01;
            header[15] = 0x01;
            header.writeInt32LE(0, 19);
            header.writeInt32LE(colLast, 23);
            chunks.push(header);

            for (const cell of cellBytes) {
                chunks.push(cell);
            }
        };

        if (headers) {
            const cells: Buffer[] = [];
            for (let c = 0; c < headers.length; c++) {
                const style = headerStyles.get(c) ?? 0;
                cells.push(this._stringCellBytes(c, style, headers[c]));
            }
            pushRow(0, cells);
        }

        let rowNum = headers ? 1 : 0;
        for (const row of rows) {
            const cells: Buffer[] = [];
            for (let c = 0; c < row.length; c++) {
                const raw = row[c];
                if (raw === null || raw === undefined) continue;
                cells.push(this._valueCellBytes(raw, c, dataStyles, dateXf));
            }
            pushRow(rowNum, cells);
            rowNum++;
        }

        return Buffer.concat(chunks);
    }

    private _valueCellBytes(
        raw: CellValue,
        col: number,
        dataStyles: Map<number, number>,
        dateXf: number | null
    ): Buffer {
        const fmtString = getFormat(raw);
        const val = fmtString !== null ? unwrapCell(raw) : raw;
        const colStyle = dataStyles.get(col) ?? 0;

        if (typeof val === 'number') {
            if (Number.isFinite(val) && Number.isInteger(val) && val >= RK_INT_LOWER && val <= RK_INT_UPPER) {
                const payload = Buffer.alloc(12);
                payload.writeUInt32LE(col, 0);
                payload.writeUInt32LE(colStyle, 4);
                payload.writeInt32LE((val << 2) | 2, 8);
                return buildRecord(0x0002, payload);
            }
            if (Number.isFinite(val)) {
                const payload = Buffer.alloc(16);
                payload.writeUInt32LE(col, 0);
                payload.writeUInt32LE(colStyle, 4);
                payload.writeDoubleLE(val, 8);
                return buildRecord(0x0005, payload);
            }
            return this._stringCellBytes(col, colStyle, val.toString());
        }
        if (typeof val === 'bigint') {
            return this._stringCellBytes(col, colStyle, val.toString());
        }
        if (typeof val === 'boolean') {
            const payload = Buffer.alloc(9);
            payload.writeUInt32LE(col, 0);
            payload.writeUInt32LE(colStyle, 4);
            payload[8] = val ? 1 : 0;
            return buildRecord(0x0004, payload);
        }
        if (val instanceof Date) {
            const oaDate = (val.getTime() - this._oaEpoch) / 86400000;
            if (Number.isFinite(oaDate)) {
                const style = colStyle > 0 ? colStyle : (dateXf !== null ? dateXf : 0);
                const payload = Buffer.alloc(16);
                payload.writeUInt32LE(col, 0);
                payload.writeUInt32LE(style, 4);
                payload.writeDoubleLE(oaDate, 8);
                return buildRecord(0x0005, payload);
            }
            return this._stringCellBytes(col, colStyle, val.toString());
        }
        if (val !== null && val !== undefined) {
            return this._stringCellBytes(col, colStyle, String(val));
        }
        return this._stringCellBytes(col, colStyle, '');
    }

    private _stringCellBytes(col: number, style: number, text: string): Buffer {
        if (this.sharedStringsBuf === null) {
            // Inline string (BrtCellSt): col(4) + xf(4) + cch(4) + utf16.
            const payload = Buffer.alloc(12 + text.length * 2);
            payload.writeUInt32LE(col, 0);
            payload.writeUInt32LE(style, 4);
            payload.writeUInt32LE(text.length, 8);
            payload.write(text, 12, text.length * 2, 'utf16le');
            return buildRecord(0x0006, payload);
        }

        let index = this.stringIndexMap.get(text);
        if (index === undefined) {
            index = this.sharedStringsValues.length;
            this.sharedStringsValues.push(text);
            this.stringIndexMap.set(text, index);
            this.sharedStringsUnique++;
            this._appendCount++;
        }
        this.sharedStringsTotal++;

        const payload = Buffer.alloc(12);
        payload.writeUInt32LE(col, 0);
        payload.writeUInt32LE(style, 4);
        payload.writeUInt32LE(index, 8);
        return buildRecord(0x0007, payload);
    }

    // ------------------------------------------------------------------
    // Rows region detection
    // ------------------------------------------------------------------

    private _findRowsRegion(
        sheetBuf: Buffer
    ): { rowsStart: number; rowsEnd: number; autoFilterRanges: number[] } {
        const autoFilterRanges: number[] = [];
        let rowsStart = -1;
        let lastCellEnd = -1;
        let lastW25 = -1;

        const rec: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
        let pos = 0;
        while (readRecord(sheetBuf, pos, rec)) {
            if (rec.id === 0x00) {
                if (rowsStart === -1) rowsStart = rec.headerStart;
            } else if (rec.id >= 0x01 && rec.id <= 0x0b) {
                lastCellEnd = rec.dataEnd;
            } else if (rec.id === 0x25 && rec.len === 6) {
                lastW25 = rec.headerStart;
            } else if (rec.id === 0x00a1 && rec.len >= 16) {
                autoFilterRanges.push(rec.dataStart);
            }
            pos = rec.dataEnd;
        }

        if (rowsStart === -1) {
            // Empty sheet: insert before the trailing wrapper (last 0x25) or at the end.
            const insertAt = lastW25 >= 0 ? lastW25 : sheetBuf.length;
            return { rowsStart: insertAt, rowsEnd: insertAt, autoFilterRanges };
        }

        // Replace only the row records. The row-block terminator records that Excel
        // writes after the last row (0x92/0x217/0x1dd/0x1dc) must be preserved,
        // otherwise Excel drops the last written row.
        const rowsEnd = lastCellEnd >= 0 ? lastCellEnd : (lastW25 >= 0 ? lastW25 : sheetBuf.length);

        return { rowsStart, rowsEnd, autoFilterRanges };
    }

    /**
     * Keep the sheet dimension record (0x98) in sync with the new data.
     * Excel uses its rwLast (payload offset 24) and colLast (payload offset 32)
     * fields to read the full row range; stale values make the last written row
     * invisible to Excel (and drop it from pivot refreshes).
     */
    private _patchDimension(sheetBuf: Buffer, lastRowIdx: number, lastCol: number): void {
        const rec: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
        let pos = 0;
        while (readRecord(sheetBuf, pos, rec)) {
            if (rec.id === 0x0098 && rec.len >= 36) {
                sheetBuf.writeInt32LE(lastRowIdx, rec.dataStart + 24);
                sheetBuf.writeInt32LE(lastCol, rec.dataStart + 32);
                return;
            }
            pos = rec.dataEnd;
        }
    }

    // ------------------------------------------------------------------
    // Pivot cache integration
    // ------------------------------------------------------------------

    private _patchPivotCaches(sheetName: string, rwLast: number, colLast: number): void {
        for (const path of this.pivotCacheDefPaths) {
            const entry = this.zip.getEntry(path);
            if (!entry) continue;
            const buf = Buffer.from(entry.getData());

            const rec: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
            let pos = 0;
            let patched = false;
            let refreshXfPatched = false;

            while (readRecord(buf, pos, rec)) {
                pos = rec.dataEnd;

                if (rec.id === 0x00b3 && rec.len >= 4 && !refreshXfPatched) {
                    // BrtBeginPivotCacheDefinition flags: bit 0x04 of payload offset 3 = refreshOnLoad.
                    if (rec.dataStart + 4 <= buf.length) {
                        const v = buf[rec.dataStart + 3];
                        if ((v & 0x04) === 0) {
                            buf[rec.dataStart + 3] = v | 0x04;
                            patched = true;
                        }
                    }
                    refreshXfPatched = true;
                }

                if (rec.id === 0x00bb && rec.len >= 23) {
                    // BrtPivotCacheSource: [3] + cch(4) + name(utf16) + rwFirst(4) rwLast(4) colFirst(4) colLast(4)
                    const cch = buf.readUInt32LE(rec.dataStart + 3);
                    const nameStart = rec.dataStart + 7;
                    const refStart = nameStart + cch * 2;
                    const name = readUtf16(buf, nameStart, cch);
                    if (name === sheetName && refStart + 16 <= buf.length) {
                        buf.writeInt32LE(rwLast, refStart + 4);
                        buf.writeInt32LE(colLast, refStart + 12);
                        patched = true;
                    }
                }
            }

            if (patched) {
                this.zip.updateFile(path, buf);
            }
        }
    }
}

export default XlsbUpdater;
