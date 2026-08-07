import * as fs from 'fs';
import AdmZip from 'adm-zip';
import { CellValue, getFormat, unwrapCell } from './Formats';
import {
    escapeXmlText,
    columnIndexToLetter,
    columnLetterToIndex,
    parseSharedStringsXml,
    trimTrailingEmptyRows,
} from './xmlUtils';

/**
 * Options for {@link XlsxUpdater.replaceSheetData}.
 */
export interface ReplaceSheetDataOptions {
    /**
     * Optional header row written into row 1 of the target sheet.
     * When omitted, row 1 is filled with the data rows.
     */
    headers?: string[];
    /**
     * Style strategy for the new cells:
     * - `'inherit'` (default): reuse the dominant cell style of each column from the
     *   existing data (and the original header style for row 1), so date/number
     *   formatting of the sheet is preserved. Falls back to General (0).
     * - `'general'`: all new cells use style 0 (General).
     */
    styleFallback?: 'inherit' | 'general';
}

interface PivotCacheDef {
    path: string;
    xml: string;
}

/**
 * Update the data of a worksheet inside an existing XLSX file without rebuilding
 * the workbook. Everything outside the target sheet's cell data — pivot tables
 * and their caches, other sheets, styles, themes, defined names — is preserved
 * byte-for-byte. This makes it possible to refresh a data sheet that other
 * Excel features (pivot tables, charts, formulas) are wired to.
 *
 * ```ts
 * const updater = new XlsxUpdater('report.xlsx');
 * updater.replaceSheetData('data1', rows, { headers: ['ID', 'NAME', 'AMOUNT'] });
 * updater.save('report_new.xlsx'); // or updater.save() to overwrite in place
 * ```
 */
export class XlsxUpdater {
    private readonly zip: AdmZip;
    private readonly sourcePath: string;

    private sheetNameToPath: Map<string, string> = new Map();
    private pivotCacheDefs: PivotCacheDef[] = [];
    private pivotTableDefPaths: string[] = [];

    private sharedStringsXml: string | null = null;
    private sharedStringsValues: string[] = [];
    private sharedStringsOriginalLength: number = 0;
    private sharedStringsCount: number = 0;
    private stringIndexMap: Map<string, number> = new Map();

    private readonly _oaEpoch: number;

    /**
     * @param path Path to an existing `.xlsx` file.
     */
    constructor(path: string) {
        this.sourcePath = path;
        this._oaEpoch = Date.UTC(1899, 11, 30);

        if (!fs.existsSync(path)) {
            throw new Error(`XlsxUpdater: file not found: ${path}`);
        }

        this.zip = new AdmZip(path);

        if (this.zip.getEntry('xl/workbook.bin')) {
            throw new Error(
                'XlsxUpdater: XLSB files are not supported yet. Convert the workbook to XLSX first.'
            );
        }
        if (!this.zip.getEntry('xl/workbook.xml')) {
            throw new Error('XlsxUpdater: not a valid XLSX workbook (xl/workbook.xml missing).');
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
            throw new Error(`XlsxUpdater: sheet "${sheetName}" not found in the workbook.`);
        }

        const entry = this.zip.getEntry(sheetPath);
        if (!entry) {
            throw new Error(`XlsxUpdater: worksheet part missing for sheet "${sheetName}".`);
        }

        const sheetXml = entry.getData().toString('utf8');
        this._ensureSharedStringsLoaded();

        let colStyles: Map<number, number> = new Map();
        let headerStyles: Map<number, number> = new Map();
        let dateStyle: number | null = null;

        if (styleFallback !== 'general') {
            const collected = this._collectExistingStyles(sheetXml);
            colStyles = collected.dataStyles;
            headerStyles = collected.headerStyles;
            dateStyle = this._findDateStyleIndex();
        }

        const lastCol = Math.max(rows.reduce((m, r) => Math.max(m, r.length), 0), headers?.length ?? 0) - 1;
        const totalRows = rows.length + (headers ? 1 : 0);

        const newSheetData = this._buildSheetData(rows, headers, colStyles, headerStyles, dateStyle);
        const newXml = this._patchSheetXml(sheetXml, newSheetData, totalRows, lastCol);
        this.zip.updateFile(sheetPath, Buffer.from(newXml, 'utf8'));

        if (this.sharedStringsXml !== null) {
            this._commitSharedStrings();
        }

        const dimRef = totalRows > 0 && lastCol >= 0
            ? `A1:${columnIndexToLetter(lastCol)}${totalRows}`
            : 'A1';
        this._updatePivotCaches(sheetName, dimRef, rows.length);
        this._addRefreshOnLoadToPivotTables();
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
        const relsEntry = this.zip.getEntry('xl/_rels/workbook.xml.rels');
        const rIdToTarget: Record<string, string> = {};
        if (relsEntry) {
            const rels = relsEntry.getData().toString('utf8');
            const relRegex = /<Relationship[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"/g;
            let match;
            while ((match = relRegex.exec(rels)) !== null) {
                rIdToTarget[match[1]] = match[2];
            }
        }

        const wbEntry = this.zip.getEntry('xl/workbook.xml');
        const wbXml = wbEntry ? wbEntry.getData().toString('utf8') : '';
        const sheetRe = /<sheet\b[^>]*>/g;
        let match;
        while ((match = sheetRe.exec(wbXml)) !== null) {
            const tag = match[0];
            const nameMatch = /name="([^"]*)"/.exec(tag);
            const rIdMatch = /r:id="([^"]*)"/.exec(tag);
            if (!nameMatch || !rIdMatch) continue;
            const name = nameMatch[1]
                .replace(/&amp;/g, '&')
                .replace(/&lt;/g, '<')
                .replace(/&gt;/g, '>')
                .replace(/&quot;/g, '"')
                .replace(/&apos;/g, "'");
            const target = rIdToTarget[rIdMatch[1]] ?? '';
            let fullPath = target;
            if (fullPath.startsWith('/')) {
                fullPath = fullPath.substring(1);
            } else if (!fullPath.startsWith('xl/')) {
                fullPath = 'xl/' + fullPath;
            }
            if (fullPath) {
                this.sheetNameToPath.set(name, fullPath);
            }
        }

        for (const entry of this.zip.getEntries()) {
            const name = entry.entryName;
            if (/^xl\/pivotCache\/pivotCacheDefinition\d*\.xml$/.test(name)) {
                this.pivotCacheDefs.push({ path: name, xml: entry.getData().toString('utf8') });
            } else if (/^xl\/pivotTables\/pivotTable\d*\.xml$/.test(name)) {
                this.pivotTableDefPaths.push(name);
            }
        }
    }

    // ------------------------------------------------------------------
    // Shared strings
    // ------------------------------------------------------------------

    private _ensureSharedStringsLoaded(): void {
        if (this.stringIndexMap.size > 0 || this.sharedStringsXml !== null) {
            return;
        }
        const entry = this.zip.getEntry('xl/sharedStrings.xml');
        if (!entry) {
            this.sharedStringsXml = null;
            return;
        }
        this.sharedStringsXml = entry.getData().toString('utf8');
        this.sharedStringsValues = parseSharedStringsXml(this.sharedStringsXml);

        const countMatch = /<sst\b[^>]*>/.exec(this.sharedStringsXml);
        const countAttr = countMatch ? /count="(\d+)"/.exec(countMatch[0]) : null;
        this.sharedStringsCount = countAttr ? parseInt(countAttr[1], 10) : 0;

        this.sharedStringsValues.forEach((value, index) => {
            if (!this.stringIndexMap.has(value)) {
                this.stringIndexMap.set(value, index);
            }
        });
        this.sharedStringsOriginalLength = this.sharedStringsValues.length;
    }

    private _commitSharedStrings(): void {
        if (this.sharedStringsXml === null || this.sharedStringsValues.length === this.sharedStringsOriginalLength) {
            return;
        }

        let xml = this.sharedStringsXml;

        const sstStart = xml.indexOf('<sst');
        const sstTagEnd = xml.indexOf('>', sstStart);
        if (sstStart !== -1 && sstTagEnd !== -1) {
            const tag = xml.substring(sstStart, sstTagEnd + 1);
            const newTag = tag
                .replace(/count="\d+"/, `count="${this.sharedStringsCount}"`)
                .replace(/uniqueCount="\d+"/, `uniqueCount="${this.sharedStringsValues.length}"`);
            xml = xml.substring(0, sstStart) + newTag + xml.substring(sstTagEnd + 1);
        }

        const sstEnd = xml.lastIndexOf('</sst>');
        if (sstEnd === -1) {
            return;
        }

        let appended = '';
        for (let i = this.sharedStringsOriginalLength; i < this.sharedStringsValues.length; i++) {
            const txt = this.sharedStringsValues[i];
            const clean = escapeXmlText(txt);
            const preserve = clean.length > 0 && (clean[0] === ' ' || clean[clean.length - 1] === ' ' || /[\t\n\r]/.test(clean));
            appended += preserve
                ? `<si><t xml:space="preserve">${clean}</t></si>`
                : `<si><t>${clean}</t></si>`;
        }

        xml = xml.substring(0, sstEnd) + appended + xml.substring(sstEnd);
        this.zip.updateFile('xl/sharedStrings.xml', Buffer.from(xml, 'utf8'));
    }

    // ------------------------------------------------------------------
    // Styles
    // ------------------------------------------------------------------

    private _collectExistingStyles(sheetXml: string): { dataStyles: Map<number, number>; headerStyles: Map<number, number> } {
        const dataCounts = new Map<number, Map<number, number>>();
        const headerStyles = new Map<number, number>();

        const sdStart = sheetXml.indexOf('<sheetData>');
        if (sdStart === -1) return { dataStyles: new Map(), headerStyles };
        const sdEnd = sheetXml.indexOf('</sheetData>', sdStart);
        if (sdEnd === -1) return { dataStyles: new Map(), headerStyles };

        const region = sheetXml.substring(sdStart, sdEnd);
        const cellRe = /<c\b[^>]*>/g;
        let m;
        while ((m = cellRe.exec(region)) !== null) {
            const tag = m[0];
            const rMatch = /r="([A-Z]+)(\d+)"/.exec(tag);
            if (!rMatch) continue;
            const col = columnLetterToIndex(rMatch[1]);
            const row = parseInt(rMatch[2], 10);
            const sMatch = /s="(\d+)"/.exec(tag);
            if (!sMatch) continue;
            const s = parseInt(sMatch[1], 10);
            if (s <= 0) continue;

            if (row === 1) {
                if (!headerStyles.has(col)) headerStyles.set(col, s);
                continue;
            }
            let colCounts = dataCounts.get(col);
            if (!colCounts) {
                colCounts = new Map();
                dataCounts.set(col, colCounts);
            }
            colCounts.set(s, (colCounts.get(s) ?? 0) + 1);
        }

        const dataStyles = new Map<number, number>();
        for (const [col, colCounts] of dataCounts) {
            let best = 0;
            let bestCount = -1;
            for (const [s, count] of colCounts) {
                if (count > bestCount) {
                    best = s;
                    bestCount = count;
                }
            }
            if (bestCount > 0) dataStyles.set(col, best);
        }

        return { dataStyles, headerStyles };
    }

    private _findDateStyleIndex(): number | null {
        const entry = this.zip.getEntry('xl/styles.xml');
        if (!entry) return null;
        const xml = entry.getData().toString('utf8');

        const customDateFormats = new Set<number>();
        const numFmtRe = /<numFmt\s+[^>]*numFmtId="(\d+)"[^>]*formatCode="([^"]*)"/g;
        let m;
        while ((m = numFmtRe.exec(xml)) !== null) {
            const id = parseInt(m[1], 10);
            const code = m[2].toLowerCase();
            if (code.includes('yy') || code.includes('mm') || code.includes('dd') || code.includes('h:mm')) {
                customDateFormats.add(id);
            }
        }

        const cellXfsStart = xml.indexOf('<cellXfs');
        if (cellXfsStart === -1) return null;
        const cellXfsEnd = xml.indexOf('</cellXfs>', cellXfsStart);
        if (cellXfsEnd === -1) return null;
        const cellXfsContent = xml.substring(cellXfsStart, cellXfsEnd);

        const xfRe = /<xf\b[^>]*>/g;
        let index = 0;
        while ((m = xfRe.exec(cellXfsContent)) !== null) {
            const nfMatch = /numFmtId="(\d+)"/.exec(m[0]);
            const nfId = nfMatch ? parseInt(nfMatch[1], 10) : 0;
            const isDate = (nfId >= 14 && nfId <= 22) ||
                (nfId >= 45 && nfId <= 47) ||
                customDateFormats.has(nfId);
            if (isDate) return index;
            index++;
        }
        return null;
    }

    // ------------------------------------------------------------------
    // Cell / row serialization
    // ------------------------------------------------------------------

    private _buildSheetData(
        rows: CellValue[][],
        headers: string[] | undefined,
        colStyles: Map<number, number>,
        headerStyles: Map<number, number>,
        dateStyle: number | null
    ): string {
        let out = '';
        let rowNum = 0;

        if (headers) {
            rowNum = 1;
            out += '<row r="1">';
            for (let c = 0; c < headers.length; c++) {
                const style = headerStyles.get(c) ?? 0;
                out += this._stringCell(headers[c], c, rowNum, style);
            }
            out += '</row>';
        }

        for (const row of rows) {
            rowNum++;
            out += `<row r="${rowNum}">`;
            for (let c = 0; c < row.length; c++) {
                const raw = row[c];
                if (raw === null || raw === undefined) continue;
                out += this._valueCell(raw, c, rowNum, colStyles, dateStyle);
            }
            out += '</row>';
        }
        return out;
    }

    private _valueCell(
        raw: CellValue,
        col: number,
        rowNum: number,
        colStyles: Map<number, number>,
        dateStyle: number | null
    ): string {
        const colRef = columnIndexToLetter(col);
        const fmtString = getFormat(raw);
        const val = fmtString !== null ? unwrapCell(raw) : raw;
        const colStyle = colStyles.get(col) ?? 0;

        if (typeof val === 'number') {
            if (Number.isFinite(val)) {
                return `<c r="${colRef}${rowNum}"${colStyle > 0 ? ` s="${colStyle}"` : ''}><v>${val}</v></c>`;
            }
            return this._stringCell(val.toString(), col, rowNum, colStyle);
        }
        if (typeof val === 'bigint') {
            return this._stringCell(val.toString(), col, rowNum, colStyle);
        }
        if (typeof val === 'boolean') {
            return `<c r="${colRef}${rowNum}" t="b"${colStyle > 0 ? ` s="${colStyle}"` : ''}><v>${val ? 1 : 0}</v></c>`;
        }
        if (val instanceof Date) {
            const oaDate = (val.getTime() - this._oaEpoch) / 86400000;
            if (Number.isFinite(oaDate)) {
                const style = colStyle > 0 ? colStyle : (dateStyle !== null ? dateStyle : 0);
                return `<c r="${colRef}${rowNum}"${style > 0 ? ` s="${style}"` : ''}><v>${oaDate}</v></c>`;
            }
            return this._stringCell(val.toString(), col, rowNum, colStyle);
        }
        if (val !== null && val !== undefined) {
            return this._stringCell(String(val), col, rowNum, colStyle);
        }
        return '';
    }

    private _stringCell(text: string, col: number, rowNum: number, style: number): string {
        const colRef = columnIndexToLetter(col);
        const styleAttr = style > 0 ? ` s="${style}"` : '';

        if (this.sharedStringsXml === null) {
            const clean = escapeXmlText(text);
            const preserve = clean.length > 0 && (clean[0] === ' ' || clean[clean.length - 1] === ' ' || /[\t\n\r]/.test(clean));
            const tAttr = preserve ? ' xml:space="preserve"' : '';
            return `<c r="${colRef}${rowNum}" t="inlineStr"${styleAttr}><is><t${tAttr}>${clean}</t></is></c>`;
        }

        let index = this.stringIndexMap.get(text);
        if (index === undefined) {
            index = this.sharedStringsValues.length;
            this.sharedStringsValues.push(text);
            this.stringIndexMap.set(text, index);
        }
        this.sharedStringsCount++;
        return `<c r="${colRef}${rowNum}" t="s"${styleAttr}><v>${index}</v></c>`;
    }

    // ------------------------------------------------------------------
    // Worksheet XML patching
    // ------------------------------------------------------------------

    private _patchSheetXml(sheetXml: string, newSheetData: string, totalRows: number, lastCol: number): string {
        let xml = sheetXml;

        const dimRef = totalRows > 0 && lastCol >= 0
            ? `A1:${columnIndexToLetter(lastCol)}${totalRows}`
            : 'A1';

        // Update the dimension element, or insert one if missing.
        const dimRe = /<dimension\b[^>]*>/;
        const dimMatch = dimRe.exec(xml);
        if (dimMatch) {
            xml = xml.substring(0, dimMatch.index) +
                dimMatch[0].replace(/ref="[^"]*"/, `ref="${dimRef}"`) +
                xml.substring(dimMatch.index + dimMatch[0].length);
        } else {
            const rootRe = /<worksheet\b[^>]*>/;
            const rootMatch = rootRe.exec(xml);
            if (rootMatch) {
                const insertAt = rootMatch.index + rootMatch[0].length;
                xml = xml.substring(0, insertAt) + `<dimension ref="${dimRef}"/>` + xml.substring(insertAt);
            }
        }

        // Replace the entire sheetData element content.
        const sdOpen = '<sheetData>';
        const sdStart = xml.indexOf(sdOpen);
        if (sdStart === -1) {
            const selfClosing = /<sheetData\s*\/\s*>/;
            const scMatch = selfClosing.exec(xml);
            if (scMatch) {
                xml = xml.substring(0, scMatch.index) +
                    '<sheetData>' + newSheetData + '</sheetData>' +
                    xml.substring(scMatch.index + scMatch[0].length);
                return xml;
            }
            throw new Error('XlsxUpdater: <sheetData> element not found in worksheet XML.');
        }
        const sdEnd = xml.indexOf('</sheetData>', sdStart);
        if (sdEnd === -1) {
            throw new Error('XlsxUpdater: <sheetData> element not closed in worksheet XML.');
        }
        xml = xml.substring(0, sdStart + sdOpen.length) + newSheetData + xml.substring(sdEnd);

        // Keep the autoFilter range in sync with the new data bounds.
        const afRe = /<autoFilter\b[^>]*>/;
        const afMatch = afRe.exec(xml);
        if (afMatch && totalRows > 0 && lastCol >= 0 && afMatch[0].includes('ref=')) {
            xml = xml.substring(0, afMatch.index) +
                afMatch[0].replace(/ref="[^"]*"/, `ref="${dimRef}"`) +
                xml.substring(afMatch.index + afMatch[0].length);
        }

        return xml;
    }

    // ------------------------------------------------------------------
    // Pivot table integration
    // ------------------------------------------------------------------

    private _updatePivotCaches(sheetName: string, dimRef: string, recordCount: number): void {
        for (const cache of this.pivotCacheDefs) {
            const wsRe = /<worksheetSource\b[^>]*>/;
            const wsMatch = wsRe.exec(cache.xml);
            if (!wsMatch) continue;
            const wsTag = wsMatch[0];
            const sheetAttr = /sheet="([^"]*)"/.exec(wsTag);
            if (!sheetAttr || sheetAttr[1] !== sheetName) continue;

            let xml = cache.xml;
            xml = xml.substring(0, wsMatch.index) +
                wsTag.replace(/ref="[^"]*"/, `ref="${dimRef}"`) +
                xml.substring(wsMatch.index + wsTag.length);

            const rootRe = /<pivotCacheDefinition\b[^>]*>/;
            const rootMatch = rootRe.exec(xml);
            if (rootMatch) {
                let rootTag = rootMatch[0];
                if (/recordCount="\d+"/.test(rootTag)) {
                    rootTag = rootTag.replace(/recordCount="\d+"/, `recordCount="${recordCount}"`);
                } else if (!/\/>$/.test(rootTag)) {
                    rootTag = rootTag.replace(/>$/, ` recordCount="${recordCount}">`);
                }
                xml = xml.substring(0, rootMatch.index) + rootTag + xml.substring(rootMatch.index + rootMatch[0].length);
            }

            this.zip.updateFile(cache.path, Buffer.from(xml, 'utf8'));
        }
    }

    private _addRefreshOnLoadToPivotTables(): void {
        for (const path of this.pivotTableDefPaths) {
            const entry = this.zip.getEntry(path);
            if (!entry) continue;
            let xml = entry.getData().toString('utf8');

            const rootRe = /<pivotTableDefinition\b[^>]*>/;
            const rootMatch = rootRe.exec(xml);
            if (!rootMatch || rootMatch[0].includes('refreshOnLoad')) continue;

            const newTag = rootMatch[0].replace(/(\/?>)$/, ' refreshOnLoad="1"$1');
            xml = xml.substring(0, rootMatch.index) + newTag + xml.substring(rootMatch.index + rootMatch[0].length);
            this.zip.updateFile(path, Buffer.from(xml, 'utf8'));
        }
    }
}

export default XlsxUpdater;
