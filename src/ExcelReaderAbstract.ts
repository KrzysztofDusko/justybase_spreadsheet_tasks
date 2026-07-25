import { CellValue } from './Formats';

/**
 * Shared reader surface for XLSB and XLSX.
 * Always use `await reader.read()` — both formats are async in v2+.
 */
export abstract class ExcelReaderAbstract {
    fieldCount: number = 0;
    rowCount: number = 0;
    actualSheetName: string = '';
    resultsCount: number = 0;

    protected _oaEpoch: number;

    constructor() {
        this._oaEpoch = Date.UTC(1899, 11, 30);
    }

    /** Open a workbook file. */
    abstract open(path: string, readSharedStrings?: boolean, updateMode?: boolean): Promise<void>;

    async close(): Promise<void> {
        // cleanup if needed
    }

    /**
     * Advance to the next row.
     * @returns true if a row is available; false at EOF
     */
    abstract read(): Promise<boolean>;

    /** Worksheet names in workbook order. */
    abstract getSheetNames(): string[];

    /** Cell value for column index `i` of the current row (0-based). */
    abstract getValue(i: number): CellValue;

    dispose(): void {
        // cleanup
    }

    getDateTimeFromOaDate(oaDate: number): Date {
        const ms = oaDate * 86400000 + this._oaEpoch;
        return new Date(ms);
    }
}

export default ExcelReaderAbstract;
