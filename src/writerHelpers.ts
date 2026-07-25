import { isFormattedCell } from './Formats';

/** Excel sheet name characters that must be replaced. */
export const INVALID_SHEET_NAME_CHARS = /[\\/*?[\]:]/g;

/**
 * Sanitize an Excel sheet name (invalid chars, max 31 chars, non-empty).
 * @param sheetCount Current sheet count before adding this sheet (used for defaults).
 */
export function sanitizeSheetName(name: string, sheetCount: number): string {
    if (!name || typeof name !== 'string') {
        return `Sheet${sheetCount + 1}`;
    }

    let sanitized = name.replace(INVALID_SHEET_NAME_CHARS, '_');

    if (sanitized.length > 31) {
        sanitized = sanitized.substring(0, 31);
    }

    if (sanitized.trim().length === 0) {
        sanitized = `Sheet${sheetCount + 1}`;
    }

    return sanitized;
}

/** Create column-width scratch array filled with -1 (unknown). */
export function initColWidths(columnCount: number): number[] {
    return new Array(columnCount).fill(-1.0);
}

/** Apply autofit widths from header labels. */
export function applyHeaderWidths(colWidths: number[], headers: string[], columnCount: number): void {
    for (let i = 0; i < columnCount; i++) {
        const len = headers[i] ? headers[i].length + 1 : 0;
        let width = 1.25 * len + 2;
        if (width > 80) width = 80;
        if (colWidths[i] < width) colWidths[i] = width;
    }
}

/**
 * Sample up to `sampleLimit` rows and expand column widths (autofit).
 * Skips null/undefined cells; Dates use a fixed display length of 20.
 */
export function updateColWidthsFromRows(
    colWidths: number[],
    rows: unknown[][],
    sampleLimit: number = 100
): void {
    for (let r = 0; r < Math.min(rows.length, sampleLimit); r++) {
        const row = rows[r];
        for (let c = 0; c < row.length; c++) {
            const raw = row[c];
            if (raw === null || raw === undefined) continue;
            const val = isFormattedCell(raw) ? raw.value : raw;
            const len = val instanceof Date ? 20 : String(val).length + 1;
            let width = 1.25 * len + 2;
            if (width > 80) width = 80;
            if (colWidths[c] < width) colWidths[c] = width;
        }
    }
}

/**
 * Resolve a stored autofit width to the value written into the file.
 * XLSB uses floored integers; XLSX keeps the float.
 */
export function defaultColWidth(width: number, floor: boolean = false): number {
    if (width > 0) {
        return floor ? Math.floor(width) : width;
    }
    return 10;
}
