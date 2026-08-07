/**
 * Shared low-level XML helpers used by readers and the XlsxUpdater.
 */

/** True when the string contains XML-invalid control characters (tab/CR/LF are allowed). */
function hasInvalidControlChars(str: string): boolean {
    for (let i = 0; i < str.length; i++) {
        const c = str.charCodeAt(i);
        if (c < 0x20 && c !== 0x09 && c !== 0x0a && c !== 0x0d) {
            return true;
        }
    }
    return false;
}

/** Remove XML-invalid control characters (tab/CR/LF and normal ranges are kept). */
function stripInvalidControlChars(str: string): string {
    let out = '';
    for (let i = 0; i < str.length; i++) {
        const c = str.charCodeAt(i);
        if (c === 0x09 || c === 0x0a || c === 0x0d || (c >= 0x20 && c <= 0xd7ff) || (c >= 0xe000 && c <= 0xfffd)) {
            out += str[i];
        }
    }
    return out;
}

/** Escape text content for use inside an XML element body. */
export function escapeXmlText(str: string): string {
    if (!/&|<|>|"|'/.test(str) && !hasInvalidControlChars(str)) {
        return str;
    }
    let out = str
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
    if (hasInvalidControlChars(out)) {
        out = stripInvalidControlChars(out);
    }
    return out;
}

/** Unescape the five predefined XML entities. */
export function unescapeXml(str: string): string {
    if (!str || str.indexOf('&') === -1) return str;
    return str.replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'");
}

/** 0-based column index -> "A", "B", ..., "AA", "AB", ... (bijective base-26). */
export function columnIndexToLetter(colIndex: number): string {
    let n = colIndex + 1;
    let result = '';
    while (n > 0) {
        const rem = (n - 1) % 26;
        result = String.fromCharCode(65 + rem) + result;
        n = Math.floor((n - 1) / 26);
    }
    return result;
}

/** "A", "AB", ... -> 0-based column index. */
export function columnLetterToIndex(letter: string): number {
    let column = 0;
    const length = letter.length;
    for (let i = 0; i < length; i++) {
        column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
    }
    return column - 1;
}

/**
 * Remove trailing rows that contain no non-empty cells (padding rows sometimes
 * produced by SQL exports). This keeps the written range aligned with the
 * actual data, so pivot tables do not pick up "(blank)" items from empty rows.
 */
export function trimTrailingEmptyRows<T>(rows: T[][]): T[][] {
    let end = rows.length;
    while (end > 0 && rows[end - 1].every(v => v === null || v === undefined)) {
        end--;
    }
    return end === rows.length ? rows : rows.slice(0, end);
}

/**
 * Parse the contents of xl/sharedStrings.xml into the list of string values.
 * Handles plain `<si><t>…</t></si>` entries as well as rich-text `<si><r><t>…</t></r></si>` runs.
 */
export function parseSharedStringsXml(xml: string): string[] {
    const result: string[] = [];
    let pos = 0;
    while (true) {
        const siStart = xml.indexOf('<si>', pos);
        if (siStart === -1) break;

        const siEnd = xml.indexOf('</si>', siStart);
        if (siEnd === -1) break;

        const content = xml.substring(siStart, siEnd);
        let val = '';
        let tPos = 0;
        while (true) {
            const tStart = content.indexOf('<t', tPos);
            if (tStart === -1) break;
            const tagEnd = content.indexOf('>', tStart);
            const tEnd = content.indexOf('</t>', tagEnd);
            if (tEnd === -1) break;

            val += content.substring(tagEnd + 1, tEnd);
            tPos = tEnd + 4;
        }

        result.push(unescapeXml(val));
        pos = siEnd + 5;
    }
    return result;
}
