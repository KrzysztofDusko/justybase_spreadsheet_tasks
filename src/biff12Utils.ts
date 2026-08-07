/**
 * Low-level BIFF12 (XLSB) record helpers shared by readers and the XlsbUpdater.
 *
 * A BIFF12 record is: VLQ(id) + VLQ(length) + payload[length].
 */

/** Decode a 7-bit VLQ integer at `pos`. Returns the value and the next byte offset. */
export function readVlq(buf: Buffer, pos: number): { value: number; nextPos: number } {
    let value = 0;
    let shift = 0;
    let byte: number;
    do {
        byte = buf[pos++];
        value |= (byte & 0x7f) << shift;
        shift += 7;
    } while ((byte & 0x80) !== 0);
    return { value, nextPos: pos };
}

/** Number of bytes needed to VLQ-encode `value`. */
export function vlqLength(value: number): number {
    let n = 1;
    while (value >= 0x80) {
        n++;
        value >>>= 7;
    }
    return n;
}

/** Encode `value` as a 7-bit VLQ byte array. */
export function vlqBytes(value: number): number[] {
    const out: number[] = [];
    while (value >= 0x80) {
        out.push((value & 0x7f) | 0x80);
        value >>>= 7;
    }
    out.push(value & 0x7f);
    return out;
}

/** A single BIFF12 record view (byte offsets into the source buffer). */
export interface Biff12Record {
    /** Byte offset of the record header (id + length). */
    headerStart: number;
    /** Byte offset just past the header — the payload starts here. */
    dataStart: number;
    /** Byte offset just past the payload. */
    dataEnd: number;
    id: number;
    len: number;
}

/**
 * Fill `rec` with the record starting at `pos`. Returns false when the buffer
 * is exhausted or the record would run past the end.
 */
export function readRecord(buf: Buffer, pos: number, rec: Biff12Record): boolean {
    const idRes = readVlq(buf, pos);
    if (idRes.nextPos > buf.length) return false;
    const lenRes = readVlq(buf, idRes.nextPos);
    if (lenRes.nextPos + lenRes.value > buf.length) return false;

    rec.headerStart = pos;
    rec.dataStart = lenRes.nextPos;
    rec.dataEnd = lenRes.nextPos + lenRes.value;
    rec.id = idRes.value;
    rec.len = lenRes.value;
    return true;
}

/** Build a BIFF12 record from an id and payload buffer. */
export function buildRecord(id: number, payload: Buffer): Buffer {
    const idBytes = vlqBytes(id);
    const lenBytes = vlqBytes(payload.length);
    const header = Buffer.alloc(idBytes.length + lenBytes.length);
    idBytes.forEach((b, i) => (header[i] = b));
    lenBytes.forEach((b, i) => (header[idBytes.length + i] = b));
    return Buffer.concat([header, payload]);
}

/** Read a UTF-16LE string of `charCount` characters starting at `start`. */
export function readUtf16(buf: Buffer, start: number, charCount: number): string {
    const end = start + charCount * 2;
    if (end > buf.length) return '';
    return buf.toString('utf16le', start, end);
}

/**
 * Parse the contents of xl/sharedStrings.bin.
 * Returns the string values in order, the SST counters, and the byte offset of
 * the BrtEndSst record where new items must be inserted.
 */
export function parseSharedStringsBin(buf: Buffer): {
    values: string[];
    total: number;
    unique: number;
    endSstOffset: number;
} {
    const values: string[] = [];
    let total = 0;
    let unique = 0;
    let endSstOffset = buf.length;

    const rec: Biff12Record = { headerStart: 0, dataStart: 0, dataEnd: 0, id: 0, len: 0 };
    let pos = 0;
    while (readRecord(buf, pos, rec)) {
        pos = rec.dataEnd;
        if (rec.id === 0x009f) {
            total = buf.readUInt32LE(rec.dataStart);
            unique = buf.readUInt32LE(rec.dataStart + 4);
        } else if (rec.id === 0x0013) {
            const cch = buf.readUInt32LE(rec.dataStart + 1);
            values.push(readUtf16(buf, rec.dataStart + 5, cch));
        } else if (rec.id === 0x00a0) {
            endSstOffset = rec.headerStart;
        }
    }
    return { values, total, unique, endSstOffset };
}
