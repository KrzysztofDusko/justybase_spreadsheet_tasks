/**
 * BiffReaderWriter Unit Tests
 */

import { BiffReaderWriter } from '../../src/BiffReaderWriter';
import { BigBuffer } from '../../src/BigBuffer';
import { Buffer } from 'buffer';

const results: { name: string, passed: boolean, message?: string }[] = [];

function log(msg: string) {
    console.log(msg);
}

function assert(condition: boolean, testName: string, message?: string): void {
    if (condition) {
        results.push({ name: testName, passed: true });
        log(`✅ ${testName}`);
    } else {
        results.push({ name: testName, passed: false, message });
        console.error(`❌ ${testName}${message ? ': ' + message : ''}`);
    }
}

function assertEqual(actual: any, expected: any, testName: string): void {
    const isEqual = JSON.stringify(actual) === JSON.stringify(expected);
    if (isEqual) {
        results.push({ name: testName, passed: true });
        log(`✅ ${testName}`);
    } else {
        const msg = `Expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`;
        results.push({ name: testName, passed: false, message: msg });
        console.error(`❌ ${testName}: ${msg}`);
    }
}

// Helper to write VLQ (Variable Length Quantity)
function writeVLQ(buf: BigBuffer, val: number): void {
    let v = val;
    while (v >= 0x80) {
        buf.writeByte((v & 0x7F) | 0x80);
        v >>>= 7;
    }
    buf.writeByte(v & 0x7F);
}

// ==================== ROW RECORD TESTS ====================

function testRowRecordParsing(): void {
    console.log('\n--- Row Record Parsing Tests ---\n');

    // Basic row record
    {
        const bb = new BigBuffer();
        writeVLQ(bb, 0x00); // Row record type
        writeVLQ(bb, 4);    // Record length
        bb.writeInt32LE(5); // Row index

        const buffer = Buffer.concat(bb.getChunks());
        const parser = new BiffReaderWriter(buffer);

        const success = parser.readWorksheet();
        assert(success, 'Row: parse success');
        assertEqual(parser._rowIndex, 5, 'Row: index value');
    }

    // Row 0
    {
        const bb = new BigBuffer();
        writeVLQ(bb, 0x00);
        writeVLQ(bb, 4);
        bb.writeInt32LE(0);

        const buffer = Buffer.concat(bb.getChunks());
        const parser = new BiffReaderWriter(buffer);
        parser.readWorksheet();
        assertEqual(parser._rowIndex, 0, 'Row: index 0');
    }

    // Large row index
    {
        const bb = new BigBuffer();
        writeVLQ(bb, 0x00);
        writeVLQ(bb, 4);
        bb.writeInt32LE(100000);

        const buffer = Buffer.concat(bb.getChunks());
        const parser = new BiffReaderWriter(buffer);
        parser.readWorksheet();
        assertEqual(parser._rowIndex, 100000, 'Row: large index');
    }
}

// ==================== NUMBER RECORD TESTS ====================

function testNumberRecordParsing(): void {
    console.log('\n--- Number Record Parsing Tests ---\n');

    // Double value (record type 0x05)
    {
        const bb = new BigBuffer();
        writeVLQ(bb, 0x05);    // Number record type
        writeVLQ(bb, 16);      // Record length
        bb.writeInt32LE(2);    // Column
        bb.writeInt32LE(0);    // Unknown
        bb.writeDoubleLE(3.14159);

        const buffer = Buffer.concat(bb.getChunks());
        const parser = new BiffReaderWriter(buffer);
        parser.readWorksheet();

        assertEqual(parser._columnNum, 2, 'Number: column');
        assert(Math.abs(parser._doubleVal - 3.14159) < 0.0001, 'Number: double value');
        assertEqual(parser._cellType, 3, 'Number: cell type');
    }

    // Zero
    {
        const bb = new BigBuffer();
        writeVLQ(bb, 0x05);
        writeVLQ(bb, 16);
        bb.writeInt32LE(0);
        bb.writeInt32LE(0);
        bb.writeDoubleLE(0);

        const buffer = Buffer.concat(bb.getChunks());
        const parser = new BiffReaderWriter(buffer);
        parser.readWorksheet();
        assertEqual(parser._doubleVal, 0, 'Number: zero');
    }

    // Negative
    {
        const bb = new BigBuffer();
        writeVLQ(bb, 0x05);
        writeVLQ(bb, 16);
        bb.writeInt32LE(0);
        bb.writeInt32LE(0);
        bb.writeDoubleLE(-123.456);

        const buffer = Buffer.concat(bb.getChunks());
        const parser = new BiffReaderWriter(buffer);
        parser.readWorksheet();
        assert(Math.abs(parser._doubleVal - (-123.456)) < 0.001, 'Number: negative');
    }

    // Very large
    {
        const bb = new BigBuffer();
        writeVLQ(bb, 0x05);
        writeVLQ(bb, 16);
        bb.writeInt32LE(0);
        bb.writeInt32LE(0);
        bb.writeDoubleLE(1.7976931348623157e+308);

        const buffer = Buffer.concat(bb.getChunks());
        const parser = new BiffReaderWriter(buffer);
        parser.readWorksheet();
        assert(parser._doubleVal > 1e308, 'Number: very large');
    }
}

// ==================== RK NUMBER TESTS ====================

function testRkNumberParsing(): void {
    console.log('\n--- RK Number Parsing Tests ---\n');

    // Integer RK (record type 0x02)
    {
        const bb = new BigBuffer();
        writeVLQ(bb, 0x02);
        writeVLQ(bb, 12);
        bb.writeInt32LE(7);    // Column
        bb.writeInt32LE(0);    // Unknown
        const rkVal = (42 << 2) | 2;  // Integer RK encoding
        bb.writeInt32LE(rkVal);

        const buffer = Buffer.concat(bb.getChunks());
        const parser = new BiffReaderWriter(buffer);
        parser.readWorksheet();

        assertEqual(parser._columnNum, 7, 'RK: column');
        assertEqual(parser._doubleVal, 42, 'RK: integer value');
    }

    // Zero RK
    {
        const bb = new BigBuffer();
        writeVLQ(bb, 0x02);
        writeVLQ(bb, 12);
        bb.writeInt32LE(0);
        bb.writeInt32LE(0);
        bb.writeInt32LE(2);  // Zero with integer flag

        const buffer = Buffer.concat(bb.getChunks());
        const parser = new BiffReaderWriter(buffer);
        parser.readWorksheet();
        assertEqual(parser._doubleVal, 0, 'RK: zero');
    }

    // Negative RK integer
    {
        const bb = new BigBuffer();
        writeVLQ(bb, 0x02);
        writeVLQ(bb, 12);
        bb.writeInt32LE(0);
        bb.writeInt32LE(0);
        const rkVal = (-100 << 2) | 2;
        bb.writeInt32LE(rkVal);

        const buffer = Buffer.concat(bb.getChunks());
        const parser = new BiffReaderWriter(buffer);
        parser.readWorksheet();
        assertEqual(parser._doubleVal, -100, 'RK: negative integer');
    }
}

// ==================== STRING RECORD TESTS ====================

function testStringRecordParsing(): void {
    console.log('\n--- String Record Parsing Tests ---\n');

    // Basic string (record type 0x06)
    {
        const bb = new BigBuffer();
        const testStr = "TestString";

        writeVLQ(bb, 0x06);
        const strBytes = Buffer.byteLength(testStr, 'utf16le');
        const recLen = 12 + strBytes;
        writeVLQ(bb, recLen);

        bb.writeInt32LE(10);        // Column
        bb.writeInt32LE(0);         // Unknown
        bb.writeInt32LE(testStr.length);
        bb.writeUtf16LE(testStr);

        const buffer = Buffer.concat(bb.getChunks());
        const parser = new BiffReaderWriter(buffer);
        parser.readWorksheet();

        assertEqual(parser._columnNum, 10, 'String: column');
        assertEqual(parser._stringValue, testStr, 'String: value');
        assertEqual(parser._cellType, 5, 'String: cell type');
    }

    // Empty string
    {
        const bb = new BigBuffer();
        const testStr = "";

        writeVLQ(bb, 0x06);
        writeVLQ(bb, 12);  // Minimum length
        bb.writeInt32LE(0);
        bb.writeInt32LE(0);
        bb.writeInt32LE(0);  // Zero length

        const buffer = Buffer.concat(bb.getChunks());
        const parser = new BiffReaderWriter(buffer);
        parser.readWorksheet();
        assertEqual(parser._stringValue, "", 'String: empty');
    }

    // Unicode string
    {
        const bb = new BigBuffer();
        const testStr = "ąęćśńźół日本語";

        writeVLQ(bb, 0x06);
        const strBytes = Buffer.byteLength(testStr, 'utf16le');
        const recLen = 12 + strBytes;
        writeVLQ(bb, recLen);

        bb.writeInt32LE(0);
        bb.writeInt32LE(0);
        bb.writeInt32LE(testStr.length);
        bb.writeUtf16LE(testStr);

        const buffer = Buffer.concat(bb.getChunks());
        const parser = new BiffReaderWriter(buffer);
        parser.readWorksheet();
        assertEqual(parser._stringValue, testStr, 'String: unicode');
    }
}

// ==================== BOOLEAN RECORD TESTS ====================

function testBoolRecordParsing(): void {
    console.log('\n--- Boolean Record Parsing Tests ---\n');

    // Boolean true (record type 0x04)
    {
        const bb = new BigBuffer();
        writeVLQ(bb, 0x04);
        writeVLQ(bb, 9);
        bb.writeInt32LE(3);    // Column
        bb.writeInt32LE(0);    // Unknown
        bb.writeByte(1);       // True

        const buffer = Buffer.concat(bb.getChunks());
        const parser = new BiffReaderWriter(buffer);
        parser.readWorksheet();

        assertEqual(parser._columnNum, 3, 'Bool: column');
        assertEqual(parser._boolValue, true, 'Bool: true value');
    }

    // Boolean false
    {
        const bb = new BigBuffer();
        writeVLQ(bb, 0x04);
        writeVLQ(bb, 9);
        bb.writeInt32LE(0);
        bb.writeInt32LE(0);
        bb.writeByte(0);       // False

        const buffer = Buffer.concat(bb.getChunks());
        const parser = new BiffReaderWriter(buffer);
        parser.readWorksheet();
        assertEqual(parser._boolValue, false, 'Bool: false value');
    }
}

// ==================== SHARED STRING INDEX TESTS ====================

function testSharedStringIndexParsing(): void {
    console.log('\n--- Shared String Index Tests ---\n');

    // SST index (record type 0x07) - requires column, xfIndex, and SST index
    {
        const bb = new BigBuffer();
        writeVLQ(bb, 0x07);     // Shared string record type
        writeVLQ(bb, 12);       // Record length (4 + 4 + 4)
        bb.writeInt32LE(5);     // Column
        bb.writeInt32LE(0);     // xfIndex
        bb.writeInt32LE(42);    // SST index

        const buffer = Buffer.concat(bb.getChunks());
        const parser = new BiffReaderWriter(buffer);
        parser.readWorksheet();

        assertEqual(parser._columnNum, 5, 'SST: column');
        assertEqual(parser._intValue, 42, 'SST: index');
        assertEqual(parser._cellType, 2, 'SST: cell type');
    }
}

// ==================== MULTIPLE RECORDS TESTS ====================

function testMultipleRecords(): void {
    console.log('\n--- Multiple Records Tests ---\n');

    // Multiple records in sequence
    {
        const bb = new BigBuffer();

        // Row record
        writeVLQ(bb, 0x00);
        writeVLQ(bb, 4);
        bb.writeInt32LE(0);

        // Number record
        writeVLQ(bb, 0x05);
        writeVLQ(bb, 16);
        bb.writeInt32LE(0);
        bb.writeInt32LE(0);
        bb.writeDoubleLE(100);

        // String record
        const testStr = "Test";
        writeVLQ(bb, 0x06);
        const strBytes = Buffer.byteLength(testStr, 'utf16le');
        writeVLQ(bb, 12 + strBytes);
        bb.writeInt32LE(1);
        bb.writeInt32LE(0);
        bb.writeInt32LE(testStr.length);
        bb.writeUtf16LE(testStr);

        const buffer = Buffer.concat(bb.getChunks());
        const parser = new BiffReaderWriter(buffer);

        // Read row
        assert(parser.readWorksheet(), 'Multi: first read');
        assertEqual(parser._rowIndex, 0, 'Multi: row index');

        // Read number
        assert(parser.readWorksheet(), 'Multi: second read');
        assertEqual(parser._doubleVal, 100, 'Multi: number value');

        // Read string
        assert(parser.readWorksheet(), 'Multi: third read');
        assertEqual(parser._stringValue, testStr, 'Multi: string value');
    }
}

// ==================== VLQ ENCODING TESTS ====================

function testVLQEncoding(): void {
    console.log('\n--- VLQ Encoding Tests ---\n');

    // Small value (single byte)
    {
        const bb = new BigBuffer();
        writeVLQ(bb, 0x7F);
        const chunks = bb.getChunks();
        assertEqual(chunks.length, 1, 'VLQ: single byte chunks');
        assertEqual(chunks[0].length, 1, 'VLQ: single byte length');
        assertEqual(chunks[0][0], 0x7F, 'VLQ: single byte value');
    }

    // Two byte value
    {
        const bb = new BigBuffer();
        writeVLQ(bb, 0x80);
        const chunks = bb.getChunks();
        assertEqual(chunks[0].length, 2, 'VLQ: two byte length');
        assertEqual(chunks[0][0] & 0x80, 0x80, 'VLQ: continuation bit set');
    }

    // Large value
    {
        const bb = new BigBuffer();
        writeVLQ(bb, 0x100000);
        const chunks = bb.getChunks();
        assert(chunks[0].length >= 3, 'VLQ: large value length');
    }
}

async function runBiffReaderWriterTests(): Promise<void> {
    console.log('='.repeat(60));
    console.log('BiffReaderWriter Unit Tests');
    console.log('='.repeat(60));

    testRowRecordParsing();
    testNumberRecordParsing();
    testRkNumberParsing();
    testStringRecordParsing();
    testBoolRecordParsing();
    testSharedStringIndexParsing();
    testMultipleRecords();
    testVLQEncoding();

    console.log('\n' + '='.repeat(60));
    const passed = results.filter(r => r.passed).length;
    const failed = results.filter(r => !r.passed).length;
    console.log(`BiffReaderWriter Tests: ${passed} passed, ${failed} failed`);
    console.log('='.repeat(60));

    if (failed > 0) {
        process.exit(1);
    }
}

runBiffReaderWriterTests().catch(console.error);
