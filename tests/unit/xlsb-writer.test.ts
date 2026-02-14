/**
 * XlsbWriter Unit Tests
 */

import { XlsbWriter } from '../../src/XlsbWriter';
import { BigBuffer } from '../../src/BigBuffer';
import { Buffer } from 'buffer';
import * as path from 'path';
import * as os from 'os';

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

// ==================== RK NUMBER TESTS ====================

function testWriteRkNumberInteger(): void {
    console.log('\n--- Write RK Number Integer Tests ---\n');

    const tempFile = path.join(os.tmpdir(), `test_rk_int_${Date.now()}.xlsb`);
    const writer = new XlsbWriter(tempFile);

    // Test writeRkNumberInteger
    {
        const bb = new BigBuffer();
        (writer as any).writeRkNumberInteger(bb, 42, 5);

        const chunks = bb.getChunks();
        const buf = Buffer.concat(chunks);

        assertEqual(buf.length, 14, 'RK Integer: buffer length');
        assertEqual(buf[0], 2, 'RK Integer: record ID');
        assertEqual(buf[1], 12, 'RK Integer: record length');
        assertEqual(buf.readInt32LE(2), 5, 'RK Integer: column');

        const rkVal = buf.readInt32LE(10);
        assertEqual(rkVal, 170, 'RK Integer: encoded value'); // (42 << 2) | 2
    }
}

function testWriteRkNumberNegative(): void {
    console.log('\n--- Write RK Number Negative Tests ---\n');

    const tempFile = path.join(os.tmpdir(), `test_rk_neg_${Date.now()}.xlsb`);
    const writer = new XlsbWriter(tempFile);

    {
        const bb = new BigBuffer();
        (writer as any).writeRkNumberInteger(bb, -100, 0);

        const chunks = bb.getChunks();
        const buf = Buffer.concat(chunks);

        assertEqual(buf[0], 2, 'RK Negative: record ID');
        assertEqual(buf.readInt32LE(2), 0, 'RK Negative: column');
    }
}

// ==================== DOUBLE TESTS ====================

function testWriteDouble(): void {
    console.log('\n--- Write Double Tests ---\n');

    const tempFile = path.join(os.tmpdir(), `test_double_${Date.now()}.xlsb`);
    const writer = new XlsbWriter(tempFile);

    {
        const bb = new BigBuffer();
        (writer as any).writeDouble(bb, 123.456, 3);

        const chunks = bb.getChunks();
        const buf = Buffer.concat(chunks);

        assertEqual(buf[0], 5, 'Double: record ID');
        assertEqual(buf[1], 16, 'Double: record length');
        assertEqual(buf.readInt32LE(2), 3, 'Double: column');

        const val = buf.readDoubleLE(10);
        assert(Math.abs(val - 123.456) < 0.0001, 'Double: value');
    }
}

function testWriteDoubleSpecial(): void {
    console.log('\n--- Write Double Special Values Tests ---\n');

    const tempFile = path.join(os.tmpdir(), `test_double_spec_${Date.now()}.xlsb`);
    const writer = new XlsbWriter(tempFile);

    // Zero
    {
        const bb = new BigBuffer();
        (writer as any).writeDouble(bb, 0, 0);
        const buf = Buffer.concat(bb.getChunks());
        assertEqual(buf.readDoubleLE(10), 0, 'Double: zero');
    }

    // Negative
    {
        const bb = new BigBuffer();
        (writer as any).writeDouble(bb, -999.999, 0);
        const buf = Buffer.concat(bb.getChunks());
        const val = buf.readDoubleLE(10);
        assert(Math.abs(val - (-999.999)) < 0.001, 'Double: negative');
    }

    // Very small
    {
        const bb = new BigBuffer();
        (writer as any).writeDouble(bb, 1e-100, 0);
        const buf = Buffer.concat(bb.getChunks());
        const val = buf.readDoubleLE(10);
        assert(val < 1e-99, 'Double: very small');
    }
}

// ==================== BOOLEAN TESTS ====================

function testWriteBool(): void {
    console.log('\n--- Write Bool Tests ---\n');

    const tempFile = path.join(os.tmpdir(), `test_bool_${Date.now()}.xlsb`);
    const writer = new XlsbWriter(tempFile);

    // True
    {
        const bb = new BigBuffer();
        (writer as any).writeBool(bb, true, 7);

        const chunks = bb.getChunks();
        const buf = Buffer.concat(chunks);

        assertEqual(buf[0], 4, 'Bool: record ID');
        assertEqual(buf[1], 9, 'Bool: record length');
        assertEqual(buf.readInt32LE(2), 7, 'Bool: column');
        assertEqual(buf[10], 1, 'Bool: true value');
    }

    // False
    {
        const bb = new BigBuffer();
        (writer as any).writeBool(bb, false, 0);

        const buf = Buffer.concat(bb.getChunks());
        assertEqual(buf[10], 0, 'Bool: false value');
    }
}

// ==================== DATETIME TESTS ====================

function testWriteDateTime(): void {
    console.log('\n--- Write DateTime Tests ---\n');

    const tempFile = path.join(os.tmpdir(), `test_date_${Date.now()}.xlsb`);
    const writer = new XlsbWriter(tempFile);

    {
        const bb = new BigBuffer();
        const date = new Date('2024-01-01T00:00:00Z');

        (writer as any).writeDateTime(bb, date, 0);
        const chunks = bb.getChunks();
        const buf = Buffer.concat(chunks);

        assertEqual(buf[0], 5, 'DateTime: uses Double record');
        const val = buf.readDoubleLE(10);
        assert(val > 40000, 'DateTime: value reasonable (Excel date)');
        assert(val < 50000, 'DateTime: value reasonable (Excel date)');
    }
}

function testWriteDateTimeEpoch(): void {
    console.log('\n--- Write DateTime Epoch Tests ---\n');

    const tempFile = path.join(os.tmpdir(), `test_date_epoch_${Date.now()}.xlsb`);
    const writer = new XlsbWriter(tempFile);

    // Different dates
    const dates = [
        new Date('1900-01-01'),
        new Date('1970-01-01'),
        new Date('2000-01-01'),
        new Date('2024-12-31'),
    ];

    for (const date of dates) {
        const bb = new BigBuffer();
        (writer as any).writeDateTime(bb, date, 0);
        const buf = Buffer.concat(bb.getChunks());
        const val = buf.readDoubleLE(10);
        assert(val > 0, `DateTime: ${date.toISOString()} has positive Excel value`);
    }
}

// ==================== STRING TESTS ====================

function testWriteString(): void {
    console.log('\n--- Write String Tests ---\n');

    const tempFile = path.join(os.tmpdir(), `test_str_${Date.now()}.xlsb`);
    const writer = new XlsbWriter(tempFile);

    {
        const bb = new BigBuffer();
        const str = "UniqueStringTests";

        (writer as any).writeString(bb, str, 2);

        const chunks = bb.getChunks();
        const buf = Buffer.concat(chunks);

        assertEqual(buf[0], 7, 'String: record ID');

        const sstDic = (writer as any).sstDic as Map<string, number>;
        assert(sstDic.has(str), 'String: SST has string');
        assertEqual(sstDic.get(str), 0, 'String: SST index correct');

        const idx = buf.readInt32LE(10);
        assertEqual(idx, 0, 'String: index in buffer');
    }
}

function testWriteStringDeduplication(): void {
    console.log('\n--- Write String Deduplication Tests ---\n');

    const tempFile = path.join(os.tmpdir(), `test_str_dup_${Date.now()}.xlsb`);
    const writer = new XlsbWriter(tempFile);

    const str = "RepeatedString";

    // Write same string multiple times
    for (let i = 0; i < 5; i++) {
        const bb = new BigBuffer();
        (writer as any).writeString(bb, str, i);
    }

    const sstDic = (writer as any).sstDic as Map<string, number>;
    assertEqual(sstDic.size, 1, 'String dedup: only one unique string');
    assertEqual(sstDic.get(str), 0, 'String dedup: index is 0');
}

function testWriteStringUnicode(): void {
    console.log('\n--- Write String Unicode Tests ---\n');

    const tempFile = path.join(os.tmpdir(), `test_str_uni_${Date.now()}.xlsb`);
    const writer = new XlsbWriter(tempFile);

    const unicodeStrings = [
        "ąęćśńźół",
        "日本語",
        "中文",
        "🎉🔥💻",
    ];

    for (const str of unicodeStrings) {
        const bb = new BigBuffer();
        (writer as any).writeString(bb, str, 0);

        const sstDic = (writer as any).sstDic as Map<string, number>;
        assert(sstDic.has(str), `String unicode: SST has "${str}"`);
    }
}

// ==================== OA EPOCH TESTS ====================

function testOaEpoch(): void {
    console.log('\n--- OA Epoch Tests ---\n');

    const tempFile = path.join(os.tmpdir(), `test_epoch_${Date.now()}.xlsb`);
    const writer = new XlsbWriter(tempFile);

    const epoch = (writer as any)._oaEpoch;
    const expectedEpoch = Date.UTC(1899, 11, 30);

    assertEqual(epoch, expectedEpoch, 'OA Epoch: correct value (Dec 30, 1899)');
}

// ==================== SHEET NAME SANITIZATION TESTS ====================

function testSheetNameSanitization(): void {
    console.log('\n--- Sheet Name Sanitization Tests ---\n');

    const tempFile = path.join(os.tmpdir(), `test_sheet_name_${Date.now()}.xlsb`);
    const writer = new XlsbWriter(tempFile);

    // Access private method
    const sanitize = (writer as any)._sanitizeSheetName.bind(writer);

    // Invalid characters
    assertEqual(sanitize('Sheet/Test'), 'Sheet_Test', 'Sheet name: slash');
    assertEqual(sanitize('Sheet*Test'), 'Sheet_Test', 'Sheet name: asterisk');
    assertEqual(sanitize('Sheet?Test'), 'Sheet_Test', 'Sheet name: question');
    assertEqual(sanitize('Sheet[Test]'), 'Sheet_Test_', 'Sheet name: brackets');
    assertEqual(sanitize('Sheet:Test'), 'Sheet_Test', 'Sheet name: colon');
    assertEqual(sanitize('Sheet\\Test'), 'Sheet_Test', 'Sheet name: backslash');

    // Empty name
    assert(sanitize('') !== '', 'Sheet name: empty gets default');

    // Long name
    const longName = 'A'.repeat(50);
    const sanitized = sanitize(longName);
    assertEqual(sanitized.length, 31, 'Sheet name: truncated to 31 chars');
}

async function runXlsbWriterUnitTests(): Promise<void> {
    console.log('='.repeat(60));
    console.log('XlsbWriter Unit Tests');
    console.log('='.repeat(60));

    testWriteRkNumberInteger();
    testWriteRkNumberNegative();
    testWriteDouble();
    testWriteDoubleSpecial();
    testWriteBool();
    testWriteDateTime();
    testWriteDateTimeEpoch();
    testWriteString();
    testWriteStringDeduplication();
    testWriteStringUnicode();
    testOaEpoch();
    testSheetNameSanitization();

    console.log('\n' + '='.repeat(60));
    const passed = results.filter(r => r.passed).length;
    const failed = results.filter(r => !r.passed).length;
    console.log(`XlsbWriter Unit Tests: ${passed} passed, ${failed} failed`);
    console.log('='.repeat(60));

    if (failed > 0) {
        process.exit(1);
    }
}

runXlsbWriterUnitTests().catch(console.error);
