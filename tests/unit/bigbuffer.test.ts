/**
 * BigBuffer Unit Tests
 */

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

// ==================== WRITE BYTE TESTS ====================

function testWriteByte(): void {
    console.log('\n--- Write Byte Tests ---\n');

    // Single byte
    {
        const bb = new BigBuffer(10);
        bb.writeByte(0xAB);
        const chunks = bb.getChunks();
        assertEqual(chunks.length, 1, 'WriteByte: single chunk');
        assertEqual(chunks[0].length, 1, 'WriteByte: chunk length');
        assertEqual(chunks[0][0], 0xAB, 'WriteByte: value');
    }

    // Zero byte
    {
        const bb = new BigBuffer(10);
        bb.writeByte(0);
        const chunks = bb.getChunks();
        assertEqual(chunks[0][0], 0, 'WriteByte: zero value');
    }

    // Max byte
    {
        const bb = new BigBuffer(10);
        bb.writeByte(255);
        const chunks = bb.getChunks();
        assertEqual(chunks[0][0], 255, 'WriteByte: max value (255)');
    }

    // Multiple bytes
    {
        const bb = new BigBuffer(10);
        bb.writeByte(1);
        bb.writeByte(2);
        bb.writeByte(3);
        const chunks = bb.getChunks();
        assertEqual(chunks[0].length, 3, 'WriteByte: multiple bytes length');
        assertEqual(chunks[0][0], 1, 'WriteByte: first byte');
        assertEqual(chunks[0][1], 2, 'WriteByte: second byte');
        assertEqual(chunks[0][2], 3, 'WriteByte: third byte');
    }
}

function testWriteByteFlush(): void {
    console.log('\n--- Write Byte Flush Tests ---\n');

    // Trigger flush
    {
        const bb = new BigBuffer(5);
        bb.writeByte(1);
        bb.writeByte(2);
        bb.writeByte(3);
        bb.writeByte(4);
        bb.writeByte(5);
        bb.writeByte(6); // Should trigger flush

        const chunks = bb.getChunks();
        assertEqual(chunks.length, 2, 'Flush: chunk count');
        assertEqual(chunks[0].length, 5, 'Flush: first chunk length');
        assertEqual(chunks[1].length, 1, 'Flush: second chunk length');
        assertEqual(chunks[0][4], 5, 'Flush: first chunk last byte');
        assertEqual(chunks[1][0], 6, 'Flush: second chunk first byte');
    }

    // Multiple flushes
    {
        const bb = new BigBuffer(3);
        for (let i = 0; i < 10; i++) {
            bb.writeByte(i);
        }

        const chunks = bb.getChunks();
        assertEqual(chunks.length, 4, 'Multiple flush: chunk count');
        
        const combined = Buffer.concat(chunks);
        assertEqual(combined.length, 10, 'Multiple flush: total length');
        
        for (let i = 0; i < 10; i++) {
            assertEqual(combined[i], i, `Multiple flush: byte ${i}`);
        }
    }
}

// ==================== WRITE INT TESTS ====================

function testWriteInt32LE(): void {
    console.log('\n--- Write Int32LE Tests ---\n');

    // Basic int32
    {
        const bb = new BigBuffer(10);
        bb.writeInt32LE(0x12345678);
        const chunks = bb.getChunks();
        assertEqual(chunks[0].length, 4, 'Int32LE: length');
        assertEqual(chunks[0].readInt32LE(0), 0x12345678, 'Int32LE: value');
    }

    // Zero
    {
        const bb = new BigBuffer(10);
        bb.writeInt32LE(0);
        const chunks = bb.getChunks();
        assertEqual(chunks[0].readInt32LE(0), 0, 'Int32LE: zero');
    }

    // Negative
    {
        const bb = new BigBuffer(10);
        bb.writeInt32LE(-12345);
        const chunks = bb.getChunks();
        assertEqual(chunks[0].readInt32LE(0), -12345, 'Int32LE: negative');
    }

    // Max int32
    {
        const bb = new BigBuffer(10);
        bb.writeInt32LE(2147483647);
        const chunks = bb.getChunks();
        assertEqual(chunks[0].readInt32LE(0), 2147483647, 'Int32LE: max value');
    }

    // Min int32
    {
        const bb = new BigBuffer(10);
        bb.writeInt32LE(-2147483648);
        const chunks = bb.getChunks();
        assertEqual(chunks[0].readInt32LE(0), -2147483648, 'Int32LE: min value');
    }
}

function testWriteUInt32LE(): void {
    console.log('\n--- Write UInt32LE Tests ---\n');

    // Basic uint32 (using int32 for positive values)
    {
        const bb = new BigBuffer(10);
        bb.writeInt32LE(0x12345678);
        const chunks = bb.getChunks();
        assertEqual(chunks[0].readUInt32LE(0), 0x12345678, 'UInt32LE: value');
    }

    // Max uint32 (using int32 buffer)
    {
        const bb = new BigBuffer(10);
        const buf = Buffer.alloc(4);
        buf.writeUInt32LE(0xFFFFFFFF, 0);
        // Write as bytes
        for (let i = 0; i < 4; i++) {
            bb.writeByte(buf[i]);
        }
        const chunks = bb.getChunks();
        assertEqual(chunks[0].readUInt32LE(0), 0xFFFFFFFF, 'UInt32LE: max value');
    }
}

// ==================== WRITE DOUBLE TESTS ====================

function testWriteDoubleLE(): void {
    console.log('\n--- Write DoubleLE Tests ---\n');

    // Basic double
    {
        const bb = new BigBuffer(20);
        const val = 123.456;
        bb.writeDoubleLE(val);
        const chunks = bb.getChunks();
        const readVal = chunks[0].readDoubleLE(0);
        assert(Math.abs(readVal - val) < 0.0001, 'DoubleLE: basic value');
    }

    // Zero
    {
        const bb = new BigBuffer(20);
        bb.writeDoubleLE(0);
        const chunks = bb.getChunks();
        assertEqual(chunks[0].readDoubleLE(0), 0, 'DoubleLE: zero');
    }

    // Negative
    {
        const bb = new BigBuffer(20);
        const val = -987.654;
        bb.writeDoubleLE(val);
        const chunks = bb.getChunks();
        const readVal = chunks[0].readDoubleLE(0);
        assert(Math.abs(readVal - val) < 0.0001, 'DoubleLE: negative');
    }

    // Very small
    {
        const bb = new BigBuffer(20);
        const val = 1e-100;
        bb.writeDoubleLE(val);
        const chunks = bb.getChunks();
        const readVal = chunks[0].readDoubleLE(0);
        assert(Math.abs(readVal - val) < 1e-110, 'DoubleLE: very small');
    }

    // Very large
    {
        const bb = new BigBuffer(20);
        const val = 1e100;
        bb.writeDoubleLE(val);
        const chunks = bb.getChunks();
        const readVal = chunks[0].readDoubleLE(0);
        assert(Math.abs(readVal - val) / val < 1e-10, 'DoubleLE: very large');
    }

    // PI
    {
        const bb = new BigBuffer(20);
        const val = Math.PI;
        bb.writeDoubleLE(val);
        const chunks = bb.getChunks();
        const readVal = chunks[0].readDoubleLE(0);
        assert(Math.abs(readVal - val) < 1e-15, 'DoubleLE: PI');
    }
}

// ==================== WRITE STRING TESTS ====================

function testWriteString(): void {
    console.log('\n--- Write String Tests ---\n');

    // Basic string
    {
        const bb = new BigBuffer(100);
        const str = "Hello World";
        bb.writeString(str);
        const chunks = bb.getChunks();
        const res = chunks[0].toString('utf8');
        assertEqual(res, str, 'String: basic');
    }

    // Empty string
    {
        const bb = new BigBuffer(100);
        bb.writeString("");
        const chunks = bb.getChunks();
        assertEqual(chunks.length, 0, 'String: empty - no chunks');
    }

    // Unicode
    {
        const bb = new BigBuffer(100);
        const str = "ąęćśńźół日本語";
        bb.writeString(str);
        const chunks = bb.getChunks();
        const combined = Buffer.concat(chunks);
        assertEqual(combined.toString('utf8'), str, 'String: unicode');
    }

    // Long string (triggers flush)
    {
        const bb = new BigBuffer(10);
        const str = "This string is definitely longer than 10 bytes";
        bb.writeString(str);
        const chunks = bb.getChunks();
        const combined = Buffer.concat(chunks);
        assertEqual(combined.toString('utf8'), str, 'String: long');
    }

    // Special characters
    {
        const bb = new BigBuffer(100);
        const str = "Line1\nLine2\tTabbed";
        bb.writeString(str);
        const chunks = bb.getChunks();
        const combined = Buffer.concat(chunks);
        assertEqual(combined.toString('utf8'), str, 'String: special chars');
    }
}

function testWriteUtf16LE(): void {
    console.log('\n--- Write UTF16LE Tests ---\n');

    // Basic UTF16
    {
        const bb = new BigBuffer(100);
        const str = "Hello";
        bb.writeUtf16LE(str);
        const chunks = bb.getChunks();
        const combined = Buffer.concat(chunks);
        assertEqual(combined.toString('utf16le'), str, 'UTF16LE: basic');
    }

    // Unicode UTF16
    {
        const bb = new BigBuffer(100);
        const str = "ąęćśńźół日本語🎉";
        bb.writeUtf16LE(str);
        const chunks = bb.getChunks();
        const combined = Buffer.concat(chunks);
        assertEqual(combined.toString('utf16le'), str, 'UTF16LE: unicode');
    }
}

// ==================== GET CHUNKS TESTS ====================

function testGetChunks(): void {
    console.log('\n--- Get Chunks Tests ---\n');

    // Empty buffer
    {
        const bb = new BigBuffer(10);
        const chunks = bb.getChunks();
        assertEqual(chunks.length, 0, 'GetChunks: empty buffer');
    }

    // Single chunk
    {
        const bb = new BigBuffer(100);
        bb.writeByte(1);
        bb.writeByte(2);
        const chunks = bb.getChunks();
        assertEqual(chunks.length, 1, 'GetChunks: single chunk');
    }

    // Multiple chunks
    {
        const bb = new BigBuffer(5);
        for (let i = 0; i < 20; i++) {
            bb.writeByte(i);
        }
        const chunks = bb.getChunks();
        assert(chunks.length >= 2, 'GetChunks: multiple chunks');
    }
}

// ==================== EDGE CASES ====================

function testEdgeCases(): void {
    console.log('\n--- Edge Cases ---\n');

    // Very small chunk size
    {
        const bb = new BigBuffer(1);
        bb.writeByte(1);
        bb.writeByte(2);
        bb.writeByte(3);
        const chunks = bb.getChunks();
        assertEqual(chunks.length, 3, 'Edge: chunk size 1');
    }

    // Large chunk size
    {
        const bb = new BigBuffer(10000);
        bb.writeByte(1);
        const chunks = bb.getChunks();
        assertEqual(chunks.length, 1, 'Edge: large chunk size');
        assertEqual(chunks[0].length, 1, 'Edge: large chunk - only written bytes');
    }

    // Mixed writes
    {
        const bb = new BigBuffer(100);
        bb.writeByte(0x01);
        bb.writeInt32LE(0x12345678);
        bb.writeDoubleLE(3.14);
        bb.writeString("test");
        
        const chunks = bb.getChunks();
        const combined = Buffer.concat(chunks);
        
        assertEqual(combined[0], 0x01, 'Edge: mixed - byte');
        assertEqual(combined.readInt32LE(1), 0x12345678, 'Edge: mixed - int32');
        assert(Math.abs(combined.readDoubleLE(5) - 3.14) < 0.01, 'Edge: mixed - double');
    }
}

async function runBigBufferTests(): Promise<void> {
    console.log('='.repeat(60));
    console.log('BigBuffer Unit Tests');
    console.log('='.repeat(60));

    testWriteByte();
    testWriteByteFlush();
    testWriteInt32LE();
    testWriteUInt32LE();
    testWriteDoubleLE();
    testWriteString();
    testWriteUtf16LE();
    testGetChunks();
    testEdgeCases();

    console.log('\n' + '='.repeat(60));
    const passed = results.filter(r => r.passed).length;
    const failed = results.filter(r => !r.passed).length;
    console.log(`BigBuffer Tests: ${passed} passed, ${failed} failed`);
    console.log('='.repeat(60));

    if (failed > 0) {
        process.exit(1);
    }
}

runBigBufferTests().catch(console.error);
