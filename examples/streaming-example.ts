/**
 * Streaming Example
 * 
 * This example demonstrates how to use the streaming API to write
 * large datasets efficiently without loading all data into memory.
 * 
 * Run with: npx ts-node examples/streaming-example.ts
 */

import { XlsbWriter } from '../dist/cjs';
import * as path from 'path';
import * as fs from 'fs';

async function streamingExample() {
    console.log('=== XlsbWriter Streaming Example ===\n');

    const outputPath = path.join(__dirname, '..', 'output', 'streaming-example.xlsb');
    fs.mkdirSync(path.dirname(outputPath), { recursive: true });
    console.log(`Output file: ${outputPath}\n`);

    const writer = new XlsbWriter(outputPath);

    // Example 1: Generate large dataset without loading all into memory
    console.log('Example 1: Generating 100,000 rows using streaming...');
    const startTime1 = Date.now();
    const initialMemory = process.memoryUsage().heapUsed / 1024 / 1024;

    writer.startSheet('LargeDataset', 5, ['ID', 'Name', 'Value', 'Date', 'Active']);

    for (let i = 0; i < 100_000; i++) {
        writer.writeRow([
            i + 1,
            `User_${i + 1}`,
            Math.random() * 10000,
            new Date(2024, 0, 1 + (i % 365)),
            i % 2 === 0
        ]);

        // Log progress every 20,000 rows
        if ((i + 1) % 20_000 === 0) {
            const currentMemory = process.memoryUsage().heapUsed / 1024 / 1024;
            const memoryDelta = (currentMemory - initialMemory).toFixed(2);
            console.log(`  Processed ${(i + 1).toLocaleString()} rows | Memory delta: ${memoryDelta} MB`);
        }
    }

    writer.endSheet();
    const elapsed1 = Date.now() - startTime1;
    const finalMemory = process.memoryUsage().heapUsed / 1024 / 1024;
    const totalMemoryUsed = (finalMemory - initialMemory).toFixed(2);

    console.log(`✓ Completed in ${elapsed1}ms`);
    console.log(`  Total memory delta: ${totalMemoryUsed} MB\n`);

    // Example 2: Multiple sheets with streaming
    console.log('Example 2: Creating multiple sheets with streaming...');
    const startTime2 = Date.now();

    // Sheet 2: Customer data
    writer.startSheet('Customers', 4, ['CustomerID', 'Email', 'SignupDate', 'Premium']);
    for (let i = 0; i < 10_000; i++) {
        writer.writeRow([
            1000 + i,
            `customer${i}@example.com`,
            new Date(2020 + (i % 5), i % 12, (i % 28) + 1),
            i % 3 === 0
        ]);
    }
    writer.endSheet();

    // Sheet 3: Transactions
    writer.startSheet('Transactions', 3, ['TransactionID', 'Amount', 'Timestamp']);
    for (let i = 0; i < 50_000; i++) {
        writer.writeRow([
            i + 1,
            parseFloat((Math.random() * 5000).toFixed(2)),
            new Date(2024, i % 12, (i % 28) + 1, i % 24, i % 60)
        ]);
    }
    writer.endSheet();

    const elapsed2 = Date.now() - startTime2;
    console.log(`✓ Multiple sheets completed in ${elapsed2}ms\n`);

    // Example 3: Mixed data types without autofilter
    console.log('Example 3: Mixed data types (no autofilter)...');
    writer.startSheet('MixedTypes', 6, undefined, { doAutofilter: false });

    for (let i = 0; i < 1000; i++) {
        writer.writeRow([
            i,
            `Text ${i}`,
            Math.PI * i,
            new Date(),
            i % 2 === 0,
            null  // null values are supported
        ]);
    }
    writer.endSheet();
    console.log('✓ Completed\n');

    // Finalize the workbook
    console.log('Finalizing workbook...');
    await writer.finalize();
    console.log('✓ Done!\n');

    console.log('=== Summary ===');
    console.log('Total sheets created: 4');
    console.log('Total rows written: ~161,000');
    console.log(`File saved to: ${outputPath}`);
}

// Memory comparison function
async function compareMemoryUsage() {
    console.log('\n=== Memory Usage Comparison ===\n');

    const rowCount = 50_000;
    const cols = 5;

    // Test 1: Streaming mode
    console.log('Test 1: Streaming mode');
    const streamingFile = path.join(__dirname, '..', 'output', 'streaming-test.xlsb');
    const writer1 = new XlsbWriter(streamingFile);

    const streamStart = Date.now();
    const streamMemStart = process.memoryUsage().heapUsed / 1024 / 1024;

    writer1.startSheet('Data', cols, ['Col1', 'Col2', 'Col3', 'Col4', 'Col5']);
    for (let i = 0; i < rowCount; i++) {
        writer1.writeRow([i, `Text${i}`, Math.random() * 1000, new Date(), i % 2 === 0]);
    }
    writer1.endSheet();
    await writer1.finalize();

    const streamEnd = Date.now();
    const streamMemEnd = process.memoryUsage().heapUsed / 1024 / 1024;
    const streamTime = streamEnd - streamStart;
    const streamMem = (streamMemEnd - streamMemStart).toFixed(2);

    console.log(`  Time: ${streamTime}ms`);
    console.log(`  Memory delta: ${streamMem} MB\n`);

    // Test 2: Batch mode (traditional writeSheet)
    console.log('Test 2: Batch mode (traditional writeSheet)');
    const batchFile = path.join(__dirname, '..', 'output', 'batch-test.xlsb');
    const writer2 = new XlsbWriter(batchFile);

    const batchStart = Date.now();
    const batchMemStart = process.memoryUsage().heapUsed / 1024 / 1024;

    // Generate all rows in memory first
    const rows: any[][] = [];
    for (let i = 0; i < rowCount; i++) {
        rows.push([i, `Text${i}`, Math.random() * 1000, new Date(), i % 2 === 0]);
    }

    writer2.addSheet('Data');
    writer2.writeSheet(rows, ['Col1', 'Col2', 'Col3', 'Col4', 'Col5']);
    await writer2.finalize();

    const batchEnd = Date.now();
    const batchMemEnd = process.memoryUsage().heapUsed / 1024 / 1024;
    const batchTime = batchEnd - batchStart;
    const batchMem = (batchMemEnd - batchMemStart).toFixed(2);

    console.log(`  Time: ${batchTime}ms`);
    console.log(`  Memory delta: ${batchMem} MB\n`);

    console.log('=== Results ===');
    console.log(`Streaming is ${(parseFloat(batchMem) / parseFloat(streamMem)).toFixed(1)}x more memory efficient`);
    console.log(`Speed difference: ${Math.abs(streamTime - batchTime)}ms ${streamTime < batchTime ? 'faster' : 'slower'}`);
}

// Run examples
(async () => {
    try {
        await streamingExample();
        await compareMemoryUsage();
    } catch (error) {
        console.error('Error:', error);
        process.exit(1);
    }
})();
