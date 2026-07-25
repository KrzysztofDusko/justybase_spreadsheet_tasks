/**
 * Benchmark for XlsbWriterNode library only (without external dependencies like ExcelJS)
 * Run with: npx ts-node benchmarks/performance.bench.ts
 */
import { XlsbWriter } from '../dist/XlsbWriter';
import { XlsxWriter } from '../dist/XlsxWriter';
import { XlsbReader } from '../dist/XlsbReader';
import { XlsxReader } from '../dist/XlsxReader';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { execSync } from 'child_process';

const ROWS = 50000;
const ITERATIONS = 20;

const outFileXlsb = path.join(os.homedir(), 'bench_our.xlsb');
const outFileXlsx = path.join(os.homedir(), 'bench_our.xlsx');

function getFilesystemType(filePath: string): string {
    try {
        if (os.platform() === 'win32') {
            const drive = path.parse(filePath).root.replace('\\', '');
            const result = execSync(`wmic logicaldisk where "DeviceID='${drive}'" get FileSystem /value`, { encoding: 'utf8' });
            const match = result.match(/FileSystem=(\w+)/);
            return match ? match[1] : 'unknown';
        } else {
            const dfResult = execSync(`df "${filePath}" | tail -1 | awk '{print $1}'`, { encoding: 'utf8' }).trim();
            const mountResult = execSync(`mount | grep "^${dfResult} " | awk '{print $5}'`, { encoding: 'utf8' }).trim();
            return mountResult || 'unknown';
        }
    } catch {
        return 'unknown';
    }
}

function getOsName(): string {
    if (os.platform() === 'win32') {
        try {
            const result = execSync('wmic os get Caption /value', { encoding: 'utf8' });
            const match = result.match(/Caption=(.+)/);
            return match ? match[1].trim() : 'Windows';
        } catch {
            return 'Windows';
        }
    } else if (os.platform() === 'linux') {
        try {
            const result = execSync('cat /etc/os-release | grep "^PRETTY_NAME=" | cut -d= -f2 | tr -d \'"\'', { encoding: 'utf8' });
            return result.trim() || 'Linux';
        } catch {
            return 'Linux';
        }
    } else if (os.platform() === 'darwin') {
        try {
            const result = execSync('sw_vers -productName && sw_vers -productVersion', { encoding: 'utf8' });
            return result.replace('\n', ' ').trim();
        } catch {
            return 'macOS';
        }
    }
    return os.platform();
}

const headers = ['ID', 'Name', 'Count', 'Score', 'Date', 'Active', 'Description'];

function getDataRow(i: number): any[] {
    return [
        i,
        `Produkt ${i} żółć`,
        Math.floor(Math.random() * 10000),
        Math.random() * 100,
        new Date(),
        i % 2 === 0,
        `Opis produktu ${i} z polskimi znakami: ąęśćńźółĄĘŚĆŃŹÓŁ oraz dłuższy tekst testowy.`
    ];
}

interface BenchmarkResult {
    name: string;
    times: number[];
    avgTime: number;
    stdDev: number;
    size: number;
}

async function benchmarkWrite(
    name: string,
    filePath: string,
    writerFactory: () => any,
    data: any[][],
    headers: string[]
): Promise<BenchmarkResult> {
    const times: number[] = [];

    for (let iter = 0; iter < ITERATIONS; iter++) {
        if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
        if (typeof global.gc === 'function') global.gc();

        const start = performance.now();
        const writer = writerFactory();
        writer.addSheet('Benchmark');
        writer.writeSheet(data, headers);
        await writer.finalize();
        const end = performance.now();

        times.push(end - start);
    }

    const sortedTimes = [...times].sort((a, b) => a - b);
    const validTimes = sortedTimes.length > 4 ? sortedTimes.slice(2, -2) : sortedTimes;
    const avgTime = validTimes.reduce((a, b) => a + b, 0) / validTimes.length;
    const variance = validTimes.reduce((a, b) => a + Math.pow(b - avgTime, 2), 0) / validTimes.length;
    const stdDev = Math.sqrt(variance);
    const size = fs.statSync(filePath).size;

    return { name, times, avgTime, stdDev, size };
}

async function benchmarkWriteStreaming(
    name: string,
    filePath: string,
    writerFactory: () => any,
    rowCount: number,
    headers: string[]
): Promise<BenchmarkResult> {
    const times: number[] = [];

    for (let iter = 0; iter < ITERATIONS; iter++) {
        if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
        if (typeof global.gc === 'function') global.gc();

        const start = performance.now();
        const writer = writerFactory();

        writer.startSheet('Benchmark', headers.length, headers);
        for (let i = 0; i < rowCount; i++) {
            writer.writeRow(getDataRow(i));
        }
        writer.endSheet();

        await writer.finalize();
        const end = performance.now();

        times.push(end - start);
    }

    const sortedTimes = [...times].sort((a, b) => a - b);
    const validTimes = sortedTimes.length > 4 ? sortedTimes.slice(2, -2) : sortedTimes;
    const avgTime = validTimes.reduce((a, b) => a + b, 0) / validTimes.length;
    const variance = validTimes.reduce((a, b) => a + Math.pow(b - avgTime, 2), 0) / validTimes.length;
    const stdDev = Math.sqrt(variance);
    const size = fs.statSync(filePath).size;

    return { name, times, avgTime, stdDev, size };
}

interface MemoryBenchmarkResult {
    name: string;
    peakMemoryMB: number;
    avgMemoryMB: number;
    memoryDeltaMB: number;
}

async function benchmarkMemoryWrite(
    name: string,
    filePath: string,
    mode: 'batch' | 'streaming',
    rowCount: number,
    headers: string[]
): Promise<MemoryBenchmarkResult> {
    if (fs.existsSync(filePath)) fs.unlinkSync(filePath);

    if (typeof global.gc === 'function') {
        global.gc();
        global.gc();
    }

    await new Promise(resolve => setTimeout(resolve, 200));

    const baselineMemory = process.memoryUsage().heapUsed / 1024 / 1024;
    let peakMemory = baselineMemory;

    const memoryMonitor = setInterval(() => {
        const currentMemory = process.memoryUsage().heapUsed / 1024 / 1024;
        if (currentMemory > peakMemory) {
            peakMemory = currentMemory;
        }
    }, 5);

    try {
        const writer = new XlsbWriter(filePath);

        if (mode === 'batch') {
            const data: any[][] = [];
            for (let i = 0; i < rowCount; i++) {
                data.push(getDataRow(i));

                if (i % 10000 === 0 && i > 0) {
                    const currentMemory = process.memoryUsage().heapUsed / 1024 / 1024;
                    if (currentMemory > peakMemory) {
                        peakMemory = currentMemory;
                    }
                }
            }

            const afterDataGenMemory = process.memoryUsage().heapUsed / 1024 / 1024;
            if (afterDataGenMemory > peakMemory) {
                peakMemory = afterDataGenMemory;
            }

            writer.addSheet('MemoryTest');
            writer.writeSheet(data, headers);
        } else {
            writer.startSheet('MemoryTest', headers.length, headers);
            for (let i = 0; i < rowCount; i++) {
                writer.writeRow(getDataRow(i));
            }
            writer.endSheet();
        }

        await writer.finalize();
    } finally {
        clearInterval(memoryMonitor);
    }

    const finalMemory = process.memoryUsage().heapUsed / 1024 / 1024;
    if (finalMemory > peakMemory) {
        peakMemory = finalMemory;
    }

    const memoryDelta = Math.max(0, peakMemory - baselineMemory);
    const avgMemory = (baselineMemory + peakMemory) / 2;

    return {
        name,
        peakMemoryMB: peakMemory,
        avgMemoryMB: avgMemory,
        memoryDeltaMB: memoryDelta
    };
}

async function benchmarkRead(
    name: string,
    filePath: string,
    readerFactory: () => any
): Promise<BenchmarkResult> {
    const times: number[] = [];

    for (let iter = 0; iter < ITERATIONS; iter++) {
        if (typeof global.gc === 'function') global.gc();

        const start = performance.now();
        const reader = readerFactory();
        await reader.open(filePath);
        let rowCount = 0;
        while (await reader.read()) {
            rowCount++;
            reader.getValue(1);
        }
        if (reader.close) await reader.close();
        const end = performance.now();

        times.push(end - start);
    }

    const sortedTimes = [...times].sort((a, b) => a - b);
    const validTimes = sortedTimes.length > 4 ? sortedTimes.slice(2, -2) : sortedTimes;
    const avgTime = validTimes.reduce((a, b) => a + b, 0) / validTimes.length;
    const variance = validTimes.reduce((a, b) => a + Math.pow(b - avgTime, 2), 0) / validTimes.length;
    const stdDev = Math.sqrt(variance);

    return { name, times, avgTime, stdDev, size: 0 };
}

function printResult(result: BenchmarkResult, log: (msg: string) => void, showSize: boolean = true): void {
    const timesStr = result.times.map(t => t.toFixed(1)).join(', ');
    log(`  ${result.name}:`);
    log(`    Times: [${timesStr}] ms`);
    log(`    Average: ${result.avgTime.toFixed(2)} ms (±${result.stdDev.toFixed(2)} ms)`);
    if (showSize && result.size > 0) {
        log(`    Size: ${(result.size / 1024 / 1024).toFixed(2)} MB`);
    }
}

async function benchmark(): Promise<void> {
    const lines: string[] = [];
    const log = (msg: string = '') => {
        console.log(msg);
        lines.push(msg);
    };

    const platformName = os.platform() === 'win32' ? 'windows' : os.platform();
    const resultFilename = `benchmark_results_${platformName}.txt`;

    const osName = getOsName();
    const fsType = getFilesystemType(outFileXlsb);

    log(`Command to reproduce: npx ts-node benchmarks/performance.bench.ts`);
    log(`\n${'='.repeat(60)}`);
    log('XlsbWriterNode Benchmark');
    log(`${'='.repeat(60)}`);
    log(`OS: ${osName}`);
    log(`Filesystem: ${fsType}`);
    log(`Rows: ${ROWS}, Iterations: ${ITERATIONS}\n`);

    log('Generating test data...');
    const data: any[][] = [];
    for (let i = 0; i < ROWS; i++) {
        data.push(getDataRow(i));
    }
    log('Data generated.\n');

    // WRITE BENCHMARKS
    log('--- WRITE BENCHMARKS (Batch Mode) ---\n');

    const xlsbWriteResult = await benchmarkWrite(
        'XlsbWriter (Batch)',
        outFileXlsb,
        () => new XlsbWriter(outFileXlsb),
        data,
        headers
    );
    printResult(xlsbWriteResult, log);

    const xlsxWriteResult = await benchmarkWrite(
        'XlsxWriter (Batch)',
        outFileXlsx,
        () => new XlsxWriter(outFileXlsx),
        data,
        headers
    );
    printResult(xlsxWriteResult, log);

    // WRITE BENCHMARKS - STREAMING MODE
    log('\n--- WRITE BENCHMARKS (Streaming Mode) ---\n');

    const outFileXlsbStreaming = path.join(os.homedir(), 'bench_our_streaming.xlsb');
    const xlsbStreamingResult = await benchmarkWriteStreaming(
        'XlsbWriter (Streaming)',
        outFileXlsbStreaming,
        () => new XlsbWriter(outFileXlsbStreaming),
        ROWS,
        headers
    );
    printResult(xlsbStreamingResult, log);

    log('\n  Streaming vs Batch:');
    const speedDiff = ((xlsbWriteResult.avgTime - xlsbStreamingResult.avgTime) / xlsbWriteResult.avgTime * 100).toFixed(1);
    if (xlsbStreamingResult.avgTime < xlsbWriteResult.avgTime) {
        log(`    Streaming is ${speedDiff}% faster than batch mode`);
    } else {
        log(`    Batch is ${Math.abs(parseFloat(speedDiff))}% faster than streaming mode`);
    }
    log(`    File size difference: ${((xlsbStreamingResult.size - xlsbWriteResult.size) / 1024).toFixed(2)} KB`);

    // MEMORY USAGE BENCHMARKS
    log('\n--- MEMORY USAGE BENCHMARKS ---\n');
    log('Testing with larger dataset to demonstrate memory efficiency...\n');

    const MEMORY_TEST_ROWS = 100_000;
    const memoryTestFile = path.join(os.homedir(), 'bench_memory_test.xlsb');

    log(`  Testing with ${MEMORY_TEST_ROWS.toLocaleString()} rows...`);

    const memoryBatchResult = await benchmarkMemoryWrite(
        'Batch Mode',
        memoryTestFile,
        'batch',
        MEMORY_TEST_ROWS,
        headers
    );

    if (fs.existsSync(memoryTestFile)) fs.unlinkSync(memoryTestFile);

    const memoryStreamingResult = await benchmarkMemoryWrite(
        'Streaming Mode',
        memoryTestFile,
        'streaming',
        MEMORY_TEST_ROWS,
        headers
    );

    log(`\n  Batch Mode (${MEMORY_TEST_ROWS.toLocaleString()} rows):`);
    log(`    Peak Memory: ${memoryBatchResult.peakMemoryMB.toFixed(2)} MB`);
    log(`    Memory Delta: ${memoryBatchResult.memoryDeltaMB.toFixed(2)} MB`);

    log(`\n  Streaming Mode (${MEMORY_TEST_ROWS.toLocaleString()} rows):`);
    log(`    Peak Memory: ${memoryStreamingResult.peakMemoryMB.toFixed(2)} MB`);
    log(`    Memory Delta: ${memoryStreamingResult.memoryDeltaMB.toFixed(2)} MB`);

    const memoryEfficiency = (memoryBatchResult.memoryDeltaMB / memoryStreamingResult.memoryDeltaMB).toFixed(2);
    log(`\n  Memory Efficiency:`);
    log(`    Streaming uses ${memoryEfficiency}x LESS memory than batch mode`);
    log(`    Memory saved: ${(memoryBatchResult.memoryDeltaMB - memoryStreamingResult.memoryDeltaMB).toFixed(2)} MB`);

    // READ BENCHMARKS
    log('\n--- READ BENCHMARKS ---\n');

    const xlsbReadResult = await benchmarkRead(
        'XlsbReader',
        outFileXlsb,
        () => new XlsbReader()
    );
    printResult(xlsbReadResult, log, false);

    const xlsxReadResult = await benchmarkRead(
        'XlsxReader',
        outFileXlsx,
        () => new XlsxReader()
    );
    printResult(xlsxReadResult, log, false);

    // Summary
    log(`\n${'='.repeat(60)}`);
    log('SUMMARY');
    log(`${'='.repeat(60)}`);
    log('\nWrite Performance (Batch Mode):');
    log(`  XlsbWriter: ${xlsbWriteResult.avgTime.toFixed(2)} ms (${(xlsbWriteResult.size / 1024 / 1024).toFixed(2)} MB)`);
    log(`  XlsxWriter: ${xlsxWriteResult.avgTime.toFixed(2)} ms (${(xlsxWriteResult.size / 1024 / 1024).toFixed(2)} MB)`);
    log(`  XLSB is ${(xlsxWriteResult.avgTime / xlsbWriteResult.avgTime).toFixed(2)}x faster than XLSX`);

    log('\nWrite Performance (Streaming Mode):');
    log(`  XlsbWriter Streaming: ${xlsbStreamingResult.avgTime.toFixed(2)} ms (${(xlsbStreamingResult.size / 1024 / 1024).toFixed(2)} MB)`);
    const streamingSpeedup = (xlsbWriteResult.avgTime / xlsbStreamingResult.avgTime).toFixed(2);
    if (parseFloat(streamingSpeedup) > 1) {
        log(`  Streaming is ${streamingSpeedup}x faster than batch mode`);
    } else {
        log(`  Batch is ${(1 / parseFloat(streamingSpeedup)).toFixed(2)}x faster than streaming mode`);
    }

    log('\nRead Performance:');
    log(`  XlsbReader: ${xlsbReadResult.avgTime.toFixed(2)} ms`);
    log(`  XlsxReader: ${xlsxReadResult.avgTime.toFixed(2)} ms`);
    log(`  XLSB is ${(xlsxReadResult.avgTime / xlsbReadResult.avgTime).toFixed(2)}x faster than XLSX`);

    log('\nMemory Usage (100K rows):');
    log(`  Batch Mode Delta: ${memoryBatchResult.memoryDeltaMB.toFixed(2)} MB`);
    log(`  Streaming Mode Delta: ${memoryStreamingResult.memoryDeltaMB.toFixed(2)} MB`);
    log(`  Streaming is ${(memoryBatchResult.memoryDeltaMB / memoryStreamingResult.memoryDeltaMB).toFixed(2)}x more memory efficient`);

    log(`\n${'='.repeat(60)}\n`);

    const resultPath = path.join(__dirname, '..', resultFilename);
    fs.writeFileSync(resultPath, lines.join('\n'));
    console.log(`Results saved to: ${resultPath}`);
}

benchmark().catch(console.error);
