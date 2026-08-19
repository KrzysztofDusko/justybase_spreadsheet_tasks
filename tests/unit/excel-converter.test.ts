/**
 * Excel converter validation tests.
 *
 * These tests deliberately stop before COM is started, so they run on every
 * platform. Excel integration is covered by tests/excel-compat/conversion.ts.
 */

import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import { convertXlsbToXlsx, convertXlsxToXlsb } from '../../src/ExcelConverter';
import { AtomicFileOperations, installTemporaryOutput } from '../../src/atomicFile';

const repoRoot = path.resolve(__dirname, '../..');
const outputDir = path.join(repoRoot, 'test-output', 'unit', 'excel-converter');

async function expectError(action: () => Promise<void>, expectedText: string): Promise<void> {
    await assert.rejects(action, (error: unknown) => {
        assert.ok(error instanceof Error);
        assert.match(error.message, new RegExp(expectedText));
        return true;
    });
}

function createMemoryOperations(initial: Record<string, string> = {}): {
    files: Map<string, string>;
    operations: AtomicFileOperations;
} {
    const files = new Map(Object.entries(initial));
    const operations: AtomicFileOperations = {
        exists: (filePath) => files.has(filePath),
        rename: (sourcePath, destinationPath) => {
            if (!files.has(sourcePath) || files.has(destinationPath)) {
                throw new Error(`rename failed: ${sourcePath} -> ${destinationPath}`);
            }
            files.set(destinationPath, files.get(sourcePath)!);
            files.delete(sourcePath);
        },
        remove: (filePath) => {
            if (!files.delete(filePath)) {
                throw new Error(`remove failed: ${filePath}`);
            }
        },
    };
    return { files, operations };
}

function testDestinationRaceWithoutOverwrite(): void {
    const temporaryPath = 'temporary.xlsx';
    const destinationPath = 'destination.xlsx';
    const memory = createMemoryOperations({ [temporaryPath]: 'converted' });
    let raceTriggered = false;
    const originalRename = memory.operations.rename;
    memory.operations.rename = (sourcePath, targetPath) => {
        if (!raceTriggered && sourcePath === temporaryPath && targetPath === destinationPath) {
            raceTriggered = true;
            memory.files.set(destinationPath, 'concurrent file');
            throw new Error('destination appeared concurrently');
        }
        originalRename(sourcePath, targetPath);
    };

    assert.throws(
        () => installTemporaryOutput(temporaryPath, destinationPath, false, memory.operations),
        /destination appeared concurrently/
    );
    assert.strictEqual(memory.files.get(destinationPath), 'concurrent file');
    assert.strictEqual(memory.files.get(temporaryPath), 'converted');
}

function testDestinationRollbackOnFailedOverwrite(): void {
    const temporaryPath = 'temporary.xlsx';
    const destinationPath = 'destination.xlsx';
    const memory = createMemoryOperations({
        [temporaryPath]: 'converted',
        [destinationPath]: 'original',
    });
    let backupPath = '';
    let finalRenameAttempted = false;
    const originalRename = memory.operations.rename;
    memory.operations.rename = (sourcePath, targetPath) => {
        if (sourcePath === destinationPath && targetPath.includes('.replacement-backup')) {
            backupPath = targetPath;
            originalRename(sourcePath, targetPath);
            return;
        }
        if (sourcePath === temporaryPath && targetPath === destinationPath && !finalRenameAttempted) {
            finalRenameAttempted = true;
            memory.files.set(destinationPath, 'concurrent file');
            throw new Error('final rename failed');
        }
        originalRename(sourcePath, targetPath);
    };

    assert.throws(
        () => installTemporaryOutput(temporaryPath, destinationPath, true, memory.operations),
        /concurrent destination was preserved/
    );
    assert.strictEqual(memory.files.get(destinationPath), 'concurrent file');
    assert.strictEqual(memory.files.get(backupPath), 'original');
    assert.strictEqual(memory.files.get(temporaryPath), 'converted');
}

async function run(): Promise<void> {
    fs.mkdirSync(outputDir, { recursive: true });

    testDestinationRaceWithoutOverwrite();
    testDestinationRollbackOnFailedOverwrite();

    await expectError(
        () => convertXlsbToXlsx('missing.xlsx', path.join(outputDir, 'output.xlsx')),
        'source must have the \\.xlsb extension'
    );

    await expectError(
        () => convertXlsxToXlsb('missing.xlsx', path.join(outputDir, 'output.xlsx')),
        'destination must have the \\.xlsb extension'
    );

    await expectError(
        () => convertXlsbToXlsx(path.join(repoRoot, 'missing.xlsb'), path.join(outputDir, 'output.xlsx')),
        'source file not found'
    );

    const existingDestination = path.join(outputDir, 'existing.xlsx');
    fs.writeFileSync(existingDestination, 'existing');
    await expectError(
        () => convertXlsbToXlsx(path.join(repoRoot, 'data_formats.xlsb'), existingDestination),
        'destination already exists'
    );

    const missingDirectory = path.join(outputDir, 'missing-directory', 'output.xlsx');
    await expectError(
        () => convertXlsbToXlsx(path.join(repoRoot, 'data_formats.xlsb'), missingDirectory),
        'destination directory not found'
    );

    fs.rmSync(outputDir, { recursive: true, force: true });
    console.log('Excel converter validation tests passed.');
}

run().catch((error: unknown) => {
    console.error(error);
    process.exit(1);
});
