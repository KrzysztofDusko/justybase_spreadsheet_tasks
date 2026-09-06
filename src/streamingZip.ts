import * as fs from 'fs';
import * as path from 'path';
import { randomUUID } from 'crypto';
import { finished } from 'stream/promises';
import { Readable } from 'stream';
import AdmZip from 'adm-zip';
import archiver from 'archiver';

/** A file-backed replacement for one ZIP member. */
export type StagedZipPart = Map<string, string>;

function readDestinationMode(target: string): number | undefined {
    try {
        if (!fs.existsSync(target)) return undefined;
        return fs.statSync(target).mode & 0o777;
    } catch {
        return undefined;
    }
}

function appendAndWait(
    archive: archiver.Archiver,
    source: NodeJS.ReadableStream,
    name: string,
): Promise<void> {
    return new Promise<void>((resolve, reject) => {
        const cleanup = (): void => {
            archive.off('entry', onEntry);
            archive.off('error', onError);
            const maybeEmitter = source as NodeJS.EventEmitter;
            if (typeof maybeEmitter.off === 'function') {
                maybeEmitter.off('error', onError);
            }
        };
        const onEntry = (): void => {
            cleanup();
            resolve();
        };
        const onError = (error: Error): void => {
            cleanup();
            reject(error);
        };
        archive.once('entry', onEntry);
        archive.once('error', onError);
        const maybeEmitter = source as NodeJS.EventEmitter;
        if (typeof maybeEmitter.once === 'function') {
            maybeEmitter.once('error', onError);
        }
        try {
            archive.append(source as never, { name });
        } catch (error) {
            cleanup();
            reject(error as Error);
        }
    });
}

/**
 * Rebuild an existing ZIP package without creating a complete output Buffer.
 * Unchanged members are copied one at a time; staged members are read from
 * temporary files.
 */
export async function writeAdmZipStreaming(
    zip: AdmZip,
    stagedParts: StagedZipPart,
    destinationPath: string,
): Promise<void> {
    const target = path.resolve(destinationPath);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    const temporaryPath = path.join(
        path.dirname(target),
        `.${path.basename(target)}.${randomUUID()}.tmp`
    );
    const destinationMode = readDestinationMode(target);

    let output: fs.WriteStream | null = null;
    let archive: archiver.Archiver | null = null;
    try {
        output = fs.createWriteStream(
            temporaryPath,
            destinationMode !== undefined ? { mode: destinationMode } : {},
        );
        archive = archiver('zip');
        archive.pipe(output);

        for (const entry of zip.getEntries()) {
            const stagedPath = stagedParts.get(entry.entryName);
            if (stagedPath) {
                await appendAndWait(archive, fs.createReadStream(stagedPath), entry.entryName);
            } else {
                // Decompress each member only after the previous one has been
                // consumed so a single buffer is retained at a time.
                await appendAndWait(archive, Readable.from([entry.getData()]), entry.entryName);
            }
        }

        await archive.finalize();
        await finished(output);

        if (destinationMode !== undefined) {
            try {
                fs.chmodSync(temporaryPath, destinationMode);
            } catch {
                // Preserve the original archive or rename error.
            }
        }
        const fd = fs.openSync(temporaryPath, 'r+');
        try {
            fs.fsyncSync(fd);
        } finally {
            fs.closeSync(fd);
        }
        fs.renameSync(temporaryPath, target);
    } catch (error) {
        try {
            archive?.abort();
        } catch {
            // Preserve the original archive error.
        }
        throw error;
    } finally {
        if (output && !output.closed) {
            output.destroy();
        }
        if (fs.existsSync(temporaryPath)) {
            try {
                fs.rmSync(temporaryPath);
            } catch {
                // Preserve the original archive or rename error.
            }
        }
    }
}
