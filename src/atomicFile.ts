import * as fs from 'fs';
import { randomUUID } from 'crypto';

/** Internal file operations seam used to test replacement races without Excel. */
export interface AtomicFileOperations {
    exists(filePath: string): boolean;
    rename(sourcePath: string, destinationPath: string): void;
    remove(filePath: string): void;
}
const realFileOperations: AtomicFileOperations = {
    exists: (filePath) => fs.existsSync(filePath),
    rename: (sourcePath, destinationPath) => fs.renameSync(sourcePath, destinationPath),
    remove: (filePath) => fs.rmSync(filePath),
};

function errorMessage(error: unknown): string {
    return error instanceof Error ? error.message : String(error);
}

function destinationExistsError(destinationPath: string): Error {
    return new Error(
        `Excel conversion: destination already exists (pass { overwrite: true } to replace it): ${destinationPath}`
    );
}

/**
 * Installs a temporary conversion result without deleting an unapproved or
 * concurrently-created destination. Existing destinations are moved aside
 * before replacement so a failed final rename can be rolled back safely.
 */
export function installTemporaryOutput(
    temporaryPath: string,
    destinationPath: string,
    overwrite: boolean,
    operations: AtomicFileOperations = realFileOperations
): void {
    const backupPath = `${destinationPath}.${randomUUID()}.replacement-backup`;
    let backupCreated = false;
    let installed = false;

    if (operations.exists(destinationPath)) {
        if (!overwrite) {
            throw destinationExistsError(destinationPath);
        }

        // A rename preserves the old destination if anything fails later.
        operations.rename(destinationPath, backupPath);
        backupCreated = true;
    }

    try {
        try {
            operations.rename(temporaryPath, destinationPath);
            installed = true;
        } catch (error) {
            // If overwrite was allowed and a destination appeared after the
            // check above, preserve it as well and retry once.
            if (overwrite && !backupCreated && operations.exists(destinationPath)) {
                operations.rename(destinationPath, backupPath);
                backupCreated = true;
                operations.rename(temporaryPath, destinationPath);
                installed = true;
            } else {
                throw error;
            }
        }

        if (backupCreated) {
            operations.remove(backupPath);
        }
    } catch (error) {
        if (backupCreated && !installed) {
            if (!operations.exists(destinationPath)) {
                try {
                    operations.rename(backupPath, destinationPath);
                    backupCreated = false;
                } catch (rollbackError) {
                    throw new Error(
                        `Excel conversion: final destination install failed (${errorMessage(error)}), ` +
                        `and restoring the original destination also failed (${errorMessage(rollbackError)}). ` +
                        `Original file remains at ${backupPath}.`
                    );
                }
            } else {
                throw new Error(
                    `Excel conversion: final destination install failed (${errorMessage(error)}). ` +
                    `A concurrent destination was preserved; the original file remains at ${backupPath}.`
                );
            }
        }

        throw error;
    }
}
