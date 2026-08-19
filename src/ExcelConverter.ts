import * as fs from 'fs';
import * as path from 'path';
import { randomUUID } from 'crypto';
import { spawn } from 'child_process';
import { installTemporaryOutput } from './atomicFile';

/** Options for an Excel format conversion. */
export interface ExcelConversionOptions {
    /** Allow replacing an existing destination file. Defaults to false. */
    overwrite?: boolean;
}

const POWERSHELL_CONVERSION_SCRIPT = `
$ErrorActionPreference = 'Stop'
$inputPath = [Environment]::GetEnvironmentVariable('JUSTYBASE_EXCEL_INPUT')
$outputPath = [Environment]::GetEnvironmentVariable('JUSTYBASE_EXCEL_OUTPUT')
$fileFormat = [int][Environment]::GetEnvironmentVariable('JUSTYBASE_EXCEL_FORMAT')
$excel = $null
$workbook = $null
$exitCode = 0

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    # Open read-only so the source workbook is never changed by the conversion.
    $workbook = $excel.Workbooks.Open($inputPath, 0, $true)
    if ($null -eq $workbook) {
        throw 'Excel did not return a workbook when opening the source file.'
    }

    $workbook.SaveAs($outputPath, $fileFormat)
} catch {
    $exitCode = 1
    [Console]::Error.WriteLine($_.Exception.ToString())
} finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch { }
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

exit $exitCode
`;

function errorMessage(error: unknown): string {
    return error instanceof Error ? error.message : String(error);
}

function validateConversionPaths(
    inputPath: string,
    outputPath: string,
    inputExtension: string,
    outputExtension: string,
    overwrite: boolean
): { input: string; output: string } {
    const input = path.resolve(inputPath);
    const output = path.resolve(outputPath);

    if (path.extname(input).toLowerCase() !== inputExtension) {
        throw new Error(
            `Excel conversion: source must have the ${inputExtension} extension: ${inputPath}`
        );
    }
    if (path.extname(output).toLowerCase() !== outputExtension) {
        throw new Error(
            `Excel conversion: destination must have the ${outputExtension} extension: ${outputPath}`
        );
    }
    if (!fs.existsSync(input)) {
        throw new Error(`Excel conversion: source file not found: ${inputPath}`);
    }
    if (!fs.statSync(input).isFile()) {
        throw new Error(`Excel conversion: source path is not a file: ${inputPath}`);
    }
    if (!fs.existsSync(path.dirname(output))) {
        throw new Error(`Excel conversion: destination directory not found: ${path.dirname(output)}`);
    }
    if (fs.existsSync(output)) {
        if (fs.statSync(output).isDirectory()) {
            throw new Error(`Excel conversion: destination path is a directory: ${outputPath}`);
        }
        if (!overwrite) {
            throw new Error(
                `Excel conversion: destination already exists (pass { overwrite: true } to replace it): ${outputPath}`
            );
        }
    }

    if (process.platform === 'win32' && input.toLowerCase() === output.toLowerCase()) {
        throw new Error('Excel conversion: source and destination must be different files.');
    }

    return { input, output };
}

function runPowerShellConversion(inputPath: string, outputPath: string, fileFormat: 50 | 51): Promise<void> {
    const encodedCommand = Buffer.from(POWERSHELL_CONVERSION_SCRIPT, 'utf16le').toString('base64');
    const environment: NodeJS.ProcessEnv = {
        ...process.env,
        JUSTYBASE_EXCEL_INPUT: inputPath,
        JUSTYBASE_EXCEL_OUTPUT: outputPath,
        JUSTYBASE_EXCEL_FORMAT: String(fileFormat),
    };

    return new Promise((resolve, reject) => {
        let stdout = '';
        let stderr = '';
        let settled = false;

        const child = spawn(
            'powershell.exe',
            [
                '-NoLogo',
                '-NoProfile',
                '-NonInteractive',
                '-ExecutionPolicy',
                'Bypass',
                '-EncodedCommand',
                encodedCommand,
            ],
            { env: environment, windowsHide: true }
        );

        child.stdout?.setEncoding('utf8');
        child.stderr?.setEncoding('utf8');
        child.stdout?.on('data', (chunk: string) => { stdout += chunk; });
        child.stderr?.on('data', (chunk: string) => { stderr += chunk; });

        child.once('error', (error) => {
            if (settled) return;
            settled = true;
            reject(new Error(`Excel conversion: unable to start PowerShell/COM: ${error.message}`));
        });

        child.once('close', (code) => {
            if (settled) return;
            settled = true;
            if (code === 0) {
                resolve();
                return;
            }

            const details = (stderr || stdout).replace(/\s+/g, ' ').trim();
            reject(new Error(
                `Excel conversion failed (PowerShell exit code ${code ?? 'unknown'})${details ? `: ${details}` : '.'}`
            ));
        });
    });
}

async function convertExcelFile(
    inputPath: string,
    outputPath: string,
    inputExtension: string,
    outputExtension: string,
    fileFormat: 50 | 51,
    options: ExcelConversionOptions = {}
): Promise<void> {
    const { input, output } = validateConversionPaths(
        inputPath,
        outputPath,
        inputExtension,
        outputExtension,
        options.overwrite === true
    );

    if (process.platform !== 'win32') {
        throw new Error('Excel conversion requires Windows with desktop Microsoft Excel and COM automation.');
    }

    const temporaryOutput = path.join(
        path.dirname(output),
        `.${path.basename(output, outputExtension)}.${randomUUID()}.conversion${outputExtension}`
    );

    try {
        await runPowerShellConversion(input, temporaryOutput, fileFormat);
        if (!fs.existsSync(temporaryOutput)) {
            throw new Error('Excel conversion completed without creating the destination file.');
        }

        installTemporaryOutput(temporaryOutput, output, options.overwrite === true);
    } catch (error) {
        throw new Error(`Excel conversion ${inputExtension} -> ${outputExtension}: ${errorMessage(error)}`);
    } finally {
        if (fs.existsSync(temporaryOutput)) {
            try {
                fs.rmSync(temporaryOutput);
            } catch {
                // Preserve the original conversion error if cleanup also fails.
            }
        }
    }
}

/** Convert an Excel Binary Workbook (.xlsb) to an Open XML Workbook (.xlsx). */
export function convertXlsbToXlsx(
    inputPath: string,
    outputPath: string,
    options?: ExcelConversionOptions
): Promise<void> {
    return convertExcelFile(inputPath, outputPath, '.xlsb', '.xlsx', 51, options);
}

/** Convert an Open XML Workbook (.xlsx) to an Excel Binary Workbook (.xlsb). */
export function convertXlsxToXlsb(
    inputPath: string,
    outputPath: string,
    options?: ExcelConversionOptions
): Promise<void> {
    return convertExcelFile(inputPath, outputPath, '.xlsx', '.xlsb', 50, options);
}
