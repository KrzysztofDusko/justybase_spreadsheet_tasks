declare module 'adm-zip' {
    export default class AdmZip {
        constructor(path?: string);
        getEntries(): any[];
        readAsText(entry: any): string;
        extractFileTo(entry: any, targetPath: string, overwrite?: boolean): void;
        getEntry(entryName: string): any;
    }
}

declare module 'archiver' {
    interface Archiver {
        pipe(stream: any): any;
        append(source: any, data?: any): any;
        finalize(): Promise<void>;
        on(event: string, listener: (...args: any[]) => void): any;
    }

    function archiver(format: string, options?: any): Archiver;

    namespace archiver {
        export { Archiver };
        export function create(format: string, options?: any): Archiver;
    }

    export = archiver;
}

declare module 'yauzl' {
    export interface Entry {
        fileName: string;
    }

    export interface ZipFile {
        on(event: 'entry', listener: (entry: Entry) => void): this;
        on(event: 'end', listener: () => void): this;
        on(event: 'error', listener: (err: Error) => void): this;
        readEntry(): void;
        openReadStream(entry: Entry, callback: (err: Error | null, readStream: NodeJS.ReadableStream) => void): void;
        close(): void;
    }

    export interface OpenOptions {
        lazyEntries?: boolean;
        autoClose?: boolean;
    }

    export function open(path: string, options: OpenOptions, callback: (err: Error | null, zipfile: ZipFile) => void): void;
}
