declare module 'yauzl' {
    interface Entry {
        fileName: string;
        compressedSize: number;
        uncompressedSize: number;
        crc32: number;
    }

    interface ZipFile {
        on(event: 'entry', listener: (entry: Entry) => void): this;
        on(event: 'end', listener: () => void): this;
        on(event: 'error', listener: (err: Error) => void): this;
        readEntry(): void;
        openReadStream(entry: Entry, callback: (err: Error | null, readStream: NodeJS.ReadableStream) => void): void;
        close(): void;
    }

    interface Options {
        lazyEntries?: boolean;
        autoClose?: boolean;
    }

    interface Yauzl {
        open(path: string, options: Options, callback: (err: Error | null, zipfile: ZipFile) => void): void;
    }

    const yauzl: Yauzl;

    export { ZipFile, Entry };
    export default yauzl;
}
