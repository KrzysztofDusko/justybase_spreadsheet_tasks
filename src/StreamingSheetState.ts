import { BigBuffer } from './BigBuffer';
import { initColWidths } from './writerHelpers';

/**
 * Shared streaming-sheet session state for Node writers.
 * Holds buffer, row cursor, column range, autofilter flag, and autofit widths.
 */
export class StreamingSheetState {
    isStreaming: boolean = false;
    rowNum: number = 0;
    startCol: number = 0;
    endCol: number = 0;
    doAutofilter: boolean = false;
    colWidths: number[] = [];
    buffer: BigBuffer | null = null;

    /**
     * Begin a streaming sheet session.
     * @throws if already streaming
     */
    begin(columnCount: number, doAutofilter: boolean, buffer: BigBuffer): void {
        if (this.isStreaming) {
            throw new Error('Already in streaming mode. Call endSheet() first.');
        }

        this.isStreaming = true;
        this.buffer = buffer;
        this.rowNum = 0;
        this.startCol = 0;
        this.endCol = columnCount;
        this.doAutofilter = doAutofilter;
        this.colWidths = initColWidths(columnCount);
    }

    /** Assert streaming mode and return the active buffer. */
    assertStreaming(): BigBuffer {
        if (!this.isStreaming || !this.buffer) {
            throw new Error('Not in streaming mode. Call startSheet() first.');
        }
        return this.buffer;
    }

    /** Reset after endSheet(). */
    end(): void {
        this.isStreaming = false;
        this.buffer = null;
        this.rowNum = 0;
        this.startCol = 0;
        this.endCol = 0;
        this.doAutofilter = false;
        this.colWidths = [];
    }
}
