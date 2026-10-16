/* tslint:disable */
/* eslint-disable */

/**
 * One engine holds any number of open workbook sessions, keyed by the
 * `sessionId` that `open` returns.
 */
export class XlsxEngine {
    free(): void;
    [Symbol.dispose](): void;
    close(sessionId: string): void;
    /**
     * Sidecar `find_cells` result for a `findWorkbookCells` request (JSON).
     */
    findCells(request: string): string;
    constructor();
    /**
     * Opens an .xlsx from its bytes. Returns the sidecar `open` result
     * (WorkbookFile-shaped JSON incl. `sessionId`).
     */
    open(bytes: Uint8Array, name?: string | null, locale?: string | null, shortDateFormat?: string | null): string;
    /**
     * Sidecar `read_formula_cells` result.
     */
    readFormulaCells(sessionId: string, sheetId: string): string;
    /**
     * The raw image bytes behind a visual or cell-image id (for a Blob URL
     * without a base64 round trip on the JS side).
     */
    readMediaBytes(sessionId: string, visualId: string): Uint8Array;
    /**
     * Sidecar `read_media` result: `{ mediaType, base64 }` for a visual or
     * cell-image id.
     */
    readMedia(sessionId: string, visualId: string): string;
    /**
     * Sidecar `read_range` result (WorkbookRangeResult-shaped JSON). The
     * first read of a sheet indexes it completely (no threads in WASM), so
     * `indexingComplete` is always true.
     */
    readRange(sessionId: string, sheetId: string, startRow: number, endRow: number, startColumn: number, endColumn: number): string;
    /**
     * Sidecar `read_row_outline` result: the sheet's row outline levels.
     */
    readRowOutline(sessionId: string, sheetId: string): string;
}
