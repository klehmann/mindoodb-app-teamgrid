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
}

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly __wbg_xlsxengine_free: (a: number, b: number) => void;
    readonly xlsxengine_close: (a: number, b: number, c: number) => [number, number];
    readonly xlsxengine_new: () => number;
    readonly xlsxengine_open: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number) => [number, number, number, number];
    readonly xlsxengine_readFormulaCells: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly xlsxengine_readMedia: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly xlsxengine_readMediaBytes: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly xlsxengine_readRange: (a: number, b: number, c: number, d: number, e: number, f: number, g: number, h: number, i: number) => [number, number, number, number];
    readonly __wbindgen_exn_store: (a: number) => void;
    readonly __externref_table_alloc: () => number;
    readonly __wbindgen_externrefs: WebAssembly.Table;
    readonly __wbindgen_malloc: (a: number, b: number) => number;
    readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
    readonly __externref_table_dealloc: (a: number) => void;
    readonly __wbindgen_free: (a: number, b: number, c: number) => void;
    readonly __wbindgen_start: () => void;
}

export type SyncInitInput = BufferSource | WebAssembly.Module;

/**
 * Instantiates the given `module`, which can either be bytes or
 * a precompiled `WebAssembly.Module`.
 *
 * @param {{ module: SyncInitInput }} module - Passing `SyncInitInput` directly is deprecated.
 *
 * @returns {InitOutput}
 */
export function initSync(module: { module: SyncInitInput } | SyncInitInput): InitOutput;

/**
 * If `module_or_path` is {RequestInfo} or {URL}, makes a request and
 * for everything else, calls `WebAssembly.instantiate` directly.
 *
 * @param {{ module_or_path: InitInput | Promise<InitInput> }} module_or_path - Passing `InitInput` directly is deprecated.
 *
 * @returns {Promise<InitOutput>}
 */
export default function __wbg_init (module_or_path?: { module_or_path: InitInput | Promise<InitInput> } | InitInput | Promise<InitInput>): Promise<InitOutput>;
