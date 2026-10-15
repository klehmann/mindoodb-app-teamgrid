// The browser counterpart of GenOffice's XlsxSidecarClient: the same requests
// and responses, served by the sidecar library compiled to WebAssembly.
export interface XlsxEngine {
  open(bytes: Uint8Array, locale: string, shortDateFormat?: string): Promise<unknown>
  readRange(input: {
    sessionId: string
    sheetId: string
    range: { startRow: number; endRow: number; startColumn: number; endColumn: number }
  }): Promise<unknown>
  readFormulaCells(input: { sessionId: string; sheetId: string }): Promise<unknown>
  /** GenOffice's findWorkbookCells: a streamed find over the whole workbook. */
  findCells(request: unknown): Promise<unknown>
  readRowOutline(input: { sessionId: string; sheetId: string }): Promise<unknown>
  readMedia(input: { sessionId: string; visualId: string }): Promise<unknown>
  /** The image bytes behind a visual, without a base64 round trip. */
  readMediaBytes(input: { sessionId: string; visualId: string }): Promise<Uint8Array>
  close(sessionId: string): Promise<void>
}

let enginePromise: Promise<XlsxEngine> | undefined

/** Loads the WASM module on first use, so the editor boots without it. */
export function loadXlsxEngine(): Promise<XlsxEngine> {
  enginePromise ??= import('./wasm-engine').then((module) => module.createWasmEngine())
  return enginePromise
}
