// XlsxEngine on top of the WASM build of GenOffice's sidecar
// (native/xlsx-wasm, built with `pnpm build:wasm`). Runs on the main thread
// for now; the first read of a large sheet indexes it synchronously, so this
// moves into a Web Worker before real use.
import init, { XlsxEngine as WasmXlsxEngine } from '../../native/xlsx-wasm/pkg/xlsx_wasm.js'
import wasmUrl from '../../native/xlsx-wasm/pkg/xlsx_wasm_bg.wasm?url'

import type { XlsxEngine } from './engine'

export async function createWasmEngine(): Promise<XlsxEngine> {
  await init({ module_or_path: wasmUrl })
  const engine = new WasmXlsxEngine()
  return {
    async open(bytes, locale, shortDateFormat) {
      return JSON.parse(engine.open(bytes, null, locale, shortDateFormat ?? null))
    },
    async readRange({ sessionId, sheetId, range }) {
      return JSON.parse(
        engine.readRange(sessionId, sheetId, range.startRow, range.endRow, range.startColumn, range.endColumn),
      )
    },
    async readFormulaCells({ sessionId, sheetId }) {
      return JSON.parse(engine.readFormulaCells(sessionId, sheetId))
    },
    async readMedia({ sessionId, visualId }) {
      return JSON.parse(engine.readMedia(sessionId, visualId))
    },
    async close(sessionId) {
      engine.close(sessionId)
    },
  }
}
