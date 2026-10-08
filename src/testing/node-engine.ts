// The xlsx engine for Node tests: the same WASM build, compiled for Node
// (`pnpm build:wasm:node` → native/xlsx-wasm/pkg-node). Tests that need it
// skip when it has not been built.
import { existsSync } from 'node:fs'
import { createRequire } from 'node:module'

import type { XlsxEngine } from '../xlsx/engine'

const PKG = new URL('../../native/xlsx-wasm/pkg-node/xlsx_wasm.js', import.meta.url).pathname

export const nodeEngineAvailable = existsSync(PKG)

export function createNodeEngine(): XlsxEngine {
  const { XlsxEngine: Engine } = createRequire(import.meta.url)(PKG) as {
    XlsxEngine: new () => {
      open(bytes: Uint8Array, name: string | null, locale: string | null, shortDate: string | null): string
      readRange(session: string, sheet: string, r0: number, r1: number, c0: number, c1: number): string
      readFormulaCells(session: string, sheet: string): string
      readMedia(session: string, visual: string): string
      readMediaBytes(session: string, visual: string): Uint8Array
      close(session: string): void
    }
  }
  const engine = new Engine()
  return {
    open: async (bytes, locale, shortDate) => JSON.parse(engine.open(bytes, null, locale, shortDate ?? null)),
    readRange: async ({ sessionId, sheetId, range }) =>
      JSON.parse(engine.readRange(sessionId, sheetId, range.startRow, range.endRow, range.startColumn, range.endColumn)),
    readFormulaCells: async ({ sessionId, sheetId }) => JSON.parse(engine.readFormulaCells(sessionId, sheetId)),
    readMedia: async ({ sessionId, visualId }) => JSON.parse(engine.readMedia(sessionId, visualId)),
    readMediaBytes: async ({ sessionId, visualId }) => engine.readMediaBytes(sessionId, visualId),
    close: async (sessionId) => engine.close(sessionId),
  }
}
