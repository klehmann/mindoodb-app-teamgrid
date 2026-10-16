import { expect, type Page } from '@playwright/test'
// Declares `window.__havenTestHost`.
import type {} from 'mindoodb-app-sdk/testing'

type AgentResult = { ok: true; result: unknown } | { ok: false; error: { code: string; message: string } }

export interface SheetInfo {
  sheetId: string
  name: string
  rows?: number
  columns?: number
}

export interface ReadCell {
  value: unknown
  text?: unknown
  formula?: string
}

/**
 * TeamGrid inside the mock Haven at `/__haven-test/`, driven through its agent tools the
 * way an AI agent in Haven would: workbooks are imported with `workbook_create` and read
 * with `read_cells`, so the tests run the editor's own import, store (MindooDB) and
 * formula engine.
 */
export class TeamGridHost {
  constructor(readonly page: Page) {}

  async open() {
    await this.page.goto('/__haven-test/?locale=en')
    await this.page.waitForFunction(
      () => window.__havenTestHost?.agent.tools().some((tool) => tool.name === 'workbook_create') ?? false,
      undefined,
      { timeout: 90_000 },
    )
    await this.page.evaluate(() => window.__havenTestHost!.agent.setConsent?.('allow'))
  }

  async call<T = unknown>(tool: string, input: Record<string, unknown> = {}): Promise<T> {
    const outcome = (await this.page.evaluate(
      ([name, args]) => window.__havenTestHost!.agent.call(name, args),
      [tool, input] as const,
    )) as AgentResult
    if (!outcome.ok) throw new Error(`${tool}: ${outcome.error.code} ${outcome.error.message}`)
    return outcome.result as T
  }

  /**
   * Imports an .xlsx as a new workbook and opens it; resolves once formulas are calculated.
   * `streaming`: too large to load whole, so formulas keep Excel's stored results.
   */
  async importWorkbook(title: string, bytes: Uint8Array): Promise<{ sheets: SheetInfo[]; streaming: boolean }> {
    try {
      return await this.importOnce(title, bytes)
    } catch (error) {
      // The mock Haven's File handed through postMessage is now and then unreadable
      // (NotReadableError) in long runs; a fresh import of the same bytes works.
      if (!String(error).includes('could not be read')) throw error
      return this.importOnce(title, bytes)
    }
  }

  private async importOnce(title: string, bytes: Uint8Array): Promise<{ sheets: SheetInfo[]; streaming: boolean }> {
    const fileRef = await this.page.evaluate(
      ([name, base64]) =>
        window.__havenTestHost!.agent.importFile({
          name,
          mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          data: Uint8Array.from(atob(base64), (c) => c.charCodeAt(0)),
        }),
      [`${title}.xlsx`, Buffer.from(bytes).toString('base64')] as const,
    )
    const context = await this.call<{ sheets: SheetInfo[]; streaming?: boolean }>('workbook_create', { title, fileRef })
    return { sheets: context.sheets, streaming: context.streaming === true }
  }

  /** The open workbook as TeamGrid exports it (`workbook_export`). */
  async exportXlsx(): Promise<Uint8Array> {
    const { fileRef } = await this.call<{ fileRef: string }>('workbook_export', { format: 'xlsx' })
    const base64 = await this.page.evaluate((ref) => {
      const file = window.__havenTestHost!.agent.getFile(ref)
      if (!file) throw new Error(`No file ${ref}`)
      const data = new Uint8Array(file.data as ArrayLike<number>)
      let binary = ''
      for (let i = 0; i < data.length; i += 0x8000) binary += String.fromCharCode(...data.subarray(i, i + 0x8000))
      return btoa(binary)
    }, fileRef)
    return new Uint8Array(Buffer.from(base64, 'base64'))
  }

  /** Cells of one sheet, 100 per `read_cells` call (the tool's limit). */
  async readCells(sheetId: string, addresses: string[]): Promise<Record<string, ReadCell>> {
    const cells: Record<string, ReadCell> = {}
    for (let i = 0; i < addresses.length; i += 100) {
      const part = await this.call<{ cells: Record<string, ReadCell> }>('read_cells', {
        sheetId,
        addresses: addresses.slice(i, i + 100),
      })
      Object.assign(cells, part.cells)
    }
    return cells
  }
}

export { expect }
