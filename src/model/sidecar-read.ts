// Reads a whole workbook through the xlsx engine (WASM sidecar) into one
// in-memory snapshot in GenOffice's read format: the common ground between
// an xlsx file, the editor and the stored TeamGrid document.
import type {
  WorkbookFile,
  WorkbookRangeResult,
} from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import type { XlsxEngine } from '../xlsx/engine'

export type SheetMetadata = WorkbookFile['sheets'][number]
export type CellRecord = WorkbookRangeResult['cells'][number]
export type RowRecord = WorkbookRangeResult['rows'][number]
export type CellArea = WorkbookRangeResult['merges'][number]

/** Sheet-wide parts of a range read, delivered complete with any block. */
type SheetWide = Pick<
  WorkbookRangeResult,
  'autoFilter' | 'autoFilterColumns' | 'rowBreaks' | 'colBreaks' | 'pageSetup'
>

export interface SidecarSheet extends Partial<SheetWide> {
  meta: SheetMetadata
  cells: CellRecord[]
  rows: RowRecord[]
  merges: CellArea[]
  hyperlinks?: WorkbookRangeResult['hyperlinks']
  conditionalRules?: WorkbookRangeResult['conditionalRules']
  dataValidations?: WorkbookRangeResult['dataValidations']
}

export interface SidecarWorkbook {
  file: WorkbookFile
  sheets: SidecarSheet[]
}

/** The engine's per-request cap (rows × columns). */
const MAX_RANGE_CELLS = 100_000

export async function readWholeWorkbook(engine: XlsxEngine, file: WorkbookFile): Promise<SidecarWorkbook> {
  const sheets: SidecarSheet[] = []
  for (const meta of file.sheets) {
    const columns = Math.max(1, meta.columnCount)
    const rowsPerRead = Math.max(1, Math.floor(MAX_RANGE_CELLS / columns))
    const cells: CellRecord[] = []
    const rows: RowRecord[] = []
    const merges = new Map<string, CellArea>()
    const hyperlinks: WorkbookRangeResult['hyperlinks'] = []
    // Rules span ranges; a rule touching several blocks comes back with each.
    const conditionalRules = new Map<string, WorkbookRangeResult['conditionalRules'][number]>()
    const dataValidations = new Map<string, WorkbookRangeResult['dataValidations'][number]>()
    let wide: Partial<SheetWide> = {}
    for (let startRow = 0; startRow < meta.rowCount; startRow += rowsPerRead) {
      const result = (await engine.readRange({
        sessionId: file.sessionId,
        sheetId: meta.id,
        range: {
          startRow,
          endRow: Math.min(meta.rowCount, startRow + rowsPerRead) - 1,
          startColumn: 0,
          endColumn: columns - 1,
        },
      })) as WorkbookRangeResult
      cells.push(...result.cells)
      rows.push(...result.rows)
      for (const merge of result.merges) {
        merges.set(`${merge.startRow}:${merge.startColumn}:${merge.endRow}:${merge.endColumn}`, merge)
      }
      hyperlinks.push(...(result.hyperlinks ?? []))
      for (const rule of result.conditionalRules ?? []) conditionalRules.set(JSON.stringify(rule), rule)
      for (const rule of result.dataValidations ?? []) dataValidations.set(JSON.stringify(rule), rule)
      if (startRow === 0) {
        wide = {
          autoFilter: result.autoFilter ?? null,
          autoFilterColumns: result.autoFilterColumns ?? [],
          rowBreaks: result.rowBreaks ?? [],
          colBreaks: result.colBreaks ?? [],
          pageSetup: result.pageSetup ?? null,
        }
      }
    }
    sheets.push({
      meta,
      cells,
      rows,
      merges: [...merges.values()],
      hyperlinks,
      conditionalRules: [...conditionalRules.values()],
      dataValidations: [...dataValidations.values()],
      ...wide,
    })
  }
  return { file, sheets }
}
