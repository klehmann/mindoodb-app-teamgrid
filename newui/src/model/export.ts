// Stored workbook → xlsx bytes.
//
// The xlsx is built the way GenOffice's CLI writes new files: start from the
// gateway's blank workbook and save the whole content onto it as if it had
// just been typed in. Sheets are created in a first pass, because the
// gateway refuses sheet operations mixed with other edits in some cases and
// the second pass needs the new sheets' ids.
import { blankXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'

import type {
  WorkbookCellEdit,
  WorkbookFile,
  WorkbookPageSetupState,
  WorkbookSaveRequest,
  WorkbookStructuralOp,
} from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { saveWorkbookInBrowser } from '../xlsx/browser-save'
import type { XlsxEngine } from '../xlsx/engine'
import { createAxesLookup, renderFormula, type SheetAxes } from './formula-refs'
import { liveIds, liveSheets, type VisualId, type Workbook, type Worksheet } from './schema'
import { visualAddition } from './visuals'

const BLANK_SHEET_NAME = 'Sheet1'

export function emptySaveRequest(sessionId: string): WorkbookSaveRequest {
  return {
    sessionId,
    mode: 'save',
    edits: [],
    structuralOps: [],
    chartEdits: [],
    visualEdits: [],
    visualAdditions: [],
    tableAdditions: [],
    pivotAdditions: [],
    sheetOps: [],
    sheetOrder: [],
    filterStates: [],
    hyperlinkEdits: [],
    cfStates: [],
    dvStates: [],
    pageSetupStates: [],
    noteStates: [],
    formulaValues: [],
    pivotCacheRefreshPaths: [],
    pivotRefreshUpdates: [],
    sheetProtections: [],
    sparklineAdditions: [],
    definedNamesState: null,
    themeState: null,
    workbookProtectionState: null,
    protectedRangeStates: [],
  } as unknown as WorkbookSaveRequest
}

/** Live axes of every sheet, for rendering formulas. */
export function workbookAxes(workbook: Workbook): SheetAxes[] {
  return liveSheets(workbook).map((sheet) => sheetAxes(sheet))
}

export function sheetAxes(sheet: Worksheet): SheetAxes {
  return {
    id: sheet.id,
    name: sheet.name,
    rowIds: liveIds(sheet.rowOrder, sheet.rowsById),
    columnIds: liveIds(sheet.columnOrder, sheet.columnsById),
    rowOrder: sheet.rowOrder,
    columnOrder: sheet.columnOrder,
  }
}

async function sheetIdsByName(engine: XlsxEngine, bytes: Uint8Array): Promise<Map<string, string>> {
  const opened = (await engine.open(bytes, 'en')) as WorkbookFile
  await engine.close(opened.sessionId)
  return new Map(opened.sheets.map((sheet) => [sheet.name, sheet.id]))
}

export interface ExportOptions {
  /** Base64 bytes of a picture's attachment; pictures without bytes are left out. */
  loadImage?: (attachment: string) => Promise<string | undefined>
}

export interface ExportResult {
  bytes: Uint8Array
  /** Per stored sheet, the visuals written, in drawing order (how the file lists them). */
  visualOrder: Map<string, VisualId[]>
}

export async function workbookToXlsx(workbook: Workbook, engine: XlsxEngine, options: ExportOptions = {}): Promise<ExportResult> {
  const sheets = liveSheets(workbook)
  let bytes = new Uint8Array(await blankXlsxBuffer(BLANK_SHEET_NAME))

  // Pass 1: sheets.
  const blankIds = await sheetIdsByName(engine, bytes)
  const blankId = blankIds.get(BLANK_SHEET_NAME)!
  const structure = emptySaveRequest('export')
  const order: string[] = []
  sheets.forEach((sheet, index) => {
    if (index === 0) {
      if (sheet.name !== BLANK_SHEET_NAME) {
        structure.sheetOps.push({ kind: 'rename-sheet', sheetId: blankId, newName: sheet.name })
      }
      order.push(blankId)
    } else {
      const id = `new-${index}`
      structure.sheetOps.push({ kind: 'add-sheet', sheetId: id, name: sheet.name })
      order.push(id)
    }
  })
  sheets.forEach((sheet, index) => {
    if (sheet.hidden) structure.sheetOps.push({ kind: 'set-sheet-hidden', sheetId: order[index]!, hidden: true })
  })
  structure.sheetOrder = order
  if (structure.sheetOps.length > 0) {
    bytes = new Uint8Array((await saveWorkbookInBrowser(bytes, new Map([[blankId, BLANK_SHEET_NAME]]), structure)).bytes)
  }

  // Pass 2: everything inside the sheets.
  const fileIds = await sheetIdsByName(engine, bytes)
  const axes = workbookAxes(workbook)
  const lookup = createAxesLookup(axes)
  const content = emptySaveRequest('export')
  const visualOrder = new Map<string, VisualId[]>()
  for (const sheet of sheets) {
    const fileSheetId = fileIds.get(sheet.name)
    if (!fileSheetId) throw new Error(`Export lost sheet ${sheet.name}.`)
    const home = lookup.bySheetId(sheet.id)!
    appendSheet(content, workbook, sheet, fileSheetId, home, lookup)
    const written: VisualId[] = []
    for (const id of new Set(sheet.visualOrder ?? [])) {
      const visual = sheet.visualsById?.[id]
      if (!visual) continue
      const image = visual.image ? await options.loadImage?.(visual.image.attachment) : undefined
      const addition = visualAddition(visual, home, lookup, image)
      if (!addition) continue
      content.visualAdditions.push({ sheetId: fileSheetId, ...addition } as WorkbookSaveRequest['visualAdditions'][number])
      written.push(id)
    }
    visualOrder.set(sheet.id, written)
  }
  const names = new Map([...fileIds].map(([name, id]) => [id, name]))
  return { bytes: (await saveWorkbookInBrowser(bytes, names, content)).bytes, visualOrder }
}

function appendSheet(
  request: WorkbookSaveRequest,
  workbook: Workbook,
  sheet: Worksheet,
  sheetId: string,
  home: SheetAxes,
  lookup: ReturnType<typeof createAxesLookup>,
): void {
  const rowIndex = new Map(home.rowIds.map((id, index) => [id, index]))
  const columnIndex = new Map(home.columnIds.map((id, index) => [id, index]))
  const styleEdit = (styleId: string | undefined) => (styleId ? workbook.stylesById[styleId] : undefined)

  const edits: WorkbookCellEdit[] = request.edits
  for (const [key, cell] of Object.entries(sheet.cellsById)) {
    const [rowId, columnId] = key.split(':') as [string, string]
    const row = rowIndex.get(rowId)
    const column = columnIndex.get(columnId)
    if (row === undefined || column === undefined) continue
    const style = styleEdit(cell.styleId)
    const formula = cell.formula ? renderFormula(cell.formula, home, lookup) : undefined
    const hasContent = formula !== undefined || (cell.value !== undefined && cell.value !== null)
    if (!hasContent && !style) continue
    edits.push({
      sheetId,
      row,
      column,
      writeValue: hasContent,
      value: cell.value ?? null,
      ...(formula === undefined ? {} : { formula: formula.startsWith('=') ? formula : `=${formula}` }),
      ...(style ? { style } : {}),
    } as WorkbookCellEdit)
  }

  const ops: WorkbookStructuralOp[] = request.structuralOps
  pushRuns(home.rowIds.length, (index) => sheet.rowsById[home.rowIds[index]!]?.height, (start, end, size) =>
    ops.push({ sheetId, kind: 'set-row-size', start, end, size }),
  )
  pushRuns(home.columnIds.length, (index) => sheet.columnsById[home.columnIds[index]!]?.width, (start, end, size) =>
    ops.push({ sheetId, kind: 'set-col-size', start, end, size }),
  )
  pushRuns(home.rowIds.length, (index) => (sheet.rowsById[home.rowIds[index]!]?.hidden ? 1 : undefined), (start, end) =>
    ops.push({ sheetId, kind: 'set-rows-hidden', start, end, hidden: true }),
  )
  pushRuns(home.columnIds.length, (index) => (sheet.columnsById[home.columnIds[index]!]?.hidden ? 1 : undefined), (start, end) =>
    ops.push({ sheetId, kind: 'set-cols-hidden', start, end, hidden: true }),
  )
  for (const merge of Object.values(sheet.mergesById)) {
    const range = {
      startRow: rowIndex.get(merge.startRowId),
      endRow: rowIndex.get(merge.endRowId),
      startColumn: columnIndex.get(merge.startColumnId),
      endColumn: columnIndex.get(merge.endColumnId),
    }
    if (Object.values(range).some((value) => value === undefined)) continue
    ops.push({ sheetId, kind: 'merge-cells', range: range as Required<{ [K in keyof typeof range]: number }> })
  }

  const view: Omit<WorkbookPageSetupState, 'sheetId'> = {}
  if (sheet.frozenRows || sheet.frozenColumns) {
    view.frozenRows = sheet.frozenRows ?? 0
    view.frozenColumns = sheet.frozenColumns ?? 0
  }
  if (sheet.showGridLines === false) view.showGridlines = false
  if (sheet.zoomScale !== undefined) view.zoomScale = sheet.zoomScale
  if (Object.keys(view).length > 0) request.pageSetupStates.push({ sheetId, ...view })
}

/** Calls `emit` for each run of equal defined values. */
function pushRuns(
  length: number,
  valueAt: (index: number) => number | undefined,
  emit: (start: number, end: number, value: number) => void,
): void {
  let start = 0
  while (start < length) {
    const value = valueAt(start)
    let end = start
    while (end + 1 < length && valueAt(end + 1) === value) end += 1
    if (value !== undefined) emit(start, end, value)
    start = end + 1
  }
}
