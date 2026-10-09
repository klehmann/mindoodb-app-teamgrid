// Builds a stored worksheet from what the editor's file holds now (a sheet in
// GenOffice's read format), given the row/column ids its positions map to.
// Used for an xlsx import (fresh ids) and after every save (ids carried over
// from the previous state through the save's structural operations).
import { parseFormula, type AxesLookup, type SheetAxes } from './formula-refs'
import {
  cellKey,
  mergeKey,
  type Cell,
  type ColumnId,
  type ColumnMeta,
  type Merge,
  type RowId,
  type RowMeta,
  type StoredStyle,
  type Worksheet,
} from './schema'
import type { SidecarSheet } from './sidecar-read'
import type { WorkbookFile } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { readSparklines, readTables } from './objects'
import { readConditionalFormats, readDataValidations } from './rules'
import { attachNotesAndLinks, readAutoFilter, readPageSetup } from './sheet-features'
import { canonicalJson, styleFields } from './styles'

export interface BuildSheetInput {
  sheet: SidecarSheet
  /** Previous stored state, or a skeleton with ids for a new sheet. */
  previous: Worksheet
  /** Raw order lists after this save's inserts, deletes and moves. */
  rowOrder: RowId[]
  columnOrder: ColumnId[]
  /** Live ids by position, aligned with the sheet's row/column indices. */
  rowIds: RowId[]
  columnIds: ColumnId[]
  /** Ids deleted by this save (tombstones). */
  deletedRowIds: ReadonlySet<RowId>
  deletedColumnIds: ReadonlySet<ColumnId>
  deletedAt: string
  /** The file the sheet was read from (differential styles, names, sheets for rules). */
  file: Pick<WorkbookFile, 'dxfStyles' | 'definedNames' | 'sheets'>
  /** Style table of the file the sheet was read from, by OOXML xf index (undefined = default). */
  styles: readonly (StoredStyle | undefined)[]
  axes: AxesLookup
}

export function buildWorksheet(input: BuildSheetInput): Worksheet {
  const { sheet, previous, rowIds, columnIds } = input
  const meta = sheet.meta
  const home: SheetAxes = {
    id: previous.id,
    name: meta.name,
    rowIds,
    columnIds,
    rowOrder: input.rowOrder,
    columnOrder: input.columnOrder,
  }

  const rowsById: Record<RowId, RowMeta> = {}
  for (const [id, row] of Object.entries(previous.rowsById)) {
    if (row.deletedAt) rowsById[id] = row
  }
  for (const id of input.deletedRowIds) rowsById[id] = { ...previous.rowsById[id], deletedAt: input.deletedAt }
  for (const record of sheet.rows) {
    const id = rowIds[record.row]
    if (!id) continue
    const row: RowMeta = {}
    if (record.height !== undefined) row.height = record.height
    if (record.customHeight) row.customHeight = true
    if (record.hidden) row.hidden = true
    Object.assign(row, styleFields(record.styleIndex === undefined ? undefined : input.styles[record.styleIndex]))
    if (Object.keys(row).length > 0) rowsById[id] = row
  }

  const columnsById: Record<ColumnId, ColumnMeta> = {}
  for (const [id, column] of Object.entries(previous.columnsById)) {
    if (column.deletedAt) columnsById[id] = column
  }
  for (const id of input.deletedColumnIds) {
    columnsById[id] = { ...previous.columnsById[id], deletedAt: input.deletedAt }
  }
  for (const span of meta.columnWidths) {
    const last = Math.min(span.endColumn, columnIds.length - 1)
    for (let index = span.startColumn; index <= last; index += 1) {
      const id = columnIds[index]!
      const column: ColumnMeta = {}
      if (span.width !== undefined) column.width = span.width
      if (span.hidden) column.hidden = true
      Object.assign(column, styleFields(span.styleIndex === undefined ? undefined : input.styles[span.styleIndex]))
      if (Object.keys(column).length > 0) columnsById[id] = column
    }
  }

  const cellsById: Worksheet['cellsById'] = {}
  // Cells of deleted rows and columns stay, so an undeleted or concurrently
  // edited row keeps its content.
  const live = new Set([...rowIds, ...columnIds])
  for (const [key, cell] of Object.entries(previous.cellsById)) {
    const [rowId, columnId] = key.split(':') as [RowId, ColumnId]
    if (!live.has(rowId) || !live.has(columnId)) cellsById[key as `${RowId}:${ColumnId}`] = cell
  }
  for (const record of sheet.cells) {
    const rowId = rowIds[record.row]
    const columnId = columnIds[record.column]
    if (!rowId || !columnId) continue
    const key = cellKey(rowId, columnId)
    const cell: Cell = {}
    if (record.formula !== undefined) {
      const formula = parseFormula(record.formula, home, input.axes)
      // Same references by id = same formula, even if its A1 text moved with
      // an inserted row: nothing is rewritten. Likewise a recalculated result
      // is not an edit: the stored value only changes with the formula, so
      // every open does not rewrite every result.
      const before = previous.cellsById[key]
      const unchanged = before?.formula && canonicalJson(before.formula) === canonicalJson(formula)
      cell.formula = unchanged ? before.formula! : formula
      const value = unchanged ? before.value : record.value
      if (value !== null && value !== undefined) cell.value = value
    } else if (record.value !== null) {
      cell.value = record.value
    }
    Object.assign(cell, styleFields(record.styleIndex === undefined ? undefined : input.styles[record.styleIndex]))
    if (Object.keys(cell).length > 0) cellsById[key] = cell
  }

  attachNotesAndLinks(cellsById, sheet, rowIds, columnIds)

  const mergesById: Record<string, Merge> = {}
  for (const area of sheet.merges) {
    const merge: Merge = {
      startRowId: rowIds[area.startRow]!,
      endRowId: rowIds[area.endRow]!,
      startColumnId: columnIds[area.startColumn]!,
      endColumnId: columnIds[area.endColumn]!,
    }
    if (merge.startRowId && merge.endRowId && merge.startColumnId && merge.endColumnId) {
      mergesById[mergeKey(merge)] = merge
    }
  }

  const next: Worksheet = {
    id: previous.id,
    name: meta.name,
    rowOrder: input.rowOrder,
    columnOrder: input.columnOrder,
    rowsById,
    columnsById,
    cellsById,
    mergesById,
    chunkOrder: previous.chunkOrder,
    visualOrder: previous.visualOrder ?? [],
    visualsById: previous.visualsById ?? {},
  }
  if (previous.viewBinding) next.viewBinding = previous.viewBinding
  const autoFilter = readAutoFilter(sheet, rowIds, columnIds)
  if (autoFilter) next.autoFilter = autoFilter
  const pageSetup = readPageSetup(sheet, rowIds, columnIds)
  if (pageSetup) next.pageSetup = pageSetup
  const conditionalFormats = readConditionalFormats(sheet, input.file.dxfStyles, rowIds, columnIds)
  if (conditionalFormats) next.conditionalFormatsById = conditionalFormats
  const dataValidations = readDataValidations(sheet, input.file, rowIds, columnIds)
  if (dataValidations) next.dataValidationsById = dataValidations
  const tables = readTables(sheet, rowIds, columnIds)
  if (tables) next.tablesById = tables
  const sparklines = readSparklines(sheet, rowIds, columnIds, input.axes)
  if (sparklines) next.sparklinesById = sparklines
  if (meta.hidden) next.hidden = true
  if (meta.tabColor) next.tabColor = meta.tabColor
  if (!meta.showGridLines) next.showGridLines = false
  if (meta.zoomScale !== undefined) next.zoomScale = meta.zoomScale
  if (meta.rightToLeft) next.rightToLeft = true
  if (meta.freeze?.frozenRows) next.frozenRows = meta.freeze.frozenRows
  if (meta.freeze?.frozenColumns) next.frozenColumns = meta.freeze.frozenColumns
  return next
}
