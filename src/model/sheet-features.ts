// Sheet features beyond cells: notes, hyperlinks, the autoFilter and print
// settings. Read from the editor's file in GenOffice's read format, stored
// with row/column ids instead of positions (so they move with inserted rows),
// and written back in GenOffice's save format on every open.
import type { WorkbookSaveRequest } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { cellKey, type ColumnId, type HeaderFooter, type IdArea, type PageSetup, type RowId, type Worksheet } from './schema'
import type { CellArea, SidecarSheet } from './sidecar-read'

type PageSetupState = WorkbookSaveRequest['pageSetupStates'][number]

/** An area by ids; undefined when a corner lies outside the sheet's rows/columns. */
export function areaToIds(area: CellArea, rowIds: readonly RowId[], columnIds: readonly ColumnId[]): IdArea | undefined {
  const startRowId = rowIds[area.startRow]
  const endRowId = rowIds[area.endRow]
  const startColumnId = columnIds[area.startColumn]
  const endColumnId = columnIds[area.endColumn]
  return startRowId && endRowId && startColumnId && endColumnId
    ? { startRowId, startColumnId, endRowId, endColumnId }
    : undefined
}

/** An id area at its current positions; undefined when a corner was deleted. */
export function areaToIndices(
  area: IdArea,
  rowIndex: ReadonlyMap<RowId, number>,
  columnIndex: ReadonlyMap<ColumnId, number>,
): CellArea | undefined {
  const startRow = rowIndex.get(area.startRowId)
  const endRow = rowIndex.get(area.endRowId)
  const startColumn = columnIndex.get(area.startColumnId)
  const endColumn = columnIndex.get(area.endColumnId)
  return startRow !== undefined && endRow !== undefined && startColumn !== undefined && endColumn !== undefined
    ? { startRow, endRow, startColumn, endColumn }
    : undefined
}

function columnIndexOf(label: string): number {
  let index = 0
  for (const char of label.toUpperCase()) index = index * 26 + (char.charCodeAt(0) - 64)
  return index - 1
}

function columnLabel(index: number): string {
  let label = ''
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) label = String.fromCharCode(65 + ((n - 1) % 26)) + label
  return label
}

/** `'Daten'!$A$1:$D$20` (first range of a list) → its area. */
export function parsePrintArea(text: string): CellArea | undefined {
  const first = text.split(',')[0] ?? ''
  const match = /\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i.exec(first.replace(/^.*!/, ''))
  if (!match) return undefined
  const startColumn = columnIndexOf(match[1]!)
  const startRow = Number(match[2]) - 1
  return {
    startRow,
    startColumn,
    endRow: match[4] ? Number(match[4]) - 1 : startRow,
    endColumn: match[3] ? columnIndexOf(match[3]) : startColumn,
  }
}

/** The repeated rows of `'Daten'!$1:$2` or `'Daten'!$A:$B,'Daten'!$1:$2`, 0-based. */
export function parsePrintTitleRows(text: string): { start: number; end: number } | undefined {
  for (const part of text.split(',')) {
    const match = /\$?(\d+):\$?(\d+)$/.exec(part.replace(/^.*!/, ''))
    if (match) return { start: Number(match[1]) - 1, end: Number(match[2]) - 1 }
  }
  return undefined
}

/** `&LLeft&C&P of &N&RRight` → its sections (`&&` is a literal ampersand). */
export function parseHeaderFooter(text: string): HeaderFooter {
  const parts: HeaderFooter = {}
  let section: keyof HeaderFooter = 'center'
  let current = ''
  const flush = () => {
    if (current) parts[section] = (parts[section] ?? '') + current
    current = ''
  }
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index]!
    const next = text[index + 1]
    if (char === '&' && (next === 'L' || next === 'C' || next === 'R')) {
      flush()
      section = next === 'L' ? 'left' : next === 'C' ? 'center' : 'right'
      index += 1
      continue
    }
    if (char === '&' && next !== undefined) {
      current += char + next
      index += 1
      continue
    }
    current += char
  }
  flush()
  return parts
}

/** Inches → the margin preset GenOffice can write (normal 0.7", wide 1", narrow 0.25"). */
function marginPreset(left: number): NonNullable<PageSetup['margins']> {
  if (left <= 0.4) return 'narrow'
  if (left >= 0.9) return 'wide'
  return 'normal'
}

/** The sheet's print settings, bound to ids; undefined when it has none. */
export function readPageSetup(
  sheet: SidecarSheet,
  rowIds: readonly RowId[],
  columnIds: readonly ColumnId[],
): PageSetup | undefined {
  const read = sheet.pageSetup
  const setup: PageSetup = {}
  if (read) {
    if (read.orientation) setup.orientation = read.orientation
    if (read.paperSize !== undefined && read.paperSize <= 118) setup.paperSize = read.paperSize
    if (read.scale !== undefined) setup.scale = read.scale
    if (read.fitToWidth !== undefined) setup.fitToWidth = Math.min(read.fitToWidth, 1_000)
    if (read.fitToHeight !== undefined) setup.fitToHeight = Math.min(read.fitToHeight, 1_000)
    if (read.fitToPage !== undefined) setup.fitToPage = read.fitToPage
    if (read.margins) setup.margins = marginPreset(read.margins.left)
    if (read.printGridlines !== undefined) setup.printGridlines = read.printGridlines
    if (read.printHeadings !== undefined) setup.printHeadings = read.printHeadings
    if (read.oddHeader) setup.header = parseHeaderFooter(read.oddHeader)
    if (read.oddFooter) setup.footer = parseHeaderFooter(read.oddFooter)
  }
  const area = sheet.meta.printArea ? parsePrintArea(sheet.meta.printArea) : undefined
  const printArea = area ? areaToIds(area, rowIds, columnIds) : undefined
  if (printArea) setup.printArea = printArea
  const titles = sheet.meta.printTitles ? parsePrintTitleRows(sheet.meta.printTitles) : undefined
  if (titles && rowIds[titles.start] && rowIds[titles.end]) {
    setup.printTitleRows = { startRowId: rowIds[titles.start]!, endRowId: rowIds[titles.end]! }
  }
  const rowBreaks = (sheet.rowBreaks ?? []).map((index) => rowIds[index]).filter((id): id is RowId => !!id)
  if (rowBreaks.length) setup.rowBreaks = rowBreaks
  const colBreaks = (sheet.colBreaks ?? []).map((index) => columnIds[index]).filter((id): id is ColumnId => !!id)
  if (colBreaks.length) setup.colBreaks = colBreaks
  return Object.keys(setup).length > 0 ? setup : undefined
}

/** Notes and hyperlinks onto the cells they belong to (cells are created as needed). */
export function attachNotesAndLinks(
  cellsById: Worksheet['cellsById'],
  sheet: SidecarSheet,
  rowIds: readonly RowId[],
  columnIds: readonly ColumnId[],
): void {
  const at = (row: number, column: number) => {
    const rowId = rowIds[row]
    const columnId = columnIds[column]
    if (!rowId || !columnId) return undefined
    const key = cellKey(rowId, columnId)
    return (cellsById[key] ??= {})
  }
  for (const comment of sheet.meta.comments ?? []) {
    const cell = at(comment.row, comment.column)
    if (cell) cell.note = { author: comment.author, text: comment.text }
  }
  for (const link of sheet.hyperlinks ?? []) {
    const cell = at(link.row, link.column)
    if (cell && link.target) cell.link = link.target
  }
}

/** The autoFilter, bound to ids. */
export function readAutoFilter(
  sheet: SidecarSheet,
  rowIds: readonly RowId[],
  columnIds: readonly ColumnId[],
): Worksheet['autoFilter'] {
  if (!sheet.autoFilter) return undefined
  const area = areaToIds(sheet.autoFilter, rowIds, columnIds)
  return area ? { area, columns: [...(sheet.autoFilterColumns ?? [])] } : undefined
}

/** Print settings as GenOffice's save writes them; undefined when there is nothing to write. */
export function pageSetupState(
  setup: PageSetup | undefined,
  sheetId: string,
  rowIndex: ReadonlyMap<RowId, number>,
  columnIndex: ReadonlyMap<ColumnId, number>,
): PageSetupState | undefined {
  if (!setup) return undefined
  const { printArea, printTitleRows, rowBreaks, colBreaks, ...plain } = setup
  const state: Record<string, unknown> = { sheetId, ...plain }
  const area = printArea ? areaToIndices(printArea, rowIndex, columnIndex) : undefined
  if (area) {
    state.printArea = `${columnLabel(area.startColumn)}${area.startRow + 1}:${columnLabel(area.endColumn)}${area.endRow + 1}`
  }
  const start = printTitleRows ? rowIndex.get(printTitleRows.startRowId) : undefined
  const end = printTitleRows ? rowIndex.get(printTitleRows.endRowId) : undefined
  if (start !== undefined && end !== undefined) state.printTitles = `${start + 1}:${end + 1}`
  const breaks = (ids: readonly string[] | undefined, index: ReadonlyMap<string, number>) =>
    (ids ?? []).map((id) => index.get(id)).filter((position): position is number => position !== undefined && position > 0)
  const rows = breaks(rowBreaks, rowIndex)
  if (rows.length) state.rowBreaks = rows
  const columns = breaks(colBreaks, columnIndex)
  if (columns.length) state.colBreaks = columns
  return Object.keys(state).length > 1 ? (state as PageSetupState) : undefined
}
