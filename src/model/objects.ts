// Tables (Excel ListObjects) and sparklines. Read from the file, kept with
// row/column ids (sparkline sources as id-bound formulas, like chart ranges)
// and written back as the gateway's table and sparkline additions. What the
// gateway cannot create is not kept: custom table styles fall back to
// Excel's default, totals rows and column stripes are not written.
import type { WorkbookSaveRequest } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { NO_HOME, parseFormula, renderFormula, type AxesLookup, type SheetAxes } from './formula-refs'
import type { ColumnId, Formula, IdArea, RowId } from './schema'
import { areaToIds, areaToIndices } from './sheet-features'
import type { SidecarSheet } from './sidecar-read'
import { canonicalJson } from './styles'

export interface StoredTable {
  name: string
  area: IdArea
  columnNames: string[]
  /** A built-in style (TableStyleMedium2, …); absent = Excel's default. */
  style?: string
  bandedRows: boolean
}

export interface StoredSparklineGroup {
  type: 'line' | 'column' | 'stacked'
  color?: string
  cells: { rowId: RowId; columnId: ColumnId; source: Formula }[]
}

const BUILT_IN_STYLE = /^TableStyle(?:Light|Medium|Dark)[1-9][0-9]?$/

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

function hash(text: string): string {
  let h1 = 0x811c9dc5
  let h2 = 0x01000193
  for (let index = 0; index < text.length; index += 1) {
    h1 = Math.imul(h1 ^ text.charCodeAt(index), 16777619) >>> 0
    h2 = Math.imul(h2 ^ text.charCodeAt(index), 2246822519) >>> 0
  }
  return `s${h1.toString(36)}${h2.toString(36)}`
}

/** The sheet's tables, keyed by lower-case name. */
export function readTables(
  sheet: SidecarSheet,
  rowIds: readonly RowId[],
  columnIds: readonly ColumnId[],
): Record<string, StoredTable> | undefined {
  const tables: Record<string, StoredTable> = {}
  for (const table of sheet.meta.tables ?? []) {
    if (!table.name || table.headerRowCount !== 1) continue
    const area = areaToIds(table.range, rowIds, columnIds)
    if (!area) continue
    const width = table.range.endColumn - table.range.startColumn + 1
    const columnNames = Array.from({ length: width }, (_, index) => table.columns?.[index] ?? `Column${index + 1}`)
    tables[table.name.toLowerCase()] = {
      name: table.name,
      area,
      columnNames,
      ...(table.styleName && BUILT_IN_STYLE.test(table.styleName) ? { style: table.styleName } : {}),
      bandedRows: table.showRowStripes,
    }
  }
  return Object.keys(tables).length ? tables : undefined
}

/** The sheet's sparkline groups, keyed by content. */
export function readSparklines(
  sheet: SidecarSheet,
  rowIds: readonly RowId[],
  columnIds: readonly ColumnId[],
  lookup: AxesLookup,
): Record<string, StoredSparklineGroup> | undefined {
  const groups: Record<string, StoredSparklineGroup> = {}
  for (const group of sheet.meta.sparklines ?? []) {
    const cells: StoredSparklineGroup['cells'] = []
    for (const entry of group.cells) {
      const match = /^\$?([A-Z]{1,3})\$?(\d+)$/i.exec(entry.cell)
      const rowId = match ? rowIds[Number(match[2]) - 1] : undefined
      const columnId = match ? columnIds[columnIndexOf(match[1]!)] : undefined
      if (!rowId || !columnId) continue
      cells.push({ rowId, columnId, source: parseFormula(entry.sourceRef.replace(/^=/, ''), NO_HOME, lookup) })
    }
    if (!cells.length) continue
    const stored: StoredSparklineGroup = { type: group.type, cells, ...(group.color ? { color: group.color } : {}) }
    groups[hash(canonicalJson(stored))] = stored
  }
  return Object.keys(groups).length ? groups : undefined
}

export function tableAdditions(
  tables: Record<string, StoredTable> | undefined,
  sheetId: string,
  rowIndex: ReadonlyMap<RowId, number>,
  columnIndex: ReadonlyMap<ColumnId, number>,
): WorkbookSaveRequest['tableAdditions'] {
  return Object.values(tables ?? {}).flatMap((table) => {
    const area = areaToIndices(table.area, rowIndex, columnIndex)
    if (!area) return []
    const width = area.endColumn - area.startColumn + 1
    const columnNames = Array.from({ length: width }, (_, index) => table.columnNames[index] ?? `Column${index + 1}`)
    return [
      {
        sheetId,
        area,
        name: table.name,
        columnNames,
        ...(table.style ? { style: table.style } : {}),
        bandedRows: table.bandedRows,
      } as WorkbookSaveRequest['tableAdditions'][number],
    ]
  })
}

export function sparklineAdditions(
  groups: Record<string, StoredSparklineGroup> | undefined,
  sheetId: string,
  home: SheetAxes,
  lookup: AxesLookup,
): WorkbookSaveRequest['sparklineAdditions'] {
  const rowIndex = new Map(home.rowIds.map((id, index) => [id, index]))
  const columnIndex = new Map(home.columnIds.map((id, index) => [id, index]))
  return Object.values(groups ?? {}).flatMap((group) => {
    const cells = group.cells.flatMap((cell) => {
      const row = rowIndex.get(cell.rowId)
      const column = columnIndex.get(cell.columnId)
      const sourceRef = renderFormula(cell.source, NO_HOME, lookup)
      if (row === undefined || column === undefined || sourceRef.includes('#REF!')) return []
      return [{ cell: `${columnLabel(column)}${row + 1}`, sourceRef }]
    })
    if (!cells.length) return []
    return [
      {
        sheetId,
        type: group.type,
        ...(group.color && /^#[0-9a-f]{6}$/i.test(group.color) ? { color: group.color } : {}),
        cells: cells.slice(0, 500),
      } as WorkbookSaveRequest['sparklineAdditions'][number],
    ]
  })
}
