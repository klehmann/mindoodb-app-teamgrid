// TeamGrid workbook document, schema version 4.
//
// One MindooDB document holds one workbook. Everything with identity is keyed
// by a stable id instead of a position, so concurrent edits from several
// people merge in Automerge without shifting each other's cells:
//
// - rows and columns are ordered by id lists (`rowOrder`, `columnOrder`);
//   deleting one leaves the id in the list and marks it with `deletedAt`
// - a cell is keyed `<rowId>:<columnId>` and patched field by field, so a
//   value change and a format change of the same cell both survive a merge
// - formula references point at row/column ids (see formula-refs.ts)
// - styles live in a workbook-wide registry keyed by a hash of their content,
//   so two people creating the same format end up with one entry
//
// Technical ids and enum-like strings are written as MindooDBAppValue.atomic()
// (see write.ts); reads return them as plain strings, which is all these
// types describe.
//
// The shapes of cell styles, values and sheet settings follow GenOffice's
// sidecar read format (vendor/.../shared/desktop-api.ts), which is what the
// editor loads from.

import type { WorkbookStyleEdit } from '@genoffice/xlsx-gateway/shared/edit-schemas'

import type { WorkbookFile } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'

export const TEAMGRID_FORM = 'teamgrid-next'
export const TEAMGRID_KIND = 'mindoodb.teamgrid.next'
export const TEAMGRID_SCHEMA_VERSION = 4

export type SheetId = string
export type RowId = string
export type ColumnId = string
export type StyleId = string
export type CellKey = `${RowId}:${ColumnId}`

export type CellScalar = string | number | boolean | null

/** GenOffice's resolved cell style, as the editor reads it from a file. */
export type CellStyle = WorkbookFile['styles'][number]
export type BorderEdge = NonNullable<CellStyle['borderTop']>

/**
 * A stored style: GenOffice's style delta on the default format, the shape
 * its save path writes. Stored in this form (not as read) because it is
 * what survives a write/read round trip unchanged, so styles keep their ids.
 */
export type StoredStyle = WorkbookStyleEdit

/** One reference inside a formula, bound to row/column ids. */
export interface FormulaRef {
  /** Absent: the sheet the formula lives on. */
  sheetId?: SheetId
  kind: 'cell' | 'range' | 'columns' | 'rows'
  startRowId?: RowId
  endRowId?: RowId
  startColumnId?: ColumnId
  endColumnId?: ColumnId
  /** `$` markers, as `[startRow, startColumn, endRow, endColumn]`. */
  absolute: [boolean, boolean, boolean, boolean]
}

export type FormulaSegment = { text: string } | { ref: FormulaRef }

export interface Formula {
  /** A1 text when it was written; informational, the segments are authoritative. */
  source: string
  segments: FormulaSegment[]
}

export interface Cell {
  value?: CellScalar
  formula?: Formula
  styleId?: StyleId
}

export interface RowMeta {
  height?: number
  customHeight?: boolean
  hidden?: boolean
  styleId?: StyleId
  deletedAt?: string
}

export interface ColumnMeta {
  width?: number
  hidden?: boolean
  styleId?: StyleId
  deletedAt?: string
}

export interface Merge {
  startRowId: RowId
  endRowId: RowId
  startColumnId: ColumnId
  endColumnId: ColumnId
}

export interface Worksheet {
  id: SheetId
  name: string
  hidden?: boolean
  tabColor?: string
  showGridLines?: boolean
  zoomScale?: number
  rightToLeft?: boolean
  frozenRows?: number
  frozenColumns?: number
  rowOrder: RowId[]
  columnOrder: ColumnId[]
  rowsById: Record<RowId, RowMeta>
  columnsById: Record<ColumnId, ColumnMeta>
  cellsById: Record<CellKey, Cell>
  /** Keyed by the merge's four corner ids, so the same merge never doubles. */
  mergesById: Record<string, Merge>
  deletedAt?: string
}

export interface Workbook {
  worksheetOrder: SheetId[]
  worksheetsById: Record<SheetId, Worksheet>
  stylesById: Record<StyleId, StoredStyle>
}

export interface TeamGridDocument {
  form: typeof TEAMGRID_FORM
  kind: typeof TEAMGRID_KIND
  subject: string
  teamgrid: {
    schemaVersion: typeof TEAMGRID_SCHEMA_VERSION
    workbook: Workbook
  }
}

export function cellKey(rowId: RowId, columnId: ColumnId): CellKey {
  return `${rowId}:${columnId}`
}

export function mergeKey(merge: Merge): string {
  return `${merge.startRowId}:${merge.startColumnId}:${merge.endRowId}:${merge.endColumnId}`
}

/** Live (not deleted) sheets in tab order, skipping duplicate entries a merge can leave. */
export function liveSheets(workbook: Workbook): Worksheet[] {
  const seen = new Set<SheetId>()
  const sheets: Worksheet[] = []
  for (const id of workbook.worksheetOrder) {
    const sheet = workbook.worksheetsById[id]
    if (!sheet || sheet.deletedAt || seen.has(id)) continue
    seen.add(id)
    sheets.push(sheet)
  }
  return sheets
}

/** Live ids of an order list, first occurrence wins (a concurrent move can duplicate one). */
export function liveIds(order: readonly string[], meta: Record<string, { deletedAt?: string }>): string[] {
  const seen = new Set<string>()
  const ids: string[] = []
  for (const id of order) {
    if (seen.has(id) || meta[id]?.deletedAt) continue
    seen.add(id)
    ids.push(id)
  }
  return ids
}

export function createId(prefix: string): string {
  return `${prefix}_${crypto.randomUUID().replace(/-/g, '').slice(0, 16)}`
}
