// TeamGrid workbook, schema version 5.
//
// A workbook is one top document plus row-block ("chunk") documents, after
// the pattern of mindoodb-word-journal's Word storage:
//
// - The top document is the workbook MindooDB lists: subject, tags, sheets
//   with their settings, column order and formats, merges, the style
//   registry, and per sheet `chunkOrder`, the order of its row blocks.
// - A chunk document (`type: "gridChunk"`, `parentId`) holds a contiguous
//   block of a sheet's rows: their order, their formats and their cells.
//   A save only touches the blocks it changed, and two people working in
//   different parts of a sheet write to different documents.
//
// Everything with identity is keyed by a stable id instead of a position, so
// concurrent edits merge in Automerge without shifting each other's cells:
//
// - rows and columns are ordered by id lists; deleting one leaves the id in
//   the list and marks it with `deletedAt`
// - a cell is keyed `<rowId>:<columnId>` and patched field by field, so a
//   value change and a format change of the same cell both survive a merge
// - formula references point at row/column ids (see formula-refs.ts)
// - styles live in a registry keyed by a hash of their content, so two people
//   creating the same format end up with one entry
// - rows and columns that only exist because the sheet grew (typing below the
//   last row) get ids derived from the id before them, so two people who
//   both type into "the next row" offline end up in the same row
//
// Every string is written as MindooDBAppValue.atomic(); reads return plain
// strings, which is all these types describe. The in-memory `Worksheet` is
// the assembled view over the top document and its chunks (store.ts).
//
// Cell styles, values and sheet settings follow GenOffice's formats
// (vendor/.../shared/desktop-api.ts), which is what the editor loads from.

import type { WorkbookStyleEdit } from '@genoffice/xlsx-gateway/shared/edit-schemas'

import type { WorkbookFile, WorkbookSaveRequest } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'

/** GenOffice's description of a visual added in the editor (save request). */
export type VisualAdd = WorkbookSaveRequest['visualAdditions'][number]
export type ChartAdd = NonNullable<VisualAdd['chart']>

export const TEAMGRID_FORM = 'teamgrid-next'
export const TEAMGRID_KIND = 'mindoodb.teamgrid.next'
export const TEAMGRID_SCHEMA_VERSION = 5
/** Form of row-block documents; workbook lists never show them. */
export const TEAMGRID_CHUNK_TYPE = 'gridChunk'

export type SheetId = string
export type RowId = string
export type ColumnId = string
export type StyleId = string
export type CellKey = `${RowId}:${ColumnId}`
/** Id of a chunk document (a MindooDB document id). */
export type ChunkId = string

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

/** One reference inside a formula, bound to row/column ids (see formula-refs.ts). */
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

/**
 * A stored formula: the text with every reference replaced by U+0001, and
 * the references in order, each packed into one string (formula-refs.ts).
 */
export interface Formula {
  t: string
  r: string[]
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

export type VisualId = string

/** Where a visual sits: its corners as row/column ids plus EMU offsets into those cells. */
export interface VisualAnchor {
  fromRowId: RowId
  fromColumnId: ColumnId
  fromRowOffset: number
  fromColumnOffset: number
  toRowId: RowId
  toColumnId: ColumnId
  toRowOffset: number
  toColumnOffset: number
}

/** A chart in GenOffice's chart-add form; series data ranges are id-bound like formulas. */
export type StoredChart = Omit<ChartAdd, 'series'> & {
  series: (Omit<ChartAdd['series'][number], 'valuesRef' | 'categoriesRef'> & {
    valuesRef?: Formula
    categoriesRef?: Formula
  })[]
}

export interface StoredVisual {
  kind: 'chart' | 'image' | 'shape'
  anchor: VisualAnchor
  chart?: StoredChart
  shape?: NonNullable<VisualAdd['shape']>
  /** The picture's bytes live in an attachment of the top document. */
  image?: { attachment: string; mediaType: 'image/png' | 'image/jpeg' | 'image/gif' }
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
  /** Charts, pictures and shapes, in drawing order (top document). */
  visualOrder: VisualId[]
  visualsById: Record<VisualId, StoredVisual>
  /** Row blocks in order (top document). */
  chunkOrder: ChunkId[]
  deletedAt?: string
}

/** Fields of a sheet that live in its chunk documents, not in the top document. */
export const CHUNK_FIELDS = ['rowOrder', 'rowsById', 'cellsById'] as const

export interface Workbook {
  worksheetOrder: SheetId[]
  worksheetsById: Record<SheetId, Worksheet>
  stylesById: Record<StyleId, StoredStyle>
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

const ID_ALPHABET = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'

/** A short random id: a one-letter prefix and 10 base-62 characters (~59 bits). */
export function createId(prefix: 'r' | 'c' | 's' | 'v'): string {
  const bytes = crypto.getRandomValues(new Uint8Array(10))
  let id = prefix
  for (const byte of bytes) id += ID_ALPHABET[byte % 62]
  return id
}

/**
 * The id of the `n`-th row/column appended after `anchor` (the last id of the
 * order list, deleted or not). Deterministic, so replicas that grow the same
 * sheet from the same state create the same ids.
 */
export function derivedId(prefix: 'r' | 'c', anchor: string, n: number): string {
  let h1 = 0xdeadbeef
  let h2 = 0x41c6ce57
  const text = `${anchor}>${n}`
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index)
    h1 = Math.imul(h1 ^ code, 2654435761)
    h2 = Math.imul(h2 ^ code, 1597334677)
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909)
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909)
  let value = 4294967296 * (2097151 & h2) + (h1 >>> 0)
  let id = ''
  for (let index = 0; index < 10; index += 1) {
    id += ID_ALPHABET[value % 62]
    value = Math.floor(value / 62)
  }
  return `${prefix}${id}`
}
