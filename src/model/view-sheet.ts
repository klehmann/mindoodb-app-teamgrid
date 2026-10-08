// Sheets filled from a MindooDB virtual view (classic TeamGrid's "virtual
// view sheet"): a bold header row with the view's visible column titles,
// then one row per view entry — categories on a light fill, documents
// below them. A refresh rewrites the sheet; rows and columns keep their ids
// by position, so formats of the sheet's rows and columns survive.
import type {
  MindooDBAppResolvedViewDefinition,
  MindooDBAppViewEntry,
  MindooDBAppViewNavigator,
} from 'mindoodb-app-sdk'

import {
  cellKey,
  createId,
  liveIds,
  type Cell,
  type CellScalar,
  type ColumnId,
  type ColumnMeta,
  type RowId,
  type RowMeta,
  type SheetId,
  type ViewBinding,
  type Worksheet,
} from './schema'
import { styleFields } from './styles'

/** Fill of category rows: document content, the same in both UI themes. */
const CATEGORY_FILL = '#F3F4F6'
const PAGE_SIZE = 1000

export interface ViewSheetSettings {
  name: string
  viewId: string
  showDocuments: boolean
  showCategories: boolean
  rootCategoryPath: string[]
}

/** `Customers\ACME` → ['Customers', 'ACME'] (the category separator Haven uses). */
export function parseCategoryPath(input: string): string[] {
  return input
    .split('\\')
    .map((part) => part.trim())
    .filter(Boolean)
}

export function viewTitle(view: MindooDBAppResolvedViewDefinition): string {
  return view.description?.trim() || view.id
}

type ViewColumn = MindooDBAppResolvedViewDefinition['columns'][number]

const COUNT_OF: Record<string, (entry: MindooDBAppViewEntry) => number> = {
  childCount: (entry) => entry.childCount ?? (entry.childCategoryCount ?? 0) + (entry.childDocumentCount ?? 0),
  childCategoryCount: (entry) => entry.childCategoryCount ?? 0,
  childDocumentCount: (entry) => entry.childDocumentCount ?? 0,
  descendantCount: (entry) =>
    entry.descendantCount ?? (entry.descendantDocumentCount ?? 0) + (entry.descendantCategoryCount ?? 0),
  descendantCategoryCount: (entry) => entry.descendantCategoryCount ?? 0,
  descendantDocumentCount: (entry) => entry.descendantDocumentCount ?? 0,
  siblingCount: (entry) => entry.siblingCount ?? 0,
}

function columnValue(entry: MindooDBAppViewEntry, column: ViewColumn): unknown {
  if (entry.kind === 'document' && column.role === 'category') return null
  const expression = column.expression
  if (expression.mode === 'formula' && expression.expression.kind === 'operation') {
    const count = COUNT_OF[(expression.expression as { op: string }).op]
    if (count) return count(entry)
  }
  return entry.columnValues[column.name]
}

/** A view value as a cell: numbers, text, dates as date serials with a date format. */
export function cellOf(value: unknown): Cell | undefined {
  if (value === null || value === undefined || value === '') return undefined
  if (typeof value === 'number') return Number.isFinite(value) ? { value } : undefined
  if (typeof value === 'boolean') return { value }
  if (typeof value === 'string') {
    const trimmed = value.trim()
    if (/^\d{4}-\d{2}-\d{2}(?:T\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:\d{2})?)?$/.test(trimmed)) {
      const time = Date.parse(trimmed.length === 10 ? `${trimmed}T00:00:00Z` : trimmed)
      if (Number.isFinite(time)) {
        const withTime = trimmed.length > 10
        return {
          value: time / 86_400_000 + 25_569,
          ...styleFields({ numberFormat: withTime ? 'yyyy-mm-dd hh:mm' : 'yyyy-mm-dd' }),
        }
      }
    }
    return { value }
  }
  if (Array.isArray(value)) return cellOf(value.map((item) => String(item)).join(', '))
  try {
    return { value: JSON.stringify(value) }
  } catch {
    return { value: String(value) }
  }
}

export interface ViewTable {
  header: string[]
  rows: { cells: (Cell | undefined)[]; category: boolean }[]
}

/** The view's entries as a table, as the sheet shows them. */
export function viewTable(
  view: MindooDBAppResolvedViewDefinition,
  entries: readonly MindooDBAppViewEntry[],
  settings: Pick<ViewSheetSettings, 'showDocuments' | 'showCategories'>,
): ViewTable {
  const columns = view.columns.filter((column) => !column.hidden)
  return {
    header: columns.map((column) => column.title.trim() || column.name),
    rows: entries
      .filter((entry) => (entry.kind === 'category' ? settings.showCategories : settings.showDocuments))
      .map((entry) => ({
        cells: columns.map((column) => cellOf(columnValue(entry, column))),
        category: entry.kind === 'category',
      })),
  }
}

/** Reads every entry of a view (expanded), page by page. */
export async function readViewEntries(navigator: MindooDBAppViewNavigator): Promise<MindooDBAppViewEntry[]> {
  await navigator.expandAll()
  const entries: MindooDBAppViewEntry[] = []
  let startPosition: string | null = null
  do {
    const page = await navigator.entriesForward({ limit: PAGE_SIZE, startPosition })
    entries.push(...page.entries)
    startPosition = page.nextPosition
  } while (startPosition)
  return entries
}

/** `count` ids for an axis: the live ones in order first, new ones after; the rest are deleted. */
function axisIds<T extends RowMeta | ColumnMeta>(
  order: readonly string[],
  meta: Record<string, T>,
  count: number,
  prefix: 'r' | 'c',
  deletedAt: string,
) {
  const live = liveIds(order, meta)
  const ids = Array.from({ length: count }, (_, index) => live[index] ?? createId(prefix))
  const byId: Record<string, T> = {}
  for (const [id, entry] of Object.entries(meta)) if (entry.deletedAt) byId[id] = entry
  for (const id of live.slice(count)) byId[id] = { ...meta[id], deletedAt } as T
  ids.slice(0, live.length).forEach((id) => {
    if (meta[id]) byId[id] = meta[id]!
  })
  return { order: [...order, ...ids.slice(live.length)], ids, byId }
}

/** Roughly what a cell shows, for sizing its column. */
function displayText(cell: Cell | undefined): string {
  if (!cell || cell.value === undefined || cell.value === null) return ''
  const format = cell['s.numberFormat']
  if (typeof format === 'string') return format
  return String(cell.value)
}

/** The sheet with the table written into it (a new sheet when `previous` is undefined). */
export function buildViewSheet(
  previous: Worksheet | undefined,
  id: SheetId,
  name: string,
  binding: ViewBinding,
  table: ViewTable,
): Worksheet {
  const deletedAt = binding.lastRefreshedAt
  const rows = axisIds<RowMeta>(previous?.rowOrder ?? [], previous?.rowsById ?? {}, table.rows.length + 1, 'r', deletedAt)
  const columns = axisIds<ColumnMeta>(
    previous?.columnOrder ?? [],
    previous?.columnsById ?? {},
    Math.max(1, table.header.length),
    'c',
    deletedAt,
  )
  const cellsById: Worksheet['cellsById'] = {}
  const put = (rowId: RowId, columnId: ColumnId, cell: Cell | undefined) => {
    if (cell && Object.keys(cell).length > 0) cellsById[cellKey(rowId, columnId)] = cell
  }
  table.header.forEach((title, index) =>
    put(rows.ids[0]!, columns.ids[index]!, { value: title as CellScalar, ...styleFields({ bold: true }) }),
  )
  table.rows.forEach((row, rowIndex) => {
    const rowId = rows.ids[rowIndex + 1]!
    columns.ids.forEach((columnId, index) => {
      const cell = row.cells[index]
      put(rowId, columnId, row.category ? { ...cell, ...styleFields({ fillColor: CATEGORY_FILL }) } : cell)
    })
  })
  // Columns nobody sized yet fit their content (character widths, as xlsx counts them).
  columns.ids.forEach((columnId, index) => {
    if (columns.byId[columnId]?.width !== undefined) return
    const texts = [table.header[index] ?? '', ...table.rows.map((row) => displayText(row.cells[index]))]
    const width = Math.min(50, Math.max(8, ...texts.map((text) => text.length + 2)))
    columns.byId[columnId] = { ...columns.byId[columnId], width }
  })
  const sheet: Worksheet = {
    ...(previous ?? {
      chunkOrder: [],
      visualOrder: [],
      visualsById: {},
      frozenRows: 1,
    }),
    id,
    name,
    viewBinding: binding,
    rowOrder: rows.order,
    rowsById: rows.byId,
    columnOrder: columns.order,
    columnsById: columns.byId,
    cellsById,
    mergesById: {},
  } as Worksheet
  delete sheet.deletedAt
  return sheet
}
