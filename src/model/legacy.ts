// TeamGrid 1.x workbooks (form `teamgrid`, schema 3) in the current shape.
//
// 1.x already keyed rows, columns and cells by id and bound formula
// references to those ids, so the conversion keeps every id: sheets, rows
// and columns of the copy are the original's. What changes is the shape —
// values and formats in GenOffice's terms, formulas as text with marks plus
// packed references, charts as chart-add visuals with id-bound data ranges.
// The original document is only read.
import type { WorkbookStyleEdit } from '@genoffice/xlsx-gateway/shared/edit-schemas'

import { createAxesLookup, packRef, parseFormula, type SheetAxes } from './formula-refs'
import {
  cellKey,
  liveIds,
  type Cell,
  nameKey,
  type CellScalar,
  type ColumnMeta,
  type DefinedName,
  type Formula,
  type FormulaRef,
  type RowMeta,
  type StoredChart,
  type StoredVisual,
  type ViewBinding,
  type Workbook,
  type Worksheet,
} from './schema'
import { styleFields } from './styles'

export const LEGACY_FORM = 'teamgrid'

// ── TeamGrid 1.x (src/features/document/lib/teamgridDocument.ts on main) ──

type LegacyReference =
  | { kind: 'cell'; worksheetId: string; rowId: string; columnId: string }
  | { kind: 'range'; worksheetId: string; startRowId: string; endRowId: string; startColumnId: string; endColumnId: string }
  | { kind: 'column'; worksheetId: string; columnId: string }

type LegacyValue =
  | { kind: 'empty' }
  | { kind: 'string'; text: string; excelNumFmt?: string }
  | { kind: 'number'; value: number; format?: string; currencyCode?: string; excelNumFmt?: string }
  | { kind: 'date'; isoDate: string; format?: string; excelNumFmt?: string }

type LegacyResult =
  | { kind: 'empty' }
  | { kind: 'string'; value: string }
  | { kind: 'number'; value: number }
  | { kind: 'date'; isoDate: string }
  | { kind: 'error'; code: string }

interface LegacyBorder {
  style: string
  color?: string
}

interface LegacyStyle {
  textColor?: string
  backgroundColor?: string
  fontFamily?: string
  fontSize?: number
  bold?: boolean
  italic?: boolean
  underline?: boolean
  horizontalAlign?: string
  verticalAlign?: string
  wrapText?: boolean
  indent?: number
  borders?: Partial<Record<'top' | 'right' | 'bottom' | 'left', LegacyBorder>>
}

interface LegacyCell {
  rowId: string
  columnId: string
  value?: LegacyValue
  formula?: {
    source: string
    segments?: ({ kind: 'text'; text: string } | { kind: 'reference'; reference: LegacyReference })[]
    cached?: LegacyResult
  }
  style?: LegacyStyle
}

interface LegacySeriesRange {
  worksheetId: string
  startRowId: string
  endRowId: string
  startColumnId: string
  endColumnId: string
}

interface LegacyChart {
  type: 'column' | 'bar' | 'line' | 'pie'
  title?: string
  series: { name?: string | LegacySeriesRange; values: LegacySeriesRange; color?: string }[]
  categoryAxis?: LegacySeriesRange
  anchor: {
    from: { rowId: string; columnId: string; rowOffsetEmu: number; colOffsetEmu: number }
    to: { rowId: string; columnId: string; rowOffsetEmu: number; colOffsetEmu: number }
  }
  legend?: { position: 'right' | 'bottom' | 'top' | 'left' | 'none' }
  style?: { colors?: string[]; showGridlines?: boolean }
  deletedAt?: string
}

interface LegacyWorksheet {
  id: string
  title: string
  rowOrder: string[]
  columnOrder: string[]
  rowsById: Record<string, { height?: number; hidden?: boolean; defaultStyle?: LegacyStyle; deletedAt?: string }>
  columnsById: Record<string, { width?: number; hidden?: boolean; defaultStyle?: LegacyStyle; deletedAt?: string }>
  cellsById: Record<string, LegacyCell>
  chartOrder?: string[]
  chartsById?: Record<string, LegacyChart>
  viewBinding?: Omit<ViewBinding, 'lastRefreshedAt'> & { lastRefreshedAt?: string }
  deletedAt?: string
}

export interface LegacyWorkbook {
  worksheetOrder: string[]
  worksheetsById: Record<string, LegacyWorksheet>
  /** Defined names; in 1.x they sit next to the workbook (`teamgrid.namedExpressionsById`). */
  namedExpressionsById?: Record<string, { name: string; reference: LegacyReference }>
}

/** The 1.x workbook of a document's data, or undefined when it is something else. */
export function legacyWorkbookOf(data: Record<string, unknown>): LegacyWorkbook | undefined {
  if (data.form !== LEGACY_FORM) return undefined
  const teamgrid = data.teamgrid as
    | { workbook?: LegacyWorkbook; namedExpressionsById?: LegacyWorkbook['namedExpressionsById'] }
    | undefined
  if (!teamgrid?.workbook?.worksheetsById) return undefined
  return { ...teamgrid.workbook, namedExpressionsById: teamgrid.namedExpressionsById ?? {} }
}

/**
 * 1.x's tab order as it read it (resolveWorksheetOrder): every id once, unknown
 * ids dropped, sheets the list misses appended in id order.
 */
function worksheetOrder(legacy: LegacyWorkbook): string[] {
  const order = [...new Set(legacy.worksheetOrder)].filter((id) => legacy.worksheetsById[id])
  const missing = Object.keys(legacy.worksheetsById)
    .filter((id) => !order.includes(id))
    .sort()
  return [...order, ...missing]
}

// ── conversion ─────────────────────────────────────────────────────────

/** 1.x sizes are screen pixels; xlsx counts column widths in characters and row heights in points. */
const PIXELS_PER_CHARACTER = 7
const POINTS_PER_PIXEL = 0.75

const NUMBER_FORMATS: Record<string, string> = { integer: '0', decimal: '0.00', percent: '0.00%' }
const DATE_FORMATS: Record<string, string> = { dateTime: 'mmm d, yyyy h:mm AM/PM', time: 'h:mm AM/PM' }
const BORDER_STYLES = new Set(['thin', 'medium', 'thick', 'dashed', 'dotted', 'double'])

function hexColor(color: string | undefined): string | undefined {
  const value = color?.trim() ?? ''
  const short = /^#([0-9a-f]{3})$/i.exec(value)
  if (short) return `#${[...short[1]!].map((digit) => digit + digit).join('').toUpperCase()}`
  return /^#[0-9a-f]{6}$/i.test(value) ? value.toUpperCase() : undefined
}

function serial(isoDate: string): number | undefined {
  const time = Date.parse(isoDate)
  return Number.isFinite(time) ? time / 86_400_000 + 25_569 : undefined
}

export function legacyStyle(style: LegacyStyle | undefined): WorkbookStyleEdit {
  const edit: Record<string, unknown> = {}
  if (!style) return edit as WorkbookStyleEdit
  if (style.bold) edit.bold = true
  if (style.italic) edit.italic = true
  if (style.underline) edit.underline = true
  if (style.wrapText) edit.wrapText = true
  if (style.fontFamily) edit.fontFamily = style.fontFamily
  if (style.fontSize && style.fontSize > 0) edit.fontSize = style.fontSize
  const fontColor = hexColor(style.textColor)
  if (fontColor) edit.fontColor = fontColor
  const fillColor = hexColor(style.backgroundColor)
  if (fillColor) edit.fillColor = fillColor
  if (style.horizontalAlign && style.horizontalAlign !== 'general') edit.horizontalAlignment = style.horizontalAlign
  if (style.verticalAlign) edit.verticalAlignment = style.verticalAlign === 'middle' ? 'center' : style.verticalAlign
  if (style.indent && style.indent > 0) edit.indent = Math.min(15, Math.round(style.indent))
  for (const [side, key] of [['top', 'borderTop'], ['bottom', 'borderBottom'], ['left', 'borderLeft'], ['right', 'borderRight']] as const) {
    const border = style.borders?.[side]
    if (!border || !BORDER_STYLES.has(border.style)) continue
    const color = hexColor(border.color)
    edit[key] = color ? { style: border.style, color } : { style: border.style }
  }
  return edit as WorkbookStyleEdit
}

function valueOf(value: LegacyValue | LegacyResult | undefined): { value?: CellScalar; numberFormat?: string } {
  if (!value) return {}
  switch (value.kind) {
    case 'string':
      return 'text' in value
        ? { value: value.text, ...(value.excelNumFmt ? { numberFormat: value.excelNumFmt } : {}) }
        : { value: value.value }
    case 'number': {
      if (!Number.isFinite(value.value)) return {}
      if (!('format' in value || 'excelNumFmt' in value)) return { value: value.value }
      const format =
        value.excelNumFmt ??
        (value.format === 'currency'
          ? value.currencyCode === 'EUR' ? '€0.00' : '$0.00'
          : value.format ? NUMBER_FORMATS[value.format] : undefined)
      return { value: value.value, ...(format ? { numberFormat: format } : {}) }
    }
    case 'date': {
      const number = serial(value.isoDate)
      if (number === undefined) return {}
      const format = 'excelNumFmt' in value && value.excelNumFmt
        ? value.excelNumFmt
        : ('format' in value && value.format && DATE_FORMATS[value.format]) || 'mmm d, yyyy'
      return { value: number, numberFormat: format }
    }
    case 'error':
      return { value: value.code }
    default:
      return {}
  }
}

function refOf(reference: LegacyReference, homeId: string): FormulaRef {
  const base = reference.worksheetId === homeId ? {} : { sheetId: reference.worksheetId }
  const absolute: FormulaRef['absolute'] = [false, false, false, false]
  if (reference.kind === 'cell') {
    return { ...base, kind: 'cell', startRowId: reference.rowId, startColumnId: reference.columnId, absolute }
  }
  if (reference.kind === 'column') {
    return { ...base, kind: 'columns', startColumnId: reference.columnId, endColumnId: reference.columnId, absolute }
  }
  return {
    ...base,
    kind: 'range',
    startRowId: reference.startRowId,
    startColumnId: reference.startColumnId,
    endRowId: reference.endRowId,
    endColumnId: reference.endColumnId,
    absolute,
  }
}

function rangeRef(range: LegacySeriesRange): Formula {
  const single = range.startRowId === range.endRowId && range.startColumnId === range.endColumnId
  return {
    t: '\u0001',
    r: [
      packRef({
        sheetId: range.worksheetId,
        kind: single ? 'cell' : 'range',
        startRowId: range.startRowId,
        startColumnId: range.startColumnId,
        ...(single ? {} : { endRowId: range.endRowId, endColumnId: range.endColumnId }),
        absolute: [true, true, true, true],
      }),
    ],
  }
}

/** Converts a 1.x workbook; `now` stamps view sheets that never recorded a refresh. */
export function convertLegacyWorkbook(legacy: LegacyWorkbook, now = new Date().toISOString()): Workbook {
  const sheets = worksheetOrder(legacy).map((id) => legacy.worksheetsById[id]!)
  const axes: SheetAxes[] = sheets.map((sheet) => ({
    id: sheet.id,
    name: sheet.title,
    rowIds: liveIds(sheet.rowOrder, sheet.rowsById),
    columnIds: liveIds(sheet.columnOrder, sheet.columnsById),
    rowOrder: sheet.rowOrder,
    columnOrder: sheet.columnOrder,
  }))
  const lookup = createAxesLookup(axes)

  /** Cell values by sheet for chart caches: what 1.x showed (cached formula results included). */
  const shown = (sheet: LegacyWorksheet, rowId: string, columnId: string): CellScalar | undefined => {
    const cell = sheet.cellsById[cellKey(rowId, columnId)]
    if (!cell) return undefined
    return valueOf(cell.formula ? cell.formula.cached : cell.value).value
  }
  /** Category labels as 1.x showed them: dates as dates, not as their serial numbers. */
  const label = (sheet: LegacyWorksheet, rowId: string, columnId: string): string => {
    const cell = sheet.cellsById[cellKey(rowId, columnId)]
    const value = cell?.formula ? cell.formula.cached : cell?.value
    if (value?.kind === 'date') return value.isoDate.slice(0, value.isoDate.includes('T00:00:00') ? 10 : 16).replace('T', ' ')
    return String(shown(sheet, rowId, columnId) ?? '')
  }
  const rangeValues = (range: LegacySeriesRange): CellScalar[] => rangeCells(range, shown).map((value) => value ?? null)
  const rangeLabels = (range: LegacySeriesRange): string[] => rangeCells(range, label)
  function rangeCells<T>(range: LegacySeriesRange, read: (sheet: LegacyWorksheet, rowId: string, columnId: string) => T): T[] {
    const sheet = legacy.worksheetsById[range.worksheetId]
    const home = axes.find((entry) => entry.id === range.worksheetId)
    if (!sheet || !home) return []
    const rows = home.rowIds.slice(home.rowIds.indexOf(range.startRowId), home.rowIds.indexOf(range.endRowId) + 1)
    const columns = home.columnIds.slice(
      home.columnIds.indexOf(range.startColumnId),
      home.columnIds.indexOf(range.endColumnId) + 1,
    )
    return rows.flatMap((rowId) => columns.map((columnId) => read(sheet, rowId, columnId)))
  }

  const worksheetsById: Workbook['worksheetsById'] = {}
  sheets.forEach((legacySheet, index) => {
    const home = axes[index]!
    const rowsById: Record<string, RowMeta> = {}
    for (const [id, row] of Object.entries(legacySheet.rowsById)) {
      const meta: RowMeta = { ...styleFields(legacyStyle(row.defaultStyle)) }
      if (row.height !== undefined) {
        meta.height = Math.round(row.height * POINTS_PER_PIXEL * 4) / 4
        meta.customHeight = true
      }
      if (row.hidden) meta.hidden = true
      if (row.deletedAt) meta.deletedAt = row.deletedAt
      if (Object.keys(meta).length > 0) rowsById[id] = meta
    }
    const columnsById: Record<string, ColumnMeta> = {}
    for (const [id, column] of Object.entries(legacySheet.columnsById)) {
      const meta: ColumnMeta = { ...styleFields(legacyStyle(column.defaultStyle)) }
      if (column.width !== undefined) meta.width = Math.max(1, Math.round((column.width / PIXELS_PER_CHARACTER) * 4) / 4)
      if (column.hidden) meta.hidden = true
      if (column.deletedAt) meta.deletedAt = column.deletedAt
      if (Object.keys(meta).length > 0) columnsById[id] = meta
    }

    const cellsById: Worksheet['cellsById'] = {}
    for (const legacyCell of Object.values(legacySheet.cellsById)) {
      const shownValue = valueOf(legacyCell.formula ? legacyCell.formula.cached : legacyCell.value)
      const ownFormat = legacyCell.formula ? valueOf(legacyCell.value).numberFormat : undefined
      const style: WorkbookStyleEdit = legacyStyle(legacyCell.style)
      const numberFormat = ownFormat ?? shownValue.numberFormat
      if (numberFormat && numberFormat !== 'General') style.numberFormat = numberFormat
      const cell: Cell = { ...styleFields(style) }
      if (legacyCell.formula) {
        const segments = legacyCell.formula.segments
        if (segments?.length) {
          let t = ''
          const r: string[] = []
          for (const segment of segments) {
            if (segment.kind === 'text') t += segment.text
            else {
              t += '\u0001'
              r.push(packRef(refOf(segment.reference, legacySheet.id)))
            }
          }
          cell.formula = { t: t.replace(/^=/, ''), r }
        } else {
          cell.formula = parseFormula(legacyCell.formula.source.replace(/^=/, ''), home, lookup)
        }
      }
      if (shownValue.value !== undefined && shownValue.value !== null && shownValue.value !== '') cell.value = shownValue.value
      if (Object.keys(cell).length > 0) cellsById[cellKey(legacyCell.rowId, legacyCell.columnId)] = cell
    }

    const visualOrder: string[] = []
    const visualsById: Record<string, StoredVisual> = {}
    for (const id of legacySheet.chartOrder ?? []) {
      const chart = legacySheet.chartsById?.[id]
      if (!chart || chart.deletedAt || chart.series.length === 0) continue
      const categories = chart.categoryAxis ? rangeLabels(chart.categoryAxis) : []
      const stored: StoredChart = {
        chartType: chart.type,
        title: chart.title ?? '',
        series: chart.series.slice(0, 24).map((series, position) => {
          const values = rangeValues(series.values).map((value) => (typeof value === 'number' ? value : 0))
          const name =
            typeof series.name === 'string'
              ? series.name
              : series.name
                ? String(rangeValues(series.name)[0] ?? '')
                : `Series ${position + 1}`
          const color = hexColor(series.color ?? chart.style?.colors?.[position])
          return {
            name,
            categories: categories.length ? categories.slice(0, values.length) : values.map((_, index) => String(index + 1)),
            values,
            valuesRef: rangeRef(series.values),
            ...(chart.categoryAxis ? { categoriesRef: rangeRef(chart.categoryAxis) } : {}),
            ...(color ? { color } : {}),
          }
        }),
        ...(chart.legend ? { legend: chart.legend.position } : {}),
        ...(chart.style?.showGridlines !== undefined ? { gridlines: chart.style.showGridlines } : {}),
      }
      visualOrder.push(id)
      visualsById[id] = {
        kind: 'chart',
        chart: stored,
        anchor: {
          fromRowId: chart.anchor.from.rowId,
          fromColumnId: chart.anchor.from.columnId,
          fromRowOffset: Math.round(chart.anchor.from.rowOffsetEmu),
          fromColumnOffset: Math.round(chart.anchor.from.colOffsetEmu),
          toRowId: chart.anchor.to.rowId,
          toColumnId: chart.anchor.to.columnId,
          toRowOffset: Math.round(chart.anchor.to.rowOffsetEmu),
          toColumnOffset: Math.round(chart.anchor.to.colOffsetEmu),
        },
      }
    }

    const sheet: Worksheet = {
      id: legacySheet.id,
      name: legacySheet.title,
      rowOrder: [...legacySheet.rowOrder],
      columnOrder: [...legacySheet.columnOrder],
      rowsById,
      columnsById,
      cellsById,
      mergesById: {},
      chunkOrder: [],
      visualOrder,
      visualsById,
    }
    if (legacySheet.viewBinding) {
      const { viewId, viewTitle, showDocuments, showCategories, rootCategoryPath, lastRefreshedAt } = legacySheet.viewBinding
      sheet.viewBinding = {
        viewId,
        viewTitle,
        showDocuments,
        showCategories,
        rootCategoryPath: [...(rootCategoryPath ?? [])],
        lastRefreshedAt: lastRefreshedAt ?? now,
      }
    }
    if (legacySheet.deletedAt) sheet.deletedAt = legacySheet.deletedAt
    worksheetsById[sheet.id] = sheet
  })
  // Defined names: always sheet-qualified and absolute, as Excel writes them.
  const namesById: Record<string, DefinedName> = {}
  for (const named of Object.values(legacy.namedExpressionsById ?? {})) {
    if (!named?.name || !named.reference || !legacy.worksheetsById[named.reference.worksheetId]) continue
    const ref = refOf(named.reference, '')
    ref.absolute = [true, true, true, true]
    namesById[nameKey(named.name)] = { name: named.name, formula: { t: '\u0001', r: [packRef(ref)] } }
  }
  return { worksheetOrder: sheets.map((sheet) => sheet.id), worksheetsById, namesById }
}
