// Charts, pictures and shapes: GenOffice's read format (what the editor's
// xlsx holds) ↔ the stored form, and the stored form → GenOffice's
// "visual add" form (how an export writes them).
//
// Anchors and chart data ranges are bound to row/column ids, so a row
// inserted above a chart moves the chart and widens its ranges the same way
// it does for formulas. Stored charts keep GenOffice's chart-add shape: that
// is what an export writes and what reads back unchanged, so an untouched
// chart is never rewritten.
import { ADDABLE_SHAPE_TYPES } from '@genoffice/xlsx-gateway/shared/shape-types'

import type { WorkbookFile } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { NO_HOME, parseFormula, renderFormula, type AxesLookup, type SheetAxes } from './formula-refs'
import type { ChartAdd, StoredChart, StoredVisual, VisualAdd, VisualAnchor } from './schema'
import type { SidecarSheet } from './sidecar-read'

export type ReadVisual = WorkbookFile['visuals'][number]
type ReadChart = NonNullable<ReadVisual['chart']>

const IMAGE_TYPES = new Set(['image/png', 'image/jpeg', 'image/gif'])
const SHAPE_TYPES = new Set<string>(ADDABLE_SHAPE_TYPES)

const EMU_PER_PX = 9525
const EMU_PER_PT = 12700

/** Column widths (EMU) and row heights (EMU) of a sheet as read, by index. */
export function sheetSizes(sheet: SidecarSheet) {
  const meta = sheet.meta
  // Excel's column unit is the default font's digit width (7 px for Calibri
  // 11) plus 5 px padding; close enough to place a picture's far corner.
  const defaultColumnChars = meta.defaultColumnWidth ?? (meta.baseColumnWidth ?? 8) + 0.43
  const columnEmu = (index: number) => {
    const span = meta.columnWidths.find((entry) => index >= entry.startColumn && index <= entry.endColumn)
    const chars = span?.width ?? defaultColumnChars
    return Math.round(chars * 7 + 5) * EMU_PER_PX
  }
  const heights = new Map(sheet.rows.filter((row) => row.height !== undefined).map((row) => [row.row, row.height!]))
  const defaultRowPt = meta.defaultRowHeight || 15
  const rowEmu = (index: number) => (heights.get(index) ?? defaultRowPt) * EMU_PER_PT
  return { columnEmu, rowEmu }
}

/**
 * A file may give a picture's size as an offset running past its cell
 * (oneCellAnchor, absoluteAnchor). Stored anchors always name the cell the
 * corner is in, which is how an export writes them back (a real `<xdr:to>`).
 */
export function normalizeAnchor(visual: ReadVisual, sizes: ReturnType<typeof sheetSizes>): ReadVisual {
  if (visual.anchor.explicitTo === true) return visual
  let { toRow, toColumn, toRowOffset, toColumnOffset } = visual.anchor
  for (let guard = 0; toColumnOffset > sizes.columnEmu(toColumn) && guard < 16_384; guard += 1) {
    toColumnOffset -= sizes.columnEmu(toColumn)
    toColumn += 1
  }
  for (let guard = 0; toRowOffset > sizes.rowEmu(toRow) && guard < 1_048_576; guard += 1) {
    toRowOffset -= sizes.rowEmu(toRow)
    toRow += 1
  }
  return {
    ...visual,
    anchor: { ...visual.anchor, toRow, toColumn, toRowOffset: Math.round(toRowOffset), toColumnOffset: Math.round(toColumnOffset), explicitTo: true },
  }
}

/** The largest row and column a visual's anchor reaches, so the sheet's ids cover it. */
export function anchorExtent(visual: ReadVisual): { rows: number; columns: number } {
  return { rows: visual.anchor.toRow + 1, columns: visual.anchor.toColumn + 1 }
}

function anchorToIds(anchor: ReadVisual['anchor'], home: SheetAxes): VisualAnchor | null {
  const fromRowId = home.rowIds[anchor.fromRow]
  const fromColumnId = home.columnIds[anchor.fromColumn]
  const toRowId = home.rowIds[anchor.toRow]
  const toColumnId = home.columnIds[anchor.toColumn]
  if (!fromRowId || !fromColumnId || !toRowId || !toColumnId) return null
  return {
    fromRowId,
    fromColumnId,
    fromRowOffset: anchor.fromRowOffset,
    fromColumnOffset: anchor.fromColumnOffset,
    toRowId,
    toColumnId,
    toRowOffset: anchor.toRowOffset,
    toColumnOffset: anchor.toColumnOffset,
  }
}

/**
 * Index of an anchor corner. A deleted row/column moves the corner to the
 * next live one (forward for the top-left corner, backward for the
 * bottom-right one), like a range in a formula.
 */
function cornerIndex(live: readonly string[], order: readonly string[] | undefined, id: string, direction: 1 | -1) {
  const direct = live.indexOf(id)
  if (direct >= 0 || !order) return direct
  for (let position = order.indexOf(id) + direction; position >= 0 && position < order.length; position += direction) {
    const found = live.indexOf(order[position]!)
    if (found >= 0) return found
  }
  return -1
}

function anchorToIndices(anchor: VisualAnchor, home: SheetAxes): VisualAdd['anchor'] | null {
  const fromRow = cornerIndex(home.rowIds, home.rowOrder, anchor.fromRowId, 1)
  const fromColumn = cornerIndex(home.columnIds, home.columnOrder, anchor.fromColumnId, 1)
  const toRow = cornerIndex(home.rowIds, home.rowOrder, anchor.toRowId, -1)
  const toColumn = cornerIndex(home.columnIds, home.columnOrder, anchor.toColumnId, -1)
  if (fromRow < 0 || fromColumn < 0 || toRow < fromRow || toColumn < fromColumn) return null
  return {
    fromRow,
    fromColumn,
    fromRowOffset: anchor.fromRowOffset,
    fromColumnOffset: anchor.fromColumnOffset,
    toRow,
    toColumn,
    toRowOffset: anchor.toRowOffset,
    toColumnOffset: anchor.toColumnOffset,
  }
}

const CHART_TYPE_BY_PLOT: Record<string, ChartAdd['chartType']> = {
  lineChart: 'line',
  line3DChart: 'line',
  areaChart: 'area',
  area3DChart: 'area',
  pieChart: 'pie',
  pie3DChart: 'pie',
  ofPieChart: 'pie',
  doughnutChart: 'doughnut',
  scatterChart: 'scatter',
  radarChart: 'radar',
}

function chartType(chart: ReadChart): ChartAdd['chartType'] | null {
  const plots = [...new Set(chart.chartTypes)]
  if (plots.length > 1) return 'combo'
  const plot = plots[0]
  if (plot === 'barChart' || plot === 'bar3DChart') return chart.barDirection === 'bar' ? 'bar' : 'column'
  return plot ? (CHART_TYPE_BY_PLOT[plot] ?? null) : null
}

const HEX = /^#[0-9A-Fa-f]{6}$/

/**
 * Chart ranges are always sheet-qualified ('Sales'!$B$2:$B$9). Parsing and
 * rendering them against no home sheet keeps the sheet in both directions.
 */
const clampInt = (value: number, min: number, max: number) => Math.round(Math.min(max, Math.max(min, value)))

function storedChart(chart: ReadChart, lookup: AxesLookup): StoredChart | null {
  const type = chartType(chart)
  if (!type || chart.series.length === 0) return null
  const stored: StoredChart = {
    chartType: type,
    title: chart.title.slice(0, 255),
    series: chart.series.slice(0, 24).map((series) => {
      const entry: StoredChart['series'][number] = {
        name: series.name.slice(0, 255),
        categories: series.categories.slice(0, 1_000).map((category) => category.slice(0, 1_024)),
        values: series.values.slice(0, 1_000),
      }
      if (series.valuesRef) entry.valuesRef = parseFormula(series.valuesRef, NO_HOME, lookup)
      if (series.categoriesRef) entry.categoriesRef = parseFormula(series.categoriesRef, NO_HOME, lookup)
      if (series.color && HEX.test(series.color)) entry.color = series.color.toUpperCase()
      const pointColors = Object.fromEntries(
        (series.pointColors ?? []).filter((point) => HEX.test(point.color)).map((point) => [String(point.index), point.color]),
      )
      if (Object.keys(pointColors).length > 0) entry.pointColors = pointColors
      if (series.explosionPct !== undefined) entry.explosionPct = clampInt(series.explosionPct, 0, 400)
      const explosions = Object.fromEntries(
        (series.pointExplosions ?? []).map((point) => [String(point.index), clampInt(point.pct, 0, 400)]),
      )
      if (Object.keys(explosions).length > 0) entry.pointExplosions = explosions
      return entry
    }),
  }
  if (chart.legend) stored.legend = chart.legend
  if (chart.dataLabels) {
    stored.dataLabels = chart.dataLabels === 'category-value-percent' ? 'category-percent' : chart.dataLabels
  }
  if (chart.dataLabelPosition) stored.dataLabelPosition = chart.dataLabelPosition
  if (chart.dataLabelFormat) stored.dataLabelFormat = chart.dataLabelFormat.slice(0, 64)
  const category = chart.axisTitles?.category ?? undefined
  const value = chart.axisTitles?.value ?? undefined
  if (category || value) {
    stored.axisTitles = {
      ...(category ? { category: category.slice(0, 255) } : {}),
      ...(value ? { value: value.slice(0, 255) } : {}),
    }
  }
  if (chart.grouping) stored.grouping = chart.grouping
  if (chart.gridlines !== undefined) stored.gridlines = chart.gridlines
  if (chart.valueAxis && (chart.valueAxis.min !== undefined || chart.valueAxis.max !== undefined)) {
    stored.valueAxis = { ...chart.valueAxis }
  }
  if (chart.gapWidthPct !== undefined) stored.gapWidthPct = clampInt(chart.gapWidthPct, 0, 500)
  if (chart.holeSizePct !== undefined) stored.holeSizePct = clampInt(chart.holeSizePct, 10, 90)
  return stored
}

/**
 * The stored form of a visual the editor's file holds. Pictures carry only
 * their media type here; the caller keeps or uploads the bytes. Returns null
 * for what cannot be stored yet (OLE objects, slicers, unsupported charts).
 */
export function storedVisual(
  visual: ReadVisual,
  home: SheetAxes,
  lookup: AxesLookup,
  attachment: string,
): StoredVisual | null {
  const anchor = anchorToIds(visual.anchor, home)
  if (!anchor) return null
  if (visual.kind === 'chart' && visual.chart) {
    const chart = storedChart(visual.chart, lookup)
    return chart ? { kind: 'chart', anchor, chart } : null
  }
  if (visual.kind === 'image' && visual.mediaType && IMAGE_TYPES.has(visual.mediaType)) {
    return {
      kind: 'image',
      anchor,
      image: { attachment, mediaType: visual.mediaType as NonNullable<StoredVisual['image']>['mediaType'] },
    }
  }
  if (visual.kind === 'shape') {
    const shapeType = visual.shapeType && SHAPE_TYPES.has(visual.shapeType) ? visual.shapeType : 'rect'
    const shape: StoredVisual['shape'] = { shapeType: shapeType as NonNullable<StoredVisual['shape']>['shapeType'] }
    if (visual.fillColor && HEX.test(visual.fillColor)) shape.fillColor = visual.fillColor.toUpperCase()
    if (visual.text) shape.text = visual.text.slice(0, 1_000)
    return { kind: 'shape', anchor, shape }
  }
  return null
}

/** A stored visual as GenOffice's save request adds it (images need their bytes as base64). */
export function visualAddition(
  visual: StoredVisual,
  home: SheetAxes,
  lookup: AxesLookup,
  imageBase64: string | undefined,
): Omit<VisualAdd, 'sheetId'> | null {
  const anchor = anchorToIndices(visual.anchor, home)
  if (!anchor) return null
  if (visual.kind === 'chart' && visual.chart) {
    const chart = {
      ...visual.chart,
      series: visual.chart.series.map(({ valuesRef, categoriesRef, ...series }) => ({
        ...series,
        ...(valuesRef ? { valuesRef: renderFormula(valuesRef, NO_HOME, lookup) } : {}),
        ...(categoriesRef ? { categoriesRef: renderFormula(categoriesRef, NO_HOME, lookup) } : {}),
      })),
    } as ChartAdd
    return { anchor, chart }
  }
  if (visual.kind === 'image' && visual.image && imageBase64) {
    return { anchor, image: { mediaType: visual.image.mediaType, base64: imageBase64 } }
  }
  if (visual.kind === 'shape' && visual.shape) return { anchor, shape: visual.shape }
  return null
}
