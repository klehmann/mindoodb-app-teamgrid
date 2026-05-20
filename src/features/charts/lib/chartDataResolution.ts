import {
  type CellValue,
  type Chart,
  type ChartSeries,
  type SeriesRange,
} from "@/features/document/lib/teamgridDocument";
import { getProjectionForWorksheet, type FormulaContext } from "@/features/formulas/lib";
import { getCell } from "@/features/grid/lib/gridProjection";
import { formatChartRange } from "@/features/charts/lib/chartRangeReferences";
import type { ResolvedChartData, ResolvedChartSeries } from "@/features/charts/lib/types";

export function resolveChartData(chart: Chart, context: FormulaContext): ResolvedChartData {
  const labels = chart.categoryAxis
    ? readRangeValues(chart.categoryAxis, context).flat().map(formatLabel)
    : inferDefaultLabels(chart, context);
  const series = chart.series.map((item, index) => resolveSeries(chart, item, context, index));
  const maxLength = Math.max(labels.length, ...series.map((item) => item.values.length), 0);
  return {
    chart,
    labels: Array.from({ length: maxLength }, (_, index) => labels[index] ?? String(index + 1)),
    series: series.map((item) => ({
      ...item,
      values: Array.from({ length: maxLength }, (_, index) => item.values[index] ?? 0),
    })),
  };
}

export function readRangeValues(range: SeriesRange, context: FormulaContext): CellValue[][] {
  const worksheet = context.workbook.worksheetsById[range.worksheetId];
  if (!worksheet || worksheet.deletedAt) {
    return [];
  }
  const projection = getProjectionForWorksheet(range.worksheetId, context);
  const startRow = projection.rowIndexById.get(range.startRowId);
  const endRow = projection.rowIndexById.get(range.endRowId);
  const startColumn = projection.columnIndexById.get(range.startColumnId);
  const endColumn = projection.columnIndexById.get(range.endColumnId);
  if (startRow == null || endRow == null || startColumn == null || endColumn == null) {
    return [];
  }
  const top = Math.min(startRow, endRow);
  const bottom = Math.max(startRow, endRow);
  const left = Math.min(startColumn, endColumn);
  const right = Math.max(startColumn, endColumn);
  const values: CellValue[][] = [];
  for (let rowIndex = top; rowIndex <= bottom; rowIndex += 1) {
    const row = projection.rows[rowIndex];
    const rowValues: CellValue[] = [];
    for (let columnIndex = left; columnIndex <= right; columnIndex += 1) {
      const column = projection.columns[columnIndex];
      if (row && column) {
        rowValues.push(getCell(worksheet, row.id, column.id).value);
      }
    }
    values.push(rowValues);
  }
  return values;
}

export function formatSeriesRange(range: SeriesRange, context: FormulaContext) {
  return formatChartRange(range, context);
}

function resolveSeries(chart: Chart, series: ChartSeries, context: FormulaContext, index: number): ResolvedChartSeries {
  return {
    id: series.id,
    name: resolveSeriesName(series, context, index),
    values: readRangeValues(series.values, context).flat().map(numberValue),
    color: series.color ?? chart.style?.colors?.[index],
  };
}

function resolveSeriesName(series: ChartSeries, context: FormulaContext, index: number) {
  if (typeof series.name === "string" && series.name.trim()) {
    return series.name;
  }
  if (series.name && typeof series.name === "object") {
    const value = readRangeValues(series.name, context).flat()[0];
    const label = formatLabel(value);
    if (label) {
      return label;
    }
  }
  return `Series ${index + 1}`;
}

function inferDefaultLabels(chart: Chart, context: FormulaContext) {
  const firstSeries = chart.series[0];
  if (!firstSeries) {
    return [];
  }
  return readRangeValues(firstSeries.values, context).flat().map((_, index) => String(index + 1));
}

function numberValue(value: CellValue) {
  if (value.kind === "number") {
    return value.value;
  }
  if (value.kind === "date") {
    return new Date(value.isoDate).getTime();
  }
  if (value.kind === "string") {
    const parsed = Number(value.text);
    return Number.isFinite(parsed) ? parsed : 0;
  }
  return 0;
}

function formatLabel(value: CellValue | undefined) {
  if (!value || value.kind === "empty") {
    return "";
  }
  if (value.kind === "string") {
    return value.text;
  }
  if (value.kind === "number") {
    return String(value.value);
  }
  return new Date(value.isoDate).toLocaleDateString();
}
