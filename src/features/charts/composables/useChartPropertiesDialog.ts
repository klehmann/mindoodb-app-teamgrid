import { computed, ref, type Ref } from "vue";
import { parseChartRangeReference, formatChartRange } from "@/features/charts/lib/chartRangeReferences";
import {
  createId,
  type Chart,
  type ChartLegend,
  type ChartSeries,
  type ChartStyle,
  type ChartType,
  type SeriesRange,
  type TwoCellAnchor,
  type Worksheet,
} from "@/features/document/lib/teamgridDocument";
import type { TeamGridAppApi } from "@/features/document/composables/useTeamGridDocument";
import type { TeamGridOperation } from "@/features/document/lib/teamgridOps";
import type { FormulaContext } from "@/features/formulas/lib";

export interface ChartSeriesDraft {
  id: string;
  name: string;
  values: string;
  color: string;
}

interface IndexedRange {
  worksheetId: string;
  minRow: number;
  maxRow: number;
  minCol: number;
  maxCol: number;
}

export interface UseChartPropertiesDialogOptions {
  app: TeamGridAppApi;
  activeWorksheet: Readonly<Ref<Worksheet | null>>;
  formulaContext: Readonly<Ref<FormulaContext | null>>;
}

export function useChartPropertiesDialog(options: UseChartPropertiesDialogOptions) {
  const { app, activeWorksheet, formulaContext } = options;

  const selectedChartId = ref<string | null>(null);
  const chartDialogVisible = ref(false);
  const chartDraftTitle = ref("");
  const chartDraftType = ref<ChartType>("column");
  const chartDraftLegendPosition = ref<ChartLegend["position"]>("right");
  const chartDraftShowGridlines = ref(true);
  const chartDraftDataRange = ref("");
  const chartDraftDataRangeSource = ref("");
  const chartDraftColors = ref("");
  const chartDraftSeries = ref<ChartSeriesDraft[]>([]);
  const chartDraftError = ref<string | null>(null);

  const selectedChart = computed(() => {
    if (!activeWorksheet.value || !selectedChartId.value) {
      return null;
    }
    const chart = activeWorksheet.value.chartsById[selectedChartId.value];
    return chart && !chart.deletedAt ? chart : null;
  });

  function selectChart(chartId: string | null) {
    selectedChartId.value = chartId;
  }

  function openChartProperties(chartId: string) {
    const chart = activeWorksheet.value?.chartsById[chartId];
    const context = formulaContext.value;
    if (!chart || chart.deletedAt || !context) {
      return;
    }
    selectedChartId.value = chartId;
    chartDraftTitle.value = chart.title ?? "";
    chartDraftType.value = chart.type;
    chartDraftLegendPosition.value = chart.legend?.position ?? "right";
    chartDraftShowGridlines.value = chart.style?.showGridlines !== false;
    chartDraftDataRange.value = formatChartDataRange(chart, context) ?? "";
    chartDraftDataRangeSource.value = chartDraftDataRange.value;
    chartDraftColors.value = (chart.style?.colors ?? []).join("\n");
    chartDraftSeries.value = chart.series.map((series, index) => ({
      id: series.id || createId("series"),
      name: typeof series.name === "string"
        ? series.name
        : series.name && !sameRange(series.name, chart.categoryAxis)
          ? formatChartRange(series.name, context)
          : "",
      values: formatChartRange(series.values, context),
      color: series.color ?? chart.style?.colors?.[index] ?? "",
    }));
    chartDraftError.value = null;
    chartDialogVisible.value = true;
  }

  function applyChartProperties() {
    const worksheet = activeWorksheet.value;
    const chartId = selectedChartId.value;
    const context = formulaContext.value;
    if (!worksheet || !chartId || app.gridReadOnly.value || !context) {
      chartDialogVisible.value = false;
      return;
    }
    const existing = worksheet.chartsById[chartId];
    if (!existing || existing.deletedAt) {
      chartDialogVisible.value = false;
      return;
    }
    const dataRangeChanged = chartDraftDataRange.value.trim() !== chartDraftDataRangeSource.value.trim();
    const parsedDrafts = parseSeriesDraftMetadata(worksheet.id, context);
    const rebuilt = dataRangeChanged
      ? buildSeriesFromDataRange(chartDraftDataRange.value, existing, parsedDrafts, worksheet.id, context)
      : null;
    if (dataRangeChanged && !rebuilt) {
      chartDraftError.value = "Enter a valid chart data range, for example Sheet 1!A1:C5.";
      return;
    }
    const nextSeries = rebuilt?.series ?? parseSeriesDrafts(worksheet.id, context);
    if (!nextSeries) {
      return;
    }
    const style = buildStyleDraft();
    const legend = chartDraftLegendPosition.value === "none"
      ? { position: "none" } satisfies ChartLegend
      : { position: chartDraftLegendPosition.value } satisfies ChartLegend;
    const title = chartDraftTitle.value.trim() || undefined;

    app.updateGrid((grid) => {
      const targetWorksheet = grid.workbook.worksheetsById[worksheet.id];
      const chart = targetWorksheet.chartsById[chartId];
      if (!chart || chart.deletedAt) {
        return [];
      }
      const operations: TeamGridOperation[] = [];
      if (chart.title !== title) {
        chart.title = title;
        operations.push({ type: "setChartTitle", worksheetId: targetWorksheet.id, chartId, title });
      }
      if (chart.type !== chartDraftType.value) {
        chart.type = chartDraftType.value;
        operations.push({ type: "setChartType", worksheetId: targetWorksheet.id, chartId, chartType: chartDraftType.value });
      }
      chart.series = nextSeries;
      operations.push({ type: "setChartSeries", worksheetId: targetWorksheet.id, chartId, series: nextSeries });
      if (dataRangeChanged && rebuilt) {
        if (rebuilt.categoryAxis) {
          chart.categoryAxis = rebuilt.categoryAxis;
          operations.push({ type: "setChartCategoryAxis", worksheetId: targetWorksheet.id, chartId, categoryAxis: rebuilt.categoryAxis });
        } else {
          delete chart.categoryAxis;
          operations.push({ type: "setChartCategoryAxis", worksheetId: targetWorksheet.id, chartId, categoryAxis: undefined });
        }
        chartDraftDataRangeSource.value = chartDraftDataRange.value.trim();
      }
      chart.legend = legend;
      operations.push({ type: "setChartLegend", worksheetId: targetWorksheet.id, chartId, legend });
      chart.style = style;
      operations.push({ type: "setChartStyle", worksheetId: targetWorksheet.id, chartId, style });
      return operations;
    });
    chartDialogVisible.value = false;
    chartDraftError.value = null;
  }

  function removeSelectedChart() {
    const worksheet = activeWorksheet.value;
    const chartId = selectedChartId.value;
    if (!worksheet || !chartId || app.gridReadOnly.value) {
      chartDialogVisible.value = false;
      return;
    }
    app.updateGrid((grid) => {
      const targetWorksheet = grid.workbook.worksheetsById[worksheet.id];
      const chart = targetWorksheet.chartsById[chartId];
      if (!chart || chart.deletedAt) {
        return [];
      }
      const deletedAt = new Date().toISOString();
      chart.deletedAt = deletedAt;
      return [{ type: "removeChart", worksheetId: targetWorksheet.id, chartId, deletedAt }];
    });
    selectedChartId.value = null;
    chartDialogVisible.value = false;
  }

  function addSeriesDraft() {
    chartDraftSeries.value = [
      ...chartDraftSeries.value,
      { id: createId("series"), name: "", values: "", color: "" },
    ];
  }

  function removeSeriesDraft(seriesId: string) {
    if (chartDraftSeries.value.length <= 1) {
      return;
    }
    chartDraftSeries.value = chartDraftSeries.value.filter((series) => series.id !== seriesId);
  }

  function setChartAnchor(chartId: string, anchor: TwoCellAnchor) {
    const worksheet = activeWorksheet.value;
    if (!worksheet || app.gridReadOnly.value) {
      return;
    }
    app.updateGrid((grid) => {
      const targetWorksheet = grid.workbook.worksheetsById[worksheet.id];
      const chart = targetWorksheet.chartsById[chartId];
      if (!chart || chart.deletedAt) {
        return [];
      }
      chart.anchor = anchor;
      return [{ type: "setChartAnchor", worksheetId: targetWorksheet.id, chartId, anchor }];
    });
  }

  function parseSeriesDrafts(worksheetId: string, context: FormulaContext): ChartSeries[] | null {
    const drafts = parseSeriesDraftMetadata(worksheetId, context);
    const series = drafts.flatMap((draft) => {
      const values = parseChartRangeReference(draft.values, worksheetId, context);
      if (!values) {
        return [];
      }
      return [toChartSeries(draft.id || createId("series"), values, draft.name, draft.color)];
    });
    if (series.length !== chartDraftSeries.value.length || series.length === 0) {
      chartDraftError.value = "Every series needs a valid values range, for example Sheet 1!B2:B5.";
      return null;
    }
    chartDraftError.value = null;
    return series;
  }

  function parseSeriesDraftMetadata(worksheetId: string, context: FormulaContext) {
    return chartDraftSeries.value.map((draft) => ({
      ...draft,
      name: parseSeriesName(draft.name, worksheetId, context),
      color: normalizeColor(draft.color) || undefined,
    }));
  }

  function parseSeriesName(name: string, worksheetId: string, context: FormulaContext): string | SeriesRange | undefined {
    const trimmed = name.trim();
    if (!trimmed) {
      return undefined;
    }
    return parseChartRangeReference(trimmed, worksheetId, context) ?? trimmed;
  }

  function buildStyleDraft(): ChartStyle | undefined {
    const colors = chartDraftColors.value
      .split(/[\n,]/)
      .map((color) => normalizeColor(color))
      .filter((color): color is string => Boolean(color));
    const style: ChartStyle = {
      showGridlines: chartDraftShowGridlines.value,
      colors: colors.length > 0 ? colors : undefined,
    };
    return style.showGridlines === true && !style.colors ? undefined : style;
  }

  return {
    selectedChartId,
    selectedChart,
    chartDialogVisible,
    chartDraftTitle,
    chartDraftType,
    chartDraftLegendPosition,
    chartDraftShowGridlines,
    chartDraftDataRange,
    chartDraftColors,
    chartDraftSeries,
    chartDraftError,
    selectChart,
    openChartProperties,
    applyChartProperties,
    removeSelectedChart,
    addSeriesDraft,
    removeSeriesDraft,
    setChartAnchor,
  };
}

function normalizeColor(value: string) {
  const trimmed = value.trim();
  return /^#[0-9a-f]{6}$/i.test(trimmed) ? trimmed : "";
}

function toChartSeries(
  id: string,
  values: SeriesRange,
  name?: string | SeriesRange,
  color?: string,
): ChartSeries {
  return {
    id,
    values,
    ...(name !== undefined ? { name } : {}),
    ...(color ? { color } : {}),
  };
}

function formatChartDataRange(chart: Chart, context: FormulaContext) {
  const ranges = chart.series.flatMap((series) => [
    series.values,
    series.name && typeof series.name === "object" ? series.name : null,
  ]).filter((range): range is SeriesRange => Boolean(range));
  if (chart.categoryAxis) {
    ranges.unshift(chart.categoryAxis);
  }
  const indexed = ranges.map((range) => indexedRange(range, context));
  if (indexed.some((range) => !range)) {
    return null;
  }
  const [first] = indexed as IndexedRange[];
  if (!first || indexed.some((range) => range?.worksheetId !== first.worksheetId)) {
    return null;
  }
  const projection = context.projectionsByWorksheetId.get(first.worksheetId);
  if (!projection) {
    return null;
  }
  const combined = rangeFromIndexes(
    first.worksheetId,
    Math.min(...indexed.map((range) => range!.minRow)),
    Math.min(...indexed.map((range) => range!.minCol)),
    Math.max(...indexed.map((range) => range!.maxRow)),
    Math.max(...indexed.map((range) => range!.maxCol)),
    context,
  );
  return combined ? formatChartRange(combined, context) : null;
}

function buildSeriesFromDataRange(
  source: string,
  chart: Chart,
  draftMetadata: Array<{ id: string; name: string | SeriesRange | undefined; color: string | undefined }>,
  currentWorksheetId: string,
  context: FormulaContext,
): { series: ChartSeries[]; categoryAxis?: SeriesRange } | null {
  const trimmed = source.trim();
  if (!trimmed) {
    return { series: chart.series };
  }
  const range = parseChartRangeReference(trimmed, currentWorksheetId, context);
  const indexed = range ? indexedRange(range, context) : null;
  if (!range || !indexed || indexed.minRow === indexed.maxRow || indexed.minCol === indexed.maxCol) {
    return null;
  }
  return usesRowsAsSeries(chart, context)
    ? buildRowSeriesFromDataRange(indexed, chart, draftMetadata, context)
    : buildColumnSeriesFromDataRange(indexed, chart, draftMetadata, context);
}

function buildRowSeriesFromDataRange(
  range: IndexedRange,
  chart: Chart,
  draftMetadata: Array<{ id: string; name: string | SeriesRange | undefined; color: string | undefined }>,
  context: FormulaContext,
) {
  const categoryAxis = rangeFromIndexes(range.worksheetId, range.minRow, range.minCol, range.minRow, range.maxCol, context);
  const series: ChartSeries[] = [];
  for (let rowIndex = range.minRow + 1; rowIndex <= range.maxRow; rowIndex += 1) {
    const seriesIndex = rowIndex - range.minRow - 1;
    const values = rangeFromIndexes(range.worksheetId, rowIndex, range.minCol, rowIndex, range.maxCol, context);
    if (!values) {
      continue;
    }
    series.push(toChartSeries(
      draftMetadata[seriesIndex]?.id || chart.series[seriesIndex]?.id || createId("series"),
      values,
      draftMetadata[seriesIndex]?.name,
      draftMetadata[seriesIndex]?.color ?? chart.series[seriesIndex]?.color,
    ));
  }
  return {
    series,
    ...(categoryAxis ? { categoryAxis } : {}),
  };
}

function buildColumnSeriesFromDataRange(
  range: IndexedRange,
  chart: Chart,
  draftMetadata: Array<{ id: string; name: string | SeriesRange | undefined; color: string | undefined }>,
  context: FormulaContext,
) {
  const categoryAxis = rangeFromIndexes(range.worksheetId, range.minRow + 1, range.minCol, range.maxRow, range.minCol, context);
  const series: ChartSeries[] = [];
  for (let columnIndex = range.minCol + 1; columnIndex <= range.maxCol; columnIndex += 1) {
    const seriesIndex = columnIndex - range.minCol - 1;
    const values = rangeFromIndexes(range.worksheetId, range.minRow + 1, columnIndex, range.maxRow, columnIndex, context);
    if (!values) {
      continue;
    }
    const headerName = rangeFromIndexes(range.worksheetId, range.minRow, columnIndex, range.minRow, columnIndex, context);
    series.push(toChartSeries(
      draftMetadata[seriesIndex]?.id || chart.series[seriesIndex]?.id || createId("series"),
      values,
      draftMetadata[seriesIndex]?.name ?? headerName ?? undefined,
      draftMetadata[seriesIndex]?.color ?? chart.series[seriesIndex]?.color,
    ));
  }
  return {
    series,
    ...(categoryAxis ? { categoryAxis } : {}),
  };
}

function usesRowsAsSeries(chart: Chart, context: FormulaContext) {
  const values = chart.series[0]?.values;
  const valuesRange = values ? indexedRange(values, context) : null;
  if (valuesRange && valuesRange.minRow === valuesRange.maxRow && valuesRange.minCol !== valuesRange.maxCol) {
    return true;
  }
  const categoryRange = chart.categoryAxis ? indexedRange(chart.categoryAxis, context) : null;
  return Boolean(categoryRange && categoryRange.minRow === categoryRange.maxRow && categoryRange.minCol !== categoryRange.maxCol);
}

function indexedRange(range: SeriesRange, context: FormulaContext): IndexedRange | null {
  const projection = context.projectionsByWorksheetId.get(range.worksheetId);
  if (!projection) {
    return null;
  }
  const startRow = projection.rowIndexById.get(range.startRowId);
  const endRow = projection.rowIndexById.get(range.endRowId);
  const startColumn = projection.columnIndexById.get(range.startColumnId);
  const endColumn = projection.columnIndexById.get(range.endColumnId);
  if (startRow == null || endRow == null || startColumn == null || endColumn == null) {
    return null;
  }
  return {
    worksheetId: range.worksheetId,
    minRow: Math.min(startRow, endRow),
    maxRow: Math.max(startRow, endRow),
    minCol: Math.min(startColumn, endColumn),
    maxCol: Math.max(startColumn, endColumn),
  };
}

function rangeFromIndexes(
  worksheetId: string,
  startRow: number,
  startColumn: number,
  endRow: number,
  endColumn: number,
  context: FormulaContext,
): SeriesRange | null {
  const projection = context.projectionsByWorksheetId.get(worksheetId);
  const startRowItem = projection?.rows[startRow];
  const endRowItem = projection?.rows[endRow];
  const startColumnItem = projection?.columns[startColumn];
  const endColumnItem = projection?.columns[endColumn];
  if (!startRowItem || !endRowItem || !startColumnItem || !endColumnItem) {
    return null;
  }
  return {
    worksheetId,
    startRowId: startRowItem.id,
    endRowId: endRowItem.id,
    startColumnId: startColumnItem.id,
    endColumnId: endColumnItem.id,
  };
}

function sameRange(left: SeriesRange | undefined, right: SeriesRange | undefined) {
  return Boolean(left && right
    && left.worksheetId === right.worksheetId
    && left.startRowId === right.startRowId
    && left.endRowId === right.endRowId
    && left.startColumnId === right.startColumnId
    && left.endColumnId === right.endColumnId);
}
