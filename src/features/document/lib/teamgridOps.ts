/**
 * Semantic operation types for Teamgrid edits and a serializer that converts
 * them into Automerge-friendly JSON patches via the App SDK.
 *
 * The UI never mutates SDK payloads directly. It records intent as a
 * {@link TeamGridOperation} batch and lets {@link serializeTeamGridOperations}
 * turn that batch into a path-scoped {@link MindooDBAppJsonPatch}. This keeps
 * concurrent edits cleanly mergeable: two users editing different cells
 * produce patches that touch disjoint stable-ID paths, so Automerge can
 * combine them without either edit clobbering the other.
 */
import type { MindooDBAppJsonPatch, MindooDBAppUpdateDocumentInput } from "mindoodb-app-sdk";

import type {
  Cell,
  CellId,
  CellStyle,
  Chart,
  ChartId,
  ChartLegend,
  ChartSeries,
  ChartStyle,
  SeriesRange,
  ColumnId,
  ColumnMeta,
  RowId,
  RowMeta,
  TwoCellAnchor,
  Worksheet,
  WorksheetId,
} from "@/features/document/lib/teamgridDocument";

/**
 * One semantic edit recorded by the UI. The shapes are intentionally close to
 * the persisted schema so the serializer can map them to JSON patch
 * operations without re-deriving anything.
 *
 * Variants:
 * - `setCell` writes one cell (value, formula, and style).
 * - `setCellsStyle` rewrites style on a contiguous selection while keeping
 *   each cell's value/formula intact.
 * - `setColumnWidth` / `setRowHeight` update display dimensions by stable ID.
 * - `setColumnHidden` / `setRowHidden` toggle Excel-compatible hide state.
 * - `insertRow` / `insertColumn` add a new ordered slot.
 * - `tombstoneRow` / `tombstoneColumn` / `tombstoneWorksheet` mark an entity
 *   deleted without physically removing it, so formulas can still report
 *   `#REF!` instead of losing identity.
 * - `addWorksheet`, `replaceWorksheet`, `renameWorksheet` handle workbook-level changes.
 * - `moveWorksheet` reorders a tab; `repairWorksheetOrder` rewrites the whole
 *   tab order and exists only for a list a merge left unsound.
 * - `setDocumentProperties` writes the title, tags, template flag, and workbook locale.
 */
export type TeamGridOperation =
  | { type: "setCell"; worksheetId: WorksheetId; cell: Cell }
  | { type: "setCellsStyle"; worksheetId: WorksheetId; cells: Cell[]; style: CellStyle }
  | { type: "setColumnWidth"; worksheetId: WorksheetId; columnId: ColumnId; width: number }
  | { type: "setRowHeight"; worksheetId: WorksheetId; rowId: RowId; height: number }
  | { type: "setColumnHidden"; worksheetId: WorksheetId; columnId: ColumnId; hidden: boolean }
  | { type: "setRowHidden"; worksheetId: WorksheetId; rowId: RowId; hidden: boolean }
  | { type: "insertRow"; worksheetId: WorksheetId; rowId: RowId; row: RowMeta; index: number }
  | { type: "tombstoneRow"; worksheetId: WorksheetId; rowId: RowId; deletedAt: string }
  | { type: "insertColumn"; worksheetId: WorksheetId; columnId: ColumnId; column: ColumnMeta; index: number }
  | { type: "tombstoneColumn"; worksheetId: WorksheetId; columnId: ColumnId; deletedAt: string }
  | { type: "addWorksheet"; worksheet: Worksheet; index: number }
  | { type: "moveWorksheet"; worksheetId: WorksheetId; fromIndex: number; toIndex: number }
  | { type: "repairWorksheetOrder"; order: WorksheetId[] }
  | { type: "replaceWorksheet"; worksheet: Worksheet }
  | { type: "renameWorksheet"; worksheetId: WorksheetId; title: string }
  | { type: "tombstoneWorksheet"; worksheetId: WorksheetId; deletedAt: string }
  | { type: "addChart"; worksheetId: WorksheetId; chart: Chart; index: number }
  | { type: "removeChart"; worksheetId: WorksheetId; chartId: ChartId; deletedAt: string }
  | { type: "setChartAnchor"; worksheetId: WorksheetId; chartId: ChartId; anchor: TwoCellAnchor }
  | { type: "setChartTitle"; worksheetId: WorksheetId; chartId: ChartId; title: string | undefined }
  | { type: "setChartType"; worksheetId: WorksheetId; chartId: ChartId; chartType: Chart["type"] }
  | { type: "setChartSeries"; worksheetId: WorksheetId; chartId: ChartId; series: ChartSeries[] }
  | { type: "setChartCategoryAxis"; worksheetId: WorksheetId; chartId: ChartId; categoryAxis: SeriesRange | undefined }
  | { type: "setChartLegend"; worksheetId: WorksheetId; chartId: ChartId; legend: ChartLegend | undefined }
  | { type: "setChartStyle"; worksheetId: WorksheetId; chartId: ChartId; style: ChartStyle | undefined }
  | { type: "setDocumentProperties"; subject: string; tags: string[]; isTemplate: boolean; locale: string };

/**
 * Convert Teamgrid semantic operations into App SDK JSON patches.
 *
 * The serializer keeps paths scoped to stable IDs, so independent edits to
 * different cells/rows/columns become independent Automerge map/list changes.
 */
export function serializeTeamGridOperations(
  operations: TeamGridOperation[],
  options: { baseHeads?: string[] } = {},
): MindooDBAppUpdateDocumentInput {
  const json: MindooDBAppJsonPatch = options.baseHeads ? { baseHeads: [...options.baseHeads] } : {};
  for (const operation of operations) {
    switch (operation.type) {
      case "setCell":
        pushSet(json, cellPath(operation.worksheetId, operation.cell.id), operation.cell);
        break;
      case "setCellsStyle":
        for (const cell of operation.cells) {
          pushSet(json, cellPath(operation.worksheetId, cell.id), withMergedStyle(cell, operation.style));
        }
        break;
      case "setColumnWidth":
        pushSet(json, [...worksheetPath(operation.worksheetId), "columnsById", operation.columnId, "width"], operation.width);
        break;
      case "setRowHeight":
        pushSet(json, [...worksheetPath(operation.worksheetId), "rowsById", operation.rowId, "height"], operation.height);
        break;
      case "setColumnHidden":
        pushSet(json, [...worksheetPath(operation.worksheetId), "columnsById", operation.columnId, "hidden"], operation.hidden);
        break;
      case "setRowHidden":
        pushSet(json, [...worksheetPath(operation.worksheetId), "rowsById", operation.rowId, "hidden"], operation.hidden);
        break;
      case "insertRow":
        pushSet(json, [...worksheetPath(operation.worksheetId), "rowsById", operation.rowId], operation.row);
        pushListInsert(json, [...worksheetPath(operation.worksheetId), "rowOrder"], operation.index, [operation.rowId]);
        break;
      case "tombstoneRow":
        pushSet(json, [...worksheetPath(operation.worksheetId), "rowsById", operation.rowId, "deletedAt"], operation.deletedAt);
        break;
      case "insertColumn":
        pushSet(json, [...worksheetPath(operation.worksheetId), "columnsById", operation.columnId], operation.column);
        pushListInsert(json, [...worksheetPath(operation.worksheetId), "columnOrder"], operation.index, [operation.columnId]);
        break;
      case "tombstoneColumn":
        pushSet(json, [...worksheetPath(operation.worksheetId), "columnsById", operation.columnId, "deletedAt"], operation.deletedAt);
        break;
      case "addWorksheet":
        pushSet(json, ["teamgrid", "workbook", "worksheetsById", operation.worksheet.id], operation.worksheet);
        pushListInsert(json, ["teamgrid", "workbook", "worksheetOrder"], operation.index, [operation.worksheet.id]);
        break;
      case "moveWorksheet":
        // Automerge has no list move, so the tab travels as a removal and a
        // re-insert of its id. `fromIndex` addresses the list as stored and
        // `toIndex` the list after the removal, which only lines up because
        // the host applies `listDelete` before `listInsert`.
        pushListDelete(json, ["teamgrid", "workbook", "worksheetOrder"], operation.fromIndex, 1);
        pushListInsert(json, ["teamgrid", "workbook", "worksheetOrder"], operation.toIndex, [operation.worksheetId]);
        break;
      case "repairWorksheetOrder":
        // A whole-list write, so it drops a tab another client added
        // concurrently. Reserved for a list that no longer reads back soundly,
        // where the indices a move would send mean nothing; never use it for an
        // ordinary reorder.
        pushSet(json, ["teamgrid", "workbook", "worksheetOrder"], operation.order);
        break;
      case "replaceWorksheet":
        pushSet(json, worksheetPath(operation.worksheet.id), operation.worksheet);
        break;
      case "renameWorksheet":
        pushSet(json, [...worksheetPath(operation.worksheetId), "title"], operation.title);
        break;
      case "tombstoneWorksheet":
        pushSet(json, [...worksheetPath(operation.worksheetId), "deletedAt"], operation.deletedAt);
        break;
      case "addChart":
        pushSet(json, chartPath(operation.worksheetId, operation.chart.id), operation.chart);
        pushListInsert(json, [...worksheetPath(operation.worksheetId), "chartOrder"], operation.index, [operation.chart.id]);
        break;
      case "removeChart":
        pushSet(json, [...chartPath(operation.worksheetId, operation.chartId), "deletedAt"], operation.deletedAt);
        break;
      case "setChartAnchor":
        pushSet(json, [...chartPath(operation.worksheetId, operation.chartId), "anchor"], operation.anchor);
        break;
      case "setChartTitle":
        pushSet(json, [...chartPath(operation.worksheetId, operation.chartId), "title"], operation.title);
        break;
      case "setChartType":
        pushSet(json, [...chartPath(operation.worksheetId, operation.chartId), "type"], operation.chartType);
        break;
      case "setChartSeries":
        pushSet(json, [...chartPath(operation.worksheetId, operation.chartId), "series"], operation.series);
        break;
      case "setChartCategoryAxis":
        pushSet(json, [...chartPath(operation.worksheetId, operation.chartId), "categoryAxis"], operation.categoryAxis);
        break;
      case "setChartLegend":
        pushSet(json, [...chartPath(operation.worksheetId, operation.chartId), "legend"], operation.legend);
        break;
      case "setChartStyle":
        pushSet(json, [...chartPath(operation.worksheetId, operation.chartId), "style"], operation.style);
        break;
      case "setDocumentProperties":
        pushSet(json, ["subject"], operation.subject);
        pushSet(json, ["tags"], operation.tags);
        pushSet(json, ["istemplate"], operation.isTemplate);
        pushSet(json, ["teamgrid", "settings", "locale"], operation.locale);
        break;
    }
  }
  return { json };
}

/** Type-narrowed predicate for a non-empty operation batch. */
export function hasTeamGridOperations(operations: TeamGridOperation[]) {
  return operations.length > 0;
}

/** Path inside the persisted document that points at one worksheet's record. */
function worksheetPath(worksheetId: WorksheetId) {
  return ["teamgrid", "workbook", "worksheetsById", worksheetId];
}

/** Path inside the persisted document that points at one cell's record. */
function cellPath(worksheetId: WorksheetId, cellId: CellId) {
  return [...worksheetPath(worksheetId), "cellsById", cellId];
}

function chartPath(worksheetId: WorksheetId, chartId: ChartId) {
  return [...worksheetPath(worksheetId), "chartsById", chartId];
}

function withMergedStyle(cell: Cell, style: CellStyle): Cell {
  const mergedStyle = {
    ...cell.style,
    ...style,
  };
  if (Object.keys(mergedStyle).length === 0) {
    const cellWithoutStyle = { ...cell };
    delete cellWithoutStyle.style;
    return cellWithoutStyle;
  }
  return {
    ...cell,
    style: mergedStyle,
  };
}

function pushSet(json: MindooDBAppJsonPatch, path: Array<string | number>, value: unknown) {
  json.set ??= [];
  json.set.push({ path, value: toJsonPatchValue(value) });
}

function pushListInsert(json: MindooDBAppJsonPatch, path: Array<string | number>, index: number, values: unknown[]) {
  json.listInsert ??= [];
  json.listInsert.push({ path, index, values: values.map(toJsonPatchValue) });
}

function pushListDelete(json: MindooDBAppJsonPatch, path: Array<string | number>, index: number, deleteCount: number) {
  json.listDelete ??= [];
  json.listDelete.push({ path, index, deleteCount });
}

function toJsonPatchValue(value: unknown) {
  if (value === undefined) {
    return null;
  }
  return JSON.parse(JSON.stringify(value)) as unknown;
}
