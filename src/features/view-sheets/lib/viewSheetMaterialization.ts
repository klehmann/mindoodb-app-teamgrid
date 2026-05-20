import type {
  MindooDBAppResolvedViewDefinition,
  MindooDBAppViewEntry,
  MindooDBAppViewNavigator,
  MindooDBAppViewNavigatorOpenOptions,
} from "mindoodb-app-sdk";

import { DEFAULT_COLUMN_WIDTH } from "@/shared/lib/gridDimensions";
import {
  createCellId,
  createId,
  type Cell,
  type CellValue,
  type ColumnId,
  type ColumnMeta,
  type RowId,
  type RowMeta,
  type ViewSheetBinding,
  type Worksheet,
} from "@/features/document/lib/teamgridDocument";

const PAGE_SIZE = 1000;
const CATEGORY_ROW_BACKGROUND = "#f3f4f6";

export interface ViewSheetSettings {
  title: string;
  viewId: string;
  showDocuments: boolean;
  showCategories: boolean;
  rootCategoryPathInput: string;
}

export interface MaterializeViewSheetOptions {
  settings: ViewSheetSettings;
  view: MindooDBAppResolvedViewDefinition;
  existingWorksheet?: Worksheet | null;
  openViewNavigator: (
    viewId: string,
    options?: MindooDBAppViewNavigatorOpenOptions,
  ) => Promise<MindooDBAppViewNavigator>;
  now?: () => Date;
}

interface MaterializedColumn {
  name: string;
  title: string;
  role: MindooDBAppResolvedViewDefinition["columns"][number]["role"];
  expression: MindooDBAppResolvedViewDefinition["columns"][number]["expression"];
}

const VIEW_ENTRY_COUNT_BY_FORMULA_OP = {
  childCount: (entry: MindooDBAppViewEntry) => entry.childCount ?? (entry.childCategoryCount ?? 0) + (entry.childDocumentCount ?? 0),
  childCategoryCount: (entry: MindooDBAppViewEntry) => entry.childCategoryCount ?? 0,
  childDocumentCount: (entry: MindooDBAppViewEntry) => entry.childDocumentCount ?? 0,
  descendantCount: (entry: MindooDBAppViewEntry) => entry.descendantCount ?? (entry.descendantDocumentCount ?? 0) + (entry.descendantCategoryCount ?? 0),
  descendantCategoryCount: (entry: MindooDBAppViewEntry) => entry.descendantCategoryCount ?? 0,
  descendantDocumentCount: (entry: MindooDBAppViewEntry) => entry.descendantDocumentCount ?? 0,
  siblingCount: (entry: MindooDBAppViewEntry) => entry.siblingCount ?? 0,
} as const;

export async function materializeViewSheet(options: MaterializeViewSheetOptions): Promise<Worksheet> {
  const rootCategoryPath = parseRootCategoryPath(options.settings.rootCategoryPathInput);
  const navigator = await options.openViewNavigator(options.settings.viewId, {
    includeCategories: options.settings.showCategories,
    includeDocuments: options.settings.showDocuments,
    hideEmptyCategories: true,
    rootCategoryPath: rootCategoryPath.length > 0 ? rootCategoryPath : undefined,
  });

  try {
    await navigator.expandAll();
    const entries = await collectNavigatorEntries(navigator);
    const viewCursor = await navigator.getViewCursor();
    return buildViewSheetWorksheet({
      settings: options.settings,
      view: options.view,
      rootCategoryPath,
      entries,
      existingWorksheet: options.existingWorksheet,
      lastViewCursor: viewCursor,
      now: options.now,
    });
  } finally {
    await navigator.dispose();
  }
}

export function parseRootCategoryPath(input: string) {
  return input
    .split("\\")
    .map((part) => part.trim())
    .filter(Boolean);
}

export function buildViewSheetWorksheet(options: {
  settings: ViewSheetSettings;
  view: MindooDBAppResolvedViewDefinition;
  rootCategoryPath: string[];
  entries: MindooDBAppViewEntry[];
  existingWorksheet?: Worksheet | null;
  lastViewCursor?: string | null;
  now?: () => Date;
}): Worksheet {
  const existing = options.existingWorksheet;
  const columns = viewSheetColumns(options.view);
  const rows = options.entries.filter((entry) =>
    (entry.kind === "category" && options.settings.showCategories)
    || (entry.kind === "document" && options.settings.showDocuments));
  const rowCount = rows.length + 1;
  const rowOrder = reuseIds(existing?.rowOrder, rowCount, "row");
  const columnOrder = reuseIds(existing?.columnOrder, columns.length, "col");
  const rowsById = Object.fromEntries(
    rowOrder.map((id, index) => [id, reuseRowMeta(existing, id, index)] satisfies [RowId, RowMeta]),
  );
  const columnsById = Object.fromEntries(
    columnOrder.map((id, index) => [id, reuseColumnMeta(existing, id, index)] satisfies [ColumnId, ColumnMeta]),
  );
  const cellsById: Record<string, Cell> = {};

  writeRow(cellsById, rowOrder[0], columnOrder, columns.map((column) => column.title), true);
  rows.forEach((entry, index) => {
    writeRow(cellsById, rowOrder[index + 1], columnOrder, rowValues(entry, columns), false, entry.kind === "category");
  });

  const binding: ViewSheetBinding = {
    kind: "mindoodbView",
    viewId: options.view.id,
    viewTitle: options.view.description?.trim() || options.view.id,
    showDocuments: options.settings.showDocuments,
    showCategories: options.settings.showCategories,
    rootCategoryPath: options.rootCategoryPath,
    lastRefreshedAt: (options.now ?? (() => new Date()))().toISOString(),
    lastViewCursor: options.lastViewCursor ?? null,
  };

  return {
    id: existing?.id ?? createId("sheet"),
    title: options.settings.title.trim() || binding.viewTitle,
    rowOrder,
    columnOrder,
    rowsById,
    columnsById,
    cellsById,
    chartOrder: [],
    chartsById: {},
    viewBinding: binding,
  };
}

export function viewSheetColumns(view: MindooDBAppResolvedViewDefinition): MaterializedColumn[] {
  return view.columns
    .filter((column) => !column.hidden)
    .map((column) => ({
      name: column.name,
      title: column.title.trim() || column.name,
      role: column.role,
      expression: column.expression,
    }));
}

export function viewValueToCellValue(value: unknown): CellValue {
  if (value == null) {
    return { kind: "empty" };
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    return { kind: "number", value };
  }
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (isIsoDateLike(trimmed)) {
      return { kind: "date", isoDate: new Date(trimmed).toISOString(), format: trimmed.includes("T") ? "dateTime" : "date" };
    }
    return { kind: "string", text: value };
  }
  if (typeof value === "boolean") {
    return { kind: "string", text: value ? "TRUE" : "FALSE" };
  }
  try {
    return { kind: "string", text: JSON.stringify(value) };
  } catch {
    return { kind: "string", text: String(value) };
  }
}

async function collectNavigatorEntries(navigator: MindooDBAppViewNavigator) {
  const entries: MindooDBAppViewEntry[] = [];
  let startPosition: string | null = null;
  do {
    const page = await navigator.entriesForward({ limit: PAGE_SIZE, startPosition });
    entries.push(...page.entries);
    startPosition = page.nextPosition;
  } while (startPosition);
  return entries;
}

function rowValues(entry: MindooDBAppViewEntry, columns: MaterializedColumn[]) {
  return columns.map((column) =>
    entry.kind === "document" && column.role === "category"
      ? null
      : viewEntryColumnValue(entry, column));
}

function viewEntryColumnValue(entry: MindooDBAppViewEntry, column: MaterializedColumn) {
  const countValue = viewEntryCountColumnValue(entry, column);
  return countValue ?? entry.columnValues[column.name];
}

function viewEntryCountColumnValue(entry: MindooDBAppViewEntry, column: MaterializedColumn) {
  if (column.expression.mode !== "formula" || column.expression.expression.kind !== "operation") {
    return null;
  }

  const readCount = VIEW_ENTRY_COUNT_BY_FORMULA_OP[
    column.expression.expression.op as keyof typeof VIEW_ENTRY_COUNT_BY_FORMULA_OP
  ];
  return readCount ? readCount(entry) : null;
}

function writeRow(
  cellsById: Record<string, Cell>,
  rowId: RowId,
  columnOrder: ColumnId[],
  values: unknown[],
  header: boolean,
  categoryRow = false,
) {
  columnOrder.forEach((columnId, index) => {
    const value = viewValueToCellValue(values[index]);
    const style = header ? { bold: true } : categoryRow ? { backgroundColor: CATEGORY_ROW_BACKGROUND } : undefined;
    if (value.kind === "empty" && !style) {
      return;
    }
    const cell: Cell = {
      id: createCellId(rowId, columnId),
      rowId,
      columnId,
      value,
      style,
    };
    cellsById[cell.id] = cell;
  });
}

function reuseIds<T extends string>(existingIds: T[] | undefined, count: number, prefix: string): T[] {
  return Array.from({ length: count }, (_, index) => existingIds?.[index] ?? createId(prefix) as T);
}

function reuseRowMeta(existing: Worksheet | null | undefined, id: RowId, index: number): RowMeta {
  const previous = existing?.rowsById[id] ?? existing?.rowsById[existing.rowOrder[index]];
  return previous ? { ...previous, id, deletedAt: undefined } : { id };
}

function reuseColumnMeta(existing: Worksheet | null | undefined, id: ColumnId, index: number): ColumnMeta {
  const previous = existing?.columnsById[id] ?? existing?.columnsById[existing.columnOrder[index]];
  return previous
    ? { ...previous, id, deletedAt: undefined }
    : { id, width: DEFAULT_COLUMN_WIDTH };
}

function isIsoDateLike(value: string) {
  if (!/^\d{4}-\d{2}-\d{2}(?:T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z?)?$/.test(value)) {
    return false;
  }
  return Number.isFinite(new Date(value).getTime());
}
