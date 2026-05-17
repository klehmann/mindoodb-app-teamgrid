import { DEFAULT_COLUMN_WIDTH } from "@/lib/gridDimensions";

export const TEAMGRID_DOCUMENT_KIND = "mindoodb.teamgrid";
export const TEAMGRID_DOCUMENT_FORM = "teamgrid";
export const TEAMGRID_SCHEMA_VERSION = 1;

export type WorkbookId = string;
export type WorksheetId = string;
export type RowId = string;
export type ColumnId = string;
export type CellId = string;

export type NumberFormat = "general" | "integer" | "decimal" | "currency" | "percent";
export type DateFormat = "date" | "dateTime" | "time";
export type CurrencyCode = "EUR" | "USD";
export type HorizontalAlign = "left" | "center" | "right";
export type VerticalAlign = "top" | "middle" | "bottom";

export interface TeamGridDocumentEnvelope {
  subject: string;
  tags: string[];
  form: typeof TEAMGRID_DOCUMENT_FORM;
  kind: typeof TEAMGRID_DOCUMENT_KIND;
  teamgrid: TeamGridDocumentV1;
}

/**
 * One MindooDB document stores one Teamgrid workbook.
 *
 * Rows, columns, cells, and worksheet tabs use stable application IDs so
 * Automerge can merge independent edits without making formulas depend on
 * mutable grid positions such as `A1`.
 */
export interface TeamGridDocumentV1 {
  schemaVersion: typeof TEAMGRID_SCHEMA_VERSION;
  workbook: Workbook;
  namedExpressionsById: Record<string, NamedExpression>;
  settings: TeamGridSettings;
}

export interface TeamGridSettings {
  locale: string;
}

export interface Workbook {
  id: WorkbookId;
  worksheetOrder: WorksheetId[];
  worksheetsById: Record<WorksheetId, Worksheet>;
}

export interface Worksheet {
  id: WorksheetId;
  title: string;
  rowOrder: RowId[];
  columnOrder: ColumnId[];
  rowsById: Record<RowId, RowMeta>;
  columnsById: Record<ColumnId, ColumnMeta>;
  cellsById: Record<CellId, Cell>;
  deletedAt?: string;
}

export interface RowMeta {
  id: RowId;
  height?: number;
  defaultStyle?: CellStyle;
  deletedAt?: string;
}

export interface ColumnMeta {
  id: ColumnId;
  width?: number;
  defaultStyle?: CellStyle;
  defaultValueKind?: Exclude<CellValue["kind"], "empty">;
  deletedAt?: string;
}

export interface Cell {
  id: CellId;
  rowId: RowId;
  columnId: ColumnId;
  value: CellValue;
  formula?: FormulaCell;
  style?: CellStyle;
}

export type CellValue =
  | { kind: "empty" }
  | { kind: "string"; text: string; excelNumFmt?: string }
  | { kind: "number"; value: number; format?: NumberFormat; currencyCode?: CurrencyCode; excelNumFmt?: string }
  | { kind: "date"; isoDate: string; format?: DateFormat; excelNumFmt?: string };

export interface CellStyle {
  textColor?: string;
  backgroundColor?: string;
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  horizontalAlign?: HorizontalAlign;
  verticalAlign?: VerticalAlign;
}

export interface FormulaCell {
  kind: "formula";
  source: string;
  references: FormulaReference[];
  cached?: FormulaResult;
  error?: FormulaErrorCode;
}

export type FormulaReference =
  | { kind: "cell"; worksheetId: WorksheetId; rowId: RowId; columnId: ColumnId }
  | {
    kind: "range";
    worksheetId: WorksheetId;
    startRowId: RowId;
    endRowId: RowId;
    startColumnId: ColumnId;
    endColumnId: ColumnId;
  }
  | { kind: "column"; worksheetId: WorksheetId; columnId: ColumnId };

export type FormulaResult =
  | { kind: "empty" }
  | { kind: "string"; value: string }
  | { kind: "number"; value: number }
  | { kind: "date"; isoDate: string }
  | { kind: "error"; code: FormulaErrorCode };

export type FormulaErrorCode = "#REF!" | "#CYCLE!" | "#VALUE!" | "#NAME?" | "#DIV/0!";

export interface NamedExpression {
  id: string;
  name: string;
  reference: FormulaReference;
}

export function createTeamGridDocument(title = "Untitled spreadsheet", tags: string[] = []): TeamGridDocumentEnvelope {
  const worksheetId = createId("sheet");
  const rowOrder = Array.from({ length: 24 }, () => createId("row"));
  const columnOrder = Array.from({ length: 12 }, () => createId("col"));

  return {
    subject: title,
    tags: normalizeTags(tags),
    form: TEAMGRID_DOCUMENT_FORM,
    kind: TEAMGRID_DOCUMENT_KIND,
    teamgrid: {
      schemaVersion: TEAMGRID_SCHEMA_VERSION,
      workbook: {
        id: createId("book"),
        worksheetOrder: [worksheetId],
        worksheetsById: {
          [worksheetId]: {
            id: worksheetId,
            title: "Sheet 1",
            rowOrder,
            columnOrder,
            rowsById: Object.fromEntries(rowOrder.map((id) => [id, { id }] satisfies [RowId, RowMeta])),
            columnsById: Object.fromEntries(columnOrder.map((id) => [id, { id, width: DEFAULT_COLUMN_WIDTH }] satisfies [ColumnId, ColumnMeta])),
            cellsById: {},
          },
        },
      },
      namedExpressionsById: {},
      settings: {
        locale: "en-US",
      },
    },
  };
}

export function migrateTeamGridDocument(data: Record<string, unknown> | null | undefined): TeamGridDocumentEnvelope {
  if (isTeamGridEnvelope(data)) {
    return {
      subject: readSubject(data) || "Untitled spreadsheet",
      tags: readTags(data),
      form: TEAMGRID_DOCUMENT_FORM,
      kind: TEAMGRID_DOCUMENT_KIND,
      teamgrid: cloneTeamGridDocument(data.teamgrid),
    };
  }
  return createTeamGridDocument(readSubject(data) || "Untitled spreadsheet", readTags(data));
}

export function isTeamGridEnvelope(data: unknown): data is TeamGridDocumentEnvelope {
  if (!data || typeof data !== "object") {
    return false;
  }
  const candidate = data as Partial<TeamGridDocumentEnvelope>;
  return candidate.form === TEAMGRID_DOCUMENT_FORM
    && candidate.kind === TEAMGRID_DOCUMENT_KIND
    && typeof candidate.subject === "string"
    && isTeamGridDocument(candidate.teamgrid);
}

export function isTeamGridDocument(data: unknown): data is TeamGridDocumentV1 {
  if (!data || typeof data !== "object") {
    return false;
  }
  const candidate = data as Partial<TeamGridDocumentV1>;
  return candidate.schemaVersion === TEAMGRID_SCHEMA_VERSION
    && Boolean(candidate.workbook)
    && typeof candidate.workbook === "object"
    && Array.isArray(candidate.workbook.worksheetOrder)
    && Boolean(candidate.workbook.worksheetsById);
}

export function getFirstVisibleWorksheet(document: TeamGridDocumentV1) {
  return document.workbook.worksheetOrder
    .map((id) => document.workbook.worksheetsById[id])
    .find((worksheet) => worksheet && !worksheet.deletedAt) ?? null;
}

export function createCellId(rowId: RowId, columnId: ColumnId): CellId {
  return `${rowId}:${columnId}`;
}

export function createEmptyCell(rowId: RowId, columnId: ColumnId): Cell {
  return {
    id: createCellId(rowId, columnId),
    rowId,
    columnId,
    value: { kind: "empty" },
  };
}

export function cloneTeamGridDocument(document: TeamGridDocumentV1): TeamGridDocumentV1 {
  const clone = JSON.parse(JSON.stringify(document)) as TeamGridDocumentV1 & {
    workbook: Workbook & { title?: string };
  };
  delete clone.workbook.title;
  return clone;
}

export function readSubject(data: Record<string, unknown> | null | undefined) {
  return typeof data?.subject === "string" ? data.subject : "";
}

export function readTags(data: Record<string, unknown> | null | undefined) {
  return normalizeTags(data?.tags);
}

/**
 * Tags are platform-level document metadata. Keeping normalization here makes
 * storage, properties editing, and Open dialog categorization agree exactly.
 */
export function normalizeTags(input: unknown) {
  if (!Array.isArray(input)) {
    return [];
  }
  const seen = new Set<string>();
  const tags: string[] = [];
  for (const value of input) {
    if (typeof value !== "string") {
      continue;
    }
    const tag = value.trim();
    if (!tag || seen.has(tag)) {
      continue;
    }
    seen.add(tag);
    tags.push(tag);
  }
  return tags;
}

export function createId(prefix: string) {
  const randomId = typeof crypto !== "undefined" && "randomUUID" in crypto
    ? crypto.randomUUID().replaceAll("-", "").slice(0, 16)
    : Math.random().toString(36).slice(2, 18);
  return `${prefix}_${randomId}`;
}
