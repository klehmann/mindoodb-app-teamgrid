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
export type HorizontalAlign = "left" | "center" | "right";
export type VerticalAlign = "top" | "middle" | "bottom";

export interface TeamGridDocumentEnvelope {
  subject: string;
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
  title: string;
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
  | { kind: "string"; text: string }
  | { kind: "number"; value: number; format?: NumberFormat }
  | { kind: "date"; isoDate: string; format?: DateFormat };

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

export function createTeamGridDocument(title = "Untitled spreadsheet"): TeamGridDocumentEnvelope {
  const worksheetId = createId("sheet");
  const rowOrder = Array.from({ length: 24 }, () => createId("row"));
  const columnOrder = Array.from({ length: 12 }, () => createId("col"));

  return {
    subject: title,
    form: TEAMGRID_DOCUMENT_FORM,
    kind: TEAMGRID_DOCUMENT_KIND,
    teamgrid: {
      schemaVersion: TEAMGRID_SCHEMA_VERSION,
      workbook: {
        id: createId("book"),
        title,
        worksheetOrder: [worksheetId],
        worksheetsById: {
          [worksheetId]: {
            id: worksheetId,
            title: "Sheet 1",
            rowOrder,
            columnOrder,
            rowsById: Object.fromEntries(rowOrder.map((id) => [id, { id }] satisfies [RowId, RowMeta])),
            columnsById: Object.fromEntries(columnOrder.map((id) => [id, { id, width: 120 }] satisfies [ColumnId, ColumnMeta])),
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
    return data;
  }
  return createTeamGridDocument(readSubject(data) || "Untitled spreadsheet");
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
  return JSON.parse(JSON.stringify(document)) as TeamGridDocumentV1;
}

export function readSubject(data: Record<string, unknown> | null | undefined) {
  return typeof data?.subject === "string" ? data.subject : "";
}

export function createId(prefix: string) {
  const randomId = typeof crypto !== "undefined" && "randomUUID" in crypto
    ? crypto.randomUUID().replaceAll("-", "").slice(0, 16)
    : Math.random().toString(36).slice(2, 18);
  return `${prefix}_${randomId}`;
}
