import { DEFAULT_COLUMN_WIDTH } from "@/shared/lib/gridDimensions";

export const TEAMGRID_DOCUMENT_KIND = "mindoodb.teamgrid";
export const TEAMGRID_DOCUMENT_FORM = "teamgrid";
export const TEAMGRID_SCHEMA_VERSION = 2;
export const DEFAULT_WORKSHEET_ROWS = 100;
export const DEFAULT_WORKSHEET_COLUMNS = 12;

export type WorkbookId = string;
export type WorksheetId = string;
export type RowId = string;
export type ColumnId = string;
export type CellId = string;
export type ChartId = string;

export type NumberFormat = "general" | "integer" | "decimal" | "currency" | "percent";
export type DateFormat = "date" | "dateTime" | "time";
export type CurrencyCode = "EUR" | "USD";
/**
 * Excel-compatible horizontal alignment options.
 *
 * `"general"` defers to the cell value: text-like values render left,
 * numeric / date / formula-numeric values render right. It is the same
 * "General" mode Excel uses by default and is the implicit alignment
 * for any cell that does not specify one.
 */
export type HorizontalAlign = "general" | "left" | "center" | "right";
export type VerticalAlign = "top" | "middle" | "bottom";
export type CellBorderStyle = "thin" | "medium" | "thick" | "dashed" | "dotted" | "double";
export type CellBorderSide = "top" | "right" | "bottom" | "left";

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
  chartOrder: ChartId[];
  chartsById: Record<ChartId, Chart>;
  deletedAt?: string;
}

export type ChartType = "column" | "bar" | "line" | "pie";

export interface Chart {
  id: ChartId;
  type: ChartType;
  title?: string;
  series: ChartSeries[];
  categoryAxis?: SeriesRange;
  anchor: TwoCellAnchor;
  legend?: ChartLegend;
  style?: ChartStyle;
  raw?: ChartRawParts;
  deletedAt?: string;
}

export interface ChartSeries {
  id: string;
  name?: string | SeriesRange;
  values: SeriesRange;
  color?: string;
}

export interface SeriesRange {
  worksheetId: WorksheetId;
  startRowId: RowId;
  endRowId: RowId;
  startColumnId: ColumnId;
  endColumnId: ColumnId;
  excelA1?: string;
}

export interface TwoCellAnchor {
  from: ChartAnchorPoint;
  to: ChartAnchorPoint;
}

export interface ChartAnchorPoint {
  rowId: RowId;
  columnId: ColumnId;
  rowOffsetEmu: number;
  colOffsetEmu: number;
}

export interface ChartLegend {
  position: "right" | "bottom" | "top" | "left" | "none";
}

export interface ChartStyle {
  colors?: string[];
  showGridlines?: boolean;
}

export interface ChartRawParts {
  chartXml: string;
  drawingXml: string;
  chartPath?: string;
  drawingPath?: string;
  drawingRelPath?: string;
  worksheetRelPath?: string;
  relationshipId?: string;
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
  /**
   * Excel's "Textumbruch" — when `true`, multi-line and
   * width-overflowing content wraps inside the cell instead of
   * spilling into the neighbouring empty cells or clipping at the
   * column edge.
   */
  wrapText?: boolean;
  /**
   * Excel's "Einzug" — horizontal text indent measured in Excel
   * indent units (1 unit ≈ one character width). Only applies when
   * `horizontalAlign` is `"left"` or `"right"`; ignored otherwise.
   */
  indent?: number;
  borders?: Partial<Record<CellBorderSide, CellBorder>>;
}

export interface CellBorder {
  style: CellBorderStyle;
  color?: string;
}

export interface FormulaCell {
  kind: "formula";
  source: string;
  segments?: FormulaSegment[];
  references: FormulaReference[];
  cached?: FormulaResult;
  error?: FormulaErrorCode;
}

export type FormulaSegment =
  | { kind: "text"; text: string }
  | { kind: "reference"; reference: FormulaReference };

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
  const rowOrder = Array.from({ length: DEFAULT_WORKSHEET_ROWS }, () => createId("row"));
  const columnOrder = Array.from({ length: DEFAULT_WORKSHEET_COLUMNS }, () => createId("col"));

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
            chartOrder: [],
            chartsById: {},
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
  return typeof candidate.schemaVersion === "number"
    && candidate.schemaVersion >= 1
    && candidate.schemaVersion <= TEAMGRID_SCHEMA_VERSION
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
  clone.schemaVersion = TEAMGRID_SCHEMA_VERSION;
  for (const worksheet of Object.values(clone.workbook.worksheetsById)) {
    worksheet.chartOrder ??= [];
    worksheet.chartsById ??= {};
  }
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
