import ExcelJS from "exceljs";

import { formulaResultToCellValue } from "@/lib/cellFormatting";
import { evaluateFormula, parseFormula } from "@/lib/formulas";
import { projectWorksheet } from "@/lib/gridProjection";
import {
  createCellId,
  createId,
  createTeamGridDocument,
  type Cell,
  type CellStyle,
  type CellValue,
  type ColumnId,
  type ColumnMeta,
  type FormulaErrorCode,
  type FormulaResult,
  type RowId,
  type RowMeta,
  type TeamGridDocumentEnvelope,
  type Worksheet,
} from "@/lib/teamgridDocument";

const DEFAULT_IMPORTED_ROWS = 24;
const DEFAULT_IMPORTED_COLUMNS = 12;
const EXCEL_INDEXED_COLORS = [
  "000000", "ffffff", "ff0000", "00ff00", "0000ff", "ffff00", "ff00ff", "00ffff",
  "000000", "ffffff", "ff0000", "00ff00", "0000ff", "ffff00", "ff00ff", "00ffff",
  "800000", "008000", "000080", "808000", "800080", "008080", "c0c0c0", "808080",
  "9999ff", "993366", "ffffcc", "ccffff", "660066", "ff8080", "0066cc", "ccccff",
  "000080", "ff00ff", "ffff00", "00ffff", "800080", "800000", "008080", "0000ff",
  "00ccff", "ccffff", "ccffcc", "ffff99", "99ccff", "ff99cc", "cc99ff", "ffcc99",
  "3366ff", "33cccc", "99cc00", "ffcc00", "ff9900", "ff6600", "666699", "969696",
  "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
] as const;
const EXCEL_THEME_COLORS = [
  "ffffff", "000000", "eeece1", "1f497d", "4f81bd", "c0504d", "9bbb59", "8064a2",
  "4bacc6", "f79646", "0000ff", "800080",
] as const;

type ExcelColor = {
  argb?: string;
  indexed?: number;
  theme?: number;
  tint?: number;
};

type ExcelPatternFill = ExcelJS.Fill & {
  type: "pattern";
  pattern?: string;
  fgColor?: ExcelColor;
  bgColor?: ExcelColor;
};

export async function importTeamGridWorkbookBuffer(buffer: ArrayBuffer, title = "Imported spreadsheet") {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return createTeamGridDocumentFromExcelWorkbook(workbook, title);
}

export function createTeamGridDocumentFromExcelWorkbook(workbook: ExcelJS.Workbook, title = "Imported spreadsheet"): TeamGridDocumentEnvelope {
  const envelope = createTeamGridDocument(title);
  const worksheets = workbook.worksheets.filter((worksheet) => worksheet.state !== "hidden" && worksheet.state !== "veryHidden");
  const importedWorksheets = (worksheets.length > 0 ? worksheets : [workbook.addWorksheet("Sheet 1")])
    .map(importWorksheet);

  envelope.teamgrid.workbook.worksheetOrder = importedWorksheets.map((worksheet) => worksheet.id);
  envelope.teamgrid.workbook.worksheetsById = Object.fromEntries(importedWorksheets.map((worksheet) => [worksheet.id, worksheet]));
  return withoutUndefinedProperties(envelope);
}

function importWorksheet(excelWorksheet: ExcelJS.Worksheet): Worksheet {
  const worksheetId = createId("sheet");
  const maxRow = Math.max(DEFAULT_IMPORTED_ROWS, excelWorksheet.actualRowCount || excelWorksheet.rowCount || 1);
  const maxColumn = Math.max(DEFAULT_IMPORTED_COLUMNS, excelWorksheet.actualColumnCount || excelWorksheet.columnCount || 1);
  const rowOrder = Array.from({ length: maxRow }, () => createId("row"));
  const columnOrder = Array.from({ length: maxColumn }, () => createId("col"));
  const rowsById = Object.fromEntries(rowOrder.map((id, index) => [id, importRowMeta(excelWorksheet.getRow(index + 1), id)] satisfies [RowId, RowMeta]));
  const columnsById = Object.fromEntries(columnOrder.map((id, index) => [id, importColumnMeta(excelWorksheet.getColumn(index + 1), id)] satisfies [ColumnId, ColumnMeta]));
  const worksheet: Worksheet = {
    id: worksheetId,
    title: excelWorksheet.name || "Sheet",
    rowOrder,
    columnOrder,
    rowsById,
    columnsById,
    cellsById: {},
  };

  for (let rowIndex = 1; rowIndex <= maxRow; rowIndex += 1) {
    const row = excelWorksheet.getRow(rowIndex);
    for (let columnIndex = 1; columnIndex <= maxColumn; columnIndex += 1) {
      const excelCell = row.getCell(columnIndex);
      const cell = importCell(excelCell, rowOrder[rowIndex - 1], columnOrder[columnIndex - 1]);
      if (cell) {
        worksheet.cellsById[cell.id] = cell;
      }
    }
  }

  importSupportedFormulas(worksheet);
  return worksheet;
}

function importRowMeta(row: ExcelJS.Row, id: RowId): RowMeta {
  const meta: RowMeta = { id };
  if (row.height != null) {
    meta.height = pointsToPixels(row.height);
  }
  return meta;
}

function importColumnMeta(column: ExcelJS.Column, id: ColumnId): ColumnMeta {
  return {
    id,
    width: column.width == null ? 120 : excelWidthToPixels(column.width),
  };
}

function importCell(excelCell: ExcelJS.Cell, rowId: RowId, columnId: ColumnId): Cell | null {
  const formulaSource = readFormulaSource(excelCell.value);
  const value = formulaSource ? excelValueToCellValue(readFormulaResult(excelCell.value), excelCell.numFmt) : excelValueToCellValue(excelCell.value, excelCell.numFmt);
  const style = excelStyleToCellStyle(excelCell);
  if (value.kind === "empty" && !formulaSource && !style) {
    return null;
  }
  return {
    id: createCellId(rowId, columnId),
    rowId,
    columnId,
    value,
    formula: formulaSource
      ? {
          kind: "formula",
          source: formulaSource,
          references: [],
          cached: cellValueToFormulaResult(value),
        }
      : undefined,
    style,
  };
}

function importSupportedFormulas(worksheet: Worksheet) {
  const projection = projectWorksheet(worksheet);
  for (const cell of Object.values(worksheet.cellsById)) {
    if (!cell.formula?.source) {
      continue;
    }
    const parsed = parseFormula(cell.formula.source, worksheet.id, projection);
    if ("code" in parsed) {
      cell.formula = undefined;
      continue;
    }
    const evaluated = evaluateFormula(cell.formula.source, worksheet, projection);
    cell.formula = {
      kind: "formula",
      source: parsed.source,
      references: evaluated.references,
      cached: evaluated.result,
      error: evaluated.result.kind === "error" ? evaluated.result.code : undefined,
    };
    cell.value = formulaResultToCellValue(evaluated.result);
  }
}

function readFormulaSource(value: ExcelJS.CellValue) {
  if (isFormulaValue(value)) {
    return value.formula ? ensureFormulaPrefix(value.formula) : null;
  }
  return null;
}

function readFormulaResult(value: ExcelJS.CellValue) {
  return isFormulaValue(value) ? value.result : value;
}

function isFormulaValue(value: ExcelJS.CellValue): value is ExcelJS.CellFormulaValue {
  return Boolean(value && typeof value === "object" && "formula" in value);
}

function ensureFormulaPrefix(source: string) {
  const trimmed = source.trim();
  return trimmed.startsWith("=") ? trimmed : `=${trimmed}`;
}

function excelValueToCellValue(value: ExcelJS.CellValue | undefined, numFmt?: string): CellValue {
  if (value == null) {
    return { kind: "empty" };
  }
  if (value instanceof Date) {
    return { kind: "date", isoDate: value.toISOString(), format: mapImportedDateFormat(numFmt) };
  }
  if (typeof value === "number") {
    return { kind: "number", value, format: mapImportedNumberFormat(numFmt) };
  }
  if (typeof value === "boolean") {
    return { kind: "string", text: value ? "TRUE" : "FALSE" };
  }
  if (typeof value === "string") {
    return { kind: "string", text: value };
  }
  if (isErrorValue(value)) {
    return { kind: "string", text: value.error };
  }
  if ("text" in value && typeof value.text === "string") {
    return { kind: "string", text: value.text };
  }
  if ("richText" in value && Array.isArray(value.richText)) {
    return { kind: "string", text: value.richText.map((part) => part.text).join("") };
  }
  if ("hyperlink" in value && "text" in value && typeof value.text === "string") {
    return { kind: "string", text: value.text };
  }
  return { kind: "empty" };
}

function isErrorValue(value: object): value is ExcelJS.CellErrorValue {
  return "error" in value && typeof value.error === "string";
}

function cellValueToFormulaResult(value: CellValue): FormulaResult {
  switch (value.kind) {
    case "empty":
      return { kind: "empty" };
    case "string":
      return { kind: "string", value: value.text };
    case "number":
      return { kind: "number", value: value.value };
    case "date":
      return { kind: "date", isoDate: value.isoDate };
  }
}

function excelStyleToCellStyle(cell: ExcelJS.Cell): CellStyle | undefined {
  const style: CellStyle = {};
  const font = cell.font;
  if (font?.name) {
    style.fontFamily = font.name;
  }
  if (font?.size) {
    style.fontSize = font.size;
  }
  if (font?.bold != null) {
    style.bold = font.bold;
  }
  if (font?.italic != null) {
    style.italic = font.italic;
  }
  if (font?.underline) {
    style.underline = true;
  }
  const textColor = argbToCssColor(font?.color?.argb);
  if (textColor) {
    style.textColor = textColor;
  }
  const fillColor = readFillColor(cell.fill);
  if (fillColor) {
    style.backgroundColor = fillColor;
  }
  if (cell.alignment?.horizontal && isSupportedHorizontalAlign(cell.alignment.horizontal)) {
    style.horizontalAlign = cell.alignment.horizontal;
  }
  if (cell.alignment?.vertical && isSupportedVerticalAlign(cell.alignment.vertical)) {
    style.verticalAlign = cell.alignment.vertical === "middle" ? "middle" : cell.alignment.vertical;
  }
  return Object.keys(style).length > 0 ? style : undefined;
}

function readFillColor(fill: ExcelJS.Fill | undefined) {
  if (!fill || fill.type !== "pattern" || fill.pattern !== "solid") {
    return null;
  }
  return excelColorToCssColor((fill as ExcelPatternFill).fgColor)
    ?? excelColorToCssColor((fill as ExcelPatternFill).bgColor);
}

function argbToCssColor(argb: string | undefined) {
  if (!argb || !/^[0-9a-f]{8}$/i.test(argb)) {
    return null;
  }
  return `#${argb.slice(2).toLowerCase()}`;
}

function excelColorToCssColor(color: ExcelColor | undefined) {
  const argbColor = argbToCssColor(color?.argb);
  if (argbColor) {
    return argbColor;
  }
  const rgb = readExcelRgb(color);
  return rgb ? `#${rgb.toLowerCase()}` : null;
}

function readExcelRgb(color: ExcelColor | undefined) {
  if (!color) {
    return null;
  }
  const baseRgb = color.indexed != null
    ? EXCEL_INDEXED_COLORS[color.indexed]
    : color.theme != null
      ? EXCEL_THEME_COLORS[color.theme]
      : null;
  if (!baseRgb) {
    return null;
  }
  return applyExcelTint(baseRgb, color.tint);
}

function applyExcelTint(rgb: string, tint = 0) {
  if (tint === 0) {
    return rgb;
  }
  const channels = rgb.match(/[0-9a-f]{2}/gi);
  if (!channels) {
    return rgb;
  }
  return channels
    .map((hex) => {
      const channel = Number.parseInt(hex, 16);
      const tinted = tint < 0
        ? channel * (1 + tint)
        : channel * (1 - tint) + 255 * tint;
      return clampColorChannel(tinted).toString(16).padStart(2, "0");
    })
    .join("");
}

function clampColorChannel(value: number) {
  return Math.max(0, Math.min(255, Math.round(value)));
}

function isSupportedHorizontalAlign(value: string): value is NonNullable<CellStyle["horizontalAlign"]> {
  return value === "left" || value === "center" || value === "right";
}

function isSupportedVerticalAlign(value: string): value is NonNullable<CellStyle["verticalAlign"]> {
  return value === "top" || value === "middle" || value === "bottom";
}

function mapImportedNumberFormat(numFmt: string | undefined) {
  if (!numFmt) {
    return undefined;
  }
  if (/%/.test(numFmt)) {
    return "percent" as const;
  }
  if (/[$€£¥]/.test(numFmt)) {
    return "currency" as const;
  }
  if (/0\.0+/.test(numFmt)) {
    return "decimal" as const;
  }
  if (/^0$|#,##0/.test(numFmt)) {
    return "integer" as const;
  }
  return "general" as const;
}

function mapImportedDateFormat(numFmt: string | undefined) {
  const normalized = numFmt?.toLowerCase() ?? "";
  if (/[hs]/.test(normalized) && /[dmy]/.test(normalized)) {
    return "dateTime" as const;
  }
  if (/[hs]/.test(normalized)) {
    return "time" as const;
  }
  return "date" as const;
}

function excelWidthToPixels(width: number) {
  return Math.max(1, Math.round(width * 7));
}

function pointsToPixels(points: number) {
  return Math.max(1, Math.round(points / 0.75));
}

function withoutUndefinedProperties<T>(value: T): T {
  if (Array.isArray(value)) {
    return value.map(withoutUndefinedProperties) as T;
  }
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value).flatMap(([key, nested]) => (
        nested === undefined ? [] : [[key, withoutUndefinedProperties(nested)]]
      )),
    ) as T;
  }
  return value;
}
