import ExcelJS from "exceljs";

import { mergeCellStyle } from "@/lib/cellFormatting";
import { DEFAULT_ROW_HEIGHT } from "@/lib/gridDimensions";
import { getCell, projectWorksheet } from "@/lib/gridProjection";
import type {
  Cell,
  CellBorder,
  CellBorderSide,
  CellBorderStyle,
  CellStyle,
  CellValue,
  DateFormat,
  FormulaErrorCode,
  NumberFormat,
  TeamGridDocumentV1,
  Worksheet,
  CurrencyCode,
} from "@/lib/teamgridDocument";

const INVALID_SHEET_NAME_CHARACTERS = /[\\/*?:[\]]/g;
const MAX_SHEET_NAME_LENGTH = 31;

type ExcelFormulaResult = string | number | boolean | Date | ExcelJS.CellErrorValue;

/**
 * Build an ExcelJS workbook from a Teamgrid document.
 *
 * This function is intentionally side-effect free so tests can inspect the
 * workbook structure without needing to create browser downloads.
 */
export function createTeamGridExcelWorkbook(document: TeamGridDocumentV1) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "MindooDB Teamgrid";
  workbook.created = new Date();
  workbook.modified = new Date();

  const usedSheetNames = new Set<string>();
  for (const worksheetId of document.workbook.worksheetOrder) {
    const worksheet = document.workbook.worksheetsById[worksheetId];
    if (!worksheet || worksheet.deletedAt) {
      continue;
    }
    appendWorksheet(workbook, worksheet, usedSheetNames);
  }

  if (workbook.worksheets.length === 0) {
    workbook.addWorksheet("Sheet 1");
  }

  return workbook;
}

/** Serialize the workbook to a browser-friendly ArrayBuffer for Blob downloads. */
export async function writeTeamGridExcelBuffer(document: TeamGridDocumentV1) {
  const workbook = createTeamGridExcelWorkbook(document);
  const buffer = await workbook.xlsx.writeBuffer() as ArrayBuffer | Uint8Array;
  return toArrayBuffer(buffer);
}

function appendWorksheet(workbook: ExcelJS.Workbook, teamgridWorksheet: Worksheet, usedSheetNames: Set<string>) {
  const projection = projectWorksheet(teamgridWorksheet);
  const worksheet = workbook.addWorksheet(createUniqueSheetName(teamgridWorksheet.title, usedSheetNames));

  projection.columns.forEach((column) => {
    worksheet.getColumn(column.index + 1).width = pixelsToExcelWidth(column.width);
  });

  projection.rows.forEach((row) => {
    const excelRow = worksheet.getRow(row.index + 1);
    excelRow.height = pixelsToPoints(row.height ?? DEFAULT_ROW_HEIGHT);

    projection.columns.forEach((column) => {
      const cell = getCell(teamgridWorksheet, row.id, column.id);
      const style = mergeCellStyle(
        teamgridWorksheet.rowsById[row.id],
        teamgridWorksheet.columnsById[column.id],
        cell,
      );

      if (shouldWriteCell(cell, style)) {
        const excelCell = excelRow.getCell(column.index + 1);
        excelCell.value = toExcelCellValue(cell);
        applyExcelStyle(excelCell, style);
        applyExcelNumberFormat(excelCell, cell.value);
      }
    });

    excelRow.commit();
  });
}

function shouldWriteCell(cell: Cell, style: CellStyle) {
  return cell.value.kind !== "empty"
    || Boolean(cell.formula)
    || Boolean(cell.style)
    || Boolean(style.backgroundColor)
    || Boolean(style.borders && Object.keys(style.borders).length > 0);
}

function toExcelCellValue(cell: Cell): ExcelJS.CellValue {
  if (cell.formula?.source) {
    const result = toExcelFormulaResult(cell.value, cell.formula.error);
    if (result == null) {
      return {
        formula: stripFormulaPrefix(cell.formula.source),
      };
    }
    return {
      formula: stripFormulaPrefix(cell.formula.source),
      result,
    };
  }
  return toExcelScalarValue(cell.value);
}

function toExcelFormulaResult(value: CellValue, error: FormulaErrorCode | undefined): ExcelFormulaResult | undefined {
  if (error) {
    return { error: mapFormulaError(error) };
  }
  switch (value.kind) {
    case "empty":
      return undefined;
    case "string":
      return value.text;
    case "number":
      return value.value;
    case "date":
      return new Date(value.isoDate);
  }
}

function mapFormulaError(error: FormulaErrorCode): ExcelJS.CellErrorValue["error"] {
  return error === "#CYCLE!" ? "#REF!" : error;
}

function toExcelScalarValue(value: CellValue): ExcelJS.CellValue {
  switch (value.kind) {
    case "empty":
      return null;
    case "string":
      return value.text;
    case "number":
      return value.value;
    case "date":
      return new Date(value.isoDate);
  }
}

function stripFormulaPrefix(source: string) {
  return source.trim().startsWith("=") ? source.trim().slice(1) : source.trim();
}

function applyExcelStyle(cell: ExcelJS.Cell, style: CellStyle) {
  const font: Partial<ExcelJS.Font> = {};
  const textColor = normalizeHexColor(style.textColor);
  if (style.fontFamily) {
    font.name = stripCssFontFallbacks(style.fontFamily);
  }
  if (style.fontSize) {
    font.size = style.fontSize;
  }
  if (style.bold != null) {
    font.bold = style.bold;
  }
  if (style.italic != null) {
    font.italic = style.italic;
  }
  if (style.underline != null) {
    font.underline = style.underline;
  }
  if (textColor) {
    font.color = { argb: textColor };
  }
  if (Object.keys(font).length > 0) {
    cell.font = font;
  }

  const backgroundColor = normalizeHexColor(style.backgroundColor);
  if (backgroundColor) {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: backgroundColor },
    };
  }

  if (style.horizontalAlign || style.verticalAlign) {
    cell.alignment = {
      horizontal: style.horizontalAlign,
      vertical: style.verticalAlign,
    };
  }

  const borders = toExcelBorders(style.borders);
  if (borders) {
    cell.border = borders;
  }
}

function toExcelBorders(borders: CellStyle["borders"]): Partial<ExcelJS.Borders> | null {
  if (!borders || Object.keys(borders).length === 0) {
    return null;
  }
  const excelBorders: Partial<ExcelJS.Borders> = {};
  for (const side of ["top", "right", "bottom", "left"] satisfies CellBorderSide[]) {
    const border = borders[side];
    if (border) {
      excelBorders[side] = toExcelBorder(border);
    }
  }
  return Object.keys(excelBorders).length > 0 ? excelBorders : null;
}

function toExcelBorder(border: CellBorder): Partial<ExcelJS.Border> {
  const excelBorder: Partial<ExcelJS.Border> = {
    style: toExcelBorderStyle(border.style),
  };
  const color = normalizeHexColor(border.color);
  if (color) {
    excelBorder.color = { argb: color };
  }
  return excelBorder;
}

function toExcelBorderStyle(style: CellBorderStyle): ExcelJS.BorderStyle {
  return style;
}

function applyExcelNumberFormat(cell: ExcelJS.Cell, value: CellValue) {
  if (value.kind === "string" && value.excelNumFmt) {
    cell.numFmt = value.excelNumFmt;
    return;
  }
  if (value.kind === "number") {
    const numberFormat = value.excelNumFmt ?? mapNumberFormat(value.format, value.currencyCode);
    if (numberFormat) {
      cell.numFmt = numberFormat;
    }
  }
  if (value.kind === "date") {
    cell.numFmt = value.excelNumFmt ?? mapDateFormat(value.format);
  }
}

function mapNumberFormat(format: NumberFormat | undefined, currencyCode: CurrencyCode | undefined) {
  switch (format) {
    case "integer":
      return "0";
    case "decimal":
      return "0.00";
    case "currency":
      return currencyCode === "EUR" ? "€0.00" : "$0.00";
    case "percent":
      return "0.00%";
    default:
      return undefined;
  }
}

function mapDateFormat(format: DateFormat | undefined) {
  switch (format) {
    case "dateTime":
      return "mmm d, yyyy h:mm AM/PM";
    case "time":
      return "h:mm AM/PM";
    default:
      return "mmm d, yyyy";
  }
}

function normalizeHexColor(color: string | undefined) {
  if (!color) {
    return null;
  }
  const normalized = color.trim();
  const shortHex = /^#([0-9a-f]{3})$/i.exec(normalized);
  if (shortHex) {
    return `FF${shortHex[1].split("").map((character) => character + character).join("").toUpperCase()}`;
  }
  const longHex = /^#([0-9a-f]{6})$/i.exec(normalized);
  if (longHex) {
    return `FF${longHex[1].toUpperCase()}`;
  }
  return null;
}

function stripCssFontFallbacks(fontFamily: string) {
  return fontFamily.split(",")[0]?.trim().replace(/^["']|["']$/g, "") || "Inter";
}

function pixelsToExcelWidth(pixels: number) {
  return Math.max(1, Math.round(((pixels - 5) / 7) * 100) / 100);
}

function pixelsToPoints(pixels: number) {
  return Math.max(1, Math.round((pixels * 0.75) * 100) / 100);
}

function createUniqueSheetName(title: string, usedSheetNames: Set<string>) {
  const baseName = sanitizeSheetName(title);
  let candidate = baseName;
  let index = 2;
  while (usedSheetNames.has(candidate.toLowerCase())) {
    const suffix = ` ${index}`;
    candidate = `${baseName.slice(0, MAX_SHEET_NAME_LENGTH - suffix.length)}${suffix}`;
    index += 1;
  }
  usedSheetNames.add(candidate.toLowerCase());
  return candidate;
}

function sanitizeSheetName(title: string) {
  const sanitized = title.replace(INVALID_SHEET_NAME_CHARACTERS, " ").trim() || "Sheet";
  return sanitized.slice(0, MAX_SHEET_NAME_LENGTH);
}

function toArrayBuffer(buffer: ArrayBuffer | Uint8Array) {
  if (buffer instanceof ArrayBuffer) {
    return buffer;
  }
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength) as ArrayBuffer;
}
