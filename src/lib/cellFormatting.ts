/**
 * Cell value coercion, locale-aware formatting, and style merging helpers.
 *
 * The schema separates a cell's semantic value (`empty`/`string`/`number`/
 * `date`) from its presentation (`CellStyle`, plus optional `format` on
 * numeric and date variants). This module centralizes the three conversions
 * the UI needs:
 *
 * - {@link coerceInputToCellValue} parses raw user input. The column's
 *   `defaultValueKind` is a hint, not a hard constraint, so a "number"
 *   column can still hold a string when the user types a label.
 * - {@link formatCellValue} / {@link formatFormulaResult} produce locale-aware
 *   display strings via `Intl.NumberFormat` and `Intl.DateTimeFormat`.
 * - {@link mergeCellStyle} layers default style, column default, row
 *   default, and cell override in that order so bulk formatting on a column
 *   does not require rewriting every cell.
 */
import type { Cell, CellStyle, CellValue, ColumnMeta, FormulaResult, NumberFormat, RowMeta } from "@/lib/teamgridDocument";

/** Fallbacks used when neither the column, row, nor cell specifies a value. */
export const DEFAULT_CELL_STYLE: Required<Pick<CellStyle, "fontFamily" | "fontSize" | "horizontalAlign" | "verticalAlign">> = {
  fontFamily: "Inter, sans-serif",
  fontSize: 14,
  horizontalAlign: "left",
  verticalAlign: "middle",
};

/**
 * Convert a raw text edit into a typed {@link CellValue}.
 *
 * Tries the column's preferred kind first (date or number), then falls back
 * to a generic numeric regex, and finally treats the input as a string.
 */
export function coerceInputToCellValue(input: string, preferredKind?: ColumnMeta["defaultValueKind"]): CellValue {
  const trimmed = input.trim();
  if (trimmed === "") {
    return { kind: "empty" };
  }
  if (preferredKind === "date") {
    const timestamp = Date.parse(trimmed);
    if (!Number.isNaN(timestamp)) {
      return { kind: "date", isoDate: new Date(timestamp).toISOString(), format: "date" };
    }
  }
  if (preferredKind === "number" || /^-?\d+(\.\d+)?$/.test(trimmed)) {
    const value = Number(trimmed);
    if (Number.isFinite(value)) {
      return preferredKind === "number"
        ? { kind: "number", value, format: "general" }
        : { kind: "number", value };
    }
  }
  return { kind: "string", text: input };
}

/** Project a formula evaluation result into the cell-value shape used for caching. */
export function formulaResultToCellValue(result: FormulaResult): CellValue {
  switch (result.kind) {
    case "number":
      return { kind: "number", value: result.value };
    case "date":
      return { kind: "date", isoDate: result.isoDate, format: "date" };
    case "string":
      return { kind: "string", text: result.value };
    default:
      return { kind: "empty" };
  }
}

/**
 * Render a cell value as a locale-formatted string for display.
 *
 * Numbers use `Intl.NumberFormat` with the cell's optional format hint,
 * dates use `Intl.DateTimeFormat`, and strings/empty cells are returned
 * unchanged.
 */
export function formatCellValue(value: CellValue, locale = "en-US") {
  switch (value.kind) {
    case "empty":
      return "";
    case "string":
      return value.text;
    case "number":
      return formatNumber(value.value, value.format, locale);
    case "date":
      return formatDate(value.isoDate, value.format, locale);
  }
}

/** Same as {@link formatCellValue}, but also renders formula error codes. */
export function formatFormulaResult(result: FormulaResult, locale = "en-US") {
  switch (result.kind) {
    case "error":
      return result.code;
    case "empty":
      return "";
    case "number":
      return formatNumber(result.value, "general", locale);
    case "date":
      return formatDate(result.isoDate, "date", locale);
    case "string":
      return result.value;
  }
}

/**
 * Layered style resolver.
 *
 * Order, lowest to highest priority: {@link DEFAULT_CELL_STYLE}, column
 * default, row default, then the cell's own overrides. Keeping these layers
 * separate means a "make column B bold" command only writes to the column's
 * `defaultStyle` rather than every existing cell.
 */
export function mergeCellStyle(row: RowMeta | undefined, column: ColumnMeta | undefined, cell: Cell): CellStyle {
  return {
    ...DEFAULT_CELL_STYLE,
    ...column?.defaultStyle,
    ...row?.defaultStyle,
    ...cell.style,
  };
}

function formatNumber(value: number, format: NumberFormat | undefined, locale: string) {
  switch (format) {
    case "integer":
      return new Intl.NumberFormat(locale, { maximumFractionDigits: 0 }).format(value);
    case "decimal":
      return new Intl.NumberFormat(locale, { maximumFractionDigits: 2 }).format(value);
    case "currency":
      return new Intl.NumberFormat(locale, { style: "currency", currency: "USD" }).format(value);
    case "percent":
      return new Intl.NumberFormat(locale, { style: "percent", maximumFractionDigits: 2 }).format(value);
    default:
      return new Intl.NumberFormat(locale, { maximumFractionDigits: 8 }).format(value);
  }
}

function formatDate(isoDate: string, format = "date", locale: string) {
  const date = new Date(isoDate);
  if (Number.isNaN(date.getTime())) {
    return "";
  }
  if (format === "time") {
    return new Intl.DateTimeFormat(locale, { timeStyle: "short" }).format(date);
  }
  if (format === "dateTime") {
    return new Intl.DateTimeFormat(locale, { dateStyle: "medium", timeStyle: "short" }).format(date);
  }
  return new Intl.DateTimeFormat(locale, { dateStyle: "medium" }).format(date);
}
