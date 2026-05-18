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
import type { Cell, CellStyle, CellValue, ColumnMeta, CurrencyCode, FormulaResult, NumberFormat, RowMeta } from "@/features/document/lib/teamgridDocument";

/**
 * Fallbacks used when neither the column, row, nor cell specifies a value.
 *
 * Matches Excel's default of Calibri 11pt. We store bare family names
 * (`"Calibri"`, `"Arial"`, etc.) the same way Excel and our XLSX import
 * do — the generic CSS fallback is appended by the render layer in
 * {@link cssFontFamily}, not baked into the document model.
 *
 * The five Microsoft Office / Windows families that show up most often
 * in `.xlsx` workbooks (Calibri, Cambria, Courier New, Arial, Times New
 * Roman) are aliased via `@font-face` rules in `main.css` to the bundled
 * metric-compatible Croscore fonts (Carlito, Caladea, Cousine, Arimo,
 * Tinos), so the default — and any cell that picks one of those Office
 * families — renders with the right glyph widths on macOS, Linux,
 * Windows, and mobile.
 */
export const DEFAULT_CELL_STYLE: Required<Pick<CellStyle, "fontFamily" | "fontSize" | "horizontalAlign" | "verticalAlign">> = {
  fontFamily: "Calibri",
  fontSize: 14,
  horizontalAlign: "left",
  verticalAlign: "middle",
};

/**
 * Strip any legacy CSS fallback chain off a stored font family.
 *
 * Older documents (and ad-hoc user typing) sometimes store the full
 * fallback chain (`"Calibri, sans-serif"`). The Format-cells combobox
 * and Excel export both want the bare family name (`"Calibri"`), so we
 * keep one canonical normalizer instead of stripping ad-hoc.
 */
export function normalizeFontFamily(fontFamily: string | undefined): string {
  if (!fontFamily) {
    return DEFAULT_CELL_STYLE.fontFamily;
  }
  const head = fontFamily.split(",")[0]?.trim().replace(/^["']|["']$/g, "");
  return head || DEFAULT_CELL_STYLE.fontFamily;
}

/**
 * Build a CSS `font-family` value from a bare family name.
 *
 * Appends a generic fallback chain so a custom or imported family that is
 * not available on the user's system (`Wingdings`, `Aptos`, etc.) still
 * falls back to the app body font rather than the user-agent default.
 * Multi-word names are quoted so they parse correctly as CSS identifiers.
 */
export function cssFontFamily(fontFamily: string | undefined): string | undefined {
  const trimmed = fontFamily?.trim();
  if (!trimmed) {
    return undefined;
  }
  // Already a CSS chain (legacy stored values); pass through unchanged
  // so we do not double-quote or double-fallback.
  if (trimmed.includes(",")) {
    return trimmed;
  }
  return `'${trimmed}', var(--font-body)`;
}

export type CellFormatKind = "text" | "general" | "integer" | "decimal" | "percent" | "currency" | "custom";

export interface CellFormatRequest {
  kind: CellFormatKind;
  currencyCode?: CurrencyCode;
  excelNumFmt?: string;
}

const DEFAULT_CURRENCY_CODE: CurrencyCode = "USD";

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
  const currencyValue = parseCurrencyInput(trimmed);
  if (currencyValue) {
    return currencyValue;
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

function parseCurrencyInput(input: string): CellValue | null {
  const currencyCode = inferCurrencyCodeFromExcelNumFmt(input);
  if (!currencyCode) {
    return null;
  }
  const numericText = input.replace(/[€$]/g, "").trim();
  const value = parseLocalizedNumberText(numericText);
  if (value == null) {
    return null;
  }
  return {
    kind: "number",
    value,
    format: "currency",
    currencyCode,
    excelNumFmt: defaultExcelNumFmtForRequest({ kind: "currency", currencyCode }),
  };
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

export function preserveCompatibleCellValueFormat(nextValue: CellValue, previousValue: CellValue): CellValue {
  if (nextValue.kind === "number" && previousValue.kind === "number") {
    return {
      ...nextValue,
      format: previousValue.format,
      currencyCode: previousValue.currencyCode,
      excelNumFmt: previousValue.excelNumFmt,
    };
  }
  if (nextValue.kind === "date" && previousValue.kind === "date") {
    return {
      ...nextValue,
      format: previousValue.format,
      excelNumFmt: previousValue.excelNumFmt,
    };
  }
  if (nextValue.kind === "string" && previousValue.kind === "string" && previousValue.excelNumFmt) {
    return {
      ...nextValue,
      excelNumFmt: previousValue.excelNumFmt,
    };
  }
  return nextValue;
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
      return formatNumber(value.value, value.format, locale, value.currencyCode);
    case "date":
      return formatDate(value.isoDate, value.format, locale);
  }
}

/** Same as {@link formatCellValue}, but also renders formula error codes. */
export function formatFormulaResult(result: FormulaResult, locale = "en-US", displayValue?: CellValue) {
  switch (result.kind) {
    case "error":
      return result.code;
    case "empty":
      return "";
    case "number":
      return formatNumber(
        result.value,
        displayValue?.kind === "number" ? displayValue.format : "general",
        locale,
        displayValue?.kind === "number" ? displayValue.currencyCode : undefined,
      );
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

export function applyCellFormat(cell: Cell, request: CellFormatRequest, locale = "en-US"): Cell {
  const nextValue = applyFormatToCellValue(cell.value, request, locale, Boolean(cell.formula));
  return nextValue === cell.value ? cell : { ...cell, value: nextValue };
}

export function inferNumberFormatFromExcelNumFmt(numFmt: string | undefined): NumberFormat | undefined {
  if (!numFmt) {
    return undefined;
  }
  if (/%/.test(numFmt)) {
    return "percent";
  }
  if (/[$€£¥]|\[\$-[^\]]+\]|\[\$[A-Z]{3}/i.test(numFmt)) {
    return "currency";
  }
  if (/0\.0+/.test(numFmt)) {
    return "decimal";
  }
  if (/^0$|#,##0/.test(numFmt)) {
    return "integer";
  }
  return "general";
}

export function inferCurrencyCodeFromExcelNumFmt(numFmt: string | undefined): CurrencyCode | undefined {
  if (!numFmt) {
    return undefined;
  }
  if (/\bEUR\b|€|\[\$EUR/i.test(numFmt)) {
    return "EUR";
  }
  if (/\bUSD\b|\$|\[\$USD/i.test(numFmt)) {
    return "USD";
  }
  return undefined;
}

export function defaultExcelNumFmtForRequest(request: CellFormatRequest) {
  switch (request.kind) {
    case "text":
      return "@";
    case "integer":
      return "0";
    case "decimal":
      return "0.00";
    case "percent":
      return "0.00%";
    case "currency":
      return request.currencyCode === "EUR" ? "€0.00" : "$0.00";
    case "custom":
      return request.excelNumFmt?.trim() || undefined;
    default:
      return undefined;
  }
}

function applyFormatToCellValue(value: CellValue, request: CellFormatRequest, locale: string, preservesFormula: boolean): CellValue {
  if (request.kind === "text") {
    if (preservesFormula) {
      return value.kind === "string" ? { ...value, excelNumFmt: "@" } : value;
    }
    return valueToTextCellValue(value, locale);
  }

  if (value.kind === "date") {
    if (request.kind === "custom") {
      const excelNumFmt = defaultExcelNumFmtForRequest(request);
      return excelNumFmt ? { ...value, excelNumFmt } : value;
    }
    return value;
  }

  const numberValue = readNumberForFormat(value, request.kind);
  if (numberValue == null) {
    return value;
  }

  if (request.kind === "currency") {
    const currencyCode = request.currencyCode ?? DEFAULT_CURRENCY_CODE;
    return {
      kind: "number",
      value: numberValue,
      format: "currency",
      currencyCode,
      excelNumFmt: defaultExcelNumFmtForRequest({ kind: "currency", currencyCode }),
    };
  }

  if (request.kind === "custom") {
    const excelNumFmt = defaultExcelNumFmtForRequest(request);
    return {
      kind: "number",
      value: numberValue,
      format: inferNumberFormatFromExcelNumFmt(excelNumFmt),
      currencyCode: inferCurrencyCodeFromExcelNumFmt(excelNumFmt),
      excelNumFmt,
    };
  }

  return {
    kind: "number",
    value: numberValue,
    format: request.kind,
    excelNumFmt: defaultExcelNumFmtForRequest(request),
  };
}

function valueToTextCellValue(value: CellValue, locale: string): CellValue {
  switch (value.kind) {
    case "empty":
      return { kind: "string", text: "", excelNumFmt: "@" };
    case "string":
      return { ...value, excelNumFmt: "@" };
    case "number":
      return { kind: "string", text: String(value.value), excelNumFmt: "@" };
    case "date":
      return { kind: "string", text: formatCellValue(value, locale), excelNumFmt: "@" };
  }
}

function readNumberForFormat(value: CellValue, kind: Exclude<CellFormatKind, "text">) {
  if (value.kind === "number") {
    return value.value;
  }
  if (value.kind !== "string") {
    return null;
  }
  return parseNumberLikeText(value.text, kind === "percent");
}

function parseNumberLikeText(text: string, percentFormat: boolean) {
  const trimmed = text.trim();
  if (!trimmed) {
    return null;
  }
  const hasPercent = trimmed.includes("%");
  const numericText = trimmed
    .replace(/[€$£¥]/g, "")
    .replace(/%/g, "")
    .trim();
  const parsed = parseLocalizedNumberText(numericText);
  if (parsed == null) {
    return null;
  }
  return hasPercent && percentFormat ? parsed / 100 : parsed;
}

function parseLocalizedNumberText(text: string) {
  const compact = text.replace(/\s/g, "");
  const decimalSeparator = inferDecimalSeparator(compact);
  const normalized = compact
    .replace(/[.,]/g, (separator) => separator === decimalSeparator ? "." : "")
    .replace(/'/g, "");
  if (!/^-?\d+(\.\d+)?$/.test(normalized)) {
    return null;
  }
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

function inferDecimalSeparator(text: string) {
  const lastComma = text.lastIndexOf(",");
  const lastDot = text.lastIndexOf(".");
  if (lastComma >= 0 && lastDot >= 0) {
    return lastComma > lastDot ? "," : ".";
  }
  if (lastComma >= 0) {
    return shouldTreatSingleSeparatorAsDecimal(text, lastComma) ? "," : "";
  }
  if (lastDot >= 0) {
    return shouldTreatSingleSeparatorAsDecimal(text, lastDot) ? "." : "";
  }
  return "";
}

function shouldTreatSingleSeparatorAsDecimal(text: string, separatorIndex: number) {
  const digitsAfterSeparator = text.length - separatorIndex - 1;
  return digitsAfterSeparator > 0 && digitsAfterSeparator <= 2;
}

function formatNumber(value: number, format: NumberFormat | undefined, locale: string, currencyCode: CurrencyCode = DEFAULT_CURRENCY_CODE) {
  switch (format) {
    case "integer":
      return new Intl.NumberFormat(locale, { maximumFractionDigits: 0, useGrouping: false }).format(value);
    case "decimal":
      return new Intl.NumberFormat(locale, { maximumFractionDigits: 2, useGrouping: false }).format(value);
    case "currency":
      return new Intl.NumberFormat(locale, { style: "currency", currency: currencyCode, useGrouping: false }).format(value);
    case "percent":
      return new Intl.NumberFormat(locale, { style: "percent", maximumFractionDigits: 2, useGrouping: false }).format(value);
    default:
      return new Intl.NumberFormat(locale, { maximumFractionDigits: 8, useGrouping: false }).format(value);
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
