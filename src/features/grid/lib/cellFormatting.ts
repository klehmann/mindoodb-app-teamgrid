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
import type { Cell, CellStyle, CellValue, ColumnMeta, CurrencyCode, DateFormat, FormulaResult, NumberFormat, RowMeta } from "@/features/document/lib/teamgridDocument";

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
  horizontalAlign: "general",
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

/**
 * Resolve Excel's "General" horizontal alignment into a concrete
 * direction based on the cell's value (or formula result):
 *
 * - text / empty values render left,
 * - numbers and dates render right,
 * - formula cells inherit from their cached result kind.
 *
 * Explicit `left` / `center` / `right` values are returned unchanged
 * so user-chosen alignment always wins over the value-driven default.
 */
export function effectiveHorizontalAlign(
  cell: Pick<Cell, "value" | "formula">,
  align: NonNullable<CellStyle["horizontalAlign"]> | undefined,
): "left" | "center" | "right" {
  if (align && align !== "general") {
    return align;
  }
  const resultKind = cell.formula?.cached?.kind;
  if (resultKind === "number" || resultKind === "date") {
    return "right";
  }
  if (cell.value.kind === "number" || cell.value.kind === "date") {
    return "right";
  }
  return "left";
}

/**
 * Excel measures indent in "characters". One indent unit is roughly
 * three space widths, which is about 0.6rem at our default 14px cell
 * font size. We clamp to the same 0..15 range Excel exposes so a stray
 * import never blows out a column.
 */
export function indentPaddingRem(indent: number | undefined): number {
  if (!indent || indent <= 0) {
    return 0;
  }
  return Math.min(15, Math.max(0, indent)) * 0.6;
}

export type CellFormatKind = "text" | "general" | "integer" | "decimal" | "percent" | "currency" | "date" | "dateTime" | "time" | "custom";

export interface CellFormatRequest {
  kind: CellFormatKind;
  currencyCode?: CurrencyCode;
  excelNumFmt?: string;
}

const DEFAULT_CURRENCY_CODE: CurrencyCode = "USD";
const MS_PER_DAY = 86400000;
const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);
const EXCEL_TIME_ONLY_YEAR = 1899;
const EXCEL_TIME_ONLY_MONTH = 12;
const EXCEL_TIME_ONLY_DAY = 30;

type DateFormatKind = Extract<CellFormatKind, "date" | "dateTime" | "time">;
type DateCellValue = Extract<CellValue, { kind: "date" }>;

export interface DateFormatOption {
  kind: DateFormatKind;
  excelNumFmt: string;
}

export const DATE_FORMAT_OPTIONS: readonly DateFormatOption[] = Object.freeze([
  { kind: "date", excelNumFmt: "dd.mm.yy" },
  { kind: "date", excelNumFmt: "dddd, d. mmmm yyyy" },
  { kind: "date", excelNumFmt: "yyyy-mm-dd" },
  { kind: "date", excelNumFmt: "d.m" },
  { kind: "date", excelNumFmt: "d.m.yy" },
  { kind: "date", excelNumFmt: "dd.mm.yyyy" },
  { kind: "date", excelNumFmt: "d. mmm." },
  { kind: "date", excelNumFmt: "d. mmm. yy" },
  { kind: "dateTime", excelNumFmt: "dd.mm.yy h:mm" },
  { kind: "dateTime", excelNumFmt: "dd.mm.yy h:mm:ss" },
  { kind: "dateTime", excelNumFmt: "dd.mm.yy h:mm AM/PM" },
  { kind: "dateTime", excelNumFmt: "yyyy-mm-dd h:mm" },
  { kind: "time", excelNumFmt: "h:mm" },
  { kind: "time", excelNumFmt: "h:mm AM/PM" },
  { kind: "time", excelNumFmt: "h:mm:ss" },
  { kind: "time", excelNumFmt: "h:mm:ss AM/PM" },
  { kind: "time", excelNumFmt: "m:ss.0" },
  { kind: "time", excelNumFmt: "[h]:mm:ss" },
]);

/**
 * Convert a raw text edit into a typed {@link CellValue}.
 *
 * Tries the column's preferred kind first (date or number), then falls back
 * to a generic numeric regex, and finally treats the input as a string.
 */
export function coerceInputToCellValue(input: string, preferredKind?: ColumnMeta["defaultValueKind"], locale = "en-US"): CellValue {
  const trimmed = input.trim();
  if (trimmed === "") {
    return { kind: "empty" };
  }
  const currencyValue = parseCurrencyInput(trimmed);
  if (currencyValue) {
    return currencyValue;
  }
  const dateValue = parseDateInput(trimmed, locale);
  if (dateValue) {
    return dateValue;
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

interface ParsedDateParts {
  year: number;
  month: number;
  day: number;
  hours?: number;
  minutes?: number;
  seconds?: number;
  timeOnly?: boolean;
}

export function parseDateInput(input: string, locale = "en-US", preferredFormat?: DateFormat): DateCellValue | null {
  const trimmed = input.trim();
  const parsed = parseDateParts(trimmed, locale);
  if (!parsed) {
    return null;
  }
  const isoDate = datePartsToIsoDate(parsed);
  if (!isoDate) {
    return null;
  }
  const hasTime = parsed.hours != null || parsed.minutes != null || parsed.seconds != null;
  const format = preferredFormat ?? (parsed.timeOnly ? "time" : hasTime ? "dateTime" : "date");
  return {
    kind: "date",
    isoDate,
    format,
    excelNumFmt: defaultExcelNumFmtForParsedDate(parsed, format),
  };
}

function parseDateParts(input: string, locale: string): ParsedDateParts | null {
  const timeMatch = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(input);
  if (timeMatch) {
    return {
      year: EXCEL_TIME_ONLY_YEAR,
      month: EXCEL_TIME_ONLY_MONTH,
      day: EXCEL_TIME_ONLY_DAY,
      hours: Number(timeMatch[1]),
      minutes: Number(timeMatch[2]),
      seconds: timeMatch[3] == null ? undefined : Number(timeMatch[3]),
      timeOnly: true,
    };
  }

  const isoMatch = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(input);
  if (isoMatch) {
    return datePartsFromMatch(isoMatch, "ymd");
  }

  const dottedMatch = /^(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(input);
  if (dottedMatch?.[3]) {
    return datePartsFromMatch(dottedMatch, "dmy");
  }

  const slashMatch = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(input);
  if (slashMatch) {
    const first = Number(slashMatch[1]);
    const second = Number(slashMatch[2]);
    const order = first > 12
      ? "dmy"
      : second > 12
        ? "mdy"
        : localePrefersMonthFirst(locale) ? "mdy" : "dmy";
    return datePartsFromMatch(slashMatch, order);
  }

  return null;
}

function datePartsFromMatch(match: RegExpExecArray, order: "ymd" | "dmy" | "mdy"): ParsedDateParts | null {
  const first = Number(match[1]);
  const second = Number(match[2]);
  const third = Number(match[3]);
  const hours = match[4] == null ? undefined : Number(match[4]);
  const minutes = match[5] == null ? undefined : Number(match[5]);
  const seconds = match[6] == null ? undefined : Number(match[6]);
  const year = normalizeInputYear(order === "ymd" ? first : third);
  const month = order === "ymd" ? second : order === "dmy" ? second : first;
  const day = order === "ymd" ? third : order === "dmy" ? first : second;
  if (year == null) {
    return null;
  }
  return { year, month, day, hours, minutes, seconds };
}

function normalizeInputYear(year: number) {
  if (!Number.isInteger(year) || year < 0) {
    return null;
  }
  return year < 100 ? 2000 + year : year;
}

function datePartsToIsoDate(parts: ParsedDateParts) {
  const hours = parts.hours ?? 0;
  const minutes = parts.minutes ?? 0;
  const seconds = parts.seconds ?? 0;
  if (
    !Number.isInteger(parts.year)
    || !Number.isInteger(parts.month)
    || !Number.isInteger(parts.day)
    || parts.month < 1
    || parts.month > 12
    || parts.day < 1
    || parts.day > 31
    || hours < 0
    || hours > 23
    || minutes < 0
    || minutes > 59
    || seconds < 0
    || seconds > 59
  ) {
    return null;
  }
  const timestamp = Date.UTC(parts.year, parts.month - 1, parts.day, hours, minutes, seconds);
  const date = new Date(timestamp);
  if (
    date.getUTCFullYear() !== parts.year
    || date.getUTCMonth() !== parts.month - 1
    || date.getUTCDate() !== parts.day
    || date.getUTCHours() !== hours
    || date.getUTCMinutes() !== minutes
    || date.getUTCSeconds() !== seconds
  ) {
    return null;
  }
  return date.toISOString();
}

function localePrefersMonthFirst(locale: string) {
  const parts = new Intl.DateTimeFormat(locale, { day: "numeric", month: "numeric", timeZone: "UTC" })
    .formatToParts(new Date(Date.UTC(2026, 0, 2)));
  return parts.find((part) => part.type === "day" || part.type === "month")?.type === "month";
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
      return formatDate(value.isoDate, value.format, locale, value.excelNumFmt);
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
      return formatDate(
        result.isoDate,
        displayValue?.kind === "date" ? displayValue.format : "date",
        locale,
        displayValue?.kind === "date" ? displayValue.excelNumFmt : undefined,
      );
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

export function inferDateFormatFromExcelNumFmt(numFmt: string | undefined): DateFormat | undefined {
  const normalized = normalizeExcelDateFormat(numFmt);
  if (!normalized) {
    return undefined;
  }
  const hasDateToken = /[dy]/.test(normalized) || /m{3,5}/.test(normalized);
  const hasTimeToken = /[hs]/.test(normalized);
  if (!hasDateToken && !hasTimeToken) {
    return undefined;
  }
  if (hasDateToken && hasTimeToken) {
    return "dateTime";
  }
  if (hasTimeToken) {
    return "time";
  }
  return "date";
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
    case "date":
    case "dateTime":
    case "time":
      return request.excelNumFmt?.trim() || defaultExcelNumFmtForDateFormat(request.kind);
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

  if (isDateFormatKind(request.kind)) {
    return applyDateFormatToCellValue(value, request, locale);
  }

  if (value.kind === "date") {
    if (request.kind === "custom") {
      const excelNumFmt = defaultExcelNumFmtForRequest(request);
      return excelNumFmt ? { ...value, format: inferDateFormatFromExcelNumFmt(excelNumFmt) ?? value.format, excelNumFmt } : value;
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

function applyDateFormatToCellValue(value: CellValue, request: CellFormatRequest, locale: string): CellValue {
  const format = dateFormatFromRequestKind(request.kind);
  const excelNumFmt = defaultExcelNumFmtForRequest(request);
  if (value.kind === "date") {
    return { ...value, format, excelNumFmt };
  }
  if (value.kind === "string") {
    const parsed = parseDateInput(value.text, locale, format);
    return parsed ? { ...parsed, format, excelNumFmt } : value;
  }
  if (value.kind === "number") {
    const isoDate = excelSerialToIsoDate(value.value);
    return isoDate ? { kind: "date", isoDate, format, excelNumFmt } : value;
  }
  return value;
}

function isDateFormatKind(kind: CellFormatKind): kind is DateFormatKind {
  return kind === "date" || kind === "dateTime" || kind === "time";
}

function dateFormatFromRequestKind(kind: CellFormatKind): DateFormat {
  return kind === "dateTime" || kind === "time" ? kind : "date";
}

function defaultExcelNumFmtForDateFormat(format: DateFormat) {
  switch (format) {
    case "dateTime":
      return "dd.mm.yy h:mm";
    case "time":
      return "h:mm";
    default:
      return "dd.mm.yy";
  }
}

function defaultExcelNumFmtForParsedDate(parts: ParsedDateParts, format: DateFormat) {
  if (format === "time") {
    return parts.seconds == null ? "h:mm" : "h:mm:ss";
  }
  if (format === "dateTime") {
    return parts.seconds == null ? "dd.mm.yy h:mm" : "dd.mm.yy h:mm:ss";
  }
  return defaultExcelNumFmtForDateFormat(format);
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

function readNumberForFormat(value: CellValue, kind: Exclude<CellFormatKind, "text" | DateFormatKind>) {
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

export function excelSerialToIsoDate(serial: number) {
  if (!Number.isFinite(serial) || serial < 0) {
    return null;
  }
  return new Date(EXCEL_EPOCH_UTC + serial * MS_PER_DAY).toISOString();
}

export function formatDatePreview(excelNumFmt: string, locale = "en-US", isoDate = "2012-03-14T13:30:55.000Z") {
  return formatDateWithExcelNumFmt(isoDate, excelNumFmt, locale) ?? formatDate(isoDate, inferDateFormatFromExcelNumFmt(excelNumFmt), locale);
}

function normalizeExcelDateFormat(numFmt: string | undefined) {
  const firstSection = numFmt?.split(";")[0]?.trim();
  if (!firstSection) {
    return "";
  }
  return firstSection
    .replace(/"[^"]*"/g, "")
    .replace(/\\./g, "")
    .replace(/\[([hms]+)]/gi, "$1")
    .replace(/\[[^\]]*]/g, "")
    .toLowerCase();
}

function formatDate(isoDate: string, format = "date", locale: string, excelNumFmt?: string) {
  const formatted = formatDateWithExcelNumFmt(isoDate, excelNumFmt, locale);
  if (formatted) {
    return formatted;
  }
  const date = new Date(isoDate);
  if (Number.isNaN(date.getTime())) {
    return "";
  }
  if (format === "time") {
    return new Intl.DateTimeFormat(locale, { timeStyle: "short", timeZone: "UTC" }).format(date);
  }
  if (format === "dateTime") {
    return new Intl.DateTimeFormat(locale, { dateStyle: "medium", timeStyle: "short", timeZone: "UTC" }).format(date);
  }
  return new Intl.DateTimeFormat(locale, { dateStyle: "medium", timeZone: "UTC" }).format(date);
}

function formatDateWithExcelNumFmt(isoDate: string, excelNumFmt: string | undefined, locale: string) {
  const normalized = normalizeExcelDateFormat(excelNumFmt);
  if (!normalized || !inferDateFormatFromExcelNumFmt(excelNumFmt)) {
    return null;
  }
  const date = new Date(isoDate);
  if (Number.isNaN(date.getTime())) {
    return "";
  }
  const values = dateFormatValues(date, locale);
  const usesAmPm = /am\/pm/i.test(excelNumFmt ?? "");
  const format = normalized.replace(/am\/pm/gi, "AM/PM");
  let output = "";
  let previousToken = "";
  for (let index = 0; index < format.length;) {
    const token = /^(yyyy|yy|dddd|ddd|dd|d|mmmm|mmm|mm|m|hh|h|ss|s|AM\/PM)/i.exec(format.slice(index))?.[0];
    if (!token) {
      output += format[index];
      index += 1;
      continue;
    }
    const lowerToken = token.toLowerCase();
    const nextToken = /^(yyyy|yy|dddd|ddd|dd|d|mmmm|mmm|mm|m|hh|h|ss|s|AM\/PM)/i.exec(format.slice(index + token.length))?.[0]?.toLowerCase() ?? "";
    output += formatDateToken(lowerToken, values, previousToken, nextToken, usesAmPm);
    previousToken = lowerToken;
    index += token.length;
  }
  return output.trim();
}

function dateFormatValues(date: Date, locale: string) {
  const hours24 = date.getUTCHours();
  const hours12 = hours24 % 12 || 12;
  return {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    weekdayShort: new Intl.DateTimeFormat(locale, { weekday: "short", timeZone: "UTC" }).format(date),
    weekdayLong: new Intl.DateTimeFormat(locale, { weekday: "long", timeZone: "UTC" }).format(date),
    monthShort: new Intl.DateTimeFormat(locale, { month: "short", timeZone: "UTC" }).format(date),
    monthLong: new Intl.DateTimeFormat(locale, { month: "long", timeZone: "UTC" }).format(date),
    hours24,
    hours12,
    minutes: date.getUTCMinutes(),
    seconds: date.getUTCSeconds(),
  };
}

function formatDateToken(
  token: string,
  values: ReturnType<typeof dateFormatValues>,
  previousToken: string,
  nextToken: string,
  usesAmPm: boolean,
) {
  switch (token) {
    case "yyyy":
      return String(values.year);
    case "yy":
      return String(values.year % 100).padStart(2, "0");
    case "dddd":
      return values.weekdayLong;
    case "ddd":
      return values.weekdayShort;
    case "dd":
      return String(values.day).padStart(2, "0");
    case "d":
      return String(values.day);
    case "mmmm":
      return values.monthLong;
    case "mmm":
      return values.monthShort;
    case "mm":
      return isMinuteToken(previousToken, nextToken)
        ? String(values.minutes).padStart(2, "0")
        : String(values.month).padStart(2, "0");
    case "m":
      return isMinuteToken(previousToken, nextToken) ? String(values.minutes) : String(values.month);
    case "hh":
      return String(usesAmPm ? values.hours12 : values.hours24).padStart(2, "0");
    case "h":
      return String(usesAmPm ? values.hours12 : values.hours24);
    case "ss":
      return String(values.seconds).padStart(2, "0");
    case "s":
      return String(values.seconds);
    case "am/pm":
      return values.hours24 < 12 ? "AM" : "PM";
    default:
      return token;
  }
}

function isMinuteToken(previousToken: string, nextToken: string) {
  return previousToken.startsWith("h") || nextToken.startsWith("s");
}
