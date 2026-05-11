import type { Cell, CellStyle, CellValue, ColumnMeta, FormulaResult, NumberFormat, RowMeta } from "@/lib/teamgridDocument";

export const DEFAULT_CELL_STYLE: Required<Pick<CellStyle, "fontFamily" | "fontSize" | "horizontalAlign" | "verticalAlign">> = {
  fontFamily: "Inter, sans-serif",
  fontSize: 14,
  horizontalAlign: "left",
  verticalAlign: "middle",
};

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
      return { kind: "number", value, format: preferredKind === "number" ? "general" : undefined };
    }
  }
  return { kind: "string", text: input };
}

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
