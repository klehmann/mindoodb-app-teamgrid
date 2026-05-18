import { describe, expect, it } from "vitest";

import { applyCellFormat, coerceInputToCellValue, formatCellValue, formatFormulaResult, preserveCompatibleCellValueFormat } from "@/features/grid/lib/cellFormatting";
import type { Cell } from "@/features/document/lib/teamgridDocument";

describe("cell formatting", () => {
  it("formats stored currencies with the selected currency code", () => {
    expect(formatCellValue({ kind: "number", value: 12, format: "currency", currencyCode: "EUR" }, "en-US")).toBe("€12.00");
    expect(formatCellValue({ kind: "number", value: 12, format: "currency", currencyCode: "USD" }, "en-US")).toBe("$12.00");
    expect(formatCellValue({ kind: "number", value: 1234, format: "currency", currencyCode: "EUR" }, "en-US")).toBe("€1234.00");
  });

  it("formats numeric formula results with the cell display format", () => {
    expect(formatFormulaResult(
      { kind: "number", value: 1234 },
      "en-US",
      { kind: "number", value: 0, format: "currency", currencyCode: "EUR" },
    )).toBe("€1234.00");
  });

  it("coerces numeric-looking text when applying numeric formats", () => {
    const cell = createCell({ kind: "string", text: "$1,234.50" });

    expect(applyCellFormat(cell, { kind: "currency", currencyCode: "EUR" }).value).toEqual({
      kind: "number",
      value: 1234.5,
      format: "currency",
      currencyCode: "EUR",
      excelNumFmt: "€0.00",
    });
  });

  it("autodetects typed currency values with suffix symbols", () => {
    expect(coerceInputToCellValue("50 €")).toEqual({
      kind: "number",
      value: 50,
      format: "currency",
      currencyCode: "EUR",
      excelNumFmt: "€0.00",
    });
    expect(coerceInputToCellValue("50,95 €")).toEqual({
      kind: "number",
      value: 50.95,
      format: "currency",
      currencyCode: "EUR",
      excelNumFmt: "€0.00",
    });
    expect(coerceInputToCellValue("50 $")).toEqual({
      kind: "number",
      value: 50,
      format: "currency",
      currencyCode: "USD",
      excelNumFmt: "$0.00",
    });
  });

  it("keeps thousands separators distinct from comma decimals while autodetecting currencies", () => {
    expect(coerceInputToCellValue("1.234,56 €")).toMatchObject({ kind: "number", value: 1234.56, currencyCode: "EUR" });
    expect(coerceInputToCellValue("1,234.56 $")).toMatchObject({ kind: "number", value: 1234.56, currencyCode: "USD" });
  });

  it("turns non-formula values into text without losing cell identity", () => {
    const cell = createCell({ kind: "number", value: 42, format: "integer" });

    expect(applyCellFormat(cell, { kind: "text" })).toEqual({
      ...cell,
      value: { kind: "string", text: "42", excelNumFmt: "@" },
    });
  });

  it("preserves formulas while applying number display formats to cached values", () => {
    const cell: Cell = {
      ...createCell({ kind: "number", value: 0 }),
      formula: {
        kind: "formula",
        source: "=A1*2",
        references: [],
        cached: { kind: "number", value: 0 },
      },
    };

    const formatted = applyCellFormat(cell, { kind: "decimal" });

    expect(formatted.formula).toBe(cell.formula);
    expect(formatted.value).toEqual({ kind: "number", value: 0, format: "decimal", excelNumFmt: "0.00" });
  });

  it("preserves compatible number formatting when a cell value changes", () => {
    expect(preserveCompatibleCellValueFormat(
      { kind: "number", value: 50 },
      { kind: "number", value: 12, format: "currency", currencyCode: "EUR", excelNumFmt: "€0.00" },
    )).toEqual({
      kind: "number",
      value: 50,
      format: "currency",
      currencyCode: "EUR",
      excelNumFmt: "€0.00",
    });
  });
});

function createCell(value: Cell["value"]): Cell {
  return {
    id: "row_1:col_1",
    rowId: "row_1",
    columnId: "col_1",
    value,
  };
}
