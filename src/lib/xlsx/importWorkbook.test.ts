import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";

import { projectWorksheet, getCell } from "@/lib/gridProjection";
import { createTeamGridDocumentFromExcelWorkbook } from "@/lib/xlsx/importWorkbook";

describe("Teamgrid XLSX import", () => {
  it("imports worksheets, typed values, dimensions, formulas, and styles", () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Budget Plan");
    worksheet.getColumn(1).width = 20;
    worksheet.getRow(1).height = 24;
    worksheet.getCell("A1").value = 12;
    worksheet.getCell("A1").numFmt = "$#,##0.00";
    worksheet.getCell("A1").font = {
      name: "Inter",
      size: 16,
      bold: true,
      color: { argb: "FFEEF2FF" },
    };
    worksheet.getCell("A1").fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF111827" },
    };
    worksheet.getCell("A1").alignment = {
      horizontal: "center",
      vertical: "bottom",
    };
    worksheet.getCell("B1").value = { formula: "A1*2", result: 24 };
    worksheet.getCell("C1").value = 7;
    worksheet.getCell("C1").numFmt = "€#,##0.00";
    worksheet.getCell("D1").value = 1.2345;
    worksheet.getCell("D1").numFmt = "0.0000";
    worksheet.getCell("A2").value = "ready";
    worksheet.getCell("B2").value = new Date("2026-05-12T00:00:00.000Z");
    worksheet.getCell("B2").numFmt = "mmm d, yyyy";
    workbook.addWorksheet("Second Sheet").getCell("A1").value = "two";

    const envelope = createTeamGridDocumentFromExcelWorkbook(workbook, "Imported Budget");
    const firstWorksheet = envelope.teamgrid.workbook.worksheetsById[envelope.teamgrid.workbook.worksheetOrder[0]];
    const secondWorksheet = envelope.teamgrid.workbook.worksheetsById[envelope.teamgrid.workbook.worksheetOrder[1]];
    const projection = projectWorksheet(firstWorksheet);
    const cellA1 = getCell(firstWorksheet, projection.rows[0].id, projection.columns[0].id);
    const cellB1 = getCell(firstWorksheet, projection.rows[0].id, projection.columns[1].id);
    const cellC1 = getCell(firstWorksheet, projection.rows[0].id, projection.columns[2].id);
    const cellD1 = getCell(firstWorksheet, projection.rows[0].id, projection.columns[3].id);
    const cellA2 = getCell(firstWorksheet, projection.rows[1].id, projection.columns[0].id);
    const cellB2 = getCell(firstWorksheet, projection.rows[1].id, projection.columns[1].id);

    expect(envelope.subject).toBe("Imported Budget");
    expect(envelope.teamgrid.workbook.worksheetOrder).toHaveLength(2);
    expect(firstWorksheet.title).toBe("Budget Plan");
    expect(secondWorksheet.title).toBe("Second Sheet");
    expect(firstWorksheet.columnsById[projection.columns[0].id].width).toBe(145);
    expect(firstWorksheet.rowsById[projection.rows[0].id].height).toBe(32);
    expect(cellA1.value).toEqual({ kind: "number", value: 12, format: "currency", currencyCode: "USD", excelNumFmt: "$#,##0.00" });
    expect(cellA1.style).toMatchObject({
      fontFamily: "Inter",
      fontSize: 16,
      bold: true,
      textColor: "#eef2ff",
      backgroundColor: "#111827",
      horizontalAlign: "center",
      verticalAlign: "bottom",
    });
    expect(cellB1.formula).toMatchObject({
      source: "=A1*2",
      cached: { kind: "number", value: 24 },
    });
    expect(cellB1.value).toEqual({ kind: "number", value: 24 });
    expect(cellC1.value).toEqual({ kind: "number", value: 7, format: "currency", currencyCode: "EUR", excelNumFmt: "€#,##0.00" });
    expect(cellD1.value).toEqual({ kind: "number", value: 1.2345, format: "decimal", excelNumFmt: "0.0000" });
    expect(cellA2.value).toEqual({ kind: "string", text: "ready" });
    expect(cellB2.value).toMatchObject({ kind: "date", format: "date", excelNumFmt: "mmm d, yyyy" });
  });

  it("keeps unsupported formulas as cached values instead of broken formulas", () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Unsupported");
    worksheet.getCell("A1").value = { formula: "XLOOKUP(A2,A:A,B:B)", result: "fallback" };

    const envelope = createTeamGridDocumentFromExcelWorkbook(workbook, "Unsupported formulas");
    const importedWorksheet = envelope.teamgrid.workbook.worksheetsById[envelope.teamgrid.workbook.worksheetOrder[0]];
    const projection = projectWorksheet(importedWorksheet);
    const cellA1 = getCell(importedWorksheet, projection.rows[0].id, projection.columns[0].id);

    expect(cellA1.formula).toBeUndefined();
    expect(cellA1.value).toEqual({ kind: "string", text: "fallback" });
  });

  it("omits optional fields instead of importing them as undefined", () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Defaults");
    worksheet.getCell("A1").value = 42;

    const envelope = createTeamGridDocumentFromExcelWorkbook(workbook);
    const importedWorksheet = envelope.teamgrid.workbook.worksheetsById[envelope.teamgrid.workbook.worksheetOrder[0]];
    const projection = projectWorksheet(importedWorksheet);
    const row = importedWorksheet.rowsById[projection.rows[0].id];
    const cellA1 = getCell(importedWorksheet, projection.rows[0].id, projection.columns[0].id);

    expect("height" in row).toBe(false);
    expect("formula" in cellA1).toBe(false);
    expect("style" in cellA1).toBe(false);
    expect(hasExplicitUndefined(envelope)).toBe(false);
  });

  it("imports solid fills that use indexed or theme colors", () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Colors");
    worksheet.getCell("A1").value = "indexed";
    worksheet.getCell("A1").fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { indexed: 22 },
    } as unknown as ExcelJS.Fill;
    worksheet.getCell("B1").value = "theme";
    worksheet.getCell("B1").fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { theme: 4, tint: 0.4 },
    } as unknown as ExcelJS.Fill;

    const envelope = createTeamGridDocumentFromExcelWorkbook(workbook);
    const importedWorksheet = envelope.teamgrid.workbook.worksheetsById[envelope.teamgrid.workbook.worksheetOrder[0]];
    const projection = projectWorksheet(importedWorksheet);
    const cellA1 = getCell(importedWorksheet, projection.rows[0].id, projection.columns[0].id);
    const cellB1 = getCell(importedWorksheet, projection.rows[0].id, projection.columns[1].id);

    expect(cellA1.style?.backgroundColor).toBe("#c0c0c0");
    expect(cellB1.style?.backgroundColor).toBe("#95b3d7");
  });
});

function hasExplicitUndefined(value: unknown): boolean {
  if (value === undefined) {
    return true;
  }
  if (!value || typeof value !== "object") {
    return false;
  }
  if (Array.isArray(value)) {
    return value.some(hasExplicitUndefined);
  }
  return Object.values(value).some(hasExplicitUndefined);
}
