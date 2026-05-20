import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";

import { projectWorksheet, getCell } from "@/features/grid/lib/gridProjection";
import { createTeamGridDocumentFromExcelWorkbook } from "@/features/xlsx/lib/importWorkbook";

describe("Teamgrid XLSX import", () => {
  it("imports worksheets, typed values, dimensions, formulas, and styles", () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Budget Plan");
    worksheet.getColumn(1).width = 20;
    worksheet.getColumn(5).hidden = true;
    worksheet.getRow(1).height = 24;
    worksheet.getRow(3).hidden = true;
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
      wrapText: true,
      indent: 3,
    };
    worksheet.getCell("A1").border = {
      top: { style: "thin", color: { argb: "FFEEF2FF" } },
      bottom: { style: "double", color: { argb: "FF111827" } },
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
    expect(firstWorksheet.columnsById[projection.columns[4].id].hidden).toBe(true);
    expect(projection.columns[4]).toMatchObject({ label: "E", width: 0, hidden: true });
    expect(firstWorksheet.rowsById[projection.rows[0].id].height).toBe(32);
    expect(firstWorksheet.rowsById[projection.rows[2].id].hidden).toBe(true);
    expect(projection.rows[2]).toMatchObject({ label: "3", height: 0, hidden: true });
    expect(cellA1.value).toEqual({ kind: "number", value: 12, format: "currency", currencyCode: "USD", excelNumFmt: "$#,##0.00" });
    expect(cellA1.style).toMatchObject({
      fontFamily: "Inter",
      fontSize: 16,
      bold: true,
      textColor: "#eef2ff",
      backgroundColor: "#111827",
      horizontalAlign: "center",
      verticalAlign: "bottom",
      wrapText: true,
      indent: 3,
      borders: {
        top: { style: "thin", color: "#eef2ff" },
        bottom: { style: "double", color: "#111827" },
      },
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

  it("imports numeric Excel serials with date formats as date cells", () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Serial dates");
    worksheet.getCell("A1").value = excelSerialForUtcDate("2026-01-01T00:00:00.000Z");
    worksheet.getCell("A1").numFmt = "dd.mm.yy";
    worksheet.getCell("B1").value = excelSerialForUtcDate("2026-01-01T14:30:00.000Z");
    worksheet.getCell("B1").numFmt = "dd.mm.yy h:mm";
    worksheet.getCell("C1").value = 42;
    worksheet.getCell("C1").numFmt = "0.00";

    const envelope = createTeamGridDocumentFromExcelWorkbook(workbook, "Serial dates");
    const importedWorksheet = envelope.teamgrid.workbook.worksheetsById[envelope.teamgrid.workbook.worksheetOrder[0]];
    const projection = projectWorksheet(importedWorksheet);
    const cellA1 = getCell(importedWorksheet, projection.rows[0].id, projection.columns[0].id);
    const cellB1 = getCell(importedWorksheet, projection.rows[0].id, projection.columns[1].id);
    const cellC1 = getCell(importedWorksheet, projection.rows[0].id, projection.columns[2].id);

    expect(cellA1.value).toEqual({ kind: "date", isoDate: "2026-01-01T00:00:00.000Z", format: "date", excelNumFmt: "dd.mm.yy" });
    expect(cellB1.value).toEqual({ kind: "date", isoDate: "2026-01-01T14:30:00.000Z", format: "dateTime", excelNumFmt: "dd.mm.yy h:mm" });
    expect(cellC1.value).toEqual({ kind: "number", value: 42, format: "decimal", excelNumFmt: "0.00" });
  });

  it("imports newly supported formulas as formulas", () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Supported");
    worksheet.getCell("A1").value = 4;
    worksheet.getCell("A2").value = 6;
    worksheet.getCell("B1").value = { formula: "IF(A2>5,ROUND(A1*2.25,1),0)", result: 9 };

    const envelope = createTeamGridDocumentFromExcelWorkbook(workbook, "Supported formulas");
    const importedWorksheet = envelope.teamgrid.workbook.worksheetsById[envelope.teamgrid.workbook.worksheetOrder[0]];
    const projection = projectWorksheet(importedWorksheet);
    const cellB1 = getCell(importedWorksheet, projection.rows[0].id, projection.columns[1].id);

    expect(cellB1.formula).toMatchObject({
      source: "=IF(A2>5,ROUND(A1*2.25,1),0)",
      cached: { kind: "number", value: 9 },
    });
    expect(cellB1.value).toEqual({ kind: "number", value: 9 });
  });

  it("imports cross-sheet formulas after all worksheets exist", () => {
    const workbook = new ExcelJS.Workbook();
    const sourceWorksheet = workbook.addWorksheet("Tabelle1");
    const summaryWorksheet = workbook.addWorksheet("Summary");
    sourceWorksheet.getCell("C10").value = 4;
    sourceWorksheet.getCell("C11").value = 6;
    summaryWorksheet.getCell("A1").value = { formula: "SUM(Tabelle1!C10:C11)", result: 10 };

    const envelope = createTeamGridDocumentFromExcelWorkbook(workbook, "Cross-sheet");
    const source = envelope.teamgrid.workbook.worksheetsById[envelope.teamgrid.workbook.worksheetOrder[0]];
    const summary = envelope.teamgrid.workbook.worksheetsById[envelope.teamgrid.workbook.worksheetOrder[1]];
    const sourceProjection = projectWorksheet(source);
    const summaryProjection = projectWorksheet(summary);
    const summaryA1 = getCell(summary, summaryProjection.rows[0].id, summaryProjection.columns[0].id);

    expect(summaryA1.formula).toMatchObject({
      source: "=SUM(Tabelle1!C10:C11)",
      cached: { kind: "number", value: 10 },
      references: [{
        kind: "range",
        worksheetId: source.id,
        startRowId: sourceProjection.rows[9].id,
        endRowId: sourceProjection.rows[10].id,
        startColumnId: sourceProjection.columns[2].id,
        endColumnId: sourceProjection.columns[2].id,
      }],
    });
    expect(summaryA1.value).toEqual({ kind: "number", value: 10 });
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

function excelSerialForUtcDate(isoDate: string) {
  return (new Date(isoDate).getTime() - Date.UTC(1899, 11, 30)) / 86400000;
}
