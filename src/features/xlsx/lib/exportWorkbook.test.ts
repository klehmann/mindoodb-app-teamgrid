import { describe, expect, it } from "vitest";

import { createTeamGridExcelWorkbook } from "@/features/xlsx/lib/exportWorkbook";
import { createCellId, createTeamGridDocument, type TeamGridDocumentV1, type Worksheet } from "@/features/document/lib/teamgridDocument";

describe("Teamgrid XLSX export", () => {
  it("exports visible worksheets with typed values, formulas, dimensions, and merged styles", () => {
    const document = createExportFixture();

    const workbook = createTeamGridExcelWorkbook(document);
    const worksheet = workbook.getWorksheet("Budget Plan 2026");

    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual(["Budget Plan 2026"]);
    expect(worksheet).toBeDefined();
    expect(worksheet!.getColumn(1).width).toBeCloseTo(20, 2);
    expect(worksheet!.getRow(1).height).toBe(24);
    expect(worksheet!.getRow(2).height).toBe(24);

    expect(worksheet!.getCell("A1").value).toBe(12);
    expect(worksheet!.getCell("A1").numFmt).toBe("$0.00");
    expect(worksheet!.getCell("B1").value).toEqual({ formula: "A1*2", result: 24 });
    expect(worksheet!.getCell("C1").value).toBe(7);
    expect(worksheet!.getCell("C1").numFmt).toBe("€0.00");
    expect(worksheet!.getCell("D1").value).toBe(1.2345);
    expect(worksheet!.getCell("D1").numFmt).toBe("0.0000");
    expect(worksheet!.getCell("A2").value).toBe("ready");
    expect(worksheet!.getCell("B2").value).toBeInstanceOf(Date);
    expect(worksheet!.getCell("B2").numFmt).toBe("mmm d, yyyy");

    expect(worksheet!.getCell("A1").font).toMatchObject({
      name: "Calibri",
      size: 16,
      bold: true,
      color: { argb: "FFEEF2FF" },
    });
    expect(worksheet!.getCell("A1").fill).toMatchObject({
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF111827" },
    });
    expect(worksheet!.getCell("A1").alignment).toMatchObject({
      horizontal: "center",
      vertical: "bottom",
    });
    expect(worksheet!.getCell("A1").border).toMatchObject({
      top: { style: "thin", color: { argb: "FFEEF2FF" } },
      bottom: { style: "double", color: { argb: "FF111827" } },
    });
  });

  it("deduplicates sanitized sheet names", () => {
    const envelope = createTeamGridDocument();
    const document = envelope.teamgrid;
    const firstWorksheet = firstVisibleWorksheet(document);
    const secondWorksheet = cloneWorksheet(firstWorksheet, "sheet_duplicate");
    firstWorksheet.title = "A/B";
    secondWorksheet.title = "A:B";
    document.workbook.worksheetOrder.push(secondWorksheet.id);
    document.workbook.worksheetsById[secondWorksheet.id] = secondWorksheet;

    const workbook = createTeamGridExcelWorkbook(document);

    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual(["A B", "A B 2"]);
  });
});

function createExportFixture() {
  const envelope = createTeamGridDocument();
  const document = envelope.teamgrid;
  const worksheet = firstVisibleWorksheet(document);
  const deletedWorksheet = cloneWorksheet(worksheet, "sheet_deleted");
  const [row1, row2] = worksheet.rowOrder;
  const [columnA, columnB, columnC, columnD] = worksheet.columnOrder;

  worksheet.title = "Budget/Plan*2026";
  worksheet.rowsById[row1].height = 32;
  worksheet.rowsById[row1].defaultStyle = { bold: true, verticalAlign: "bottom" };
  worksheet.columnsById[columnA].width = 145;
  worksheet.columnsById[columnA].defaultStyle = { backgroundColor: "#111827", horizontalAlign: "center" };

  worksheet.cellsById[createCellId(row1, columnA)] = {
    id: createCellId(row1, columnA),
    rowId: row1,
    columnId: columnA,
    value: { kind: "number", value: 12, format: "currency" },
    style: {
      textColor: "#eef2ff",
      fontSize: 16,
      borders: {
        top: { style: "thin", color: "#eef2ff" },
        bottom: { style: "double", color: "#111827" },
      },
    },
  };
  worksheet.cellsById[createCellId(row1, columnB)] = {
    id: createCellId(row1, columnB),
    rowId: row1,
    columnId: columnB,
    value: { kind: "number", value: 24 },
    formula: {
      kind: "formula",
      source: "=A1*2",
      references: [{ kind: "cell", worksheetId: worksheet.id, rowId: row1, columnId: columnA }],
      cached: { kind: "number", value: 24 },
    },
  };
  worksheet.cellsById[createCellId(row1, columnC)] = {
    id: createCellId(row1, columnC),
    rowId: row1,
    columnId: columnC,
    value: { kind: "number", value: 7, format: "currency", currencyCode: "EUR" },
  };
  worksheet.cellsById[createCellId(row1, columnD)] = {
    id: createCellId(row1, columnD),
    rowId: row1,
    columnId: columnD,
    value: { kind: "number", value: 1.2345, format: "decimal", excelNumFmt: "0.0000" },
  };
  worksheet.cellsById[createCellId(row2, columnA)] = {
    id: createCellId(row2, columnA),
    rowId: row2,
    columnId: columnA,
    value: { kind: "string", text: "ready" },
  };
  worksheet.cellsById[createCellId(row2, columnB)] = {
    id: createCellId(row2, columnB),
    rowId: row2,
    columnId: columnB,
    value: { kind: "date", isoDate: "2026-05-12T00:00:00.000Z", format: "date" },
  };

  deletedWorksheet.deletedAt = "2026-05-12T00:00:00.000Z";
  document.workbook.worksheetOrder.push(deletedWorksheet.id);
  document.workbook.worksheetsById[deletedWorksheet.id] = deletedWorksheet;
  return document;
}

function firstVisibleWorksheet(document: TeamGridDocumentV1) {
  const worksheet = document.workbook.worksheetsById[document.workbook.worksheetOrder[0]];
  if (!worksheet) {
    throw new Error("Expected a fixture worksheet");
  }
  return worksheet;
}

function cloneWorksheet(worksheet: Worksheet, id: string): Worksheet {
  return {
    ...structuredClone(worksheet),
    id,
    title: worksheet.title,
    rowOrder: [...worksheet.rowOrder],
    columnOrder: [...worksheet.columnOrder],
    rowsById: structuredClone(worksheet.rowsById),
    columnsById: structuredClone(worksheet.columnsById),
    cellsById: {},
  };
}
