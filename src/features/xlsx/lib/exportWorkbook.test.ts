import { describe, expect, it } from "vitest";

import { createTeamGridExcelWorkbook, writeTeamGridExcelBuffer } from "@/features/xlsx/lib/exportWorkbook";
import { createCellId, createTeamGridDocument, type TeamGridDocumentV1, type Worksheet } from "@/features/document/lib/teamgridDocument";
import { getCell, projectWorksheet } from "@/features/grid/lib/gridProjection";
import { createTeamGridDocumentFromExcelWorkbook } from "@/features/xlsx/lib/importWorkbook";
import { createFormulaContext, evaluateFormula, renderFormulaSource } from "@/features/formulas/lib";
import { readOoxmlZip, readZipText } from "@/features/xlsx/lib/ooxmlZip";

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
      wrapText: true,
      indent: 2,
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

  it("exports cross-sheet formulas from stable references using exported sheet names", () => {
    const envelope = createTeamGridDocument();
    const document = envelope.teamgrid;
    const firstWorksheet = firstVisibleWorksheet(document);
    const secondWorksheet = cloneWorksheet(firstWorksheet, "sheet_summary");
    firstWorksheet.title = "Tabelle/1";
    secondWorksheet.title = "Summary";
    document.workbook.worksheetOrder.push(secondWorksheet.id);
    document.workbook.worksheetsById[secondWorksheet.id] = secondWorksheet;
    const firstProjection = projectWorksheet(firstWorksheet);
    const secondProjection = projectWorksheet(secondWorksheet);
    const c10 = createCellId(firstProjection.rows[9].id, firstProjection.columns[2].id);
    const c11 = createCellId(firstProjection.rows[10].id, firstProjection.columns[2].id);
    const a1 = createCellId(secondProjection.rows[0].id, secondProjection.columns[0].id);
    firstWorksheet.cellsById[c10] = { id: c10, rowId: firstProjection.rows[9].id, columnId: firstProjection.columns[2].id, value: { kind: "number", value: 4 } };
    firstWorksheet.cellsById[c11] = { id: c11, rowId: firstProjection.rows[10].id, columnId: firstProjection.columns[2].id, value: { kind: "number", value: 6 } };
    const context = createFormulaContext(document.workbook);
    const evaluated = evaluateFormula("=SUM('Tabelle/1'!C10:C11)", secondWorksheet.id, context);
    secondWorksheet.cellsById[a1] = {
      id: a1,
      rowId: secondProjection.rows[0].id,
      columnId: secondProjection.columns[0].id,
      value: { kind: "number", value: 10 },
      formula: {
        kind: "formula",
        source: renderFormulaSource({ source: "=SUM('Tabelle/1'!C10:C11)", segments: evaluated.segments }, secondWorksheet.id, context),
        segments: evaluated.segments,
        references: evaluated.references,
        cached: evaluated.result,
      },
    };

    const workbook = createTeamGridExcelWorkbook(document);

    expect(workbook.worksheets.map((sheet) => sheet.name)).toEqual(["Tabelle 1", "Summary"]);
    expect(workbook.getWorksheet("Summary")!.getCell("A1").value).toEqual({ formula: "SUM('Tabelle 1'!C10:C11)", result: 10 });
  });

  it("round-trips date cells through the Excel workbook representation", () => {
    const document = createExportFixture();
    const workbook = createTeamGridExcelWorkbook(document);

    const envelope = createTeamGridDocumentFromExcelWorkbook(workbook, "Round-tripped");
    const importedWorksheet = envelope.teamgrid.workbook.worksheetsById[envelope.teamgrid.workbook.worksheetOrder[0]];
    const projection = projectWorksheet(importedWorksheet);
    const cellB2 = getCell(importedWorksheet, projection.rows[1].id, projection.columns[1].id);

    expect(cellB2.value).toMatchObject({
      kind: "date",
      isoDate: "2026-05-12T00:00:00.000Z",
      format: "date",
      excelNumFmt: "mmm d, yyyy",
    });
  });

  it("injects TeamGrid charts into exported XLSX drawing parts", async () => {
    const document = createExportFixture();
    const worksheet = firstVisibleWorksheet(document);
    const projection = projectWorksheet(worksheet);
    const [row1, row2, row3, row4, row5] = projection.rows;
    const [colA, colB, colC, colD] = projection.columns;
    worksheet.chartsById.chart_1 = {
      id: "chart_1",
      type: "column",
      title: "Revenue",
      categoryAxis: { worksheetId: worksheet.id, startRowId: row2.id, endRowId: row3.id, startColumnId: colA.id, endColumnId: colA.id },
      series: [{
        id: "series_1",
        name: "Actual",
        values: { worksheetId: worksheet.id, startRowId: row2.id, endRowId: row3.id, startColumnId: colB.id, endColumnId: colB.id },
      }],
      anchor: {
        from: { rowId: row1.id, columnId: colC.id, rowOffsetEmu: 0, colOffsetEmu: 0 },
        to: { rowId: row5.id, columnId: colD.id, rowOffsetEmu: 0, colOffsetEmu: 0 },
      },
    };
    worksheet.chartOrder.push("chart_1");

    const zip = readOoxmlZip(await writeTeamGridExcelBuffer(document));
    const chartXml = readZipText(zip, "xl/charts/chart1.xml");
    const drawingXml = readZipText(zip, "xl/drawings/drawing1.xml");
    const worksheetXml = readZipText(zip, "xl/worksheets/sheet1.xml");

    expect(chartXml).toContain("<c:barDir val=\"col\"/>");
    expect(chartXml).toContain("'Budget Plan 2026'!A2:A3");
    expect(chartXml).toContain("'Budget Plan 2026'!B2:B3");
    expect(drawingXml).toContain("<xdr:twoCellAnchor>");
    expect(worksheetXml).toContain("<drawing r:id=");
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
      wrapText: true,
      indent: 2,
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
