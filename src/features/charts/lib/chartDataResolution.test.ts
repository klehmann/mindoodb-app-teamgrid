import { describe, expect, it } from "vitest";

import { resolveChartData } from "@/features/charts/lib/chartDataResolution";
import { createCellId, createTeamGridDocument, type Chart } from "@/features/document/lib/teamgridDocument";
import { createFormulaContext } from "@/features/formulas/lib";
import { projectWorksheet } from "@/features/grid/lib/gridProjection";

describe("chart data resolution", () => {
  it("reads labels and numeric values through stable-ID ranges", () => {
    const envelope = createTeamGridDocument();
    const workbook = envelope.teamgrid.workbook;
    const worksheet = workbook.worksheetsById[workbook.worksheetOrder[0]];
    const projection = projectWorksheet(worksheet);
    const [row1, row2, row3] = projection.rows;
    const [colA, colB] = projection.columns;
    worksheet.cellsById[createCellId(row2.id, colA.id)] = { id: createCellId(row2.id, colA.id), rowId: row2.id, columnId: colA.id, value: { kind: "string", text: "Q1" } };
    worksheet.cellsById[createCellId(row3.id, colA.id)] = { id: createCellId(row3.id, colA.id), rowId: row3.id, columnId: colA.id, value: { kind: "string", text: "Q2" } };
    worksheet.cellsById[createCellId(row1.id, colB.id)] = { id: createCellId(row1.id, colB.id), rowId: row1.id, columnId: colB.id, value: { kind: "string", text: "Revenue" } };
    worksheet.cellsById[createCellId(row2.id, colB.id)] = { id: createCellId(row2.id, colB.id), rowId: row2.id, columnId: colB.id, value: { kind: "number", value: 12 } };
    worksheet.cellsById[createCellId(row3.id, colB.id)] = { id: createCellId(row3.id, colB.id), rowId: row3.id, columnId: colB.id, value: { kind: "number", value: 18 } };
    const chart: Chart = {
      id: "chart_1",
      type: "column",
      title: "Revenue",
      categoryAxis: { worksheetId: worksheet.id, startRowId: row2.id, endRowId: row3.id, startColumnId: colA.id, endColumnId: colA.id },
      series: [{
        id: "series_1",
        name: { worksheetId: worksheet.id, startRowId: row1.id, endRowId: row1.id, startColumnId: colB.id, endColumnId: colB.id },
        values: { worksheetId: worksheet.id, startRowId: row2.id, endRowId: row3.id, startColumnId: colB.id, endColumnId: colB.id },
      }],
      anchor: {
        from: { rowId: row1.id, columnId: colB.id, rowOffsetEmu: 0, colOffsetEmu: 0 },
        to: { rowId: row3.id, columnId: colB.id, rowOffsetEmu: 0, colOffsetEmu: 0 },
      },
    };

    const data = resolveChartData(chart, createFormulaContext(workbook));

    expect(data.labels).toEqual(["Q1", "Q2"]);
    expect(data.series).toEqual([{ id: "series_1", name: "Revenue", values: [12, 18], color: undefined }]);
  });
});
