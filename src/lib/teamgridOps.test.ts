import { describe, expect, it } from "vitest";

import { serializeTeamGridOperations } from "@/lib/teamgridOps";

describe("teamgrid operation serialization", () => {
  it("serializes cell edits without rewriting the workbook", () => {
    expect(serializeTeamGridOperations([{
      type: "setCell",
      worksheetId: "sheet_1",
      cell: {
        id: "row_1:col_1",
        rowId: "row_1",
        columnId: "col_1",
        value: { kind: "number", value: 42 },
      },
    }])).toEqual({
      json: {
        set: [{
          path: ["teamgrid", "workbook", "worksheetsById", "sheet_1", "cellsById", "row_1:col_1"],
          value: {
            id: "row_1:col_1",
            rowId: "row_1",
            columnId: "col_1",
            value: { kind: "number", value: 42 },
          },
        }],
      },
    });
  });

  it("serializes row insertion as metadata plus order-list insert", () => {
    expect(serializeTeamGridOperations([{
      type: "insertRow",
      worksheetId: "sheet_1",
      rowId: "row_2",
      row: { id: "row_2" },
      index: 1,
    }], { baseHeads: ["head_1"] })).toEqual({
      json: {
        baseHeads: ["head_1"],
        set: [{
          path: ["teamgrid", "workbook", "worksheetsById", "sheet_1", "rowsById", "row_2"],
          value: { id: "row_2" },
        }],
        listInsert: [{
          path: ["teamgrid", "workbook", "worksheetsById", "sheet_1", "rowOrder"],
          index: 1,
          values: ["row_2"],
        }],
      },
    });
  });

  it("serializes range style changes as per-cell patches", () => {
    const patch = serializeTeamGridOperations([{
      type: "setCellsStyle",
      worksheetId: "sheet_1",
      style: { bold: true },
      cells: [{
        id: "row_1:col_1",
        rowId: "row_1",
        columnId: "col_1",
        value: { kind: "empty" },
      }],
    }]);

    expect(patch.json?.set).toEqual([{
      path: ["teamgrid", "workbook", "worksheetsById", "sheet_1", "cellsById", "row_1:col_1"],
      value: {
        id: "row_1:col_1",
        rowId: "row_1",
        columnId: "col_1",
        value: { kind: "empty" },
        style: { bold: true },
      },
    }]);
  });

  it("serializes document properties as top-level metadata patches", () => {
    expect(serializeTeamGridOperations([{
      type: "setDocumentProperties",
      subject: "Budget",
      tags: ["Finance\\Q1", "Planning"],
    }])).toEqual({
      json: {
        set: [
          { path: ["subject"], value: "Budget" },
          { path: ["tags"], value: ["Finance\\Q1", "Planning"] },
        ],
      },
    });
  });
});
