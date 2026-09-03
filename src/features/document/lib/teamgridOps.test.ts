import { describe, expect, it } from "vitest";

import { serializeTeamGridOperations } from "@/features/document/lib/teamgridOps";

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

  it("removes undefined properties from JSON patch values", () => {
    const patch = serializeTeamGridOperations([{
      type: "setCell",
      worksheetId: "sheet_1",
      cell: {
        id: "row_1:col_1",
        rowId: "row_1",
        columnId: "col_1",
        value: { kind: "number", value: 42, format: undefined },
      },
    }]);

    expect(patch.json?.set?.[0]?.value).toEqual({
      id: "row_1:col_1",
      rowId: "row_1",
      columnId: "col_1",
      value: { kind: "number", value: 42 },
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

  it("serializes cleared range styles without empty background overrides", () => {
    const patch = serializeTeamGridOperations([{
      type: "setCellsStyle",
      worksheetId: "sheet_1",
      style: {},
      cells: [{
        id: "row_1:col_1",
        rowId: "row_1",
        columnId: "col_1",
        value: { kind: "empty" },
        style: { bold: true },
      }, {
        id: "row_1:col_2",
        rowId: "row_1",
        columnId: "col_2",
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
    }, {
      path: ["teamgrid", "workbook", "worksheetsById", "sheet_1", "cellsById", "row_1:col_2"],
      value: {
        id: "row_1:col_2",
        rowId: "row_1",
        columnId: "col_2",
        value: { kind: "empty" },
      },
    }]);
  });

  it("serializes border changes and omits empty border objects", () => {
    const patch = serializeTeamGridOperations([{
      type: "setCellsStyle",
      worksheetId: "sheet_1",
      style: {},
      cells: [{
        id: "row_1:col_1",
        rowId: "row_1",
        columnId: "col_1",
        value: { kind: "empty" },
        style: {
          borders: {
            top: { style: "thin", color: "#111827" },
          },
        },
      }, {
        id: "row_1:col_2",
        rowId: "row_1",
        columnId: "col_2",
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
        style: {
          borders: {
            top: { style: "thin", color: "#111827" },
          },
        },
      },
    }, {
      path: ["teamgrid", "workbook", "worksheetsById", "sheet_1", "cellsById", "row_1:col_2"],
      value: {
        id: "row_1:col_2",
        rowId: "row_1",
        columnId: "col_2",
        value: { kind: "empty" },
      },
    }]);
  });

  it("serializes column width and row height changes by stable ID", () => {
    expect(serializeTeamGridOperations([
      {
        type: "setColumnWidth",
        worksheetId: "sheet_1",
        columnId: "col_1",
        width: 180,
      },
      {
        type: "setRowHeight",
        worksheetId: "sheet_1",
        rowId: "row_1",
        height: 44,
      },
    ])).toEqual({
      json: {
        set: [
          {
            path: ["teamgrid", "workbook", "worksheetsById", "sheet_1", "columnsById", "col_1", "width"],
            value: 180,
          },
          {
            path: ["teamgrid", "workbook", "worksheetsById", "sheet_1", "rowsById", "row_1", "height"],
            value: 44,
          },
        ],
      },
    });
  });

  it("serializes row and column hide state by stable ID", () => {
    expect(serializeTeamGridOperations([
      {
        type: "setColumnHidden",
        worksheetId: "sheet_1",
        columnId: "col_1",
        hidden: true,
      },
      {
        type: "setRowHidden",
        worksheetId: "sheet_1",
        rowId: "row_1",
        hidden: false,
      },
    ])).toEqual({
      json: {
        set: [
          {
            path: ["teamgrid", "workbook", "worksheetsById", "sheet_1", "columnsById", "col_1", "hidden"],
            value: true,
          },
          {
            path: ["teamgrid", "workbook", "worksheetsById", "sheet_1", "rowsById", "row_1", "hidden"],
            value: false,
          },
        ],
      },
    });
  });

  it("serializes document properties and locale patches", () => {
    expect(serializeTeamGridOperations([{
      type: "setDocumentProperties",
      subject: "Budget",
      tags: ["Finance\\Q1", "Planning"],
      isTemplate: true,
      locale: "de-DE",
    }])).toEqual({
      json: {
        set: [
          { path: ["subject"], value: "Budget" },
          { path: ["tags"], value: ["Finance\\Q1", "Planning"] },
          { path: ["istemplate"], value: true },
          { path: ["teamgrid", "settings", "locale"], value: "de-DE" },
        ],
      },
    });
  });

  it("serializes a tab move as a removal plus a re-insert of its id", () => {
    expect(serializeTeamGridOperations([{
      type: "moveWorksheet",
      worksheetId: "sheet_3",
      fromIndex: 2,
      toIndex: 0,
    }], { baseHeads: ["head_1"] })).toEqual({
      json: {
        baseHeads: ["head_1"],
        listDelete: [{
          path: ["teamgrid", "workbook", "worksheetOrder"],
          index: 2,
          deleteCount: 1,
        }],
        listInsert: [{
          path: ["teamgrid", "workbook", "worksheetOrder"],
          index: 0,
          values: ["sheet_3"],
        }],
      },
    });
  });

  it("serializes an order repair as one whole-list write", () => {
    expect(serializeTeamGridOperations([{
      type: "repairWorksheetOrder",
      order: ["sheet_2", "sheet_1"],
    }])).toEqual({
      json: {
        set: [{
          path: ["teamgrid", "workbook", "worksheetOrder"],
          value: ["sheet_2", "sheet_1"],
        }],
      },
    });
  });

  it("serializes generated worksheet replacements at the worksheet path", () => {
    expect(serializeTeamGridOperations([{
      type: "replaceWorksheet",
      worksheet: {
        id: "sheet_view",
        title: "Contacts",
        rowOrder: ["row_1"],
        columnOrder: ["col_1"],
        rowsById: { row_1: { id: "row_1" } },
        columnsById: { col_1: { id: "col_1" } },
        cellsById: {},
        chartOrder: [],
        chartsById: {},
        viewBinding: {
          kind: "mindoodbView",
          viewId: "contacts_flat",
          viewTitle: "Contacts",
          showDocuments: true,
          showCategories: true,
          rootCategoryPath: ["Customers"],
          lastRefreshedAt: "2026-05-20T12:00:00.000Z",
          lastViewCursor: "cursor_1",
        },
      },
    }])).toEqual({
      json: {
        set: [{
          path: ["teamgrid", "workbook", "worksheetsById", "sheet_view"],
          value: {
            id: "sheet_view",
            title: "Contacts",
            rowOrder: ["row_1"],
            columnOrder: ["col_1"],
            rowsById: { row_1: { id: "row_1" } },
            columnsById: { col_1: { id: "col_1" } },
            cellsById: {},
            chartOrder: [],
            chartsById: {},
            viewBinding: {
              kind: "mindoodbView",
              viewId: "contacts_flat",
              viewTitle: "Contacts",
              showDocuments: true,
              showCategories: true,
              rootCategoryPath: ["Customers"],
              lastRefreshedAt: "2026-05-20T12:00:00.000Z",
              lastViewCursor: "cursor_1",
            },
          },
        }],
      },
    });
  });

  it("serializes chart creation and anchor edits by stable ID", () => {
    const chart = {
      id: "chart_1",
      type: "column" as const,
      title: "Revenue",
      series: [{
        id: "series_1",
        values: {
          worksheetId: "sheet_1",
          startRowId: "row_2",
          endRowId: "row_3",
          startColumnId: "col_2",
          endColumnId: "col_2",
        },
      }],
      anchor: {
        from: { rowId: "row_1", columnId: "col_4", rowOffsetEmu: 0, colOffsetEmu: 0 },
        to: { rowId: "row_10", columnId: "col_8", rowOffsetEmu: 0, colOffsetEmu: 0 },
      },
    };

    expect(serializeTeamGridOperations([{
      type: "addChart",
      worksheetId: "sheet_1",
      chart,
      index: 0,
    }, {
      type: "setChartAnchor",
      worksheetId: "sheet_1",
      chartId: "chart_1",
      anchor: chart.anchor,
    }]).json).toMatchObject({
      set: [
        { path: ["teamgrid", "workbook", "worksheetsById", "sheet_1", "chartsById", "chart_1"], value: chart },
        { path: ["teamgrid", "workbook", "worksheetsById", "sheet_1", "chartsById", "chart_1", "anchor"], value: chart.anchor },
      ],
      listInsert: [{
        path: ["teamgrid", "workbook", "worksheetsById", "sheet_1", "chartOrder"],
        index: 0,
        values: ["chart_1"],
      }],
    });
  });

  it("serializes chart property edits by stable ID", () => {
    const categoryAxis = {
      worksheetId: "sheet_1",
      startRowId: "row_2",
      endRowId: "row_3",
      startColumnId: "col_1",
      endColumnId: "col_1",
    };

    expect(serializeTeamGridOperations([{
      type: "setChartTitle",
      worksheetId: "sheet_1",
      chartId: "chart_1",
      title: "Pipeline",
    }, {
      type: "setChartType",
      worksheetId: "sheet_1",
      chartId: "chart_1",
      chartType: "line",
    }, {
      type: "setChartCategoryAxis",
      worksheetId: "sheet_1",
      chartId: "chart_1",
      categoryAxis,
    }, {
      type: "setChartLegend",
      worksheetId: "sheet_1",
      chartId: "chart_1",
      legend: { position: "bottom" },
    }, {
      type: "setChartStyle",
      worksheetId: "sheet_1",
      chartId: "chart_1",
      style: { showGridlines: false, colors: ["#4472C4"] },
    }, {
      type: "removeChart",
      worksheetId: "sheet_1",
      chartId: "chart_1",
      deletedAt: "2026-05-20T12:00:00.000Z",
    }]).json).toMatchObject({
      set: [
        { path: ["teamgrid", "workbook", "worksheetsById", "sheet_1", "chartsById", "chart_1", "title"], value: "Pipeline" },
        { path: ["teamgrid", "workbook", "worksheetsById", "sheet_1", "chartsById", "chart_1", "type"], value: "line" },
        { path: ["teamgrid", "workbook", "worksheetsById", "sheet_1", "chartsById", "chart_1", "categoryAxis"], value: categoryAxis },
        { path: ["teamgrid", "workbook", "worksheetsById", "sheet_1", "chartsById", "chart_1", "legend"], value: { position: "bottom" } },
        { path: ["teamgrid", "workbook", "worksheetsById", "sheet_1", "chartsById", "chart_1", "style"], value: { showGridlines: false, colors: ["#4472C4"] } },
        { path: ["teamgrid", "workbook", "worksheetsById", "sheet_1", "chartsById", "chart_1", "deletedAt"], value: "2026-05-20T12:00:00.000Z" },
      ],
    });
  });
});
