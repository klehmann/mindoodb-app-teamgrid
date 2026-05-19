import { describe, expect, it } from "vitest";
import { createFormulaCellKey, createFormulaContext, evaluateFormula, parseFormula, buildDependencyGraph, collectDirtyFormulaCells, renderFormulaSource } from "@/features/formulas/lib";
import { SUPPORTED_FORMULA_NAMES } from "@/features/formulas/lib/fastFormulaParserAdapter";
import { FUNCTION_REGISTRY, suggestFunctions } from "@/features/formulas/lib/functionRegistry";
import { projectWorksheet } from "@/features/grid/lib/gridProjection";
import { createCellId, createId, createTeamGridDocument, getFirstVisibleWorksheet, type Worksheet } from "@/features/document/lib/teamgridDocument";

describe("formula subsystem", () => {
  it("exposes every supported formula in content assist", () => {
    expect(SUPPORTED_FORMULA_NAMES.length).toBeGreaterThan(250);
    expect(Object.keys(FUNCTION_REGISTRY).sort()).toEqual([...SUPPORTED_FORMULA_NAMES].sort());
    expect(suggestFunctions("UPP").map((definition) => definition.name)).toContain("UPPER");
    expect(suggestFunctions("COUNTI").map((definition) => definition.name)).toEqual(["COUNTIF"]);
    expect(FUNCTION_REGISTRY.ABS.description).toMatch(/absolute value/i);
    expect(FUNCTION_REGISTRY.ACOS.description).toMatch(/arccosine/i);
  });

  it("normalizes visible cell references to stable row and column ids", () => {
    const envelope = createTeamGridDocument();
    const worksheet = getFirstVisibleWorksheet(envelope.teamgrid)!;
    const projection = projectWorksheet(worksheet);
    const context = createFormulaContext(envelope.teamgrid.workbook);

    const parsed = parseFormula("=A1+B2", worksheet.id, context);

    expect("references" in parsed ? parsed.references : []).toEqual([
      { kind: "cell", worksheetId: worksheet.id, rowId: projection.rows[0].id, columnId: projection.columns[0].id },
      { kind: "cell", worksheetId: worksheet.id, rowId: projection.rows[1].id, columnId: projection.columns[1].id },
    ]);
  });

  it("evaluates arithmetic and functions over ranges", () => {
    const envelope = createTeamGridDocument();
    const worksheet = getFirstVisibleWorksheet(envelope.teamgrid)!;
    const projection = projectWorksheet(worksheet);
    const context = createFormulaContext(envelope.teamgrid.workbook);
    const a1 = createCellId(projection.rows[0].id, projection.columns[0].id);
    const a2 = createCellId(projection.rows[1].id, projection.columns[0].id);
    worksheet.cellsById[a1] = { id: a1, rowId: projection.rows[0].id, columnId: projection.columns[0].id, value: { kind: "number", value: 2 } };
    worksheet.cellsById[a2] = { id: a2, rowId: projection.rows[1].id, columnId: projection.columns[0].id, value: { kind: "number", value: 3 } };

    expect(evaluateFormula("=A1*5", worksheet.id, context).result).toEqual({ kind: "number", value: 10 });
    expect(evaluateFormula("=SUM(A1:A2)", worksheet.id, context).result).toEqual({ kind: "number", value: 5 });
    expect(evaluateFormula("=MIN(A1:A2)", worksheet.id, context).result).toEqual({ kind: "number", value: 2 });
    expect(evaluateFormula("=MAX(A1:A2)", worksheet.id, context).result).toEqual({ kind: "number", value: 3 });
  });

  it("evaluates a broader Excel-compatible formula subset", () => {
    const envelope = createTeamGridDocument();
    const worksheet = getFirstVisibleWorksheet(envelope.teamgrid)!;
    const projection = projectWorksheet(worksheet);
    const context = createFormulaContext(envelope.teamgrid.workbook);
    const a1 = createCellId(projection.rows[0].id, projection.columns[0].id);
    const a2 = createCellId(projection.rows[1].id, projection.columns[0].id);
    const b1 = createCellId(projection.rows[0].id, projection.columns[1].id);
    const b13 = createCellId(projection.rows[12].id, projection.columns[1].id);
    worksheet.cellsById[a1] = { id: a1, rowId: projection.rows[0].id, columnId: projection.columns[0].id, value: { kind: "number", value: 2.45 } };
    worksheet.cellsById[a2] = { id: a2, rowId: projection.rows[1].id, columnId: projection.columns[0].id, value: { kind: "number", value: 8 } };
    worksheet.cellsById[b1] = { id: b1, rowId: projection.rows[0].id, columnId: projection.columns[1].id, value: { kind: "string", text: "ready" } };
    worksheet.cellsById[b13] = { id: b13, rowId: projection.rows[12].id, columnId: projection.columns[1].id, value: { kind: "string", text: "mixed case" } };

    expect(evaluateFormula("=IF(A2>5,\"yes\",\"no\")", worksheet.id, context).result).toEqual({ kind: "string", value: "yes" });
    expect(evaluateFormula("=ROUND(A1,1)", worksheet.id, context).result).toEqual({ kind: "number", value: 2.5 });
    expect(evaluateFormula("=COUNTIF(A1:A2,\">3\")", worksheet.id, context).result).toEqual({ kind: "number", value: 1 });
    expect(evaluateFormula("=LOWER(B1)", worksheet.id, context).result).toEqual({ kind: "string", value: "ready" });
    expect(evaluateFormula("=UPPER(B13)", worksheet.id, context).result).toEqual({ kind: "string", value: "MIXED CASE" });
    expect(evaluateFormula("=YEAR(DATE(2026,5,19))", worksheet.id, context).result).toEqual({ kind: "number", value: 2026 });
  });

  it("maps formula errors into Teamgrid error codes", () => {
    const envelope = createTeamGridDocument();
    const worksheet = getFirstVisibleWorksheet(envelope.teamgrid)!;
    const context = createFormulaContext(envelope.teamgrid.workbook);

    expect(evaluateFormula("=1/0", worksheet.id, context).result).toEqual({ kind: "error", code: "#DIV/0!" });
    expect(evaluateFormula("=UNKNOWN_FN(1)", worksheet.id, context).result).toEqual({ kind: "error", code: "#NAME?" });
    expect(evaluateFormula("=A999+1", worksheet.id, context).result).toEqual({ kind: "error", code: "#REF!" });
    expect(evaluateFormula("=\"text\"+1", worksheet.id, context).result).toEqual({ kind: "error", code: "#VALUE!" });
  });

  it("keeps formula-to-formula cycles bounded", () => {
    const envelope = createTeamGridDocument();
    const worksheet = getFirstVisibleWorksheet(envelope.teamgrid)!;
    const projection = projectWorksheet(worksheet);
    const context = createFormulaContext(envelope.teamgrid.workbook);
    const a1 = createCellId(projection.rows[0].id, projection.columns[0].id);
    const b1 = createCellId(projection.rows[0].id, projection.columns[1].id);
    const a1Formula = evaluateFormula("=B1+1", worksheet.id, context);
    const b1Formula = evaluateFormula("=A1+1", worksheet.id, context);
    worksheet.cellsById[a1] = {
      id: a1,
      rowId: projection.rows[0].id,
      columnId: projection.columns[0].id,
      value: { kind: "empty" },
      formula: { kind: "formula", source: "=B1+1", references: a1Formula.references, cached: a1Formula.result },
    };
    worksheet.cellsById[b1] = {
      id: b1,
      rowId: projection.rows[0].id,
      columnId: projection.columns[1].id,
      value: { kind: "empty" },
      formula: { kind: "formula", source: "=A1+1", references: b1Formula.references, cached: b1Formula.result },
    };

    expect(evaluateFormula("=B1+1", worksheet.id, createFormulaContext(envelope.teamgrid.workbook)).result).toEqual({ kind: "error", code: "#CYCLE!" });
  });

  it("tracks dirty dependents from stable formula references", () => {
    const envelope = createTeamGridDocument();
    const worksheet = getFirstVisibleWorksheet(envelope.teamgrid)!;
    const projection = projectWorksheet(worksheet);
    const context = createFormulaContext(envelope.teamgrid.workbook);
    const a1 = createCellId(projection.rows[0].id, projection.columns[0].id);
    const b1 = createCellId(projection.rows[0].id, projection.columns[1].id);
    const evaluated = evaluateFormula("=A1+1", worksheet.id, context);
    worksheet.cellsById[b1] = {
      id: b1,
      rowId: projection.rows[0].id,
      columnId: projection.columns[1].id,
      value: { kind: "number", value: 1 },
      formula: {
        kind: "formula",
        source: "=A1+1",
        references: evaluated.references,
        cached: evaluated.result,
      },
    };

    const graph = buildDependencyGraph(envelope.teamgrid.workbook);

    expect(collectDirtyFormulaCells(graph, [createFormulaCellKey(worksheet.id, a1)])).toEqual([createFormulaCellKey(worksheet.id, b1)]);
  });

  it("maps whole-column references to stable column dependencies", () => {
    const envelope = createTeamGridDocument();
    const worksheet = getFirstVisibleWorksheet(envelope.teamgrid)!;
    const projection = projectWorksheet(worksheet);
    const context = createFormulaContext(envelope.teamgrid.workbook);

    const parsed = parseFormula("=SUM(A:A)", worksheet.id, context);

    expect("references" in parsed ? parsed.references : []).toEqual([
      { kind: "column", worksheetId: worksheet.id, columnId: projection.columns[0].id },
    ]);
  });

  it("evaluates and renders cross-sheet references from stable ids", () => {
    const envelope = createTeamGridDocument();
    const firstWorksheet = getFirstVisibleWorksheet(envelope.teamgrid)!;
    const secondWorksheet = cloneWorksheet(firstWorksheet, "sheet_second", "Sheet 2");
    envelope.teamgrid.workbook.worksheetOrder.push(secondWorksheet.id);
    envelope.teamgrid.workbook.worksheetsById[secondWorksheet.id] = secondWorksheet;
    const firstProjection = projectWorksheet(firstWorksheet);
    const secondProjection = projectWorksheet(secondWorksheet);
    const c10 = createCellId(firstProjection.rows[9].id, firstProjection.columns[2].id);
    const c11 = createCellId(firstProjection.rows[10].id, firstProjection.columns[2].id);
    firstWorksheet.cellsById[c10] = { id: c10, rowId: firstProjection.rows[9].id, columnId: firstProjection.columns[2].id, value: { kind: "number", value: 4 } };
    firstWorksheet.cellsById[c11] = { id: c11, rowId: firstProjection.rows[10].id, columnId: firstProjection.columns[2].id, value: { kind: "number", value: 6 } };
    const context = createFormulaContext(envelope.teamgrid.workbook);

    const evaluated = evaluateFormula("=sum(Sheet 1!C10:C11)", secondWorksheet.id, context);

    expect(evaluated.result).toEqual({ kind: "number", value: 10 });
    expect(evaluated.references).toEqual([{
      kind: "range",
      worksheetId: firstWorksheet.id,
      startRowId: firstProjection.rows[9].id,
      endRowId: firstProjection.rows[10].id,
      startColumnId: firstProjection.columns[2].id,
      endColumnId: firstProjection.columns[2].id,
    }]);

    firstWorksheet.title = "Renamed Sheet";
    const insertedRowId = createId("row");
    firstWorksheet.rowOrder.splice(0, 0, insertedRowId);
    firstWorksheet.rowsById[insertedRowId] = { id: insertedRowId };
    const renamedContext = createFormulaContext(envelope.teamgrid.workbook);
    expect(renderFormulaSource({ source: "=sum(Sheet 1!C10:C11)", segments: evaluated.segments }, secondWorksheet.id, renamedContext)).toBe("=sum('Renamed Sheet'!C11:C12)");
    expect(secondProjection.rows).toHaveLength(24);
  });

  it("does not highlight local ranges for unquoted cross-sheet names with spaces", () => {
    const envelope = createTeamGridDocument();
    const firstWorksheet = getFirstVisibleWorksheet(envelope.teamgrid)!;
    const secondWorksheet = cloneWorksheet(firstWorksheet, "sheet_second", "Sheet 2");
    envelope.teamgrid.workbook.worksheetOrder.push(secondWorksheet.id);
    envelope.teamgrid.workbook.worksheetsById[secondWorksheet.id] = secondWorksheet;
    const firstProjection = projectWorksheet(firstWorksheet);
    const secondProjection = projectWorksheet(secondWorksheet);
    const context = createFormulaContext(envelope.teamgrid.workbook);

    const parsed = parseFormula("=sum(Sheet 2!A1:C1)", firstWorksheet.id, context);

    expect("references" in parsed ? parsed.references : []).toEqual([{
      kind: "range",
      worksheetId: secondWorksheet.id,
      startRowId: secondProjection.rows[0].id,
      endRowId: secondProjection.rows[0].id,
      startColumnId: secondProjection.columns[0].id,
      endColumnId: secondProjection.columns[2].id,
    }]);
    expect("references" in parsed ? parsed.references[0]?.worksheetId : null).not.toBe(firstWorksheet.id);
    expect(firstProjection.rows).toHaveLength(24);
  });
});

function cloneWorksheet(worksheet: Worksheet, id: string, title: string): Worksheet {
  return {
    ...structuredClone(worksheet),
    id,
    title,
    rowOrder: [...worksheet.rowOrder],
    columnOrder: [...worksheet.columnOrder],
    rowsById: structuredClone(worksheet.rowsById),
    columnsById: structuredClone(worksheet.columnsById),
    cellsById: {},
  };
}
