import { describe, expect, it } from "vitest";
import { evaluateFormula, parseFormula, buildDependencyGraph, collectDirtyFormulaCells } from "@/lib/formulas";
import { projectWorksheet } from "@/lib/gridProjection";
import { createCellId, createTeamGridDocument, getFirstVisibleWorksheet } from "@/lib/teamgridDocument";

describe("formula subsystem", () => {
  it("normalizes visible cell references to stable row and column ids", () => {
    const envelope = createTeamGridDocument();
    const worksheet = getFirstVisibleWorksheet(envelope.teamgrid)!;
    const projection = projectWorksheet(worksheet);

    const parsed = parseFormula("=A1+B2", worksheet.id, projection);

    expect("references" in parsed ? parsed.references : []).toEqual([
      { kind: "cell", worksheetId: worksheet.id, rowId: projection.rows[0].id, columnId: projection.columns[0].id },
      { kind: "cell", worksheetId: worksheet.id, rowId: projection.rows[1].id, columnId: projection.columns[1].id },
    ]);
  });

  it("evaluates arithmetic and functions over ranges", () => {
    const envelope = createTeamGridDocument();
    const worksheet = getFirstVisibleWorksheet(envelope.teamgrid)!;
    const projection = projectWorksheet(worksheet);
    const a1 = createCellId(projection.rows[0].id, projection.columns[0].id);
    const a2 = createCellId(projection.rows[1].id, projection.columns[0].id);
    worksheet.cellsById[a1] = { id: a1, rowId: projection.rows[0].id, columnId: projection.columns[0].id, value: { kind: "number", value: 2 } };
    worksheet.cellsById[a2] = { id: a2, rowId: projection.rows[1].id, columnId: projection.columns[0].id, value: { kind: "number", value: 3 } };

    expect(evaluateFormula("=A1*5", worksheet, projection).result).toEqual({ kind: "number", value: 10 });
    expect(evaluateFormula("=SUM(A1:A2)", worksheet, projection).result).toEqual({ kind: "number", value: 5 });
  });

  it("tracks dirty dependents from stable formula references", () => {
    const envelope = createTeamGridDocument();
    const worksheet = getFirstVisibleWorksheet(envelope.teamgrid)!;
    const projection = projectWorksheet(worksheet);
    const a1 = createCellId(projection.rows[0].id, projection.columns[0].id);
    const b1 = createCellId(projection.rows[0].id, projection.columns[1].id);
    const evaluated = evaluateFormula("=A1+1", worksheet, projection);
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

    const graph = buildDependencyGraph(worksheet);

    expect(collectDirtyFormulaCells(graph, [a1])).toEqual([b1]);
  });
});
