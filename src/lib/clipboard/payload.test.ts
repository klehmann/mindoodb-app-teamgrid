import { describe, expect, it } from "vitest";

import { decodePayload, decodeTsvPayload, rewriteFormulaSource, serializeRange } from "@/lib/clipboard";
import type { Cell } from "@/lib/teamgridDocument";

const cells: Record<string, Cell> = {
  "0:0": { id: "r1:c1", rowId: "r1", columnId: "c1", value: { kind: "number", value: 10 } },
  "0:1": {
    id: "r1:c2",
    rowId: "r1",
    columnId: "c2",
    value: { kind: "number", value: 20 },
    formula: { kind: "formula", source: "=A1*2", references: [], cached: { kind: "number", value: 20 } },
    style: { bold: true },
  },
};

describe("clipboard payloads", () => {
  it("round-trips rich HTML payloads", () => {
    const serialized = serializeRange(
      { worksheetId: "sheet-1", startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
      (row, col) => cells[`${row}:${col}`],
      "cut",
    );

    const decoded = decodePayload(serialized.html, serialized.tsv);

    expect(decoded?.source.rows).toBe(1);
    expect(decoded?.source.cols).toBe(2);
    expect(decoded?.mode).toBe("cut");
    expect(decoded?.cutCellIds).toEqual(["r1:c1", "r1:c2"]);
    expect(decoded?.cells[1].formulaSource).toBe("=A1*2");
    expect(decoded?.cells[1].style).toEqual({ bold: true });
  });

  it("exports Excel-compatible relative formulas in HTML", () => {
    const serialized = serializeRange(
      { worksheetId: "sheet-1", startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
      (row, col) => cells[`${row}:${col}`],
      "copy",
    );

    expect(serialized.html).toContain('x:fmla="=RC[-1]*2"');
    expect(serialized.html).toContain(">20</td>");
  });

  it("prefers Teamgrid JSON over Excel formula metadata", () => {
    const serialized = serializeRange(
      { worksheetId: "sheet-1", startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
      (row, col) => cells[`${row}:${col}`],
      "copy",
    );
    const decoded = decodePayload(serialized.html.replace("=RC[-1]*2", "=R[99]C[99]"), serialized.tsv);

    expect(decoded?.cells[1].formulaSource).toBe("=A1*2");
  });

  it("imports Excel HTML formulas into Teamgrid payloads", () => {
    const decoded = decodePayload(
      '<table><tbody><tr><td>10</td><td x:fmla="=RC[-1]*2">20</td></tr></tbody></table>',
      "10\t20",
    );
    const formulaCell = decoded?.cells[1];

    expect(formulaCell?.formulaSource).toBeTruthy();
    expect(formulaCell?.formulaSource).not.toContain("RC");
    expect(rewriteFormulaSource(formulaCell?.formulaSource ?? "", {
      rows: -decoded!.source.anchor.row,
      cols: 1 - decoded!.source.anchor.col,
    })).toBe("=B1*2");
  });

  it("decodes TSV-only clipboard data", () => {
    const decoded = decodeTsvPayload("1\tHello\n3.5\t");

    expect(decoded?.source.rows).toBe(2);
    expect(decoded?.source.cols).toBe(2);
    expect(decoded?.cells.map((cell) => cell.value)).toEqual([
      { kind: "number", value: 1 },
      { kind: "string", text: "Hello" },
      { kind: "number", value: 3.5 },
      { kind: "empty" },
    ]);
  });
});
