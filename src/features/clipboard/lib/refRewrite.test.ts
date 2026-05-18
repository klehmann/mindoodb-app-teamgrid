import { describe, expect, it } from "vitest";

import { rewriteFormulaSource } from "@/features/clipboard/lib";

describe("formula reference rewriting", () => {
  it("shifts relative cell and range references", () => {
    expect(rewriteFormulaSource("=SUM(A1:B2)+C3", { rows: 2, cols: 1 })).toBe("=SUM(B3:C4)+D5");
  });

  it("turns out-of-bounds shifted references into #REF!", () => {
    expect(rewriteFormulaSource("=A1+B2", { rows: -1, cols: 0 })).toBe("=#REF!+B1");
  });

  it("does not rewrite references in string literals", () => {
    expect(rewriteFormulaSource('=CONCAT("A1", A1)', { rows: 1, cols: 1 })).toBe('=CONCAT("A1", B2)');
  });

  it("moves only references inside the source range when requested", () => {
    expect(rewriteFormulaSource("=A1+B1+C1", { rows: 2, cols: 0 }, {
      insideRange: { startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
    })).toBe("=A3+B3+C1");
  });
});
