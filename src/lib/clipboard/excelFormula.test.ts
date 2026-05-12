import { describe, expect, it } from "vitest";

import { a1FormulaToRelativeR1C1, excelFormulaToA1 } from "@/lib/clipboard";

describe("Excel formula conversion", () => {
  it("exports same-row A1 references as relative R1C1", () => {
    expect(a1FormulaToRelativeR1C1("=B3*5", { row: 2, col: 2 })).toBe("=RC[-1]*5");
  });

  it("exports row and column offsets as relative R1C1", () => {
    expect(a1FormulaToRelativeR1C1("=B1*5", { row: 2, col: 2 })).toBe("=R[-2]C[-1]*5");
  });

  it("exports A1 ranges as relative R1C1 ranges", () => {
    expect(a1FormulaToRelativeR1C1("=SUM(A1:B2)", { row: 2, col: 2 })).toBe("=SUM(R[-2]C[-2]:R[-1]C[-1])");
  });

  it("does not rewrite references inside strings", () => {
    expect(a1FormulaToRelativeR1C1('=CONCAT("A1", A1)', { row: 2, col: 2 })).toBe('=CONCAT("A1", R[-2]C[-2])');
  });

  it("imports relative R1C1 references as A1", () => {
    expect(excelFormulaToA1("=RC[-1]*5", { row: 2, col: 2 })).toBe("=B3*5");
    expect(excelFormulaToA1("=R[-2]C[-1]*5", { row: 2, col: 2 })).toBe("=B1*5");
  });

  it("imports relative R1C1 ranges as A1 ranges", () => {
    expect(excelFormulaToA1("=SUM(R[-2]C[-2]:R[-1]C[-1])", { row: 2, col: 2 })).toBe("=SUM(A1:B2)");
  });
});
