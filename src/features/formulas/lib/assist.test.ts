import { describe, expect, it } from "vitest";

import {
  getActiveFunctionFragment,
  getEnclosingFunctionName,
  insertFunctionAtCaret,
} from "@/features/formulas/lib/assist";

describe("formula assist helpers", () => {
  it("detects function fragments at expression boundaries", () => {
    expect(getActiveFunctionFragment("=SU")).toBe("SU");
    expect(getActiveFunctionFragment("=A1+AV")).toBe("AV");
    expect(getActiveFunctionFragment("=SUM(AV")).toBe("AV");
  });

  it("does not treat cell references as function fragments", () => {
    expect(getActiveFunctionFragment("=B1")).toBe("");
    expect(getActiveFunctionFragment("=B1*5")).toBe("");
  });

  it("finds the innermost enclosing function", () => {
    expect(getEnclosingFunctionName("=SUM(A1, AVERAGE(B1:B3", 24)).toBe("AVERAGE");
    expect(getEnclosingFunctionName("=SUM(A1)", 8)).toBe(null);
  });

  it("ignores parentheses in string literals", () => {
    expect(getEnclosingFunctionName('=CONCAT(")", A1', 15)).toBe("CONCAT");
  });

  it("inserts a function call at the caret", () => {
    expect(insertFunctionAtCaret("=SU", 3, "SUM")).toEqual({ next: "=SUM(", nextCaret: 5 });
    expect(insertFunctionAtCaret("=A1+", 4, "AVERAGE")).toEqual({ next: "=A1+AVERAGE(", nextCaret: 12 });
  });
});
