import { describe, expect, it } from "vitest";
import { canMutateGrid, isGridSessionReadOnly } from "@/lib/capabilities";

describe("capability gates", () => {
  it("opens time travel and historical revisions read-only", () => {
    expect(isGridSessionReadOnly({ timeTravelDate: Date.now(), viewingHistorical: false })).toBe(true);
    expect(isGridSessionReadOnly({ timeTravelDate: null, viewingHistorical: true })).toBe(true);
    expect(isGridSessionReadOnly({ timeTravelDate: null, viewingHistorical: false })).toBe(false);
  });

  it("requires update access, a document, dirty state, and a live revision before saving", () => {
    expect(canMutateGrid({
      canUpdate: true,
      hasDocument: true,
      dirty: true,
      timeTravelDate: null,
      viewingHistorical: false,
    })).toBe(true);
    expect(canMutateGrid({
      canUpdate: true,
      hasDocument: true,
      dirty: true,
      timeTravelDate: Date.now(),
      viewingHistorical: false,
    })).toBe(false);
  });
});
