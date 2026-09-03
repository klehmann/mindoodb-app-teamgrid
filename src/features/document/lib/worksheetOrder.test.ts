import { describe, expect, it } from "vitest";

import type { Workbook, Worksheet, WorksheetId } from "@/features/document/lib/teamgridDocument";
import {
  isWorksheetOrderSound,
  planWorksheetMove,
  planWorksheetMoveTo,
  planWorksheetNudge,
  resolveWorksheetOrder,
} from "@/features/document/lib/worksheetOrder";

function createWorkbook(
  order: WorksheetId[],
  sheets: readonly { id: WorksheetId; deleted?: boolean }[] = order.map((id) => ({ id })),
): Workbook {
  return {
    id: "book",
    worksheetOrder: order,
    worksheetsById: Object.fromEntries(sheets.map((sheet) => [
      sheet.id,
      { id: sheet.id, title: sheet.id, deletedAt: sheet.deleted ? "2026-01-01" : undefined } as Worksheet,
    ])),
  };
}

/** What the host does with the patch a move serializes to. */
function applyMovePatch(order: readonly WorksheetId[], plan: { fromIndex: number; toIndex: number }) {
  const next = [...order];
  const [moved] = next.splice(plan.fromIndex, 1);
  next.splice(plan.toIndex, 0, moved);
  return next;
}

describe("resolveWorksheetOrder", () => {
  it("leaves a sound list alone, tombstones included", () => {
    const workbook = createWorkbook(["a", "b", "c"], [{ id: "a" }, { id: "b", deleted: true }, { id: "c" }]);
    expect(resolveWorksheetOrder(workbook)).toEqual(["a", "b", "c"]);
    expect(isWorksheetOrderSound(workbook)).toBe(true);
  });

  it("keeps the first of a doubled id, which is what two concurrent moves produce", () => {
    const workbook = createWorkbook(["a", "b", "a", "c"], [{ id: "a" }, { id: "b" }, { id: "c" }]);
    expect(resolveWorksheetOrder(workbook)).toEqual(["a", "b", "c"]);
    expect(isWorksheetOrderSound(workbook)).toBe(false);
  });

  it("drops ids the workbook does not know", () => {
    const workbook = createWorkbook(["a", "ghost", "b"], [{ id: "a" }, { id: "b" }]);
    expect(resolveWorksheetOrder(workbook)).toEqual(["a", "b"]);
  });

  it("appends worksheets missing from the list so no sheet is left without a tab", () => {
    const workbook = createWorkbook(["b"], [{ id: "a" }, { id: "b" }, { id: "c" }]);
    expect(resolveWorksheetOrder(workbook)).toEqual(["b", "a", "c"]);
  });
});

describe("planWorksheetMove", () => {
  it("lands after the target when dragging right", () => {
    const plan = planWorksheetMove(["a", "b", "c"], "a", "c");
    expect(plan?.order).toEqual(["b", "c", "a"]);
  });

  it("lands before the target when dragging left", () => {
    const plan = planWorksheetMove(["a", "b", "c"], "c", "a");
    expect(plan?.order).toEqual(["c", "a", "b"]);
  });

  it("refuses a move that goes nowhere or names an id the list lacks", () => {
    expect(planWorksheetMove(["a", "b"], "a", "a")).toBeNull();
    expect(planWorksheetMove(["a", "b"], "ghost", "a")).toBeNull();
    expect(planWorksheetMove(["a", "b"], "a", "ghost")).toBeNull();
  });

  it("emits indices whose delete-then-insert reproduces the planned order", () => {
    const order = ["a", "b", "c", "d"];
    for (const dragged of order) {
      for (const target of order) {
        const plan = planWorksheetMove(order, dragged, target);
        if (plan) {
          expect(applyMovePatch(order, plan)).toEqual(plan.order);
        }
      }
    }
  });
});

describe("planWorksheetMoveTo", () => {
  it("puts the tab at the index it was sent to", () => {
    const plan = planWorksheetMoveTo(["a", "b", "c", "d"], "a", 2);

    expect(plan?.order).toEqual(["b", "c", "a", "d"]);
    expect(plan?.order.indexOf("a")).toBe(2);
  });

  it("refuses an index outside the list, or the one it already has", () => {
    expect(planWorksheetMoveTo(["a", "b"], "a", 0)).toBeNull();
    expect(planWorksheetMoveTo(["a", "b"], "a", 2)).toBeNull();
    expect(planWorksheetMoveTo(["a", "b"], "a", -1)).toBeNull();
    expect(planWorksheetMoveTo(["a", "b"], "ghost", 1)).toBeNull();
  });

  it("reports a toIndex that is where the tab ends up, which is what a drop sends back", () => {
    const order = ["a", "b", "c", "d"];
    for (const id of order) {
      for (let toIndex = 0; toIndex < order.length; toIndex += 1) {
        const plan = planWorksheetMoveTo(order, id, toIndex);
        if (plan) {
          expect(plan.order.indexOf(id)).toBe(plan.toIndex);
          expect(applyMovePatch(order, plan)).toEqual(plan.order);
        }
      }
    }
  });
});

describe("planWorksheetNudge", () => {
  it("swaps a tab with its neighbour", () => {
    const workbook = createWorkbook(["a", "b", "c"]);
    expect(planWorksheetNudge(workbook, "b", 1)?.order).toEqual(["a", "c", "b"]);
    expect(planWorksheetNudge(workbook, "b", -1)?.order).toEqual(["b", "a", "c"]);
  });

  it("steps past a tombstone rather than into its slot", () => {
    const workbook = createWorkbook(["a", "gone", "b"], [{ id: "a" }, { id: "gone", deleted: true }, { id: "b" }]);
    const plan = planWorksheetNudge(workbook, "a", 1);

    expect(plan?.order).toEqual(["gone", "b", "a"]);
    expect(plan?.order.filter((id) => id !== "gone")).toEqual(["b", "a"]);
  });

  it("stays put at either edge", () => {
    const workbook = createWorkbook(["a", "b"]);
    expect(planWorksheetNudge(workbook, "a", -1)).toBeNull();
    expect(planWorksheetNudge(workbook, "b", 1)).toBeNull();
  });

  it("plans against the repaired list, so an unsound one still moves", () => {
    const workbook = createWorkbook(["a", "b", "a"], [{ id: "a" }, { id: "b" }]);
    expect(planWorksheetNudge(workbook, "a", 1)?.order).toEqual(["b", "a"]);
  });
});
