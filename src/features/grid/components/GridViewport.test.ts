import { mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";
import { nextTick } from "vue";

import GridViewport from "@/features/grid/components/GridViewport.vue";
import { projectWorksheet } from "@/features/grid/lib/gridProjection";
import { createCellId, createTeamGridDocument } from "@/features/document/lib/teamgridDocument";
import type { CellSelectionRange } from "@/features/grid/composables/useSelection";

describe("GridViewport inline editing", () => {
  it("commits a plain typed edit and selects the clicked cell", async () => {
    const { wrapper, firstCellId, secondCellId } = mountGrid();

    await wrapper.find(`[data-test-cell-id="${firstCellId}"]`).trigger("keydown", { key: "a" });
    expect(wrapper.find("textarea.grid-cell__editor").exists()).toBe(true);

    await wrapper.find(`[data-test-cell-id="${secondCellId}"]`).trigger("mousedown", { button: 0 });

    expect(wrapper.emitted("commit")).toHaveLength(1);
    expect(wrapper.emitted("commit")![0][1]).toBe("a");
    expect(wrapper.emitted("select")!.at(-1)?.[0]).toMatchObject({ id: secondCellId });
    expect(wrapper.find("textarea.grid-cell__editor").exists()).toBe(false);
  });

  it("keeps formula picking active when clicking another cell during formula edit", async () => {
    const { wrapper, firstCellId, secondCellId } = mountGrid();

    await wrapper.find(`[data-test-cell-id="${firstCellId}"]`).trigger("keydown", { key: "=" });
    await wrapper.find(`[data-test-cell-id="${secondCellId}"]`).trigger("mousedown", { button: 0 });

    expect(wrapper.emitted("commit")).toBeUndefined();
    expect((wrapper.find("textarea.grid-cell__editor").element as HTMLTextAreaElement).value).toBe("=B1");
  });

  it("flushes an active inline edit for toolbar save actions", async () => {
    const { wrapper, firstCellId } = mountGrid();

    await wrapper.find(`[data-test-cell-id="${firstCellId}"]`).trigger("keydown", { key: "a" });
    await wrapper.find("textarea.grid-cell__editor").setValue("saved draft");

    expect((wrapper.vm as unknown as { flushPendingEdit: () => boolean }).flushPendingEdit()).toBe(true);

    expect(wrapper.emitted("commit")).toHaveLength(1);
    expect(wrapper.emitted("commit")![0][1]).toBe("saved draft");
    await nextTick();
    expect(wrapper.find("textarea.grid-cell__editor").exists()).toBe(false);
  });
});

describe("GridViewport cell context menu", () => {
  it("emits the current range when right-clicking inside the selection", async () => {
    const { wrapper, firstCellId, secondCellId } = mountGrid();

    await wrapper.setProps({ selectedRange: { startCellId: firstCellId, endCellId: secondCellId } });

    await wrapper.find(`[data-test-cell-id="${secondCellId}"]`).trigger("contextmenu");

    const eventPayload = wrapper.emitted("cell-context")?.[0]?.[0];
    expect(eventPayload).toMatchObject({
      cell: { id: secondCellId },
      address: "B1",
      range: { startCellId: firstCellId, endCellId: secondCellId },
    });
  });

  it("emits a single-cell range when right-clicking outside the selection", async () => {
    const { wrapper, firstCellId, secondCellId } = mountGrid();

    await wrapper.find(`[data-test-cell-id="${secondCellId}"]`).trigger("contextmenu");

    const eventPayload = wrapper.emitted("cell-context")?.[0]?.[0];
    expect(eventPayload).toMatchObject({
      cell: { id: secondCellId },
      address: "B1",
      range: { startCellId: secondCellId, endCellId: secondCellId },
    });
    expect(eventPayload).not.toMatchObject({
      range: { startCellId: firstCellId, endCellId: firstCellId },
    });
  });
});

describe("GridViewport mouse range selection", () => {
  it("extends the selection from the active cell when shift+clicking another cell", async () => {
    const { wrapper, firstCellId, secondRowSecondCellId } = mountGrid();

    await wrapper.find(`[data-test-cell-id="${secondRowSecondCellId}"]`).trigger("mousedown", { button: 0, shiftKey: true });

    expect(wrapper.emitted("select")!.at(-1)?.[0]).toMatchObject({ id: secondRowSecondCellId });
    expect(wrapper.emitted("select-range")!.at(-1)?.[0]).toEqual({
      startCellId: firstCellId,
      endCellId: secondRowSecondCellId,
    });
  });

  it("keeps the original range anchor when shift+clicking after an existing range", async () => {
    const { wrapper, firstCellId, secondCellId, secondRowSecondCellId } = mountGrid();
    await wrapper.setProps({
      selectedCellId: secondCellId,
      selectedRange: { startCellId: firstCellId, endCellId: secondCellId },
    });

    await wrapper.find(`[data-test-cell-id="${secondRowSecondCellId}"]`).trigger("mousedown", { button: 0, shiftKey: true });

    expect(wrapper.emitted("select-range")!.at(-1)?.[0]).toEqual({
      startCellId: firstCellId,
      endCellId: secondRowSecondCellId,
    });
  });

  it("pushes the prior range onto additional ranges when ctrl+clicking another cell", async () => {
    const { wrapper, firstCellId, secondRowSecondCellId } = mountGrid();

    await wrapper.find(`[data-test-cell-id="${secondRowSecondCellId}"]`).trigger("mousedown", { button: 0, ctrlKey: true });

    expect(wrapper.emitted("add-range")!.at(-1)?.[0]).toEqual({
      startCellId: firstCellId,
      endCellId: firstCellId,
    });
    expect(wrapper.emitted("select")!.at(-1)?.[0]).toMatchObject({ id: secondRowSecondCellId });
    expect(wrapper.emitted("select-range")!.at(-1)?.[0]).toEqual({
      startCellId: secondRowSecondCellId,
      endCellId: secondRowSecondCellId,
    });
    expect(wrapper.emitted("clear-additional-ranges")).toBeUndefined();
  });

  it("treats meta+click the same as ctrl+click for macOS users", async () => {
    const { wrapper, firstCellId, secondRowSecondCellId } = mountGrid();

    await wrapper.find(`[data-test-cell-id="${secondRowSecondCellId}"]`).trigger("mousedown", { button: 0, metaKey: true });

    expect(wrapper.emitted("add-range")!.at(-1)?.[0]).toEqual({
      startCellId: firstCellId,
      endCellId: firstCellId,
    });
    expect(wrapper.emitted("select-range")!.at(-1)?.[0]).toEqual({
      startCellId: secondRowSecondCellId,
      endCellId: secondRowSecondCellId,
    });
  });

  it("highlights cells in additional ranges alongside the primary range", async () => {
    const { wrapper, firstCellId, secondRowSecondCellId } = mountGrid();

    await wrapper.setProps({
      additionalRanges: [{ startCellId: firstCellId, endCellId: firstCellId }],
      selectedCellId: secondRowSecondCellId,
      selectedRange: { startCellId: secondRowSecondCellId, endCellId: secondRowSecondCellId },
    });

    expect(wrapper.find(`[data-test-cell-id="${firstCellId}"]`).classes()).toContain("grid-cell--range-selected");
    expect(wrapper.find(`[data-test-cell-id="${secondRowSecondCellId}"]`).classes()).toContain("grid-cell--range-selected");
  });

  it("clears the additional ranges on a plain click", async () => {
    const { wrapper, firstCellId, secondCellId } = mountGrid();
    await wrapper.setProps({
      additionalRanges: [{ startCellId: firstCellId, endCellId: firstCellId }],
    });

    await wrapper.find(`[data-test-cell-id="${secondCellId}"]`).trigger("mousedown", { button: 0 });

    expect(wrapper.emitted("clear-additional-ranges")).toHaveLength(1);
    expect(wrapper.emitted("add-range")).toBeUndefined();
  });

  it("deselects a single cell from the primary range when ctrl+clicking it", async () => {
    const { wrapper, firstRow, secondRow, firstColumn, secondColumn, thirdColumn } = mountGrid();
    // Primary covers A1:C2; ctrl+click the middle cell (B1) of row 1.
    const a1 = createCellId(firstRow.id, firstColumn.id);
    const b1 = createCellId(firstRow.id, secondColumn.id);
    const c2 = createCellId(secondRow.id, thirdColumn.id);
    await wrapper.setProps({
      selectedCellId: a1,
      selectedRange: { startCellId: a1, endCellId: c2 },
    });

    await wrapper.find(`[data-test-cell-id="${b1}"]`).trigger("mousedown", { button: 0, ctrlKey: true });

    // The original A1:C2 rectangle is split around B1 into:
    //   - bottom strip (row 2): A2:C2
    //   - left sliver (row 1, col A): A1:A1
    //   - right sliver (row 1, col C): C1:C1
    const a2 = createCellId(secondRow.id, firstColumn.id);
    const c1 = createCellId(firstRow.id, thirdColumn.id);
    expect(wrapper.emitted("select-range")!.at(-1)?.[0]).toEqual({ startCellId: a2, endCellId: c2 });
    expect(wrapper.emitted("set-additional-ranges")!.at(-1)?.[0]).toEqual([
      { startCellId: a1, endCellId: a1 },
      { startCellId: c1, endCellId: c1 },
    ]);
  });

  it("deselects a cell that lives in an additional range and leaves the primary intact", async () => {
    const { wrapper, firstRow, firstColumn, secondColumn } = mountGrid();
    const a1 = createCellId(firstRow.id, firstColumn.id);
    const b1 = createCellId(firstRow.id, secondColumn.id);
    await wrapper.setProps({
      selectedCellId: a1,
      selectedRange: { startCellId: a1, endCellId: a1 },
      additionalRanges: [{ startCellId: b1, endCellId: b1 }],
    });

    await wrapper.find(`[data-test-cell-id="${b1}"]`).trigger("mousedown", { button: 0, ctrlKey: true });

    expect(wrapper.emitted("set-additional-ranges")!.at(-1)?.[0]).toEqual([]);
    expect(wrapper.emitted("select-range")!.at(-1)?.[0]).toEqual({ startCellId: a1, endCellId: a1 });
  });

  it("refuses to deselect the only remaining selected cell", async () => {
    const { wrapper, firstCellId } = mountGrid();

    await wrapper.find(`[data-test-cell-id="${firstCellId}"]`).trigger("mousedown", { button: 0, ctrlKey: true });

    expect(wrapper.emitted("set-additional-ranges")).toBeUndefined();
    expect(wrapper.emitted("add-range")).toBeUndefined();
  });

  it("moves the active cell when ctrl+clicking deselects the current active cell", async () => {
    const { wrapper, firstRow, secondRow, firstColumn } = mountGrid();
    const a1 = createCellId(firstRow.id, firstColumn.id);
    const a2 = createCellId(secondRow.id, firstColumn.id);
    await wrapper.setProps({
      selectedCellId: a1,
      selectedRange: { startCellId: a1, endCellId: a2 },
    });

    await wrapper.find(`[data-test-cell-id="${a1}"]`).trigger("mousedown", { button: 0, ctrlKey: true });

    // A1:A2 minus A1 collapses to the single-cell A2 (only a "bottom strip" survives).
    expect(wrapper.emitted("select-range")!.at(-1)?.[0]).toEqual({ startCellId: a2, endCellId: a2 });
    expect(wrapper.emitted("select")!.at(-1)?.[0]).toMatchObject({ id: a2 });
  });

  it("deselects the intersection cell of a row + column multi-selection", async () => {
    const { wrapper, firstRow, secondRow, firstColumn, secondColumn, lastRow, lastColumn } = mountGrid();
    // Primary covers row 2 (entire row); additional covers column B (entire column).
    // The intersection is the cell at (row 2, col B).
    const rowStart = createCellId(secondRow.id, firstColumn.id);
    const rowEnd = createCellId(secondRow.id, lastColumn.id);
    const columnStart = createCellId(firstRow.id, secondColumn.id);
    const columnEnd = createCellId(lastRow.id, secondColumn.id);
    const intersection = createCellId(secondRow.id, secondColumn.id);
    await wrapper.setProps({
      selectedCellId: rowStart,
      selectedRange: { startCellId: rowStart, endCellId: rowEnd },
      additionalRanges: [{ startCellId: columnStart, endCellId: columnEnd }],
    });

    await wrapper.find(`[data-test-cell-id="${intersection}"]`).trigger("mousedown", { button: 0, ctrlKey: true });

    // Row 2 split around (row 2, col B) leaves the row-2-left sliver and the row-2-right strip.
    // Column B split around (row 2, col B) leaves the column-B-above and column-B-below strips.
    // The first surviving fragment becomes the new primary; the rest go into additional.
    const lastRangePayload = wrapper.emitted("set-additional-ranges")!.at(-1)?.[0] as CellSelectionRange[];
    const newPrimary = wrapper.emitted("select-range")!.at(-1)?.[0] as CellSelectionRange;
    const allFragments = [newPrimary, ...lastRangePayload];
    // The deselected cell must not appear in any fragment.
    for (const fragment of allFragments) {
      expect(fragment.startCellId).not.toBe(intersection);
      expect(fragment.endCellId).not.toBe(intersection);
    }
    // We expect four fragments total (row left, row right, col top, col bottom).
    expect(allFragments).toHaveLength(4);
  });
});

describe("GridViewport row and column header selection", () => {
  it("selects the whole row when clicking its header", async () => {
    const { wrapper, firstRow, firstColumn, lastColumn } = mountGrid();

    const rowHeaders = wrapper.findAll(".grid-row-header");
    await rowHeaders[0].trigger("mousedown");

    const firstCellId = createCellId(firstRow.id, firstColumn.id);
    const lastCellId = createCellId(firstRow.id, lastColumn.id);
    expect(wrapper.emitted("select")!.at(-1)?.[0]).toMatchObject({ id: firstCellId });
    expect(wrapper.emitted("select-range")!.at(-1)?.[0]).toEqual({
      startCellId: firstCellId,
      endCellId: lastCellId,
    });
    expect(wrapper.emitted("clear-additional-ranges")).toHaveLength(1);
  });

  it("extends the row selection from the anchor row when shift+clicking another row header", async () => {
    const { wrapper, firstRow, secondRow, firstColumn, lastColumn } = mountGrid();
    const firstRowFirstCellId = createCellId(firstRow.id, firstColumn.id);
    const firstRowLastCellId = createCellId(firstRow.id, lastColumn.id);
    await wrapper.setProps({
      selectedCellId: firstRowFirstCellId,
      selectedRange: { startCellId: firstRowFirstCellId, endCellId: firstRowLastCellId },
    });

    const rowHeaders = wrapper.findAll(".grid-row-header");
    await rowHeaders[1].trigger("mousedown", { shiftKey: true });

    expect(wrapper.emitted("select-range")!.at(-1)?.[0]).toEqual({
      startCellId: firstRowFirstCellId,
      endCellId: createCellId(secondRow.id, lastColumn.id),
    });
    expect(wrapper.emitted("add-range")).toBeUndefined();
  });

  it("appends the previous range to additional ranges when ctrl+clicking another row header", async () => {
    const { wrapper, firstRow, secondRow, firstColumn, lastColumn } = mountGrid();
    const firstRowFirstCellId = createCellId(firstRow.id, firstColumn.id);
    const firstRowLastCellId = createCellId(firstRow.id, lastColumn.id);
    await wrapper.setProps({
      selectedCellId: firstRowFirstCellId,
      selectedRange: { startCellId: firstRowFirstCellId, endCellId: firstRowLastCellId },
    });

    const rowHeaders = wrapper.findAll(".grid-row-header");
    await rowHeaders[1].trigger("mousedown", { ctrlKey: true });

    expect(wrapper.emitted("add-range")!.at(-1)?.[0]).toEqual({
      startCellId: firstRowFirstCellId,
      endCellId: firstRowLastCellId,
    });
    expect(wrapper.emitted("select-range")!.at(-1)?.[0]).toEqual({
      startCellId: createCellId(secondRow.id, firstColumn.id),
      endCellId: createCellId(secondRow.id, lastColumn.id),
    });
    expect(wrapper.emitted("clear-additional-ranges")).toBeUndefined();
  });

  it("extends the column selection from the anchor column when shift+clicking another column header", async () => {
    const { wrapper, firstRow, firstColumn, secondColumn, lastRow } = mountGrid();
    const firstColumnFirstCellId = createCellId(firstRow.id, firstColumn.id);
    const firstColumnLastCellId = createCellId(lastRow.id, firstColumn.id);
    await wrapper.setProps({
      selectedCellId: firstColumnFirstCellId,
      selectedRange: { startCellId: firstColumnFirstCellId, endCellId: firstColumnLastCellId },
    });

    const columnHeaders = wrapper.findAll(".grid-column-header");
    await columnHeaders[1].trigger("mousedown", { shiftKey: true });

    expect(wrapper.emitted("select-range")!.at(-1)?.[0]).toEqual({
      startCellId: firstColumnFirstCellId,
      endCellId: createCellId(lastRow.id, secondColumn.id),
    });
  });

  it("appends the previous range to additional ranges when ctrl+clicking another column header", async () => {
    const { wrapper, firstRow, firstColumn, secondColumn, lastRow } = mountGrid();
    const firstColumnFirstCellId = createCellId(firstRow.id, firstColumn.id);
    const firstColumnLastCellId = createCellId(lastRow.id, firstColumn.id);
    await wrapper.setProps({
      selectedCellId: firstColumnFirstCellId,
      selectedRange: { startCellId: firstColumnFirstCellId, endCellId: firstColumnLastCellId },
    });

    const columnHeaders = wrapper.findAll(".grid-column-header");
    await columnHeaders[1].trigger("mousedown", { ctrlKey: true });

    expect(wrapper.emitted("add-range")!.at(-1)?.[0]).toEqual({
      startCellId: firstColumnFirstCellId,
      endCellId: firstColumnLastCellId,
    });
    expect(wrapper.emitted("select-range")!.at(-1)?.[0]).toEqual({
      startCellId: createCellId(firstRow.id, secondColumn.id),
      endCellId: createCellId(lastRow.id, secondColumn.id),
    });
  });

  it("extends a cell-based selection to a whole row when shift+clicking a row header", async () => {
    const { wrapper, firstRow, secondRow, firstColumn, secondColumn, lastColumn } = mountGrid();
    const cellB1 = createCellId(firstRow.id, secondColumn.id);
    await wrapper.setProps({
      selectedCellId: cellB1,
      selectedRange: { startCellId: cellB1, endCellId: cellB1 },
    });

    const rowHeaders = wrapper.findAll(".grid-row-header");
    await rowHeaders[1].trigger("mousedown", { shiftKey: true });

    expect(wrapper.emitted("select-range")!.at(-1)?.[0]).toEqual({
      startCellId: createCellId(firstRow.id, firstColumn.id),
      endCellId: createCellId(secondRow.id, lastColumn.id),
    });
  });

  it("appends a cell selection to additional ranges when ctrl+clicking a row header afterwards", async () => {
    const { wrapper, secondRow, firstColumn, secondColumn, lastColumn } = mountGrid();
    const cellB2 = createCellId(secondRow.id, secondColumn.id);
    await wrapper.setProps({
      selectedCellId: cellB2,
      selectedRange: { startCellId: cellB2, endCellId: cellB2 },
    });

    const rowHeaders = wrapper.findAll(".grid-row-header");
    await rowHeaders[1].trigger("mousedown", { ctrlKey: true });

    expect(wrapper.emitted("add-range")!.at(-1)?.[0]).toEqual({
      startCellId: cellB2,
      endCellId: cellB2,
    });
    expect(wrapper.emitted("select-range")!.at(-1)?.[0]).toEqual({
      startCellId: createCellId(secondRow.id, firstColumn.id),
      endCellId: createCellId(secondRow.id, lastColumn.id),
    });
  });
});

describe("GridViewport editor navigation", () => {
  it("commits and moves the active cell down one row on Enter", async () => {
    const { wrapper, firstCellId, secondRowFirstCellId } = mountGrid();

    await wrapper.find(`[data-test-cell-id="${firstCellId}"]`).trigger("keydown", { key: "a" });
    const input = wrapper.find("textarea.grid-cell__editor");
    await input.setValue("entered");
    await input.trigger("keydown", { key: "Enter" });

    expect(wrapper.emitted("commit")).toHaveLength(1);
    expect(wrapper.emitted("commit")![0][1]).toBe("entered");
    expect(wrapper.emitted("select")!.at(-1)?.[0]).toMatchObject({ id: secondRowFirstCellId });
    expect(wrapper.emitted("select-range")!.at(-1)?.[0]).toEqual({
      startCellId: secondRowFirstCellId,
      endCellId: secondRowFirstCellId,
    });
  });

  it("commits and moves the active cell right one column on Tab", async () => {
    const { wrapper, firstCellId, secondCellId } = mountGrid();

    await wrapper.find(`[data-test-cell-id="${firstCellId}"]`).trigger("keydown", { key: "a" });
    const input = wrapper.find("textarea.grid-cell__editor");
    await input.setValue("tabbed");
    await input.trigger("keydown", { key: "Tab" });

    expect(wrapper.emitted("commit")).toHaveLength(1);
    expect(wrapper.emitted("commit")![0][1]).toBe("tabbed");
    expect(wrapper.emitted("select")!.at(-1)?.[0]).toMatchObject({ id: secondCellId });
    expect(wrapper.emitted("select-range")!.at(-1)?.[0]).toEqual({
      startCellId: secondCellId,
      endCellId: secondCellId,
    });
  });

  it("commits and moves the active cell up on Shift+Enter", async () => {
    const { wrapper, firstCellId, secondRowFirstCellId } = mountGrid();
    await wrapper.setProps({
      selectedCellId: secondRowFirstCellId,
      selectedRange: { startCellId: secondRowFirstCellId, endCellId: secondRowFirstCellId },
    });

    await wrapper.find(`[data-test-cell-id="${secondRowFirstCellId}"]`).trigger("keydown", { key: "a" });
    const input = wrapper.find("textarea.grid-cell__editor");
    await input.setValue("shift-enter");
    await input.trigger("keydown", { key: "Enter", shiftKey: true });

    expect(wrapper.emitted("commit")![0][1]).toBe("shift-enter");
    expect(wrapper.emitted("select")!.at(-1)?.[0]).toMatchObject({ id: firstCellId });
  });

  it("commits and moves the active cell left on Shift+Tab", async () => {
    const { wrapper, firstCellId, secondCellId } = mountGrid();
    await wrapper.setProps({
      selectedCellId: secondCellId,
      selectedRange: { startCellId: secondCellId, endCellId: secondCellId },
    });

    await wrapper.find(`[data-test-cell-id="${secondCellId}"]`).trigger("keydown", { key: "a" });
    const input = wrapper.find("textarea.grid-cell__editor");
    await input.setValue("shift-tab");
    await input.trigger("keydown", { key: "Tab", shiftKey: true });

    expect(wrapper.emitted("commit")![0][1]).toBe("shift-tab");
    expect(wrapper.emitted("select")!.at(-1)?.[0]).toMatchObject({ id: firstCellId });
  });

  it("inserts a newline at the caret on Alt+Enter without committing", async () => {
    const { wrapper, firstCellId } = mountGrid();

    await wrapper.find(`[data-test-cell-id="${firstCellId}"]`).trigger("keydown", { key: "a" });
    const input = wrapper.find("textarea.grid-cell__editor");
    await input.setValue("Line 1");

    const inputEl = input.element as HTMLTextAreaElement;
    inputEl.setSelectionRange(inputEl.value.length, inputEl.value.length);
    await input.trigger("keydown", { key: "Enter", altKey: true });

    expect(wrapper.emitted("commit")).toBeUndefined();
    expect(wrapper.find("textarea.grid-cell__editor").exists()).toBe(true);
    await input.setValue(`${(input.element as HTMLTextAreaElement).value}Line 2`);

    await input.trigger("keydown", { key: "Enter" });
    expect(wrapper.emitted("commit")).toHaveLength(1);
    expect(wrapper.emitted("commit")![0][1]).toBe("Line 1\nLine 2");
  });

  it("does not commit or block caret placement when clicking inside the active editor", async () => {
    const { wrapper, firstCellId } = mountGrid();

    await wrapper.find(`[data-test-cell-id="${firstCellId}"]`).trigger("keydown", { key: "a" });
    const input = wrapper.find("textarea.grid-cell__editor");
    await input.setValue("hello world");

    // Mousedowns originating from the textarea (target === textarea, not
    // the parent <td>) must not be intercepted by the cell's range
    // selection handler. Otherwise the browser's default caret /
    // text-selection behaviour on the textarea is preventDefault-ed and
    // the user cannot click inside the editor to reposition the caret.
    const event = new MouseEvent("mousedown", { bubbles: true, cancelable: true, button: 0 });
    (input.element as HTMLTextAreaElement).dispatchEvent(event);

    expect(event.defaultPrevented).toBe(false);
    expect(wrapper.emitted("commit")).toBeUndefined();
    expect(wrapper.find("textarea.grid-cell__editor").exists()).toBe(true);
  });
});

describe("GridViewport keyboard selection", () => {
  it("moves the selected cell with arrow keys", async () => {
    const { wrapper, firstCellId, secondCellId } = mountGrid();

    await wrapper.find(`[data-test-cell-id="${firstCellId}"]`).trigger("keydown", { key: "ArrowRight" });

    expect(wrapper.emitted("select")!.at(-1)?.[0]).toMatchObject({ id: secondCellId });
    expect(wrapper.emitted("select-range")!.at(-1)?.[0]).toEqual({
      startCellId: secondCellId,
      endCellId: secondCellId,
    });
  });

  it("keeps keyboard focus on the moved-to cell for repeated arrow navigation", async () => {
    const { wrapper, firstCellId, secondCellId, thirdCellId } = mountGrid();
    const firstCell = wrapper.find(`[data-test-cell-id="${firstCellId}"]`);

    (firstCell.element as HTMLElement).focus();
    await firstCell.trigger("keydown", { key: "ArrowRight" });
    await wrapper.setProps({
      selectedCellId: secondCellId,
      selectedRange: { startCellId: secondCellId, endCellId: secondCellId },
    });
    await nextTick();

    expect(document.activeElement).toBe(wrapper.find(`[data-test-cell-id="${secondCellId}"]`).element);

    await wrapper.find(`[data-test-cell-id="${secondCellId}"]`).trigger("keydown", { key: "ArrowRight" });

    expect(wrapper.emitted("select")!.at(-1)?.[0]).toMatchObject({ id: thirdCellId });
  });

  it("moves the selected cell right on Tab when not editing", async () => {
    const { wrapper, firstCellId, secondCellId } = mountGrid();

    await wrapper.find(`[data-test-cell-id="${firstCellId}"]`).trigger("keydown", { key: "Tab" });

    expect(wrapper.emitted("select")!.at(-1)?.[0]).toMatchObject({ id: secondCellId });
    expect(wrapper.emitted("select-range")!.at(-1)?.[0]).toEqual({
      startCellId: secondCellId,
      endCellId: secondCellId,
    });
  });

  it("moves the selected cell left on Shift+Tab when not editing", async () => {
    const { wrapper, firstCellId, secondCellId } = mountGrid();
    await wrapper.setProps({
      selectedCellId: secondCellId,
      selectedRange: { startCellId: secondCellId, endCellId: secondCellId },
    });

    await wrapper.find(`[data-test-cell-id="${secondCellId}"]`).trigger("keydown", { key: "Tab", shiftKey: true });

    expect(wrapper.emitted("select")!.at(-1)?.[0]).toMatchObject({ id: firstCellId });
  });

  it("extends the current range with Shift plus arrow keys", async () => {
    const { wrapper, firstCellId, secondCellId, secondRowSecondCellId } = mountGrid();
    await wrapper.setProps({
      selectedCellId: secondCellId,
      selectedRange: { startCellId: firstCellId, endCellId: secondCellId },
    });

    await wrapper.find(`[data-test-cell-id="${secondCellId}"]`).trigger("keydown", { key: "ArrowDown", shiftKey: true });

    expect(wrapper.emitted("select")!.at(-1)?.[0]).toMatchObject({ id: secondRowSecondCellId });
    expect(wrapper.emitted("select-range")!.at(-1)?.[0]).toEqual({
      startCellId: firstCellId,
      endCellId: secondRowSecondCellId,
    });
  });
});

describe("GridViewport resizing", () => {
  it("emits column and row resize events after pointer release", async () => {
    const { wrapper, firstColumn, firstRow } = mountGrid();

    wrapper.find(".grid-resize-handle--column").element.dispatchEvent(new MouseEvent("pointerdown", { bubbles: true, button: 0, clientX: 120 }));
    window.dispatchEvent(new MouseEvent("pointermove", { clientX: 185 }));
    window.dispatchEvent(new MouseEvent("pointerup"));
    await nextTick();

    expect(wrapper.emitted("resize-column")).toEqual([[{ columnId: firstColumn.id, width: 185 }]]);

    wrapper.find(".grid-resize-handle--row").element.dispatchEvent(new MouseEvent("pointerdown", { bubbles: true, button: 0, clientY: 32 }));
    window.dispatchEvent(new MouseEvent("pointermove", { clientY: 50 }));
    window.dispatchEvent(new MouseEvent("pointerup"));
    await nextTick();

    expect(wrapper.emitted("resize-row")).toEqual([[{ rowId: firstRow.id, height: 50 }]]);
  });

  it("does not render resize handles in read-only mode", () => {
    const { wrapper } = mountGrid({ readonly: true });

    expect(wrapper.find(".grid-resize-handle--column").exists()).toBe(false);
    expect(wrapper.find(".grid-resize-handle--row").exists()).toBe(false);
    expect(wrapper.emitted("resize-column")).toBeUndefined();
    expect(wrapper.emitted("resize-row")).toBeUndefined();
  });
});

describe("GridViewport text overflow", () => {
  it("lets text overflow across empty cells and stops before the next content cell", () => {
    const { wrapper, firstCellId } = mountGrid(({ worksheet, firstRow, firstColumn, thirdColumn }) => {
      worksheet.cellsById[createCellId(firstRow.id, firstColumn.id)] = {
        id: createCellId(firstRow.id, firstColumn.id),
        rowId: firstRow.id,
        columnId: firstColumn.id,
        value: { kind: "string", text: "A long value that should overflow" },
      };
      worksheet.cellsById[createCellId(firstRow.id, thirdColumn.id)] = {
        id: createCellId(firstRow.id, thirdColumn.id),
        rowId: firstRow.id,
        columnId: thirdColumn.id,
        value: { kind: "string", text: "stop" },
      };
    });

    expect(wrapper.find(`[data-test-cell-id="${firstCellId}"] .grid-cell__value`).attributes("style")).toContain("width: 240px");
  });

  it("keeps right-aligned numbers inside their own cell", () => {
    const { wrapper, firstCellId } = mountGrid(({ worksheet, firstRow, firstColumn }) => {
      worksheet.cellsById[createCellId(firstRow.id, firstColumn.id)] = {
        id: createCellId(firstRow.id, firstColumn.id),
        rowId: firstRow.id,
        columnId: firstColumn.id,
        value: { kind: "number", value: 30 },
      };
    });

    expect(wrapper.find(`[data-test-cell-id="${firstCellId}"] .grid-cell__value`).attributes("style")).toContain("width: 120px");
  });
});

function mountGrid(
  configureOrOptions?: MountGridOptions | ((context: MountGridConfigureContext) => void),
) {
  const document = createTeamGridDocument().teamgrid;
  const worksheet = document.workbook.worksheetsById[document.workbook.worksheetOrder[0]];
  const initialProjection = projectWorksheet(worksheet);
  const firstRow = initialProjection.rows[0];
  const secondRow = initialProjection.rows[1];
  const lastRow = initialProjection.rows[initialProjection.rows.length - 1];
  const firstColumn = initialProjection.columns[0];
  const secondColumn = initialProjection.columns[1];
  const thirdColumn = initialProjection.columns[2];
  const lastColumn = initialProjection.columns[initialProjection.columns.length - 1];
  const options = typeof configureOrOptions === "function" ? {} : configureOrOptions ?? {};
  const configure = typeof configureOrOptions === "function" ? configureOrOptions : configureOrOptions?.configure;

  configure?.({ worksheet, firstRow, secondRow, firstColumn, secondColumn, thirdColumn });

  const projection = projectWorksheet(worksheet);
  const firstCellId = createCellId(firstRow.id, firstColumn.id);
  const secondCellId = createCellId(firstRow.id, secondColumn.id);
  const thirdCellId = createCellId(firstRow.id, thirdColumn.id);
  const secondRowFirstCellId = createCellId(secondRow.id, firstColumn.id);
  const secondRowSecondCellId = createCellId(secondRow.id, secondColumn.id);
  const wrapper = mount(GridViewport, {
    props: {
      worksheet,
      projection,
      selectedCellId: firstCellId,
      selectedRange: { startCellId: firstCellId, endCellId: firstCellId },
      clipboardRange: null,
      highlightedCellIds: [],
      readonly: options.readonly ?? false,
      locale: "en-US",
    },
    attachTo: documentBody(),
  });

  wrapper.findAll("td.grid-cell").forEach((cellWrapper, index) => {
    const row = projection.rows[Math.floor(index / projection.columns.length)];
    const column = projection.columns[index % projection.columns.length];
    const cellId = row && column ? createCellId(row.id, column.id) : "";
    if (cellId) {
      cellWrapper.element.setAttribute("data-test-cell-id", cellId);
    }
  });

  return { wrapper, worksheet, firstRow, secondRow, lastRow, firstColumn, secondColumn, thirdColumn, lastColumn, firstCellId, secondCellId, thirdCellId, secondRowFirstCellId, secondRowSecondCellId };
}

interface MountGridOptions {
  readonly?: boolean;
  configure?: (context: MountGridConfigureContext) => void;
}

interface MountGridConfigureContext {
  worksheet: ReturnType<typeof createTeamGridDocument>["teamgrid"]["workbook"]["worksheetsById"][string];
  firstRow: ReturnType<typeof projectWorksheet>["rows"][number];
  secondRow: ReturnType<typeof projectWorksheet>["rows"][number];
  firstColumn: ReturnType<typeof projectWorksheet>["columns"][number];
  secondColumn: ReturnType<typeof projectWorksheet>["columns"][number];
  thirdColumn: ReturnType<typeof projectWorksheet>["columns"][number];
}

function documentBody() {
  const host = document.createElement("div");
  document.body.append(host);
  return host;
}
