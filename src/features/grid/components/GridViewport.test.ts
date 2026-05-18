import { mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";
import { nextTick } from "vue";

import GridViewport from "@/features/grid/components/GridViewport.vue";
import { projectWorksheet } from "@/features/grid/lib/gridProjection";
import { createCellId, createTeamGridDocument } from "@/features/document/lib/teamgridDocument";

describe("GridViewport inline editing", () => {
  it("commits a plain typed edit and selects the clicked cell", async () => {
    const { wrapper, firstCellId, secondCellId } = mountGrid();

    await wrapper.find(`[data-test-cell-id="${firstCellId}"]`).trigger("keydown", { key: "a" });
    expect(wrapper.find("input.grid-cell__editor").exists()).toBe(true);

    await wrapper.find(`[data-test-cell-id="${secondCellId}"]`).trigger("mousedown", { button: 0 });

    expect(wrapper.emitted("commit")).toHaveLength(1);
    expect(wrapper.emitted("commit")![0][1]).toBe("a");
    expect(wrapper.emitted("select")!.at(-1)?.[0]).toMatchObject({ id: secondCellId });
    expect(wrapper.find("input.grid-cell__editor").exists()).toBe(false);
  });

  it("keeps formula picking active when clicking another cell during formula edit", async () => {
    const { wrapper, firstCellId, secondCellId } = mountGrid();

    await wrapper.find(`[data-test-cell-id="${firstCellId}"]`).trigger("keydown", { key: "=" });
    await wrapper.find(`[data-test-cell-id="${secondCellId}"]`).trigger("mousedown", { button: 0 });

    expect(wrapper.emitted("commit")).toBeUndefined();
    expect((wrapper.find("input.grid-cell__editor").element as HTMLInputElement).value).toBe("=B1");
  });

  it("flushes an active inline edit for toolbar save actions", async () => {
    const { wrapper, firstCellId } = mountGrid();

    await wrapper.find(`[data-test-cell-id="${firstCellId}"]`).trigger("keydown", { key: "a" });
    await wrapper.find("input.grid-cell__editor").setValue("saved draft");

    expect((wrapper.vm as unknown as { flushPendingEdit: () => boolean }).flushPendingEdit()).toBe(true);

    expect(wrapper.emitted("commit")).toHaveLength(1);
    expect(wrapper.emitted("commit")![0][1]).toBe("saved draft");
    await nextTick();
    expect(wrapper.find("input.grid-cell__editor").exists()).toBe(false);
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
});

function mountGrid(
  configureOrOptions?: MountGridOptions | ((context: MountGridConfigureContext) => void),
) {
  const document = createTeamGridDocument().teamgrid;
  const worksheet = document.workbook.worksheetsById[document.workbook.worksheetOrder[0]];
  const initialProjection = projectWorksheet(worksheet);
  const firstRow = initialProjection.rows[0];
  const secondRow = initialProjection.rows[1];
  const firstColumn = initialProjection.columns[0];
  const secondColumn = initialProjection.columns[1];
  const thirdColumn = initialProjection.columns[2];
  const options = typeof configureOrOptions === "function" ? {} : configureOrOptions ?? {};
  const configure = typeof configureOrOptions === "function" ? configureOrOptions : configureOrOptions?.configure;

  configure?.({ worksheet, firstRow, secondRow, firstColumn, secondColumn, thirdColumn });

  const projection = projectWorksheet(worksheet);
  const firstCellId = createCellId(firstRow.id, firstColumn.id);
  const secondCellId = createCellId(firstRow.id, secondColumn.id);
  const thirdCellId = createCellId(firstRow.id, thirdColumn.id);
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

  return { wrapper, worksheet, firstRow, secondRow, firstColumn, secondColumn, thirdColumn, firstCellId, secondCellId, thirdCellId, secondRowSecondCellId };
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
