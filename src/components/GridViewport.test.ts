import { mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";
import { nextTick } from "vue";

import GridViewport from "@/components/GridViewport.vue";
import { projectWorksheet } from "@/lib/gridProjection";
import { createCellId, createTeamGridDocument } from "@/lib/teamgridDocument";

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

function mountGrid() {
  const document = createTeamGridDocument().teamgrid;
  const worksheet = document.workbook.worksheetsById[document.workbook.worksheetOrder[0]];
  const projection = projectWorksheet(worksheet);
  const firstRow = projection.rows[0];
  const secondRow = projection.rows[1];
  const firstColumn = projection.columns[0];
  const secondColumn = projection.columns[1];
  const thirdColumn = projection.columns[2];
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
      readonly: false,
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

  return { wrapper, firstCellId, secondCellId, thirdCellId, secondRowSecondCellId };
}

function documentBody() {
  const host = document.createElement("div");
  document.body.append(host);
  return host;
}
