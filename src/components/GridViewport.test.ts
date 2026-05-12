import { mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";

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

function mountGrid() {
  const document = createTeamGridDocument().teamgrid;
  const worksheet = document.workbook.worksheetsById[document.workbook.worksheetOrder[0]];
  const projection = projectWorksheet(worksheet);
  const firstRow = projection.rows[0];
  const firstColumn = projection.columns[0];
  const secondColumn = projection.columns[1];
  const firstCellId = createCellId(firstRow.id, firstColumn.id);
  const secondCellId = createCellId(firstRow.id, secondColumn.id);
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
    const cellId = index === 0 ? firstCellId : index === 1 ? secondCellId : "";
    if (cellId) {
      cellWrapper.element.setAttribute("data-test-cell-id", cellId);
    }
  });

  return { wrapper, firstCellId, secondCellId };
}

function documentBody() {
  const host = document.createElement("div");
  document.body.append(host);
  return host;
}
