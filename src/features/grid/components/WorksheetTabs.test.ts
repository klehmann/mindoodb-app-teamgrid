import { mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";
import PrimeVue from "primevue/config";
import type { MenuItem } from "primevue/menuitem";

import WorksheetTabs from "@/features/grid/components/WorksheetTabs.vue";
import { createTeamGridDocument } from "@/features/document/lib/teamgridDocument";

function mountTabs(readonly = false) {
  const grid = createTeamGridDocument("Test").teamgrid;
  return {
    grid,
    wrapper: mount(WorksheetTabs, {
      props: {
        grid,
        activeWorksheetId: grid.workbook.worksheetOrder[0],
        readonly,
      },
      global: { plugins: [PrimeVue] },
      attachTo: document.body,
    }),
  };
}

describe("WorksheetTabs", () => {
  it("renders a tab per live worksheet and an add button", () => {
    const { wrapper, grid } = mountTabs();
    const tabs = wrapper.findAll(".worksheet-tab");
    expect(tabs).toHaveLength(grid.workbook.worksheetOrder.length);
    expect(wrapper.find('[aria-label="Add worksheet"]').exists()).toBe(true);
    wrapper.unmount();
  });

  it("emits select when a tab is clicked", async () => {
    const { wrapper, grid } = mountTabs();
    const tabs = wrapper.findAll(".worksheet-tab");
    await tabs[0].trigger("click");

    expect(wrapper.emitted<[string]>("select")?.[0]?.[0]).toBe(grid.workbook.worksheetOrder[0]);
    wrapper.unmount();
  });

  it("offers Add Sheet from the plus menu", () => {
    const { wrapper } = mountTabs();
    const contextMenus = wrapper.findAllComponents({ name: "ContextMenu" });
    const addMenuItems = contextMenus[1].props("model") as MenuItem[];

    addMenuItems.find((item) => item.label === "Add Sheet")?.command?.({} as never);

    expect(wrapper.emitted("add")).toHaveLength(1);
    wrapper.unmount();
  });

  it("offers Add Virtual View Sheet from the plus menu", () => {
    const { wrapper } = mountTabs();
    const contextMenus = wrapper.findAllComponents({ name: "ContextMenu" });
    const addMenuItems = contextMenus[1].props("model") as MenuItem[];

    addMenuItems.find((item) => item.label === "Add Virtual View Sheet")?.command?.({} as never);

    expect(wrapper.emitted("add-view")).toHaveLength(1);
    wrapper.unmount();
  });

  it("emits rename on double-click", async () => {
    const { wrapper, grid } = mountTabs();
    await wrapper.findAll(".worksheet-tab")[0].trigger("dblclick");
    expect(wrapper.emitted<[string]>("rename")?.[0]?.[0]).toBe(grid.workbook.worksheetOrder[0]);
    wrapper.unmount();
  });

  it("disables interactions in read-only mode", async () => {
    const { wrapper } = mountTabs(true);
    expect(wrapper.find<HTMLButtonElement>('[aria-label="Add worksheet"]').attributes("disabled")).toBeDefined();

    await wrapper.findAll(".worksheet-tab")[0].trigger("dblclick");
    expect(wrapper.emitted("rename")).toBeUndefined();
    wrapper.unmount();
  });

  it("offers settings for view-backed worksheets", async () => {
    const { wrapper, grid } = mountTabs();
    const worksheetId = grid.workbook.worksheetOrder[0];
    grid.workbook.worksheetsById[worksheetId].viewBinding = {
      kind: "mindoodbView",
      viewId: "contacts_flat",
      viewTitle: "Contacts",
      showDocuments: true,
      showCategories: true,
      rootCategoryPath: [],
    };
    wrapper.vm.$forceUpdate();
    await wrapper.find(".worksheet-tab").trigger("contextmenu");
    const contextMenus = wrapper.findAllComponents({ name: "ContextMenu" });
    const tabMenuItems = contextMenus[0].props("model") as MenuItem[];
    tabMenuItems.find((item) => item.label === "View Sheet Settings...")?.command?.({} as never);

    expect(wrapper.emitted<[string]>("configure-view")?.[0]?.[0]).toBe(worksheetId);
    wrapper.unmount();
  });
});
