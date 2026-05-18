import { mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";
import PrimeVue from "primevue/config";

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

  it("emits add when the plus button is clicked", async () => {
    const { wrapper } = mountTabs();
    await wrapper.find('[aria-label="Add worksheet"]').trigger("click");
    expect(wrapper.emitted("add")).toHaveLength(1);
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
});
