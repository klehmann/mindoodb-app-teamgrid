import { mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";
import { nextTick } from "vue";
import PrimeVue from "primevue/config";
import type { MenuItem } from "primevue/menuitem";

import WorksheetTabs from "@/features/grid/components/WorksheetTabs.vue";
import { createTeamGridDocument, type Worksheet } from "@/features/document/lib/teamgridDocument";

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

/**
 * Three tabs whose buttons report distinct, non-overlapping x ranges, so the
 * component can tell which one the pointer is over.
 */
function mountThreeTabs(readonly = false) {
  const grid = createTeamGridDocument("Test").teamgrid;
  const first = grid.workbook.worksheetOrder[0];
  for (const id of ["sheet_b", "sheet_c"]) {
    grid.workbook.worksheetOrder.push(id);
    grid.workbook.worksheetsById[id] = { ...grid.workbook.worksheetsById[first], id, title: id } as Worksheet;
  }
  const wrapper = mount(WorksheetTabs, {
    props: { grid, activeWorksheetId: first, readonly },
    global: { plugins: [PrimeVue] },
    attachTo: document.body,
  });
  // Measured from the slot the tab currently occupies, not the one it had at
  // mount, so the geometry keeps up once a drag has shuffled the strip.
  wrapper.findAll(".worksheet-tab").forEach((tab) => {
    tab.element.getBoundingClientRect = () => {
      const slots = Array.from(tab.element.parentElement?.querySelectorAll("[data-worksheet-id]") ?? []);
      const index = slots.indexOf(tab.element);
      return {
        left: index * 100,
        right: index * 100 + 100,
        top: 0,
        bottom: 30,
        width: 100,
        height: 30,
        x: index * 100,
        y: 0,
        toJSON: () => ({}),
      } as DOMRect;
    };
  });
  return { grid, wrapper, ids: [first, "sheet_b", "sheet_c"] };
}

/**
 * `trigger` cannot set `clientX` — jsdom exposes it as a getter — so pointer
 * events are built and dispatched directly.
 */
function pointer(
  element: Element,
  type: "pointerdown" | "pointermove" | "pointerup",
  options: { clientX: number; pointerType?: string },
) {
  const event = new MouseEvent(type, { clientX: options.clientX, button: 0, bubbles: true });
  Object.defineProperty(event, "pointerId", { value: 1 });
  Object.defineProperty(event, "pointerType", { value: options.pointerType ?? "mouse" });
  element.dispatchEvent(event);
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

  it("reports the drop index once, on release", () => {
    const { wrapper, ids } = mountThreeTabs();
    const tab = wrapper.findAll(".worksheet-tab")[0].element;

    pointer(tab, "pointerdown", { clientX: 50 });
    pointer(tab, "pointermove", { clientX: 250 });
    expect(wrapper.emitted("move")).toBeUndefined();

    pointer(tab, "pointerup", { clientX: 250 });

    expect(wrapper.emitted<[string, number]>("move")).toEqual([[ids[0], 2]]);
    wrapper.unmount();
  });

  it("shows the reorder while dragging without touching the document", async () => {
    const { wrapper, grid, ids } = mountThreeTabs();
    const tab = wrapper.findAll(".worksheet-tab")[0].element;

    pointer(tab, "pointerdown", { clientX: 50 });
    pointer(tab, "pointermove", { clientX: 250 });
    await nextTick();

    expect(wrapper.findAll(".worksheet-tab").map((item) => item.attributes("data-worksheet-id")))
      .toEqual([ids[1], ids[2], ids[0]]);
    expect(grid.workbook.worksheetOrder).toEqual(ids);
    wrapper.unmount();
  });

  /**
   * A drag over several tabs has to add up to one move — as a step per tab it
   * would depend on how the host orders a patch's deletes and inserts.
   */
  it("keeps a drag across several tabs a single move", async () => {
    const { wrapper, ids } = mountThreeTabs();
    const tab = wrapper.findAll(".worksheet-tab")[0].element;

    // Each move gets its own render, the way separate pointer events do in a
    // browser, so the tab the next one lands on is the one actually under it.
    pointer(tab, "pointerdown", { clientX: 50 });
    for (const clientX of [150, 250, 150]) {
      pointer(tab, "pointermove", { clientX });
      await nextTick();
    }
    pointer(tab, "pointerup", { clientX: 150 });

    expect(wrapper.emitted<[string, number]>("move")).toEqual([[ids[0], 1]]);
    wrapper.unmount();
  });

  it("treats a press below the threshold as a click, not a reorder", async () => {
    const { wrapper, ids } = mountThreeTabs();
    const tab = wrapper.findAll(".worksheet-tab")[0];

    pointer(tab.element, "pointerdown", { clientX: 50 });
    pointer(tab.element, "pointermove", { clientX: 53 });
    pointer(tab.element, "pointerup", { clientX: 53 });
    await tab.trigger("click");

    expect(wrapper.emitted("move")).toBeUndefined();
    expect(wrapper.emitted<[string]>("select")?.[0]?.[0]).toBe(ids[0]);
    wrapper.unmount();
  });

  it("keeps the dragged tab selected rather than the one now under the pointer", async () => {
    const { wrapper, ids } = mountThreeTabs();
    const tab = wrapper.findAll(".worksheet-tab")[0];

    pointer(tab.element, "pointerdown", { clientX: 50 });
    pointer(tab.element, "pointermove", { clientX: 250 });
    pointer(tab.element, "pointerup", { clientX: 250 });
    await tab.trigger("click");

    expect(wrapper.emitted<[string]>("select")).toEqual([[ids[0]]]);
    wrapper.unmount();
  });

  it("leaves touch drags to the strip's own scrolling", () => {
    const { wrapper } = mountThreeTabs();
    const tab = wrapper.findAll(".worksheet-tab")[0].element;

    pointer(tab, "pointerdown", { clientX: 50, pointerType: "touch" });
    pointer(tab, "pointermove", { clientX: 250, pointerType: "touch" });

    expect(wrapper.emitted("move")).toBeUndefined();
    wrapper.unmount();
  });

  it("does not reorder in read-only mode", () => {
    const { wrapper } = mountThreeTabs(true);
    const tab = wrapper.findAll(".worksheet-tab")[0].element;

    pointer(tab, "pointerdown", { clientX: 50 });
    pointer(tab, "pointermove", { clientX: 250 });

    expect(wrapper.emitted("move")).toBeUndefined();
    wrapper.unmount();
  });

  it("offers move entries in the context menu and greys out the ones at the edge", async () => {
    const { wrapper, ids } = mountThreeTabs();
    await wrapper.findAll(".worksheet-tab")[0].trigger("contextmenu");
    const items = wrapper.findAllComponents({ name: "ContextMenu" })[0].props("model") as MenuItem[];

    expect(items.find((item) => item.label === "Move Left")?.disabled).toBe(true);
    items.find((item) => item.label === "Move Right")?.command?.({} as never);

    expect(wrapper.emitted<[string, number]>("nudge")?.[0]).toEqual([ids[0], 1]);
    wrapper.unmount();
  });

  it("moves a tab with Alt and an arrow key", async () => {
    const { wrapper, ids } = mountThreeTabs();

    await wrapper.findAll(".worksheet-tab")[1].trigger("keydown", { key: "ArrowLeft", altKey: true });

    expect(wrapper.emitted<[string, number]>("nudge")?.[0]).toEqual([ids[1], -1]);
    wrapper.unmount();
  });

  it("leaves a bare arrow key to the browser's tab navigation", async () => {
    const { wrapper } = mountThreeTabs();

    await wrapper.findAll(".worksheet-tab")[1].trigger("keydown", { key: "ArrowLeft" });

    expect(wrapper.emitted("nudge")).toBeUndefined();
    wrapper.unmount();
  });
});
