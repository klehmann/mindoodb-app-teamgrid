import { mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";

import TagTreeList from "@/features/document/components/TagTreeList.vue";
import type { OpenCategoryNode } from "@/features/document/lib/viewOpen";

function buildNodes(): OpenCategoryNode[] {
  return [
    {
      key: "all",
      label: "All spreadsheets",
      count: 5,
      children: [],
    } as OpenCategoryNode,
    {
      key: "Work",
      label: "Work",
      count: 3,
      children: [
        {
          key: "Work\\Planning",
          label: "Planning",
          count: 2,
          children: [],
        } as OpenCategoryNode,
      ],
    } as OpenCategoryNode,
  ];
}

describe("TagTreeList", () => {
  it("renders top-level and nested nodes with their counts", () => {
    const wrapper = mount(TagTreeList, {
      props: { nodes: buildNodes(), selectedKey: "all" },
    });

    const buttons = wrapper.findAll(".tag-tree-list__button");
    expect(buttons).toHaveLength(3);
    expect(buttons[0].text()).toContain("All spreadsheets");
    expect(buttons[0].text()).toContain("5");
    expect(buttons[2].text()).toContain("Planning");
  });

  it("emits select with the clicked node key (including for nested nodes)", async () => {
    const wrapper = mount(TagTreeList, {
      props: { nodes: buildNodes(), selectedKey: "all" },
    });

    const buttons = wrapper.findAll(".tag-tree-list__button");
    await buttons[1].trigger("click");
    await buttons[2].trigger("click");

    const events = wrapper.emitted<[string]>("select")!;
    expect(events.map((payload) => payload[0])).toEqual(["Work", "Work\\Planning"]);
  });

  it("marks the active node with the selected modifier class", () => {
    const wrapper = mount(TagTreeList, {
      props: { nodes: buildNodes(), selectedKey: "Work" },
    });

    const selected = wrapper.find(".tag-tree-list__button--selected");
    expect(selected.exists()).toBe(true);
    expect(selected.text()).toContain("Work");
  });
});
