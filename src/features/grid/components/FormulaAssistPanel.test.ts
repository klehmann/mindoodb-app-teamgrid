import { flushPromises, mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";

import FormulaAssistPanel from "@/features/grid/components/FormulaAssistPanel.vue";

function makeAnchor() {
  const host = document.createElement("input");
  document.body.append(host);
  return host;
}

function mountPanel(props: Partial<{ visible: boolean; draft: string; caretPos: number; anchorEl: HTMLElement | null; readonly: boolean }> = {}) {
  return mount(FormulaAssistPanel, {
    props: {
      visible: true,
      draft: "",
      caretPos: 0,
      anchorEl: null,
      readonly: false,
      ...props,
    },
    attachTo: document.body,
  });
}

describe("FormulaAssistPanel", () => {
  it("renders suggestions when visible with an anchor", async () => {
    const anchor = makeAnchor();
    const wrapper = mountPanel({ draft: "=SU", caretPos: 3, anchorEl: anchor });
    await flushPromises();

    const suggestions = document.body.querySelectorAll(".formula-assist__suggestion");
    expect(suggestions.length).toBeGreaterThan(0);
    expect(suggestions[0].textContent).toMatch(/SUM/i);
    expect(suggestions[0].querySelector("small")?.textContent).toBeTruthy();
    wrapper.unmount();
    anchor.remove();
  });

  it("emits select when a suggestion is clicked", async () => {
    const anchor = makeAnchor();
    const wrapper = mountPanel({ draft: "=SUM", caretPos: 4, anchorEl: anchor });
    await flushPromises();

    const first = document.body.querySelector<HTMLButtonElement>(".formula-assist__suggestion");
    first?.click();
    await flushPromises();

    const events = wrapper.emitted<[{ name: string }]>("select");
    expect(events?.[0]?.[0]?.name).toBeTruthy();
    wrapper.unmount();
    anchor.remove();
  });

  it("keeps editor focus stable before selecting a suggestion", async () => {
    const anchor = makeAnchor();
    const wrapper = mountPanel({ draft: "=SUM", caretPos: 4, anchorEl: anchor });
    anchor.focus();
    await flushPromises();

    const first = document.body.querySelector<HTMLButtonElement>(".formula-assist__suggestion");
    const pointerEvent = new PointerEvent("pointerdown", { bubbles: true, cancelable: true });
    first?.dispatchEvent(pointerEvent);
    first?.click();
    await flushPromises();

    expect(pointerEvent.defaultPrevented).toBe(true);
    expect(document.activeElement).toBe(anchor);
    expect(wrapper.emitted("select")).toHaveLength(1);
    wrapper.unmount();
    anchor.remove();
  });

  it("emits dismiss when the close button is clicked", async () => {
    const anchor = makeAnchor();
    const wrapper = mountPanel({ draft: "=SUM", caretPos: 4, anchorEl: anchor });
    await flushPromises();

    document.body.querySelector<HTMLButtonElement>(".formula-assist__close")?.click();
    await flushPromises();

    expect(wrapper.emitted("dismiss")).toHaveLength(1);
    expect(wrapper.emitted<[boolean]>("update:visible")?.at(-1)?.[0]).toBe(false);
    wrapper.unmount();
    anchor.remove();
  });

  it("hides itself when visible=false", () => {
    const wrapper = mountPanel({ visible: false });
    expect(document.body.querySelector(".formula-assist")).toBeNull();
    wrapper.unmount();
  });
});
