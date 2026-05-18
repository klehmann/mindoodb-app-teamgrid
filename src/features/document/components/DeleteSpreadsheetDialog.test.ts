import { flushPromises, mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";
import PrimeVue from "primevue/config";

import DeleteSpreadsheetDialog from "@/features/document/components/DeleteSpreadsheetDialog.vue";

function mountDialog(props: { visible: boolean }) {
  return mount(DeleteSpreadsheetDialog, {
    props,
    global: { plugins: [PrimeVue] },
    attachTo: document.body,
  });
}

function findButtonByLabel(label: string): HTMLButtonElement | undefined {
  return Array.from(document.body.querySelectorAll<HTMLButtonElement>("button"))
    .find((node) => node.textContent?.trim() === label || node.getAttribute("aria-label") === label);
}

describe("DeleteSpreadsheetDialog", () => {
  it("emits confirm and clears visible when the delete button is clicked", async () => {
    const wrapper = mountDialog({ visible: true });
    await flushPromises();

    const deleteButton = findButtonByLabel("Delete");
    expect(deleteButton).toBeDefined();
    deleteButton?.click();

    expect(wrapper.emitted("confirm")).toHaveLength(1);
    const visibleEvents = wrapper.emitted<[boolean]>("update:visible");
    expect(visibleEvents?.some((payload) => payload[0] === false)).toBe(true);
    wrapper.unmount();
  });

  it("dismisses without confirming when the cancel button is clicked", async () => {
    const wrapper = mountDialog({ visible: true });
    await flushPromises();

    const cancelButton = findButtonByLabel("Cancel");
    expect(cancelButton).toBeDefined();
    cancelButton?.click();

    expect(wrapper.emitted("confirm")).toBeUndefined();
    const visibleEvents = wrapper.emitted<[boolean]>("update:visible");
    expect(visibleEvents?.some((payload) => payload[0] === false)).toBe(true);
    wrapper.unmount();
  });
});
