import { flushPromises, mount } from "@vue/test-utils";
import { describe, expect, it, vi } from "vitest";
import PrimeVue from "primevue/config";
import { ref } from "vue";

import RenameWorksheetDialog from "@/features/document/components/RenameWorksheetDialog.vue";
import type { useWorksheetDialogs } from "@/features/document/composables/useWorksheetDialogs";

function makeController() {
  const renameDialogVisible = ref(true);
  const renameDraft = ref("Sheet1");
  const applyWorksheetRename = vi.fn();
  return {
    renameDialogVisible,
    renameDraft,
    applyWorksheetRename,
  } as unknown as ReturnType<typeof useWorksheetDialogs>;
}

function findButtonByLabel(label: string) {
  return Array.from(document.body.querySelectorAll<HTMLButtonElement>("button"))
    .find((node) => node.textContent?.trim() === label || node.getAttribute("aria-label") === label);
}

describe("RenameWorksheetDialog", () => {
  it("renders the rename draft and applies on click", async () => {
    const controller = makeController();
    const wrapper = mount(RenameWorksheetDialog, {
      props: { controller },
      global: { plugins: [PrimeVue] },
      attachTo: document.body,
    });
    await flushPromises();

    const input = document.body.querySelector<HTMLInputElement>("input.native-input");
    expect(input?.value).toBe("Sheet1");

    findButtonByLabel("Rename")?.click();
    await flushPromises();

    expect((controller as unknown as { applyWorksheetRename: () => void }).applyWorksheetRename).toHaveBeenCalledTimes(1);
    wrapper.unmount();
  });

  it("disables the rename button when the draft is empty", async () => {
    const controller = makeController();
    (controller as unknown as { renameDraft: { value: string } }).renameDraft.value = "   ";
    mount(RenameWorksheetDialog, {
      props: { controller },
      global: { plugins: [PrimeVue] },
      attachTo: document.body,
    });
    await flushPromises();

    const rename = findButtonByLabel("Rename");
    expect(rename?.hasAttribute("disabled")).toBe(true);
  });
});
