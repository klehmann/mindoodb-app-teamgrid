import { flushPromises, mount } from "@vue/test-utils";
import { describe, expect, it, vi } from "vitest";
import PrimeVue from "primevue/config";
import { ref } from "vue";

import ErrorDialog from "@/features/document/components/ErrorDialog.vue";
import type { useErrorDialog } from "@/features/document/composables/useErrorDialog";

function makeController() {
  const errorDialogVisible = ref(true);
  const dismissErrorDialog = vi.fn(() => {
    errorDialogVisible.value = false;
  });
  return { errorDialogVisible, dismissErrorDialog } as unknown as ReturnType<typeof useErrorDialog>;
}

function findButtonByLabel(label: string) {
  return Array.from(document.body.querySelectorAll<HTMLButtonElement>("button"))
    .find((node) => node.textContent?.trim() === label || node.getAttribute("aria-label") === label);
}

describe("ErrorDialog", () => {
  it("shows the provided message and dismisses through the controller", async () => {
    const controller = makeController();
    const onHide = vi.fn();
    const wrapper = mount(ErrorDialog, {
      props: { controller, message: "Save failed: invalid revision", onHide },
      global: { plugins: [PrimeVue] },
      attachTo: document.body,
    });
    await flushPromises();

    expect(document.body.textContent).toContain("Save failed: invalid revision");

    findButtonByLabel("OK")?.click();
    await flushPromises();

    expect((controller as unknown as { dismissErrorDialog: () => void }).dismissErrorDialog).toHaveBeenCalledTimes(1);
    wrapper.unmount();
  });
});
