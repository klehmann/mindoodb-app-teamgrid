import { flushPromises, mount } from "@vue/test-utils";
import { describe, expect, it, vi } from "vitest";
import PrimeVue from "primevue/config";
import { ref } from "vue";

import DocumentPropertiesDialog from "@/features/document/components/DocumentPropertiesDialog.vue";
import type { useDocumentPropertiesDialog } from "@/features/document/composables/useDocumentPropertiesDialog";

function makeController() {
  const propertiesDialogVisible = ref(true);
  const propertiesTitleDraft = ref("Q3 plan");
  const propertiesTagsDraft = ref("Work\\Planning\nFinance");
  const propertiesIsTemplateDraft = ref(true);
  const propertiesLocaleDraft = ref("de-DE");
  const propertiesLocaleOptions = ref([
    { value: "en-US", label: "English (United States)" },
    { value: "de-DE", label: "German (Germany)" },
  ]);
  const applyDocumentProperties = vi.fn();
  const resetPropertiesDraft = vi.fn();
  return {
    propertiesDialogVisible,
    propertiesTitleDraft,
    propertiesTagsDraft,
    propertiesIsTemplateDraft,
    propertiesLocaleDraft,
    propertiesLocaleOptions,
    applyDocumentProperties,
    resetPropertiesDraft,
  } as unknown as ReturnType<typeof useDocumentPropertiesDialog>;
}

function findButtonByLabel(label: string) {
  return Array.from(document.body.querySelectorAll<HTMLButtonElement>("button"))
    .find((node) => node.textContent?.trim() === label || node.getAttribute("aria-label") === label);
}

describe("DocumentPropertiesDialog", () => {
  it("applies title and tag drafts when Apply is clicked", async () => {
    const controller = makeController();
    const wrapper = mount(DocumentPropertiesDialog, {
      props: { controller, readOnly: false },
      global: { plugins: [PrimeVue] },
      attachTo: document.body,
    });
    await flushPromises();

    expect(document.body.querySelector<HTMLInputElement>("input.native-input")?.value).toBe("Q3 plan");
    expect(document.body.querySelector<HTMLInputElement>('input[type="checkbox"]')?.checked).toBe(true);
    expect(document.body.querySelector<HTMLSelectElement>("select.native-input")?.value).toBe("de-DE");

    findButtonByLabel("Apply")?.click();
    await flushPromises();

    expect((controller as unknown as { applyDocumentProperties: () => void }).applyDocumentProperties).toHaveBeenCalledTimes(1);
    wrapper.unmount();
  });

  it("disables Apply and clears the draft on Cancel when readOnly is true", async () => {
    const controller = makeController();
    mount(DocumentPropertiesDialog, {
      props: { controller, readOnly: true },
      global: { plugins: [PrimeVue] },
      attachTo: document.body,
    });
    await flushPromises();

    expect(findButtonByLabel("Apply")?.hasAttribute("disabled")).toBe(true);

    findButtonByLabel("Cancel")?.click();
    await flushPromises();

    expect((controller as unknown as { resetPropertiesDraft: () => void }).resetPropertiesDraft).toHaveBeenCalledTimes(1);
    expect((controller as unknown as { propertiesDialogVisible: { value: boolean } }).propertiesDialogVisible.value).toBe(false);
  });
});
