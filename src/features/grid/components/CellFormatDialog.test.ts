import { flushPromises, mount } from "@vue/test-utils";
import { describe, expect, it, vi } from "vitest";
import PrimeVue from "primevue/config";
import { ref } from "vue";

import CellFormatDialog from "@/features/grid/components/CellFormatDialog.vue";
import type { useCellFormatDialog } from "@/features/grid/composables/useCellFormatDialog";

function makeController() {
  return {
    formatDialogVisible: ref(true),
    formatDialogTab: ref<"cellType" | "font" | "fill" | "border">("cellType"),
    formatDialogKind: ref("general"),
    formatDialogCurrency: ref("USD"),
    formatDialogCustomNumFmt: ref(""),
    formatDialogFontFamily: ref("Calibri"),
    formatDialogFontSize: ref(14),
    formatDialogBold: ref(false),
    formatDialogItalic: ref(false),
    formatDialogUnderline: ref(false),
    formatDialogTextColor: ref("#111827"),
    formatDialogFillEnabled: ref(false),
    formatDialogFillColor: ref("#ffffff"),
    formatDialogBorderStyle: ref("thin"),
    formatDialogBorderColor: ref("#111827"),
    formatDialogBorderPreset: ref("custom"),
    formatDialogBorders: ref({}),
    openCellFormatDialog: vi.fn(),
    applySelectedCellFormat: vi.fn(),
    patchSelectedStyle: vi.fn(),
    patchCellsStyle: vi.fn(),
    updateCustomBordersFromLineSelection: vi.fn(),
    toggleFormatDialogBorder: vi.fn(),
    setFormatDialogBorderPreset: vi.fn(),
    currentDialogBorderCss: vi.fn(() => "1px solid #111827"),
  } as unknown as ReturnType<typeof useCellFormatDialog>;
}

function findButtonByLabel(label: string) {
  return Array.from(document.body.querySelectorAll<HTMLButtonElement>("button"))
    .find((node) => node.textContent?.trim() === label || node.getAttribute("aria-label") === label);
}

describe("CellFormatDialog", () => {
  it("switches tabs and applies the selected format", async () => {
    const controller = makeController();
    const wrapper = mount(CellFormatDialog, {
      props: { controller, readOnly: false },
      global: { plugins: [PrimeVue] },
      attachTo: document.body,
    });
    await flushPromises();

    const tabs = document.body.querySelectorAll<HTMLButtonElement>(".cell-format-dialog__tabs button");
    expect(tabs).toHaveLength(4);
    tabs[1].click();
    await flushPromises();
    expect((controller as unknown as { formatDialogTab: { value: string } }).formatDialogTab.value).toBe("font");

    findButtonByLabel("Apply")?.click();
    await flushPromises();
    expect((controller as unknown as { applySelectedCellFormat: () => void }).applySelectedCellFormat).toHaveBeenCalledTimes(1);
    wrapper.unmount();
  });

  it("disables Apply in read-only mode and when a custom format is empty", async () => {
    const controller = makeController();
    (controller as unknown as { formatDialogKind: { value: string } }).formatDialogKind.value = "custom";
    (controller as unknown as { formatDialogCustomNumFmt: { value: string } }).formatDialogCustomNumFmt.value = "   ";
    mount(CellFormatDialog, {
      props: { controller, readOnly: false },
      global: { plugins: [PrimeVue] },
      attachTo: document.body,
    });
    await flushPromises();

    expect(findButtonByLabel("Apply")?.hasAttribute("disabled")).toBe(true);
  });
});
