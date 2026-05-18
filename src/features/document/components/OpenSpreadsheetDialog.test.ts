import { flushPromises, mount } from "@vue/test-utils";
import { afterEach, describe, expect, it, vi } from "vitest";
import PrimeVue from "primevue/config";
import { ref } from "vue";
import type { MindooDBAppDatabaseInfo } from "mindoodb-app-sdk";

import OpenSpreadsheetDialog from "@/features/document/components/OpenSpreadsheetDialog.vue";
import type { useOpenDialog } from "@/features/document/composables/useOpenDialog";
import type { OpenCategoryNode, OpenDocumentRow } from "@/features/document/lib/viewOpen";

afterEach(() => {
  document.body.innerHTML = "";
});

function makeController(documents: OpenDocumentRow[]) {
  return {
    openDialogVisible: ref(true),
    selectedOpenDocId: ref(""),
    selectedOpenCategoryKey: ref("__all__"),
    openCategoryNodes: ref<OpenCategoryNode[]>([
      { key: "__all__", label: "All", count: documents.length, children: [] } as OpenCategoryNode,
    ]),
    openDialogDocuments: ref(documents),
    handleOpenDatabaseChange: vi.fn(),
    selectOpenCategory: vi.fn(),
    openSelectedDocument: vi.fn(),
    disposeOpenNavigator: vi.fn(),
  } as unknown as ReturnType<typeof useOpenDialog>;
}

function findButtonByLabel(label: string) {
  return Array.from(document.body.querySelectorAll<HTMLButtonElement>("button"))
    .find((node) => node.textContent?.trim() === label || node.getAttribute("aria-label") === label);
}

describe("OpenSpreadsheetDialog", () => {
  it("lists documents and opens the selected one", async () => {
    const controller = makeController([
      { id: "doc-1", key: "doc-1", title: "Q1 plan", detail: "Updated yesterday", tags: [] } as OpenDocumentRow,
      { id: "doc-2", key: "doc-2", title: "Roadmap", detail: "Updated last week", tags: [] } as OpenDocumentRow,
    ]);
    const databases: MindooDBAppDatabaseInfo[] = [
      { id: "db-a", title: "Primary", capabilities: ["read", "views"] } as MindooDBAppDatabaseInfo,
    ];
    const wrapper = mount(OpenSpreadsheetDialog, {
      props: {
        controller,
        selectedDatabaseId: ref("db-a"),
        readableDatabases: ref(databases),
      },
      global: { plugins: [PrimeVue] },
      attachTo: document.body,
    });
    await flushPromises();

    const rows = document.body.querySelectorAll<HTMLButtonElement>(".document-row");
    expect(rows).toHaveLength(2);
    rows[1].click();
    await flushPromises();

    const openButton = findButtonByLabel("Open");
    expect(openButton?.hasAttribute("disabled")).toBe(false);
    openButton?.click();
    await flushPromises();

    expect((controller as unknown as { openSelectedDocument: () => void }).openSelectedDocument).toHaveBeenCalledTimes(1);
    expect((controller as unknown as { selectedOpenDocId: { value: string } }).selectedOpenDocId.value).toBe("doc-2");
    wrapper.unmount();
  });

  it("renders the empty-state placeholder when there are no documents", async () => {
    const controller = makeController([]);
    mount(OpenSpreadsheetDialog, {
      props: {
        controller,
        selectedDatabaseId: ref("db-a"),
        readableDatabases: ref([]),
      },
      global: { plugins: [PrimeVue] },
      attachTo: document.body,
    });
    await flushPromises();

    expect(document.body.textContent).toContain("No spreadsheets in this category.");
    expect(findButtonByLabel("Open")?.hasAttribute("disabled")).toBe(true);
  });
});
