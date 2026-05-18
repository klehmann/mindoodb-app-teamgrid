import { flushPromises, mount } from "@vue/test-utils";
import { describe, expect, it } from "vitest";
import PrimeVue from "primevue/config";
import type { MindooDBAppDocumentHistoryEntry } from "mindoodb-app-sdk";

import DocumentRevisionDialog from "@/features/document/components/DocumentRevisionDialog.vue";

function makeEntry(overrides: Partial<MindooDBAppDocumentHistoryEntry> = {}): MindooDBAppDocumentHistoryEntry {
  return {
    revisionId: "rev-1",
    timestamp: Date.UTC(2025, 5, 10, 9, 0, 0),
    publicKey: "key-1",
    identityLabel: "Alice",
    publicKeyFingerprint: "fp-1",
    isCurrent: false,
    isDeleted: false,
    summary: "Snapshot",
    ...overrides,
  } as MindooDBAppDocumentHistoryEntry;
}

function findButtonByLabel(label: string) {
  return Array.from(document.body.querySelectorAll<HTMLButtonElement>("button"))
    .find((node) => node.textContent?.trim() === label || node.getAttribute("aria-label") === label);
}

describe("DocumentRevisionDialog", () => {
  it("shows a loading state while history is fetching", async () => {
    const wrapper = mount(DocumentRevisionDialog, {
      props: {
        visible: true,
        entries: [],
        loading: true,
        errorMessage: null,
        currentRevisionId: null,
      },
      global: { plugins: [PrimeVue] },
      attachTo: document.body,
    });
    await flushPromises();

    expect(document.body.textContent).toContain("Loading revisions...");
    wrapper.unmount();
  });

  it("emits select with the chosen revision when Open revision is clicked", async () => {
    const entries = [
      makeEntry({ revisionId: "rev-1", isCurrent: true }),
      makeEntry({ revisionId: "rev-2", timestamp: Date.UTC(2025, 5, 9), summary: "Edit" }),
    ];
    const wrapper = mount(DocumentRevisionDialog, {
      props: {
        visible: true,
        entries,
        loading: false,
        errorMessage: null,
        currentRevisionId: "rev-1",
      },
      global: { plugins: [PrimeVue] },
      attachTo: document.body,
    });
    await flushPromises();

    const rows = document.body.querySelectorAll<HTMLButtonElement>(".revision-row");
    expect(rows).toHaveLength(2);
    rows[1].click();
    await flushPromises();

    findButtonByLabel("Open revision")?.click();
    await flushPromises();

    expect(wrapper.emitted<[string]>("select")?.[0]?.[0]).toBe("rev-2");
    wrapper.unmount();
  });

  it("emits cancel when dismissed via the cancel button", async () => {
    const wrapper = mount(DocumentRevisionDialog, {
      props: {
        visible: true,
        entries: [makeEntry()],
        loading: false,
        errorMessage: null,
        currentRevisionId: null,
      },
      global: { plugins: [PrimeVue] },
      attachTo: document.body,
    });
    await flushPromises();

    findButtonByLabel("Cancel")?.click();
    await flushPromises();

    expect(wrapper.emitted("cancel")).toHaveLength(1);
    expect(wrapper.emitted<[boolean]>("update:visible")?.at(-1)?.[0]).toBe(false);
    wrapper.unmount();
  });
});
