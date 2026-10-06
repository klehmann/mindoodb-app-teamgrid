import { flushPromises, mount } from "@vue/test-utils";
import { defineComponent } from "vue";
import { beforeEach, describe, expect, it, vi } from "vitest";

import { useTeamGridDocument } from "@/features/document/composables/useTeamGridDocument";
import { createTeamGridDocument } from "@/features/document/lib/teamgridDocument";

vi.mock("@/shared/lib/theme", () => ({
  applyAppTheme: vi.fn(),
}));

const fakeDatabase = {
  documents: {
    create: vi.fn(),
    get: vi.fn(),
    list: vi.fn(),
    update: vi.fn(),
    delete: vi.fn(),
    history: vi.fn(),
    getRevision: vi.fn(),
  },
};

const fakeSession = {
  getLaunchContext: vi.fn(),
  onThemeChange: vi.fn(() => vi.fn()),
  onUiPreferencesChange: vi.fn(() => vi.fn()),
  onLocaleChange: vi.fn(() => vi.fn()),
  openDatabase: vi.fn(),
  createViewNavigator: vi.fn(),
  disconnect: vi.fn(),
  onBeforeClose: vi.fn(() => vi.fn()),
  embedding: {
    complete: vi.fn(),
    cancel: vi.fn(),
    setDirty: vi.fn(async () => {}),
    onSaveRequest: vi.fn((_handler: () => Promise<void>) => vi.fn()),
  },
};

vi.mock("mindoodb-app-sdk", () => ({
  abbreviateCanonicalName: (value: string) => value,
  createMindooDBAppBridge: vi.fn(() => ({
    connect: vi.fn(async () => fakeSession),
  })),
}));

describe("useTeamGridDocument open sessions", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    fakeSession.getLaunchContext.mockResolvedValue({
      databases: [{
        id: "db1",
        name: "Database",
        capabilities: ["read", "create", "update", "delete", "history"],
      }],
      preferredDatabaseId: "db1",
      runtime: "iframe",
      theme: {},
      uiPreferences: { iosMultitaskingOptimized: false, reduceMotion: false },
      locale: "en",
      user: { id: "u1", username: "cn=Test/o=Acme" },
    });
    fakeSession.openDatabase.mockResolvedValue(fakeDatabase);
    let nextDocumentId = 1;
    fakeDatabase.documents.create.mockImplementation(async ({ set }) => ({
      id: `doc-${nextDocumentId++}`,
      data: structuredClone(set),
      heads: [`head-${nextDocumentId}`],
    }));
  });

  it("creates imported documents as separate open sessions and switches between them", async () => {
    const wrapper = mountHarness();
    await flushPromises();
    const app = wrapper.vm.app;

    await app.createDocumentFromEnvelope(createTeamGridDocument("First import"));
    await app.createDocumentFromEnvelope(createTeamGridDocument("Second import"));

    expect(app.openSessions.value.map((session) => session.title)).toEqual(["First import", "Second import"]);
    expect(app.activeSubject.value).toBe("Second import");

    app.switchToOpenSession(app.openSessions.value[0].id);

    expect(app.activeSubject.value).toBe("First import");
    expect(app.openSessions.value.map((session) => session.isActive)).toEqual([true, false]);
    wrapper.unmount();
  });

  it("creates new spreadsheet documents with the teamgrid form marker", async () => {
    const wrapper = mountHarness();
    await flushPromises();
    const app = wrapper.vm.app;

    await app.createNewDocument();

    expect(fakeDatabase.documents.create).toHaveBeenCalledWith({
      set: expect.objectContaining({
        form: "teamgrid",
      }),
    });
    wrapper.unmount();
  });

  it("requires explicit discard confirmation before closing dirty sessions", async () => {
    const wrapper = mountHarness();
    await flushPromises();
    const app = wrapper.vm.app;

    await app.createDocumentFromEnvelope(createTeamGridDocument("Dirty spreadsheet"));
    app.updateGrid((_grid, envelope) => {
      envelope.subject = "Unsaved title";
      return [{ type: "setDocumentProperties", subject: "Unsaved title", tags: [], isTemplate: envelope.istemplate, locale: envelope.teamgrid.settings.locale }];
    });

    const sessionId = app.activeSpreadsheetSessionId.value;

    expect(app.closeOpenSession(sessionId)).toBe(false);
    expect(app.openSessions.value).toHaveLength(1);
    expect(app.status.value).toBe("Save the spreadsheet before closing its window.");

    expect(app.closeOpenSession(sessionId, { discardChanges: true })).toBe(true);
    expect(app.openSessions.value).toHaveLength(0);
    expect(app.currentDocument.value).toBeNull();
    wrapper.unmount();
  });

  it("creates a normal spreadsheet copy from a template", async () => {
    const wrapper = mountHarness();
    await flushPromises();
    const app = wrapper.vm.app;
    const template = createTeamGridDocument("Budget template", ["Finance"], "en-US", true);
    fakeDatabase.documents.get.mockResolvedValueOnce({
      id: "template-1",
      data: template,
      heads: ["template-head"],
    });

    await app.createDocumentFromTemplate("template-1");

    expect(fakeDatabase.documents.get).toHaveBeenCalledWith("template-1");
    expect(fakeDatabase.documents.create).toHaveBeenCalledWith({
      set: expect.objectContaining({
        subject: "Copy of Budget template",
        tags: ["Finance"],
        istemplate: false,
        form: "teamgrid",
      }),
    });
    expect(app.activeSubject.value).toBe("Copy of Budget template");
    wrapper.unmount();
  });

  it("surfaces save failures as user-visible errors", async () => {
    const wrapper = mountHarness();
    await flushPromises();
    const app = wrapper.vm.app;

    await app.createDocumentFromEnvelope(createTeamGridDocument("Save failure"));
    app.updateGrid((_grid, envelope) => {
      envelope.subject = "Unsaved title";
      return [{ type: "setDocumentProperties", subject: "Unsaved title", tags: [], isTemplate: envelope.istemplate, locale: envelope.teamgrid.settings.locale }];
    });
    fakeDatabase.documents.update.mockRejectedValueOnce(new Error("JSON patch failed"));

    await app.saveDocument();

    expect(app.lastErrorMessage.value).toBe("JSON patch failed");
    expect(app.status.value).toBe("JSON patch failed");
    expect(app.isDirty.value).toBe(true);
    wrapper.unmount();
  });
});

describe("useTeamGridDocument embedded as a component", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    fakeSession.getLaunchContext.mockResolvedValue({
      databases: [{ id: "embed", name: "CRM", capabilities: ["read", "update"] }],
      preferredDatabaseId: "embed",
      runtime: "iframe",
      theme: {},
      uiPreferences: { iosMultitaskingOptimized: false, reduceMotion: false },
      locale: "en",
      user: { id: "u1", username: "cn=Test/o=Acme" },
      embed: {
        embedId: "e1",
        componentId: "spreadsheet",
        intent: "edit",
        databaseId: "embed",
        docId: "root1",
        hostAppLabel: "CRM",
        features: { create: false, open: false },
      },
    });
    fakeSession.openDatabase.mockResolvedValue(fakeDatabase);
    // What a host creates from the component's `create` + `match` fields.
    fakeDatabase.documents.get.mockResolvedValue({
      id: "root1",
      data: { form: "teamgrid", kind: "mindoodb.teamgrid", subject: "", tags: [], istemplate: false, crm: { folders: ["Offers"] } },
      heads: ["h1"],
    });
    fakeDatabase.documents.update.mockImplementation(async (id: string, patch: { set: Record<string, unknown> }) => ({
      id,
      data: { form: "teamgrid", crm: { folders: ["Offers"] }, ...structuredClone(patch.set) },
      heads: ["h2"],
    }));
  });

  it("writes the workbook into the host's empty root and opens it", async () => {
    const wrapper = mountHarness();
    await flushPromises();
    const app = wrapper.vm.app;

    expect(app.embedded.value).toBe(true);
    expect(fakeDatabase.documents.update).toHaveBeenCalledTimes(1);
    const [docId, patch] = fakeDatabase.documents.update.mock.calls[0]!;
    expect(docId).toBe("root1");
    // Only the component's own fields: the host's match fields and namespace stay untouched.
    expect(Object.keys(patch.set).sort()).toEqual(["istemplate", "kind", "subject", "tags", "teamgrid"]);
    expect(app.activeGrid.value?.workbook.worksheetOrder).toHaveLength(1);
  });

  it("tells the host about unsaved edits and saves when the host asks", async () => {
    const wrapper = mountHarness();
    await flushPromises();
    const app = wrapper.vm.app;
    expect(fakeSession.embedding.setDirty).toHaveBeenLastCalledWith(false);
    const saveForHost = fakeSession.embedding.onSaveRequest.mock.calls[0]![0];

    app.updateGrid((_grid, envelope) => {
      envelope.subject = "Offer 2027";
      return [{ type: "setDocumentProperties", subject: "Offer 2027", tags: [], isTemplate: envelope.istemplate, locale: envelope.teamgrid.settings.locale }];
    });
    await flushPromises();
    expect(fakeSession.embedding.setDirty).toHaveBeenLastCalledWith(true);

    // A failed save keeps the edits, and the host learns why.
    fakeDatabase.documents.update.mockRejectedValueOnce(new Error("Disk full"));
    await expect(saveForHost()).rejects.toThrow("Disk full");
    expect(app.isDirty.value).toBe(true);

    await saveForHost();
    await flushPromises();
    expect(app.isDirty.value).toBe(false);
    expect(fakeSession.embedding.setDirty).toHaveBeenLastCalledWith(false);
    wrapper.unmount();
  });

  it("hands control back to the host when done", async () => {
    const wrapper = mountHarness();
    await flushPromises();
    await wrapper.vm.app.finishEmbedding();
    expect(fakeSession.embedding.complete).toHaveBeenCalledWith({ docId: "root1", subject: "Untitled spreadsheet" });
  });
});

function mountHarness() {
  return mount(defineComponent({
    setup(_, { expose }) {
      const app = useTeamGridDocument();
      expose({ app });
      return () => null;
    },
  })) as ReturnType<typeof mount> & { vm: { app: ReturnType<typeof useTeamGridDocument> } };
}
