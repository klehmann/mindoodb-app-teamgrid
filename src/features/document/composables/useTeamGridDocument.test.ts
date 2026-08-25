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
      uiPreferences: { iosMultitaskingOptimized: false },
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

function mountHarness() {
  return mount(defineComponent({
    setup(_, { expose }) {
      const app = useTeamGridDocument();
      expose({ app });
      return () => null;
    },
  })) as ReturnType<typeof mount> & { vm: { app: ReturnType<typeof useTeamGridDocument> } };
}
