import { flushPromises, mount } from "@vue/test-utils";
import { defineComponent } from "vue";
import { beforeEach, describe, expect, it, vi } from "vitest";

import { useTeamGridDocument } from "@/composables/useTeamGridDocument";
import { createTeamGridDocument } from "@/lib/teamgridDocument";

vi.mock("@/lib/theme", () => ({
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
  openDatabase: vi.fn(),
  createViewNavigator: vi.fn(),
  disconnect: vi.fn(),
};

vi.mock("mindoodb-app-sdk", () => ({
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

  it("surfaces save failures as user-visible errors", async () => {
    const wrapper = mountHarness();
    await flushPromises();
    const app = wrapper.vm.app;

    await app.createDocumentFromEnvelope(createTeamGridDocument("Save failure"));
    app.updateGrid((_grid, envelope) => {
      envelope.subject = "Unsaved title";
      return [{ type: "setDocumentProperties", subject: "Unsaved title", tags: [] }];
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
