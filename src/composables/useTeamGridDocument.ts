import { computed, onBeforeUnmount, onMounted, ref } from "vue";
import {
  createMindooDBAppBridge,
  type MindooDBAppDatabase,
  type MindooDBAppDatabaseInfo,
  type MindooDBAppDocument,
  type MindooDBAppDocumentHistoryEntry,
  type MindooDBAppDocumentRevisionId,
  type MindooDBAppDocumentSummary,
  type MindooDBAppHistoricalDocument,
  type MindooDBAppRuntime,
  type MindooDBAppSession,
  type MindooDBAppUiPreferences,
} from "mindoodb-app-sdk";

import { applyAppTheme } from "@/lib/theme";
import { canMutateGrid, isGridSessionReadOnly } from "@/lib/capabilities";
import {
  hasTeamGridOperations,
  serializeTeamGridOperations,
  type TeamGridOperation,
} from "@/lib/teamgridOps";
import {
  createTeamGridDocument,
  cloneTeamGridDocument,
  migrateTeamGridDocument,
  readSubject,
  type TeamGridDocumentEnvelope,
  type TeamGridDocumentV1,
} from "@/lib/teamgridDocument";

export function useTeamGridDocument() {
  const session = ref<MindooDBAppSession | null>(null);
  const databases = ref<MindooDBAppDatabaseInfo[]>([]);
  const selectedDatabaseId = ref("");
  const currentDatabase = ref<MindooDBAppDatabase | null>(null);
  const currentDatabaseId = ref("");
  const documents = ref<MindooDBAppDocumentSummary[]>([]);
  const currentDocument = ref<MindooDBAppDocument | null>(null);
  const currentEnvelope = ref<TeamGridDocumentEnvelope | null>(null);
  const viewingHistoricalSnapshot = ref<MindooDBAppHistoricalDocument | null>(null);
  const currentRuntime = ref<MindooDBAppRuntime>("iframe");
  const hostUiPreferences = ref<MindooDBAppUiPreferences>({ iosMultitaskingOptimized: false });
  const launchTimeTravelDate = ref<number | null>(null);
  const status = ref("Connecting to Haven...");
  const isDirty = ref(false);
  const revisionEntries = ref<MindooDBAppDocumentHistoryEntry[]>([]);
  const revisionLoading = ref(false);
  const revisionErrorMessage = ref<string | null>(null);
  const pendingOps = ref<TeamGridOperation[]>([]);
  const pendingOpsBaseHeads = ref<string[]>([]);

  let cleanupTheme: (() => void) | null = null;
  let cleanupUiPreferences: (() => void) | null = null;

  const readableDatabases = computed(() => databases.value.filter((database) => database.capabilities.includes("read")));
  const creatableDatabases = computed(() => databases.value.filter((database) => database.capabilities.includes("create")));
  const selectedDatabaseInfo = computed(() => databases.value.find((database) => database.id === selectedDatabaseId.value) ?? null);
  const currentDatabaseInfo = computed(() => databases.value.find((database) => database.id === currentDatabaseId.value) ?? null);
  const currentCanUpdate = computed(() => currentDatabaseInfo.value?.capabilities.includes("update") ?? false);
  const currentCanDelete = computed(() => currentDatabaseInfo.value?.capabilities.includes("delete") ?? false);
  const currentCanBrowseHistory = computed(() => currentDatabaseInfo.value?.capabilities.includes("history") ?? false);
  const isTimeTravelActive = computed(() => launchTimeTravelDate.value != null);
  const isViewingHistorical = computed(() => viewingHistoricalSnapshot.value !== null);
  const gridReadOnly = computed(() => isGridSessionReadOnly({
    timeTravelDate: launchTimeTravelDate.value,
    viewingHistorical: isViewingHistorical.value,
  }));
  const activeGrid = computed(() => currentEnvelope.value?.teamgrid ?? null);
  const activeSubject = computed(() => currentEnvelope.value?.subject ?? "");
  const hasCurrentDocument = computed(() => Boolean(currentDocument.value && currentEnvelope.value));
  const canCreate = computed(() => !isTimeTravelActive.value && creatableDatabases.value.length > 0);
  const canSave = computed(() => Boolean(currentDatabase.value && canMutateGrid({
    canUpdate: currentCanUpdate.value,
    hasDocument: hasCurrentDocument.value,
    dirty: isDirty.value,
    timeTravelDate: launchTimeTravelDate.value,
    viewingHistorical: isViewingHistorical.value,
  })));
  const canDelete = computed(() => Boolean(currentDocument.value && currentCanDelete.value && currentDatabase.value && !gridReadOnly.value));
  const canRefresh = computed(() => Boolean(currentDatabase.value && currentDocument.value));
  const currentRevisionId = computed(() => revisionEntries.value.find((entry) => entry.isCurrent)?.revisionId ?? viewingHistoricalSnapshot.value?.revisionId ?? null);
  const timeTravelDateLabel = computed(() =>
    launchTimeTravelDate.value == null ? "" : new Date(launchTimeTravelDate.value).toLocaleString(),
  );

  onMounted(async () => {
    try {
      const bridge = createMindooDBAppBridge();
      const nextSession = await bridge.connect();
      session.value = nextSession;
      const context = await nextSession.getLaunchContext();
      launchTimeTravelDate.value = context.timeTravelDate ?? null;
      applyAppTheme(context.theme);
      cleanupTheme = nextSession.onThemeChange((theme) => applyAppTheme(theme));
      currentRuntime.value = context.runtime;
      hostUiPreferences.value = { ...context.uiPreferences };
      cleanupUiPreferences = nextSession.onUiPreferencesChange((uiPreferences) => {
        hostUiPreferences.value = { ...uiPreferences };
      });
      databases.value = context.databases;
      selectedDatabaseId.value = context.preferredDatabaseId
        ?? readableDatabases.value[0]?.id
        ?? context.databases[0]?.id
        ?? "";
      status.value = "Connected. Choose File / New or File / Open.";
    } catch (error) {
      status.value = readError(error);
    }
  });

  onBeforeUnmount(async () => {
    cleanupTheme?.();
    cleanupUiPreferences?.();
    await session.value?.disconnect();
  });

  async function openDatabaseById(databaseId: string) {
    if (!session.value || !databaseId) {
      throw new Error("Select a database first.");
    }
    return await session.value.openDatabase(databaseId);
  }

  async function refreshDocuments() {
    const database = await openDatabaseById(selectedDatabaseId.value);
    const result = await database.documents.list({
      status: "existing",
      limit: 100,
    });
    documents.value = [...result.items].sort((left, right) =>
      readDocumentSummaryLabel(left).localeCompare(readDocumentSummaryLabel(right), undefined, {
        sensitivity: "base",
        numeric: true,
      }));
  }

  async function createNewDocument() {
    try {
      const targetDatabaseInfo = selectedDatabaseInfo.value?.capabilities.includes("create")
        ? selectedDatabaseInfo.value
        : creatableDatabases.value[0];
      if (!targetDatabaseInfo) {
        throw new Error("No writable database is available.");
      }
      selectedDatabaseId.value = targetDatabaseInfo.id;
      const database = await openDatabaseById(targetDatabaseInfo.id);
      const envelope = createTeamGridDocument("Untitled spreadsheet");
      const document = await database.documents.create({
        set: envelope as unknown as Record<string, unknown>,
      });
      loadDocument(database, targetDatabaseInfo.id, document);
      status.value = `Created ${document.id}.`;
    } catch (error) {
      status.value = readError(error);
    }
  }

  async function openDocument(documentId: string) {
    try {
      const databaseId = selectedDatabaseId.value;
      const database = await openDatabaseById(databaseId);
      const document = await database.documents.get(documentId);
      if (!document) {
        throw new Error("Select a document to open.");
      }
      loadDocument(database, databaseId, document);
      status.value = `Opened ${document.id}.`;
    } catch (error) {
      status.value = readError(error);
    }
  }

  async function refreshCurrentDocument() {
    if (!currentDatabase.value || !currentDocument.value) {
      status.value = "Open a document before refreshing.";
      return;
    }
    try {
      const document = await currentDatabase.value.documents.get(currentDocument.value.id);
      if (!document) {
        throw new Error("The document is no longer available.");
      }
      loadDocument(currentDatabase.value, currentDatabaseId.value, document);
      status.value = "Reloaded the current spreadsheet.";
    } catch (error) {
      status.value = readError(error);
    }
  }

  async function saveDocument() {
    if (!currentDatabase.value || !currentDocument.value || !currentEnvelope.value) {
      status.value = "Open or create a spreadsheet before saving.";
      return;
    }
    if (gridReadOnly.value) {
      status.value = isTimeTravelActive.value
        ? "Time travel mode is read-only."
        : "Historical revisions are read-only. Return to the current version before saving.";
      return;
    }
    if (!currentCanUpdate.value) {
      status.value = "This application does not have write access to the current database.";
      return;
    }
    try {
      if (!hasTeamGridOperations(pendingOps.value)) {
        status.value = "No granular spreadsheet changes are pending.";
        isDirty.value = false;
        return;
      }
      const optimisticEnvelope = currentEnvelope.value;
      const updated = await currentDatabase.value.documents.update(
        currentDocument.value.id,
        serializeTeamGridOperations(pendingOps.value, { baseHeads: pendingOpsBaseHeads.value }),
      );
      const returnedEnvelope = migrateTeamGridDocument(updated.data);
      loadDocument(currentDatabase.value, currentDatabaseId.value, updated);
      const reconciled = JSON.stringify(returnedEnvelope.teamgrid) !== JSON.stringify(optimisticEnvelope.teamgrid);
      status.value = reconciled ? "Saved and merged concurrent changes." : "Saved.";
    } catch (error) {
      status.value = readError(error);
    }
  }

  async function deleteCurrentDocument() {
    if (!currentDatabase.value || !currentDocument.value) {
      status.value = "Open a spreadsheet before deleting.";
      return;
    }
    if (!canDelete.value) {
      status.value = "This application cannot delete the current spreadsheet.";
      return;
    }
    try {
      const deletedId = currentDocument.value.id;
      await currentDatabase.value.documents.delete(deletedId);
      currentDocument.value = null;
      currentEnvelope.value = null;
      viewingHistoricalSnapshot.value = null;
      isDirty.value = false;
      status.value = `Deleted ${deletedId}.`;
    } catch (error) {
      status.value = readError(error);
    }
  }

  async function openRevisionDialog() {
    if (!currentDatabase.value || !currentDocument.value || !currentCanBrowseHistory.value) {
      return;
    }
    revisionLoading.value = true;
    revisionErrorMessage.value = null;
    try {
      revisionEntries.value = await currentDatabase.value.documents.listHistory(currentDocument.value.id);
    } catch (error) {
      revisionErrorMessage.value = readError(error);
    } finally {
      revisionLoading.value = false;
    }
  }

  async function loadHistoricalRevision(revisionId: MindooDBAppDocumentRevisionId) {
    if (!currentDatabase.value || !currentDocument.value) {
      return;
    }
    if (revisionEntries.value.find((entry) => entry.revisionId === revisionId)?.isCurrent) {
      returnToCurrent();
      return;
    }
    try {
      const snapshot = await currentDatabase.value.documents.getAtRevision(currentDocument.value.id, revisionId);
      if (snapshot.state !== "exists" || !snapshot.data) {
        status.value = snapshot.state === "deleted"
          ? "That revision is a deletion marker and cannot be opened."
          : "That revision is no longer available.";
        return;
      }
      viewingHistoricalSnapshot.value = snapshot;
      currentEnvelope.value = migrateTeamGridDocument(snapshot.data);
      pendingOps.value = [];
      pendingOpsBaseHeads.value = [];
      isDirty.value = false;
      status.value = "Opened historical revision read-only.";
    } catch (error) {
      status.value = readError(error);
    }
  }

  function returnToCurrent() {
    if (currentDocument.value) {
      currentEnvelope.value = migrateTeamGridDocument(currentDocument.value.data);
    }
    viewingHistoricalSnapshot.value = null;
    pendingOps.value = [];
    pendingOpsBaseHeads.value = [];
    isDirty.value = false;
    status.value = "Returned to the current spreadsheet.";
  }

  function updateGrid(
    mutator: (grid: TeamGridDocumentV1, envelope: TeamGridDocumentEnvelope) => TeamGridOperation[] | void,
  ) {
    if (!currentEnvelope.value || gridReadOnly.value) {
      return;
    }
    const nextEnvelope = {
      ...currentEnvelope.value,
      teamgrid: cloneTeamGridDocument(currentEnvelope.value.teamgrid),
    };
    const operations = mutator(nextEnvelope.teamgrid, nextEnvelope) ?? [];
    if (operations.length > 0) {
      if (pendingOps.value.length === 0) {
        pendingOpsBaseHeads.value = currentDocument.value?.heads ? [...currentDocument.value.heads] : [];
      }
      pendingOps.value.push(...operations);
    }
    currentEnvelope.value = nextEnvelope;
    isDirty.value = true;
  }

  function loadDocument(database: MindooDBAppDatabase, databaseId: string, document: MindooDBAppDocument) {
    currentDatabase.value = database;
    currentDatabaseId.value = databaseId;
    currentDocument.value = document;
    currentEnvelope.value = migrateTeamGridDocument(document.data);
    viewingHistoricalSnapshot.value = null;
    pendingOps.value = [];
    pendingOpsBaseHeads.value = [];
    isDirty.value = false;
  }

  return {
    databases,
    selectedDatabaseId,
    documents,
    currentDocument,
    currentEnvelope,
    currentRuntime,
    hostUiPreferences,
    status,
    isDirty,
    revisionEntries,
    revisionLoading,
    revisionErrorMessage,
    pendingOps,
    pendingOpsBaseHeads,
    readableDatabases,
    selectedDatabaseInfo,
    currentDatabaseInfo,
    currentCanBrowseHistory,
    isTimeTravelActive,
    isViewingHistorical,
    gridReadOnly,
    activeGrid,
    activeSubject,
    canCreate,
    canSave,
    canDelete,
    canRefresh,
    currentRevisionId,
    timeTravelDateLabel,
    refreshDocuments,
    createNewDocument,
    openDocument,
    refreshCurrentDocument,
    saveDocument,
    deleteCurrentDocument,
    openRevisionDialog,
    loadHistoricalRevision,
    returnToCurrent,
    updateGrid,
  };
}

export function readDocumentSummaryLabel(summary: MindooDBAppDocumentSummary) {
  return readSubject(summary.data) || summary.id;
}

function readError(error: unknown) {
  return error instanceof Error ? error.message : String(error);
}
