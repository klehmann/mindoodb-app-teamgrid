/**
 * Composable that owns Teamgrid's bridge connection, document lifecycle,
 * capability gating, time travel, and granular-save bookkeeping.
 *
 * This is the central public surface of the sample app: the root `App.vue`
 * binds to the refs returned here, and every UI mutation goes through
 * {@link useTeamGridDocument.updateGrid} so that an Automerge-friendly JSON
 * patch can be produced when the user saves.
 *
 * Design notes that are easy to miss when reading the code top-to-bottom:
 *
 * - The composable keeps two parallel representations of the document. The
 *   `currentDocument` ref is the raw `MindooDBAppDocument` we got from Haven,
 *   used for IDs and Automerge `heads`. The `currentEnvelope` ref is a typed
 *   {@link TeamGridDocumentEnvelope} clone the UI freely mutates without
 *   touching the SDK's payload object.
 * - Each call to {@link useTeamGridDocument.updateGrid} clones the workbook
 *   before applying the mutator. That makes Vue's reactivity see a new object
 *   reference and, more importantly, prevents the mutator from accidentally
 *   modifying state that is shared with the previous `currentEnvelope`.
 * - The first mutation in a save batch captures `currentDocument.heads` into
 *   `pendingOpsBaseHeads`. That value is later attached as `baseHeads` on the
 *   JSON patch so the Haven host (and ultimately Automerge) can replay the
 *   patch at the document version the user actually saw, instead of at HEAD.
 *   Without this, two users inserting rows concurrently would race on list
 *   indices and last-writer-wins.
 * - Time travel and historical-revision snapshots flip the grid to read-only
 *   via {@link isGridSessionReadOnly}. Mutators silently short-circuit in
 *   those modes rather than throwing, because the UI already disables the
 *   triggering menu/toolbar items via {@link canMutateGrid}.
 */
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
  type MindooDBAppCreateViewNavigatorInput,
  type MindooDBAppResolvedViewDefinition,
  type MindooDBAppViewNavigator,
  type MindooDBAppViewNavigatorOpenOptions,
  type MindooDBAppRuntime,
  type MindooDBAppSession,
  type MindooDBAppUiPreferences,
} from "mindoodb-app-sdk";

import { applyAppTheme } from "@/shared/lib/theme";
import { canMutateGrid, isGridSessionReadOnly } from "@/shared/lib/capabilities";
import {
  hasTeamGridOperations,
  serializeTeamGridOperations,
  type TeamGridOperation,
} from "@/features/document/lib/teamgridOps";
import {
  createTeamGridDocument,
  cloneTeamGridDocument,
  migrateTeamGridDocument,
  readSubject,
  type TeamGridDocumentEnvelope,
  type TeamGridDocumentV1,
} from "@/features/document/lib/teamgridDocument";

interface OpenSpreadsheetSession {
  id: string;
  databaseId: string;
  database: MindooDBAppDatabase;
  documentId: string;
  document: MindooDBAppDocument;
  envelope: TeamGridDocumentEnvelope;
  pendingOps: TeamGridOperation[];
  pendingOpsBaseHeads: string[];
  isDirty: boolean;
}

/**
 * Construct the Teamgrid document composable.
 *
 * Returns refs and helpers that `App.vue` and its child components bind to.
 * The returned API is intentionally flat: every interesting piece of state
 * is a {@link Ref} that templates can read directly, and every transition is
 * an `async` helper that swallows errors into the human-readable `status` ref.
 */
export function useTeamGridDocument() {
  const session = ref<MindooDBAppSession | null>(null);
  const databases = ref<MindooDBAppDatabaseInfo[]>([]);
  const configuredViews = ref<MindooDBAppResolvedViewDefinition[]>([]);
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
  const lastErrorMessage = ref<string | null>(null);
  const isDirty = ref(false);
  const revisionEntries = ref<MindooDBAppDocumentHistoryEntry[]>([]);
  const revisionLoading = ref(false);
  const revisionErrorMessage = ref<string | null>(null);
  const pendingOps = ref<TeamGridOperation[]>([]);
  const pendingOpsBaseHeads = ref<string[]>([]);
  const openSpreadsheetSessions = ref<OpenSpreadsheetSession[]>([]);
  const activeSpreadsheetSessionId = ref("");

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
  const activeTags = computed(() => currentEnvelope.value?.tags ?? []);
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
  const openSessions = computed(() => openSpreadsheetSessions.value.map((session) => ({
    id: session.id,
    documentId: session.documentId,
    databaseId: session.databaseId,
    title: session.envelope.subject || session.documentId,
    isActive: session.id === activeSpreadsheetSessionId.value,
    isDirty: session.id === activeSpreadsheetSessionId.value ? isDirty.value : session.isDirty,
  })));

  /**
   * Connect to the Haven host bridge once Vue mounts the consuming component.
   *
   * The launch context tells the app which databases it is allowed to see,
   * which runtime it runs under (iframe vs. native), whether the host opened
   * the database in time-travel mode, and the initial theme. Theme and
   * UI-preference changes are then pushed by the host via the listeners
   * registered below, which we tear down in {@link onBeforeUnmount}.
   */
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
      configuredViews.value = context.views;
      selectedDatabaseId.value = context.preferredDatabaseId
        ?? readableDatabases.value[0]?.id
        ?? context.databases[0]?.id
        ?? "";
      status.value = "Connected. Choose File / New or File / Open.";
    } catch (error) {
      showError(error);
    }
  });

  onBeforeUnmount(async () => {
    cleanupTheme?.();
    cleanupUiPreferences?.();
    await session.value?.disconnect();
  });

  /**
   * Resolve a database handle via the SDK session.
   *
   * Throws when the session is not yet connected or when the caller did not
   * pick a database id. The session caches handles internally, so calling
   * this repeatedly with the same id is cheap.
   */
  async function openDatabaseById(databaseId: string) {
    if (!session.value || !databaseId) {
      throw new Error("Select a database first.");
    }
    return await session.value.openDatabase(databaseId);
  }

  /**
   * Create a dynamic view navigator on the current session.
   *
   * Exposed so the File/Open dialog in `App.vue` can build a tag-categorized
   * tree without having to import the bridge directly.
   */
  async function createViewNavigator(input: MindooDBAppCreateViewNavigatorInput): Promise<MindooDBAppViewNavigator> {
    if (!session.value) {
      throw new Error("Connect to Haven before opening a view.");
    }
    return await session.value.createViewNavigator(input);
  }

  async function openViewNavigator(viewId: string, options?: MindooDBAppViewNavigatorOpenOptions): Promise<MindooDBAppViewNavigator> {
    if (!session.value) {
      throw new Error("Connect to Haven before opening a view.");
    }
    return await session.value.openViewNavigator(viewId, options);
  }

  /**
   * Refresh the legacy flat document list for the currently selected database.
   *
   * The view-backed Open dialog uses {@link createViewNavigator} instead, but
   * this helper is still used by callers that just need a sorted summary list
   * (for example to resolve a previously selected document id).
   */
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

  /**
   * Create a brand-new spreadsheet document.
   *
   * Falls back to the first creatable database when the currently selected
   * one does not advertise the `create` capability, so the menu's "New"
   * action remains useful even when the user only has read access to the
   * primary database.
   */
  async function createNewDocument(envelopeOverride?: TeamGridDocumentEnvelope) {
    try {
      const targetDatabaseInfo = selectedDatabaseInfo.value?.capabilities.includes("create")
        ? selectedDatabaseInfo.value
        : creatableDatabases.value[0];
      if (!targetDatabaseInfo) {
        throw new Error("No writable database is available.");
      }
      selectedDatabaseId.value = targetDatabaseInfo.id;
      const database = await openDatabaseById(targetDatabaseInfo.id);
      const envelope = envelopeOverride ?? createTeamGridDocument("Untitled spreadsheet");
      const document = await database.documents.create({
        set: envelope as unknown as Record<string, unknown>,
      });
      loadDocument(database, targetDatabaseInfo.id, document);
      status.value = `Created ${document.id}.`;
    } catch (error) {
      showError(error);
    }
  }

  async function createDocumentFromEnvelope(envelope: TeamGridDocumentEnvelope) {
    await createNewDocument(envelope);
  }

  async function createDocumentFromTemplate(templateDocumentId: string) {
    try {
      const sourceDatabaseId = selectedDatabaseId.value;
      const sourceDatabase = await openDatabaseById(sourceDatabaseId);
      const templateDocument = await sourceDatabase.documents.get(templateDocumentId);
      if (!templateDocument) {
        throw new Error("Select a template to use.");
      }
      const templateEnvelope = migrateTeamGridDocument(templateDocument.data);
      const envelope: TeamGridDocumentEnvelope = {
        ...templateEnvelope,
        subject: `Copy of ${templateEnvelope.subject || "Untitled spreadsheet"}`,
        istemplate: false,
        teamgrid: cloneTeamGridDocument(templateEnvelope.teamgrid),
      };
      await createNewDocument(envelope);
    } catch (error) {
      showError(error);
    }
  }

  /** Open an existing document by id from the currently selected database. */
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
      showError(error);
    }
  }

  /**
   * Re-read the currently open document from Haven and discard local edits.
   *
   * This drops any pending granular operations and resets the dirty flag, so
   * the UI should warn the user before calling this when there are unsaved
   * changes.
   */
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
      showError(error);
    }
  }

  /**
   * Flush all pending granular operations to Haven.
   *
   * The pending operations are serialized into one {@link MindooDBAppJsonPatch}
   * with `baseHeads` set to the document heads captured when the batch
   * started. Haven applies the patch against that historical version using
   * `Automerge.changeAt`, then merges the result with any concurrent writes.
   *
   * If the returned payload differs from the optimistic envelope we sent
   * (for example, because another collaborator inserted a row), we surface
   * "Saved and merged concurrent changes." so the user knows the visible
   * sheet has been reconciled rather than just round-tripped.
   */
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
      isDirty.value = false;
      snapshotActiveSession();
      const reconciled = JSON.stringify(returnedEnvelope) !== JSON.stringify(optimisticEnvelope);
      status.value = reconciled ? "Saved and merged concurrent changes." : "Saved.";
    } catch (error) {
      showError(error);
    }
  }

  /** Hard-delete the currently open document from its database. */
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
      const deletedSessionId = activeSpreadsheetSessionId.value;
      await currentDatabase.value.documents.delete(deletedId);
      openSpreadsheetSessions.value = openSpreadsheetSessions.value.filter((session) => session.id !== deletedSessionId);
      const nextSession = openSpreadsheetSessions.value[0] ?? null;
      if (nextSession) {
        activateSession(nextSession);
      } else {
        clearActiveDocumentState();
      }
      status.value = `Deleted ${deletedId}.`;
    } catch (error) {
      showError(error);
    }
  }

  /**
   * Load the document history list shown in the Revisions dialog.
   *
   * No-ops when the host did not advertise the `history` capability so the
   * sample app still works against minimal database bindings.
   */
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

  /**
   * Switch the editor to a historical snapshot.
   *
   * The grid becomes read-only via {@link gridReadOnly}, pending operations
   * are cleared, and the snapshot data is migrated through
   * {@link migrateTeamGridDocument} so older schemas still render.
   */
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
      showError(error);
    }
  }

  /**
   * Leave historical-revision mode and reattach the editor to the current
   * version. The current document's data is re-migrated so we never carry
   * stale grid state from a previously displayed snapshot.
   */
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

  /**
   * Apply a UI mutation to the workbook and record the equivalent granular
   * operations for the next save.
   *
   * `mutator` receives a freshly cloned workbook so it can mutate it in place
   * without worrying about reactivity. It must return the list of semantic
   * {@link TeamGridOperation} entries that describe the same change, so the
   * saver can later serialize them into an Automerge-friendly JSON patch.
   *
   * The first mutation in a batch captures `currentDocument.heads` into
   * `pendingOpsBaseHeads`. Subsequent mutations append to the same batch
   * without re-capturing heads, so the entire batch is replayed at the
   * version the user saw when they started editing.
   *
   * Silently no-ops in read-only modes; the UI is expected to disable the
   * triggering controls in those modes (see {@link gridReadOnly}).
   */
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
    snapshotActiveSession();
  }

  /**
   * Replace the editor's state with a freshly fetched document.
   *
   * Resets pending operations and the dirty flag because the new document
   * starts clean. Callers should make sure they don't overwrite unsaved
   * local edits before invoking this.
   */
  function loadDocument(database: MindooDBAppDatabase, databaseId: string, document: MindooDBAppDocument) {
    snapshotActiveSession();
    const id = createSessionId(databaseId, document.id);
    const session: OpenSpreadsheetSession = {
      id,
      databaseId,
      database,
      documentId: document.id,
      document,
      envelope: migrateTeamGridDocument(document.data),
      pendingOps: [],
      pendingOpsBaseHeads: [],
      isDirty: false,
    };
    const existingIndex = openSpreadsheetSessions.value.findIndex((candidate) => candidate.id === id);
    if (existingIndex >= 0) {
      openSpreadsheetSessions.value.splice(existingIndex, 1, session);
    } else {
      openSpreadsheetSessions.value.push(session);
    }
    activateSession(session);
  }

  function switchToOpenSession(sessionId: string) {
    snapshotActiveSession();
    const session = openSpreadsheetSessions.value.find((candidate) => candidate.id === sessionId);
    if (!session) {
      status.value = "That spreadsheet window is no longer open.";
      return;
    }
    activateSession(session);
    status.value = `Switched to ${session.envelope.subject || session.documentId}.`;
  }

  function closeOpenSession(sessionId: string) {
    snapshotActiveSession();
    const session = openSpreadsheetSessions.value.find((candidate) => candidate.id === sessionId);
    if (!session) {
      return;
    }
    if (session.isDirty) {
      status.value = "Save the spreadsheet before closing its window.";
      return;
    }
    openSpreadsheetSessions.value = openSpreadsheetSessions.value.filter((candidate) => candidate.id !== sessionId);
    if (activeSpreadsheetSessionId.value === sessionId) {
      const nextSession = openSpreadsheetSessions.value[0] ?? null;
      if (nextSession) {
        activateSession(nextSession);
      } else {
        clearActiveDocumentState();
      }
    }
  }

  function activateSession(session: OpenSpreadsheetSession) {
    activeSpreadsheetSessionId.value = session.id;
    currentDatabase.value = session.database;
    currentDatabaseId.value = session.databaseId;
    currentDocument.value = session.document;
    currentEnvelope.value = session.envelope;
    viewingHistoricalSnapshot.value = null;
    pendingOps.value = [...session.pendingOps];
    pendingOpsBaseHeads.value = [...session.pendingOpsBaseHeads];
    isDirty.value = session.isDirty;
  }

  function snapshotActiveSession() {
    const session = openSpreadsheetSessions.value.find((candidate) => candidate.id === activeSpreadsheetSessionId.value);
    if (!session || !currentDocument.value || !currentEnvelope.value) {
      return;
    }
    if (currentDatabase.value) {
      session.database = currentDatabase.value;
    }
    session.document = currentDocument.value;
    session.envelope = currentEnvelope.value;
    session.pendingOps = [...pendingOps.value];
    session.pendingOpsBaseHeads = [...pendingOpsBaseHeads.value];
    session.isDirty = isDirty.value;
  }

  function clearActiveDocumentState() {
    activeSpreadsheetSessionId.value = "";
    currentDatabase.value = null;
    currentDatabaseId.value = "";
    currentDocument.value = null;
    currentEnvelope.value = null;
    viewingHistoricalSnapshot.value = null;
    pendingOps.value = [];
    pendingOpsBaseHeads.value = [];
    isDirty.value = false;
  }

  function createSessionId(databaseId: string, documentId: string) {
    return `${databaseId}:${documentId}`;
  }

  function clearLastError() {
    lastErrorMessage.value = null;
  }

  function showError(error: unknown) {
    const message = readError(error);
    status.value = message;
    lastErrorMessage.value = message;
  }

  return {
    databases,
    configuredViews,
    selectedDatabaseId,
    documents,
    currentDocument,
    currentEnvelope,
    currentRuntime,
    hostUiPreferences,
    status,
    lastErrorMessage,
    isDirty,
    revisionEntries,
    revisionLoading,
    revisionErrorMessage,
    pendingOps,
    pendingOpsBaseHeads,
    openSessions,
    activeSpreadsheetSessionId,
    readableDatabases,
    selectedDatabaseInfo,
    currentDatabaseInfo,
    currentCanBrowseHistory,
    isTimeTravelActive,
    isViewingHistorical,
    gridReadOnly,
    activeGrid,
    activeSubject,
    activeTags,
    canCreate,
    canSave,
    canDelete,
    canRefresh,
    currentRevisionId,
    timeTravelDateLabel,
    refreshDocuments,
    createViewNavigator,
    openViewNavigator,
    createNewDocument,
    createDocumentFromEnvelope,
    createDocumentFromTemplate,
    openDocument,
    switchToOpenSession,
    closeOpenSession,
    refreshCurrentDocument,
    saveDocument,
    deleteCurrentDocument,
    openRevisionDialog,
    loadHistoricalRevision,
    returnToCurrent,
    updateGrid,
    clearLastError,
  };
}

/**
 * Pick a user-facing label for a document summary, preferring the saved
 * subject and falling back to the document id when no title is set.
 */
export function readDocumentSummaryLabel(summary: MindooDBAppDocumentSummary) {
  return readSubject(summary.data) || summary.id;
}

/**
 * Public surface of {@link useTeamGridDocument}, exported so the focused
 * composables in `features/grid` and `features/document` can express their
 * dependencies on the document lifecycle without re-declaring the entire
 * shape.
 */
export type TeamGridAppApi = ReturnType<typeof useTeamGridDocument>;

/** Normalize an unknown thrown value into a human-readable status message. */
function readError(error: unknown) {
  return error instanceof Error ? error.message : String(error);
}
