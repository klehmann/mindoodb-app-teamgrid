/**
 * State and orchestration for the File / Open dialog.
 *
 * Owns the view-navigator lifecycle (build, refresh, dispose) so the dialog
 * never leaks server-side navigator sessions when the user switches
 * databases or dismisses the dialog. The navigator is built with
 * `category_then_document` ordering so each category appears once with its
 * full descendant document count.
 */
import { ref } from "vue";
import type { MindooDBAppViewNavigator } from "mindoodb-app-sdk";
import {
  ALL_SPREADSHEETS_NODE_KEY,
  buildOpenCategoryTree,
  collectNavigatorEntries,
  createOpenViewDefinition,
  dedupeDocumentEntries,
  mapDocumentEntries,
  type OpenCategoryNode,
  type OpenDocumentRow,
} from "@/features/document/lib/viewOpen";
import type { TeamGridAppApi } from "@/features/document/composables/useTeamGridDocument";

export interface UseOpenDialogOptions {
  app: TeamGridAppApi;
  onError: (error: unknown) => void;
}

export function useOpenDialog(options: UseOpenDialogOptions) {
  const { app, onError } = options;
  const openDialogVisible = ref(false);
  const selectedOpenDocId = ref("");
  const selectedOpenCategoryKey = ref(ALL_SPREADSHEETS_NODE_KEY);
  const openCategoryNodes = ref<OpenCategoryNode[]>([]);
  const openDialogDocuments = ref<OpenDocumentRow[]>([]);
  const allOpenDialogDocuments = ref<OpenDocumentRow[]>([]);
  const openNavigator = ref<MindooDBAppViewNavigator | null>(null);

  async function openFileDialog() {
    try {
      await rebuildOpenNavigator();
      openDialogVisible.value = true;
    } catch (error) {
      onError(error);
    }
  }

  /**
   * Re-open the navigator when the user switches databases inside the dialog.
   *
   * If the previously selected document still exists in the new database we
   * keep the selection so the keyboard focus stays predictable; otherwise we
   * fall back to the first document.
   */
  async function handleOpenDatabaseChange() {
    try {
      const previousDocumentId = selectedOpenDocId.value;
      await rebuildOpenNavigator();
      selectedOpenDocId.value = openDialogDocuments.value.some((document) => document.id === previousDocumentId)
        ? previousDocumentId
        : openDialogDocuments.value[0]?.id ?? "";
    } catch (error) {
      onError(error);
    }
  }

  async function selectOpenCategory(key: string) {
    selectedOpenCategoryKey.value = key;
    await refreshOpenDocumentsForSelectedCategory();
  }

  /**
   * Confirm the Open dialog: load the selected document and tear down the
   * navigator. The navigator must be disposed even on the success path
   * because the dialog's `@hide` only fires for dismissal-style closes.
   */
  async function openSelectedDocument() {
    if (!selectedOpenDocId.value) {
      return;
    }
    await app.openDocument(selectedOpenDocId.value);
    openDialogVisible.value = false;
    await disposeOpenNavigator();
  }

  async function rebuildOpenNavigator() {
    const databaseInfo = app.selectedDatabaseInfo.value;
    if (!databaseInfo?.capabilities.includes("views")) {
      throw new Error("This database does not expose the views capability required for categorized Open.");
    }
    await disposeOpenNavigator();
    const navigator = await app.createViewNavigator({
      databaseIds: [app.selectedDatabaseId.value],
      definition: createOpenViewDefinition(),
      categorizationStyle: "category_then_document",
      options: {
        includeCategories: true,
        includeDocuments: true,
        hideEmptyCategories: true,
      },
    });
    await navigator.expandAll();
    const entries = await collectNavigatorEntries(navigator);
    const documents = dedupeDocumentEntries(entries);
    const categories = buildOpenCategoryTree(entries.filter((entry) => entry.kind === "category"), documents.length);
    openNavigator.value = navigator;
    openCategoryNodes.value = categories.roots;
    selectedOpenCategoryKey.value = ALL_SPREADSHEETS_NODE_KEY;
    allOpenDialogDocuments.value = documents;
    openDialogDocuments.value = documents;
    selectedOpenDocId.value = documents[0]?.id ?? "";
  }

  async function refreshOpenDocumentsForSelectedCategory() {
    const navigator = openNavigator.value;
    if (!navigator || selectedOpenCategoryKey.value === ALL_SPREADSHEETS_NODE_KEY) {
      openDialogDocuments.value = allOpenDialogDocuments.value;
    } else {
      openDialogDocuments.value = mapDocumentEntries(await navigator.childDocuments(selectedOpenCategoryKey.value));
    }
    if (!openDialogDocuments.value.some((document) => document.id === selectedOpenDocId.value)) {
      selectedOpenDocId.value = openDialogDocuments.value[0]?.id ?? "";
    }
  }

  async function disposeOpenNavigator() {
    await openNavigator.value?.dispose();
    openNavigator.value = null;
  }

  return {
    openDialogVisible,
    selectedOpenDocId,
    selectedOpenCategoryKey,
    openCategoryNodes,
    openDialogDocuments,
    openFileDialog,
    handleOpenDatabaseChange,
    selectOpenCategory,
    openSelectedDocument,
    disposeOpenNavigator,
  };
}
