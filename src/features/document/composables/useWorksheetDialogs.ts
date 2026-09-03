/**
 * Worksheet lifecycle dialogs and mutations: add, rename (modal-driven so
 * we never depend on `window.prompt`, which Haven's sandboxed iframe
 * suppresses), and tombstone.
 */
import { ref } from "vue";
import { DEFAULT_COLUMN_WIDTH } from "@/shared/lib/gridDimensions";
import {
  createId,
  DEFAULT_WORKSHEET_COLUMNS,
  DEFAULT_WORKSHEET_ROWS,
  getFirstVisibleWorksheet,
  type TeamGridDocumentV1,
  type WorksheetId,
} from "@/features/document/lib/teamgridDocument";
import type { TeamGridAppApi } from "@/features/document/composables/useTeamGridDocument";
import {
  isWorksheetOrderSound,
  planWorksheetMoveTo,
  planWorksheetNudge,
  resolveWorksheetOrder,
  type WorksheetMovePlan,
} from "@/features/document/lib/worksheetOrder";
import {
  materializeViewSheet,
  type ViewSheetSettings,
} from "@/features/view-sheets/lib/viewSheetMaterialization";

export interface UseWorksheetDialogsOptions {
  app: TeamGridAppApi;
  /** Mutated when the active worksheet should change as a side effect. */
  activeWorksheetId: { value: WorksheetId | null };
}

export function useWorksheetDialogs(options: UseWorksheetDialogsOptions) {
  const { app, activeWorksheetId } = options;
  const renameDialogVisible = ref(false);
  const renameTargetId = ref<WorksheetId | null>(null);
  const renameDraft = ref("");
  const viewSheetDialogVisible = ref(false);
  const viewSheetTargetId = ref<WorksheetId | null>(null);
  const viewSheetNameDraft = ref("");
  const viewSheetViewIdDraft = ref("");
  const viewSheetShowDocuments = ref(true);
  const viewSheetShowCategories = ref(true);
  const viewSheetRootCategoryPathDraft = ref("");
  const viewSheetErrorMessage = ref<string | null>(null);
  const viewSheetSaving = ref(false);

  /**
   * Append a new worksheet to the workbook seeded with the standard blank
   * worksheet dimensions and switch the UI to the new tab.
   */
  function addWorksheet() {
    app.updateGrid((grid) => {
      const worksheetId = createId("sheet");
      const rowOrder = Array.from({ length: DEFAULT_WORKSHEET_ROWS }, () => createId("row"));
      const columnOrder = Array.from({ length: DEFAULT_WORKSHEET_COLUMNS }, () => createId("col"));
      const worksheet = {
        id: worksheetId,
        title: nextWorksheetTitle(grid),
        rowOrder,
        columnOrder,
        rowsById: Object.fromEntries(rowOrder.map((id) => [id, { id }])),
        columnsById: Object.fromEntries(columnOrder.map((id) => [id, { id, width: DEFAULT_COLUMN_WIDTH }])),
        cellsById: {},
        chartOrder: [],
        chartsById: {},
      };
      grid.workbook.worksheetOrder.push(worksheetId);
      grid.workbook.worksheetsById[worksheetId] = worksheet;
      activeWorksheetId.value = worksheetId;
      return [{ type: "addWorksheet", worksheet, index: grid.workbook.worksheetOrder.length - 1 }];
    });
  }

  /**
   * Drop a tab onto another one, or step it one place with the keyboard.
   *
   * The plan is made against the repaired list so its two indices and its
   * resulting order always describe the same move. Only a list that already
   * read back soundly can carry the minimal delete-plus-insert patch — for any
   * other, those indices would address entries that are not there, so the
   * whole order is rewritten once and the document is sound again afterwards.
   */
  function moveWorksheet(worksheetId: WorksheetId, toIndex: number) {
    applyWorksheetMove((grid) =>
      planWorksheetMoveTo(resolveWorksheetOrder(grid.workbook), worksheetId, toIndex));
  }

  function nudgeWorksheet(worksheetId: WorksheetId, offset: -1 | 1) {
    applyWorksheetMove((grid) => planWorksheetNudge(grid.workbook, worksheetId, offset));
  }

  function applyWorksheetMove(plan: (grid: TeamGridDocumentV1) => WorksheetMovePlan | null) {
    app.updateGrid((grid) => {
      const move = plan(grid);
      if (!move) {
        return [];
      }
      const wasSound = isWorksheetOrderSound(grid.workbook);
      const worksheetId = move.order[move.toIndex];
      grid.workbook.worksheetOrder = move.order;
      return wasSound && worksheetId
        ? [{ type: "moveWorksheet", worksheetId, fromIndex: move.fromIndex, toIndex: move.toIndex }]
        : [{ type: "repairWorksheetOrder", order: [...move.order] }];
    });
  }

  /**
   * Generate a "Sheet N" title that does not collide with any existing
   * worksheet, including tombstoned ones (their entries still live in the
   * workbook and may be restored, so we don't want to recycle their names).
   */
  function nextWorksheetTitle(grid: TeamGridDocumentV1) {
    const usedTitles = new Set(
      Object.values(grid.workbook.worksheetsById).map((worksheet) => worksheet.title),
    );
    for (let index = 1; index < 1000; index += 1) {
      const candidate = `Sheet ${index}`;
      if (!usedTitles.has(candidate)) {
        return candidate;
      }
    }
    return `Sheet ${Date.now()}`;
  }

  /** Open the rename dialog for the given worksheet, seeded with its current title. */
  function renameWorksheet(worksheetId: WorksheetId) {
    if (app.gridReadOnly.value) {
      return;
    }
    const currentTitle = app.activeGrid.value?.workbook.worksheetsById[worksheetId]?.title ?? "";
    renameTargetId.value = worksheetId;
    renameDraft.value = currentTitle;
    renameDialogVisible.value = true;
  }

  function applyWorksheetRename() {
    const worksheetId = renameTargetId.value;
    const nextTitle = renameDraft.value.trim();
    if (!worksheetId || !nextTitle) {
      renameDialogVisible.value = false;
      return;
    }
    app.updateGrid((grid) => {
      const worksheet = grid.workbook.worksheetsById[worksheetId];
      if (!worksheet || worksheet.title === nextTitle) {
        return [];
      }
      worksheet.title = nextTitle;
      return [{ type: "renameWorksheet", worksheetId, title: nextTitle }];
    });
    renameDialogVisible.value = false;
    renameTargetId.value = null;
  }

  /**
   * Tombstone a worksheet rather than removing it from the workbook entirely.
   *
   * Keeping the entry preserves stable references for cells and formulas
   * that pointed at this worksheet so the change merges cleanly with other
   * collaborators. The active worksheet jumps to the first remaining tab so
   * the grid is never left blank.
   */
  function deleteWorksheet(worksheetId: WorksheetId) {
    app.updateGrid((grid) => {
      const deletedAt = new Date().toISOString();
      grid.workbook.worksheetsById[worksheetId].deletedAt = deletedAt;
      activeWorksheetId.value = getFirstVisibleWorksheet(grid)?.id ?? null;
      return [{ type: "tombstoneWorksheet", worksheetId, deletedAt }];
    });
  }

  function openCreateViewSheetDialog() {
    if (app.gridReadOnly.value) {
      return;
    }
    const firstView = app.configuredViews.value[0] ?? null;
    viewSheetTargetId.value = null;
    viewSheetViewIdDraft.value = firstView?.id ?? "";
    viewSheetNameDraft.value = "";
    viewSheetShowDocuments.value = true;
    viewSheetShowCategories.value = true;
    viewSheetRootCategoryPathDraft.value = "";
    viewSheetErrorMessage.value = app.configuredViews.value.length === 0
      ? "No configured MindooDB views are available to this app."
      : null;
    viewSheetDialogVisible.value = true;
  }

  function openViewSheetSettings(worksheetId: WorksheetId) {
    if (app.gridReadOnly.value) {
      return;
    }
    const worksheet = app.activeGrid.value?.workbook.worksheetsById[worksheetId];
    const binding = worksheet?.viewBinding;
    if (!worksheet || !binding) {
      return;
    }
    viewSheetTargetId.value = worksheetId;
    viewSheetNameDraft.value = worksheet.title;
    viewSheetViewIdDraft.value = binding.viewId;
    viewSheetShowDocuments.value = binding.showDocuments;
    viewSheetShowCategories.value = binding.showCategories;
    viewSheetRootCategoryPathDraft.value = binding.rootCategoryPath.join("\\");
    viewSheetErrorMessage.value = null;
    viewSheetDialogVisible.value = true;
  }

  async function applyViewSheetSettings() {
    const settings = readViewSheetSettings();
    const view = app.configuredViews.value.find((candidate) => candidate.id === settings.viewId);
    if (!settings.title.trim()) {
      viewSheetErrorMessage.value = "Enter a sheet name.";
      return;
    }
    if (!view) {
      viewSheetErrorMessage.value = "Select a configured view.";
      return;
    }
    if (!settings.showDocuments && !settings.showCategories) {
      viewSheetErrorMessage.value = "Show documents, categories, or both.";
      return;
    }

    viewSheetSaving.value = true;
    viewSheetErrorMessage.value = null;
    try {
      const targetId = viewSheetTargetId.value;
      const existingWorksheet = targetId ? app.activeGrid.value?.workbook.worksheetsById[targetId] ?? null : null;
      const worksheet = await materializeViewSheet({
        settings,
        view,
        existingWorksheet,
        openViewNavigator: app.openViewNavigator,
      });
      app.updateGrid((grid) => {
        if (targetId) {
          grid.workbook.worksheetsById[worksheet.id] = worksheet;
          activeWorksheetId.value = worksheet.id;
          return [{ type: "replaceWorksheet", worksheet }];
        }
        grid.workbook.worksheetOrder.push(worksheet.id);
        grid.workbook.worksheetsById[worksheet.id] = worksheet;
        activeWorksheetId.value = worksheet.id;
        return [{ type: "addWorksheet", worksheet, index: grid.workbook.worksheetOrder.length - 1 }];
      });
      app.status.value = `Refreshed ${worksheet.title}.`;
      viewSheetDialogVisible.value = false;
      viewSheetTargetId.value = null;
    } catch (error) {
      viewSheetErrorMessage.value = error instanceof Error ? error.message : String(error);
    } finally {
      viewSheetSaving.value = false;
    }
  }

  function readViewSheetSettings(): ViewSheetSettings {
    return {
      title: viewSheetNameDraft.value,
      viewId: viewSheetViewIdDraft.value,
      showDocuments: viewSheetShowDocuments.value,
      showCategories: viewSheetShowCategories.value,
      rootCategoryPathInput: viewSheetRootCategoryPathDraft.value,
    };
  }

  return {
    renameDialogVisible,
    renameTargetId,
    renameDraft,
    viewSheetDialogVisible,
    viewSheetTargetId,
    viewSheetNameDraft,
    viewSheetViewIdDraft,
    viewSheetShowDocuments,
    viewSheetShowCategories,
    viewSheetRootCategoryPathDraft,
    viewSheetErrorMessage,
    viewSheetSaving,
    configuredViews: app.configuredViews,
    addWorksheet,
    moveWorksheet,
    nudgeWorksheet,
    renameWorksheet,
    applyWorksheetRename,
    deleteWorksheet,
    openCreateViewSheetDialog,
    openViewSheetSettings,
    applyViewSheetSettings,
  };
}
