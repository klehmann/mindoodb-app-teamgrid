/**
 * Worksheet lifecycle dialogs and mutations: add, rename (modal-driven so
 * we never depend on `window.prompt`, which Haven's sandboxed iframe
 * suppresses), and tombstone.
 */
import { ref } from "vue";
import { DEFAULT_COLUMN_WIDTH } from "@/shared/lib/gridDimensions";
import {
  createId,
  getFirstVisibleWorksheet,
  type TeamGridDocumentV1,
  type WorksheetId,
} from "@/features/document/lib/teamgridDocument";
import type { TeamGridAppApi } from "@/features/document/composables/useTeamGridDocument";

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

  /**
   * Append a new worksheet to the workbook seeded with 24 rows x 12 columns
   * (a reasonable Excel-like default) and switch the UI to the new tab.
   */
  function addWorksheet() {
    app.updateGrid((grid) => {
      const worksheetId = createId("sheet");
      const rowOrder = Array.from({ length: 24 }, () => createId("row"));
      const columnOrder = Array.from({ length: 12 }, () => createId("col"));
      const worksheet = {
        id: worksheetId,
        title: nextWorksheetTitle(grid),
        rowOrder,
        columnOrder,
        rowsById: Object.fromEntries(rowOrder.map((id) => [id, { id }])),
        columnsById: Object.fromEntries(columnOrder.map((id) => [id, { id, width: DEFAULT_COLUMN_WIDTH }])),
        cellsById: {},
      };
      grid.workbook.worksheetOrder.push(worksheetId);
      grid.workbook.worksheetsById[worksheetId] = worksheet;
      activeWorksheetId.value = worksheetId;
      return [{ type: "addWorksheet", worksheet, index: grid.workbook.worksheetOrder.length - 1 }];
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

  return {
    renameDialogVisible,
    renameTargetId,
    renameDraft,
    addWorksheet,
    renameWorksheet,
    applyWorksheetRename,
    deleteWorksheet,
  };
}
