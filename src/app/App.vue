<script setup lang="ts">
/**
 * Root component for the TeamGrid sample app.
 *
 * `App.vue` is intentionally thin: it wires focused composables together
 * and forwards their state into PrimeVue widgets. The heavy lifting lives
 * in the composables next to each feature:
 *
 * - Document lifecycle, Haven bridge, save / history: `useTeamGridDocument`
 *   (`features/document/composables`).
 * - Selection state and helpers: `useSelection`.
 * - Formula-bar editing, reference highlighting, formula picking:
 *   `useFormulaBarEditing`.
 * - Floating formula-assist panel routing: `useFormulaAssistRouter`.
 * - Copy / cut / paste pipeline: `useGridClipboard`.
 * - Excel-like format dialog: `useCellFormatDialog`.
 * - File / Open dialog and view navigator: `useOpenDialog`.
 * - Spreadsheet properties dialog: `useDocumentPropertiesDialog`.
 * - Worksheet add / rename / delete dialogs: `useWorksheetDialogs`.
 * - Error dialog: `useErrorDialog`.
 *
 * What stays here is the orchestration that ties the grid to the
 * composables: cell commits, row / column / cell mutations triggered from
 * the menu, save with pending-edit flush, and the PrimeVue menu models.
 */
import { computed, nextTick, ref, watch } from "vue";
import Button from "primevue/button";
import ContextMenu from "primevue/contextmenu";
import Menubar from "primevue/menubar";
import type { MenuItem } from "primevue/menuitem";

import DocumentRevisionDialog from "@/features/document/components/DocumentRevisionDialog.vue";
import DocumentPropertiesDialog from "@/features/document/components/DocumentPropertiesDialog.vue";
import DeleteSpreadsheetDialog from "@/features/document/components/DeleteSpreadsheetDialog.vue";
import ErrorDialog from "@/features/document/components/ErrorDialog.vue";
import FormulaAssistPanel from "@/features/grid/components/FormulaAssistPanel.vue";
import FormulaBar from "@/features/grid/components/FormulaBar.vue";
import GridViewport from "@/features/grid/components/GridViewport.vue";
import OpenSpreadsheetDialog from "@/features/document/components/OpenSpreadsheetDialog.vue";
import RenameWorksheetDialog from "@/features/document/components/RenameWorksheetDialog.vue";
import CellFormatDialog from "@/features/grid/components/CellFormatDialog.vue";
import WorksheetTabs from "@/features/grid/components/WorksheetTabs.vue";

import { useTeamGridDocument } from "@/features/document/composables/useTeamGridDocument";
import { useErrorDialog } from "@/features/document/composables/useErrorDialog";
import { useOpenDialog } from "@/features/document/composables/useOpenDialog";
import { useDocumentPropertiesDialog } from "@/features/document/composables/useDocumentPropertiesDialog";
import { useWorksheetDialogs } from "@/features/document/composables/useWorksheetDialogs";

import { useSelection, type CellSelectionRange } from "@/features/grid/composables/useSelection";
import { useFormulaAssistRouter } from "@/features/grid/composables/useFormulaAssistRouter";
import { useFormulaBarEditing } from "@/features/grid/composables/useFormulaBarEditing";
import { useGridClipboard } from "@/features/grid/composables/useGridClipboard";
import { useCellFormatDialog } from "@/features/grid/composables/useCellFormatDialog";

import {
  coerceInputToCellValue,
  formulaResultToCellValue,
  preserveCompatibleCellValueFormat,
} from "@/features/grid/lib/cellFormatting";
import { evaluateFormula } from "@/features/formulas/lib";
import { DEFAULT_COLUMN_WIDTH } from "@/shared/lib/gridDimensions";
import {
  createId,
  getFirstVisibleWorksheet,
  type Cell,
  type ColumnId,
  type RowId,
  type WorksheetId,
} from "@/features/document/lib/teamgridDocument";
import { projectWorksheet } from "@/features/grid/lib/gridProjection";
import type { TeamGridOperation } from "@/features/document/lib/teamgridOps";
import { importTeamGridWorkbookBuffer } from "@/features/xlsx/lib/importWorkbook";
import { writeTeamGridExcelBuffer } from "@/features/xlsx/lib/exportWorkbook";

const app = useTeamGridDocument();

const cellContextMenu = ref<InstanceType<typeof ContextMenu> | null>(null);
const formulaBarComponent = ref<InstanceType<typeof FormulaBar> | null>(null);
const gridViewportComponent = ref<InstanceType<typeof GridViewport> | null>(null);
const xlsxImportInput = ref<HTMLInputElement | null>(null);

const deleteDialogVisible = ref(false);
const revisionDialogVisible = ref(false);
const saveInFlight = ref(false);

const activeWorksheetId = ref<WorksheetId | null>(null);
const cellContextRange = ref<CellSelectionRange | null>(null);

/**
 * The worksheet currently displayed in the grid.
 *
 * Defaults to whichever worksheet the user explicitly picked via the tab
 * strip, falling back to the first non-tombstoned worksheet so deleting
 * the active tab does not leave the UI without a worksheet to render.
 */
const activeWorksheet = computed(() => {
  if (!app.activeGrid.value) {
    return null;
  }
  const explicit = activeWorksheetId.value
    ? app.activeGrid.value.workbook.worksheetsById[activeWorksheetId.value]
    : null;
  return explicit && !explicit.deletedAt ? explicit : getFirstVisibleWorksheet(app.activeGrid.value);
});

/**
 * Render-time projection of {@link activeWorksheet} (ordered rows/columns
 * and A1 address lookups). Recomputed automatically when the worksheet
 * changes so consumers never have to track invalidation themselves.
 */
const projection = computed(() => activeWorksheet.value ? projectWorksheet(activeWorksheet.value) : null);

const locale = computed(() => app.activeGrid.value?.settings.locale);

const {
  selectedCellId,
  selectedRange,
  additionalRanges,
  allSelectedRanges,
  selectedCellAddress,
  selectedCell,
  selectedCells,
  hasSelection,
  findCellCoordinates,
  cellsForRange,
  boundsForRange,
} = useSelection({ activeWorksheet, projection });

const {
  formulaAssistOpen,
  formulaAssistAnchor,
  formulaAssistDraft,
  formulaAssistCaretPos,
  inlineCellEditing,
  inlineCellDraft,
  openFormulaAssist,
  handleFormulaAssistSelect,
  handleInlineEditState,
} = useFormulaAssistRouter({
  formulaBarTarget: formulaBarComponent,
  gridViewportTarget: gridViewportComponent,
});

const {
  formulaDraft,
  formulaError,
  formulaEditing,
  highlightedCellIds,
  appendPickedAddress,
  commitFormulaBar,
  cancelFormulaEdit,
} = useFormulaBarEditing({
  selectedCell,
  activeWorksheet,
  projection,
  locale,
  commitCell: (cell, rawValue) => commitCell(cell, rawValue),
  closeFormulaAssist: () => { formulaAssistOpen.value = false; },
});

const {
  clipboardSourceRange,
  internalClipboard,
  handleGridClipboardCopy,
  handleGridClipboardCut,
  handleGridClipboardPaste,
  clearClipboardMarquee,
  copySelectionFromMenu,
  pasteFromMenu,
} = useGridClipboard({
  app,
  activeWorksheet,
  projection,
  selectedCell,
  selectedRange,
  findCellCoordinates,
});

const formatDialog = useCellFormatDialog({
  app,
  activeWorksheet,
  selectedCell,
  selectedRange,
  allSelectedRanges,
  selectedCells,
  cellsForRange,
  boundsForRange,
});
const { openCellFormatDialog } = formatDialog;

const openDialog = useOpenDialog({
  app,
  onError: (error) => showAppError(error),
});
const { openFileDialog } = openDialog;

const propertiesDialog = useDocumentPropertiesDialog({ app });
const { openPropertiesDialog } = propertiesDialog;

const worksheetDialogs = useWorksheetDialogs({ app, activeWorksheetId });
const { addWorksheet, renameWorksheet, deleteWorksheet } = worksheetDialogs;

const errorDialog = useErrorDialog({
  lastErrorMessage: app.lastErrorMessage,
  clearLastError: app.clearLastError,
});

/** Document subject with a sensible default for the toolbar title button. */
const documentTitle = computed(() => app.activeSubject.value || "Untitled spreadsheet");

/**
 * True when at least one formula editor (bar or in-cell) is open with a
 * formula draft. Drives the discreet "Press Ctrl+Space for function help"
 * hint in the status line.
 */
const showFormulaAssistHint = computed(() =>
  (formulaEditing.value && formulaDraft.value.trim().startsWith("="))
  || (inlineCellEditing.value && inlineCellDraft.value.trim().startsWith("=")));

/**
 * Status line text. Appends the assist hint while the user is editing a
 * formula so the keyboard shortcut stays discoverable without rendering
 * a full popover by default.
 */
const statusLineText = computed(() => showFormulaAssistHint.value
  ? `${app.status.value} \u00B7 Press Ctrl+Space for function help`
  : app.status.value);

/**
 * Right-hand toolbar badge that summarizes the current document mode:
 * read-only revision, active time-travel cursor, or live-edit + dirty
 * state.
 */
const statusBadgeLabel = computed(() => {
  if (app.isViewingHistorical.value) {
    return "Historical \u00B7 read-only";
  }
  if (app.isTimeTravelActive.value) {
    return `Time travel \u00B7 ${app.timeTravelDateLabel.value}`;
  }
  return `Current \u00B7 ${app.isDirty.value ? "Unsaved" : "Saved"}`;
});

// Keep the explicit `activeWorksheetId` ref in sync with whatever
// `activeWorksheet` resolved to. This matters when the previously active
// worksheet was deleted and we fell back to the first visible one.
watch(
  () => activeWorksheet.value?.id,
  (worksheetId) => {
    activeWorksheetId.value = worksheetId ?? null;
  },
);

/**
 * PrimeVue Menubar model.
 *
 * Built from `computed` rather than declared as a constant so menu item
 * `disabled` flags and `command` closures stay reactive to the document
 * lifecycle (read-only mode, selection, dirty state, etc.).
 */
const menuItems = computed<MenuItem[]>(() => [
  {
    label: "File",
    items: [
      { label: "New", icon: "pi pi-file-plus", disabled: !app.canCreate.value, command: () => void app.createNewDocument() },
      { label: "Open", icon: "pi pi-folder-open", command: () => void openFileDialog() },
      { separator: true },
      { label: "Save", icon: "pi pi-save", disabled: !app.canSave.value || saveInFlight.value, command: () => void saveCurrentDocument() },
      { label: "Import XLSX...", icon: "pi pi-upload", disabled: !app.canCreate.value, command: () => xlsxImportInput.value?.click() },
      { label: "Export XLSX", icon: "pi pi-download", disabled: !app.activeGrid.value, command: () => void exportCurrentWorkbook() },
      { label: "Delete", icon: "pi pi-trash", disabled: !app.canDelete.value, command: () => { deleteDialogVisible.value = true; } },
    ],
  },
  {
    label: "Edit",
    items: [
      { label: "Copy", icon: "pi pi-copy", shortcut: "\u2318C", disabled: !hasSelection.value, command: () => void copySelectionFromMenu("copy") },
      { label: "Cut", icon: "pi pi-file-export", shortcut: "\u2318X", disabled: app.gridReadOnly.value || !hasSelection.value, command: () => void copySelectionFromMenu("cut") },
      { label: "Paste", icon: "pi pi-clipboard", shortcut: "\u2318V", disabled: app.gridReadOnly.value || (!internalClipboard.value && !navigator.clipboard), command: () => void pasteFromMenu() },
      { separator: true },
      { label: "Insert row above", icon: "pi pi-arrow-up", disabled: app.gridReadOnly.value || !selectedCell.value, command: () => insertRow("before") },
      { label: "Insert row below", icon: "pi pi-arrow-down", disabled: app.gridReadOnly.value || !selectedCell.value, command: () => insertRow("after") },
      { label: "Insert column left", icon: "pi pi-arrow-left", disabled: app.gridReadOnly.value || !selectedCell.value, command: () => insertColumn("before") },
      { label: "Insert column right", icon: "pi pi-arrow-right", disabled: app.gridReadOnly.value || !selectedCell.value, command: () => insertColumn("after") },
      { separator: true },
      { label: "Delete row", icon: "pi pi-minus", disabled: app.gridReadOnly.value || !selectedCell.value, command: deleteSelectedRow },
      { label: "Delete column", icon: "pi pi-minus", disabled: app.gridReadOnly.value || !selectedCell.value, command: deleteSelectedColumn },
    ],
  },
  {
    label: "Format",
    items: [
      { label: "Format cells...", icon: "pi pi-sliders-h", disabled: app.gridReadOnly.value || !hasSelection.value, command: () => openCellFormatDialog(selectedRange.value) },
    ],
  },
  {
    label: "History",
    items: [
      { label: "Browse revisions", icon: "pi pi-history", disabled: !app.currentCanBrowseHistory.value || !app.currentDocument.value, command: () => void openRevisionDialog() },
      { label: "Return to current", icon: "pi pi-refresh", disabled: !app.isViewingHistorical.value, command: app.returnToCurrent },
    ],
  },
  {
    label: "Window",
    items: [
      ...app.openSessions.value.map((session) => ({
        label: `${session.isActive ? "\u2713 " : ""}${session.title}${session.isDirty ? " *" : ""}`,
        icon: session.isActive ? "pi pi-check" : "pi pi-window-maximize",
        disabled: session.isActive,
        command: () => app.switchToOpenSession(session.id),
      })),
      ...(app.openSessions.value.length > 0 ? [{ separator: true } satisfies MenuItem] : []),
      {
        label: "Close current spreadsheet",
        icon: "pi pi-times",
        disabled: !app.currentDocument.value,
        command: () => app.activeSpreadsheetSessionId.value && app.closeOpenSession(app.activeSpreadsheetSessionId.value),
      },
    ],
  },
]);

const cellContextMenuItems = computed<MenuItem[]>(() => [
  {
    label: "Copy",
    icon: "pi pi-copy",
    disabled: !cellContextRange.value,
    command: () => void copySelectionFromMenu("copy", cellContextRange.value),
  },
  {
    label: "Cut",
    icon: "pi pi-file-export",
    disabled: app.gridReadOnly.value || !cellContextRange.value,
    command: () => void copySelectionFromMenu("cut", cellContextRange.value),
  },
  {
    label: "Paste",
    icon: "pi pi-clipboard",
    disabled: app.gridReadOnly.value || (!internalClipboard.value && !navigator.clipboard),
    command: () => void pasteFromMenu(),
  },
  { separator: true },
  {
    label: "Format cells...",
    icon: "pi pi-sliders-h",
    disabled: app.gridReadOnly.value || !cellContextRange.value,
    command: () => openCellFormatDialog(cellContextRange.value),
  },
  { separator: true },
  {
    label: "Insert row above",
    icon: "pi pi-arrow-up",
    disabled: app.gridReadOnly.value || !selectedCell.value,
    command: () => insertRow("before"),
  },
  {
    label: "Insert row below",
    icon: "pi pi-arrow-down",
    disabled: app.gridReadOnly.value || !selectedCell.value,
    command: () => insertRow("after"),
  },
  {
    label: "Insert column left",
    icon: "pi pi-arrow-left",
    disabled: app.gridReadOnly.value || !selectedCell.value,
    command: () => insertColumn("before"),
  },
  {
    label: "Insert column right",
    icon: "pi pi-arrow-right",
    disabled: app.gridReadOnly.value || !selectedCell.value,
    command: () => insertColumn("after"),
  },
]);

/** Export the current workbook as a local .xlsx download. */
async function exportCurrentWorkbook() {
  const grid = app.activeGrid.value;
  if (!grid) {
    return;
  }
  try {
    const buffer = await writeTeamGridExcelBuffer(grid);
    const filename = `${sanitizeDownloadFilename(documentTitle.value)}.xlsx`;
    downloadBlob(buffer, filename);
    app.status.value = `Exported ${filename}`;
  } catch (error) {
    showAppError(error);
  }
}

async function importXlsxFile(event: Event) {
  const input = event.target as HTMLInputElement;
  const file = input.files?.[0];
  input.value = "";
  if (!file) {
    return;
  }
  try {
    const title = file.name.replace(/\.xlsx$/i, "") || "Imported spreadsheet";
    const envelope = await importTeamGridWorkbookBuffer(await file.arrayBuffer(), title);
    await app.createDocumentFromEnvelope(envelope);
    app.status.value = `Imported ${file.name}.`;
  } catch (error) {
    showAppError(error);
  }
}

/** Normalize an unknown thrown value into a status-bar string. */
function readError(error: unknown) {
  return error instanceof Error ? error.message : String(error);
}

function showAppError(error: unknown) {
  const message = readError(error);
  app.status.value = message;
  app.lastErrorMessage.value = message;
}

/** Trigger a browser download for generated export content. */
function downloadBlob(buffer: ArrayBuffer, filename: string) {
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

function sanitizeDownloadFilename(filename: string) {
  const sanitized = filename
    .trim()
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "-")
    .replace(/\.+$/g, "")
    .slice(0, 120);
  return sanitized || "Untitled spreadsheet";
}

/**
 * Pre-fetch the revision history through the composable and show the
 * revisions dialog once the data is in hand, so users never see an empty
 * shell while the network request is in flight.
 */
async function openRevisionDialog() {
  await app.openRevisionDialog();
  revisionDialogVisible.value = true;
}

/** Close the revisions dialog and time-travel to the chosen revision. */
function handleRevisionSelect(revisionId: string) {
  revisionDialogVisible.value = false;
  void app.loadHistoricalRevision(revisionId);
}

/**
 * Handle a cell click from the grid.
 *
 * The branch matters: while the user is actively editing a formula that
 * starts with `=`, clicking another cell appends its address to the draft
 * (Excel-style formula picking). Otherwise it leaves formula-edit mode
 * and makes the clicked cell the new active cell.
 *
 * The selection range is **not** touched here. Every gesture path in
 * `useGridSelectionGestures` emits a `select-range` alongside its
 * `select`, so by the time we get here the range is already correct
 * (e.g. `A:F` for shift+click, `{F,F}` for a plain click, the same
 * multi-cell range from before when keyboard-edit starts). Clobbering
 * the range to `{cell, cell}` here would silently collapse a freshly
 * extended shift+click range.
 */
function selectCell(cell: Cell, address: string) {
  if (
    formulaEditing.value
    && !app.gridReadOnly.value
    && selectedCellId.value
    && selectedCellId.value !== cell.id
    && formulaDraft.value.trim().startsWith("=")
  ) {
    formulaDraft.value = appendPickedAddress(formulaDraft.value, address);
    return;
  }
  formulaEditing.value = false;
  formulaAssistOpen.value = false;
  selectedCellId.value = cell.id;
  selectedCellAddress.value = address;
}

/**
 * Update the multi-cell selection.
 *
 * We deliberately ignore range changes while the user is editing a
 * formula so that drag selecting inside the grid does not accidentally
 * clobber the picked-references list in the formula bar.
 */
function selectRange(range: CellSelectionRange) {
  if (formulaEditing.value) {
    return;
  }
  selectedRange.value = range;
}

/**
 * Append a Ctrl/Meta+click sub-range to the disjoint extra-selection
 * list, keeping the primary `selectedRange` free for the freshly clicked
 * cell. Skipped during formula editing for the same reason as
 * {@link selectRange}.
 */
function addRange(range: CellSelectionRange) {
  if (formulaEditing.value) {
    return;
  }
  additionalRanges.value = [...additionalRanges.value, range];
}

/** Drop every Ctrl/Meta+click sub-range; the primary range stays untouched. */
function clearAdditionalRanges() {
  if (formulaEditing.value) {
    return;
  }
  if (additionalRanges.value.length > 0) {
    additionalRanges.value = [];
  }
}

/**
 * Atomically replace the disjoint extra-selection list. Used by
 * Ctrl/Meta+click deselection, which subtracts a cell from every
 * enclosing rectangle and produces a fresh fragmented sub-range list.
 */
function setAdditionalRanges(ranges: CellSelectionRange[]) {
  if (formulaEditing.value) {
    return;
  }
  additionalRanges.value = ranges;
}

/** Open the cell context menu, preserving an existing range when right-clicked inside it. */
function openCellContextMenu(payload: { event: MouseEvent; cell: Cell; address: string; range: CellSelectionRange }) {
  if (!selectedRange.value || selectedRange.value.startCellId !== payload.range.startCellId || selectedRange.value.endCellId !== payload.range.endCellId) {
    formulaEditing.value = false;
    formulaAssistOpen.value = false;
    selectedCellId.value = payload.cell.id;
    selectedCellAddress.value = payload.address;
    selectedRange.value = payload.range;
    additionalRanges.value = [];
  }
  cellContextRange.value = payload.range;
  cellContextMenu.value?.show(payload.event);
}

async function saveCurrentDocument() {
  if (saveInFlight.value) {
    return;
  }
  if (gridViewportComponent.value?.flushPendingEdit()) {
    await nextTick();
  }
  saveInFlight.value = true;
  try {
    await app.saveDocument();
  } finally {
    saveInFlight.value = false;
  }
}

/**
 * Write a cell value coming from either the formula bar or the in-grid
 * editor.
 *
 * If the input starts with `=` we parse it as a formula, evaluate it
 * against the current worksheet, and cache both the AST references and
 * the latest result on the cell so dependent recomputation can skip
 * re-parsing. Otherwise we coerce the input through
 * `coerceInputToCellValue` so a column-typed cell still keeps its
 * preferred shape.
 */
function commitCell(cell: Cell, rawValue: string) {
  if (!activeWorksheet.value || !projection.value) {
    return;
  }
  app.updateGrid((grid) => {
    const worksheet = grid.workbook.worksheetsById[activeWorksheet.value!.id];
    const targetCell: Cell = {
      ...cell,
      value: preserveCompatibleCellValueFormat(
        coerceInputToCellValue(rawValue, worksheet.columnsById[cell.columnId]?.defaultValueKind),
        cell.value,
      ),
      formula: undefined,
    };
    if (rawValue.trim().startsWith("=")) {
      const evaluated = evaluateFormula(rawValue, worksheet, projectWorksheet(worksheet));
      targetCell.formula = {
        kind: "formula",
        source: rawValue,
        references: evaluated.references,
        cached: evaluated.result,
        error: evaluated.result.kind === "error" ? evaluated.result.code : undefined,
      };
      targetCell.value = preserveCompatibleCellValueFormat(formulaResultToCellValue(evaluated.result), cell.value);
      formulaError.value = evaluated.errorMessage ?? null;
    } else {
      formulaError.value = null;
    }
    worksheet.cellsById[targetCell.id] = targetCell;
    return [{ type: "setCell", worksheetId: worksheet.id, cell: targetCell }];
  });
}

/**
 * Insert a new row immediately before or after the row containing the
 * current selection. Emits an `insertRow` operation so the granular patch
 * carries the inserted row id and target index.
 */
function insertRow(position: "before" | "after") {
  if (!activeWorksheet.value || !selectedCell.value) return;
  app.updateGrid((grid) => {
    const worksheet = grid.workbook.worksheetsById[activeWorksheet.value!.id];
    const rowId = createId("row");
    const selectedIndex = worksheet.rowOrder.indexOf(selectedCell.value!.rowId);
    const index = position === "before" ? selectedIndex : selectedIndex + 1;
    const row = { id: rowId };
    worksheet.rowsById[rowId] = row;
    worksheet.rowOrder.splice(index, 0, rowId);
    return [{ type: "insertRow", worksheetId: worksheet.id, rowId, row, index }];
  });
}

/**
 * Insert a new column immediately before or after the column containing
 * the current selection. Defaults the new column to the standard 120px
 * width used elsewhere in the app.
 */
function insertColumn(position: "before" | "after") {
  if (!activeWorksheet.value || !selectedCell.value) return;
  app.updateGrid((grid) => {
    const worksheet = grid.workbook.worksheetsById[activeWorksheet.value!.id];
    const columnId = createId("col");
    const selectedIndex = worksheet.columnOrder.indexOf(selectedCell.value!.columnId);
    const index = position === "before" ? selectedIndex : selectedIndex + 1;
    const column = { id: columnId, width: DEFAULT_COLUMN_WIDTH };
    worksheet.columnsById[columnId] = column;
    worksheet.columnOrder.splice(index, 0, columnId);
    return [{ type: "insertColumn", worksheetId: worksheet.id, columnId, column, index }];
  });
}

/**
 * Tombstone the row containing the current selection. Like worksheet
 * deletion, we keep the row entry so collaborators with concurrent edits
 * to the same row do not see their work disappear after a merge.
 */
function deleteSelectedRow() {
  if (!activeWorksheet.value || !selectedCell.value) return;
  app.updateGrid((grid) => {
    const worksheet = grid.workbook.worksheetsById[activeWorksheet.value!.id];
    const deletedAt = new Date().toISOString();
    worksheet.rowsById[selectedCell.value!.rowId].deletedAt = deletedAt;
    return [{ type: "tombstoneRow", worksheetId: worksheet.id, rowId: selectedCell.value!.rowId, deletedAt }];
  });
}

/** Column-side counterpart of {@link deleteSelectedRow}. */
function deleteSelectedColumn() {
  if (!activeWorksheet.value || !selectedCell.value) return;
  app.updateGrid((grid) => {
    const worksheet = grid.workbook.worksheetsById[activeWorksheet.value!.id];
    const deletedAt = new Date().toISOString();
    worksheet.columnsById[selectedCell.value!.columnId].deletedAt = deletedAt;
    return [{ type: "tombstoneColumn", worksheetId: worksheet.id, columnId: selectedCell.value!.columnId, deletedAt }];
  });
}

/** Commit a released header drag into the local dirty document. */
function resizeColumn(payload: { columnId: ColumnId; width: number }) {
  if (!activeWorksheet.value || app.gridReadOnly.value) return;
  app.updateGrid((grid) => {
    const worksheet = grid.workbook.worksheetsById[activeWorksheet.value!.id];
    const column = worksheet.columnsById[payload.columnId];
    if (!column || column.width === payload.width) {
      return [];
    }
    column.width = payload.width;
    return [{ type: "setColumnWidth", worksheetId: worksheet.id, columnId: payload.columnId, width: payload.width }];
  });
}

/** Commit a released row-header drag into the local dirty document. */
function resizeRow(payload: { rowId: RowId; height: number }) {
  if (!activeWorksheet.value || app.gridReadOnly.value) return;
  app.updateGrid((grid) => {
    const worksheet = grid.workbook.worksheetsById[activeWorksheet.value!.id];
    const row = worksheet.rowsById[payload.rowId];
    if (!row || row.height === payload.height) {
      return [];
    }
    row.height = payload.height;
    return [{ type: "setRowHeight", worksheetId: worksheet.id, rowId: payload.rowId, height: payload.height }];
  });
}

</script>

<template>
  <main class="teamgrid-shell">
    <input
      ref="xlsxImportInput"
      type="file"
      accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      hidden
      @change="importXlsxFile"
    >
    <header class="toolbar glass-card" :class="{ 'toolbar--ios-multitasking': app.hostUiPreferences.value.iosMultitaskingOptimized }">
      <div class="toolbar__leading">
        <span class="toolbar__title">TeamGrid</span>
        <Menubar :model="menuItems" class="toolbar__menubar" />
        <Button
          icon="pi pi-save"
          class="toolbar__mobile-save"
          text
          rounded
          severity="secondary"
          aria-label="Save spreadsheet"
          :disabled="!app.canSave.value || saveInFlight"
          @click="void saveCurrentDocument()"
        />
        <Button
          :icon="app.isViewingHistorical.value ? 'pi pi-history' : 'pi pi-refresh'"
          text
          rounded
          severity="secondary"
          :aria-label="app.isViewingHistorical.value ? 'Return to current version' : 'Refresh spreadsheet'"
          :disabled="!app.canRefresh.value || saveInFlight"
          @click="app.isViewingHistorical.value ? app.returnToCurrent() : app.refreshCurrentDocument()"
        />
      </div>
      <button
        v-if="app.currentEnvelope.value"
        class="toolbar__document-title"
        type="button"
        :disabled="app.gridReadOnly.value"
        @click="openPropertiesDialog"
      >
        {{ documentTitle }}
      </button>
      <button
        v-if="app.currentCanBrowseHistory.value && app.currentDocument.value"
        class="toolbar__status-badge"
        type="button"
        @click="openRevisionDialog"
      >
        {{ statusBadgeLabel }}
      </button>
      <span v-else class="toolbar__status-badge">{{ statusBadgeLabel }}</span>
    </header>

    <section class="workspace">
      <template v-if="app.activeGrid.value && activeWorksheet && projection">
        <div v-if="app.isTimeTravelActive.value" class="history-banner">
          <i class="pi pi-clock" aria-hidden="true" />
          <span>Time travel mode is active as of {{ app.timeTravelDateLabel.value }} - read-only.</span>
        </div>
        <div v-if="app.isViewingHistorical.value" class="history-banner">
          <i class="pi pi-history" aria-hidden="true" />
          <span>You're viewing a historical revision - read-only.</span>
          <button type="button" @click="app.returnToCurrent">Return to current</button>
        </div>

        <FormulaBar
          ref="formulaBarComponent"
          v-model="formulaDraft"
          :active-address="selectedCellAddress"
          :readonly="app.gridReadOnly.value || !selectedCell"
          :error-message="formulaError"
          @begin-edit="formulaEditing = true"
          @commit="commitFormulaBar"
          @cancel="cancelFormulaEdit"
          @request-help="openFormulaAssist('formulaBar', $event)"
        />
        <GridViewport
          ref="gridViewportComponent"
          :worksheet="activeWorksheet"
          :projection="projection"
          :selected-cell-id="selectedCellId"
          :selected-range="selectedRange"
          :additional-ranges="additionalRanges"
          :clipboard-range="clipboardSourceRange"
          :highlighted-cell-ids="highlightedCellIds"
          :readonly="app.gridReadOnly.value"
          :locale="app.activeGrid.value.settings.locale"
          @select="selectCell"
          @select-range="selectRange"
          @add-range="addRange"
          @clear-additional-ranges="clearAdditionalRanges"
          @set-additional-ranges="setAdditionalRanges"
          @commit="commitCell"
          @cell-context="openCellContextMenu"
          @request-help="openFormulaAssist('inlineCell', $event)"
          @edit-state="handleInlineEditState"
          @clipboard-copy="handleGridClipboardCopy"
          @clipboard-cut="handleGridClipboardCut"
          @clipboard-paste="handleGridClipboardPaste"
          @clipboard-clear="clearClipboardMarquee"
          @resize-column="resizeColumn"
          @resize-row="resizeRow"
        />
        <WorksheetTabs
          :grid="app.activeGrid.value"
          :active-worksheet-id="activeWorksheet.id"
          :readonly="app.gridReadOnly.value"
          @select="activeWorksheetId = $event"
          @add="addWorksheet"
          @rename="renameWorksheet"
          @delete="deleteWorksheet"
        />
        <ContextMenu ref="cellContextMenu" :model="cellContextMenuItems" />
      </template>
      <section v-else class="empty-state">
        <h1>Collaborative spreadsheets</h1>
        <p>Create a new spreadsheet or open an existing one. Edit cells, write formulas, organize your work across multiple sheet tabs, and see your teammates' changes in real time.</p>
        <div class="empty-state__actions">
          <Button label="New spreadsheet" icon="pi pi-file-plus" :disabled="!app.canCreate.value" @click="() => app.createNewDocument()" />
          <Button label="Open spreadsheet" icon="pi pi-folder-open" severity="secondary" @click="openFileDialog" />
        </div>
      </section>
    </section>

    <footer class="status-line">{{ statusLineText }}</footer>

    <CellFormatDialog :controller="formatDialog" :read-only="app.gridReadOnly.value" />

    <ErrorDialog
      :controller="errorDialog"
      :message="app.lastErrorMessage.value"
      :on-hide="app.clearLastError"
    />

    <FormulaAssistPanel
      v-model:visible="formulaAssistOpen"
      :draft="formulaAssistDraft"
      :caret-pos="formulaAssistCaretPos"
      :anchor-el="formulaAssistAnchor"
      :readonly="app.gridReadOnly.value"
      @select="handleFormulaAssistSelect"
      @dismiss="formulaAssistOpen = false"
    />

    <OpenSpreadsheetDialog
      :controller="openDialog"
      :selected-database-id="app.selectedDatabaseId"
      :readable-databases="app.readableDatabases"
    />

    <DocumentPropertiesDialog
      :controller="propertiesDialog"
      :read-only="app.gridReadOnly.value"
    />

    <DeleteSpreadsheetDialog
      v-model:visible="deleteDialogVisible"
      @confirm="app.deleteCurrentDocument"
    />

    <RenameWorksheetDialog :controller="worksheetDialogs" />

    <DocumentRevisionDialog
      v-model:visible="revisionDialogVisible"
      :entries="app.revisionEntries.value"
      :loading="app.revisionLoading.value"
      :error-message="app.revisionErrorMessage.value"
      :current-revision-id="app.currentRevisionId.value"
      @select="handleRevisionSelect"
      @cancel="revisionDialogVisible = false"
    />
  </main>
</template>
