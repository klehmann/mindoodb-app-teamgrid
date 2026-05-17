<script setup lang="ts">
/**
 * Root component for the Teamgrid sample app.
 *
 * Architecture notes
 * ------------------
 *
 * Composables vs. local state. The document lifecycle (connection, open,
 * save, history, etc.) lives in `useTeamGridDocument`. The component-level
 * refs declared here are scoped to UI concerns the composable should not
 * know about: which dialog is currently open, which cell is selected, the
 * formula-bar draft string, and so on.
 *
 * Selection model. A single `selectedCellId` drives "current cell" and an
 * optional `selectedRange` drives multi-cell selections (e.g. for the
 * "format selection" actions). The `GridViewport` component owns the
 * pointer/keyboard interactions and emits semantic events back here.
 *
 * Formula editing state machine. When the user starts editing the formula
 * bar (`formulaEditing` flips to `true`) we enter "formula picking" mode:
 * clicking grid cells appends their `A1` address to the draft instead of
 * changing selection. Commit / cancel both clear `formulaEditing` and run
 * the appropriate `updateGrid` mutation through `useTeamGridDocument`.
 *
 * Open dialog & view navigator lifecycle. The File / Open dialog is backed
 * by a dynamic view navigator (see `viewOpen.ts`). `rebuildOpenNavigator`
 * disposes any previous navigator before creating a new one, so we never
 * leak navigator sessions when the user switches databases or reopens the
 * dialog. The dialog's `@hide` handler also calls `disposeOpenNavigator`,
 * so an Escape or backdrop dismissal does not leave a navigator running.
 *
 * Properties dialog. Title and tag edits are batched into a single
 * `setDocumentProperties` operation. We never apply title or tag changes
 * straight to the envelope; they always travel through `updateGrid` so
 * they participate in the granular-save / `baseHeads` machinery.
 */
import { computed, nextTick, ref, watch } from "vue";
import Button from "primevue/button";
import ContextMenu from "primevue/contextmenu";
import Dialog from "primevue/dialog";
import Menubar from "primevue/menubar";
import type { MenuItem } from "primevue/menuitem";
import type { MindooDBAppViewNavigator } from "mindoodb-app-sdk";

import DocumentRevisionDialog from "@/components/DocumentRevisionDialog.vue";
import FormulaAssistPanel from "@/components/FormulaAssistPanel.vue";
import FormulaBar from "@/components/FormulaBar.vue";
import GridViewport from "@/components/GridViewport.vue";
import TagTreeList from "@/components/TagTreeList.vue";
import WorksheetTabs from "@/components/WorksheetTabs.vue";
import { useTeamGridDocument, readDocumentSummaryLabel } from "@/composables/useTeamGridDocument";
import { applyCellFormat, coerceInputToCellValue, formatCellValue, formulaResultToCellValue, preserveCompatibleCellValueFormat, type CellFormatKind, type CellFormatRequest } from "@/lib/cellFormatting";
import {
  decodePayload,
  rewriteFormulaSource,
  serializeRange,
  type ClipboardPayload,
  type ClipboardRange,
  type SerializedClipboardPayload,
} from "@/lib/clipboard";
import { evaluateFormula, parseFormula, type FunctionDefinition } from "@/lib/formulas";
import { DEFAULT_COLUMN_WIDTH } from "@/lib/gridDimensions";
import { createCellId, createId, getFirstVisibleWorksheet, normalizeTags, type Cell, type CellStyle, type ColumnId, type CurrencyCode, type RowId, type WorksheetId } from "@/lib/teamgridDocument";
import { getCell, projectWorksheet } from "@/lib/gridProjection";
import type { TeamGridOperation } from "@/lib/teamgridOps";
import { importTeamGridWorkbookBuffer } from "@/lib/xlsx/importWorkbook";
import { writeTeamGridExcelBuffer } from "@/lib/xlsx/exportWorkbook";
import {
  ALL_SPREADSHEETS_NODE_KEY,
  buildOpenCategoryTree,
  collectNavigatorEntries,
  createOpenViewDefinition,
  dedupeDocumentEntries,
  mapDocumentEntries,
  type OpenCategoryNode,
  type OpenDocumentRow,
} from "@/lib/viewOpen";

/** Inclusive rectangular cell range used for multi-cell selection state. */
interface CellSelectionRange {
  startCellId: string;
  endCellId: string;
}

/**
 * Identifies which editor the formula assist panel is anchored against, so
 * `handleFormulaAssistSelect` can route the chosen function back into the
 * right input element.
 */
type FormulaAssistEditor = "formulaBar" | "inlineCell";

/**
 * Payload emitted by either editor when the user requests formula assist
 * (Ctrl+Space or the `fx` button). The panel uses `anchorEl` to compute its
 * floating coordinates and `draft` / `caretPos` to derive context-aware
 * suggestions.
 */
interface FormulaAssistRequest {
  anchorEl: HTMLElement;
  draft: string;
  caretPos: number;
}

const app = useTeamGridDocument();
const cellContextMenu = ref<InstanceType<typeof ContextMenu> | null>(null);
const formulaBarComponent = ref<InstanceType<typeof FormulaBar> | null>(null);
const gridViewportComponent = ref<InstanceType<typeof GridViewport> | null>(null);
const xlsxImportInput = ref<HTMLInputElement | null>(null);

const openDialogVisible = ref(false);
const propertiesDialogVisible = ref(false);
const deleteDialogVisible = ref(false);
const revisionDialogVisible = ref(false);
const renameDialogVisible = ref(false);
const errorDialogVisible = ref(false);
const formatDialogVisible = ref(false);
const renameTargetId = ref<WorksheetId | null>(null);
const renameDraft = ref("");
const formatDialogKind = ref<CellFormatKind>("general");
const formatDialogCurrency = ref<CurrencyCode>("USD");
const formatDialogCustomNumFmt = ref("");
const selectedOpenDocId = ref("");
const selectedOpenCategoryKey = ref(ALL_SPREADSHEETS_NODE_KEY);
const openCategoryNodes = ref<OpenCategoryNode[]>([]);
const openDialogDocuments = ref<OpenDocumentRow[]>([]);
const allOpenDialogDocuments = ref<OpenDocumentRow[]>([]);
const openNavigator = ref<MindooDBAppViewNavigator | null>(null);
const activeWorksheetId = ref<WorksheetId | null>(null);
const selectedCellId = ref<string | null>(null);
const selectedRange = ref<CellSelectionRange | null>(null);
const selectedCellAddress = ref("");
const cellContextRange = ref<CellSelectionRange | null>(null);
const formulaDraft = ref("");
const formulaError = ref<string | null>(null);
const formulaEditing = ref(false);
const inlineCellEditing = ref(false);
const inlineCellDraft = ref("");
const formulaAssistOpen = ref(false);
const formulaAssistEditor = ref<FormulaAssistEditor>("formulaBar");
const formulaAssistAnchor = ref<HTMLElement | null>(null);
const formulaAssistDraft = ref("");
const formulaAssistCaretPos = ref(0);
const saveInFlight = ref(false);
const clipboardSourceRange = ref<ClipboardRange | null>(null);
const internalClipboard = ref<SerializedClipboardPayload | null>(null);
const propertiesTitleDraft = ref("");
const propertiesTagsDraft = ref("");

/**
 * The worksheet currently displayed in the grid.
 *
 * Defaults to whichever worksheet the user explicitly picked via the tab
 * strip, falling back to the first non-tombstoned worksheet so deleting the
 * active tab does not leave the UI without a worksheet to render.
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
 * Render-time projection of {@link activeWorksheet} (ordered rows/columns and
 * A1 address lookups). Recomputed automatically when the worksheet changes
 * so consumers never have to track invalidation themselves.
 */
const projection = computed(() => activeWorksheet.value ? projectWorksheet(activeWorksheet.value) : null);

/**
 * Persisted cell record for {@link selectedCellId}.
 *
 * Returns `null` when nothing is selected or when the worksheet is not yet
 * loaded. The lookup walks the projection so we get a fresh `Cell` (synthetic
 * empty if the cell has never been persisted) rather than a stale reference.
 */
const selectedCell = computed(() => {
  if (!activeWorksheet.value || !projection.value || !selectedCellId.value) {
    return null;
  }
  for (const row of projection.value.rows) {
    for (const column of projection.value.columns) {
      const cell = getCell(activeWorksheet.value, row.id, column.id);
      if (cell.id === selectedCellId.value) {
        return cell;
      }
    }
  }
  return activeWorksheet.value.cellsById[selectedCellId.value] ?? null;
});

/**
 * All cells covered by the current {@link selectedRange}, or just the single
 * {@link selectedCell} when no range is active. Used by every action that
 * applies to "everything the user picked", e.g. the format toolbar.
 */
const selectedCells = computed(() => {
  if (!activeWorksheet.value || !projection.value || !selectedRange.value) {
    return selectedCell.value ? [selectedCell.value] : [];
  }
  const start = findCellCoordinates(selectedRange.value.startCellId);
  const end = findCellCoordinates(selectedRange.value.endCellId);
  if (!start || !end) {
    return selectedCell.value ? [selectedCell.value] : [];
  }
  const cells: Cell[] = [];
  for (let rowIndex = Math.min(start.rowIndex, end.rowIndex); rowIndex <= Math.max(start.rowIndex, end.rowIndex); rowIndex += 1) {
    for (let columnIndex = Math.min(start.columnIndex, end.columnIndex); columnIndex <= Math.max(start.columnIndex, end.columnIndex); columnIndex += 1) {
      const row = projection.value.rows[rowIndex];
      const column = projection.value.columns[columnIndex];
      if (row && column) {
        cells.push(getCell(activeWorksheet.value, row.id, column.id));
      }
    }
  }
  return cells;
});

/** Convenience flag used to gate selection-dependent menu commands. */
const hasSelection = computed(() => selectedCells.value.length > 0);

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
  ? `${app.status.value} · Press Ctrl+Space for function help`
  : app.status.value);

/**
 * Right-hand toolbar badge that summarizes the current document mode:
 * read-only revision, active time-travel cursor, or live-edit + dirty state.
 */
const statusBadgeLabel = computed(() => {
  if (app.isViewingHistorical.value) {
    return "Historical · read-only";
  }
  if (app.isTimeTravelActive.value) {
    return `Time travel · ${app.timeTravelDateLabel.value}`;
  }
  return `Current · ${app.isDirty.value ? "Unsaved" : "Saved"}`;
});

watch(
  () => app.lastErrorMessage.value,
  (message) => {
    if (message) {
      errorDialogVisible.value = true;
    }
  },
);

/**
 * Cell ids that the grid should overlay while the user is composing a
 * formula in the formula bar. Each parsed reference contributes one or more
 * cells so the user gets instant visual feedback for formula targets.
 */
const highlightedCellIds = computed(() => {
  if (!activeWorksheet.value || !projection.value || !formulaDraft.value.trim().startsWith("=")) {
    return [];
  }
  const parsed = parseFormula(formulaDraft.value, activeWorksheet.value.id, projection.value);
  if ("code" in parsed) {
    return [];
  }
  return parsed.references.flatMap((reference) => {
    if (reference.kind === "cell") {
      return [createCellId(reference.rowId, reference.columnId)];
    }
    if (reference.kind === "range") {
      const startRow = projection.value?.rowIndexById.get(reference.startRowId);
      const endRow = projection.value?.rowIndexById.get(reference.endRowId);
      const startColumn = projection.value?.columnIndexById.get(reference.startColumnId);
      const endColumn = projection.value?.columnIndexById.get(reference.endColumnId);
      if (startRow == null || endRow == null || startColumn == null || endColumn == null || !projection.value) {
        return [];
      }
      const cellIds: string[] = [];
      for (let rowIndex = Math.min(startRow, endRow); rowIndex <= Math.max(startRow, endRow); rowIndex += 1) {
        for (let columnIndex = Math.min(startColumn, endColumn); columnIndex <= Math.max(startColumn, endColumn); columnIndex += 1) {
          const row = projection.value.rows[rowIndex];
          const column = projection.value.columns[columnIndex];
          if (row && column) {
            cellIds.push(createCellId(row.id, column.id));
          }
        }
      }
      return cellIds;
    }
    return [];
  });
});

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
      { label: "Copy", icon: "pi pi-copy", shortcut: "⌘C", disabled: !hasSelection.value, command: () => void copySelectionFromMenu("copy") },
      { label: "Cut", icon: "pi pi-file-export", shortcut: "⌘X", disabled: app.gridReadOnly.value || !hasSelection.value, command: () => void copySelectionFromMenu("cut") },
      { label: "Paste", icon: "pi pi-clipboard", shortcut: "⌘V", disabled: app.gridReadOnly.value || (!internalClipboard.value && !navigator.clipboard), command: () => void pasteFromMenu() },
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
      { label: "Bold", icon: "pi pi-bold", disabled: app.gridReadOnly.value || !hasSelection.value, command: () => patchSelectedStyle({ bold: !selectedCell.value?.style?.bold }) },
      { label: "Italic", icon: "pi pi-italic", disabled: app.gridReadOnly.value || !hasSelection.value, command: () => patchSelectedStyle({ italic: !selectedCell.value?.style?.italic }) },
      { label: "Underline", icon: "pi pi-underline", disabled: app.gridReadOnly.value || !hasSelection.value, command: () => patchSelectedStyle({ underline: !selectedCell.value?.style?.underline }) },
      { separator: true },
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
        label: `${session.isActive ? "✓ " : ""}${session.title}${session.isDirty ? " *" : ""}`,
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

// Keep the explicit `activeWorksheetId` ref in sync with whatever
// `activeWorksheet` resolved to. This matters when the previously active
// worksheet was deleted and we fell back to the first visible one.
watch(
  () => activeWorksheet.value?.id,
  (worksheetId) => {
    activeWorksheetId.value = worksheetId ?? null;
  },
);

// Mirror the selected cell into the formula bar draft. Formula cells expose
// their `source` (e.g. `=SUM(A1:A10)`); plain cells get their formatted
// display value so the user can edit it as text.
watch(
  selectedCell,
  (cell) => {
    if (!cell) {
      formulaDraft.value = "";
      return;
    }
    formulaDraft.value = cell.formula?.source ?? formatCellValue(cell.value, app.activeGrid.value?.settings.locale);
  },
);

// Guarantee a valid selection whenever the worksheet or its projection
// changes: clear the selection when there is no worksheet, keep the current
// selection when it still maps to a visible cell, otherwise snap to the
// top-left cell of the new projection.
watch(
  [activeWorksheet, projection],
  () => {
    if (!activeWorksheet.value || !projection.value) {
      selectedCellId.value = null;
      selectedRange.value = null;
      selectedCellAddress.value = "";
      return;
    }
    if (selectedCellId.value && findCellCoordinates(selectedCellId.value)) {
      return;
    }
    const firstRow = projection.value.rows[0];
    const firstColumn = projection.value.columns[0];
    if (!firstRow || !firstColumn) {
      selectedCellId.value = null;
      selectedRange.value = null;
      selectedCellAddress.value = "";
      return;
    }
    const cell = getCell(activeWorksheet.value, firstRow.id, firstColumn.id);
    selectedCellId.value = cell.id;
    selectedRange.value = { startCellId: cell.id, endCellId: cell.id };
    selectedCellAddress.value = projection.value.cellAddressById.get(cell.id) ?? "";
  },
);

// Reset the Properties dialog draft whenever the active document changes,
// but never clobber values while the dialog is open mid-edit.
watch(
  () => app.currentEnvelope.value,
  () => {
    if (!propertiesDialogVisible.value) {
      resetPropertiesDraft();
    }
  },
  { immediate: true },
);

/** Open the File/Open dialog after building a fresh view navigator. */
async function openFileDialog() {
  try {
    await rebuildOpenNavigator();
    openDialogVisible.value = true;
  } catch (error) {
    showAppError(error);
  }
}

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
    showAppError(error);
  }
}

/** User picked a category node; refresh the document list for that subtree. */
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

/**
 * Construct a fresh view navigator for the currently selected database.
 *
 * Disposes any existing navigator first so we never leak server-side
 * sessions when the user reopens the dialog. The navigator is built with
 * `category_then_document` ordering so each category entry appears once
 * with its full descendant document count, while documents fan out under
 * every matching category.
 *
 * After the navigator is materialized we collect every entry, dedupe the
 * document fan-out for the "All spreadsheets" view, and build the
 * category tree for the left pane.
 */
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

/**
 * Re-populate the document pane from the navigator using the currently
 * selected category. The synthetic "All spreadsheets" node returns the
 * deduped list from the initial collection pass instead of asking the
 * navigator for it again.
 */
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

/** Release the dynamic navigator and clear local references. */
async function disposeOpenNavigator() {
  await openNavigator.value?.dispose();
  openNavigator.value = null;
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

function dismissErrorDialog() {
  errorDialogVisible.value = false;
  app.clearLastError();
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
 * (Excel-style formula picking). Otherwise it leaves formula-edit mode and
 * makes the clicked cell the new selection.
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
  selectedRange.value = { startCellId: cell.id, endCellId: cell.id };
  selectedCellAddress.value = address;
}

/**
 * Update the multi-cell selection.
 *
 * We deliberately ignore range changes while the user is editing a formula
 * so that drag selecting inside the grid does not accidentally clobber the
 * picked-references list in the formula bar.
 */
function selectRange(range: CellSelectionRange) {
  if (formulaEditing.value) {
    return;
  }
  selectedRange.value = range;
}

/** Open the cell context menu, preserving an existing range when right-clicked inside it. */
function openCellContextMenu(payload: { event: MouseEvent; cell: Cell; address: string; range: CellSelectionRange }) {
  if (!selectedRange.value || selectedRange.value.startCellId !== payload.range.startCellId || selectedRange.value.endCellId !== payload.range.endCellId) {
    formulaEditing.value = false;
    formulaAssistOpen.value = false;
    selectedCellId.value = payload.cell.id;
    selectedCellAddress.value = payload.address;
    selectedRange.value = payload.range;
  }
  cellContextRange.value = payload.range;
  cellContextMenu.value?.show(payload.event);
}

/** Open the Excel-like value-format dialog for the active selection. */
function openCellFormatDialog(range: CellSelectionRange | null) {
  const targetRange = range ?? (selectedCell.value ? { startCellId: selectedCell.value.id, endCellId: selectedCell.value.id } : null);
  if (app.gridReadOnly.value || !targetRange) {
    return;
  }
  selectedRange.value = targetRange;
  seedFormatDialogFromCell(selectedCell.value);
  formatDialogVisible.value = true;
}

function seedFormatDialogFromCell(cell: Cell | null) {
  const value = cell?.value;
  const excelNumFmt = value && "excelNumFmt" in value ? value.excelNumFmt : undefined;
  formatDialogKind.value = excelNumFmt && isCustomExcelNumFmt(excelNumFmt)
    ? "custom"
    : value?.kind === "string"
      ? "text"
      : value?.kind === "number"
        ? (value.format ?? "general")
        : "general";
  formatDialogCurrency.value = value?.kind === "number" && value.currencyCode ? value.currencyCode : "USD";
  formatDialogCustomNumFmt.value = value?.kind === "number" || value?.kind === "date" || value?.kind === "string"
    ? excelNumFmt ?? ""
    : "";
}

function isCustomExcelNumFmt(numFmt: string) {
  return !new Set(["@", "0", "0.00", "0.00%", "$#,##0.00", "€#,##0.00"]).has(numFmt);
}

function currentCellFormatRequest(): CellFormatRequest {
  if (formatDialogKind.value === "currency") {
    return { kind: "currency", currencyCode: formatDialogCurrency.value };
  }
  if (formatDialogKind.value === "custom") {
    return { kind: "custom", excelNumFmt: formatDialogCustomNumFmt.value };
  }
  return { kind: formatDialogKind.value };
}

/** Apply the chosen value format to every selected cell through granular cell patches. */
function applySelectedCellFormat() {
  if (!activeWorksheet.value || app.gridReadOnly.value || selectedCells.value.length === 0) {
    formatDialogVisible.value = false;
    return;
  }
  const request = currentCellFormatRequest();
  const cellsToFormat = selectedCells.value;
  const locale = app.activeGrid.value?.settings.locale ?? "en-US";
  app.updateGrid((grid) => {
    const worksheet = grid.workbook.worksheetsById[activeWorksheet.value!.id];
    const operations: TeamGridOperation[] = [];
    for (const cell of cellsToFormat) {
      const existing = worksheet.cellsById[cell.id] ?? cell;
      const formatted = applyCellFormat(existing, request, locale);
      if (formatted === existing) {
        continue;
      }
      worksheet.cellsById[formatted.id] = formatted;
      operations.push({ type: "setCell", worksheetId: worksheet.id, cell: formatted });
    }
    return operations;
  });
  formatDialogVisible.value = false;
}

/** Grid-emitted `copy` event: serialize the selection into the clipboard. */
function handleGridClipboardCopy(payload: { range: CellSelectionRange | null; event: ClipboardEvent }) {
  writeSelectionToClipboard(payload.range, payload.event, "copy");
}

/**
 * Grid-emitted `cut` event: same as copy, but mark the cells as "cut" so the
 * paste handler knows to clear the source range and rewrite incoming
 * formulas using move-tracking deltas.
 */
function handleGridClipboardCut(payload: { range: CellSelectionRange | null; event: ClipboardEvent }) {
  writeSelectionToClipboard(payload.range, payload.event, "cut");
}

/**
 * Grid-emitted `paste` event.
 *
 * Tries the OS clipboard first (so Excel-originated HTML wins over an older
 * internal copy), then falls back to the in-memory clipboard for sandboxed
 * hosts that block `navigator.clipboard`.
 */
function handleGridClipboardPaste(payload: { event: ClipboardEvent }) {
  const clipboardPayload = readClipboardPayload(payload.event) ?? internalClipboard.value?.payload ?? null;
  const anchor = selectedCell.value ? findCellCoordinates(selectedCell.value.id) : null;
  if (!clipboardPayload || !anchor) {
    return;
  }
  applyPasteAtAnchor(clipboardPayload, { row: anchor.rowIndex, col: anchor.columnIndex });
}

/** Stop drawing the "marching ants" marquee, e.g. on Escape or paste. */
function clearClipboardMarquee() {
  clipboardSourceRange.value = null;
}

/**
 * Serialize the current selection into the native `ClipboardEvent` plus the
 * in-memory fallback. We always write both `text/html` (rich Teamgrid +
 * Excel-compatible payload) and `text/plain` (TSV) so paste targets across
 * the ecosystem get the richest data they can consume.
 */
function writeSelectionToClipboard(range: CellSelectionRange | null, event: ClipboardEvent, mode: "copy" | "cut") {
  const serialized = serializeSelection(range, mode);
  if (!serialized) {
    return;
  }
  event.clipboardData?.setData("text/html", serialized.html);
  event.clipboardData?.setData("text/plain", serialized.tsv);
  internalClipboard.value = serialized;
  clipboardSourceRange.value = serialized.payload.source.worksheetId ? {
    worksheetId: serialized.payload.source.worksheetId,
    startRow: serialized.payload.source.anchor.row,
    startCol: serialized.payload.source.anchor.col,
    endRow: serialized.payload.source.anchor.row + serialized.payload.source.rows - 1,
    endCol: serialized.payload.source.anchor.col + serialized.payload.source.cols - 1,
  } : null;
}

/** Decode a paste-event clipboard, preferring Teamgrid JSON over Excel HTML over TSV. */
function readClipboardPayload(event: ClipboardEvent) {
  const html = event.clipboardData?.getData("text/html") ?? "";
  const tsv = event.clipboardData?.getData("text/plain") ?? "";
  return decodePayload(html, tsv);
}

/**
 * Convert the given (or active) selection into the serialized clipboard
 * payload used by both the native event and the in-memory fallback.
 * Returns `null` when there is no worksheet or no usable range.
 */
function serializeSelection(range: CellSelectionRange | null, mode: "copy" | "cut") {
  const clipboardRange = selectionToClipboardRange(range);
  if (!clipboardRange || !activeWorksheet.value || !projection.value) {
    return null;
  }
  return serializeRange(
    clipboardRange,
    (rowIndex, columnIndex) => {
      const row = projection.value!.rows[rowIndex];
      const column = projection.value!.columns[columnIndex];
      return getCell(activeWorksheet.value!, row.id, column.id);
    },
    mode,
  );
}

/**
 * Edit menu entry point for Copy / Cut.
 *
 * The clipboard `copy`/`cut` events only fire for keyboard shortcuts and the
 * browser's own menu, so this path uses the async `navigator.clipboard`
 * write API. Sandboxed Haven hosts often reject this API, so we always
 * populate the in-memory clipboard as a fallback before attempting the
 * async write.
 */
async function copySelectionFromMenu(mode: "copy" | "cut", range = selectedRange.value) {
  const serialized = serializeSelection(range, mode);
  if (!serialized) {
    return;
  }
  internalClipboard.value = serialized;
  clipboardSourceRange.value = payloadSourceRange(serialized.payload);
  try {
    if (typeof ClipboardItem !== "undefined" && navigator.clipboard.write) {
      await navigator.clipboard.write([new ClipboardItem({
        "text/html": new Blob([serialized.html], { type: "text/html" }),
        "text/plain": new Blob([serialized.tsv], { type: "text/plain" }),
      })]);
      return;
    }
    await navigator.clipboard.writeText(serialized.tsv);
  } catch {
    // Sandboxed hosts often block navigator.clipboard; internalClipboard keeps the menu useful.
  }
}

/**
 * Edit menu entry point for Paste.
 *
 * Reads the OS clipboard via the async API (richer than the synchronous
 * `ClipboardEvent` path because we can pull both `text/html` and
 * `text/plain`), falling back to the in-memory clipboard when the host
 * blocks the API.
 */
async function pasteFromMenu() {
  if (app.gridReadOnly.value) {
    return;
  }
  const anchor = selectedCell.value ? findCellCoordinates(selectedCell.value.id) : null;
  if (!anchor) {
    return;
  }
  const payload = await readNavigatorClipboardPayload() ?? internalClipboard.value?.payload ?? null;
  if (!payload) {
    return;
  }
  applyPasteAtAnchor(payload, { row: anchor.rowIndex, col: anchor.columnIndex });
}

/**
 * Read the OS clipboard via `navigator.clipboard.read()` (preferred because
 * it exposes HTML data) and decode the first item that yields a usable
 * Teamgrid/Excel/TSV payload. Returns `null` on permission errors or empty
 * clipboards.
 */
async function readNavigatorClipboardPayload() {
  try {
    if (navigator.clipboard.read) {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        const html = item.types.includes("text/html")
          ? await (await item.getType("text/html")).text()
          : "";
        const text = item.types.includes("text/plain")
          ? await (await item.getType("text/plain")).text()
          : "";
        const payload = decodePayload(html, text);
        if (payload) {
          return payload;
        }
      }
    }
    const text = await navigator.clipboard.readText();
    return decodePayload("", text);
  } catch {
    return null;
  }
}

/**
 * Project a `CellSelectionRange` (selection by stable cell ids) onto a
 * worksheet-relative {@link ClipboardRange} (selection by row/column index).
 * Returns `null` when there is nothing to copy, e.g. before the worksheet
 * is loaded.
 */
function selectionToClipboardRange(range: CellSelectionRange | null): ClipboardRange | null {
  if (!activeWorksheet.value || !projection.value) {
    return null;
  }
  const activeRange = range ?? (selectedCell.value ? { startCellId: selectedCell.value.id, endCellId: selectedCell.value.id } : null);
  if (!activeRange) {
    return null;
  }
  const start = findCellCoordinates(activeRange.startCellId);
  const end = findCellCoordinates(activeRange.endCellId);
  if (!start || !end) {
    return null;
  }
  return {
    worksheetId: activeWorksheet.value.id,
    startRow: Math.min(start.rowIndex, end.rowIndex),
    startCol: Math.min(start.columnIndex, end.columnIndex),
    endRow: Math.max(start.rowIndex, end.rowIndex),
    endCol: Math.max(start.columnIndex, end.columnIndex),
  };
}

/**
 * Apply a decoded clipboard payload starting at the given anchor cell.
 *
 * Performs the heavy lifting of the paste pipeline:
 * 1. Auto-extends the worksheet so the entire payload fits.
 * 2. Writes each clipboard cell into its target position, rewriting any
 *    formula references by the copy delta so relative references shift
 *    Excel-style.
 * 3. For cut payloads, clears the source cells and rewrites formulas
 *    elsewhere in the sheet that pointed at the cut range (move tracking).
 * 4. Bundles every change into one `updateGrid` mutation so the granular
 *    save / `baseHeads` machinery treats the paste as a single edit.
 *
 * A cut that resolves back to its own source range is a no-op apart from
 * clearing the marquee.
 */
function applyPasteAtAnchor(payload: ClipboardPayload, anchor: { row: number; col: number }) {
  if (!activeWorksheet.value) {
    return;
  }
  const activeWorksheetBeforePaste = activeWorksheet.value;
  const sourceRange = payloadSourceRange(payload);
  const destinationRange = {
    worksheetId: activeWorksheetBeforePaste.id,
    startRow: anchor.row,
    startCol: anchor.col,
    endRow: anchor.row + payload.source.rows - 1,
    endCol: anchor.col + payload.source.cols - 1,
  };
  if (payload.mode === "cut" && sourceRange && rangesEqual(sourceRange, destinationRange)) {
    clipboardSourceRange.value = null;
    return;
  }

  app.updateGrid((grid) => {
    const worksheet = grid.workbook.worksheetsById[activeWorksheetBeforePaste.id];
    const operations: TeamGridOperation[] = [];
    ensureGridSize(worksheet, destinationRange.endRow + 1, destinationRange.endCol + 1, operations);
    const nextProjection = projectWorksheet(worksheet);
    const destinationIds = new Set<string>();
    const formulaDelta = {
      rows: payload.mode === "copy" ? anchor.row - payload.source.anchor.row : 0,
      cols: payload.mode === "copy" ? anchor.col - payload.source.anchor.col : 0,
    };

    for (const clipboardCell of payload.cells) {
      const targetRow = nextProjection.rows[anchor.row + clipboardCell.rowOffset];
      const targetColumn = nextProjection.columns[anchor.col + clipboardCell.colOffset];
      if (!targetRow || !targetColumn) {
        continue;
      }
      const targetCellId = createCellId(targetRow.id, targetColumn.id);
      destinationIds.add(targetCellId);
      const formulaSource = clipboardCell.formulaSource
        ? rewriteFormulaSource(clipboardCell.formulaSource, formulaDelta)
        : undefined;
      const nextCell: Cell = {
        id: targetCellId,
        rowId: targetRow.id,
        columnId: targetColumn.id,
        value: cloneCellValue(clipboardCell.value),
        style: clipboardCell.style ? { ...clipboardCell.style } : undefined,
        formula: undefined,
      };
      if (formulaSource) {
        applyFormulaToCell(nextCell, formulaSource, worksheet, nextProjection);
      }
      worksheet.cellsById[targetCellId] = nextCell;
      operations.push({ type: "setCell", worksheetId: worksheet.id, cell: nextCell });
    }

    if (payload.mode === "cut" && payload.source.worksheetId === worksheet.id && payload.cutCellIds) {
      for (const cellId of payload.cutCellIds) {
        if (destinationIds.has(cellId)) {
          continue;
        }
        const existing = worksheet.cellsById[cellId];
        if (!existing) {
          continue;
        }
        const emptyCell: Cell = { ...existing, value: { kind: "empty" }, formula: undefined };
        worksheet.cellsById[cellId] = emptyCell;
        operations.push({ type: "setCell", worksheetId: worksheet.id, cell: emptyCell });
      }
    }

    if (payload.mode === "cut" && sourceRange && payload.source.worksheetId === worksheet.id) {
      operations.push(...applyMoveTracking(worksheet, nextProjection, sourceRange, destinationRange));
    }

    return operations;
  });

  selectedRange.value = {
    startCellId: cellIdAt(destinationRange.startRow, destinationRange.startCol),
    endCellId: cellIdAt(destinationRange.endRow, destinationRange.endCol),
  };
  clipboardSourceRange.value = null;
}

/**
 * Append rows/columns to the worksheet until the projection reaches at
 * least `minRows` rows and `minCols` columns. Mutations are appended to
 * the caller's `operations` array so they participate in the granular
 * patch generated by the enclosing `updateGrid` call.
 */
function ensureGridSize(worksheet: NonNullable<typeof activeWorksheet.value>, minRows: number, minCols: number, operations: TeamGridOperation[]) {
  while (projectWorksheet(worksheet).rows.length < minRows) {
    const rowId = createId("row");
    const row = { id: rowId };
    worksheet.rowsById[rowId] = row;
    worksheet.rowOrder.push(rowId);
    operations.push({ type: "insertRow", worksheetId: worksheet.id, rowId, row, index: worksheet.rowOrder.length - 1 });
  }
  while (projectWorksheet(worksheet).columns.length < minCols) {
    const columnId = createId("col");
    const column = { id: columnId, width: DEFAULT_COLUMN_WIDTH };
    worksheet.columnsById[columnId] = column;
    worksheet.columnOrder.push(columnId);
    operations.push({ type: "insertColumn", worksheetId: worksheet.id, columnId, column, index: worksheet.columnOrder.length - 1 });
  }
}

/**
 * Parse, evaluate, and cache a formula on the given cell.
 *
 * Mutates `cell` in place so callers can build a fresh cell record and pipe
 * it through a single setCell operation. The cached `references` allow the
 * dependency tracker to skip re-parsing on subsequent reads.
 */
function applyFormulaToCell(cell: Cell, formulaSource: string, worksheet: NonNullable<typeof activeWorksheet.value>, worksheetProjection: NonNullable<typeof projection.value>) {
  const evaluated = evaluateFormula(formulaSource, worksheet, worksheetProjection);
  cell.formula = {
    kind: "formula",
    source: formulaSource,
    references: evaluated.references,
    cached: evaluated.result,
    error: evaluated.result.kind === "error" ? evaluated.result.code : undefined,
  };
  cell.value = formulaResultToCellValue(evaluated.result);
}

/**
 * Rewrite formulas across the worksheet that referenced the cut range so
 * they follow the moved cells (Excel-style "move tracking").
 *
 * Cells inside the source or destination range are skipped because they
 * have already been rewritten by `applyPasteAtAnchor` with the standard
 * copy delta. Cells whose formula source does not change are not touched
 * to keep the generated patch minimal.
 */
function applyMoveTracking(
  worksheet: NonNullable<typeof activeWorksheet.value>,
  worksheetProjection: NonNullable<typeof projection.value>,
  sourceRange: ClipboardRange,
  destinationRange: ClipboardRange,
) {
  const operations: TeamGridOperation[] = [];
  const delta = {
    rows: destinationRange.startRow - sourceRange.startRow,
    cols: destinationRange.startCol - sourceRange.startCol,
  };
  for (const cell of Object.values(worksheet.cellsById)) {
    const coordinates = findCellCoordinatesInProjection(cell.id, worksheet, worksheetProjection);
    if (!coordinates || rangeContains(sourceRange, coordinates.rowIndex, coordinates.columnIndex) || rangeContains(destinationRange, coordinates.rowIndex, coordinates.columnIndex) || !cell.formula?.source) {
      continue;
    }
    const nextSource = rewriteFormulaSource(cell.formula.source, delta, {
      insideRange: {
        startRow: sourceRange.startRow,
        startCol: sourceRange.startCol,
        endRow: sourceRange.endRow,
        endCol: sourceRange.endCol,
      },
    });
    if (nextSource === cell.formula.source) {
      continue;
    }
    const nextCell: Cell = { ...cell, formula: undefined };
    applyFormulaToCell(nextCell, nextSource, worksheet, worksheetProjection);
    worksheet.cellsById[nextCell.id] = nextCell;
    operations.push({ type: "setCell", worksheetId: worksheet.id, cell: nextCell });
  }
  return operations;
}

/** Resolve a stable cell id to its `{rowIndex, columnIndex}` in the current projection. */
function findCellCoordinates(cellId: string) {
  if (!activeWorksheet.value || !projection.value) {
    return null;
  }
  return findCellCoordinatesInProjection(cellId, activeWorksheet.value, projection.value);
}

/**
 * Variant of {@link findCellCoordinates} that operates on an explicit
 * worksheet + projection pair. Used inside `updateGrid` mutations where the
 * `worksheet` argument is a draft copy that the reactive `activeWorksheet`
 * has not yet observed.
 */
function findCellCoordinatesInProjection(
  cellId: string,
  worksheet: NonNullable<typeof activeWorksheet.value>,
  worksheetProjection: NonNullable<typeof projection.value>,
) {
  for (const row of worksheetProjection.rows) {
    for (const column of worksheetProjection.columns) {
      if (getCell(worksheet, row.id, column.id).id === cellId) {
        return { rowIndex: row.index, columnIndex: column.index };
      }
    }
  }
  return null;
}

/** Inverse of {@link findCellCoordinates}: turn `(row, col)` into a stable cell id. */
function cellIdAt(rowIndex: number, columnIndex: number) {
  if (!activeWorksheet.value || !projection.value) {
    return "";
  }
  const row = projection.value.rows[rowIndex];
  const column = projection.value.columns[columnIndex];
  return row && column ? createCellId(row.id, column.id) : "";
}

/**
 * Deep-clone a cell value. The clipboard payload is shared across paste
 * destinations, so cloning prevents accidental aliasing when the same
 * structured value (e.g. `{ kind: "number", value: 42 }`) ends up in many
 * target cells.
 */
function cloneCellValue(value: Cell["value"]): Cell["value"] {
  return JSON.parse(JSON.stringify(value)) as Cell["value"];
}

/**
 * Reconstruct the source range of a clipboard payload in worksheet-relative
 * coordinates. Returns `null` for payloads that did not come from this app
 * (no `worksheetId`), which avoids accidentally treating an Excel paste as
 * a cut-from-self.
 */
function payloadSourceRange(payload: ClipboardPayload): ClipboardRange | null {
  if (!payload.source.worksheetId) {
    return null;
  }
  return {
    worksheetId: payload.source.worksheetId,
    startRow: payload.source.anchor.row,
    startCol: payload.source.anchor.col,
    endRow: payload.source.anchor.row + payload.source.rows - 1,
    endCol: payload.source.anchor.col + payload.source.cols - 1,
  };
}

/** Structural equality for `ClipboardRange`. */
function rangesEqual(left: ClipboardRange, right: ClipboardRange) {
  return left.worksheetId === right.worksheetId
    && left.startRow === right.startRow
    && left.startCol === right.startCol
    && left.endRow === right.endRow
    && left.endCol === right.endCol;
}

/** Test whether `(rowIndex, columnIndex)` falls inside the inclusive `range`. */
function rangeContains(range: ClipboardRange, rowIndex: number, columnIndex: number) {
  return rowIndex >= range.startRow
    && rowIndex <= range.endRow
    && columnIndex >= range.startCol
    && columnIndex <= range.endCol;
}

/**
 * Append a picked cell address to the current draft.
 *
 * Inserts a `+` operator when the draft ends with an identifier or closing
 * paren, so picking `B1` after typing `=A1` produces `=A1+B1` instead of
 * the syntactically invalid `=A1B1`.
 */
function appendPickedAddress(source: string, address: string) {
  if (/[\w)]$/.test(source.trimEnd())) {
    return `${source}+${address}`;
  }
  return `${source}${address}`;
}

/** Open content assist against whichever formula editor is active. */
function openFormulaAssist(editor: FormulaAssistEditor, request: FormulaAssistRequest) {
  formulaAssistEditor.value = editor;
  formulaAssistAnchor.value = request.anchorEl;
  formulaAssistDraft.value = request.draft;
  formulaAssistCaretPos.value = request.caretPos;
  formulaAssistOpen.value = true;
}

/**
 * Route a function pick from the assist panel back into the editor that
 * triggered it. The panel itself is intentionally agnostic of the editor
 * implementation; this coordinator owns the routing decision.
 */
function handleFormulaAssistSelect(definition: FunctionDefinition) {
  if (formulaAssistEditor.value === "formulaBar") {
    void formulaBarComponent.value?.applyFormulaAssistSuggestion(definition);
  } else {
    void gridViewportComponent.value?.applyFormulaAssistSuggestion(definition);
  }
  formulaAssistOpen.value = false;
}

/**
 * Track the inline cell editor's open/close state and current draft so the
 * status-line hint stays in sync and so closing the editor also closes the
 * formula assist panel that was anchored to it.
 */
function handleInlineEditState(payload: { editing: boolean; draft: string }) {
  inlineCellEditing.value = payload.editing;
  inlineCellDraft.value = payload.draft;
  if (!payload.editing && formulaAssistEditor.value === "inlineCell") {
    formulaAssistOpen.value = false;
  }
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

/** Commit the formula bar to the selected cell and leave formula-edit mode. */
function commitFormulaBar(value: string) {
  if (!selectedCell.value) {
    return;
  }
  formulaEditing.value = false;
  formulaAssistOpen.value = false;
  commitCell(selectedCell.value, value);
}

/**
 * Cancel a formula-bar edit, reset the draft to the cell's current source,
 * and leave formula-edit mode without writing anything to the document.
 */
function cancelFormulaEdit() {
  formulaEditing.value = false;
  formulaAssistOpen.value = false;
  formulaError.value = null;
  formulaDraft.value = selectedCell.value?.formula?.source
    ?? (selectedCell.value ? formatCellValue(selectedCell.value.value, app.activeGrid.value?.settings.locale) : "");
}

/**
 * Open the Properties (title + tags) dialog.
 *
 * Guarded so we never offer edits for read-only revisions or empty states
 * where there is no document to edit.
 */
function openPropertiesDialog() {
  if (!app.currentEnvelope.value || app.gridReadOnly.value) {
    return;
  }
  resetPropertiesDraft();
  propertiesDialogVisible.value = true;
}

/**
 * Persist the title and tag edits from the Properties dialog.
 *
 * Title and tags are stored as top-level document fields (`subject` and
 * `tags`), so the view-backed Open dialog and the title button can read
 * them with simple field expressions. The change is routed through
 * `updateGrid` so it joins the granular save batch like any other edit.
 */
function applyDocumentProperties() {
  if (!app.currentEnvelope.value) {
    return;
  }
  const subject = propertiesTitleDraft.value.trim() || "Untitled spreadsheet";
  const tags = normalizeTags(propertiesTagsDraft.value.split(/\r?\n/));
  app.updateGrid((_grid, envelope) => {
    envelope.subject = subject;
    envelope.tags = tags;
    return [{ type: "setDocumentProperties", subject, tags }];
  });
  propertiesDialogVisible.value = false;
}

/**
 * Reset the Properties dialog inputs to the current document state. Called
 * both when the dialog opens and when the active document changes while the
 * dialog is closed, so reopening always shows fresh values.
 */
function resetPropertiesDraft() {
  propertiesTitleDraft.value = app.activeSubject.value || "Untitled spreadsheet";
  propertiesTagsDraft.value = app.activeTags.value.join("\n");
}

/**
 * Write a cell value coming from either the formula bar or the in-grid
 * editor.
 *
 * If the input starts with `=` we parse it as a formula, evaluate it
 * against the current worksheet, and cache both the AST references and
 * the latest result on the cell so dependent recomputation can skip
 * re-parsing. Otherwise we coerce the input through
 * {@link coerceInputToCellValue} so a column-typed cell still keeps its
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
 * Append a new worksheet to the workbook seeded with 24 rows × 12 columns
 * (a reasonable Excel-like default) and switch the UI to the new tab.
 *
 * The title is chosen by {@link nextWorksheetTitle} to avoid collisions
 * with existing or tombstoned worksheets.
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
function nextWorksheetTitle(grid: { workbook: { worksheetsById: Record<string, { title: string }> } }) {
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

/**
 * Open the rename dialog for the given worksheet.
 *
 * We can't rely on `window.prompt` here because Haven's app shell renders
 * the iframe with a sandbox that does not include `allow-modals`, which
 * silently suppresses native prompts. A regular PrimeVue Dialog works
 * in every host configuration.
 */
function renameWorksheet(worksheetId: WorksheetId) {
  if (app.gridReadOnly.value) {
    return;
  }
  const currentTitle = app.activeGrid.value?.workbook.worksheetsById[worksheetId]?.title ?? "";
  renameTargetId.value = worksheetId;
  renameDraft.value = currentTitle;
  renameDialogVisible.value = true;
}

/** Persist the rename dialog's draft title for the captured worksheet. */
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

/**
 * Apply a partial style patch to every cell in the current selection.
 *
 * The patch is shallow-merged onto each cell so the user can, for example,
 * change only `fontWeight` without losing previously set `backgroundColor`.
 * A single `setCellsStyle` operation is emitted so the persisted patch is
 * compact and Automerge sees one logical edit per selection.
 */
function patchSelectedStyle(style: CellStyle) {
  if (!activeWorksheet.value || selectedCells.value.length === 0) return;
  const cellsToPatch = selectedCells.value;
  app.updateGrid((grid) => {
    const worksheet = grid.workbook.worksheetsById[activeWorksheet.value!.id];
    const patchedCells: Cell[] = [];
    for (const cell of cellsToPatch) {
      const existing = worksheet.cellsById[cell.id] ?? cell;
      const patchedCell = {
        ...existing,
        style: {
          ...existing.style,
          ...style,
        },
      };
      worksheet.cellsById[existing.id] = patchedCell;
      patchedCells.push(patchedCell);
    }
    return [{ type: "setCellsStyle", worksheetId: worksheet.id, cells: patchedCells, style }];
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

        <div class="format-toolbar">
          <label>
            Text
            <input type="color" :disabled="app.gridReadOnly.value || !hasSelection" :value="selectedCell?.style?.textColor ?? '#eef2ff'" @input="patchSelectedStyle({ textColor: ($event.target as HTMLInputElement).value })">
          </label>
          <label>
            Fill
            <input type="color" :disabled="app.gridReadOnly.value || !hasSelection" :value="selectedCell?.style?.backgroundColor ?? '#111827'" @input="patchSelectedStyle({ backgroundColor: ($event.target as HTMLInputElement).value })">
          </label>
          <label>
            Font size
            <input class="format-toolbar__number" type="number" min="8" max="48" :disabled="app.gridReadOnly.value || !hasSelection" :value="selectedCell?.style?.fontSize ?? 14" @change="patchSelectedStyle({ fontSize: Number(($event.target as HTMLInputElement).value) })">
          </label>
          <span class="format-toolbar__divider" aria-hidden="true" />
          <button
            type="button"
            class="format-toolbar__text-button format-toolbar__text-button--bold"
            aria-label="Bold"
            :disabled="app.gridReadOnly.value || !hasSelection"
            :class="{ 'format-toolbar__button--active': selectedCell?.style?.bold }"
            @click="patchSelectedStyle({ bold: !selectedCell?.style?.bold })"
          >
            B
          </button>
          <button
            type="button"
            class="format-toolbar__text-button format-toolbar__text-button--italic"
            aria-label="Italic"
            :disabled="app.gridReadOnly.value || !hasSelection"
            :class="{ 'format-toolbar__button--active': selectedCell?.style?.italic }"
            @click="patchSelectedStyle({ italic: !selectedCell?.style?.italic })"
          >
            I
          </button>
          <button
            type="button"
            class="format-toolbar__text-button format-toolbar__text-button--underline"
            aria-label="Underline"
            :disabled="app.gridReadOnly.value || !hasSelection"
            :class="{ 'format-toolbar__button--active': selectedCell?.style?.underline }"
            @click="patchSelectedStyle({ underline: !selectedCell?.style?.underline })"
          >
            U
          </button>
          <span class="format-toolbar__divider" aria-hidden="true" />
          <button
            type="button"
            class="format-toolbar__value-button"
            :disabled="app.gridReadOnly.value || !hasSelection"
            @click="openCellFormatDialog(selectedRange)"
          >
            Format cells
          </button>
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
          :clipboard-range="clipboardSourceRange"
          :highlighted-cell-ids="highlightedCellIds"
          :readonly="app.gridReadOnly.value"
          :locale="app.activeGrid.value.settings.locale"
          @select="selectCell"
          @select-range="selectRange"
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

    <Dialog v-model:visible="formatDialogVisible" modal header="Format cells" :style="{ width: '32rem', maxWidth: '96vw' }">
      <div class="cell-format-dialog">
        <label class="field">
          Format
          <select v-model="formatDialogKind" class="native-input">
            <option value="text">Text</option>
            <option value="general">Standard</option>
            <option value="integer">Integer</option>
            <option value="decimal">Decimal</option>
            <option value="percent">Percent</option>
            <option value="currency">Currency</option>
            <option value="custom">Custom Excel format</option>
          </select>
        </label>
        <label v-if="formatDialogKind === 'currency'" class="field">
          Currency
          <select v-model="formatDialogCurrency" class="native-input">
            <option value="USD">USD</option>
            <option value="EUR">EUR</option>
          </select>
        </label>
        <label v-if="formatDialogKind === 'custom'" class="field">
          Excel number format
          <input
            v-model="formatDialogCustomNumFmt"
            class="native-input"
            type="text"
            autocomplete="off"
            placeholder="$#,##0.00;[Red]-$#,##0.00"
            @keyup.enter="applySelectedCellFormat"
          >
        </label>
        <p class="cell-format-dialog__hint">Formatting is applied to the current selection. Numeric-looking text is converted for number, percent, and currency formats; incompatible values are left unchanged.</p>
      </div>
      <template #footer>
        <Button label="Cancel" text @click="formatDialogVisible = false" />
        <Button label="Apply" icon="pi pi-check" :disabled="app.gridReadOnly.value || (formatDialogKind === 'custom' && !formatDialogCustomNumFmt.trim())" @click="applySelectedCellFormat" />
      </template>
    </Dialog>

    <Dialog
      v-model:visible="errorDialogVisible"
      modal
      header="Something went wrong"
      :style="{ width: '32rem', maxWidth: '96vw' }"
      @hide="app.clearLastError"
    >
      <p>{{ app.lastErrorMessage.value }}</p>
      <template #footer>
        <Button label="OK" icon="pi pi-check" @click="dismissErrorDialog" />
      </template>
    </Dialog>

    <FormulaAssistPanel
      v-model:visible="formulaAssistOpen"
      :draft="formulaAssistDraft"
      :caret-pos="formulaAssistCaretPos"
      :anchor-el="formulaAssistAnchor"
      :readonly="app.gridReadOnly.value"
      @select="handleFormulaAssistSelect"
      @dismiss="formulaAssistOpen = false"
    />

    <Dialog v-model:visible="openDialogVisible" modal header="Open spreadsheet" :style="{ width: '58rem', maxWidth: '96vw' }" @hide="disposeOpenNavigator">
      <div class="open-dialog">
        <label class="field">
          Database
          <select v-model="app.selectedDatabaseId.value" class="native-input" @change="handleOpenDatabaseChange">
            <option v-for="database in app.readableDatabases.value" :key="database.id" :value="database.id">{{ database.title || database.id }}</option>
          </select>
        </label>
        <div class="open-dialog__browser">
          <aside class="open-dialog__tree" aria-label="Spreadsheet tags">
            <TagTreeList
              :nodes="openCategoryNodes"
              :selected-key="selectedOpenCategoryKey"
              @select="selectOpenCategory"
            />
          </aside>
          <div class="document-list">
            <button
              v-for="document in openDialogDocuments"
              :key="document.id"
              class="document-row"
              :class="{ 'document-row--selected': document.id === selectedOpenDocId }"
              type="button"
              @click="selectedOpenDocId = document.id"
              @dblclick="openSelectedDocument"
            >
              <strong>{{ document.title }}</strong>
              <small>{{ document.detail }}</small>
              <small>{{ document.id }}</small>
            </button>
            <p v-if="openDialogDocuments.length === 0" class="document-list__empty">No spreadsheets in this category.</p>
          </div>
        </div>
      </div>
      <template #footer>
        <Button label="Cancel" text @click="openDialogVisible = false" />
        <Button label="Open" icon="pi pi-folder-open" :disabled="!selectedOpenDocId" @click="openSelectedDocument" />
      </template>
    </Dialog>

    <Dialog v-model:visible="propertiesDialogVisible" modal header="Spreadsheet properties" :style="{ width: '34rem', maxWidth: '96vw' }">
      <div class="properties-dialog">
        <label class="field">
          Title
          <input v-model="propertiesTitleDraft" class="native-input" type="text" autocomplete="off">
        </label>
        <label class="field">
          Tags
          <textarea v-model="propertiesTagsDraft" class="native-input native-input--textarea" rows="6" placeholder="Work\Planning&#10;Finance" />
        </label>
        <p class="properties-dialog__hint">Enter one tag per line. Use a backslash to create hierarchy, for example <code>Work\Planning</code>.</p>
      </div>
      <template #footer>
        <Button label="Cancel" text @click="propertiesDialogVisible = false; resetPropertiesDraft()" />
        <Button label="Apply" icon="pi pi-check" :disabled="app.gridReadOnly.value" @click="applyDocumentProperties" />
      </template>
    </Dialog>

    <Dialog v-model:visible="deleteDialogVisible" modal header="Delete spreadsheet" :style="{ width: '28rem', maxWidth: '96vw' }">
      <p>This deletes the current spreadsheet document from the selected MindooDB database.</p>
      <template #footer>
        <Button label="Cancel" text @click="deleteDialogVisible = false" />
        <Button label="Delete" icon="pi pi-trash" severity="danger" @click="deleteDialogVisible = false; app.deleteCurrentDocument()" />
      </template>
    </Dialog>

    <Dialog v-model:visible="renameDialogVisible" modal header="Rename worksheet" :style="{ width: '24rem', maxWidth: '96vw' }">
      <label class="field">
        Name
        <input
          v-model="renameDraft"
          class="native-input"
          type="text"
          autocomplete="off"
          autofocus
          @keyup.enter="applyWorksheetRename"
        >
      </label>
      <template #footer>
        <Button label="Cancel" text @click="renameDialogVisible = false" />
        <Button label="Rename" icon="pi pi-check" :disabled="!renameDraft.trim()" @click="applyWorksheetRename" />
      </template>
    </Dialog>

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
