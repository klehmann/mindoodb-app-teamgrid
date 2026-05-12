<script setup lang="ts">
/**
 * Spreadsheet grid renderer.
 *
 * Responsibilities:
 * - Render rows/columns from the supplied {@link GridProjection}.
 * - Manage in-cell editing (double-click or start typing to enter edit
 *   mode, Enter to commit, Escape to cancel).
 * - Maintain rectangular range selections driven by mouse drag.
 * - Forward header clicks so the parent can switch to row/column selection.
 *
 * The component does not own document state; it operates on a `worksheet`
 * prop and emits semantic events that the parent translates into
 * {@link TeamGridOperation}s via `useTeamGridDocument.updateGrid`.
 *
 * Props:
 * - `worksheet`: persisted worksheet to render.
 * - `projection`: derived view of the worksheet (visible rows/columns plus
 *   `A1` address lookups).
 * - `selectedCellId`: id of the currently active cell or `null`.
 * - `selectedRange`: rectangular range selection, when one is active.
 * - `highlightedCellIds`: cells highlighted because the formula bar is
 *   picking references; rendered with a translucent overlay.
 * - `readonly`: disables editing.
 * - `locale`: BCP-47 locale for number/date formatting.
 *
 * Emits:
 * - `select(cell, address)`: cell click.
 * - `select-range(range)`: drag selection completed.
 * - `commit(cell, rawValue)`: user pressed Enter or blurred a cell edit.
 */
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from "vue";
import {
  formatCellValue,
  formatFormulaResult,
  mergeCellStyle,
} from "@/lib/cellFormatting";
import { evaluateFormula } from "@/lib/formulas";
import { insertFunctionAtCaret } from "@/lib/formulas/assist";
import { getCell, getCellAddress, type GridProjection } from "@/lib/gridProjection";
import type { Cell, CellId, ColumnId, RowId, Worksheet } from "@/lib/teamgridDocument";
import type { FunctionDefinition } from "@/lib/formulas";

/** Inclusive rectangular cell range, addressed by stable cell ids. */
export interface CellSelectionRange {
  startCellId: CellId;
  endCellId: CellId;
}

/** Coordinate range used to draw the clipboard source marquee. */
export interface GridClipboardRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

const props = defineProps<{
  worksheet: Worksheet;
  projection: GridProjection;
  selectedCellId: string | null;
  selectedRange: CellSelectionRange | null;
  clipboardRange: GridClipboardRange | null;
  highlightedCellIds: string[];
  readonly: boolean;
  locale: string;
}>();

const emit = defineEmits<{
  select: [cell: Cell, address: string];
  "select-range": [range: CellSelectionRange];
  commit: [cell: Cell, rawValue: string];
  "request-help": [payload: { anchorEl: HTMLElement; draft: string; caretPos: number }];
  "edit-state": [payload: { editing: boolean; draft: string }];
  "clipboard-copy": [payload: { range: CellSelectionRange | null; event: ClipboardEvent }];
  "clipboard-cut": [payload: { range: CellSelectionRange | null; event: ClipboardEvent }];
  "clipboard-paste": [payload: { event: ClipboardEvent }];
  "clipboard-clear": [];
}>();

const editingCellId = ref<string | null>(null);
const editDraft = ref("");
const gridViewport = ref<HTMLElement | null>(null);
const editorInputEl = ref<HTMLInputElement | HTMLInputElement[] | null>(null);
const draggingRangeStart = ref<CellId | null>(null);
let suppressNextBlurCommit = false;

const highlighted = computed(() => new Set(props.highlightedCellIds));
const selectedRangeIds = computed(() => new Set(props.selectedRange ? getRangeCellIds(props.selectedRange) : []));

watch(
  () => props.selectedCellId,
  () => {
    if (editingCellId.value && editingCellId.value !== props.selectedCellId) {
      editingCellId.value = null;
    }
  },
);

onMounted(() => {
  window.addEventListener("keydown", handleWindowKeydown);
  window.addEventListener("mouseup", endRangeDrag);
});

onBeforeUnmount(() => {
  window.removeEventListener("keydown", handleWindowKeydown);
  window.removeEventListener("mouseup", endRangeDrag);
});

function displayCell(cell: Cell) {
  if (cell.formula) {
    return formatFormulaResult(evaluateFormula(cell.formula.source, props.worksheet, props.projection).result, props.locale);
  }
  return formatCellValue(cell.value, props.locale);
}

function selectCell(rowId: RowId, columnId: ColumnId) {
  const cell = getCell(props.worksheet, rowId, columnId);
  emit("select", cell, getCellAddress(props.projection, rowId, columnId));
}

async function startEditing(rowId: RowId, columnId: ColumnId, initialValue?: string) {
  if (props.readonly) {
    return;
  }
  const cell = getCell(props.worksheet, rowId, columnId);
  editingCellId.value = cell.id;
  editDraft.value = initialValue ?? cell.formula?.source ?? displayCell(cell);
  emit("edit-state", { editing: true, draft: editDraft.value });
  selectCell(rowId, columnId);
  await nextTick();
  getEditorInputEl()?.focus();
}

function commitEdit(cell: Cell) {
  emit("commit", cell, editDraft.value);
  editingCellId.value = null;
  emit("edit-state", { editing: false, draft: "" });
}

function commitEditFromKeyboard(event: KeyboardEvent, cell: Cell) {
  suppressNextBlurCommit = true;
  commitEdit(cell);
  (event.currentTarget as HTMLInputElement | null)?.blur();
  queueMicrotask(() => {
    suppressNextBlurCommit = false;
  });
}

function handleEditorKeydown(event: KeyboardEvent, cell: Cell) {
  if ((event.ctrlKey || event.metaKey) && event.code === "Space") {
    event.preventDefault();
    event.stopPropagation();
    requestFormulaHelp();
    return;
  }
  if (event.key === "Enter" || event.code === "NumpadEnter" || event.keyCode === 13) {
    event.preventDefault();
    event.stopPropagation();
    commitEditFromKeyboard(event, cell);
    return;
  }
  if (event.key === "Escape") {
    event.preventDefault();
    event.stopPropagation();
    editingCellId.value = null;
    emit("edit-state", { editing: false, draft: "" });
  }
}

function updateEditDraft(value: string) {
  editDraft.value = value;
  emit("edit-state", { editing: Boolean(editingCellId.value), draft: editDraft.value });
}

async function applyFormulaAssistSuggestion(definition: FunctionDefinition) {
  if (!editingCellId.value) {
    return;
  }
  const input = getEditorInputEl();
  const caretPos = input?.selectionStart ?? editDraft.value.length;
  const inserted = insertFunctionAtCaret(editDraft.value, caretPos, definition.name);
  editDraft.value = inserted.next;
  await nextTick();
  getEditorInputEl()?.focus();
  getEditorInputEl()?.setSelectionRange(inserted.nextCaret, inserted.nextCaret);
}

function requestFormulaHelp() {
  const input = getEditorInputEl();
  if (props.readonly || !input || !editDraft.value.trim().startsWith("=")) {
    return;
  }
  emit("request-help", {
    anchorEl: input,
    draft: editDraft.value,
    caretPos: input.selectionStart ?? editDraft.value.length,
  });
}

function commitEditFromBlur(cell: Cell) {
  if (suppressNextBlurCommit) {
    return;
  }
  commitEdit(cell);
}

function cellStyle(cell: Cell) {
  const mergedStyle = mergeCellStyle(props.worksheet.rowsById[cell.rowId], props.worksheet.columnsById[cell.columnId], cell);
  return {
    color: mergedStyle.textColor,
    backgroundColor: mergedStyle.backgroundColor,
    fontFamily: mergedStyle.fontFamily,
    fontSize: mergedStyle.fontSize ? `${mergedStyle.fontSize}px` : undefined,
    fontWeight: mergedStyle.bold ? "700" : undefined,
    fontStyle: mergedStyle.italic ? "italic" : undefined,
    textDecoration: mergedStyle.underline ? "underline" : undefined,
    textAlign: mergedStyle.horizontalAlign,
    verticalAlign: mergedStyle.verticalAlign,
  };
}

function handleCellClick(event: MouseEvent, rowId: RowId, columnId: ColumnId) {
  selectCell(rowId, columnId);
  (event.currentTarget as HTMLElement | null)?.focus();
}

function startRangeSelection(event: MouseEvent, rowId: RowId, columnId: ColumnId) {
  if (event.button !== 0) {
    return;
  }
  const cell = getCell(props.worksheet, rowId, columnId);
  if (editingCellId.value) {
    if (editDraft.value.trim().startsWith("=")) {
      appendPickedCellToInlineFormula(cell);
      return;
    }
    const editingCell = findEditingCell();
    if (editingCell) {
      suppressNextBlurCommit = true;
      commitEdit(editingCell);
      queueMicrotask(() => {
        suppressNextBlurCommit = false;
      });
    }
  }
  draggingRangeStart.value = cell.id;
  emit("select-range", { startCellId: cell.id, endCellId: cell.id });
  handleCellClick(event, rowId, columnId);
}

function selectWholeRow(rowId: RowId) {
  if (editingCellId.value || props.projection.columns.length === 0) {
    return;
  }
  const firstColumn = props.projection.columns[0];
  const lastColumn = props.projection.columns[props.projection.columns.length - 1];
  const firstCell = getCell(props.worksheet, rowId, firstColumn.id);
  const lastCell = getCell(props.worksheet, rowId, lastColumn.id);
  emit("select", firstCell, getCellAddress(props.projection, firstCell.rowId, firstCell.columnId));
  emit("select-range", { startCellId: firstCell.id, endCellId: lastCell.id });
}

function selectWholeColumn(columnId: ColumnId) {
  if (editingCellId.value || props.projection.rows.length === 0) {
    return;
  }
  const firstRow = props.projection.rows[0];
  const lastRow = props.projection.rows[props.projection.rows.length - 1];
  const firstCell = getCell(props.worksheet, firstRow.id, columnId);
  const lastCell = getCell(props.worksheet, lastRow.id, columnId);
  emit("select", firstCell, getCellAddress(props.projection, firstCell.rowId, firstCell.columnId));
  emit("select-range", { startCellId: firstCell.id, endCellId: lastCell.id });
}

function appendPickedCellToInlineFormula(cell: Cell) {
  if (!editDraft.value.trim().startsWith("=") || cell.id === editingCellId.value) {
    return;
  }
  const address = getCellAddress(props.projection, cell.rowId, cell.columnId);
  editDraft.value = /[\w)]$/.test(editDraft.value.trimEnd())
    ? `${editDraft.value}+${address}`
    : `${editDraft.value}${address}`;
}

function extendRangeSelection(rowId: RowId, columnId: ColumnId) {
  if (!draggingRangeStart.value) {
    return;
  }
  const cell = getCell(props.worksheet, rowId, columnId);
  emit("select-range", { startCellId: draggingRangeStart.value, endCellId: cell.id });
}

function endRangeDrag() {
  draggingRangeStart.value = null;
}

function handleWindowKeydown(event: KeyboardEvent) {
  if (event.defaultPrevented || isTypingInAnotherEditor(event.target) || !props.selectedCellId) {
    return;
  }
  const selectedCell = findSelectedCell();
  if (!selectedCell) {
    return;
  }
  handleEditKey(event, selectedCell.rowId, selectedCell.columnId);
}

function handleEditKey(event: KeyboardEvent, rowId: RowId, columnId: ColumnId) {
  if (props.readonly || editingCellId.value) {
    return;
  }
  if (event.key === "Enter" || event.key === "F2") {
    event.preventDefault();
    void startEditing(rowId, columnId);
    return;
  }
  if (event.key === "Backspace" || event.key === "Delete") {
    event.preventDefault();
    void startEditing(rowId, columnId, "");
    return;
  }
  if (isPrintableKey(event)) {
    event.preventDefault();
    void startEditing(rowId, columnId, event.key);
  }
}

function findSelectedCell() {
  if (!props.selectedCellId) {
    return null;
  }
  for (const row of props.projection.rows) {
    for (const column of props.projection.columns) {
      const cell = getCell(props.worksheet, row.id, column.id);
      if (cell.id === props.selectedCellId) {
        return cell;
      }
    }
  }
  return null;
}

function findEditingCell() {
  if (!editingCellId.value) {
    return null;
  }
  for (const row of props.projection.rows) {
    for (const column of props.projection.columns) {
      const cell = getCell(props.worksheet, row.id, column.id);
      if (cell.id === editingCellId.value) {
        return cell;
      }
    }
  }
  return null;
}

function getRangeCellIds(range: CellSelectionRange) {
  const start = findCellCoordinates(range.startCellId);
  const end = findCellCoordinates(range.endCellId);
  if (!start || !end) {
    return [];
  }
  const ids: CellId[] = [];
  for (let rowIndex = Math.min(start.rowIndex, end.rowIndex); rowIndex <= Math.max(start.rowIndex, end.rowIndex); rowIndex += 1) {
    for (let columnIndex = Math.min(start.columnIndex, end.columnIndex); columnIndex <= Math.max(start.columnIndex, end.columnIndex); columnIndex += 1) {
      const row = props.projection.rows[rowIndex];
      const column = props.projection.columns[columnIndex];
      if (row && column) {
        ids.push(getCell(props.worksheet, row.id, column.id).id);
      }
    }
  }
  return ids;
}

function isWholeRowSelected(rowId: RowId) {
  return props.projection.columns.length > 0
    && props.projection.columns.every((column) => selectedRangeIds.value.has(getCell(props.worksheet, rowId, column.id).id));
}

function isWholeColumnSelected(columnId: ColumnId) {
  return props.projection.rows.length > 0
    && props.projection.rows.every((row) => selectedRangeIds.value.has(getCell(props.worksheet, row.id, columnId).id));
}

function isInClipboardRange(rowIndex: number, columnIndex: number) {
  if (!props.clipboardRange) {
    return false;
  }
  const startRow = Math.min(props.clipboardRange.startRow, props.clipboardRange.endRow);
  const endRow = Math.max(props.clipboardRange.startRow, props.clipboardRange.endRow);
  const startCol = Math.min(props.clipboardRange.startCol, props.clipboardRange.endCol);
  const endCol = Math.max(props.clipboardRange.startCol, props.clipboardRange.endCol);
  return rowIndex >= startRow && rowIndex <= endRow && columnIndex >= startCol && columnIndex <= endCol;
}

function findCellCoordinates(cellId: CellId) {
  for (const row of props.projection.rows) {
    for (const column of props.projection.columns) {
      if (getCell(props.worksheet, row.id, column.id).id === cellId) {
        return { rowIndex: row.index, columnIndex: column.index };
      }
    }
  }
  return null;
}

function isTypingInAnotherEditor(target: EventTarget | null) {
  if (!(target instanceof HTMLElement)) {
    return false;
  }
  return target.matches("input, textarea, select, [contenteditable='true']");
}

function getEditorInputEl() {
  return Array.isArray(editorInputEl.value) ? editorInputEl.value[0] : editorInputEl.value;
}

function isPrintableKey(event: KeyboardEvent) {
  return event.key.length === 1 && !event.metaKey && !event.ctrlKey && !event.altKey;
}

function handleViewportCopy(event: ClipboardEvent) {
  if (isTypingInAnotherEditor(event.target)) {
    return;
  }
  event.preventDefault();
  emit("clipboard-copy", { range: props.selectedRange, event });
}

function handleViewportCut(event: ClipboardEvent) {
  if (isTypingInAnotherEditor(event.target) || props.readonly) {
    return;
  }
  event.preventDefault();
  emit("clipboard-cut", { range: props.selectedRange, event });
}

function handleViewportPaste(event: ClipboardEvent) {
  if (isTypingInAnotherEditor(event.target) || props.readonly) {
    return;
  }
  event.preventDefault();
  emit("clipboard-paste", { event });
}

function handleViewportKeydown(event: KeyboardEvent) {
  if (event.key === "Escape" && !editingCellId.value) {
    emit("clipboard-clear");
  }
}

defineExpose({ applyFormulaAssistSuggestion });
</script>

<template>
  <div
    ref="gridViewport"
    class="grid-viewport"
    tabindex="0"
    @copy="handleViewportCopy"
    @cut="handleViewportCut"
    @paste="handleViewportPaste"
    @keydown="handleViewportKeydown"
  >
    <table class="grid-table" aria-label="Spreadsheet grid">
      <thead>
        <tr>
          <th class="grid-corner" scope="col" />
          <th
            v-for="column in projection.columns"
            :key="column.id"
            class="grid-column-header"
            :class="{ 'grid-axis-header--selected': isWholeColumnSelected(column.id) }"
            scope="col"
            :style="{ width: `${column.width}px`, minWidth: `${column.width}px` }"
            @mousedown.prevent="selectWholeColumn(column.id)"
          >
            {{ column.label }}
          </th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="row in projection.rows" :key="row.id" :style="{ height: row.height ? `${row.height}px` : undefined }">
          <th
            class="grid-row-header"
            :class="{ 'grid-axis-header--selected': isWholeRowSelected(row.id) }"
            scope="row"
            @mousedown.prevent="selectWholeRow(row.id)"
          >
            {{ row.label }}
          </th>
          <td
            v-for="column in projection.columns"
            :key="column.id"
            class="grid-cell"
            :class="{
              'grid-cell--selected': getCell(worksheet, row.id, column.id).id === selectedCellId,
              'grid-cell--range-selected': selectedRangeIds.has(getCell(worksheet, row.id, column.id).id),
              'grid-cell--highlighted': highlighted.has(getCell(worksheet, row.id, column.id).id),
              'grid-cell--formula': Boolean(getCell(worksheet, row.id, column.id).formula),
              'grid-cell--clipboard-source': isInClipboardRange(row.index, column.index),
            }"
            :style="cellStyle(getCell(worksheet, row.id, column.id))"
            tabindex="0"
            @mousedown.prevent="startRangeSelection($event, row.id, column.id)"
            @mouseenter="extendRangeSelection(row.id, column.id)"
            @dblclick="startEditing(row.id, column.id)"
            @keydown="handleEditKey($event, row.id, column.id)"
          >
            <input
              v-if="editingCellId === getCell(worksheet, row.id, column.id).id"
              ref="editorInputEl"
              v-model="editDraft"
              class="grid-cell__editor"
              autofocus
              @input="updateEditDraft(($event.target as HTMLInputElement).value)"
              @keydown="handleEditorKeydown($event, getCell(worksheet, row.id, column.id))"
              @blur="commitEditFromBlur(getCell(worksheet, row.id, column.id))"
            >
            <span v-else>{{ displayCell(getCell(worksheet, row.id, column.id)) }}</span>
          </td>
        </tr>
      </tbody>
    </table>
  </div>
</template>

<style scoped>
.grid-viewport {
  overflow: auto;
  min-height: 0;
  flex: 1 1 auto;
}

.grid-table {
  border-collapse: collapse;
  width: max-content;
  min-width: 100%;
  table-layout: fixed;
  font-size: 0.92rem;
}

.grid-corner,
.grid-column-header,
.grid-row-header {
  position: sticky;
  z-index: 2;
  border: 1px solid var(--grid-border);
  background: var(--grid-header-bg);
  color: var(--muted);
  font-weight: 600;
  user-select: none;
}

.grid-corner {
  top: 0;
  left: 0;
  z-index: 3;
  width: 3rem;
  min-width: 3rem;
}

.grid-column-header {
  top: 0;
  height: 2rem;
}

.grid-row-header {
  left: 0;
  width: 3rem;
  min-width: 3rem;
}

.grid-column-header,
.grid-row-header {
  cursor: pointer;
}

.grid-axis-header--selected {
  background: rgb(212 160 23 / 0.18);
  color: var(--text);
}

.grid-cell {
  height: 2rem;
  max-width: 18rem;
  padding: 0.25rem 0.45rem;
  overflow: hidden;
  border: 1px solid var(--grid-border);
  background: var(--grid-cell-bg);
  text-overflow: ellipsis;
  white-space: nowrap;
  cursor: cell;
}

.grid-cell--selected {
  outline: 2px solid var(--accent);
  outline-offset: -2px;
}

.grid-cell--range-selected {
  background: rgb(212 160 23 / 0.12);
}

.grid-cell--highlighted {
  background: rgb(212 160 23 / 0.18);
}

.grid-cell--formula {
  color: var(--formula-text);
}

.grid-cell__editor {
  width: 100%;
  height: 100%;
  border: 0;
  outline: 0;
  background: transparent;
  color: inherit;
  font: inherit;
}
</style>
