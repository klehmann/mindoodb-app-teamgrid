<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from "vue";
import {
  formatCellValue,
  formatFormulaResult,
  mergeCellStyle,
} from "@/lib/cellFormatting";
import { evaluateFormula } from "@/lib/formulas";
import { getCell, getCellAddress, type GridProjection } from "@/lib/gridProjection";
import type { Cell, ColumnId, RowId, Worksheet } from "@/lib/teamgridDocument";

const props = defineProps<{
  worksheet: Worksheet;
  projection: GridProjection;
  selectedCellId: string | null;
  highlightedCellIds: string[];
  readonly: boolean;
  locale: string;
}>();

const emit = defineEmits<{
  select: [cell: Cell, address: string];
  commit: [cell: Cell, rawValue: string];
}>();

const editingCellId = ref<string | null>(null);
const editDraft = ref("");
const gridViewport = ref<HTMLElement | null>(null);

const highlighted = computed(() => new Set(props.highlightedCellIds));

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
});

onBeforeUnmount(() => {
  window.removeEventListener("keydown", handleWindowKeydown);
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
  selectCell(rowId, columnId);
  await nextTick();
  gridViewport.value?.querySelector<HTMLInputElement>(".grid-cell__editor")?.focus();
}

function commitEdit(cell: Cell) {
  emit("commit", cell, editDraft.value);
  editingCellId.value = null;
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

function isTypingInAnotherEditor(target: EventTarget | null) {
  if (!(target instanceof HTMLElement)) {
    return false;
  }
  return target.matches("input, textarea, select, [contenteditable='true']");
}

function isPrintableKey(event: KeyboardEvent) {
  return event.key.length === 1 && !event.metaKey && !event.ctrlKey && !event.altKey;
}
</script>

<template>
  <div ref="gridViewport" class="grid-viewport">
    <table class="grid-table" aria-label="Spreadsheet grid">
      <thead>
        <tr>
          <th class="grid-corner" scope="col" />
          <th
            v-for="column in projection.columns"
            :key="column.id"
            class="grid-column-header"
            scope="col"
            :style="{ width: `${column.width}px`, minWidth: `${column.width}px` }"
          >
            {{ column.label }}
          </th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="row in projection.rows" :key="row.id" :style="{ height: row.height ? `${row.height}px` : undefined }">
          <th class="grid-row-header" scope="row">{{ row.label }}</th>
          <td
            v-for="column in projection.columns"
            :key="column.id"
            class="grid-cell"
            :class="{
              'grid-cell--selected': getCell(worksheet, row.id, column.id).id === selectedCellId,
              'grid-cell--highlighted': highlighted.has(getCell(worksheet, row.id, column.id).id),
              'grid-cell--formula': Boolean(getCell(worksheet, row.id, column.id).formula),
            }"
            :style="cellStyle(getCell(worksheet, row.id, column.id))"
            tabindex="0"
            @click="handleCellClick($event, row.id, column.id)"
            @dblclick="startEditing(row.id, column.id)"
            @keydown="handleEditKey($event, row.id, column.id)"
          >
            <input
              v-if="editingCellId === getCell(worksheet, row.id, column.id).id"
              v-model="editDraft"
              class="grid-cell__editor"
              autofocus
              @keydown.enter.prevent="commitEdit(getCell(worksheet, row.id, column.id))"
              @keydown.escape.prevent="editingCellId = null"
              @blur="commitEdit(getCell(worksheet, row.id, column.id))"
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
