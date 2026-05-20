<script setup lang="ts">
/**
 * GridViewport - the spreadsheet surface that renders rows, columns, and
 * cells of a {@link Worksheet} and turns user gestures into semantic
 * editing events.
 *
 * # Where this fits in TeamGrid
 *
 * TeamGrid stores a workbook as an Automerge document with stable
 * application IDs for rows, columns, and cells (see the project README,
 * "Concurrency Model"). Visible `A1` notation, cell positions, and even
 * the cell-by-cell DOM layout are computed on every render from the
 * current worksheet plus a derived {@link GridProjection}.
 *
 * GridViewport is intentionally a "dumb" view component:
 *
 * - It does **not** own document state; the parent passes the persisted
 *   {@link Worksheet} and a {@link GridProjection} as props.
 * - It does **not** mutate the document; instead, it emits semantic
 *   events that the parent translates into `TeamGridOperation`s via
 *   `useTeamGridDocument.updateGrid`.
 * - The only mutable state living here is transient UI state - the
 *   active inline editor, the in-progress drag-range, and the
 *   in-progress row/column resize - and each of those is delegated to a
 *   focused composable so the SFC stays a thin orchestration layer:
 *
 *     - `useInlineCellEditor`     editor draft + IME + commit/cancel
 *     - `useGridSelectionGestures` mouse/keyboard selection
 *     - `useColumnRowResize`       pointer-driven resize handles
 *     - `useGridClipboardBridge`   native copy/cut/paste forwarding
 *
 * # Exposed methods
 *
 * - `applyFormulaAssistSuggestion(definition)`: insert a function chosen
 *   from the formula assist panel into the in-cell draft.
 * - `flushPendingEdit()`: synchronously commit any pending inline edit.
 *   Used by toolbar actions to make sure the user's typing reaches the
 *   document before another operation runs.
 */
import { computed, nextTick, ref, toRef, watch, type CSSProperties } from "vue";
import ChartOverlay from "@/features/charts/components/ChartOverlay.vue";
import { cssFontFamily, effectiveHorizontalAlign, indentPaddingRem, mergeCellStyle } from "@/features/grid/lib/cellFormatting";
import { getCell, type GridProjection } from "@/features/grid/lib/gridProjection";
import {
  type Cell,
  type CellBorder,
  type CellBorderSide,
  type CellId,
  type ChartId,
  type ColumnId,
  type RowId,
  type TwoCellAnchor,
  type Worksheet,
} from "@/features/document/lib/teamgridDocument";
import { useInlineCellEditor } from "@/features/grid/composables/useInlineCellEditor";
import { useGridSelectionGestures } from "@/features/grid/composables/useGridSelectionGestures";
import {
  MIN_COLUMN_WIDTH,
  useColumnRowResize,
} from "@/features/grid/composables/useColumnRowResize";
import { useGridClipboardBridge } from "@/features/grid/composables/useGridClipboardBridge";
import type { CellSelectionRange } from "@/features/grid/composables/useSelection";
import { createSingleWorksheetFormulaContext, type FormulaContext } from "@/features/formulas/lib";

/**
 * Coordinate range used to draw the clipboard source marquee ("marching
 * ants"). Indexes into the live projection rather than stable ids so the
 * marquee follows the current visible layout - if the user inserts a row
 * above the source, the marquee shifts with it.
 */
export interface GridClipboardRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

const props = withDefaults(defineProps<{
  worksheet: Worksheet;
  formulaContext?: FormulaContext | null;
  projection: GridProjection;
  selectedCellId: string | null;
  selectedRange: CellSelectionRange | null;
  /**
   * Disjoint extra ranges accumulated by Ctrl/Meta+click multi-selection.
   * Defaults to an empty array so consumers that do not (yet) plumb the
   * multi-range state through still work unchanged.
   */
  additionalRanges?: CellSelectionRange[];
  clipboardRange: GridClipboardRange | null;
  highlightedCellIds: string[];
  selectedChartId?: ChartId | null;
  readonly: boolean;
  locale: string;
}>(), {
  additionalRanges: () => [],
  selectedChartId: null,
});

const emit = defineEmits<{
  select: [cell: Cell, address: string];
  "select-range": [range: CellSelectionRange];
  "add-range": [range: CellSelectionRange];
  "clear-additional-ranges": [];
  "set-additional-ranges": [ranges: CellSelectionRange[]];
  commit: [cell: Cell, rawValue: string];
  "request-help": [payload: { anchorEl: HTMLElement; draft: string; caretPos: number }];
  "edit-state": [payload: { editing: boolean; draft: string }];
  "cell-context": [payload: { event: MouseEvent; cell: Cell; address: string; range: CellSelectionRange }];
  "clipboard-copy": [payload: { range: CellSelectionRange | null; event: ClipboardEvent }];
  "clipboard-cut": [payload: { range: CellSelectionRange | null; event: ClipboardEvent }];
  "clipboard-paste": [payload: { event: ClipboardEvent }];
  "clipboard-clear": [];
  "resize-column": [payload: { columnId: ColumnId; width: number }];
  "resize-row": [payload: { rowId: RowId; height: number }];
  "select-chart": [chartId: ChartId | null];
  "edit-chart": [chartId: ChartId];
  "chart-context": [payload: { event: MouseEvent; chartId: ChartId }];
  "resize-chart": [payload: { chartId: ChartId; anchor: TwoCellAnchor }];
  "delete-chart": [chartId: ChartId];
  /**
   * Delete/Backspace pressed while more than one cell is selected.
   * The parent is expected to clear every selected cell in a single
   * `updateGrid` mutation.
   */
  "clear-selection": [];
}>();

const BORDER_SIDES: CellBorderSide[] = ["top", "right", "bottom", "left"];

const gridViewport = ref<HTMLElement | null>(null);

const worksheetRef = toRef(props, "worksheet");
const formulaContextRef = computed(() => props.formulaContext ?? createSingleWorksheetFormulaContext(props.worksheet));
const projectionRef = toRef(props, "projection");
const selectedCellIdRef = toRef(props, "selectedCellId");
const selectedRangeRef = toRef(props, "selectedRange");
const additionalRangesRef = toRef(props, "additionalRanges");
const localeRef = toRef(props, "locale");
const readonlyRef = toRef(props, "readonly");

const editor = useInlineCellEditor({
  worksheet: worksheetRef,
  formulaContext: formulaContextRef,
  projection: projectionRef,
  selectedCellId: selectedCellIdRef,
  locale: localeRef,
  readonly: readonlyRef,
  viewportEl: gridViewport,
  onCommit: (cell, rawValue) => emit("commit", cell, rawValue),
  onEditState: (payload) => emit("edit-state", payload),
  onRequestHelp: (payload) => emit("request-help", payload),
  onSelect: (cell, address) => emit("select", cell, address),
  onSelectRange: (range) => emit("select-range", range),
});

const {
  editingCellId,
  editDraft,
  editorInputEl,
  displayCell,
  startEditing,
  commitEditFromBlur,
  handleEditorKeydown,
  updateEditDraft,
  applyFormulaAssistSuggestion,
  flushPendingEdit,
  appendPickedCellToInlineFormula,
  commitEditingExternally,
} = editor;

const {
  selectedRangeIds,
  startRangeSelection,
  openCellContextMenu,
  selectWholeRow,
  selectWholeColumn,
  extendRangeSelection,
  handleEditKey,
  isWholeRowSelected,
  isWholeColumnSelected,
} = useGridSelectionGestures({
  worksheet: worksheetRef,
  projection: projectionRef,
  selectedCellId: selectedCellIdRef,
  selectedRange: selectedRangeRef,
  additionalRanges: additionalRangesRef,
  readonly: readonlyRef,
  editingCellId,
  editDraft,
  viewportEl: gridViewport,
  appendPickedCellToInlineFormula,
  commitEditingExternally,
  startEditing,
  onSelect: (cell, address) => emit("select", cell, address),
  onSelectRange: (range) => emit("select-range", range),
  onAddRange: (range) => emit("add-range", range),
  onClearAdditionalRanges: () => emit("clear-additional-ranges"),
  onSetAdditionalRanges: (ranges) => emit("set-additional-ranges", ranges),
  onCellContext: (payload) => emit("cell-context", payload),
  onClearSelectedCells: () => emit("clear-selection"),
});

const {
  startColumnResize,
  startRowResize,
  columnPixelWidth,
  rowPixelHeight,
} = useColumnRowResize({
  readonly: readonlyRef,
  onColumnResize: (payload) => emit("resize-column", payload),
  onRowResize: (payload) => emit("resize-row", payload),
});

const {
  handleViewportCopy,
  handleViewportCut,
  handleViewportPaste,
  handleViewportKeydown,
} = useGridClipboardBridge({
  readonly: readonlyRef,
  selectedRange: selectedRangeRef,
  editingCellId,
  onCopy: (payload) => emit("clipboard-copy", payload),
  onCut: (payload) => emit("clipboard-cut", payload),
  onPaste: (payload) => emit("clipboard-paste", payload),
  onClear: () => emit("clipboard-clear"),
});

const highlighted = computed(() => new Set(props.highlightedCellIds));

/**
 * Forward keystrokes into the inline editor draft and grow the
 * textarea to fit its content, so Alt+Enter line breaks expand the
 * cell vertically the way Excel does instead of being clipped by the
 * default single-row height.
 */
function onEditorInput(event: Event, cell: Cell) {
  const target = event.target as HTMLTextAreaElement;
  updateEditDraft(target.value);
  autoGrowEditor(target);
  void cell;
}

/** Resize a textarea so its rendered height matches its content. */
function autoGrowEditor(target: HTMLTextAreaElement | HTMLInputElement | null) {
  if (!(target instanceof HTMLTextAreaElement)) {
    return;
  }
  target.style.height = "auto";
  target.style.height = `${target.scrollHeight}px`;
}

// When the editor mounts (i.e. the user starts editing) the textarea
// is initialised with the draft string; we measure its scrollHeight on
// the next tick so multi-line drafts loaded from a date or formula
// source open at their natural height instead of a single-row stub.
watch(editingCellId, async (next) => {
  if (!next) {
    return;
  }
  await nextTick();
  const el = Array.isArray(editorInputEl.value) ? editorInputEl.value[0] : editorInputEl.value;
  autoGrowEditor(el ?? null);
});

// Style derivation: cell styling cascades row default -> column default
// -> per-cell override (see `mergeCellStyle`). Keeping that merge here
// means the grid never has to flatten styles into every cell at write
// time, which would defeat the granular CRDT save model.

function cellStyle(cell: Cell) {
  const mergedStyle = mergeCellStyle(props.worksheet.rowsById[cell.rowId], props.worksheet.columnsById[cell.columnId], cell);
  const resolvedHorizontal = effectiveHorizontalAlign(cell, mergedStyle.horizontalAlign);
  return {
    color: mergedStyle.textColor,
    backgroundColor: mergedStyle.backgroundColor,
    fontFamily: cssFontFamily(mergedStyle.fontFamily),
    fontSize: mergedStyle.fontSize ? `${mergedStyle.fontSize}px` : undefined,
    fontWeight: mergedStyle.bold ? "700" : undefined,
    fontStyle: mergedStyle.italic ? "italic" : undefined,
    textDecoration: mergedStyle.underline ? "underline" : undefined,
    textAlign: resolvedHorizontal,
    verticalAlign: mergedStyle.verticalAlign,
  };
}

function cellBorderOverlayStyle(cell: Cell, side: CellBorderSide) {
  const mergedStyle = mergeCellStyle(props.worksheet.rowsById[cell.rowId], props.worksheet.columnsById[cell.columnId], cell);
  const border = cssBorder(mergedStyle.borders?.[side]);
  if (!border) {
    return { display: "none" };
  }
  switch (side) {
    case "top":
      return { borderTop: border };
    case "right":
      return { borderRight: border };
    case "bottom":
      return { borderBottom: border };
    case "left":
      return { borderLeft: border };
  }
}

function cssBorder(border: CellBorder | undefined) {
  if (!border) {
    return undefined;
  }
  const width = border.style === "thick"
    ? "3px"
    : border.style === "medium" || border.style === "double"
      ? "2px"
      : "1px";
  const lineStyle = border.style === "dashed" || border.style === "dotted" || border.style === "double"
    ? border.style
    : "solid";
  return `${width} ${lineStyle} ${border.color ?? "currentColor"}`;
}

function cellDisplayStyle(cell: Cell): CSSProperties {
  const mergedStyle = mergeCellStyle(props.worksheet.rowsById[cell.rowId], props.worksheet.columnsById[cell.columnId], cell);
  const resolvedHorizontal = effectiveHorizontalAlign(cell, mergedStyle.horizontalAlign);
  const wrap = Boolean(mergedStyle.wrapText);
  const indent = indentPaddingRem(mergedStyle.indent);
  // Wrapping cells need their own column width so the line break point
  // is deterministic. Non-wrapping cells keep the Excel-style "spill
  // into the empty neighbour" behaviour driven by `cellOverflowWidth`.
  const ownWidth = props.worksheet.columnsById[cell.columnId]?.width ?? MIN_COLUMN_WIDTH;
  const width = wrap || resolvedHorizontal !== "left" ? ownWidth : cellOverflowWidth(cell.rowId, cell.columnId);
  return {
    width: `${width}px`,
    justifyContent: horizontalFlexAlignment(resolvedHorizontal),
    alignItems: verticalFlexAlignment(mergedStyle.verticalAlign),
    whiteSpace: wrap ? "pre-wrap" : "pre",
    wordBreak: wrap ? "break-word" : undefined,
    paddingLeft: resolvedHorizontal === "left" && indent ? `calc(0.45rem + ${indent}rem)` : undefined,
    paddingRight: resolvedHorizontal === "right" && indent ? `calc(0.45rem + ${indent}rem)` : undefined,
  };
}

function columnHeaderStyle(column: GridProjection["columns"][number]) {
  const width = columnPixelWidth(column);
  return { width: `${width}px`, minWidth: `${width}px` };
}

function rowStyle(row: GridProjection["rows"][number]) {
  return { height: `${rowPixelHeight(row)}px` };
}

/**
 * Excel-style text overflow: when a cell's text is wider than its column
 * the value span extends into adjacent columns as long as those columns
 * have empty cells in the same row. The first non-empty neighbour stops
 * the spill.
 */
function cellOverflowWidth(rowId: RowId, columnId: ColumnId) {
  const startIndex = props.projection.columnIndexById.get(columnId);
  if (startIndex == null) {
    return props.worksheet.columnsById[columnId]?.width ?? MIN_COLUMN_WIDTH;
  }
  let width = 0;
  for (let index = startIndex; index < props.projection.columns.length; index += 1) {
    const column = props.projection.columns[index];
    if (!column) {
      break;
    }
    if (index !== startIndex && !isEmptyCell(getCell(props.worksheet, rowId, column.id))) {
      break;
    }
    width += columnPixelWidth(column);
  }
  return width;
}

function isEmptyCell(cell: Cell) {
  return cell.value.kind === "empty" && !cell.formula;
}

function horizontalFlexAlignment(align: ReturnType<typeof mergeCellStyle>["horizontalAlign"]) {
  if (align === "center") {
    return "center";
  }
  if (align === "right") {
    return "flex-end";
  }
  return "flex-start";
}

function verticalFlexAlignment(align: ReturnType<typeof mergeCellStyle>["verticalAlign"]) {
  if (align === "top") {
    return "flex-start";
  }
  if (align === "bottom") {
    return "flex-end";
  }
  return "center";
}

/**
 * Whether a cell falls inside the current "marching ants" clipboard
 * marquee. Endpoints are normalized so the marquee draws correctly even
 * if the source range was selected in any direction.
 */
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

defineExpose({ applyFormulaAssistSuggestion, flushPendingEdit });

// `CellId` is re-exported here purely for backward compatibility with
// callers that imported the type from this file before the composable
// split. Strip later if no external consumer relies on it.
export type { CellId };
</script>

<!--
  Template structure
  ──────────────────
  The grid is a single semantic <table>:
    - <thead> holds the top-left corner spacer and the column headers
      (each with a resize handle at the right edge).
    - <tbody> is one <tr> per visible row. Each row starts with a
      sticky <th> row header (with a row-resize handle) and one <td>
      per visible column.
    - The active cell either renders an `<input class="grid-cell__editor">`
      or a `<span class="grid-cell__value">`, never both.
  The viewport <div> owns the native clipboard listeners so paste/copy
  shortcuts work no matter which cell has focus.
-->
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
            :style="columnHeaderStyle(column)"
            @mousedown.prevent="selectWholeColumn($event, column.id)"
          >
            <span class="grid-axis-header__label">{{ column.label }}</span>
            <span
              v-if="!readonly"
              class="grid-resize-handle grid-resize-handle--column"
              role="separator"
              :aria-label="`Resize column ${column.label}`"
              aria-orientation="vertical"
              @pointerdown="startColumnResize($event, column.id, columnPixelWidth(column))"
            />
          </th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="row in projection.rows" :key="row.id" :style="rowStyle(row)">
          <th
            class="grid-row-header"
            :class="{ 'grid-axis-header--selected': isWholeRowSelected(row.id) }"
            scope="row"
            @mousedown.prevent="selectWholeRow($event, row.id)"
          >
            <span class="grid-axis-header__label">{{ row.label }}</span>
            <span
              v-if="!readonly"
              class="grid-resize-handle grid-resize-handle--row"
              role="separator"
              :aria-label="`Resize row ${row.label}`"
              aria-orientation="horizontal"
              @pointerdown="startRowResize($event, row.id, rowPixelHeight(row))"
            />
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
            :data-cell-id="getCell(worksheet, row.id, column.id).id"
            tabindex="0"
            @mousedown.self.prevent="startRangeSelection($event, row.id, column.id)"
            @contextmenu.prevent="openCellContextMenu($event, row.id, column.id)"
            @mouseenter="extendRangeSelection(row.id, column.id)"
            @dblclick="startEditing(row.id, column.id)"
            @keydown="handleEditKey($event, row.id, column.id)"
          >
            <span
              v-for="side in BORDER_SIDES"
              :key="side"
              class="grid-cell__border-overlay"
              :class="`grid-cell__border-overlay--${side}`"
              :style="cellBorderOverlayStyle(getCell(worksheet, row.id, column.id), side)"
              aria-hidden="true"
            />
            <textarea
              v-if="editingCellId === getCell(worksheet, row.id, column.id).id"
              ref="editorInputEl"
              v-model="editDraft"
              class="grid-cell__editor"
              autofocus
              rows="1"
              wrap="soft"
              spellcheck="false"
              @input="onEditorInput($event, getCell(worksheet, row.id, column.id))"
              @keydown="handleEditorKeydown($event, getCell(worksheet, row.id, column.id))"
              @blur="commitEditFromBlur(getCell(worksheet, row.id, column.id))"
            />
            <span
              v-else
              class="grid-cell__value"
              :style="cellDisplayStyle(getCell(worksheet, row.id, column.id))"
            >
              {{ displayCell(getCell(worksheet, row.id, column.id)) }}
            </span>
          </td>
        </tr>
      </tbody>
    </table>
    <ChartOverlay
      v-if="formulaContextRef"
      :worksheet="worksheet"
      :projection="projection"
      :formula-context="formulaContextRef"
      :selected-chart-id="selectedChartId"
      :readonly="readonly"
      @select="emit('select-chart', $event)"
      @edit="emit('edit-chart', $event)"
      @chart-context="emit('chart-context', $event)"
      @resize-chart="emit('resize-chart', $event)"
      @delete-chart="emit('delete-chart', $event)"
    />
  </div>
</template>

<style scoped>
.grid-viewport {
  overflow: auto;
  min-height: 0;
  flex: 1 1 auto;
  position: relative;
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

.grid-axis-header__label {
  pointer-events: none;
}

.grid-resize-handle {
  position: absolute;
  z-index: 4;
  background: transparent;
}

.grid-resize-handle:hover {
  background: rgb(212 160 23 / 0.28);
}

.grid-resize-handle--column {
  top: 0;
  right: -3px;
  bottom: 0;
  width: 6px;
  cursor: col-resize;
}

.grid-resize-handle--row {
  right: 0;
  bottom: -3px;
  left: 0;
  height: 6px;
  cursor: row-resize;
}

.grid-axis-header--selected {
  background: rgb(212 160 23 / 0.18);
  color: var(--text);
}

.grid-cell {
  position: relative;
  height: 2rem;
  padding: 0;
  overflow: visible;
  border: 1px solid var(--grid-border);
  background: var(--grid-cell-bg);
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

.grid-cell__border-overlay {
  position: absolute;
  z-index: 3;
  pointer-events: none;
}

.grid-cell__border-overlay--top {
  top: 0;
  right: 0;
  left: 0;
}

.grid-cell__border-overlay--right {
  top: 0;
  right: 0;
  bottom: 0;
}

.grid-cell__border-overlay--bottom {
  right: 0;
  bottom: 0;
  left: 0;
}

.grid-cell__border-overlay--left {
  top: 0;
  bottom: 0;
  left: 0;
}

.grid-cell__editor {
  /*
   * The editor is a `<textarea>` that auto-grows vertically as the user
   * presses Alt+Enter to insert a newline. We start at exactly one row's
   * height and let the JS `autoGrowEditor` helper bump `style.height` to
   * `scrollHeight` so additional lines push the editor downward over the
   * grid (the cell's `overflow: visible` allows this), the same way
   * Excel expands the in-cell editor while typing.
   */
  position: relative;
  z-index: 2;
  display: block;
  width: 100%;
  min-height: 100%;
  margin: 0;
  padding: 0.25rem 0.45rem;
  overflow: hidden;
  border: 0;
  outline: 0;
  background: var(--grid-cell-bg);
  resize: none;
  color: inherit;
  font: inherit;
  line-height: 1.3;
  /*
   * Default to soft-wrap so long single-line drafts wrap into the
   * editor instead of producing a horizontal scrollbar. Alt+Enter
   * newlines are honoured because `pre-wrap` preserves explicit `\n`.
   */
  white-space: pre-wrap;
  word-break: break-word;
}

.grid-cell__value {
  position: absolute;
  inset: 0 auto 0 0;
  z-index: 1;
  display: flex;
  box-sizing: border-box;
  padding: 0.25rem 0.45rem;
  overflow: hidden;
  /*
   * Preserve explicit newlines (Alt+Enter) without wrapping on plain
   * spaces. This matches Excel's "Wrap text: off" behaviour: multi-line
   * cells show each `\n`-separated segment on its own line as far as
   * the row height allows, while ordinary long values still spill into
   * neighbouring empty cells via {@link cellOverflowWidth}.
   */
  white-space: pre;
  pointer-events: none;
}

.grid-cell--clipboard-source {
  background-image:
    linear-gradient(90deg, var(--accent) 50%, transparent 50%),
    linear-gradient(90deg, var(--accent) 50%, transparent 50%),
    linear-gradient(0deg, var(--accent) 50%, transparent 50%),
    linear-gradient(0deg, var(--accent) 50%, transparent 50%);
  background-position: 0 0, 0 100%, 0 0, 100% 0;
  background-repeat: repeat-x, repeat-x, repeat-y, repeat-y;
  background-size: 8px 2px, 8px 2px, 2px 8px, 2px 8px;
  animation: teamgrid-clipboard-marquee 0.65s linear infinite;
}

@media (prefers-reduced-motion: reduce) {
  .grid-cell--clipboard-source {
    animation: none;
    outline: 1px dashed var(--accent);
    outline-offset: -2px;
  }
}
</style>
