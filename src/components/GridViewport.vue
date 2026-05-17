<script setup lang="ts">
/**
 * GridViewport — the spreadsheet surface that renders rows, columns, and
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
 * - It does **not** mutate the document; instead, it emits semantic events
 *   that the parent translates into `TeamGridOperation`s via
 *   `useTeamGridDocument.updateGrid`. The parent records those operations
 *   against the document's `baseHeads` so saves merge cleanly with
 *   concurrent edits.
 * - The only mutable state living here is transient UI state — the active
 *   inline editor, the in-progress drag-range, and the in-progress
 *   row/column resize.
 *
 * # Responsibilities
 *
 * - Render the table of rows and columns from `projection`.
 * - Manage in-cell editing: double-click to edit, start typing to replace,
 *   F2 / Enter to enter edit mode on the keyboard, Enter to commit,
 *   Escape to cancel, blur to commit. The `commit` event carries the raw
 *   text the user typed (with or without a leading `=`); the parent is
 *   responsible for parsing it into a value or a formula.
 * - Drive rectangular range selection via mouse drag and arrow / shift+arrow
 *   keys.
 * - Support "formula picking": while the user is editing a formula, clicking
 *   another cell appends its `A1` address into the in-flight draft instead
 *   of committing the edit.
 * - Implement column / row resize through pointer-event drag handles.
 * - Bridge native clipboard events (`copy` / `cut` / `paste`) up to the
 *   parent so it can speak the TeamGrid + Excel + TSV clipboard payloads
 *   documented in `src/lib/clipboard/`.
 * - Forward header clicks so the parent can switch to row/column-wide
 *   selections.
 *
 * # Props
 *
 * - `worksheet`: the persisted worksheet snapshot rendered.
 * - `projection`: derived view of `worksheet` — only the rows and columns
 *   that should be visible, plus `A1` address lookups and stable-id ↔
 *   index maps used for keyboard navigation and range iteration.
 * - `selectedCellId`: id of the currently active cell, or `null`.
 * - `selectedRange`: inclusive rectangular range, addressed by stable cell
 *   ids. `null` when only a single cell is selected.
 * - `clipboardRange`: visible "marching ants" marquee for the most recent
 *   copy/cut, in projection coordinates.
 * - `highlightedCellIds`: cells the formula bar wants painted as references
 *   while editing a formula.
 * - `readonly`: disables every mutating affordance (editing, resize,
 *   paste, cut). Used for time-travel snapshots and capability-restricted
 *   sessions.
 * - `locale`: BCP-47 locale forwarded to `formatCellValue` /
 *   `formatFormulaResult`.
 *
 * # Emits
 *
 * Selection / editing:
 * - `select(cell, address)`: a cell became the active cell.
 * - `select-range(range)`: a rectangular selection was started, extended,
 *   or completed.
 * - `commit(cell, rawValue)`: the user finished editing a cell. `rawValue`
 *   is the literal draft string; the parent decides whether it parses
 *   into a number, date, string, or formula.
 * - `edit-state({ editing, draft })`: the inline editor opened, closed, or
 *   the draft text changed. The parent uses this to keep the formula bar
 *   synchronized with the in-cell editor.
 * - `request-help({ anchorEl, draft, caretPos })`: the user pressed
 *   ⌘/Ctrl+Space inside a formula draft to open the formula assist panel.
 *
 * Context menu / clipboard:
 * - `cell-context({ event, cell, address, range })`: the user right-clicked
 *   a cell. `range` is the existing selection if the click happened inside
 *   it, otherwise a single-cell range so the menu always operates on the
 *   right target.
 * - `clipboard-copy` / `clipboard-cut` / `clipboard-paste`: native clipboard
 *   events forwarded with the current selection. The parent is responsible
 *   for serializing/deserializing the payloads.
 * - `clipboard-clear`: user pressed Escape outside an inline edit (clear
 *   the marching ants).
 *
 * Resize:
 * - `resize-column({ columnId, width })`: user finished dragging a column
 *   resize handle. Only emitted if the size actually changed.
 * - `resize-row({ rowId, height })`: same, for rows.
 *
 * # Exposed methods
 *
 * The component exposes two methods to its parent via `defineExpose`:
 *
 * - `applyFormulaAssistSuggestion(definition)`: insert the chosen function
 *   from the formula assist panel into the in-cell draft, place the caret
 *   inside the parameter list, and refocus the editor.
 * - `flushPendingEdit()`: synchronously commit any pending inline edit.
 *   Used by toolbar actions (Save, undo/redo, file commands) to make sure
 *   the user's in-flight typing reaches the document before another
 *   operation runs. Returns `true` if a commit was emitted.
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
import { DEFAULT_ROW_HEIGHT } from "@/lib/gridDimensions";
import { createCellId, type Cell, type CellId, type ColumnId, type RowId, type Worksheet } from "@/lib/teamgridDocument";
import type { FunctionDefinition } from "@/lib/formulas";

/**
 * Inclusive rectangular cell range, addressed by stable cell ids.
 *
 * `start` and `end` are not required to be top-left / bottom-right. The
 * range is whatever rectangle covers both ids in the current projection,
 * so a drag from B5 to A1 produces a valid range that iterates A1..B5.
 */
export interface CellSelectionRange {
  startCellId: CellId;
  endCellId: CellId;
}

/**
 * Coordinate range used to draw the clipboard source marquee ("marching
 * ants"). Indexes into the live projection rather than stable ids so the
 * marquee follows the current visible layout — if the user inserts a row
 * above the source, the marquee shifts with it.
 */
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
  "cell-context": [payload: { event: MouseEvent; cell: Cell; address: string; range: CellSelectionRange }];
  "clipboard-copy": [payload: { range: CellSelectionRange | null; event: ClipboardEvent }];
  "clipboard-cut": [payload: { range: CellSelectionRange | null; event: ClipboardEvent }];
  "clipboard-paste": [payload: { event: ClipboardEvent }];
  "clipboard-clear": [];
  "resize-column": [payload: { columnId: ColumnId; width: number }];
  "resize-row": [payload: { rowId: RowId; height: number }];
}>();

// ─── Transient UI state ─────────────────────────────────────────────────
// Nothing in this section is persisted; it lives only as long as the user's
// in-flight gesture (editing, drag-selecting, or dragging a resize handle).

/** Stable id of the cell currently being edited inline, or `null`. */
const editingCellId = ref<string | null>(null);

/**
 * Raw text in the inline editor. Mirrors what the formula bar would show.
 * The parent receives the same string via `edit-state` so the two editors
 * stay synchronized character-by-character.
 */
const editDraft = ref("");

/** The scrollable `<div>` wrapping the table. Used to focus cells imperatively. */
const gridViewport = ref<HTMLElement | null>(null);

/**
 * Vue normally returns either a single ref or an array of refs depending
 * on whether the bound element is rendered inside a `v-for`. The inline
 * editor is rendered conditionally inside the cell loop, so the ref can
 * appear as both shapes; {@link getEditorInputEl} normalizes that.
 */
const editorInputEl = ref<HTMLInputElement | HTMLInputElement[] | null>(null);

/**
 * Anchor cell of an in-progress drag selection. Set on `mousedown`,
 * cleared on global `mouseup`. While set, hovering cells extends the
 * range live.
 */
const draggingRangeStart = ref<CellId | null>(null);

/** In-progress column/row resize drag, if any. */
const activeResize = ref<ResizeDrag | null>(null);

/**
 * Guard that suppresses one upcoming `blur` commit.
 *
 * When the user presses Enter or clicks another cell mid-edit, we want to
 * call {@link commitEdit} exactly once. Both the keydown handler and the
 * subsequent blur fire, which would emit `commit` twice. Setting this flag
 * before the explicit commit and clearing it on the next microtask means
 * the blur handler sees the guard and bails.
 */
let suppressNextBlurCommit = false;

/** Per-key delta applied by {@link handleSelectionKey} for arrow navigation. */
const KEYBOARD_SELECTION_DELTAS: Record<string, { rows: number; cols: number }> = {
  ArrowUp: { rows: -1, cols: 0 },
  ArrowDown: { rows: 1, cols: 0 },
  ArrowLeft: { rows: 0, cols: -1 },
  ArrowRight: { rows: 0, cols: 1 },
};

/** Resize handles cannot drag a column narrower than this. */
const MIN_COLUMN_WIDTH = 48;
/** Resize handles cannot drag a row shorter than this. */
const MIN_ROW_HEIGHT = 24;

/** State for an in-progress resize drag of a single column or row. */
interface ResizeDrag {
  kind: "column" | "row";
  id: ColumnId | RowId;
  /** Pointer position (clientX for columns, clientY for rows) at drag start. */
  startPointer: number;
  /** Size in pixels at drag start; used as the baseline for the delta. */
  startSize: number;
  /**
   * Live size during the drag. Rendered by {@link columnPixelWidth} /
   * {@link rowPixelHeight} so the user sees feedback without firing a
   * `resize-*` event on every pointer move.
   */
  currentSize: number;
}

/** Set lookups for `:class` bindings; recomputed only when inputs change. */
const highlighted = computed(() => new Set(props.highlightedCellIds));
const selectedRangeIds = computed(() => new Set(props.selectedRange ? getRangeCellIds(props.selectedRange) : []));

/**
 * If the parent moves the active cell while we are editing a different
 * cell (e.g. after a programmatic `setSelection` call from the toolbar),
 * close the inline editor. Without this the input would stay open over
 * a non-active cell and look stuck.
 */
watch(
  () => props.selectedCellId,
  () => {
    if (editingCellId.value && editingCellId.value !== props.selectedCellId) {
      editingCellId.value = null;
    }
  },
);

// ─── Lifecycle: window-level listeners ─────────────────────────────────
// Some interactions need to receive events outside the table — keyboard
// navigation when only the cell `<td>` has focus, mouseup that ends a
// drag-selection that started inside the table but ended outside it, and
// pointermove/up for resize drags. These are attached on the window and
// torn down on unmount so they don't leak between component instances.

onMounted(() => {
  window.addEventListener("keydown", handleWindowKeydown);
  window.addEventListener("mouseup", endRangeDrag);
});

onBeforeUnmount(() => {
  window.removeEventListener("keydown", handleWindowKeydown);
  window.removeEventListener("mouseup", endRangeDrag);
  removeResizeListeners();
});

// ─── Display ────────────────────────────────────────────────────────────

/**
 * Render a cell's display string. Formulas are evaluated each render
 * (results are cached upstream by the formula engine) so dependent cells
 * always show fresh values without an explicit re-evaluation step.
 */
function displayCell(cell: Cell) {
  if (cell.formula) {
    return formatFormulaResult(evaluateFormula(cell.formula.source, props.worksheet, props.projection).result, props.locale);
  }
  return formatCellValue(cell.value, props.locale);
}

// ─── Editing ────────────────────────────────────────────────────────────

/** Tell the parent which cell became active. Does not change selection range. */
function selectCell(rowId: RowId, columnId: ColumnId) {
  const cell = getCell(props.worksheet, rowId, columnId);
  emit("select", cell, getCellAddress(props.projection, rowId, columnId));
}

/**
 * Open the inline editor for a cell.
 *
 * The starting draft depends on how editing was triggered:
 * - `initialValue` provided → user started typing (replace mode) or hit
 *   Backspace/Delete (empty draft).
 * - cell has a formula → show its source so the user can edit `=A1+B1`,
 *   not the rendered number.
 * - otherwise → show the formatted display value so editing a date does
 *   not reveal the underlying ISO string.
 *
 * Focus is moved to the input on the next tick because the `<input>` is
 * only rendered when `editingCellId` matches the cell.
 */
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

/**
 * Emit the final draft to the parent and close the inline editor. The
 * parent decides how to interpret the raw text (number vs date vs string
 * vs formula) — see `useTeamGridDocument.commitCellEdit`.
 */
function commitEdit(cell: Cell) {
  emit("commit", cell, editDraft.value);
  editingCellId.value = null;
  emit("edit-state", { editing: false, draft: "" });
}

/**
 * Variant called when the commit is triggered by a key press (typically
 * Enter). We commit explicitly **and** force-blur the input so focus
 * returns to the cell `<td>`; {@link suppressNextBlurCommit} prevents the
 * resulting `blur` event from firing a duplicate commit.
 */
function commitEditFromKeyboard(event: KeyboardEvent, cell: Cell) {
  suppressNextBlurCommit = true;
  commitEdit(cell);
  (event.currentTarget as HTMLInputElement | null)?.blur();
  queueMicrotask(() => {
    suppressNextBlurCommit = false;
  });
}

/**
 * Keys handled while the inline editor has focus.
 *
 * - ⌘/Ctrl+Space → open the formula assist panel (only meaningful inside
 *   a formula draft).
 * - Enter (incl. NumpadEnter and the legacy keyCode 13 for browsers that
 *   still need it) → commit.
 * - Escape → cancel without committing.
 *
 * `stopPropagation` keeps these keys from also reaching the cell-level
 * keydown handler (which would try to start editing again or move the
 * selection).
 */
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

/**
 * Forward each keystroke to the parent so the formula bar mirrors the
 * in-cell editor in real time.
 */
function updateEditDraft(value: string) {
  editDraft.value = value;
  emit("edit-state", { editing: Boolean(editingCellId.value), draft: editDraft.value });
}

/**
 * Insert a function chosen from the formula assist panel into the current
 * inline draft. Exposed via `defineExpose` so the parent can wire the
 * assist panel's `select` event directly to whichever editor is active.
 */
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

/**
 * Open the formula assist panel anchored on the inline editor input.
 * No-op if we are not actually inside a formula draft (`= ...`), so an
 * accidental ⌘/Ctrl+Space inside a plain number cell does nothing.
 */
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

/**
 * Default blur behavior: commit the draft. The guard exists because
 * keyboard / click commits also force a blur and we must not double-emit
 * — see {@link suppressNextBlurCommit}.
 */
function commitEditFromBlur(cell: Cell) {
  if (suppressNextBlurCommit) {
    return;
  }
  commitEdit(cell);
}

/**
 * Synchronously commit any pending inline edit. Exposed to the parent so
 * toolbar actions (Save, file commands, undo/redo) can flush the user's
 * unsaved typing before they run. Returns `true` if a commit was emitted.
 */
function flushPendingEdit() {
  const editingCell = findEditingCell();
  if (!editingCell) {
    return false;
  }
  suppressNextBlurCommit = true;
  commitEdit(editingCell);
  queueMicrotask(() => {
    suppressNextBlurCommit = false;
  });
  return true;
}

// ─── Style derivation ──────────────────────────────────────────────────
// Cell styling cascades row default → column default → per-cell override
// (see `mergeCellStyle`). Keeping that merge here means the grid never
// has to flatten styles into every cell at write time, which would defeat
// the granular CRDT save model.

/**
 * Build the inline CSS for the `<td>` itself. The `<td>` paints the
 * background and contributes the border; the actual text lives in the
 * `<span class="grid-cell__value">` (see {@link cellDisplayStyle}) so that
 * long strings can overflow into adjacent empty cells.
 */
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

/**
 * CSS for the value `<span>` overlaid on a cell. The width is computed
 * by {@link cellOverflowWidth} so unformatted text can spill into empty
 * neighbours (Excel-style overflow), while alignment is mapped from the
 * cell's logical alignment to the corresponding flex alignment.
 */
function cellDisplayStyle(cell: Cell) {
  const mergedStyle = mergeCellStyle(props.worksheet.rowsById[cell.rowId], props.worksheet.columnsById[cell.columnId], cell);
  return {
    width: `${cellOverflowWidth(cell.rowId, cell.columnId)}px`,
    justifyContent: horizontalFlexAlignment(mergedStyle.horizontalAlign),
    alignItems: verticalFlexAlignment(mergedStyle.verticalAlign),
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
 * Width of a column, accounting for an in-progress resize drag. While
 * dragging we render the live `currentSize` so the user sees feedback
 * without writing to the document on every pixel of motion.
 */
function columnPixelWidth(column: GridProjection["columns"][number]) {
  return activeResize.value?.kind === "column" && activeResize.value.id === column.id
    ? activeResize.value.currentSize
    : column.width;
}

/** Mirror of {@link columnPixelWidth} for rows; falls back to {@link DEFAULT_ROW_HEIGHT}. */
function rowPixelHeight(row: GridProjection["rows"][number]) {
  return activeResize.value?.kind === "row" && activeResize.value.id === row.id
    ? activeResize.value.currentSize
    : row.height ?? DEFAULT_ROW_HEIGHT;
}

/**
 * Excel-style text overflow: when a cell's text is wider than its column
 * the value span extends into adjacent columns as long as those columns
 * have empty cells in the same row. The first non-empty neighbour stops
 * the spill.
 *
 * The overflow happens at the visual layer only — it does not move data,
 * does not change the underlying cell, and stops at the rightmost visible
 * column.
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

/**
 * A cell is "empty" only if it has neither a value nor a formula. A cell
 * with a formula whose result is empty is still considered occupied so
 * that text from the left will not bleed across it.
 */
function isEmptyCell(cell: Cell) {
  return cell.value.kind === "empty" && !cell.formula;
}

/** Translate logical horizontal alignment to a flex `justify-content` value. */
function horizontalFlexAlignment(align: ReturnType<typeof mergeCellStyle>["horizontalAlign"]) {
  if (align === "center") {
    return "center";
  }
  if (align === "right") {
    return "flex-end";
  }
  return "flex-start";
}

/** Translate logical vertical alignment to a flex `align-items` value. */
function verticalFlexAlignment(align: ReturnType<typeof mergeCellStyle>["verticalAlign"]) {
  if (align === "top") {
    return "flex-start";
  }
  if (align === "bottom") {
    return "flex-end";
  }
  return "center";
}

// ─── Column / row resize ───────────────────────────────────────────────
// Resize handles use pointer events rather than mouse events so they
// behave consistently with touch and pen input. Window-level listeners
// keep the drag alive even if the cursor moves outside the resize handle
// or outside the grid entirely.

function startColumnResize(event: PointerEvent, columnId: ColumnId, width: number) {
  startResize(event, {
    kind: "column",
    id: columnId,
    startPointer: event.clientX,
    startSize: width,
    currentSize: width,
  });
}

function startRowResize(event: PointerEvent, rowId: RowId, height?: number) {
  const startSize = height ?? DEFAULT_ROW_HEIGHT;
  startResize(event, {
    kind: "row",
    id: rowId,
    startPointer: event.clientY,
    startSize,
    currentSize: startSize,
  });
}

/**
 * Begin a resize drag. We `preventDefault` and `stopPropagation` so the
 * pointer-down does not also bubble up and trigger row/column selection
 * on the underlying header.
 */
function startResize(event: PointerEvent, drag: ResizeDrag) {
  if (props.readonly || event.button !== 0) {
    return;
  }
  event.preventDefault();
  event.stopPropagation();
  activeResize.value = drag;
  window.addEventListener("pointermove", handleResizePointerMove);
  window.addEventListener("pointerup", finishResize);
  window.addEventListener("pointercancel", cancelResize);
}

/**
 * Update the live `currentSize` in response to pointer motion. Clamping
 * to a minimum size keeps users from collapsing a column/row into a state
 * where the resize handle becomes unreachable.
 */
function handleResizePointerMove(event: PointerEvent) {
  const drag = activeResize.value;
  if (!drag) {
    return;
  }
  const pointer = drag.kind === "column" ? event.clientX : event.clientY;
  const minSize = drag.kind === "column" ? MIN_COLUMN_WIDTH : MIN_ROW_HEIGHT;
  drag.currentSize = Math.max(minSize, Math.round(drag.startSize + pointer - drag.startPointer));
}

/**
 * End-of-drag commit. Emit a single `resize-*` event only when the size
 * actually changed so a stray click on the handle does not push a no-op
 * operation through the document save path.
 */
function finishResize() {
  const drag = activeResize.value;
  if (!drag) {
    return;
  }
  if (drag.currentSize !== drag.startSize) {
    if (drag.kind === "column") {
      emit("resize-column", { columnId: drag.id as ColumnId, width: drag.currentSize });
    } else {
      emit("resize-row", { rowId: drag.id as RowId, height: drag.currentSize });
    }
  }
  activeResize.value = null;
  removeResizeListeners();
}

/** Pointer was cancelled (e.g. browser focus loss); discard any preview. */
function cancelResize() {
  activeResize.value = null;
  removeResizeListeners();
}

function removeResizeListeners() {
  window.removeEventListener("pointermove", handleResizePointerMove);
  window.removeEventListener("pointerup", finishResize);
  window.removeEventListener("pointercancel", cancelResize);
}

// ─── Selection & range drag ─────────────────────────────────────────────

/**
 * Move focus to the clicked `<td>` so the keyboard shortcuts handled by
 * {@link handleEditKey} apply to the new active cell.
 */
function handleCellClick(event: MouseEvent, rowId: RowId, columnId: ColumnId) {
  selectCell(rowId, columnId);
  (event.currentTarget as HTMLElement | null)?.focus();
}

/**
 * Mousedown on a cell. Three branches matter:
 *
 * 1. **Formula picking** — if we're mid-edit and the draft starts with
 *    `=`, treat the click as "pick this cell as a reference" instead of
 *    a selection change. This is the primary way users build formulas
 *    in TeamGrid (and Excel).
 * 2. **Mid-edit, non-formula click** — commit the in-flight edit first
 *    (with the blur-commit guard so we don't double-commit), then start
 *    a fresh range selection on the clicked cell.
 * 3. **Default** — start a new range anchored on the clicked cell.
 */
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

/**
 * Right-click handling. If the click happens inside the existing
 * selection we keep that selection so the menu's actions (copy, delete,
 * formatting…) operate on the whole range. Otherwise we collapse the
 * range to the clicked cell so menu actions don't accidentally apply to
 * cells the user didn't intend.
 */
function openCellContextMenu(event: MouseEvent, rowId: RowId, columnId: ColumnId) {
  const cell = getCell(props.worksheet, rowId, columnId);
  const address = getCellAddress(props.projection, rowId, columnId);
  const range = props.selectedRange && selectedRangeIds.value.has(cell.id)
    ? props.selectedRange
    : { startCellId: cell.id, endCellId: cell.id };
  emit("cell-context", { event, cell, address, range });
}

/**
 * Select every cell in a row by clicking the row header. The range spans
 * the first and last visible columns in the projection — we do not
 * include tombstoned columns even if their ids still exist in
 * `columnOrder`.
 */
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

/** Mirror of {@link selectWholeRow} for columns. */
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

/**
 * Insert the picked cell's `A1` address into the in-flight formula draft.
 *
 * The simple heuristic — append a `+` between two adjacent references —
 * keeps formulas like `=A1+B1+C1` typeable purely by clicking, without a
 * full operator-aware editor. Picking the cell that's currently being
 * edited is rejected so you can't create `=A1+A1` accidentally on the
 * mousedown that started the edit.
 */
function appendPickedCellToInlineFormula(cell: Cell) {
  if (!editDraft.value.trim().startsWith("=") || cell.id === editingCellId.value) {
    return;
  }
  const address = getCellAddress(props.projection, cell.rowId, cell.columnId);
  editDraft.value = /[\w)]$/.test(editDraft.value.trimEnd())
    ? `${editDraft.value}+${address}`
    : `${editDraft.value}${address}`;
}

/** Mouseenter while a drag is active extends the range to the hovered cell. */
function extendRangeSelection(rowId: RowId, columnId: ColumnId) {
  if (!draggingRangeStart.value) {
    return;
  }
  const cell = getCell(props.worksheet, rowId, columnId);
  emit("select-range", { startCellId: draggingRangeStart.value, endCellId: cell.id });
}

/** Listener bound on `window` so a drag that ends outside the grid is still cleaned up. */
function endRangeDrag() {
  draggingRangeStart.value = null;
}

// ─── Keyboard handling ──────────────────────────────────────────────────

/**
 * Window-level keydown handler. Lets the user move and edit the active
 * cell even when only the grid (rather than a specific `<td>`) has focus.
 *
 * It deliberately bails when the event target is another editor (formula
 * bar, dialog input, etc.), so typing in the formula bar does not also
 * move the active cell.
 */
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

/**
 * Single keydown router for the active cell. The branching matters:
 *
 * 1. While the inline editor is open, this handler does nothing — the
 *    editor's own keydown handler takes over.
 * 2. Arrow keys move (or extend with shift) the selection regardless of
 *    `readonly`, because navigation is not a mutation.
 * 3. After that we gate on `readonly`: Enter/F2 enters edit mode,
 *    Backspace/Delete enters edit mode with an empty draft (delete-on-
 *    commit), and any printable key enters edit mode using that key as
 *    the first character (replace-the-cell typing).
 */
function handleEditKey(event: KeyboardEvent, rowId: RowId, columnId: ColumnId) {
  if (editingCellId.value) {
    return;
  }
  if (handleSelectionKey(event, rowId, columnId)) {
    return;
  }
  if (props.readonly) {
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

/**
 * Process arrow-key navigation.
 *
 * Returns `true` if it consumed the event (so {@link handleEditKey} can
 * stop processing). Without shift the active cell moves to the neighbour;
 * with shift the existing range's start anchor stays put and only the
 * end is moved, matching Excel's selection-extend behaviour.
 */
function handleSelectionKey(event: KeyboardEvent, rowId: RowId, columnId: ColumnId) {
  const delta = KEYBOARD_SELECTION_DELTAS[event.key];
  if (!delta) {
    return false;
  }
  const currentRowIndex = props.projection.rowIndexById.get(rowId);
  const currentColumnIndex = props.projection.columnIndexById.get(columnId);
  if (currentRowIndex == null || currentColumnIndex == null) {
    return false;
  }

  event.preventDefault();
  const targetRow = props.projection.rows[clampIndex(currentRowIndex + delta.rows, props.projection.rows.length)];
  const targetColumn = props.projection.columns[clampIndex(currentColumnIndex + delta.cols, props.projection.columns.length)];
  if (!targetRow || !targetColumn) {
    return true;
  }

  const targetCell = getCell(props.worksheet, targetRow.id, targetColumn.id);
  const range = event.shiftKey
    ? { startCellId: props.selectedRange?.startCellId ?? createCellId(rowId, columnId), endCellId: targetCell.id }
    : { startCellId: targetCell.id, endCellId: targetCell.id };
  emit("select", targetCell, getCellAddress(props.projection, targetRow.id, targetColumn.id));
  emit("select-range", range);
  void focusCell(targetCell.id);
  return true;
}

/** Clamp `index` to `[0, length - 1]`, i.e. block over- and underflow. */
function clampIndex(index: number, length: number) {
  return Math.max(0, Math.min(index, length - 1));
}

/**
 * Move DOM focus to a cell by id. We do this on the next tick because
 * the parent might still be reacting to the `select-range` emission and
 * have not yet re-rendered the table.
 */
async function focusCell(cellId: CellId) {
  await nextTick();
  const cells = gridViewport.value?.querySelectorAll<HTMLElement>("[data-cell-id]") ?? [];
  for (const cell of cells) {
    if (cell.dataset.cellId === cellId) {
      cell.focus();
      return;
    }
  }
}

// ─── Cell / range lookups ───────────────────────────────────────────────
// These are linear scans of the projection rather than direct map
// lookups. The sample sheet sizes (≈ a few thousand cells) make that
// trivially fast; a real product would maintain `cellId → coordinates`
// indexes. See the README's "Current Scope" — virtualization and
// indexing are deferred.

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

/**
 * Expand a range to the full set of cell ids it covers. The range
 * endpoints are normalized into `(min, max)` first so a "drag from
 * bottom-right to top-left" range still produces the correct rectangle.
 */
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

/** True if the current selection covers every visible column on this row. */
function isWholeRowSelected(rowId: RowId) {
  return props.projection.columns.length > 0
    && props.projection.columns.every((column) => selectedRangeIds.value.has(getCell(props.worksheet, rowId, column.id).id));
}

/** True if the current selection covers every visible row on this column. */
function isWholeColumnSelected(columnId: ColumnId) {
  return props.projection.rows.length > 0
    && props.projection.rows.every((row) => selectedRangeIds.value.has(getCell(props.worksheet, row.id, columnId).id));
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

/**
 * True if the keyboard event originated inside another editable element
 * (formula bar input, modal dialog field, contenteditable…). Used to
 * stop window-level keyboard navigation from stealing keystrokes from
 * those editors.
 */
function isTypingInAnotherEditor(target: EventTarget | null) {
  if (!(target instanceof HTMLElement)) {
    return false;
  }
  return target.matches("input, textarea, select, [contenteditable='true']");
}

/** Normalize the polymorphic ref returned for the `v-for`-rendered editor. */
function getEditorInputEl() {
  return Array.isArray(editorInputEl.value) ? editorInputEl.value[0] : editorInputEl.value;
}

/**
 * Whether a key, if pressed on a non-editing cell, should start an inline
 * edit using that key as the first character. Modifier keys are excluded
 * so shortcuts like ⌘C / Ctrl+S do not turn into typed input.
 */
function isPrintableKey(event: KeyboardEvent) {
  return event.key.length === 1 && !event.metaKey && !event.ctrlKey && !event.altKey;
}

// ─── Native clipboard bridging ─────────────────────────────────────────
// We intercept the browser's native copy/cut/paste events on the grid
// element and forward them to the parent with the current selection.
// `preventDefault` is essential: it lets the parent write its own custom
// payloads (TeamGrid JSON + Excel HTML + TSV) into the OS clipboard
// instead of whatever the browser would put there by default.

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

/**
 * Local Escape handler used to dismiss the clipboard marquee. The Escape
 * for cancelling an inline edit is handled by {@link handleEditorKeydown}
 * before this fires, so we only emit `clipboard-clear` when no edit is
 * active.
 */
function handleViewportKeydown(event: KeyboardEvent) {
  if (event.key === "Escape" && !editingCellId.value) {
    emit("clipboard-clear");
  }
}

defineExpose({ applyFormulaAssistSuggestion, flushPendingEdit });
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
            @mousedown.prevent="selectWholeColumn(column.id)"
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
            @mousedown.prevent="selectWholeRow(row.id)"
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
            @mousedown.prevent="startRangeSelection($event, row.id, column.id)"
            @contextmenu.prevent="openCellContextMenu($event, row.id, column.id)"
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

.grid-cell__editor {
  width: 100%;
  height: 100%;
  padding: 0.25rem 0.45rem;
  border: 0;
  outline: 0;
  background: transparent;
  color: inherit;
  font: inherit;
}

.grid-cell__value {
  position: absolute;
  inset: 0 auto 0 0;
  z-index: 1;
  display: flex;
  box-sizing: border-box;
  padding: 0.25rem 0.45rem;
  overflow: hidden;
  white-space: nowrap;
  pointer-events: none;
}
</style>
