/**
 * Inline cell editor — opens an `<input>` over the active cell, mirrors
 * its draft into the parent so the formula bar can stay in sync, and
 * handles the various commit / cancel pathways (Enter, blur, Escape,
 * Ctrl+Space for formula assist).
 *
 * Mid-edit clicks on other cells are routed through
 * {@link UseInlineCellEditorReturn.appendPickedCellToInlineFormula} so
 * users can build formulas by clicking references instead of typing them.
 *
 * Enter / Tab also advance the active cell (down or right respectively,
 * up / left with Shift) after committing, matching Excel and Google
 * Sheets behavior.
 */
import { nextTick, ref, watch, type Ref } from "vue";
import { formatCellValue, formatFormulaResult } from "@/features/grid/lib/cellFormatting";
import { evaluateFormula } from "@/features/formulas/lib";
import { insertFunctionAtCaret } from "@/features/formulas/lib/assist";
import { getCell, getCellAddress, type GridProjection } from "@/features/grid/lib/gridProjection";
import type { Cell, CellId, ColumnId, RowId, Worksheet } from "@/features/document/lib/teamgridDocument";
import type { FunctionDefinition } from "@/features/formulas/lib";
import type { CellSelectionRange } from "@/features/grid/composables/useSelection";

export interface UseInlineCellEditorOptions {
  worksheet: Readonly<Ref<Worksheet>>;
  projection: Readonly<Ref<GridProjection>>;
  selectedCellId: Readonly<Ref<string | null>>;
  locale: Readonly<Ref<string>>;
  readonly: Readonly<Ref<boolean>>;
  /**
   * Grid viewport element used to locate the `<td>` of the cell we
   * want to focus after Enter/Tab navigation. Optional so existing
   * callers that don't yet plumb the ref through keep working; in
   * that case the new cell is selected but not DOM-focused.
   */
  viewportEl?: Readonly<Ref<HTMLElement | null>>;
  onCommit: (cell: Cell, rawValue: string) => void;
  onEditState: (payload: { editing: boolean; draft: string }) => void;
  onRequestHelp: (payload: { anchorEl: HTMLElement; draft: string; caretPos: number }) => void;
  onSelect: (cell: Cell, address: string) => void;
  /**
   * Collapse the selection range to the post-Enter/Tab cell. Optional
   * so tests can mount the editor without wiring range emission.
   */
  onSelectRange?: (range: CellSelectionRange) => void;
}

export function useInlineCellEditor(options: UseInlineCellEditorOptions) {
  const editingCellId = ref<string | null>(null);
  const editDraft = ref("");
  const editorInputEl = ref<HTMLInputElement | HTMLInputElement[] | null>(null);
  /**
   * Guard that suppresses one upcoming `blur` commit.
   *
   * When the user presses Enter or clicks another cell mid-edit, we want
   * to call {@link commitEdit} exactly once. Both the keydown handler
   * and the subsequent blur fire, which would emit `commit` twice. Setting
   * this flag before the explicit commit and clearing it on the next
   * microtask means the blur handler sees the guard and bails.
   */
  let suppressNextBlurCommit = false;

  /**
   * Render a cell's display string. Formulas are evaluated each render
   * (results are cached upstream by the formula engine) so dependent
   * cells always show fresh values without an explicit re-evaluation
   * step.
   */
  function displayCell(cell: Cell) {
    if (cell.formula) {
      return formatFormulaResult(
        evaluateFormula(cell.formula.source, options.worksheet.value, options.projection.value).result,
        options.locale.value,
        cell.value,
      );
    }
    return formatCellValue(cell.value, options.locale.value);
  }

  function selectCellByCoordinates(rowId: RowId, columnId: ColumnId) {
    const cell = getCell(options.worksheet.value, rowId, columnId);
    options.onSelect(cell, getCellAddress(options.projection.value, rowId, columnId));
  }

  /**
   * Open the inline editor for a cell.
   *
   * The starting draft depends on how editing was triggered:
   * - `initialValue` provided -> user started typing (replace mode) or
   *   hit Backspace/Delete (empty draft).
   * - cell has a formula -> show its source so the user can edit
   *   `=A1+B1`, not the rendered number.
   * - otherwise -> show the formatted display value so editing a date
   *   does not reveal the underlying ISO string.
   */
  async function startEditing(rowId: RowId, columnId: ColumnId, initialValue?: string) {
    if (options.readonly.value) {
      return;
    }
    const cell = getCell(options.worksheet.value, rowId, columnId);
    editingCellId.value = cell.id;
    editDraft.value = initialValue ?? cell.formula?.source ?? displayCell(cell);
    options.onEditState({ editing: true, draft: editDraft.value });
    selectCellByCoordinates(rowId, columnId);
    await nextTick();
    getEditorInputEl()?.focus();
  }

  function commitEdit(cell: Cell) {
    options.onCommit(cell, editDraft.value);
    editingCellId.value = null;
    options.onEditState({ editing: false, draft: "" });
  }

  /**
   * Variant called when the commit is triggered by a key press
   * (typically Enter / Tab). We commit explicitly and force-blur the
   * input so focus returns to the cell `<td>`; the blur guard prevents
   * the resulting `blur` event from firing a duplicate commit.
   *
   * When `delta` is non-zero the active cell advances after the commit
   * (Enter -> down, Tab -> right, with Shift inverting the direction)
   * and DOM focus moves to the new `<td>`, matching the spreadsheet
   * navigation conventions in Excel / Sheets.
   */
  function commitEditFromKeyboard(
    event: KeyboardEvent,
    cell: Cell,
    delta?: { rows: number; cols: number },
  ) {
    suppressNextBlurCommit = true;
    commitEdit(cell);
    (event.currentTarget as HTMLInputElement | null)?.blur();
    queueMicrotask(() => {
      suppressNextBlurCommit = false;
    });
    if (delta && (delta.rows !== 0 || delta.cols !== 0)) {
      void moveActiveCellRelativeTo(cell, delta);
    }
  }

  /**
   * Advance the active cell by `delta` rows/columns relative to
   * `originCell`, clamping at the edges of the visible projection. The
   * new cell is emitted via {@link UseInlineCellEditorOptions.onSelect}
   * and {@link UseInlineCellEditorOptions.onSelectRange} (so the range
   * collapses to the new active cell), and the matching `<td>` is
   * focused on the next render tick.
   */
  async function moveActiveCellRelativeTo(originCell: Cell, delta: { rows: number; cols: number }) {
    const rowIndex = options.projection.value.rowIndexById.get(originCell.rowId);
    const columnIndex = options.projection.value.columnIndexById.get(originCell.columnId);
    if (rowIndex == null || columnIndex == null) {
      return;
    }
    const targetRow = options.projection.value.rows[clampIndex(rowIndex + delta.rows, options.projection.value.rows.length)];
    const targetColumn = options.projection.value.columns[clampIndex(columnIndex + delta.cols, options.projection.value.columns.length)];
    if (!targetRow || !targetColumn) {
      return;
    }
    const targetCell = getCell(options.worksheet.value, targetRow.id, targetColumn.id);
    options.onSelect(targetCell, getCellAddress(options.projection.value, targetRow.id, targetColumn.id));
    options.onSelectRange?.({ startCellId: targetCell.id, endCellId: targetCell.id });
    await nextTick();
    focusCellById(targetCell.id);
  }

  function focusCellById(cellId: CellId) {
    const cells = options.viewportEl?.value?.querySelectorAll<HTMLElement>("[data-cell-id]") ?? [];
    for (const candidate of cells) {
      if (candidate.dataset.cellId === cellId) {
        candidate.focus();
        return;
      }
    }
  }

  /**
   * Default blur behavior: commit the draft. The guard exists because
   * keyboard / click commits also force a blur and we must not
   * double-emit.
   */
  function commitEditFromBlur(cell: Cell) {
    if (suppressNextBlurCommit) {
      return;
    }
    commitEdit(cell);
  }

  function handleEditorKeydown(event: KeyboardEvent, cell: Cell) {
    if ((event.ctrlKey || event.metaKey) && event.code === "Space") {
      event.preventDefault();
      event.stopPropagation();
      requestFormulaHelp();
      return;
    }
    if (event.key === "Tab") {
      // Tab commits and advances to the cell to the right
      // (Shift+Tab moves left), matching Excel / Sheets and
      // preventing the browser's default focus traversal.
      event.preventDefault();
      event.stopPropagation();
      commitEditFromKeyboard(event, cell, { rows: 0, cols: event.shiftKey ? -1 : 1 });
      return;
    }
    if (event.key === "Enter" || event.code === "NumpadEnter" || event.keyCode === 13) {
      // Enter commits and advances one row down (Shift+Enter moves
      // up). Plain Enter without movement only on the last visible
      // row is handled implicitly by the clamping in
      // `moveActiveCellRelativeTo`.
      event.preventDefault();
      event.stopPropagation();
      commitEditFromKeyboard(event, cell, { rows: event.shiftKey ? -1 : 1, cols: 0 });
      return;
    }
    if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      editingCellId.value = null;
      options.onEditState({ editing: false, draft: "" });
    }
  }

  function updateEditDraft(value: string) {
    editDraft.value = value;
    options.onEditState({ editing: Boolean(editingCellId.value), draft: editDraft.value });
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

  /**
   * Open the formula assist panel anchored on the inline editor input.
   * No-op if we are not actually inside a formula draft (`=...`), so an
   * accidental Ctrl+Space inside a plain number cell does nothing.
   */
  function requestFormulaHelp() {
    const input = getEditorInputEl();
    if (options.readonly.value || !input || !editDraft.value.trim().startsWith("=")) {
      return;
    }
    options.onRequestHelp({
      anchorEl: input,
      draft: editDraft.value,
      caretPos: input.selectionStart ?? editDraft.value.length,
    });
  }

  /**
   * Synchronously commit any pending inline edit. Used by toolbar actions
   * (Save, file commands, undo/redo) to flush the user's unsaved typing
   * before they run. Returns `true` if a commit was emitted.
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

  /**
   * Insert the picked cell's `A1` address into the in-flight formula
   * draft. Simple heuristic: append a `+` between adjacent references so
   * `=A1` followed by clicking B1 becomes `=A1+B1`. Clicking the
   * currently-being-edited cell is rejected to avoid `=A1+A1` accidents.
   */
  function appendPickedCellToInlineFormula(cell: Cell) {
    if (!editDraft.value.trim().startsWith("=") || cell.id === editingCellId.value) {
      return;
    }
    const address = getCellAddress(options.projection.value, cell.rowId, cell.columnId);
    editDraft.value = /[\w)]$/.test(editDraft.value.trimEnd())
      ? `${editDraft.value}+${address}`
      : `${editDraft.value}${address}`;
  }

  function findEditingCell() {
    if (!editingCellId.value) {
      return null;
    }
    for (const row of options.projection.value.rows) {
      for (const column of options.projection.value.columns) {
        const cell = getCell(options.worksheet.value, row.id, column.id);
        if (cell.id === editingCellId.value) {
          return cell;
        }
      }
    }
    return null;
  }

  function commitEditingExternally() {
    const editingCell = findEditingCell();
    if (!editingCell) {
      return;
    }
    suppressNextBlurCommit = true;
    commitEdit(editingCell);
    queueMicrotask(() => {
      suppressNextBlurCommit = false;
    });
  }

  function getEditorInputEl() {
    return Array.isArray(editorInputEl.value) ? editorInputEl.value[0] : editorInputEl.value;
  }

  /**
   * If the parent moves the active cell while we are editing a different
   * cell (e.g. after a programmatic `setSelection` call from the
   * toolbar), close the inline editor.
   */
  watch(
    () => options.selectedCellId.value,
    () => {
      if (editingCellId.value && editingCellId.value !== options.selectedCellId.value) {
        editingCellId.value = null;
      }
    },
  );

  return {
    editingCellId,
    editDraft,
    editorInputEl,
    displayCell,
    startEditing,
    commitEdit,
    commitEditFromBlur,
    handleEditorKeydown,
    updateEditDraft,
    applyFormulaAssistSuggestion,
    flushPendingEdit,
    appendPickedCellToInlineFormula,
    findEditingCell,
    commitEditingExternally,
  };
}

function clampIndex(index: number, length: number) {
  return Math.max(0, Math.min(index, length - 1));
}
