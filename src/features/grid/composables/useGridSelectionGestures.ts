/**
 * Mouse / keyboard gestures that drive the grid's selection.
 *
 * - Drag-selection across cells (mousedown + mouseenter + global mouseup).
 * - Header click to select an entire row or column.
 * - Arrow-key navigation with optional Shift extend.
 * - Right-click context menu plumbing.
 * - Mid-edit clicks: route to formula picking when the draft begins with
 *   `=`, otherwise commit the edit before starting the new range.
 *
 * Window-level listeners are installed on mount so a drag that ends
 * outside the viewport still releases cleanly.
 */
import { computed, onBeforeUnmount, onMounted, nextTick, ref, type Ref } from "vue";
import { getCell, getCellAddress, type GridProjection } from "@/features/grid/lib/gridProjection";
import {
  createCellId,
  type Cell,
  type CellId,
  type ColumnId,
  type RowId,
  type Worksheet,
} from "@/features/document/lib/teamgridDocument";
import type { CellSelectionRange } from "@/features/grid/composables/useSelection";

const KEYBOARD_SELECTION_DELTAS: Record<string, { rows: number; cols: number }> = {
  ArrowUp: { rows: -1, cols: 0 },
  ArrowDown: { rows: 1, cols: 0 },
  ArrowLeft: { rows: 0, cols: -1 },
  ArrowRight: { rows: 0, cols: 1 },
};

export interface UseGridSelectionGesturesOptions {
  worksheet: Readonly<Ref<Worksheet>>;
  projection: Readonly<Ref<GridProjection>>;
  selectedCellId: Readonly<Ref<string | null>>;
  selectedRange: Readonly<Ref<CellSelectionRange | null>>;
  /**
   * Disjoint ranges accumulated by Ctrl/Meta+click multi-selection.
   * Included in {@link selectedRangeIds} so the viewport highlights
   * cells from every range, not just the primary one.
   */
  additionalRanges: Readonly<Ref<CellSelectionRange[]>>;
  readonly: Readonly<Ref<boolean>>;
  editingCellId: Readonly<Ref<string | null>>;
  editDraft: Readonly<Ref<string>>;
  viewportEl: Readonly<Ref<HTMLElement | null>>;
  appendPickedCellToInlineFormula: (cell: Cell) => void;
  commitEditingExternally: () => void;
  startEditing: (rowId: RowId, columnId: ColumnId, initialValue?: string) => Promise<void> | void;
  onSelect: (cell: Cell, address: string) => void;
  onSelectRange: (range: CellSelectionRange) => void;
  /** Push a range onto the disjoint extra-selection list (Ctrl/Meta+click). */
  onAddRange: (range: CellSelectionRange) => void;
  /** Drop all extra Ctrl/Meta+click ranges (default click, keyboard nav, etc.). */
  onClearAdditionalRanges: () => void;
  /**
   * Atomically replace the entire disjoint extra-selection list. Used
   * by Ctrl/Meta+click toggling, which needs to swap the whole list in
   * one go after subtracting the clicked cell from each enclosing
   * rectangle.
   */
  onSetAdditionalRanges: (ranges: CellSelectionRange[]) => void;
  onCellContext: (payload: { event: MouseEvent; cell: Cell; address: string; range: CellSelectionRange }) => void;
}

export function useGridSelectionGestures(options: UseGridSelectionGesturesOptions) {
  const draggingRangeStart = ref<CellId | null>(null);

  /**
   * Union of cell ids covered by the primary selection range and every
   * disjoint Ctrl/Meta+click range. Drives cell highlighting in the
   * viewport so multi-selections look the same as single ranges.
   */
  const selectedRangeIds = computed(() => {
    const ids = new Set<CellId>();
    if (options.selectedRange.value) {
      for (const id of getRangeCellIds(options.selectedRange.value)) {
        ids.add(id);
      }
    }
    for (const range of options.additionalRanges.value) {
      for (const id of getRangeCellIds(range)) {
        ids.add(id);
      }
    }
    return ids;
  });

  function selectCellByCoordinates(rowId: RowId, columnId: ColumnId) {
    const cell = getCell(options.worksheet.value, rowId, columnId);
    options.onSelect(cell, getCellAddress(options.projection.value, rowId, columnId));
  }

  function handleCellClick(event: MouseEvent, rowId: RowId, columnId: ColumnId) {
    selectCellByCoordinates(rowId, columnId);
    (event.currentTarget as HTMLElement | null)?.focus();
  }

  /**
   * Mousedown on a cell. Six branches matter:
   *
   * 1. Formula picking - if we're mid-edit and the draft starts with
   *    `=`, treat the click as "pick this cell as a reference" instead
   *    of a selection change.
   * 2. Mid-edit, non-formula click - commit the in-flight edit first,
   *    then start a fresh range selection on the clicked cell.
   * 3. Shift+click - extend the existing range from its current anchor
   *    (the start of the live range, or the active cell if none) to
   *    the clicked cell. The active cell follows the click, mirroring
   *    the shift+arrow keyboard gesture. Disjoint Ctrl/Meta+click
   *    ranges accumulated so far are preserved.
   * 4. Ctrl/Meta+click on an already-selected cell - subtract the
   *    clicked cell from every enclosing rectangle, splitting each
   *    affected range into up to four sub-rectangles. Refuses to
   *    deselect the very last selected cell. Matches Excel's
   *    "Ctrl+click to deselect" behavior.
   * 5. Ctrl/Meta+click on a non-selected cell - push the current
   *    primary range onto the disjoint extra-range list, then start a
   *    new single-cell primary range at the clicked cell. A subsequent
   *    drag (mouseenter) only extends the new primary, leaving the
   *    older sub-ranges alone, so the user can paint several
   *    non-contiguous rectangles.
   * 6. Default - drop any disjoint sub-ranges and start a new primary
   *    range anchored on the clicked cell.
   */
  function startRangeSelection(event: MouseEvent, rowId: RowId, columnId: ColumnId) {
    if (event.button !== 0) {
      return;
    }
    const cell = getCell(options.worksheet.value, rowId, columnId);
    if (options.editingCellId.value) {
      if (options.editDraft.value.trim().startsWith("=")) {
        options.appendPickedCellToInlineFormula(cell);
        return;
      }
      options.commitEditingExternally();
    }
    if (event.shiftKey && options.selectedCellId.value) {
      const anchorId = options.selectedRange.value?.startCellId ?? options.selectedCellId.value;
      draggingRangeStart.value = anchorId;
      options.onSelectRange({ startCellId: anchorId, endCellId: cell.id });
      handleCellClick(event, rowId, columnId);
      return;
    }
    if ((event.ctrlKey || event.metaKey) && options.selectedCellId.value) {
      if (selectedRangeIds.value.has(cell.id)) {
        deselectCellFromSelection(event, cell);
        return;
      }
      if (options.selectedRange.value) {
        options.onAddRange(options.selectedRange.value);
        draggingRangeStart.value = cell.id;
        options.onSelectRange({ startCellId: cell.id, endCellId: cell.id });
        handleCellClick(event, rowId, columnId);
        return;
      }
    }
    options.onClearAdditionalRanges();
    draggingRangeStart.value = cell.id;
    options.onSelectRange({ startCellId: cell.id, endCellId: cell.id });
    handleCellClick(event, rowId, columnId);
  }

  /**
   * Remove `cell` from the multi-range selection by subtracting it
   * from every enclosing rectangle. Each affected range is split into
   * up to four edge sub-rectangles (top / bottom strips spanning the
   * full width, plus single-row left / right slivers around the
   * deselected cell), and ranges that did not contain the cell are
   * preserved untouched.
   *
   * The first surviving range becomes the new primary; the rest go
   * into the additional-range list. If the deselected cell was the
   * active cell, the active cell moves to the start of the new
   * primary so keyboard nav has a sensible anchor. We refuse to
   * deselect when doing so would leave the selection empty (Excel
   * does the same).
   */
  function deselectCellFromSelection(event: MouseEvent, cell: Cell) {
    const allRanges: CellSelectionRange[] = [];
    if (options.selectedRange.value) {
      allRanges.push(options.selectedRange.value);
    }
    allRanges.push(...options.additionalRanges.value);

    const remaining: CellSelectionRange[] = [];
    for (const range of allRanges) {
      if (cellInsideRange(range, cell.id)) {
        remaining.push(...splitRangeExcludingCell(range, cell.id));
      } else {
        remaining.push(range);
      }
    }

    if (remaining.length === 0) {
      // Refuse to deselect the only selected cell - keeping at least
      // one cell selected mirrors Excel and stops "Format cells" /
      // copy commands from suddenly having nothing to operate on.
      return;
    }

    draggingRangeStart.value = null;
    const newPrimary = remaining[0];
    const newAdditional = remaining.slice(1);
    options.onSetAdditionalRanges(newAdditional);
    options.onSelectRange(newPrimary);

    if (options.selectedCellId.value === cell.id) {
      const startCoords = findCellCoordinates(newPrimary.startCellId);
      if (startCoords) {
        const targetRow = options.projection.value.rows[startCoords.rowIndex];
        const targetColumn = options.projection.value.columns[startCoords.columnIndex];
        if (targetRow && targetColumn) {
          const replacementCell = getCell(options.worksheet.value, targetRow.id, targetColumn.id);
          options.onSelect(replacementCell, getCellAddress(options.projection.value, targetRow.id, targetColumn.id));
        }
      }
    }
    (event.currentTarget as HTMLElement | null)?.focus();
  }

  function cellInsideRange(range: CellSelectionRange, cellId: CellId): boolean {
    const start = findCellCoordinates(range.startCellId);
    const end = findCellCoordinates(range.endCellId);
    const target = findCellCoordinates(cellId);
    if (!start || !end || !target) {
      return false;
    }
    return target.rowIndex >= Math.min(start.rowIndex, end.rowIndex)
      && target.rowIndex <= Math.max(start.rowIndex, end.rowIndex)
      && target.columnIndex >= Math.min(start.columnIndex, end.columnIndex)
      && target.columnIndex <= Math.max(start.columnIndex, end.columnIndex);
  }

  /**
   * Split a rectangular range into up to four sub-rectangles whose
   * union equals the original range minus the cell at `excludedCellId`.
   * Returns `[range]` unchanged if any coordinate fails to resolve, so
   * a stale id cannot accidentally erase the entire selection.
   */
  function splitRangeExcludingCell(range: CellSelectionRange, excludedCellId: CellId): CellSelectionRange[] {
    const start = findCellCoordinates(range.startCellId);
    const end = findCellCoordinates(range.endCellId);
    const target = findCellCoordinates(excludedCellId);
    if (!start || !end || !target) {
      return [range];
    }
    const minRow = Math.min(start.rowIndex, end.rowIndex);
    const maxRow = Math.max(start.rowIndex, end.rowIndex);
    const minCol = Math.min(start.columnIndex, end.columnIndex);
    const maxCol = Math.max(start.columnIndex, end.columnIndex);
    const r = target.rowIndex;
    const c = target.columnIndex;
    const out: CellSelectionRange[] = [];
    if (minRow < r) {
      out.push(rangeFromIndices(minRow, minCol, r - 1, maxCol));
    }
    if (maxRow > r) {
      out.push(rangeFromIndices(r + 1, minCol, maxRow, maxCol));
    }
    if (minCol < c) {
      out.push(rangeFromIndices(r, minCol, r, c - 1));
    }
    if (maxCol > c) {
      out.push(rangeFromIndices(r, c + 1, r, maxCol));
    }
    return out;
  }

  function rangeFromIndices(minRow: number, minCol: number, maxRow: number, maxCol: number): CellSelectionRange {
    const startRow = options.projection.value.rows[minRow];
    const startColumn = options.projection.value.columns[minCol];
    const endRow = options.projection.value.rows[maxRow];
    const endColumn = options.projection.value.columns[maxCol];
    return {
      startCellId: createCellId(startRow.id, startColumn.id),
      endCellId: createCellId(endRow.id, endColumn.id),
    };
  }

  function openCellContextMenu(event: MouseEvent, rowId: RowId, columnId: ColumnId) {
    const cell = getCell(options.worksheet.value, rowId, columnId);
    const address = getCellAddress(options.projection.value, rowId, columnId);
    const range = options.selectedRange.value && selectedRangeIds.value.has(cell.id)
      ? options.selectedRange.value
      : { startCellId: cell.id, endCellId: cell.id };
    options.onCellContext({ event, cell, address, range });
  }

  /**
   * Select every cell in a row by clicking the row header. The range
   * spans the first and last visible columns in the projection - we do
   * not include tombstoned columns even if their ids still exist in
   * `columnOrder`.
   *
   * Modifier branches mirror the cell-level selection gestures:
   * - Shift+click extends from the anchor row of the current primary
   *   range to the clicked row (full-width rows between the two).
   * - Ctrl/Meta+click pushes the current primary range into the
   *   additional-ranges list and starts a new primary range covering
   *   the clicked row, so users can build "row 2 + row 5" or "row 2 +
   *   cell C8"-style disjoint selections.
   * - Default click clears any disjoint extras and selects only the
   *   clicked row.
   */
  function selectWholeRow(event: MouseEvent, rowId: RowId) {
    if (options.editingCellId.value || options.projection.value.columns.length === 0) {
      return;
    }
    const firstColumn = options.projection.value.columns[0];
    const lastColumn = options.projection.value.columns[options.projection.value.columns.length - 1];
    const firstCell = getCell(options.worksheet.value, rowId, firstColumn.id);
    const lastCell = getCell(options.worksheet.value, rowId, lastColumn.id);

    if (event.shiftKey && options.selectedCellId.value) {
      const anchorRowId = anchorRowIdFromSelection() ?? firstCell.rowId;
      const anchorFirstCell = getCell(options.worksheet.value, anchorRowId, firstColumn.id);
      options.onSelect(anchorFirstCell, getCellAddress(options.projection.value, anchorFirstCell.rowId, anchorFirstCell.columnId));
      options.onSelectRange({ startCellId: anchorFirstCell.id, endCellId: lastCell.id });
      return;
    }
    if ((event.ctrlKey || event.metaKey) && options.selectedRange.value && options.selectedCellId.value) {
      options.onAddRange(options.selectedRange.value);
      options.onSelect(firstCell, getCellAddress(options.projection.value, firstCell.rowId, firstCell.columnId));
      options.onSelectRange({ startCellId: firstCell.id, endCellId: lastCell.id });
      return;
    }
    options.onClearAdditionalRanges();
    options.onSelect(firstCell, getCellAddress(options.projection.value, firstCell.rowId, firstCell.columnId));
    options.onSelectRange({ startCellId: firstCell.id, endCellId: lastCell.id });
  }

  function selectWholeColumn(event: MouseEvent, columnId: ColumnId) {
    if (options.editingCellId.value || options.projection.value.rows.length === 0) {
      return;
    }
    const firstRow = options.projection.value.rows[0];
    const lastRow = options.projection.value.rows[options.projection.value.rows.length - 1];
    const firstCell = getCell(options.worksheet.value, firstRow.id, columnId);
    const lastCell = getCell(options.worksheet.value, lastRow.id, columnId);

    if (event.shiftKey && options.selectedCellId.value) {
      const anchorColumnId = anchorColumnIdFromSelection() ?? firstCell.columnId;
      const anchorFirstCell = getCell(options.worksheet.value, firstRow.id, anchorColumnId);
      options.onSelect(anchorFirstCell, getCellAddress(options.projection.value, anchorFirstCell.rowId, anchorFirstCell.columnId));
      options.onSelectRange({ startCellId: anchorFirstCell.id, endCellId: lastCell.id });
      return;
    }
    if ((event.ctrlKey || event.metaKey) && options.selectedRange.value && options.selectedCellId.value) {
      options.onAddRange(options.selectedRange.value);
      options.onSelect(firstCell, getCellAddress(options.projection.value, firstCell.rowId, firstCell.columnId));
      options.onSelectRange({ startCellId: firstCell.id, endCellId: lastCell.id });
      return;
    }
    options.onClearAdditionalRanges();
    options.onSelect(firstCell, getCellAddress(options.projection.value, firstCell.rowId, firstCell.columnId));
    options.onSelectRange({ startCellId: firstCell.id, endCellId: lastCell.id });
  }

  /**
   * Resolve the row id of the current selection's anchor (its primary
   * range's start cell, or the active cell as a fallback). Used by
   * shift+click on row headers to know which row to extend from.
   */
  function anchorRowIdFromSelection(): RowId | null {
    const anchorCellId = options.selectedRange.value?.startCellId ?? options.selectedCellId.value;
    if (!anchorCellId) {
      return null;
    }
    const coords = findCellCoordinates(anchorCellId);
    return coords ? options.projection.value.rows[coords.rowIndex]?.id ?? null : null;
  }

  function anchorColumnIdFromSelection(): ColumnId | null {
    const anchorCellId = options.selectedRange.value?.startCellId ?? options.selectedCellId.value;
    if (!anchorCellId) {
      return null;
    }
    const coords = findCellCoordinates(anchorCellId);
    return coords ? options.projection.value.columns[coords.columnIndex]?.id ?? null : null;
  }

  function extendRangeSelection(rowId: RowId, columnId: ColumnId) {
    if (!draggingRangeStart.value) {
      return;
    }
    const cell = getCell(options.worksheet.value, rowId, columnId);
    options.onSelectRange({ startCellId: draggingRangeStart.value, endCellId: cell.id });
  }

  function endRangeDrag() {
    draggingRangeStart.value = null;
  }

  function isWholeRowSelected(rowId: RowId) {
    return options.projection.value.columns.length > 0
      && options.projection.value.columns.every((column) => selectedRangeIds.value.has(getCell(options.worksheet.value, rowId, column.id).id));
  }

  function isWholeColumnSelected(columnId: ColumnId) {
    return options.projection.value.rows.length > 0
      && options.projection.value.rows.every((row) => selectedRangeIds.value.has(getCell(options.worksheet.value, row.id, columnId).id));
  }

  /**
   * Window-level keydown router. Lets the user move and edit the active
   * cell even when only the grid (rather than a specific `<td>`) has
   * focus. Bails when the event target is another editor so typing in
   * the formula bar does not also move the active cell.
   */
  function handleWindowKeydown(event: KeyboardEvent) {
    if (event.defaultPrevented || isTypingInAnotherEditor(event.target) || !options.selectedCellId.value) {
      return;
    }
    const selectedCell = findSelectedCell();
    if (!selectedCell) {
      return;
    }
    handleEditKey(event, selectedCell.rowId, selectedCell.columnId);
  }

  /**
   * Single keydown router for the active cell.
   *
   * 1. While the inline editor is open, this handler does nothing.
   * 2. Arrow keys move (or extend with Shift) the selection regardless
   *    of `readonly`.
   * 3. After that we gate on `readonly`: Enter/F2 enters edit mode,
   *    Backspace/Delete enters edit mode with an empty draft, and any
   *    printable key enters edit mode using that key as the first
   *    character.
   */
  function handleEditKey(event: KeyboardEvent, rowId: RowId, columnId: ColumnId) {
    if (options.editingCellId.value) {
      return;
    }
    if (handleSelectionKey(event, rowId, columnId)) {
      return;
    }
    if (options.readonly.value) {
      return;
    }
    if (event.key === "Enter" || event.key === "F2") {
      event.preventDefault();
      void options.startEditing(rowId, columnId);
      return;
    }
    if (event.key === "Backspace" || event.key === "Delete") {
      event.preventDefault();
      void options.startEditing(rowId, columnId, "");
      return;
    }
    if (isPrintableKey(event)) {
      event.preventDefault();
      void options.startEditing(rowId, columnId, event.key);
    }
  }

  function handleSelectionKey(event: KeyboardEvent, rowId: RowId, columnId: ColumnId) {
    const navigation = resolveNavigationKey(event);
    if (!navigation) {
      return false;
    }
    const currentRowIndex = options.projection.value.rowIndexById.get(rowId);
    const currentColumnIndex = options.projection.value.columnIndexById.get(columnId);
    if (currentRowIndex == null || currentColumnIndex == null) {
      return false;
    }

    event.preventDefault();
    const targetRow = options.projection.value.rows[clampIndex(currentRowIndex + navigation.delta.rows, options.projection.value.rows.length)];
    const targetColumn = options.projection.value.columns[clampIndex(currentColumnIndex + navigation.delta.cols, options.projection.value.columns.length)];
    if (!targetRow || !targetColumn) {
      return true;
    }

    const targetCell = getCell(options.worksheet.value, targetRow.id, targetColumn.id);
    const range = navigation.extend
      ? { startCellId: options.selectedRange.value?.startCellId ?? createCellId(rowId, columnId), endCellId: targetCell.id }
      : { startCellId: targetCell.id, endCellId: targetCell.id };
    // Arrow-/Tab-key navigation (with or without Shift) collapses the
    // selection to a single contiguous rectangle, matching Excel /
    // Sheets behaviour where keyboard motion ends a Ctrl+click
    // multi-selection.
    options.onClearAdditionalRanges();
    options.onSelect(targetCell, getCellAddress(options.projection.value, targetRow.id, targetColumn.id));
    options.onSelectRange(range);
    void focusCell(targetCell.id);
    return true;
  }

  /**
   * Map a keydown event to a navigation delta plus whether the gesture
   * should extend the existing range (Shift+Arrow) or collapse to the
   * new cell (everything else). Tab and Shift+Tab always collapse
   * because that is how spreadsheets traditionally handle them.
   */
  function resolveNavigationKey(event: KeyboardEvent): { delta: { rows: number; cols: number }; extend: boolean } | null {
    const arrowDelta = KEYBOARD_SELECTION_DELTAS[event.key];
    if (arrowDelta) {
      return { delta: arrowDelta, extend: event.shiftKey };
    }
    if (event.key === "Tab") {
      return { delta: { rows: 0, cols: event.shiftKey ? -1 : 1 }, extend: false };
    }
    return null;
  }

  /**
   * Move DOM focus to a cell by id. We do this on the next tick because
   * the parent might still be reacting to the `select-range` emission
   * and have not yet re-rendered the table.
   */
  async function focusCell(cellId: CellId) {
    await nextTick();
    const cells = options.viewportEl.value?.querySelectorAll<HTMLElement>("[data-cell-id]") ?? [];
    for (const cell of cells) {
      if (cell.dataset.cellId === cellId) {
        cell.focus();
        return;
      }
    }
  }

  function findSelectedCell() {
    if (!options.selectedCellId.value) {
      return null;
    }
    for (const row of options.projection.value.rows) {
      for (const column of options.projection.value.columns) {
        const cell = getCell(options.worksheet.value, row.id, column.id);
        if (cell.id === options.selectedCellId.value) {
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
        const row = options.projection.value.rows[rowIndex];
        const column = options.projection.value.columns[columnIndex];
        if (row && column) {
          ids.push(getCell(options.worksheet.value, row.id, column.id).id);
        }
      }
    }
    return ids;
  }

  function findCellCoordinates(cellId: CellId) {
    for (const row of options.projection.value.rows) {
      for (const column of options.projection.value.columns) {
        if (getCell(options.worksheet.value, row.id, column.id).id === cellId) {
          return { rowIndex: row.index, columnIndex: column.index };
        }
      }
    }
    return null;
  }

  onMounted(() => {
    window.addEventListener("keydown", handleWindowKeydown);
    window.addEventListener("mouseup", endRangeDrag);
  });

  onBeforeUnmount(() => {
    window.removeEventListener("keydown", handleWindowKeydown);
    window.removeEventListener("mouseup", endRangeDrag);
  });

  return {
    selectedRangeIds,
    startRangeSelection,
    openCellContextMenu,
    selectWholeRow,
    selectWholeColumn,
    extendRangeSelection,
    handleEditKey,
    isWholeRowSelected,
    isWholeColumnSelected,
  };
}

function clampIndex(index: number, length: number) {
  return Math.max(0, Math.min(index, length - 1));
}

function isPrintableKey(event: KeyboardEvent) {
  return event.key.length === 1 && !event.metaKey && !event.ctrlKey && !event.altKey;
}

function isTypingInAnotherEditor(target: EventTarget | null) {
  if (!(target instanceof HTMLElement)) {
    return false;
  }
  return target.matches("input, textarea, select, [contenteditable='true']");
}
