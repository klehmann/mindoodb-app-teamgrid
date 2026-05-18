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
  readonly: Readonly<Ref<boolean>>;
  editingCellId: Readonly<Ref<string | null>>;
  editDraft: Readonly<Ref<string>>;
  viewportEl: Readonly<Ref<HTMLElement | null>>;
  appendPickedCellToInlineFormula: (cell: Cell) => void;
  commitEditingExternally: () => void;
  startEditing: (rowId: RowId, columnId: ColumnId, initialValue?: string) => Promise<void> | void;
  onSelect: (cell: Cell, address: string) => void;
  onSelectRange: (range: CellSelectionRange) => void;
  onCellContext: (payload: { event: MouseEvent; cell: Cell; address: string; range: CellSelectionRange }) => void;
}

export function useGridSelectionGestures(options: UseGridSelectionGesturesOptions) {
  const draggingRangeStart = ref<CellId | null>(null);

  const selectedRangeIds = computed(() => new Set(options.selectedRange.value
    ? getRangeCellIds(options.selectedRange.value)
    : []));

  function selectCellByCoordinates(rowId: RowId, columnId: ColumnId) {
    const cell = getCell(options.worksheet.value, rowId, columnId);
    options.onSelect(cell, getCellAddress(options.projection.value, rowId, columnId));
  }

  function handleCellClick(event: MouseEvent, rowId: RowId, columnId: ColumnId) {
    selectCellByCoordinates(rowId, columnId);
    (event.currentTarget as HTMLElement | null)?.focus();
  }

  /**
   * Mousedown on a cell. Three branches matter:
   *
   * 1. Formula picking - if we're mid-edit and the draft starts with
   *    `=`, treat the click as "pick this cell as a reference" instead
   *    of a selection change.
   * 2. Mid-edit, non-formula click - commit the in-flight edit first,
   *    then start a fresh range selection on the clicked cell.
   * 3. Default - start a new range anchored on the clicked cell.
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
    draggingRangeStart.value = cell.id;
    options.onSelectRange({ startCellId: cell.id, endCellId: cell.id });
    handleCellClick(event, rowId, columnId);
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
   */
  function selectWholeRow(rowId: RowId) {
    if (options.editingCellId.value || options.projection.value.columns.length === 0) {
      return;
    }
    const firstColumn = options.projection.value.columns[0];
    const lastColumn = options.projection.value.columns[options.projection.value.columns.length - 1];
    const firstCell = getCell(options.worksheet.value, rowId, firstColumn.id);
    const lastCell = getCell(options.worksheet.value, rowId, lastColumn.id);
    options.onSelect(firstCell, getCellAddress(options.projection.value, firstCell.rowId, firstCell.columnId));
    options.onSelectRange({ startCellId: firstCell.id, endCellId: lastCell.id });
  }

  function selectWholeColumn(columnId: ColumnId) {
    if (options.editingCellId.value || options.projection.value.rows.length === 0) {
      return;
    }
    const firstRow = options.projection.value.rows[0];
    const lastRow = options.projection.value.rows[options.projection.value.rows.length - 1];
    const firstCell = getCell(options.worksheet.value, firstRow.id, columnId);
    const lastCell = getCell(options.worksheet.value, lastRow.id, columnId);
    options.onSelect(firstCell, getCellAddress(options.projection.value, firstCell.rowId, firstCell.columnId));
    options.onSelectRange({ startCellId: firstCell.id, endCellId: lastCell.id });
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
    const delta = KEYBOARD_SELECTION_DELTAS[event.key];
    if (!delta) {
      return false;
    }
    const currentRowIndex = options.projection.value.rowIndexById.get(rowId);
    const currentColumnIndex = options.projection.value.columnIndexById.get(columnId);
    if (currentRowIndex == null || currentColumnIndex == null) {
      return false;
    }

    event.preventDefault();
    const targetRow = options.projection.value.rows[clampIndex(currentRowIndex + delta.rows, options.projection.value.rows.length)];
    const targetColumn = options.projection.value.columns[clampIndex(currentColumnIndex + delta.cols, options.projection.value.columns.length)];
    if (!targetRow || !targetColumn) {
      return true;
    }

    const targetCell = getCell(options.worksheet.value, targetRow.id, targetColumn.id);
    const range = event.shiftKey
      ? { startCellId: options.selectedRange.value?.startCellId ?? createCellId(rowId, columnId), endCellId: targetCell.id }
      : { startCellId: targetCell.id, endCellId: targetCell.id };
    options.onSelect(targetCell, getCellAddress(options.projection.value, targetRow.id, targetColumn.id));
    options.onSelectRange(range);
    void focusCell(targetCell.id);
    return true;
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
