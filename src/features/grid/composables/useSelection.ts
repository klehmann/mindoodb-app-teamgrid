/**
 * Selection state for the spreadsheet grid.
 *
 * Owns the single "active cell" (`selectedCellId`) and the optional
 * rectangular `selectedRange` used by multi-cell actions like Copy or
 * Format. Selection is addressed by stable cell ids; visible `A1`-style
 * addresses are derived from the current {@link GridProjection} on the fly.
 *
 * The composable also guarantees the selection stays valid as the
 * underlying worksheet changes: when the active worksheet or its projection
 * is replaced, the selection is either kept (if it still maps to a visible
 * cell) or snapped to the top-left cell of the new projection.
 */
import { computed, ref, watch, type Ref } from "vue";
import { createCellId, type Cell, type Worksheet } from "@/features/document/lib/teamgridDocument";
import { getCell, type GridProjection } from "@/features/grid/lib/gridProjection";

/** Inclusive rectangular cell range used for multi-cell selection state. */
export interface CellSelectionRange {
  startCellId: string;
  endCellId: string;
}

export interface UseSelectionOptions {
  activeWorksheet: Readonly<Ref<Worksheet | null>>;
  projection: Readonly<Ref<GridProjection | null>>;
}

export function useSelection(options: UseSelectionOptions) {
  const { activeWorksheet, projection } = options;
  const selectedCellId = ref<string | null>(null);
  const selectedRange = ref<CellSelectionRange | null>(null);
  const selectedCellAddress = ref("");

  /**
   * Persisted cell record for {@link selectedCellId}. Returns `null` when
   * nothing is selected or when the worksheet is not yet loaded. The lookup
   * walks the projection so we get a fresh `Cell` (synthetic empty if the
   * cell has never been persisted) rather than a stale reference.
   */
  const selectedCell = computed<Cell | null>(() => {
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

  /** All cells covered by the current selection, or the single active cell. */
  const selectedCells = computed(() => cellsForRange(selectedRange.value).map(({ cell }) => cell));

  /** Convenience flag used to gate selection-dependent menu commands. */
  const hasSelection = computed(() => selectedCells.value.length > 0);

  /** Resolve a stable cell id to its `{rowIndex, columnIndex}` in the current projection. */
  function findCellCoordinates(cellId: string) {
    if (!activeWorksheet.value || !projection.value) {
      return null;
    }
    return findCellCoordinatesInProjection(cellId, activeWorksheet.value, projection.value);
  }

  function cellsForRange(range: CellSelectionRange | null) {
    if (!activeWorksheet.value || !projection.value || !range) {
      const coordinates = selectedCell.value ? findCellCoordinates(selectedCell.value.id) : null;
      return selectedCell.value
        ? [{ cell: selectedCell.value, rowIndex: coordinates?.rowIndex ?? 0, columnIndex: coordinates?.columnIndex ?? 0 }]
        : [];
    }
    const start = findCellCoordinates(range.startCellId);
    const end = findCellCoordinates(range.endCellId);
    if (!start || !end) {
      const coordinates = selectedCell.value ? findCellCoordinates(selectedCell.value.id) : null;
      return selectedCell.value
        ? [{ cell: selectedCell.value, rowIndex: coordinates?.rowIndex ?? 0, columnIndex: coordinates?.columnIndex ?? 0 }]
        : [];
    }
    const cells: Array<{ cell: Cell; rowIndex: number; columnIndex: number }> = [];
    for (let rowIndex = Math.min(start.rowIndex, end.rowIndex); rowIndex <= Math.max(start.rowIndex, end.rowIndex); rowIndex += 1) {
      for (let columnIndex = Math.min(start.columnIndex, end.columnIndex); columnIndex <= Math.max(start.columnIndex, end.columnIndex); columnIndex += 1) {
        const row = projection.value.rows[rowIndex];
        const column = projection.value.columns[columnIndex];
        if (row && column) {
          cells.push({ cell: getCell(activeWorksheet.value, row.id, column.id), rowIndex, columnIndex });
        }
      }
    }
    return cells;
  }

  function boundsForRange(range: CellSelectionRange | null) {
    const cells = cellsForRange(range);
    if (cells.length === 0) {
      return null;
    }
    return {
      minRow: Math.min(...cells.map((cell) => cell.rowIndex)),
      maxRow: Math.max(...cells.map((cell) => cell.rowIndex)),
      minCol: Math.min(...cells.map((cell) => cell.columnIndex)),
      maxCol: Math.max(...cells.map((cell) => cell.columnIndex)),
    };
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

  return {
    selectedCellId,
    selectedRange,
    selectedCellAddress,
    selectedCell,
    selectedCells,
    hasSelection,
    findCellCoordinates,
    cellsForRange,
    boundsForRange,
    cellIdAt,
  };
}

/**
 * Variant of `findCellCoordinates` that operates on an explicit worksheet
 * and projection pair. Used inside `updateGrid` mutators where the
 * worksheet argument is a draft copy that the reactive `activeWorksheet`
 * has not yet observed.
 */
export function findCellCoordinatesInProjection(
  cellId: string,
  worksheet: Worksheet,
  worksheetProjection: GridProjection,
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
