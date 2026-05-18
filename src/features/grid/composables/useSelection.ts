/**
 * Selection state for the spreadsheet grid.
 *
 * Owns the single "active cell" (`selectedCellId`) and the optional
 * primary rectangular `selectedRange` used by multi-cell actions like
 * Copy or Format, plus an array of disjoint `additionalRanges` produced
 * by Ctrl/Meta+click multi-selection. Selection is addressed by stable
 * cell ids; visible `A1`-style addresses are derived from the current
 * {@link GridProjection} on the fly.
 *
 * `selectedCells` returns the deduplicated union of cells across the
 * primary and additional ranges, so style/format operations that iterate
 * the selection naturally cover the whole multi-range selection.
 *
 * The composable also guarantees the selection stays valid as the
 * underlying worksheet changes: when the active worksheet or its projection
 * is replaced, the selection is either kept (if it still maps to a visible
 * cell) or snapped to the top-left cell of the new projection, and the
 * additional ranges are dropped because their cells may no longer exist.
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
  /**
   * Extra disjoint ranges produced by Ctrl/Meta+click multi-selection.
   * The primary `selectedRange` always owns the active cell; everything
   * else lives here so callers that only care about the active rectangle
   * (e.g. clipboard) keep working unchanged.
   */
  const additionalRanges = ref<CellSelectionRange[]>([]);
  const selectedCellAddress = ref("");

  /**
   * Primary range followed by additional ranges. Used by anything that
   * needs to iterate every selected rectangle (highlighting, format
   * dialog application, etc.).
   */
  const allSelectedRanges = computed<CellSelectionRange[]>(() => {
    const ranges: CellSelectionRange[] = [];
    if (selectedRange.value) {
      ranges.push(selectedRange.value);
    }
    ranges.push(...additionalRanges.value);
    return ranges;
  });

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

  /**
   * All cells covered by the current selection.
   *
   * Returns the deduplicated union of cells across the primary range and
   * any additional Ctrl/Meta+click ranges. Falls back to the single
   * active cell when no range is set.
   */
  const selectedCells = computed(() => {
    if (allSelectedRanges.value.length === 0) {
      return cellsForRange(null).map(({ cell }) => cell);
    }
    const seenIds = new Set<string>();
    const cells: Cell[] = [];
    for (const range of allSelectedRanges.value) {
      for (const { cell } of cellsForRange(range)) {
        if (seenIds.has(cell.id)) {
          continue;
        }
        seenIds.add(cell.id);
        cells.push(cell);
      }
    }
    return cells;
  });

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
        additionalRanges.value = [];
        selectedCellAddress.value = "";
        return;
      }
      if (selectedCellId.value && findCellCoordinates(selectedCellId.value)) {
        // Selection still maps to a visible cell; just prune any extra
        // Ctrl/Meta+click sub-ranges whose endpoints no longer resolve
        // (e.g. their row or column was deleted). Surviving sub-ranges
        // are kept so a row-insert from a collaborator does not wipe
        // the user's multi-selection.
        if (additionalRanges.value.length > 0) {
          const pruned = additionalRanges.value.filter((range) =>
            findCellCoordinates(range.startCellId) && findCellCoordinates(range.endCellId));
          if (pruned.length !== additionalRanges.value.length) {
            additionalRanges.value = pruned;
          }
        }
        return;
      }
      const firstRow = projection.value.rows[0];
      const firstColumn = projection.value.columns[0];
      if (!firstRow || !firstColumn) {
        selectedCellId.value = null;
        selectedRange.value = null;
        additionalRanges.value = [];
        selectedCellAddress.value = "";
        return;
      }
      const cell = getCell(activeWorksheet.value, firstRow.id, firstColumn.id);
      selectedCellId.value = cell.id;
      selectedRange.value = { startCellId: cell.id, endCellId: cell.id };
      additionalRanges.value = [];
      selectedCellAddress.value = projection.value.cellAddressById.get(cell.id) ?? "";
    },
  );

  return {
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
