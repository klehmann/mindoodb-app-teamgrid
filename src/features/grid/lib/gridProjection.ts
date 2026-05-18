/**
 * Project a stable-ID worksheet into the visible row/column grid the UI
 * renders.
 *
 * The persisted document never uses positional row/column references. Rows
 * and columns are addressed by opaque stable IDs so that concurrent inserts
 * never silently retarget formulas. The UI, however, needs ordered
 * Excel-style addresses (`A1`, `B12`) for selection, formula display, and
 * keyboard navigation. This module is the bridge: every render pass produces
 * a {@link GridProjection} that maps `rowId`/`columnId` to indices and
 * `A1`-style strings, while {@link parseCellAddress} walks back the other
 * direction when the formula parser encounters a user-typed reference.
 *
 * Helpers:
 * - {@link projectWorksheet} builds the visible grid from a worksheet.
 * - {@link getCell} returns the persisted cell or a synthetic empty one.
 * - {@link parseCellAddress} resolves an `A1`-style string against a projection.
 * - {@link columnIndexToLabel} / {@link columnLabelToIndex} translate the
 *   alphabetic column header used in the spreadsheet UI.
 */
import {
  createCellId,
  createEmptyCell,
  type Cell,
  type ColumnId,
  type RowId,
  type Worksheet,
} from "@/features/document/lib/teamgridDocument";
import { DEFAULT_COLUMN_WIDTH } from "@/shared/lib/gridDimensions";

/** One column slot in the projected grid (after dedupe and tombstone filtering). */
export interface VisibleColumn {
  id: ColumnId;
  index: number;
  label: string;
  width: number;
}

/** One row slot in the projected grid (after dedupe and tombstone filtering). */
export interface VisibleRow {
  id: RowId;
  index: number;
  label: string;
  height?: number;
}

/**
 * Render-time view of a worksheet: ordered rows/columns plus lookup tables
 * for converting stable IDs to indices and to user-facing `A1` addresses.
 */
export interface GridProjection {
  rows: VisibleRow[];
  columns: VisibleColumn[];
  rowIndexById: Map<RowId, number>;
  columnIndexById: Map<ColumnId, number>;
  cellAddressById: Map<string, string>;
}

/**
 * Build a rendered grid view from stable CRDT-friendly IDs.
 *
 * The order arrays may contain duplicate IDs after concurrent move operations.
 * For a deterministic sample-app projection, the first visible occurrence wins
 * and later duplicates are ignored.
 */
export function projectWorksheet(worksheet: Worksheet): GridProjection {
  const rows = dedupeIds(worksheet.rowOrder)
    .filter((id) => !worksheet.rowsById[id]?.deletedAt)
    .map<VisibleRow>((id, index) => ({
      id,
      index,
      label: String(index + 1),
      height: worksheet.rowsById[id]?.height,
    }));

  const columns = dedupeIds(worksheet.columnOrder)
    .filter((id) => !worksheet.columnsById[id]?.deletedAt)
    .map<VisibleColumn>((id, index) => ({
      id,
      index,
      label: columnIndexToLabel(index),
      width: worksheet.columnsById[id]?.width ?? DEFAULT_COLUMN_WIDTH,
    }));

  const rowIndexById = new Map(rows.map((row) => [row.id, row.index] as const));
  const columnIndexById = new Map(columns.map((column) => [column.id, column.index] as const));
  const cellAddressById = new Map<string, string>();
  for (const row of rows) {
    for (const column of columns) {
      cellAddressById.set(createCellId(row.id, column.id), `${column.label}${row.label}`);
    }
  }

  return {
    rows,
    columns,
    rowIndexById,
    columnIndexById,
    cellAddressById,
  };
}

/**
 * Return the persisted cell for a given row/column, or a synthetic empty one.
 *
 * The sample storage model only persists cells that have content, so this
 * keeps the rest of the code from having to special-case `undefined`.
 */
export function getCell(worksheet: Worksheet, rowId: RowId, columnId: ColumnId): Cell {
  return worksheet.cellsById[createCellId(rowId, columnId)] ?? createEmptyCell(rowId, columnId);
}

/** Look up the `A1`-style address for a stable cell id, or `#REF!` if missing. */
export function getCellAddress(projection: GridProjection, rowId: RowId, columnId: ColumnId) {
  return projection.cellAddressById.get(createCellId(rowId, columnId)) ?? "#REF!";
}

/**
 * Resolve an `A1`-style address typed by the user into stable row/column IDs.
 *
 * Returns `null` for malformed input or when the referenced row/column is
 * outside the current projection, so callers can surface `#REF!`.
 */
export function parseCellAddress(address: string, projection: GridProjection) {
  const match = /^([A-Z]+)([1-9][0-9]*)$/i.exec(address.trim());
  if (!match) {
    return null;
  }
  const columnIndex = columnLabelToIndex(match[1].toUpperCase());
  const rowIndex = Number.parseInt(match[2], 10) - 1;
  const row = projection.rows[rowIndex];
  const column = projection.columns[columnIndex];
  if (!row || !column) {
    return null;
  }
  return { rowId: row.id, columnId: column.id };
}

/**
 * Convert a zero-based column index to an Excel-style label.
 *
 * Examples: 0 -> "A", 25 -> "Z", 26 -> "AA", 701 -> "ZZ".
 */
export function columnIndexToLabel(index: number) {
  let value = index + 1;
  let label = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

/** Inverse of {@link columnIndexToLabel}: parse "AA" back to 26, etc. */
export function columnLabelToIndex(label: string) {
  let value = 0;
  for (const character of label) {
    value = value * 26 + character.charCodeAt(0) - 64;
  }
  return value - 1;
}

function dedupeIds<T extends string>(ids: T[]) {
  const seen = new Set<T>();
  const result: T[] = [];
  for (const id of ids) {
    if (seen.has(id)) {
      continue;
    }
    seen.add(id);
    result.push(id);
  }
  return result;
}
