import {
  createCellId,
  createEmptyCell,
  type Cell,
  type ColumnId,
  type RowId,
  type Worksheet,
} from "@/lib/teamgridDocument";

export interface VisibleColumn {
  id: ColumnId;
  index: number;
  label: string;
  width: number;
}

export interface VisibleRow {
  id: RowId;
  index: number;
  label: string;
  height?: number;
}

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
      width: worksheet.columnsById[id]?.width ?? 120,
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

export function getCell(worksheet: Worksheet, rowId: RowId, columnId: ColumnId): Cell {
  return worksheet.cellsById[createCellId(rowId, columnId)] ?? createEmptyCell(rowId, columnId);
}

export function getCellAddress(projection: GridProjection, rowId: RowId, columnId: ColumnId) {
  return projection.cellAddressById.get(createCellId(rowId, columnId)) ?? "#REF!";
}

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
