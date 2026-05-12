import type { MindooDBAppJsonPatch, MindooDBAppUpdateDocumentInput } from "mindoodb-app-sdk";

import type {
  Cell,
  CellId,
  CellStyle,
  ColumnId,
  ColumnMeta,
  RowId,
  RowMeta,
  Worksheet,
  WorksheetId,
} from "@/lib/teamgridDocument";

export type TeamGridOperation =
  | { type: "setCell"; worksheetId: WorksheetId; cell: Cell }
  | { type: "setCellsStyle"; worksheetId: WorksheetId; cells: Cell[]; style: CellStyle }
  | { type: "insertRow"; worksheetId: WorksheetId; rowId: RowId; row: RowMeta; index: number }
  | { type: "tombstoneRow"; worksheetId: WorksheetId; rowId: RowId; deletedAt: string }
  | { type: "insertColumn"; worksheetId: WorksheetId; columnId: ColumnId; column: ColumnMeta; index: number }
  | { type: "tombstoneColumn"; worksheetId: WorksheetId; columnId: ColumnId; deletedAt: string }
  | { type: "addWorksheet"; worksheet: Worksheet; index: number }
  | { type: "renameWorksheet"; worksheetId: WorksheetId; title: string }
  | { type: "tombstoneWorksheet"; worksheetId: WorksheetId; deletedAt: string }
  | { type: "setWorkbookTitle"; subject: string; title: string };

/**
 * Convert Teamgrid semantic operations into App SDK JSON patches.
 *
 * The serializer keeps paths scoped to stable IDs, so independent edits to
 * different cells/rows/columns become independent Automerge map/list changes.
 */
export function serializeTeamGridOperations(
  operations: TeamGridOperation[],
  options: { baseHeads?: string[] } = {},
): MindooDBAppUpdateDocumentInput {
  const json: MindooDBAppJsonPatch = options.baseHeads ? { baseHeads: [...options.baseHeads] } : {};
  for (const operation of operations) {
    switch (operation.type) {
      case "setCell":
        pushSet(json, cellPath(operation.worksheetId, operation.cell.id), operation.cell);
        break;
      case "setCellsStyle":
        for (const cell of operation.cells) {
          pushSet(json, cellPath(operation.worksheetId, cell.id), {
            ...cell,
            style: {
              ...cell.style,
              ...operation.style,
            },
          });
        }
        break;
      case "insertRow":
        pushSet(json, [...worksheetPath(operation.worksheetId), "rowsById", operation.rowId], operation.row);
        pushListInsert(json, [...worksheetPath(operation.worksheetId), "rowOrder"], operation.index, [operation.rowId]);
        break;
      case "tombstoneRow":
        pushSet(json, [...worksheetPath(operation.worksheetId), "rowsById", operation.rowId, "deletedAt"], operation.deletedAt);
        break;
      case "insertColumn":
        pushSet(json, [...worksheetPath(operation.worksheetId), "columnsById", operation.columnId], operation.column);
        pushListInsert(json, [...worksheetPath(operation.worksheetId), "columnOrder"], operation.index, [operation.columnId]);
        break;
      case "tombstoneColumn":
        pushSet(json, [...worksheetPath(operation.worksheetId), "columnsById", operation.columnId, "deletedAt"], operation.deletedAt);
        break;
      case "addWorksheet":
        pushSet(json, ["teamgrid", "workbook", "worksheetsById", operation.worksheet.id], operation.worksheet);
        pushListInsert(json, ["teamgrid", "workbook", "worksheetOrder"], operation.index, [operation.worksheet.id]);
        break;
      case "renameWorksheet":
        pushSet(json, [...worksheetPath(operation.worksheetId), "title"], operation.title);
        break;
      case "tombstoneWorksheet":
        pushSet(json, [...worksheetPath(operation.worksheetId), "deletedAt"], operation.deletedAt);
        break;
      case "setWorkbookTitle":
        pushSet(json, ["subject"], operation.subject);
        pushSet(json, ["teamgrid", "workbook", "title"], operation.title);
        break;
    }
  }
  return { json };
}

export function hasTeamGridOperations(operations: TeamGridOperation[]) {
  return operations.length > 0;
}

function worksheetPath(worksheetId: WorksheetId) {
  return ["teamgrid", "workbook", "worksheetsById", worksheetId];
}

function cellPath(worksheetId: WorksheetId, cellId: CellId) {
  return [...worksheetPath(worksheetId), "cellsById", cellId];
}

function pushSet(json: MindooDBAppJsonPatch, path: Array<string | number>, value: unknown) {
  json.set ??= [];
  json.set.push({ path, value });
}

function pushListInsert(json: MindooDBAppJsonPatch, path: Array<string | number>, index: number, values: unknown[]) {
  json.listInsert ??= [];
  json.listInsert.push({ path, index, values });
}
