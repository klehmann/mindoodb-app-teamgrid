/**
 * Semantic operation types for Teamgrid edits and a serializer that converts
 * them into Automerge-friendly JSON patches via the App SDK.
 *
 * The UI never mutates SDK payloads directly. It records intent as a
 * {@link TeamGridOperation} batch and lets {@link serializeTeamGridOperations}
 * turn that batch into a path-scoped {@link MindooDBAppJsonPatch}. This keeps
 * concurrent edits cleanly mergeable: two users editing different cells
 * produce patches that touch disjoint stable-ID paths, so Automerge can
 * combine them without either edit clobbering the other.
 */
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

/**
 * One semantic edit recorded by the UI. The shapes are intentionally close to
 * the persisted schema so the serializer can map them to JSON patch
 * operations without re-deriving anything.
 *
 * Variants:
 * - `setCell` writes one cell (value, formula, and style).
 * - `setCellsStyle` rewrites style on a contiguous selection while keeping
 *   each cell's value/formula intact.
 * - `insertRow` / `insertColumn` add a new ordered slot.
 * - `tombstoneRow` / `tombstoneColumn` / `tombstoneWorksheet` mark an entity
 *   deleted without physically removing it, so formulas can still report
 *   `#REF!` instead of losing identity.
 * - `addWorksheet`, `renameWorksheet` handle workbook-level changes.
 * - `setDocumentProperties` writes the top-level `subject` and `tags` fields.
 */
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
  | { type: "setDocumentProperties"; subject: string; tags: string[] };

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
      case "setDocumentProperties":
        pushSet(json, ["subject"], operation.subject);
        pushSet(json, ["tags"], operation.tags);
        break;
    }
  }
  return { json };
}

/** Type-narrowed predicate for a non-empty operation batch. */
export function hasTeamGridOperations(operations: TeamGridOperation[]) {
  return operations.length > 0;
}

/** Path inside the persisted document that points at one worksheet's record. */
function worksheetPath(worksheetId: WorksheetId) {
  return ["teamgrid", "workbook", "worksheetsById", worksheetId];
}

/** Path inside the persisted document that points at one cell's record. */
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
