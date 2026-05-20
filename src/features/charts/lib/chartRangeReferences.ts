import { createCellId, type SeriesRange, type WorksheetId } from "@/features/document/lib/teamgridDocument";
import type { FormulaContext } from "@/features/formulas/lib";
import { columnLabelToIndex } from "@/features/grid/lib/gridProjection";

export function formatChartRange(range: SeriesRange, context: FormulaContext) {
  const projection = context.projectionsByWorksheetId.get(range.worksheetId);
  const sheetName = context.sheetNameById.get(range.worksheetId) ?? "#REF";
  const start = projection?.cellAddressById.get(createCellId(range.startRowId, range.startColumnId));
  const end = projection?.cellAddressById.get(createCellId(range.endRowId, range.endColumnId));
  if (!projection || !start || !end) {
    return "#REF!";
  }
  const reference = start === end ? start : `${start}:${end}`;
  return `${quoteSheetName(sheetName)}!${reference}`;
}

export function parseChartRangeReference(source: string, currentWorksheetId: WorksheetId, context: FormulaContext): SeriesRange | null {
  const match = /^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?\$?([A-Z]{1,3})\$?([1-9][0-9]*)(?::\$?([A-Z]{1,3})\$?([1-9][0-9]*))?$/i.exec(source.trim());
  if (!match) {
    return null;
  }
  const sheetName = (match[1]?.replace(/''/g, "'") ?? match[2])?.trim();
  const worksheetId = sheetName
    ? context.worksheetIdsByName.get(sheetName.toLowerCase())?.[0]
    : currentWorksheetId;
  const projection = worksheetId ? context.projectionsByWorksheetId.get(worksheetId) : null;
  if (!worksheetId || !projection) {
    return null;
  }
  const startColumn = projection.columns[columnLabelToIndex(match[3].toUpperCase())];
  const startRow = projection.rows[Number.parseInt(match[4], 10) - 1];
  const endColumn = projection.columns[columnLabelToIndex((match[5] ?? match[3]).toUpperCase())];
  const endRow = projection.rows[Number.parseInt(match[6] ?? match[4], 10) - 1];
  if (!startColumn || !startRow || !endColumn || !endRow) {
    return null;
  }
  return {
    worksheetId,
    startRowId: startRow.id,
    endRowId: endRow.id,
    startColumnId: startColumn.id,
    endColumnId: endColumn.id,
    excelA1: source.trim(),
  };
}

function quoteSheetName(sheetName: string) {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName)
    ? sheetName
    : `'${sheetName.replace(/'/g, "''")}'`;
}
