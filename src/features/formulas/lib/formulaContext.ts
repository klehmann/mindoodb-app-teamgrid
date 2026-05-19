import FormulaParser from "fast-formula-parser";

import {
  createCellId,
  type ColumnId,
  type FormulaCell,
  type FormulaReference,
  type FormulaSegment,
  type RowId,
  type Workbook,
  type Worksheet,
  type WorksheetId,
} from "@/features/document/lib/teamgridDocument";
import { columnIndexToLabel, columnLabelToIndex, projectWorksheet, type GridProjection } from "@/features/grid/lib/gridProjection";

const ParserFormulaError = FormulaParser.FormulaError;

export interface FormulaContext {
  workbook: Workbook;
  projectionsByWorksheetId: Map<WorksheetId, GridProjection>;
  sheetNameById: Map<WorksheetId, string>;
  worksheetIdsByName: Map<string, WorksheetId[]>;
}

export interface ParsedFormulaSegments {
  source: string;
  segments: FormulaSegment[];
  references: FormulaReference[];
}

export function createFormulaContext(workbook: Workbook, sheetNameById?: Map<WorksheetId, string> | Record<WorksheetId, string>): FormulaContext {
  const projectionsByWorksheetId = new Map<WorksheetId, GridProjection>();
  const resolvedSheetNameById = new Map<WorksheetId, string>();
  const worksheetIdsByName = new Map<string, WorksheetId[]>();

  for (const worksheetId of workbook.worksheetOrder) {
    const worksheet = workbook.worksheetsById[worksheetId];
    if (!worksheet || worksheet.deletedAt) {
      continue;
    }
    projectionsByWorksheetId.set(worksheetId, projectWorksheet(worksheet));
    const sheetName = readSheetName(sheetNameById, worksheetId) ?? worksheet.title;
    resolvedSheetNameById.set(worksheetId, sheetName);
    const normalized = normalizeSheetName(sheetName);
    worksheetIdsByName.set(normalized, [...(worksheetIdsByName.get(normalized) ?? []), worksheetId]);
  }

  return {
    workbook,
    projectionsByWorksheetId,
    sheetNameById: resolvedSheetNameById,
    worksheetIdsByName,
  };
}

export function createSingleWorksheetFormulaContext(worksheet: Worksheet): FormulaContext {
  return createFormulaContext({
    id: "book_formula_context",
    worksheetOrder: [worksheet.id],
    worksheetsById: { [worksheet.id]: worksheet },
  } as Workbook);
}

export function parseFormulaSegments(source: string, currentWorksheetId: WorksheetId, context: FormulaContext): ParsedFormulaSegments {
  const normalizedSource = normalizeFormulaSource(source);
  const body = stripFormulaPrefix(normalizedSource);
  const bodySegments = parseFormulaBodySegments(body, currentWorksheetId, context);
  return {
    source: normalizedSource,
    segments: withFormulaPrefix(bodySegments),
    references: dedupeFormulaReferences(bodySegments.flatMap((segment) => segment.kind === "reference" ? [segment.reference] : [])),
  };
}

export function renderFormulaSource(formula: Pick<FormulaCell, "source" | "segments">, currentWorksheetId: WorksheetId, context: FormulaContext) {
  if (!formula.segments || formula.segments.length === 0) {
    return normalizeFormulaSource(formula.source);
  }
  return renderFormulaSegments(formula.segments, currentWorksheetId, context);
}

export function renderFormulaSegments(segments: FormulaSegment[], currentWorksheetId: WorksheetId, context: FormulaContext) {
  return segments.map((segment) => (
    segment.kind === "text" ? segment.text : formatFormulaReference(segment.reference, currentWorksheetId, context)
  )).join("");
}

export function renderFormulaParserSource(segments: FormulaSegment[], currentWorksheetId: WorksheetId, context: FormulaContext) {
  return segments.map((segment) => (
    segment.kind === "text" ? segment.text : formatFormulaReference(segment.reference, currentWorksheetId, context, "id")
  )).join("");
}

export function resolveWorksheetForFormulaSheet(sheet: string | undefined, currentWorksheetId: WorksheetId, context: FormulaContext) {
  if (!sheet) {
    return getVisibleWorksheet(currentWorksheetId, context);
  }
  if (context.workbook.worksheetsById[sheet as WorksheetId] && !context.workbook.worksheetsById[sheet as WorksheetId]?.deletedAt) {
    return context.workbook.worksheetsById[sheet as WorksheetId];
  }
  const matches = context.worksheetIdsByName.get(normalizeSheetName(sheet)) ?? [];
  if (matches.length !== 1) {
    throw ParserFormulaError.REF;
  }
  return getVisibleWorksheet(matches[0], context);
}

export function getProjectionForWorksheet(worksheetId: WorksheetId, context: FormulaContext) {
  const projection = context.projectionsByWorksheetId.get(worksheetId);
  if (!projection) {
    throw ParserFormulaError.REF;
  }
  return projection;
}

export function normalizeFormulaSource(source: string) {
  const trimmed = source.trim();
  return trimmed.startsWith("=") ? trimmed : `=${trimmed}`;
}

export function stripFormulaPrefix(source: string) {
  const trimmed = source.trim();
  return trimmed.startsWith("=") ? trimmed.slice(1) : trimmed;
}

function parseFormulaBodySegments(body: string, currentWorksheetId: WorksheetId, context: FormulaContext): FormulaSegment[] {
  const segments: FormulaSegment[] = [];
  let textStart = 0;
  let index = 0;

  while (index < body.length) {
    if (body[index] === "\"") {
      index = skipStringLiteral(body, index);
      continue;
    }
    if (!isReferenceStartBoundary(body, index)) {
      index += 1;
      continue;
    }
    const parsed = parseReferenceToken(body, index, currentWorksheetId, context);
    if (!parsed) {
      index += 1;
      continue;
    }
    if (textStart < index) {
      segments.push({ kind: "text", text: body.slice(textStart, index) });
    }
    segments.push({ kind: "reference", reference: parsed.reference });
    index = parsed.end;
    textStart = parsed.end;
  }

  if (textStart < body.length) {
    segments.push({ kind: "text", text: body.slice(textStart) });
  }
  return segments;
}

function parseReferenceToken(body: string, start: number, currentWorksheetId: WorksheetId, context: FormulaContext) {
  const sheet = parseSheetPrefix(body, start);
  const referenceStart = sheet?.end ?? start;
  const referenceText = body.slice(referenceStart);
  const worksheet = resolveWorksheetForFormulaSheet(sheet?.name, currentWorksheetId, context);
  const projection = getProjectionForWorksheet(worksheet.id, context);
  const columnRangeMatch = /^\$?([A-Z]{1,3}):\$?([A-Z]{1,3})(?![A-Z0-9_])/i.exec(referenceText);
  if (columnRangeMatch) {
    const reference = parseColumnRangeReference(columnRangeMatch[1], columnRangeMatch[2], worksheet.id, projection);
    return { end: referenceStart + columnRangeMatch[0].length, reference };
  }

  const cellRangeMatch = /^\$?([A-Z]{1,3})\$?([1-9][0-9]*)(?::\$?([A-Z]{1,3})\$?([1-9][0-9]*))?(?![A-Z0-9_])/i.exec(referenceText);
  if (!cellRangeMatch) {
    return null;
  }
  const reference = parseCellOrRangeReference(cellRangeMatch, worksheet.id, projection);
  return { end: referenceStart + cellRangeMatch[0].length, reference };
}

function parseSheetPrefix(body: string, start: number) {
  if (body[start] === "'") {
    let name = "";
    for (let index = start + 1; index < body.length; index += 1) {
      if (body[index] !== "'") {
        name += body[index];
        continue;
      }
      if (body[index + 1] === "'") {
        name += "'";
        index += 1;
        continue;
      }
      if (body[index + 1] === "!") {
        return { name, end: index + 2 };
      }
      return null;
    }
    return null;
  }

  const bangIndex = body.indexOf("!", start);
  if (bangIndex > start) {
    const candidate = body.slice(start, bangIndex).trim();
    const afterBang = body.slice(bangIndex + 1);
    if (candidate && /^[A-Za-z_][A-Za-z0-9_. ]*$/.test(candidate) && startsWithReference(afterBang)) {
      return { name: candidate, end: bangIndex + 1 };
    }
  }

  const match = /^([A-Za-z_][A-Za-z0-9_.]*)!/.exec(body.slice(start));
  return match ? { name: match[1], end: start + match[0].length } : null;
}

function startsWithReference(source: string) {
  return /^\$?[A-Z]{1,3}(:\$?[A-Z]{1,3}|\$?[1-9][0-9]*(?::\$?[A-Z]{1,3}\$?[1-9][0-9]*)?)(?![A-Z0-9_])/i.test(source);
}

function parseCellOrRangeReference(match: RegExpExecArray, worksheetId: WorksheetId, projection: GridProjection): FormulaReference {
  const startRowIndex = Number.parseInt(match[2], 10) - 1;
  const startColumnIndex = columnLabelToIndex(match[1].toUpperCase());
  const startRow = projection.rows[startRowIndex];
  const startColumn = projection.columns[startColumnIndex];
  if (!startRow || !startColumn) {
    throw ParserFormulaError.REF;
  }
  if (!match[3] || !match[4]) {
    return { kind: "cell", worksheetId, rowId: startRow.id, columnId: startColumn.id };
  }

  const endRowIndex = Number.parseInt(match[4], 10) - 1;
  const endColumnIndex = columnLabelToIndex(match[3].toUpperCase());
  const endRow = projection.rows[endRowIndex];
  const endColumn = projection.columns[endColumnIndex];
  if (!endRow || !endColumn) {
    throw ParserFormulaError.REF;
  }
  return {
    kind: "range",
    worksheetId,
    startRowId: startRow.id,
    endRowId: endRow.id,
    startColumnId: startColumn.id,
    endColumnId: endColumn.id,
  };
}

function parseColumnRangeReference(startColumnLabel: string, endColumnLabel: string, worksheetId: WorksheetId, projection: GridProjection): FormulaReference {
  const startColumn = projection.columns[columnLabelToIndex(startColumnLabel.toUpperCase())];
  const endColumn = projection.columns[columnLabelToIndex(endColumnLabel.toUpperCase())];
  if (!startColumn || !endColumn) {
    throw ParserFormulaError.REF;
  }
  if (startColumn.id === endColumn.id) {
    return { kind: "column", worksheetId, columnId: startColumn.id };
  }
  const firstRow = projection.rows[0];
  const lastRow = projection.rows[projection.rows.length - 1];
  if (!firstRow || !lastRow) {
    throw ParserFormulaError.REF;
  }
  return {
    kind: "range",
    worksheetId,
    startRowId: firstRow.id,
    endRowId: lastRow.id,
    startColumnId: startColumn.id,
    endColumnId: endColumn.id,
  };
}

function formatFormulaReference(reference: FormulaReference, currentWorksheetId: WorksheetId, context: FormulaContext, sheetMode: "title" | "id" = "title") {
  const projection = getProjectionForWorksheet(reference.worksheetId, context);
  const sheetName = sheetMode === "id" ? reference.worksheetId : context.sheetNameById.get(reference.worksheetId) ?? "#REF";
  const prefix = reference.worksheetId === currentWorksheetId ? "" : `${quoteSheetName(sheetName)}!`;
  if (reference.kind === "cell") {
    return `${prefix}${formatCellReference(projection, reference.rowId, reference.columnId)}`;
  }
  if (reference.kind === "column") {
    const columnIndex = projection.columnIndexById.get(reference.columnId);
    if (columnIndex == null) {
      return "#REF!";
    }
    const label = columnIndexToLabel(columnIndex);
    return `${prefix}${label}:${label}`;
  }
  return `${prefix}${formatCellReference(projection, reference.startRowId, reference.startColumnId)}:${formatCellReference(projection, reference.endRowId, reference.endColumnId)}`;
}

function formatCellReference(projection: GridProjection, rowId: RowId, columnId: ColumnId) {
  const address = projection.cellAddressById.get(createCellId(rowId, columnId));
  if (!address) {
    return "#REF!";
  }
  return address;
}

function quoteSheetName(sheetName: string) {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName)
    ? sheetName
    : `'${sheetName.replace(/'/g, "''")}'`;
}

function withFormulaPrefix(segments: FormulaSegment[]) {
  if (segments[0]?.kind === "text" && segments[0].text.startsWith("=")) {
    return segments;
  }
  return [{ kind: "text", text: "=" }, ...segments] satisfies FormulaSegment[];
}

function skipStringLiteral(source: string, start: number) {
  for (let index = start + 1; index < source.length; index += 1) {
    if (source[index] !== "\"") {
      continue;
    }
    if (source[index + 1] === "\"") {
      index += 1;
      continue;
    }
    return index + 1;
  }
  return source.length;
}

function isReferenceStartBoundary(source: string, index: number) {
  const previous = source[index - 1];
  return !previous || !/[A-Za-z0-9_.$]/.test(previous);
}

function getVisibleWorksheet(worksheetId: WorksheetId, context: FormulaContext) {
  const worksheet = context.workbook.worksheetsById[worksheetId];
  if (!worksheet || worksheet.deletedAt) {
    throw ParserFormulaError.REF;
  }
  return worksheet;
}

function normalizeSheetName(sheetName: string) {
  return sheetName.trim().toLowerCase();
}

function readSheetName(sheetNameById: Map<WorksheetId, string> | Record<WorksheetId, string> | undefined, worksheetId: WorksheetId) {
  if (!sheetNameById) {
    return undefined;
  }
  return sheetNameById instanceof Map ? sheetNameById.get(worksheetId) : sheetNameById[worksheetId];
}

function dedupeFormulaReferences(references: FormulaReference[]) {
  const seen = new Set<string>();
  return references.filter((reference) => {
    const key = JSON.stringify(reference);
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}
