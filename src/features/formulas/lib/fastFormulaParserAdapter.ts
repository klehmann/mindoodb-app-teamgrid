import FormulaParser, { type FormulaCellRef, type FormulaRangeRef } from "fast-formula-parser";

import { createCellId, type CellValue, type FormulaErrorCode, type FormulaReference, type FormulaResult, type WorksheetId } from "@/features/document/lib/teamgridDocument";
import { getCell, type GridProjection } from "@/features/grid/lib/gridProjection";
import {
  getProjectionForWorksheet,
  normalizeFormulaSource,
  parseFormulaSegments,
  renderFormulaParserSource,
  renderFormulaSource,
  resolveWorksheetForFormulaSheet,
  stripFormulaPrefix,
  type FormulaContext,
} from "@/features/formulas/lib/formulaContext";

export interface AdapterParsedFormula {
  source: string;
  segments: ReturnType<typeof parseFormulaSegments>["segments"];
  references: FormulaReference[];
}

export interface AdapterFormulaError {
  code: FormulaErrorCode;
  message: string;
}

export interface AdapterEvaluatedFormula {
  result: FormulaResult;
  segments?: AdapterParsedFormula["segments"];
  references: FormulaReference[];
  errorMessage?: string;
}

const DEFAULT_POSITION = { row: 1, col: 1 };
const ParserFormulaError = FormulaParser.FormulaError;
const CUSTOM_FUNCTIONS = {
  MIN: (...args: unknown[]) => aggregateNumbers(args, (numbers) => Math.min(...numbers)),
  MAX: (...args: unknown[]) => aggregateNumbers(args, (numbers) => Math.max(...numbers)),
  UPPER: (value: unknown) => String(unwrapParserArg(value) ?? "").toUpperCase(),
};
export const SUPPORTED_FORMULA_NAMES = Object.freeze([
  ...new Set([
    ...new FormulaParser(undefined, true).supportedFunctions(),
    ...Object.keys(CUSTOM_FUNCTIONS),
  ]),
].sort());
const SUPPORTED_FUNCTION_NAMES = new Set(SUPPORTED_FORMULA_NAMES);

export { normalizeFormulaSource };

export function parseFormulaWithFastParser(source: string, worksheetId: WorksheetId, context: FormulaContext): AdapterParsedFormula | AdapterFormulaError {
  try {
    const parsedSegments = parseFormulaSegments(source, worksheetId, context);
    const renderedSource = renderFormulaSource(parsedSegments, worksheetId, context);
    const parserSource = renderFormulaParserSource(parsedSegments.segments, worksheetId, context);
    const body = stripFormulaPrefix(parserSource);
    validateFormula(body, worksheetId);
    return {
      source: renderedSource,
      segments: parsedSegments.segments,
      references: parsedSegments.references,
    };
  } catch (error) {
    return adapterErrorFromUnknown(error);
  }
}

export function evaluateFormulaWithFastParser(source: string, worksheetId: WorksheetId, context: FormulaContext): AdapterEvaluatedFormula {
  return evaluateFormulaInternal(source, worksheetId, context, new Set());
}

function evaluateFormulaInternal(source: string, worksheetId: WorksheetId, context: FormulaContext, stack: Set<string>): AdapterEvaluatedFormula {
  const parsed = parseFormulaWithFastParser(source, worksheetId, context);
  if ("code" in parsed) {
    return {
      result: { kind: "error", code: parsed.code },
      segments: [],
      references: [],
      errorMessage: parsed.message,
    };
  }

  try {
    const parser = new FormulaParser({
      functions: CUSTOM_FUNCTIONS,
      onVariable: () => null,
      onCell: (ref) => readParserCell(ref, worksheetId, context, stack),
      onRange: (ref) => readParserRange(ref, worksheetId, context, stack),
    });
    const value = parser.parse(normalizeSupportedFunctionNames(stripFormulaPrefix(renderFormulaParserSource(parsed.segments, worksheetId, context))), { ...DEFAULT_POSITION, sheet: worksheetId });
    const result = parserValueToFormulaResult(value);
    return {
      result,
      segments: parsed.segments,
      references: parsed.references,
      errorMessage: result.kind === "error" ? result.code : undefined,
    };
  } catch (error) {
    const adapterError = adapterErrorFromUnknown(error);
    return {
      result: { kind: "error", code: adapterError.code },
      segments: parsed.segments,
      references: parsed.references,
      errorMessage: adapterError.message,
    };
  }
}

function validateFormula(body: string, worksheetId: WorksheetId) {
  if (body.trim() === "") {
    throw new Error("Formula must not be empty.");
  }
  assertSupportedFunctions(body);
  const parserBody = normalizeSupportedFunctionNames(body);
  const depParser = new FormulaParser.DepParser({ onVariable: () => null });
  depParser.parse(parserBody, { ...DEFAULT_POSITION, sheet: worksheetId });
}

function readParserCell(ref: FormulaCellRef, currentWorksheetId: WorksheetId, context: FormulaContext, stack: Set<string>): unknown {
  const worksheet = resolveWorksheetForFormulaSheet(ref.sheet, currentWorksheetId, context);
  const projection = getProjectionForWorksheet(worksheet.id, context);
  const row = projection.rows[ref.row - 1];
  const column = projection.columns[ref.col - 1];
  if (!row || !column) {
    throw ParserFormulaError.REF;
  }
  const cell = getCell(worksheet, row.id, column.id);
  if (!cell.formula) {
    return cellValueToParserValue(cell.value);
  }

  const stackKey = `${worksheet.id}:${cell.id}`;
  if (stack.has(stackKey)) {
    throw new Error("#CYCLE!");
  }
  stack.add(stackKey);
  const evaluated = evaluateFormulaInternal(renderFormulaSource(cell.formula, worksheet.id, context), worksheet.id, context, stack);
  stack.delete(stackKey);
  if (evaluated.result.kind === "error") {
    throw new Error(evaluated.result.code);
  }
  return formulaResultToParserValue(evaluated.result);
}

function readParserRange(ref: FormulaRangeRef, currentWorksheetId: WorksheetId, context: FormulaContext, stack: Set<string>) {
  const worksheet = resolveWorksheetForFormulaSheet(ref.sheet, currentWorksheetId, context);
  const projection = getProjectionForWorksheet(worksheet.id, context);
  const bounds = normalizeRangeBounds(ref, projection);
  const values: unknown[][] = [];
  for (let rowIndex = bounds.startRow; rowIndex <= bounds.endRow; rowIndex += 1) {
    const rowValues: unknown[] = [];
    for (let columnIndex = bounds.startColumn; columnIndex <= bounds.endColumn; columnIndex += 1) {
      rowValues.push(readParserCell({ row: rowIndex + 1, col: columnIndex + 1, sheet: ref.sheet }, currentWorksheetId, context, stack));
    }
    values.push(rowValues);
  }
  return values;
}

function normalizeRangeBounds(ref: FormulaRangeRef, projection: GridProjection) {
  const fromRow = ref.from.row ?? 1;
  const fromColumn = ref.from.col ?? 1;
  const toRow = ref.to.row ?? FormulaParser.MAX_ROW;
  const toColumn = ref.to.col ?? FormulaParser.MAX_COLUMN;
  const coversAllRows = fromRow === 1 && toRow >= FormulaParser.MAX_ROW;
  const coversAllColumns = fromColumn === 1 && toColumn >= FormulaParser.MAX_COLUMN;
  const endRow = coversAllRows ? projection.rows.length : toRow;
  const endColumn = coversAllColumns ? projection.columns.length : toColumn;

  if (fromRow < 1 || fromColumn < 1 || endRow < fromRow || endColumn < fromColumn) {
      throw ParserFormulaError.REF;
  }
  if (!coversAllRows && endRow > projection.rows.length) {
    throw ParserFormulaError.REF;
  }
  if (!coversAllColumns && endColumn > projection.columns.length) {
    throw ParserFormulaError.REF;
  }
  if (projection.rows.length === 0 || projection.columns.length === 0) {
    throw ParserFormulaError.REF;
  }

  return {
    startRow: fromRow - 1,
    endRow: endRow - 1,
    startColumn: fromColumn - 1,
    endColumn: endColumn - 1,
    coversAllRows,
    coversAllColumns,
  };
}

function cellValueToParserValue(value: CellValue): unknown {
  switch (value.kind) {
    case "empty":
      return null;
    case "number":
      return value.value;
    case "string":
      return value.text;
    case "date":
      return dateToExcelSerial(new Date(value.isoDate));
  }
}

function formulaResultToParserValue(result: FormulaResult): unknown {
  switch (result.kind) {
    case "empty":
      return null;
    case "number":
      return result.value;
    case "string":
      return result.value;
    case "date":
      return dateToExcelSerial(new Date(result.isoDate));
    case "error":
      throw new Error(result.code);
  }
}

function parserValueToFormulaResult(value: unknown): FormulaResult {
  if (isParserFormulaError(value)) {
    return { kind: "error", code: formulaErrorCodeFromUnknown(value) };
  }
  if (typeof value === "number") {
    return Number.isFinite(value) ? { kind: "number", value } : { kind: "error", code: "#VALUE!" };
  }
  if (typeof value === "boolean") {
    return { kind: "string", value: value ? "TRUE" : "FALSE" };
  }
  if (typeof value === "string") {
    return { kind: "string", value };
  }
  if (value instanceof Date) {
    return { kind: "date", isoDate: value.toISOString() };
  }
  if (value == null) {
    return { kind: "empty" };
  }
  if (Array.isArray(value)) {
    return parserValueToFormulaResult(value[0]?.[0]);
  }
  return { kind: "error", code: "#VALUE!" };
}

function adapterErrorFromUnknown(error: unknown): AdapterFormulaError {
  const code = formulaErrorCodeFromUnknown(error);
  return {
    code,
    message: error instanceof Error ? error.message : code,
  };
}

function formulaErrorCodeFromUnknown(error: unknown): FormulaErrorCode {
  const text = error instanceof Error
    ? `${error.name} ${error.message}`
    : typeof error === "object" && error !== null && "name" in error
      ? String((error as { name?: unknown }).name)
      : String(error);
  if (text.includes("#CYCLE!")) return "#CYCLE!";
  if (text.includes("#DIV/0!")) return "#DIV/0!";
  if (text.includes("#REF!")) return "#REF!";
  if (text.includes("#NAME?") || text.includes("not implemented") || text.includes("NOT_IMPLEMENTED")) return "#NAME?";
  return "#VALUE!";
}

function isParserFormulaError(value: unknown): value is InstanceType<typeof ParserFormulaError> {
  return value instanceof ParserFormulaError
    || (typeof value === "object" && value !== null && "name" in value && String((value as { name?: unknown }).name).startsWith("#"));
}

function assertSupportedFunctions(body: string) {
  for (const name of collectFunctionNames(body)) {
    const normalizedName = name.toUpperCase().replace(/^_XLFN\./, "");
    if (!SUPPORTED_FUNCTION_NAMES.has(normalizedName)) {
      throw ParserFormulaError.NAME;
    }
  }
}

function collectFunctionNames(body: string) {
  const names = new Set<string>();
  let inString = false;
  for (let index = 0; index < body.length; index += 1) {
    const character = body[index];
    if (character === "\"") {
      if (body[index + 1] === "\"") {
        index += 1;
      } else {
        inString = !inString;
      }
      continue;
    }
    if (inString || !/[A-Za-z_]/.test(character)) {
      continue;
    }
    const match = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(body.slice(index));
    if (!match) {
      continue;
    }
    const nextIndex = index + match[0].length;
    const nextNonSpace = body.slice(nextIndex).match(/^\s*/)?.[0].length ?? 0;
    if (body[nextIndex + nextNonSpace] === "(") {
      names.add(match[0]);
    }
    index = nextIndex - 1;
  }
  return names;
}

function normalizeSupportedFunctionNames(body: string) {
  return rewriteOutsideStrings(body, (segment) =>
    segment.replace(/\b([A-Za-z_][A-Za-z0-9_.]*)\b(?=\s*\()/g, (match) => {
      const normalizedName = match.toUpperCase().replace(/^_XLFN\./, "");
      return SUPPORTED_FUNCTION_NAMES.has(normalizedName) ? match.toUpperCase() : match;
    }));
}

function rewriteOutsideStrings(source: string, rewriteSegment: (segment: string) => string) {
  let result = "";
  let segmentStart = 0;
  let inString = false;

  for (let index = 0; index < source.length; index += 1) {
    if (source[index] !== "\"") {
      continue;
    }
    if (!inString) {
      result += rewriteSegment(source.slice(segmentStart, index));
      segmentStart = index;
      inString = true;
    } else if (source[index + 1] === "\"") {
      index += 1;
    } else {
      result += source.slice(segmentStart, index + 1);
      segmentStart = index + 1;
      inString = false;
    }
  }
  return result + (inString ? source.slice(segmentStart) : rewriteSegment(source.slice(segmentStart)));
}

function aggregateNumbers(args: unknown[], aggregate: (numbers: number[]) => number) {
  const numbers = flattenParserArgs(args).filter((value): value is number => typeof value === "number" && Number.isFinite(value));
  if (numbers.length === 0) {
    throw ParserFormulaError.VALUE;
  }
  return aggregate(numbers);
}

function flattenParserArgs(values: unknown[]): unknown[] {
  return values.flatMap((value) => {
    const unwrapped = unwrapParserArg(value);
    return Array.isArray(unwrapped) ? flattenParserArgs(unwrapped) : [unwrapped];
  });
}

function unwrapParserArg(value: unknown): unknown {
  if (value && typeof value === "object" && "value" in value) {
    return (value as { value: unknown }).value;
  }
  return value;
}

function dateToExcelSerial(date: Date) {
  const excelEpoch = Date.UTC(1899, 11, 30);
  return (Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()) - excelEpoch) / 86400000;
}
