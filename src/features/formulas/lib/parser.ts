/**
 * Teamgrid formula parser facade.
 *
 * `fast-formula-parser` owns the Excel-like grammar. This module preserves
 * Teamgrid's public parse shape and converts positional parser dependencies
 * back into stable row/column IDs for Automerge storage.
 */
import type { FormulaErrorCode, FormulaReference, FormulaSegment, WorksheetId } from "@/features/document/lib/teamgridDocument";
import { parseFormulaWithFastParser } from "@/features/formulas/lib/fastFormulaParserAdapter";
import type { FormulaContext } from "@/features/formulas/lib/formulaContext";

/** Placeholder AST retained for compatibility with the previous public type. */
export type FormulaNode = { kind: "external"; source: string };

/**
 * Successful parse result. `source` is canonicalized to always start with `=`
 * so it round-trips back into the formula bar.
 */
export interface ParsedFormula {
  source: string;
  ast: FormulaNode;
  segments: FormulaSegment[];
  references: FormulaReference[];
}

/** Returned instead of an AST when the source is malformed or unsupported. */
export interface FormulaParseError {
  code: FormulaErrorCode;
  message: string;
}

/**
 * Parse a formula string and collect stable-ID references.
 *
 * Accepts both `=SUM(A1:A3)` and `SUM(A1:A3)` shapes. Errors are returned as
 * values (`FormulaParseError`) rather than thrown, which keeps formula-bar
 * reactivity straightforward.
 */
export function parseFormula(source: string, worksheetId: WorksheetId, context: FormulaContext): ParsedFormula | FormulaParseError {
  const parsed = parseFormulaWithFastParser(source, worksheetId, context);
  if ("code" in parsed) {
    return parsed;
  }
  return {
    source: parsed.source,
    ast: { kind: "external", source: parsed.source },
    segments: parsed.segments,
    references: parsed.references,
  };
}
