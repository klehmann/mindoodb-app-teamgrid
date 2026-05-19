/**
 * Teamgrid formula evaluator facade.
 *
 * Evaluation is delegated to `fast-formula-parser` through a local adapter so
 * third-party parser details do not leak into UI, clipboard, or XLSX code.
 */
import type { FormulaReference, FormulaResult, FormulaSegment, WorksheetId } from "@/features/document/lib/teamgridDocument";
import { evaluateFormulaWithFastParser } from "@/features/formulas/lib/fastFormulaParserAdapter";
import type { FormulaContext } from "@/features/formulas/lib/formulaContext";

/** Outcome of evaluating one cell's formula. */
export interface EvaluatedFormula {
  result: FormulaResult;
  segments?: FormulaSegment[];
  references: FormulaReference[];
  errorMessage?: string;
}

/**
 * Parse and evaluate a formula source string against a worksheet snapshot.
 *
 * Returns both the resolved {@link FormulaResult} and the collected stable-ID
 * references. Callers persist the references so the dependency graph can later
 * recompute dependents efficiently.
 */
export function evaluateFormula(source: string, worksheetId: WorksheetId, context: FormulaContext): EvaluatedFormula {
  return evaluateFormulaWithFastParser(source, worksheetId, context);
}
