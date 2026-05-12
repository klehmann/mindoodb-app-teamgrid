/**
 * Tree-walking evaluator for Teamgrid formulas.
 *
 * Parses the source through {@link parseFormula} and then recurses over the
 * resulting AST. Each node returns a {@link FormulaRuntimeValue}, which is a
 * superset of {@link CellValue} that also carries error codes so the first
 * error short-circuits its enclosing function call.
 *
 * Cycle detection
 * ---------------
 * When a `cell` node references another cell that itself contains a formula,
 * the evaluator recurses into that formula. The `stack` set tracks which
 * cell ids are currently being evaluated, so a self-referential or
 * mutually-recursive formula yields `#CYCLE!` instead of blowing the JS
 * call stack.
 *
 * Range handling
 * --------------
 * Ranges flatten to an array of cell values. Functions decide whether they
 * want to consume a range (`SUM`, `AVERAGE`, ...) or treat it as a scalar.
 * Top-level ranges outside of function arguments are uncommon in
 * spreadsheets and currently produce a placeholder runtime string; functions
 * receive ranges as expanded arguments via {@link evaluateNode}'s `flatMap`.
 */
import { getCell, type GridProjection } from "@/lib/gridProjection";
import type { CellValue, FormulaErrorCode, FormulaReference, FormulaResult, Worksheet } from "@/lib/teamgridDocument";
import { FUNCTION_REGISTRY, type FormulaRuntimeValue } from "@/lib/formulas/functionRegistry";
import { type FormulaNode, parseFormula } from "@/lib/formulas/parser";

/** Outcome of evaluating one cell's formula. */
export interface EvaluatedFormula {
  result: FormulaResult;
  references: FormulaReference[];
  errorMessage?: string;
}

/**
 * Parse and evaluate a formula source string against a worksheet snapshot.
 *
 * Returns both the resolved {@link FormulaResult} and the collected
 * {@link FormulaReference}s. Callers persist the references so the
 * dependency graph can later recompute dependents efficiently.
 */
export function evaluateFormula(source: string, worksheet: Worksheet, projection: GridProjection): EvaluatedFormula {
  const parsed = parseFormula(source, worksheet.id, projection);
  if ("code" in parsed) {
    return {
      result: { kind: "error", code: parsed.code },
      references: [],
      errorMessage: parsed.message,
    };
  }
  const value = evaluateNode(parsed.ast, worksheet, projection, new Set());
  return {
    result: runtimeToFormulaResult(value),
    references: parsed.references,
  };
}

function evaluateNode(node: FormulaNode, worksheet: Worksheet, projection: GridProjection, stack: Set<string>): FormulaRuntimeValue {
  switch (node.kind) {
    case "number":
      return { kind: "number", value: node.value };
    case "string":
      return { kind: "string", value: node.value };
    case "cell": {
      const cell = getCell(worksheet, node.reference.rowId, node.reference.columnId);
      if (stack.has(cell.id)) {
        return { kind: "error", code: "#CYCLE!" };
      }
      if (cell.formula) {
        stack.add(cell.id);
        const evaluated = evaluateFormula(cell.formula.source, worksheet, projection);
        stack.delete(cell.id);
        return formulaResultToRuntime(evaluated.result);
      }
      return cellValueToRuntime(cell.value);
    }
    case "range":
      return { kind: "string", value: JSON.stringify(expandRange(node, worksheet, projection).map(cellValueToRuntime)) };
    case "function": {
      const definition = FUNCTION_REGISTRY[node.name];
      if (!definition) {
        return { kind: "error", code: "#NAME?" };
      }
      if (node.args.length < definition.minArgs || (definition.maxArgs != null && node.args.length > definition.maxArgs)) {
        return { kind: "error", code: "#VALUE!" };
      }
      const args = node.args.flatMap((arg) => arg.kind === "range"
        ? expandRange(arg, worksheet, projection).map(cellValueToRuntime)
        : [evaluateNode(arg, worksheet, projection, stack)]);
      const firstError = args.find((arg) => arg.kind === "error");
      return firstError ?? definition.evaluate(args);
    }
    case "binary":
      return evaluateBinary(
        node.operator,
        evaluateNode(node.left, worksheet, projection, stack),
        evaluateNode(node.right, worksheet, projection, stack),
      );
  }
}

function expandRange(node: Extract<FormulaNode, { kind: "range" }>, worksheet: Worksheet, projection: GridProjection): CellValue[] {
  const startRow = projection.rowIndexById.get(node.reference.startRowId);
  const endRow = projection.rowIndexById.get(node.reference.endRowId);
  const startColumn = projection.columnIndexById.get(node.reference.startColumnId);
  const endColumn = projection.columnIndexById.get(node.reference.endColumnId);
  if (startRow == null || endRow == null || startColumn == null || endColumn == null) {
    return [{ kind: "string", text: "#REF!" }];
  }

  const rowMin = Math.min(startRow, endRow);
  const rowMax = Math.max(startRow, endRow);
  const columnMin = Math.min(startColumn, endColumn);
  const columnMax = Math.max(startColumn, endColumn);
  const values: CellValue[] = [];
  for (let rowIndex = rowMin; rowIndex <= rowMax; rowIndex += 1) {
    for (let columnIndex = columnMin; columnIndex <= columnMax; columnIndex += 1) {
      const row = projection.rows[rowIndex];
      const column = projection.columns[columnIndex];
      if (row && column) {
        values.push(getCell(worksheet, row.id, column.id).value);
      }
    }
  }
  return values;
}

function evaluateBinary(operator: "+" | "-" | "*" | "/", left: FormulaRuntimeValue, right: FormulaRuntimeValue): FormulaRuntimeValue {
  if (left.kind === "error") return left;
  if (right.kind === "error") return right;
  const leftNumber = runtimeToNumber(left);
  const rightNumber = runtimeToNumber(right);
  if (leftNumber == null || rightNumber == null) {
    return { kind: "error", code: "#VALUE!" };
  }
  if (operator === "/" && rightNumber === 0) {
    return { kind: "error", code: "#DIV/0!" };
  }
  switch (operator) {
    case "+":
      return { kind: "number", value: leftNumber + rightNumber };
    case "-":
      return { kind: "number", value: leftNumber - rightNumber };
    case "*":
      return { kind: "number", value: leftNumber * rightNumber };
    case "/":
      return { kind: "number", value: leftNumber / rightNumber };
  }
}

function cellValueToRuntime(value: CellValue): FormulaRuntimeValue {
  switch (value.kind) {
    case "empty":
      return { kind: "empty" };
    case "string":
      return { kind: "string", value: value.text };
    case "number":
      return { kind: "number", value: value.value };
    case "date":
      return { kind: "date", isoDate: value.isoDate };
  }
}

function formulaResultToRuntime(result: FormulaResult): FormulaRuntimeValue {
  switch (result.kind) {
    case "empty":
      return { kind: "empty" };
    case "string":
      return { kind: "string", value: result.value };
    case "number":
      return { kind: "number", value: result.value };
    case "date":
      return { kind: "date", isoDate: result.isoDate };
    case "error":
      return { kind: "error", code: result.code };
  }
}

function runtimeToFormulaResult(value: FormulaRuntimeValue): FormulaResult {
  switch (value.kind) {
    case "empty":
      return { kind: "empty" };
    case "string":
      return { kind: "string", value: value.value };
    case "number":
      return { kind: "number", value: value.value };
    case "date":
      return { kind: "date", isoDate: value.isoDate };
    case "error":
      return { kind: "error", code: value.code as FormulaErrorCode };
  }
}

function runtimeToNumber(value: FormulaRuntimeValue) {
  if (value.kind === "number") {
    return value.value;
  }
  if (value.kind === "empty") {
    return 0;
  }
  if (value.kind === "string") {
    const number = Number(value.value);
    return Number.isFinite(number) ? number : null;
  }
  return null;
}
