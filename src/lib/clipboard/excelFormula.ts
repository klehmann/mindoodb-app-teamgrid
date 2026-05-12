/**
 * Excel formula conversion helpers for clipboard interop.
 *
 * Teamgrid stores formulas as user-facing A1 strings (for example `=B3*5`).
 * Excel's HTML clipboard format expects formulas in relative R1C1 notation
 * inside an `x:fmla` attribute on each `<td>`. R1C1 is intrinsically relative
 * to the cell that contains the formula, so Excel can shift references
 * automatically when the user pastes the HTML at a different target.
 *
 * This module owns the two-way conversion:
 *
 * - {@link a1FormulaToRelativeR1C1} turns a Teamgrid A1 formula into an
 *   `RC[-1]`-style string positioned for the formula's source cell. Used by
 *   the HTML serializer in {@link ./payload.ts}.
 * - {@link excelFormulaToA1} turns either an R1C1 formula or a "Excel A1
 *   relative to the top-left of the copied table" formula back into Teamgrid
 *   A1. Used by the Excel HTML importer in {@link ./payload.ts}.
 *
 * Both helpers preserve quoted string literals byte-for-byte so a formula
 * such as `=CONCAT("A1", A1)` never has its string content rewritten.
 *
 * Scope (matches the rest of the Teamgrid formula surface):
 * - Same-sheet references only. No cross-sheet syntax is supported.
 * - Absolute references (`$A$1`) are accepted on input but converted to plain
 *   relative addresses on output because the Teamgrid evaluator does not yet
 *   support absolute pinning.
 * - The functions operate purely on strings; the evaluator is what ultimately
 *   binds them to stable IDs after a paste.
 */
import { columnIndexToLabel, columnLabelToIndex } from "@/lib/gridProjection";

/** Zero-based row/column position in a worksheet projection. */
export interface FormulaPosition {
  row: number;
  col: number;
}

/** Optional configuration accepted by {@link excelFormulaToA1}. */
export interface ExcelFormulaToA1Options {
  /**
   * Translation applied to A1 formulas that Excel reports relative to the
   * copied HTML table's top-left cell. The Excel HTML clipboard sometimes
   * surfaces A1 references where `A1` means "first cell of the copied table"
   * rather than absolute sheet `A1`; the importer offsets every reference by
   * this `origin` so the resulting Teamgrid formula stays valid when applied
   * at any anchor.
   */
  origin?: FormulaPosition;
}

/**
 * Convert a Teamgrid A1 formula into Excel relative R1C1 notation.
 *
 * The returned string is anchored at the supplied `formulaCellPosition` so
 * Excel can paste it at any target and have references shift relative to the
 * new location.
 *
 * @example
 * ```ts
 * a1FormulaToRelativeR1C1("=B3*5", { row: 2, col: 2 }) // "=RC[-1]*5"
 * a1FormulaToRelativeR1C1("=B1*5", { row: 2, col: 2 }) // "=R[-2]C[-1]*5"
 * ```
 */
export function a1FormulaToRelativeR1C1(source: string, formulaCellPosition: FormulaPosition) {
  return rewriteOutsideStrings(source, (segment) =>
    segment.replace(a1RangePattern, (match, startCol: string, startRow: string, endCol?: string, endRow?: string) => {
      const start = a1ToPosition(startCol, startRow);
      const end = endCol && endRow ? a1ToPosition(endCol, endRow) : null;
      if (!start) {
        return match;
      }
      const startRef = positionToRelativeR1C1(start, formulaCellPosition);
      if (!end) {
        return startRef;
      }
      return `${startRef}:${positionToRelativeR1C1(end, formulaCellPosition)}`;
    }));
}

/**
 * Convert an Excel formula string into Teamgrid A1 notation.
 *
 * Excel may hand us two distinct shapes inside an `x:fmla` attribute:
 *
 * 1. R1C1 references (e.g. `=RC[-1]*5`). These are resolved against the
 *    supplied `formulaCellPosition`, which represents the cell that contains
 *    the formula in the Teamgrid coordinate system.
 * 2. A1 references that Excel reports relative to the copied HTML table's
 *    top-left cell. The {@link ExcelFormulaToA1Options.origin} option offsets
 *    each reference back to true sheet coordinates so the formula remains
 *    valid after paste.
 *
 * The function autodetects R1C1 by looking for at least one `R...C...`
 * reference outside of string literals; otherwise it falls back to the A1
 * branch.
 */
export function excelFormulaToA1(source: string, formulaCellPosition: FormulaPosition, options: ExcelFormulaToA1Options = {}) {
  if (hasR1C1Reference(source)) {
    return rewriteOutsideStrings(source, (segment) => rewriteR1C1Segment(segment, formulaCellPosition));
  }
  const origin = options.origin ?? { row: 0, col: 0 };
  return rewriteOutsideStrings(source, (segment) =>
    segment.replace(a1RangePattern, (match, startCol: string, startRow: string, endCol?: string, endRow?: string) => {
      const start = a1ToPosition(startCol, startRow);
      const end = endCol && endRow ? a1ToPosition(endCol, endRow) : null;
      if (!start) {
        return match;
      }
      const shiftedStart = positionToA1({ row: start.row + origin.row, col: start.col + origin.col });
      if (!end) {
        return shiftedStart;
      }
      return `${shiftedStart}:${positionToA1({ row: end.row + origin.row, col: end.col + origin.col })}`;
    }));
}

/**
 * Regex for A1 cell or range references with optional `$` absolute markers.
 * The `$` characters are accepted but dropped because Teamgrid formulas are
 * always relative.
 */
const a1RangePattern = /\$?([A-Z]+)\$?([1-9][0-9]*)(?::\$?([A-Z]+)\$?([1-9][0-9]*))?\b/gi;

/**
 * Regex for R1C1 cell or range references. Both row and column can be
 * relative (`R[-2]`), absolute (`R5`), or current (`R` / `C`). Capture groups
 * follow `start row marker, start row delta, start row absolute, start column
 * marker, start column delta, start column absolute, end row marker, ...`
 * which {@link r1c1ToPosition} consumes positionally.
 */
const r1c1RangePattern = /R(\[(-?[0-9]+)\]|([1-9][0-9]*))?C(\[(-?[0-9]+)\]|([1-9][0-9]*))?(?::R(\[(-?[0-9]+)\]|([1-9][0-9]*))?C(\[(-?[0-9]+)\]|([1-9][0-9]*))?)?/gi;

/**
 * Sniff for at least one R1C1 reference in the formula, ignoring string
 * literals. Used to pick the correct branch in {@link excelFormulaToA1}.
 */
function hasR1C1Reference(source: string) {
  let found = false;
  rewriteOutsideStrings(source, (segment) => {
    r1c1RangePattern.lastIndex = 0;
    found ||= r1c1RangePattern.test(segment);
    return segment;
  });
  return found;
}

/**
 * Replace every R1C1 cell/range reference in `segment` with its A1
 * equivalent, resolved against `formulaCellPosition`.
 */
function rewriteR1C1Segment(segment: string, formulaCellPosition: FormulaPosition) {
  r1c1RangePattern.lastIndex = 0;
  return segment.replace(r1c1RangePattern, (match, startRowRelative: string | undefined, startRowDelta: string | undefined, startRowAbsolute: string | undefined, startColRelative: string | undefined, startColDelta: string | undefined, startColAbsolute: string | undefined, endRowRelative?: string, endRowDelta?: string, endRowAbsolute?: string, endColRelative?: string, endColDelta?: string, endColAbsolute?: string) => {
    const start = r1c1ToPosition(formulaCellPosition, startRowRelative, startRowDelta, startRowAbsolute, startColRelative, startColDelta, startColAbsolute);
    if (!start) {
      return match;
    }
    if (!endRowRelative && !endRowDelta && !endRowAbsolute && !endColRelative && !endColDelta && !endColAbsolute) {
      return positionToA1(start);
    }
    const end = r1c1ToPosition(formulaCellPosition, endRowRelative, endRowDelta, endRowAbsolute, endColRelative, endColDelta, endColAbsolute);
    return end ? `${positionToA1(start)}:${positionToA1(end)}` : match;
  });
}

/**
 * Resolve one R1C1 row/column pair into a zero-based {@link FormulaPosition}.
 *
 * The capture groups (`rowPart`, `rowDelta`, `rowAbsolute`, and column twins)
 * map directly to {@link r1c1RangePattern}:
 *
 * - `rowPart == null` -> the reference omitted the row marker, so it inherits
 *   the formula cell's row (e.g. `RC[-1]` shares the row of its parent cell).
 * - `rowDelta != null` -> relative `R[delta]`. We add the delta to the base.
 * - `rowAbsolute != null` -> absolute `R5`. R1C1 is one-based, hence the `-1`.
 *
 * Returns `null` when the resolved position is negative, which the caller
 * surfaces as an unchanged literal so we never emit an invalid Teamgrid
 * formula.
 */
function r1c1ToPosition(
  base: FormulaPosition,
  rowPart: string | undefined,
  rowDelta: string | undefined,
  rowAbsolute: string | undefined,
  colPart: string | undefined,
  colDelta: string | undefined,
  colAbsolute: string | undefined,
) {
  const row = rowPart == null
    ? base.row
    : rowDelta != null
      ? base.row + Number(rowDelta)
      : Number(rowAbsolute) - 1;
  const col = colPart == null
    ? base.col
    : colDelta != null
      ? base.col + Number(colDelta)
      : Number(colAbsolute) - 1;
  return row < 0 || col < 0 ? null : { row, col };
}

/** Parse the column-letter / row-number captures of {@link a1RangePattern}. */
function a1ToPosition(columnLabel: string, rowText: string) {
  const row = Number(rowText) - 1;
  const col = columnLabelToIndex(columnLabel.toUpperCase());
  return row < 0 || col < 0 ? null : { row, col };
}

/** Format a zero-based position as an A1-style address (e.g. `B3`). */
function positionToA1(position: FormulaPosition) {
  return `${columnIndexToLabel(position.col)}${position.row + 1}`;
}

/**
 * Build a relative R1C1 reference (e.g. `R[-2]C[-1]`) from `position`
 * measured against the formula cell. A zero delta collapses to a bare `R` or
 * `C`, matching Excel's preferred shorthand.
 */
function positionToRelativeR1C1(position: FormulaPosition, formulaCellPosition: FormulaPosition) {
  const rowDelta = position.row - formulaCellPosition.row;
  const colDelta = position.col - formulaCellPosition.col;
  return `${relativePart("R", rowDelta)}${relativePart("C", colDelta)}`;
}

/** Emit one half (row or column) of a relative R1C1 reference. */
function relativePart(prefix: "R" | "C", delta: number) {
  return delta === 0 ? prefix : `${prefix}[${delta}]`;
}

/**
 * Apply `rewriteSegment` to every span of `source` that is *outside* a
 * double-quoted string literal. Excel formulas embed string arguments
 * verbatim, and we must not rewrite, for example, `"A1"` inside `CONCAT`.
 *
 * Doubled quotes (`""`) inside a string are Excel's escape sequence for a
 * literal quote character; we keep them as-is and stay inside the string.
 */
function rewriteOutsideStrings(source: string, rewriteSegment: (segment: string) => string) {
  let result = "";
  let segmentStart = 0;
  let inString = false;

  for (let index = 0; index < source.length; index += 1) {
    if (source[index] !== "\"") {
      continue;
    }
    if (source[index + 1] === "\"") {
      index += 1;
      continue;
    }
    if (!inString) {
      result += rewriteSegment(source.slice(segmentStart, index));
      segmentStart = index;
      inString = true;
    } else {
      result += source.slice(segmentStart, index + 1);
      segmentStart = index + 1;
      inString = false;
    }
  }
  result += inString ? source.slice(segmentStart) : rewriteSegment(source.slice(segmentStart));
  return result;
}
