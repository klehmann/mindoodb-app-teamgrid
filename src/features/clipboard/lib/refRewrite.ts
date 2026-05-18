/**
 * Formula reference rewriting for copy/paste and cut move-tracking.
 *
 * Teamgrid's parser stores stable IDs after commit, but clipboard formulas
 * travel as user-facing `A1` source text. During paste we rewrite those source
 * references by row/column deltas before re-parsing them against the target
 * worksheet projection.
 */
import { columnIndexToLabel, columnLabelToIndex } from "@/features/grid/lib/gridProjection";

export interface ReferenceDelta {
  rows: number;
  cols: number;
}

export interface ReferenceRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export interface RewriteOptions {
  insideRange?: ReferenceRange;
}

export function rewriteFormulaSource(source: string, delta: ReferenceDelta, options: RewriteOptions = {}) {
  return rewriteOutsideStrings(source, (segment) =>
    segment.replace(/\b([A-Z]+)([1-9][0-9]*)(?::([A-Z]+)([1-9][0-9]*))?\b/gi, (match, startColLabel: string, startRowText: string, endColLabel?: string, endRowText?: string) => {
      const start = parsePosition(startColLabel, startRowText);
      const end = endColLabel && endRowText ? parsePosition(endColLabel, endRowText) : null;

      if (options.insideRange && (!isInsideRange(start, options.insideRange) || (end && !isInsideRange(end, options.insideRange)))) {
        return match;
      }

      const shiftedStart = shiftPosition(start, delta);
      const shiftedEnd = end ? shiftPosition(end, delta) : null;
      if (!shiftedStart || (end && !shiftedEnd)) {
        return "#REF!";
      }
      return shiftedEnd ? `${formatPosition(shiftedStart)}:${formatPosition(shiftedEnd)}` : formatPosition(shiftedStart);
    }));
}

function rewriteOutsideStrings(source: string, rewriteSegment: (segment: string) => string) {
  let result = "";
  let segmentStart = 0;
  let inString = false;

  for (let index = 0; index < source.length; index += 1) {
    if (source[index] !== '"') {
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

function parsePosition(columnLabel: string, rowText: string) {
  return {
    row: Number.parseInt(rowText, 10) - 1,
    col: columnLabelToIndex(columnLabel.toUpperCase()),
  };
}

function shiftPosition(position: { row: number; col: number }, delta: ReferenceDelta) {
  const row = position.row + delta.rows;
  const col = position.col + delta.cols;
  return row < 0 || col < 0 ? null : { row, col };
}

function formatPosition(position: { row: number; col: number }) {
  return `${columnIndexToLabel(position.col)}${position.row + 1}`;
}

function isInsideRange(position: { row: number; col: number }, range: ReferenceRange) {
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);
  return position.row >= startRow && position.row <= endRow && position.col >= startCol && position.col <= endCol;
}
