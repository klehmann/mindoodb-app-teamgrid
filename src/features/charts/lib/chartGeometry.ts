import { DEFAULT_ROW_HEIGHT } from "@/shared/lib/gridDimensions";
import type { GridProjection } from "@/features/grid/lib/gridProjection";
import type { TwoCellAnchor, Worksheet } from "@/features/document/lib/teamgridDocument";

const EMUS_PER_PIXEL = 9525;
const ROW_HEADER_WIDTH = 48;
const COLUMN_HEADER_HEIGHT = 32;

export interface ChartRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export function chartAnchorToRect(anchor: TwoCellAnchor, worksheet: Worksheet, projection: GridProjection): ChartRect | null {
  const fromColumnIndex = projection.columnIndexById.get(anchor.from.columnId);
  const toColumnIndex = projection.columnIndexById.get(anchor.to.columnId);
  const fromRowIndex = projection.rowIndexById.get(anchor.from.rowId);
  const toRowIndex = projection.rowIndexById.get(anchor.to.rowId);
  if (fromColumnIndex == null || toColumnIndex == null || fromRowIndex == null || toRowIndex == null) {
    return null;
  }
  const x1 = ROW_HEADER_WIDTH + sumColumnWidths(projection, 0, fromColumnIndex) + emuToPixels(anchor.from.colOffsetEmu);
  const x2 = ROW_HEADER_WIDTH + sumColumnWidths(projection, 0, toColumnIndex) + emuToPixels(anchor.to.colOffsetEmu);
  const y1 = COLUMN_HEADER_HEIGHT + sumRowHeights(worksheet, projection, 0, fromRowIndex) + emuToPixels(anchor.from.rowOffsetEmu);
  const y2 = COLUMN_HEADER_HEIGHT + sumRowHeights(worksheet, projection, 0, toRowIndex) + emuToPixels(anchor.to.rowOffsetEmu);
  return {
    x: Math.min(x1, x2),
    y: Math.min(y1, y2),
    width: Math.max(120, Math.abs(x2 - x1)),
    height: Math.max(90, Math.abs(y2 - y1)),
  };
}

export function rectToChartAnchor(rect: ChartRect, worksheet: Worksheet, projection: GridProjection): TwoCellAnchor | null {
  const from = pixelPointToAnchor(rect.x - ROW_HEADER_WIDTH, rect.y - COLUMN_HEADER_HEIGHT, worksheet, projection);
  const to = pixelPointToAnchor(rect.x + rect.width - ROW_HEADER_WIDTH, rect.y + rect.height - COLUMN_HEADER_HEIGHT, worksheet, projection);
  return from && to ? { from, to } : null;
}

export function emuToPixels(emu: number) {
  return emu / EMUS_PER_PIXEL;
}

export function pixelsToEmu(pixels: number) {
  return Math.max(0, Math.round(pixels * EMUS_PER_PIXEL));
}

function pixelPointToAnchor(x: number, y: number, worksheet: Worksheet, projection: GridProjection) {
  const column = pointToColumn(x, projection);
  const row = pointToRow(y, worksheet, projection);
  if (!column || !row) {
    return null;
  }
  return {
    rowId: row.id,
    columnId: column.id,
    rowOffsetEmu: pixelsToEmu(row.offset),
    colOffsetEmu: pixelsToEmu(column.offset),
  };
}

function pointToColumn(x: number, projection: GridProjection) {
  let cursor = 0;
  for (const column of projection.columns) {
    const next = cursor + column.width;
    if (x <= next) {
      return { id: column.id, offset: Math.max(0, x - cursor) };
    }
    cursor = next;
  }
  const last = projection.columns[projection.columns.length - 1];
  return last ? { id: last.id, offset: last.width } : null;
}

function pointToRow(y: number, worksheet: Worksheet, projection: GridProjection) {
  let cursor = 0;
  for (const row of projection.rows) {
    const height = worksheet.rowsById[row.id]?.height ?? DEFAULT_ROW_HEIGHT;
    const next = cursor + height;
    if (y <= next) {
      return { id: row.id, offset: Math.max(0, y - cursor) };
    }
    cursor = next;
  }
  const last = projection.rows[projection.rows.length - 1];
  return last ? { id: last.id, offset: worksheet.rowsById[last.id]?.height ?? DEFAULT_ROW_HEIGHT } : null;
}

function sumColumnWidths(projection: GridProjection, start: number, end: number) {
  let total = 0;
  for (let index = start; index < end; index += 1) {
    total += projection.columns[index]?.width ?? 0;
  }
  return total;
}

function sumRowHeights(worksheet: Worksheet, projection: GridProjection, start: number, end: number) {
  let total = 0;
  for (let index = start; index < end; index += 1) {
    const row = projection.rows[index];
    total += row ? worksheet.rowsById[row.id]?.height ?? DEFAULT_ROW_HEIGHT : 0;
  }
  return total;
}
