/**
 * Grid clipboard pipeline: serialize selections to OS-clipboard payloads,
 * decode pastes coming from native events / the menu / the in-memory
 * fallback, and apply pastes (including cut-style move tracking) as a
 * single granular `updateGrid` mutation.
 *
 * The composable hosts three parallel encodings on copy/cut:
 *
 * - A TeamGrid-native JSON payload embedded in `text/html` for lossless
 *   TeamGrid-to-TeamGrid paste.
 * - An Excel-compatible HTML body with `x:fmla`/`x:num`/`x:str` attributes.
 * - A plain TSV fallback in `text/plain` for everything else.
 *
 * Paste decodes in the same priority order. When the host blocks
 * `navigator.clipboard` (common in sandboxed Haven iframes) the in-memory
 * {@link UseGridClipboardOptions} clipboard keeps Edit menu copy + paste
 * useful inside the app.
 */
import { ref, type Ref } from "vue";
import { DEFAULT_COLUMN_WIDTH } from "@/shared/lib/gridDimensions";
import {
  decodePayload,
  rewriteFormulaSource,
  serializeRange,
  type ClipboardPayload,
  type ClipboardRange,
  type SerializedClipboardPayload,
} from "@/features/clipboard/lib";
import { createFormulaContext, evaluateFormula, renderFormulaSource, type FormulaContext } from "@/features/formulas/lib";
import { formulaResultToCellValue } from "@/features/grid/lib/cellFormatting";
import {
  createCellId,
  createId,
  type Cell,
  type Worksheet,
} from "@/features/document/lib/teamgridDocument";
import type { TeamGridOperation } from "@/features/document/lib/teamgridOps";
import { getCell, projectWorksheet, type GridProjection } from "@/features/grid/lib/gridProjection";
import {
  findCellCoordinatesInProjection,
  type CellSelectionRange,
} from "@/features/grid/composables/useSelection";
import type { TeamGridAppApi } from "@/features/document/composables/useTeamGridDocument";

export interface UseGridClipboardOptions {
  app: TeamGridAppApi;
  activeWorksheet: Readonly<Ref<Worksheet | null>>;
  projection: Readonly<Ref<GridProjection | null>>;
  selectedCell: Readonly<Ref<Cell | null>>;
  selectedRange: Ref<CellSelectionRange | null>;
  findCellCoordinates: (cellId: string) => { rowIndex: number; columnIndex: number } | null;
}

export function useGridClipboard(options: UseGridClipboardOptions) {
  const { app, activeWorksheet, projection, selectedCell, selectedRange, findCellCoordinates } = options;
  const clipboardSourceRange = ref<ClipboardRange | null>(null);
  const internalClipboard = ref<SerializedClipboardPayload | null>(null);

  /** Grid-emitted `copy` event: serialize the selection into the clipboard. */
  function handleGridClipboardCopy(payload: { range: CellSelectionRange | null; event: ClipboardEvent }) {
    writeSelectionToClipboard(payload.range, payload.event, "copy");
  }

  /**
   * Grid-emitted `cut` event: same as copy, but mark the cells as "cut" so
   * the paste handler knows to clear the source range and rewrite incoming
   * formulas using move-tracking deltas.
   */
  function handleGridClipboardCut(payload: { range: CellSelectionRange | null; event: ClipboardEvent }) {
    writeSelectionToClipboard(payload.range, payload.event, "cut");
  }

  /**
   * Grid-emitted `paste` event.
   *
   * Tries the OS clipboard first (so Excel-originated HTML wins over an
   * older internal copy), then falls back to the in-memory clipboard for
   * sandboxed hosts that block `navigator.clipboard`.
   */
  function handleGridClipboardPaste(payload: { event: ClipboardEvent }) {
    const clipboardPayload = readClipboardPayload(payload.event) ?? internalClipboard.value?.payload ?? null;
    const anchor = selectedCell.value ? findCellCoordinates(selectedCell.value.id) : null;
    if (!clipboardPayload || !anchor) {
      return;
    }
    applyPasteAtAnchor(clipboardPayload, { row: anchor.rowIndex, col: anchor.columnIndex });
  }

  /** Stop drawing the "marching ants" marquee, e.g. on Escape or paste. */
  function clearClipboardMarquee() {
    clipboardSourceRange.value = null;
  }

  /**
   * Serialize the current selection into the native `ClipboardEvent` plus
   * the in-memory fallback. We always write both `text/html` (rich TeamGrid
   * + Excel-compatible payload) and `text/plain` (TSV) so paste targets
   * across the ecosystem get the richest data they can consume.
   */
  function writeSelectionToClipboard(range: CellSelectionRange | null, event: ClipboardEvent, mode: "copy" | "cut") {
    const serialized = serializeSelection(range, mode);
    if (!serialized) {
      return;
    }
    event.clipboardData?.setData("text/html", serialized.html);
    event.clipboardData?.setData("text/plain", serialized.tsv);
    internalClipboard.value = serialized;
    clipboardSourceRange.value = serialized.payload.source.worksheetId ? {
      worksheetId: serialized.payload.source.worksheetId,
      startRow: serialized.payload.source.anchor.row,
      startCol: serialized.payload.source.anchor.col,
      endRow: serialized.payload.source.anchor.row + serialized.payload.source.rows - 1,
      endCol: serialized.payload.source.anchor.col + serialized.payload.source.cols - 1,
    } : null;
  }

  /** Decode a paste-event clipboard, preferring TeamGrid JSON over Excel HTML over TSV. */
  function readClipboardPayload(event: ClipboardEvent) {
    const html = event.clipboardData?.getData("text/html") ?? "";
    const tsv = event.clipboardData?.getData("text/plain") ?? "";
    return decodePayload(html, tsv);
  }

  /**
   * Convert the given (or active) selection into the serialized clipboard
   * payload used by both the native event and the in-memory fallback.
   * Returns `null` when there is no worksheet or no usable range.
   */
  function serializeSelection(range: CellSelectionRange | null, mode: "copy" | "cut") {
    const clipboardRange = selectionToClipboardRange(range);
    if (!clipboardRange || !activeWorksheet.value || !projection.value || !app.activeGrid.value) {
      return null;
    }
    const formulaContext = createFormulaContext(app.activeGrid.value.workbook);
    return serializeRange(
      clipboardRange,
      (rowIndex, columnIndex) => {
        const row = projection.value!.rows[rowIndex];
        const column = projection.value!.columns[columnIndex];
        return getCell(activeWorksheet.value!, row.id, column.id);
      },
      mode,
      {
        formulaSource: (cell) => cell.formula
          ? renderFormulaSource(cell.formula, activeWorksheet.value!.id, formulaContext)
          : "",
      },
    );
  }

  /**
   * Edit menu entry point for Copy / Cut.
   *
   * The clipboard `copy`/`cut` events only fire for keyboard shortcuts and
   * the browser's own menu, so this path uses the async
   * `navigator.clipboard` write API. Sandboxed Haven hosts often reject
   * this API, so we always populate the in-memory clipboard as a fallback
   * before attempting the async write.
   */
  async function copySelectionFromMenu(mode: "copy" | "cut", range = selectedRange.value) {
    const serialized = serializeSelection(range, mode);
    if (!serialized) {
      return;
    }
    internalClipboard.value = serialized;
    clipboardSourceRange.value = payloadSourceRange(serialized.payload);
    try {
      if (typeof ClipboardItem !== "undefined" && navigator.clipboard.write) {
        await navigator.clipboard.write([new ClipboardItem({
          "text/html": new Blob([serialized.html], { type: "text/html" }),
          "text/plain": new Blob([serialized.tsv], { type: "text/plain" }),
        })]);
        return;
      }
      await navigator.clipboard.writeText(serialized.tsv);
    } catch {
      // Sandboxed hosts often block navigator.clipboard; internalClipboard keeps the menu useful.
    }
  }

  /**
   * Edit menu entry point for Paste.
   *
   * Reads the OS clipboard via the async API (richer than the synchronous
   * `ClipboardEvent` path because we can pull both `text/html` and
   * `text/plain`), falling back to the in-memory clipboard when the host
   * blocks the API.
   */
  async function pasteFromMenu() {
    if (app.gridReadOnly.value) {
      return;
    }
    const anchor = selectedCell.value ? findCellCoordinates(selectedCell.value.id) : null;
    if (!anchor) {
      return;
    }
    const payload = await readNavigatorClipboardPayload() ?? internalClipboard.value?.payload ?? null;
    if (!payload) {
      return;
    }
    applyPasteAtAnchor(payload, { row: anchor.rowIndex, col: anchor.columnIndex });
  }

  /**
   * Read the OS clipboard via `navigator.clipboard.read()` (preferred
   * because it exposes HTML data) and decode the first item that yields a
   * usable TeamGrid/Excel/TSV payload. Returns `null` on permission errors
   * or empty clipboards.
   */
  async function readNavigatorClipboardPayload() {
    try {
      if (navigator.clipboard.read) {
        const items = await navigator.clipboard.read();
        for (const item of items) {
          const html = item.types.includes("text/html")
            ? await (await item.getType("text/html")).text()
            : "";
          const text = item.types.includes("text/plain")
            ? await (await item.getType("text/plain")).text()
            : "";
          const payload = decodePayload(html, text);
          if (payload) {
            return payload;
          }
        }
      }
      const text = await navigator.clipboard.readText();
      return decodePayload("", text);
    } catch {
      return null;
    }
  }

  /**
   * Project a `CellSelectionRange` (selection by stable cell ids) onto a
   * worksheet-relative {@link ClipboardRange} (selection by row/column
   * index). Returns `null` when there is nothing to copy, e.g. before the
   * worksheet is loaded.
   */
  function selectionToClipboardRange(range: CellSelectionRange | null): ClipboardRange | null {
    if (!activeWorksheet.value || !projection.value) {
      return null;
    }
    const activeRange = range ?? (selectedCell.value ? { startCellId: selectedCell.value.id, endCellId: selectedCell.value.id } : null);
    if (!activeRange) {
      return null;
    }
    const start = findCellCoordinates(activeRange.startCellId);
    const end = findCellCoordinates(activeRange.endCellId);
    if (!start || !end) {
      return null;
    }
    return {
      worksheetId: activeWorksheet.value.id,
      startRow: Math.min(start.rowIndex, end.rowIndex),
      startCol: Math.min(start.columnIndex, end.columnIndex),
      endRow: Math.max(start.rowIndex, end.rowIndex),
      endCol: Math.max(start.columnIndex, end.columnIndex),
    };
  }

  /**
   * Apply a decoded clipboard payload starting at the given anchor cell.
   *
   * Performs the heavy lifting of the paste pipeline:
   * 1. Auto-extends the worksheet so the entire payload fits.
   * 2. Writes each clipboard cell into its target position, rewriting any
   *    formula references by the copy delta so relative references shift
   *    Excel-style.
   * 3. For cut payloads, clears the source cells and rewrites formulas
   *    elsewhere in the sheet that pointed at the cut range (move tracking).
   * 4. Bundles every change into one `updateGrid` mutation so the granular
   *    save / `baseHeads` machinery treats the paste as a single edit.
   *
   * A cut that resolves back to its own source range is a no-op apart from
   * clearing the marquee.
   */
  function applyPasteAtAnchor(payload: ClipboardPayload, anchor: { row: number; col: number }) {
    if (!activeWorksheet.value) {
      return;
    }
    const activeWorksheetBeforePaste = activeWorksheet.value;
    const sourceRange = payloadSourceRange(payload);
    const destinationRange = {
      worksheetId: activeWorksheetBeforePaste.id,
      startRow: anchor.row,
      startCol: anchor.col,
      endRow: anchor.row + payload.source.rows - 1,
      endCol: anchor.col + payload.source.cols - 1,
    };
    if (payload.mode === "cut" && sourceRange && rangesEqual(sourceRange, destinationRange)) {
      clipboardSourceRange.value = null;
      return;
    }

    app.updateGrid((grid) => {
      const worksheet = grid.workbook.worksheetsById[activeWorksheetBeforePaste.id];
      const operations: TeamGridOperation[] = [];
      ensureGridSize(worksheet, destinationRange.endRow + 1, destinationRange.endCol + 1, operations);
      const nextProjection = projectWorksheet(worksheet);
      const destinationIds = new Set<string>();
      const formulaDelta = {
        rows: payload.mode === "copy" ? anchor.row - payload.source.anchor.row : 0,
        cols: payload.mode === "copy" ? anchor.col - payload.source.anchor.col : 0,
      };

      for (const clipboardCell of payload.cells) {
        const targetRow = nextProjection.rows[anchor.row + clipboardCell.rowOffset];
        const targetColumn = nextProjection.columns[anchor.col + clipboardCell.colOffset];
        if (!targetRow || !targetColumn) {
          continue;
        }
        const targetCellId = createCellId(targetRow.id, targetColumn.id);
        destinationIds.add(targetCellId);
        const formulaSource = clipboardCell.formulaSource
          ? rewriteFormulaSource(clipboardCell.formulaSource, formulaDelta)
          : undefined;
        const nextCell: Cell = {
          id: targetCellId,
          rowId: targetRow.id,
          columnId: targetColumn.id,
          value: cloneCellValue(clipboardCell.value),
          style: clipboardCell.style ? { ...clipboardCell.style } : undefined,
          formula: undefined,
        };
        if (formulaSource) {
          applyFormulaToCell(nextCell, formulaSource, worksheet, createFormulaContext(grid.workbook));
        }
        worksheet.cellsById[targetCellId] = nextCell;
        operations.push({ type: "setCell", worksheetId: worksheet.id, cell: nextCell });
      }

      if (payload.mode === "cut" && payload.source.worksheetId === worksheet.id && payload.cutCellIds) {
        for (const cellId of payload.cutCellIds) {
          if (destinationIds.has(cellId)) {
            continue;
          }
          const existing = worksheet.cellsById[cellId];
          if (!existing) {
            continue;
          }
          const emptyCell: Cell = { ...existing, value: { kind: "empty" }, formula: undefined };
          worksheet.cellsById[cellId] = emptyCell;
          operations.push({ type: "setCell", worksheetId: worksheet.id, cell: emptyCell });
        }
      }

      if (payload.mode === "cut" && sourceRange && payload.source.worksheetId === worksheet.id) {
        operations.push(...applyMoveTracking(worksheet, nextProjection, sourceRange, destinationRange, createFormulaContext(grid.workbook)));
      }

      return operations;
    });

    selectedRange.value = {
      startCellId: cellIdAt(destinationRange.startRow, destinationRange.startCol),
      endCellId: cellIdAt(destinationRange.endRow, destinationRange.endCol),
    };
    clipboardSourceRange.value = null;
  }

  /**
   * Append rows/columns to the worksheet until the projection reaches at
   * least `minRows` rows and `minCols` columns. Mutations are appended to
   * the caller's `operations` array so they participate in the granular
   * patch generated by the enclosing `updateGrid` call.
   */
  function ensureGridSize(worksheet: Worksheet, minRows: number, minCols: number, operations: TeamGridOperation[]) {
    while (projectWorksheet(worksheet).rows.length < minRows) {
      const rowId = createId("row");
      const row = { id: rowId };
      worksheet.rowsById[rowId] = row;
      worksheet.rowOrder.push(rowId);
      operations.push({ type: "insertRow", worksheetId: worksheet.id, rowId, row, index: worksheet.rowOrder.length - 1 });
    }
    while (projectWorksheet(worksheet).columns.length < minCols) {
      const columnId = createId("col");
      const column = { id: columnId, width: DEFAULT_COLUMN_WIDTH };
      worksheet.columnsById[columnId] = column;
      worksheet.columnOrder.push(columnId);
      operations.push({ type: "insertColumn", worksheetId: worksheet.id, columnId, column, index: worksheet.columnOrder.length - 1 });
    }
  }

  /**
   * Parse, evaluate, and cache a formula on the given cell.
   *
   * Mutates `cell` in place so callers can build a fresh cell record and
   * pipe it through a single setCell operation. The cached `references`
   * allow the dependency tracker to skip re-parsing on subsequent reads.
   */
  function applyFormulaToCell(cell: Cell, formulaSource: string, worksheet: Worksheet, formulaContext: FormulaContext) {
    const evaluated = evaluateFormula(formulaSource, worksheet.id, formulaContext);
    const renderedSource = evaluated.segments
      ? renderFormulaSource({ source: formulaSource, segments: evaluated.segments }, worksheet.id, formulaContext)
      : formulaSource;
    cell.formula = {
      kind: "formula",
      source: renderedSource,
      segments: evaluated.segments,
      references: evaluated.references,
      cached: evaluated.result,
      error: evaluated.result.kind === "error" ? evaluated.result.code : undefined,
    };
    cell.value = formulaResultToCellValue(evaluated.result);
  }

  /**
   * Rewrite formulas across the worksheet that referenced the cut range so
   * they follow the moved cells (Excel-style "move tracking").
   */
  function applyMoveTracking(
    worksheet: Worksheet,
    worksheetProjection: GridProjection,
    sourceRange: ClipboardRange,
    destinationRange: ClipboardRange,
    formulaContext: FormulaContext,
  ) {
    const operations: TeamGridOperation[] = [];
    const delta = {
      rows: destinationRange.startRow - sourceRange.startRow,
      cols: destinationRange.startCol - sourceRange.startCol,
    };
    for (const cell of Object.values(worksheet.cellsById)) {
      const coordinates = findCellCoordinatesInProjection(cell.id, worksheet, worksheetProjection);
      if (!coordinates || rangeContains(sourceRange, coordinates.rowIndex, coordinates.columnIndex) || rangeContains(destinationRange, coordinates.rowIndex, coordinates.columnIndex) || !cell.formula?.source) {
        continue;
      }
      const renderedSource = renderFormulaSource(cell.formula, worksheet.id, formulaContext);
      const nextSource = rewriteFormulaSource(renderedSource, delta, {
        insideRange: {
          startRow: sourceRange.startRow,
          startCol: sourceRange.startCol,
          endRow: sourceRange.endRow,
          endCol: sourceRange.endCol,
        },
      });
      if (nextSource === renderedSource) {
        continue;
      }
      const nextCell: Cell = { ...cell, formula: undefined };
      applyFormulaToCell(nextCell, nextSource, worksheet, formulaContext);
      worksheet.cellsById[nextCell.id] = nextCell;
      operations.push({ type: "setCell", worksheetId: worksheet.id, cell: nextCell });
    }
    return operations;
  }

  function cellIdAt(rowIndex: number, columnIndex: number) {
    if (!activeWorksheet.value || !projection.value) {
      return "";
    }
    const row = projection.value.rows[rowIndex];
    const column = projection.value.columns[columnIndex];
    return row && column ? createCellId(row.id, column.id) : "";
  }

  return {
    clipboardSourceRange,
    internalClipboard,
    handleGridClipboardCopy,
    handleGridClipboardCut,
    handleGridClipboardPaste,
    clearClipboardMarquee,
    copySelectionFromMenu,
    pasteFromMenu,
  };
}

/**
 * Deep-clone a cell value. The clipboard payload is shared across paste
 * destinations, so cloning prevents accidental aliasing when the same
 * structured value (e.g. `{ kind: "number", value: 42 }`) ends up in many
 * target cells.
 */
function cloneCellValue(value: Cell["value"]): Cell["value"] {
  return JSON.parse(JSON.stringify(value)) as Cell["value"];
}

/**
 * Reconstruct the source range of a clipboard payload in worksheet-relative
 * coordinates. Returns `null` for payloads that did not come from this app
 * (no `worksheetId`), which avoids accidentally treating an Excel paste as
 * a cut-from-self.
 */
function payloadSourceRange(payload: ClipboardPayload): ClipboardRange | null {
  if (!payload.source.worksheetId) {
    return null;
  }
  return {
    worksheetId: payload.source.worksheetId,
    startRow: payload.source.anchor.row,
    startCol: payload.source.anchor.col,
    endRow: payload.source.anchor.row + payload.source.rows - 1,
    endCol: payload.source.anchor.col + payload.source.cols - 1,
  };
}

/** Structural equality for `ClipboardRange`. */
function rangesEqual(left: ClipboardRange, right: ClipboardRange) {
  return left.worksheetId === right.worksheetId
    && left.startRow === right.startRow
    && left.startCol === right.startCol
    && left.endRow === right.endRow
    && left.endCol === right.endCol;
}

/** Test whether `(rowIndex, columnIndex)` falls inside the inclusive `range`. */
function rangeContains(range: ClipboardRange, rowIndex: number, columnIndex: number) {
  return rowIndex >= range.startRow
    && rowIndex <= range.endRow
    && columnIndex >= range.startCol
    && columnIndex <= range.endCol;
}
