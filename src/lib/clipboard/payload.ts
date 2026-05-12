/**
 * Clipboard payload encoding for Teamgrid ranges.
 *
 * A copy/cut operation produces three parallel representations so that the
 * widest possible set of paste targets can consume the data:
 *
 * 1. `text/plain` TSV - the lowest common denominator for spreadsheets and
 *    plain editors that ignore HTML.
 * 2. `text/html` table with Excel-compatible attributes (`x:fmla`, `x:num`,
 *    `x:str`) so Excel and other Office clients keep relative formulas.
 * 3. A Teamgrid-specific JSON payload embedded as the
 *    `data-teamgrid-payload` attribute on the root `<table>`. This carries
 *    stable cell IDs, styles, formula source text, and cut metadata so a
 *    Teamgrid-to-Teamgrid paste preserves everything losslessly.
 *
 * On decode the precedence is reversed: Teamgrid JSON wins, Excel HTML is
 * the second choice, and TSV is the fallback.
 */
import { a1FormulaToRelativeR1C1, excelFormulaToA1 } from "@/lib/clipboard/excelFormula";
import type { Cell, CellStyle, CellValue, WorksheetId } from "@/lib/teamgridDocument";

/** Whether a clipboard payload was produced by a copy or by a cut. */
export type ClipboardMode = "copy" | "cut";

/** Zero-based anchor coordinate inside the source worksheet. */
export interface ClipboardAnchor {
  row: number;
  col: number;
}

/** Inclusive rectangular range of cells in the source worksheet. */
export interface ClipboardRange {
  worksheetId: WorksheetId;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

/**
 * One cell inside a {@link ClipboardPayload}.
 *
 * Coordinates are stored as offsets from the payload's anchor (top-left) so
 * the same payload can be applied at any target anchor without re-coordinating
 * the entries.
 */
export interface ClipboardCell {
  rowOffset: number;
  colOffset: number;
  value: CellValue;
  formulaSource?: string;
  style?: CellStyle;
  text: string;
}

/**
 * Internal Teamgrid clipboard payload.
 *
 * Persisted into the HTML clipboard as base64-encoded JSON inside
 * `data-teamgrid-payload` so it survives a round trip through the OS
 * clipboard. Outside callers should treat this as opaque metadata: it is the
 * paste pipeline in {@link ../../App.vue} that knows how to interpret it.
 */
export interface ClipboardPayload {
  version: 1;
  origin: "teamgrid";
  mode: ClipboardMode;
  source: {
    worksheetId: WorksheetId | null;
    rows: number;
    cols: number;
    anchor: ClipboardAnchor;
  };
  cells: ClipboardCell[];
  cutCellIds?: string[];
}

/** Result of {@link serializeRange}: an HTML body, a TSV body, and the raw payload. */
export interface SerializedClipboardPayload {
  html: string;
  tsv: string;
  payload: ClipboardPayload;
}

/**
 * Anchor used when we synthesize a payload from foreign Excel HTML.
 *
 * Excel's R1C1 references in `x:fmla` can point outside the copied range
 * (for example `=R[-2]C[-1]`). Resolving them at coordinates `(0, 0)` would
 * produce negative cell addresses, which `excelFormulaToA1` would have to
 * drop. We therefore pretend the imported table sits at a large offset so
 * every R1C1 reference resolves to a non-negative A1 address. The paste path
 * then applies the existing `rewriteFormulaSource(...)` to shift everything
 * back to the user's actual target anchor.
 */
const EXCEL_HTML_VIRTUAL_ORIGIN = { row: 10000, col: 10000 };

/**
 * Build a clipboard payload for the supplied rectangular range.
 *
 * `getCellAt(row, col)` is supplied by the caller (see
 * {@link ../../App.vue}) and resolves grid coordinates to {@link Cell} records.
 * The returned `html` carries Excel-compatible formula metadata plus the
 * embedded Teamgrid payload, `tsv` is the plain-text fallback, and `payload`
 * is the raw structure for internal fallbacks when `navigator.clipboard` is
 * unavailable.
 */
export function serializeRange(
  range: ClipboardRange,
  getCellAt: (row: number, col: number) => Cell,
  mode: ClipboardMode,
): SerializedClipboardPayload {
  const bounds = normalizeRange(range);
  const rows = bounds.endRow - bounds.startRow + 1;
  const cols = bounds.endCol - bounds.startCol + 1;
  const cells: ClipboardCell[] = [];
  const tsvRows: string[][] = [];
  const htmlRows: string[] = [];
  const cutCellIds: string[] = [];

  for (let rowIndex = bounds.startRow; rowIndex <= bounds.endRow; rowIndex += 1) {
    const tsvRow: string[] = [];
    const htmlCells: string[] = [];
    for (let colIndex = bounds.startCol; colIndex <= bounds.endCol; colIndex += 1) {
      const cell = getCellAt(rowIndex, colIndex);
      const text = cellToClipboardText(cell);
      const clipboardCell: ClipboardCell = {
        rowOffset: rowIndex - bounds.startRow,
        colOffset: colIndex - bounds.startCol,
        value: cell.value,
        text,
      };
      if (cell.formula?.source) {
        clipboardCell.formulaSource = cell.formula.source;
      }
      if (cell.style) {
        clipboardCell.style = cell.style;
      }
      cells.push(clipboardCell);
      tsvRow.push(text);
      htmlCells.push(createHtmlCell(cell, text, { row: rowIndex, col: colIndex }));
      cutCellIds.push(cell.id);
    }
    tsvRows.push(tsvRow);
    htmlRows.push(`<tr>${htmlCells.join("")}</tr>`);
  }

  const payload: ClipboardPayload = {
    version: 1,
    origin: "teamgrid",
    mode,
    source: {
      worksheetId: range.worksheetId,
      rows,
      cols,
      anchor: { row: bounds.startRow, col: bounds.startCol },
    },
    cells,
    cutCellIds: mode === "cut" ? cutCellIds : undefined,
  };
  const encodedPayload = encodePayload(payload);
  return {
    payload,
    tsv: tsvRows.map((row) => row.map(escapeTsvCell).join("\t")).join("\n"),
    html: `<table xmlns:x="urn:schemas-microsoft-com:office:excel" data-teamgrid-payload="${encodedPayload}"><tbody>${htmlRows.join("")}</tbody></table>`,
  };
}

/**
 * Decode a clipboard event's HTML and TSV payloads into a Teamgrid payload.
 *
 * Resolution order:
 *
 * 1. Teamgrid's embedded JSON payload (`data-teamgrid-payload`). Wins
 *    whenever present, including in Teamgrid-to-Teamgrid paste.
 * 2. Excel HTML formula metadata (`x:fmla` and friends). Used when a foreign
 *    app such as Excel pasted the clipboard.
 * 3. Plain TSV. Final fallback for clipboards that contain only text.
 *
 * Returns `null` when none of the inputs carry usable data.
 */
export function decodePayload(html: string, tsv: string): ClipboardPayload | null {
  const encoded = /data-teamgrid-payload=["']([^"']+)["']/.exec(html)?.[1];
  if (encoded) {
    const payload = decodeEmbeddedPayload(encoded);
    if (payload) {
      return payload;
    }
  }
  const excelPayload = decodeExcelHtmlPayload(html);
  if (excelPayload) {
    return excelPayload;
  }
  return decodeTsvPayload(tsv);
}

/**
 * Build a value-only payload from TSV text.
 *
 * Numeric-looking cells become `number` values; everything else becomes a
 * string. There is no formula or style information in TSV, so the paste
 * pipeline will commit plain values.
 */
export function decodeTsvPayload(tsv: string): ClipboardPayload | null {
  if (!tsv.trim()) {
    return null;
  }
  const rows = tsv.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  const cols = Math.max(...rows.map((row) => row.split("\t").length));
  const cells: ClipboardCell[] = [];
  rows.forEach((row, rowOffset) => {
    row.split("\t").forEach((text, colOffset) => {
      cells.push({
        rowOffset,
        colOffset,
        text,
        value: parseClipboardText(text),
      });
    });
  });
  return {
    version: 1,
    origin: "teamgrid",
    mode: "copy",
    source: {
      worksheetId: null,
      rows: rows.length,
      cols,
      anchor: { row: 0, col: 0 },
    },
    cells,
  };
}

/** Sort the four range corners so `start* <= end*`. */
function normalizeRange(range: ClipboardRange) {
  return {
    startRow: Math.min(range.startRow, range.endRow),
    endRow: Math.max(range.startRow, range.endRow),
    startCol: Math.min(range.startCol, range.endCol),
    endCol: Math.max(range.startCol, range.endCol),
  };
}

/**
 * Stringify a cell for the TSV body and the Teamgrid embedded payload.
 *
 * Formula cells contribute their `source` so a Teamgrid-aware target sees the
 * original A1 expression; plain-value cells contribute their formatted text.
 */
function cellToClipboardText(cell: Cell) {
  if (cell.formula?.source) {
    return cell.formula.source;
  }
  switch (cell.value.kind) {
    case "number":
      return String(cell.value.value);
    case "string":
      return cell.value.text;
    case "date":
      return cell.value.isoDate;
    default:
      return "";
  }
}

/**
 * Stringify only the cached value of a cell.
 *
 * Used inside the visible HTML body so foreign spreadsheets show the cached
 * formula result while Excel-aware ones recompute from the `x:fmla` metadata.
 */
function cellToDisplayText(cell: Cell) {
  switch (cell.value.kind) {
    case "number":
      return String(cell.value.value);
    case "string":
      return cell.value.text;
    case "date":
      return cell.value.isoDate;
    default:
      return "";
  }
}

/**
 * Render one source cell as an HTML `<td>` with Excel-compatible metadata.
 *
 * Adds:
 * - `x:fmla` carrying the formula converted to relative R1C1 so Excel will
 *   shift references on paste.
 * - `x:num` for numeric cells so Excel keeps the precise value.
 * - `x:str` for explicit string cells.
 *
 * The text content of the `<td>` is the cached/displayed value of the cell,
 * which is what spreadsheet apps fall back to when they do not understand
 * the Excel metadata.
 */
function createHtmlCell(cell: Cell, fallbackText: string, position: { row: number; col: number }) {
  const attributes: string[] = [];
  const displayText = cell.formula?.source ? cellToDisplayText(cell) : fallbackText;
  if (cell.formula?.source) {
    attributes.push(`x:fmla="${escapeAttribute(a1FormulaToRelativeR1C1(cell.formula.source, position))}"`);
  }
  if (cell.value.kind === "number") {
    attributes.push(`x:num="${escapeAttribute(String(cell.value.value))}"`);
  } else if (cell.value.kind === "string") {
    attributes.push("x:str");
  }
  return `<td${attributes.length > 0 ? ` ${attributes.join(" ")}` : ""}>${escapeHtml(displayText)}</td>`;
}

/**
 * Convert a foreign Excel HTML clipboard string into a {@link ClipboardPayload}.
 *
 * Each `<td>`'s text becomes the cell value and the Excel `x:fmla` attribute
 * (or any namespaced equivalent) becomes the Teamgrid `formulaSource`. The
 * payload is anchored at {@link EXCEL_HTML_VIRTUAL_ORIGIN} so R1C1 references
 * never resolve to negative addresses; the paste pipeline shifts them to the
 * user's real target anchor afterward.
 *
 * Returns `null` when the HTML has no `<tr>` elements, or when running outside
 * a browser (no `DOMParser`).
 */
function decodeExcelHtmlPayload(html: string): ClipboardPayload | null {
  if (!html.trim() || typeof DOMParser === "undefined") {
    return null;
  }
  const document = new DOMParser().parseFromString(html, "text/html");
  const rows = Array.from(document.querySelectorAll("tr"));
  if (rows.length === 0) {
    return null;
  }

  const cells: ClipboardCell[] = [];
  let maxCols = 0;
  rows.forEach((row, rowOffset) => {
    const tableCells = Array.from(row.querySelectorAll("td, th"));
    maxCols = Math.max(maxCols, tableCells.length);
    tableCells.forEach((tableCell, colOffset) => {
      const text = tableCell.textContent ?? "";
      const formula = getExcelFormulaAttribute(tableCell);
      const clipboardCell: ClipboardCell = {
        rowOffset,
        colOffset,
        text,
        value: parseClipboardText(text),
      };
      if (formula) {
        clipboardCell.formulaSource = excelFormulaToA1(
          formula,
          {
            row: EXCEL_HTML_VIRTUAL_ORIGIN.row + rowOffset,
            col: EXCEL_HTML_VIRTUAL_ORIGIN.col + colOffset,
          },
          { origin: EXCEL_HTML_VIRTUAL_ORIGIN },
        );
      }
      cells.push(clipboardCell);
    });
  });

  if (cells.length === 0) {
    return null;
  }
  return {
    version: 1,
    origin: "teamgrid",
    mode: "copy",
    source: {
      worksheetId: null,
      rows: rows.length,
      cols: maxCols,
      anchor: EXCEL_HTML_VIRTUAL_ORIGIN,
    },
    cells,
  };
}

/**
 * Read the formula attribute from an HTML table cell.
 *
 * Different Excel versions and the `DOMParser` text/html mode can serialize
 * the namespaced attribute as `x:fmla`, `fmla` (after namespace stripping),
 * or any `prefix:fmla` shape. We accept all of those so the importer is
 * resilient.
 */
function getExcelFormulaAttribute(element: Element) {
  for (const attribute of Array.from(element.attributes)) {
    const name = attribute.name.toLowerCase();
    if (name === "fmla" || name === "x:fmla" || name.endsWith(":fmla")) {
      return attribute.value;
    }
  }
  return "";
}

/** Heuristically convert TSV cell text into a {@link CellValue}. */
function parseClipboardText(text: string): CellValue {
  if (text === "") {
    return { kind: "empty" };
  }
  const number = Number(text);
  if (Number.isFinite(number) && text.trim() !== "") {
    return { kind: "number", value: number };
  }
  return { kind: "string", text };
}

/**
 * Encode a Teamgrid payload as base64 JSON.
 *
 * `btoa` only accepts Latin-1 strings so we route the JSON through
 * `TextEncoder` first to safely cover the full Unicode range (workbook
 * subjects, tag labels, and cell values are all free-form text).
 */
function encodePayload(payload: ClipboardPayload) {
  const bytes = new TextEncoder().encode(JSON.stringify(payload));
  let binary = "";
  bytes.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary);
}

/**
 * Inverse of {@link encodePayload}. Validates the version and origin so we
 * never accept a foreign string that happened to look like base64 JSON.
 */
function decodeEmbeddedPayload(encoded: string): ClipboardPayload | null {
  try {
    const binary = atob(encoded);
    const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
    const payload = JSON.parse(new TextDecoder().decode(bytes)) as ClipboardPayload;
    return payload?.version === 1 && payload.origin === "teamgrid" ? payload : null;
  } catch {
    return null;
  }
}

/** HTML-escape user-controlled cell text and attribute fragments. */
function escapeHtml(value: string) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

/**
 * Attribute-escape including single quotes.
 *
 * `escapeHtml` keeps single quotes intact because they are safe inside double
 * quoted text content, but our attributes use double quotes too. Encoding `'`
 * as `&#39;` guards against any tooling that might re-emit attributes with
 * single quotes.
 */
function escapeAttribute(value: string) {
  return escapeHtml(value).replaceAll("'", "&#39;");
}

/**
 * Make a cell safe for TSV: tabs and newlines would corrupt the row/column
 * structure, so we replace them with spaces. Excel applies the same heuristic
 * when generating its plain-text clipboard.
 */
function escapeTsvCell(value: string) {
  return value.replaceAll("\t", " ").replaceAll(/\r?\n/g, " ");
}
