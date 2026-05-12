/**
 * Public surface of the Teamgrid clipboard subsystem.
 *
 * The module is split into three concern-focused files:
 *
 * - {@link ./payload.ts} owns the serialization and deserialization of
 *   clipboard content (HTML with embedded Teamgrid JSON, TSV, and Excel
 *   `x:fmla` metadata).
 * - {@link ./excelFormula.ts} owns the A1 to relative R1C1 formula
 *   conversion required by the Excel HTML interop, in both directions.
 * - {@link ./refRewrite.ts} owns the post-paste rewriting of formula source
 *   strings to shift references by the anchor delta and to handle
 *   move-tracking for cut + paste.
 */
export {
  a1FormulaToRelativeR1C1,
  excelFormulaToA1,
  type ExcelFormulaToA1Options,
  type FormulaPosition,
} from "@/lib/clipboard/excelFormula";
export {
  decodePayload,
  decodeTsvPayload,
  serializeRange,
  type ClipboardAnchor,
  type ClipboardCell,
  type ClipboardMode,
  type ClipboardPayload,
  type ClipboardRange,
  type SerializedClipboardPayload,
} from "@/lib/clipboard/payload";
export {
  rewriteFormulaSource,
  type ReferenceDelta,
  type ReferenceRange,
  type RewriteOptions,
} from "@/lib/clipboard/refRewrite";
