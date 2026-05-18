import { describe, expect, it } from "vitest";

import { createTeamGridExcelWorkbook } from "@/features/xlsx/lib/exportWorkbook";
import { importTeamGridWorkbookBuffer } from "@/features/xlsx/lib/importWorkbook";
import { getCell, projectWorksheet } from "@/features/grid/lib/gridProjection";
import {
  createCellId,
  createTeamGridDocument,
  type CellStyle,
  type HorizontalAlign,
  type VerticalAlign,
} from "@/features/document/lib/teamgridDocument";

interface AlignmentCase {
  /** Excel-style A1 address inside the first worksheet of the fixture. */
  address: string;
  /** Style stamped onto the source cell. */
  style: CellStyle;
  /**
   * Alignment fields the round-trip is contractually required to
   * preserve. `verticalAlign` defaults to `"middle"` because the
   * exporter writes the merged style (row default + column default
   * + cell + DEFAULT_CELL_STYLE), which is what an Excel reader
   * needs to render the cell identically.
   */
  expected: {
    horizontalAlign?: HorizontalAlign;
    verticalAlign: VerticalAlign;
    wrapText: boolean;
    indent: number;
  };
}

/**
 * Exhaustive set of alignment combinations we want Excel interop to
 * preserve. Each row covers one orthogonal piece of the new Alignment
 * tab so a regression on any single field shows up as a focused test
 * failure instead of a single monolithic mismatch.
 */
const ALIGNMENT_CASES: AlignmentCase[] = [
  {
    address: "A1",
    style: { horizontalAlign: "left", verticalAlign: "top" },
    expected: { horizontalAlign: "left", verticalAlign: "top", wrapText: false, indent: 0 },
  },
  {
    address: "B1",
    style: { horizontalAlign: "center", verticalAlign: "middle" },
    expected: { horizontalAlign: "center", verticalAlign: "middle", wrapText: false, indent: 0 },
  },
  {
    address: "C1",
    style: { horizontalAlign: "right", verticalAlign: "bottom" },
    expected: { horizontalAlign: "right", verticalAlign: "bottom", wrapText: false, indent: 0 },
  },
  {
    // Wrap-only cell: vertical defaults to "middle" via the merged
    // export style so an Excel reader knows where to place the
    // wrapped text vertically.
    address: "D1",
    style: { wrapText: true },
    expected: { verticalAlign: "middle", wrapText: true, indent: 0 },
  },
  {
    address: "A2",
    style: { horizontalAlign: "left", indent: 3 },
    expected: { horizontalAlign: "left", verticalAlign: "middle", wrapText: false, indent: 3 },
  },
  {
    address: "B2",
    style: { horizontalAlign: "right", indent: 7, verticalAlign: "top", wrapText: true },
    expected: { horizontalAlign: "right", verticalAlign: "top", wrapText: true, indent: 7 },
  },
  {
    // "general" is Excel's implicit default; the exporter intentionally
    // omits `horizontal` from the xlsx so the file matches what
    // Excel itself produces for an untouched cell.
    address: "C2",
    style: { horizontalAlign: "general", verticalAlign: "middle" },
    expected: { verticalAlign: "middle", wrapText: false, indent: 0 },
  },
  {
    // Indent on Center is preserved even though it has no visual
    // effect — Excel stores the value the same way, so the
    // round-trip stays lossless.
    address: "D2",
    style: { horizontalAlign: "center", indent: 5 },
    expected: { horizontalAlign: "center", verticalAlign: "middle", wrapText: false, indent: 5 },
  },
  {
    address: "A3",
    style: { horizontalAlign: "left", verticalAlign: "bottom", wrapText: true, indent: 15 },
    expected: { horizontalAlign: "left", verticalAlign: "bottom", wrapText: true, indent: 15 },
  },
];

describe("XLSX alignment round-trip", () => {
  it("preserves horizontal, vertical, wrap text, and indent through Teamgrid \u2192 ExcelJS \u2192 Teamgrid", async () => {
    const envelope = createTeamGridDocument("Alignment round-trip");
    const document = envelope.teamgrid;
    const worksheet = document.workbook.worksheetsById[document.workbook.worksheetOrder[0]];
    const projection = projectWorksheet(worksheet);

    function rowColFromAddress(address: string) {
      const match = /^([A-Z]+)(\d+)$/.exec(address);
      if (!match) {
        throw new Error(`Bad fixture address: ${address}`);
      }
      const columnIndex = [...match[1]].reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - "A".charCodeAt(0) + 1), 0) - 1;
      const rowIndex = Number.parseInt(match[2], 10) - 1;
      return { rowId: projection.rows[rowIndex].id, columnId: projection.columns[columnIndex].id };
    }

    for (const fixture of ALIGNMENT_CASES) {
      const { rowId, columnId } = rowColFromAddress(fixture.address);
      worksheet.cellsById[createCellId(rowId, columnId)] = {
        id: createCellId(rowId, columnId),
        rowId,
        columnId,
        value: { kind: "string", text: `Cell ${fixture.address}` },
        style: { ...fixture.style },
      };
    }

    const workbook = createTeamGridExcelWorkbook(document);
    const buffer = await workbook.xlsx.writeBuffer();
    const imported = await importTeamGridWorkbookBuffer(buffer as ArrayBuffer, "Alignment round-trip");
    const importedWorksheet = imported.teamgrid.workbook.worksheetsById[imported.teamgrid.workbook.worksheetOrder[0]];
    const importedProjection = projectWorksheet(importedWorksheet);

    for (const fixture of ALIGNMENT_CASES) {
      const { rowId, columnId } = rowColFromAddress(fixture.address);
      // The imported worksheet preserves the row/column order so the
      // same A1 address maps to the same (rowIndex, columnIndex) pair.
      const importedRow = importedProjection.rows[projection.rowIndexById.get(rowId)!];
      const importedColumn = importedProjection.columns[projection.columnIndexById.get(columnId)!];
      const cell = getCell(importedWorksheet, importedRow.id, importedColumn.id);
      const style = cell.style ?? {};

      expect(style.horizontalAlign, `${fixture.address} horizontalAlign`).toBe(fixture.expected.horizontalAlign);
      expect(style.verticalAlign, `${fixture.address} verticalAlign`).toBe(fixture.expected.verticalAlign);
      expect(Boolean(style.wrapText), `${fixture.address} wrapText`).toBe(fixture.expected.wrapText);
      expect(style.indent ?? 0, `${fixture.address} indent`).toBe(fixture.expected.indent);
    }
  });
});
