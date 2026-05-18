import { describe, expect, it } from "vitest";
import { columnIndexToLabel, getCellAddress, parseCellAddress, projectWorksheet } from "@/features/grid/lib/gridProjection";
import { createTeamGridDocument, getFirstVisibleWorksheet } from "@/features/document/lib/teamgridDocument";

describe("grid projection", () => {
  it("projects stable row and column ids into spreadsheet addresses", () => {
    const envelope = createTeamGridDocument();
    const worksheet = getFirstVisibleWorksheet(envelope.teamgrid)!;
    const projection = projectWorksheet(worksheet);

    expect(projection.rows[0].label).toBe("1");
    expect(projection.columns[0].label).toBe("A");
    expect(getCellAddress(projection, projection.rows[0].id, projection.columns[0].id)).toBe("A1");
  });

  it("dedupes duplicate row and column ids from concurrent move projections", () => {
    const envelope = createTeamGridDocument();
    const worksheet = getFirstVisibleWorksheet(envelope.teamgrid)!;
    worksheet.rowOrder.splice(1, 0, worksheet.rowOrder[0]);
    worksheet.columnOrder.splice(1, 0, worksheet.columnOrder[0]);

    const projection = projectWorksheet(worksheet);

    expect(projection.rows.filter((row) => row.id === worksheet.rowOrder[0])).toHaveLength(1);
    expect(projection.columns.filter((column) => column.id === worksheet.columnOrder[0])).toHaveLength(1);
  });

  it("parses visible addresses back to stable ids", () => {
    const envelope = createTeamGridDocument();
    const worksheet = getFirstVisibleWorksheet(envelope.teamgrid)!;
    const projection = projectWorksheet(worksheet);

    expect(parseCellAddress("B2", projection)).toEqual({
      rowId: projection.rows[1].id,
      columnId: projection.columns[1].id,
    });
  });

  it("labels columns after Z", () => {
    expect(columnIndexToLabel(0)).toBe("A");
    expect(columnIndexToLabel(25)).toBe("Z");
    expect(columnIndexToLabel(26)).toBe("AA");
  });
});
