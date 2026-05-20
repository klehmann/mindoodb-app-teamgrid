import { describe, expect, it } from "vitest";

import { createTeamGridDocument } from "@/features/document/lib/teamgridDocument";
import { importChartsFromXlsx } from "@/features/xlsx/lib/chartImporter";
import { writeOoxmlZip, writeZipText, type OoxmlZip } from "@/features/xlsx/lib/ooxmlZip";

describe("XLSX chart importer", () => {
  it("imports a basic column chart from OOXML drawing parts", () => {
    const envelope = createTeamGridDocument();
    const worksheet = envelope.teamgrid.workbook.worksheetsById[envelope.teamgrid.workbook.worksheetOrder[0]];

    importChartsFromXlsx(createChartWorkbookZip(), envelope.teamgrid.workbook);

    expect(worksheet.chartOrder).toHaveLength(1);
    const chart = worksheet.chartsById[worksheet.chartOrder[0]];
    expect(chart.type).toBe("column");
    expect(chart.title).toBe("Revenue");
    expect(chart.legend).toEqual({ position: "right" });
    expect(chart.series[0].values).toMatchObject({
      worksheetId: worksheet.id,
      startRowId: worksheet.rowOrder[1],
      endRowId: worksheet.rowOrder[2],
      startColumnId: worksheet.columnOrder[1],
      endColumnId: worksheet.columnOrder[1],
    });
    expect(chart.categoryAxis).toMatchObject({
      startColumnId: worksheet.columnOrder[0],
      endColumnId: worksheet.columnOrder[0],
    });
    expect(chart.anchor.from).toMatchObject({ rowId: worksheet.rowOrder[0], columnId: worksheet.columnOrder[3] });
  });
});

function createChartWorkbookZip() {
  const zip: OoxmlZip = {};
  writeZipText(zip, "xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet 1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`);
  writeZipText(zip, "xl/_rels/workbook.xml.rels", relationships(`<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/>`));
  writeZipText(zip, "xl/worksheets/sheet1.xml", "<worksheet/>");
  writeZipText(zip, "xl/worksheets/_rels/sheet1.xml.rels", relationships(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>`));
  writeZipText(zip, "xl/drawings/drawing1.xml", `<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame>
  </xdr:twoCellAnchor>
</xdr:wsDr>`);
  writeZipText(zip, "xl/drawings/_rels/drawing1.xml.rels", relationships(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>`));
  writeZipText(zip, "xl/charts/chart1.xml", `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:tx><c:v>Actual</c:v></c:tx><c:cat><c:strRef><c:f>'Sheet 1'!$A$2:$A$3</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>'Sheet 1'!$B$2:$B$3</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea>
    <c:legend><c:legendPos val="r"/></c:legend>
  </c:chart>
</c:chartSpace>`);
  return writeOoxmlZip(zip);
}

function relationships(inner: string) {
  return `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;
}
