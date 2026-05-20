import {
  type Chart,
  type ChartType,
  type SeriesRange,
  type TeamGridDocumentV1,
  type TwoCellAnchor,
  type Worksheet,
  type WorksheetId,
} from "@/features/document/lib/teamgridDocument";
import { createFormulaContext } from "@/features/formulas/lib";
import { formatSeriesRange } from "@/features/charts/lib/chartDataResolution";
import { readOoxmlZip, readZipText, relsPathForPart, writeOoxmlZip, writeZipText } from "@/features/xlsx/lib/ooxmlZip";
import { descendants, parseXml, readRelationships, serializeXml } from "@/features/xlsx/lib/ooxmlXml";

const REL_TYPE_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const REL_TYPE_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const CONTENT_TYPE_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml";
const CONTENT_TYPE_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const CATEGORY_AXIS_ID = 123456;
const VALUE_AXIS_ID = 123457;

interface SheetExportPart {
  worksheetId: WorksheetId;
  sheetName: string;
  sheetPath: string;
}

export function injectChartsIntoXlsxBuffer(buffer: ArrayBuffer, document: TeamGridDocumentV1) {
  const chartsByWorksheet = visibleWorksheets(document)
    .map(({ worksheet }) => ({ worksheet, charts: worksheet.chartOrder.map((id) => worksheet.chartsById[id]).filter((chart) => chart && !chart.deletedAt) }))
    .filter((entry) => entry.charts.length > 0);
  if (chartsByWorksheet.length === 0) {
    return buffer;
  }

  const zip = readOoxmlZip(buffer);
  const sheetParts = readGeneratedSheetParts(zip, document);
  const context = createFormulaContext(document.workbook, Object.fromEntries(sheetParts.map((part) => [part.worksheetId, part.sheetName])));
  let nextChartIndex = nextPartIndex(zip, /^xl\/charts\/chart(\d+)\.xml$/);
  let nextDrawingIndex = nextPartIndex(zip, /^xl\/drawings\/drawing(\d+)\.xml$/);

  for (const entry of chartsByWorksheet) {
    const sheetPart = sheetParts.find((part) => part.worksheetId === entry.worksheet.id);
    if (!sheetPart) {
      continue;
    }
    const drawingPath = `xl/drawings/drawing${nextDrawingIndex}.xml`;
    nextDrawingIndex += 1;
    const drawingRelsPath = relsPathForPart(drawingPath);
    const anchors: string[] = [];
    const drawingRelationships: string[] = [];

    entry.charts.forEach((chart, index) => {
      const chartPath = `xl/charts/chart${nextChartIndex}.xml`;
      nextChartIndex += 1;
      const relationshipId = `rId${index + 1}`;
      writeZipText(zip, chartPath, chart.raw?.chartXml ? patchChartXml(chart, chart.raw.chartXml, context) : createChartXml(chart, context));
      anchors.push(createDrawingAnchorXml(chart, relationshipId, entry.worksheet, index + 2));
      drawingRelationships.push(createRelationshipXml(relationshipId, REL_TYPE_CHART, `../charts/${chartPath.split("/").pop()}`));
      addContentTypeOverride(zip, chartPath, CONTENT_TYPE_CHART);
    });

    writeZipText(zip, drawingPath, createDrawingXml(anchors.join("")));
    writeZipText(zip, drawingRelsPath, createRelationshipsXml(drawingRelationships.join("")));
    addContentTypeOverride(zip, drawingPath, CONTENT_TYPE_DRAWING);
    attachDrawingToWorksheet(zip, sheetPart.sheetPath, drawingPath);
  }

  return writeOoxmlZip(zip);
}

function visibleWorksheets(document: TeamGridDocumentV1) {
  return document.workbook.worksheetOrder.flatMap((worksheetId) => {
    const worksheet = document.workbook.worksheetsById[worksheetId];
    return worksheet && !worksheet.deletedAt ? [{ worksheetId, worksheet }] : [];
  });
}

function readGeneratedSheetParts(zip: ReturnType<typeof readOoxmlZip>, document: TeamGridDocumentV1): SheetExportPart[] {
  const workbookXml = readZipText(zip, "xl/workbook.xml");
  const relsXml = readZipText(zip, "xl/_rels/workbook.xml.rels");
  if (!workbookXml || !relsXml) {
    return [];
  }
  const workbookDocument = parseXml(workbookXml);
  const relationships = readRelationships(relsXml);
  const sheets = descendants(workbookDocument, "sheet");
  return visibleWorksheets(document).flatMap(({ worksheetId }, index) => {
    const relationshipId = sheets[index] ? Array.from(sheets[index].attributes).find((attribute) => attribute.localName === "id")?.value : null;
    const sheetName = sheets[index]?.getAttribute("name");
    const relationship = relationshipId ? relationships.find((item) => item.id === relationshipId) : null;
    return relationship && sheetName ? [{ worksheetId, sheetName, sheetPath: `xl/${relationship.target.replace(/^\/?xl\//, "")}` }] : [];
  });
}

function patchChartXml(chart: Chart, chartXml: string, context: ReturnType<typeof createFormulaContext>) {
  const chartDocument = parseXml(chartXml);
  const titleText = descendants(chartDocument, "title")[0] ? descendants(descendants(chartDocument, "title")[0], "t")[0] : null;
  if (titleText && chart.title) {
    titleText.textContent = chart.title;
  }
  const formulas = descendants(chartDocument, "f");
  const ranges = chart.series.flatMap((series) => [chart.categoryAxis, series.name && typeof series.name === "object" ? series.name : null, series.values].filter(Boolean) as SeriesRange[]);
  formulas.forEach((formula, index) => {
    const range = ranges[index];
    if (range) {
      formula.textContent = formatSeriesRange(range, context);
    }
  });
  return serializeXml(chartDocument);
}

function createChartXml(chart: Chart, context: ReturnType<typeof createFormulaContext>) {
  const typeTag = chartTypeTag(chart.type);
  const typeSpecific = chart.type === "bar"
    ? '<c:barDir val="bar"/><c:grouping val="clustered"/>'
    : chart.type === "column"
      ? '<c:barDir val="col"/><c:grouping val="clustered"/>'
      : chart.type === "line"
        ? '<c:grouping val="standard"/>'
      : "";
  const series = chart.series.map((item, index) => createSeriesXml(chart, itemName(item.name, context), formatSeriesRange(item.values, context), index, item.color ?? chart.style?.colors?.[index])).join("");
  const categories = chart.categoryAxis ? `<c:cat><c:strRef><c:f>${escapeXml(formatSeriesRange(chart.categoryAxis, context))}</c:f></c:strRef></c:cat>` : "";
  const axisRefs = chart.type === "pie" ? "" : `<c:axId val="${CATEGORY_AXIS_ID}"/><c:axId val="${VALUE_AXIS_ID}"/>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">
  <c:chart>
    ${chart.title ? createTitleXml(chart.title) : ""}
    <c:plotArea>
      <${typeTag}>
        ${typeSpecific}
        ${chart.type === "pie" ? '<c:varyColors val="1"/>' : ""}
        ${series.replaceAll("<c:cat/>", categories)}
        ${axisRefs}
      </${typeTag}>
      ${chart.type === "pie" ? "" : createAxisXml(chart)}
    </c:plotArea>
    ${chart.legend?.position && chart.legend.position !== "none" ? `<c:legend><c:legendPos val="${legendPosition(chart.legend.position)}"/></c:legend>` : ""}
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;
}

function createSeriesXml(chart: Chart, name: string, values: string, index: number, color: string | undefined) {
  const category = chart.categoryAxis ? "<c:cat/>" : "";
  const shapeProperties = color ? `<c:spPr><a:solidFill><a:srgbClr val="${escapeXml(color.replace(/^#/, ""))}"/></a:solidFill></c:spPr>` : "";
  return `<c:ser>
    <c:idx val="${index}"/><c:order val="${index}"/>
    <c:tx><c:v>${escapeXml(name)}</c:v></c:tx>
    ${shapeProperties}
    ${category}
    <c:val><c:numRef><c:f>${escapeXml(values)}</c:f></c:numRef></c:val>
  </c:ser>`;
}

function createAxisXml(chart: Chart) {
  const categoryPosition = chart.type === "bar" ? "l" : "b";
  const valuePosition = chart.type === "bar" ? "b" : "l";
  const gridlines = chart.style?.showGridlines === false ? "" : "<c:majorGridlines/>";
  return `<c:catAx>
    <c:axId val="${CATEGORY_AXIS_ID}"/>
    <c:scaling><c:orientation val="minMax"/></c:scaling>
    <c:delete val="0"/>
    <c:axPos val="${categoryPosition}"/>
    <c:tickLblPos val="nextTo"/>
    <c:crossAx val="${VALUE_AXIS_ID}"/>
    <c:crosses val="autoZero"/>
    <c:auto val="1"/>
    <c:lblAlgn val="ctr"/>
    <c:lblOffset val="100"/>
  </c:catAx>
  <c:valAx>
    <c:axId val="${VALUE_AXIS_ID}"/>
    <c:scaling><c:orientation val="minMax"/></c:scaling>
    <c:delete val="0"/>
    <c:axPos val="${valuePosition}"/>
    ${gridlines}
    <c:numFmt formatCode="General" sourceLinked="1"/>
    <c:tickLblPos val="nextTo"/>
    <c:crossAx val="${CATEGORY_AXIS_ID}"/>
    <c:crosses val="autoZero"/>
    <c:crossBetween val="between"/>
  </c:valAx>`;
}

function itemName(name: Chart["series"][number]["name"], context: ReturnType<typeof createFormulaContext>) {
  if (typeof name === "string") {
    return name;
  }
  return name ? formatSeriesRange(name, context) : "Series";
}

function createDrawingXml(anchors: string) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">${anchors}</xdr:wsDr>`;
}

function createDrawingAnchorXml(chart: Chart, relationshipId: string, worksheet: Worksheet, objectId: number) {
  const anchor = anchorToIndexes(chart.anchor, worksheet);
  return `<xdr:twoCellAnchor>
    <xdr:from><xdr:col>${anchor.from.col}</xdr:col><xdr:colOff>${anchor.from.colOff}</xdr:colOff><xdr:row>${anchor.from.row}</xdr:row><xdr:rowOff>${anchor.from.rowOff}</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>${anchor.to.col}</xdr:col><xdr:colOff>${anchor.to.colOff}</xdr:colOff><xdr:row>${anchor.to.row}</xdr:row><xdr:rowOff>${anchor.to.rowOff}</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="${objectId}" name="${escapeXml(chart.title || `Chart ${objectId}`)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" xmlns:r="${NS_R}" r:id="${relationshipId}"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>`;
}

function anchorToIndexes(anchor: TwoCellAnchor, worksheet: Worksheet) {
  return {
    from: {
      col: Math.max(0, worksheet.columnOrder.indexOf(anchor.from.columnId)),
      row: Math.max(0, worksheet.rowOrder.indexOf(anchor.from.rowId)),
      colOff: anchor.from.colOffsetEmu,
      rowOff: anchor.from.rowOffsetEmu,
    },
    to: {
      col: Math.max(0, worksheet.columnOrder.indexOf(anchor.to.columnId)),
      row: Math.max(0, worksheet.rowOrder.indexOf(anchor.to.rowId)),
      colOff: anchor.to.colOffsetEmu,
      rowOff: anchor.to.rowOffsetEmu,
    },
  };
}

function attachDrawingToWorksheet(zip: ReturnType<typeof readOoxmlZip>, sheetPath: string, drawingPath: string) {
  const sheetXml = readZipText(zip, sheetPath);
  if (!sheetXml) {
    return;
  }
  const relPath = relsPathForPart(sheetPath);
  const existingRels = readRelationships(readZipText(zip, relPath) ?? "");
  const relationshipId = nextRelationshipId(existingRels.map((relationship) => relationship.id));
  const target = `../drawings/${drawingPath.split("/").pop()}`;
  writeZipText(zip, relPath, createRelationshipsXml([
    ...existingRels.map((relationship) => createRelationshipXml(relationship.id, relationship.type, relationship.target)),
    createRelationshipXml(relationshipId, REL_TYPE_DRAWING, target),
  ].join("")));
  const withNamespace = sheetXml.includes("xmlns:r=") ? sheetXml : sheetXml.replace("<worksheet ", `<worksheet xmlns:r="${NS_R}" `);
  const drawingTag = `<drawing r:id="${relationshipId}"/>`;
  const updatedSheet = withNamespace.includes("</worksheet>")
    ? withNamespace.replace("</worksheet>", `${drawingTag}</worksheet>`)
    : withNamespace;
  writeZipText(zip, sheetPath, updatedSheet);
}

function addContentTypeOverride(zip: ReturnType<typeof readOoxmlZip>, partPath: string, contentType: string) {
  const xml = readZipText(zip, "[Content_Types].xml");
  if (!xml || xml.includes(`PartName="/${partPath}"`)) {
    return;
  }
  const override = `<Override PartName="/${partPath}" ContentType="${contentType}"/>`;
  writeZipText(zip, "[Content_Types].xml", xml.replace("</Types>", `${override}</Types>`));
}

function createRelationshipsXml(relationships: string) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`;
}

function createRelationshipXml(id: string, type: string, target: string) {
  return `<Relationship Id="${escapeXml(id)}" Type="${escapeXml(type)}" Target="${escapeXml(target)}"/>`;
}

function createTitleXml(title: string) {
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeXml(title)}</a:t></a:r></a:p></c:rich></c:tx></c:title>`;
}

function chartTypeTag(type: ChartType) {
  return type === "line" ? "c:lineChart" : type === "pie" ? "c:pieChart" : "c:barChart";
}

function legendPosition(position: NonNullable<Chart["legend"]>["position"]) {
  return position === "right" ? "r" : position === "left" ? "l" : position === "top" ? "t" : "b";
}

function nextPartIndex(zip: ReturnType<typeof readOoxmlZip>, pattern: RegExp) {
  const indexes = Object.keys(zip).flatMap((path) => {
    const match = pattern.exec(path);
    return match ? [Number.parseInt(match[1], 10)] : [];
  });
  return Math.max(0, ...indexes) + 1;
}

function nextRelationshipId(ids: string[]) {
  const indexes = ids.flatMap((id) => {
    const match = /^rId(\d+)$/.exec(id);
    return match ? [Number.parseInt(match[1], 10)] : [];
  });
  return `rId${Math.max(0, ...indexes) + 1}`;
}

function escapeXml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
