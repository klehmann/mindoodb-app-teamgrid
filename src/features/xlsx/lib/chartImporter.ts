import {
  createId,
  type Chart,
  type ChartLegend,
  type ChartSeries,
  type ChartType,
  type SeriesRange,
  type TwoCellAnchor,
  type Workbook,
  type WorksheetId,
} from "@/features/document/lib/teamgridDocument";
import { createFormulaContext } from "@/features/formulas/lib";
import { columnLabelToIndex } from "@/features/grid/lib/gridProjection";
import { readOoxmlZip, readZipText, relsPathForPart, resolveZipTarget } from "@/features/xlsx/lib/ooxmlZip";
import { attr, child, children, descendants, parseXml, readRelationships, textOf } from "@/features/xlsx/lib/ooxmlXml";

const REL_TYPE_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const REL_TYPE_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

interface WorkbookSheetPart {
  name: string;
  path: string;
  state?: string;
}

export function importChartsFromXlsx(buffer: ArrayBuffer | Uint8Array, workbook: Workbook) {
  const zip = readOoxmlZip(buffer);
  const workbookSheets = readWorkbookSheetParts(zip);
  const visibleSheets = workbookSheets.filter((sheet) => sheet.state !== "hidden" && sheet.state !== "veryHidden");
  const worksheetIdByPath = new Map<string, WorksheetId>();
  visibleSheets.forEach((sheet, index) => {
    const worksheetId = workbook.worksheetOrder[index];
    if (worksheetId) {
      worksheetIdByPath.set(sheet.path, worksheetId);
    }
  });

  const formulaContext = createFormulaContext(workbook);
  for (const sheet of visibleSheets) {
    const worksheetId = worksheetIdByPath.get(sheet.path);
    const worksheet = worksheetId ? workbook.worksheetsById[worksheetId] : null;
    if (!worksheet || worksheet.deletedAt) {
      continue;
    }
    const worksheetRels = readRelationships(readZipText(zip, relsPathForPart(sheet.path)) ?? "");
    const drawingRel = worksheetRels.find((relationship) => relationship.type === REL_TYPE_DRAWING);
    if (!drawingRel) {
      continue;
    }
    const drawingPath = resolveZipTarget(sheet.path, drawingRel.target);
    const drawingXml = readZipText(zip, drawingPath);
    const drawingRelsXml = readZipText(zip, relsPathForPart(drawingPath));
    if (!drawingXml || !drawingRelsXml) {
      continue;
    }
    const drawingRels = readRelationships(drawingRelsXml);
    const drawingDocument = parseXml(drawingXml);
    for (const anchor of descendants(drawingDocument, "twoCellAnchor")) {
      const chartRef = descendants(anchor, "chart")[0];
      const relationshipId = attr(chartRef, "id");
      if (!relationshipId) {
        continue;
      }
      const chartRel = drawingRels.find((relationship) => relationship.id === relationshipId && relationship.type === REL_TYPE_CHART);
      if (!chartRel) {
        continue;
      }
      const chartPath = resolveZipTarget(drawingPath, chartRel.target);
      const chartXml = readZipText(zip, chartPath);
      if (!chartXml) {
        continue;
      }
      const chart = parseChart(chartXml, chartPath, drawingXml, drawingPath, relsPathForPart(drawingPath), relsPathForPart(sheet.path), relationshipId, anchor, worksheet.id, formulaContext);
      if (chart) {
        worksheet.chartsById[chart.id] = chart;
        worksheet.chartOrder.push(chart.id);
      }
    }
  }
}

function parseChart(
  chartXml: string,
  chartPath: string,
  drawingXml: string,
  drawingPath: string,
  drawingRelPath: string,
  worksheetRelPath: string,
  relationshipId: string,
  anchorElement: Element,
  currentWorksheetId: WorksheetId,
  context: ReturnType<typeof createFormulaContext>,
): Chart | null {
  const chartDocument = parseXml(chartXml);
  const plotArea = descendants(chartDocument, "plotArea")[0];
  const chartTypeElement = plotArea ? children(plotArea, "barChart")[0] ?? children(plotArea, "lineChart")[0] ?? children(plotArea, "pieChart")[0] : null;
  if (!chartTypeElement) {
    return null;
  }
  const type = readChartType(chartTypeElement);
  if (!type) {
    return null;
  }
  const anchor = parseTwoCellAnchor(anchorElement, context.workbook.worksheetsById[currentWorksheetId]);
  if (!anchor) {
    return null;
  }
  const series = children(chartTypeElement, "ser").flatMap((seriesElement, index) => (
    parseSeries(seriesElement, index, currentWorksheetId, context)
  ));
  if (series.length === 0) {
    return null;
  }
  const categoryAxis = readSeriesFormula(children(chartTypeElement, "ser")[0] ? child(children(chartTypeElement, "ser")[0], "cat") : null, currentWorksheetId, context);
  const title = readChartTitle(chartDocument);
  const legend = readLegend(chartDocument);
  return {
    id: createId("chart"),
    type,
    series,
    anchor,
    raw: {
      chartXml,
      drawingXml,
      chartPath,
      drawingPath,
      drawingRelPath,
      worksheetRelPath,
      relationshipId,
    },
    ...(title ? { title } : {}),
    ...(categoryAxis ? { categoryAxis } : {}),
    ...(legend ? { legend } : {}),
  };
}

function readWorkbookSheetParts(zip: ReturnType<typeof readOoxmlZip>): WorkbookSheetPart[] {
  const workbookXml = readZipText(zip, "xl/workbook.xml");
  const relsXml = readZipText(zip, "xl/_rels/workbook.xml.rels");
  if (!workbookXml || !relsXml) {
    return [];
  }
  const workbookDocument = parseXml(workbookXml);
  const relationships = readRelationships(relsXml);
  return descendants(workbookDocument, "sheet").flatMap((sheet) => {
    const relationshipId = attr(sheet, "id");
    const relationship = relationshipId ? relationships.find((item) => item.id === relationshipId) : null;
    const name = sheet.getAttribute("name");
    if (!relationship || !name) {
      return [];
    }
    return [{
      name,
      state: sheet.getAttribute("state") ?? undefined,
      path: resolveZipTarget("xl/workbook.xml", relationship.target),
    }];
  });
}

function readChartType(element: Element): ChartType | null {
  if (element.localName === "lineChart") {
    return "line";
  }
  if (element.localName === "pieChart") {
    return "pie";
  }
  if (element.localName === "barChart") {
    const barDir = child(element, "barDir")?.getAttribute("val");
    return barDir === "bar" ? "bar" : "column";
  }
  return null;
}

function parseSeries(seriesElement: Element, index: number, currentWorksheetId: WorksheetId, context: ReturnType<typeof createFormulaContext>): ChartSeries[] {
  const values = readSeriesFormula(child(seriesElement, "val"), currentWorksheetId, context);
  if (!values) {
    return [];
  }
  const tx = child(seriesElement, "tx");
  const nameRange = readSeriesFormula(tx, currentWorksheetId, context);
  const literalName = textOf(child(tx ?? seriesElement, "strRef"), "v") ?? textOf(tx ?? seriesElement, "v");
  const color = readSeriesColor(seriesElement);
  const name = nameRange ?? literalName ?? undefined;
  return [{
    id: `series_${index + 1}`,
    values,
    ...(name !== undefined ? { name } : {}),
    ...(color ? { color } : {}),
  }];
}

function readSeriesFormula(element: Element | null, currentWorksheetId: WorksheetId, context: ReturnType<typeof createFormulaContext>): SeriesRange | null {
  if (!element) {
    return null;
  }
  const formula = descendants(element, "f")[0]?.textContent?.trim();
  return formula ? parseA1Range(formula, currentWorksheetId, context) : null;
}

export function parseA1Range(source: string, currentWorksheetId: WorksheetId, context: ReturnType<typeof createFormulaContext>): SeriesRange | null {
  const match = /^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?\$?([A-Z]{1,3})\$?([1-9][0-9]*)(?::\$?([A-Z]{1,3})\$?([1-9][0-9]*))?$/i.exec(source.trim());
  if (!match) {
    return null;
  }
  const sheetName = (match[1]?.replace(/''/g, "'") ?? match[2])?.trim();
  const worksheetId = sheetName ? context.worksheetIdsByName.get(sheetName.toLowerCase())?.[0] : currentWorksheetId;
  const projection = worksheetId ? context.projectionsByWorksheetId.get(worksheetId) : null;
  if (!worksheetId || !projection) {
    return null;
  }
  const startColumn = projection.columns[columnLabelToIndex(match[3].toUpperCase())];
  const startRow = projection.rows[Number.parseInt(match[4], 10) - 1];
  const endColumn = projection.columns[columnLabelToIndex((match[5] ?? match[3]).toUpperCase())];
  const endRow = projection.rows[Number.parseInt(match[6] ?? match[4], 10) - 1];
  if (!startColumn || !startRow || !endColumn || !endRow) {
    return null;
  }
  return {
    worksheetId,
    startRowId: startRow.id,
    endRowId: endRow.id,
    startColumnId: startColumn.id,
    endColumnId: endColumn.id,
    excelA1: source,
  };
}

function parseTwoCellAnchor(anchor: Element, worksheet: { rowOrder: string[]; columnOrder: string[] }): TwoCellAnchor | null {
  const from = child(anchor, "from");
  const to = child(anchor, "to");
  if (!from || !to) {
    return null;
  }
  const fromPoint = parseAnchorPoint(from, worksheet);
  const toPoint = parseAnchorPoint(to, worksheet);
  return fromPoint && toPoint ? { from: fromPoint, to: toPoint } : null;
}

function parseAnchorPoint(point: Element, worksheet: { rowOrder: string[]; columnOrder: string[] }) {
  const columnIndex = Number.parseInt(textOf(point, "col") ?? "", 10);
  const rowIndex = Number.parseInt(textOf(point, "row") ?? "", 10);
  const columnId = worksheet.columnOrder[columnIndex];
  const rowId = worksheet.rowOrder[rowIndex];
  if (!rowId || !columnId) {
    return null;
  }
  return {
    rowId,
    columnId,
    rowOffsetEmu: Number.parseInt(textOf(point, "rowOff") ?? "0", 10) || 0,
    colOffsetEmu: Number.parseInt(textOf(point, "colOff") ?? "0", 10) || 0,
  };
}

function readChartTitle(document: Document) {
  const title = descendants(document, "title")[0];
  const pieces = title ? descendants(title, "t").map((text) => text.textContent ?? "").join("") : "";
  return pieces.trim() || null;
}

function readLegend(document: Document): ChartLegend | null {
  const legendPos = descendants(document, "legendPos")[0]?.getAttribute("val");
  if (!legendPos) {
    return null;
  }
  const position = legendPos === "r"
    ? "right"
    : legendPos === "l"
      ? "left"
      : legendPos === "t"
        ? "top"
        : legendPos === "b"
          ? "bottom"
          : "none";
  return { position };
}

function readSeriesColor(seriesElement: Element) {
  const solidFill = descendants(seriesElement, "solidFill")[0];
  const rgb = solidFill ? descendants(solidFill, "srgbClr")[0]?.getAttribute("val") : null;
  return rgb && /^[0-9a-f]{6}$/i.test(rgb) ? `#${rgb.toLowerCase()}` : null;
}
