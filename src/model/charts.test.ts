// Charts beyond GenOffice's chart-add: the chart part is kept (with its style
// part) and put back over the chart an export adds, its references moving
// with inserted rows. Skipped without pkg-node.
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'

import { blankXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'

import type { WorkbookFile, WorkbookSaveRequest } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { createWorkbook, loadWorkbook, readAttachmentBase64 } from '../haven/store'
import { AutomergeTestHost } from '../testing/automerge-host'
import { createNodeEngine, nodeEngineAvailable } from '../testing/node-engine'
import { saveWorkbookInBrowser } from '../xlsx/browser-save'
import { extractCharts } from '../xlsx/chart-parts'
import type { XlsxEngine } from '../xlsx/engine'
import { emptySaveRequest, workbookToXlsx } from './export'
import { createId } from './schema'
import { readWholeWorkbook } from './sidecar-read'
import { emptyStoredWorkbook, syncWorkbook } from './sync'

const SHEET = 'sheet-1'

const BUBBLE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Blasen</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:plotArea><c:bubbleChart><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:spPr><a:solidFill><a:srgbClr val="7030A0"/></a:solidFill></c:spPr><c:xVal><c:numRef><c:f>Daten!$A$2:$A$4</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>Daten!$B$2:$B$4</c:f></c:numRef></c:yVal><c:bubbleSize><c:numRef><c:f>Daten!$C$2:$C$4</c:f></c:numRef></c:bubbleSize><c:bubble3D val="0"/></c:ser><c:bubbleScale val="100"/><c:axId val="1"/><c:axId val="2"/></c:bubbleChart><c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`
const STYLE = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="201"/>'

/** A file with a bubble chart: GenOffice adds a scatter chart, its part is swapped for a bubble chart. */
async function bubbleXlsx(): Promise<Uint8Array> {
  const request = emptySaveRequest('source')
  const rows = [['X', 'Y', 'Größe'], [1, 10, 3], [2, 14, 5], [3, 9, 8]]
  rows.forEach((row, rowIndex) =>
    row.forEach((value, column) => request.edits.push({ sheetId: SHEET, row: rowIndex, column, writeValue: true, value } as never)),
  )
  request.visualAdditions.push({
    sheetId: SHEET,
    anchor: { fromRow: 1, fromColumn: 4, fromRowOffset: 0, fromColumnOffset: 0, toRow: 12, toColumn: 10, toRowOffset: 0, toColumnOffset: 0 },
    chart: { chartType: 'scatter', title: 'Platzhalter', series: [{ name: 'Y', categories: [], values: [10, 14, 9], valuesRef: 'Daten!$B$2:$B$4' }] },
  } as WorkbookSaveRequest['visualAdditions'][number])
  const blank = new Uint8Array(await blankXlsxBuffer('Daten'))
  const saved = (await saveWorkbookInBrowser(blank, new Map([[SHEET, 'Daten']]), request)).bytes
  const zip = await JSZip.loadAsync(saved)
  const chartPath = Object.keys(zip.files).find((name) => /^xl\/charts\/chart[^/]*\.xml$/.test(name))!
  zip.file(chartPath, BUBBLE)
  zip.file(chartPath.replace(/chart(\d*)\.xml$/, '_rels/chart$1.xml.rels'),
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style9.xml"/></Relationships>')
  zip.file('xl/charts/style9.xml', STYLE)
  const types = (await zip.file('[Content_Types].xml')!.async('text')).replace('</Types>', '<Override PartName="/xl/charts/style9.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/></Types>')
  zip.file('[Content_Types].xml', types)
  return zip.generateAsync({ type: 'uint8array' })
}

async function read(engine: XlsxEngine, bytes: Uint8Array) {
  const file = (await engine.open(bytes, 'de')) as WorkbookFile
  const result = await readWholeWorkbook(engine, file)
  await engine.close(file.sessionId)
  return result
}

describe('kept chart parts', () => {
  it.skipIf(!nodeEngineAvailable)('keeps a bubble chart with its style, moved with an inserted row', async () => {
    const engine = createNodeEngine()
    const host = new AutomergeTestHost()
    const bytes = await bubbleXlsx()
    const charts = await extractCharts(bytes)
    const [source] = [...charts.values()]
    expect(source!.refs).toEqual(['Daten!$A$2:$A$4', 'Daten!$B$2:$B$4', 'Daten!$C$2:$C$4'])

    const { next, writes } = syncWorkbook({ stored: emptyStoredWorkbook(), storedIdByFileId: new Map(), read: await read(engine, bytes), charts })
    const id = await createWorkbook(host.connection, 'Blasen', next, writes)
    const loaded = await loadWorkbook(host.connection, id)
    const sheet = Object.values(loaded.stored.workbook.worksheetsById)[0]!
    expect(Object.values(sheet.visualsById).map((visual) => visual.chart?.raw?.refs.length)).toEqual([3])
    const loadImage = (attachment: string) => readAttachmentBase64(host.connection, id, attachment)

    const back = [...(await extractCharts((await workbookToXlsx(loaded.stored.workbook, engine, { loadImage })).bytes)).values()]
    expect(back).toEqual([source])

    sheet.rowOrder = [createId('r'), ...sheet.rowOrder]
    const moved = [...(await extractCharts((await workbookToXlsx(loaded.stored.workbook, engine, { loadImage })).bytes)).values()]
    expect(moved[0]!.refs).toEqual(['Daten!$A$3:$A$5', 'Daten!$B$3:$B$5', 'Daten!$C$3:$C$5'])
    expect(moved[0]!.parts).toEqual(source!.parts)
  })
})
