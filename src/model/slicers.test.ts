// Table slicers: GenOffice cannot write them, so their parts are kept and put
// back into the file an export builds, the frame at its id-bound position.
// Skipped without pkg-node.
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'

import { blankXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'

import type { WorkbookFile, WorkbookSaveRequest } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { createWorkbook, loadWorkbook, readAttachmentBase64 } from '../haven/store'
import { AutomergeTestHost } from '../testing/automerge-host'
import { createNodeEngine, nodeEngineAvailable } from '../testing/node-engine'
import { saveWorkbookInBrowser } from '../xlsx/browser-save'
import type { XlsxEngine } from '../xlsx/engine'
import { extractSlicers } from '../xlsx/slicer-parts'
import { emptySaveRequest, workbookToXlsx } from './export'
import { createId } from './schema'
import { readWholeWorkbook } from './sidecar-read'
import { emptyStoredWorkbook, syncWorkbook } from './sync'

const SHEET = 'sheet-1'
const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const X15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'
const CACHE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<slicerCacheDefinition xmlns="${X14}" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="Datenschnitt_Region" sourceName="Region"><extLst><x:ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}" xmlns:x15="${X15}"><x15:tableSlicerCache tableId="TABLE_ID" column="2"/></x:ext></extLst></slicerCacheDefinition>`
const SLICER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<slicers xmlns="${X14}"><slicer name="Region" cache="Datenschnitt_Region" caption="Region" rowHeight="241300"/></slicers>`
const FRAME =
  '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:sle15="http://schemas.microsoft.com/office/drawing/2012/slicer" Requires="sle15">' +
  '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="Region"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer"><sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="Region"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>' +
  '<mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp></mc:Fallback></mc:AlternateContent>'

/** A table with a slicer on its Region column, made the way Excel stores it. */
async function slicerXlsx(): Promise<Uint8Array> {
  const request = emptySaveRequest('source')
  const rows = [['Monat', 'Region', 'Umsatz'], ['Jan', 'Nord', 10], ['Feb', 'Süd', 12], ['Mär', 'Nord', 9]]
  rows.forEach((row, rowIndex) =>
    row.forEach((value, column) => request.edits.push({ sheetId: SHEET, row: rowIndex, column, writeValue: true, value } as never)),
  )
  request.tableAdditions.push({
    sheetId: SHEET,
    area: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 2 },
    name: 'Umsatz',
    columnNames: ['Monat', 'Region', 'Umsatz'],
    bandedRows: true,
  } as WorkbookSaveRequest['tableAdditions'][number])
  request.visualAdditions.push({
    sheetId: SHEET,
    anchor: { fromRow: 1, fromColumn: 4, fromRowOffset: 0, fromColumnOffset: 0, toRow: 10, toColumn: 6, toRowOffset: 0, toColumnOffset: 0 },
    shape: { shapeType: 'rect' },
  } as WorkbookSaveRequest['visualAdditions'][number])
  const blank = new Uint8Array(await blankXlsxBuffer('Daten'))
  const zip = await JSZip.loadAsync((await saveWorkbookInBrowser(blank, new Map([[SHEET, 'Daten']]), request)).bytes)

  const tablePath = Object.keys(zip.files).find((name) => /^xl\/tables\/[^/]+\.xml$/.test(name))!
  const tableId = /<table\b[^>]*\sid="(\d+)"/.exec(await zip.file(tablePath)!.async('text'))![1]!
  zip.file('xl/slicerCaches/slicerCache1.xml', CACHE.replace('TABLE_ID', tableId))
  zip.file('xl/slicers/slicer1.xml', SLICER)
  let types = await zip.file('[Content_Types].xml')!.async('text')
  types = types.replace('</Types>', '<Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/><Override PartName="/xl/slicers/slicer1.xml" ContentType="application/vnd.ms-excel.slicer+xml"/></Types>')
  zip.file('[Content_Types].xml', types)
  const workbookRels = (await zip.file('xl/_rels/workbook.xml.rels')!.async('text')).replace('</Relationships>', '<Relationship Id="rId90" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/slicerCache1.xml"/></Relationships>')
  zip.file('xl/_rels/workbook.xml.rels', workbookRels)
  const workbookExt = `<ext uri="{46BE6895-7355-4a93-B00E-2C351335B9C9}" xmlns:x15="${X15}"><x15:slicerCaches xmlns:x14="${X14}"><x14:slicerCache r:id="rId90"/></x15:slicerCaches></ext>`
  let workbook = (await zip.file('xl/workbook.xml')!.async('text')).replace('</workbook>', `<extLst>${workbookExt}</extLst></workbook>`)
  workbook = workbook.replace('</sheets>', '</sheets><definedNames><definedName name="Datenschnitt_Region">#N/A</definedName></definedNames>')
  zip.file('xl/workbook.xml', workbook)
  const sheetPath = 'xl/worksheets/sheet1.xml'
  const sheetRelsPath = 'xl/worksheets/_rels/sheet1.xml.rels'
  zip.file(sheetRelsPath, (await zip.file(sheetRelsPath)!.async('text')).replace('</Relationships>', '<Relationship Id="rId91" Type="http://schemas.microsoft.com/office/2007/relationships/slicer" Target="../slicers/slicer1.xml"/></Relationships>'))
  const sheetExt = `<ext uri="{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}" xmlns:x14="${X14}"><x14:slicerList xmlns:x15="${X15}"><x14:slicer r:id="rId91"/></x14:slicerList></ext>`
  zip.file(sheetPath, (await zip.file(sheetPath)!.async('text')).replace('</worksheet>', `<extLst>${sheetExt}</extLst></worksheet>`))
  const drawingPath = Object.keys(zip.files).find((name) => /^xl\/drawings\/drawing[^/]*\.xml$/.test(name))!
  const drawing = (await zip.file(drawingPath)!.async('text')).replace(/(<\/xdr:to>)[\s\S]*(<xdr:clientData\b)/, `$1${FRAME}$2`)
  zip.file(drawingPath, drawing)
  return zip.generateAsync({ type: 'uint8array' })
}

async function read(engine: XlsxEngine, bytes: Uint8Array) {
  const file = (await engine.open(bytes, 'de')) as WorkbookFile
  const result = await readWholeWorkbook(engine, file)
  await engine.close(file.sessionId)
  return result
}

describe('slicers', () => {
  it.skipIf(!nodeEngineAvailable)('keeps a table slicer and writes it back at its position', async () => {
    const engine = createNodeEngine()
    const host = new AutomergeTestHost()
    const bytes = await slicerXlsx()
    const slicers = await extractSlicers(bytes)
    expect(slicers?.caches[0]?.xml).toContain('tableId="{{table:Umsatz}}"')
    expect(Object.keys(slicers!.frames)).toEqual(['Region'])
    const source = await read(engine, bytes)
    expect(source.file.visuals.map((visual) => visual.kind)).toEqual(['slicer'])

    const { next, writes } = syncWorkbook({ stored: emptyStoredWorkbook(), storedIdByFileId: new Map(), read: source, slicers: slicers ?? null })
    const id = await createWorkbook(host.connection, 'Datenschnitt', next, writes)
    const loaded = await loadWorkbook(host.connection, id)
    const sheet = Object.values(loaded.stored.workbook.worksheetsById)[0]!
    expect(Object.values(sheet.visualsById).map((visual) => visual.slicer)).toEqual([{ name: 'Region' }])
    const loadImage = (attachment: string) => readAttachmentBase64(host.connection, id, attachment)

    const back = (await workbookToXlsx(loaded.stored.workbook, engine, { loadImage })).bytes
    expect(await extractSlicers(back)).toEqual(slicers)
    const backRead = await read(engine, back)
    expect(backRead.file.visuals.map(({ kind, anchor, text }) => ({ kind, anchor, text }))).toEqual(
      source.file.visuals.map(({ kind, anchor, text }) => ({ kind, anchor, text })),
    )
    expect(backRead.file.definedNames).toEqual(source.file.definedNames)
    expect(next.namesById).toEqual({})

    // A row inserted above moves the frame; the slicer stays.
    sheet.rowOrder = [createId('r'), ...sheet.rowOrder]
    const moved = await read(engine, (await workbookToXlsx(loaded.stored.workbook, engine, { loadImage })).bytes)
    expect(moved.file.visuals[0]).toMatchObject({ kind: 'slicer', anchor: { fromRow: 2 } })
  })
})
