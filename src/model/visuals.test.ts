// Charts, pictures and shapes through the real pipeline: GenOffice's gateway
// writes the xlsx, the WASM sidecar (Node build) reads it back. Skipped when
// native/xlsx-wasm/pkg-node has not been built (`pnpm build:wasm:node`).
import { describe, expect, it } from 'vitest'

import { blankXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'

import type { WorkbookFile, WorkbookSaveRequest } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { createWorkbook, loadWorkbook, readAttachmentBase64, writeWorkbook, type LoadedWorkbook } from '../haven/store'
import { AutomergeTestHost } from '../testing/automerge-host'
import { createNodeEngine, nodeEngineAvailable } from '../testing/node-engine'
import { saveWorkbookInBrowser } from '../xlsx/browser-save'
import type { XlsxEngine } from '../xlsx/engine'
import { emptySaveRequest, workbookToXlsx } from './export'
import { readWholeWorkbook } from './sidecar-read'
import { emptyStoredWorkbook, syncWorkbook, visualFileKey } from './sync'

const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
const anchor = (row: number) => ({
  fromRow: row,
  fromColumn: 2,
  fromRowOffset: 0,
  fromColumnOffset: 0,
  toRow: row + 3,
  toColumn: 5,
  toRowOffset: 0,
  toColumnOffset: 0,
})

/** An xlsx with three numbers, a picture, a chart over the numbers and a shape. */
async function sourceXlsx(): Promise<Uint8Array> {
  const request = emptySaveRequest('source')
  for (let row = 0; row < 3; row += 1) {
    request.edits.push({ sheetId: 'sheet-1', row, column: 0, writeValue: true, value: (row + 1) * 10 } as never)
  }
  request.visualAdditions.push(
    { sheetId: 'sheet-1', anchor: anchor(0), image: { mediaType: 'image/png', base64: PNG } },
    {
      sheetId: 'sheet-1',
      anchor: anchor(5),
      chart: {
        chartType: 'column',
        title: 'Numbers',
        series: [{ name: 'S', categories: ['a', 'b', 'c'], values: [10, 20, 30], valuesRef: 'Sheet1!$A$1:$A$3' }],
      },
    },
    { sheetId: 'sheet-1', anchor: anchor(10), shape: { shapeType: 'rect', fillColor: '#FF0000', text: 'Hi' } },
  )
  const blank = new Uint8Array(await blankXlsxBuffer('Sheet1'))
  return (await saveWorkbookInBrowser(blank, new Map([['sheet-1', 'Sheet1']]), request)).bytes
}

async function importXlsx(host: AutomergeTestHost, engine: XlsxEngine, bytes: Uint8Array) {
  const file = (await engine.open(bytes, 'en')) as WorkbookFile
  const { next, writes } = syncWorkbook({
    stored: emptyStoredWorkbook(),
    storedIdByFileId: new Map(),
    read: await readWholeWorkbook(engine, file),
  })
  for (const image of writes.newImages) {
    image.bytes = await engine.readMediaBytes({ sessionId: file.sessionId, visualId: image.fileVisualId })
  }
  return createWorkbook(host.connection, 'Visuals', next, writes)
}

/** Opens a stored workbook the way the editor does: export, then read the file. */
async function openSession(host: AutomergeTestHost, engine: XlsxEngine, loaded: LoadedWorkbook) {
  const { bytes, visualOrder } = await workbookToXlsx(loaded.stored.workbook, engine, {
    loadImage: (name) => readAttachmentBase64(host.connection, loaded.id, name),
  })
  const file = (await engine.open(bytes, 'en')) as WorkbookFile
  const sheetId = loaded.stored.workbook.worksheetOrder[0]!
  const fileSheetId = file.sheets[0]!.id
  const ids = visualOrder.get(sheetId) ?? []
  const byFileKey = new Map(file.visuals.map((visual, index) => [visualFileKey(visual), ids[index]!]))
  return { bytes, file, fileSheetId, storedIdByFileId: new Map([[fileSheetId, sheetId]]), byFileKey, visualOrder }
}

/** The editor saves `request` on an open session; returns what the store would write. */
async function save(
  host: AutomergeTestHost,
  engine: XlsxEngine,
  loaded: LoadedWorkbook,
  build: (session: Awaited<ReturnType<typeof openSession>>) => WorkbookSaveRequest,
) {
  const session = await openSession(host, engine, loaded)
  const request = build(session)
  const names = new Map([[session.fileSheetId, session.file.sheets[0]!.name]])
  const saved = (await saveWorkbookInBrowser(session.bytes, names, request)).bytes
  const file = (await engine.open(saved, 'en')) as WorkbookFile
  const { writes } = syncWorkbook({
    stored: loaded.stored,
    storedIdByFileId: session.storedIdByFileId,
    request,
    read: await readWholeWorkbook(engine, file),
    visuals: { byFileKey: session.byFileKey, orderBySheet: session.visualOrder },
  })
  for (const image of writes.newImages) {
    image.bytes = await engine.readMediaBytes({ sessionId: file.sessionId, visualId: image.fileVisualId })
  }
  await writeWorkbook(host.connection, loaded, writes)
  return writes
}

const visualsOf = (loaded: LoadedWorkbook) => Object.values(loaded.stored.workbook.worksheetsById)[0]!

describe.skipIf(!nodeEngineAvailable)('visuals', () => {
  it('stores a picture, a chart and a shape and opens them again', async () => {
    const host = new AutomergeTestHost()
    const engine = createNodeEngine()
    const id = await importXlsx(host, engine, await sourceXlsx())
    const loaded = await loadWorkbook(host.connection, id)
    const sheet = visualsOf(loaded)
    expect(sheet.visualOrder.map((visualId) => sheet.visualsById[visualId]!.kind)).toEqual(['image', 'chart', 'shape'])

    const session = await openSession(host, engine, loaded)
    expect(session.file.visuals.map((visual) => visual.kind)).toEqual(['image', 'chart', 'shape'])
    const chart = session.file.visuals[1]!.chart!
    expect(chart.title).toBe('Numbers')
    expect(chart.series[0]!.valuesRef).toBe('Sheet1!$A$1:$A$3')
    const media = await engine.readMediaBytes({ sessionId: session.file.sessionId, visualId: session.file.visuals[0]!.id })
    expect(media.length).toBeGreaterThan(0)
  })

  it('moves visuals and chart ranges with an inserted row, without rewriting them', async () => {
    const host = new AutomergeTestHost()
    const engine = createNodeEngine()
    const id = await importXlsx(host, engine, await sourceXlsx())
    const before = await loadWorkbook(host.connection, id)
    const writes = await save(host, engine, before, (session) => {
      const request = emptySaveRequest(session.file.sessionId)
      request.structuralOps.push({ sheetId: session.fileSheetId, kind: 'insert-rows', index: 0, count: 1 })
      return request
    })
    expect(writes.top.set.filter((op) => op.path.includes('visualsById'))).toEqual([])
    const after = await loadWorkbook(host.connection, id)
    expect(visualsOf(after).visualsById).toEqual(visualsOf(before).visualsById)
    const session = await openSession(host, engine, after)
    expect(session.file.visuals.map((visual) => visual.anchor.fromRow)).toEqual([1, 6, 11])
    expect(session.file.visuals[1]!.chart!.series[0]!.valuesRef).toBe('Sheet1!$A$2:$A$4')
  })

  it('removes one visual and keeps the ids of the others', async () => {
    const host = new AutomergeTestHost()
    const engine = createNodeEngine()
    const id = await importXlsx(host, engine, await sourceXlsx())
    const before = await loadWorkbook(host.connection, id)
    const [pictureId, chartId, shapeId] = visualsOf(before).visualOrder
    await save(host, engine, before, (session) => {
      const request = emptySaveRequest(session.file.sessionId)
      const picture = session.file.visuals[0]!
      request.visualEdits.push({ drawingPath: picture.drawingPath!, drawingIndex: picture.drawingIndex!, remove: true })
      request.visualAdditions.push({ sheetId: session.fileSheetId, anchor: anchor(20), shape: { shapeType: 'ellipse' } })
      return request
    })
    const after = visualsOf(await loadWorkbook(host.connection, id))
    expect(after.visualOrder.slice(0, 2)).toEqual([chartId, shapeId])
    expect(after.visualOrder).toHaveLength(3)
    expect(after.visualOrder).not.toContain(pictureId)
    expect(after.visualsById[after.visualOrder[2]!]!.shape?.shapeType).toBe('ellipse')
  })
})
