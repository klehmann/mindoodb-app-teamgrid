// Notes, hyperlinks, the autoFilter, print settings and tab colors through the
// real pipeline: GenOffice's gateway writes the xlsx, the WASM sidecar (Node
// build) reads it back. Skipped without native/xlsx-wasm/pkg-node.
import { describe, expect, it } from 'vitest'

import { blankXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'

import type { WorkbookFile, WorkbookSaveRequest } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { createNodeEngine, nodeEngineAvailable } from '../testing/node-engine'
import { saveWorkbookInBrowser } from '../xlsx/browser-save'
import type { XlsxEngine } from '../xlsx/engine'
import { emptySaveRequest, workbookToXlsx } from './export'
import { parseHeaderFooter } from './sheet-features'
import { readWholeWorkbook, type SidecarWorkbook } from './sidecar-read'
import { emptyStoredWorkbook, syncWorkbook } from './sync'

const SHEET = 'sheet-1'
const area = { startRow: 0, startColumn: 0, endRow: 3, endColumn: 2 }

async function sourceXlsx(): Promise<Uint8Array> {
  const request = emptySaveRequest('source')
  const rows = [['Name', 'Team', 'Wert'], ['a', 'x', 1], ['b', 'y', 2], ['c', 'x', 3]]
  rows.forEach((row, rowIndex) =>
    row.forEach((value, column) =>
      request.edits.push({ sheetId: SHEET, row: rowIndex, column, writeValue: true, value } as never),
    ),
  )
  request.noteStates.push({ sheetId: SHEET, notes: [{ row: 1, column: 0, author: 'Karsten', text: 'Prüfen' }] })
  request.hyperlinkEdits.push({ sheetId: SHEET, row: 2, column: 0, target: 'https://mindoo.de/' })
  request.pageSetupStates.push({
    sheetId: SHEET,
    orientation: 'landscape',
    printArea: 'A1:C4',
    printTitles: '1:1',
    header: { center: 'Bericht &D' },
    rowBreaks: [2],
  } as WorkbookSaveRequest['pageSetupStates'][number])
  request.filterStates.push({
    sheetId: SHEET,
    filter: { range: area, columns: [{ colId: 1, values: ['x'] }] },
    hiddenRows: [2],
    visibilityRange: area,
  } as WorkbookSaveRequest['filterStates'][number])
  const blank = new Uint8Array(await blankXlsxBuffer('Sheet1'))
  return (await saveWorkbookInBrowser(blank, new Map([[SHEET, 'Sheet1']]), request)).bytes
}

async function read(engine: XlsxEngine, bytes: Uint8Array): Promise<SidecarWorkbook> {
  const file = (await engine.open(bytes, 'de')) as WorkbookFile
  const workbook = await readWholeWorkbook(engine, file)
  await engine.close(file.sessionId)
  return workbook
}

describe('sheet features', () => {
  it('splits headers and footers into their sections', () => {
    expect(parseHeaderFooter('&LLinks&C&P von &N&RRechts && mehr')).toEqual({
      left: 'Links',
      center: '&P von &N',
      right: 'Rechts && mehr',
    })
    expect(parseHeaderFooter('Nur Mitte')).toEqual({ center: 'Nur Mitte' })
  })

  it.skipIf(!nodeEngineAvailable)('stores them by id and writes them back', async () => {
    const engine = createNodeEngine()
    const source = await read(engine, await sourceXlsx())
    const { next } = syncWorkbook({ stored: emptyStoredWorkbook(), storedIdByFileId: new Map(), read: source })
    const sheet = Object.values(next.worksheetsById)[0]!
    const [r1, r2, r3, r4] = sheet.rowOrder
    const [cA, , cC] = sheet.columnOrder

    expect(sheet.cellsById[`${r2}:${cA}`]?.note).toEqual({ author: 'Karsten', text: 'Prüfen' })
    expect(sheet.cellsById[`${r3}:${cA}`]?.link).toBe('https://mindoo.de/')
    expect(sheet.autoFilter).toEqual({
      area: { startRowId: r1, startColumnId: cA, endRowId: r4, endColumnId: cC },
      columns: [{ colId: 1, values: ['x'] }],
    })
    expect(sheet.pageSetup).toMatchObject({
      orientation: 'landscape',
      header: { center: 'Bericht &D' },
      printArea: { startRowId: r1, startColumnId: cA, endRowId: r4, endColumnId: cC },
      printTitleRows: { startRowId: r1, endRowId: r1 },
      rowBreaks: [r3],
    })

    sheet.tabColor = '#FF8800'
    const { bytes } = await workbookToXlsx(next, engine)
    const back = (await read(engine, bytes)).sheets[0]!
    expect(back.meta.tabColor?.toUpperCase()).toContain('FF8800')
    expect(back.meta.comments).toEqual([{ row: 1, column: 0, author: 'Karsten', text: 'Prüfen' }])
    expect(back.hyperlinks).toEqual([{ row: 2, column: 0, target: 'https://mindoo.de/' }])
    expect(back.autoFilter).toEqual(area)
    expect(back.autoFilterColumns).toEqual([{ colId: 1, values: ['x'] }])
    expect(back.pageSetup).toMatchObject({ orientation: 'landscape', oddHeader: '&CBericht &D' })
    expect(back.meta.printArea).toMatch(/\$A\$1:\$C\$4$/)
    expect(back.meta.printTitles).toMatch(/\$1:\$1$/)
    expect(back.rowBreaks).toEqual([2])
  })
})
