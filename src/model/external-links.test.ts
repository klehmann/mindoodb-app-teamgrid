// Links to other workbooks: formulas keep their `[n]` references verbatim and
// the externalLink parts are stored and written back in order.
import { describe, expect, it } from 'vitest'

import { blankXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'

import type { WorkbookFile } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { createWorkbook, loadWorkbook, readAttachmentBase64 } from '../haven/store'
import { AutomergeTestHost } from '../testing/automerge-host'
import { createNodeEngine, nodeEngineAvailable } from '../testing/node-engine'
import { saveWorkbookInBrowser } from '../xlsx/browser-save'
import type { XlsxEngine } from '../xlsx/engine'
import { applyExternalLinks, extractExternalLinks } from '../xlsx/external-links'
import { emptySaveRequest, workbookToXlsx } from './export'
import { createAxesLookup, parseFormula, renderFormula } from './formula-refs'
import { readWholeWorkbook } from './sidecar-read'
import { emptyStoredWorkbook, syncWorkbook } from './sync'

const SHEET = 'sheet-1'
const LINK = {
  xml:
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<externalBook r:id="rId1"><sheetNames><sheetName val="Kurse"/></sheetNames><sheetDataSet><sheetData sheetId="0">' +
    '<row r="2"><cell r="B2"><v>42</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>',
  rels:
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="Kurse.xlsx" TargetMode="External"/></Relationships>',
}

async function read(engine: XlsxEngine, bytes: Uint8Array) {
  const file = (await engine.open(bytes, 'de')) as WorkbookFile
  const result = await readWholeWorkbook(engine, file)
  await engine.close(file.sessionId)
  return result
}

describe('external links', () => {
  it('keeps references into another workbook verbatim', () => {
    const home = { id: 's1', name: 'Kurse', rowIds: ['r1', 'r2'], columnIds: ['c1', 'c2'] }
    const lookup = createAxesLookup([home])
    const formula = parseFormula('[1]Kurse!B2+B2+SUM([2]Daten!$A$1:$A$9)+[1]!Zins', home, lookup)
    expect(formula.r).toHaveLength(1)
    expect(renderFormula(formula, home, lookup)).toBe('[1]Kurse!B2+B2+SUM([2]Daten!$A$1:$A$9)+[1]!Zins')
  })

  it.skipIf(!nodeEngineAvailable)('stores the links and writes them back', async () => {
    const engine = createNodeEngine()
    const host = new AutomergeTestHost()
    const request = emptySaveRequest('source')
    request.edits.push({ sheetId: SHEET, row: 0, column: 0, writeValue: true, value: 'Kurs' } as never)
    request.edits.push({ sheetId: SHEET, row: 0, column: 1, writeValue: true, value: null, formula: '=[1]Kurse!B2*2' } as never)
    request.formulaValues.push({ sheetId: SHEET, row: 0, column: 1, value: 84 })
    const blank = new Uint8Array(await blankXlsxBuffer('Kurse'))
    const saved = (await saveWorkbookInBrowser(blank, new Map([[SHEET, 'Kurse']]), request)).bytes
    const bytes = await applyExternalLinks(saved, [LINK])
    expect(await extractExternalLinks(bytes)).toEqual([LINK])
    const source = await read(engine, bytes)
    expect(source.sheets[0]!.cells.find((cell) => cell.column === 1)?.formula).toBe('=[1]Kurse!B2*2')

    const { next, writes } = syncWorkbook({
      stored: emptyStoredWorkbook(),
      storedIdByFileId: new Map(),
      read: source,
      externalLinks: await extractExternalLinks(bytes),
    })
    const id = await createWorkbook(host.connection, 'Kurse', next, writes)
    const loaded = await loadWorkbook(host.connection, id)
    expect(loaded.stored.workbook.externalLinks?.attachment).toMatch(/^links-/)
    const loadImage = (attachment: string) => readAttachmentBase64(host.connection, id, attachment)

    const back = (await workbookToXlsx(loaded.stored.workbook, engine, { loadImage })).bytes
    expect(await extractExternalLinks(back)).toEqual([LINK])
    const cell = (await read(engine, back)).sheets[0]!.cells.find((entry) => entry.column === 1)
    expect(cell).toMatchObject({ formula: '=[1]Kurse!B2*2', value: 84 })
  })
})
