// VBA macros of an .xlsm: the project and its code names are stored and put
// back, so an export is again an .xlsm. Skipped without pkg-node.
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'

import { blankXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'

import type { WorkbookFile } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { createWorkbook, loadWorkbook, readAttachmentBase64 } from '../haven/store'
import { AutomergeTestHost } from '../testing/automerge-host'
import { createNodeEngine, nodeEngineAvailable } from '../testing/node-engine'
import { saveWorkbookInBrowser } from '../xlsx/browser-save'
import type { XlsxEngine } from '../xlsx/engine'
import { applyMacros, extractMacros, hasMacros, type MacroParts } from '../xlsx/macros'
import { emptySaveRequest, workbookToXlsx } from './export'
import { readWholeWorkbook } from './sidecar-read'
import { emptyStoredWorkbook, syncWorkbook } from './sync'

const SHEET = 'sheet-1'
const MACROS: MacroParts = {
  project: btoa('VBA-PROJEKT-BYTES'),
  signatures: [
    {
      type: 'http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature',
      contentType: 'application/vnd.ms-office.vbaProjectSignature',
      data: btoa('SIGNATUR'),
    },
  ],
  workbookCodeName: 'DieseArbeitsmappe',
  sheetCodeNames: { Daten: 'Tabelle1' },
}

async function read(engine: XlsxEngine, bytes: Uint8Array) {
  const file = (await engine.open(bytes, 'de')) as WorkbookFile
  const result = await readWholeWorkbook(engine, file)
  await engine.close(file.sessionId)
  return result
}

describe('macros', () => {
  it.skipIf(!nodeEngineAvailable)('stores the VBA project and writes an .xlsm back', async () => {
    const engine = createNodeEngine()
    const host = new AutomergeTestHost()
    const request = emptySaveRequest('source')
    request.edits.push({ sheetId: SHEET, row: 0, column: 0, writeValue: true, value: 'Makro' } as never)
    const blank = new Uint8Array(await blankXlsxBuffer('Daten'))
    const plain = (await saveWorkbookInBrowser(blank, new Map([[SHEET, 'Daten']]), request)).bytes
    expect(hasMacros(plain)).toBe(false)
    const bytes = await applyMacros(plain, MACROS)
    expect(hasMacros(bytes)).toBe(true)
    expect(await extractMacros(bytes)).toEqual(MACROS)

    const { next, writes } = syncWorkbook({
      stored: emptyStoredWorkbook(),
      storedIdByFileId: new Map(),
      read: await read(engine, bytes),
      macros: (await extractMacros(bytes)) ?? null,
    })
    const id = await createWorkbook(host.connection, 'Makros', next, writes)
    const loaded = await loadWorkbook(host.connection, id)
    const loadImage = (attachment: string) => readAttachmentBase64(host.connection, id, attachment)
    const back = (await workbookToXlsx(loaded.stored.workbook, engine, { loadImage })).bytes
    expect(await extractMacros(back)).toEqual(MACROS)
    const types = await (await JSZip.loadAsync(back)).file('[Content_Types].xml')!.async('text')
    expect(types).toContain('application/vnd.ms-excel.sheet.macroEnabled.main+xml')

    // The editor's own save keeps them.
    const file = (await engine.open(back, 'de')) as WorkbookFile
    const edit = emptySaveRequest(file.sessionId)
    edit.edits.push({ sheetId: file.sheets[0]!.id, row: 1, column: 0, writeValue: true, value: 'neu' } as never)
    const edited = (await saveWorkbookInBrowser(back, new Map([[file.sheets[0]!.id, 'Daten']]), edit)).bytes
    await engine.close(file.sessionId)
    expect(await extractMacros(edited)).toEqual(MACROS)
  })
})
