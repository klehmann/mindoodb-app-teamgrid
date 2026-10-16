// Document properties (title, author, company, …): stored at import, kept by
// saves and written into every exported file. Skipped without pkg-node.
import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'

import { blankXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'

import type { WorkbookFile } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { createWorkbook, loadWorkbook } from '../haven/store'
import { AutomergeTestHost } from '../testing/automerge-host'
import { createNodeEngine, nodeEngineAvailable } from '../testing/node-engine'
import { saveWorkbookInBrowser } from '../xlsx/browser-save'
import { applyDocProps, extractDocProps, type DocumentProperties } from '../xlsx/doc-props'
import type { XlsxEngine } from '../xlsx/engine'
import { emptySaveRequest, workbookToXlsx } from './export'
import { readWholeWorkbook } from './sidecar-read'
import { emptyStoredWorkbook, syncWorkbook } from './sync'

const SHEET = 'sheet-1'
const PROPERTIES: DocumentProperties = {
  title: 'Umsatz <Q3> & Planung',
  subject: 'Vertrieb',
  creator: 'Jürgen Müller',
  keywords: 'umsatz; planung',
  description: 'Zeile 1\nZeile 2',
  lastModifiedBy: 'Anna',
  category: 'Berichte',
  created: '2024-05-01T08:00:00Z',
  company: 'Mindoo GmbH',
  manager: 'Chef',
}

async function read(engine: XlsxEngine, bytes: Uint8Array) {
  const file = (await engine.open(bytes, 'de')) as WorkbookFile
  const result = await readWholeWorkbook(engine, file)
  await engine.close(file.sessionId)
  return result
}

describe('document properties', () => {
  it('read back what was written, with a fresh modified date', async () => {
    const blank = new Uint8Array(await blankXlsxBuffer('Daten'))
    const bytes = await applyDocProps(blank, PROPERTIES, new Date('2026-10-16T09:00:00.123Z'))
    expect(await extractDocProps(bytes)).toEqual(PROPERTIES)
    const zip = await JSZip.loadAsync(bytes)
    expect(await zip.file('docProps/core.xml')!.async('text')).toContain('<dcterms:modified xsi:type="dcterms:W3CDTF">2026-10-16T09:00:00Z</dcterms:modified>')
    expect(await zip.file('[Content_Types].xml')!.async('text')).toContain('core-properties+xml')
    expect(await zip.file('_rels/.rels')!.async('text')).toContain('docProps/app.xml')
    // Written again (an export of an export): one part and one relationship each.
    const again = await JSZip.loadAsync(await applyDocProps(bytes, { title: 'Neu' }, new Date()))
    expect((await again.file('_rels/.rels')!.async('text')).match(/docProps\/core\.xml/g)).toHaveLength(1)
    expect(await extractDocProps(await again.generateAsync({ type: 'uint8array' }))).toMatchObject({ title: 'Neu' })
  })

  it.skipIf(!nodeEngineAvailable)('are stored at import, kept by a save and exported', async () => {
    const engine = createNodeEngine()
    const host = new AutomergeTestHost()
    const request = emptySaveRequest('source')
    request.edits.push({ sheetId: SHEET, row: 0, column: 0, writeValue: true, value: 'Wert' } as never)
    const blank = new Uint8Array(await blankXlsxBuffer('Daten'))
    const plain = (await saveWorkbookInBrowser(blank, new Map([[SHEET, 'Daten']]), request)).bytes
    const bytes = await applyDocProps(plain, PROPERTIES, new Date())

    const { next, writes } = syncWorkbook({
      stored: emptyStoredWorkbook(),
      storedIdByFileId: new Map(),
      read: await read(engine, bytes),
      properties: (await extractDocProps(bytes)) ?? null,
    })
    const id = await createWorkbook(host.connection, 'Eigenschaften', next, writes)
    const loaded = await loadWorkbook(host.connection, id)
    expect(loaded.stored.workbook.properties).toEqual(PROPERTIES)

    const back = (await workbookToXlsx(loaded.stored.workbook, engine)).bytes
    expect(await extractDocProps(back)).toEqual(PROPERTIES)

    // A save (no properties passed) keeps the stored ones.
    const saved = syncWorkbook({ stored: loaded.stored, storedIdByFileId: new Map(), read: await read(engine, back) })
    expect(saved.next.properties).toEqual(PROPERTIES)
  })
})
