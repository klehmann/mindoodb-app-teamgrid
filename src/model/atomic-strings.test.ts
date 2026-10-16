// TeamGrid has no collaborative text: every string it stores (ids, enums,
// style values, cell text, formulas, names) is an atomic value, so a
// concurrent change replaces it as a whole instead of interleaving its
// characters. This checks the stored Automerge documents (top document and
// gridChunk documents) after an Excel import, an edit saved through the real
// pipeline and a property change. Skipped without native/xlsx-wasm/pkg-node.
import { readFileSync } from 'node:fs'

import { describe, expect, it } from 'vitest'

import type { WorkbookFile } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { writeProperties } from '../haven/connection'
import { createWorkbook, loadWorkbook, readAttachmentBase64, writeWorkbook } from '../haven/store'
import { AutomergeTestHost } from '../testing/automerge-host'
import { createNodeEngine, nodeEngineAvailable } from '../testing/node-engine'
import { saveWorkbookInBrowser } from '../xlsx/browser-save'
import { extractCharts } from '../xlsx/chart-parts'
import type { XlsxEngine } from '../xlsx/engine'
import { extractExternalLinks } from '../xlsx/external-links'
import { extractMacros } from '../xlsx/macros'
import { extractPivots } from '../xlsx/pivot-parts'
import { extractSlicers } from '../xlsx/slicer-parts'
import { extractTableParts } from '../xlsx/table-parts'
import { emptySaveRequest, workbookToXlsx } from './export'
import { loadRuleConverters } from './rules'
import { liveSheets } from './schema'
import { readWholeWorkbook } from './sidecar-read'
import { emptyStoredWorkbook, syncWorkbook } from './sync'

/** Workbooks that bring most of what the schema stores. */
const CORPUS = [
  'rxw-autofilter02',
  'rxw-chart-combined01',
  'rxw-comment01',
  'rxw-cond-format02',
  'rxw-data-validation02',
  'rxw-defined-name02',
  'rxw-dynamic-array01',
  'rxw-hyperlink02',
  'rxw-image01',
  'rxw-merge-range02',
  'rxw-outline02',
  'rxw-print-options02',
  'rxw-properties02',
  'rxw-protect02',
  'rxw-repeat02',
  'rxw-table02',
  'rxw-tab-color01',
  'rxw-textbox01',
]

const corpus = (name: string) =>
  new Uint8Array(readFileSync(new URL(`../../e2e/fidelity/corpus/${name}.xlsx`, import.meta.url)))

async function importExcel(host: AutomergeTestHost, engine: XlsxEngine, bytes: Uint8Array): Promise<string> {
  await loadRuleConverters()
  const file = (await engine.open(bytes, 'en')) as WorkbookFile
  const { next, writes } = syncWorkbook({
    stored: emptyStoredWorkbook(),
    storedIdByFileId: new Map(),
    read: await readWholeWorkbook(engine, file),
    pivots: await extractPivots(bytes),
    tableParts: await extractTableParts(bytes),
    charts: await extractCharts(bytes),
    externalLinks: await extractExternalLinks(bytes),
    slicers: (await extractSlicers(bytes)) ?? null,
    macros: (await extractMacros(bytes)) ?? null,
  })
  for (const image of writes.newImages) {
    if (!image.bytes) image.bytes = await engine.readMediaBytes({ sessionId: file.sessionId, visualId: image.fileVisualId })
  }
  return createWorkbook(host.connection, 'Excel', next, writes, { tags: ['import'] })
}

/** Opens the workbook in the editor, inserts a row, types text and a formula and saves. */
async function editAndSave(host: AutomergeTestHost, engine: XlsxEngine, id: string) {
  const loaded = await loadWorkbook(host.connection, id)
  const { bytes } = await workbookToXlsx(loaded.stored.workbook, engine, {
    loadImage: (name) => readAttachmentBase64(host.connection, loaded.id, name),
  })
  const file = (await engine.open(bytes, 'en')) as WorkbookFile
  const live = liveSheets(loaded.stored.workbook)
  const storedIdByFileId = new Map(file.sheets.map((sheet, index) => [sheet.id, live[index]!.id]))
  const request = emptySaveRequest(file.sessionId)
  const sheetId = file.sheets[0]!.id
  request.structuralOps.push({ sheetId, kind: 'insert-rows', index: 0, count: 1 })
  request.edits.push(
    { sheetId, row: 0, column: 0, writeValue: true, value: 'typed text' },
    { sheetId, row: 0, column: 1, writeValue: true, value: null, formula: 'A1&"!"' },
  )
  const names = new Map(file.sheets.map((sheet) => [sheet.id, sheet.name]))
  const saved = (await saveWorkbookInBrowser(bytes, names, request)).bytes
  const reopened = (await engine.open(saved, 'en')) as WorkbookFile
  const { writes } = syncWorkbook({
    stored: loaded.stored,
    storedIdByFileId,
    request,
    read: await readWholeWorkbook(engine, reopened),
  })
  for (const image of writes.newImages) {
    if (!image.bytes) image.bytes = await engine.readMediaBytes({ sessionId: reopened.sessionId, visualId: image.fileVisualId })
  }
  await writeWorkbook(host.connection, loaded, writes)
}

function textFields(host: AutomergeTestHost) {
  return Object.fromEntries(
    host
      .documentIds()
      .map((id) => [id, host.textPaths(id)] as const)
      .filter(([, paths]) => paths.length > 0),
  )
}

describe.skipIf(!nodeEngineAvailable)('stored strings', { timeout: 120_000 }, () => {
  it.each(CORPUS)('are all atomic after importing, editing and retitling %s', async (name) => {
    const engine = createNodeEngine()
    const host = new AutomergeTestHost()
    const id = await importExcel(host, engine, corpus(name))
    expect(host.documentIds().length).toBeGreaterThan(1)
    expect(textFields(host)).toEqual({})

    await editAndSave(host, engine, id)
    await writeProperties(host.connection, id, { subject: 'Retitled', tags: ['a', 'b'], istemplate: true })
    expect(textFields(host)).toEqual({})

    // Reads still see plain strings.
    const loaded = await loadWorkbook(host.connection, id)
    expect(loaded.subject).toBe('Retitled')
    const sheet = liveSheets(loaded.stored.workbook)[0]!
    expect(Object.values(sheet.cellsById).some((cell) => cell.value === 'typed text')).toBe(true)
  })

  it('reports a string stored as collaborative text (the check itself works)', async () => {
    const host = new AutomergeTestHost()
    await host.connection.database.documents.create({ id: 'plain', set: { form: 'x', nested: { list: ['y'] } } })
    expect(host.textPaths('plain').sort()).toEqual(['form', 'nested.list[0]'])
  })
})
