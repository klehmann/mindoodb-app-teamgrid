// Two people edit an Excel workbook (from the fidelity corpus) offline and sync, in both
// directions, through the real pipeline: import, the editor's file built from the store,
// GenOffice's save, the WASM sidecar reading it back and the save becoming MindooDB
// writes. Both devices must end up with the same workbook, both edits must survive and
// what the Excel file brought (formulas, tables, conditional formats, merges, comments,
// hyperlinks) must still be written when the merged workbook opens again. Skipped
// without native/xlsx-wasm/pkg-node.
import { readFileSync } from 'node:fs'

import { describe, expect, it } from 'vitest'

import type { WorkbookFile, WorkbookSaveRequest } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { readWorkbookModel, type WorkbookModel } from '../../e2e/fidelity/workbookModel'
import { createWorkbook, loadWorkbook, readAttachmentBase64, writeWorkbook, type LoadedWorkbook } from '../haven/store'
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
import { emptyStoredWorkbook, syncWorkbook, visualFileKey } from './sync'

const corpus = (name: string) =>
  new Uint8Array(readFileSync(new URL(`../../e2e/fidelity/corpus/${name}.xlsx`, import.meta.url)))

/** The import TeamGrid runs for an .xlsx (xlsx/sessions.ts importXlsxAsDocument). */
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
  return createWorkbook(host.connection, 'Excel', next, writes)
}

/** The editor's file for a stored workbook, and how its ids map back to the store. */
async function openSession(host: AutomergeTestHost, engine: XlsxEngine, loaded: LoadedWorkbook) {
  const { bytes, visualOrder } = await workbookToXlsx(loaded.stored.workbook, engine, {
    loadImage: (name) => readAttachmentBase64(host.connection, loaded.id, name),
  })
  const file = (await engine.open(bytes, 'en')) as WorkbookFile
  const live = liveSheets(loaded.stored.workbook)
  const storedIdByFileId = new Map(file.sheets.map((sheet, index) => [sheet.id, live[index]!.id]))
  const byFileKey = new Map<string, string>()
  for (const [fileSheetId, storedId] of storedIdByFileId) {
    const ids = visualOrder.get(storedId) ?? []
    file.visuals.filter((visual) => visual.sheetId === fileSheetId).forEach((visual, index) => {
      if (ids[index]) byFileKey.set(visualFileKey(visual), ids[index]!)
    })
  }
  return { bytes, file, storedIdByFileId, byFileKey, visualOrder }
}

type Session = Awaited<ReturnType<typeof openSession>>

/** One person opens the workbook, edits it in the editor and saves. */
async function edit(
  host: AutomergeTestHost,
  engine: XlsxEngine,
  id: string,
  build: (session: Session, request: WorkbookSaveRequest) => void,
) {
  const loaded = await loadWorkbook(host.connection, id)
  const session = await openSession(host, engine, loaded)
  const request = emptySaveRequest(session.file.sessionId)
  build(session, request)
  const names = new Map(session.file.sheets.map((sheet) => [sheet.id, sheet.name]))
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
    if (!image.bytes) image.bytes = await engine.readMediaBytes({ sessionId: file.sessionId, visualId: image.fileVisualId })
  }
  await writeWorkbook(host.connection, loaded, writes)
}

/** What the workbook writes when it opens: our parser's model of the editor's file. */
async function opened(host: AutomergeTestHost, engine: XlsxEngine, id: string): Promise<WorkbookModel> {
  const loaded = await loadWorkbook(host.connection, id)
  return readWorkbookModel((await openSession(host, engine, loaded)).bytes)
}

/** A and B edit offline, then sync both ways; returns the merged workbook (same on both). */
async function concurrently(
  host: AutomergeTestHost,
  engine: XlsxEngine,
  id: string,
  a: (session: Session, request: WorkbookSaveRequest) => void,
  b: (session: Session, request: WorkbookSaveRequest) => void,
): Promise<WorkbookModel> {
  const deviceA = host.clone('a'.repeat(32))
  const deviceB = host.clone('b'.repeat(32))
  await edit(deviceA, engine, id, a)
  await edit(deviceB, engine, id, b)
  const syncedA = deviceA.clone()
  const syncedB = deviceB.clone()
  syncedA.syncFrom(deviceB)
  syncedB.syncFrom(deviceA)
  const left = await opened(syncedA, engine, id)
  const right = await opened(syncedB, engine, id)
  expect(snapshot(left)).toEqual(snapshot(right))
  return left
}

/** A comparable form of the model (maps as sorted entries). */
function snapshot(model: WorkbookModel) {
  const entries = <K, V>(map: Map<K, V>) => [...map].sort(([a], [b]) => String(a).localeCompare(String(b)))
  return {
    names: model.definedNames,
    sheets: model.sheets.map((sheet) => ({
      ...sheet,
      cells: entries(sheet.cells),
      formats: entries(sheet.formats),
      columns: entries(sheet.columns),
      rows: entries(sheet.rows),
      features: entries(sheet.features),
    })),
  }
}

/**
 * Rows or columns inserted at the top/left, as the editor saves them: Univer moves the
 * notes with the cells and the save rewrites the sheet's notes at their new places
 * (GenOffice's collectNoteStates), so the file's comments do not stay behind.
 */
function insert(session: Session, request: WorkbookSaveRequest, axis: 'rows' | 'cols', count: number) {
  const sheet = session.file.sheets[0]!
  request.structuralOps.push({ sheetId: sheet.id, kind: axis === 'rows' ? 'insert-rows' : 'insert-cols', index: 0, count })
  const comments = (sheet as unknown as { comments?: Array<{ row: number; column: number; author: string; text: string }> }).comments ?? []
  if (comments.length === 0) return
  request.noteStates.push({
    sheetId: sheet.id,
    notes: comments.map((comment) => ({
      row: comment.row + (axis === 'rows' ? count : 0),
      column: comment.column + (axis === 'cols' ? count : 0),
      author: comment.author,
      text: comment.text,
    })),
  } as WorkbookSaveRequest['noteStates'][number])
}

const value = (request: WorkbookSaveRequest, sheetId: string, address: string, content: string | number) => {
  const column = address.charCodeAt(0) - 65
  const row = Number(address.slice(1)) - 1
  request.edits.push(
    (typeof content === 'string' && content.startsWith('=')
      ? { sheetId, row, column, writeValue: true, value: null, formula: content }
      : { sheetId, row, column, writeValue: true, value: content }) as WorkbookSaveRequest['edits'][number],
  )
}

describe.skipIf(!nodeEngineAvailable)('concurrent edits of Excel workbooks', { timeout: 60_000 }, () => {
  it('shifts a new formula with a row someone else inserted above it (IronCalc EDATE)', async () => {
    const engine = createNodeEngine()
    const host = new AutomergeTestHost()
    const id = await importExcel(host, engine, corpus('ironcalc-edate'))
    const before = await opened(host, engine, id)
    const sheet = before.sheets[0]!.name
    const merged = await concurrently(
      host,
      engine,
      id,
      (session, request) => {
        value(request, session.file.sheets[0]!.id, 'F3', '=EDATE(A3,1)')
        value(request, session.file.sheets[0]!.id, 'F4', 'A was here')
      },
      (session, request) => {
        request.structuralOps.push({ sheetId: session.file.sheets[0]!.id, kind: 'insert-rows', index: 0, count: 1 })
      },
    )
    const cells = merged.sheets.find((s) => s.name === sheet)!.cells
    expect(cells.get('F4')).toBe('f:EDATE(A4,1)')
    expect(cells.get('F5')).toBe('v:s:A was here')
    // Every formula Excel wrote moved down one row with its references.
    for (const [address, content] of before.sheets[0]!.cells) {
      if (!content.startsWith('f:')) continue
      const moved = address.replace(/\d+$/, (row) => String(Number(row) + 1))
      expect(cells.get(moved), address).toBeDefined()
    }
  })

  it('keeps an edit of a cell that moved down with rows someone else inserted above it', async () => {
    const engine = createNodeEngine()
    const host = new AutomergeTestHost()
    const id = await importExcel(host, engine, corpus('rxw-comment01'))
    const merged = await concurrently(
      host,
      engine,
      id,
      (session, request) => {
        request.structuralOps.push({ sheetId: session.file.sheets[0]!.id, kind: 'insert-rows', index: 0, count: 2 })
      },
      (session, request) => {
        value(request, session.file.sheets[0]!.id, 'A1', 'B changed this')
      },
    )
    const cells = merged.sheets[0]!.cells
    expect(cells.get('A3')).toBe('v:s:B changed this')
    expect(cells.get('A1')).toBeUndefined()
    expect([...cells.values()].filter((v) => v.includes('changed'))).toHaveLength(1)
  })

  it('keeps array formulas and moves them with a row someone else inserted (Excel array_formula01, IronCalc MMULT)', async () => {
    const engine = createNodeEngine()
    for (const name of ['rxw-array-formula01', 'ironcalc-mmult']) {
      const host = new AutomergeTestHost()
      const id = await importExcel(host, engine, corpus(name))
      const before = (await opened(host, engine, id)).sheets[0]!
      const arrays = [...before.cells].filter(([, content]) => content.startsWith('f:{=}'))
      expect(arrays.length, name).toBeGreaterThan(0)
      const merged = await concurrently(
        host,
        engine,
        id,
        (session, request) => insert(session, request, 'rows', 1),
        (session, request) => value(request, session.file.sheets[0]!.id, 'Z1', 'B was here'),
      )
      const after = merged.sheets[0]!.cells
      const down = (ref: string) => ref.replace(/([A-Z]+)(\d+)/g, (_, column: string, row: string) => `${column}${Number(row) + 1}`)
      for (const [address, content] of arrays) {
        const [, ref, formula] = /^f:\{=\}([^:]+(?::[A-Z]+\d+)?):(.*)$/.exec(content)!
        expect(after.get(down(address)), `${name} ${address}`).toBe(`f:{=}${down(ref!)}:${down(formula!)}`)
      }
      expect(after.get('Z2')).toBe('v:s:B was here')
    }
  })

  it('keeps a table when one person widens a column and the other writes into the table (Excel table01)', async () => {
    const engine = createNodeEngine()
    const host = new AutomergeTestHost()
    const id = await importExcel(host, engine, corpus('rxw-table01'))
    const merged = await concurrently(
      host,
      engine,
      id,
      (session, request) => {
        request.structuralOps.push({ sheetId: session.file.sheets[0]!.id, kind: 'set-col-size', start: 2, end: 2, size: 200 } as never)
      },
      (session, request) => {
        value(request, session.file.sheets[0]!.id, 'D3', 'Umsatz')
        value(request, session.file.sheets[0]!.id, 'C4', 42)
      },
    )
    const sheet = merged.sheets[0]!
    expect(sheet.features.get('table')).toHaveLength(1)
    expect(sheet.cells.get('D3')).toBe('v:s:Umsatz')
    expect(sheet.cells.get('C4')).toBe('v:n:42')
    expect(sheet.columns.get(3)).toBeDefined()
  })

  it('moves conditional formats, merges, comments and hyperlinks with rows and columns inserted concurrently', async () => {
    const engine = createNodeEngine()
    for (const name of ['rxw-cond-format01', 'rxw-merge-range01', 'rxw-comment01', 'rxw-hyperlink01']) {
      const host = new AutomergeTestHost()
      const id = await importExcel(host, engine, corpus(name))
      const before = (await opened(host, engine, id)).sheets[0]!
      const merged = await concurrently(
        host,
        engine,
        id,
        (session, request) => insert(session, request, 'rows', 2),
        (session, request) => insert(session, request, 'cols', 1),
      )
      const after = merged.sheets[0]!
      const shift = (ref: string) =>
        ref.replace(/([A-Z]+)(\d+)/g, (_, column: string, row: string) => `${String.fromCharCode(column.charCodeAt(0) + 1)}${Number(row) + 2}`)
      expect(after.merges, name).toEqual(before.merges.map(shift))
      for (const feature of ['conditionalFormat', 'comment', 'hyperlink']) {
        const was = before.features.get(feature) ?? []
        expect((after.features.get(feature) ?? []).length, `${name} ${feature}: ${JSON.stringify(after.features.get(feature))} was ${JSON.stringify(was)}`).toBe(was.length)
        expect(after.features.get(feature) ?? [], `${name} ${feature}`).toEqual(was.map((entry) => shift(entry.split(' ')[0]!) + entry.slice(entry.indexOf(' ') < 0 ? entry.length : entry.indexOf(' '))).sort())
      }
    }
  })
})
