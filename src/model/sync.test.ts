// Concurrent editing of stored workbooks: two replicas save against the same
// base (as if both were offline), in both orders, and every replica must end
// up with the same rows, in the same order, with nothing lost.
import { describe, expect, it } from 'vitest'

import type { WorkbookSaveRequest } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { changedSinceLoad, createWorkbook, loadWorkbook, writeWorkbook, type LoadedWorkbook } from '../haven/store'
import { listRevisions, loadWorkbookAt } from '../haven/history'
import { AutomergeTestHost } from '../testing/automerge-host'
import { createAxesLookup, renderFormula } from './formula-refs'
import { liveIds } from './schema'
import type { SheetMetadata, SidecarWorkbook } from './sidecar-read'
import { sheetAxes, workbookAxes } from './export'
import { CHUNK_SOFT_LIMIT, emptyStoredWorkbook, syncWorkbook } from './sync'

const FILE_SHEET = 'f1'
type Grid = (string | number | null)[][]

/** What the editor's file holds: one sheet "Data" with these rows (formulas start with "="). */
function read(rows: Grid): SidecarWorkbook {
  const columnCount = Math.max(1, ...rows.map((row) => row.length))
  const meta = {
    id: FILE_SHEET,
    name: 'Data',
    rowCount: Math.max(1, rows.length),
    columnCount,
    columnWidths: [],
    defaultRowHeight: null,
    defaultColumnWidth: null,
    freeze: null,
    hidden: false,
    tabColor: null,
    showGridLines: true,
    tables: [],
  } as unknown as SheetMetadata
  const cells: SidecarWorkbook['sheets'][number]['cells'] = []
  rows.forEach((row, rowIndex) =>
    row.forEach((value, column) => {
      if (value === null) return
      if (typeof value === 'string' && value.startsWith('=')) cells.push({ row: rowIndex, column, value: null, formula: value })
      else cells.push({ row: rowIndex, column, value })
    }),
  )
  return {
    file: { sessionId: 'x', styles: [], visuals: [] } as unknown as SidecarWorkbook['file'],
    sheets: [{ meta, cells, rows: [], merges: [] }],
  }
}

function request(structuralOps: object[] = []): WorkbookSaveRequest {
  return { sessionId: 'x', sheetOps: [], structuralOps } as unknown as WorkbookSaveRequest
}

async function importGrid(host: AutomergeTestHost, rows: Grid) {
  const { next, writes } = syncWorkbook({ stored: emptyStoredWorkbook(), storedIdByFileId: new Map(), read: read(rows) })
  return createWorkbook(host.connection, 'Test', next, writes)
}

async function save(host: AutomergeTestHost, loaded: LoadedWorkbook, rows: Grid, ops: object[] = []) {
  const sheetId = loaded.stored.workbook.worksheetOrder[0]!
  const { writes } = syncWorkbook({
    stored: loaded.stored,
    storedIdByFileId: new Map([[FILE_SHEET, sheetId]]),
    request: request(ops),
    read: read(rows),
  })
  await writeWorkbook(host.connection, loaded, writes)
}

/** The sheet as the editor would show it: values (and A1 formulas) by visible row. */
async function visible(host: AutomergeTestHost, id: string): Promise<Grid> {
  const { stored } = await loadWorkbook(host.connection, id)
  const sheet = Object.values(stored.workbook.worksheetsById)[0]!
  const axes = sheetAxes(sheet)
  const lookup = createAxesLookup(workbookAxes(stored.workbook))
  return liveIds(sheet.rowOrder, sheet.rowsById).map((rowId) =>
    axes.columnIds.map((columnId) => {
      const cell = sheet.cellsById[`${rowId}:${columnId}`]
      if (cell?.formula) return renderFormula(cell.formula, axes, lookup)
      return (cell?.value ?? null) as string | number | null
    }),
  )
}

/**
 * Two devices start from the same state, each saves its own edit offline,
 * then they sync, in both directions. Both must show the same sheet.
 */
async function concurrently(
  host: AutomergeTestHost,
  id: string,
  a: (host: AutomergeTestHost, loaded: LoadedWorkbook) => Promise<void>,
  b: (host: AutomergeTestHost, loaded: LoadedWorkbook) => Promise<void>,
) {
  const deviceA = host.clone('a'.repeat(32))
  const deviceB = host.clone('b'.repeat(32))
  await a(deviceA, await loadWorkbook(deviceA.connection, id))
  await b(deviceB, await loadWorkbook(deviceB.connection, id))
  const syncedA = deviceA.clone()
  const syncedB = deviceB.clone()
  syncedA.syncFrom(deviceB)
  syncedB.syncFrom(deviceA)
  const [left, right] = [await visible(syncedA, id), await visible(syncedB, id)]
  expect(left).toEqual(right)
  return left
}

describe('stored workbook', () => {
  it('round-trips values and formulas', async () => {
    const host = new AutomergeTestHost()
    const id = await importGrid(host, [
      ['Region', 'Q1'],
      ['North', 10],
      ['South', 20],
      ['Total', '=SUM(B2:B3)'],
    ])
    expect(await visible(host, id)).toEqual([
      ['Region', 'Q1'],
      ['North', 10],
      ['South', 20],
      ['Total', '=SUM(B2:B3)'],
    ])
  })

  it('keeps both rows when two people insert a row at the same place', async () => {
    const host = new AutomergeTestHost()
    const id = await importGrid(host, [
      ['a', 1],
      ['b', 2],
      ['sum', '=SUM(B1:B2)'],
    ])
    const result = await concurrently(
      host,
      id,
      (h, loaded) =>
        save(h, loaded, [['a', 1], ['A', 5], ['b', 2], ['sum', '=SUM(B1:B3)']], [
          { sheetId: FILE_SHEET, kind: 'insert-rows', index: 1, count: 1 },
        ]),
      (h, loaded) =>
        save(h, loaded, [['a', 1], ['B', 7], ['b', 2], ['sum', '=SUM(B1:B3)']], [
          { sheetId: FILE_SHEET, kind: 'insert-rows', index: 1, count: 1 },
        ]),
    )
    expect(result).toHaveLength(5)
    expect(result.map((row) => row[0])).toEqual(expect.arrayContaining(['a', 'A', 'B', 'b', 'sum']))
    expect(result[0]![0]).toBe('a')
    expect(result[3]![0]).toBe('b')
    // The range spans both inserted rows on every replica.
    expect(result[4]![1]).toBe('=SUM(B1:B4)')
  })

  it('puts two people typing into the next empty row into the same row', async () => {
    const host = new AutomergeTestHost()
    const id = await importGrid(host, [['a', 1]])
    const result = await concurrently(
      host,
      id,
      (h, loaded) => save(h, loaded, [['a', 1], ['from A', null]]),
      (h, loaded) => save(h, loaded, [['a', 1], [null, 'from B']]),
    )
    expect(result).toEqual([
      ['a', 1],
      ['from A', 'from B'],
    ])
  })

  it('merges rows appended past a full chunk into one new chunk', async () => {
    const host = new AutomergeTestHost()
    const base: Grid = Array.from({ length: CHUNK_SOFT_LIMIT }, (_, index) => [`r${index}`])
    const id = await importGrid(host, base)
    const result = await concurrently(
      host,
      id,
      (h, loaded) => save(h, loaded, [...base, ['A1'], ['A2']]),
      (h, loaded) => save(h, loaded, [...base, ['B1']]),
    )
    // Same derived ids: A's and B's first appended row are one row, so the
    // same cell holds one of the two values (on both devices the same one).
    expect(result).toHaveLength(CHUNK_SOFT_LIMIT + 2)
    expect(['A1', 'B1']).toContain(result[CHUNK_SOFT_LIMIT]![0])
    expect(result[CHUNK_SOFT_LIMIT + 1]![0]).toBe('A2')
  })

  it('keeps an edit to a row that another person deleted out of view, without losing the deletion', async () => {
    const host = new AutomergeTestHost()
    const id = await importGrid(host, [['a'], ['b'], ['c']])
    const result = await concurrently(
      host,
      id,
      (h, loaded) => save(h, loaded, [['a'], ['c']], [{ sheetId: FILE_SHEET, kind: 'remove-rows', index: 1, count: 1 }]),
      (h, loaded) => save(h, loaded, [['a'], ['b!'], ['c']]),
    )
    expect(result).toEqual([['a'], ['c']])
  })

  it('writes only the chunk that changed', async () => {
    const host = new AutomergeTestHost()
    const base: Grid = Array.from({ length: CHUNK_SOFT_LIMIT * 2 + 10 }, (_, index) => [`r${index}`])
    const id = await importGrid(host, base)
    const loaded = await loadWorkbook(host.connection, id)
    const edited = base.map((row, index) => (index === CHUNK_SOFT_LIMIT + 5 ? ['changed'] : row))
    const sheetId = loaded.stored.workbook.worksheetOrder[0]!
    const { writes } = syncWorkbook({
      stored: loaded.stored,
      storedIdByFileId: new Map([[FILE_SHEET, sheetId]]),
      request: request(),
      read: read(edited),
    })
    expect(writes.chunks.size).toBe(1)
    expect(writes.top.set).toHaveLength(0)
    expect(writes.createdChunks.size).toBe(0)
  })

  it('notices when someone else changed one of its row blocks', async () => {
    const host = new AutomergeTestHost()
    const id = await importGrid(host, [['a'], ['b']])
    const watching = await loadWorkbook(host.connection, id)
    expect(await changedSinceLoad(host.connection, watching)).toBe(false)
    const other = await loadWorkbook(host.connection, id)
    host.actor = 'c'.repeat(32)
    await save(host, other, [['a'], ['b, edited elsewhere']])
    expect(await changedSinceLoad(host.connection, watching)).toBe(true)
    expect(await changedSinceLoad(host.connection, watching)).toBe(false)
  })
})

describe('revisions', () => {
  /** Values by visible row of the workbook's first sheet. */
  function values(workbook: import('./schema').Workbook): Grid {
    const sheet = Object.values(workbook.worksheetsById)[0]!
    const axes = sheetAxes(sheet)
    return liveIds(sheet.rowOrder, sheet.rowsById).map((rowId) =>
      axes.columnIds.map((columnId) => (sheet.cellsById[`${rowId}:${columnId}`]?.value ?? null) as string | number | null),
    )
  }

  it('lists one revision per save and reads the workbook as it was then', async () => {
    const host = new AutomergeTestHost()
    const id = await importGrid(host, [['a', 1], ['b', 2]])
    host.now += 60_000
    await save(host, await loadWorkbook(host.connection, id), [['a', 10], ['b', 2]])
    host.now += 60_000
    host.actor = 'bb'.repeat(16)
    await save(host, await loadWorkbook(host.connection, id), [['a', 10], ['b', 2], ['c', 3]])

    const revisions = await listRevisions(host.connection, await loadWorkbook(host.connection, id))
    expect(revisions.map((revision) => revision.current)).toEqual([true, false, false])
    expect(revisions[0]!.timestamp).toBeGreaterThan(revisions[1]!.timestamp)

    const at = async (index: number) => values((await loadWorkbookAt(host.connection, id, revisions[index]!.timestamp)).workbook)
    expect(await at(2)).toEqual([['a', 1], ['b', 2]])
    expect(await at(1)).toEqual([['a', 10], ['b', 2]])
    expect(await at(0)).toEqual([['a', 10], ['b', 2], ['c', 3]])
  })
})
