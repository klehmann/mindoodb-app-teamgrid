// Two people edit the same workbook offline and sync, in both directions.
// Each scenario checks that both devices end up with the same workbook and
// that nothing either of them did is lost where both edits can coexist.
import { describe, expect, it } from 'vitest'

import type { WorkbookSaveRequest } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { createWorkbook, loadWorkbook, writeWorkbook, type LoadedWorkbook } from '../haven/store'
import { AutomergeTestHost } from '../testing/automerge-host'
import { sheetAxes } from './export'
import { liveIds, liveSheets, type CellStyle, type Workbook } from './schema'
import type { SheetMetadata, SidecarWorkbook } from './sidecar-read'
import { styleOf } from './styles'
import { emptyStoredWorkbook, syncWorkbook } from './sync'

/** A cell: a value, or a value with formatting. */
type Value = string | number | null
type Fmt = { v: Value; bold?: boolean; italic?: boolean; fill?: string }
type Grid = (Value | Fmt)[][]
interface SheetSpec {
  name: string
  rows: Grid
  merges?: [number, number, number, number][]
}

const BASE_STYLE: CellStyle = {
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  wrapText: false,
  diagonalUp: false,
  diagonalDown: false,
}

/** The editor's file: these sheets, file sheet ids f0, f1, … in order. */
function read(sheets: SheetSpec[]): SidecarWorkbook {
  const styles: CellStyle[] = [BASE_STYLE]
  const styleIndex = (cell: Fmt) => {
    const style: CellStyle = { ...BASE_STYLE, bold: !!cell.bold, italic: !!cell.italic }
    if (cell.fill) style.fillColor = cell.fill
    const key = JSON.stringify(style)
    const found = styles.findIndex((entry) => JSON.stringify(entry) === key)
    if (found >= 0) return found
    styles.push(style)
    return styles.length - 1
  }
  return {
    file: { sessionId: 'x', styles, visuals: [] } as unknown as SidecarWorkbook['file'],
    sheets: sheets.map((sheet, index) => {
      const cells: SidecarWorkbook['sheets'][number]['cells'] = []
      sheet.rows.forEach((row, rowIndex) =>
        row.forEach((entry, column) => {
          const cell: Fmt = entry !== null && typeof entry === 'object' ? entry : { v: entry }
          const formatted = !!(cell.bold || cell.italic || cell.fill)
          if (cell.v === null && !formatted) return
          cells.push({
            row: rowIndex,
            column,
            value: cell.v,
            ...(formatted ? { styleIndex: styleIndex(cell) } : {}),
          })
        }),
      )
      const meta = {
        id: `f${index}`,
        name: sheet.name,
        rowCount: Math.max(1, sheet.rows.length),
        columnCount: Math.max(1, ...sheet.rows.map((row) => row.length)),
        columnWidths: [],
        freeze: null,
        hidden: false,
        tabColor: null,
        showGridLines: true,
        tables: [],
      } as unknown as SheetMetadata
      const merges = (sheet.merges ?? []).map(([startRow, startColumn, endRow, endColumn]) => ({
        startRow,
        startColumn,
        endRow,
        endColumn,
      }))
      return { meta, cells, rows: [], merges }
    }),
  }
}

async function importSheets(host: AutomergeTestHost, sheets: SheetSpec[]) {
  const { next, writes } = syncWorkbook({ stored: emptyStoredWorkbook(), storedIdByFileId: new Map(), read: read(sheets) })
  return createWorkbook(host.connection, 'Test', next, writes)
}

/** The editor opened `loaded` (file ids f0, f1, … for its live sheets) and saved `after`. */
async function save(
  host: AutomergeTestHost,
  loaded: LoadedWorkbook,
  after: SheetSpec[],
  request: Partial<WorkbookSaveRequest> = {},
) {
  const live = liveSheets(loaded.stored.workbook)
  const { writes } = syncWorkbook({
    stored: loaded.stored,
    storedIdByFileId: new Map(live.map((sheet, index) => [`f${index}`, sheet.id])),
    request: { sessionId: 'x', sheetOps: [], structuralOps: [], ...request } as WorkbookSaveRequest,
    read: read(after),
  })
  await writeWorkbook(host.connection, loaded, writes)
}

/** What the editor would show: per sheet its name, values and formatting by position. */
function view(workbook: Workbook) {
  return liveSheets(workbook).map((sheet) => {
    const axes = sheetAxes(sheet)
    const rows = liveIds(sheet.rowOrder, sheet.rowsById).map((rowId) =>
      axes.columnIds.map((columnId) => {
        const cell = sheet.cellsById[`${rowId}:${columnId}`]
        const style = styleOf(cell)
        const fmt = [style?.bold ? 'b' : '', style?.italic ? 'i' : '', style?.fillColor ? `fill${style.fillColor}` : '']
          .filter(Boolean)
          .join('+')
        const value = (cell?.value ?? null) as Value
        return fmt ? `${value ?? ''}[${fmt}]` : value
      }),
    )
    const rowIndex = new Map(axes.rowIds.map((id, index) => [id, index]))
    const columnIndex = new Map(axes.columnIds.map((id, index) => [id, index]))
    const merges = Object.values(sheet.mergesById)
      .map((merge) => [
        rowIndex.get(merge.startRowId),
        columnIndex.get(merge.startColumnId),
        rowIndex.get(merge.endRowId),
        columnIndex.get(merge.endColumnId),
      ])
      .sort()
    return { name: sheet.name, rows, merges }
  })
}

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
  const left = view((await loadWorkbook(syncedA.connection, id)).stored.workbook)
  const right = view((await loadWorkbook(syncedB.connection, id)).stored.workbook)
  expect(left).toEqual(right)
  lastSynced = syncedA
  return left
}

/** Device A after the last `concurrently`, with B's changes synced in. */
let lastSynced: AutomergeTestHost

const DATA: SheetSpec = { name: 'Daten', rows: [['a', 1], ['b', 2]] }

describe('concurrent edits', () => {
  it('keeps two different sheets added at the same time', async () => {
    const host = new AutomergeTestHost()
    const id = await importSheets(host, [DATA])
    const result = await concurrently(
      host,
      id,
      (h, l) => save(h, l, [DATA, { name: 'Plan', rows: [['p']] }], { sheetOps: [{ kind: 'add-sheet', sheetId: 'u1', name: 'Plan' }], sheetOrder: ['f0', 'u1'] }),
      (h, l) => save(h, l, [DATA, { name: 'Budget', rows: [['x']] }], { sheetOps: [{ kind: 'add-sheet', sheetId: 'u1', name: 'Budget' }], sheetOrder: ['f0', 'u1'] }),
    )
    expect(result.map((sheet) => sheet.name).sort()).toEqual(['Budget', 'Daten', 'Plan'])
    expect(result[0]!.name).toBe('Daten')
  })

  it('keeps both sheets when both add one with the same name, and gives them distinct names', async () => {
    const host = new AutomergeTestHost()
    const id = await importSheets(host, [DATA])
    const result = await concurrently(
      host,
      id,
      (h, l) => save(h, l, [DATA, { name: 'Tabelle2', rows: [['from A']] }], { sheetOps: [{ kind: 'add-sheet', sheetId: 'u1', name: 'Tabelle2' }], sheetOrder: ['f0', 'u1'] }),
      (h, l) => save(h, l, [DATA, { name: 'Tabelle2', rows: [['from B']] }], { sheetOps: [{ kind: 'add-sheet', sheetId: 'u1', name: 'Tabelle2' }], sheetOrder: ['f0', 'u1'] }),
    )
    expect(result).toHaveLength(3)
    expect(new Set(result.map((sheet) => sheet.name)).size).toBe(3)
    expect(result.flatMap((sheet) => sheet.rows.flat()).filter((value) => String(value).startsWith('from'))).toEqual(
      expect.arrayContaining(['from A', 'from B']),
    )
    // The next save writes the distinct name, so older replicas see it too.
    const loaded = await loadWorkbook(lastSynced.connection, id)
    await save(lastSynced, loaded, result.map((sheet) => ({ name: sheet.name, rows: sheet.rows as Grid })))
    const top = (await loadWorkbook(lastSynced.connection, id)).stored.top
    expect(new Set(Object.values(top.worksheetsById).map((sheet) => sheet.name)).size).toBe(3)
  })

  it('keeps edits to cells of a sheet another person renamed', async () => {
    const host = new AutomergeTestHost()
    const id = await importSheets(host, [DATA])
    const result = await concurrently(
      host,
      id,
      (h, l) => save(h, l, [{ ...DATA, name: 'Umsatz' }], { sheetOps: [{ kind: 'rename-sheet', sheetId: 'f0', newName: 'Umsatz' }], sheetOrder: ['f0'] }),
      (h, l) => save(h, l, [{ name: 'Daten', rows: [['a', 10], ['b', 2]] }]),
    )
    expect(result).toEqual([{ name: 'Umsatz', rows: [['a', 10], ['b', 2]], merges: [] }])
  })

  it('combines formatting of different cells', async () => {
    const host = new AutomergeTestHost()
    const id = await importSheets(host, [DATA])
    const result = await concurrently(
      host,
      id,
      (h, l) => save(h, l, [{ name: 'Daten', rows: [[{ v: 'a', bold: true }, 1], ['b', 2]] }]),
      (h, l) => save(h, l, [{ name: 'Daten', rows: [['a', 1], ['b', { v: 2, fill: '#FFFF00' }]] }]),
    )
    expect(result[0]!.rows).toEqual([['a[b]', 1], ['b', '2[fill#FFFF00]']])
  })

  it('keeps a value change and a format change of the same cell', async () => {
    const host = new AutomergeTestHost()
    const id = await importSheets(host, [DATA])
    const result = await concurrently(
      host,
      id,
      (h, l) => save(h, l, [{ name: 'Daten', rows: [['a', 100], ['b', 2]] }]),
      (h, l) => save(h, l, [{ name: 'Daten', rows: [['a', { v: 1, bold: true }], ['b', 2]] }]),
    )
    expect(result[0]!.rows).toEqual([['a', '100[b]'], ['b', 2]])
  })

  it('combines different format properties set on the same cell', async () => {
    const host = new AutomergeTestHost()
    const id = await importSheets(host, [DATA])
    const result = await concurrently(
      host,
      id,
      (h, l) => save(h, l, [{ name: 'Daten', rows: [[{ v: 'a', bold: true }, 1], ['b', 2]] }]),
      (h, l) => save(h, l, [{ name: 'Daten', rows: [[{ v: 'a', fill: '#FFFF00' }, 1], ['b', 2]] }]),
    )
    expect(result[0]!.rows[0]![0]).toBe('a[b+fill#FFFF00]')
  })

  it('keeps both columns when two people insert one at the same place', async () => {
    const host = new AutomergeTestHost()
    const id = await importSheets(host, [DATA])
    const result = await concurrently(
      host,
      id,
      (h, l) => save(h, l, [{ name: 'Daten', rows: [['a', 'A', 1], ['b', null, 2]] }], { structuralOps: [{ sheetId: 'f0', kind: 'insert-cols', index: 1, count: 1 }] }),
      (h, l) => save(h, l, [{ name: 'Daten', rows: [['a', 'B', 1], ['b', null, 2]] }], { structuralOps: [{ sheetId: 'f0', kind: 'insert-cols', index: 1, count: 1 }] }),
    )
    expect(result[0]!.rows[0]).toHaveLength(4)
    expect(result[0]!.rows[0]![0]).toBe('a')
    expect(result[0]!.rows[0]![3]).toBe(1)
    expect(result[0]!.rows[0]!.slice(1, 3).sort()).toEqual(['A', 'B'])
  })

  it('deletes a row once when two people delete it', async () => {
    const host = new AutomergeTestHost()
    const id = await importSheets(host, [{ name: 'Daten', rows: [['a'], ['b'], ['c']] }])
    const deleteB = (h: AutomergeTestHost, l: LoadedWorkbook) =>
      save(h, l, [{ name: 'Daten', rows: [['a'], ['c']] }], { structuralOps: [{ sheetId: 'f0', kind: 'remove-rows', index: 1, count: 1 }] })
    const result = await concurrently(host, id, deleteB, deleteB)
    expect(result[0]!.rows).toEqual([['a'], ['c']])
  })

  it('does not end up with overlapping merged cells', async () => {
    const host = new AutomergeTestHost()
    const id = await importSheets(host, [{ name: 'Daten', rows: [['a', 'b'], ['c', 'd']] }])
    const result = await concurrently(
      host,
      id,
      (h, l) => save(h, l, [{ name: 'Daten', rows: [['a', null], ['c', 'd']], merges: [[0, 0, 0, 1]] }]),
      (h, l) => save(h, l, [{ name: 'Daten', rows: [['a', 'b'], [null, 'd']], merges: [[0, 0, 1, 0]] }]),
    )
    expect(result[0]!.merges).toHaveLength(1)
  })
})
