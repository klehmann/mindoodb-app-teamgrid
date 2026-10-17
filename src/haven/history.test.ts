// Time travel's timelines: the whole workbook, or one sheet (only the saves
// that changed it), and a sheet read back as it was while the others stay.
import { describe, expect, it } from 'vitest'

import type { WorkbookSaveRequest } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { sheetAxes } from '../model/export'
import { liveIds, type Worksheet } from '../model/schema'
import type { SheetMetadata, SidecarWorkbook } from '../model/sidecar-read'
import { emptyStoredWorkbook, syncWorkbook } from '../model/sync'
import { AutomergeTestHost } from '../testing/automerge-host'
import { listRevisions, listSheetRevisions, loadSheetAt } from './history'
import { createWorkbook, loadWorkbook, writeWorkbook } from './store'

type Grid = (string | number | null)[][]

/** What the editor's file holds: sheets "One" (file id f1) and "Two" (f2). */
function read(one: Grid, two: Grid): SidecarWorkbook {
  const sheet = (id: string, name: string, rows: Grid) => ({
    meta: {
      id,
      name,
      rowCount: Math.max(1, rows.length),
      columnCount: Math.max(1, ...rows.map((row) => row.length)),
      columnWidths: [],
      defaultRowHeight: null,
      defaultColumnWidth: null,
      freeze: null,
      hidden: false,
      tabColor: null,
      showGridLines: true,
      tables: [],
    } as unknown as SheetMetadata,
    cells: rows.flatMap((row, rowIndex) =>
      row.flatMap((value, column) => (value === null ? [] : [{ row: rowIndex, column, value }])),
    ),
    rows: [],
    merges: [],
  })
  return {
    file: { sessionId: 'x', styles: [], visuals: [] } as unknown as SidecarWorkbook['file'],
    sheets: [sheet('f1', 'One', one), sheet('f2', 'Two', two)],
  }
}

const request = { sessionId: 'x', sheetOps: [], structuralOps: [] } as unknown as WorkbookSaveRequest

function values(sheet: Worksheet): Grid {
  const axes = sheetAxes(sheet)
  return liveIds(sheet.rowOrder, sheet.rowsById).map((rowId) =>
    axes.columnIds.map((columnId) => (sheet.cellsById[`${rowId}:${columnId}`]?.value ?? null) as string | number | null),
  )
}

/** A workbook saved four times a minute apart: created, sheet One, sheet Two, sheet One again. */
async function fourSaves() {
  const host = new AutomergeTestHost()
  const first = syncWorkbook({ stored: emptyStoredWorkbook(), storedIdByFileId: new Map(), read: read([['a']], [['x']]) })
  const id = await createWorkbook(host.connection, 'Test', first.next, first.writes)
  const save = async (one: Grid, two: Grid) => {
    host.now += 60_000
    const loaded = await loadWorkbook(host.connection, id)
    const [firstSheet, secondSheet] = loaded.stored.workbook.worksheetOrder
    const { writes } = syncWorkbook({
      stored: loaded.stored,
      storedIdByFileId: new Map([
        ['f1', firstSheet!],
        ['f2', secondSheet!],
      ]),
      request,
      read: read(one, two),
    })
    await writeWorkbook(host.connection, loaded, writes)
    return host.now
  }
  const created = host.now
  const oneChanged = await save([['b']], [['x']])
  const twoChanged = await save([['b']], [['y']])
  const oneAgain = await save([['c']], [['y']])
  const loaded = await loadWorkbook(host.connection, id)
  const [one, two] = loaded.stored.workbook.worksheetOrder as [string, string]
  return { host, id, loaded, one, two, times: { created, oneChanged, twoChanged, oneAgain } }
}

describe('time travel', () => {
  it('a sheet’s timeline holds only the saves that changed it', async () => {
    const { host, loaded, one, two, times } = await fourSaves()
    const all = await listRevisions(host.connection, loaded)
    expect(all.map((revision) => revision.timestamp)).toEqual([times.oneAgain, times.twoChanged, times.oneChanged, times.created])

    const ofOne = await listSheetRevisions(host.connection, loaded, one)
    expect(ofOne.map((revision) => revision.timestamp)).toEqual([times.oneAgain, times.oneChanged, times.created])
    const ofTwo = await listSheetRevisions(host.connection, loaded, two)
    expect(ofTwo.map((revision) => revision.timestamp)).toEqual([times.twoChanged, times.created])
  })

  it('reads one sheet as it was and every other sheet as it is', async () => {
    const { host, loaded, one, two, times } = await fourSaves()
    const { workbook } = await loadSheetAt(host.connection, loaded, one, times.oneChanged)
    expect(values(workbook.worksheetsById[one]!)).toEqual([['b']])
    expect(values(workbook.worksheetsById[two]!)).toEqual([['y']])

    const before = await loadSheetAt(host.connection, loaded, two, times.created)
    expect(values(before.workbook.worksheetsById[two]!)).toEqual([['x']])
    expect(values(before.workbook.worksheetsById[one]!)).toEqual([['c']])
  })
})
