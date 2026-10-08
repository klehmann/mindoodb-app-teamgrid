import type { MindooDBAppResolvedViewDefinition, MindooDBAppViewEntry } from 'mindoodb-app-sdk'
import { describe, expect, it } from 'vitest'

import { createWorkbook, loadWorkbook, writeWorkbook } from '../haven/store'
import { AutomergeTestHost } from '../testing/automerge-host'
import { sheetAxes } from './export'
import { liveIds, liveSheets, type Workbook } from './schema'
import { styleOf } from './styles'
import { newWorkbook, writesFor } from './sync'
import { buildViewSheet, cellOf, parseCategoryPath, viewTable } from './view-sheet'

const column = (name: string, title: string, role: 'category' | 'display' = 'display') =>
  ({ id: name, name, title, role, expression: { mode: 'field', field: name }, sorting: 'none', totalMode: 'none', hidden: false }) as MindooDBAppResolvedViewDefinition['columns'][number]

const VIEW = {
  id: 'contacts',
  description: 'Kontakte',
  categorizationStyle: 'category_then_document',
  previewMode: 'table',
  sources: [],
  filter: { mode: 'rules', match: 'all', rules: [] },
  columns: [column('company', 'Firma', 'category'), column('name', 'Name'), column('since', 'Seit')],
} as unknown as MindooDBAppResolvedViewDefinition

const entry = (kind: 'category' | 'document', values: Record<string, unknown>) =>
  ({ kind, columnValues: values }) as unknown as MindooDBAppViewEntry

const ENTRIES = [
  entry('category', { company: 'ACME' }),
  entry('document', { company: 'ACME', name: 'Alice', since: '2026-05-20' }),
  entry('document', { company: 'ACME', name: 'Bob', since: 3 }),
]

function show(workbook: Workbook, name: string) {
  const sheet = liveSheets(workbook).find((candidate) => candidate.name === name)!
  const { columnIds } = sheetAxes(sheet)
  return liveIds(sheet.rowOrder, sheet.rowsById).map((rowId) =>
    columnIds.map((columnId) => {
      const cell = sheet.cellsById[`${rowId}:${columnId}`]
      const style = styleOf(cell)
      return cell ? `${cell.value ?? ''}${style?.bold ? '[b]' : ''}${style?.fillColor ? '[fill]' : ''}${style?.numberFormat ? `[${style.numberFormat}]` : ''}` : null
    }),
  )
}

describe('view sheets', () => {
  it('reads category paths and view values', () => {
    expect(parseCategoryPath(' Customers \\ ACME ')).toEqual(['Customers', 'ACME'])
    expect(cellOf('2026-05-20')).toMatchObject({ value: 46162, 's.numberFormat': 'yyyy-mm-dd' })
    expect(cellOf(['a', 'b'])).toEqual({ value: 'a, b' })
    expect(cellOf(null)).toBeUndefined()
  })

  it('stores a view sheet next to the others and refreshes it in place', async () => {
    const host = new AutomergeTestHost()
    const id = await createWorkbook(host.connection, 'Test', newWorkbook('Tabelle1'), {
      newImages: [],
      top: { set: [], unset: [], listInsert: [], listDelete: [] },
      createdChunks: new Map(),
      chunks: new Map(),
      changed: true,
    })
    const binding = { viewId: 'contacts', viewTitle: 'Kontakte', showDocuments: true, showCategories: true, rootCategoryPath: [], lastRefreshedAt: '2026-10-08T00:00:00.000Z' }

    let loaded = await loadWorkbook(host.connection, id)
    let sheet = buildViewSheet(undefined, 'sview', 'Kontakte', binding, viewTable(VIEW, ENTRIES, binding))
    let next: Workbook = {
      ...loaded.stored.workbook,
      worksheetOrder: [...loaded.stored.workbook.worksheetOrder, sheet.id],
      worksheetsById: { ...loaded.stored.workbook.worksheetsById, [sheet.id]: sheet },
    }
    await writeWorkbook(host.connection, loaded, writesFor(loaded.stored, next, [sheet.id]))

    loaded = await loadWorkbook(host.connection, id)
    expect(liveSheets(loaded.stored.workbook).map((s) => s.name)).toEqual(['Tabelle1', 'Kontakte'])
    expect(loaded.stored.workbook.worksheetsById.sview!.viewBinding).toEqual(binding)
    expect(show(loaded.stored.workbook, 'Kontakte')).toEqual([
      ['Firma[b]', 'Name[b]', 'Seit[b]'],
      ['ACME[fill]', '[fill]', '[fill]'],
      [null, 'Alice', '46162[yyyy-mm-dd]'],
      [null, 'Bob', '3'],
    ])

    // Refresh with only documents: one row less, the remaining rows keep their ids.
    const before = loaded.stored.workbook.worksheetsById.sview!
    const documentsOnly = { ...binding, showCategories: false, lastRefreshedAt: '2026-10-09T00:00:00.000Z' }
    sheet = buildViewSheet(before, 'sview', 'Kontakte', documentsOnly, viewTable(VIEW, ENTRIES, documentsOnly))
    next = { ...loaded.stored.workbook, worksheetsById: { ...loaded.stored.workbook.worksheetsById, sview: sheet } }
    await writeWorkbook(host.connection, loaded, writesFor(loaded.stored, next, ['sview']))

    loaded = await loadWorkbook(host.connection, id)
    const after = loaded.stored.workbook.worksheetsById.sview!
    expect(liveIds(after.rowOrder, after.rowsById)).toEqual(liveIds(before.rowOrder, before.rowsById).slice(0, 3))
    expect(show(loaded.stored.workbook, 'Kontakte')).toEqual([
      ['Firma[b]', 'Name[b]', 'Seit[b]'],
      [null, 'Alice', '46162[yyyy-mm-dd]'],
      [null, 'Bob', '3'],
    ])
    expect(after.viewBinding?.showCategories).toBe(false)
  })
})
