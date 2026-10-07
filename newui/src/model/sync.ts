// Turns what the editor saved into stored-document changes.
//
// After a save the session's xlsx is re-read in full. Each sheet's rows and
// columns are matched to their ids by replaying the save's structural
// operations on the stored order lists, so an inserted row gets a new id and
// every other row keeps its own. The re-read sheet is then built into the
// stored shape (build.ts) and compared with the stored state, producing a
// JSON patch that touches only what changed.
import { MindooDBAppValue, type MindooDBAppJsonPatch } from 'mindoodb-app-sdk'

import type { WorkbookSaveRequest } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { buildWorksheet } from './build'
import { createAxesLookup, type SheetAxes } from './formula-refs'
import { extendAxis, diffList, replayAxis, type AxisOp, type ReplayResult } from './replay'
import { createId, liveIds, liveSheets, type Workbook, type Worksheet } from './schema'
import type { SidecarWorkbook } from './sidecar-read'
import { styleTable } from './styles'

const WORKBOOK_PATH = ['teamgrid', 'workbook']

/** Every string becomes an atomic value: this schema has no collaborative text. */
export function toStored(value: unknown): unknown {
  if (typeof value === 'string') return MindooDBAppValue.atomic(value)
  if (Array.isArray(value)) return value.map(toStored)
  if (value && typeof value === 'object') {
    return Object.fromEntries(Object.entries(value).map(([key, entry]) => [key, toStored(entry)]))
  }
  return value
}

function emptySheet(id: string, name: string): Worksheet {
  return { id, name, rowOrder: [], columnOrder: [], rowsById: {}, columnsById: {}, cellsById: {}, mergesById: {} }
}

interface SheetPlan {
  stored: Worksheet
  rows: ReplayResult
  columns: ReplayResult
}

function planAxes(stored: Worksheet, rowOps: AxisOp[], columnOps: AxisOp[], rowCount: number, columnCount: number) {
  const rows = replayAxis(stored.rowOrder, liveIds(stored.rowOrder, stored.rowsById), rowOps, 'row')
  const columns = replayAxis(stored.columnOrder, liveIds(stored.columnOrder, stored.columnsById), columnOps, 'col')
  extendAxis(rows, rowCount, 'row')
  extendAxis(columns, columnCount, 'col')
  return { rows, columns }
}

function buildAll(
  workbook: Workbook,
  read: SidecarWorkbook,
  plans: Map<string, SheetPlan>,
  storedIdByName: Map<string, string>,
) {
  const { ids: styleIds, byId } = styleTable(read.file.styles)
  const stylesById = { ...workbook.stylesById, ...byId }
  const axes: SheetAxes[] = read.sheets.map((sheet) => {
    const plan = plans.get(storedIdByName.get(sheet.meta.name)!)!
    return {
      id: plan.stored.id,
      name: sheet.meta.name,
      rowIds: plan.rows.live,
      columnIds: plan.columns.live,
      rowOrder: plan.rows.order,
      columnOrder: plan.columns.order,
    }
  })
  const lookup = createAxesLookup(axes)
  const deletedAt = new Date().toISOString()
  const sheets = read.sheets.map((sheet) => {
    const plan = plans.get(storedIdByName.get(sheet.meta.name)!)!
    return buildWorksheet({
      sheet,
      previous: plan.stored,
      rowOrder: plan.rows.order,
      columnOrder: plan.columns.order,
      rowIds: plan.rows.live,
      columnIds: plan.columns.live,
      deletedRowIds: plan.rows.deleted,
      deletedColumnIds: plan.columns.deleted,
      deletedAt,
      styleIds,
      axes: lookup,
    })
  })
  return { sheets, stylesById }
}

/** A freshly imported workbook: every row, column and sheet gets a new id. */
export function importWorkbook(read: SidecarWorkbook): Workbook {
  const plans = new Map<string, SheetPlan>()
  const storedIdByName = new Map<string, string>()
  for (const sheet of read.sheets) {
    const stored = emptySheet(createId('sheet'), sheet.meta.name)
    storedIdByName.set(sheet.meta.name, stored.id)
    plans.set(stored.id, { stored, ...planAxes(stored, [], [], sheet.meta.rowCount, sheet.meta.columnCount) })
  }
  const { sheets, stylesById } = buildAll(emptyWorkbook(), read, plans, storedIdByName)
  return {
    worksheetOrder: sheets.map((sheet) => sheet.id),
    worksheetsById: Object.fromEntries(sheets.map((sheet) => [sheet.id, sheet])),
    stylesById: pruneStyles(stylesById, sheets),
  }
}

export function emptyWorkbook(): Workbook {
  return { worksheetOrder: [], worksheetsById: {}, stylesById: {} }
}

/** Styles any cell, row or column still uses (an import carries the file's whole table). */
function pruneStyles(stylesById: Workbook['stylesById'], sheets: readonly Worksheet[]) {
  const used = new Set<string>()
  for (const sheet of sheets) {
    for (const cell of Object.values(sheet.cellsById)) if (cell.styleId) used.add(cell.styleId)
    for (const row of Object.values(sheet.rowsById)) if (row.styleId) used.add(row.styleId)
    for (const column of Object.values(sheet.columnsById)) if (column.styleId) used.add(column.styleId)
  }
  return Object.fromEntries(Object.entries(stylesById).filter(([id]) => used.has(id)))
}

export interface SyncInput {
  stored: Workbook
  /** The session's file sheet ids → stored sheet ids, as loaded. */
  storedIdByFileId: ReadonlyMap<string, string>
  request: WorkbookSaveRequest
  /** The saved xlsx, re-read in full. */
  read: SidecarWorkbook
}

export interface SyncResult {
  next: Workbook
  patch: Omit<MindooDBAppJsonPatch, 'baseHeads'>
  changed: boolean
}

export function syncWorkbook({ stored, storedIdByFileId, request, read }: SyncInput): SyncResult {
  // Sheet identity: file ids of this session → stored ids; sheets the editor
  // added get new ids. The re-read file is matched by final sheet name.
  const storedIdByEditorId = new Map(storedIdByFileId)
  const removed = new Set<string>()
  const finalName = new Map<string, string>()
  for (const sheet of liveSheets(stored)) finalName.set(sheet.id, sheet.name)
  for (const op of request.sheetOps) {
    if (op.kind === 'reorder-sheets') continue
    if (op.kind === 'add-sheet' || op.kind === 'duplicate-sheet') {
      const id = createId('sheet')
      storedIdByEditorId.set(op.sheetId, id)
      finalName.set(id, op.name)
      continue
    }
    const id = storedIdByEditorId.get(op.sheetId)
    if (!id) continue
    if (op.kind === 'rename-sheet') finalName.set(id, op.newName)
    else if (op.kind === 'remove-sheet') removed.add(id)
  }
  const storedIdByName = new Map<string, string>()
  for (const [id, name] of finalName) if (!removed.has(id)) storedIdByName.set(name, id)

  const rowOps = new Map<string, AxisOp[]>()
  const columnOps = new Map<string, AxisOp[]>()
  for (const op of request.structuralOps) {
    const id = storedIdByEditorId.get(op.sheetId)
    if (!id) continue
    const push = (target: Map<string, AxisOp[]>, axisOp: AxisOp) => target.set(id, [...(target.get(id) ?? []), axisOp])
    if (op.kind === 'insert-rows') push(rowOps, { kind: 'insert', index: op.index, count: op.count })
    else if (op.kind === 'remove-rows') push(rowOps, { kind: 'remove', index: op.index, count: op.count })
    else if (op.kind === 'move-rows') push(rowOps, { kind: 'move', index: op.index, count: op.count, before: op.before })
    else if (op.kind === 'insert-cols') push(columnOps, { kind: 'insert', index: op.index, count: op.count })
    else if (op.kind === 'remove-cols') push(columnOps, { kind: 'remove', index: op.index, count: op.count })
  }

  const plans = new Map<string, SheetPlan>()
  for (const sheet of read.sheets) {
    const id = storedIdByName.get(sheet.meta.name)
    if (!id) throw new Error(`Saved sheet ${sheet.meta.name} has no stored counterpart.`)
    const previous = stored.worksheetsById[id] ?? emptySheet(id, sheet.meta.name)
    plans.set(id, {
      stored: previous,
      ...planAxes(previous, rowOps.get(id) ?? [], columnOps.get(id) ?? [], sheet.meta.rowCount, sheet.meta.columnCount),
    })
  }
  const { sheets, stylesById } = buildAll(stored, read, plans, storedIdByName)

  const deletedAt = new Date().toISOString()
  const worksheetsById: Workbook['worksheetsById'] = { ...stored.worksheetsById }
  for (const sheet of sheets) worksheetsById[sheet.id] = sheet
  for (const id of removed) {
    const sheet = worksheetsById[id]
    if (sheet && !sheet.deletedAt) worksheetsById[id] = { ...sheet, deletedAt }
  }
  const liveOrder = read.sheets.map((sheet) => storedIdByName.get(sheet.meta.name)!)
  const next: Workbook = {
    worksheetOrder: [...liveOrder, ...stored.worksheetOrder.filter((id) => !liveOrder.includes(id))],
    worksheetsById,
    stylesById,
  }
  return { next, ...diffWorkbook(stored, next) }
}

// ── diff ────────────────────────────────────────────────────────────────

type Path = (string | number)[]

function equal(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

class PatchBuilder {
  readonly set: { path: Path; value: unknown }[] = []
  readonly unset: { path: Path }[] = []
  readonly listInsert: { path: Path; index: number; values: unknown[] }[] = []
  readonly listDelete: { path: Path; index: number; deleteCount: number }[] = []

  put(path: Path, value: unknown) {
    this.set.push({ path, value: toStored(value) })
  }

  list(path: Path, before: readonly string[], after: readonly string[]) {
    const { deletes, inserts } = diffList(before, after)
    for (const remove of deletes) this.listDelete.push({ path, ...remove })
    for (const insert of inserts) {
      this.listInsert.push({ path, index: insert.index, values: insert.values.map(toStored) })
    }
  }

  /** Field-by-field diff of one object (a cell, row or column). */
  fields(path: Path, before: Record<string, unknown> | undefined, after: Record<string, unknown> | undefined) {
    if (equal(before, after)) return
    if (!after) {
      this.unset.push({ path })
      return
    }
    if (!before) {
      this.put(path, after)
      return
    }
    for (const key of new Set([...Object.keys(before), ...Object.keys(after)])) {
      if (equal(before[key], after[key])) continue
      if (after[key] === undefined) this.unset.push({ path: [...path, key] })
      else this.put([...path, key], after[key])
    }
  }

  map<T>(path: Path, before: Record<string, T>, after: Record<string, T>, deep: boolean) {
    for (const key of new Set([...Object.keys(before), ...Object.keys(after)])) {
      const left = before[key]
      const right = after[key]
      if (deep) this.fields([...path, key], left as Record<string, unknown>, right as Record<string, unknown>)
      else if (!equal(left, right)) {
        if (right === undefined) this.unset.push({ path: [...path, key] })
        else this.put([...path, key], right)
      }
    }
  }

  get empty() {
    return !this.set.length && !this.unset.length && !this.listInsert.length && !this.listDelete.length
  }
}

const SHEET_COLLECTIONS = ['rowsById', 'columnsById', 'cellsById'] as const

function diffWorkbook(before: Workbook, after: Workbook): Omit<SyncResult, 'next'> {
  const patch = new PatchBuilder()
  patch.list([...WORKBOOK_PATH, 'worksheetOrder'], before.worksheetOrder, after.worksheetOrder)
  patch.map([...WORKBOOK_PATH, 'stylesById'], before.stylesById, after.stylesById, false)
  for (const [id, sheet] of Object.entries(after.worksheetsById)) {
    const path = [...WORKBOOK_PATH, 'worksheetsById', id]
    const previous = before.worksheetsById[id]
    if (!previous) {
      patch.put(path, sheet)
      continue
    }
    if (equal(previous, sheet)) continue
    patch.list([...path, 'rowOrder'], previous.rowOrder, sheet.rowOrder)
    patch.list([...path, 'columnOrder'], previous.columnOrder, sheet.columnOrder)
    for (const collection of SHEET_COLLECTIONS) {
      patch.map([...path, collection], previous[collection], sheet[collection], true)
    }
    patch.map([...path, 'mergesById'], previous.mergesById, sheet.mergesById, false)
    const skip = new Set<string>(['rowOrder', 'columnOrder', 'mergesById', ...SHEET_COLLECTIONS])
    const scalars = (value: Worksheet) =>
      Object.fromEntries(Object.entries(value).filter(([key]) => !skip.has(key))) as Record<string, unknown>
    patch.fields(path, scalars(previous), scalars(sheet))
  }
  const { set, unset, listInsert, listDelete } = patch
  return { patch: { set, unset, listInsert, listDelete }, changed: !patch.empty }
}
