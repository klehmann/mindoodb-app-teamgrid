// Turns what the editor saved into stored-document changes.
//
// After a save the session's xlsx is re-read in full. Each sheet's rows and
// columns are matched to their ids by replaying the save's structural
// operations on the stored order lists, so an inserted row gets a new id and
// every other row keeps its own. The re-read sheet is then built into the
// stored shape (build.ts), its rows are placed into row blocks (chunks), and
// the result is compared with what is stored: one JSON patch for the top
// document and one per chunk that changed.
import { MindooDBAppValue, type MindooDBAppJsonPatch } from 'mindoodb-app-sdk'

import type { WorkbookSaveRequest } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { buildWorksheet } from './build'
import { createAxesLookup, NO_HOME, parseFormula, type AxesLookup, type SheetAxes } from './formula-refs'
import { diffList, grow, replayAxis, type AxisOp, type ReplayResult } from './replay'
import {
  CHUNK_FIELDS,
  createId,
  nameKey,
  type DefinedName,
  liveIds,
  liveSheets,
  type Cell,
  type CellKey,
  type ChunkId,
  type Merge,
  type RowId,
  type RowMeta,
  type SheetId,
  type VisualId,
  type Workbook,
  type Worksheet,
} from './schema'
import type { SidecarWorkbook } from './sidecar-read'
import { canonicalJson, toStyleEdit, withoutLegacyStyle } from './styles'
import { featureExtent } from './sheet-features'
import { anchorExtent, normalizeAnchor, sheetSizes, storedVisual } from './visuals'

/** Rows a chunk takes before appending at its end opens a new one. */
export const CHUNK_SOFT_LIMIT = 256

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

/** A chunk document's content as stored. */
export interface ChunkContent {
  sheetId: SheetId
  rowOrder: RowId[]
  rowsById: Record<RowId, RowMeta>
  cellsById: Record<CellKey, Cell>
}

/** A stored workbook as loaded: the assembled view plus where each row lives. */
export interface StoredWorkbook {
  /** What the editor works on: merged, repaired (see assembleWorkbook). */
  workbook: Workbook
  /** The top document's workbook exactly as stored; saves diff against it, so repairs get written. */
  top: Workbook
  chunks: Map<ChunkId, ChunkContent>
  /** Per sheet: row id → the chunk that holds it (the first one, if a merge left copies). */
  rowHome: Map<SheetId, Map<RowId, ChunkId>>
}

export type JsonPatchBody = Required<Pick<MindooDBAppJsonPatch, 'set' | 'unset' | 'listInsert' | 'listDelete'>>

/** A picture added in the editor: its bytes go into an attachment of the top document. */
export interface NewImage {
  attachment: string
  mediaType: string
  /** The visual's id in the saved file, to read its bytes. */
  fileVisualId: string
  bytes?: Uint8Array
}

export interface WorkbookWrites {
  newImages: NewImage[]
  top: JsonPatchBody
  /** Chunks to create first (empty skeleton, derived id), then patch like the others. */
  createdChunks: Map<ChunkId, SheetId>
  chunks: Map<ChunkId, JsonPatchBody>
  changed: boolean
}

/** The empty content a chunk document is created with: equal on every replica. */
export function chunkSkeleton(sheetId: SheetId): ChunkContent {
  return { sheetId, rowOrder: [], rowsById: {}, cellsById: {} }
}

/**
 * Id of a chunk opened after `after` for `firstRow`. Derived, not random, so
 * two replicas that append the same derived rows after the same full chunk
 * open the same chunk document and their writes merge instead of forking.
 * MindooDB document ids allow lowercase letters, digits and `_`.
 */
export function chunkIdFor(sheetId: SheetId, after: ChunkId | undefined, firstRow: RowId): ChunkId {
  let hash = 0xcbf29ce4 >>> 0
  let hash2 = 0x84222325 >>> 0
  const text = `${sheetId}|${after ?? ''}|${firstRow}`
  for (let index = 0; index < text.length; index += 1) {
    hash = Math.imul(hash ^ text.charCodeAt(index), 16777619) >>> 0
    hash2 = Math.imul(hash2 ^ text.charCodeAt(index), 2246822519) >>> 0
  }
  return `gc_${hash.toString(36)}${hash2.toString(36)}`
}

/** Assembles the in-memory workbook from the top document's workbook and its chunks. */
export function assembleWorkbook(top: Workbook, chunks: Map<ChunkId, ChunkContent>): StoredWorkbook {
  const legacy = top.stylesById
  const rowHome = new Map<SheetId, Map<RowId, ChunkId>>()
  const worksheetsById: Workbook['worksheetsById'] = {}
  for (const [sheetId, sheet] of Object.entries(top.worksheetsById)) {
    const home = new Map<RowId, ChunkId>()
    const rowOrder: RowId[] = []
    const rowsById: Record<RowId, RowMeta> = {}
    const cellsById: Record<CellKey, Cell> = {}
    const chunkOrder = [...new Set(sheet.chunkOrder ?? [])]
    for (const chunkId of chunkOrder) {
      const chunk = chunks.get(chunkId)
      if (!chunk) continue
      for (const rowId of chunk.rowOrder) {
        if (home.has(rowId)) continue
        home.set(rowId, chunkId)
        rowOrder.push(rowId)
        const row = chunk.rowsById[rowId]
        if (row) rowsById[rowId] = withoutLegacyStyle(row, legacy)
      }
    }
    // Cells of every copy count: a row two replicas appended under the same
    // derived id may hold cells in two chunks. The home chunk wins a tie.
    for (const chunkId of [...chunkOrder].reverse()) {
      const chunk = chunks.get(chunkId)
      if (chunk) Object.assign(cellsById, chunk.cellsById)
    }
    for (const chunkId of chunkOrder) {
      const chunk = chunks.get(chunkId)
      if (!chunk) continue
      for (const [key, cell] of Object.entries(chunk.cellsById)) {
        if (home.get(key.split(':')[0]!) === chunkId) cellsById[key as CellKey] = cell
      }
    }
    if (legacy) for (const [key, cell] of Object.entries(cellsById)) cellsById[key as CellKey] = withoutLegacyStyle(cell, legacy)
    const columnsById = Object.fromEntries(
      Object.entries(sheet.columnsById).map(([id, column]) => [id, withoutLegacyStyle(column, legacy)]),
    )
    worksheetsById[sheetId] = {
      ...sheet,
      columnsById,
      mergesById: withoutOverlaps(sheet.mergesById, rowOrder, rowsById, sheet.columnOrder, sheet.columnsById),
      visualOrder: sheet.visualOrder ?? [],
      visualsById: sheet.visualsById ?? {},
      chunkOrder,
      rowOrder,
      rowsById,
      cellsById,
    }
    rowHome.set(sheetId, home)
  }
  const { stylesById: _legacy, ...rest } = top
  const workbook: Workbook = { ...rest, worksheetsById }
  distinctSheetNames(workbook)
  return { workbook, top, chunks, rowHome }
}

/**
 * Two people who each added a sheet with the same name offline end up with
 * two sheets of that name, which a workbook cannot hold. Later sheets in tab
 * order get a number appended, the same way on every replica; the next save
 * writes the new name.
 */
function distinctSheetNames(workbook: Workbook): void {
  const taken = new Set<string>()
  for (const sheet of liveSheets(workbook)) {
    let name = sheet.name
    for (let n = 2; taken.has(name.toLowerCase()); n += 1) name = `${sheet.name} (${n})`
    taken.add(name.toLowerCase())
    if (name !== sheet.name) workbook.worksheetsById[sheet.id] = { ...sheet, name }
  }
}

/**
 * Merged areas two people created offline may overlap, which a sheet cannot
 * show. The first by key wins, the same way on every replica; the next save
 * removes the others.
 */
function withoutOverlaps(
  mergesById: Record<string, Merge>,
  rowOrder: readonly RowId[],
  rowsById: Record<RowId, RowMeta>,
  columnOrder: readonly string[],
  columnsById: Worksheet['columnsById'],
): Record<string, Merge> {
  const rowIndex = new Map(liveIds(rowOrder, rowsById).map((id, index) => [id, index]))
  const columnIndex = new Map(liveIds(columnOrder, columnsById).map((id, index) => [id, index]))
  const kept: Record<string, Merge> = {}
  const areas: [number, number, number, number][] = []
  for (const key of Object.keys(mergesById).sort()) {
    const merge = mergesById[key]!
    const area = [
      rowIndex.get(merge.startRowId),
      columnIndex.get(merge.startColumnId),
      rowIndex.get(merge.endRowId),
      columnIndex.get(merge.endColumnId),
    ]
    if (area.some((value) => value === undefined)) {
      // A corner was deleted: the export skips it; keep it for an undo elsewhere.
      kept[key] = merge
      continue
    }
    const [top, left, bottom, right] = area as [number, number, number, number]
    if (areas.some(([t, l, b, r]) => top <= b && t <= bottom && left <= r && l <= right)) continue
    areas.push([top, left, bottom, right])
    kept[key] = merge
  }
  return kept
}

function emptySheet(id: string, name: string): Worksheet {
  return {
    id,
    name,
    rowOrder: [],
    columnOrder: [],
    rowsById: {},
    columnsById: {},
    cellsById: {},
    mergesById: {},
    chunkOrder: [],
    visualOrder: [],
    visualsById: {},
  }
}

interface SheetPlan {
  stored: Worksheet
  rows: ReplayResult
  columns: ReplayResult
}

function planAxes(stored: Worksheet, rowOps: AxisOp[], columnOps: AxisOp[], rowCount: number, columnCount: number) {
  const rows = replayAxis(stored.rowOrder, liveIds(stored.rowOrder, stored.rowsById), rowOps, 'r', `${stored.id}:rows`)
  const columns = replayAxis(
    stored.columnOrder,
    liveIds(stored.columnOrder, stored.columnsById),
    columnOps,
    'c',
    `${stored.id}:columns`,
  )
  grow(rows, rowCount, 'r', `${stored.id}:rows`)
  grow(columns, columnCount, 'c', `${stored.id}:columns`)
  return { rows, columns }
}

function buildAll(
  workbook: Workbook,
  read: SidecarWorkbook,
  plans: Map<string, SheetPlan>,
  storedIdByName: Map<string, string>,
) {
  const styles = read.file.styles.map(toStyleEdit)
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
      styles,
      file: read.file,
      axes: lookup,
    })
  })
  return { sheets, lookup }
}

export function emptyStoredWorkbook(): StoredWorkbook {
  const workbook: Workbook = { worksheetOrder: [], worksheetsById: {} }
  return { workbook, top: workbook, chunks: new Map(), rowHome: new Map() }
}

/** A new workbook with one empty sheet. */
export function newWorkbook(sheetName: string): Workbook {
  const sheet = emptySheet(createId('s'), sheetName)
  return { worksheetOrder: [sheet.id], worksheetsById: { [sheet.id]: sheet }, namesById: {} }
}

export interface SyncInput {
  stored: StoredWorkbook
  /** The session's file sheet ids → stored sheet ids, as loaded. */
  storedIdByFileId: ReadonlyMap<string, string>
  /** The editor's save; omit for an import (everything is new). */
  request?: WorkbookSaveRequest
  /** The saved xlsx, re-read in full. */
  read: SidecarWorkbook
  /** The session's visuals: file visual (`drawingPath#drawingIndex`) → stored id, and per sheet the file order. */
  visuals?: { byFileKey: ReadonlyMap<string, VisualId>; orderBySheet: ReadonlyMap<SheetId, VisualId[]> }
}

export interface SyncResult {
  next: Workbook
  writes: WorkbookWrites
}

export function visualFileKey(visual: { drawingPath?: string | undefined; drawingIndex?: number | undefined; id: string }) {
  return visual.drawingPath === undefined ? visual.id : `${visual.drawingPath}#${visual.drawingIndex ?? 0}`
}

/**
 * The file's defined names, bound to ids. A name whose references did not
 * change keeps its stored formula (as cell formulas do), so a save does not
 * rewrite every name. Hidden and built-in names (_xlnm.*) are the file's own
 * business and are not stored.
 */
function definedNames(
  read: SidecarWorkbook,
  previous: Record<string, DefinedName>,
  scopeOf: (sheetIndex: number) => SheetId | undefined,
  lookup: AxesLookup,
): Record<string, DefinedName> {
  const names: Record<string, DefinedName> = {}
  for (const entry of read.file.definedNames ?? []) {
    if (entry.name.startsWith('_xlnm.')) continue
    const scopeSheetId = entry.sheetIndex === undefined ? undefined : scopeOf(entry.sheetIndex)
    if (entry.sheetIndex !== undefined && !scopeSheetId) continue
    const key = nameKey(entry.name, scopeSheetId)
    const formula = parseFormula(entry.formula.replace(/^=/, ''), NO_HOME, lookup)
    const before = previous[key]
    const name: DefinedName = {
      name: entry.name,
      formula: before && canonicalJson(before.formula) === canonicalJson(formula) ? before.formula : formula,
    }
    if (scopeSheetId) name.scopeSheetId = scopeSheetId
    names[key] = name
  }
  return names
}

/**
 * Writes for a workbook changed outside the editor (say a view sheet filled
 * from MindooDB): `next` is the stored workbook with some sheets replaced or
 * added; only those sheets' rows are placed into chunks and diffed.
 */
export function writesFor(stored: StoredWorkbook, next: Workbook, changedSheetIds: readonly SheetId[]): WorkbookWrites {
  const createdChunks = new Map<ChunkId, SheetId>()
  for (const id of changedSheetIds) {
    const sheet = next.worksheetsById[id]
    if (sheet) placeRows(sheet, stored.rowHome.get(id) ?? new Map(), stored.chunks, createdChunks)
  }
  return { ...diffWorkbook(stored, next, createdChunks), newImages: [] }
}

export function syncWorkbook({ stored, storedIdByFileId, request, read: fileRead, visuals }: SyncInput): SyncResult {
  const workbook = stored.workbook
  // Visual anchors with their far corner in the cell it falls into.
  const sizesBySheet = new Map(fileRead.sheets.map((sheet) => [sheet.meta.id, sheetSizes(sheet)]))
  const read: SidecarWorkbook = {
    ...fileRead,
    file: {
      ...fileRead.file,
      visuals: fileRead.file.visuals.map((visual) => {
        const sizes = sizesBySheet.get(visual.sheetId)
        return sizes ? normalizeAnchor(visual, sizes) : visual
      }),
    },
  }
  // Sheet identity: file ids of this session → stored ids; sheets the editor
  // added get new ids. The re-read file is matched by final sheet name.
  const storedIdByEditorId = new Map(storedIdByFileId)
  const removed = new Set<string>()
  const finalName = new Map<string, string>()
  for (const sheet of liveSheets(workbook)) finalName.set(sheet.id, sheet.name)
  if (!request) {
    for (const sheet of read.sheets) finalName.set(createId('s'), sheet.meta.name)
  }
  for (const op of request?.sheetOps ?? []) {
    if (op.kind === 'reorder-sheets') continue
    if (op.kind === 'add-sheet' || op.kind === 'duplicate-sheet') {
      const id = createId('s')
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
  for (const op of request?.structuralOps ?? []) {
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
    const previous = workbook.worksheetsById[id] ?? emptySheet(id, sheet.meta.name)
    // Ids must also cover the cells visuals are anchored to, and the areas of
    // rules, the filter, links and notes (a rule may reach past the last value).
    const features = featureExtent(sheet)
    let rowCount = Math.max(sheet.meta.rowCount, features.rows)
    let columnCount = Math.max(sheet.meta.columnCount, features.columns)
    for (const visual of read.file.visuals) {
      if (visual.sheetId !== sheet.meta.id) continue
      const extent = anchorExtent(visual)
      rowCount = Math.max(rowCount, extent.rows)
      columnCount = Math.max(columnCount, extent.columns)
    }
    plans.set(id, {
      stored: previous,
      ...planAxes(previous, rowOps.get(id) ?? [], columnOps.get(id) ?? [], rowCount, columnCount),
    })
  }
  const { sheets, lookup } = buildAll(workbook, read, plans, storedIdByName)

  // Visuals keep their ids: the file lists a sheet's visuals in drawing
  // order, which is the session's order minus the removed ones plus the
  // added ones at the end. If that does not add up, ids start over.
  const removedVisuals = new Set<VisualId>()
  for (const edit of request?.visualEdits ?? []) {
    const id = edit.remove ? visuals?.byFileKey.get(visualFileKey({ ...edit, id: '' })) : undefined
    if (id) removedVisuals.add(id)
  }
  const newImages: NewImage[] = []
  for (const [index, sheet] of sheets.entries()) {
    const fileSheet = read.sheets[index]!
    const fileVisuals = read.file.visuals.filter((visual) => visual.sheetId === fileSheet.meta.id)
    const added = (request?.visualAdditions ?? []).filter((visual) => storedIdByEditorId.get(visual.sheetId) === sheet.id)
    let ids = [
      ...(visuals?.orderBySheet.get(sheet.id) ?? []).filter((id) => !removedVisuals.has(id)),
      ...added.map(() => createId('v')),
    ]
    if (ids.length !== fileVisuals.length) ids = fileVisuals.map(() => createId('v'))
    const home = lookup.bySheetId(sheet.id)!
    const previous = plans.get(sheet.id)!.stored
    sheet.visualOrder = []
    sheet.visualsById = {}
    fileVisuals.forEach((fileVisual, position) => {
      const id = ids[position]!
      const before = previous.visualsById?.[id]
      const extension = fileVisual.mediaType?.split('/')[1] ?? 'bin'
      const visual = storedVisual(fileVisual, home, lookup, `${id}.${extension}`)
      if (!visual) return
      if (visual.image) {
        if (before?.image) visual.image = before.image
        else newImages.push({ attachment: visual.image.attachment, mediaType: visual.image.mediaType, fileVisualId: fileVisual.id })
      }
      sheet.visualOrder.push(id)
      sheet.visualsById[id] = visual
    })
  }

  const namesById = definedNames(read, workbook.namesById ?? {}, (index) => {
    const sheet = read.sheets[index]
    return sheet ? storedIdByName.get(sheet.meta.name) : undefined
  }, lookup)

  const deletedAt = new Date().toISOString()
  const worksheetsById: Workbook['worksheetsById'] = { ...workbook.worksheetsById }
  for (const sheet of sheets) worksheetsById[sheet.id] = sheet
  for (const id of removed) {
    const sheet = worksheetsById[id]
    if (sheet && !sheet.deletedAt) worksheetsById[id] = { ...sheet, deletedAt }
  }
  const liveOrder = read.sheets.map((sheet) => storedIdByName.get(sheet.meta.name)!)
  const next: Workbook = {
    worksheetOrder: [...liveOrder, ...workbook.worksheetOrder.filter((id) => !liveOrder.includes(id))],
    worksheetsById,
    namesById,
  }
  const createdChunks = new Map<ChunkId, SheetId>()
  for (const sheet of sheets) placeRows(sheet, stored.rowHome.get(sheet.id) ?? new Map(), stored.chunks, createdChunks)
  return { next, writes: { ...diffWorkbook(stored, next, createdChunks), newImages } }
}

// ── placement ──────────────────────────────────────────────────────────

/** Row id → chunk, for the sheet's next state. Kept on the sheet object during a save. */
const placements = new WeakMap<Worksheet, Map<RowId, ChunkId>>()

/**
 * Decides which chunk each row of `sheet` lives in, and opens new chunks
 * where needed (updating `sheet.chunkOrder`):
 *
 * - a row keeps its chunk unless the order forces it out (a move across
 *   chunks); existing rows are never redistributed, so a concurrent edit to
 *   a row always finds it where it was
 * - a new row joins the chunk of the row before it; when that chunk is full
 *   and the row comes after its last row (appending), a new chunk opens
 *   right after it. Inserting in the middle of a full chunk lets it grow.
 */
function placeRows(
  sheet: Worksheet,
  previousHome: Map<RowId, ChunkId>,
  chunks: Map<ChunkId, ChunkContent>,
  createdChunks: Map<ChunkId, SheetId>,
): void {
  const chunkOrder = [...sheet.chunkOrder]
  const position = (chunkId: ChunkId) => chunkOrder.indexOf(chunkId)
  const lastStoredAt = new Map<ChunkId, number>()
  sheet.rowOrder.forEach((rowId, index) => {
    const chunkId = previousHome.get(rowId)
    if (chunkId) lastStoredAt.set(chunkId, index)
  })
  const size = new Map<ChunkId, number>()
  for (const chunkId of chunkOrder) size.set(chunkId, 0)
  const home = new Map<RowId, ChunkId>()
  let current: ChunkId | undefined
  const open = (after: ChunkId | undefined, firstRow: RowId) => {
    const chunkId = chunkIdFor(sheet.id, after, firstRow)
    chunkOrder.splice(after === undefined ? 0 : position(after) + 1, 0, chunkId)
    size.set(chunkId, 0)
    if (!chunks.has(chunkId)) createdChunks.set(chunkId, sheet.id)
    return chunkId
  }
  sheet.rowOrder.forEach((rowId, index) => {
    const stored = previousHome.get(rowId)
    if (stored !== undefined && (current === undefined || position(stored) >= position(current!))) {
      current = stored
    } else if (current === undefined) {
      current = chunkOrder[0] ?? open(undefined, rowId)
    } else if ((size.get(current) ?? 0) >= CHUNK_SOFT_LIMIT && index > (lastStoredAt.get(current) ?? -1)) {
      current = open(current, rowId)
    }
    home.set(rowId, current)
    size.set(current, (size.get(current) ?? 0) + 1)
  })
  sheet.chunkOrder = chunkOrder
  placements.set(sheet, home)
}

// ── diff ────────────────────────────────────────────────────────────────

type Path = (string | number)[]

/** Stored documents come back with sorted keys, so content is compared key-order-free. */
function equal(left: unknown, right: unknown): boolean {
  return canonicalJson(left) === canonicalJson(right)
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

  body(): JsonPatchBody {
    return { set: this.set, unset: this.unset, listInsert: this.listInsert, listDelete: this.listDelete }
  }
}

/** The top-document part of a sheet: everything but the chunk fields. */
function topPart(sheet: Worksheet): Record<string, unknown> {
  const skip = new Set<string>(CHUNK_FIELDS)
  return Object.fromEntries(Object.entries(sheet).filter(([key]) => !skip.has(key)))
}

function diffWorkbook(
  stored: StoredWorkbook,
  next: Workbook,
  createdChunks: Map<ChunkId, SheetId>,
): Omit<WorkbookWrites, 'newImages'> {
  const before = stored.top
  const top = new PatchBuilder()
  top.list([...WORKBOOK_PATH, 'worksheetOrder'], before.worksheetOrder, next.worksheetOrder)
  // Formats moved onto the cells (assembleWorkbook): the registry goes.
  if (before.stylesById) top.unset.push({ path: [...WORKBOOK_PATH, 'stylesById'] })
  if (before.namesById) top.map([...WORKBOOK_PATH, 'namesById'], before.namesById, next.namesById ?? {}, false)
  else if (next.namesById && Object.keys(next.namesById).length > 0) top.put([...WORKBOOK_PATH, 'namesById'], next.namesById)

  const chunkPatches = new Map<ChunkId, PatchBuilder>()
  const chunkPatch = (chunkId: ChunkId) => {
    let patch = chunkPatches.get(chunkId)
    if (!patch) chunkPatches.set(chunkId, (patch = new PatchBuilder()))
    return patch
  }

  for (const [id, sheet] of Object.entries(next.worksheetsById)) {
    const path = [...WORKBOOK_PATH, 'worksheetsById', id]
    const previous = before.worksheetsById[id]
    const nextTop = topPart(sheet)
    if (!previous) {
      top.put(path, nextTop)
    } else {
      const previousTop = topPart(previous)
      top.list([...path, 'columnOrder'], previous.columnOrder, sheet.columnOrder)
      top.list([...path, 'chunkOrder'], previous.chunkOrder ?? [], sheet.chunkOrder)
      top.map([...path, 'columnsById'], previous.columnsById, sheet.columnsById, true)
      top.map([...path, 'mergesById'], previous.mergesById, sheet.mergesById, false)
      top.list([...path, 'visualOrder'], previous.visualOrder ?? [], sheet.visualOrder)
      top.map([...path, 'visualsById'], previous.visualsById ?? {}, sheet.visualsById, false)
      for (const field of ['conditionalFormatsById', 'dataValidationsById'] as const) {
        const before = previous[field]
        const after = sheet[field]
        // A map that is missing on one side is put or removed whole.
        if (before && after) top.map([...path, field], before, after, false)
        else if (before) top.unset.push({ path: [...path, field] })
        else if (after) top.put([...path, field], after)
      }
      const scalars = (value: Record<string, unknown>) => {
        const {
          columnOrder: _c,
          chunkOrder: _k,
          columnsById: _b,
          mergesById: _m,
          visualOrder: _vo,
          visualsById: _v,
          conditionalFormatsById: _cf,
          dataValidationsById: _dv,
          ...rest
        } = value
        return rest
      }
      top.fields(path, scalars(previousTop), scalars(nextTop))
    }

    // Chunks: each gets the rows placed in it, their formats and cells.
    const home = placements.get(sheet)
    if (!home) continue
    const nextByChunk = new Map<ChunkId, ChunkContent>()
    for (const chunkId of sheet.chunkOrder) nextByChunk.set(chunkId, chunkSkeleton(id))
    for (const rowId of sheet.rowOrder) {
      const content = nextByChunk.get(home.get(rowId)!)!
      content.rowOrder.push(rowId)
      const row = sheet.rowsById[rowId]
      if (row) content.rowsById[rowId] = row
    }
    for (const [key, cell] of Object.entries(sheet.cellsById)) {
      const chunkId = home.get(key.split(':')[0]!)
      if (chunkId) nextByChunk.get(chunkId)!.cellsById[key as CellKey] = cell
    }
    for (const [chunkId, content] of nextByChunk) {
      const old = stored.chunks.get(chunkId) ?? chunkSkeleton(id)
      const patch = chunkPatch(chunkId)
      patch.list(['rowOrder'], old.rowOrder, content.rowOrder)
      patch.map(['rowsById'], old.rowsById, content.rowsById, true)
      patch.map(['cellsById'], old.cellsById, content.cellsById, true)
    }
    // Chunks of this sheet the order no longer lists keep their content
    // (another replica may still write to them); they are just not loaded.
  }

  const chunks = new Map<ChunkId, JsonPatchBody>()
  for (const [chunkId, patch] of chunkPatches) if (!patch.empty) chunks.set(chunkId, patch.body())
  return {
    top: top.body(),
    createdChunks,
    chunks,
    changed: !top.empty || chunks.size > 0 || createdChunks.size > 0,
  }
}
