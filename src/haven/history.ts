// Older states of a workbook, read-only.
//
// A workbook is a top document plus row-block (chunk) documents, and a save
// writes only the documents it changed, so the workbook's timeline is the
// merge of all their timelines: the writes of one save land within a moment
// of each other by the same person and form one revision. A revision is read
// back as every document's state at the revision's time.
import type { MindooDBAppDocumentHistoryEntry } from 'mindoodb-app-sdk'

import type { Workbook } from '../model/schema'
import { assembleWorkbook, type ChunkContent } from '../model/sync'
import { topWorkbookOf, type HavenConnection } from './connection'
import { chunkContentOf, type LoadedWorkbook } from './store'

/** Writes of one save land within this time of its last write (ms). */
const SAVE_WINDOW_MS = 1_500

export interface WorkbookRevision {
  /** The revision's state is every document as of this time. */
  timestamp: number
  author?: string
  current: boolean
}

export function canBrowseHistory(haven: HavenConnection): boolean {
  const database = haven.context.databases.find((entry) => entry.id === haven.databaseId)
  return database?.capabilities.includes('history') ?? false
}

/** The workbook's revisions, newest first. */
export async function listRevisions(haven: HavenConnection, loaded: LoadedWorkbook): Promise<WorkbookRevision[]> {
  const ids = [loaded.id, ...loaded.stored.chunks.keys()]
  return revisionsOf((await Promise.all(ids.map((id) => historyOf(haven, id)))).flat())
}

/** Top-document writes a sheet's timeline looks at, newest first (each is read back). */
const SHEET_TOP_ENTRIES = 300

/**
 * One sheet's revisions, newest first: the writes of its row blocks, and of
 * the top document where they changed this sheet (columns, merges, formats,
 * name, …), not the ones that only changed other sheets.
 */
export async function listSheetRevisions(
  haven: HavenConnection,
  loaded: LoadedWorkbook,
  sheetId: string,
): Promise<WorkbookRevision[]> {
  const chunkIds = loaded.stored.workbook.worksheetsById[sheetId]?.chunkOrder ?? []
  const [top, ...chunks] = await Promise.all([loaded.id, ...chunkIds].map((id) => historyOf(haven, id)))
  const topEntries = top!
    .filter((entry) => !entry.isDeleted)
    .sort((left, right) => right.timestamp - left.timestamp)
    .slice(0, SHEET_TOP_ENTRIES)
  const sheetParts = await Promise.all(
    topEntries.map(async (entry) => {
      const state = await haven.database.documents.getAtTimestamp(loaded.id, entry.timestamp).catch(() => null)
      if (state?.state !== 'exists' || !state.data) return undefined
      try {
        const sheet = topWorkbookOf({ data: state.data }).worksheetsById[sheetId]
        return sheet && !sheet.deletedAt ? JSON.stringify(sheet) : undefined
      } catch {
        return undefined
      }
    }),
  )
  // A top write counts when this sheet's part differs from the write before it (the next older one).
  const changedSheet = topEntries.filter((_, index) => {
    const part = sheetParts[index]
    return part !== undefined && part !== sheetParts[index + 1]
  })
  return revisionsOf([...changedSheet, ...chunks.flat()])
}

function historyOf(haven: HavenConnection, id: string): Promise<MindooDBAppDocumentHistoryEntry[]> {
  return haven.database.documents.listHistory(id).catch((): MindooDBAppDocumentHistoryEntry[] => [])
}

/** Merges document timelines into revisions: the writes of one save count once. */
function revisionsOf(timelines: MindooDBAppDocumentHistoryEntry[]): WorkbookRevision[] {
  const entries = timelines
    .filter((entry) => !entry.isDeleted)
    .sort((left, right) => right.timestamp - left.timestamp)
  const revisions: (WorkbookRevision & { publicKey: string })[] = []
  for (const entry of entries) {
    const last = revisions[revisions.length - 1]
    if (last && last.publicKey === entry.publicKey && last.timestamp - entry.timestamp <= SAVE_WINDOW_MS) continue
    revisions.push({
      timestamp: entry.timestamp,
      publicKey: entry.publicKey,
      ...(entry.identityLabel ? { author: entry.identityLabel } : {}),
      current: revisions.length === 0,
    })
  }
  return revisions.map(({ timestamp, author, current }) => ({ timestamp, current, ...(author ? { author } : {}) }))
}

/** The workbook as it was at `timestamp`: the top document and its chunks at that time. */
export async function loadWorkbookAt(
  haven: HavenConnection,
  id: string,
  timestamp: number,
): Promise<{ subject: string; workbook: Workbook }> {
  const documents = haven.database.documents
  const top = await documents.getAtTimestamp(id, timestamp)
  if (top.state !== 'exists' || !top.data) throw new Error('The workbook did not exist at that time.')
  const workbook = topWorkbookOf({ data: top.data })
  const chunkIds = new Set(Object.values(workbook.worksheetsById).flatMap((sheet) => sheet.chunkOrder ?? []))
  const chunks = new Map<string, ChunkContent>()
  await Promise.all(
    [...chunkIds].map(async (chunkId) => {
      const chunk = await documents.getAtTimestamp(chunkId, timestamp).catch(() => null)
      if (chunk?.state === 'exists' && chunk.data) chunks.set(chunkId, chunkContentOf({ data: chunk.data }))
    }),
  )
  return {
    subject: typeof top.data.subject === 'string' ? top.data.subject : id,
    workbook: assembleWorkbook(workbook, chunks).workbook,
  }
}

/**
 * The workbook as it is, with one sheet as it was at `timestamp` (its columns,
 * formats and rows then). Throws when the sheet did not exist at that time.
 */
export async function loadSheetAt(
  haven: HavenConnection,
  current: LoadedWorkbook,
  sheetId: string,
  timestamp: number,
): Promise<{ subject: string; workbook: Workbook }> {
  const then = await loadWorkbookAt(haven, current.id, timestamp)
  const sheet = then.workbook.worksheetsById[sheetId]
  if (!sheet || sheet.deletedAt) throw new Error('The sheet did not exist at that time.')
  const now = current.stored.workbook
  return {
    subject: current.subject,
    workbook: { ...now, worksheetsById: { ...now.worksheetsById, [sheetId]: sheet } },
  }
}
