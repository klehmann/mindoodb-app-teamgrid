// Open workbook sessions of the browser shim.
//
// Every session is an xlsx in memory, opened through the WASM sidecar: the
// editor reads its cells from there and each save patches it (the desktop
// app's snapshot file). A session opened from MindooDB additionally knows its
// document: on save, the patched xlsx is re-read and its changes go to the
// document as a JSON patch (model/sync.ts); the editor then continues on the
// merged document, which includes concurrent changes from other people.
import JSZip from 'jszip'

import { parsePivotDefinition } from '@genoffice/xlsx-gateway/gateway/xlsx-pivot'

import {
  workbookFileSchema,
  type WorkbookFile,
  type WorkbookSaveRequest,
} from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import {
  createWorkbookDocument,
  updateWorkbookDocument,
  workbookOf,
  type HavenConnection,
} from '../haven/connection'
import { workbookToXlsx } from '../model/export'
import { createId, type Workbook } from '../model/schema'
import { readWholeWorkbook } from '../model/sidecar-read'
import { importWorkbook, syncWorkbook } from '../model/sync'
import { saveWorkbookInBrowser } from './browser-save'
import { loadXlsxEngine } from './engine'

export interface StoredDocument {
  haven: HavenConnection
  id: string
  subject: string
  heads: string[]
  workbook: Workbook
  /** File sheet ids of this session → stored sheet ids. */
  storedIdByFileId: Map<string, string>
}

interface Session {
  name: string
  bytes: Uint8Array
  sheetNames: Map<string, string>
  document?: StoredDocument
}

const sessions = new Map<string, Session>()

async function openBytes(
  bytes: Uint8Array,
  name: string,
  locale: string,
  document?: Omit<StoredDocument, 'storedIdByFileId'>,
) {
  const engine = await loadXlsxEngine()
  const opened = (await engine.open(bytes, locale)) as Omit<WorkbookFile, 'sha256' | 'readOnly'>
  const session: Session = {
    name,
    bytes,
    sheetNames: new Map(opened.sheets.map((sheet) => [sheet.id, sheet.name])),
  }
  if (document) {
    const storedIdByName = new Map(
      Object.values(document.workbook.worksheetsById)
        .filter((sheet) => !sheet.deletedAt)
        .map((sheet) => [sheet.name, sheet.id]),
    )
    session.document = {
      ...document,
      storedIdByFileId: new Map(opened.sheets.map((sheet) => [sheet.id, storedIdByName.get(sheet.name)!])),
    }
  }
  sessions.set(opened.sessionId, session)
  const file = workbookFileSchema.parse({
    ...opened,
    name,
    sha256: await sha256Hex(bytes),
    fileBytes: bytes.byteLength,
    readOnly: document ? !document.haven.canWrite : false,
    // A stored workbook saves in place; a plain xlsx asks where to put it.
    needsSaveAs: !document,
  })
  return { file, session }
}

export async function openWorkbookBytes(bytes: Uint8Array, name: string, locale: string): Promise<WorkbookFile> {
  return (await openBytes(bytes, name, locale)).file
}

/** Opens a stored workbook: renders it to xlsx and opens that in the editor. */
export async function openStoredWorkbook(haven: HavenConnection, id: string, locale: string): Promise<WorkbookFile> {
  const document = await haven.database.documents.get(id)
  if (!document) throw new Error('The workbook no longer exists.')
  const workbook = workbookOf(document)
  const subject = typeof document.data.subject === 'string' && document.data.subject ? document.data.subject : id
  const bytes = await workbookToXlsx(workbook, await loadXlsxEngine())
  const { file } = await openBytes(bytes, subject, locale, {
    haven,
    id,
    subject,
    heads: document.heads ?? [],
    workbook,
  })
  return file
}

/** Imports an xlsx as a new stored workbook and returns the new document id. */
export async function importXlsxAsDocument(haven: HavenConnection, bytes: Uint8Array, subject: string, locale: string) {
  const engine = await loadXlsxEngine()
  const opened = (await engine.open(bytes, locale)) as WorkbookFile
  try {
    const workbook = importWorkbook(await readWholeWorkbook(engine, opened))
    return (await createWorkbookDocument(haven, subject, workbook)).id
  } finally {
    await engine.close(opened.sessionId)
  }
}

export async function createEmptyDocument(haven: HavenConnection, subject: string, sheetName: string) {
  const id = createId('sheet')
  const workbook: Workbook = {
    worksheetOrder: [id],
    worksheetsById: {
      [id]: {
        id,
        name: sheetName,
        rowOrder: [],
        columnOrder: [],
        rowsById: {},
        columnsById: {},
        cellsById: {},
        mergesById: {},
      },
    },
    stylesById: {},
  }
  return (await createWorkbookDocument(haven, subject, workbook)).id
}

export function sessionFor(sessionId: string): Session {
  const session = sessions.get(sessionId)
  if (!session) throw new Error('Unknown workbook session.')
  return session
}

export function storedDocumentOf(sessionId: string): StoredDocument | undefined {
  return sessions.get(sessionId)?.document
}

export async function closeSession(sessionId: string): Promise<void> {
  if (!sessions.delete(sessionId)) return
  await (await loadXlsxEngine()).close(sessionId)
}

export async function readPivotDefinition(sessionId: string, path: string, cachePath: string) {
  const zip = await JSZip.loadAsync(sessionFor(sessionId).bytes)
  const read = async (entry: string) => {
    const file = zip.file(entry)
    if (!file) throw new Error(`Workbook is missing ${entry}.`)
    return file.async('text')
  }
  const [pivotXml, cacheXml] = await Promise.all([read(path), read(cachePath)])
  return parsePivotDefinition(pivotXml, cacheXml)
}

export interface SaveOutcome {
  file: WorkbookFile
  touchedEntries: readonly string[]
  /** Set for a plain xlsx session: the bytes to hand to the user. */
  download?: Uint8Array
  /** Stored workbooks: whether the document changed at all. */
  stored?: { changed: boolean }
}

/**
 * Saves a session's pending edits. The session's xlsx is patched like the
 * desktop app patches its file; a stored workbook then writes the difference
 * to its document and reopens from the merged result.
 */
export async function saveSession(request: WorkbookSaveRequest, locale: string): Promise<SaveOutcome> {
  const session = sessionFor(request.sessionId)
  const { bytes, touchedEntries } = await saveWorkbookInBrowser(session.bytes, session.sheetNames, request)
  const document = session.document
  if (!document) {
    const file = await openWorkbookBytes(bytes, session.name, locale)
    await closeSession(request.sessionId)
    return { file, touchedEntries, download: bytes }
  }

  const engine = await loadXlsxEngine()
  const saved = (await engine.open(bytes, locale)) as WorkbookFile
  let sync
  try {
    sync = syncWorkbook({
      stored: document.workbook,
      storedIdByFileId: document.storedIdByFileId,
      request,
      read: await readWholeWorkbook(engine, saved),
    })
  } finally {
    await engine.close(saved.sessionId)
  }
  const merged = sync.changed
    ? await updateWorkbookDocument(document.haven, document.id, { baseHeads: document.heads, ...sync.patch })
    : await document.haven.database.documents.get(document.id)
  if (!merged) throw new Error('The workbook was deleted while saving.')
  const workbook = workbookOf(merged)
  const reopened = await openBytes(await workbookToXlsx(workbook, engine), document.subject, locale, {
    haven: document.haven,
    id: document.id,
    subject: document.subject,
    heads: merged.heads ?? [],
    workbook,
  })
  await closeSession(request.sessionId)
  return { file: reopened.file, touchedEntries, stored: { changed: sync.changed } }
}

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', bytes as BufferSource)
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, '0')).join('')
}
