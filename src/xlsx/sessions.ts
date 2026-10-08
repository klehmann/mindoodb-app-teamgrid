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
import type { HavenConnection } from '../haven/connection'
import {
  createWorkbook,
  loadWorkbook,
  readAttachmentBase64,
  writeWorkbook,
  type LoadedWorkbook,
} from '../haven/store'
import { workbookToXlsx } from '../model/export'
import { readWholeWorkbook } from '../model/sidecar-read'
import { emptyStoredWorkbook, newWorkbook, syncWorkbook, visualFileKey, type WorkbookWrites } from '../model/sync'
import type { XlsxEngine } from './engine'
import { saveWorkbookInBrowser } from './browser-save'
import { loadXlsxEngine } from './engine'

export interface StoredDocument {
  haven: HavenConnection
  loaded: LoadedWorkbook
  /** File sheet ids of this session → stored sheet ids. */
  storedIdByFileId: Map<string, string>
  /** This session's visuals: file visual → stored id, and per stored sheet the file order. */
  visuals: { byFileKey: Map<string, string>; orderBySheet: Map<string, string[]> }
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
  document?: Omit<StoredDocument, 'storedIdByFileId' | 'visuals'> & { visualOrder: Map<string, string[]> },
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
      Object.values(document.loaded.stored.workbook.worksheetsById)
        .filter((sheet) => !sheet.deletedAt)
        .map((sheet) => [sheet.name, sheet.id]),
    )
    const storedIdByFileId = new Map(opened.sheets.map((sheet) => [sheet.id, storedIdByName.get(sheet.name)!]))
    // The export wrote each sheet's visuals in this order; the file lists them the same way.
    const byFileKey = new Map<string, string>()
    for (const [fileSheetId, storedSheetId] of storedIdByFileId) {
      const ids = document.visualOrder.get(storedSheetId) ?? []
      const fileVisuals = opened.visuals.filter((visual) => visual.sheetId === fileSheetId)
      if (fileVisuals.length !== ids.length) continue
      fileVisuals.forEach((visual, index) => byFileKey.set(visualFileKey(visual), ids[index]!))
    }
    session.document = {
      haven: document.haven,
      loaded: document.loaded,
      storedIdByFileId,
      visuals: { byFileKey, orderBySheet: document.visualOrder },
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
  return (await openLoaded(haven, await loadWorkbook(haven, id), await loadXlsxEngine(), locale)).file
}

async function openLoaded(haven: HavenConnection, loaded: LoadedWorkbook, engine: XlsxEngine, locale: string) {
  const { bytes, visualOrder } = await workbookToXlsx(loaded.stored.workbook, engine, {
    loadImage: (attachment) => readAttachmentBase64(haven, loaded.id, attachment),
  })
  return openBytes(bytes, loaded.subject, locale, { haven, loaded, visualOrder })
}

/** Reads the bytes of pictures a save or import adds, while the file's session is open. */
async function readNewImages(engine: XlsxEngine, sessionId: string, writes: WorkbookWrites) {
  for (const image of writes.newImages) {
    image.bytes = await engine.readMediaBytes({ sessionId, visualId: image.fileVisualId })
  }
}

/** Imports an xlsx as a new stored workbook and returns the new document id. */
export async function importXlsxAsDocument(haven: HavenConnection, bytes: Uint8Array, subject: string, locale: string) {
  const engine = await loadXlsxEngine()
  const opened = (await engine.open(bytes, locale)) as WorkbookFile
  try {
    const { next, writes } = syncWorkbook({
      stored: emptyStoredWorkbook(),
      storedIdByFileId: new Map(),
      read: await readWholeWorkbook(engine, opened),
    })
    await readNewImages(engine, opened.sessionId, writes)
    return await createWorkbook(haven, subject, next, writes)
  } finally {
    await engine.close(opened.sessionId)
  }
}

/** A new workbook with a template's content (pictures included) and fresh ids. */
export async function createFromTemplate(haven: HavenConnection, templateId: string, subject: string, locale: string) {
  const template = await loadWorkbook(haven, templateId)
  const { bytes } = await workbookToXlsx(template.stored.workbook, await loadXlsxEngine(), {
    loadImage: (attachment) => readAttachmentBase64(haven, templateId, attachment),
  })
  return importXlsxAsDocument(haven, bytes, subject, locale)
}

export async function createEmptyDocument(haven: HavenConnection, subject: string, sheetName: string) {
  const workbook = newWorkbook(sheetName)
  return createWorkbook(haven, subject, workbook, {
    newImages: [],
    top: { set: [], unset: [], listInsert: [], listDelete: [] },
    createdChunks: new Map(),
    chunks: new Map(),
    changed: true,
  })
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
      stored: document.loaded.stored,
      storedIdByFileId: document.storedIdByFileId,
      request,
      read: await readWholeWorkbook(engine, saved),
      visuals: document.visuals,
    })
    await readNewImages(engine, saved.sessionId, sync.writes)
  } finally {
    await engine.close(saved.sessionId)
  }
  if (sync.writes.changed) await writeWorkbook(document.haven, document.loaded, sync.writes)
  // Continue on the merged state, which includes other people's changes.
  const loaded = await loadWorkbook(document.haven, document.loaded.id)
  const reopened = await openLoaded(document.haven, loaded, engine, locale)
  await closeSession(request.sessionId)
  return { file: reopened.file, touchedEntries, stored: { changed: sync.writes.changed } }
}

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', bytes as BufferSource)
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, '0')).join('')
}
