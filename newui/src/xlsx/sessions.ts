// Open workbook sessions of the browser shim: the source bytes each session
// was opened from (the save base, like the desktop app's snapshot file) and
// the sheet id → name map the save mapping resolves through.
import JSZip from 'jszip'

import { parsePivotDefinition } from '@genoffice/xlsx-gateway/gateway/xlsx-pivot'

import {
  workbookFileSchema,
  type WorkbookFile,
  type WorkbookSaveRequest,
} from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { saveWorkbookInBrowser } from './browser-save'
import { loadXlsxEngine } from './engine'

interface Session {
  name: string
  bytes: Uint8Array
  sheetNames: Map<string, string>
}

const sessions = new Map<string, Session>()

export async function openWorkbookBytes(bytes: Uint8Array, name: string, locale: string): Promise<WorkbookFile> {
  const engine = await loadXlsxEngine()
  const opened = (await engine.open(bytes, locale)) as Omit<WorkbookFile, 'sha256' | 'readOnly'>
  sessions.set(opened.sessionId, {
    name,
    bytes,
    sheetNames: new Map(opened.sheets.map((sheet) => [sheet.id, sheet.name])),
  })
  return workbookFileSchema.parse({
    ...opened,
    name,
    sha256: await sha256Hex(bytes),
    fileBytes: bytes.byteLength,
    readOnly: false,
    // No path on disk: the first save always asks where to put the file.
    needsSaveAs: true,
  })
}

export function sessionFor(sessionId: string): Session {
  const session = sessions.get(sessionId)
  if (!session) throw new Error('Unknown workbook session.')
  return session
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

/**
 * Saves a session's pending edits: patches its source bytes, then reopens the
 * result as the new session (the desktop app re-bases on the saved file the
 * same way) and returns the bytes for the caller to store or download.
 */
export async function saveSession(request: WorkbookSaveRequest, locale: string) {
  const session = sessionFor(request.sessionId)
  const { bytes, touchedEntries } = await saveWorkbookInBrowser(session.bytes, session.sheetNames, request)
  const file = await openWorkbookBytes(bytes, session.name, locale)
  await closeSession(request.sessionId)
  return { bytes, file, touchedEntries }
}

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', bytes as BufferSource)
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, '0')).join('')
}
