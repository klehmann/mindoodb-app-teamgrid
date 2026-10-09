// Reads and writes a stored workbook: the top document and its row-block
// (chunk) documents. Every write is a JSON patch at the heads it was computed
// against (`baseHeads`), so MindooDB merges it with concurrent changes.
import type { MindooDBAppDocument } from 'mindoodb-app-sdk'

import {
  CHUNK_FIELDS,
  TEAMGRID_CHUNK_TYPE,
  TEAMGRID_FORM,
  TEAMGRID_KIND,
  TEAMGRID_SCHEMA_VERSION,
  type ChunkId,
  type Workbook,
  type Worksheet,
} from '../model/schema'
import {
  assembleWorkbook,
  chunkSkeleton,
  toStored,
  type ChunkContent,
  type JsonPatchBody,
  type StoredWorkbook,
  type WorkbookWrites,
} from '../model/sync'
import { topWorkbookOf, type HavenConnection } from './connection'

export interface LoadedWorkbook {
  id: string
  subject: string
  stored: StoredWorkbook
  /** Heads every write is based on: the top document's and each chunk's. */
  heads: Map<string, string[]>
  decryptionKeyId?: string
  /** Changefeed position when it was loaded: later changes mean someone else wrote. */
  feedCursor: string | null
}

export function chunkContentOf(document: Pick<MindooDBAppDocument, 'data'>): ChunkContent {
  const data = document.data as Partial<ChunkContent>
  return {
    sheetId: data.sheetId ?? '',
    rowOrder: Array.isArray(data.rowOrder) ? data.rowOrder : [],
    rowsById: data.rowsById ?? {},
    cellsById: data.cellsById ?? {},
  }
}

export async function loadWorkbook(haven: HavenConnection, id: string): Promise<LoadedWorkbook> {
  const { cursor: feedCursor } = await haven.database.documents.getHeadCursor()
  const top = await haven.database.documents.get(id)
  if (!top) throw new Error('The workbook no longer exists.')
  const workbook = topWorkbookOf(top)
  const heads = new Map<string, string[]>([[id, top.heads ?? []]])
  const chunkIds = new Set(Object.values(workbook.worksheetsById).flatMap((sheet) => sheet.chunkOrder ?? []))
  const chunks = new Map<ChunkId, ChunkContent>()
  await Promise.all(
    [...chunkIds].map(async (chunkId) => {
      const document = await haven.database.documents.get(chunkId)
      // A chunk another replica listed but whose document has not synced yet
      // simply is not there; its rows appear with the next reload.
      if (!document) return
      chunks.set(chunkId, chunkContentOf(document))
      heads.set(chunkId, document.heads ?? [])
    }),
  )
  return {
    id,
    subject: typeof top.data.subject === 'string' && top.data.subject ? top.data.subject : id,
    stored: assembleWorkbook(workbook, chunks),
    heads,
    feedCursor,
    ...(top.decryptionKeyId ? { decryptionKeyId: top.decryptionKeyId } : {}),
  }
}

/**
 * Whether the workbook (its top document or any of its chunks) changed since
 * it was loaded or last checked, by reading the database's changefeed.
 */
export async function changedSinceLoad(haven: HavenConnection, loaded: LoadedWorkbook): Promise<boolean> {
  let changed = false
  for (;;) {
    const page = await haven.database.documents.list({
      cursor: loaded.feedCursor,
      metadataOnly: true,
      status: 'all',
      limit: 200,
    })
    if (page.items.some((item) => loaded.heads.has(item.id))) changed = true
    if (!page.nextCursor || page.items.length === 0) break
    loaded.feedCursor = page.nextCursor
  }
  return changed
}

/** A workbook's top-document part: sheets without rows and cells. */
export function topWorkbook(workbook: Workbook): Workbook {
  const worksheetsById: Record<string, Worksheet> = {}
  for (const [id, sheet] of Object.entries(workbook.worksheetsById)) {
    const top = { ...sheet } as Partial<Worksheet>
    for (const field of CHUNK_FIELDS) delete top[field]
    worksheetsById[id] = top as Worksheet
  }
  return { ...workbook, worksheetsById }
}

function patchOf(body: JsonPatchBody, baseHeads: string[] | undefined) {
  return { ...(baseHeads && baseHeads.length > 0 ? { baseHeads } : {}), ...body }
}

/** Creates a chunk document with the same content on every replica (see chunkIdFor). */
async function createChunk(haven: HavenConnection, chunkId: ChunkId, parentId: string, sheetId: string, decryptionKeyId?: string) {
  const created = await haven.database.documents.create({
    id: chunkId,
    set: toStored({ type: TEAMGRID_CHUNK_TYPE, parentId, ...chunkSkeleton(sheetId) }) as Record<string, unknown>,
    ...(decryptionKeyId ? { decryptionKeyId } : {}),
  })
  return created.heads ?? []
}

/** An attachment of a document as base64, or undefined when it is missing. */
export async function readAttachmentBase64(haven: HavenConnection, docId: string, name: string) {
  try {
    const stream = await haven.database.attachments.openReadStream(docId, name)
    const chunks: Uint8Array[] = []
    for (let chunk = await stream.read(); chunk; chunk = await stream.read()) chunks.push(chunk)
    await stream.close()
    let binary = ''
    for (const chunk of chunks) {
      for (let index = 0; index < chunk.length; index += 0x8000) {
        binary += String.fromCharCode(...chunk.subarray(index, index + 0x8000))
      }
    }
    return btoa(binary)
  } catch (error) {
    console.warn(`[newui] picture ${name} could not be read`, error)
    return undefined
  }
}

async function writeAttachment(haven: HavenConnection, docId: string, name: string, bytes: Uint8Array, contentType: string) {
  const stream = await haven.database.attachments.openWriteStream(docId, name, contentType)
  try {
    for (let offset = 0; offset < bytes.length; offset += 256 * 1024) {
      await stream.write(bytes.subarray(offset, offset + 256 * 1024))
    }
    await stream.close()
  } catch (error) {
    await stream.abort().catch(() => {})
    throw error
  }
}

/**
 * Writes a save: new chunks first, then the chunks' patches, then new
 * pictures, then the top document, so the top document never refers to a
 * chunk or picture that does not exist.
 */
export async function writeWorkbook(haven: HavenConnection, loaded: LoadedWorkbook, writes: WorkbookWrites): Promise<void> {
  const heads = new Map(loaded.heads)
  for (const [chunkId, sheetId] of writes.createdChunks) {
    heads.set(chunkId, await createChunk(haven, chunkId, loaded.id, sheetId, loaded.decryptionKeyId))
  }
  await Promise.all(
    [...writes.chunks].map(([chunkId, body]) =>
      haven.database.documents.update(chunkId, { json: patchOf(body, heads.get(chunkId)) }),
    ),
  )
  for (const image of writes.newImages) {
    if (!image.bytes) throw new Error(`The picture ${image.attachment} has no bytes to store.`)
    await writeAttachment(haven, loaded.id, image.attachment, image.bytes, image.mediaType)
  }
  const top = writes.top
  if (top.set.length || top.unset.length || top.listInsert.length || top.listDelete.length) {
    await haven.database.documents.update(loaded.id, { json: patchOf(top, heads.get(loaded.id)) })
  }
}

/** Document fields of a new workbook beyond its title. */
export interface NewWorkbookFields {
  tags?: string[]
  istemplate?: boolean
  /** The TeamGrid 1.x document this workbook was copied from. */
  copiedFrom?: string
}

/** Creates a workbook: the top document, then its chunks with their content. */
export async function createWorkbook(
  haven: HavenConnection,
  subject: string,
  next: Workbook,
  writes: WorkbookWrites,
  fields: NewWorkbookFields = {},
) {
  const created = await haven.database.documents.create({
    set: toStored({
      form: TEAMGRID_FORM,
      kind: TEAMGRID_KIND,
      subject,
      tags: fields.tags ?? [],
      istemplate: fields.istemplate ?? false,
      ...(fields.copiedFrom ? { copiedFrom: fields.copiedFrom } : {}),
      teamgrid: { schemaVersion: TEAMGRID_SCHEMA_VERSION, workbook: topWorkbook(next) },
    }) as Record<string, unknown>,
  })
  const loaded: LoadedWorkbook = {
    id: created.id,
    subject,
    stored: assembleWorkbook(topWorkbook(next), new Map()),
    heads: new Map([[created.id, created.heads ?? []]]),
    feedCursor: null,
    ...(created.decryptionKeyId ? { decryptionKeyId: created.decryptionKeyId } : {}),
  }
  // The top document already holds everything but the chunks.
  await writeWorkbook(haven, loaded, { ...writes, top: { set: [], unset: [], listInsert: [], listDelete: [] } })
  return created.id
}
