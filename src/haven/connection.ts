// Haven host bridge: the one place that talks to the MindooDB App SDK.
// Outside a Haven frame (plain `pnpm dev`) there is no host, and the editor
// falls back to opening and downloading xlsx files.
import {
  MindooDBAppValue,
  createMindooDBAppBridge,
  type MindooDBAppDatabase,
  type MindooDBAppDocument,
  type MindooDBAppLaunchContext,
  type MindooDBAppSession,
} from 'mindoodb-app-sdk'

import { LEGACY_FORM } from '../model/legacy'
import { TEAMGRID_FORM, TEAMGRID_SCHEMA_VERSION, type Workbook } from '../model/schema'

export interface WorkbookSummary {
  id: string
  subject: string
  istemplate: boolean
  updatedAt?: string
  /** A TeamGrid 1.x workbook: opening it opens a copy in the current format. */
  legacy?: boolean
}

export interface HavenConnection {
  session: MindooDBAppSession
  context: MindooDBAppLaunchContext
  database: MindooDBAppDatabase
  databaseId: string
  /** False in Haven's time travel, which opens every database read-only. */
  canWrite: boolean
  /** May create workbooks (the `create` capability; never in Haven's time travel). */
  canCreate: boolean
}

let connection: Promise<HavenConnection | null> | undefined

function insideHost(): boolean {
  // `?standalone` runs without Haven even inside a frame (tests, demos).
  if (new URLSearchParams(window.location.search).has('standalone')) return false
  return window.parent !== window || window.opener != null
}

export function connectHaven(): Promise<HavenConnection | null> {
  connection ??= (async () => {
    if (!insideHost()) return null
    const session = await createMindooDBAppBridge().connect()
    const context = await session.getLaunchContext()
    const databaseId = context.preferredDatabaseId ?? context.databases[0]?.id
    if (!databaseId) throw new Error('Haven granted this app no database.')
    const database = await session.openDatabase(databaseId)
    const info = context.databases.find((entry) => entry.id === databaseId)
    const canWrite = !context.timeTravelDate && (info?.capabilities.includes('update') ?? false)
    const canCreate = !context.timeTravelDate && (info?.capabilities.includes('create') ?? false)
    return { session, context, database, databaseId, canWrite, canCreate }
  })()
  return connection
}

/** Fields a workbook list shows; all are top-level scalars, so the summary buffer covers them. */
const LIST_FIELDS = ['subject', 'form', 'istemplate', 'copiedFrom']
/** Rows per `documents.query` page. */
const QUERY_PAGE = 200

interface FormItem {
  id: string
  data: Record<string, unknown>
  updatedAt?: string
}

/** A `documents.query` filter: every field equals its value (formula source, parsed by the SDK). */
export function equalsFilter(fields: Record<string, string>): string {
  const terms = Object.entries(fields).map(([field, value]) => `v.eq(v.field(${JSON.stringify(field)}), ${JSON.stringify(value)})`)
  return terms.length === 1 ? terms[0]! : `v.and(${terms.join(', ')})`
}

/**
 * Documents matching `fields`, from the database's summary buffer
 * (`documents.query`): no document is loaded. `complete` is false while the
 * summary is still being built, when the answer may miss documents
 * (MindooDB best practices §4.3). A host without query support is answered
 * by walking the database instead.
 */
export async function queryDocuments(
  haven: HavenConnection,
  fields: Record<string, string>,
  options: { limit?: number } = {},
): Promise<{ items: FormItem[]; complete: boolean }> {
  try {
    const items: FormItem[] = []
    let complete = true
    for (let offset = 0; ; ) {
      const page = await haven.database.documents.query({
        filter: equalsFilter(fields),
        fields: LIST_FIELDS,
        limit: Math.min(QUERY_PAGE, options.limit ?? QUERY_PAGE),
        offset,
      })
      if (page.coverage === 'rebuilding') complete = false
      for (const row of page.rows) {
        items.push({ id: row.docId, data: row.fields, updatedAt: new Date(row.lastModified).toISOString() })
      }
      offset += page.rows.length
      if (page.rows.length === 0 || offset >= page.total || (options.limit !== undefined && offset >= options.limit)) break
    }
    return { items, complete }
  } catch (error) {
    console.warn('[newui] documents.query failed; walking the database instead', error)
    return { items: await scanDocuments(haven, fields), complete: true }
  }
}

/** The same by walking the database (loads every document): hosts without `documents.query`. */
export async function scanDocuments(haven: HavenConnection, fields: Record<string, string>): Promise<FormItem[]> {
  const items: FormItem[] = []
  let cursor: string | null = null
  do {
    const page = await haven.database.documents.list({
      cursor,
      limit: 200,
      filter: fields,
      fields: LIST_FIELDS,
    })
    for (const item of page.items) {
      if (Object.entries(fields).every(([field, value]) => item.data?.[field] === value)) {
        items.push({ id: item.id, data: item.data ?? {}, ...(item.updatedAt ? { updatedAt: item.updatedAt } : {}) })
      }
    }
    cursor = page.items.length > 0 ? page.nextCursor : null
  } while (cursor)
  return items
}

export interface WorkbookList {
  workbooks: WorkbookSummary[]
  /** The summary buffer is still being built: workbooks may be missing (show a hint, not an error). */
  incomplete: boolean
}

/**
 * The workbooks to offer: the current ones, plus TeamGrid 1.x workbooks that
 * have no copy yet (opening one copies it, see legacy-copy.ts).
 */
export async function listWorkbooks(haven: HavenConnection): Promise<WorkbookList> {
  const [currentList, legacyList] = await Promise.all([
    queryDocuments(haven, { form: TEAMGRID_FORM }),
    queryDocuments(haven, { form: LEGACY_FORM }),
  ])
  const current = currentList.items
  const legacy = legacyList.items
  const copied = new Set(current.map((item) => item.data.copiedFrom).filter((id): id is string => typeof id === 'string'))
  const summary = (item: (typeof current)[number], isLegacy: boolean): WorkbookSummary => ({
    id: item.id,
    subject: typeof item.data.subject === 'string' && item.data.subject ? item.data.subject : item.id,
    istemplate: item.data.istemplate === true,
    ...(item.updatedAt ? { updatedAt: item.updatedAt } : {}),
    ...(isLegacy ? { legacy: true } : {}),
  })
  return {
    workbooks: [
      ...current.map((item) => summary(item, false)),
      ...legacy.filter((item) => !copied.has(item.id)).map((item) => summary(item, true)),
    ].sort((left, right) => left.subject.localeCompare(right.subject)),
    incomplete: !currentList.complete || !legacyList.complete,
  }
}

export interface WorkbookProperties {
  subject: string
  tags: string[]
  istemplate: boolean
}

export async function readProperties(haven: HavenConnection, id: string): Promise<WorkbookProperties> {
  const document = await haven.database.documents.get(id)
  const data = document?.data ?? {}
  return {
    subject: typeof data.subject === 'string' ? data.subject : '',
    tags: Array.isArray(data.tags) ? data.tags.filter((tag): tag is string => typeof tag === 'string') : [],
    istemplate: data.istemplate === true,
  }
}

/** Writes title, tags and template flag; the tag list is replaced as a whole. */
export async function writeProperties(haven: HavenConnection, id: string, properties: WorkbookProperties) {
  const document = await haven.database.documents.get(id)
  await haven.database.documents.update(id, {
    json: {
      ...(document?.heads?.length ? { baseHeads: document.heads } : {}),
      set: [
        { path: ['subject'], value: MindooDBAppValue.atomic(properties.subject) },
        { path: ['tags'], value: properties.tags.map((tag) => MindooDBAppValue.atomic(tag)) },
        { path: ['istemplate'], value: properties.istemplate },
      ],
    },
  })
}

/** The top-document part of a stored workbook (sheets without their chunk fields). */
export function topWorkbookOf(document: Pick<MindooDBAppDocument, 'data'>): Workbook {
  const teamgrid = document.data.teamgrid as { schemaVersion?: number; workbook?: Workbook } | undefined
  if (!teamgrid?.workbook || teamgrid.schemaVersion !== TEAMGRID_SCHEMA_VERSION) {
    throw new Error('This document is not a TeamGrid workbook of this version.')
  }
  return teamgrid.workbook
}
