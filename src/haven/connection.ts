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
    return { session, context, database, databaseId, canWrite }
  })()
  return connection
}

/** Documents of one form: id, the listed fields, last change. */
async function listForm(haven: HavenConnection, form: string) {
  const items: { id: string; data: Record<string, unknown>; updatedAt?: string }[] = []
  let cursor: string | null = null
  do {
    const page = await haven.database.documents.list({
      cursor,
      limit: 200,
      filter: { form },
      fields: ['subject', 'form', 'istemplate', 'copiedFrom'],
    })
    for (const item of page.items) {
      if (item.data?.form === form) items.push({ id: item.id, data: item.data, ...(item.updatedAt ? { updatedAt: item.updatedAt } : {}) })
    }
    cursor = page.items.length > 0 ? page.nextCursor : null
  } while (cursor)
  return items
}

/**
 * The workbooks to offer: the current ones, plus TeamGrid 1.x workbooks that
 * have no copy yet (opening one copies it, see legacy-copy.ts).
 */
export async function listWorkbooks(haven: HavenConnection): Promise<WorkbookSummary[]> {
  const [current, legacy] = await Promise.all([listForm(haven, TEAMGRID_FORM), listForm(haven, LEGACY_FORM)])
  const copied = new Set(current.map((item) => item.data.copiedFrom).filter((id): id is string => typeof id === 'string'))
  const summary = (item: (typeof current)[number], isLegacy: boolean): WorkbookSummary => ({
    id: item.id,
    subject: typeof item.data.subject === 'string' && item.data.subject ? item.data.subject : item.id,
    istemplate: item.data.istemplate === true,
    ...(item.updatedAt ? { updatedAt: item.updatedAt } : {}),
    ...(isLegacy ? { legacy: true } : {}),
  })
  return [
    ...current.map((item) => summary(item, false)),
    ...legacy.filter((item) => !copied.has(item.id)).map((item) => summary(item, true)),
  ].sort((left, right) => left.subject.localeCompare(right.subject))
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
