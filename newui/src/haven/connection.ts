// Haven host bridge: the one place that talks to the MindooDB App SDK.
// Outside a Haven frame (plain `pnpm dev`) there is no host, and the editor
// falls back to opening and downloading xlsx files.
import {
  createMindooDBAppBridge,
  type MindooDBAppDatabase,
  type MindooDBAppDocument,
  type MindooDBAppLaunchContext,
  type MindooDBAppSession,
} from 'mindoodb-app-sdk'

import { TEAMGRID_FORM, TEAMGRID_SCHEMA_VERSION, type Workbook } from '../model/schema'

export interface WorkbookSummary {
  id: string
  subject: string
  updatedAt?: string
}

export interface HavenConnection {
  session: MindooDBAppSession
  context: MindooDBAppLaunchContext
  database: MindooDBAppDatabase
  canWrite: boolean
}

let connection: Promise<HavenConnection | null> | undefined

function insideHost(): boolean {
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
    return { session, context, database, canWrite: info?.capabilities.includes('update') ?? false }
  })()
  return connection
}

export async function listWorkbooks(haven: HavenConnection): Promise<WorkbookSummary[]> {
  const workbooks: WorkbookSummary[] = []
  let cursor: string | null = null
  do {
    const page = await haven.database.documents.list({
      cursor,
      limit: 200,
      filter: { form: TEAMGRID_FORM },
      fields: ['subject', 'form'],
    })
    for (const item of page.items) {
      if (item.data?.form !== TEAMGRID_FORM) continue
      workbooks.push({
        id: item.id,
        subject: typeof item.data.subject === 'string' && item.data.subject ? item.data.subject : item.id,
        ...(item.updatedAt ? { updatedAt: item.updatedAt } : {}),
      })
    }
    cursor = page.items.length > 0 ? page.nextCursor : null
  } while (cursor)
  return workbooks.sort((left, right) => left.subject.localeCompare(right.subject))
}

/** The top-document part of a stored workbook (sheets without their chunk fields). */
export function topWorkbookOf(document: MindooDBAppDocument): Workbook {
  const teamgrid = document.data.teamgrid as { schemaVersion?: number; workbook?: Workbook } | undefined
  if (!teamgrid?.workbook || teamgrid.schemaVersion !== TEAMGRID_SCHEMA_VERSION) {
    throw new Error('This document is not a TeamGrid workbook of this version.')
  }
  return teamgrid.workbook
}
