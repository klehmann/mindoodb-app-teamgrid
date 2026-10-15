// Opening a TeamGrid 1.x workbook opens a copy of it in the current format;
// the original stays as it is. The copy records where it came from
// (`copiedFrom`), so the next open finds it instead of copying again, and
// the workbook list shows the copy instead of the original.
import { convertLegacyWorkbook, legacyWorkbookOf } from '../model/legacy'
import { TEAMGRID_FORM } from '../model/schema'
import { emptyStoredWorkbook, writesFor } from '../model/sync'
import type { HavenConnection } from './connection'
import { createWorkbook } from './store'

/** The id to open for `id`: itself, or the current-format copy of a 1.x workbook (made on first open). */
export async function workbookToOpen(haven: HavenConnection, id: string): Promise<string> {
  const document = await haven.database.documents.get(id)
  const data = (document?.data ?? {}) as Record<string, unknown>
  const legacy = legacyWorkbookOf(data)
  if (!legacy) return id
  const existing = await findCopy(haven, id)
  if (existing) return existing
  if (!haven.canWrite) throw new Error('This TeamGrid 1 workbook can only be opened by someone who may create workbooks.')
  const workbook = convertLegacyWorkbook(legacy)
  const tags = Array.isArray(data.tags) ? data.tags.filter((tag): tag is string => typeof tag === 'string') : []
  return createWorkbook(
    haven,
    typeof data.subject === 'string' ? data.subject : '',
    workbook,
    writesFor(emptyStoredWorkbook(), workbook, workbook.worksheetOrder),
    { tags, istemplate: data.istemplate === true, copiedFrom: id },
  )
}

async function findCopy(haven: HavenConnection, legacyId: string): Promise<string | undefined> {
  let cursor: string | null = null
  do {
    const page = await haven.database.documents.list({
      cursor,
      limit: 200,
      filter: { form: TEAMGRID_FORM, copiedFrom: legacyId },
      fields: ['form', 'copiedFrom'],
    })
    const hit = page.items.find((item) => item.data?.copiedFrom === legacyId)
    if (hit) return hit.id
    cursor = page.items.length > 0 ? page.nextCursor : null
  } while (cursor)
  return undefined
}
