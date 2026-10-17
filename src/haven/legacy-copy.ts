// Opening a TeamGrid 1.x workbook opens a copy of it in the current format;
// the original stays as it is. The copy records where it came from
// (`copiedFrom`), so the next open finds it instead of copying again, and
// the workbook list shows the copy instead of the original.
import { convertLegacyWorkbook, legacyWorkbookOf } from '../model/legacy'
import { TEAMGRID_FORM } from '../model/schema'
import { emptyStoredWorkbook, writesFor } from '../model/sync'
import { queryDocuments, scanDocuments, type HavenConnection } from './connection'
import { createWorkbook } from './store'

/** The id to open for `id`: itself, or the current-format copy of a 1.x workbook (made on first open). */
export async function workbookToOpen(haven: HavenConnection, id: string): Promise<string> {
  const document = await haven.database.documents.get(id)
  const data = (document?.data ?? {}) as Record<string, unknown>
  const legacy = legacyWorkbookOf(data)
  if (!legacy) return id
  const existing = await findCopy(haven, id)
  if (existing) return existing
  if (!haven.canCreate) throw new Error('This TeamGrid 1 workbook can only be opened by someone who may create workbooks.')
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
  const fields = { form: TEAMGRID_FORM, copiedFrom: legacyId }
  const { items, complete } = await queryDocuments(haven, fields, { limit: 1 })
  if (items[0]) return items[0].id
  // A summary still being built may not know the copy yet; a second copy must not be made.
  if (!complete) return (await scanDocuments(haven, fields))[0]?.id
  return undefined
}
