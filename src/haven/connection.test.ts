// Workbook lists come from the summary buffer (documents.query), with a walk
// of the database only for hosts that cannot query.
import { describe, expect, it } from 'vitest'

import { LEGACY } from '../testing/legacy-sample'
import { equalsFilter, listWorkbooks, queryDocuments, type HavenConnection } from './connection'
import { workbookToOpen } from './legacy-copy'

interface Doc {
  id: string
  data: Record<string, unknown>
  lastModified: number
}

const DOCS: Doc[] = [
  { id: 'b', data: { form: 'teamgrid-next', subject: 'Budget', istemplate: false }, lastModified: 2_000 },
  { id: 'a', data: { form: 'teamgrid-next', subject: 'Angebot', istemplate: true }, lastModified: 1_000 },
  { id: 'old1', data: { form: 'teamgrid', subject: 'Umsatz 2025' }, lastModified: 500 },
  { id: 'old2', data: { form: 'teamgrid', subject: 'Plan 2024' }, lastModified: 400 },
  { id: 'copy', data: { form: 'teamgrid-next', subject: 'Plan 2024', copiedFrom: 'old2' }, lastModified: 3_000 },
  { id: 'note', data: { form: 'note', subject: 'Something else' }, lastModified: 100 },
]

/** The filter's field/value pairs back from the formula source equalsFilter writes. */
function parse(filter: string): Record<string, string> {
  return Object.fromEntries([...filter.matchAll(/v\.field\(("[^"]*")\), ("[^"]*")\)/g)].map((m) => [JSON.parse(m[1]!), JSON.parse(m[2]!)]))
}

function haven(
  options: { query?: 'full' | 'rebuilding' | 'missing'; scanned?: string[]; docs?: Doc[] } = {},
): HavenConnection {
  const docs = options.docs ?? DOCS
  const matching = (fields: Record<string, unknown>) =>
    docs.filter((doc) => Object.entries(fields).every(([field, value]) => doc.data[field] === value))
  return {
    database: {
      documents: {
        async query(query: { filter: string; limit: number; offset?: number }) {
          if (options.query === 'missing') throw new Error('unknown method documents.query')
          const all = matching(parse(query.filter))
          const page = all.slice(query.offset ?? 0, (query.offset ?? 0) + query.limit)
          return {
            rows: page.map((doc) => ({ docId: doc.id, fields: doc.data, lastModified: doc.lastModified })),
            total: all.length,
            coverage: options.query ?? 'full',
          }
        },
        async list(query: { filter: Record<string, unknown> }) {
          options.scanned?.push(JSON.stringify(query.filter))
          return {
            items: docs.map((doc) => ({ id: doc.id, data: doc.data, updatedAt: new Date(doc.lastModified).toISOString() })),
            nextCursor: null,
          }
        },
        async get(id: string) {
          const doc = docs.find((candidate) => candidate.id === id)
          return doc ? { id, data: doc.data, heads: [] } : null
        },
        async create() {
          throw new Error('no copy may be created here')
        },
      },
    },
  } as unknown as HavenConnection
}

describe('workbook lists', () => {
  it('writes equality filters as formula source', () => {
    expect(equalsFilter({ form: 'teamgrid-next' })).toBe('v.eq(v.field("form"), "teamgrid-next")')
    expect(equalsFilter({ form: 'x', copiedFrom: 'a"b' })).toBe('v.and(v.eq(v.field("form"), "x"), v.eq(v.field("copiedFrom"), "a\\"b"))')
  })

  it('lists current workbooks and 1.x workbooks without a copy, by title, from the summary', async () => {
    const scanned: string[] = []
    const list = await listWorkbooks(haven({ scanned }))
    expect(list.workbooks.map((workbook) => [workbook.id, workbook.subject, workbook.istemplate, workbook.legacy ?? false])).toEqual([
      ['a', 'Angebot', true, false],
      ['b', 'Budget', false, false],
      ['copy', 'Plan 2024', false, false],
      ['old1', 'Umsatz 2025', false, true],
    ])
    expect(list.workbooks.find((workbook) => workbook.id === 'b')?.updatedAt).toBe(new Date(2_000).toISOString())
    expect(list.incomplete).toBe(false)
    expect(scanned).toEqual([])
  })

  it('says when the summary is still being built', async () => {
    expect((await listWorkbooks(haven({ query: 'rebuilding' }))).incomplete).toBe(true)
  })

  it('walks the database when the host cannot query', async () => {
    const scanned: string[] = []
    const list = await listWorkbooks(haven({ query: 'missing', scanned }))
    expect(list.workbooks.map((workbook) => workbook.id)).toEqual(['a', 'b', 'copy', 'old1'])
    expect(scanned.length).toBe(2)
  })

  it('a 1.x workbook opens as the copy the summary knows', async () => {
    const docs: Doc[] = [
      { id: 'legacy', data: LEGACY as Record<string, unknown>, lastModified: 1 },
      { id: 'its-copy', data: { form: 'teamgrid-next', subject: 'Umsatz 2025', copiedFrom: 'legacy' }, lastModified: 2 },
    ]
    expect(await workbookToOpen(haven({ docs }), 'legacy')).toBe('its-copy')
    // While the summary is rebuilt and does not know the copy yet, the database is walked.
    const scanned: string[] = []
    const unknown = haven({ docs, query: 'rebuilding', scanned })
    ;(unknown.database.documents as unknown as { query: unknown }).query = async () => ({ rows: [], total: 0, coverage: 'rebuilding' })
    expect(await workbookToOpen(unknown, 'legacy')).toBe('its-copy')
    expect(scanned).toEqual([JSON.stringify({ form: 'teamgrid-next', copiedFrom: 'legacy' })])
  })

  it('stops at the limit', async () => {
    const { items } = await queryDocuments(haven(), { form: 'teamgrid-next' }, { limit: 1 })
    expect(items.map((item) => item.id)).toEqual(['b'])
  })
})
