// A minimal MindooDB stand-in for unit tests: documents are real Automerge
// documents, and JSON patches are applied the way MindooDB applies them
// (BaseMindooDB.applyJsonPatch: `changeAt(baseHeads)` when given, operations
// in the order set, unset, listDelete, listInsert; `$mindoo` atomic values
// become ImmutableStrings). A host is one device's replica: `clone()` gives a
// second device with the same state, `syncFrom()` merges another device's
// documents in, as MindooDB's sync does. Tests use this to play two people
// editing the same base offline.
import * as A from '@automerge/automerge'
import type { MindooDBAppDocument, MindooDBAppJsonPatch } from 'mindoodb-app-sdk'

import type { HavenConnection } from '../haven/connection'

type Doc = A.Doc<Record<string, unknown>>

function hydrate(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(hydrate)
  if (value && typeof value === 'object') {
    const record = value as Record<string, unknown>
    if (record.$mindoo === 'atomic') return new A.ImmutableString(String(record.value))
    return Object.fromEntries(Object.entries(record).map(([key, entry]) => [key, hydrate(entry)]))
  }
  return value
}

function plain(value: unknown): unknown {
  if (value instanceof A.ImmutableString) return value.toString()
  if (Array.isArray(value)) return value.map(plain)
  if (value && typeof value === 'object') {
    return Object.fromEntries(Object.entries(value).map(([key, entry]) => [key, plain(entry)]))
  }
  return value
}

function parentOf(root: Record<string, unknown>, path: (string | number)[], create: boolean) {
  let parent: Record<string | number, unknown> = root
  for (const segment of path.slice(0, -1)) {
    if (parent[segment] == null) {
      if (!create) return undefined
      parent[segment] = {}
    }
    parent = parent[segment] as Record<string | number, unknown>
  }
  return parent
}

function applyPatch(doc: Record<string, unknown>, patch: MindooDBAppJsonPatch) {
  for (const { path, value } of patch.set ?? []) parentOf(doc, path, true)![path.at(-1)!] = hydrate(value)
  for (const { path } of patch.unset ?? []) {
    const parent = parentOf(doc, path, false)
    if (parent) delete parent[path.at(-1)!]
  }
  for (const { path, index, deleteCount } of patch.listDelete ?? []) {
    ;(parentOf(doc, path, false)![path.at(-1)!] as unknown[]).splice(index, deleteCount)
  }
  for (const { path, index, values } of patch.listInsert ?? []) {
    const parent = parentOf(doc, path, true)!
    if (parent[path.at(-1)!] == null) parent[path.at(-1)!] = []
    ;(parent[path.at(-1)!] as unknown[]).splice(index, 0, ...(hydrate(values) as unknown[]))
  }
}

const HOST_ACTOR = 'aa'.repeat(16)

export class AutomergeTestHost {
  private docs = new Map<string, Doc>()
  private feed: string[] = []
  private counter = 0
  /** Actor of the next writes; set per replica in tests. */
  actor = HOST_ACTOR

  /** Merges another replica's documents into this one (both directions give the same result). */
  syncFrom(other: AutomergeTestHost): void {
    for (const [id, doc] of other.docs) {
      const mine = this.docs.get(id)
      this.docs.set(id, mine ? A.merge(A.clone(mine, { actor: this.actor }), doc) : A.clone(doc, { actor: this.actor }))
      this.changed(id)
    }
  }

  clone(actor?: string): AutomergeTestHost {
    const copy = new AutomergeTestHost()
    for (const [id, doc] of this.docs) copy.docs.set(id, A.clone(doc))
    copy.feed = [...this.feed]
    copy.counter = this.counter
    copy.actor = actor ?? this.actor
    for (const [id, doc] of copy.docs) copy.docs.set(id, A.clone(doc, { actor: copy.actor }))
    return copy
  }

  private snapshot(id: string): MindooDBAppDocument {
    const doc = this.docs.get(id)!
    return { id, data: plain(doc) as Record<string, unknown>, heads: A.getHeads(doc) }
  }

  private changed(id: string) {
    this.feed.push(id)
  }

  readonly connection: HavenConnection = {
    canWrite: true,
    session: {} as HavenConnection['session'],
    context: {} as HavenConnection['context'],
    database: {
      documents: {
        get: async (id: string) => (this.docs.has(id) ? this.snapshot(id) : null),
        create: async (input: { id?: string; set?: Record<string, unknown> }) => {
          const id = input.id ?? `doc${++this.counter}`
          if (!this.docs.has(id)) {
            // Like MindooDB, equal caller ids start from a shared seed.
            const seed = A.from<Record<string, unknown>>(hydrate(input.set ?? {}) as Record<string, unknown>, {
              actor: input.id ? '00'.repeat(16) : this.actor,
            })
            this.docs.set(id, A.clone(seed, { actor: this.actor }))
            this.changed(id)
          }
          return this.snapshot(id)
        },
        update: async (id: string, input: { json?: MindooDBAppJsonPatch }) => {
          const patch = input.json ?? {}
          const current = A.clone(this.docs.get(id)!, { actor: this.actor })
          const next = patch.baseHeads?.length
            ? A.changeAt(current, patch.baseHeads, (doc) => applyPatch(doc, patch)).newDoc
            : A.change(current, (doc) => applyPatch(doc, patch))
          this.docs.set(id, A.merge(this.docs.get(id)!, next))
          this.changed(id)
          return this.snapshot(id)
        },
        getHeadCursor: async () => ({ cursor: String(this.feed.length) }),
        list: async (query: { cursor?: string | null }) => {
          const from = Number(query.cursor ?? 0)
          const items = this.feed.slice(from).map((docId) => ({ id: docId }))
          return { items, nextCursor: items.length ? String(this.feed.length) : null }
        },
      },
    } as unknown as HavenConnection['database'],
  }
}
