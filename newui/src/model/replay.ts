// Replays a save's structural operations (GenOffice's journal: inserts,
// removals and moves by position, each in the coordinates the previous ones
// produced) on a sheet's id order list. The result says which ids are new,
// which were deleted, and how to turn the stored list into the new one with
// list patches.
import { createId } from './schema'

export interface AxisOp {
  kind: 'insert' | 'remove' | 'move'
  index: number
  count: number
  /** move: pre-move position the block goes in front of. */
  before?: number
}

export interface ReplayResult {
  /** New raw order (deleted ids stay in place). */
  order: string[]
  /** Live ids by position. */
  live: string[]
  created: Set<string>
  deleted: Set<string>
}

export function replayAxis(
  baseOrder: readonly string[],
  baseLive: readonly string[],
  ops: readonly AxisOp[],
  prefix: string,
): ReplayResult {
  const order = [...baseOrder]
  const live = [...baseLive]
  const created = new Set<string>()
  const deleted = new Set<string>()
  const fresh = (count: number) =>
    Array.from({ length: count }, () => {
      const id = createId(prefix)
      created.add(id)
      return id
    })
  // Raw position in front of which ids for live position `index` belong.
  const rawPosition = (index: number) => (index < live.length ? order.indexOf(live[index]!) : order.length)
  const ensureLength = (length: number) => {
    if (live.length >= length) return
    const ids = fresh(length - live.length)
    order.push(...ids)
    live.push(...ids)
  }
  for (const op of ops) {
    if (op.kind === 'insert') {
      ensureLength(op.index)
      const ids = fresh(op.count)
      order.splice(rawPosition(op.index), 0, ...ids)
      live.splice(op.index, 0, ...ids)
    } else if (op.kind === 'remove') {
      ensureLength(op.index + op.count)
      for (const id of live.splice(op.index, op.count)) {
        if (created.has(id)) {
          created.delete(id)
          order.splice(order.indexOf(id), 1)
        } else {
          deleted.add(id)
        }
      }
    } else {
      const before = op.before ?? 0
      ensureLength(Math.max(op.index + op.count, before))
      const target = before < live.length ? live[before] : undefined
      const block = live.splice(op.index, op.count)
      for (const id of block) order.splice(order.indexOf(id), 1)
      const liveAt = target === undefined ? live.length : live.indexOf(target)
      live.splice(liveAt, 0, ...block)
      const rawAt = target === undefined ? order.length : order.indexOf(target)
      order.splice(rawAt, 0, ...block)
    }
  }
  return { order, live, created, deleted }
}

/** Appends ids until the live list covers `length` positions. */
export function extendAxis(result: ReplayResult, length: number, prefix: string): void {
  while (result.live.length < length) {
    const id = createId(prefix)
    result.created.add(id)
    result.order.push(id)
    result.live.push(id)
  }
}

export interface ListPatch {
  deletes: { index: number; deleteCount: number }[]
  inserts: { index: number; values: string[] }[]
}

/**
 * List patches that turn `base` into `next`, for a host that applies all
 * deletes before all inserts, each kind in the given order. Ids present in
 * both keep their relative order (deletes only drop moved-away ids), so the
 * diff is: delete ids that moved, then insert new and moved ids.
 */
export function diffList(base: readonly string[], next: readonly string[]): ListPatch {
  const nextPositions = new Map(next.map((id, index) => [id, index]))
  // Ids that stay put: the longest run of base ids appearing in `next` in
  // increasing position (patience LIS); every other base id is moved/removed.
  const candidates = base.filter((id) => nextPositions.has(id))
  const keep = longestIncreasing(candidates.map((id) => nextPositions.get(id)!)).map((i) => candidates[i]!)
  const kept = new Set(keep)
  const deletes: ListPatch['deletes'] = []
  for (let index = base.length - 1; index >= 0; index -= 1) {
    if (kept.has(base[index]!)) continue
    const last = deletes[deletes.length - 1]
    if (last && last.index === index + 1) {
      last.index = index
      last.deleteCount += 1
    } else {
      deletes.push({ index, deleteCount: 1 })
    }
  }
  const inserts: ListPatch['inserts'] = []
  for (let index = 0; index < next.length; index += 1) {
    const id = next[index]!
    if (kept.has(id)) continue
    const last = inserts[inserts.length - 1]
    if (last && last.index + last.values.length === index) last.values.push(id)
    else inserts.push({ index, values: [id] })
  }
  return { deletes, inserts }
}

/** Indices (into `values`) of a longest strictly increasing subsequence. */
function longestIncreasing(values: readonly number[]): number[] {
  const tails: number[] = []
  const previous: number[] = new Array(values.length).fill(-1)
  for (let index = 0; index < values.length; index += 1) {
    let low = 0
    let high = tails.length
    while (low < high) {
      const middle = (low + high) >> 1
      if (values[tails[middle]!]! < values[index]!) low = middle + 1
      else high = middle
    }
    if (low > 0) previous[index] = tails[low - 1]!
    tails[low] = index
  }
  const result: number[] = []
  for (let index = tails[tails.length - 1] ?? -1; index >= 0; index = previous[index]!) result.push(index)
  return result.reverse()
}
