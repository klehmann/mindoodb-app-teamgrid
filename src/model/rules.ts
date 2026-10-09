// Conditional formats and data validation. The editor (Univer) and
// GenOffice's save speak Univer's rule JSON; the file read gives OOXML-like
// rules. GenOffice's own converters turn a read rule into Univer JSON (the
// conditional-format one through Univer's rule builder, which needs no
// running editor). Stored are that JSON and the rule's areas by row/column
// ids, keyed by a hash of both: two people adding different rules keep both,
// and a changed rule is a removed plus an added one.
import type { WorkbookFile, WorkbookSaveRequest } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import type { ColumnId, IdArea, RowId } from './schema'
import { areaToIds, areaToIndices } from './sheet-features'
import type { SidecarSheet } from './sidecar-read'
import { canonicalJson } from './styles'

export interface StoredRule {
  areas: IdArea[]
  /** Univer's rule JSON without its ranges and ids. */
  rule: Record<string, unknown>
  /** Evaluation order (conditional formats; lower first). */
  priority?: number
  stopIfTrue?: boolean
}

function hash(text: string): string {
  let h1 = 0x811c9dc5
  let h2 = 0x01000193
  for (let index = 0; index < text.length; index += 1) {
    h1 = Math.imul(h1 ^ text.charCodeAt(index), 16777619) >>> 0
    h2 = Math.imul(h2 ^ text.charCodeAt(index), 2246822519) >>> 0
  }
  return `r${h1.toString(36)}${h2.toString(36)}`
}

function plain(value: unknown): Record<string, unknown> {
  return JSON.parse(JSON.stringify(value ?? {})) as Record<string, unknown>
}

function keyed(rules: StoredRule[]): Record<string, StoredRule> {
  const byKey: Record<string, StoredRule> = {}
  for (const rule of rules) byKey[hash(canonicalJson({ areas: rule.areas, rule: rule.rule }))] = rule
  return byKey
}

/**
 * GenOffice's converters live in the editor's renderer, which may only load
 * once the browser shim is installed: they are loaded on demand, before the
 * first read (`loadRuleConverters`), and a read without them fails loudly
 * rather than storing a sheet without its rules.
 */
interface Converters {
  buildConditionalRule: typeof import('../../vendor/genoffice/apps/sheets/src/renderer/univer-sync').buildConditionalRule
  toUniverDvRule: typeof import('../../vendor/genoffice/apps/sheets/src/renderer/univer-sync').toUniverDvRule
  /** A worksheet stand-in for GenOffice's builder: all it asks for is a new rule builder. */
  factory: { newConditionalFormattingRule(): unknown }
}

let converters: Converters | undefined

export async function loadRuleConverters(): Promise<void> {
  if (converters) return
  const [sync, facade] = await Promise.all([
    import('../../vendor/genoffice/apps/sheets/src/renderer/univer-sync'),
    import('@univerjs/sheets-conditional-formatting/facade'),
  ])
  converters = {
    buildConditionalRule: sync.buildConditionalRule,
    toUniverDvRule: sync.toUniverDvRule,
    factory: { newConditionalFormattingRule: () => new facade.FConditionalFormattingBuilder() },
  }
}

function loaded(): Converters {
  if (!converters) throw new Error('loadRuleConverters() must run before reading conditional formats.')
  return converters
}

export function readConditionalFormats(
  sheet: SidecarSheet,
  dxfStyles: WorkbookFile['dxfStyles'],
  rowIds: readonly RowId[],
  columnIds: readonly ColumnId[],
): Record<string, StoredRule> | undefined {
  const rules: StoredRule[] = []
  for (const read of sheet.conditionalRules ?? []) {
    const areas = read.ranges.map((area) => areaToIds(area, rowIds, columnIds)).filter((area): area is IdArea => !!area)
    if (areas.length === 0) continue
    let rule: unknown
    try {
      const { buildConditionalRule, factory } = loaded()
      rule = (buildConditionalRule(factory as never, dxfStyles ?? [], read) as { rule?: unknown } | null)?.rule
    } catch {
      rule = undefined
    }
    if (!rule) continue
    rules.push({ areas, rule: plain(rule), priority: read.priority, ...(read.stopIfTrue ? { stopIfTrue: true } : {}) })
  }
  return rules.length ? keyed(rules) : undefined
}

export function readDataValidations(
  sheet: SidecarSheet,
  file: Pick<WorkbookFile, 'definedNames' | 'sheets'>,
  rowIds: readonly RowId[],
  columnIds: readonly ColumnId[],
): Record<string, StoredRule> | undefined {
  const rules: StoredRule[] = []
  for (const read of sheet.dataValidations ?? []) {
    const areas = read.ranges.map((area) => areaToIds(area, rowIds, columnIds)).filter((area): area is IdArea => !!area)
    if (areas.length === 0) continue
    const built = loaded().toUniverDvRule(read, 'dv', file)
    if (!built) continue
    const { ranges: _ranges, uid: _uid, ...rule } = built
    rules.push({ areas, rule: plain(rule) })
  }
  return rules.length ? keyed(rules) : undefined
}

function indexed(
  rules: Record<string, StoredRule> | undefined,
  rowIndex: ReadonlyMap<RowId, number>,
  columnIndex: ReadonlyMap<ColumnId, number>,
) {
  return Object.values(rules ?? {})
    .sort((left, right) => (left.priority ?? 0) - (right.priority ?? 0))
    .flatMap((stored) => {
      const ranges = stored.areas
        .map((area) => areaToIndices(area, rowIndex, columnIndex))
        .filter((area): area is NonNullable<typeof area> => !!area)
      return ranges.length ? [{ ranges, stored }] : []
    })
}

export function cfState(
  rules: Record<string, StoredRule> | undefined,
  sheetId: string,
  rowIndex: ReadonlyMap<RowId, number>,
  columnIndex: ReadonlyMap<ColumnId, number>,
): WorkbookSaveRequest['cfStates'][number] | undefined {
  const entries = indexed(rules, rowIndex, columnIndex)
  if (!entries.length) return undefined
  return {
    sheetId,
    rules: entries.map(({ ranges, stored }) => ({ ranges, stopIfTrue: stored.stopIfTrue === true, rule: stored.rule })),
  }
}

export function dvState(
  rules: Record<string, StoredRule> | undefined,
  sheetId: string,
  rowIndex: ReadonlyMap<RowId, number>,
  columnIndex: ReadonlyMap<ColumnId, number>,
): WorkbookSaveRequest['dvStates'][number] | undefined {
  const entries = indexed(rules, rowIndex, columnIndex)
  if (!entries.length) return undefined
  return { sheetId, rules: entries.map(({ ranges, stored }) => ({ ranges, rule: stored.rule })) }
}
