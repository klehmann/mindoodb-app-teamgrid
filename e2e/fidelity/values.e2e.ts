/**
 * Formula results against Excel: each corpus workbook is imported into TeamGrid in the
 * mock Haven (`workbook_create`), every formula cell is read back with `read_cells` and its
 * value, as TeamGrid's formula engine computed it, compared with the result Excel stored
 * in the file. The displayed text is not used: where the engine fails, TeamGrid shows
 * Excel's cached result, which would hide the gap.
 *
 * A workbook fails when a cell that matched at the last accepted baseline
 * (values-baseline.json) no longer does. Without a baseline entry it only records.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { test as base } from '@playwright/test'

import { expect, TeamGridHost } from '../testHost'
import { mismatchKind, sameResult } from './compare'
import { corpus, corpusDir, valuesBaselinePath, valuesReportDir } from './paths'
import { functionsOf, readWorkbookCells, VOLATILE } from './workbookCells'

export interface ValueMismatch {
  ref: string
  formula: string
  expected: unknown
  got: unknown
  text?: unknown
  kind: ReturnType<typeof mismatchKind>
}

export interface ValuesResult {
  id: string
  tags: string[]
  application: string
  cells: number
  matched: number
  volatile: number
  /** Per function: formula cells calling it, and how many of those matched. */
  functions: Record<string, { cells: number; matched: number }>
  mismatches: ValueMismatch[]
  /** Too large for TeamGrid to load whole: its formulas show Excel's stored results, nothing to compare. */
  streaming?: boolean
  error?: string
}

const baseline: Record<string, string[]> = existsSync(valuesBaselinePath)
  ? JSON.parse(readFileSync(valuesBaselinePath, 'utf8'))
  : {}

// One mock Haven per worker: booting it costs seconds, a workbook import about one.
const test = base.extend<object, { host: TeamGridHost }>({
  host: [
    async ({ browser }, use) => {
      const page = await browser.newPage({ viewport: { width: 1400, height: 900 }, locale: 'en-US' })
      const host = new TeamGridHost(page)
      await host.open()
      await use(host)
      await page.close()
    },
    { scope: 'worker' },
  ],
})

test.describe.configure({ mode: 'parallel' })

for (const entry of corpus.cases.filter((c) => c.axes.includes('values'))) {
  test(`values ${entry.id} ${entry.tags.map((tag) => `@${tag}`).join(' ')}`, async ({ host }) => {
    test.fixme(!!entry.skip, entry.skip)
    const bytes = new Uint8Array(readFileSync(join(corpusDir, `${entry.id}.xlsx`)))
    const workbook = await readWorkbookCells(bytes)
    const result: ValuesResult = {
      id: entry.id,
      tags: entry.tags,
      application: workbook.application,
      cells: 0,
      matched: 0,
      volatile: 0,
      functions: {},
      mismatches: [],
    }
    try {
      const { sheets, streaming } = await host.importWorkbook(entry.id, bytes)
      if (streaming) result.streaming = true
      for (const sheet of streaming ? [] : workbook.sheets) {
        if (sheet.cells.length === 0) continue
        const sheetId = sheets.find((s) => s.name === sheet.name)?.sheetId
        if (!sheetId) throw new Error(`TeamGrid has no sheet "${sheet.name}"`)
        const read = await host.readCells(
          sheetId,
          sheet.cells.map((cell) => cell.address),
        )
        for (const cell of sheet.cells) {
          const got = read[cell.address]
          const functions = functionsOf(cell.formula)
          const volatile = functions.some((name) => VOLATILE.has(name))
          const ok = volatile
            ? !(typeof got?.value === 'string' && got.value.startsWith('#')) || cell.error
            : sameResult(got?.value, cell.value)
          result.cells++
          if (volatile) result.volatile++
          if (ok) result.matched++
          for (const name of functions) {
            const stats = (result.functions[name] ??= { cells: 0, matched: 0 })
            stats.cells++
            if (ok) stats.matched++
          }
          if (!ok) {
            result.mismatches.push({
              ref: `${sheet.name}!${cell.address}`,
              formula: cell.formula,
              expected: cell.value,
              got: got?.value ?? null,
              ...(got?.text !== undefined ? { text: got.text } : {}),
              kind: mismatchKind(got?.value, cell.value, cell.error),
            })
          }
        }
      }
    } catch (error) {
      result.error = error instanceof Error ? error.message : String(error)
    } finally {
      mkdirSync(valuesReportDir, { recursive: true })
      writeFileSync(join(valuesReportDir, `${entry.id}.json`), JSON.stringify(result, null, 1))
    }

    test.info().annotations.push({
      type: 'values',
      description: result.streaming
        ? 'streamed (too large to load whole), formulas not recalculated'
        : `${result.matched}/${result.cells} formula results match Excel`,
    })
    expect(result.error, 'import and read').toBeUndefined()
    const known = baseline[entry.id]
    if (!known) return
    const accepted = new Set(known)
    const regressions = result.mismatches.filter((m) => !accepted.has(m.ref))
    expect(
      regressions.map((m) => `${m.ref} ${m.formula}: Excel ${JSON.stringify(m.expected)}, TeamGrid ${JSON.stringify(m.got)}`),
      'cells that matched Excel at the baseline',
    ).toEqual([])
  })
}
