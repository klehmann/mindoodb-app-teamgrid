/**
 * Round trip through TeamGrid's storage: each corpus workbook is imported into the mock
 * Haven (`workbook_create`, so it is stored as MindooDB documents), exported again
 * (`workbook_export`) and both files are read with our own parser (workbookModel.ts):
 * values, formulas, cell formats, merges, column widths, row heights, views, sheet
 * features (filters, validation, conditional formats, hyperlinks, tables, comments,
 * drawings) and part types. A workbook fails when a category shows more differences
 * than accepted in roundtrip-baseline.json; without an entry it only records.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { test as base } from '@playwright/test'

import { expect, TeamGridHost } from '../testHost'
import { corpus, corpusDir, roundtripBaselinePath, roundtripReportDir } from './paths'
import { compareModels, readWorkbookModel, type Differences } from './workbookModel'

export interface RoundtripResult {
  id: string
  tags: string[]
  /** category → number of differences */
  counts: Record<string, number>
  /** category → the first differences, for the report */
  samples: Differences
  error?: string
}

const MAX_SAMPLES = 12

const baseline: Record<string, Record<string, number>> = existsSync(roundtripBaselinePath)
  ? JSON.parse(readFileSync(roundtripBaselinePath, 'utf8'))
  : {}

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

for (const entry of corpus.cases.filter((c) => c.axes.includes('roundtrip'))) {
  test(`roundtrip ${entry.id} ${entry.tags.map((tag) => `@${tag}`).join(' ')}`, async ({ host }) => {
    test.fixme(!!(entry.skip ?? entry.skipRoundtrip), entry.skip ?? entry.skipRoundtrip)
    const bytes = new Uint8Array(readFileSync(join(corpusDir, `${entry.id}.xlsx`)))
    const result: RoundtripResult = { id: entry.id, tags: entry.tags, counts: {}, samples: {} }
    try {
      await host.importWorkbook(entry.id, bytes)
      const exported = await host.exportXlsx()
      const differences = compareModels(await readWorkbookModel(bytes), await readWorkbookModel(exported))
      for (const [category, entries] of Object.entries(differences)) {
        result.counts[category] = entries.length
        result.samples[category] = entries.slice(0, MAX_SAMPLES)
      }
    } catch (error) {
      result.error = error instanceof Error ? error.message : String(error)
    } finally {
      mkdirSync(roundtripReportDir, { recursive: true })
      writeFileSync(join(roundtripReportDir, `${entry.id}.json`), JSON.stringify(result, null, 1))
    }
    const total = Object.values(result.counts).reduce((a, b) => a + b, 0)
    test.info().annotations.push({ type: 'roundtrip', description: total ? `${total} differences` : 'no differences' })
    expect(result.error, 'import and export').toBeUndefined()
    const accepted = baseline[entry.id]
    if (!accepted) return
    const worse = Object.entries(result.counts)
      .filter(([category, count]) => count > (accepted[category] ?? 0))
      .map(([category, count]) => `${category}: ${accepted[category] ?? 0} → ${count} (${result.samples[category]?.[0] ?? ''})`)
    expect(worse, 'categories with more differences than at the baseline').toEqual([])
  })
}
