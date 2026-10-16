import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

export const fidelityDir = dirname(fileURLToPath(import.meta.url))
export const corpusDir = join(fidelityDir, 'corpus')
/** Per workbook the formula cells that differed from Excel when the baseline was accepted. */
export const valuesBaselinePath = join(fidelityDir, 'values-baseline.json')
/** Results of the last run; the global teardown turns them into fidelity-report/index.html. */
export const reportDir = join(fidelityDir, '../../fidelity-report')
export const valuesReportDir = join(reportDir, 'values')
/** Per workbook the differences after import and export that were accepted (category → count). */
export const roundtripBaselinePath = join(fidelityDir, 'roundtrip-baseline.json')
export const roundtripReportDir = join(reportDir, 'roundtrip')

export interface CorpusCase {
  id: string
  source: string
  file: string
  tags: string[]
  /** Which tests use the workbook: `values` (results stored by Excel) and/or `roundtrip`. */
  axes: Array<'values' | 'roundtrip'>
  /** Why the workbook is skipped (a known failure that would block every run). */
  skip?: string
  /** Why only the round trip skips it. */
  skipRoundtrip?: string
  /** Why the value test only reports differences beyond the baseline (results vary between runs). */
  unstable?: string
}

export const corpus: { cases: CorpusCase[] } = JSON.parse(readFileSync(join(fidelityDir, 'corpus.json'), 'utf8'))
