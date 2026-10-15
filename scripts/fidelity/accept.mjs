#!/usr/bin/env node
/**
 * `pnpm fidelity:accept`: the last run's differences (fidelity-report/values) become the
 * baseline (e2e/fidelity/values-baseline.json); from then on a workbook fails when another
 * cell starts to differ from Excel. Workbooks that failed to import keep their entry.
 *
 *   node scripts/fidelity/accept.mjs [--case <id>]... [--tag <tag>]...
 */
import { existsSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { fidelityDir, parseArgs, reportDir, sortKeys } from './common.mjs'

const args = parseArgs(process.argv.slice(2))
const baselinePath = join(fidelityDir, 'values-baseline.json')
const dataDir = join(reportDir, 'values')

const baseline = existsSync(baselinePath) ? JSON.parse(readFileSync(baselinePath, 'utf8')) : {}
let accepted = 0
for (const file of existsSync(dataDir) ? readdirSync(dataDir) : []) {
  if (!file.endsWith('.json')) continue
  const result = JSON.parse(readFileSync(join(dataDir, file), 'utf8'))
  if (args.cases.length && !args.cases.includes(result.id)) continue
  if (args.tags.length && !result.tags.some((t) => args.tags.includes(t))) continue
  if (result.error || result.streaming) continue
  baseline[result.id] = result.mismatches.map((m) => m.ref).sort()
  accepted++
}
writeFileSync(baselinePath, JSON.stringify(sortKeys(baseline), null, 1) + '\n')
console.log(`Accepted ${accepted} workbooks into ${baselinePath}`)
