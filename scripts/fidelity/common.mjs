/** Paths and helpers shared by the fidelity scripts (see e2e/fidelity/README.md). */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

export const root = join(dirname(fileURLToPath(import.meta.url)), '../..')
export const fidelityDir = join(root, 'e2e/fidelity')
export const corpusDir = join(fidelityDir, 'corpus')
export const reportDir = join(root, 'fidelity-report')

export function loadCorpus() {
  return JSON.parse(readFileSync(join(fidelityDir, 'corpus.json'), 'utf8'))
}

/** `--force`, `--case <id>` (repeatable) and `--tag <tag>` (repeatable). */
export function parseArgs(argv) {
  const out = { force: false, cases: [], tags: [] }
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--force') out.force = true
    else if (argv[i] === '--case') out.cases.push(argv[++i])
    else if (argv[i] === '--tag') out.tags.push(argv[++i])
    else throw new Error(`Unknown argument ${argv[i]}`)
  }
  return out
}

export function selectCases(corpus, args) {
  return corpus.cases.filter(
    (c) => (!args.cases.length || args.cases.includes(c.id)) && (!args.tags.length || c.tags.some((t) => args.tags.includes(t))),
  )
}

export function sortKeys(record) {
  return Object.fromEntries(Object.keys(record).sort().map((k) => [k, record[k]]))
}
