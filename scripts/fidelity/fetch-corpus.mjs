#!/usr/bin/env node
/**
 * Downloads the fidelity corpus (e2e/fidelity/corpus.json) from its sources at the pinned
 * commits: every workbook to e2e/fidelity/corpus/<id>.xlsx, each source's license to
 * e2e/fidelity/corpus/LICENSE-<source>, and the SHA-256 of every workbook to
 * e2e/fidelity/corpus/SHA256SUMS. Files already present are kept unless --force.
 *
 *   node scripts/fidelity/fetch-corpus.mjs [--force] [--case <id>]... [--tag <tag>]...
 */
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { corpusDir, loadCorpus, parseArgs, selectCases } from './common.mjs'

const args = parseArgs(process.argv.slice(2))
const corpus = loadCorpus()
mkdirSync(corpusDir, { recursive: true })

const raw = (source, path) => `https://raw.githubusercontent.com/${source.repo}/${source.commit}/${path}`

async function download(url) {
  for (let attempt = 1; ; attempt++) {
    const response = await fetch(url)
    if (response.ok) return new Uint8Array(await response.arrayBuffer())
    if (attempt >= 4 || response.status === 404) throw new Error(`${url}: HTTP ${response.status}`)
    await new Promise((resolve) => setTimeout(resolve, 1000 * 2 ** attempt))
  }
}

for (const [name, source] of Object.entries(corpus.sources)) {
  const target = join(corpusDir, `LICENSE-${name}`)
  if (args.force || !existsSync(target)) {
    const text = new TextDecoder().decode(await download(raw(source, source.licenseFile)))
    writeFileSync(target, `${source.repo} at ${source.commit} (${source.license})\n\n${text}`)
  }
}

const sums = new Map()
const sumsPath = join(corpusDir, 'SHA256SUMS')
if (existsSync(sumsPath)) {
  for (const line of readFileSync(sumsPath, 'utf8').split('\n')) {
    const [hash, file] = line.split(/\s+/)
    if (hash && file) sums.set(file, hash)
  }
}

for (const entry of selectCases(corpus, args)) {
  const source = corpus.sources[entry.source]
  const file = `${entry.id}.xlsx`
  const target = join(corpusDir, file)
  if (!args.force && existsSync(target)) {
    if (!sums.has(file)) sums.set(file, sha256(readFileSync(target)))
    continue
  }
  const bytes = await download(raw(source, `${source.dir}/${entry.file}`))
  writeFileSync(target, bytes)
  sums.set(file, sha256(bytes))
  console.log(`${entry.id}: ${bytes.length} bytes`)
}

writeFileSync(
  sumsPath,
  [...sums].sort(([a], [b]) => a.localeCompare(b)).map(([file, hash]) => `${hash}  ${file}`).join('\n') + '\n',
)

function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex')
}
