#!/usr/bin/env node
// Copies the GenOffice Sheets sources TeamGrid's new UI builds on into
// vendor/genoffice, mirroring GenOffice's own directory layout so its
// relative imports keep working. Run again to take an upstream update:
//   node scripts/sync-genoffice.mjs [path-to-genoffice-checkout]
// Local changes to vendored files belong in patches/, never in vendor/.
import { execFileSync } from 'node:child_process'
import { cpSync, existsSync, mkdirSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const source = resolve(process.argv[2] ?? join(here, '../../../genoffice'))
const target = resolve(here, '../vendor/genoffice')

const PATHS = [
  'LICENSE',
  'NOTICE',
  'apps/sheets/src/renderer',
  'apps/sheets/src/shared',
  'apps/sheets/src/ai',
  'apps/sheets/src/types',
  'apps/sheets/native/xlsx-engine/Cargo.toml',
  'apps/sheets/native/xlsx-engine/Cargo.lock',
  'apps/sheets/native/xlsx-engine/src',
  ...[
    'xlsx-gateway', 'ui', 'i18n', 'agent-core', 'ai-provider', 'docx-engine',
    'pptx-render', 'pptx-engine', 'zip-gate', 'project-store', 'electron-utils',
  ].flatMap((pkg) => [`packages/${pkg}/package.json`, `packages/${pkg}/src`]),
]

if (!existsSync(join(source, 'apps/sheets'))) {
  console.error(`not a genoffice checkout: ${source}`)
  process.exit(1)
}
rmSync(target, { recursive: true, force: true })
for (const path of PATHS) {
  const from = join(source, path)
  if (!existsSync(from)) continue
  mkdirSync(dirname(join(target, path)), { recursive: true })
  cpSync(from, join(target, path), {
    recursive: true,
    // tests and fixtures stay upstream; they need GenOffice's full workspace
    filter: (file) => !/\.test\.tsx?$|\/(__tests__|fixtures)\//.test(file),
  })
}
// Local changes to vendored files, applied in name order. Each patch is
// relative to vendor/genoffice (`patch -p1`); keep them small so upstream
// updates rebase cleanly.
const patches = join(here, '../patches')
for (const name of existsSync(patches) ? readdirSync(patches).filter((n) => n.endsWith('.patch')).sort() : []) {
  execFileSync('patch', ['-p1', '--forward', '--no-backup-if-mismatch', '-d', target, '-i', join(patches, name)], { stdio: 'inherit' })
}
const commit = execFileSync('git', ['-C', source, 'rev-parse', 'HEAD'], { encoding: 'utf8' }).trim()
writeFileSync(join(target, 'VENDORED_FROM'), `https://github.com/genspark-ai/genoffice @ ${commit}\n`)
console.log(`vendored genoffice @ ${commit.slice(0, 8)} into ${target}`)
