#!/usr/bin/env node
// Smoke test for the xlsx-wasm build (Node target in ./pkg-node):
//   wasm-pack build --target nodejs --release --out-dir pkg-node
//   node smoke.mjs [file.xlsx] [--sidecar path/to/native/xlsx-sidecar]
// With --sidecar, the same reads go through the native sidecar too and the
// JSON results are compared (sessionId aside).
import { spawn } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { basename, dirname, join, resolve } from 'node:path'
import { createInterface } from 'node:readline'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const require = createRequire(import.meta.url)
const { XlsxEngine } = require(join(here, 'pkg-node/xlsx_wasm.js'))

const args = process.argv.slice(2)
const sidecarAt = args.indexOf('--sidecar')
const sidecarBin = sidecarAt >= 0 ? args.splice(sidecarAt, 2)[1] : null
const file = resolve(args[0] ?? join(here, '../../../../../genoffice/apps/shell/build/shell-new/blank.xlsx'))

const ms = (start) => `${(performance.now() - start).toFixed(1)} ms`
const bytes = readFileSync(file)
console.log(`file: ${file} (${bytes.length} bytes)`)

const engine = new XlsxEngine()
let t = performance.now()
const workbook = JSON.parse(engine.open(bytes, basename(file), 'en', null))
console.log(`open: ${ms(t)}  session=${workbook.sessionId}`)
console.log(`sheets: ${workbook.sheets.map((s) => `${s.name} (${s.rowCount}x${s.columnCount})`).join(', ')}`)
console.log(`styles=${workbook.styles.length} visuals=${workbook.visuals.length} ` +
  `(${workbook.visuals.map((v) => v.kind ?? v.type).join(',')}) definedNames=${workbook.definedNames.length}`)

const reads = []
for (const sheet of workbook.sheets) {
  const range = {
    startRow: 0, endRow: Math.min(sheet.rowCount, 500) - 1,
    startColumn: 0, endColumn: Math.min(sheet.columnCount, 50) - 1,
  }
  t = performance.now()
  const result = JSON.parse(engine.readRange(workbook.sessionId, sheet.id,
    range.startRow, range.endRow, range.startColumn, range.endColumn))
  console.log(`readRange ${sheet.name} rows 0..${range.endRow}: ${result.cells.length} cells, ` +
    `${result.merges.length} merges, complete=${result.indexingComplete} in ${ms(t)}`)
  t = performance.now()
  const formulas = JSON.parse(engine.readFormulaCells(workbook.sessionId, sheet.id))
  console.log(`readFormulaCells ${sheet.name}: ${formulas.cells.length} in ${ms(t)}`)
  reads.push({ sheet, range, result, formulas })
}
const media = workbook.visuals.find((v) => v.mediaPath || v.fillMediaPath)
if (media) {
  const image = engine.readMediaBytes(workbook.sessionId, media.id)
  console.log(`readMediaBytes ${media.id}: ${image.length} bytes`)
}
engine.close(workbook.sessionId)
try {
  engine.readRange(workbook.sessionId, workbook.sheets[0].id, 0, 0, 0, 0)
} catch (error) {
  console.log(`after close: ${error.code}: ${error.message}`)
}

if (sidecarBin) await compareWithSidecar(sidecarBin)

async function compareWithSidecar(bin) {
  const child = spawn(bin, [], { stdio: ['pipe', 'pipe', 'inherit'] })
  const lines = createInterface({ input: child.stdout })[Symbol.asyncIterator]()
  let id = 0
  const call = async (command) => {
    child.stdin.write(`${JSON.stringify({ version: 1, requestId: String(++id), ...command })}\n`)
    const response = JSON.parse((await lines.next()).value)
    if (!response.ok) throw new Error(`${response.error.code}: ${response.error.message}`)
    return response.result
  }
  const native = await call({ command: 'open', path: file, locale: 'en' })
  let same = compare('open', workbook, native, ['.sessionId'])
  for (const { sheet, range, result, formulas } of reads) {
    let nativeRange
    do {
      nativeRange = await call({ command: 'read_range', sessionId: native.sessionId, sheetId: sheet.id, range })
    } while (!nativeRange.indexingComplete)
    const nativeFormulas = await call({ command: 'read_formula_cells', sessionId: native.sessionId, sheetId: sheet.id })
    same = compare(`readRange ${sheet.name}`, result, nativeRange) && same
    same = compare(`readFormulaCells ${sheet.name}`, formulas, nativeFormulas) && same
  }
  await call({ command: 'close', sessionId: native.sessionId })
  console.log(`native sidecar parity: ${same ? 'all results identical' : 'DIFFERENCES above'}`)
  child.stdin.end()
  if (!same) process.exitCode = 1
}

// Logs the paths where wasm and native JSON differ; true when identical.
function compare(label, wasm, native, ignore = []) {
  const diffs = []
  const walk = (a, b, path) => {
    if (ignore.includes(path) || JSON.stringify(a) === JSON.stringify(b)) return
    if (a && b && typeof a === 'object' && typeof b === 'object') {
      for (const key of new Set([...Object.keys(a), ...Object.keys(b)])) walk(a[key], b[key], `${path}.${key}`)
    } else {
      diffs.push(`  ${path}: wasm=${JSON.stringify(a)?.slice(0, 80)} native=${JSON.stringify(b)?.slice(0, 80)}`)
    }
  }
  walk(wasm, native, '')
  console.log(`sidecar ${label}: ${diffs.length ? `DIFFERENT\n${diffs.slice(0, 10).join('\n')}` : 'identical'}`)
  return diffs.length === 0
}
