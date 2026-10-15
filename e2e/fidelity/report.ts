/**
 * fidelity-report/index.html from the last run's results (fidelity-report/values/*.json):
 * formula results per Excel function and per category, and per workbook the cells that
 * differ from Excel. Written by the global teardown after every run.
 */
import { existsSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { reportDir, valuesReportDir } from './paths'
import type { ValuesResult } from './values.e2e'

const MAX_CELLS_SHOWN = 40

const escape = (value: unknown) =>
  String(typeof value === 'string' ? value : JSON.stringify(value)).replace(
    /[&<>"]/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!,
  )

const percent = (matched: number, total: number) => (total ? `${((100 * matched) / total).toFixed(1)} %` : '–')

function rate(matched: number, total: number): string {
  const share = total ? matched / total : 1
  const tone = share === 1 ? 'ok' : share >= 0.9 ? 'warn' : 'bad'
  return `<td class="num ${tone}">${percent(matched, total)}</td>`
}

export function writeReport(): void {
  if (!existsSync(valuesReportDir)) return
  const results: ValuesResult[] = readdirSync(valuesReportDir)
    .filter((file) => file.endsWith('.json'))
    .map((file) => JSON.parse(readFileSync(join(valuesReportDir, file), 'utf8')) as ValuesResult)
    .sort((a, b) => a.id.localeCompare(b.id))
  if (results.length === 0) return

  const functions = new Map<string, { cells: number; matched: number; workbooks: Set<string> }>()
  const tags = new Map<string, { cells: number; matched: number; workbooks: number }>()
  let cells = 0
  let matched = 0
  for (const result of results) {
    cells += result.cells
    matched += result.matched
    for (const [name, stats] of Object.entries(result.functions)) {
      const entry = functions.get(name) ?? { cells: 0, matched: 0, workbooks: new Set<string>() }
      entry.cells += stats.cells
      entry.matched += stats.matched
      entry.workbooks.add(result.id)
      functions.set(name, entry)
    }
    for (const tag of result.tags) {
      const entry = tags.get(tag) ?? { cells: 0, matched: 0, workbooks: 0 }
      entry.cells += result.cells
      entry.matched += result.matched
      entry.workbooks++
      tags.set(tag, entry)
    }
  }
  const byGap = [...functions].sort(
    ([an, a], [bn, b]) => a.matched / a.cells - b.matched / b.cells || b.cells - a.cells || an.localeCompare(bn),
  )
  const imperfect = byGap.filter(([, f]) => f.matched < f.cells).length

  const functionRows = byGap
    .map(
      ([name, f]) =>
        `<tr class="${f.matched === f.cells ? 'perfect' : ''}"><td>${escape(name)}</td><td class="num">${f.cells}</td><td class="num">${f.cells - f.matched}</td>${rate(f.matched, f.cells)}<td>${[...f.workbooks].map((id) => `<a href="#${id}">${escape(id.replace(/^ironcalc-/, ''))}</a>`).join(' ')}</td></tr>`,
    )
    .join('\n')
  const tagRows = [...tags]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([tag, t]) => `<tr><td>${escape(tag)}</td><td class="num">${t.workbooks}</td><td class="num">${t.cells}</td>${rate(t.matched, t.cells)}</tr>`)
    .join('\n')
  const streamed = results.filter((r) => r.streaming)
  const streamedList = streamed.length
    ? `<h2>Streamed workbooks</h2><p>Too large for TeamGrid to load whole: their formulas show Excel's stored results and are not recalculated, so they are not compared.</p><p>${streamed.map((r) => escape(r.id)).join(', ')}</p>`
    : ''
  const workbookSections = results
    .filter((r) => r.error || r.mismatches.length)
    .map((r) => {
      const rows = r.mismatches
        .slice(0, MAX_CELLS_SHOWN)
        .map(
          (m) =>
            `<tr><td>${escape(m.ref)}</td><td><code>${escape(m.formula)}</code></td><td>${escape(m.expected)}</td><td>${escape(m.got)}</td><td>${escape(m.kind)}</td></tr>`,
        )
        .join('\n')
      const more = r.mismatches.length > MAX_CELLS_SHOWN ? `<p>… ${r.mismatches.length - MAX_CELLS_SHOWN} more in values/${r.id}.json</p>` : ''
      return `<section id="${r.id}"><h3>${escape(r.id)} <small>${r.matched}/${r.cells} · ${escape(r.application)}</small></h3>${r.error ? `<p class="bad">${escape(r.error)}</p>` : ''}${rows ? `<table><tr><th>Cell</th><th>Formula</th><th>Excel</th><th>TeamGrid</th><th>Kind</th></tr>${rows}</table>${more}` : ''}</section>`
    })
    .join('\n')

  const html = `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>TeamGrid vs Excel</title>
<style>
body { font: 14px/1.4 system-ui, sans-serif; margin: 24px; color: #1f2328; background: #fff; }
table { border-collapse: collapse; margin: 8px 0 24px; }
th, td { border-bottom: 1px solid #d0d7de; padding: 4px 8px; text-align: left; vertical-align: top; }
td.num { text-align: right; font-variant-numeric: tabular-nums; }
.ok { color: #1a7f37; } .warn { color: #9a6700; } .bad { color: #cf222e; }
code { font-size: 12px; } small { color: #656d76; font-weight: normal; }
label { display: block; margin: 8px 0; }
body.hide-perfect tr.perfect { display: none; }
</style></head>
<body>
<h1>TeamGrid vs Excel: formula results</h1>
<p>${matched} of ${cells} formula cells in ${results.length - streamed.length} workbooks match Excel (${percent(matched, cells)}); ${imperfect} of ${functions.size} functions have differences.${streamed.length ? ` ${streamed.length} streamed workbooks are not compared.` : ''}</p>
<h2>By category</h2>
<table><tr><th>Category</th><th>Workbooks</th><th>Cells</th><th>Match</th></tr>${tagRows}</table>
<h2>By function</h2>
<label><input type="checkbox" id="hide" checked> Only functions with differences</label>
<table><tr><th>Function</th><th>Cells</th><th>Differ</th><th>Match</th><th>Workbooks</th></tr>${functionRows}</table>
${streamedList}
<h2>Differences per workbook</h2>
${workbookSections}
<script>
const hide = document.getElementById('hide')
const apply = () => document.body.classList.toggle('hide-perfect', hide.checked)
hide.addEventListener('change', apply); apply()
</script>
</body></html>
`
  writeFileSync(join(reportDir, 'index.html'), html)
  console.log(`Fidelity report: ${join(reportDir, 'index.html')} (${percent(matched, cells)} of ${cells} formula cells match Excel)`)
}
