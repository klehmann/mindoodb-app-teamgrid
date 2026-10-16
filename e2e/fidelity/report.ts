/**
 * fidelity-report/index.html from the last run's results: formula results against Excel
 * (fidelity-report/values/*.json) per category and function, and the round trip through
 * TeamGrid's storage (fidelity-report/roundtrip/*.json) per kind of difference, each with
 * the workbooks and cells that differ. Written by the global teardown after every run.
 */
import { existsSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

import { reportDir, roundtripReportDir, valuesReportDir } from './paths'
import type { RoundtripResult } from './roundtrip.e2e'
import type { ValuesResult } from './values.e2e'

const MAX_CELLS_SHOWN = 40

const escape = (value: unknown) =>
  String(typeof value === 'string' ? value : JSON.stringify(value)).replace(
    /[&<>"]/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!,
  )

const percent = (matched: number, total: number) => (total ? `${((100 * matched) / total).toFixed(1)} %` : '–')

function tone(share: number): string {
  return share >= 1 ? 'ok' : share >= 0.9 ? 'warn' : 'bad'
}

function rate(matched: number, total: number): string {
  const share = total ? matched / total : 1
  return `<td class="num"><span class="bar ${tone(share)}" style="--share:${(share * 100).toFixed(1)}%"></span>${percent(matched, total)}</td>`
}

function readResults<T extends { id: string }>(dir: string): T[] {
  if (!existsSync(dir)) return []
  return readdirSync(dir)
    .filter((file) => file.endsWith('.json'))
    .map((file) => JSON.parse(readFileSync(join(dir, file), 'utf8')) as T)
    .sort((a, b) => a.id.localeCompare(b.id))
}

function card(label: string, value: string, detail: string, state: string): string {
  return `<div class="card ${state}"><div class="label">${label}</div><div class="value">${value}</div><div class="detail">${detail}</div></div>`
}

function valuesSection(results: ValuesResult[]): { card: string; body: string } | null {
  if (results.length === 0) return null
  const compared = results.filter((r) => !r.streaming && !r.error)
  const functions = new Map<string, { cells: number; matched: number; workbooks: Set<string> }>()
  const tags = new Map<string, { cells: number; matched: number; workbooks: number }>()
  let cells = 0
  let matched = 0
  for (const result of compared) {
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
    ([an, a], [bn, b]) => (b.cells - b.matched) - (a.cells - a.matched) || a.matched / a.cells - b.matched / b.cells || an.localeCompare(bn),
  )
  const imperfect = byGap.filter(([, f]) => f.matched < f.cells).length
  const functionRows = byGap
    .map(
      ([name, f]) =>
        `<tr class="${f.matched === f.cells ? 'perfect' : ''}"><td><code>${escape(name)}</code></td><td class="num">${f.cells}</td><td class="num">${f.cells - f.matched || ''}</td>${rate(f.matched, f.cells)}<td class="links">${[...f.workbooks].map((id) => `<a href="#v-${id}">${escape(id.replace(/^ironcalc-/, ''))}</a>`).join(' ')}</td></tr>`,
    )
    .join('\n')
  const tagRows = [...tags]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([tag, t]) => `<tr><td>${escape(tag)}</td><td class="num">${t.workbooks}</td><td class="num">${t.cells}</td>${rate(t.matched, t.cells)}</tr>`)
    .join('\n')
  const streamed = results.filter((r) => r.streaming)
  const failed = results.filter((r) => r.error)
  const workbookSections = results
    .filter((r) => r.error || r.mismatches.length)
    .map((r) => {
      const rows = r.mismatches
        .slice(0, MAX_CELLS_SHOWN)
        .map(
          (m) =>
            `<tr><td>${escape(m.ref)}</td><td><code>${escape(m.formula)}</code></td><td>${escape(m.expected)}</td><td>${escape(m.got)}</td><td><span class="tag ${m.kind}">${escape(m.kind)}</span></td></tr>`,
        )
        .join('\n')
      const more = r.mismatches.length > MAX_CELLS_SHOWN ? `<p class="muted">… ${r.mismatches.length - MAX_CELLS_SHOWN} more in values/${r.id}.json</p>` : ''
      return `<details id="v-${r.id}"><summary><b>${escape(r.id)}</b> <span class="muted">${r.matched}/${r.cells} · ${escape(r.application)}</span></summary>${r.error ? `<p class="bad">${escape(r.error)}</p>` : ''}${rows ? `<table><tr><th>Cell</th><th>Formula</th><th>Excel</th><th>TeamGrid</th><th>Kind</th></tr>${rows}</table>${more}` : ''}</details>`
    })
    .join('\n')
  const share = cells ? matched / cells : 1
  return {
    card: card(
      'Formula results',
      percent(matched, cells),
      `${matched.toLocaleString('en-US')} of ${cells.toLocaleString('en-US')} cells · ${imperfect} of ${functions.size} functions differ${failed.length ? ` · ${failed.length} failed` : ''}`,
      tone(share),
    ),
    body: `<section id="values"><h2>Formula results against Excel</h2>
<p class="muted">The value TeamGrid's formula engine computed for every formula cell, against the result Excel stored in the file. ${compared.length} workbooks compared${streamed.length ? `; ${streamed.length} too large to load whole (formulas keep Excel's results): ${streamed.map((r) => escape(r.id)).join(', ')}` : ''}.</p>
<div class="grid"><div><h3>By category</h3><table><tr><th>Category</th><th>Workbooks</th><th>Cells</th><th>Match</th></tr>${tagRows}</table></div>
<div><h3>By function <label class="toggle"><input type="checkbox" class="hide-perfect" checked> only with differences</label></h3><table class="functions"><tr><th>Function</th><th>Cells</th><th>Differ</th><th>Match</th><th>Workbooks</th></tr>${functionRows}</table></div></div>
<h3>Differences per workbook</h3>${workbookSections}</section>`,
  }
}

function roundtripSection(results: RoundtripResult[]): { card: string; body: string } | null {
  if (results.length === 0) return null
  const categories = new Map<string, { workbooks: Set<string>; count: number }>()
  for (const result of results) {
    for (const [category, count] of Object.entries(result.counts)) {
      const entry = categories.get(category) ?? { workbooks: new Set<string>(), count: 0 }
      entry.workbooks.add(result.id)
      entry.count += count
      categories.set(category, entry)
    }
  }
  const clean = results.filter((r) => !r.error && Object.keys(r.counts).length === 0)
  const failed = results.filter((r) => r.error)
  const categoryRows = [...categories]
    .sort(([, a], [, b]) => b.workbooks.size - a.workbooks.size)
    .map(
      ([category, c]) =>
        `<tr><td>${escape(category)}</td><td class="num">${c.workbooks.size}</td><td class="num">${c.count.toLocaleString('en-US')}</td><td class="links">${[...c.workbooks].slice(0, 30).map((id) => `<a href="#r-${id}">${escape(id.replace(/^(rxw|ironcalc)-/, ''))}</a>`).join(' ')}${c.workbooks.size > 30 ? ' …' : ''}</td></tr>`,
    )
    .join('\n')
  const workbookSections = results
    .filter((r) => r.error || Object.keys(r.counts).length)
    .map((r) => {
      const blocks = Object.entries(r.samples)
        .map(([category, entries]) => `<h4>${escape(category)} <span class="muted">${r.counts[category]}</span></h4><ul>${entries.map((e) => `<li><code>${escape(e)}</code></li>`).join('')}</ul>`)
        .join('')
      const total = Object.values(r.counts).reduce((a, b) => a + b, 0)
      return `<details id="r-${r.id}"><summary><b>${escape(r.id)}</b> <span class="muted">${r.error ? 'failed' : `${total} differences in ${Object.keys(r.counts).join(', ')}`}</span></summary>${r.error ? `<p class="bad">${escape(r.error)}</p>` : ''}${blocks}</details>`
    })
    .join('\n')
  const share = results.length ? clean.length / results.length : 1
  return {
    card: card(
      'Round trip',
      `${clean.length}/${results.length}`,
      `workbooks unchanged after import and export${failed.length ? ` · ${failed.length} failed` : ''}`,
      failed.length ? 'bad' : tone(share),
    ),
    body: `<section id="roundtrip"><h2>Round trip through TeamGrid's storage</h2>
<p class="muted">Each workbook imported (stored as MindooDB documents) and exported again, both files read with our own parser: values, formulas, formats, merges, sizes, views, sheet features, objects and part types.</p>
<h3>Kinds of differences</h3><table><tr><th>Kind</th><th>Workbooks</th><th>Differences</th><th>Workbooks</th></tr>${categoryRows}</table>
<h3>Per workbook</h3>${workbookSections}</section>`,
  }
}

export function writeReport(): void {
  const values = valuesSection(readResults<ValuesResult>(valuesReportDir))
  const roundtrip = roundtripSection(readResults<RoundtripResult>(roundtripReportDir))
  if (!values && !roundtrip) return
  const html = `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>TeamGrid vs Excel</title>
<style>
:root { --bg: #f6f8fa; --panel: #fff; --text: #1f2328; --muted: #656d76; --line: #d8dee4; --ok: #1a7f37; --warn: #9a6700; --bad: #cf222e; --accent: #0969da; }
@media (prefers-color-scheme: dark) { :root { --bg: #0d1117; --panel: #161b22; --text: #e6edf3; --muted: #8d96a0; --line: #30363d; --ok: #3fb950; --warn: #d29922; --bad: #f85149; --accent: #4493f8; } }
* { box-sizing: border-box; }
body { font: 14px/1.45 system-ui, -apple-system, sans-serif; margin: 0; color: var(--text); background: var(--bg); }
header { padding: 24px 32px 8px; } h1 { margin: 0 0 4px; font-size: 22px; }
nav a { color: var(--accent); margin-right: 16px; text-decoration: none; }
main { padding: 8px 32px 48px; }
.cards { display: flex; gap: 16px; flex-wrap: wrap; margin: 16px 0; }
.card { background: var(--panel); border: 1px solid var(--line); border-radius: 8px; padding: 16px 20px; min-width: 260px; border-top: 4px solid var(--line); }
.card.ok { border-top-color: var(--ok); } .card.warn { border-top-color: var(--warn); } .card.bad { border-top-color: var(--bad); }
.card .label { color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: .04em; }
.card .value { font-size: 30px; font-weight: 600; margin: 4px 0; } .card .detail { color: var(--muted); }
section { background: var(--panel); border: 1px solid var(--line); border-radius: 8px; padding: 8px 24px 24px; margin: 24px 0; }
.grid { display: grid; grid-template-columns: minmax(280px, 1fr) 2fr; gap: 24px; align-items: start; }
@media (max-width: 900px) { .grid { grid-template-columns: 1fr; } }
table { border-collapse: collapse; width: 100%; margin: 8px 0 16px; }
th, td { border-bottom: 1px solid var(--line); padding: 5px 8px; text-align: left; vertical-align: top; }
th { color: var(--muted); font-weight: 600; font-size: 12px; }
td.num { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
.bar { display: inline-block; width: 60px; height: 6px; border-radius: 3px; margin-right: 8px; vertical-align: middle; background: linear-gradient(to right, currentColor var(--share), var(--line) var(--share)); }
.ok { color: var(--ok); } .warn { color: var(--warn); } .bad { color: var(--bad); } .muted { color: var(--muted); }
code { font: 12px ui-monospace, SFMono-Regular, Menlo, monospace; word-break: break-word; }
.links a { color: var(--accent); text-decoration: none; margin-right: 6px; font-size: 12px; }
details { border-bottom: 1px solid var(--line); padding: 6px 0; } summary { cursor: pointer; }
.tag { font-size: 11px; padding: 1px 6px; border-radius: 10px; background: var(--line); }
.tag.missing { background: var(--bad); color: #fff; } .tag.error { background: var(--warn); color: #fff; }
.toggle { font-weight: normal; font-size: 12px; color: var(--muted); margin-left: 12px; }
section.hide-perfect tr.perfect { display: none; }
ul { margin: 4px 0 8px; padding-left: 20px; } h4 { margin: 8px 0 0; font-size: 13px; }
</style></head>
<body>
<header><h1>TeamGrid vs Excel</h1><nav>${values ? '<a href="#values">Formula results</a>' : ''}${roundtrip ? '<a href="#roundtrip">Round trip</a>' : ''}</nav></header>
<main>
<div class="cards">${values?.card ?? ''}${roundtrip?.card ?? ''}</div>
${values?.body ?? ''}
${roundtrip?.body ?? ''}
</main>
<script>
for (const box of document.querySelectorAll('input.hide-perfect')) {
  const section = box.closest('section')
  const apply = () => section.classList.toggle('hide-perfect', box.checked)
  box.addEventListener('change', apply); apply()
}
</script>
</body></html>
`
  writeFileSync(join(reportDir, 'index.html'), html)
  console.log(`Fidelity report: ${join(reportDir, 'index.html')}`)
}
