// Charts as the file has them: GenOffice's chart-add form covers the common
// types and options only, so the chart part itself (with its style and
// color parts) is kept and put back over the chart an export adds. Data
// references (`<c:f>`) are stored separately, bound to ids by the caller.
import JSZip from 'jszip'

import { attribute, relationships, relsPathOf, sheetParts } from './pivot-parts'

export interface ChartParts {
  /** The chart part with each `<c:f>` reference replaced by `{{ref:n}}`. */
  chart: string
  style?: string
  colors?: string
}

export interface ExtractedChart {
  parts: ChartParts
  /** The references, in placeholder order, as written in the file. */
  refs: string[]
}

const STYLE_REL = 'http://schemas.microsoft.com/office/2011/relationships/chartStyle'
const COLORS_REL = 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle'
const STYLE_TYPE = 'application/vnd.ms-office.chartstyle+xml'
const COLORS_TYPE = 'application/vnd.ms-office.chartcolorstyle+xml'

const unescapeXml = (text: string) =>
  text.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')
const escapeText = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

/** Elements that point at parts not kept (embedded workbook, drawn-on shapes). */
const DANGLING = /<(\w+:)?(?:externalData|userShapes)\b[^>]*\/>|<(\w+:)?(?:externalData|userShapes)\b[^>]*>[\s\S]*?<\/\1?(?:externalData|userShapes)>/g

/** The file's charts by chart part path (the `chartPath` the engine reports). */
export async function extractCharts(bytes: Uint8Array): Promise<Map<string, ExtractedChart>> {
  const zip = await JSZip.loadAsync(bytes)
  const charts = new Map<string, ExtractedChart>()
  for (const path of Object.keys(zip.files).filter((name) => /^xl\/charts\/chart[^/]*\.xml$/.test(name))) {
    const xml = await zip.file(path)?.async('text')
    if (!xml) continue
    const refs: string[] = []
    const chart = xml
      .replace(DANGLING, '')
      .replace(/<((?:\w+:)?f)>([\s\S]*?)<\/\1>/g, (_match, tag: string, text: string) => {
        refs.push(unescapeXml(text))
        return `<${tag}>{{ref:${refs.length - 1}}}</${tag}>`
      })
    const parts: ChartParts = { chart }
    for (const rel of await relationships(zip, path)) {
      const text = await zip.file(rel.target)?.async('text')
      if (!text) continue
      if (rel.type === STYLE_REL) parts.style = text
      else if (rel.type === COLORS_REL) parts.colors = text
    }
    charts.set(path, { parts, refs })
  }
  return charts
}

/** A stored chart to put over the n-th chart (drawing order) of a sheet. */
export interface ChartPlacement {
  sheetName: string
  index: number
  parts: ChartParts
  /** The references, rendered for the file, in placeholder order. */
  refs: string[]
}

/** Puts kept chart parts over the charts the export added. */
export async function applyCharts(bytes: Uint8Array, placements: readonly ChartPlacement[]): Promise<Uint8Array> {
  if (placements.length === 0) return bytes
  const zip = await JSZip.loadAsync(bytes)
  const sheets = await sheetParts(zip)
  let types = (await zip.file('[Content_Types].xml')?.async('text')) ?? ''
  let number = Math.max(
    0,
    ...Object.keys(zip.files).map((name) => Number(/^xl\/charts\/(?:style|colors)(\d+)\.xml$/.exec(name)?.[1] ?? 0)),
  )
  const chartsOf = new Map<string, string[]>()
  for (const placement of placements) {
    const sheetPart = sheets.get(placement.sheetName)
    if (!sheetPart) continue
    let charts = chartsOf.get(sheetPart)
    if (!charts) {
      charts = []
      const drawing = (await relationships(zip, sheetPart)).find((rel) => rel.type.endsWith('/drawing'))
      const drawingXml = drawing ? await zip.file(drawing.target)?.async('text') : undefined
      if (drawing && drawingXml) {
        const targets = new Map((await relationships(zip, drawing.target)).map((rel) => [rel.id, rel.target]))
        for (const match of drawingXml.matchAll(/<(?:\w+:)?chart\b[^>]*\br:id="[^"]*"[^>]*>/g)) {
          const target = targets.get(attribute(match[0], 'r:id') ?? '')
          if (target) charts.push(target)
        }
      }
      chartsOf.set(sheetPart, charts)
    }
    const chartPath = charts[placement.index]
    if (!chartPath) continue
    zip.file(
      chartPath,
      placement.parts.chart.replace(/\{\{ref:(\d+)\}\}/g, (_match, n: string) => escapeText(placement.refs[Number(n)] ?? '')),
    )
    const rels: string[] = []
    for (const [kind, text, rel, type] of [
      ['style', placement.parts.style, STYLE_REL, STYLE_TYPE],
      ['colors', placement.parts.colors, COLORS_REL, COLORS_TYPE],
    ] as const) {
      if (!text) continue
      number += 1
      const name = `${kind}${number}.xml`
      zip.file(`xl/charts/${name}`, text)
      rels.push(`<Relationship Id="rId${rels.length + 1}" Type="${rel}" Target="${name}"/>`)
      types = types.replace('</Types>', `<Override PartName="/xl/charts/${name}" ContentType="${type}"/></Types>`)
    }
    const relsPath = relsPathOf(chartPath)
    if (rels.length) {
      zip.file(
        relsPath,
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
          `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`,
      )
    } else {
      zip.remove(relsPath)
    }
  }
  zip.file('[Content_Types].xml', types)
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
}
