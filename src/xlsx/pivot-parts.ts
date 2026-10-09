// PivotTables as raw xlsx parts. GenOffice's save can only write a pivot the
// editor creates; a pivot read from a file it keeps verbatim in that file.
// TeamGrid builds every file from scratch, so it takes a pivot's parts
// (table definition, cache definition, cache records) out of the saved file
// and puts them back into the next one, with the pivot's position and
// source range rewritten to where those cells are now.
import JSZip from 'jszip'

import type { CellArea } from '../model/sidecar-read'

export interface PivotParts {
  table: string
  cache: string
  records?: string
}

export interface ExtractedPivot {
  /** The sheet the pivot's output sits on. */
  sheetName: string
  name: string
  location: CellArea
  /** The source range on its sheet; absent for named or external sources (kept verbatim). */
  source?: { sheetName: string; area: CellArea }
  parts: PivotParts
}

const MAIN = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const CONTENT_TYPE = {
  table: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
  cache: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
  records: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
}

function columnIndexOf(label: string): number {
  let index = 0
  for (const char of label.toUpperCase()) index = index * 26 + (char.charCodeAt(0) - 64)
  return index - 1
}

function columnLabel(index: number): string {
  let label = ''
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) label = String.fromCharCode(65 + ((n - 1) % 26)) + label
  return label
}

export function parseRef(ref: string): CellArea | undefined {
  const match = /^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i.exec(ref.trim())
  if (!match) return undefined
  const startColumn = columnIndexOf(match[1]!)
  const startRow = Number(match[2]) - 1
  return {
    startRow,
    startColumn,
    endRow: match[4] ? Number(match[4]) - 1 : startRow,
    endColumn: match[3] ? columnIndexOf(match[3]) : startColumn,
  }
}

export function formatRef(area: CellArea): string {
  return `${columnLabel(area.startColumn)}${area.startRow + 1}:${columnLabel(area.endColumn)}${area.endRow + 1}`
}

const unescapeXml = (text: string) =>
  text.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')
const escapeXml = (text: string) =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

function attribute(element: string, name: string): string | undefined {
  const match = new RegExp(`\\s${name.replace(':', '\\:')}="([^"]*)"`).exec(element)
  return match ? unescapeXml(match[1]!) : undefined
}

function setAttribute(element: string, name: string, value: string): string {
  const pattern = new RegExp(`(\\s${name.replace(':', '\\:')}=")[^"]*(")`)
  if (pattern.test(element)) return element.replace(pattern, `$1${escapeXml(value)}$2`)
  return element.replace(/^(<[\w:]+)/, `$1 ${name}="${escapeXml(value)}"`)
}

/** `xl/worksheets/sheet1.xml` → `xl/worksheets/_rels/sheet1.xml.rels`. */
function relsPathOf(part: string): string {
  const slash = part.lastIndexOf('/')
  return `${part.slice(0, slash)}/_rels/${part.slice(slash + 1)}.rels`
}

/** A relationship target relative to its source part, as a zip path. */
function resolveTarget(source: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const parts = source.split('/').slice(0, -1)
  for (const segment of target.split('/')) {
    if (segment === '..') parts.pop()
    else if (segment !== '.') parts.push(segment)
  }
  return parts.join('/')
}

async function relationships(zip: JSZip, part: string) {
  const xml = await zip.file(relsPathOf(part))?.async('text')
  return [...(xml ?? '').matchAll(/<Relationship\b[^>]*>/g)].map((match) => ({
    id: attribute(match[0], 'Id') ?? '',
    type: attribute(match[0], 'Type') ?? '',
    target: resolveTarget(part, attribute(match[0], 'Target') ?? ''),
  }))
}

/** Worksheet name → its part path. */
async function sheetParts(zip: JSZip): Promise<Map<string, string>> {
  const workbook = (await zip.file('xl/workbook.xml')?.async('text')) ?? ''
  const rels = new Map((await relationships(zip, 'xl/workbook.xml')).map((rel) => [rel.id, rel.target]))
  const parts = new Map<string, string>()
  for (const match of workbook.matchAll(/<sheet\b[^>]*>/g)) {
    const name = attribute(match[0], 'name')
    const target = rels.get(attribute(match[0], 'r:id') ?? '')
    if (name !== undefined && target) parts.set(name, target)
  }
  return parts
}

export async function extractPivots(bytes: Uint8Array): Promise<ExtractedPivot[]> {
  const zip = await JSZip.loadAsync(bytes)
  const pivots: ExtractedPivot[] = []
  for (const [sheetName, sheetPart] of await sheetParts(zip)) {
    for (const rel of await relationships(zip, sheetPart)) {
      if (!rel.type.endsWith('/pivotTable')) continue
      const table = await zip.file(rel.target)?.async('text')
      if (!table) continue
      const cacheRel = (await relationships(zip, rel.target)).find((entry) => entry.type.endsWith('/pivotCacheDefinition'))
      const cache = cacheRel ? await zip.file(cacheRel.target)?.async('text') : undefined
      if (!cacheRel || !cache) continue
      const recordsRel = (await relationships(zip, cacheRel.target)).find((entry) => entry.type.endsWith('/pivotCacheRecords'))
      const records = recordsRel ? await zip.file(recordsRel.target)?.async('text') : undefined
      const head = /<pivotTableDefinition\b[^>]*>/.exec(table)?.[0] ?? ''
      const locationElement = /<location\b[^>]*>/.exec(table)?.[0] ?? ''
      const location = parseRef(attribute(locationElement, 'ref') ?? '')
      if (!location) continue
      const sourceElement = /<worksheetSource\b[^>]*>/.exec(cache)?.[0]
      const sourceRef = sourceElement ? attribute(sourceElement, 'ref') : undefined
      const sourceSheet = sourceElement ? attribute(sourceElement, 'sheet') : undefined
      const sourceArea = sourceRef ? parseRef(sourceRef) : undefined
      pivots.push({
        sheetName,
        name: attribute(head, 'name') ?? `Pivot${pivots.length + 1}`,
        location,
        ...(sourceArea && sourceSheet !== undefined ? { source: { sheetName: sourceSheet, area: sourceArea } } : {}),
        parts: { table, cache, ...(records ? { records } : {}) },
      })
    }
  }
  return pivots
}

export interface PivotPlacement {
  sheetName: string
  location: CellArea
  source?: { sheetName: string; area: CellArea }
  parts: PivotParts
}

/** Puts pivots into a file, each with a cache of its own. */
export async function injectPivots(bytes: Uint8Array, pivots: readonly PivotPlacement[]): Promise<Uint8Array> {
  if (pivots.length === 0) return bytes
  const zip = await JSZip.loadAsync(bytes)
  const sheets = await sheetParts(zip)
  let workbook = (await zip.file('xl/workbook.xml')?.async('text')) ?? ''
  let workbookRels = (await zip.file('xl/_rels/workbook.xml.rels')?.async('text')) ?? ''
  let types = (await zip.file('[Content_Types].xml')?.async('text')) ?? ''
  const used = (pattern: RegExp) =>
    Math.max(0, ...Object.keys(zip.files).map((name) => Number(pattern.exec(name)?.[1] ?? 0)))
  let tableNumber = used(/^xl\/pivotTables\/pivotTable(\d+)\.xml$/)
  let cacheNumber = Math.max(
    used(/^xl\/pivotCache\/pivotCacheDefinition(\d+)\.xml$/),
    used(/^xl\/pivotCache\/pivotCacheRecords(\d+)\.xml$/),
  )
  let cacheId = Math.max(0, ...[...workbook.matchAll(/<pivotCache\b[^>]*>/g)].map((match) => Number(attribute(match[0], 'cacheId') ?? 0)))
  let relNumber = Math.max(0, ...[...workbookRels.matchAll(/\bId="rId(\d+)"/g)].map((match) => Number(match[1])))
  const caches: string[] = []
  const sheetRels = new Map<string, string[]>()

  for (const pivot of pivots) {
    const sheetPart = sheets.get(pivot.sheetName)
    if (!sheetPart) continue
    tableNumber += 1
    cacheNumber += 1
    cacheId += 1
    relNumber += 1
    const tablePath = `xl/pivotTables/pivotTable${tableNumber}.xml`
    const cachePath = `xl/pivotCache/pivotCacheDefinition${cacheNumber}.xml`
    const recordsPath = `xl/pivotCache/pivotCacheRecords${cacheNumber}.xml`

    let table = pivot.parts.table.replace(/<pivotTableDefinition\b[^>]*>/, (head) => setAttribute(head, 'cacheId', String(cacheId)))
    table = table.replace(/<location\b[^>]*>/, (element) => setAttribute(element, 'ref', formatRef(pivot.location)))
    let cache = pivot.parts.cache
    if (pivot.source) {
      const source = pivot.source
      cache = cache.replace(/<worksheetSource\b[^>]*>/, (element) =>
        setAttribute(setAttribute(element, 'ref', formatRef(source.area)), 'sheet', source.sheetName),
      )
    }
    zip.file(tablePath, table)
    zip.file(
      relsPathOf(tablePath),
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${MAIN}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition${cacheNumber}.xml"/></Relationships>`,
    )
    if (pivot.parts.records) {
      cache = cache.replace(/<pivotCacheDefinition\b[^>]*>/, (head) => setAttribute(head, 'r:id', 'rId1'))
      zip.file(recordsPath, pivot.parts.records)
      zip.file(
        relsPathOf(cachePath),
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${MAIN}/pivotCacheRecords" Target="pivotCacheRecords${cacheNumber}.xml"/></Relationships>`,
      )
      types = types.replace('</Types>', `<Override PartName="/${recordsPath}" ContentType="${CONTENT_TYPE.records}"/></Types>`)
    } else {
      cache = cache.replace(/<pivotCacheDefinition\b[^>]*>/, (head) => head.replace(/\sr:id="[^"]*"/, ''))
    }
    zip.file(cachePath, cache)
    types = types.replace(
      '</Types>',
      `<Override PartName="/${cachePath}" ContentType="${CONTENT_TYPE.cache}"/><Override PartName="/${tablePath}" ContentType="${CONTENT_TYPE.table}"/></Types>`,
    )
    workbookRels = workbookRels.replace(
      '</Relationships>',
      `<Relationship Id="rId${relNumber}" Type="${MAIN}/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition${cacheNumber}.xml"/></Relationships>`,
    )
    caches.push(`<pivotCache cacheId="${cacheId}" r:id="rId${relNumber}"/>`)
    sheetRels.set(sheetPart, [...(sheetRels.get(sheetPart) ?? []), `../pivotTables/pivotTable${tableNumber}.xml`])
  }

  if (caches.length > 0) {
    if (/<pivotCaches\b[^>]*>/.test(workbook)) {
      workbook = workbook.replace('</pivotCaches>', `${caches.join('')}</pivotCaches>`)
    } else {
      // pivotCaches comes after calcPr and before the rest of the tail (CT_Workbook order).
      const tail = /<(?:smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b/.exec(workbook)
      const insert = `<pivotCaches>${caches.join('')}</pivotCaches>`
      workbook = tail
        ? `${workbook.slice(0, tail.index)}${insert}${workbook.slice(tail.index)}`
        : workbook.replace('</workbook>', `${insert}</workbook>`)
    }
  }
  for (const [sheetPart, targets] of sheetRels) {
    const path = relsPathOf(sheetPart)
    let rels =
      (await zip.file(path)?.async('text')) ??
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    let next = Math.max(0, ...[...rels.matchAll(/\bId="rId(\d+)"/g)].map((match) => Number(match[1])))
    for (const target of targets) {
      next += 1
      rels = rels.replace('</Relationships>', `<Relationship Id="rId${next}" Type="${MAIN}/pivotTable" Target="${target}"/></Relationships>`)
    }
    zip.file(path, rels)
  }
  zip.file('xl/workbook.xml', workbook)
  zip.file('xl/_rels/workbook.xml.rels', workbookRels)
  zip.file('[Content_Types].xml', types)
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
}
