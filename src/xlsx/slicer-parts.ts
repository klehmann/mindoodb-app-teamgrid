// Slicers: GenOffice draws them read-only and cannot write them, so their
// parts are kept as the file has them: the slicer caches (workbook), the
// slicer parts (per sheet), the extLst entries that register both, and each
// slicer's drawing frame. Cross references are kept as placeholders — table
// ids by table name, sheet tab ids by sheet name — and filled in again for
// the file an export builds. Timelines are not kept.
import JSZip from 'jszip'

import { attribute, relationships, relsPathOf, sheetParts } from './pivot-parts'

export interface SlicerParts {
  /** Slicer cache parts, `{{table:Name}}` / `{{tab:Sheet}}` in place of ids. */
  caches: { name: string; xml: string }[]
  /** Workbook extLst `<ext>` elements, `{{cache:i}}` in place of relationship ids. */
  workbookExts: string[]
  /** Per sheet: its slicer parts and extLst `<ext>` elements (`{{slicer:i}}`). */
  sheets: { sheet: string; parts: string[]; exts: string[] }[]
  /** Each slicer's drawing frame (the anchor's content between its position and clientData), by slicer name. */
  frames: Record<string, string>
}

const CACHE_REL = 'http://schemas.microsoft.com/office/2007/relationships/slicerCache'
const SLICER_REL = 'http://schemas.microsoft.com/office/2007/relationships/slicer'
const CACHE_TYPE = 'application/vnd.ms-excel.slicerCache+xml'
const SLICER_TYPE = 'application/vnd.ms-excel.slicer+xml'
const ANCHOR = /<(xdr:(?:twoCellAnchor|oneCellAnchor|absoluteAnchor))\b[^>]*>[\s\S]*?<\/\1>/g
const EXT = /<ext\b[^>]*>[\s\S]*?<\/ext>/g

function extsOf(xml: string): string[] {
  const list = /<extLst>([\s\S]*?)<\/extLst>\s*<\/(?:workbook|worksheet)>/.exec(xml)?.[1] ?? ''
  return [...list.matchAll(EXT)].map((match) => match[0])
}

function rootNamespaces(xml: string): string[] {
  const root = /<[\w:]+\b[^>]*>/.exec(xml.replace(/^<\?xml[^>]*\?>\s*/, ''))?.[0] ?? ''
  return [...root.matchAll(/\sxmlns:\w+="[^"]*"/g)].map((match) => match[0])
}

/** The file's slicers, or undefined when it has none. */
export async function extractSlicers(bytes: Uint8Array): Promise<SlicerParts | undefined> {
  const zip = await JSZip.loadAsync(bytes)
  const workbook = (await zip.file('xl/workbook.xml')?.async('text')) ?? ''
  const cacheRels = (await relationships(zip, 'xl/workbook.xml')).filter((rel) => rel.type === CACHE_REL)
  if (cacheRels.length === 0) return undefined

  const sheetNameByTab = new Map<string, string>()
  for (const match of workbook.matchAll(/<sheet\b[^>]*>/g)) {
    const name = attribute(match[0], 'name')
    const tab = attribute(match[0], 'sheetId')
    if (name !== undefined && tab) sheetNameByTab.set(tab, name)
  }
  const tableNameById = new Map<string, string>()
  for (const path of Object.keys(zip.files).filter((name) => /^xl\/tables\/[^/]+\.xml$/.test(name))) {
    const head = /<(?:\w+:)?table\b[^>]*\sref="[^"]*"[^>]*>/.exec((await zip.file(path)?.async('text')) ?? '')?.[0] ?? ''
    const id = attribute(head, 'id')
    const name = attribute(head, 'displayName') ?? attribute(head, 'name')
    if (id && name) tableNameById.set(id, name)
  }

  const parts: SlicerParts = { caches: [], workbookExts: [], sheets: [], frames: {} }
  const cacheIndexByRel = new Map<string, number>()
  for (const rel of cacheRels) {
    let xml = (await zip.file(rel.target)?.async('text')) ?? ''
    xml = xml.replace(/(<(?:\w+:)?tableSlicerCache\b[^>]*\btableId=")(\d+)(")/g, (match, head: string, id: string, tail: string) => {
      const name = tableNameById.get(id)
      return name ? `${head}{{table:${name}}}${tail}` : match
    })
    xml = xml.replace(/(<(?:\w+:)?pivotTable\b[^>]*\btabId=")(\d+)(")/g, (match, head: string, tab: string, tail: string) => {
      const name = sheetNameByTab.get(tab)
      return name ? `${head}{{tab:${name}}}${tail}` : match
    })
    const name = attribute(/<(?:\w+:)?slicerCacheDefinition\b[^>]*>/.exec(xml)?.[0] ?? '', 'name') ?? ''
    cacheIndexByRel.set(rel.id, parts.caches.length)
    parts.caches.push({ name, xml })
  }
  for (const ext of extsOf(workbook)) {
    if (!/slicerCache/.test(ext)) continue
    parts.workbookExts.push(
      ext.replace(/\br:id="([^"]+)"/g, (match, id: string) => {
        const index = cacheIndexByRel.get(id)
        return index === undefined ? match : `r:id="{{cache:${index}}}"`
      }),
    )
  }

  for (const [sheet, sheetPart] of await sheetParts(zip)) {
    const rels = await relationships(zip, sheetPart)
    const slicerRels = rels.filter((rel) => rel.type === SLICER_REL)
    if (slicerRels.length === 0) continue
    const sheetXml = (await zip.file(sheetPart)?.async('text')) ?? ''
    const entry: SlicerParts['sheets'][number] = { sheet, parts: [], exts: [] }
    const indexByRel = new Map<string, number>()
    for (const rel of slicerRels) {
      indexByRel.set(rel.id, entry.parts.length)
      entry.parts.push((await zip.file(rel.target)?.async('text')) ?? '')
    }
    for (const ext of extsOf(sheetXml)) {
      if (!/slicerList/.test(ext)) continue
      entry.exts.push(
        ext.replace(/\br:id="([^"]+)"/g, (match, id: string) => {
          const index = indexByRel.get(id)
          return index === undefined ? match : `r:id="{{slicer:${index}}}"`
        }),
      )
    }
    parts.sheets.push(entry)

    const drawing = rels.find((rel) => rel.type.endsWith('/drawing'))
    const drawingXml = drawing ? await zip.file(drawing.target)?.async('text') : undefined
    if (!drawingXml) continue
    const namespaces = rootNamespaces(drawingXml)
    for (const anchor of drawingXml.matchAll(ANCHOR)) {
      const name = /<(?:\w+:)?slicer\b[^>]*\bname="([^"]*)"/.exec(anchor[0])?.[1]
      const frame = /<\/xdr:(?:to|ext|pos)>([\s\S]*)<xdr:clientData\b/.exec(anchor[0])?.[1]
      if (!name || !frame || /timeslicer/.test(anchor[0])) continue
      // The frame's prefixes are declared on the drawing's root: carry them along.
      const own = frame.trim()
      const missing = namespaces.filter((declaration) => !own.includes(declaration.trim()))
      parts.frames[name] = own.replace(/^<([\w:]+)/, `<$1${missing.join('')}`)
    }
  }
  return parts
}

/** A slicer drawn on a sheet: the n-th anchor of the sheet's drawing holds its placeholder. */
export interface SlicerFrame {
  sheet: string
  anchorIndex: number
  name: string
}

function appendExts(xml: string, root: 'workbook' | 'worksheet', exts: readonly string[]): string {
  if (exts.length === 0) return xml
  const closing = `</${root}>`
  if (/<extLst>[\s\S]*?<\/extLst>\s*<\/(?:workbook|worksheet)>/.test(xml)) {
    return xml.replace(/<\/extLst>(\s*<\/(?:workbook|worksheet)>)/, `${exts.join('')}</extLst>$1`)
  }
  return xml.replace(closing, `<extLst>${exts.join('')}</extLst>${closing}`)
}

/** Puts the slicers back into a file built from scratch; unchanged bytes when a table or sheet they need is gone. */
export async function applySlicers(
  bytes: Uint8Array,
  slicers: SlicerParts,
  frames: readonly SlicerFrame[],
): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  let workbook = (await zip.file('xl/workbook.xml')?.async('text')) ?? ''
  let workbookRels = (await zip.file('xl/_rels/workbook.xml.rels')?.async('text')) ?? ''
  let types = (await zip.file('[Content_Types].xml')?.async('text')) ?? ''
  const sheets = await sheetParts(zip)

  const tabByName = new Map<string, string>()
  for (const match of workbook.matchAll(/<sheet\b[^>]*>/g)) {
    const name = attribute(match[0], 'name')
    const tab = attribute(match[0], 'sheetId')
    if (name !== undefined && tab) tabByName.set(name, tab)
  }
  const tableIdByName = new Map<string, string>()
  for (const path of Object.keys(zip.files).filter((name) => /^xl\/tables\/[^/]+\.xml$/.test(name))) {
    const head = /<(?:\w+:)?table\b[^>]*\sref="[^"]*"[^>]*>/.exec((await zip.file(path)?.async('text')) ?? '')?.[0] ?? ''
    const id = attribute(head, 'id')
    const name = attribute(head, 'displayName') ?? attribute(head, 'name')
    if (id && name) tableIdByName.set(name.toLowerCase(), id)
  }
  let missing = false
  const caches = slicers.caches.map((cache) =>
    cache.xml
      .replace(/\{\{table:([^}]*)\}\}/g, (_match, name: string) => tableIdByName.get(name.toLowerCase()) ?? ((missing = true), ''))
      .replace(/\{\{tab:([^}]*)\}\}/g, (_match, name: string) => tabByName.get(name) ?? ((missing = true), '')),
  )
  if (missing || slicers.sheets.some((entry) => !sheets.has(entry.sheet))) return bytes

  let relNumber = Math.max(0, ...[...workbookRels.matchAll(/\bId="rId(\d+)"/g)].map((match) => Number(match[1])))
  const cacheRelIds: string[] = []
  caches.forEach((xml, index) => {
    const path = `xl/slicerCaches/slicerCache${index + 1}.xml`
    zip.file(path, xml)
    relNumber += 1
    cacheRelIds.push(`rId${relNumber}`)
    workbookRels = workbookRels.replace(
      '</Relationships>',
      `<Relationship Id="rId${relNumber}" Type="${CACHE_REL}" Target="slicerCaches/slicerCache${index + 1}.xml"/></Relationships>`,
    )
    types = types.replace('</Types>', `<Override PartName="/${path}" ContentType="${CACHE_TYPE}"/></Types>`)
  })
  workbook = appendExts(
    workbook,
    'workbook',
    slicers.workbookExts.map((ext) => ext.replace(/\{\{cache:(\d+)\}\}/g, (_match, index: string) => cacheRelIds[Number(index)] ?? '')),
  )
  // Each cache has a defined name of its own (=#N/A), as Excel writes it.
  const names = slicers.caches
    .filter((cache) => cache.name && !new RegExp(`<definedName\\b[^>]*\\bname="${cache.name}"`).test(workbook))
    .map((cache) => `<definedName name="${cache.name}">#N/A</definedName>`)
  if (names.length) {
    if (/<\/definedNames>/.test(workbook)) workbook = workbook.replace('</definedNames>', `${names.join('')}</definedNames>`)
    else {
      const after = /<\/externalReferences>|<\/functionGroups>|<\/sheets>/.exec(workbook)
      if (after) workbook = workbook.replace(after[0], `${after[0]}<definedNames>${names.join('')}</definedNames>`)
    }
  }

  let slicerNumber = 0
  for (const entry of slicers.sheets) {
    const sheetPart = sheets.get(entry.sheet)!
    const relsPath = relsPathOf(sheetPart)
    let rels =
      (await zip.file(relsPath)?.async('text')) ??
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    let sheetRel = Math.max(0, ...[...rels.matchAll(/\bId="rId(\d+)"/g)].map((match) => Number(match[1])))
    const relIds = entry.parts.map((xml) => {
      slicerNumber += 1
      const path = `xl/slicers/slicer${slicerNumber}.xml`
      zip.file(path, xml)
      sheetRel += 1
      rels = rels.replace(
        '</Relationships>',
        `<Relationship Id="rId${sheetRel}" Type="${SLICER_REL}" Target="../slicers/slicer${slicerNumber}.xml"/></Relationships>`,
      )
      types = types.replace('</Types>', `<Override PartName="/${path}" ContentType="${SLICER_TYPE}"/></Types>`)
      return `rId${sheetRel}`
    })
    zip.file(relsPath, rels)
    const sheetXml = (await zip.file(sheetPart)?.async('text')) ?? ''
    zip.file(
      sheetPart,
      appendExts(
        sheetXml,
        'worksheet',
        entry.exts.map((ext) => ext.replace(/\{\{slicer:(\d+)\}\}/g, (_match, index: string) => relIds[Number(index)] ?? '')),
      ),
    )
  }

  // Frames: the export drew a placeholder shape at each slicer's position.
  const bySheet = new Map<string, SlicerFrame[]>()
  for (const frame of frames) bySheet.set(frame.sheet, [...(bySheet.get(frame.sheet) ?? []), frame])
  for (const [sheet, list] of bySheet) {
    const sheetPart = sheets.get(sheet)
    const drawing = sheetPart ? (await relationships(zip, sheetPart)).find((rel) => rel.type.endsWith('/drawing')) : undefined
    let drawingXml = drawing ? await zip.file(drawing.target)?.async('text') : undefined
    if (!drawing || !drawingXml) continue
    let position = -1
    drawingXml = drawingXml.replace(ANCHOR, (anchor) => {
      position += 1
      const frame = list.find((entry) => entry.anchorIndex === position)
      const xml = frame ? slicers.frames[frame.name] : undefined
      if (!xml) return anchor
      return anchor.replace(/(<\/xdr:(?:to|ext|pos)>)[\s\S]*(<xdr:clientData\b)/, `$1${xml}$2`)
    })
    zip.file(drawing.target, drawingXml)
  }

  zip.file('xl/workbook.xml', workbook)
  zip.file('xl/_rels/workbook.xml.rels', workbookRels)
  zip.file('[Content_Types].xml', types)
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
}
