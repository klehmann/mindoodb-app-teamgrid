// Links to other workbooks ([1]Sheet1!A1): the externalLink parts (the other
// file's path and the cached values Excel shows until it updates) are kept
// as they are and put back in the same order, so `[n]` in formulas still
// names the same file.
import JSZip from 'jszip'

import { attribute, relationships, relsPathOf } from './pivot-parts'

export interface ExternalLinkPart {
  xml: string
  /** The part's relationships file (the other file's path), if any. */
  rels?: string
}

const LINK_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink'
const LINK_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml'

/** The workbook's external links in `[n]` order. */
export async function extractExternalLinks(bytes: Uint8Array): Promise<ExternalLinkPart[]> {
  const zip = await JSZip.loadAsync(bytes)
  const workbook = (await zip.file('xl/workbook.xml')?.async('text')) ?? ''
  const targets = new Map((await relationships(zip, 'xl/workbook.xml')).map((rel) => [rel.id, rel.target]))
  const links: ExternalLinkPart[] = []
  for (const match of workbook.matchAll(/<externalReference\b[^>]*>/g)) {
    const target = targets.get(attribute(match[0], 'r:id') ?? '')
    const xml = target ? await zip.file(target)?.async('text') : undefined
    if (!target || !xml) continue
    const rels = await zip.file(relsPathOf(target))?.async('text')
    links.push({ xml, ...(rels ? { rels } : {}) })
  }
  return links
}

/** Adds the links to a file that has none, as `[1]`, `[2]`, … in the given order. */
export async function applyExternalLinks(bytes: Uint8Array, links: readonly ExternalLinkPart[]): Promise<Uint8Array> {
  if (links.length === 0) return bytes
  const zip = await JSZip.loadAsync(bytes)
  let workbook = (await zip.file('xl/workbook.xml')?.async('text')) ?? ''
  let workbookRels = (await zip.file('xl/_rels/workbook.xml.rels')?.async('text')) ?? ''
  let types = (await zip.file('[Content_Types].xml')?.async('text')) ?? ''
  if (/<externalReferences\b/.test(workbook)) return bytes
  let relNumber = Math.max(0, ...[...workbookRels.matchAll(/\bId="rId(\d+)"/g)].map((match) => Number(match[1])))
  const references: string[] = []
  links.forEach((link, index) => {
    const path = `xl/externalLinks/externalLink${index + 1}.xml`
    zip.file(path, link.xml)
    if (link.rels) zip.file(relsPathOf(path), link.rels)
    relNumber += 1
    workbookRels = workbookRels.replace(
      '</Relationships>',
      `<Relationship Id="rId${relNumber}" Type="${LINK_REL}" Target="externalLinks/externalLink${index + 1}.xml"/></Relationships>`,
    )
    types = types.replace('</Types>', `<Override PartName="/${path}" ContentType="${LINK_TYPE}"/></Types>`)
    references.push(`<externalReference r:id="rId${relNumber}"/>`)
  })
  // Schema order: externalReferences follows sheets (and functionGroups).
  const element = `<externalReferences>${references.join('')}</externalReferences>`
  workbook = /<\/functionGroups>|<functionGroups\b[^>]*\/>/.test(workbook)
    ? workbook.replace(/(<\/functionGroups>|<functionGroups\b[^>]*\/>)/, `$1${element}`)
    : workbook.replace(/(<\/sheets>)/, `$1${element}`)
  if (!/xmlns:r=/.test(/<workbook\b[^>]*>/.exec(workbook)?.[0] ?? '')) {
    workbook = workbook.replace(/<workbook\b/, '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
  }
  zip.file('xl/workbook.xml', workbook)
  zip.file('xl/_rels/workbook.xml.rels', workbookRels)
  zip.file('[Content_Types].xml', types)
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
}
