// The file's document properties (File → Properties in Excel: title, author,
// keywords, company, …) from docProps/core.xml and docProps/app.xml. GenOffice's
// gateway writes neither part, so they are kept with the stored workbook and
// written into every file built from it.
import JSZip from 'jszip'

export interface DocumentProperties {
  title?: string
  subject?: string
  creator?: string
  keywords?: string
  description?: string
  lastModifiedBy?: string
  category?: string
  contentStatus?: string
  /** W3CDTF, e.g. 2024-05-01T08:00:00Z. */
  created?: string
  company?: string
  manager?: string
}

const CORE_TYPE = 'application/vnd.openxmlformats-package.core-properties+xml'
const APP_TYPE = 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
const CORE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
const APP_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'

/** Property → element in core.xml or app.xml. */
const CORE: Array<[keyof DocumentProperties, string]> = [
  ['title', 'dc:title'],
  ['subject', 'dc:subject'],
  ['creator', 'dc:creator'],
  ['keywords', 'cp:keywords'],
  ['description', 'dc:description'],
  ['lastModifiedBy', 'cp:lastModifiedBy'],
  ['category', 'cp:category'],
  ['contentStatus', 'cp:contentStatus'],
]
const APP: Array<[keyof DocumentProperties, string]> = [
  ['manager', 'Manager'],
  ['company', 'Company'],
]

const unescapeXml = (text: string) =>
  text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, decimal: string) => String.fromCodePoint(Number(decimal)))
    .replace(/&amp;/g, '&')
const escapeXml = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

function elementText(xml: string, name: string): string | undefined {
  const match = new RegExp(`<${name}(?:\\s[^>]*)?>([\\s\\S]*?)</${name}>`).exec(xml)
  const text = match ? unescapeXml(match[1]!) : undefined
  return text ? text : undefined
}

/** The properties a package holds; undefined when it holds none. */
export async function extractDocProps(bytes: Uint8Array): Promise<DocumentProperties | undefined> {
  const zip = await JSZip.loadAsync(bytes)
  const core = (await zip.file('docProps/core.xml')?.async('text')) ?? ''
  const app = (await zip.file('docProps/app.xml')?.async('text')) ?? ''
  const properties: DocumentProperties = {}
  for (const [key, name] of CORE) {
    const value = elementText(core, name)
    if (value !== undefined) properties[key] = value
  }
  const created = elementText(core, 'dcterms:created')
  if (created) properties.created = created
  for (const [key, name] of APP) {
    const value = elementText(app, name)
    if (value !== undefined) properties[key] = value
  }
  return Object.keys(properties).length ? properties : undefined
}

/**
 * Writes docProps/core.xml and docProps/app.xml (replacing any) with the stored
 * properties; `modified` is the time the file is built, `created` falls back to it.
 */
export async function applyDocProps(bytes: Uint8Array, properties: DocumentProperties, modified: Date): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  const now = modified.toISOString().replace(/\.\d{3}Z$/, 'Z')
  const core =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
    'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    CORE.map(([key, name]) => (properties[key] ? `<${name}>${escapeXml(properties[key]!)}</${name}>` : '')).join('') +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${escapeXml(properties.created ?? now)}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
    '</cp:coreProperties>'
  const app =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    '<Application>Mindoo TeamGrid</Application>' +
    APP.map(([key, name]) => (properties[key] ? `<${name}>${escapeXml(properties[key]!)}</${name}>` : '')).join('') +
    '</Properties>'
  zip.file('docProps/core.xml', core)
  zip.file('docProps/app.xml', app)

  let types = (await zip.file('[Content_Types].xml')?.async('text')) ?? ''
  for (const [part, type] of [
    ['/docProps/core.xml', CORE_TYPE],
    ['/docProps/app.xml', APP_TYPE],
  ] as const) {
    if (!types.includes(`PartName="${part}"`)) types = types.replace('</Types>', `<Override PartName="${part}" ContentType="${type}"/></Types>`)
  }
  zip.file('[Content_Types].xml', types)

  let rels = (await zip.file('_rels/.rels')?.async('text')) ?? ''
  let next = Math.max(0, ...[...rels.matchAll(/\bId="rId(\d+)"/g)].map((match) => Number(match[1]))) + 1
  for (const [type, target] of [
    [CORE_REL, 'docProps/core.xml'],
    [APP_REL, 'docProps/app.xml'],
  ] as const) {
    if (!rels.includes(`Type="${type}"`)) rels = rels.replace('</Relationships>', `<Relationship Id="rId${next++}" Type="${type}" Target="${target}"/></Relationships>`)
  }
  zip.file('_rels/.rels', rels)
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
}
