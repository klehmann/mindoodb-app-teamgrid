// What the engine does not report about tables and the gateway's table
// addition cannot write, read from and written into the package directly:
// the first/last column and filter button switches (read only), and custom
// table styles, kept as their style elements with the formats (dxfs) inlined.
import JSZip from 'jszip'

import { attribute, escapeXml, setAttribute } from './pivot-parts'

/** A custom table style: its elements in file order, each with its format as raw `<dxf>` XML. */
export interface StoredTableStyle {
  name: string
  /** tableStyle/@pivot and @table as written; absent = the default (both true). */
  pivot?: boolean
  table?: boolean
  elements: { type: string; size?: number; dxf?: string }[]
}

export interface TableExtras {
  firstColumn: boolean
  lastColumn: boolean
  /** False when the table has no autoFilter (no filter buttons). */
  filterButton: boolean
  styleName?: string
}

export interface ExtractedTableParts {
  /** By lower-case table name. */
  tables: Map<string, TableExtras>
  /** Custom table styles by name. */
  styles: Map<string, StoredTableStyle>
}

const TABLE_PART = /^xl\/tables\/[^/]+\.xml$/
const isOn = (value: string | undefined) => value === '1' || value === 'true'

function dxfList(stylesXml: string): string[] {
  const block = /<dxfs\b[^>]*>([\s\S]*?)<\/dxfs>/.exec(stylesXml)?.[1] ?? ''
  return [...block.matchAll(/<dxf\b[^>]*\/>|<dxf\b[^>]*>[\s\S]*?<\/dxf>/g)].map((match) => match[0])
}

export async function extractTableParts(bytes: Uint8Array): Promise<ExtractedTableParts> {
  const zip = await JSZip.loadAsync(bytes)
  const tables = new Map<string, TableExtras>()
  for (const path of Object.keys(zip.files).filter((name) => TABLE_PART.test(name))) {
    const xml = (await zip.file(path)?.async('text')) ?? ''
    const head = /<(?:\w+:)?table\b[^>]*\sref="[^"]*"[^>]*>/.exec(xml)?.[0] ?? ''
    const name = attribute(head, 'displayName') ?? attribute(head, 'name')
    if (!name) continue
    const info = /<(?:\w+:)?tableStyleInfo\b[^>]*>/.exec(xml)?.[0] ?? ''
    const styleName = attribute(info, 'name')
    tables.set(name.toLowerCase(), {
      firstColumn: isOn(attribute(info, 'showFirstColumn')),
      lastColumn: isOn(attribute(info, 'showLastColumn')),
      filterButton: /<(?:\w+:)?autoFilter\b/.test(xml),
      ...(styleName ? { styleName } : {}),
    })
  }

  const styles = new Map<string, StoredTableStyle>()
  const stylesXml = (await zip.file('xl/styles.xml')?.async('text')) ?? ''
  const dxfs = dxfList(stylesXml)
  for (const match of stylesXml.matchAll(/<tableStyle\b[^>]*\/>|<tableStyle\b[^>]*>[\s\S]*?<\/tableStyle>/g)) {
    const head = /^<tableStyle\b[^>]*>/.exec(match[0])![0]
    const name = attribute(head, 'name')
    if (!name) continue
    const style: StoredTableStyle = { name, elements: [] }
    const pivot = attribute(head, 'pivot')
    const table = attribute(head, 'table')
    if (pivot !== undefined) style.pivot = isOn(pivot)
    if (table !== undefined) style.table = isOn(table)
    for (const element of match[0].matchAll(/<tableStyleElement\b[^>]*>/g)) {
      const type = attribute(element[0], 'type')
      if (!type) continue
      const size = attribute(element[0], 'size')
      const dxf = dxfs[Number(attribute(element[0], 'dxfId') ?? -1)]
      style.elements.push({ type, ...(size ? { size: Number(size) } : {}), ...(dxf ? { dxf } : {}) })
    }
    styles.set(name, style)
  }
  return { tables, styles }
}

/**
 * Adds the custom styles to styles.xml (formats appended to the dxfs) and
 * points each table using one (table name → style name) at it.
 */
export async function applyCustomTableStyles(
  bytes: Uint8Array,
  styles: readonly StoredTableStyle[],
  styleByTable: ReadonlyMap<string, string>,
): Promise<Uint8Array> {
  const used = new Set(styleByTable.values())
  const written = styles.filter((style) => used.has(style.name))
  if (written.length === 0) return bytes
  const zip = await JSZip.loadAsync(bytes)
  {
    let stylesXml = (await zip.file('xl/styles.xml')?.async('text')) ?? ''
    const dxfs = dxfList(stylesXml)
    const added: string[] = []
    const styleXml = written.map((style) => {
      const elements = style.elements.map((element) => {
        let dxfId = ''
        if (element.dxf) {
          dxfId = ` dxfId="${dxfs.length + added.length}"`
          added.push(element.dxf)
        }
        return `<tableStyleElement type="${escapeXml(element.type)}"${element.size ? ` size="${element.size}"` : ''}${dxfId}/>`
      })
      const flags =
        (style.pivot === undefined ? '' : ` pivot="${style.pivot ? 1 : 0}"`) +
        (style.table === undefined ? '' : ` table="${style.table ? 1 : 0}"`)
      return `<tableStyle name="${escapeXml(style.name)}"${flags} count="${elements.length}">${elements.join('')}</tableStyle>`
    })
    const allDxfs = [...dxfs, ...added]
    const dxfsXml = `<dxfs count="${allDxfs.length}">${allDxfs.join('')}</dxfs>`
    if (/<dxfs\b[^>]*\/>|<dxfs\b[^>]*>[\s\S]*?<\/dxfs>/.test(stylesXml)) {
      stylesXml = stylesXml.replace(/<dxfs\b[^>]*\/>|<dxfs\b[^>]*>[\s\S]*?<\/dxfs>/, dxfsXml)
    } else {
      // Schema order: dxfs follows cellStyles.
      stylesXml = stylesXml.replace(/(<\/cellStyles>|<cellStyles\b[^>]*\/>)/, `$1${dxfsXml}`)
    }
    const existing = /<tableStyles\b[^>]*\/>|<tableStyles\b[^>]*>([\s\S]*?)<\/tableStyles>/.exec(stylesXml)
    const kept = [...(existing?.[1] ?? '').matchAll(/<tableStyle\b[^>]*\/>|<tableStyle\b[^>]*>[\s\S]*?<\/tableStyle>/g)]
      .map((match) => match[0])
      .filter((xml) => !written.some((style) => attribute(xml, 'name') === style.name))
    const head = existing ? /^<tableStyles\b[^>]*?\/?>/.exec(existing[0])![0].replace(/\/>$/, '>') : '<tableStyles>'
    let tableStyles = setAttribute(head.replace(/>$/, ''), 'count', String(kept.length + styleXml.length)) + '>'
    tableStyles += [...kept, ...styleXml].join('') + '</tableStyles>'
    if (!/defaultTableStyle=/.test(tableStyles)) {
      tableStyles = tableStyles.replace(/^<tableStyles/, '<tableStyles defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"')
    }
    stylesXml = existing
      ? stylesXml.replace(existing[0], tableStyles)
      : stylesXml.replace(/(<dxfs\b[^>]*>[\s\S]*?<\/dxfs>)/, `$1${tableStyles}`)
    zip.file('xl/styles.xml', stylesXml)
  }

  const byName = new Map([...styleByTable].map(([table, style]) => [table.toLowerCase(), style]))
  for (const path of Object.keys(zip.files).filter((name) => TABLE_PART.test(name))) {
    let xml = (await zip.file(path)?.async('text')) ?? ''
    const head = /<(?:\w+:)?table\b[^>]*\sref="[^"]*"[^>]*>/.exec(xml)?.[0] ?? ''
    const style = byName.get((attribute(head, 'displayName') ?? attribute(head, 'name') ?? '').toLowerCase())
    if (!style) continue
    xml = xml.replace(/<(?:\w+:)?tableStyleInfo\b[^>]*?\/?>/, (info) => setAttribute(info, 'name', style))
    zip.file(path, xml)
  }
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
}
