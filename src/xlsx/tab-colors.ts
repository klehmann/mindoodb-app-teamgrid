// Sheet tab colors and veryHidden sheets, written straight into the package:
// GenOffice's save keeps them verbatim from the opened file but cannot set
// them, and TeamGrid builds every file from scratch.
import JSZip from 'jszip'

/** `#RRGGBB` / `RRGGBB` / `AARRGGBB` → OOXML's ARGB, or undefined. */
function argb(color: string): string | undefined {
  const hex = color.replace(/^#/, '').toUpperCase()
  if (/^[0-9A-F]{8}$/.test(hex)) return hex
  if (/^[0-9A-F]{6}$/.test(hex)) return `FF${hex}`
  return undefined
}

/** Sets the tab color of each named sheet (sheet name → color). */
export async function applyTabColors(bytes: Uint8Array, colors: ReadonlyMap<string, string>): Promise<Uint8Array> {
  if (colors.size === 0) return bytes
  const zip = await JSZip.loadAsync(bytes)
  const workbookXml = await zip.file('xl/workbook.xml')?.async('text')
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('text')
  if (!workbookXml || !relsXml) return bytes
  const targets = new Map<string, string>()
  for (const match of relsXml.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /\bId="([^"]+)"/.exec(match[0])?.[1]
    const target = /\bTarget="([^"]+)"/.exec(match[0])?.[1]
    if (id && target) targets.set(id, target.startsWith('/') ? target.slice(1) : `xl/${target}`)
  }
  const unescape = (text: string) =>
    text.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')
  let changed = false
  for (const match of workbookXml.matchAll(/<sheet\b[^>]*>/g)) {
    const name = /\bname="([^"]*)"/.exec(match[0])?.[1]
    const relId = /\br:id="([^"]+)"/.exec(match[0])?.[1]
    const color = name !== undefined ? colors.get(unescape(name)) : undefined
    const path = relId ? targets.get(relId) : undefined
    const rgb = color ? argb(color) : undefined
    if (!rgb || !path) continue
    const part = zip.file(path)
    const xml = await part?.async('text')
    if (!xml) continue
    const tab = `<tabColor rgb="${rgb}"/>`
    let next: string
    if (/<sheetPr\b[^>]*\/>/.test(xml)) {
      next = xml.replace(/<sheetPr\b([^>]*)\/>/, `<sheetPr$1>${tab}</sheetPr>`)
    } else if (/<sheetPr\b[^>]*>/.test(xml)) {
      next = xml.replace(/<tabColor\b[^>]*\/>/, '').replace(/(<sheetPr\b[^>]*>)/, `$1${tab}`)
    } else {
      next = xml.replace(/(<worksheet\b[^>]*>)/, `$1<sheetPr>${tab}</sheetPr>`)
    }
    zip.file(path, next)
    changed = true
  }
  return changed ? new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })) : bytes
}

/** Marks the named sheets veryHidden (only a macro can show them again). */
export async function applyVeryHidden(bytes: Uint8Array, names: ReadonlySet<string>): Promise<Uint8Array> {
  if (names.size === 0) return bytes
  const zip = await JSZip.loadAsync(bytes)
  const workbookXml = await zip.file('xl/workbook.xml')?.async('text')
  if (!workbookXml) return bytes
  const unescape = (text: string) =>
    text.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')
  const next = workbookXml.replace(/<sheet\b[^>]*?\/?>/g, (element) => {
    const name = /\bname="([^"]*)"/.exec(element)?.[1]
    if (name === undefined || !names.has(unescape(name))) return element
    const bare = element.replace(/\sstate="[^"]*"/, '')
    return bare.replace(/^<sheet\b/, '<sheet state="veryHidden"')
  })
  if (next === workbookXml) return bytes
  zip.file('xl/workbook.xml', next)
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
}
