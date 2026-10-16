// Legacy array (CSE) formulas, written straight into the package: GenOffice's save
// keeps `<f t="array" ref>` of a file it opened but cannot create one, and TeamGrid
// builds every file from scratch.
import JSZip from 'jszip'

const unescape = (text: string) =>
  text.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')

/** Marks formula cells as array formulas: sheet name → (address → ref, e.g. `A1` → `A1:A3`). */
export async function applyArrayFormulas(bytes: Uint8Array, arrays: ReadonlyMap<string, ReadonlyMap<string, string>>): Promise<Uint8Array> {
  if (arrays.size === 0) return bytes
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
  let changed = false
  for (const match of workbookXml.matchAll(/<sheet\b[^>]*>/g)) {
    const name = /\bname="([^"]*)"/.exec(match[0])?.[1]
    const relId = /\br:id="([^"]+)"/.exec(match[0])?.[1]
    const cells = name !== undefined ? arrays.get(unescape(name)) : undefined
    const path = relId ? targets.get(relId) : undefined
    if (!cells || !path) continue
    const xml = await zip.file(path)?.async('text')
    if (!xml) continue
    // GenOffice marks a formula with a spilling function as a one-cell array (ref = its own
    // address); the stored range wins there too.
    const next = xml.replace(/(<c\b[^>]*\br="([A-Z]+\d+)"[^>]*>)<f(?: t="array" ref="[^"]*")?>/g, (whole, open: string, address: string) => {
      const ref = cells.get(address)
      return ref ? `${open}<f t="array" ref="${ref}">` : whole
    })
    if (next !== xml) {
      zip.file(path, next)
      changed = true
    }
  }
  return changed ? new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })) : bytes
}
