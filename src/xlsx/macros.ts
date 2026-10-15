// VBA macros of an .xlsm: they never run in the browser, but the project
// (vbaProject.bin with its signatures) and the code names that bind its
// modules to the workbook and its sheets are kept, so an export is again an
// .xlsm with working macros.
import JSZip from 'jszip'

import { attribute, relationships, relsPathOf, setAttribute, sheetParts } from './pivot-parts'

export interface MacroParts {
  /** xl/vbaProject.bin, base64. */
  project: string
  /** Its signature parts (relationship type → base64). */
  signatures: { type: string; contentType: string; data: string }[]
  workbookCodeName?: string
  /** Sheet code names (Tabelle1, Sheet1, …) by sheet name; the model keys them by stored sheet id. */
  sheetCodeNames: Record<string, string>
}

const VBA_REL = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'
const VBA_TYPE = 'application/vnd.ms-office.vbaProject'
const MAIN = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
const MACRO_MAIN = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'

function toBase64(bytes: Uint8Array): string {
  let binary = ''
  for (let index = 0; index < bytes.length; index += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(index, index + 0x8000))
  }
  return btoa(binary)
}

function fromBase64(text: string): Uint8Array {
  const binary = atob(text)
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index)
  return bytes
}

/** Whether a package holds a VBA project (it then has to be saved as .xlsm). */
export function hasMacros(bytes: Uint8Array): boolean {
  const needle = 'xl/vbaProject.bin'
  outer: for (let index = 0; index <= bytes.length - needle.length; index += 1) {
    for (let offset = 0; offset < needle.length; offset += 1) {
      if (bytes[index + offset] !== needle.charCodeAt(offset)) continue outer
    }
    return true
  }
  return false
}

export async function extractMacros(bytes: Uint8Array): Promise<MacroParts | undefined> {
  const zip = await JSZip.loadAsync(bytes)
  const rel = (await relationships(zip, 'xl/workbook.xml')).find((entry) => entry.type === VBA_REL)
  const project = rel ? await zip.file(rel.target)?.async('uint8array') : undefined
  if (!rel || !project) return undefined
  const types = (await zip.file('[Content_Types].xml')?.async('text')) ?? ''
  const signatures: MacroParts['signatures'] = []
  for (const signature of await relationships(zip, rel.target)) {
    const data = await zip.file(signature.target)?.async('uint8array')
    const contentType = new RegExp(`PartName="/${signature.target}"[^>]*ContentType="([^"]+)"`).exec(types)?.[1]
    if (data && contentType) signatures.push({ type: signature.type, contentType, data: toBase64(data) })
  }
  const workbook = (await zip.file('xl/workbook.xml')?.async('text')) ?? ''
  const workbookCodeName = attribute(/<workbookPr\b[^>]*>/.exec(workbook)?.[0] ?? '', 'codeName')
  const sheetCodeNames: Record<string, string> = {}
  for (const [name, part] of await sheetParts(zip)) {
    const xml = (await zip.file(part)?.async('text')) ?? ''
    const codeName = attribute(/<sheetPr\b[^>]*>/.exec(xml)?.[0] ?? '', 'codeName')
    if (codeName) sheetCodeNames[name] = codeName
  }
  return { project: toBase64(project), signatures, ...(workbookCodeName ? { workbookCodeName } : {}), sheetCodeNames }
}

/** Puts the VBA project back and makes the package an .xlsm. */
export async function applyMacros(bytes: Uint8Array, macros: MacroParts): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes)
  zip.file('xl/vbaProject.bin', fromBase64(macros.project))
  let types = (await zip.file('[Content_Types].xml')?.async('text')) ?? ''
  types = types.replace(`ContentType="${MAIN}"`, `ContentType="${MACRO_MAIN}"`)
  if (!/<Default\b[^>]*Extension="bin"/.test(types)) {
    types = types.replace(/(<Types\b[^>]*>)/, `$1<Default Extension="bin" ContentType="${VBA_TYPE}"/>`)
  }
  if (macros.signatures.length) {
    const rels = macros.signatures.map((signature, index) => {
      const name = `vbaProjectSignature${index + 1}.bin`
      zip.file(`xl/${name}`, fromBase64(signature.data))
      types = types.replace('</Types>', `<Override PartName="/xl/${name}" ContentType="${signature.contentType}"/></Types>`)
      return `<Relationship Id="rId${index + 1}" Type="${signature.type}" Target="${name}"/>`
    })
    zip.file(
      relsPathOf('xl/vbaProject.bin'),
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`,
    )
  }
  zip.file('[Content_Types].xml', types)

  let workbookRels = (await zip.file('xl/_rels/workbook.xml.rels')?.async('text')) ?? ''
  const relNumber = Math.max(0, ...[...workbookRels.matchAll(/\bId="rId(\d+)"/g)].map((match) => Number(match[1]))) + 1
  workbookRels = workbookRels.replace('</Relationships>', `<Relationship Id="rId${relNumber}" Type="${VBA_REL}" Target="vbaProject.bin"/></Relationships>`)
  zip.file('xl/_rels/workbook.xml.rels', workbookRels)

  if (macros.workbookCodeName) {
    let workbook = (await zip.file('xl/workbook.xml')?.async('text')) ?? ''
    const codeName = macros.workbookCodeName
    workbook = /<workbookPr\b/.test(workbook)
      ? workbook.replace(/<workbookPr\b[^>]*?(\/?)>/, (element) => setAttribute(element.replace(/\/?>$/, ''), 'codeName', codeName) + (element.endsWith('/>') ? '/>' : '>'))
      : workbook.replace(
          // Schema order: workbookPr follows fileVersion and fileSharing.
          [/<\/fileSharing>|<fileSharing\b[^>]*\/>/, /<fileVersion\b[^>]*\/>/, /<workbook\b[^>]*>/].find((pattern) => pattern.test(workbook))!,
          (element) => `${element}<workbookPr codeName="${codeName}"/>`,
        )
    zip.file('xl/workbook.xml', workbook)
  }
  for (const [name, part] of await sheetParts(zip)) {
    const codeName = macros.sheetCodeNames[name]
    if (!codeName) continue
    let xml = (await zip.file(part)?.async('text')) ?? ''
    xml = /<sheetPr\b/.test(xml)
      ? xml.replace(/<sheetPr\b[^>]*?(\/?)>/, (element) => setAttribute(element.replace(/\/?>$/, ''), 'codeName', codeName) + (element.endsWith('/>') ? '/>' : '>'))
      : xml.replace(/(<worksheet\b[^>]*>)/, `$1<sheetPr codeName="${codeName}"/>`)
    zip.file(part, xml)
  }
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
}
