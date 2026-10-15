/**
 * The formula cells of an xlsx with the results Excel stored for them (`<c><f>…</f><v>…</v>`):
 * the expectations of the value tests. Read with our own small parser, not TeamGrid's
 * xlsx reader, so a reader bug cannot hide on both sides.
 */
import { XMLParser } from 'fast-xml-parser'
import JSZip from 'jszip'

export type CellResult = number | string | boolean | null

export interface FormulaCell {
  address: string
  /** The formula as written; for a shared formula's other cells the master's formula. */
  formula: string
  /** Excel's result: numbers as numbers, errors as their text (`#N/A`). */
  value: CellResult
  error: boolean
}

export interface WorkbookCells {
  /** `docProps/app.xml`'s Application and AppVersion: which program computed the results. */
  application: string
  sheets: Array<{ name: string; cells: FormulaCell[] }>
}

const ARRAYS = new Set(['sheet', 'row', 'c', 'si', 'r', 'Relationship'])
const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  textNodeName: '#text',
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
  // `r` and `si` are also attribute names (cell address, shared formula index)
  isArray: (name, _path, _leaf, isAttribute) => !isAttribute && ARRAYS.has(name),
})

type Node = string | { '#text'?: string; [key: string]: unknown } | undefined

function text(node: Node): string {
  if (node === undefined || node === null) return ''
  return typeof node === 'object' ? String(node['#text'] ?? '') : String(node)
}

async function xml(zip: JSZip, path: string): Promise<Record<string, any>> {
  const file = zip.file(path)
  if (!file) throw new Error(`The workbook has no ${path}`)
  return parser.parse(await file.async('text'))
}

export async function readWorkbookCells(bytes: Uint8Array): Promise<WorkbookCells> {
  const zip = await JSZip.loadAsync(bytes)
  const workbook = await xml(zip, 'xl/workbook.xml')
  const rels = await xml(zip, 'xl/_rels/workbook.xml.rels')
  const targets = new Map<string, string>(
    rels.Relationships.Relationship.map((r: { Id: string; Target: string }) => [r.Id, r.Target]),
  )
  const shared: string[] = zip.file('xl/sharedStrings.xml')
    ? ((await xml(zip, 'xl/sharedStrings.xml')).sst.si ?? []).map((si: { t?: Node; r?: Array<{ t?: Node }> }) =>
        si.t !== undefined ? text(si.t) : (si.r ?? []).map((run) => text(run.t)).join(''),
      )
    : []
  const app = zip.file('docProps/app.xml') ? await zip.file('docProps/app.xml')!.async('text') : ''
  const application = [/<Application>([^<]*)</.exec(app)?.[1], /<AppVersion>([^<]*)</.exec(app)?.[1]]
    .filter(Boolean)
    .join(' ')

  const sheets: WorkbookCells['sheets'] = []
  for (const sheet of [workbook.workbook.sheets.sheet].flat() as Array<Record<string, string>>) {
    const target = targets.get(sheet['r:id']!)
    if (!target) continue
    const path = target.startsWith('/') ? target.slice(1) : target.startsWith('xl/') ? target : `xl/${target}`
    if (!zip.file(path)) continue // chartsheets, dialog sheets
    const worksheet = (await xml(zip, path)).worksheet
    const masters = new Map<string, string>()
    const cells: FormulaCell[] = []
    for (const row of worksheet?.sheetData?.row ?? []) {
      for (const c of row.c ?? []) {
        if (c.f === undefined) continue
        const f = c.f as Node
        let formula = text(f)
        if (typeof f === 'object' && f.t === 'shared' && typeof f.si === 'string') {
          if (formula) masters.set(f.si, formula)
          else formula = masters.get(f.si) ?? ''
        }
        const raw = c.v === undefined ? null : text(c.v)
        let value: CellResult = raw
        if (raw !== null) {
          if (c.t === 's') value = shared[Number(raw)] ?? ''
          else if (c.t === 'b') value = raw === '1'
          else if (c.t !== 'str' && c.t !== 'e' && c.t !== 'inlineStr') value = Number(raw)
        } else if (c.t === 'str') {
          value = ''
        }
        cells.push({ address: c.r, formula, value, error: c.t === 'e' })
      }
    }
    sheets.push({ name: sheet.name!, cells })
  }
  return { application, sheets }
}

/** Functions a formula calls, upper case, without Excel's `_xlfn.`/`_xlws.` prefixes. */
export function functionsOf(formula: string): string[] {
  const code = formula.replace(/"(?:[^"]|"")*"/g, '""')
  const names = new Set<string>()
  for (const match of code.matchAll(/(?<![A-Za-z0-9._])(?:_xl[a-z]+\.)*([A-Za-z][A-Za-z0-9._]*)\s*\(/g)) {
    names.add(match[1]!.toUpperCase())
  }
  return [...names]
}

/** Results that differ on every calculation; only checked for being no error. */
export const VOLATILE = new Set(['RAND', 'RANDBETWEEN', 'RANDARRAY', 'NOW', 'TODAY', 'INFO', 'CELL'])
