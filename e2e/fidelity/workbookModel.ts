/**
 * What a user sees in an xlsx, read with our own parser (not TeamGrid's reader, so a
 * reader bug cannot hide on both sides): cell values and formulas, resolved cell formats,
 * layout, sheet features and the objects on each sheet. The round-trip test compares
 * this model before and after TeamGrid imported and exported a workbook.
 */
import { XMLParser } from 'fast-xml-parser'
import JSZip from 'jszip'

import { offsetFormulaRefs } from '../../vendor/genoffice/packages/xlsx-gateway/src/domain/formula-shift'

export interface SheetModel {
  name: string
  state: string
  /** address → `v:<value>` or `f:<formula>` */
  cells: Map<string, string>
  /** address → resolved format, only for cells whose format is not the default */
  formats: Map<string, string>
  merges: string[]
  /** column number → `width/hidden` for columns with a custom width or hidden */
  columns: Map<number, string>
  /** row number → `height/hidden` for rows with a custom height or hidden */
  rows: Map<number, string>
  view: string
  features: Map<string, string[]>
}

export interface WorkbookModel {
  sheets: SheetModel[]
  definedNames: string[]
  /** content type → number of parts, e.g. charts, drawings, comments, tables */
  partTypes: Map<string, number>
}

const ARRAYS = new Set([
  'sheet', 'row', 'c', 'si', 'r', 'Relationship', 'Override', 'Default', 'xf', 'font', 'fill', 'border', 'numFmt',
  'mergeCell', 'col', 'definedName', 'dataValidation', 'conditionalFormatting', 'cfRule', 'hyperlink', 'pane',
  'selection', 'comment', 'tablePart', 'formula', 'sheetView', 'brk',
])
const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  textNodeName: '#text',
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
  isArray: (name, _path, _leaf, isAttribute) => !isAttribute && ARRAYS.has(name),
})

type Xml = Record<string, any>

const text = (node: unknown): string =>
  node === undefined || node === null ? '' : typeof node === 'object' ? String((node as Xml)['#text'] ?? '') : String(node)

/** Excel's built-in number formats that matter for display. */
const BUILTIN_FORMATS: Record<string, string> = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?',
  13: '# ??/??', 14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM',
  20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yyyy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0',
  49: '@',
}

/** Colors that all mean the default (automatic) black. */
const DEFAULT_COLORS = new Set(['', 'auto', '000000', 'theme1', 'idx8'])

function color(node: Xml | undefined): string {
  const value = rawColor(node)
  return DEFAULT_COLORS.has(value) ? '' : value
}

function rawColor(node: Xml | undefined): string {
  if (!node) return ''
  if (node.rgb) return String(node.rgb).toUpperCase().replace(/^FF(?=[0-9A-F]{6}$)/, '')
  if (node.theme !== undefined) return `theme${node.theme}${node.tint ? `~${Number(node.tint).toFixed(2)}` : ''}`
  if (node.indexed !== undefined) return node.indexed === '64' ? '' : `idx${node.indexed}`
  return node.auto ? 'auto' : ''
}

const flag = (node: unknown) => node !== undefined && (node as Xml)?.val !== '0' && (node as Xml)?.val !== 'false'

/** One string per cell format (xf index): number format, font, fill, borders, alignment. */
function formatSignatures(styles: Xml | undefined): string[] {
  const sheet = styles?.styleSheet
  if (!sheet) return []
  const numFmts = new Map<string, string>(
    (sheet.numFmts?.numFmt ?? []).map((n: Xml) => [String(n.numFmtId), String(n.formatCode)]),
  )
  const fonts: Xml[] = sheet.fonts?.font ?? []
  const fills: Xml[] = sheet.fills?.fill ?? []
  const borders: Xml[] = sheet.borders?.border ?? []
  const fontOf = (f: Xml | undefined) =>
    f
      ? [
          f.name?.val ?? '',
          f.sz?.val ?? '',
          flag(f.b) ? 'b' : '',
          flag(f.i) ? 'i' : '',
          f.u ? `u${f.u.val ?? 'single'}` : '',
          flag(f.strike) ? 's' : '',
          f.vertAlign?.val ?? '',
          color(f.color),
        ].join(',')
      : ''
  const fillOf = (f: Xml | undefined) => {
    const p = f?.patternFill
    if (f?.gradientFill) return 'gradient'
    if (!p || !p.patternType || p.patternType === 'none') return ''
    return `${p.patternType}:${color(p.fgColor)}:${p.patternType === 'solid' ? '' : color(p.bgColor)}`
  }
  const borderOf = (b: Xml | undefined) =>
    b
      ? ['left', 'right', 'top', 'bottom', 'diagonal']
          .map((side) => (b[side]?.style ? `${side[0]}${b[side].style}${color(b[side].color)}` : ''))
          .filter(Boolean)
          .join(' ')
      : ''
  return (sheet.cellXfs?.xf ?? []).map((xf: Xml) => {
    const a = xf.alignment ?? {}
    const id = String(xf.numFmtId ?? '0')
    return [
      numFmts.get(id) ?? BUILTIN_FORMATS[id] ?? `fmt${id}`,
      fontOf(fonts[Number(xf.fontId ?? 0)]),
      fillOf(fills[Number(xf.fillId ?? 0)]),
      borderOf(borders[Number(xf.borderId ?? 0)]),
      [a.horizontal ?? '', a.vertical ?? '', a.wrapText === '1' || a.wrapText === 'true' ? 'wrap' : '', a.textRotation ?? '', a.indent ?? ''].join(','),
    ].join(' | ')
  })
}

async function xml(zip: JSZip, path: string): Promise<Xml | undefined> {
  const file = zip.file(path)
  return file ? parser.parse(await file.async('text')) : undefined
}

function resolve(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const parts = base.split('/').slice(0, -1)
  for (const segment of target.split('/')) {
    if (segment === '..') parts.pop()
    else if (segment !== '.') parts.push(segment)
  }
  return parts.join('/')
}

async function relationships(zip: JSZip, part: string): Promise<Array<{ id: string; type: string; target: string; external: boolean }>> {
  const dir = part.split('/').slice(0, -1).join('/')
  const name = part.split('/').pop()
  const rels = await xml(zip, `${dir}/_rels/${name}.rels`)
  return (rels?.Relationships?.Relationship ?? []).map((r: Xml) => ({
    id: String(r.Id),
    type: String(r.Type).split('/').pop()!,
    target: r.TargetMode === 'External' ? String(r.Target) : resolve(part, String(r.Target)),
    external: r.TargetMode === 'External',
  }))
}

const normalizeFormula = (formula: string) => formula.replace(/"(?:[^"]|"")*"|\s+/g, (m) => (m.startsWith('"') ? m : ''))

function value(c: Xml, shared: string[]): string {
  const raw = c.v === undefined ? undefined : text(c.v)
  if (c.t === 'inlineStr') return `s:${text(c.is?.t) || (c.is?.r ?? []).map((r: Xml) => text(r.t)).join('')}`
  if (raw === undefined) return ''
  if (c.t === 's') return `s:${shared[Number(raw)] ?? ''}`
  if (c.t === 'str') return `s:${raw}`
  if (c.t === 'b') return `b:${raw === '1' ? 'TRUE' : 'FALSE'}`
  if (c.t === 'e') return `e:${raw}`
  const number = Number(raw)
  // 15 significant digits, as Excel shows them
  return `n:${Number.isFinite(number) ? Number(number.toPrecision(15)) : raw}`
}

export async function readWorkbookModel(bytes: Uint8Array): Promise<WorkbookModel> {
  const zip = await JSZip.loadAsync(bytes)
  const types = await xml(zip, '[Content_Types].xml')
  const overrides = new Map<string, string>(
    (types?.Types?.Override ?? []).map((o: Xml) => [String(o.PartName).replace(/^\//, ''), String(o.ContentType)]),
  )
  const defaults = new Map<string, string>((types?.Types?.Default ?? []).map((d: Xml) => [String(d.Extension).toLowerCase(), String(d.ContentType)]))
  const partTypes = new Map<string, number>()
  for (const path of Object.keys(zip.files)) {
    if (zip.files[path]!.dir || path.endsWith('.rels') || path === '[Content_Types].xml') continue
    const type = overrides.get(path) ?? defaults.get(path.split('.').pop()!.toLowerCase()) ?? 'unknown'
    const short = type.replace(/^application\/vnd\.(openxmlformats-officedocument\.|ms-(office|excel)\.)?/, '').replace(/\+xml$/, '')
    // Pictures count as one kind, whatever their format; shared strings and the calculation
    // chain are storage choices, not content
    if (/sharedStrings|calcChain/.test(short)) continue
    const key = /^image\//.test(type) ? 'image' : short
    partTypes.set(key, (partTypes.get(key) ?? 0) + 1)
  }

  const workbookPath = 'xl/workbook.xml'
  const workbook = (await xml(zip, workbookPath))?.workbook
  const workbookRels = await relationships(zip, workbookPath)
  const sharedPath = workbookRels.find((r) => r.type === 'sharedStrings')?.target
  const shared: string[] = sharedPath
    ? ((await xml(zip, sharedPath))?.sst?.si ?? []).map((si: Xml) =>
        si.t !== undefined ? text(si.t) : (si.r ?? []).map((r: Xml) => text(r.t)).join(''),
      )
    : []
  const stylesPath = workbookRels.find((r) => r.type === 'styles')?.target
  const formats = formatSignatures(stylesPath ? await xml(zip, stylesPath) : undefined)
  const defaultFormat = formats[0] ?? ''

  const sheetEntries: Xml[] = [workbook?.sheets?.sheet ?? []].flat()
  const definedNames = (workbook?.definedNames?.definedName ?? [])
    .filter((d: Xml) => !String(d.name).startsWith('_xlnm._FilterDatabase'))
    .map((d: Xml) => `${d.localSheetId !== undefined ? `${sheetEntries[Number(d.localSheetId)]?.name ?? d.localSheetId}!` : ''}${d.name}=${normalizeFormula(text(d))}`)
    .sort()

  const sheets: SheetModel[] = []
  for (const entry of sheetEntries) {
    const target = workbookRels.find((r) => r.id === entry['r:id'])
    if (!target || target.type !== 'worksheet') continue
    const worksheet = (await xml(zip, target.target))?.worksheet
    if (!worksheet) continue
    const rels = await relationships(zip, target.target)
    const sheet: SheetModel = {
      name: String(entry.name),
      state: String(entry.state ?? 'visible'),
      cells: new Map(),
      formats: new Map(),
      merges: (worksheet.mergeCells?.mergeCell ?? []).map((m: Xml) => String(m.ref)).sort(),
      columns: new Map(),
      rows: new Map(),
      view: '',
      features: new Map(),
    }
    const masters = new Map<string, { formula: string; row: number; column: number }>()
    for (const row of worksheet.sheetData?.row ?? []) {
      const rowNumber = Number(row.r)
      if (row.customHeight === '1' || row.customHeight === 'true' || row.hidden === '1' || row.hidden === 'true') {
        sheet.rows.set(rowNumber, `${row.customHeight === '1' || row.customHeight === 'true' ? Number(row.ht).toFixed(1) : ''}/${row.hidden === '1' || row.hidden === 'true' ? 'hidden' : ''}`)
      }
      for (const c of row.c ?? []) {
        const address = String(c.r)
        const match = /^([A-Z]+)(\d+)$/.exec(address)!
        const column = [...match[1]!].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)
        if (c.f !== undefined) {
          let formula = text(c.f)
          const f = c.f as Xml
          if (typeof f === 'object' && f.t === 'shared' && f.si !== undefined) {
            if (formula) masters.set(String(f.si), { formula, row: rowNumber, column })
            else {
              const master = masters.get(String(f.si))
              formula = master ? offsetFormulaRefs(master.formula, rowNumber - master.row, column - master.column) : ''
            }
          }
          // Array (CSE) formulas keep their range: `{=SUM(A1:B1*A2:B2)}` over A1:A3
          const array = typeof f === 'object' && f.t === 'array' ? `{=}${f.ref ?? address}:` : ''
          sheet.cells.set(address, `f:${array}${normalizeFormula(formula)}`)
        } else {
          const v = value(c, shared)
          if (v) sheet.cells.set(address, `v:${v}`)
        }
        const format = formats[Number(c.s ?? 0)] ?? defaultFormat
        if (format !== defaultFormat) sheet.formats.set(address, format)
      }
    }
    for (const col of worksheet.cols?.col ?? []) {
      const custom = col.customWidth === '1' || col.customWidth === 'true'
      const hidden = col.hidden === '1' || col.hidden === 'true'
      if (!custom && !hidden) continue
      for (let n = Number(col.min); n <= Math.min(Number(col.max), Number(col.min) + 200); n++) {
        sheet.columns.set(n, `${custom ? Number(col.width).toFixed(2) : ''}/${hidden ? 'hidden' : ''}`)
      }
    }
    const view = worksheet.sheetViews?.sheetView?.[0]
    const pane = view?.pane?.[0]
    sheet.view = [
      pane?.state === 'frozen' || pane?.state === 'frozenSplit' ? `frozen ${pane.xSplit ?? 0},${pane.ySplit ?? 0}` : '',
      view?.showGridLines === '0' || view?.showGridLines === 'false' ? 'no-gridlines' : '',
      view?.rightToLeft === '1' ? 'rtl' : '',
      worksheet.sheetPr?.tabColor ? `tab ${color(worksheet.sheetPr.tabColor)}` : '',
    ]
      .filter(Boolean)
      .join(' ')
    const features = sheet.features
    const add = (key: string, values: string[]) => {
      if (values.length) features.set(key, values.sort())
    }
    add('autoFilter', worksheet.autoFilter ? [String(worksheet.autoFilter.ref)] : [])
    add(
      'dataValidation',
      (worksheet.dataValidations?.dataValidation ?? []).map(
        (d: Xml) => `${d.sqref} ${d.type ?? 'any'} ${d.operator ?? ''} ${normalizeFormula(text(d.formula1))} ${normalizeFormula(text(d.formula2))}`.trim(),
      ),
    )
    add(
      'conditionalFormat',
      (worksheet.conditionalFormatting ?? []).flatMap((cf: Xml) =>
        (cf.cfRule ?? []).map((rule: Xml) => `${cf.sqref} ${rule.type}${rule.operator ? ` ${rule.operator}` : ''}${(rule.formula ?? []).map((f: unknown) => ` ${normalizeFormula(text(f))}`).join('')}`),
      ),
    )
    add(
      'hyperlink',
      (worksheet.hyperlinks?.hyperlink ?? []).map((h: Xml) => {
        const rel = rels.find((r) => r.id === h['r:id'])
        return `${h.ref} ${rel ? rel.target : `#${h.location ?? ''}`}`
      }),
    )
    add('pageBreaks', [...(worksheet.rowBreaks?.brk ?? []).map((b: Xml) => `row ${b.id}`), ...(worksheet.colBreaks?.brk ?? []).map((b: Xml) => `col ${b.id}`)])
    add('protection', worksheet.sheetProtection ? ['protected'] : [])
    add('printSetup', worksheet.pageSetup?.orientation ? [`orientation ${worksheet.pageSetup.orientation}`] : [])
    add('headerFooter', worksheet.headerFooter ? [text(worksheet.headerFooter.oddHeader), text(worksheet.headerFooter.oddFooter)].filter(Boolean) : [])
    for (const rel of rels) {
      if (rel.type === 'table') {
        const table = (await xml(zip, rel.target))?.table
        add('table', [...(features.get('table') ?? []), `${table?.displayName ?? table?.name} ${table?.ref}`])
      } else if (rel.type === 'comments') {
        const comments = (await xml(zip, rel.target))?.comments?.commentList?.comment ?? []
        add('comment', comments.map((c: Xml) => String(c.ref)))
      } else if (rel.type === 'drawing') {
        const drawingRels = await relationships(zip, rel.target)
        add(
          'drawing',
          drawingRels.filter((r) => r.type === 'chart' || r.type === 'image').map((r) => r.type),
        )
      }
    }
    sheets.push(sheet)
  }
  return { sheets, definedNames, partTypes }
}

export type Differences = Record<string, string[]>

/**
 * What changed from `before` to `after`, by category; each entry names the place and
 * both states. Categories: sheets, values, formulas, formats, merges, columns, rows,
 * view, definedNames, parts and each sheet feature.
 */
export function compareModels(before: WorkbookModel, after: WorkbookModel): Differences {
  const out: Differences = {}
  const push = (category: string, entry: string) => (out[category] ??= []).push(entry)
  const sheetNames = (m: WorkbookModel) => m.sheets.map((s) => `${s.name}${s.state === 'visible' ? '' : ` (${s.state})`}`)
  if (sheetNames(before).join('\n') !== sheetNames(after).join('\n')) {
    push('sheets', `${sheetNames(before).join(', ')} → ${sheetNames(after).join(', ')}`)
  }
  for (const sheet of before.sheets) {
    const other = after.sheets.find((s) => s.name === sheet.name)
    if (!other) continue
    for (const [address, content] of sheet.cells) {
      const now = other.cells.get(address)
      if (now === content) continue
      const category = content.startsWith('f:') ? 'formulas' : 'values'
      push(category, `${sheet.name}!${address}: ${content.slice(2)} → ${now ? now.slice(2) : '(empty)'}`)
    }
    for (const [address, content] of other.cells) {
      if (!sheet.cells.has(address)) push(content.startsWith('f:') ? 'formulas' : 'values', `${sheet.name}!${address}: (empty) → ${content.slice(2)}`)
    }
    for (const [address, format] of sheet.formats) {
      const now = other.formats.get(address)
      if (now !== format) push('formats', `${sheet.name}!${address}: ${format} → ${now ?? '(default)'}`)
    }
    for (const [address, format] of other.formats) {
      if (!sheet.formats.has(address) && sheet.cells.has(address)) push('formats', `${sheet.name}!${address}: (default) → ${format}`)
    }
    const diffList = (category: string, a: string[], b: string[]) => {
      const missing = a.filter((x) => !b.includes(x))
      const added = b.filter((x) => !a.includes(x))
      for (const x of missing) push(category, `${sheet.name}: lost ${x}`)
      for (const x of added) push(category, `${sheet.name}: added ${x}`)
    }
    diffList('merges', sheet.merges, other.merges)
    const mapDiff = (category: string, a: Map<number, string>, b: Map<number, string>) => {
      for (const [key, v] of a) if (b.get(key) !== v) push(category, `${sheet.name}: ${key} ${v} → ${b.get(key) ?? '(default)'}`)
    }
    mapDiff('columns', sheet.columns, other.columns)
    mapDiff('rows', sheet.rows, other.rows)
    if (sheet.view !== other.view) push('view', `${sheet.name}: ${sheet.view || '(default)'} → ${other.view || '(default)'}`)
    for (const [feature, values] of sheet.features) diffList(feature, values, other.features.get(feature) ?? [])
    for (const [feature, values] of other.features) if (!sheet.features.has(feature)) diffList(feature, [], values)
  }
  const missingNames = before.definedNames.filter((n) => !after.definedNames.includes(n))
  for (const n of missingNames) push('definedNames', `lost ${n}`)
  for (const [type, count] of before.partTypes) {
    const now = after.partTypes.get(type) ?? 0
    if (now < count) push('parts', `${type}: ${count} → ${now}`)
  }
  return out
}
