// A1 formula text ↔ stored formulas with id-bound references.
//
// Stored formulas keep their references as row/column ids, so a row another
// person inserts above a referenced cell moves the reference with it instead
// of leaving it pointing at the old position. At load time the references are
// rendered back to A1 against the current row/column order.
//
// A reference outside the materialized grid (no id at that position) stays
// plain text; so do defined names, structured table references and anything
// else that is not an A1 cell, range, column or row reference.
import type { ColumnId, Formula, FormulaRef, RowId, SheetId } from './schema'

/** Index ↔ id lookup of one sheet's live rows and columns. */
export interface SheetAxes {
  id: SheetId
  name: string
  rowIds: readonly RowId[]
  columnIds: readonly ColumnId[]
  /** Full order lists including deleted ids, to shrink ranges past a deleted corner. */
  rowOrder?: readonly RowId[]
  columnOrder?: readonly ColumnId[]
}

export interface AxesLookup {
  bySheetId(id: SheetId): SheetAxes | undefined
  bySheetName(name: string): SheetAxes | undefined
}

export function createAxesLookup(sheets: readonly SheetAxes[]): AxesLookup {
  const byId = new Map(sheets.map((sheet) => [sheet.id, sheet]))
  const byName = new Map(sheets.map((sheet) => [sheet.name.toLowerCase(), sheet]))
  return {
    bySheetId: (id) => byId.get(id),
    bySheetName: (name) => byName.get(name.toLowerCase()),
  }
}

/** Position lookups built once per sheet: id → index. */
const indexCache = new WeakMap<readonly string[], Map<string, number>>()

function indexOf(ids: readonly string[], id: string | undefined): number {
  if (id === undefined) return -1
  let map = indexCache.get(ids)
  if (!map) {
    map = new Map(ids.map((value, index) => [value, index]))
    indexCache.set(ids, map)
  }
  return map.get(id) ?? -1
}

// sheet prefix: 'quoted name'! or bare name!
const SHEET = String.raw`(?:'((?:[^']|'')+)'|([A-Za-z_À-￿][\w.À-￿]*))!`
const CELL = String.raw`(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})`
const COLS = String.raw`(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})`
const ROWS = String.raw`(\$?)(\d{1,7}):(\$?)(\d{1,7})`
// One alternation, tried at every position the scanner stops at.
const REFERENCE = new RegExp(
  String.raw`^(?:${SHEET})?(?:${CELL}(?::${CELL})?|${COLS}|${ROWS})`,
)

/**
 * Index of a range corner. A deleted corner moves inwards to the nearest live
 * row/column (forward for a start corner, backward for an end corner).
 */
function cornerIndex(
  live: readonly string[],
  order: readonly string[] | undefined,
  id: string | undefined,
  direction: 1 | -1,
): number {
  const direct = indexOf(live, id)
  if (direct >= 0 || !order || id === undefined) return direct
  for (let position = order.indexOf(id) + direction; position >= 0 && position < order.length; position += direction) {
    const found = indexOf(live, order[position])
    if (found >= 0) return found
  }
  return -1
}

function columnIndex(label: string): number {
  let index = 0
  for (const char of label.toUpperCase()) index = index * 26 + (char.charCodeAt(0) - 64)
  return index - 1
}

function columnLabel(index: number): string {
  let label = ''
  for (let value = index + 1; value > 0; value = Math.floor((value - 1) / 26)) {
    label = String.fromCharCode(65 + ((value - 1) % 26)) + label
  }
  return label
}

/** Marks a reference's place in a stored formula's text. */
const REF_MARK = '\u0001'
const KIND_CODE: Record<FormulaRef['kind'], string> = { cell: 'c', range: 'g', columns: 'C', rows: 'R' }
const KIND_OF: Record<string, FormulaRef['kind']> = { c: 'cell', g: 'range', C: 'columns', R: 'rows' }

/** One reference as one string: `<kind><$ bits>|sheet|startRow|startColumn|endRow|endColumn`. */
export function packRef(ref: FormulaRef): string {
  const bits = ref.absolute.reduce((value, flag, index) => value | (flag ? 1 << index : 0), 0)
  return [
    `${KIND_CODE[ref.kind]}${bits.toString(16)}`,
    ref.sheetId ?? '',
    ref.startRowId ?? '',
    ref.startColumnId ?? '',
    ref.endRowId ?? '',
    ref.endColumnId ?? '',
  ].join('|')
}

export function unpackRef(packed: string): FormulaRef | null {
  const [head = '', sheetId, startRowId, startColumnId, endRowId, endColumnId] = packed.split('|')
  const kind = KIND_OF[head[0] ?? '']
  if (!kind) return null
  const bits = parseInt(head.slice(1), 16) || 0
  const ref: FormulaRef = { kind, absolute: [!!(bits & 1), !!(bits & 2), !!(bits & 4), !!(bits & 8)] }
  if (sheetId) ref.sheetId = sheetId
  if (startRowId) ref.startRowId = startRowId
  if (startColumnId) ref.startColumnId = startColumnId
  if (endRowId) ref.endRowId = endRowId
  if (endColumnId) ref.endColumnId = endColumnId
  return ref
}

/**
 * Splits A1 formula text into text and id-bound references, resolved against
 * the sheet the formula lives on (`home`) and the workbook's other sheets.
 */
export function parseFormula(source: string, home: SheetAxes, sheets: AxesLookup): Formula {
  let t = ''
  const r: string[] = []
  let index = 0
  while (index < source.length) {
    const char = source[index]!
    // string literals and structured references are copied verbatim
    if (char === '"' || char === '[') {
      const close = char === '"' ? '"' : ']'
      let end = index + 1
      while (end < source.length) {
        if (source[end] === close) {
          if (close === '"' && source[end + 1] === '"') {
            end += 2
            continue
          }
          break
        }
        end += 1
      }
      t += source.slice(index, end + 1)
      index = end + 1
      continue
    }
    const previous = index > 0 ? source[index - 1]! : ''
    if (/[\w.$\u00C0-\uFFFF]/.test(previous) && !/['!]/.test(previous)) {
      t += char
      index += 1
      continue
    }
    const match = REFERENCE.exec(source.slice(index))
    const after = match ? source[index + match[0].length] ?? '' : ''
    // A name like LOG10( or A1B is not a reference.
    if (!match || /[\w(.\u00C0-\uFFFF]/.test(after)) {
      t += char
      index += 1
      continue
    }
    const ref = toRef(match, home, sheets)
    if (ref) {
      t += REF_MARK
      r.push(packRef(ref))
    } else {
      t += match[0]
    }
    index += match[0].length
  }
  return { t, r }
}

function toRef(match: RegExpExecArray, home: SheetAxes, sheets: AxesLookup): FormulaRef | null {
  const sheetName = match[1]?.replace(/''/g, "'") ?? match[2]
  const target = sheetName === undefined ? home : sheets.bySheetName(sheetName)
  if (!target) return null
  const sheetId = target === home ? undefined : target.id
  const row = (digits: string) => target.rowIds[Number(digits) - 1]
  const column = (label: string) => target.columnIds[columnIndex(label)]
  const base = sheetId === undefined ? {} : { sheetId }
  if (match[4] !== undefined) {
    const startRowId = row(match[6]!)
    const startColumnId = column(match[4])
    if (!startRowId || !startColumnId) return null
    if (match[8] === undefined) {
      const absolute: FormulaRef['absolute'] = [match[5] === '$', match[3] === '$', match[5] === '$', match[3] === '$']
      return { ...base, kind: 'cell', startRowId, startColumnId, absolute }
    }
    const endRowId = row(match[10]!)
    const endColumnId = column(match[8])
    if (!endRowId || !endColumnId) return null
    return {
      ...base,
      kind: 'range',
      startRowId,
      startColumnId,
      endRowId,
      endColumnId,
      absolute: [match[5] === '$', match[3] === '$', match[9] === '$', match[7] === '$'],
    }
  }
  if (match[12] !== undefined) {
    const startColumnId = column(match[12])
    const endColumnId = column(match[14]!)
    if (!startColumnId || !endColumnId) return null
    return { ...base, kind: 'columns', startColumnId, endColumnId, absolute: [false, match[11] === '$', false, match[13] === '$'] }
  }
  const startRowId = row(match[16]!)
  const endRowId = row(match[18]!)
  if (!startRowId || !endRowId) return null
  return { ...base, kind: 'rows', startRowId, endRowId, absolute: [match[15] === '$', false, match[17] === '$', false] }
}

/** Renders a stored formula back to A1 text against the current row/column order. */
export function renderFormula(formula: Formula, home: SheetAxes, sheets: AxesLookup): string {
  const parts = formula.t.split(REF_MARK)
  let text = parts[0] ?? ''
  for (let index = 1; index < parts.length; index += 1) {
    const ref = unpackRef(formula.r[index - 1] ?? '')
    text += (ref ? renderRef(ref, home, sheets) : '#REF!') + parts[index]
  }
  return text
}

function quoteSheetName(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name)
    ? name
    : `'${name.replace(/'/g, "''")}'`
}

function renderRef(ref: FormulaRef, home: SheetAxes, sheets: AxesLookup): string {
  const target = ref.sheetId === undefined ? home : sheets.bySheetId(ref.sheetId)
  if (!target) return '#REF!'
  const prefix = target === home ? '' : `${quoteSheetName(target.name)}!`
  const [absStartRow, absStartColumn, absEndRow, absEndColumn] = ref.absolute
  const rowText = (index: number, absolute: boolean) => `${absolute ? '$' : ''}${index + 1}`
  const columnText = (index: number, absolute: boolean) => `${absolute ? '$' : ''}${columnLabel(index)}`
  let startRow = indexOf(target.rowIds, ref.startRowId)
  let endRow = indexOf(target.rowIds, ref.endRowId)
  let startColumn = indexOf(target.columnIds, ref.startColumnId)
  let endColumn = indexOf(target.columnIds, ref.endColumnId)
  switch (ref.kind) {
    case 'cell':
      if (startRow < 0 || startColumn < 0) return '#REF!'
      return `${prefix}${columnText(startColumn, absStartColumn)}${rowText(startRow, absStartRow)}`
    case 'columns':
      if (startColumn < 0 || endColumn < 0) return '#REF!'
      return `${prefix}${columnText(startColumn, absStartColumn)}:${columnText(endColumn, absEndColumn)}`
    case 'rows':
      if (startRow < 0 || endRow < 0) return '#REF!'
      return `${prefix}${rowText(startRow, absStartRow)}:${rowText(endRow, absEndRow)}`
    case 'range':
      // A deleted corner shrinks the range to what is left, as Excel does;
      // only a range with nothing left turns into #REF!.
      startRow = cornerIndex(target.rowIds, target.rowOrder, ref.startRowId, 1)
      endRow = cornerIndex(target.rowIds, target.rowOrder, ref.endRowId, -1)
      startColumn = cornerIndex(target.columnIds, target.columnOrder, ref.startColumnId, 1)
      endColumn = cornerIndex(target.columnIds, target.columnOrder, ref.endColumnId, -1)
      if (startRow < 0 || endRow < 0 || startColumn < 0 || endColumn < 0) return '#REF!'
      if (startRow > endRow || startColumn > endColumn) return '#REF!'
      return (
        `${prefix}${columnText(startColumn, absStartColumn)}${rowText(startRow, absStartRow)}` +
        `:${columnText(endColumn, absEndColumn)}${rowText(endRow, absEndRow)}`
      )
  }
}
