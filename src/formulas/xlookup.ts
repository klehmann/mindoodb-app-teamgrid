/**
 * XLOOKUP as in Excel for one lookup value and vector ranges: Univer's next-smaller /
 * next-larger matches, its binary searches and its argument checks differ (it took the
 * first larger item instead of the smallest larger one, and answered #N/A where Excel
 * finds one). The search follows IronCalc (base/src/functions/xlookup.rs and
 * binary_search.rs, Apache-2.0/MIT), whose results match Excel's in its test workbook.
 * Several lookup values and two-dimensional return arrays stay with Univer.
 */
import { type ArrayValueObject, BaseFunction, type BaseValueObject, ErrorType, ErrorValueObject, NumberValueObject } from '@univerjs/engine-formula'

type Cell =
  | { kind: 'number'; value: number }
  | { kind: 'text'; value: string }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'empty' }
  | { kind: 'error'; value: string }

function cellOf(value: BaseValueObject | null | undefined): Cell {
  if (!value || value.isNull()) return { kind: 'empty' }
  if (value.isError()) return { kind: 'error', value: String(value.getValue()) }
  if (value.isNumber()) return { kind: 'number', value: Number(value.getValue()) }
  if (value.isBoolean()) return { kind: 'boolean', value: Boolean(value.getValue()) }
  return { kind: 'text', value: String(value.getValue()) }
}

const precise = (value: number) => Number(value.toPrecision(15))
const ERROR_RANK = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A']

function compareText(a: string, b: string): number {
  const x = a.toUpperCase()
  const y = b.toUpperCase()
  return x === y ? 0 : x < y ? -1 : 1
}

/**
 * Excel's comparison for the linear search: numbers < text < booleans, errors last; an
 * empty cell counts as 0, "" or FALSE, whichever it is compared with. -1, 0 or 1.
 */
export function compareValues(left: Cell, right: Cell): number {
  if (left.kind === 'empty' && right.kind === 'empty') return 0
  if (left.kind === 'empty') return compareValues(right.kind === 'text' ? { kind: 'text', value: '' } : right.kind === 'boolean' ? { kind: 'boolean', value: false } : { kind: 'number', value: 0 }, right)
  if (right.kind === 'empty') return compareValues(left, left.kind === 'text' ? { kind: 'text', value: '' } : left.kind === 'boolean' ? { kind: 'boolean', value: false } : { kind: 'number', value: 0 })
  if (left.kind === 'error' && right.kind === 'error') return Math.sign(ERROR_RANK.indexOf(left.value) - ERROR_RANK.indexOf(right.value))
  if (left.kind === 'error') return 1
  if (right.kind === 'error') return -1
  const rank = { number: 0, text: 1, boolean: 2 } as const
  if (left.kind !== right.kind) return rank[left.kind] < rank[right.kind] ? -1 : 1
  if (left.kind === 'number') {
    const a = precise(left.value)
    const b = precise((right as { value: number }).value)
    return Math.abs(a - b) < Number.EPSILON ? 0 : a < b ? -1 : 1
  }
  if (left.kind === 'text') return compareText(left.value, (right as { value: string }).value)
  const a = left.value
  const b = (right as { value: boolean }).value
  return a === b ? 0 : a ? 1 : -1
}

/** The order the binary searches assume: numbers < text < booleans < empty (errors after). */
function order(left: Cell, right: Cell): number {
  const rank = { number: 0, text: 1, boolean: 2, empty: 3, error: 4 } as const
  if (left.kind !== right.kind) return rank[left.kind] < rank[right.kind] ? -1 : 1
  if (left.kind === 'empty' || left.kind === 'error') return 0
  if (left.kind === 'number') {
    const b = (right as { value: number }).value
    return Math.abs(b - left.value) < Number.EPSILON ? 0 : left.value < b ? -1 : 1
  }
  if (left.kind === 'text') return compareText(left.value, (right as { value: string }).value)
  const b = (right as { value: boolean }).value
  return left.value === b ? 0 : left.value ? 1 : -1
}

/** Ascending: the match or the largest smaller item. */
function ascendingOrSmaller(target: Cell, array: Cell[]): number | undefined {
  let l = 0
  let r = array.length
  while (l < r) {
    const m = (l + r) >> 1
    if (order(array[m]!, target) < 0) l = m + 1
    else r = m
  }
  if (l === array.length) return l - 1
  if (order(array[l]!, target) === 0) return l
  return l === 0 ? undefined : l - 1
}

/** Ascending: the match or the smallest larger item. */
function ascendingOrGreater(target: Cell, array: Cell[]): number | undefined {
  let l = 0
  let r = array.length
  while (l < r) {
    let m = (l + r) >> 1
    const c = order(array[m]!, target)
    if (c < 0) l = m + 1
    else if (c > 0) r = m
    else {
      while (m > 1 && order(array[m - 1]!, target) === 0) m--
      return m
    }
  }
  return r === array.length ? undefined : r
}

/** Descending: the match (last of equal ones) or the largest smaller item. */
function descending(target: Cell, array: Cell[], greater: boolean): number | undefined {
  const n = array.length
  let l = 0
  let r = n
  while (l < r) {
    const m = (l + r) >> 1
    let index = n - m - 1
    const c = order(array[index]!, target)
    if (c < 0) l = m + 1
    else if (c > 0) r = m
    else {
      while (index < n - 1 && order(array[index + 1]!, target) === 0) index++
      return index
    }
  }
  if (greater) return r === n ? undefined : n - r - 1
  return l === 0 ? undefined : n - l
}

/** Excel's wildcards (* ? and ~ as escape), the whole text, ignoring case. */
function wildcard(pattern: string): RegExp {
  let source = ''
  for (let index = 0; index < pattern.length; index++) {
    const ch = pattern[index]!
    if (ch === '~' && index + 1 < pattern.length) source += pattern[++index]!.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    else if (ch === '*') source += '.*'
    else if (ch === '?') source += '.'
    else source += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  }
  return new RegExp(`^${source}$`, 'is')
}

function linear(target: Cell, array: Cell[], reverse: boolean, matchMode: number): number | undefined {
  const indexes = array.map((_, index) => (reverse ? array.length - 1 - index : index))
  if (matchMode === 2) {
    const matches = target.kind === 'text' ? ((pattern: RegExp) => (cell: Cell) => cell.kind === 'text' && pattern.test(cell.value))(wildcard(target.value)) : (cell: Cell) => compareValues(cell, target) === 0
    return indexes.find((index) => matches(array[index]!))
  }
  let found: number | undefined
  for (const index of indexes) {
    const c = compareValues(array[index]!, target)
    if (c === 0) return index
    if (matchMode !== 0 && c === matchMode && (found === undefined || compareValues(array[found]!, array[index]!) === matchMode)) found = index
  }
  return found
}

/** Index of the match in `array`, or undefined; errors are thrown as their result. */
export function xlookupIndex(target: Cell, array: Cell[], matchMode: number, searchMode: number): number | undefined {
  if (searchMode === 1 || searchMode === -1) return linear(target, array, searchMode === -1, matchMode)
  const index =
    matchMode === 1
      ? searchMode === 2 ? ascendingOrGreater(target, array) : descending(target, array, true)
      : searchMode === 2 ? ascendingOrSmaller(target, array) : descending(target, array, false)
  if (index === undefined) return undefined
  if (matchMode === 0) return compareValues(array[index]!, target) === 0 ? index : undefined
  if (matchMode === 2) throw ErrorValueObject.create(ErrorType.VALUE)
  return index
}

const isVector = (array: ArrayValueObject) => array.getRowCount() === 1 || array.getColumnCount() === 1

function vector(array: ArrayValueObject): BaseValueObject[] {
  const values: BaseValueObject[] = []
  array.iterator((value) => {
    values.push(value as BaseValueObject)
  })
  return values
}

/** A mode argument: omitted or empty = the default; a number (floored); else undefined (error). */
function mode(value: BaseValueObject | undefined, fallback: number): number | BaseValueObject {
  let one = value
  if (one?.isArray()) {
    const array = one as ArrayValueObject
    if (array.getRowCount() !== 1 || array.getColumnCount() !== 1) return ErrorValueObject.create(ErrorType.VALUE)
    one = array.get(0, 0) ?? undefined
  }
  if (!one || one.isNull()) return fallback
  if (one.isError()) return one
  const number = one.isBoolean() ? Number(one.getValue()) : Number(one.getValue())
  if (!Number.isFinite(number) || (one.isString() && String(one.getValue()).trim() === '')) return ErrorValueObject.create(ErrorType.VALUE)
  return Math.floor(number)
}

export function createXlookup(base: (name: string) => new (name: string) => BaseFunction): BaseFunction {
  const Base = base('XLOOKUP')
  return new (class extends Base {
    override calculate(
      lookupValue: BaseValueObject,
      lookupArray: BaseValueObject,
      returnArray: BaseValueObject,
      ifNotFound?: BaseValueObject,
      matchModeArg?: BaseValueObject,
      searchModeArg?: BaseValueObject,
    ) {
      const fallback = () => super.calculate(lookupValue, lookupArray, returnArray, ifNotFound as BaseValueObject, matchModeArg as BaseValueObject, searchModeArg as BaseValueObject)
      let target: BaseValueObject | undefined = lookupValue
      if (target.isArray()) {
        const array = target as ArrayValueObject
        if (array.getRowCount() !== 1 || array.getColumnCount() !== 1) return fallback()
        target = array.get(0, 0) ?? undefined
      }
      if (!lookupArray.isArray() || !returnArray.isArray()) return fallback()
      const lookups = lookupArray as ArrayValueObject
      const returns = returnArray as ArrayValueObject
      if (!isVector(lookups)) return ErrorValueObject.create(ErrorType.VALUE)
      const byRow = lookups.getColumnCount() === 1 && lookups.getRowCount() > 1
      const length = byRow ? lookups.getRowCount() : lookups.getColumnCount()
      // A two-dimensional return array gives whole rows or columns: Univer's part.
      if (!isVector(returns) || (byRow ? returns.getColumnCount() !== 1 : returns.getRowCount() !== 1)) return fallback()
      if ((byRow ? returns.getRowCount() : returns.getColumnCount()) !== length) return ErrorValueObject.create(ErrorType.VALUE)
      if (target?.isError()) return target
      const matchMode = mode(matchModeArg, 0)
      if (typeof matchMode !== 'number') return matchMode
      if (![-1, 0, 1, 2].includes(matchMode)) return ErrorValueObject.create(ErrorType.VALUE)
      const searchMode = mode(searchModeArg, 1)
      if (typeof searchMode !== 'number') return searchMode
      if (![-2, -1, 1, 2].includes(searchMode)) return ErrorValueObject.create(ErrorType.VALUE)
      let index: number | undefined
      try {
        index = xlookupIndex(cellOf(target), vector(lookups).map(cellOf), matchMode, searchMode)
      } catch (error) {
        if (error instanceof ErrorValueObject) return error
        throw error
      }
      if (index === undefined) {
        if (!ifNotFound || ifNotFound.isNull()) return ErrorValueObject.create(ErrorType.NA)
        return ifNotFound
      }
      const found = vector(returns)[index]
      return !found || found.isNull() ? NumberValueObject.create(0) : found
    }
  })('XLOOKUP')
}
