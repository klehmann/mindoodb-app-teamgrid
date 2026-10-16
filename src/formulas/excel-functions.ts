/**
 * Corrections to Univer's formula functions where TeamGrid's results differ from Excel's
 * (found by the fidelity tests, e2e/fidelity). Each one subclasses Univer's own executor
 * and changes only what differs, so argument handling and array broadcasting stay
 * Univer's; `installExcelFunctions` registers them on every Univer the editor creates.
 */
import type { Univer } from '@univerjs/core'
import {
  type ArrayValueObject,
  BaseFunction,
  type BaseValueObject,
  ErrorType,
  ErrorValueObject,
  functionCompatibility,
  functionDate,
  functionEngineering,
  functionFinancial,
  functionLookup,
  functionMath,
  functionStatistical,
  functionText,
  IFunctionService,
  NumberValueObject,
  StringValueObject,
} from '@univerjs/engine-formula'

import { criteriaMatcher, type CriteriaCell } from './criteria'
import { binomialCdf, binomialInverse, binomialPdf, negativeBinomialCdf, negativeBinomialPdf } from './distributions'
import { addMonths, MAX_DATE_SERIAL } from './dates'

type Executor = new (name: string) => BaseFunction
type CalcResult = ReturnType<BaseFunction['calculate']>
/** Univer executors that broadcast arrays and hand each element combination to `_handleSingleObject`. */
interface SingleObjectFunction extends BaseFunction {
  _handleSingleObject(...values: BaseValueObject[]): BaseValueObject
}
type SingleObjectExecutor = new (name: string) => SingleObjectFunction

const GROUPS = [functionCompatibility, functionDate, functionEngineering, functionFinancial, functionLookup, functionMath, functionStatistical, functionText] as unknown as Array<Array<[Executor, string]>>

/** Univer's executor class for a function name. */
function univerExecutor(name: string): Executor {
  for (const group of GROUPS) {
    for (const [executor, functionName] of group) if (functionName === name) return executor
  }
  throw new Error(`Univer has no ${name}`)
}

/**
 * A numeric argument as Univer coerces it (numbers, numeric text, booleans, empty = 0),
 * else undefined (errors, other text, ranges), which the caller leaves to Univer.
 */
function plainNumber(value: BaseValueObject | undefined): number | undefined {
  if (value?.isArray()) {
    // A single-cell reference arrives as a 1×1 array.
    const array = value as ArrayValueObject
    return array.getRowCount() === 1 && array.getColumnCount() === 1 ? plainNumber(array.get(0, 0) ?? undefined) : undefined
  }
  if (!value || value.isError()) return undefined
  if (value.isNull()) return 0
  if (value.isNumber() || value.isBoolean()) return Number(value.getValue())
  if (value.isString()) {
    const text = String(value.getValue()).trim()
    const parsed = text === '' ? Number.NaN : Number(text)
    return Number.isFinite(parsed) ? parsed : undefined
  }
  return undefined
}

/** A truthy/falsy argument, coerced the same way. */
function plainFlag(value: BaseValueObject | undefined): boolean | undefined {
  const number = plainNumber(value)
  return number === undefined ? undefined : number !== 0
}

/** A 1×1 array (a single-cell reference) as its value. */
function single(value: BaseValueObject): BaseValueObject {
  if (!value?.isArray()) return value
  const array = value as ArrayValueObject
  return array.getRowCount() === 1 && array.getColumnCount() === 1 ? (array.get(0, 0) ?? value) : value
}

/** A date result past 9999-12-31, which Excel answers with #NUM!. */
function outsideDates(result: BaseValueObject): boolean {
  return !result.isArray() && result.isNumber() && Number(result.getValue()) > MAX_DATE_SERIAL
}

const num = (value: number) => NumberValueObject.create(value)
const NUM_ERROR = () => ErrorValueObject.create(ErrorType.NUM)

/**
 * BINOM.DIST: Univer sums the probability mass term by term and builds the binomial
 * coefficient by multiplication, which loops for billions of trials (TeamGrid stopped
 * responding) and underflows for large ones. Valid numeric arguments are computed
 * with log-gamma and the regularized incomplete beta function, like Excel.
 */
function createBinomDist(): BaseFunction {
  const Base = univerExecutor('BINOM.DIST') as unknown as SingleObjectExecutor
  return new (class extends Base {
    override _handleSingleObject(numberS: BaseValueObject, trials: BaseValueObject, probability: BaseValueObject, cumulative: BaseValueObject) {
      const s = plainNumber(numberS)
      const n = plainNumber(trials)
      const p = plainNumber(probability)
      const c = plainFlag(cumulative)
      if (s === undefined || n === undefined || p === undefined || c === undefined) {
        return super._handleSingleObject(numberS, trials, probability, cumulative)
      }
      const k = Math.floor(s)
      const count = Math.floor(n)
      // Excel answers #NUM! for trial counts beyond exact doubles (1E+30 in its own tests).
      if (k < 0 || k > count || p < 0 || p > 1 || count > Number.MAX_SAFE_INTEGER) return NUM_ERROR()
      return num(c ? binomialCdf(k, count, p) : binomialPdf(k, count, p))
    }
  })('BINOM.DIST')
}

/**
 * BINOM.INV and CRITBINOM: Univer tries k = 0, 1, 2, … and sums the distribution anew
 * for each, quadratic in the trials (100,000 trials froze TeamGrid). A binary search on
 * the cumulative distribution finds the same smallest k.
 */
function createBinomInv(name: 'BINOM.INV' | 'CRITBINOM'): BaseFunction {
  const Base = univerExecutor(name) as unknown as SingleObjectExecutor
  return new (class extends Base {
    override _handleSingleObject(trials: BaseValueObject, probability: BaseValueObject, alpha: BaseValueObject) {
      const n = plainNumber(trials)
      const p = plainNumber(probability)
      const a = plainNumber(alpha)
      if (n === undefined || p === undefined || a === undefined) return super._handleSingleObject(trials, probability, alpha)
      const count = Math.floor(n)
      if (count < 0 || p <= 0 || p >= 1 || a <= 0 || a >= 1 || count > Number.MAX_SAFE_INTEGER) return NUM_ERROR()
      return num(binomialInverse(count, p, a))
    }
  })(name)
}

/** NEGBINOM.DIST: the same term-by-term sums as BINOM.DIST. */
function createNegbinomDist(): BaseFunction {
  const Base = univerExecutor('NEGBINOM.DIST') as unknown as SingleObjectExecutor
  return new (class extends Base {
    override _handleSingleObject(numberF: BaseValueObject, numberS: BaseValueObject, probability: BaseValueObject, cumulative: BaseValueObject) {
      const f = plainNumber(numberF)
      const s = plainNumber(numberS)
      const p = plainNumber(probability)
      const c = plainFlag(cumulative)
      if (f === undefined || s === undefined || p === undefined || c === undefined) {
        return super._handleSingleObject(numberF, numberS, probability, cumulative)
      }
      const failures = Math.floor(f)
      const successes = Math.floor(s)
      if (failures < 0 || successes < 1 || p <= 0 || p >= 1 || failures + successes > Number.MAX_SAFE_INTEGER) return NUM_ERROR()
      return num(c ? negativeBinomialCdf(failures, successes, p) : negativeBinomialPdf(failures, successes, p))
    }
  })('NEGBINOM.DIST')
}

/**
 * WORKDAY.INTL: Univer walks day by day, so a huge day count (Excel answers #NUM!, the
 * result is past 9999-12-31) froze TeamGrid. A count larger than every valid date is
 * answered right away.
 */
function createWorkdayIntl(): BaseFunction {
  const Base = univerExecutor('WORKDAY.INTL') as unknown as SingleObjectExecutor
  return new (class extends Base {
    override _handleSingleObject(startDate: BaseValueObject, days: BaseValueObject, weekend: BaseValueObject, holidays: BaseValueObject) {
      const count = plainNumber(days)
      if (count !== undefined && Math.abs(count) > MAX_DATE_SERIAL) return NUM_ERROR()
      // Univer reads the day count from the argument itself, which for a cell reference
      // is a 1×1 array (giving 0): hand it the single values.
      const result = super._handleSingleObject(single(startDate), single(days), weekend, holidays)
      return outsideDates(result) ? NUM_ERROR() : result
    }
  })('WORKDAY.INTL')
}

/** WORKDAY: the same day-by-day walk as WORKDAY.INTL. */
function createWorkday(): BaseFunction {
  const Base = univerExecutor('WORKDAY')
  return new (class extends Base {
    override calculate(startDate: BaseValueObject, days: BaseValueObject, holidays?: BaseValueObject): CalcResult {
      const count = plainNumber(days)
      if (count !== undefined && Math.abs(count) > MAX_DATE_SERIAL) return NUM_ERROR()
      const result = (holidays ? super.calculate(startDate, days, holidays) : super.calculate(startDate, days)) as BaseValueObject
      return outsideDates(result) ? NUM_ERROR() : result
    }
  })('WORKDAY')
}

/**
 * ACCRINT: Univer reads calc_method straight from the argument, which for a cell
 * reference is a 1×1 array (giving 0, FALSE): the accrued interest then always counted
 * from the first interest date. Single values are handed over instead.
 */
function createAccrint(): BaseFunction {
  const Base = univerExecutor('ACCRINT')
  return new (class extends Base {
    override calculate(...args: BaseValueObject[]): CalcResult {
      return super.calculate(...args.map((arg) => (arg ? single(arg) : arg)))
    }
  })('ACCRINT')
}

/** A cell value as the criteria matcher sees it. */
function criteriaCell(value: BaseValueObject | null | undefined | void): CriteriaCell {
  if (!value || value.isNull()) return { kind: 'empty' }
  if (value.isError()) return { kind: 'error', value: String(value.getValue()) }
  if (value.isBoolean()) return { kind: 'boolean', value: Boolean(value.getValue()) }
  if (value.isNumber()) return { kind: 'number', value: Number(value.getValue()) }
  return { kind: 'text', value: String(value.getValue() ?? '') }
}

interface RangeReference extends BaseValueObject {
  getRangePosition(): { startRow: number; endRow: number; startColumn: number; endColumn: number }
  toArrayValueObject(): ArrayValueObject
  isColumn(): boolean
  isRow(): boolean
}

/** Excel's grid, which whole-column and whole-row references span (Univer's sheet is smaller). */
const EXCEL_ROWS = 1_048_576
const EXCEL_COLUMNS = 16_384

/**
 * COUNTIF: Univer leaves out empty cells for `<>x` (and the rows of a whole-column range
 * past the data), coerces number-like text for `<>`, counts `<=` with an empty operand
 * and differs on error criteria. Counted here with Excel's criteria (criteria.ts);
 * criteria written as dates stay with Univer.
 */
function createCountif(): BaseFunction {
  const Base = univerExecutor('COUNTIF') as unknown as SingleObjectExecutor
  return new (class extends Base {
    override _handleSingleObject(range: BaseValueObject, criteria: BaseValueObject) {
      const matcher = criteriaMatcher(criteriaCell(single(criteria)))
      if (!matcher || !range.isReferenceObject()) return super._handleSingleObject(range, criteria)
      const reference = range as unknown as RangeReference
      const { startRow, endRow, startColumn, endColumn } = reference.getRangePosition()
      const rows = reference.isColumn() ? EXCEL_ROWS : endRow - startRow + 1
      const columns = reference.isRow() ? EXCEL_COLUMNS : endColumn - startColumn + 1
      const total = rows * columns
      const array = reference.toArrayValueObject()
      let count = 0
      let seen = 0
      array.iterator((value) => {
        seen++
        if (matcher(criteriaCell(value))) count++
      })
      // Cells of the range past the sheet's data are empty.
      if (total > seen && matcher({ kind: 'empty' })) count += total - seen
      return num(count)
    }
  })('COUNTIF')
}

/**
 * COUNTIFS: the same differences as COUNTIF, per criteria pair. Counted here when every
 * criterion is one value Excel's rules cover and the ranges have one shape; the cells of
 * whole-column (or whole-row) ranges past the sheet's data are empty.
 */
function createCountifs(): BaseFunction {
  const Base = univerExecutor('COUNTIFS')
  return new (class extends Base {
    override calculate(...variants: BaseValueObject[]): CalcResult {
      if (variants.length < 2 || variants.length % 2 !== 0) return super.calculate(...variants)
      const pairs: Array<{ array: ArrayValueObject; matcher: (cell: CriteriaCell) => boolean }> = []
      let rows = 0
      let columns = 0
      for (let index = 0; index < variants.length; index += 2) {
        const range = variants[index]!
        const given = variants[index + 1]!
        const criterion = single(given.isReferenceObject() ? (given as unknown as RangeReference).toArrayValueObject() : given)
        const matcher = criterion.isArray() ? undefined : criteriaMatcher(criteriaCell(criterion))
        if (!matcher || !range.isReferenceObject()) return super.calculate(...variants)
        const reference = range as unknown as RangeReference
        const { startRow, endRow, startColumn, endColumn } = reference.getRangePosition()
        const height = reference.isColumn() ? EXCEL_ROWS : endRow - startRow + 1
        const width = reference.isRow() ? EXCEL_COLUMNS : endColumn - startColumn + 1
        if (pairs.length && (height !== rows || width !== columns)) return ErrorValueObject.create(ErrorType.VALUE)
        rows = height
        columns = width
        pairs.push({ array: reference.toArrayValueObject(), matcher })
      }
      // Rows and columns any range has data in; past them every cell is empty.
      const dataRows = Math.min(rows, Math.max(...pairs.map((pair) => pair.array.getRowCount())))
      const dataColumns = Math.min(columns, Math.max(...pairs.map((pair) => pair.array.getColumnCount())))
      const at = (array: ArrayValueObject, row: number, column: number): CriteriaCell =>
        row < array.getRowCount() && column < array.getColumnCount() ? criteriaCell(array.get(row, column)) : { kind: 'empty' }
      let count = 0
      for (let row = 0; row < dataRows; row++) {
        for (let column = 0; column < dataColumns; column++) {
          if (pairs.every((pair) => pair.matcher(at(pair.array, row, column)))) count++
        }
      }
      const rest = rows * columns - dataRows * dataColumns
      if (rest > 0 && pairs.every((pair) => pair.matcher({ kind: 'empty' }))) count += rest
      return num(count)
    }
  })('COUNTIFS')
}

/** ISO.CEILING: Univer knows the name but has no executor (#NAME?); Excel computes it as CEILING.PRECISE. */
function createIsoCeiling(): BaseFunction {
  return new (univerExecutor('CEILING.PRECISE'))('ISO.CEILING')
}

/**
 * EDATE: Univer lets the day overflow into the next month (Jan 31 + 1 month = Mar 3)
 * where Excel stops at the month's last day, and returns dates before 1900 instead of
 * #NUM!.
 */
function createEdate(): BaseFunction {
  const Base = univerExecutor('EDATE')
  return new (class extends Base {
    override calculate(startDate: BaseValueObject, months: BaseValueObject): CalcResult {
      const result = super.calculate(startDate, months) as BaseValueObject
      const fix = (start: BaseValueObject, shift: BaseValueObject, computed: BaseValueObject): BaseValueObject => {
        const serial = plainNumber(start)
        const offset = plainNumber(shift)
        if (serial === undefined || offset === undefined || computed.isError()) return computed
        const value = addMonths(serial, Math.trunc(offset))
        return value === null ? NUM_ERROR() : NumberValueObject.create(value, (computed as NumberValueObject).getPattern?.() || undefined)
      }
      if (!result.isArray()) return fix(startDate, months, result)
      // Univer answers a 1×1 array even for two single values.
      if (!startDate.isArray() && !months.isArray()) {
        return fix(startDate, months, (result as ArrayValueObject).get(0, 0) ?? result)
      }
      return result.mapValue((computed, row, column) => {
        const start = startDate.isArray() ? (startDate as ArrayValueObject).get(row, column) : startDate
        const shift = months.isArray() ? (months as ArrayValueObject).get(row, column) : months
        return start && shift ? fix(start, shift, computed) : computed
      })
    }
  })('EDATE')
}

/**
 * OFFSET: Univer adds fractional offsets and sizes to the row and column numbers as they
 * are, and the reference with fractional rows (OFFSET(S8:X19, 3.5, 7.2, 14.9, 2.3)) makes
 * building its array throw, which aborts the whole calculation run: every later result
 * of the workbook stays empty and the engine never reports the run as finished. Excel
 * truncates offsets and sizes toward zero, and a size between 0 and 1 counts as 1.
 */
function createOffset(): BaseFunction {
  const Base = univerExecutor('OFFSET') as unknown as SingleObjectExecutor
  const whole = (value: BaseValueObject, size: boolean): BaseValueObject => {
    const number = plainNumber(value)
    if (number === undefined || !Number.isFinite(number) || Number.isInteger(number)) return value
    const truncated = Math.trunc(number)
    return NumberValueObject.create(size && truncated === 0 ? Math.sign(number) : truncated)
  }
  return new (class extends Base {
    override _handleSingleObject(reference: BaseValueObject, rows: BaseValueObject, columns: BaseValueObject, height: BaseValueObject, width: BaseValueObject) {
      return super._handleSingleObject(reference, whole(rows, false), whole(columns, false), whole(height, true), whole(width, true))
    }
  })('OFFSET')
}

/**
 * COMPLEX: an empty suffix (a blank cell or "") means "i" in Excel, as an omitted one
 * does; Univer answered #VALUE!, and every IM* function over the result failed with it.
 */
function createComplex(): BaseFunction {
  const Base = univerExecutor('COMPLEX')
  const blank = (value: BaseValueObject | undefined): boolean => {
    if (!value) return false
    if (value.isArray()) {
      const array = value as ArrayValueObject
      return array.getRowCount() === 1 && array.getColumnCount() === 1 && blank(array.get(0, 0) ?? undefined)
    }
    return value.isNull() || (value.isString() && value.getValue() === '')
  }
  return new (class extends Base {
    override calculate(realNum: BaseValueObject, iNum: BaseValueObject, suffix?: BaseValueObject): CalcResult {
      return super.calculate(realNum, iNum, blank(suffix) ? StringValueObject.create('i') : (suffix as BaseValueObject))
    }
  })('COMPLEX')
}

/**
 * UNICODE of empty text: Excel answers #VALUE! (as CODE does in Univer too), Univer's
 * character code of nothing (NaN) showed as an empty cell.
 */
function createUnicode(): BaseFunction {
  const Base = univerExecutor('UNICODE') as unknown as SingleObjectExecutor
  return new (class extends Base {
    override _handleSingleObject(text: BaseValueObject) {
      return text.isString() && text.getValue() === '' ? ErrorValueObject.create(ErrorType.VALUE) : super._handleSingleObject(text)
    }
  })('UNICODE')
}

export function createExcelFunctionExecutors(): BaseFunction[] {
  return [createBinomDist(), createBinomInv('BINOM.INV'), createBinomInv('CRITBINOM'), createNegbinomDist(), createWorkdayIntl(), createWorkday(), createAccrint(), createCountif(), createCountifs(), createIsoCeiling(), createEdate(), createOffset(), createComplex(), createUnicode()]
}

/**
 * An exception in an executor aborts the engine's whole run (every later result stays
 * empty): one of ours answers #VALUE! instead and reports it.
 */
function guarded(executor: BaseFunction): BaseFunction {
  const calculate = executor.calculate.bind(executor)
  executor.calculate = (...variants: BaseValueObject[]) => {
    try {
      return calculate(...variants)
    } catch (error) {
      console.error(`[TeamGrid] ${executor.name} failed`, error)
      return ErrorValueObject.create(ErrorType.VALUE)
    }
  }
  return executor
}

/** Registers the corrected executors on every Univer the editor creates (patches/sheets-univer-hook.patch). */
export function installExcelFunctions(): void {
  const hook = window as unknown as { teamGridOnUniver?: (univer: Univer) => void }
  const previous = hook.teamGridOnUniver
  hook.teamGridOnUniver = (univer) => {
    previous?.(univer)
    const functionService = univer.__getInjector().get(IFunctionService)
    const ours = new Map(createExcelFunctionExecutors().map((executor) => [executor.name, guarded(executor)]))
    // The formula plugin registers Univer's own executors again while it starts, after
    // this hook ran: whoever registers one of these names gets ours.
    const register = functionService.registerExecutors.bind(functionService)
    functionService.registerExecutors = (...executors: BaseFunction[]) =>
      register(...executors.map((executor) => ours.get(executor.name) ?? executor))
    functionService.registerExecutors(...ours.values())
  }
}
