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
  functionMath,
  functionStatistical,
  IFunctionService,
  NumberValueObject,
} from '@univerjs/engine-formula'

import { binomialCdf, binomialInverse, binomialPdf, negativeBinomialCdf, negativeBinomialPdf } from './distributions'
import { addMonths, MAX_DATE_SERIAL } from './dates'

type Executor = new (name: string) => BaseFunction
type CalcResult = ReturnType<BaseFunction['calculate']>
/** Univer executors that broadcast arrays and hand each element combination to `_handleSingleObject`. */
interface SingleObjectFunction extends BaseFunction {
  _handleSingleObject(...values: BaseValueObject[]): BaseValueObject
}
type SingleObjectExecutor = new (name: string) => SingleObjectFunction

const GROUPS = [functionCompatibility, functionDate, functionMath, functionStatistical] as unknown as Array<Array<[Executor, string]>>

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

export function createExcelFunctionExecutors(): BaseFunction[] {
  return [createBinomDist(), createBinomInv('BINOM.INV'), createBinomInv('CRITBINOM'), createNegbinomDist(), createWorkdayIntl(), createWorkday(), createIsoCeiling(), createEdate()]
}

/** Registers the corrected executors on every Univer the editor creates (patches/sheets-univer-hook.patch). */
export function installExcelFunctions(): void {
  const hook = window as unknown as { teamGridOnUniver?: (univer: Univer) => void }
  const previous = hook.teamGridOnUniver
  hook.teamGridOnUniver = (univer) => {
    previous?.(univer)
    const functionService = univer.__getInjector().get(IFunctionService)
    const ours = new Map(createExcelFunctionExecutors().map((executor) => [executor.name, executor]))
    // The formula plugin registers Univer's own executors again while it starts, after
    // this hook ran: whoever registers one of these names gets ours.
    const register = functionService.registerExecutors.bind(functionService)
    functionService.registerExecutors = (...executors: BaseFunction[]) =>
      register(...executors.map((executor) => ours.get(executor.name) ?? executor))
    functionService.registerExecutors(...ours.values())
  }
}
