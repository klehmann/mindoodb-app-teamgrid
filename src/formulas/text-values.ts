/**
 * Text in formulas as Excel reads it.
 *
 * Empty text from a formula (`=""`, `=IF(…,"",…)`) as Excel reads it: a text value, not
 * a blank cell. Univer's references skip every cell whose value is "" (core's isNullCell),
 * so a formula over such a cell saw a blank: CEILING("") gave 0 where Excel answers
 * #VALUE!, COUNTA and AVERAGEA left the cell out.
 *
 * Only results of the current calculation run are looked at (the engine's runtime data,
 * which holds formula results only); stored cells with "" and no formula stay blank, like
 * the empty cells an import writes for formatting.
 */
import { BaseReferenceObject, type BaseValueObject, ErrorType, ErrorValueObject, StringValueObject } from '@univerjs/engine-formula'

type Iterator = BaseReferenceObject['iterator']
let installed = false

export function installTextValues(): void {
  if (installed) return
  installed = true
  const prototype = BaseReferenceObject.prototype
  const iterator: Iterator = prototype.iterator
  prototype.iterator = function (this: BaseReferenceObject, callback) {
    const runtime = this.getCurrentRuntimeSheetData()
    if (!runtime) return iterator.call(this, callback)
    return iterator.call(this, (value, row, column) => {
      if (value == null && runtime.getValue(row, column)?.v === '') {
        return callback(StringValueObject.create('') as BaseValueObject, row, column)
      }
      return callback(value, row, column)
    })
  }

  // Text where a number is expected: Excel trims spaces around a number or an amount
  // (" 10 " → 10, " €100 " → 100) but not around a date or time (" 2024-01-10 " →
  // #VALUE!; Univer read it as the date).
  const strings = StringValueObject.prototype
  const convert = strings.convertToNumberObjectValue
  strings.convertToNumberObjectValue = function (this: StringValueObject) {
    const result = convert.call(this)
    const text = String(this.getValue())
    if (text === text.trim() || result.isError()) return result
    const pattern = (result as BaseValueObject & { getPattern?(): string }).getPattern?.() ?? ''
    return DATE_TIME.test(pattern.replace(/"[^"]*"|\[[^\]]*\]|\\./g, '')) ? ErrorValueObject.create(ErrorType.VALUE) : result
  }
}

/** A number format that shows a date or a time (quoted text, [colors] and escapes removed). */
const DATE_TIME = /[ymdhs]/i
