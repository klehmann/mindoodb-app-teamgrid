import { BooleanValueObject, NumberValueObject, type BaseValueObject } from '@univerjs/engine-formula'
import { describe, expect, it } from 'vitest'

import { addMonths, dateToSerial, serialToDate } from './dates'
import { binomialCdf, negativeBinomialCdf, regularizedBeta } from './distributions'
import { createExcelFunctionExecutors } from './excel-functions'

const executor = (name: string) => {
  const found = createExcelFunctionExecutors().find((e) => e.name === name)
  if (!found) throw new Error(name)
  return found
}
const n = (value: number) => NumberValueObject.create(value)
const b = (value: boolean) => BooleanValueObject.create(value)
const run = (name: string, ...args: BaseValueObject[]) => (executor(name).calculate(...args) as BaseValueObject).getValue()

describe('Excel date serials', () => {
  it('keep 1900-02-29 like Excel', () => {
    expect(serialToDate(1)).toEqual({ year: 1900, month: 1, day: 1 })
    expect(serialToDate(60)).toEqual({ year: 1900, month: 2, day: 29 })
    expect(serialToDate(61)).toEqual({ year: 1900, month: 3, day: 1 })
    expect(dateToSerial(2025, 1, 10)).toBe(45667)
    expect(dateToSerial(1900, 3, 1)).toBe(61)
  })

  it('EDATE stops at the last day of the month', () => {
    expect(addMonths(dateToSerial(2023, 1, 31), 1)).toBe(dateToSerial(2023, 2, 28))
    expect(addMonths(dateToSerial(2024, 1, 31), 1)).toBe(dateToSerial(2024, 2, 29))
    expect(addMonths(dateToSerial(2024, 3, 31), -1)).toBe(dateToSerial(2024, 2, 29))
    expect(addMonths(dateToSerial(1900, 1, 31), 1)).toBe(60)
    expect(addMonths(10, -1)).toBeNull()
    expect(run('EDATE', n(dateToSerial(2023, 1, 31)), n(1))).toBe(dateToSerial(2023, 2, 28))
    expect(run('EDATE', n(100), n(-12))).toBe('#NUM!')
  })
})

describe('binomial distributions', () => {
  it('match Excel for small and huge arguments without looping', () => {
    expect(binomialCdf(5, 10, 0.5)).toBeCloseTo(0.623046875, 14)
    expect(binomialCdf(1000, 10000, 0.1)).toBeCloseTo(0.50842103926526505, 10)
    expect(binomialCdf(9999, 10000, 0.999)).toBeCloseTo(0.99995482665402291, 12)
    expect(negativeBinomialCdf(1e6, 10, 0.5)).toBeCloseTo(1, 12)
    expect(regularizedBeta(0.5, 2, 2)).toBeCloseTo(0.5, 14)
    const started = Date.now()
    expect(run('BINOM.DIST', n(2), n(1e10), n(0.3), b(true))).toBe(0)
    expect(run('NEGBINOM.DIST', n(2), n(1e10), n(0.3), b(true))).toBe(0)
    expect(Date.now() - started).toBeLessThan(100)
  })

  it('BINOM.INV finds the smallest k fast', () => {
    expect(run('BINOM.INV', n(10), n(0.7), n(0.8))).toBe(8)
    expect(run('BINOM.INV', n(10), n(0.5), n(0.5))).toBe(5)
    expect(run('BINOM.INV', n(100000), n(0.7), n(0.8))).toBe(70122)
    expect(run('BINOM.INV', n(100000), n(0.1), n(0.5))).toBe(10000)
    expect(run('CRITBINOM', n(10), n(0.5), n(0.9))).toBe(7)
    expect(run('BINOM.INV', n(10), n(0.7), n(0))).toBe('#NUM!')
  })

  it('answer #NUM! like Excel', () => {
    expect(run('BINOM.DIST', n(2), n(10), n(-2), b(true))).toBe('#NUM!')
    expect(run('BINOM.DIST', n(10), n(2), n(0.3), b(true))).toBe('#NUM!')
    expect(run('NEGBINOM.DIST', n(0), n(0), n(0.5), b(true))).toBe('#NUM!')
  })
})

describe('other corrections', () => {
  it('WORKDAY.INTL answers a day count past every date at once', () => {
    expect(run('WORKDAY.INTL', n(45667), n(-1e15))).toBe('#NUM!')
    expect(run('WORKDAY', n(45667), n(1e15))).toBe('#NUM!')
    expect(run('WORKDAY', n(45667), n(10))).toBe(45681)
  })

  it('ACCRINT takes calc_method from a cell (a 1×1 array)', async () => {
    const { ArrayValueObject } = await import('@univerjs/engine-formula')
    const cell = (value: boolean) => ArrayValueObject.create({ calculateValueList: [[b(value)]], rowCount: 1, columnCount: 1, unitId: '', sheetId: '', row: 0, column: 0 })
    const args = [n(43435), n(44531), n(43983), n(0.1), n(1000), n(1), n(0)]
    expect(run('ACCRINT', ...args, cell(true))).toBeCloseTo(150, 9)
    expect(run('ACCRINT', ...args, cell(false))).toBeCloseTo(-50, 9)
  })

  it('ISO.CEILING exists and computes like CEILING.PRECISE', () => {
    expect(run('ISO.CEILING', n(4.3), n(2))).toBe(6)
    expect(run('ISO.CEILING', n(-4.3), n(2))).toBe(-4)
  })
})

describe('COUNTIF criteria', async () => {
  const { criteriaMatcher, readNumber } = await import('./criteria')
  const cells = [
    { kind: 'number', value: 23 },
    { kind: 'text', value: '23' },
    { kind: 'boolean', value: true },
    { kind: 'number', value: 1 },
    { kind: 'text', value: 'Mr. Brown' },
    { kind: 'error', value: '#N/A' },
    { kind: 'text', value: '' },
    { kind: 'empty' },
  ] as const
  const count = (criterion: Parameters<typeof criteriaMatcher>[0]) => cells.filter(criteriaMatcher(criterion)!).length
  const text = (value: string) => ({ kind: 'text', value }) as const

  it('follows Excel on IronCalc countif.xlsx cases', () => {
    expect(count({ kind: 'number', value: 23 })).toBe(2)
    expect(count(text('=23'))).toBe(2)
    expect(count(text('<>23'))).toBe(7)
    expect(count(text('<23'))).toBe(1)
    expect(count(text('<>*brown'))).toBe(7)
    expect(count(text('=*BROWN'))).toBe(1)
    expect(count(text('#N/A'))).toBe(1)
    expect(count(text(''))).toBe(2)
    expect(count(text('<>'))).toBe(7)
    expect(count(text('<='))).toBe(0)
    expect(count({ kind: 'empty' })).toBe(0)
    expect(count(text('TRUE'))).toBe(1)
    expect(readNumber('50%')).toBe(0.5)
    expect(criteriaMatcher(text('>1/1/2020'))).toBeUndefined()
  })
})
