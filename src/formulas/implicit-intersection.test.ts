import { ArrayValueObject, NumberValueObject, type BaseValueObject } from '@univerjs/engine-formula'
import { describe, expect, it } from 'vitest'

import { formulaAreas, implicitIntersection } from './implicit-intersection'

const column = (values: number[]) =>
  ArrayValueObject.create({
    calculateValueList: values.map((v) => [NumberValueObject.create(v)]),
    rowCount: values.length,
    columnCount: 1,
    unitId: '',
    sheetId: '',
    row: 0,
    column: 0,
  }) as unknown as BaseValueObject & { isReferenceObject(): boolean }
const row = (values: number[]) =>
  ArrayValueObject.create({
    calculateValueList: [values.map((v) => NumberValueObject.create(v))],
    rowCount: 1,
    columnCount: values.length,
    unitId: '',
    sheetId: '',
    row: 0,
    column: 0,
  }) as unknown as BaseValueObject & { isReferenceObject(): boolean }

describe('implicit intersection', () => {
  it('reads the areas a formula refers to', () => {
    expect(formulaAreas('SUM($A$2:A4)+"B1:B9"')).toEqual([{ startRow: 1, endRow: 3, startColumn: 0, endColumn: 0 }])
    expect(formulaAreas('BINOM.DIST(G:G,B73,C73,D73)')).toEqual([{ startRow: 0, endRow: 1_048_575, startColumn: 6, endColumn: 6 }])
    expect(formulaAreas('ROW(3:5)')).toEqual([{ startRow: 2, endRow: 4, startColumn: 0, endColumn: 16_383 }])
    expect(formulaAreas('LOG10(A1)')).toEqual([])
  })

  it('takes the element in the formula’s row or column, or #VALUE!', () => {
    // =ARABIC(E39:E42) in row 41 (index 40) reads E41
    expect(implicitIntersection(column([1, 2, 3, 4]), 'ARABIC(E39:E42)', 40, 1).getValue()).toBe(3)
    // in row 23, outside A25:A27
    expect(implicitIntersection(column([1, 2, 3]), 'CHOOSE(A25:A27,1,2,3)', 22, 1).getValue()).toBe('#VALUE!')
    // a whole column, cut to the sheet by Univer
    expect(implicitIntersection(column([5, 6, 7, 8]), 'BINOM.DIST(G:G,1,0.5,TRUE)', 2, 4).getValue()).toBe(7)
    // a row range in column C
    expect(implicitIntersection(row([1, 2, 3]), 'A1:C1*2', 5, 2).getValue()).toBe(3)
    // an array without a position: its first element
    expect(implicitIntersection(row([2, 3]), '{1,2}+1', 0, 0).getValue()).toBe(2)
  })
})
