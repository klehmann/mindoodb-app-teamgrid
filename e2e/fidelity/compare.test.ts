import { describe, expect, it } from 'vitest'

import { mismatchKind, sameResult } from './compare'
import { functionsOf } from './workbookCells'

describe('comparing formula results with Excel', () => {
  it('allows the last bits of a long computation to differ', () => {
    expect(sameResult(0.1 + 0.2, 0.3)).toBe(true)
    expect(sameResult(1.0000001, 1)).toBe(false)
    expect(sameResult(1e-17, 0)).toBe(true)
    expect(sameResult('42', 42)).toBe(true)
    expect(sameResult(null, 0)).toBe(false)
  })

  it('compares errors, text and booleans exactly', () => {
    expect(sameResult('#N/A', '#N/A')).toBe(true)
    expect(sameResult('#N/A', '#DIV/0!')).toBe(false)
    expect(sameResult('TRUE', true)).toBe(true)
    expect(sameResult(false, false)).toBe(true)
    expect(sameResult(null, '')).toBe(true)
    expect(sameResult('abc', 'ABC')).toBe(false)
  })

  it('tells missing functions from wrong results', () => {
    expect(mismatchKind('#NAME?', 3, false)).toBe('missing')
    expect(mismatchKind('#SPILL!', 1, false)).toBe('error')
    expect(mismatchKind('#N/A', '#DIV/0!', true)).toBe('value')
    expect(mismatchKind(null, 1, false)).toBe('empty')
    expect(mismatchKind(2, 1, false)).toBe('value')
  })

  it('names the functions a formula calls', () => {
    expect(functionsOf('_xlfn.XLOOKUP(A1,B:B,C:C,"IF(x)")+SUM(1)')).toEqual(['XLOOKUP', 'SUM'])
    expect(functionsOf('IFERROR(_xlfn._xlws.FILTER(A1:A3,B1:B3>0),"")')).toEqual(['IFERROR', 'FILTER'])
    expect(functionsOf('A1+B1')).toEqual([])
  })
})
