import { describe, expect, it } from 'vitest'

import { compareValues, xlookupIndex } from './xlookup'

const n = (value: number) => ({ kind: 'number' as const, value })
const t = (value: string) => ({ kind: 'text' as const, value })

describe('XLOOKUP search as in Excel', () => {
  const numbers = [10, 30, 20, 50, 40].map(n)

  it('finds the smallest larger and the largest smaller item, not the first one', () => {
    expect(xlookupIndex(n(25), numbers, 1, 1)).toBe(1) // 30
    expect(xlookupIndex(n(25), numbers, -1, 1)).toBe(2) // 20
    expect(xlookupIndex(n(60), numbers, 1, 1)).toBeUndefined()
  })

  it('searches sorted ranges by bisection in both directions', () => {
    const ascending = [1, 3, 5, 7].map(n)
    const descending = [7, 5, 3, 1].map(n)
    expect(xlookupIndex(n(4), ascending, -1, 2)).toBe(1)
    expect(xlookupIndex(n(4), ascending, 1, 2)).toBe(2)
    expect(xlookupIndex(n(4), descending, -1, -2)).toBe(2)
    expect(xlookupIndex(n(4), descending, 1, -2)).toBe(1)
    expect(xlookupIndex(n(4), ascending, 0, 2)).toBeUndefined()
  })

  it('matches wildcards and compares text without case', () => {
    const words = ['Apple', 'banana', 'Cherry'].map(t)
    expect(xlookupIndex(t('b*'), words, 2, 1)).toBe(1)
    expect(xlookupIndex(t('c?erry'), words, 2, -1)).toBe(2)
    expect(compareValues(t('apple'), t('APPLE'))).toBe(0)
    expect(compareValues(n(1), t('a'))).toBe(-1)
    expect(compareValues({ kind: 'empty' }, n(0))).toBe(0)
  })
})
