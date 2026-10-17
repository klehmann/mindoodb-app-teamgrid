import { describe, expect, it } from 'vitest'

import { formatComplex, parseComplex } from './complex'

describe('complex numbers as Excel writes them', () => {
  it('parses every form, exponents included', () => {
    expect(parseComplex('1+2i')).toEqual({ x: 1, y: 2, suffix: 'i' })
    expect(parseComplex('2i')).toEqual({ x: 0, y: 2, suffix: 'i' })
    expect(parseComplex('-7.5')).toEqual({ x: -7.5, y: 0, suffix: 'i' })
    expect(parseComplex('7-i')).toEqual({ x: 7, y: -1, suffix: 'i' })
    expect(parseComplex('-j')).toEqual({ x: 0, y: -1, suffix: 'j' })
    expect(parseComplex('1.2E-45-3i')).toEqual({ x: 1.2e-45, y: -3, suffix: 'i' })
    expect(parseComplex('1+2k')).toBeNull()
    expect(parseComplex('1+-2i')).toBeNull()
  })

  it('writes the results Excel writes', () => {
    const cases: Array<[number, number, 'i' | 'j', string]> = [
      [12, 6, 'i', '12+6i'],
      [1, -1, 'i', '1-i'],
      [0, -5, 'i', '-5i'],
      [0, 1, 'j', 'j'],
      [1.2e-45, -3, 'i', '1.2E-45-3i'],
      [1e-16, 0, 'i', '0.0000000000000001'],
      [1e-18, 0, 'i', '0.000000000000000001'],
      [1e-20, 0, 'i', '1E-20'],
      [0, 0, 'i', '0'],
      [-23.1234, -45.899, 'i', '-23.1234-45.899i'],
      [0.1 + 0.2, 0, 'i', '0.3'],
    ]
    for (const [x, y, suffix, text] of cases) expect(formatComplex({ x, y, suffix })).toBe(text)
  })
})
