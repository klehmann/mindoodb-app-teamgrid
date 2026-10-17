import { describe, expect, it } from 'vitest'

import { convertUnits } from './convert'

describe('CONVERT units as in Excel', () => {
  it('converts with prefixes where Excel allows them', () => {
    expect(convertUnits(1, 'km', 'm')).toBe(1000)
    expect(convertUnits(1, 'km2', 'm2')).toBe(1e6)
    expect(convertUnits(1, 'kibyte', 'byte')).toBe(1024)
    expect(convertUnits(1, 'pc', 'm')).toBe(3.08567758128155e16) // parsec, not pico-calories
    expect(convertUnits(1, 'kmi', 'mi')).toBeUndefined()
    expect(convertUnits(1, 'kC', 'C')).toBeUndefined()
  })

  it('converts temperatures by formula and rejects mixed kinds', () => {
    expect(convertUnits(100, 'C', 'F')).toBeCloseTo(212, 10)
    expect(convertUnits(0, 'K', 'C')).toBeCloseTo(-273.15, 10)
    expect(convertUnits(1, 'm', 'g')).toBeUndefined()
  })
})
