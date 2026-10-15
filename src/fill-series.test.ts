// Localized fill series: the lists, choosing among overlapping ones, and filling.
import { describe, expect, it } from 'vitest'

import { chooseSeries, fillSeries, seriesLists } from './fill-series'

const cells = (...values: string[]) => values.map((v) => ({ v, s: 'style-1' }))
const values = (filled: Array<{ v?: unknown }>) => filled.map((cell) => cell.v)

describe('localized fill series', () => {
  it('knows month and weekday names of all TeamGrid languages', () => {
    const keys = seriesLists().map((list) => list.key)
    expect(keys).toEqual(expect.arrayContaining(['de-month-long', 'de-weekday-short', 'fr-month-long', 'pl-weekday-long']))
    expect(chooseSeries(['Januar'], 'de')?.key).toBe('de-month-long')
    expect(chooseSeries(['Mrz'], 'de')?.names[2]).toBe('Mrz')
    expect(chooseSeries(['Teamsitzung'], 'de')).toBeNull()
  })

  it('continues German months and weekdays', () => {
    expect(values(fillSeries(cells('Januar'), 3, false, 'de'))).toEqual(['Februar', 'März', 'April'])
    expect(values(fillSeries(cells('Montag', 'Mittwoch'), 3, false, 'de'))).toEqual(['Freitag', 'Sonntag', 'Dienstag'])
    expect(values(fillSeries(cells('Mo'), 2, false, 'de'))).toEqual(['Di', 'Mi'])
    expect(values(fillSeries(cells('Nov', 'Dez'), 2, false, 'de'))).toEqual(['Jan', 'Feb'])
  })

  it('picks the list holding all cells: January…May goes on with June, not Jun', () => {
    expect(values(fillSeries(cells('January', 'February', 'March', 'April', 'May'), 2, false, 'de'))).toEqual(['June', 'July'])
    expect(values(fillSeries(cells('Jan', 'Feb'), 1, false, 'en'))).toEqual(['Mar'])
  })

  it('fills upwards, keeps the writing and the format, and copies when steps differ', () => {
    expect(values(fillSeries(cells('März', 'April'), 2, true, 'de'))).toEqual(['Januar', 'Februar'])
    expect(values(fillSeries(cells('JANUAR'), 1, false, 'de'))).toEqual(['FEBRUAR'])
    expect(values(fillSeries(cells('januar'), 1, false, 'de'))).toEqual(['februar'])
    expect(fillSeries(cells('Januar'), 1, false, 'de')[0]).toEqual({ v: 'Februar', s: 'style-1' })
    expect(values(fillSeries(cells('Januar', 'März', 'April'), 3, false, 'de'))).toEqual(['Januar', 'März', 'April'])
  })

  it('prefers the UI language where names overlap (Mai is German and French)', () => {
    expect(chooseSeries(['mai'], 'fr')?.language).toBe('fr')
    expect(chooseSeries(['Mai'], 'de')?.language).toBe('de')
  })
})
