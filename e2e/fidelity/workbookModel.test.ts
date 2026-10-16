import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'

import { describe, expect, it } from 'vitest'

import { corpusDir } from './paths'
import { compareModels, readWorkbookModel } from './workbookModel'

describe('the round-trip model', () => {
  it('reads every corpus workbook and finds no difference to itself', async () => {
    const files = readdirSync(corpusDir).filter((file) => file.startsWith('rxw-') && file.endsWith('.xlsx'))
    expect(files.length).toBeGreaterThan(100)
    for (const file of files) {
      const bytes = new Uint8Array(readFileSync(join(corpusDir, file)))
      const model = await readWorkbookModel(bytes)
      expect(model.sheets.length, file).toBeGreaterThan(0)
      expect(compareModels(model, await readWorkbookModel(bytes)), file).toEqual({})
    }
  })

  it('expands shared formulas and resolves formats', async () => {
    const model = await readWorkbookModel(new Uint8Array(readFileSync(join(corpusDir, 'ironcalc-accrint-cases.xlsx'))))
    expect(model.sheets[0]!.cells.get('I212')).toBe('f:ACCRINT(A212,B212,C212,D212,E212,F212,G212,H212)')
    const formats = await readWorkbookModel(new Uint8Array(readFileSync(join(corpusDir, 'rxw-format01.xlsx'))))
    expect([...formats.sheets.flatMap((s) => [...s.formats.values()])].length).toBeGreaterThan(0)
  })
})
