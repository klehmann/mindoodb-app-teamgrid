// Notes, hyperlinks, the autoFilter, print settings and tab colors through the
// real pipeline: GenOffice's gateway writes the xlsx, the WASM sidecar (Node
// build) reads it back. Skipped without native/xlsx-wasm/pkg-node.
import { describe, expect, it } from 'vitest'

import { blankXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'

import type { WorkbookFile, WorkbookSaveRequest } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { FConditionalFormattingBuilder } from '@univerjs/sheets-conditional-formatting/facade'

import { toUniverDvRule } from '../../vendor/genoffice/apps/sheets/src/renderer/univer-sync'
import { createNodeEngine, nodeEngineAvailable } from '../testing/node-engine'
import { saveWorkbookInBrowser } from '../xlsx/browser-save'
import type { XlsxEngine } from '../xlsx/engine'
import { emptySaveRequest, workbookToXlsx } from './export'
import { loadRuleConverters } from './rules'
import { areaToIds, areaToIndices, parseHeaderFooter } from './sheet-features'
import { readWholeWorkbook, type SidecarWorkbook } from './sidecar-read'
import { emptyStoredWorkbook, syncWorkbook } from './sync'

const SHEET = 'sheet-1'
const area = { startRow: 0, startColumn: 0, endRow: 3, endColumn: 2 }

async function sourceXlsx(): Promise<Uint8Array> {
  const request = emptySaveRequest('source')
  const rows = [['Name', 'Team', 'Wert'], ['a', 'x', 1], ['b', 'y', 2], ['c', 'x', 3]]
  rows.forEach((row, rowIndex) =>
    row.forEach((value, column) =>
      request.edits.push({ sheetId: SHEET, row: rowIndex, column, writeValue: true, value } as never),
    ),
  )
  request.noteStates.push({ sheetId: SHEET, notes: [{ row: 1, column: 0, author: 'Karsten', text: 'Prüfen' }] })
  request.hyperlinkEdits.push({ sheetId: SHEET, row: 2, column: 0, target: 'https://mindoo.de/' })
  request.pageSetupStates.push({
    sheetId: SHEET,
    orientation: 'landscape',
    printArea: 'A1:C4',
    printTitles: '1:1',
    header: { center: 'Bericht &D' },
    rowBreaks: [2],
  } as WorkbookSaveRequest['pageSetupStates'][number])
  request.filterStates.push({
    sheetId: SHEET,
    filter: { range: area, columns: [{ colId: 1, values: ['x'] }] },
    hiddenRows: [2],
    visibilityRange: area,
  } as WorkbookSaveRequest['filterStates'][number])
  const highlight = new FConditionalFormattingBuilder()
    .whenNumberGreaterThan(1)
    .setBackground('#FFFF00')
    .setBold(true)
    .setRanges([{ startRow: 1, startColumn: 2, endRow: 3, endColumn: 2 }])
    .build()
  const scale = new FConditionalFormattingBuilder()
    .setColorScale([
      { index: 0, color: '#F8696B', value: { type: 'min' } },
      { index: 1, color: '#63BE7B', value: { type: 'max' } },
    ] as never)
    .setRanges([{ startRow: 1, startColumn: 2, endRow: 3, endColumn: 2 }])
    .build()
  request.cfStates.push({
    sheetId: SHEET,
    rules: [highlight, scale].map((rule) => ({ ranges: rule.ranges, stopIfTrue: false, rule: rule.rule as never })),
  })
  const list = toUniverDvRule(
    {
      ranges: [{ startRow: 1, startColumn: 1, endRow: 3, endColumn: 1 }],
      ruleType: 'list',
      formulas: ['"x,y,z"'],
      allowBlank: true,
      suppressDropdown: false,
      showInputMessage: false,
      showErrorMessage: true,
    } as never,
    'dv1',
  )!
  const { ranges: dvRanges, uid: _uid, ...dvRule } = list as { ranges: never; uid: string }
  request.dvStates.push({ sheetId: SHEET, rules: [{ ranges: dvRanges, rule: dvRule }] })
  const blank = new Uint8Array(await blankXlsxBuffer('Sheet1'))
  return (await saveWorkbookInBrowser(blank, new Map([[SHEET, 'Sheet1']]), request)).bytes
}

async function read(engine: XlsxEngine, bytes: Uint8Array): Promise<SidecarWorkbook> {
  const file = (await engine.open(bytes, 'de')) as WorkbookFile
  const workbook = await readWholeWorkbook(engine, file)
  await engine.close(file.sessionId)
  return workbook
}

describe('sheet features', () => {
  it('splits headers and footers into their sections', () => {
    expect(parseHeaderFooter('&LLinks&C&P von &N&RRechts && mehr')).toEqual({
      left: 'Links',
      center: '&P von &N',
      right: 'Rechts && mehr',
    })
    expect(parseHeaderFooter('Nur Mitte')).toEqual({ center: 'Nur Mitte' })
  })

  it('keeps whole-column areas reaching to the sheet end', () => {
    const area = areaToIds({ startRow: 0, startColumn: 1, endRow: 1_048_575, endColumn: 1 }, ['r1', 'r2'], ['c1', 'c2'])
    expect(area).toEqual({ startRowId: 'r1', startColumnId: 'c2', endRowId: 'r2', endColumnId: 'c2', rowsToEnd: true })
    expect(areaToIndices(area!, new Map([['r1', 0], ['r2', 1]]), new Map([['c1', 0], ['c2', 1]]))).toEqual({
      startRow: 0,
      startColumn: 1,
      endRow: 1_048_575,
      endColumn: 1,
    })
  })

  it.skipIf(!nodeEngineAvailable)('stores them by id and writes them back', async () => {
    const engine = createNodeEngine()
    await loadRuleConverters()
    const source = await read(engine, await sourceXlsx())
    const { next } = syncWorkbook({ stored: emptyStoredWorkbook(), storedIdByFileId: new Map(), read: source })
    const sheet = Object.values(next.worksheetsById)[0]!
    const [r1, r2, r3, r4] = sheet.rowOrder
    const [cA, , cC] = sheet.columnOrder

    expect(sheet.cellsById[`${r2}:${cA}`]?.note).toEqual({ author: 'Karsten', text: 'Prüfen' })
    expect(sheet.cellsById[`${r3}:${cA}`]?.link).toBe('https://mindoo.de/')
    expect(sheet.autoFilter).toEqual({
      area: { startRowId: r1, startColumnId: cA, endRowId: r4, endColumnId: cC },
      columns: [{ colId: 1, values: ['x'] }],
    })
    expect(sheet.pageSetup).toMatchObject({
      orientation: 'landscape',
      header: { center: 'Bericht &D' },
      printArea: { startRowId: r1, startColumnId: cA, endRowId: r4, endColumnId: cC },
      printTitleRows: { startRowId: r1, endRowId: r1 },
      rowBreaks: [r3],
    })

    sheet.tabColor = '#FF8800'
    const { bytes } = await workbookToXlsx(next, engine)
    const backWorkbook = await read(engine, bytes)
    const back = backWorkbook.sheets[0]!
    expect(back.meta.tabColor?.toUpperCase()).toContain('FF8800')
    expect(back.meta.comments).toEqual([{ row: 1, column: 0, author: 'Karsten', text: 'Prüfen' }])
    expect(back.hyperlinks).toEqual([{ row: 2, column: 0, target: 'https://mindoo.de/' }])
    expect(back.autoFilter).toEqual(area)
    expect(back.autoFilterColumns).toEqual([{ colId: 1, values: ['x'] }])
    expect(back.pageSetup).toMatchObject({ orientation: 'landscape', oddHeader: '&CBericht &D' })
    expect(back.meta.printArea).toMatch(/\$A\$1:\$C\$4$/)
    expect(back.meta.printTitles).toMatch(/\$1:\$1$/)
    expect(back.rowBreaks).toEqual([2])

    // Conditional formats and validation: stored by id, written back the same.
    expect(Object.keys(sheet.conditionalFormatsById ?? {})).toHaveLength(2)
    expect(Object.keys(sheet.dataValidationsById ?? {})).toHaveLength(1)
    const withoutDxf = (rules: SidecarWorkbook['sheets'][number]['conditionalRules']) =>
      (rules ?? []).map(({ dxfIndex: _dxf, ...rule }) => rule)
    expect(withoutDxf(back.conditionalRules)).toEqual(withoutDxf(source.sheets[0]!.conditionalRules))
    expect(back.dataValidations).toEqual(source.sheets[0]!.dataValidations)
    const highlight = back.conditionalRules?.find((rule) => rule.dxfIndex !== undefined)
    expect(backWorkbook.file.dxfStyles[highlight!.dxfIndex!]).toMatchObject({ bold: true, fillColor: expect.stringMatching(/FFFF00$/i) })
  })
})
