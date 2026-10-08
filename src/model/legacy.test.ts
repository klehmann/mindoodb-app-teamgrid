import { describe, expect, it } from 'vitest'

import type { WorkbookFile } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { createWorkbook, loadWorkbook } from '../haven/store'
import { AutomergeTestHost } from '../testing/automerge-host'
import { createNodeEngine, nodeEngineAvailable } from '../testing/node-engine'
import { sheetAxes, workbookAxes, workbookToXlsx } from './export'
import { createAxesLookup, renderFormula } from './formula-refs'
import { convertLegacyWorkbook, legacyWorkbookOf, type LegacyWorkbook } from './legacy'
import { readWholeWorkbook } from './sidecar-read'
import { styleOf } from './styles'
import { emptyStoredWorkbook, writesFor } from './sync'
import { LEGACY, rows } from '../testing/legacy-sample'

describe('TeamGrid 1.x workbooks', () => {
  const legacy = legacyWorkbookOf(LEGACY) as LegacyWorkbook
  const workbook = convertLegacyWorkbook(legacy, '2026-10-08T00:00:00.000Z')
  const lookup = createAxesLookup(workbookAxes(workbook))
  const main = workbook.worksheetsById.sheet_main!
  const view = workbook.worksheetsById.sheet_view!

  it('recognises only 1.x documents', () => {
    expect(legacyWorkbookOf({ form: 'teamgrid-next' })).toBeUndefined()
    expect(legacy.worksheetOrder).toEqual(['sheet_main', 'sheet_view'])
  })

  it('keeps sheets, rows, columns and their ids', () => {
    expect(workbook.worksheetOrder).toEqual(['sheet_main', 'sheet_view'])
    expect(main.name).toBe('Umsatz')
    expect(sheetAxes(main).rowIds).toEqual(rows)
    expect(main.rowsById.row_gone?.deletedAt).toBe('2025-01-01T00:00:00Z')
    expect(main.columnsById.col_a?.width).toBe(20)
  })

  it('converts values, number formats and styles', () => {
    expect(main.cellsById['row_1:col_a']).toMatchObject({ value: 'Monat' })
    expect(styleOf(main.cellsById['row_1:col_a'])).toEqual({ bold: true, fillColor: '#FFFF00' })
    expect(main.cellsById['row_2:col_a']).toMatchObject({ value: 45658 })
    expect(styleOf(main.cellsById['row_2:col_a'])).toEqual({ numberFormat: 'mmm d, yyyy' })
    expect(styleOf(main.cellsById['row_2:col_b'])).toEqual({ numberFormat: '€0.00' })
    expect(styleOf(main.cellsById['row_3:col_b'])).toEqual({
      horizontalAlignment: 'center',
      verticalAlignment: 'center',
      borderBottom: { style: 'thin', color: '#000000' },
    })
  })

  it('keeps formulas bound to the same ids, across sheets too', () => {
    const total = main.cellsById['row_4:col_b']!
    expect(total.value).toBe(2000)
    expect(renderFormula(total.formula!, sheetAxes(main), lookup)).toBe('SUM(B2:B3)')
    expect(renderFormula(view.cellsById['vrow_1:vcol_1']!.formula!, sheetAxes(view), lookup)).toBe('Umsatz!B4')
  })

  it('turns charts into chart visuals and keeps view sheets bound', () => {
    expect(main.visualOrder).toEqual(['chart_1'])
    const chart = main.visualsById.chart_1!.chart!
    expect(chart).toMatchObject({ chartType: 'column', title: 'Umsatz', legend: 'bottom' })
    expect(chart.series[0]).toMatchObject({ name: 'Betrag', values: [1200, 800], categories: ['2025-01-01', 'Feb'] })
    expect(view.viewBinding).toEqual({
      viewId: 'contacts',
      viewTitle: 'Kontakte',
      showDocuments: true,
      showCategories: false,
      rootCategoryPath: ['ACME'],
      lastRefreshedAt: '2026-10-08T00:00:00.000Z',
    })
  })

  it('stores the copy and opens it in the editor', async () => {
    const host = new AutomergeTestHost()
    const copy = convertLegacyWorkbook(legacy)
    const id = await createWorkbook(host.connection, 'Umsatz 2025', copy, writesFor(emptyStoredWorkbook(), copy, copy.worksheetOrder))
    const loaded = await loadWorkbook(host.connection, id)
    expect(loaded.stored.workbook.worksheetsById.sheet_main!.cellsById['row_4:col_b']?.value).toBe(2000)
    if (!nodeEngineAvailable) return
    const engine = createNodeEngine()
    const { bytes } = await workbookToXlsx(loaded.stored.workbook, engine)
    const file = (await engine.open(bytes, 'de')) as WorkbookFile
    const read = await readWholeWorkbook(engine, file)
    expect(read.sheets.map((sheet) => sheet.meta.name)).toEqual(['Umsatz', 'Kontakte'])
    const total = read.sheets[0]!.cells.find((cell) => cell.row === 3 && cell.column === 1)
    expect(total?.formula?.replace(/^=/, '')).toBe('SUM(B2:B3)')
    expect(read.file.visuals.filter((visual) => visual.kind === 'chart')).toHaveLength(1)
  })
})
