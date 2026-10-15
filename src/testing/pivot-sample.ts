// An xlsx with a PivotTable, written by GenOffice's gateway the way the
// editor saves a new pivot: source data A1:C5 on "Daten", the pivot (sum of
// Wert by Team) baked at E1 on the same sheet.
import { blankXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'

import type { WorkbookSaveRequest } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { emptySaveRequest } from '../model/export'
import { saveWorkbookInBrowser } from '../xlsx/browser-save'

export const PIVOT_SHEET = 'sheet-1'

export async function pivotSampleXlsx(): Promise<Uint8Array> {
  const blank = new Uint8Array(await blankXlsxBuffer('Daten'))
  const request = emptySaveRequest('pivot')
  const rows = [['Monat', 'Team', 'Wert'], ['Jan', 'x', 1], ['Jan', 'y', 2], ['Feb', 'x', 3], ['Feb', 'y', 4]]
  const baked = [['Team', 'Summe von Wert'], ['x', 4], ['y', 6], ['Gesamtergebnis', 10]]
  const put = (row: number, column: number, value: string | number) =>
    request.edits.push({ sheetId: PIVOT_SHEET, row, column, writeValue: true, value } as never)
  rows.forEach((row, index) => row.forEach((value, column) => put(index, column, value)))
  baked.forEach((row, index) => row.forEach((value, column) => put(index, column + 4, value)))
  request.pivotAdditions.push({
    sheetId: PIVOT_SHEET,
    sourceSheetId: PIVOT_SHEET,
    sourceArea: { startRow: 0, startColumn: 0, endRow: 4, endColumn: 2 },
    location: { startRow: 0, startColumn: 4, endRow: 3, endColumn: 5 },
    name: 'PivotUmsatz',
    fieldNames: ['Monat', 'Team', 'Wert'],
    rowFieldIndices: [1],
    rowItems: ['x', 'y'],
    values: [{ fieldIndex: 2, agg: 'sum' }],
  } as WorkbookSaveRequest['pivotAdditions'][number])
  return (await saveWorkbookInBrowser(blank, new Map([[PIVOT_SHEET, 'Daten']]), request)).bytes
}
