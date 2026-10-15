// A TeamGrid 1.x workbook document (form `teamgrid`, schema 3) for tests and
// the test host: two sheets, values, formats, formulas, a chart, a view sheet.
export const rows = ['row_1', 'row_2', 'row_3', 'row_4']
const columns = ['col_a', 'col_b', 'col_c', 'col_d', 'col_e', 'col_f']
const range = (startRowId: string, endRowId: string, column: string) => ({
  worksheetId: 'sheet_main',
  startRowId,
  endRowId,
  startColumnId: column,
  endColumnId: column,
})

/** A 1.x workbook as classic TeamGrid stored it (the document's data). */
export const LEGACY = {
  form: 'teamgrid',
  kind: 'mindoodb.teamgrid',
  subject: 'Umsatz 2025',
  tags: ['Finanzen'],
  istemplate: false,
  teamgrid: {
    schemaVersion: 3,
    workbook: {
      id: 'book_1',
      worksheetOrder: ['sheet_main', 'sheet_view'],
      worksheetsById: {
        sheet_main: {
          id: 'sheet_main',
          title: 'Umsatz',
          rowOrder: [...rows, 'row_gone'],
          columnOrder: columns,
          rowsById: { ...Object.fromEntries(rows.map((id) => [id, { id }])), row_gone: { id: 'row_gone', deletedAt: '2025-01-01T00:00:00Z' } },
          columnsById: { ...Object.fromEntries(columns.map((id) => [id, { id, width: 120 }])), col_a: { id: 'col_a', width: 140 } },
          cellsById: {
            'row_1:col_a': { id: 'row_1:col_a', rowId: 'row_1', columnId: 'col_a', value: { kind: 'string', text: 'Monat' }, style: { bold: true, backgroundColor: '#ff0' } },
            'row_1:col_b': { id: 'row_1:col_b', rowId: 'row_1', columnId: 'col_b', value: { kind: 'string', text: 'Betrag' }, style: { bold: true } },
            'row_2:col_a': { id: 'row_2:col_a', rowId: 'row_2', columnId: 'col_a', value: { kind: 'date', isoDate: '2025-01-01T00:00:00.000Z', format: 'date' } },
            'row_2:col_b': { id: 'row_2:col_b', rowId: 'row_2', columnId: 'col_b', value: { kind: 'number', value: 1200, format: 'currency', currencyCode: 'EUR' } },
            'row_3:col_a': { id: 'row_3:col_a', rowId: 'row_3', columnId: 'col_a', value: { kind: 'string', text: 'Feb' } },
            'row_3:col_b': { id: 'row_3:col_b', rowId: 'row_3', columnId: 'col_b', value: { kind: 'number', value: 800 }, style: { horizontalAlign: 'center', verticalAlign: 'middle', borders: { bottom: { style: 'thin', color: '#000000' } } } },
            'row_4:col_a': {
              id: 'row_4:col_a',
              rowId: 'row_4',
              columnId: 'col_a',
              value: { kind: 'empty' },
              formula: { kind: 'formula', source: '=MAX(Betraege)', segments: [{ kind: 'text', text: '=MAX(Betraege)' }], references: [], cached: { kind: 'number', value: 1200 } },
            },
            'row_4:col_b': {
              id: 'row_4:col_b',
              rowId: 'row_4',
              columnId: 'col_b',
              value: { kind: 'empty' },
              formula: {
                kind: 'formula',
                source: '=SUM(B2:B3)',
                segments: [
                  { kind: 'text', text: '=SUM(' },
                  { kind: 'reference', reference: { kind: 'range', ...range('row_2', 'row_3', 'col_b') } },
                  { kind: 'text', text: ')' },
                ],
                references: [],
                cached: { kind: 'number', value: 2000 },
              },
            },
          },
          chartOrder: ['chart_1'],
          chartsById: {
            chart_1: {
              id: 'chart_1',
              type: 'column',
              title: 'Umsatz',
              series: [{ id: 's1', name: 'Betrag', values: range('row_2', 'row_3', 'col_b') }],
              categoryAxis: range('row_2', 'row_3', 'col_a'),
              anchor: {
                from: { rowId: 'row_1', columnId: 'col_c', rowOffsetEmu: 0, colOffsetEmu: 0 },
                to: { rowId: 'row_4', columnId: 'col_f', rowOffsetEmu: 150_000, colOffsetEmu: 300_000 },
              },
              legend: { position: 'bottom' },
            },
          },
        },
        sheet_view: {
          id: 'sheet_view',
          title: 'Kontakte',
          rowOrder: ['vrow_1'],
          columnOrder: ['vcol_1'],
          rowsById: { vrow_1: { id: 'vrow_1' } },
          columnsById: { vcol_1: { id: 'vcol_1' } },
          cellsById: {
            'vrow_1:vcol_1': {
              id: 'vrow_1:vcol_1',
              rowId: 'vrow_1',
              columnId: 'vcol_1',
              value: { kind: 'empty' },
              formula: {
                kind: 'formula',
                source: '=Umsatz!B4',
                segments: [
                  { kind: 'text', text: '=' },
                  { kind: 'reference', reference: { kind: 'cell', worksheetId: 'sheet_main', rowId: 'row_4', columnId: 'col_b' } },
                ],
                references: [],
                cached: { kind: 'number', value: 2000 },
              },
            },
          },
          chartOrder: [],
          chartsById: {},
          viewBinding: { kind: 'mindoodbView', viewId: 'contacts', viewTitle: 'Kontakte', showDocuments: true, showCategories: false, rootCategoryPath: ['ACME'] },
        },
      },
    },
    namedExpressionsById: {
      ne_1: { id: 'ne_1', name: 'Betraege', reference: { kind: 'range', ...range('row_2', 'row_3', 'col_b') } },
    },
    settings: { locale: 'de-DE' },
  },
}
