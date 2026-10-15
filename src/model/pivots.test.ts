// PivotTables through the real pipeline: the gateway writes a pivot, it is
// stored (parts as an attachment, position and source by id) and written
// back into a file built from scratch. Skipped without pkg-node.
import { describe, expect, it } from 'vitest'

import type { WorkbookFile } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'
import { createWorkbook, loadWorkbook, readAttachmentBase64 } from '../haven/store'
import { AutomergeTestHost } from '../testing/automerge-host'
import { createNodeEngine, nodeEngineAvailable } from '../testing/node-engine'
import { pivotSampleXlsx } from '../testing/pivot-sample'
import type { XlsxEngine } from '../xlsx/engine'
import { extractPivots, formatRef, parseRef } from '../xlsx/pivot-parts'
import { workbookToXlsx } from './export'
import { createId } from './schema'
import { readWholeWorkbook } from './sidecar-read'
import { emptyStoredWorkbook, syncWorkbook } from './sync'

async function meta(engine: XlsxEngine, bytes: Uint8Array) {
  const file = (await engine.open(bytes, 'de')) as WorkbookFile
  const read = await readWholeWorkbook(engine, file)
  await engine.close(file.sessionId)
  return read
}

describe('pivot tables', () => {
  it('reads and formats cell references', () => {
    expect(parseRef('$E$1:$F$4')).toEqual({ startRow: 0, startColumn: 4, endRow: 3, endColumn: 5 })
    expect(formatRef({ startRow: 9, startColumn: 26, endRow: 10, endColumn: 27 })).toBe('AA10:AB11')
  })

  it('finds a pivot, its source and its parts in a file', async () => {
    const [pivot] = await extractPivots(await pivotSampleXlsx())
    expect(pivot).toMatchObject({
      sheetName: 'Daten',
      name: 'PivotUmsatz',
      location: { startRow: 0, startColumn: 4, endRow: 3, endColumn: 5 },
      source: { sheetName: 'Daten', area: { startRow: 0, startColumn: 0, endRow: 4, endColumn: 2 } },
    })
    expect(pivot!.parts.records).toContain('<pivotCacheRecords')
  })

  it.skipIf(!nodeEngineAvailable)('stores a pivot and writes it back, moved with an inserted row', async () => {
    const engine = createNodeEngine()
    const host = new AutomergeTestHost()
    const bytes = await pivotSampleXlsx()
    const source = await meta(engine, bytes)
    const { next, writes } = syncWorkbook({
      stored: emptyStoredWorkbook(),
      storedIdByFileId: new Map(),
      read: source,
      pivots: await extractPivots(bytes),
    })
    const id = await createWorkbook(host.connection, 'Pivot', next, writes)
    const loaded = await loadWorkbook(host.connection, id)
    const sheet = Object.values(loaded.stored.workbook.worksheetsById)[0]!
    expect(Object.keys(sheet.pivotsById ?? {})).toEqual(['pivotumsatz'])
    const loadImage = (attachment: string) => readAttachmentBase64(host.connection, id, attachment)

    // Written back unchanged: the editor reads the same pivot.
    const back = await meta(engine, (await workbookToXlsx(loaded.stored.workbook, engine, { loadImage })).bytes)
    const describe = (read: typeof source) =>
      read.sheets[0]!.meta.pivotTables.map(({ path: _path, cachePath: _cache, ...pivot }) => pivot)
    expect(describe(back)).toEqual(describe(source))

    // Someone inserted a row at the top: output and source move down one row.
    sheet.rowOrder = [createId('r'), ...sheet.rowOrder]
    const moved = await extractPivots((await workbookToXlsx(loaded.stored.workbook, engine, { loadImage })).bytes)
    expect(moved[0]).toMatchObject({
      location: { startRow: 1, startColumn: 4, endRow: 4, endColumn: 5 },
      source: { sheetName: 'Daten', area: { startRow: 1, startColumn: 0, endRow: 5, endColumn: 2 } },
    })
  })
})
