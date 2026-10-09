// PivotTables in the stored workbook: their raw parts (pivot-parts.ts) as an
// attachment of the top document, named after a hash of the content, so an
// unchanged pivot is never uploaded again; position and source range as
// row/column ids, so both move with rows and columns inserted elsewhere.
import type { ExtractedPivot, PivotParts, PivotPlacement } from '../xlsx/pivot-parts'
import type { AxesLookup } from './formula-refs'
import type { IdArea, SheetId, Worksheet } from './schema'
import { areaToIds, areaToIndices } from './sheet-features'

export interface StoredPivot {
  name: string
  location: IdArea
  source?: { sheetId: SheetId; area: IdArea }
  /** Attachment of the top document holding the parts as JSON. */
  attachment: string
}

/** A pivot attachment to upload. */
export interface NewPivotParts {
  attachment: string
  bytes: Uint8Array
}

function hash(text: string): string {
  let h1 = 0x811c9dc5
  let h2 = 0x01000193
  for (let index = 0; index < text.length; index += 1) {
    h1 = Math.imul(h1 ^ text.charCodeAt(index), 16777619) >>> 0
    h2 = Math.imul(h2 ^ text.charCodeAt(index), 2246822519) >>> 0
  }
  return `${h1.toString(36)}${h2.toString(36)}`
}

/** The pivots of one sheet, keyed by lower-case name; new attachments are added to `uploads`. */
export function storedPivots(
  sheetName: string,
  pivots: readonly ExtractedPivot[],
  previous: Worksheet['pivotsById'],
  lookup: AxesLookup,
  uploads: NewPivotParts[],
): Worksheet['pivotsById'] {
  const home = lookup.bySheetName(sheetName)
  if (!home) return undefined
  const byKey: NonNullable<Worksheet['pivotsById']> = {}
  for (const pivot of pivots) {
    if (pivot.sheetName !== sheetName) continue
    const location = areaToIds(pivot.location, home.rowIds, home.columnIds)
    if (!location) continue
    const sourceSheet = pivot.source ? lookup.bySheetName(pivot.source.sheetName) : undefined
    const sourceArea = pivot.source && sourceSheet ? areaToIds(pivot.source.area, sourceSheet.rowIds, sourceSheet.columnIds) : undefined
    const json = JSON.stringify(pivot.parts)
    const attachment = `pivot-${hash(json)}.json`
    const key = pivot.name.toLowerCase()
    if (previous?.[key]?.attachment !== attachment && !uploads.some((upload) => upload.attachment === attachment)) {
      uploads.push({ attachment, bytes: new TextEncoder().encode(json) })
    }
    byKey[key] = {
      name: pivot.name,
      location,
      ...(sourceArea && sourceSheet ? { source: { sheetId: sourceSheet.id, area: sourceArea } } : {}),
      attachment,
    }
  }
  return Object.keys(byKey).length ? byKey : undefined
}

/** Where each stored pivot goes in a file built from the workbook. */
export async function pivotPlacements(
  sheets: readonly Worksheet[],
  lookup: AxesLookup,
  loadParts: (attachment: string) => Promise<PivotParts | undefined>,
): Promise<PivotPlacement[]> {
  const placements: PivotPlacement[] = []
  for (const sheet of sheets) {
    const home = lookup.bySheetId(sheet.id)
    if (!home) continue
    const rowIndex = new Map(home.rowIds.map((id, index) => [id, index]))
    const columnIndex = new Map(home.columnIds.map((id, index) => [id, index]))
    for (const pivot of Object.values(sheet.pivotsById ?? {})) {
      const location = areaToIndices(pivot.location, rowIndex, columnIndex)
      const parts = location ? await loadParts(pivot.attachment) : undefined
      if (!location || !parts) continue
      let source: PivotPlacement['source']
      if (pivot.source) {
        const sourceSheet = lookup.bySheetId(pivot.source.sheetId)
        const area = sourceSheet
          ? areaToIndices(
              pivot.source.area,
              new Map(sourceSheet.rowIds.map((id, index) => [id, index])),
              new Map(sourceSheet.columnIds.map((id, index) => [id, index])),
            )
          : undefined
        // A deleted source sheet or range: the pivot keeps its last cache, Excel shows it as is.
        if (sourceSheet && area) source = { sheetName: sourceSheet.name, area }
      }
      placements.push({ sheetName: sheet.name, location, ...(source ? { source } : {}), parts })
    }
  }
  return placements
}
