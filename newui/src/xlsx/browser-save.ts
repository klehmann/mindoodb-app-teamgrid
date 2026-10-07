// Applies a GenOffice WorkbookSaveRequest to the workbook's source bytes in
// the browser and returns the new xlsx bytes.
//
// The request → gateway mapping is a port of `writeWorkbookTo` in GenOffice's
// apps/sheets/src/main/sheets-main.ts (Apache-2.0); the planning itself is the
// vendored xlsx-gateway, unchanged. Only the archive assembly differs: the
// desktop app streams through the Rust sidecar, here JSZip rebuilds the
// package in memory.
import JSZip from 'jszip'

import {
  createBufferEntrySource,
  planCellEditsToXlsx,
  type CellEdit,
  type MutationPlan,
  type SheetStructuralOps,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { SheetEditPlan } from '@genoffice/xlsx-gateway/gateway/xlsx-sheets'

import type { WorkbookSaveRequest } from '../../vendor/genoffice/apps/sheets/src/shared/desktop-api'

export async function saveWorkbookInBrowser(
  source: Uint8Array,
  sheetNames: ReadonlyMap<string, string>,
  request: WorkbookSaveRequest,
): Promise<{ bytes: Uint8Array; touchedEntries: readonly string[] }> {
  const addedSheetNames = new Map<string, string>()
  const duplicateSources = new Map<string, string>()
  const renames: { sheetName: string; newName: string }[] = []
  const removals: string[] = []
  const hiddenChanges: { sheetName: string; hidden: boolean }[] = []
  let orderChanged = false
  for (const op of request.sheetOps) {
    if (op.kind === 'add-sheet') {
      addedSheetNames.set(op.sheetId, op.name)
      continue
    }
    if (op.kind === 'duplicate-sheet') {
      const sourceName = sheetNames.get(op.sourceSheetId)
      if (!sourceName) throw new Error(`Unknown duplicate source ${op.sourceSheetId}.`)
      addedSheetNames.set(op.sheetId, op.name)
      duplicateSources.set(op.sheetId, sourceName)
      continue
    }
    if (op.kind === 'reorder-sheets') {
      orderChanged = true
      continue
    }
    const sheetName = addedSheetNames.get(op.sheetId) ?? sheetNames.get(op.sheetId)
    if (!sheetName) throw new Error(`Unknown worksheet ${op.sheetId}.`)
    if (op.kind === 'rename-sheet') renames.push({ sheetName, newName: op.newName })
    else if (op.kind === 'set-sheet-hidden') hiddenChanges.push({ sheetName, hidden: op.hidden })
    else removals.push(sheetName)
  }
  const renameByOriginal = new Map(renames.map((rename) => [rename.sheetName, rename.newName]))
  const resolveSheetName = (sheetId: string): string => {
    const sheetName = addedSheetNames.get(sheetId) ?? sheetNames.get(sheetId)
    if (!sheetName) throw new Error(`Unknown worksheet ${sheetId}.`)
    return sheetName
  }
  let sheetPlan: SheetEditPlan | undefined
  if (request.sheetOps.length > 0) {
    sheetPlan = {
      renames,
      additions: [...addedSheetNames].map(([sheetId, name]) => ({
        name,
        sourceSheetName: duplicateSources.get(sheetId),
      })),
      removals,
      hiddenChanges,
      orderChanged,
      order: request.sheetOrder.map((sheetId) => {
        const original = resolveSheetName(sheetId)
        return addedSheetNames.has(sheetId) ? original : (renameByOriginal.get(original) ?? original)
      }),
    }
  }

  const edits: CellEdit[] = request.edits.map((edit) => ({
    sheetName: resolveSheetName(edit.sheetId),
    row: edit.row,
    column: edit.column,
    writeValue: edit.writeValue,
    cell: { value: edit.value, formula: edit.formula },
    style: edit.style,
    rich: edit.rich,
    styleReset: edit.styleReset,
  }))
  const bulkConstantFills = (request.bulkConstantFills ?? []).map(({ sheetId, ...fill }) => ({
    sheetName: resolveSheetName(sheetId),
    ...fill,
  }))
  const opsBySheet = new Map<string, SheetStructuralOps['ops'][number][]>()
  for (const op of request.structuralOps) {
    const sheetName = resolveSheetName(op.sheetId)
    const sheetOps = opsBySheet.get(sheetName) ?? []
    if ('range' in op) {
      sheetOps.push({ kind: op.kind, range: op.range })
    } else if ('size' in op) {
      sheetOps.push({ kind: op.kind, start: op.start, end: op.end, size: op.size })
    } else if ('level' in op) {
      sheetOps.push({
        kind: op.kind,
        start: op.start,
        end: op.end,
        level: op.level,
        ...(op.collapsed === undefined ? {} : { collapsed: op.collapsed }),
      })
    } else if ('hidden' in op) {
      sheetOps.push({ kind: op.kind, start: op.start, end: op.end, hidden: op.hidden })
    } else if ('style' in op) {
      sheetOps.push({ kind: op.kind, start: op.start, end: op.end, style: op.style })
    } else if ('before' in op) {
      sheetOps.push({ kind: op.kind, index: op.index, count: op.count, before: op.before })
    } else {
      sheetOps.push({ kind: op.kind, index: op.index, count: op.count })
    }
    opsBySheet.set(sheetName, sheetOps)
  }
  const structuralOps: SheetStructuralOps[] = [...opsBySheet].map(([sheetName, ops]) => ({
    sheetName,
    ops,
  }))
  const linksBySheet = new Map<string, { row: number; column: number; target: string | null }[]>()
  for (const link of request.hyperlinkEdits) {
    const sheetName = resolveSheetName(link.sheetId)
    const sheetLinks = linksBySheet.get(sheetName) ?? []
    sheetLinks.push({ row: link.row, column: link.column, target: link.target })
    linksBySheet.set(sheetName, sheetLinks)
  }
  const formulaValuesBySheet = new Map<
    string,
    { row: number; column: number; value: string | number | boolean | null | { error: string } }[]
  >()
  for (const cell of request.formulaValues) {
    const sheetName = resolveSheetName(cell.sheetId)
    const list = formulaValuesBySheet.get(sheetName) ?? []
    list.push({ row: cell.row, column: cell.column, value: cell.value })
    formulaValuesBySheet.set(sheetName, list)
  }

  const plan = await planCellEditsToXlsx(
    await createBufferEntrySource(source as never),
    edits,
    structuralOps,
    request.chartEdits,
    sheetPlan,
    request.filterStates.map((state) => ({
      sheetName: resolveSheetName(state.sheetId),
      filter: state.filter,
      hiddenRows: state.hiddenRows,
      visibilityRange: state.visibilityRange,
    })),
    [...linksBySheet].map(([sheetName, links]) => ({ sheetName, edits: links })),
    request.cfStates.map((state) => ({ sheetName: resolveSheetName(state.sheetId), rules: state.rules })),
    request.dvStates.map((state) => ({ sheetName: resolveSheetName(state.sheetId), rules: state.rules })),
    request.sheetProtections.map((state) => ({
      sheetName: resolveSheetName(state.sheetId),
      protected: state.protected,
    })),
    request.definedNamesState,
    request.visualAdditions.map((addition) => ({
      sheetName: resolveSheetName(addition.sheetId),
      anchor: addition.anchor,
      chart: addition.chart,
      shape: addition.shape,
      image: addition.image,
    })),
    request.pageSetupStates.map(({ sheetId, ...state }) => ({ sheetName: resolveSheetName(sheetId), ...state })),
    request.noteStates.map(({ sheetId, notes }) => ({ sheetName: resolveSheetName(sheetId), notes })),
    request.tableAdditions.map((table) => ({
      sheetName: resolveSheetName(table.sheetId),
      area: table.area,
      name: table.name,
      columnNames: table.columnNames,
      style: table.style,
      bandedRows: table.bandedRows,
    })),
    request.pivotAdditions.map(({ sheetId, sourceSheetId, ...pivot }) => ({
      ...pivot,
      sheetName: resolveSheetName(sheetId),
      sourceSheetName: resolveSheetName(sourceSheetId),
    })),
    request.pivotCacheRefreshPaths,
    request.pivotRefreshUpdates.map((update) => ({
      cachePath: update.cachePath,
      sheetName: resolveSheetName(update.sheetId),
      newOutputRef: update.newOutputRef,
      ...(update.relayout === undefined
        ? {}
        : {
            relayout: (({ sheetId: _sheetId, sourceSheetId, ...rest }) => ({
              ...rest,
              sourceSheetName: resolveSheetName(sourceSheetId),
            }))(update.relayout),
          }),
    })),
    request.visualEdits,
    request.sparklineAdditions.map(({ sheetId, ...group }) => ({ sheetName: resolveSheetName(sheetId), ...group })),
    [...formulaValuesBySheet].map(([sheetName, cells]) => ({ sheetName, cells })),
    request.themeState,
    request.workbookProtectionState,
    request.protectedRangeStates.map((state) => ({
      sheetName: resolveSheetName(state.sheetId),
      ranges: state.ranges,
    })),
    bulkConstantFills,
  )
  return { bytes: await assemble(source, plan), touchedEntries: plan.touchedEntries }
}

async function assemble(source: Uint8Array, plan: MutationPlan): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(source, { checkCRC32: true })
  for (const path of plan.removedEntries) zip.remove(path)
  for (const [path, content] of plan.replaced) zip.file(path, content, { createFolders: false })
  for (const [path, content] of plan.added) zip.file(path, content, { createFolders: false })
  for (const [path, bytes] of plan.addedBinary) zip.file(path, bytes, { createFolders: false })
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } })
}
