import { createCellId, type CellId, type FormulaReference, type Worksheet } from "@/lib/teamgridDocument";

export interface DependencyGraph {
  dependentsByCellId: Map<CellId, Set<CellId>>;
  dependenciesByCellId: Map<CellId, Set<CellId>>;
}

/**
 * Build a simple formula dependency graph from persisted stable-ID references.
 * This is intentionally separate from evaluation so the sample app can explain
 * recalculation invalidation without binding itself to a formula engine license.
 */
export function buildDependencyGraph(worksheet: Worksheet): DependencyGraph {
  const dependentsByCellId = new Map<CellId, Set<CellId>>();
  const dependenciesByCellId = new Map<CellId, Set<CellId>>();

  for (const cell of Object.values(worksheet.cellsById)) {
    if (!cell.formula) {
      continue;
    }
    const dependencies = new Set(expandReferences(worksheet, cell.formula.references));
    dependenciesByCellId.set(cell.id, dependencies);
    for (const dependency of dependencies) {
      const dependents = dependentsByCellId.get(dependency) ?? new Set<CellId>();
      dependents.add(cell.id);
      dependentsByCellId.set(dependency, dependents);
    }
  }

  return { dependentsByCellId, dependenciesByCellId };
}

export function collectDirtyFormulaCells(graph: DependencyGraph, changedCellIds: CellId[]) {
  const dirty = new Set<CellId>();
  const queue = [...changedCellIds];
  while (queue.length > 0) {
    const cellId = queue.shift();
    if (!cellId) {
      continue;
    }
    for (const dependent of graph.dependentsByCellId.get(cellId) ?? []) {
      if (dirty.has(dependent)) {
        continue;
      }
      dirty.add(dependent);
      queue.push(dependent);
    }
  }
  return [...dirty];
}

function expandReferences(worksheet: Worksheet, references: FormulaReference[]) {
  return references.flatMap<CellId>((reference) => {
    if (reference.kind === "cell") {
      return [createCellId(reference.rowId, reference.columnId)];
    }
    if (reference.kind === "column") {
      return worksheet.rowOrder.map((rowId) => createCellId(rowId, reference.columnId));
    }
    const startRow = worksheet.rowOrder.indexOf(reference.startRowId);
    const endRow = worksheet.rowOrder.indexOf(reference.endRowId);
    const startColumn = worksheet.columnOrder.indexOf(reference.startColumnId);
    const endColumn = worksheet.columnOrder.indexOf(reference.endColumnId);
    if (startRow < 0 || endRow < 0 || startColumn < 0 || endColumn < 0) {
      return [];
    }
    const cellIds: CellId[] = [];
    for (let rowIndex = Math.min(startRow, endRow); rowIndex <= Math.max(startRow, endRow); rowIndex += 1) {
      for (let columnIndex = Math.min(startColumn, endColumn); columnIndex <= Math.max(startColumn, endColumn); columnIndex += 1) {
        cellIds.push(createCellId(worksheet.rowOrder[rowIndex], worksheet.columnOrder[columnIndex]));
      }
    }
    return cellIds;
  });
}
