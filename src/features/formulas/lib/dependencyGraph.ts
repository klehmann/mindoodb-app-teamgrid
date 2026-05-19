import { createCellId, type CellId, type FormulaReference, type Workbook, type Worksheet, type WorksheetId } from "@/features/document/lib/teamgridDocument";

export type FormulaCellKey = `${WorksheetId}:${CellId}`;

export interface DependencyGraph {
  dependentsByCellId: Map<FormulaCellKey, Set<FormulaCellKey>>;
  dependenciesByCellId: Map<FormulaCellKey, Set<FormulaCellKey>>;
}

/**
 * Build a simple formula dependency graph from persisted stable-ID references.
 * This is intentionally separate from evaluation so the sample app can explain
 * recalculation invalidation without binding itself to a formula engine license.
 */
export function buildDependencyGraph(workbookOrWorksheet: Workbook | Worksheet): DependencyGraph {
  const dependentsByCellId = new Map<FormulaCellKey, Set<FormulaCellKey>>();
  const dependenciesByCellId = new Map<FormulaCellKey, Set<FormulaCellKey>>();
  const worksheets = "worksheetsById" in workbookOrWorksheet
    ? workbookOrWorksheet.worksheetOrder.map((worksheetId) => workbookOrWorksheet.worksheetsById[worksheetId]).filter((worksheet): worksheet is Worksheet => Boolean(worksheet && !worksheet.deletedAt))
    : [workbookOrWorksheet];

  for (const worksheet of worksheets) {
    for (const cell of Object.values(worksheet.cellsById)) {
      if (!cell.formula) {
        continue;
      }
      const dependencies = new Set(expandReferences(workbookOrWorksheet, cell.formula.references));
      const cellKey = createFormulaCellKey(worksheet.id, cell.id);
      dependenciesByCellId.set(cellKey, dependencies);
      for (const dependency of dependencies) {
        const dependents = dependentsByCellId.get(dependency) ?? new Set<FormulaCellKey>();
        dependents.add(cellKey);
        dependentsByCellId.set(dependency, dependents);
      }
    }
  }

  return { dependentsByCellId, dependenciesByCellId };
}

export function collectDirtyFormulaCells(graph: DependencyGraph, changedCellIds: FormulaCellKey[]) {
  const dirty = new Set<FormulaCellKey>();
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

export function createFormulaCellKey(worksheetId: WorksheetId, cellId: CellId): FormulaCellKey {
  return `${worksheetId}:${cellId}`;
}

function expandReferences(workbookOrWorksheet: Workbook | Worksheet, references: FormulaReference[]) {
  return references.flatMap<FormulaCellKey>((reference) => {
    const worksheet = "worksheetsById" in workbookOrWorksheet
      ? workbookOrWorksheet.worksheetsById[reference.worksheetId]
      : workbookOrWorksheet;
    if (!worksheet || worksheet.deletedAt || worksheet.id !== reference.worksheetId) {
      return [] satisfies FormulaCellKey[];
    }
    if (reference.kind === "cell") {
      return [createFormulaCellKey(reference.worksheetId, createCellId(reference.rowId, reference.columnId))];
    }
    if (reference.kind === "column") {
      return worksheet.rowOrder.map((rowId) => createFormulaCellKey(reference.worksheetId, createCellId(rowId, reference.columnId)));
    }
    const startRow = worksheet.rowOrder.indexOf(reference.startRowId);
    const endRow = worksheet.rowOrder.indexOf(reference.endRowId);
    const startColumn = worksheet.columnOrder.indexOf(reference.startColumnId);
    const endColumn = worksheet.columnOrder.indexOf(reference.endColumnId);
    if (startRow < 0 || endRow < 0 || startColumn < 0 || endColumn < 0) {
      return [] satisfies FormulaCellKey[];
    }
    const cellIds: FormulaCellKey[] = [];
    for (let rowIndex = Math.min(startRow, endRow); rowIndex <= Math.max(startRow, endRow); rowIndex += 1) {
      for (let columnIndex = Math.min(startColumn, endColumn); columnIndex <= Math.max(startColumn, endColumn); columnIndex += 1) {
        cellIds.push(createFormulaCellKey(reference.worksheetId, createCellId(worksheet.rowOrder[rowIndex], worksheet.columnOrder[columnIndex])));
      }
    }
    return cellIds;
  });
}
