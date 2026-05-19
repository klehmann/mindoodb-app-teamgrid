/**
 * Formula-bar draft state and Excel-style "formula picking".
 *
 * The composable mirrors the active cell into a draft string the user can
 * edit in the formula bar, parses the draft live to highlight referenced
 * cells inside the grid, and exposes the `appendPickedAddress` helper so
 * clicking another cell while editing a formula appends its `A1` address
 * to the draft (with an inserted `+` when necessary).
 *
 * Commit / cancel both drop out of edit mode and close the floating
 * formula-assist panel through {@link UseFormulaBarEditingOptions.closeFormulaAssist}.
 */
import { computed, ref, watch, type Ref } from "vue";
import { formatCellValue } from "@/features/grid/lib/cellFormatting";
import { parseFormula, renderFormulaSource, type FormulaContext } from "@/features/formulas/lib";
import { createCellId, type Cell, type Worksheet } from "@/features/document/lib/teamgridDocument";
import type { GridProjection } from "@/features/grid/lib/gridProjection";

export interface UseFormulaBarEditingOptions {
  selectedCell: Readonly<Ref<Cell | null>>;
  activeWorksheet: Readonly<Ref<Worksheet | null>>;
  formulaContext: Readonly<Ref<FormulaContext | null>>;
  projection: Readonly<Ref<GridProjection | null>>;
  locale: Readonly<Ref<string | undefined>>;
  commitCell: (cell: Cell, rawValue: string) => void;
  closeFormulaAssist: () => void;
}

export function useFormulaBarEditing(options: UseFormulaBarEditingOptions) {
  const { selectedCell, activeWorksheet, formulaContext, projection, locale, commitCell, closeFormulaAssist } = options;
  const formulaDraft = ref("");
  const formulaError = ref<string | null>(null);
  const formulaEditing = ref(false);

  /**
   * Cell ids that the grid should overlay while the user is composing a
   * formula in the formula bar. Each parsed reference contributes one or
   * more cells so the user gets instant visual feedback for formula targets.
   */
  const highlightedCellIds = computed(() => {
    if (!activeWorksheet.value || !formulaContext.value || !projection.value || !formulaDraft.value.trim().startsWith("=")) {
      return [];
    }
    const parsed = parseFormula(formulaDraft.value, activeWorksheet.value.id, formulaContext.value);
    if ("code" in parsed) {
      return [];
    }
    return parsed.references.flatMap((reference) => {
      if (reference.worksheetId !== activeWorksheet.value?.id) {
        return [];
      }
      if (reference.kind === "cell") {
        return [createCellId(reference.rowId, reference.columnId)];
      }
      if (reference.kind === "range") {
        const startRow = projection.value?.rowIndexById.get(reference.startRowId);
        const endRow = projection.value?.rowIndexById.get(reference.endRowId);
        const startColumn = projection.value?.columnIndexById.get(reference.startColumnId);
        const endColumn = projection.value?.columnIndexById.get(reference.endColumnId);
        if (startRow == null || endRow == null || startColumn == null || endColumn == null || !projection.value) {
          return [];
        }
        const cellIds: string[] = [];
        for (let rowIndex = Math.min(startRow, endRow); rowIndex <= Math.max(startRow, endRow); rowIndex += 1) {
          for (let columnIndex = Math.min(startColumn, endColumn); columnIndex <= Math.max(startColumn, endColumn); columnIndex += 1) {
            const row = projection.value.rows[rowIndex];
            const column = projection.value.columns[columnIndex];
            if (row && column) {
              cellIds.push(createCellId(row.id, column.id));
            }
          }
        }
        return cellIds;
      }
      return [];
    });
  });

  /**
   * Append a picked cell address to the current draft.
   *
   * Inserts a `+` operator when the draft ends with an identifier or
   * closing paren, so picking `B1` after typing `=A1` produces `=A1+B1`
   * instead of the syntactically invalid `=A1B1`.
   */
  function appendPickedAddress(source: string, address: string) {
    if (/[\w)]$/.test(source.trimEnd())) {
      return `${source}+${address}`;
    }
    return `${source}${address}`;
  }

  function commitFormulaBar(value: string) {
    if (!selectedCell.value) {
      return;
    }
    formulaEditing.value = false;
    closeFormulaAssist();
    commitCell(selectedCell.value, value);
  }

  function cancelFormulaEdit() {
    formulaEditing.value = false;
    closeFormulaAssist();
    formulaError.value = null;
    formulaDraft.value = selectedCell.value?.formula && activeWorksheet.value && formulaContext.value
      ? renderFormulaSource(selectedCell.value.formula, activeWorksheet.value.id, formulaContext.value)
      : (selectedCell.value ? formatCellValue(selectedCell.value.value, locale.value) : "");
  }

  // Mirror the selected cell into the formula bar draft. Formula cells
  // expose their `source` (e.g. `=SUM(A1:A10)`); plain cells get their
  // formatted display value so the user can edit it as text.
  watch(
    [selectedCell, activeWorksheet, formulaContext],
    ([cell]) => {
      if (!cell) {
        formulaDraft.value = "";
        return;
      }
      formulaDraft.value = cell.formula && activeWorksheet.value && formulaContext.value
        ? renderFormulaSource(cell.formula, activeWorksheet.value.id, formulaContext.value)
        : formatCellValue(cell.value, locale.value);
    },
  );

  return {
    formulaDraft,
    formulaError,
    formulaEditing,
    highlightedCellIds,
    appendPickedAddress,
    commitFormulaBar,
    cancelFormulaEdit,
  };
}
