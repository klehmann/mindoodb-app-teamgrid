/**
 * State and orchestration for the Excel-like "Format cells" dialog.
 *
 * The dialog has four tabs (cell type, font, fill, border). Each tab owns
 * a handful of reactive draft refs; `applySelectedCellFormat` flattens
 * them into a single `setCellsStyle` operation so the persisted patch is
 * compact and the granular-save machinery sees one logical edit.
 *
 * Border editing is the most fiddly part: it supports both per-side
 * toggling and the four preset selections (None / Outline / Inside / All).
 * The composable computes the right per-cell border patch based on the
 * cell's position inside the bounds of the selection so e.g. "Outline"
 * only paints the outer border of the rectangle.
 */
import { ref, type Ref } from "vue";
import {
  applyCellFormat,
  type CellFormatKind,
  type CellFormatRequest,
} from "@/features/grid/lib/cellFormatting";
import {
  type Cell,
  type CellBorder,
  type CellBorderSide,
  type CellBorderStyle,
  type CellStyle,
  type CurrencyCode,
  type Worksheet,
} from "@/features/document/lib/teamgridDocument";
import type { TeamGridOperation } from "@/features/document/lib/teamgridOps";
import type { CellSelectionRange } from "@/features/grid/composables/useSelection";
import type { TeamGridAppApi } from "@/features/document/composables/useTeamGridDocument";

export type CellBorders = NonNullable<CellStyle["borders"]>;
export type CellBorderPatch = Partial<Record<CellBorderSide, CellBorder | null>>;
export type CellStylePatch = Omit<{
  [Key in keyof CellStyle]?: CellStyle[Key] | null;
}, "borders"> & {
  borders?: CellBorderPatch | null;
};
export type FormatDialogTab = "cellType" | "font" | "fill" | "border";
export type BorderPreset = "custom" | "none" | "outline" | "inside" | "all";

interface CellsWithCoordinates {
  cell: Cell;
  rowIndex: number;
  columnIndex: number;
}

interface RangeBounds {
  minRow: number;
  maxRow: number;
  minCol: number;
  maxCol: number;
}

export interface UseCellFormatDialogOptions {
  app: TeamGridAppApi;
  activeWorksheet: Readonly<Ref<Worksheet | null>>;
  selectedCell: Readonly<Ref<Cell | null>>;
  selectedRange: Ref<CellSelectionRange | null>;
  selectedCells: Readonly<Ref<Cell[]>>;
  cellsForRange: (range: CellSelectionRange | null) => CellsWithCoordinates[];
  boundsForRange: (range: CellSelectionRange | null) => RangeBounds | null;
}

export function useCellFormatDialog(options: UseCellFormatDialogOptions) {
  const { app, activeWorksheet, selectedCell, selectedRange, selectedCells, cellsForRange, boundsForRange } = options;

  const formatDialogVisible = ref(false);
  const formatDialogTab = ref<FormatDialogTab>("cellType");
  const formatDialogKind = ref<CellFormatKind>("general");
  const formatDialogCurrency = ref<CurrencyCode>("USD");
  const formatDialogCustomNumFmt = ref("");
  const formatDialogFontFamily = ref("Inter, sans-serif");
  const formatDialogFontSize = ref(14);
  const formatDialogBold = ref(false);
  const formatDialogItalic = ref(false);
  const formatDialogUnderline = ref(false);
  const formatDialogTextColor = ref("#111827");
  const formatDialogFillEnabled = ref(false);
  const formatDialogFillColor = ref("#ffffff");
  const formatDialogBorderStyle = ref<CellBorderStyle | "none">("thin");
  const formatDialogBorderColor = ref("#111827");
  const formatDialogBorderPreset = ref<BorderPreset>("custom");
  const formatDialogBorders = ref<CellBorders>({});

  /** Open the Excel-like value-format dialog for the active selection. */
  function openCellFormatDialog(range: CellSelectionRange | null) {
    const targetRange = range ?? (selectedCell.value ? { startCellId: selectedCell.value.id, endCellId: selectedCell.value.id } : null);
    if (app.gridReadOnly.value || !targetRange) {
      return;
    }
    selectedRange.value = targetRange;
    seedFormatDialogFromCell(selectedCell.value);
    formatDialogVisible.value = true;
  }

  function seedFormatDialogFromCell(cell: Cell | null) {
    const value = cell?.value;
    const excelNumFmt = value && "excelNumFmt" in value ? value.excelNumFmt : undefined;
    const style = cell?.style;
    formatDialogTab.value = "cellType";
    formatDialogKind.value = excelNumFmt && isCustomExcelNumFmt(excelNumFmt)
      ? "custom"
      : value?.kind === "string"
        ? "text"
        : value?.kind === "number"
          ? (value.format ?? "general")
          : "general";
    formatDialogCurrency.value = value?.kind === "number" && value.currencyCode ? value.currencyCode : "USD";
    formatDialogCustomNumFmt.value = value?.kind === "number" || value?.kind === "date" || value?.kind === "string"
      ? excelNumFmt ?? ""
      : "";
    formatDialogFontFamily.value = style?.fontFamily ?? "Inter, sans-serif";
    formatDialogFontSize.value = style?.fontSize ?? 14;
    formatDialogBold.value = Boolean(style?.bold);
    formatDialogItalic.value = Boolean(style?.italic);
    formatDialogUnderline.value = Boolean(style?.underline);
    formatDialogTextColor.value = style?.textColor ?? "#111827";
    formatDialogFillEnabled.value = Boolean(style?.backgroundColor);
    formatDialogFillColor.value = style?.backgroundColor ?? "#ffffff";
    formatDialogBorderStyle.value = firstBorderStyle(style?.borders) ?? "thin";
    formatDialogBorderColor.value = firstBorderColor(style?.borders) ?? "#111827";
    formatDialogBorderPreset.value = "custom";
    formatDialogBorders.value = { ...style?.borders };
  }

  function isCustomExcelNumFmt(numFmt: string) {
    return !new Set(["@", "0", "0.00", "0.00%", "$0.00", "\u20AC0.00", "$#,##0.00", "\u20AC#,##0.00"]).has(numFmt);
  }

  function currentCellFormatRequest(): CellFormatRequest {
    if (formatDialogKind.value === "currency") {
      return { kind: "currency", currencyCode: formatDialogCurrency.value };
    }
    if (formatDialogKind.value === "custom") {
      return { kind: "custom", excelNumFmt: formatDialogCustomNumFmt.value };
    }
    return { kind: formatDialogKind.value };
  }

  /** Apply the chosen value and style format to every selected cell. */
  function applySelectedCellFormat() {
    if (!activeWorksheet.value || app.gridReadOnly.value || selectedCells.value.length === 0) {
      formatDialogVisible.value = false;
      return;
    }
    const request = currentCellFormatRequest();
    const range = selectedRange.value;
    const cellsToFormat = cellsForRange(range);
    const locale = app.activeGrid.value?.settings.locale ?? "en-US";
    app.updateGrid((grid) => {
      const worksheet = grid.workbook.worksheetsById[activeWorksheet.value!.id];
      const operations: TeamGridOperation[] = [];
      const patchedCells: Cell[] = [];
      const bounds = boundsForRange(range);
      for (const { cell, rowIndex, columnIndex } of cellsToFormat) {
        const existing = worksheet.cellsById[cell.id] ?? cell;
        const formatted = applyCellFormat(existing, request, locale);
        const stylePatch = formatDialogStylePatchForCell(rowIndex, columnIndex, bounds);
        const nextStyle = applyCellStylePatch(formatted.style, stylePatch);
        const nextCell: Cell = { ...formatted };
        if (nextStyle) {
          nextCell.style = nextStyle;
        } else {
          delete nextCell.style;
        }
        worksheet.cellsById[nextCell.id] = nextCell;
        patchedCells.push(nextCell);
      }
      operations.push({ type: "setCellsStyle", worksheetId: worksheet.id, cells: patchedCells, style: {} });
      return operations;
    });
    formatDialogVisible.value = false;
  }

  /**
   * Apply a partial style patch to every cell in the current selection.
   *
   * The patch is shallow-merged onto each cell so callers can, for example,
   * change only `fontWeight` without losing previously set `backgroundColor`.
   * A `null` patch value removes that cell-level override.
   * A single `setCellsStyle` operation is emitted so the persisted patch is
   * compact and Automerge sees one logical edit per selection.
   */
  function patchSelectedStyle(style: CellStylePatch) {
    patchCellsStyle(selectedRange.value, style);
  }

  function patchCellsStyle(range: CellSelectionRange | null, style: CellStylePatch) {
    if (!activeWorksheet.value || app.gridReadOnly.value) return;
    const cellsToPatch = cellsForRange(range).map(({ cell }) => cell);
    if (cellsToPatch.length === 0) return;
    app.updateGrid((grid) => {
      const worksheet = grid.workbook.worksheetsById[activeWorksheet.value!.id];
      const patchedCells: Cell[] = [];
      for (const cell of cellsToPatch) {
        const existing = worksheet.cellsById[cell.id] ?? cell;
        const nextStyle = applyCellStylePatch(existing.style, style);
        const patchedCell: Cell = { ...existing };
        if (nextStyle) {
          patchedCell.style = nextStyle;
        } else {
          delete patchedCell.style;
        }
        worksheet.cellsById[existing.id] = patchedCell;
        patchedCells.push(patchedCell);
      }
      return [{ type: "setCellsStyle", worksheetId: worksheet.id, cells: patchedCells, style: compactCellStylePatch(style) }];
    });
  }

  function applyCellStylePatch(currentStyle: CellStyle | undefined, patch: CellStylePatch) {
    const nextStyle: CellStyle = { ...currentStyle };
    for (const key of Object.keys(patch) as Array<keyof CellStyle>) {
      const value = patch[key];
      if (value == null) {
        delete nextStyle[key];
      } else if (key === "borders") {
        nextStyle.borders = applyCellBorderPatch(nextStyle.borders, value as CellBorderPatch);
        if (!nextStyle.borders) {
          delete nextStyle.borders;
        }
      } else {
        nextStyle[key] = value as never;
      }
    }
    return Object.keys(nextStyle).length > 0 ? nextStyle : undefined;
  }

  function compactCellStylePatch(patch: CellStylePatch): CellStyle {
    const compact: CellStyle = {};
    for (const key of Object.keys(patch) as Array<keyof CellStyle>) {
      const value = patch[key];
      if (value != null) {
        compact[key] = value as never;
      }
    }
    return compact;
  }

  function applyCellBorderPatch(currentBorders: CellBorders | undefined, patch: CellBorderPatch) {
    const nextBorders: CellBorders = { ...currentBorders };
    for (const side of Object.keys(patch) as CellBorderSide[]) {
      const border = patch[side];
      if (border == null) {
        delete nextBorders[side];
      } else {
        nextBorders[side] = border;
      }
    }
    return Object.keys(nextBorders).length > 0 ? nextBorders : undefined;
  }

  function formatDialogStylePatchForCell(rowIndex: number, columnIndex: number, bounds: RangeBounds | null): CellStylePatch {
    return {
      fontFamily: formatDialogFontFamily.value.trim() || null,
      fontSize: formatDialogFontSize.value || null,
      bold: formatDialogBold.value,
      italic: formatDialogItalic.value,
      underline: formatDialogUnderline.value,
      textColor: formatDialogTextColor.value,
      backgroundColor: formatDialogFillEnabled.value ? formatDialogFillColor.value : null,
      borders: borderPatchForCell(rowIndex, columnIndex, bounds),
    };
  }

  function borderPatchForCell(rowIndex: number, columnIndex: number, bounds: RangeBounds | null): CellBorderPatch | null {
    if (formatDialogBorderPreset.value === "none") {
      return { top: null, right: null, bottom: null, left: null };
    }
    if (formatDialogBorderStyle.value === "none") {
      return formatDialogBorderPreset.value === "custom" ? normalizeBorderPatch(formatDialogBorders.value) : null;
    }
    const border = currentDialogBorder();
    if (formatDialogBorderPreset.value === "outline" && bounds) {
      return {
        top: rowIndex === bounds.minRow ? border : null,
        right: columnIndex === bounds.maxCol ? border : null,
        bottom: rowIndex === bounds.maxRow ? border : null,
        left: columnIndex === bounds.minCol ? border : null,
      };
    }
    if (formatDialogBorderPreset.value === "inside" && bounds) {
      return {
        top: rowIndex > bounds.minRow ? border : null,
        left: columnIndex > bounds.minCol ? border : null,
      };
    }
    if (formatDialogBorderPreset.value === "all") {
      return { top: border, right: border, bottom: border, left: border };
    }
    return normalizeBorderPatch(formatDialogBorders.value);
  }

  function normalizeBorderPatch(borders: CellBorders): CellBorderPatch | null {
    return Object.keys(borders).length > 0 ? { ...borders } : null;
  }

  function currentDialogBorder(): CellBorder {
    return {
      style: formatDialogBorderStyle.value === "none" ? "thin" : formatDialogBorderStyle.value,
      color: formatDialogBorderColor.value,
    };
  }

  function currentDialogBorderCss(border: CellBorder | undefined) {
    if (!border) {
      return undefined;
    }
    const width = border.style === "thick"
      ? "3px"
      : border.style === "medium" || border.style === "double"
        ? "2px"
        : "1px";
    const style = border.style === "dashed" || border.style === "dotted" || border.style === "double"
      ? border.style
      : "solid";
    return `${width} ${style} ${border.color ?? formatDialogBorderColor.value}`;
  }

  function updateCustomBordersFromLineSelection() {
    if (formatDialogBorderPreset.value !== "custom") {
      return;
    }
    const sides = Object.keys(formatDialogBorders.value) as CellBorderSide[];
    if (formatDialogBorderStyle.value === "none") {
      formatDialogBorders.value = {};
      return;
    }
    if (sides.length === 0) {
      formatDialogBorderPreset.value = "outline";
      return;
    }
    const border = currentDialogBorder();
    formatDialogBorders.value = Object.fromEntries(sides.map((side) => [side, border])) as CellBorders;
  }

  function firstBorderStyle(borders: CellBorders | undefined) {
    return Object.values(borders ?? {})[0]?.style;
  }

  function firstBorderColor(borders: CellBorders | undefined) {
    return Object.values(borders ?? {})[0]?.color;
  }

  function toggleFormatDialogBorder(side: CellBorderSide) {
    formatDialogBorderPreset.value = "custom";
    const nextBorders = { ...formatDialogBorders.value };
    if (formatDialogBorderStyle.value === "none" || nextBorders[side]) {
      delete nextBorders[side];
    } else {
      nextBorders[side] = currentDialogBorder();
    }
    formatDialogBorders.value = nextBorders;
  }

  function setFormatDialogBorderPreset(preset: BorderPreset) {
    formatDialogBorderPreset.value = preset;
    if (preset === "none") {
      formatDialogBorders.value = {};
    }
  }

  return {
    formatDialogVisible,
    formatDialogTab,
    formatDialogKind,
    formatDialogCurrency,
    formatDialogCustomNumFmt,
    formatDialogFontFamily,
    formatDialogFontSize,
    formatDialogBold,
    formatDialogItalic,
    formatDialogUnderline,
    formatDialogTextColor,
    formatDialogFillEnabled,
    formatDialogFillColor,
    formatDialogBorderStyle,
    formatDialogBorderColor,
    formatDialogBorderPreset,
    formatDialogBorders,
    openCellFormatDialog,
    applySelectedCellFormat,
    patchSelectedStyle,
    patchCellsStyle,
    updateCustomBordersFromLineSelection,
    toggleFormatDialogBorder,
    setFormatDialogBorderPreset,
    currentDialogBorderCss,
  };
}
