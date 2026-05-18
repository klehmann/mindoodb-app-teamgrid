/**
 * Pointer-driven column and row resize for the grid.
 *
 * Window-level listeners keep the drag alive even when the cursor moves
 * outside the resize handle (or outside the grid entirely). A live
 * `currentSize` is rendered during the drag so the user sees feedback
 * without writing to the document until they release the pointer.
 */
import { onBeforeUnmount, ref, type Ref } from "vue";
import { DEFAULT_ROW_HEIGHT } from "@/shared/lib/gridDimensions";
import type { ColumnId, RowId } from "@/features/document/lib/teamgridDocument";

/** Resize handles cannot drag a column narrower than this. */
export const MIN_COLUMN_WIDTH = 48;
/** Resize handles cannot drag a row shorter than this. */
export const MIN_ROW_HEIGHT = 24;

/** State for an in-progress resize drag of a single column or row. */
interface ResizeDrag {
  kind: "column" | "row";
  id: ColumnId | RowId;
  /** Pointer position (clientX for columns, clientY for rows) at drag start. */
  startPointer: number;
  /** Size in pixels at drag start; used as the baseline for the delta. */
  startSize: number;
  /** Live size during the drag (rendered while the user holds the handle). */
  currentSize: number;
}

export interface UseColumnRowResizeOptions {
  readonly: Readonly<Ref<boolean>>;
  onColumnResize: (payload: { columnId: ColumnId; width: number }) => void;
  onRowResize: (payload: { rowId: RowId; height: number }) => void;
}

export function useColumnRowResize(options: UseColumnRowResizeOptions) {
  const activeResize = ref<ResizeDrag | null>(null);

  function startColumnResize(event: PointerEvent, columnId: ColumnId, width: number) {
    startResize(event, {
      kind: "column",
      id: columnId,
      startPointer: event.clientX,
      startSize: width,
      currentSize: width,
    });
  }

  function startRowResize(event: PointerEvent, rowId: RowId, height?: number) {
    const startSize = height ?? DEFAULT_ROW_HEIGHT;
    startResize(event, {
      kind: "row",
      id: rowId,
      startPointer: event.clientY,
      startSize,
      currentSize: startSize,
    });
  }

  /**
   * Begin a resize drag. We `preventDefault` and `stopPropagation` so the
   * pointer-down does not also bubble up and trigger row/column selection
   * on the underlying header.
   */
  function startResize(event: PointerEvent, drag: ResizeDrag) {
    if (options.readonly.value || event.button !== 0) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    activeResize.value = drag;
    window.addEventListener("pointermove", handleResizePointerMove);
    window.addEventListener("pointerup", finishResize);
    window.addEventListener("pointercancel", cancelResize);
  }

  /**
   * Update the live `currentSize` in response to pointer motion. Clamping
   * to a minimum size keeps users from collapsing a column/row into a
   * state where the resize handle becomes unreachable.
   */
  function handleResizePointerMove(event: PointerEvent) {
    const drag = activeResize.value;
    if (!drag) {
      return;
    }
    const pointer = drag.kind === "column" ? event.clientX : event.clientY;
    const minSize = drag.kind === "column" ? MIN_COLUMN_WIDTH : MIN_ROW_HEIGHT;
    drag.currentSize = Math.max(minSize, Math.round(drag.startSize + pointer - drag.startPointer));
  }

  /**
   * End-of-drag commit. Emit a single `resize-*` event only when the size
   * actually changed so a stray click on the handle does not push a no-op
   * operation through the document save path.
   */
  function finishResize() {
    const drag = activeResize.value;
    if (!drag) {
      return;
    }
    if (drag.currentSize !== drag.startSize) {
      if (drag.kind === "column") {
        options.onColumnResize({ columnId: drag.id as ColumnId, width: drag.currentSize });
      } else {
        options.onRowResize({ rowId: drag.id as RowId, height: drag.currentSize });
      }
    }
    activeResize.value = null;
    removeResizeListeners();
  }

  /** Pointer was cancelled (e.g. browser focus loss); discard any preview. */
  function cancelResize() {
    activeResize.value = null;
    removeResizeListeners();
  }

  function removeResizeListeners() {
    window.removeEventListener("pointermove", handleResizePointerMove);
    window.removeEventListener("pointerup", finishResize);
    window.removeEventListener("pointercancel", cancelResize);
  }

  /**
   * Effective rendered width of a column, accounting for the live preview
   * size during an in-progress drag.
   */
  function columnPixelWidth(column: { id: ColumnId; width: number }) {
    return activeResize.value?.kind === "column" && activeResize.value.id === column.id
      ? activeResize.value.currentSize
      : column.width;
  }

  /** Mirror of {@link columnPixelWidth} for rows; falls back to {@link DEFAULT_ROW_HEIGHT}. */
  function rowPixelHeight(row: { id: RowId; height?: number }) {
    return activeResize.value?.kind === "row" && activeResize.value.id === row.id
      ? activeResize.value.currentSize
      : row.height ?? DEFAULT_ROW_HEIGHT;
  }

  onBeforeUnmount(removeResizeListeners);

  return {
    activeResize,
    startColumnResize,
    startRowResize,
    columnPixelWidth,
    rowPixelHeight,
  };
}
