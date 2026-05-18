/**
 * Bridge native clipboard events on the grid viewport up to the parent.
 *
 * The actual serialization / deserialization of the OS clipboard lives in
 * `useGridClipboard`; this composable only intercepts the browser's
 * `copy` / `cut` / `paste` events on the grid element, calls
 * `preventDefault` so the parent can write its own payloads, and forwards
 * them with the current selection.
 */
import type { Ref } from "vue";
import type { CellSelectionRange } from "@/features/grid/composables/useSelection";

export interface UseGridClipboardBridgeOptions {
  readonly: Readonly<Ref<boolean>>;
  selectedRange: Readonly<Ref<CellSelectionRange | null>>;
  editingCellId: Readonly<Ref<string | null>>;
  onCopy: (payload: { range: CellSelectionRange | null; event: ClipboardEvent }) => void;
  onCut: (payload: { range: CellSelectionRange | null; event: ClipboardEvent }) => void;
  onPaste: (payload: { event: ClipboardEvent }) => void;
  onClear: () => void;
}

export function useGridClipboardBridge(options: UseGridClipboardBridgeOptions) {
  function handleViewportCopy(event: ClipboardEvent) {
    if (isTypingInAnotherEditor(event.target)) {
      return;
    }
    event.preventDefault();
    options.onCopy({ range: options.selectedRange.value, event });
  }

  function handleViewportCut(event: ClipboardEvent) {
    if (isTypingInAnotherEditor(event.target) || options.readonly.value) {
      return;
    }
    event.preventDefault();
    options.onCut({ range: options.selectedRange.value, event });
  }

  function handleViewportPaste(event: ClipboardEvent) {
    if (isTypingInAnotherEditor(event.target) || options.readonly.value) {
      return;
    }
    event.preventDefault();
    options.onPaste({ event });
  }

  /**
   * Local Escape handler used to dismiss the clipboard marquee. The
   * Escape for cancelling an inline edit is handled by the editor before
   * this fires, so we only emit `clear` when no edit is active.
   */
  function handleViewportKeydown(event: KeyboardEvent) {
    if (event.key === "Escape" && !options.editingCellId.value) {
      options.onClear();
    }
  }

  return {
    handleViewportCopy,
    handleViewportCut,
    handleViewportPaste,
    handleViewportKeydown,
  };
}

/**
 * True if the keyboard event originated inside another editable element
 * (formula bar input, modal dialog field, contenteditable). Used to stop
 * clipboard interception when the focus is in a non-grid editor.
 */
function isTypingInAnotherEditor(target: EventTarget | null) {
  if (!(target instanceof HTMLElement)) {
    return false;
  }
  return target.matches("input, textarea, select, [contenteditable='true']");
}
