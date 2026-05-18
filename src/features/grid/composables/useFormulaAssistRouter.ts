/**
 * Coordinate the floating formula-assist panel across the two editors that
 * can request it: the formula bar and the inline cell editor.
 *
 * The component-side editors only need to know that the panel might be
 * routed elsewhere; this composable owns the routing decision and exposes
 * the imperative `applyFormulaAssistSuggestion` method on the active editor
 * via its `Ref` so the parent never has to special-case the call site.
 */
import { ref, type Ref } from "vue";
import type { FunctionDefinition } from "@/features/formulas/lib";

export type FormulaAssistEditor = "formulaBar" | "inlineCell";

export interface FormulaAssistRequest {
  anchorEl: HTMLElement;
  draft: string;
  caretPos: number;
}

export interface FormulaAssistTarget {
  applyFormulaAssistSuggestion: (definition: FunctionDefinition) => void;
}

export interface UseFormulaAssistRouterOptions {
  formulaBarTarget: Ref<FormulaAssistTarget | null>;
  gridViewportTarget: Ref<FormulaAssistTarget | null>;
}

export function useFormulaAssistRouter(options: UseFormulaAssistRouterOptions) {
  const formulaAssistOpen = ref(false);
  const formulaAssistEditor = ref<FormulaAssistEditor>("formulaBar");
  const formulaAssistAnchor = ref<HTMLElement | null>(null);
  const formulaAssistDraft = ref("");
  const formulaAssistCaretPos = ref(0);
  const inlineCellEditing = ref(false);
  const inlineCellDraft = ref("");

  function openFormulaAssist(editor: FormulaAssistEditor, request: FormulaAssistRequest) {
    formulaAssistEditor.value = editor;
    formulaAssistAnchor.value = request.anchorEl;
    formulaAssistDraft.value = request.draft;
    formulaAssistCaretPos.value = request.caretPos;
    formulaAssistOpen.value = true;
  }

  function handleFormulaAssistSelect(definition: FunctionDefinition) {
    const target = formulaAssistEditor.value === "formulaBar"
      ? options.formulaBarTarget.value
      : options.gridViewportTarget.value;
    target?.applyFormulaAssistSuggestion(definition);
    formulaAssistOpen.value = false;
  }

  function handleInlineEditState(payload: { editing: boolean; draft: string }) {
    inlineCellEditing.value = payload.editing;
    inlineCellDraft.value = payload.draft;
    if (!payload.editing && formulaAssistEditor.value === "inlineCell") {
      formulaAssistOpen.value = false;
    }
  }

  return {
    formulaAssistOpen,
    formulaAssistEditor,
    formulaAssistAnchor,
    formulaAssistDraft,
    formulaAssistCaretPos,
    inlineCellEditing,
    inlineCellDraft,
    openFormulaAssist,
    handleFormulaAssistSelect,
    handleInlineEditState,
  };
}
