<script setup lang="ts">
/**
 * Excel-style formula bar with content assist for the registered functions.
 *
 * The component is a controlled input: the parent owns the draft string via
 * `v-model:modelValue` and decides what to do when the user commits or
 * cancels. The bar emits an explicit `begin-edit` event the first time the
 * user starts typing, which the parent uses to enter "formula picking" mode
 * (clicks on grid cells append `A1` references into the draft).
 *
 * Props:
 * - `modelValue`: current draft text (with or without leading `=`).
 * - `activeAddress`: address label shown to the left of the input.
 * - `readonly`: disables editing and the commit/cancel buttons.
 * - `errorMessage` (optional): error string surfaced below the bar.
 *
 * Emits:
 * - `update:modelValue(value)`: standard `v-model` write.
 * - `begin-edit()`: first non-trivial edit in the current selection.
 * - `commit(value)`: user pressed Enter or clicked the check icon.
 * - `cancel()`: user pressed Escape or clicked the x icon.
 */
import { nextTick, onMounted, ref, watch } from "vue";
import Button from "primevue/button";
import { insertFunctionAtCaret } from "@/features/formulas/lib/assist";
import type { FunctionDefinition } from "@/features/formulas/lib";

const props = defineProps<{
  modelValue: string;
  activeAddress: string;
  readonly: boolean;
  errorMessage?: string | null;
}>();

const emit = defineEmits<{
  "update:modelValue": [value: string];
  "begin-edit": [];
  commit: [value: string];
  cancel: [];
  "request-help": [payload: { anchorEl: HTMLElement; draft: string; caretPos: number }];
}>();

const draft = ref(props.modelValue);
const inputEl = ref<HTMLTextAreaElement | null>(null);

watch(
  () => props.modelValue,
  (value) => {
    draft.value = value;
    void nextTick(() => autoGrowInput());
  },
);

onMounted(() => {
  autoGrowInput();
});

function autoGrowInput() {
  const el = inputEl.value;
  if (!el) {
    return;
  }
  el.style.height = "auto";
  el.style.height = `${el.scrollHeight}px`;
}

function updateDraft(value: string) {
  draft.value = value;
  emit("begin-edit");
  emit("update:modelValue", draft.value);
  autoGrowInput();
}

async function applyFormulaAssistSuggestion(definition: FunctionDefinition) {
  const caretPos = inputEl.value?.selectionStart ?? draft.value.length;
  const inserted = insertFunctionAtCaret(draft.value, caretPos, definition.name);
  draft.value = inserted.next;
  emit("begin-edit");
  emit("update:modelValue", draft.value);
  await nextTick();
  inputEl.value?.focus();
  inputEl.value?.setSelectionRange(inserted.nextCaret, inserted.nextCaret);
}

function commit() {
  emit("commit", draft.value);
}

function cancel() {
  emit("cancel");
}

function requestHelp() {
  if (props.readonly || !inputEl.value) {
    return;
  }
  emit("begin-edit");
  emit("request-help", {
    anchorEl: inputEl.value,
    draft: draft.value,
    caretPos: inputEl.value.selectionStart ?? draft.value.length,
  });
}

function handleKeydown(event: KeyboardEvent) {
  if ((event.ctrlKey || event.metaKey) && event.code === "Space") {
    event.preventDefault();
    event.stopPropagation();
    requestHelp();
    return;
  }
  if (event.key === "Enter" || event.code === "NumpadEnter" || event.keyCode === 13) {
    // Alt+Enter / Option+Enter inserts a literal newline so the
    // formula bar can build multi-line string cells the same way
    // Excel's formula bar does. Plain Enter still commits.
    if (event.altKey) {
      event.preventDefault();
      event.stopPropagation();
      void insertNewlineAtCaret();
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    commit();
    return;
  }
  if (event.key === "Escape") {
    event.preventDefault();
    event.stopPropagation();
    cancel();
  }
}

async function insertNewlineAtCaret() {
  const el = inputEl.value;
  if (!el) {
    updateDraft(`${draft.value}\n`);
    return;
  }
  const start = el.selectionStart ?? draft.value.length;
  const end = el.selectionEnd ?? draft.value.length;
  const next = `${draft.value.slice(0, start)}\n${draft.value.slice(end)}`;
  updateDraft(next);
  await nextTick();
  inputEl.value?.setSelectionRange(start + 1, start + 1);
}

defineExpose({ applyFormulaAssistSuggestion });
</script>

<template>
  <section class="formula-bar" aria-label="Formula editor">
    <span class="formula-bar__address">{{ activeAddress || "A1" }}</span>
    <textarea
      ref="inputEl"
      v-model="draft"
      class="formula-bar__input"
      rows="1"
      wrap="soft"
      spellcheck="false"
      :readonly="readonly"
      placeholder="Enter a value or formula, for example =SUM(A1:B4)"
      @focus="emit('begin-edit')"
      @input="updateDraft(($event.target as HTMLTextAreaElement).value)"
      @keydown="handleKeydown"
    />
    <div class="formula-bar__actions">
      <Button label="fx" text rounded severity="secondary" aria-label="Show formula help" :disabled="readonly" @click="requestHelp" />
      <Button icon="pi pi-times" text rounded severity="secondary" aria-label="Cancel editing" :disabled="readonly" @click="cancel" />
      <Button icon="pi pi-check" text rounded aria-label="Apply formula" :disabled="readonly" @click="commit" />
    </div>
    <p v-if="errorMessage" class="formula-bar__error">{{ errorMessage }}</p>
  </section>
</template>

<style scoped>
.formula-bar {
  position: relative;
  display: grid;
  grid-template-columns: minmax(4rem, auto) 1fr auto;
  gap: 0.5rem;
  align-items: center;
  padding: 0.6rem;
  border-bottom: 1px solid var(--border);
  background: rgb(255 255 255 / 0.04);
}

.formula-bar__address {
  padding: 0.45rem 0.65rem;
  border: 1px solid var(--border);
  border-radius: 0.6rem;
  color: var(--muted);
  font-family: var(--font-code);
  font-size: 0.85rem;
}

.formula-bar__input {
  width: 100%;
  min-height: 2.2rem;
  padding: 0.55rem 0.7rem;
  border: 1px solid var(--border);
  border-radius: 0.65rem;
  background: rgb(255 255 255 / 0.06);
  color: var(--text);
  /*
   * The bar is a `<textarea>` so it can hold the multi-line content
   * Alt+Enter produces, but visually we want it to feel like a normal
   * single-row input that simply grows downward when wrapping is
   * needed. `field-sizing: content` covers Chrome/Edge automatically
   * and the `autoGrowInput` helper handles Safari/Firefox.
   */
  field-sizing: content;
  resize: none;
  overflow-y: auto;
  font: inherit;
  line-height: 1.3;
  white-space: pre-wrap;
  word-break: break-word;
}

.formula-bar__actions {
  display: inline-flex;
  align-items: center;
  gap: 0.15rem;
}

.formula-bar__error {
  grid-column: 2 / 4;
  margin: -0.2rem 0 0;
  color: var(--danger);
  font-size: 0.82rem;
}
</style>
