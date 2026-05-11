<script setup lang="ts">
import { computed, ref, watch } from "vue";
import Button from "primevue/button";
import { FUNCTION_REGISTRY, suggestFunctions, type FunctionDefinition } from "@/lib/formulas";

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
}>();

const draft = ref(props.modelValue);
const assistDismissed = ref(false);

const activeFunctionQuery = computed(() => {
  const match = /(?:^|[=(,+\-*/])\s*([A-Za-z_]*)$/.exec(draft.value);
  return match ? match[1] : "";
});

const suggestions = computed(() => {
  if (!draft.value.trim().startsWith("=")) {
    return [];
  }
  return suggestFunctions(activeFunctionQuery.value).slice(0, 6);
});

const activeFunctionHelp = computed(() => {
  const match = /([A-Za-z_][A-Za-z0-9_]*)\([^()]*$/.exec(draft.value);
  if (!match) {
    return null;
  }
  return FUNCTION_REGISTRY[match[1].toUpperCase()] ?? null;
});

const showAssist = computed(() =>
  !assistDismissed.value
  && !props.readonly
  && (suggestions.value.length > 0 || Boolean(activeFunctionHelp.value) || Boolean(props.errorMessage)));

watch(
  () => props.modelValue,
  (value) => {
    draft.value = value;
  },
);

function updateDraft(value: string) {
  draft.value = value;
  assistDismissed.value = false;
  emit("begin-edit");
  emit("update:modelValue", draft.value);
}

function chooseSuggestion(definition: FunctionDefinition) {
  const prefixMatch = /^(.*?)([A-Za-z_]*)$/.exec(draft.value);
  const prefix = prefixMatch?.[1] ?? draft.value;
  draft.value = `${prefix}${definition.name}(`;
  assistDismissed.value = false;
  emit("begin-edit");
  emit("update:modelValue", draft.value);
}

function commit() {
  assistDismissed.value = true;
  emit("commit", draft.value);
}

function cancel() {
  assistDismissed.value = true;
  emit("cancel");
}
</script>

<template>
  <section class="formula-bar" aria-label="Formula editor">
    <span class="formula-bar__address">{{ activeAddress || "A1" }}</span>
    <input
      v-model="draft"
      class="formula-bar__input"
      type="text"
      :readonly="readonly"
      placeholder="Enter a value or formula, for example =SUM(A1:B4)"
      @focus="emit('begin-edit')"
      @input="updateDraft(($event.target as HTMLInputElement).value)"
      @keydown.enter.prevent="commit"
      @keydown.escape.prevent="cancel"
    >
    <div class="formula-bar__actions">
      <Button icon="pi pi-times" text rounded severity="secondary" aria-label="Cancel editing" :disabled="readonly" @click="cancel" />
      <Button icon="pi pi-check" text rounded aria-label="Apply formula" :disabled="readonly" @click="commit" />
    </div>
    <div v-if="showAssist" class="formula-assist">
      <button class="formula-assist__close" type="button" aria-label="Hide formula help" @click="assistDismissed = true">
        <i class="pi pi-times" aria-hidden="true" />
      </button>
      <p v-if="activeFunctionHelp" class="formula-assist__help">
        <strong>{{ activeFunctionHelp.signature }}</strong>
        <span>{{ activeFunctionHelp.description }}</span>
      </p>
      <p v-if="errorMessage" class="formula-assist__error">{{ errorMessage }}</p>
      <button
        v-for="suggestion in suggestions"
        :key="suggestion.name"
        class="formula-assist__suggestion"
        type="button"
        :disabled="readonly"
        @click="chooseSuggestion(suggestion)"
      >
        <strong>{{ suggestion.name }}</strong>
        <span>{{ suggestion.signature }}</span>
      </button>
    </div>
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
  padding: 0.55rem 0.7rem;
  border: 1px solid var(--border);
  border-radius: 0.65rem;
  background: rgb(255 255 255 / 0.06);
  color: var(--text);
}

.formula-bar__actions {
  display: inline-flex;
  align-items: center;
  gap: 0.15rem;
}

.formula-assist {
  position: absolute;
  z-index: 4;
  top: calc(100% - 0.2rem);
  left: 5.6rem;
  right: 3rem;
  display: grid;
  gap: 0.35rem;
  padding: 0.75rem 2.35rem 0.75rem 0.75rem;
  border: 1px solid var(--border);
  border-radius: 0.8rem;
  background: var(--bg-elevated);
  box-shadow: var(--shadow);
}

.formula-assist__close {
  position: absolute;
  top: 0.5rem;
  right: 0.5rem;
  display: inline-grid;
  place-items: center;
  width: 1.8rem;
  height: 1.8rem;
  border: 0;
  border-radius: 999px;
  background: transparent;
  color: var(--muted);
  cursor: pointer;
}

.formula-assist__close:hover {
  background: rgb(255 255 255 / 0.08);
  color: var(--text);
}

.formula-assist__help,
.formula-assist__error {
  margin: 0;
  display: grid;
  gap: 0.2rem;
  color: var(--muted);
  font-size: 0.85rem;
}

.formula-assist__error {
  color: var(--danger);
}

.formula-assist__suggestion {
  display: grid;
  gap: 0.15rem;
  padding: 0.45rem 0.55rem;
  border: 0;
  border-radius: 0.55rem;
  background: transparent;
  color: var(--text);
  cursor: pointer;
  text-align: left;
}

.formula-assist__suggestion:hover {
  background: rgb(255 255 255 / 0.07);
}

.formula-assist__suggestion span {
  color: var(--muted);
  font-size: 0.82rem;
}
</style>
