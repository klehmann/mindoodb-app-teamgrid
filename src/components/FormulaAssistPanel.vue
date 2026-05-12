<script setup lang="ts">
/**
 * Floating formula content-assist panel.
 *
 * The panel is editor-agnostic: callers pass the active draft, caret position,
 * and the input element to anchor against. This lets the formula bar and the
 * in-cell editor share exactly the same help UI.
 */
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from "vue";

import { getActiveFunctionFragment, getEnclosingFunctionName } from "@/lib/formulas/assist";
import { FUNCTION_REGISTRY, suggestFunctions, type FunctionDefinition } from "@/lib/formulas/functionRegistry";

const props = defineProps<{
  visible: boolean;
  draft: string;
  caretPos: number;
  anchorEl: HTMLElement | null;
  readonly: boolean;
}>();

const emit = defineEmits<{
  "update:visible": [value: boolean];
  select: [definition: FunctionDefinition];
  dismiss: [];
}>();

const panelEl = ref<HTMLElement | null>(null);
const coordinates = ref({ top: 0, left: 0, width: 320 });

const activeFragment = computed(() => getActiveFunctionFragment(props.draft, props.caretPos));
const activeFunctionHelp = computed(() => {
  const name = getEnclosingFunctionName(props.draft, props.caretPos);
  return name ? FUNCTION_REGISTRY[name] ?? null : null;
});
const suggestions = computed(() => suggestFunctions(activeFragment.value));

const panelStyle = computed(() => ({
  top: `${coordinates.value.top}px`,
  left: `${coordinates.value.left}px`,
  minWidth: `${coordinates.value.width}px`,
}));

watch(
  () => [props.visible, props.anchorEl] as const,
  () => {
    if (props.visible) {
      void nextTick(updatePosition);
    }
  },
);

onMounted(() => {
  document.addEventListener("keydown", handleDocumentKeydown, true);
  document.addEventListener("pointerdown", handleDocumentPointerDown, true);
  window.addEventListener("resize", updatePosition);
  window.addEventListener("scroll", updatePosition, true);
});

onBeforeUnmount(() => {
  document.removeEventListener("keydown", handleDocumentKeydown, true);
  document.removeEventListener("pointerdown", handleDocumentPointerDown, true);
  window.removeEventListener("resize", updatePosition);
  window.removeEventListener("scroll", updatePosition, true);
});

function updatePosition() {
  if (!props.anchorEl) {
    return;
  }
  const rect = props.anchorEl.getBoundingClientRect();
  coordinates.value = {
    top: rect.bottom + window.scrollY + 6,
    left: rect.left + window.scrollX,
    width: Math.max(rect.width, 320),
  };
}

function dismiss() {
  emit("update:visible", false);
  emit("dismiss");
}

function choose(definition: FunctionDefinition) {
  if (props.readonly) {
    return;
  }
  emit("select", definition);
}

function handleDocumentKeydown(event: KeyboardEvent) {
  if (!props.visible) {
    return;
  }
  if (event.key === "Escape") {
    event.preventDefault();
    event.stopPropagation();
    dismiss();
  }
}

function handleDocumentPointerDown(event: PointerEvent) {
  if (!props.visible) {
    return;
  }
  const target = event.target;
  if (!(target instanceof Node)) {
    return;
  }
  if (panelEl.value?.contains(target) || props.anchorEl?.contains(target)) {
    return;
  }
  dismiss();
}
</script>

<template>
  <Teleport to="body">
    <div
      v-if="visible"
      ref="panelEl"
      class="formula-assist"
      :style="panelStyle"
      role="dialog"
      aria-label="Formula help"
    >
      <button class="formula-assist__close" type="button" aria-label="Hide formula help" @click="dismiss">
        <i class="pi pi-times" aria-hidden="true" />
      </button>
      <p v-if="activeFunctionHelp" class="formula-assist__help">
        <strong>{{ activeFunctionHelp.signature }}</strong>
        <span>{{ activeFunctionHelp.description }}</span>
      </p>
      <button
        v-for="suggestion in suggestions"
        :key="suggestion.name"
        class="formula-assist__suggestion"
        type="button"
        :disabled="readonly"
        @click="choose(suggestion)"
      >
        <strong>{{ suggestion.name }}</strong>
        <span>{{ suggestion.signature }}</span>
      </button>
    </div>
  </Teleport>
</template>

<style scoped>
.formula-assist {
  position: absolute;
  z-index: 1000;
  display: grid;
  gap: 0.35rem;
  max-width: min(32rem, calc(100vw - 2rem));
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

.formula-assist__help {
  margin: 0;
  display: grid;
  gap: 0.2rem;
  color: var(--muted);
  font-size: 0.85rem;
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
