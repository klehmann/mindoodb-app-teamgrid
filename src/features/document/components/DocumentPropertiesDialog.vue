<script setup lang="ts">
/**
 * Title, tags, and locale editor for the current spreadsheet document.
 *
 * Edits go through `useDocumentPropertiesDialog`, which batches them into a
 * single `setDocumentProperties` operation so they participate in the
 * granular-save machinery rather than going around it.
 */
import Button from "primevue/button";
import Dialog from "primevue/dialog";
import type { useDocumentPropertiesDialog } from "@/features/document/composables/useDocumentPropertiesDialog";

const props = defineProps<{
  controller: ReturnType<typeof useDocumentPropertiesDialog>;
  readOnly: boolean;
}>();

const {
  propertiesDialogVisible,
  propertiesTitleDraft,
  propertiesTagsDraft,
  propertiesIsTemplateDraft,
  propertiesLocaleDraft,
  propertiesLocaleOptions,
  applyDocumentProperties,
  resetPropertiesDraft,
} = props.controller;

function cancel() {
  propertiesDialogVisible.value = false;
  resetPropertiesDraft();
}
</script>

<template>
  <Dialog
    v-model:visible="propertiesDialogVisible"
    modal
    header="Spreadsheet properties"
    :style="{ width: '34rem', maxWidth: '96vw' }"
  >
    <div class="properties-dialog">
      <label class="field">
        Title
        <input v-model="propertiesTitleDraft" class="native-input" type="text" autocomplete="off">
      </label>
      <label class="field">
        Tags
        <textarea v-model="propertiesTagsDraft" class="native-input native-input--textarea" rows="6" placeholder="Work\Planning&#10;Finance" />
      </label>
      <p class="properties-dialog__hint">Enter one tag per line. Use a backslash to create hierarchy, for example <code>Work\Planning</code>.</p>
      <label class="field">
        Locale
        <select v-model="propertiesLocaleDraft" class="native-input">
          <option v-for="option in propertiesLocaleOptions" :key="option.value" :value="option.value">{{ option.label }}</option>
        </select>
      </label>
      <label class="properties-dialog__checkbox">
        <input v-model="propertiesIsTemplateDraft" type="checkbox">
        Use this spreadsheet as a template
      </label>
    </div>
    <template #footer>
      <Button label="Cancel" text @click="cancel" />
      <Button label="Apply" icon="pi pi-check" :disabled="readOnly" @click="applyDocumentProperties" />
    </template>
  </Dialog>
</template>

<style scoped>
.properties-dialog {
  display: grid;
  gap: 1rem;
}

.properties-dialog__hint {
  margin: 0;
  color: var(--muted);
  line-height: 1.5;
}

.properties-dialog__checkbox {
  display: inline-flex;
  align-items: center;
  gap: 0.5rem;
  color: var(--text);
}

.properties-dialog code {
  color: var(--formula-text);
}
</style>
