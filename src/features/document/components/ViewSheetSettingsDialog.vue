<script setup lang="ts">
import Button from "primevue/button";
import Dialog from "primevue/dialog";
import type { useWorksheetDialogs } from "@/features/document/composables/useWorksheetDialogs";

const props = defineProps<{
  controller: ReturnType<typeof useWorksheetDialogs>;
}>();

const {
  viewSheetDialogVisible,
  viewSheetTargetId,
  viewSheetNameDraft,
  viewSheetViewIdDraft,
  viewSheetShowDocuments,
  viewSheetShowCategories,
  viewSheetRootCategoryPathDraft,
  viewSheetErrorMessage,
  viewSheetSaving,
  configuredViews,
  applyViewSheetSettings,
} = props.controller;
</script>

<template>
  <Dialog
    v-model:visible="viewSheetDialogVisible"
    modal
    :header="viewSheetTargetId ? 'Virtual View Sheet settings' : 'Add Virtual View Sheet'"
    :style="{ width: '32rem', maxWidth: '96vw' }"
  >
    <div class="view-sheet-settings">
      <label class="field">
        Sheet name
        <input
          v-model="viewSheetNameDraft"
          class="native-input"
          type="text"
          autocomplete="off"
          autofocus
          @keyup.enter="applyViewSheetSettings"
        >
      </label>

      <label class="field">
        View
        <select v-model="viewSheetViewIdDraft" class="native-input">
          <option value="" disabled>Select a view</option>
          <option v-for="view in configuredViews" :key="view.id" :value="view.id">
            {{ view.description || view.id }}
          </option>
        </select>
      </label>

      <div class="view-sheet-settings__checks" aria-label="Included view rows">
        <label>
          <input v-model="viewSheetShowDocuments" type="checkbox">
          Show documents
        </label>
        <label>
          <input v-model="viewSheetShowCategories" type="checkbox">
          Show categories
        </label>
      </div>

      <label class="field">
        Top level category
        <input
          v-model="viewSheetRootCategoryPathDraft"
          class="native-input"
          type="text"
          autocomplete="off"
          placeholder="somecategory\\sublevel1"
        >
      </label>

      <p v-if="viewSheetErrorMessage" class="view-sheet-settings__error">{{ viewSheetErrorMessage }}</p>
      <p class="view-sheet-settings__hint">
        A virtual view sheet can read data from virtual views configured for this MindooDB application in the application properties.
      </p>
    </div>

    <template #footer>
      <Button label="Cancel" text :disabled="viewSheetSaving" @click="viewSheetDialogVisible = false" />
      <Button
        :label="viewSheetTargetId ? 'Apply and refresh' : 'Create'"
        icon="pi pi-check"
        :loading="viewSheetSaving"
        :disabled="!viewSheetNameDraft.trim() || !viewSheetViewIdDraft || (!viewSheetShowDocuments && !viewSheetShowCategories)"
        @click="applyViewSheetSettings"
      />
    </template>
  </Dialog>
</template>

<style scoped>
.view-sheet-settings {
  display: grid;
  gap: 1rem;
}

.view-sheet-settings__checks {
  display: flex;
  flex-wrap: wrap;
  gap: 1rem;
  color: var(--text);
}

.view-sheet-settings__checks label {
  display: inline-flex;
  align-items: center;
  gap: 0.4rem;
}

.view-sheet-settings__error {
  margin: 0;
  color: var(--danger);
}

.view-sheet-settings__hint {
  margin: 0;
  color: var(--muted);
  font-size: 0.9rem;
  line-height: 1.35;
}
</style>
