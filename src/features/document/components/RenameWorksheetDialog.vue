<script setup lang="ts">
import { useI18n } from "vue-i18n";

/**
 * Modal-driven rename input for a worksheet tab.
 *
 * Haven's iframe sandbox suppresses `window.prompt`, so we render a proper
 * PrimeVue dialog and let `useWorksheetDialogs` own the draft and the
 * apply action.
 */
import Button from "primevue/button";
import Dialog from "primevue/dialog";
import type { useWorksheetDialogs } from "@/features/document/composables/useWorksheetDialogs";

const { t } = useI18n();
const props = defineProps<{
  controller: ReturnType<typeof useWorksheetDialogs>;
}>();

const {
  renameDialogVisible,
  renameDraft,
  applyWorksheetRename,
} = props.controller;
</script>

<template>
  <Dialog
    v-model:visible="renameDialogVisible"
    modal
    :header="t('app.renameSheet.title')"
    :style="{ width: '24rem', maxWidth: '96vw' }"
  >
    <label class="field">
      {{ t('app.renameSheet.name') }}
      <input
        v-model="renameDraft"
        class="native-input"
        type="text"
        autocomplete="off"
        autofocus
        @keyup.enter="applyWorksheetRename"
      >
    </label>
    <template #footer>
      <Button :label="t('common.cancel')" text @click="renameDialogVisible = false" />
      <Button :label="t('common.rename')" icon="pi pi-check" :disabled="!renameDraft.trim()" @click="applyWorksheetRename" />
    </template>
  </Dialog>
</template>
