<script setup lang="ts">
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
    header="Rename worksheet"
    :style="{ width: '24rem', maxWidth: '96vw' }"
  >
    <label class="field">
      Name
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
      <Button label="Cancel" text @click="renameDialogVisible = false" />
      <Button label="Rename" icon="pi pi-check" :disabled="!renameDraft.trim()" @click="applyWorksheetRename" />
    </template>
  </Dialog>
</template>
