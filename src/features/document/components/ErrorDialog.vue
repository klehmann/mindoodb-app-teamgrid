<script setup lang="ts">
/**
 * Generic error dialog that surfaces unexpected failures from the document
 * composable to the user. The visible state and message are owned by
 * `useErrorDialog`; this component only renders the modal and forwards the
 * dismissal action.
 */
import Button from "primevue/button";
import Dialog from "primevue/dialog";
import type { useErrorDialog } from "@/features/document/composables/useErrorDialog";

const props = defineProps<{
  controller: ReturnType<typeof useErrorDialog>;
  message: string | null;
  onHide: () => void;
}>();

const { errorDialogVisible, dismissErrorDialog } = props.controller;
</script>

<template>
  <Dialog
    v-model:visible="errorDialogVisible"
    modal
    header="Something went wrong"
    :style="{ width: '32rem', maxWidth: '96vw' }"
    @hide="onHide"
  >
    <p>{{ message }}</p>
    <template #footer>
      <Button label="OK" icon="pi pi-check" @click="dismissErrorDialog" />
    </template>
  </Dialog>
</template>
