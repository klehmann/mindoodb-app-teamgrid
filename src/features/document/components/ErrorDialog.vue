<script setup lang="ts">
import { useI18n } from "vue-i18n";

/**
 * Generic error dialog that surfaces unexpected failures from the document
 * composable to the user. The visible state and message are owned by
 * `useErrorDialog`; this component only renders the modal and forwards the
 * dismissal action.
 */
import Button from "primevue/button";
import Dialog from "primevue/dialog";
import type { useErrorDialog } from "@/features/document/composables/useErrorDialog";

const { t } = useI18n();
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
    :header="t('app.error.title')"
    :style="{ width: '32rem', maxWidth: '96vw' }"
    @hide="onHide"
  >
    <p>{{ message }}</p>
    <template #footer>
      <Button :label="t('common.ok')" icon="pi pi-check" @click="dismissErrorDialog" />
    </template>
  </Dialog>
</template>
