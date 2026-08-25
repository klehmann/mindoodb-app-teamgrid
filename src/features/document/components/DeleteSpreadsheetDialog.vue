<script setup lang="ts">
import { useI18n } from "vue-i18n";

/**
 * Confirmation dialog for File / Delete. Visibility is owned by the parent
 * via `v-model:visible`; confirming emits `confirm` so the parent can wire
 * it to `useTeamGridDocument.deleteCurrentDocument`.
 */
import Button from "primevue/button";
import Dialog from "primevue/dialog";

const { t } = useI18n();
const visible = defineModel<boolean>("visible", { required: true });

const emit = defineEmits<{
  (event: "confirm"): void;
}>();

function confirm() {
  visible.value = false;
  emit("confirm");
}
</script>

<template>
  <Dialog
    v-model:visible="visible"
    modal
    :header="t('app.confirm.deleteSpreadsheetTitle')"
    :style="{ width: '28rem', maxWidth: '96vw' }"
  >
    <p>{{ t('app.confirm.deleteSpreadsheetBody') }}</p>
    <template #footer>
      <Button :label="t('common.cancel')" text @click="visible = false" />
      <Button :label="t('common.delete')" icon="pi pi-trash" severity="danger" @click="confirm" />
    </template>
  </Dialog>
</template>
