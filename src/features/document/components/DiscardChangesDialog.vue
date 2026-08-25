<script setup lang="ts">
import { useI18n } from "vue-i18n";

/**
 * Confirmation dialog shown before closing a dirty spreadsheet session.
 * Confirming intentionally discards local pending operations for that window.
 */
import Button from "primevue/button";
import Dialog from "primevue/dialog";

const { t } = useI18n();
defineProps<{
  title: string;
}>();

const visible = defineModel<boolean>("visible", { required: true });

const emit = defineEmits<{
  (event: "discard"): void;
}>();

function discard() {
  visible.value = false;
  emit("discard");
}
</script>

<template>
  <Dialog
    v-model:visible="visible"
    modal
    :header="t('app.confirm.discardTitle')"
    :style="{ width: '28rem', maxWidth: '96vw' }"
  >
    <p>{{ t('app.confirm.discardBody', { title }) }}</p>
    <template #footer>
      <Button :label="t('common.keepEditing')" text @click="visible = false" />
      <Button :label="t('app.confirm.discardConfirm')" icon="pi pi-times" severity="danger" @click="discard" />
    </template>
  </Dialog>
</template>
