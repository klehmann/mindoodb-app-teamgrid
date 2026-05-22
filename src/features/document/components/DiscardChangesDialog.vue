<script setup lang="ts">
/**
 * Confirmation dialog shown before closing a dirty spreadsheet session.
 * Confirming intentionally discards local pending operations for that window.
 */
import Button from "primevue/button";
import Dialog from "primevue/dialog";

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
    header="Discard changes?"
    :style="{ width: '28rem', maxWidth: '96vw' }"
  >
    <p>{{ title }} has unsaved changes. Close it and discard those changes?</p>
    <template #footer>
      <Button label="Keep editing" text @click="visible = false" />
      <Button label="Discard changes" icon="pi pi-times" severity="danger" @click="discard" />
    </template>
  </Dialog>
</template>
