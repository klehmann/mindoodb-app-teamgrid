<script setup lang="ts">
/**
 * Confirmation dialog for File / Delete. Visibility is owned by the parent
 * via `v-model:visible`; confirming emits `confirm` so the parent can wire
 * it to `useTeamGridDocument.deleteCurrentDocument`.
 */
import Button from "primevue/button";
import Dialog from "primevue/dialog";

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
    header="Delete spreadsheet"
    :style="{ width: '28rem', maxWidth: '96vw' }"
  >
    <p>This deletes the current spreadsheet document from the selected MindooDB database.</p>
    <template #footer>
      <Button label="Cancel" text @click="visible = false" />
      <Button label="Delete" icon="pi pi-trash" severity="danger" @click="confirm" />
    </template>
  </Dialog>
</template>
