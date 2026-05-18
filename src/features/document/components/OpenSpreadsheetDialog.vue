<script setup lang="ts">
/**
 * File / Open dialog backed by a dynamic view navigator.
 *
 * All navigator orchestration (build, refresh on database change, dispose
 * on close) lives in `useOpenDialog`. The dialog itself only wires the
 * controller's state to PrimeVue widgets and the category tree.
 */
import type { Ref } from "vue";
import Button from "primevue/button";
import Dialog from "primevue/dialog";
import type { MindooDBAppDatabaseInfo } from "mindoodb-app-sdk";
import TagTreeList from "@/features/document/components/TagTreeList.vue";
import type { useOpenDialog } from "@/features/document/composables/useOpenDialog";

const props = defineProps<{
  controller: ReturnType<typeof useOpenDialog>;
  selectedDatabaseId: Ref<string>;
  readableDatabases: Ref<MindooDBAppDatabaseInfo[]>;
}>();

const {
  openDialogVisible,
  selectedOpenDocId,
  selectedOpenCategoryKey,
  openCategoryNodes,
  openDialogDocuments,
  handleOpenDatabaseChange,
  selectOpenCategory,
  openSelectedDocument,
  disposeOpenNavigator,
} = props.controller;
</script>

<template>
  <Dialog
    v-model:visible="openDialogVisible"
    modal
    header="Open spreadsheet"
    :style="{ width: '58rem', maxWidth: '96vw' }"
    @hide="disposeOpenNavigator"
  >
    <div class="open-dialog">
      <label class="field">
        Database
        <select v-model="selectedDatabaseId.value" class="native-input" @change="handleOpenDatabaseChange">
          <option v-for="database in readableDatabases.value" :key="database.id" :value="database.id">{{ database.title || database.id }}</option>
        </select>
      </label>
      <div class="open-dialog__browser">
        <aside class="open-dialog__tree" aria-label="Spreadsheet tags">
          <TagTreeList
            :nodes="openCategoryNodes"
            :selected-key="selectedOpenCategoryKey"
            @select="selectOpenCategory"
          />
        </aside>
        <div class="document-list">
          <button
            v-for="document in openDialogDocuments"
            :key="document.id"
            class="document-row"
            :class="{ 'document-row--selected': document.id === selectedOpenDocId }"
            type="button"
            @click="selectedOpenDocId = document.id"
            @dblclick="openSelectedDocument"
          >
            <strong>{{ document.title }}</strong>
            <small>{{ document.detail }}</small>
            <small>{{ document.id }}</small>
          </button>
          <p v-if="openDialogDocuments.length === 0" class="document-list__empty">No spreadsheets in this category.</p>
        </div>
      </div>
    </div>
    <template #footer>
      <Button label="Cancel" text @click="openDialogVisible = false" />
      <Button label="Open" icon="pi pi-folder-open" :disabled="!selectedOpenDocId" @click="openSelectedDocument" />
    </template>
  </Dialog>
</template>

<style scoped>
.open-dialog {
  display: grid;
  gap: 1rem;
}

.open-dialog__browser {
  min-height: 22rem;
  display: grid;
  grid-template-columns: minmax(12rem, 0.8fr) minmax(0, 1.4fr);
  gap: 0.8rem;
}

.open-dialog__tree {
  max-height: 24rem;
  overflow: auto;
  padding: 0.35rem;
  border: 1px solid var(--border);
  border-radius: 0.85rem;
  background: rgb(255 255 255 / 0.025);
}

.document-list {
  max-height: 22rem;
  overflow: auto;
  display: grid;
  align-content: start;
  gap: 0.45rem;
}

.document-row {
  display: grid;
  gap: 0.2rem;
  padding: 0.75rem 0.85rem;
  border: 1px solid var(--border);
  border-radius: 0.85rem;
  background: rgb(255 255 255 / 0.03);
  color: inherit;
  cursor: pointer;
  text-align: left;
}

.document-row--selected,
.document-row:hover {
  border-color: var(--accent);
  background: rgb(255 255 255 / 0.07);
}

.document-row small {
  color: var(--muted);
}

.document-list__empty {
  margin: 0;
  color: var(--muted);
  line-height: 1.5;
}

@media (max-width: 720px) {
  .open-dialog__browser {
    grid-template-columns: 1fr;
  }
}
</style>
