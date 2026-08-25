<script setup lang="ts">
import { useI18n } from "vue-i18n";

/**
 * Modal dialog that lists historical revisions of the current document.
 *
 * The dialog is presentational; the parent provides the entries list (loaded
 * via `MindooDBAppDocument.history()`), shows a loading state while the
 * request is in flight, and decides what to do when the user picks an
 * entry. The current document's revision id, when known, is rendered with
 * a special "Current revision" chip.
 *
 * Props:
 * - `visible`: standard `v-model:visible` flag.
 * - `entries`: history entries returned by the SDK.
 * - `loading`: show a loader while history is being fetched.
 * - `errorMessage`: error string returned by the history fetch, if any.
 * - `currentRevisionId`: id of the revision currently open in the editor.
 *
 * Emits:
 * - `update:visible(value)`: standard `v-model:visible` write.
 * - `select(revisionId)`: user confirmed switching to a historical revision.
 * - `cancel()`: user dismissed the dialog without picking a revision.
 */
import { computed, ref, watch } from "vue";
import Button from "primevue/button";
import Dialog from "primevue/dialog";
import type { MindooDBAppDocumentHistoryEntry } from "mindoodb-app-sdk";

const { t } = useI18n();
const props = defineProps<{
  visible: boolean;
  entries: MindooDBAppDocumentHistoryEntry[];
  loading: boolean;
  errorMessage: string | null;
  currentRevisionId: string | null;
}>();

const emit = defineEmits<{
  "update:visible": [value: boolean];
  select: [revisionId: string];
  cancel: [];
}>();

const selectedRevisionId = ref<string | null>(null);
const selectedEntry = computed(() =>
  props.entries.find((entry) => entry.revisionId === selectedRevisionId.value) ?? null);

function formatRevisionDate(timestamp: number) {
  return new Intl.DateTimeFormat(undefined, {
    dateStyle: "medium",
    timeStyle: "short",
  }).format(new Date(timestamp));
}

function authorLabel(entry: MindooDBAppDocumentHistoryEntry) {
  return entry.identityLabel ?? entry.publicKeyFingerprint ?? entry.publicKey;
}

function chipLabel(entry: MindooDBAppDocumentHistoryEntry) {
  if (entry.isCurrent) return t("revisions.current");
  if (entry.isDeleted) return t("revisions.deleted");
  return entry.summary ?? t("revisions.snapshot");
}

function selectEntry(entry: MindooDBAppDocumentHistoryEntry) {
  selectedRevisionId.value = entry.revisionId;
}

function confirmSelection() {
  if (selectedRevisionId.value) {
    emit("select", selectedRevisionId.value);
  }
}

function cancel() {
  emit("cancel");
  emit("update:visible", false);
}

function handleVisibleChange(value: boolean) {
  if (!value) {
    cancel();
    return;
  }
  emit("update:visible", value);
}

watch(
  () => [props.visible, props.entries, props.currentRevisionId] as const,
  ([visible]) => {
    if (!visible) return;
    selectedRevisionId.value = props.currentRevisionId
      ?? props.entries.find((entry) => entry.isCurrent)?.revisionId
      ?? props.entries[0]?.revisionId
      ?? null;
  },
  { immediate: true },
);
</script>

<template>
  <Dialog
    :visible="visible"
    modal
    :header="t('revisions.title')"
    :style="{ width: '38rem', maxWidth: '96vw' }"
    @update:visible="handleVisibleChange"
  >
    <div class="revision-dialog">
      <p class="revision-dialog__intro">{{ t('revisions.intro') }}</p>

      <p v-if="loading" class="revision-dialog__state">{{ t("revisions.loading") }}</p>
      <p v-else-if="errorMessage" class="revision-dialog__state revision-dialog__state--error">{{ errorMessage }}</p>
      <p v-else-if="entries.length === 0" class="revision-dialog__state">{{ t("revisions.empty") }}</p>

      <div v-else class="revision-list" role="listbox" :aria-label="t('revisions.listAria')">
        <button
          v-for="entry in entries"
          :key="entry.revisionId"
          class="revision-row"
          :class="{
            'revision-row--selected': entry.revisionId === selectedRevisionId,
            'revision-row--current': entry.revisionId === currentRevisionId,
          }"
          type="button"
          role="option"
          :aria-selected="entry.revisionId === selectedRevisionId"
          @click="selectEntry(entry)"
          @dblclick="emit('select', entry.revisionId)"
        >
          <span class="revision-row__main">
            <strong>{{ formatRevisionDate(entry.timestamp) }}</strong>
            <small>{{ authorLabel(entry) }}</small>
          </span>
          <span class="revision-row__chip">{{ chipLabel(entry) }}</span>
        </button>
      </div>
    </div>

    <template #footer>
      <Button :label="t('common.cancel')" text @click="cancel" />
      <Button :label="t('revisions.open')" icon="pi pi-history" :disabled="loading || !selectedEntry" @click="confirmSelection" />
    </template>
  </Dialog>
</template>

<style scoped>
.revision-dialog {
  display: grid;
  gap: 0.85rem;
}

.revision-dialog__intro,
.revision-dialog__state {
  margin: 0;
  color: var(--muted);
}

.revision-dialog__state--error {
  color: var(--danger);
}

.revision-list {
  max-height: min(26rem, 62vh);
  overflow: auto;
  display: grid;
  gap: 0.45rem;
}

.revision-row {
  width: 100%;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.75rem;
  padding: 0.75rem 0.85rem;
  border: 1px solid var(--border);
  border-radius: 0.85rem;
  background: rgb(255 255 255 / 0.03);
  color: inherit;
  cursor: pointer;
  font: inherit;
  text-align: left;
}

.revision-row:hover,
.revision-row--selected {
  border-color: var(--accent);
  background: rgb(255 255 255 / 0.07);
}

.revision-row--current {
  box-shadow: inset 3px 0 0 var(--accent);
}

.revision-row__main {
  min-width: 0;
  display: grid;
  gap: 0.2rem;
}

.revision-row__main small {
  overflow: hidden;
  color: var(--muted);
  text-overflow: ellipsis;
  white-space: nowrap;
}

.revision-row__chip {
  flex: 0 0 auto;
  padding: 0.2rem 0.45rem;
  border-radius: 999px;
  background: rgb(255 255 255 / 0.08);
  color: var(--muted);
  font-size: 0.78rem;
}
</style>
