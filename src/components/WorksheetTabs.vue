<script setup lang="ts">
/**
 * Tab strip rendered below the grid for switching between worksheets.
 *
 * The component is fully controlled: the parent owns the active worksheet
 * and reacts to events to mutate the workbook through the
 * `useTeamGridDocument` composable.
 *
 * Props:
 * - `grid`: full workbook used to render tabs in `worksheetOrder`.
 * - `activeWorksheetId`: currently visible worksheet, or `null` if none.
 * - `readonly`: disables rename/delete affordances and the add button.
 *
 * Emits:
 * - `select(id)`: user clicked a tab.
 * - `add()`: user clicked the plus button.
 * - `rename(id)`: user double-clicked a tab to rename it.
 * - `delete(id)`: user requested deletion of a tab.
 */
import Button from "primevue/button";
import type { TeamGridDocumentV1, WorksheetId } from "@/lib/teamgridDocument";

defineProps<{
  grid: TeamGridDocumentV1;
  activeWorksheetId: WorksheetId | null;
  readonly: boolean;
}>();

const emit = defineEmits<{
  select: [worksheetId: WorksheetId];
  add: [];
  rename: [worksheetId: WorksheetId];
  delete: [worksheetId: WorksheetId];
}>();
</script>

<template>
  <nav class="worksheet-tabs" aria-label="Worksheets">
    <button
      v-for="worksheetId in grid.workbook.worksheetOrder"
      :key="worksheetId"
      class="worksheet-tab"
      :class="{ 'worksheet-tab--active': worksheetId === activeWorksheetId }"
      type="button"
      :disabled="Boolean(grid.workbook.worksheetsById[worksheetId]?.deletedAt)"
      @click="emit('select', worksheetId)"
      @dblclick="!readonly && emit('rename', worksheetId)"
    >
      {{ grid.workbook.worksheetsById[worksheetId]?.title ?? "Deleted sheet" }}
      <span v-if="grid.workbook.worksheetsById[worksheetId]?.deletedAt" class="worksheet-tab__deleted">deleted</span>
      <span
        v-else-if="worksheetId === activeWorksheetId && !readonly && grid.workbook.worksheetOrder.length > 1"
        class="worksheet-tab__delete"
        role="button"
        tabindex="0"
        aria-label="Delete worksheet"
        @click.stop="emit('delete', worksheetId)"
        @keyup.enter.stop="emit('delete', worksheetId)"
      >
        x
      </span>
    </button>
    <Button icon="pi pi-plus" text rounded size="small" aria-label="Add worksheet" :disabled="readonly" @click="emit('add')" />
  </nav>
</template>

<style scoped>
.worksheet-tabs {
  display: flex;
  align-items: center;
  gap: 0.35rem;
  padding: 0.5rem;
  overflow-x: auto;
  border-top: 1px solid var(--border);
  background: rgb(255 255 255 / 0.03);
}

.worksheet-tab {
  display: inline-flex;
  align-items: center;
  gap: 0.35rem;
  min-width: 5rem;
  padding: 0.45rem 0.8rem;
  border: 1px solid var(--border);
  border-radius: 999px;
  background: rgb(255 255 255 / 0.04);
  color: var(--text);
  cursor: pointer;
}

.worksheet-tab--active {
  border-color: var(--accent);
  background: rgb(212 160 23 / 0.18);
}

.worksheet-tab__deleted {
  color: var(--muted);
  font-size: 0.75rem;
}

.worksheet-tab__delete {
  display: inline-grid;
  place-items: center;
  width: 1rem;
  height: 1rem;
  border-radius: 999px;
  color: var(--muted);
  font-size: 0.78rem;
}

.worksheet-tab__delete:hover {
  background: rgb(255 255 255 / 0.12);
  color: var(--text);
}
</style>
