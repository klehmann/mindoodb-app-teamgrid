<script setup lang="ts">
/**
 * Tab strip rendered below the grid for switching between worksheets.
 *
 * The component is fully controlled: the parent owns the active worksheet
 * and reacts to events to mutate the workbook through the
 * `useTeamGridDocument` composable.
 *
 * Deleted worksheets are kept in the document's `worksheetOrder` as
 * tombstones so concurrent edits remain merge-safe, but the tab strip
 * only renders live worksheets to keep the UI clean.
 *
 * Interaction:
 * - Click a tab to activate it.
 * - Double-click a tab (or right-click → Rename) to rename it.
 * - Right-click a tab to open a context menu with Rename / Delete.
 *
 * Props:
 * - `grid`: full workbook used to render tabs in `worksheetOrder`.
 * - `activeWorksheetId`: currently visible worksheet, or `null` if none.
 * - `readonly`: disables rename/delete affordances and the add button.
 *
 * Emits:
 * - `select(id)`: user clicked a tab.
 * - `add()`: user clicked the plus button.
 * - `rename(id)`: user requested a rename for a tab.
 * - `delete(id)`: user requested deletion of a tab.
 */
import { computed, ref } from "vue";
import Button from "primevue/button";
import ContextMenu from "primevue/contextmenu";
import type { MenuItem } from "primevue/menuitem";
import type { TeamGridDocumentV1, WorksheetId } from "@/lib/teamgridDocument";

const props = defineProps<{
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

const visibleWorksheetIds = computed(() =>
  props.grid.workbook.worksheetOrder.filter(
    (id) => !props.grid.workbook.worksheetsById[id]?.deletedAt,
  ),
);

const contextMenu = ref<InstanceType<typeof ContextMenu> | null>(null);
const contextWorksheetId = ref<WorksheetId | null>(null);

const contextMenuItems = computed<MenuItem[]>(() => {
  const targetId = contextWorksheetId.value;
  const canDelete = targetId !== null && visibleWorksheetIds.value.length > 1;
  return [
    {
      label: "Rename",
      icon: "pi pi-pencil",
      disabled: props.readonly || targetId === null,
      command: () => {
        if (targetId !== null) {
          emit("rename", targetId);
        }
      },
    },
    {
      label: "Delete",
      icon: "pi pi-trash",
      disabled: props.readonly || !canDelete,
      command: () => {
        if (targetId !== null) {
          emit("delete", targetId);
        }
      },
    },
  ];
});

function openContextMenu(event: MouseEvent, worksheetId: WorksheetId) {
  contextWorksheetId.value = worksheetId;
  contextMenu.value?.show(event);
}
</script>

<template>
  <nav class="worksheet-tabs" aria-label="Worksheets">
    <button
      v-for="worksheetId in visibleWorksheetIds"
      :key="worksheetId"
      class="worksheet-tab"
      :class="{ 'worksheet-tab--active': worksheetId === activeWorksheetId }"
      type="button"
      @click="emit('select', worksheetId)"
      @dblclick="!readonly && emit('rename', worksheetId)"
      @contextmenu.prevent="openContextMenu($event, worksheetId)"
    >
      {{ grid.workbook.worksheetsById[worksheetId]?.title }}
      <span
        v-if="worksheetId === activeWorksheetId && !readonly && visibleWorksheetIds.length > 1"
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
    <ContextMenu ref="contextMenu" :model="contextMenuItems" />
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
