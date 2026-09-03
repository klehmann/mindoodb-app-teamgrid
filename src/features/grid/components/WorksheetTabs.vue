<script setup lang="ts">
import { useI18n } from "vue-i18n";

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
 * - Right-click a tab to open a context menu with Rename / Move / Delete.
 * - Drag a tab sideways to reorder it.
 *
 * Reordering is a mouse or pen drag only. On touch the same gesture is how the
 * strip is scrolled when the tabs overflow, so taking it over would cost more
 * than it gives; the context menu's move entries are the touch path, and they
 * double as the keyboard one.
 *
 * Props:
 * - `grid`: full workbook used to render tabs in `worksheetOrder`.
 * - `activeWorksheetId`: currently visible worksheet, or `null` if none.
 * - `readonly`: disables rename/delete affordances and the add button.
 *
 * Emits:
 * - `select(id)`: user clicked a tab.
 * - `add()`: user picked Add Sheet from the plus menu.
 * - `addView()`: user picked Add Virtual View Sheet from the plus menu.
 * - `rename(id)`: user requested a rename for a tab.
 * - `configureView(id)`: user requested View Sheet settings.
 * - `delete(id)`: user requested deletion of a tab.
 * - `move(id, toIndex)`: user finished dragging a tab to a new place.
 * - `nudge(id, offset)`: user moved a tab one place with menu or keyboard.
 */
import { computed, ref } from "vue";
import Button from "primevue/button";
import ContextMenu from "primevue/contextmenu";
import type { MenuItem } from "primevue/menuitem";
import type { TeamGridDocumentV1, WorksheetId } from "@/features/document/lib/teamgridDocument";
import { planWorksheetMove, resolveWorksheetOrder } from "@/features/document/lib/worksheetOrder";

const { t } = useI18n();
const props = defineProps<{
  grid: TeamGridDocumentV1;
  activeWorksheetId: WorksheetId | null;
  readonly: boolean;
}>();

const emit = defineEmits<{
  select: [worksheetId: WorksheetId];
  add: [];
  "add-view": [];
  rename: [worksheetId: WorksheetId];
  "configure-view": [worksheetId: WorksheetId];
  delete: [worksheetId: WorksheetId];
  move: [worksheetId: WorksheetId, toIndex: number];
  nudge: [worksheetId: WorksheetId, offset: -1 | 1];
}>();

/** Minimum sideways travel before a press on a tab becomes a reorder. */
const DRAG_THRESHOLD_PX = 6;

/**
 * The order the tabs follow while a drag is in flight.
 *
 * A drag reorders this list on its own so the strip under the pointer is
 * already the answer, and reports the one move it added up to on release. The
 * document is left alone until then, which keeps a drag across several tabs a
 * single edit rather than one per tab it passed.
 */
const dragOrder = ref<WorksheetId[] | null>(null);

const worksheetOrder = computed(() => dragOrder.value ?? resolveWorksheetOrder(props.grid.workbook));

const visibleWorksheetIds = computed(() =>
  worksheetOrder.value.filter((id) => !props.grid.workbook.worksheetsById[id]?.deletedAt),
);

const contextMenu = ref<InstanceType<typeof ContextMenu> | null>(null);
const addMenu = ref<InstanceType<typeof ContextMenu> | null>(null);
const contextWorksheetId = ref<WorksheetId | null>(null);
const tabsRef = ref<HTMLElement | null>(null);
const draggedWorksheetId = ref<WorksheetId | null>(null);

const drag = {
  pointerId: null as number | null,
  worksheetId: null as WorksheetId | null,
  startX: 0,
  moved: false,
  /** Set by a finished reorder, consumed by the click that follows it. */
  swallowClick: false,
};

function canMove(worksheetId: WorksheetId | null, offset: -1 | 1) {
  if (props.readonly || worksheetId === null) {
    return false;
  }
  const index = visibleWorksheetIds.value.indexOf(worksheetId);
  return index >= 0 && index + offset >= 0 && index + offset < visibleWorksheetIds.value.length;
}

const contextMenuItems = computed<MenuItem[]>(() => {
  const targetId = contextWorksheetId.value;
  const canDelete = targetId !== null && visibleWorksheetIds.value.length > 1;
  return [
    {
      label: t("grid.sheets.rename"),
      icon: "pi pi-pencil",
      disabled: props.readonly || targetId === null,
      command: () => {
        if (targetId !== null) {
          emit("rename", targetId);
        }
      },
    },
    {
      label: t("grid.sheets.viewSettings"),
      icon: "pi pi-table",
      disabled: props.readonly || targetId === null || !props.grid.workbook.worksheetsById[targetId]?.viewBinding,
      command: () => {
        if (targetId !== null) {
          emit("configure-view", targetId);
        }
      },
    },
    {
      label: t("grid.sheets.moveLeft"),
      icon: "pi pi-arrow-left",
      disabled: !canMove(targetId, -1),
      command: () => {
        if (targetId !== null) {
          emit("nudge", targetId, -1);
        }
      },
    },
    {
      label: t("grid.sheets.moveRight"),
      icon: "pi pi-arrow-right",
      disabled: !canMove(targetId, 1),
      command: () => {
        if (targetId !== null) {
          emit("nudge", targetId, 1);
        }
      },
    },
    {
      label: t("common.delete"),
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

const addMenuItems = computed<MenuItem[]>(() => [
  {
    label: t("grid.sheets.addSheet"),
    icon: "pi pi-plus",
    disabled: props.readonly,
    command: () => emit("add"),
  },
  {
    label: t("grid.sheets.addViewSheet"),
    icon: "pi pi-table",
    disabled: props.readonly,
    command: () => emit("add-view"),
  },
]);

function openContextMenu(event: MouseEvent, worksheetId: WorksheetId) {
  contextWorksheetId.value = worksheetId;
  contextMenu.value?.show(event);
}

function openAddMenu(event: MouseEvent) {
  addMenu.value?.show(event);
}

function startDrag(worksheetId: WorksheetId, event: PointerEvent) {
  if (props.readonly || event.pointerType === "touch" || event.button !== 0) {
    return;
  }
  drag.pointerId = event.pointerId;
  drag.worksheetId = worksheetId;
  drag.startX = event.clientX;
  drag.moved = false;
  // A previous reorder whose click never arrived must not eat this one.
  drag.swallowClick = false;
}

function moveDrag(event: PointerEvent) {
  if (drag.pointerId !== event.pointerId || drag.worksheetId === null) {
    return;
  }
  if (!drag.moved) {
    if (Math.abs(event.clientX - drag.startX) < DRAG_THRESHOLD_PX) {
      return;
    }
    drag.moved = true;
    draggedWorksheetId.value = drag.worksheetId;
    dragOrder.value = [...worksheetOrder.value];
    // Held on the tab rather than the element under the pointer, so the moves
    // keep coming once the neighbours have shuffled out from under it.
    if (event.currentTarget instanceof HTMLElement) {
      try {
        event.currentTarget.setPointerCapture?.(event.pointerId);
      } catch {
        // Capture is a nicety; the reorder still works from the move events.
      }
    }
  }
  const target = worksheetUnderPointer(event.clientX);
  if (!target || target === drag.worksheetId) {
    return;
  }
  const plan = planWorksheetMove(worksheetOrder.value, drag.worksheetId, target);
  if (plan) {
    dragOrder.value = plan.order;
  }
}

function endDrag(event: PointerEvent) {
  if (drag.pointerId !== event.pointerId) {
    return;
  }
  const worksheetId = drag.worksheetId;
  const dropped = drag.moved ? dragOrder.value : null;
  drag.pointerId = null;
  drag.worksheetId = null;
  drag.moved = false;
  draggedWorksheetId.value = null;
  dragOrder.value = null;
  if (!dropped || worksheetId === null) {
    return;
  }
  // The tab the pointer started on is the one that moved, so it is the one to
  // activate. Without pointer capture the click would land on whichever tab
  // has since slid under the pointer, hence swallowing it rather than letting
  // both fire.
  drag.swallowClick = true;
  emit("select", worksheetId);
  const toIndex = dropped.indexOf(worksheetId);
  if (toIndex >= 0) {
    emit("move", worksheetId, toIndex);
  }
}

function worksheetUnderPointer(clientX: number): WorksheetId | null {
  for (const tab of tabsRef.value?.querySelectorAll<HTMLElement>("[data-worksheet-id]") ?? []) {
    const rect = tab.getBoundingClientRect();
    if (clientX >= rect.left && clientX <= rect.right) {
      return (tab.dataset.worksheetId ?? null) as WorksheetId | null;
    }
  }
  return null;
}

/** A click that ended a reorder already picked its tab in {@link endDrag}. */
function handleClick(worksheetId: WorksheetId) {
  if (drag.swallowClick) {
    drag.swallowClick = false;
    return;
  }
  emit("select", worksheetId);
}

function handleKeydown(worksheetId: WorksheetId, event: KeyboardEvent) {
  const offset = event.key === "ArrowLeft" ? -1 : event.key === "ArrowRight" ? 1 : null;
  if (offset === null || !event.altKey || !canMove(worksheetId, offset)) {
    return;
  }
  event.preventDefault();
  emit("nudge", worksheetId, offset);
}
</script>

<template>
  <nav ref="tabsRef" class="worksheet-tabs" :aria-label="t('grid.sheets.aria')">
    <button
      v-for="worksheetId in visibleWorksheetIds"
      :key="worksheetId"
      :data-worksheet-id="worksheetId"
      class="worksheet-tab"
      :class="{
        'worksheet-tab--active': worksheetId === activeWorksheetId,
        'worksheet-tab--dragging': worksheetId === draggedWorksheetId,
      }"
      type="button"
      @click="handleClick(worksheetId)"
      @dblclick="!readonly && emit('rename', worksheetId)"
      @contextmenu.prevent="openContextMenu($event, worksheetId)"
      @keydown="handleKeydown(worksheetId, $event)"
      @pointerdown="startDrag(worksheetId, $event)"
      @pointermove="moveDrag"
      @pointerup="endDrag"
      @pointercancel="endDrag"
    >
      {{ grid.workbook.worksheetsById[worksheetId]?.title }}
      <span
        v-if="worksheetId === activeWorksheetId && !readonly && visibleWorksheetIds.length > 1"
        class="worksheet-tab__delete"
        role="button"
        tabindex="0"
        :aria-label="t('grid.sheets.delete')"
        @click.stop="emit('delete', worksheetId)"
        @keyup.enter.stop="emit('delete', worksheetId)"
      >
        x
      </span>
    </button>
    <Button
      icon="pi pi-plus"
      text
      rounded
      size="small"
      :aria-label="t('grid.sheets.add')"
      :disabled="readonly"
      @click="openAddMenu"
    />
    <ContextMenu ref="contextMenu" :model="contextMenuItems" />
    <ContextMenu ref="addMenu" :model="addMenuItems" />
  </nav>
</template>

<style scoped>
.worksheet-tabs {
  flex: 0 0 auto;
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

.worksheet-tab--dragging {
  opacity: 0.65;
  cursor: grabbing;
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
