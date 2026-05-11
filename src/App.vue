<script setup lang="ts">
import { computed, ref, watch } from "vue";
import Button from "primevue/button";
import Dialog from "primevue/dialog";
import Menubar from "primevue/menubar";
import type { MenuItem } from "primevue/menuitem";

import DocumentRevisionDialog from "@/components/DocumentRevisionDialog.vue";
import FormulaBar from "@/components/FormulaBar.vue";
import GridViewport from "@/components/GridViewport.vue";
import WorksheetTabs from "@/components/WorksheetTabs.vue";
import { useTeamGridDocument, readDocumentSummaryLabel } from "@/composables/useTeamGridDocument";
import { coerceInputToCellValue, formatCellValue, formulaResultToCellValue } from "@/lib/cellFormatting";
import { evaluateFormula, parseFormula } from "@/lib/formulas";
import { createCellId, createId, getFirstVisibleWorksheet, type Cell, type CellStyle, type WorksheetId } from "@/lib/teamgridDocument";
import { projectWorksheet } from "@/lib/gridProjection";

const app = useTeamGridDocument();

const openDialogVisible = ref(false);
const deleteDialogVisible = ref(false);
const revisionDialogVisible = ref(false);
const selectedOpenDocId = ref("");
const activeWorksheetId = ref<WorksheetId | null>(null);
const selectedCellId = ref<string | null>(null);
const selectedCellAddress = ref("");
const formulaDraft = ref("");
const formulaError = ref<string | null>(null);
const formulaEditing = ref(false);

const activeWorksheet = computed(() => {
  if (!app.activeGrid.value) {
    return null;
  }
  const explicit = activeWorksheetId.value
    ? app.activeGrid.value.workbook.worksheetsById[activeWorksheetId.value]
    : null;
  return explicit && !explicit.deletedAt ? explicit : getFirstVisibleWorksheet(app.activeGrid.value);
});

const projection = computed(() => activeWorksheet.value ? projectWorksheet(activeWorksheet.value) : null);

const selectedCell = computed(() => {
  if (!activeWorksheet.value || !selectedCellId.value) {
    return null;
  }
  return activeWorksheet.value.cellsById[selectedCellId.value] ?? null;
});

const statusBadgeLabel = computed(() => {
  if (app.isViewingHistorical.value) {
    return "Historical · read-only";
  }
  if (app.isTimeTravelActive.value) {
    return `Time travel · ${app.timeTravelDateLabel.value}`;
  }
  return `Current · ${app.isDirty.value ? "Unsaved" : "Saved"}`;
});

const highlightedCellIds = computed(() => {
  if (!activeWorksheet.value || !projection.value || !formulaDraft.value.trim().startsWith("=")) {
    return [];
  }
  const parsed = parseFormula(formulaDraft.value, activeWorksheet.value.id, projection.value);
  if ("code" in parsed) {
    return [];
  }
  return parsed.references.flatMap((reference) => {
    if (reference.kind === "cell") {
      return [createCellId(reference.rowId, reference.columnId)];
    }
    if (reference.kind === "range") {
      const startRow = projection.value?.rowIndexById.get(reference.startRowId);
      const endRow = projection.value?.rowIndexById.get(reference.endRowId);
      const startColumn = projection.value?.columnIndexById.get(reference.startColumnId);
      const endColumn = projection.value?.columnIndexById.get(reference.endColumnId);
      if (startRow == null || endRow == null || startColumn == null || endColumn == null || !projection.value) {
        return [];
      }
      const cellIds: string[] = [];
      for (let rowIndex = Math.min(startRow, endRow); rowIndex <= Math.max(startRow, endRow); rowIndex += 1) {
        for (let columnIndex = Math.min(startColumn, endColumn); columnIndex <= Math.max(startColumn, endColumn); columnIndex += 1) {
          const row = projection.value.rows[rowIndex];
          const column = projection.value.columns[columnIndex];
          if (row && column) {
            cellIds.push(createCellId(row.id, column.id));
          }
        }
      }
      return cellIds;
    }
    return [];
  });
});

const menuItems = computed<MenuItem[]>(() => [
  {
    label: "File",
    items: [
      { label: "New", icon: "pi pi-file-plus", disabled: !app.canCreate.value, command: () => void app.createNewDocument() },
      { label: "Open", icon: "pi pi-folder-open", command: () => void openFileDialog() },
      { separator: true },
      { label: "Save", icon: "pi pi-save", disabled: !app.canSave.value, command: () => void app.saveDocument() },
      { label: "Delete", icon: "pi pi-trash", disabled: !app.canDelete.value, command: () => { deleteDialogVisible.value = true; } },
    ],
  },
  {
    label: "Edit",
    items: [
      { label: "Insert row below", icon: "pi pi-arrow-down", disabled: app.gridReadOnly.value || !selectedCell.value, command: () => insertRow("after") },
      { label: "Insert column right", icon: "pi pi-arrow-right", disabled: app.gridReadOnly.value || !selectedCell.value, command: () => insertColumn("after") },
      { separator: true },
      { label: "Delete row", icon: "pi pi-minus", disabled: app.gridReadOnly.value || !selectedCell.value, command: deleteSelectedRow },
      { label: "Delete column", icon: "pi pi-minus", disabled: app.gridReadOnly.value || !selectedCell.value, command: deleteSelectedColumn },
    ],
  },
  {
    label: "Format",
    items: [
      { label: "Bold", icon: "pi pi-bold", disabled: app.gridReadOnly.value || !selectedCell.value, command: () => patchSelectedStyle({ bold: !selectedCell.value?.style?.bold }) },
      { label: "Italic", icon: "pi pi-italic", disabled: app.gridReadOnly.value || !selectedCell.value, command: () => patchSelectedStyle({ italic: !selectedCell.value?.style?.italic }) },
      { label: "Underline", icon: "pi pi-underline", disabled: app.gridReadOnly.value || !selectedCell.value, command: () => patchSelectedStyle({ underline: !selectedCell.value?.style?.underline }) },
    ],
  },
  {
    label: "History",
    items: [
      { label: "Browse revisions", icon: "pi pi-history", disabled: !app.currentCanBrowseHistory.value || !app.currentDocument.value, command: () => void openRevisionDialog() },
      { label: "Return to current", icon: "pi pi-refresh", disabled: !app.isViewingHistorical.value, command: app.returnToCurrent },
    ],
  },
]);

watch(
  () => activeWorksheet.value?.id,
  (worksheetId) => {
    activeWorksheetId.value = worksheetId ?? null;
  },
);

watch(
  selectedCell,
  (cell) => {
    if (!cell) {
      formulaDraft.value = "";
      return;
    }
    formulaDraft.value = cell.formula?.source ?? formatCellValue(cell.value, app.activeGrid.value?.settings.locale);
  },
);

async function openFileDialog() {
  await app.refreshDocuments();
  selectedOpenDocId.value = app.documents.value[0]?.id ?? "";
  openDialogVisible.value = true;
}

async function openSelectedDocument() {
  if (!selectedOpenDocId.value) {
    return;
  }
  await app.openDocument(selectedOpenDocId.value);
  openDialogVisible.value = false;
}

async function openRevisionDialog() {
  await app.openRevisionDialog();
  revisionDialogVisible.value = true;
}

function handleRevisionSelect(revisionId: string) {
  revisionDialogVisible.value = false;
  void app.loadHistoricalRevision(revisionId);
}

function selectCell(cell: Cell, address: string) {
  if (
    formulaEditing.value
    && !app.gridReadOnly.value
    && selectedCellId.value
    && selectedCellId.value !== cell.id
    && formulaDraft.value.trim().startsWith("=")
  ) {
    formulaDraft.value = appendPickedAddress(formulaDraft.value, address);
    return;
  }
  formulaEditing.value = false;
  selectedCellId.value = cell.id;
  selectedCellAddress.value = address;
}

function appendPickedAddress(source: string, address: string) {
  if (/[\w)]$/.test(source.trimEnd())) {
    return `${source}+${address}`;
  }
  return `${source}${address}`;
}

function commitFormulaBar(value: string) {
  if (!selectedCell.value) {
    return;
  }
  formulaEditing.value = false;
  commitCell(selectedCell.value, value);
}

function cancelFormulaEdit() {
  formulaEditing.value = false;
  formulaError.value = null;
  formulaDraft.value = selectedCell.value?.formula?.source
    ?? (selectedCell.value ? formatCellValue(selectedCell.value.value, app.activeGrid.value?.settings.locale) : "");
}

function commitCell(cell: Cell, rawValue: string) {
  if (!activeWorksheet.value || !projection.value) {
    return;
  }
  app.updateGrid((grid) => {
    const worksheet = grid.workbook.worksheetsById[activeWorksheet.value!.id];
    const targetCell: Cell = {
      ...cell,
      value: coerceInputToCellValue(rawValue, worksheet.columnsById[cell.columnId]?.defaultValueKind),
      formula: undefined,
    };
    if (rawValue.trim().startsWith("=")) {
      const evaluated = evaluateFormula(rawValue, worksheet, projectWorksheet(worksheet));
      targetCell.formula = {
        kind: "formula",
        source: rawValue,
        references: evaluated.references,
        cached: evaluated.result,
        error: evaluated.result.kind === "error" ? evaluated.result.code : undefined,
      };
      targetCell.value = formulaResultToCellValue(evaluated.result);
      formulaError.value = evaluated.errorMessage ?? null;
    } else {
      formulaError.value = null;
    }
    worksheet.cellsById[targetCell.id] = targetCell;
  });
}

function addWorksheet() {
  app.updateGrid((grid) => {
    const worksheetId = createId("sheet");
    const rowOrder = Array.from({ length: 24 }, () => createId("row"));
    const columnOrder = Array.from({ length: 12 }, () => createId("col"));
    grid.workbook.worksheetOrder.push(worksheetId);
    grid.workbook.worksheetsById[worksheetId] = {
      id: worksheetId,
      title: `Sheet ${grid.workbook.worksheetOrder.length}`,
      rowOrder,
      columnOrder,
      rowsById: Object.fromEntries(rowOrder.map((id) => [id, { id }])),
      columnsById: Object.fromEntries(columnOrder.map((id) => [id, { id, width: 120 }])),
      cellsById: {},
    };
    activeWorksheetId.value = worksheetId;
  });
}

function renameWorksheet(worksheetId: WorksheetId) {
  const nextTitle = window.prompt("Worksheet name", app.activeGrid.value?.workbook.worksheetsById[worksheetId]?.title ?? "");
  if (!nextTitle?.trim()) {
    return;
  }
  app.updateGrid((grid) => {
    grid.workbook.worksheetsById[worksheetId].title = nextTitle.trim();
  });
}

function deleteWorksheet(worksheetId: WorksheetId) {
  app.updateGrid((grid) => {
    grid.workbook.worksheetsById[worksheetId].deletedAt = new Date().toISOString();
    activeWorksheetId.value = getFirstVisibleWorksheet(grid)?.id ?? null;
  });
}

function insertRow(position: "before" | "after") {
  if (!activeWorksheet.value || !selectedCell.value) return;
  app.updateGrid((grid) => {
    const worksheet = grid.workbook.worksheetsById[activeWorksheet.value!.id];
    const rowId = createId("row");
    const selectedIndex = worksheet.rowOrder.indexOf(selectedCell.value!.rowId);
    worksheet.rowsById[rowId] = { id: rowId };
    worksheet.rowOrder.splice(position === "before" ? selectedIndex : selectedIndex + 1, 0, rowId);
  });
}

function insertColumn(position: "before" | "after") {
  if (!activeWorksheet.value || !selectedCell.value) return;
  app.updateGrid((grid) => {
    const worksheet = grid.workbook.worksheetsById[activeWorksheet.value!.id];
    const columnId = createId("col");
    const selectedIndex = worksheet.columnOrder.indexOf(selectedCell.value!.columnId);
    worksheet.columnsById[columnId] = { id: columnId, width: 120 };
    worksheet.columnOrder.splice(position === "before" ? selectedIndex : selectedIndex + 1, 0, columnId);
  });
}

function deleteSelectedRow() {
  if (!activeWorksheet.value || !selectedCell.value) return;
  app.updateGrid((grid) => {
    grid.workbook.worksheetsById[activeWorksheet.value!.id].rowsById[selectedCell.value!.rowId].deletedAt = new Date().toISOString();
  });
}

function deleteSelectedColumn() {
  if (!activeWorksheet.value || !selectedCell.value) return;
  app.updateGrid((grid) => {
    grid.workbook.worksheetsById[activeWorksheet.value!.id].columnsById[selectedCell.value!.columnId].deletedAt = new Date().toISOString();
  });
}

function patchSelectedStyle(style: CellStyle) {
  if (!selectedCell.value) return;
  app.updateGrid((grid) => {
    const worksheet = grid.workbook.worksheetsById[activeWorksheet.value!.id];
    const existing = worksheet.cellsById[selectedCell.value!.id] ?? selectedCell.value!;
    worksheet.cellsById[existing.id] = {
      ...existing,
      style: {
        ...existing.style,
        ...style,
      },
    };
  });
}
</script>

<template>
  <main class="teamgrid-shell">
    <header class="toolbar glass-card" :class="{ 'toolbar--ios-multitasking': app.hostUiPreferences.value.iosMultitaskingOptimized }">
      <div class="toolbar__leading">
        <span class="toolbar__title">Teamgrid</span>
        <Menubar :model="menuItems" class="toolbar__menubar" />
        <Button
          :icon="app.isViewingHistorical.value ? 'pi pi-history' : 'pi pi-refresh'"
          text
          rounded
          severity="secondary"
          :aria-label="app.isViewingHistorical.value ? 'Return to current version' : 'Refresh spreadsheet'"
          :disabled="!app.canRefresh.value"
          @click="app.isViewingHistorical.value ? app.returnToCurrent() : app.refreshCurrentDocument()"
        />
      </div>
      <button
        v-if="app.currentCanBrowseHistory.value && app.currentDocument.value"
        class="toolbar__status-badge"
        type="button"
        @click="openRevisionDialog"
      >
        {{ statusBadgeLabel }}
      </button>
      <span v-else class="toolbar__status-badge">{{ statusBadgeLabel }}</span>
    </header>

    <section class="workspace glass-card">
      <template v-if="app.activeGrid.value && activeWorksheet && projection">
        <div v-if="app.isTimeTravelActive.value" class="history-banner">
          <i class="pi pi-clock" aria-hidden="true" />
          <span>Time travel mode is active as of {{ app.timeTravelDateLabel.value }} - read-only.</span>
        </div>
        <div v-if="app.isViewingHistorical.value" class="history-banner">
          <i class="pi pi-history" aria-hidden="true" />
          <span>You're viewing a historical revision - read-only.</span>
          <button type="button" @click="app.returnToCurrent">Return to current</button>
        </div>

        <div class="format-toolbar">
          <label>
            Text
            <input type="color" :disabled="app.gridReadOnly.value || !selectedCell" :value="selectedCell?.style?.textColor ?? '#eef2ff'" @input="patchSelectedStyle({ textColor: ($event.target as HTMLInputElement).value })">
          </label>
          <label>
            Fill
            <input type="color" :disabled="app.gridReadOnly.value || !selectedCell" :value="selectedCell?.style?.backgroundColor ?? '#111827'" @input="patchSelectedStyle({ backgroundColor: ($event.target as HTMLInputElement).value })">
          </label>
          <label>
            Font size
            <input class="format-toolbar__number" type="number" min="8" max="48" :disabled="app.gridReadOnly.value || !selectedCell" :value="selectedCell?.style?.fontSize ?? 14" @change="patchSelectedStyle({ fontSize: Number(($event.target as HTMLInputElement).value) })">
          </label>
        </div>

        <FormulaBar
          v-model="formulaDraft"
          :active-address="selectedCellAddress"
          :readonly="app.gridReadOnly.value || !selectedCell"
          :error-message="formulaError"
          @begin-edit="formulaEditing = true"
          @commit="commitFormulaBar"
          @cancel="cancelFormulaEdit"
        />
        <GridViewport
          :worksheet="activeWorksheet"
          :projection="projection"
          :selected-cell-id="selectedCellId"
          :highlighted-cell-ids="highlightedCellIds"
          :readonly="app.gridReadOnly.value"
          :locale="app.activeGrid.value.settings.locale"
          @select="selectCell"
          @commit="commitCell"
        />
        <WorksheetTabs
          :grid="app.activeGrid.value"
          :active-worksheet-id="activeWorksheet.id"
          :readonly="app.gridReadOnly.value"
          @select="activeWorksheetId = $event"
          @add="addWorksheet"
          @rename="renameWorksheet"
          @delete="deleteWorksheet"
        />
      </template>
      <section v-else class="empty-state">
        <h1>Collaborative spreadsheets for Haven</h1>
        <p>Create or open a Teamgrid spreadsheet. Each file is stored as one MindooDB Automerge document with stable rows, columns, worksheet tabs, and formulas.</p>
        <div class="empty-state__actions">
          <Button label="New spreadsheet" icon="pi pi-file-plus" :disabled="!app.canCreate.value" @click="app.createNewDocument" />
          <Button label="Open spreadsheet" icon="pi pi-folder-open" severity="secondary" @click="openFileDialog" />
        </div>
      </section>
    </section>

    <footer class="status-line">{{ app.status.value }}</footer>

    <Dialog v-model:visible="openDialogVisible" modal header="Open spreadsheet" :style="{ width: '34rem', maxWidth: '96vw' }">
      <div class="open-dialog">
        <label class="field">
          Database
          <select v-model="app.selectedDatabaseId.value" class="native-input" @change="app.refreshDocuments">
            <option v-for="database in app.readableDatabases.value" :key="database.id" :value="database.id">{{ database.title || database.id }}</option>
          </select>
        </label>
        <div class="document-list">
          <button
            v-for="document in app.documents.value"
            :key="document.id"
            class="document-row"
            :class="{ 'document-row--selected': document.id === selectedOpenDocId }"
            type="button"
            @click="selectedOpenDocId = document.id"
            @dblclick="openSelectedDocument"
          >
            <strong>{{ readDocumentSummaryLabel(document) }}</strong>
            <small>{{ document.updatedAt ?? document.id }}</small>
          </button>
        </div>
      </div>
      <template #footer>
        <Button label="Cancel" text @click="openDialogVisible = false" />
        <Button label="Open" icon="pi pi-folder-open" :disabled="!selectedOpenDocId" @click="openSelectedDocument" />
      </template>
    </Dialog>

    <Dialog v-model:visible="deleteDialogVisible" modal header="Delete spreadsheet" :style="{ width: '28rem', maxWidth: '96vw' }">
      <p>This deletes the current spreadsheet document from the selected MindooDB database.</p>
      <template #footer>
        <Button label="Cancel" text @click="deleteDialogVisible = false" />
        <Button label="Delete" icon="pi pi-trash" severity="danger" @click="deleteDialogVisible = false; app.deleteCurrentDocument()" />
      </template>
    </Dialog>

    <DocumentRevisionDialog
      v-model:visible="revisionDialogVisible"
      :entries="app.revisionEntries.value"
      :loading="app.revisionLoading.value"
      :error-message="app.revisionErrorMessage.value"
      :current-revision-id="app.currentRevisionId.value"
      @select="handleRevisionSelect"
      @cancel="revisionDialogVisible = false"
    />
  </main>
</template>
