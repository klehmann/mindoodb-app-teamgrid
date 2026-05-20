<script setup lang="ts">
import Button from "primevue/button";
import Dialog from "primevue/dialog";
import type { useChartPropertiesDialog } from "@/features/charts/composables/useChartPropertiesDialog";

const props = defineProps<{
  controller: ReturnType<typeof useChartPropertiesDialog>;
  readOnly: boolean;
}>();

const {
  chartDialogVisible,
  chartDraftTitle,
  chartDraftType,
  chartDraftLegendPosition,
  chartDraftShowGridlines,
  chartDraftDataRange,
  chartDraftColors,
  chartDraftSeries,
  chartDraftError,
  applyChartProperties,
  removeSelectedChart,
  addSeriesDraft,
  removeSeriesDraft,
} = props.controller;
</script>

<template>
  <Dialog v-model:visible="chartDialogVisible" modal header="Chart properties" :style="{ width: '42rem', maxWidth: '96vw' }">
    <div class="chart-properties-dialog">
      <label class="field">
        Title
        <input v-model="chartDraftTitle" class="native-input" type="text" autocomplete="off" @keyup.enter="applyChartProperties">
      </label>

      <div class="chart-properties-dialog__grid">
        <label class="field">
          Chart type
          <select v-model="chartDraftType" class="native-input">
            <option value="column">Column</option>
            <option value="bar">Bar</option>
            <option value="line">Line</option>
            <option value="pie">Pie</option>
          </select>
        </label>
        <label class="field">
          Legend
          <select v-model="chartDraftLegendPosition" class="native-input">
            <option value="right">Right</option>
            <option value="bottom">Bottom</option>
            <option value="top">Top</option>
            <option value="left">Left</option>
            <option value="none">None</option>
          </select>
        </label>
      </div>

      <label class="field field--inline">
        <input v-model="chartDraftShowGridlines" type="checkbox">
        Show gridlines
      </label>

      <label class="field">
        Chart data range
        <input v-model="chartDraftDataRange" class="native-input" type="text" autocomplete="off" placeholder="Sheet 1!A1:C5">
      </label>

      <section class="chart-properties-dialog__series">
        <div class="chart-properties-dialog__section-header">
          <h3>Series</h3>
          <Button label="Add series" icon="pi pi-plus" size="small" text :disabled="readOnly" @click="addSeriesDraft" />
        </div>
        <div
          v-for="(series, index) in chartDraftSeries"
          :key="series.id"
          class="chart-properties-dialog__series-row"
        >
          <label class="field">
            Name
            <input v-model="series.name" class="native-input" type="text" autocomplete="off" :placeholder="`Series ${index + 1}`">
          </label>
          <label class="field">
            Values
            <input v-model="series.values" class="native-input" type="text" autocomplete="off" placeholder="Sheet 1!B2:B5">
          </label>
          <label class="field">
            Color
            <input v-model="series.color" class="native-input" type="text" autocomplete="off" placeholder="#4472C4">
          </label>
          <Button
            icon="pi pi-times"
            severity="secondary"
            text
            rounded
            :aria-label="`Remove series ${index + 1}`"
            :disabled="readOnly || chartDraftSeries.length <= 1"
            @click="removeSeriesDraft(series.id)"
          />
        </div>
      </section>

      <label class="field">
        Palette colors
        <textarea
          v-model="chartDraftColors"
          class="native-input chart-properties-dialog__textarea"
          rows="3"
          placeholder="#4472C4&#10;#ED7D31&#10;#A5A5A5"
        />
      </label>

      <p v-if="chartDraftError" class="chart-properties-dialog__error">{{ chartDraftError }}</p>
    </div>

    <template #footer>
      <Button label="Delete chart" icon="pi pi-trash" severity="danger" text :disabled="readOnly" @click="removeSelectedChart" />
      <span class="chart-properties-dialog__footer-spacer" />
      <Button label="Cancel" text @click="chartDialogVisible = false" />
      <Button label="Apply" icon="pi pi-check" :disabled="readOnly || chartDraftSeries.length === 0" @click="applyChartProperties" />
    </template>
  </Dialog>
</template>

<style scoped>
.chart-properties-dialog {
  display: flex;
  flex-direction: column;
  gap: 1rem;
}

.chart-properties-dialog__grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 1rem;
}

.field--inline {
  flex-direction: row;
  align-items: center;
  gap: 0.5rem;
}

.chart-properties-dialog__series {
  display: flex;
  flex-direction: column;
  gap: 0.75rem;
}

.chart-properties-dialog__section-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.75rem;
}

.chart-properties-dialog__section-header h3 {
  margin: 0;
  font-size: 0.95rem;
}

.chart-properties-dialog__series-row {
  display: grid;
  grid-template-columns: 1fr 1.3fr 7rem auto;
  gap: 0.75rem;
  align-items: end;
}

.chart-properties-dialog__textarea {
  min-height: 5rem;
  resize: vertical;
}

.chart-properties-dialog__error {
  margin: 0;
  color: #b91c1c;
  font-size: 0.9rem;
}

.chart-properties-dialog__footer-spacer {
  flex: 1 1 auto;
}

@media (max-width: 640px) {
  .chart-properties-dialog__grid,
  .chart-properties-dialog__series-row {
    grid-template-columns: 1fr;
  }
}
</style>
