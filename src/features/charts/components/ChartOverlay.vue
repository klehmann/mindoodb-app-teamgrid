<script setup lang="ts">
import { computed, toRef } from "vue";
import BarChart from "@/features/charts/components/charts/BarChart.vue";
import ColumnChart from "@/features/charts/components/charts/ColumnChart.vue";
import LineChart from "@/features/charts/components/charts/LineChart.vue";
import PieChart from "@/features/charts/components/charts/PieChart.vue";
import { useChartsLayer } from "@/features/charts/composables/useChartsLayer";
import { chartAnchorToRect } from "@/features/charts/lib/chartGeometry";
import { resolveChartData } from "@/features/charts/lib/chartDataResolution";
import type { Chart, Worksheet } from "@/features/document/lib/teamgridDocument";
import type { FormulaContext } from "@/features/formulas/lib";
import type { GridProjection } from "@/features/grid/lib/gridProjection";

const props = defineProps<{
  worksheet: Worksheet;
  projection: GridProjection;
  formulaContext: FormulaContext;
}>();

const worksheetRef = toRef(props, "worksheet");
const charts = useChartsLayer(worksheetRef);
const renderedCharts = computed(() => charts.value.flatMap((chart) => {
  const rect = chartAnchorToRect(chart.anchor, props.worksheet, props.projection);
  if (!rect) {
    return [];
  }
  return [{ chart, rect, data: resolveChartData(chart, props.formulaContext) }];
}));

function chartComponent(chart: Chart) {
  switch (chart.type) {
    case "bar":
      return BarChart;
    case "line":
      return LineChart;
    case "pie":
      return PieChart;
    case "column":
    default:
      return ColumnChart;
  }
}
</script>

<template>
  <div class="chart-overlay" aria-hidden="false">
    <div
      v-for="item in renderedCharts"
      :key="item.chart.id"
      class="chart-overlay__item"
      :style="{
        transform: `translate(${item.rect.x}px, ${item.rect.y}px)`,
        width: `${item.rect.width}px`,
        height: `${item.rect.height}px`,
      }"
    >
      <component
        :is="chartComponent(item.chart)"
        :data="item.data"
        :width="item.rect.width"
        :height="item.rect.height"
      />
    </div>
  </div>
</template>

<style scoped>
.chart-overlay {
  position: absolute;
  inset: 0;
  min-width: max-content;
  pointer-events: none;
  z-index: 5;
}

.chart-overlay__item {
  position: absolute;
  pointer-events: auto;
  border: 1px solid rgb(156 163 175 / 0.7);
  border-radius: 4px;
  box-shadow: 0 6px 18px rgb(15 23 42 / 0.12);
  overflow: hidden;
  background: #fff;
}
</style>
