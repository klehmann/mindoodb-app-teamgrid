<script setup lang="ts">
import { computed, ref, toRef } from "vue";
import BarChart from "@/features/charts/components/charts/BarChart.vue";
import ColumnChart from "@/features/charts/components/charts/ColumnChart.vue";
import LineChart from "@/features/charts/components/charts/LineChart.vue";
import PieChart from "@/features/charts/components/charts/PieChart.vue";
import { useChartsLayer } from "@/features/charts/composables/useChartsLayer";
import {
  chartAnchorToRect,
  rectToChartAnchor,
  type ChartLayoutOrigin,
  type ChartRect,
} from "@/features/charts/lib/chartGeometry";
import { resolveChartData } from "@/features/charts/lib/chartDataResolution";
import type { Chart, ChartId, TwoCellAnchor, Worksheet } from "@/features/document/lib/teamgridDocument";
import type { FormulaContext } from "@/features/formulas/lib";
import type { GridProjection } from "@/features/grid/lib/gridProjection";

const props = withDefaults(defineProps<{
  worksheet: Worksheet;
  projection: GridProjection;
  formulaContext: FormulaContext;
  selectedChartId: ChartId | null;
  readonly: boolean;
  /** `body` when the overlay lives inside the scrolling cell region (below column headers). */
  layoutOrigin?: ChartLayoutOrigin;
}>(), {
  layoutOrigin: "body",
});

const emit = defineEmits<{
  select: [chartId: ChartId | null];
  edit: [chartId: ChartId];
  "chart-context": [payload: { event: MouseEvent; chartId: ChartId }];
  "resize-chart": [payload: { chartId: ChartId; anchor: TwoCellAnchor }];
  "delete-chart": [chartId: ChartId];
}>();

const worksheetRef = toRef(props, "worksheet");
const charts = useChartsLayer(worksheetRef);
const resizing = ref<{ chartId: ChartId; rect: ChartRect } | null>(null);
const renderedCharts = computed(() => charts.value.flatMap((chart) => {
  const rect = chartAnchorToRect(chart.anchor, props.worksheet, props.projection, props.layoutOrigin);
  if (!rect) {
    return [];
  }
  return [{
    chart,
    rect: resizing.value?.chartId === chart.id ? resizing.value.rect : rect,
    data: resolveChartData(chart, props.formulaContext),
  }];
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

function selectChart(event: MouseEvent, chartId: ChartId) {
  event.stopPropagation();
  emit("select", chartId);
}

function openChartContext(event: MouseEvent, chartId: ChartId) {
  event.preventDefault();
  event.stopPropagation();
  emit("select", chartId);
  emit("chart-context", { event, chartId });
}

function handleChartKeydown(event: KeyboardEvent, chartId: ChartId) {
  if (props.readonly) {
    return;
  }
  if (event.key === "Enter") {
    emit("edit", chartId);
    return;
  }
  if (event.key === "Backspace" || event.key === "Delete") {
    event.preventDefault();
    emit("delete-chart", chartId);
  }
}

function startResize(event: PointerEvent, chartId: ChartId, rect: ChartRect) {
  if (props.readonly) {
    return;
  }
  event.preventDefault();
  event.stopPropagation();
  emit("select", chartId);
  const startX = event.clientX;
  const startY = event.clientY;
  const startRect = { ...rect };
  const pointerId = event.pointerId;
  const target = event.currentTarget instanceof HTMLElement ? event.currentTarget : null;
  target?.setPointerCapture(pointerId);
  const move = (moveEvent: PointerEvent) => {
    const width = Math.max(120, startRect.width + moveEvent.clientX - startX);
    const height = Math.max(90, startRect.height + moveEvent.clientY - startY);
    resizing.value = { chartId, rect: { ...startRect, width, height } };
  };
  const finish = () => {
    window.removeEventListener("pointermove", move);
    window.removeEventListener("pointerup", finish);
    target?.releasePointerCapture(pointerId);
    const rectToCommit = resizing.value?.chartId === chartId ? resizing.value.rect : startRect;
    resizing.value = null;
    const anchor = rectToChartAnchor(rectToCommit, props.worksheet, props.projection, props.layoutOrigin);
    if (anchor) {
      emit("resize-chart", { chartId, anchor });
    }
  };
  window.addEventListener("pointermove", move);
  window.addEventListener("pointerup", finish, { once: true });
}
</script>

<template>
  <div class="chart-overlay" aria-hidden="false" @mousedown.self="emit('select', null)">
    <div
      v-for="item in renderedCharts"
      :key="item.chart.id"
      class="chart-overlay__item"
      :class="{ 'chart-overlay__item--selected': item.chart.id === selectedChartId }"
      :style="{
        transform: `translate(${item.rect.x}px, ${item.rect.y}px)`,
        width: `${item.rect.width}px`,
        height: `${item.rect.height}px`,
      }"
      tabindex="0"
      role="button"
      :aria-label="`${item.chart.title || 'Chart'} properties`"
      @mousedown="selectChart($event, item.chart.id)"
      @contextmenu="openChartContext($event, item.chart.id)"
      @dblclick.stop="emit('edit', item.chart.id)"
      @keydown="handleChartKeydown($event, item.chart.id)"
    >
      <component
        :is="chartComponent(item.chart)"
        :data="item.data"
        :width="item.rect.width"
        :height="item.rect.height"
      />
      <span
        v-if="!readonly && item.chart.id === selectedChartId"
        class="chart-overlay__resize-handle"
        role="separator"
        aria-label="Resize chart"
        @pointerdown="startResize($event, item.chart.id, item.rect)"
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

.chart-overlay__item--selected {
  border-color: var(--accent);
  box-shadow: 0 0 0 2px rgb(212 160 23 / 0.28), 0 6px 18px rgb(15 23 42 / 0.12);
}

.chart-overlay__resize-handle {
  position: absolute;
  right: 0;
  bottom: 0;
  width: 14px;
  height: 14px;
  cursor: nwse-resize;
  background: var(--accent);
  border-top: 2px solid #fff;
  border-left: 2px solid #fff;
}
</style>
