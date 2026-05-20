<script setup lang="ts">
import { computed } from "vue";
import { createCategoryScale, createHorizontalValueScale } from "@/features/charts/lib/d3Scales";
import { valueAxisTicks } from "@/features/charts/lib/d3Axes";
import { chartSeriesColor } from "@/features/charts/lib/excelPalette";
import type { ResolvedChartData } from "@/features/charts/lib/types";

const props = defineProps<{ data: ResolvedChartData; width: number; height: number }>();
const margin = { top: 36, right: 18, bottom: 28, left: 72 };
const plot = computed(() => ({
  width: Math.max(1, props.width - margin.left - margin.right),
  height: Math.max(1, props.height - margin.top - margin.bottom),
}));
const y = computed(() => createCategoryScale(props.data.labels, plot.value.height));
const x = computed(() => createHorizontalValueScale(props.data.series.flatMap((series) => series.values), plot.value.width));
const ticks = computed(() => valueAxisTicks(x.value));
const groupHeight = computed(() => y.value.bandwidth());
const seriesHeight = computed(() => Math.max(1, groupHeight.value / Math.max(1, props.data.series.length)));
</script>

<template>
  <svg class="chart-svg" :viewBox="`0 0 ${width} ${height}`" role="img" :aria-label="data.chart.title || 'Bar chart'">
    <rect :width="width" :height="height" rx="4" fill="#fff" />
    <text v-if="data.chart.title" :x="width / 2" y="20" text-anchor="middle" class="chart-title">{{ data.chart.title }}</text>
    <g :transform="`translate(${margin.left}, ${margin.top})`">
      <line
        v-for="tick in ticks"
        :key="tick.value"
        :x1="tick.position"
        :x2="tick.position"
        y1="0"
        :y2="plot.height"
        class="chart-gridline"
      />
      <text
        v-for="tick in ticks"
        :key="`label-${tick.value}`"
        :x="tick.position"
        :y="plot.height + 18"
        text-anchor="middle"
        class="chart-axis-label"
      >{{ tick.label }}</text>
      <g v-for="(series, seriesIndex) in data.series" :key="series.id">
        <rect
          v-for="(value, index) in series.values"
          :key="`${series.id}-${index}`"
          :x="x(Math.min(0, value))"
          :y="(y(data.labels[index]) ?? 0) + seriesIndex * seriesHeight"
          :width="Math.abs(x(value) - x(0))"
          :height="Math.max(1, seriesHeight - 2)"
          :fill="chartSeriesColor(seriesIndex, series.color)"
        />
      </g>
      <line :x1="x(0)" :x2="x(0)" y1="0" :y2="plot.height" class="chart-axis" />
      <text
        v-for="label in data.labels"
        :key="label"
        x="-8"
        :y="(y(label) ?? 0) + groupHeight / 2 + 4"
        text-anchor="end"
        class="chart-axis-label"
      >{{ label }}</text>
    </g>
  </svg>
</template>

<style scoped>
.chart-svg { display: block; width: 100%; height: 100%; font-family: Calibri, Arial, sans-serif; }
.chart-title { font-size: 14px; font-weight: 600; fill: #1f2937; }
.chart-gridline { stroke: #d9e2f3; stroke-width: 1; }
.chart-axis { stroke: #6b7280; stroke-width: 1; }
.chart-axis-label { font-size: 10px; fill: #4b5563; }
</style>
