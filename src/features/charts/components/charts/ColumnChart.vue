<script setup lang="ts">
import { computed } from "vue";
import { createCategoryScale, createValueScale } from "@/features/charts/lib/d3Scales";
import { valueAxisTicks } from "@/features/charts/lib/d3Axes";
import { chartSeriesColor } from "@/features/charts/lib/excelPalette";
import type { ResolvedChartData } from "@/features/charts/lib/types";

const props = defineProps<{
  data: ResolvedChartData;
  width: number;
  height: number;
}>();

const margin = { top: 36, right: 18, bottom: 44, left: 48 };
const plot = computed(() => ({
  width: Math.max(1, props.width - margin.left - margin.right),
  height: Math.max(1, props.height - margin.top - margin.bottom),
}));
const x = computed(() => createCategoryScale(props.data.labels, plot.value.width));
const y = computed(() => createValueScale(props.data.series.flatMap((series) => series.values), plot.value.height));
const ticks = computed(() => valueAxisTicks(y.value));
const groupWidth = computed(() => x.value.bandwidth());
const seriesWidth = computed(() => Math.max(1, groupWidth.value / Math.max(1, props.data.series.length)));
</script>

<template>
  <svg class="chart-svg" :viewBox="`0 0 ${width} ${height}`" role="img" :aria-label="data.chart.title || 'Column chart'">
    <rect :width="width" :height="height" rx="4" fill="#fff" />
    <text v-if="data.chart.title" :x="width / 2" y="20" text-anchor="middle" class="chart-title">{{ data.chart.title }}</text>
    <g :transform="`translate(${margin.left}, ${margin.top})`">
      <line
        v-for="tick in ticks"
        :key="tick.value"
        x1="0"
        :x2="plot.width"
        :y1="tick.position"
        :y2="tick.position"
        class="chart-gridline"
      />
      <text
        v-for="tick in ticks"
        :key="`label-${tick.value}`"
        x="-8"
        :y="tick.position + 4"
        text-anchor="end"
        class="chart-axis-label"
      >{{ tick.label }}</text>
      <g v-for="(series, seriesIndex) in data.series" :key="series.id">
        <rect
          v-for="(value, index) in series.values"
          :key="`${series.id}-${index}`"
          :x="(x(data.labels[index]) ?? 0) + seriesIndex * seriesWidth"
          :y="y(Math.max(0, value))"
          :width="Math.max(1, seriesWidth - 2)"
          :height="Math.abs(y(value) - y(0))"
          :fill="chartSeriesColor(seriesIndex, series.color)"
        />
      </g>
      <line x1="0" :x2="plot.width" :y1="y(0)" :y2="y(0)" class="chart-axis" />
      <text
        v-for="label in data.labels"
        :key="label"
        :x="(x(label) ?? 0) + groupWidth / 2"
        :y="plot.height + 18"
        text-anchor="middle"
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
