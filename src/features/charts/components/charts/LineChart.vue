<script setup lang="ts">
import { computed } from "vue";
import { line } from "d3-shape";
import { scalePoint } from "d3-scale";
import { createValueScale } from "@/features/charts/lib/d3Scales";
import { valueAxisTicks } from "@/features/charts/lib/d3Axes";
import { chartSeriesColor } from "@/features/charts/lib/excelPalette";
import type { ResolvedChartData } from "@/features/charts/lib/types";

const props = defineProps<{ data: ResolvedChartData; width: number; height: number }>();
const margin = { top: 36, right: 18, bottom: 44, left: 48 };
const plot = computed(() => ({
  width: Math.max(1, props.width - margin.left - margin.right),
  height: Math.max(1, props.height - margin.top - margin.bottom),
}));
const x = computed(() => scalePoint<string>().domain(props.data.labels).range([0, plot.value.width]).padding(0.4));
const y = computed(() => createValueScale(props.data.series.flatMap((series) => series.values), plot.value.height));
const ticks = computed(() => valueAxisTicks(y.value));
const paths = computed(() => props.data.series.map((series) => line<number>()
  .x((_, index) => x.value(props.data.labels[index]) ?? 0)
  .y((value) => y.value(value))(series.values) ?? ""));
</script>

<template>
  <svg class="chart-svg" :viewBox="`0 0 ${width} ${height}`" role="img" :aria-label="data.chart.title || 'Line chart'">
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
        <path :d="paths[seriesIndex]" fill="none" :stroke="chartSeriesColor(seriesIndex, series.color)" stroke-width="2" />
        <circle
          v-for="(value, index) in series.values"
          :key="`${series.id}-${index}`"
          :cx="x(data.labels[index]) ?? 0"
          :cy="y(value)"
          r="3"
          :fill="chartSeriesColor(seriesIndex, series.color)"
        />
      </g>
      <line x1="0" :x2="plot.width" :y1="y(0)" :y2="y(0)" class="chart-axis" />
      <text
        v-for="label in data.labels"
        :key="label"
        :x="x(label) ?? 0"
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
