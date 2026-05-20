<script setup lang="ts">
import { computed } from "vue";
import { line } from "d3-shape";
import { scalePoint } from "d3-scale";
import { createValueScale } from "@/features/charts/lib/d3Scales";
import { valueAxisTicks } from "@/features/charts/lib/d3Axes";
import { chartSeriesColor } from "@/features/charts/lib/excelPalette";
import { chartLegendTransform, chartMarginWithLegend, shouldShowSeriesLegend } from "@/features/charts/lib/chartLegend";
import type { ResolvedChartData } from "@/features/charts/lib/types";

const props = defineProps<{ data: ResolvedChartData; width: number; height: number }>();
const margin = computed(() => chartMarginWithLegend(
  { top: 36, right: 18, bottom: 44, left: 48 },
  props.data.chart.legend?.position,
  props.data.series.length,
));
const showLegend = computed(() => shouldShowSeriesLegend(props.data.chart.legend?.position, props.data.series.length));
const legendTransform = computed(() => chartLegendTransform(props.width, props.height, props.data.chart.legend?.position, props.data.series.length));
const plot = computed(() => ({
  width: Math.max(1, props.width - margin.value.left - margin.value.right),
  height: Math.max(1, props.height - margin.value.top - margin.value.bottom),
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
        v-if="data.chart.style?.showGridlines !== false"
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
        <path :d="paths[seriesIndex]" fill="none" :stroke="chartSeriesColor(seriesIndex, series.color ?? data.chart.style?.colors?.[seriesIndex])" stroke-width="2" />
        <circle
          v-for="(value, index) in series.values"
          :key="`${series.id}-${index}`"
          :cx="x(data.labels[index]) ?? 0"
          :cy="y(value)"
          r="3"
          :fill="chartSeriesColor(seriesIndex, series.color ?? data.chart.style?.colors?.[seriesIndex])"
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
    <g v-if="showLegend" :transform="legendTransform">
      <g v-for="(series, index) in data.series" :key="`legend-${series.id}`" :transform="`translate(0, ${index * 18})`">
        <rect width="10" height="10" :fill="chartSeriesColor(index, series.color ?? data.chart.style?.colors?.[index])" />
        <text x="16" y="9" class="chart-legend-label">{{ series.name }}</text>
      </g>
    </g>
  </svg>
</template>

<style scoped>
.chart-svg { display: block; width: 100%; height: 100%; font-family: Calibri, Arial, sans-serif; }
.chart-title { font-size: 14px; font-weight: 600; fill: #1f2937; }
.chart-gridline { stroke: #d9e2f3; stroke-width: 1; }
.chart-axis { stroke: #6b7280; stroke-width: 1; }
.chart-axis-label { font-size: 10px; fill: #4b5563; }
.chart-legend-label { font-size: 10px; fill: #4b5563; }
</style>
