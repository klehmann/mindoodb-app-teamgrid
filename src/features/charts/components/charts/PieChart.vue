<script setup lang="ts">
import { computed } from "vue";
import { arc, pie } from "d3-shape";
import type { PieArcDatum } from "d3-shape";
import { chartSeriesColor } from "@/features/charts/lib/excelPalette";
import type { ResolvedChartData } from "@/features/charts/lib/types";

const props = defineProps<{ data: ResolvedChartData; width: number; height: number }>();
const radius = computed(() => Math.max(20, Math.min(props.width, props.height) / 2 - 42));
const center = computed(() => ({ x: props.width / 2, y: props.height / 2 + 8 }));
const values = computed(() => props.data.series[0]?.values ?? []);
const slices = computed(() => pie<number>().sort(null).value((value) => Math.max(0, value))(values.value));
const path = computed(() => arc<PieArcDatum<number>>().innerRadius(0).outerRadius(radius.value));
</script>

<template>
  <svg class="chart-svg" :viewBox="`0 0 ${width} ${height}`" role="img" :aria-label="data.chart.title || 'Pie chart'">
    <rect :width="width" :height="height" rx="4" fill="#fff" />
    <text v-if="data.chart.title" :x="width / 2" y="20" text-anchor="middle" class="chart-title">{{ data.chart.title }}</text>
    <g :transform="`translate(${center.x}, ${center.y})`">
      <path
        v-for="(slice, index) in slices"
        :key="index"
        :d="path(slice) ?? ''"
        :fill="chartSeriesColor(index, data.series[0]?.color)"
        stroke="#fff"
        stroke-width="1"
      />
    </g>
    <g :transform="`translate(${Math.max(12, width - 120)}, 40)`">
      <g v-for="(label, index) in data.labels" :key="label" :transform="`translate(0, ${index * 18})`">
        <rect width="10" height="10" :fill="chartSeriesColor(index)" />
        <text x="16" y="9" class="chart-legend-label">{{ label }}</text>
      </g>
    </g>
  </svg>
</template>

<style scoped>
.chart-svg { display: block; width: 100%; height: 100%; font-family: Calibri, Arial, sans-serif; }
.chart-title { font-size: 14px; font-weight: 600; fill: #1f2937; }
.chart-legend-label { font-size: 10px; fill: #4b5563; }
</style>
