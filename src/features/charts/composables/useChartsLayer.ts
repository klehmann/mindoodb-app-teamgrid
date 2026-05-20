import { computed, type Ref } from "vue";
import type { Worksheet } from "@/features/document/lib/teamgridDocument";

export function useChartsLayer(worksheet: Ref<Worksheet>) {
  return computed(() => worksheet.value.chartOrder
    .map((id) => worksheet.value.chartsById[id])
    .filter((chart) => chart && !chart.deletedAt));
}
