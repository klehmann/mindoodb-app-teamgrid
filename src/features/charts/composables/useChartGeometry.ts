import { computed, type Ref } from "vue";
import { chartAnchorToRect } from "@/features/charts/lib/chartGeometry";
import type { Chart, Worksheet } from "@/features/document/lib/teamgridDocument";
import type { GridProjection } from "@/features/grid/lib/gridProjection";

export function useChartGeometry(chart: Ref<Chart>, worksheet: Ref<Worksheet>, projection: Ref<GridProjection>) {
  return computed(() => chartAnchorToRect(chart.value.anchor, worksheet.value, projection.value));
}
