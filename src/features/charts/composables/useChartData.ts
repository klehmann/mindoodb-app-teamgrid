import { computed, type Ref } from "vue";
import { resolveChartData } from "@/features/charts/lib/chartDataResolution";
import type { Chart } from "@/features/document/lib/teamgridDocument";
import type { FormulaContext } from "@/features/formulas/lib";

export function useChartData(chart: Ref<Chart>, formulaContext: Ref<FormulaContext>) {
  return computed(() => resolveChartData(chart.value, formulaContext.value));
}
