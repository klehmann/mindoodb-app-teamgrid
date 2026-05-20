import type { Chart } from "@/features/document/lib/teamgridDocument";

export interface ResolvedChartSeries {
  id: string;
  name: string;
  values: number[];
  color?: string;
}

export interface ResolvedChartData {
  chart: Chart;
  labels: string[];
  series: ResolvedChartSeries[];
}
