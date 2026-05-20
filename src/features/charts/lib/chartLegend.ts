import type { ChartLegend } from "@/features/document/lib/teamgridDocument";

export interface ChartMargin {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

const LEGEND_WIDTH = 120;
const LEGEND_ROW_HEIGHT = 18;
const LEGEND_PADDING = 12;

export function chartMarginWithLegend(base: ChartMargin, position: ChartLegend["position"] | undefined, itemCount: number): ChartMargin {
  const legendHeight = Math.min(96, itemCount * LEGEND_ROW_HEIGHT + LEGEND_PADDING);
  switch (position) {
    case "left":
      return { ...base, left: base.left + LEGEND_WIDTH };
    case "top":
      return { ...base, top: base.top + legendHeight };
    case "bottom":
      return { ...base, bottom: base.bottom + legendHeight };
    case "none":
      return base;
    case "right":
    default:
      return { ...base, right: base.right + LEGEND_WIDTH };
  }
}

export function chartLegendTransform(width: number, height: number, position: ChartLegend["position"] | undefined, itemCount: number) {
  const legendHeight = Math.min(96, itemCount * LEGEND_ROW_HEIGHT + LEGEND_PADDING);
  switch (position) {
    case "left":
      return "translate(12, 40)";
    case "top":
      return "translate(12, 32)";
    case "bottom":
      return `translate(12, ${Math.max(40, height - legendHeight)})`;
    case "right":
    default:
      return `translate(${Math.max(12, width - LEGEND_WIDTH + 8)}, 40)`;
  }
}

export function shouldShowSeriesLegend(position: ChartLegend["position"] | undefined, itemCount: number) {
  return position !== "none" && itemCount > 0;
}
