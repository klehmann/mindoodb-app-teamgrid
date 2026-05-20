export const EXCEL_SERIES_COLORS = [
  "#4472c4",
  "#ed7d31",
  "#a5a5a5",
  "#ffc000",
  "#5b9bd5",
  "#70ad47",
  "#264478",
  "#9e480e",
] as const;

export function chartSeriesColor(index: number, explicitColor?: string) {
  return explicitColor || EXCEL_SERIES_COLORS[index % EXCEL_SERIES_COLORS.length];
}
