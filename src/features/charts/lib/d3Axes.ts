import type { ScaleLinear } from "d3-scale";

export interface AxisTick {
  value: number;
  label: string;
  position: number;
}

export function valueAxisTicks(scale: ScaleLinear<number, number>, count = 5): AxisTick[] {
  return scale.ticks(count).map((value) => ({
    value,
    label: formatTick(value),
    position: scale(value),
  }));
}

function formatTick(value: number) {
  const absolute = Math.abs(value);
  if (absolute >= 1000) {
    return value.toLocaleString("en-US", { maximumFractionDigits: 0 });
  }
  if (absolute >= 10 || Number.isInteger(value)) {
    return String(Math.round(value * 100) / 100);
  }
  return String(Math.round(value * 1000) / 1000);
}
