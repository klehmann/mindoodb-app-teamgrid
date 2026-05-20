import { scaleBand, scaleLinear } from "d3-scale";

export function createCategoryScale(labels: string[], width: number, padding = 0.18) {
  return scaleBand<string>()
    .domain(labels)
    .range([0, width])
    .padding(padding);
}

export function createValueScale(values: number[], height: number) {
  const min = Math.min(0, ...values);
  const max = Math.max(0, ...values, 1);
  return scaleLinear()
    .domain([min, max])
    .nice()
    .range([height, 0]);
}

export function createHorizontalValueScale(values: number[], width: number) {
  const min = Math.min(0, ...values);
  const max = Math.max(0, ...values, 1);
  return scaleLinear()
    .domain([min, max])
    .nice()
    .range([0, width]);
}
