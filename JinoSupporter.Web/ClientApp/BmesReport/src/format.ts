import type { MetricUnit } from "./contract";

const number0 = new Intl.NumberFormat("en-US", {
  minimumFractionDigits: 0,
  maximumFractionDigits: 0,
});
const number2 = new Intl.NumberFormat("en-US", {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});
const usd0 = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
  currencyDisplay: "narrowSymbol",
  minimumFractionDigits: 0,
  maximumFractionDigits: 0,
});
const usd2 = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
  currencyDisplay: "narrowSymbol",
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

const safe = (value: number | null | undefined): value is number =>
  typeof value === "number" && Number.isFinite(value);

export function roundHalfToEven(value: number, digits = 0): number {
  const factor = 10 ** digits;
  const scaled = value * factor;
  const lower = Math.floor(scaled);
  const fraction = scaled - lower;
  const tolerance = Number.EPSILON * Math.max(1, Math.abs(scaled)) * 4;
  if (Math.abs(fraction - 0.5) <= tolerance) {
    return (lower % 2 === 0 ? lower : lower + 1) / factor;
  }
  return Math.round(scaled) / factor;
}

export const formatCount = (value: number | null | undefined): string => (safe(value) ? number0.format(roundHalfToEven(value)) : "-");
export const formatPpm = (value: number | null | undefined): string => (safe(value) ? number0.format(roundHalfToEven(value)) : "-");
export const formatPercent = (value: number | null | undefined): string =>
  safe(value) ? `${number2.format(roundHalfToEven(value, 2))}%` : "-";
export const formatShareRatio = (value: number | null | undefined): string =>
  safe(value) ? `${number2.format(roundHalfToEven(value * 100, 2))}%` : "-";
export const formatUsd = (value: number | null | undefined, cents = false): string =>
  safe(value) ? (cents ? usd2 : usd0).format(roundHalfToEven(value, cents ? 2 : 0)) : "-";
export const formatVnd = (value: number | null | undefined): string =>
  safe(value) ? `${number0.format(roundHalfToEven(value))} VND` : "-";
export const formatNumber2 = (value: number | null | undefined): string => (safe(value) ? number2.format(roundHalfToEven(value, 2)) : "-");

export function formatMetric(value: number | null | undefined, unit: MetricUnit): string {
  if (unit === "percent") return formatPercent(value);
  if (unit === "usd") return formatUsd(value);
  if (unit === "ppm") return formatPpm(value);
  return safe(value) ? number2.format(roundHalfToEven(value, 2)) : "-";
}
