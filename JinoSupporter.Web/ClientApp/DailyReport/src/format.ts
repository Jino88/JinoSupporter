import type { DailyReportCause } from "./contract";

const ppmFormatter = new Intl.NumberFormat("en-US", {
  minimumFractionDigits: 0,
  maximumFractionDigits: 0,
  useGrouping: true,
});
const percentFormatter = new Intl.NumberFormat("en-US", {
  minimumFractionDigits: 0,
  maximumFractionDigits: 1,
  useGrouping: false,
});

export function formatPpm(value: number): string {
  return ppmFormatter.format(value);
}

export function formatSignedPpm(value: number): string {
  if (value === 0) return "0";
  return `${value > 0 ? "+" : "−"}${formatPpm(Math.abs(value))}`;
}

export function formatPercent(value: number): string {
  return `(${value > 0 ? "+" : ""}${percentFormatter.format(value)}%)`;
}

export type Direction = "worse" | "better" | "flat";

export function direction(value: number, isNew = false): Direction {
  if (isNew || value > 0) return "worse";
  if (value < 0) return "better";
  return "flat";
}

export function directionClass(value: number, isNew = false): string {
  return `dr-is-${direction(value, isNew)}`;
}

export function formatCauseDelta(cause: DailyReportCause): string {
  if (cause.isNew) return "NEW";
  if (cause.deltaPpm === 0) return "–";
  return formatSignedPpm(cause.deltaPpm);
}

export function processTypeLabel(processType: string): string {
  const value = processType.trim().toUpperCase();
  if (value.startsWith("MAIN")) return "MAIN";
  if (value.startsWith("SUB")) return "SUB";
  if (value.startsWith("FUNC")) return "FUNC";
  if (value.startsWith("VIS")) return "VIS";
  return value.slice(0, 4) || "—";
}

export function isMainProcessType(processType: string): boolean {
  return processType.trim().toUpperCase().startsWith("MAIN");
}
