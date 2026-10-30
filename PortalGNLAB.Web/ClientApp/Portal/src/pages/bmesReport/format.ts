// Number / text helpers shared by the report pages. They mirror the .NET formats the
// Blazor pages used ("N0", "0.00", ordinal / ordinal-ignore-case ordering).

const grouped0 = new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 });

/** Math.Round(double): ties go to the even neighbour. */
export function roundHalfEven(v: number): number {
  if (Math.abs(v % 1) === 0.5) return 2 * Math.round(v / 2);
  return Math.round(v);
}

/** ((long)Math.Round(v)).ToString("N0") */
export function n0(v: number): string {
  return grouped0.format(roundHalfEven(v) || 0);
}

/** decimal.ToString(format) with the default away-from-zero rounding of number formatting. */
export function fixedGrouped(v: number, digits: number): string {
  return new Intl.NumberFormat("en-US", { minimumFractionDigits: digits, maximumFractionDigits: digits }).format(v);
}

/** String.Compare(a, b, StringComparison.Ordinal) */
export function cmpOrdinal(a: string, b: string): number {
  return a < b ? -1 : a > b ? 1 : 0;
}

/** String.Compare(a, b, StringComparison.OrdinalIgnoreCase) */
export function cmpIgnoreCase(a: string, b: string): number {
  return cmpOrdinal(a.toUpperCase(), b.toUpperCase());
}

/** Distinct(StringComparer.OrdinalIgnoreCase): first spelling wins. */
export function distinctIgnoreCase(values: Iterable<string>): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const v of values) {
    const key = v.toUpperCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(v);
  }
  return out;
}

/** GroupBy that keeps the order in which keys first appear (like LINQ). */
export function groupBy<T>(items: Iterable<T>, keyOf: (item: T) => string): Map<string, T[]> {
  const map = new Map<string, T[]>();
  for (const item of items) {
    const key = keyOf(item);
    const list = map.get(key);
    if (list) list.push(item);
    else map.set(key, [item]);
  }
  return map;
}

export function pad2(n: number): string {
  return String(n).padStart(2, "0");
}

/** yyyy-MM-dd of a local date. */
export function isoDate(d: Date): string {
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

/** yyyyMMdd_HHmmss / yyyyMMdd-HHmm style stamps for download names. */
export function stamp(d: Date, withSeconds: boolean, separator = "_"): string {
  const date = `${d.getFullYear()}${pad2(d.getMonth() + 1)}${pad2(d.getDate())}`;
  const time = `${pad2(d.getHours())}${pad2(d.getMinutes())}${withSeconds ? pad2(d.getSeconds()) : ""}`;
  return date + separator + time;
}

/** CsvExportUtility.Escape */
export function csvEscape(value: string | null | undefined): string {
  const text = value ?? "";
  if (!/["\r\n,]/.test(text)) return text;
  return '"' + text.replaceAll('"', '""') + '"';
}

/** CsvExportUtility.Build: one line per row, CRLF, UTF-8 BOM. */
export function csvText(rows: (string | null | undefined)[][]): string {
  return "﻿" + rows.map((r) => r.map(csvEscape).join(",") + "\r\n").join("");
}

export function htmlEscape(value: string | null | undefined): string {
  return (value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

export function base64ToBlob(base64: string, type: string): Blob {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return new Blob([bytes], { type });
}
