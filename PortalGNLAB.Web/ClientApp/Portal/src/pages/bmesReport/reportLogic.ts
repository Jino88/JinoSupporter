// Report page logic: port of the view helpers of NgRateSetupPanel / NgRateModelGroupPicker.razor.
import type { PickerEntry, PickerGroup, Selection, Stage } from "../../bmesReportApi";

export const selectedCount = (s: Selection) => s.groupIds.length + s.subEntries.length;

export function toggleGroup(s: Selection, id: number): Selection {
  return s.groupIds.includes(id)
    ? { ...s, groupIds: s.groupIds.filter((g) => g !== id) }
    : { ...s, groupIds: [...s.groupIds, id] };
}

export function toggleSubEntry(s: Selection, key: string): Selection {
  return s.subEntries.includes(key)
    ? { ...s, subEntries: s.subEntries.filter((k) => k !== key) }
    : { ...s, subEntries: [...s.subEntries, key] };
}

/** "MODEL" for a Mid entry, else "SUB n". */
export const entryLevelLabel = (entry: PickerEntry) =>
  entry.key.includes("|mid|") ? "MODEL" : `SUB ${Math.max(1, entry.depth - 1)}`;

/** Selected groups of one product-group tab (the count in the tab). */
export function selectedInProductGroup(groups: PickerGroup[], s: Selection, productGroup: string): number {
  return s.groupIds.filter((id) => groups.some((g) => g.id === id && g.productGroup === productGroup)).length;
}

/** Same checks, in the same order, as BmesReportPage.GetReportAsync. */
export function validateReport(start: string, end: string, credentialsConfigured: boolean, s: Selection): string | null {
  if (end < start) return "End date must be on or after start date.";
  if (!credentialsConfigured) return "BMES credentials are not configured.";
  if (selectedCount(s) === 0) return "Select at least one sub group.";
  return null;
}

export const pct = (value: number) => Math.round(Math.min(100, Math.max(0, value)) * 10) / 10 + "%";

export function formatElapsed(seconds: number): string {
  return seconds < 60 ? seconds.toFixed(1) + "s" : `${Math.floor(seconds / 60)}m ${String(Math.floor(seconds % 60)).padStart(2, "0")}s`;
}

export const stageIcon = (s: Stage) => (s.status === "Done" ? "✓" : s.status === "Failed" ? "!" : s.status === "Running" ? "•" : "○");

export function logClass(line: string): string {
  const l = line.toUpperCase();
  if (l.includes("[ERROR]")) return "err";
  if (l.includes("[WARN]")) return "warn";
  if (l.includes("DONE") || l.includes("REPORT COMPLETE")) return "ok";
  return "";
}

/** MM/dd/yyyy of a yyyy-MM-dd value, for the collapsed setup bar. */
export function usDate(iso: string): string {
  const [y, m, d] = iso.split("-");
  return y && m && d ? `${m}/${d}/${y}` : iso;
}
