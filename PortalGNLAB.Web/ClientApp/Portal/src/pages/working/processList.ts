// Process List (공정편성 LIST) of a model (Endpoints/WorkingProcessListEndpoints.cs, /process-list): one row per process
// with its people, standard time, machines, jigs and testers; typed in or read from a JIG / Machine / Tester list workbook.
import { loginUrl } from "../../api";
import { getJson, toolPath, toolSend } from "../../toolsApi";

export type ProcessRow = {
  id: string; no: string; process: string; workers: number | null; stdTime: number | null;
  machine: string; jig: string; tester: string; material: string; note: string;
};
export type ProcessList = { rows: ProcessRow[]; revision: number; updatedBy: string; updatedAt: string | null; source: string };

/** Where the AI took the process order from. */
export type MatchBasis = "instructions" | "process-list" | "workbook";
/** AI match result: rows (id "", in process order) for review, nothing saved. */
export type ProcessMatch = { rows: ProcessRow[]; note: string; unmatched: string[]; basis: MatchBasis; model: string };

export const BASIS_TEXT: Record<MatchBasis, string> = {
  instructions: "Process order from the Work Instructions",
  "process-list": "Process order from the current Process List",
  workbook: "Process order from the workbook's process numbers",
};

const path = (name: string) => `working/models/${encodeURIComponent(name)}/process-list`;

export const processListApi = {
  get: (name: string) => getJson<ProcessList>(toolPath(path(name))),
  save: (name: string, rows: ProcessRow[], revision: number) => toolSend<ProcessList>(path(name), { rows, revision }, "PUT"),
  /** Reads a JIG / Machine / Tester list workbook into the rows, in place of what is there. */
  import: (name: string, file: File, revision: number): Promise<ProcessList> =>
    upload<ProcessList>(`${path(name)}/import`, file, { revision: String(revision) }),
  /** AI matches the workbook's machines, jigs and testers to the processes, in process order. Saves nothing; takes 1-3 minutes. */
  aiMatch: (name: string, file: File): Promise<ProcessMatch> => upload<ProcessMatch>(`${path(name)}/ai-match`, file),
};

function upload<T>(route: string, file: File, fields: Record<string, string> = {}): Promise<T> {
  const form = new FormData();
  form.append("file", file, file.name);
  for (const [k, v] of Object.entries(fields)) form.append(k, v);
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open("POST", toolPath(route));
    xhr.withCredentials = true;
    xhr.setRequestHeader("Accept", "application/json");
    xhr.setRequestHeader("X-Requested-With", "PortalGNLAB");
    xhr.onload = () => {
      if (xhr.status === 401) window.location.assign(loginUrl(location.pathname));
      let body: { error?: string } | null = null;
      try { body = JSON.parse(xhr.responseText); } catch { /* not JSON */ }
      if (xhr.status >= 200 && xhr.status < 300) resolve(body as T);
      else reject(new Error(body?.error ?? `Request failed (HTTP ${xhr.status})`));
    };
    xhr.onerror = () => reject(new Error("The upload failed. Check the connection and try again."));
    xhr.send(form);
  });
}

export const blankRow = (no = ""): ProcessRow =>
  ({ id: "", no, process: "", workers: null, stdTime: null, machine: "", jig: "", tester: "", material: "", note: "" });

/** The number after the last numbered row ("07" after "06"), in the same width; "" when no row has a number. */
export function nextNo(rows: ProcessRow[]): string {
  const nums = rows.map((r) => /^\d+$/.test(r.no.trim()) ? r.no.trim() : "").filter(Boolean);
  if (nums.length === 0) return "";
  const last = nums[nums.length - 1];
  return String(Number(last) + 1).padStart(last.length, "0");
}

/** People, standard time and the slowest process (the line's bottleneck) of the rows. */
export function totals(rows: ProcessRow[]) {
  let workers = 0, time = 0;
  let slowest: ProcessRow | null = null;
  for (const r of rows) {
    workers += r.workers ?? 0;
    time += r.stdTime ?? 0;
    if (r.stdTime != null && (slowest === null || r.stdTime > (slowest.stdTime ?? 0))) slowest = r;
  }
  return { processes: rows.length, workers, time: Math.round(time * 1000) / 1000, slowest };
}

/** A number typed into a cell: "" is none, else a non-negative number (whole for people). */
export function cellNumber(text: string, whole: boolean): number | null | undefined {
  const t = text.trim().replace(",", ".");
  if (t === "") return null;
  const n = Number(t);
  if (!Number.isFinite(n) || n < 0 || (whole && !Number.isInteger(n))) return undefined;
  return n;
}

/** Problems that keep the rows from being saved, or null. */
export function rowsProblem(rows: ProcessRow[]): string | null {
  for (const [i, r] of rows.entries())
    if (!r.no.trim() && !r.process.trim()) return `Row ${i + 1} needs a number or a process name.`;
  return null;
}
