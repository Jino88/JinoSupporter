// Setup Plan of a model (Endpoints/WorkingEndpoints.cs, /setup): one row per process with its machine line and setup
// progress. A row's problems are ordinary issues whose setupRowId is the row.
import { getJson, toolPath, toolSend } from "../../toolsApi";

/** "" (not set), O (OK), X (not yet set up) or Δ (setting up now). */
export type SetupState = "" | "O" | "X" | "Δ";
export const SETUP_STATES: { value: SetupState; label: string }[] = [
  { value: "O", label: "OK" },
  { value: "X", label: "Not yet setup" },
  { value: "Δ", label: "Now setup" },
];

export interface SetupFields {
  /** 0 = after the last row (new rows only). */
  no: number;
  layoutNo: string;
  line: string;
  processName: string;
  machinePlan: string;
  machineState: SetupState;
  machineNote: string;
  setupPlan: string;
  setupFinish: string;
  setupStatus: SetupState;
}

export interface SetupRow extends SetupFields {
  id: string;
  modelName: string;
  revision: number;
  createdBy: string;
  createdAt: string;
  updatedBy: string;
  updatedAt: string;
}

const setupPath = (name: string) => `working/models/${encodeURIComponent(name)}/setup`;

export const setupApi = {
  rows: (name: string) => getJson<{ rows: SetupRow[] }>(toolPath(setupPath(name))).then((r) => r.rows),
  add: (name: string, f: SetupFields) => toolSend<SetupRow>(setupPath(name), f),
  import: (name: string, rows: SetupFields[]) => toolSend<{ rows: SetupRow[] }>(`${setupPath(name)}/import`, { rows }).then((r) => r.rows),
  save: (name: string, row: SetupRow, f: SetupFields) =>
    toolSend<SetupRow>(`${setupPath(name)}/${encodeURIComponent(row.id)}`, { ...f, revision: row.revision }, "PUT"),
  remove: (name: string, id: string) => toolSend<void>(`${setupPath(name)}/${encodeURIComponent(id)}`, undefined, "DELETE"),
};

export const blankSetup = (line = ""): SetupFields => ({
  no: 0, layoutNo: "", line, processName: "", machinePlan: "", machineState: "", machineNote: "",
  setupPlan: "", setupFinish: "", setupStatus: "",
});

export const setupFieldsOf = (r: SetupRow): SetupFields => ({
  no: r.no, layoutNo: r.layoutNo, line: r.line, processName: r.processName, machinePlan: r.machinePlan,
  machineState: r.machineState, machineNote: r.machineNote, setupPlan: r.setupPlan, setupFinish: r.setupFinish,
  setupStatus: r.setupStatus,
});

/** The No. a new row gets unless it is changed: one after the highest of the rows and the rows being added. */
export const nextNo = (rows: { no: number }[]) => rows.reduce((m, r) => Math.max(m, r.no), 0) + 1;

/**
 * One more row to add after the ones already being added: the next No. and the line of the row above it (the last new
 * row, else the line shown, else the last saved row).
 */
export function newSetupRow(rows: SetupRow[], adding: SetupFields[], shownLine: string): SetupFields {
  const line = adding.at(-1)?.line ?? (shownLine || rowsInOrder(rows).at(-1)?.line || "");
  return { ...blankSetup(line), no: nextNo([...rows, ...adding]) };
}

/** The rows being added, trimmed, or the first problem among them ("Row 2: …" when there are several). */
export function setupBatch(adding: SetupFields[]): { rows: SetupFields[]; problem: string } {
  for (let i = 0; i < adding.length; i++) {
    const problem = setupProblem(adding[i]);
    if (problem) return { rows: [], problem: adding.length > 1 ? `Row ${i + 1}: ${problem}` : problem };
  }
  return { rows: adding.map((f) => ({ ...f, processName: f.processName.trim() })), problem: "" };
}

export const rowsInOrder =(rows: SetupRow[]) =>
  [...rows].sort((a, b) => a.no - b.no || a.createdAt.localeCompare(b.createdAt) || a.id.localeCompare(b.id));

const ymd = /^\d{4}-\d{2}-\d{2}$/;

/** Message for a row that cannot be saved yet, or "". */
export function setupProblem(f: SetupFields): string {
  if (!f.processName.trim()) return "Process name is required.";
  if (f.processName.trim().length > 200) return "Process name must be 200 characters or fewer.";
  if (!Number.isInteger(f.no) || f.no < 0 || f.no > 100000) return "No. must be a whole number up to 100000.";
  if (f.layoutNo.trim().length > 20) return "No. layout must be 20 characters or fewer.";
  if (f.line.trim().length > 50) return "Line must be 50 characters or fewer.";
  if (f.machineNote.trim().length > 100) return "Machine note must be 100 characters or fewer.";
  for (const [d, label] of [[f.machinePlan, "Machine plan"], [f.setupPlan, "Setup plan"], [f.setupFinish, "Setup finish"]] as const)
    if (d && !ymd.test(d)) return `${label} must be a date.`;
  return "";
}

export interface SetupSummaryLine {
  line: string;
  total: number;
  machineFinish: number;
  setupFinish: number;
}

/** Per line (in the order the lines first appear) and in total: processes, machines OK, setups OK. */
export function setupSummary(rows: SetupRow[]): { lines: SetupSummaryLine[]; total: SetupSummaryLine } {
  const by = new Map<string, SetupSummaryLine>();
  const total: SetupSummaryLine = { line: "Total", total: 0, machineFinish: 0, setupFinish: 0 };
  for (const r of rowsInOrder(rows)) {
    const key = r.line.trim() || "(No line)";
    let s = by.get(key);
    if (!s) by.set(key, (s = { line: key, total: 0, machineFinish: 0, setupFinish: 0 }));
    for (const t of [s, total]) {
      t.total++;
      if (r.machineState === "O") t.machineFinish++;
      if (r.setupStatus === "O") t.setupFinish++;
    }
  }
  return { lines: [...by.values()], total };
}

export const rate = (n: number, of: number) => `${of ? ((n / of) * 100).toFixed(1) : "0.0"}%`;

const MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];

/** 20-Nov (this year) or 20-Nov-27 (another year), as in the Excel sheet; "" stays "". */
export function shortDate(d: string, today: string): string {
  if (!ymd.test(d)) return d;
  const [y, m, day] = d.split("-").map(Number);
  const s = `${day}-${MONTHS[m - 1][0].toUpperCase()}${MONTHS[m - 1].slice(1)}`;
  return d.slice(0, 4) === today.slice(0, 4) ? s : `${s}-${String(y).slice(2)}`;
}

/**
 * A date as typed in Excel → yyyy-MM-dd: 2026-11-20, 2026/11/20, 20-Nov, 20-Nov-26, 11/20, 11/20/2026.
 * Without a year the year of today is used. "" → "", anything else → null.
 */
export function parseSheetDate(text: string, today: string): string | null {
  const t = text.trim();
  if (!t) return "";
  const year = Number(today.slice(0, 4));
  const fullYear = (y: string | undefined) => (y === undefined ? year : y.length <= 2 ? 2000 + Number(y) : Number(y));
  const make = (y: number, m: number, d: number) => {
    const dt = new Date(Date.UTC(y, m - 1, d));
    if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
    return `${String(y).padStart(4, "0")}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  };
  let m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(t);
  if (m) return make(Number(m[1]), Number(m[2]), Number(m[3]));
  m = /^(\d{1,2})[-\s]([A-Za-z]{3})[A-Za-z]*(?:[-\s](\d{2}|\d{4}))?$/.exec(t);
  if (m) {
    const mon = MONTHS.indexOf(m[2].toLowerCase());
    return mon < 0 ? null : make(fullYear(m[3]), mon + 1, Number(m[1]));
  }
  m = /^(\d{1,2})\/(\d{1,2})(?:\/(\d{2}|\d{4}))?$/.exec(t);
  if (m) return make(fullYear(m[3]), Number(m[1]), Number(m[2]));
  return null;
}

/** O / X / Δ and their look-alikes typed in Excel; null when it is none of them. */
export function parseState(text: string): SetupState | null {
  const t = text.trim();
  if (!t) return "";
  if (["O", "o", "0", "○", "◯", "●"].includes(t)) return "O";
  if (["X", "x", "×", "✕"].includes(t)) return "X";
  if (["Δ", "△", "▲", "δ"].includes(t)) return "Δ";
  return null;
}

export interface PastedSetup {
  rows: (SetupFields & { note: string })[];
  problem: string;
}

/**
 * Rows copied from the Setup Plan sheet (tab separated): No | No. Layout | LINE | Process Name | Machine Plan | State | Note |
 * Setup Plan | Finish | Status | Note. Header lines are skipped, an empty LINE takes the one above, and the last Note
 * (the problem text) comes back per row.
 */
export function parseSetupPaste(text: string, today: string): PastedSetup {
  const rows: PastedSetup["rows"] = [];
  let line = "";
  const lines = text.replace(/\r\n?/g, "\n").split("\n");
  for (let i = 0; i < lines.length; i++) {
    const c = lines[i].split("\t").map((v) => v.trim());
    const process = c[3] ?? "";
    if (!process || /^process\s*name$/i.test(process)) continue;
    const at = `Line ${i + 1}`;
    const no = /^\d+$/.test(c[0] ?? "") ? Number(c[0]) : 0;
    if (c[2]) line = c[2];
    const machineState = parseState(c[5] ?? "");
    const setupStatus = parseState(c[9] ?? "");
    if (machineState === null) return { rows: [], problem: `${at}: machine state "${c[5]}" is not O, X or Δ.` };
    if (setupStatus === null) return { rows: [], problem: `${at}: setup status "${c[9]}" is not O, X or Δ.` };
    const dates: string[] = [];
    for (const [k, label] of [[4, "machine plan"], [7, "setup plan"], [8, "setup finish"]] as const) {
      const d = parseSheetDate(c[k] ?? "", today);
      if (d === null) return { rows: [], problem: `${at}: ${label} "${c[k]}" is not a date.` };
      dates.push(d);
    }
    const row = {
      no, layoutNo: c[1] ?? "", line, processName: process, machinePlan: dates[0], machineState, machineNote: c[6] ?? "",
      setupPlan: dates[1], setupFinish: dates[2], setupStatus, note: (c[10] ?? "").replace(/^[-–•]\s*/, ""),
    };
    const problem = setupProblem(row);
    if (problem) return { rows: [], problem: `${at}: ${problem}` };
    rows.push(row);
  }
  return { rows, problem: rows.length ? "" : "No rows found. Copy the rows from No. to Note in Excel and paste them here." };
}
