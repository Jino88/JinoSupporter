// Setup Plan of a model (Endpoints/WorkingEndpoints.cs, /setup): one row per process with its machine line and setup
// progress. A row's problems are ordinary issues whose setupRowId is the row.
import { getJson, toolPath, toolSend } from "../../toolsApi";
import { COLORS, type PlanEvent } from "./plan";

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

/** Issue "Process" of a problem of a setup row: its process name (not its line), cut to the 60 characters the field takes. */
export const issueProcessOf = (r: { processName: string }) => r.processName.trim().slice(0, 60);

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

/** The line a row is on, "(No line)" when it has none. */
export const lineOf = (r: { line: string }) => r.line.trim() || "(No line)";

/**
 * The rows as plan events, the way Dev Plan shows them: a line per event (in the order the lines first appear) and its
 * processes as tasks, each from its machine plan day to its setup plan day (one of them when only one is set), done when
 * its setup is OK or has a finish day.
 */
export function setupEvents(rows: SetupRow[]): PlanEvent[] {
  const ordered = rowsInOrder(rows);
  const lines = [...new Set(ordered.map(lineOf))];
  return lines.map((line, i) => ({
    id: "line:" + line, name: line, start: "", end: "", color: COLORS[(i + 1) % COLORS.length], rows: [], bomFrom: "",
    tasks: ordered.filter((r) => lineOf(r) === line).map((r) => ({
      id: r.id, name: r.processName || `No. ${r.no}`, owner: "", note: "",
      start: r.machinePlan || r.setupPlan, end: r.setupPlan || r.machinePlan, done: r.setupStatus === "O" || !!r.setupFinish,
    })),
  }));
}

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

/** Columns of a Setup Plan row in the sheet, No. to Note. */
const SHEET_COLUMNS = 11;

/**
 * The cells of a pasted line without tabs (copied through a viewer that turns tabs into spaces): split where two or
 * more spaces stand, empty cells past the 11 sheet columns dropped.
 */
export function cellsOf(line: string): string[] {
  const c = (line.includes("\t") ? line.split("\t") : line.trim().split(/\s{2,}/)).map((v) => v.trim());
  while (c.length > SHEET_COLUMNS && c[c.length - 1] === "") c.pop();
  return c;
}

const ENTITIES: Record<string, string> = { nbsp: " ", amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };
const textOf = (html: string) =>
  html.replace(/<br\s*\/?>/gi, " ").replace(/<[^>]*>/g, "")
    .replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e: string) =>
      e[0] === "#" ? String.fromCodePoint(e[1].toLowerCase() === "x" ? parseInt(e.slice(2), 16) : Number(e.slice(1))) : ENTITIES[e.toLowerCase()] ?? m)
    .replace(/\s+/g, " ").trim();

/**
 * The rows of the first HTML table on the clipboard (Excel and Google Sheets put one there) as tab-separated lines;
 * a cell spanning several columns is followed by empty ones, as in the sheet. "" when there is no table.
 */
export function tableToTsv(html: string): string {
  const table = /<table[\s\S]*?<\/table>/i.exec(html);
  if (!table) return "";
  return (table[0].match(/<tr[\s\S]*?<\/tr>/gi) ?? []).map((tr) =>
    (tr.match(/<t[dh]\b[^>]*>[\s\S]*?<\/t[dh]>/gi) ?? []).flatMap((td) => {
      const span = Number(/colspan\s*=\s*["']?(\d+)/i.exec(td)?.[1] ?? 1);
      return [textOf(td), ...Array(Math.max(span, 1) - 1).fill("")];
    }).join("\t")).join("\n");
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
/**
 * Records of tab-separated text as Excel copies it: a cell with a line break, tab or quote comes in quotes ("No.
 * Layout"), with "" for a quote inside.
 */
export function tsvRecords(text: string): string[][] {
  const records: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch !== '"') cell += ch;
      else if (text[i + 1] === '"') { cell += '"'; i++; }
      else quoted = false;
    } else if (ch === '"' && cell === "") quoted = true;
    else if (ch === "\t") { row.push(cell); cell = ""; }
    else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      row.push(cell);
      records.push(row);
      row = [];
      cell = "";
    } else cell += ch;
  }
  if (cell || row.length) { row.push(cell); records.push(row); }
  return records.map((r) => r.map((c) => c.trim()));
}

/** Where each field is in a pasted row: No, No. layout, LINE, process, machine plan / state / note, setup plan / finish / status, note. */
export interface SheetColumns {
  no: number; layout: number; line: number; process: number;
  machinePlan: number; machineState: number; machineNote: number;
  setupPlan: number; setupFinish: number; setupStatus: number; note: number;
}

const PLAIN_COLUMNS: SheetColumns = {
  no: 0, layout: 1, line: 2, process: 3, machinePlan: 4, machineState: 5, machineNote: 6, setupPlan: 7, setupFinish: 8, setupStatus: 9, note: 10,
};

/**
 * Columns read from the sheet's header row (No | No. Layout | LINE | Process Name | Machine Line | SETUP | Note), so cells
 * merged over several columns (LINE over two, Process Name over four in the Setup Plan sheet) still line up; null when
 * the row is not that header.
 */
export function headerColumns(header: string[]): SheetColumns | null {
  const at = (re: RegExp, from = 0) => header.findIndex((c, i) => i >= from && re.test(c));
  const process = at(/^process\s*name$/i);
  const machine = at(/^machine/i);
  const setup = at(/^setup$/i, machine + 1);
  if (process < 0 || machine < 0 || setup < 0) return null;
  const line = at(/^line$/i);
  const layout = at(/layout/i);
  const note = at(/^note$/i, setup + 3);
  return {
    no: Math.max(at(/^no\.?$/i), 0), layout: layout < 0 ? 1 : layout, line: line < 0 ? 2 : line, process,
    machinePlan: machine, machineState: machine + 1, machineNote: machine + 2,
    setupPlan: setup, setupFinish: setup + 1, setupStatus: setup + 2, note: note < 0 ? setup + 3 : note,
  };
}

/**
 * Without a header: rows wider than the 11 sheet columns lose the columns that are empty in every row, from the left
 * after LINE (where merged LINE and Process Name cells leave them), until 11 are left.
 */
function withoutMergeGaps(rows: string[][]): string[][] {
  const width = Math.max(0, ...rows.map((r) => r.length));
  if (width <= SHEET_COLUMNS) return rows;
  const drop = new Set<number>();
  for (let k = 3; k < width && drop.size < width - SHEET_COLUMNS; k++)
    if (rows.every((r) => !(r[k] ?? ""))) drop.add(k);
  return rows.map((r) => r.filter((_, k) => !drop.has(k)));
}

/**
 * Rows copied from the Setup Plan sheet: No | No. Layout | LINE | Process Name | Machine Plan | State | Note | Setup Plan |
 * Finish | Status | Note. Header lines are skipped (and used to find the columns when copied with the rows), an empty
 * LINE takes the one above, and the last Note (the problem text) comes back per row. Text without tabs is split at runs
 * of two or more spaces.
 */
export function parseSetupPaste(text: string, today: string): PastedSetup {
  let records = text.includes("\t") ? tsvRecords(text) : text.replace(/\r\n?/g, "\n").split("\n").map(cellsOf);
  const headerAt = records.findIndex((r) => r.some((c) => /^process\s*name$/i.test(c)));
  const fromHeader = headerAt >= 0 ? headerColumns(records[headerAt]) : null;
  const col = fromHeader ?? PLAIN_COLUMNS;
  if (!fromHeader && text.includes("\t")) {
    const data = records.filter((r) => r.some(Boolean));
    const fixed = withoutMergeGaps(data);
    records = records.map((r) => (r.some(Boolean) ? fixed[data.indexOf(r)] : r));
  }
  const rows: PastedSetup["rows"] = [];
  let line = "";
  for (let i = 0; i < records.length; i++) {
    const c = records[i];
    const get = (k: number) => c[k] ?? "";
    const process = get(col.process);
    if (!process || /^process\s*name$/i.test(process)) continue;
    const at = `Line ${i + 1}`;
    const no = /^\d+$/.test(get(col.no)) ? Number(get(col.no)) : 0;
    if (get(col.line)) line = get(col.line);
    const machineState = parseState(get(col.machineState));
    const setupStatus = parseState(get(col.setupStatus));
    if (machineState === null) return { rows: [], problem: `${at}: machine state "${get(col.machineState)}" is not O, X or Δ.` };
    if (setupStatus === null) return { rows: [], problem: `${at}: setup status "${get(col.setupStatus)}" is not O, X or Δ.` };
    const dates: string[] = [];
    for (const [k, label] of [[col.machinePlan, "machine plan"], [col.setupPlan, "setup plan"], [col.setupFinish, "setup finish"]] as const) {
      const d = parseSheetDate(get(k), today);
      if (d === null) return { rows: [], problem: `${at}: ${label} "${get(k)}" is not a date.` };
      dates.push(d);
    }
    const row = {
      no, layoutNo: get(col.layout), line, processName: process, machinePlan: dates[0], machineState, machineNote: get(col.machineNote),
      setupPlan: dates[1], setupFinish: dates[2], setupStatus, note: get(col.note).replace(/^[-–•]\s*/, ""),
    };
    const problem = setupProblem(row);
    if (problem) return { rows: [], problem: `${at}: ${problem}` };
    rows.push(row);
  }
  return {
    rows,
    problem: rows.length ? "" : "No rows found. Copy the rows from No. to Note in Excel (with the header rows if you like) and paste them here.",
  };
}
