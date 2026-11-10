// The process names and NG names of a Daily NG Rate sheet, read from its header rows copied in Excel (the process row
// and the Input / NG items / NG rate (%) row), for the model group entered by hand. Same rules as
// tools/working_process_ng_cli.py.
import { blankProcess, type ManualGroup, type ManualProcess } from "./manual";
import { tsvRecords } from "./setup";

export interface SheetProcess {
  name: string;
  localName: string;
  ngItems: string[];
}

export interface SheetHeader {
  processes: SheetProcess[];
  warnings: string[];
  /** Why nothing could be read, or "". */
  problem: string;
}

const flat = (s: string) => s.split(/\s+/).filter(Boolean).join(" ");
const key = (s: string) => flat(s).toLowerCase();
const isInput = (s: string) => s.trim().toLowerCase() === "input";
const isRate = (s: string) => /^ng\s*rate\b/i.test(s);

/** "N. NAME" with the second line of the cell as the local name. */
function process(cell: string): { name: string; localName: string } {
  const lines = cell.split("\n").map(flat).filter(Boolean);
  const name = (lines[0] ?? "").replace(/^\d+\s*[.)]\s*/, "").trim();
  return { name: name || (lines[0] ?? ""), localName: lines.slice(1).join(" ") };
}

/** NG items of a process: every cell but Input and NG rate (%), once each (any case). */
function items(cells: string[]): string[] {
  const out: string[] = [];
  for (const c of cells) {
    const t = flat(c);
    if (!t || isInput(t) || isRate(t)) continue;
    if (!out.some((x) => key(x) === key(t))) out.push(t);
  }
  return out;
}

/** The processes (name, local name, NG items) of the copied header rows. */
export function parseNgHeader(text: string): SheetHeader {
  const rows = tsvRecords(text).filter((r) => r.some((c) => c.trim()));
  const at = rows.findIndex((r) => r.some(isInput));
  if (at < 1) {
    return { processes: [], warnings: [], problem: "No row with Input cells under a process row. Copy the process row and the Input / NG / NG rate (%) row together." };
  }
  const head = rows[at - 1];
  const cells = rows[at];
  const first = cells.findIndex(isInput);
  const procs = head.map((c, i) => [i, c] as const).filter(([i, c]) => c.trim() && i >= first);
  if (!procs.length) return { processes: [], warnings: [], problem: "The process row has no process names above the Input cells." };

  const starts = cells.flatMap((c, i) => (isInput(c) ? [i] : []));
  const warnings: string[] = [];
  let spans: [number, number][];
  if (starts.length === procs.length) {
    // one Input per process: the NG items are the cells up to the next Input
    spans = starts.map((s, k) => [s, starts[k + 1] ?? cells.length]);
  } else {
    // some process has no Input (or one more): the columns under each process cell
    warnings.push(`${procs.length} processes but ${starts.length} Input cells; NG items were taken from the columns under each process name.`);
    const cols = procs.map(([i]) => i);
    spans = cols.map((c, k) => [c, cols[k + 1] ?? Math.max(cells.length, c + 1)]);
  }

  const processes = procs.map(([, cell], k) => {
    const p = { ...process(cell), ngItems: items(cells.slice(spans[k][0], spans[k][1])) };
    if (!p.ngItems.length) warnings.push(`${p.name} has no NG item.`);
    return p;
  });
  const seen = new Map<string, number>();
  for (const p of processes) seen.set(key(p.name), (seen.get(key(p.name)) ?? 0) + 1);
  for (const p of processes) {
    const n = seen.get(key(p.name)) ?? 0;
    if (n > 1) { warnings.push(`${p.name} appears ${n} times.`); seen.delete(key(p.name)); }
  }
  return { processes, warnings, problem: "" };
}

/**
 * The group with the sheet's processes in their order. The material list stays; a process with the same name as one
 * already there keeps its id, type and materials. Also returns the names of the processes dropped.
 */
export function mergeSheet(g: ManualGroup, sheet: SheetProcess[]): { group: ManualGroup; dropped: string[] } {
  const used = new Set<number>();
  const processes = sheet.map((s): ManualProcess => {
    const at = g.processes.findIndex((o, i) => !used.has(i) && key(o.name) === key(s.name));
    if (at >= 0) used.add(at);
    const old = at >= 0 ? g.processes[at] : blankProcess();
    return { ...old, name: s.name, localName: s.localName || old.localName, ngItems: s.ngItems };
  });
  const dropped = g.processes.filter((_, i) => !used.has(i)).map((p) => p.name);
  return { group: { materials: g.materials, processes }, dropped };
}
