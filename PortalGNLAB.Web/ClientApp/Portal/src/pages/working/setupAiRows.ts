// AI input of the Setup Plan (SetupAi.tsx): the request sent, the answer read leniently, and the rows previewed as New / Change.
import type { SetupFields, SetupRow, SetupState } from "./setup";
import { setupFieldsOf } from "./setup";

export interface SetupAiAnswer {
  /** id "" = a new row, else the row to change. */
  rows: (SetupFields & { id: string })[];
  note: string;
  model: string;
}

export function setupAiRequest(text: string, today: string, rows: SetupRow[], images: { name: string; dataUrl: string }[] = []) {
  return {
    text,
    today,
    rows: rows.map((r) => ({
      id: r.id, no: r.no, layoutNo: r.layoutNo, line: r.line, processName: r.processName, machinePlan: r.machinePlan,
      machineState: r.machineState, machineNote: r.machineNote, setupPlan: r.setupPlan, setupFinish: r.setupFinish, setupStatus: r.setupStatus,
    })),
    images,
  };
}

const str = (v: unknown) => (typeof v === "string" ? v : typeof v === "number" ? String(v) : "");
const state = (v: unknown): SetupState => { const s = str(v).trim(); return s === "O" || s === "X" || s === "Δ" ? s : ""; };

/** The server answer as rows; anything that is not a row is left out and a missing field is "". */
export function readSetupAnswer(v: unknown): SetupAiAnswer {
  const o = v && typeof v === "object" ? (v as Record<string, unknown>) : {};
  const rows = (Array.isArray(o.rows) ? o.rows : []).flatMap((r): SetupAiAnswer["rows"] => {
    if (!r || typeof r !== "object") return [];
    const x = r as Record<string, unknown>;
    const no = Number(x.no);
    return [{
      id: str(x.id).trim(), no: Number.isInteger(no) && no > 0 ? no : 0, layoutNo: str(x.layoutNo).trim(), line: str(x.line).trim(),
      processName: str(x.processName).trim(), machinePlan: str(x.machinePlan).trim(), machineState: state(x.machineState),
      machineNote: str(x.machineNote).trim(), setupPlan: str(x.setupPlan).trim(), setupFinish: str(x.setupFinish).trim(), setupStatus: state(x.setupStatus),
    }];
  });
  return { rows, note: str(o.note).trim(), model: str(o.model).trim() };
}

export interface SetupAiItem {
  key: string;
  /** The row it changes, or null for a new row. */
  row: SetupRow | null;
  fields: SetupFields;
  title: string;
  detail: string;
}

const LABEL: [keyof SetupFields, string][] = [
  ["no", "No."], ["layoutNo", "Layout"], ["line", "Line"], ["processName", "Process"], ["machinePlan", "Machine plan"],
  ["machineState", "Machine"], ["machineNote", "Machine note"], ["setupPlan", "Setup plan"], ["setupFinish", "Setup finish"], ["setupStatus", "Setup"],
];

/**
 * One item per answered row. A row whose id is a current row is a change (an answer with nothing different is left out);
 * any other id is a new row.
 */
export function setupAiItems(answer: SetupAiAnswer, rows: SetupRow[]): SetupAiItem[] {
  const byId = new Map(rows.map((r) => [r.id, r]));
  const items: SetupAiItem[] = [];
  answer.rows.forEach((a, i) => {
    const { id, ...fields } = a;
    const row = byId.get(id) ?? null;
    const title = `${fields.line ? fields.line + " · " : ""}${fields.processName || row?.processName || ""}`;
    if (!row) {
      const detail = LABEL.filter(([k]) => !["line", "processName"].includes(k) && fields[k] !== "" && fields[k] !== 0).map(([k, l]) => `${l} ${fields[k]}`).join(" · ");
      items.push({ key: `row.${i}`, row: null, fields, title, detail });
      return;
    }
    // a changed row keeps its own process name when the answer left it out
    const next: SetupFields = { ...fields, processName: fields.processName || row.processName, no: fields.no || row.no };
    const before = setupFieldsOf(row);
    const diff = LABEL.filter(([k]) => next[k] !== before[k]);
    if (!diff.length) return;
    items.push({ key: `row.${i}`, row, fields: next, title, detail: diff.map(([k, l]) => `${l} ${before[k] === "" ? "–" : before[k]} → ${next[k] === "" ? "–" : next[k]}`).join(" · ") });
  });
  return items;
}

/** The picked items split into the rows to add and the rows to change. */
export function splitSetupAi(items: SetupAiItem[], picked: Set<string>) {
  const kept = items.filter((i) => picked.has(i.key));
  return {
    added: kept.filter((i) => !i.row).map((i) => i.fields),
    changed: kept.flatMap((i) => (i.row ? [{ row: i.row, fields: i.fields }] : [])),
  };
}
