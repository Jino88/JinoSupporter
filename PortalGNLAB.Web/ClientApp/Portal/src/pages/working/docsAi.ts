// Model Info › Documents: the AI makes the Process table (processes and materials) from the Work Instructions and fills the
// JIG / MC List from a JIG / Machine / Tester workbook. Pure helpers (summary text, replace confirmation, run order) live here.
import { loginUrl } from "../../api";
import { toolPath, toolSend } from "../../toolsApi";
import { aiAnswer, readManual, type ManualGroup } from "./manual";
import { jigTotals, type ProcessRow } from "./processList";

export type ProcessAiResponse = { group: unknown; note: string; model: string };

/** AI makes the process table from the model's uploaded Work Instructions (not saved here). */
export const manualAiFromInstructions = (name: string, group: ManualGroup, signal?: AbortSignal) =>
  toolSend<ProcessAiResponse>(`working/models/${encodeURIComponent(name)}/manual/ai-from-instructions`, { group }, "POST", signal);

export type ProcessResult = { processes: number; materials: number; note: string; model: string; saved: boolean; group?: ManualGroup };
export type JigResult = { processes: number; jigKinds: number; note: string; unmatched: number; saved: boolean; rows?: ProcessRow[] };
export type DocsResult = {
  process?: ProcessResult; processError?: string;
  jig?: JigResult; jigError?: string;
};

/** What the AI answered for the process table, read against the current one. */
export function readProcessAnswer(current: ManualGroup, r: ProcessAiResponse): { group: ManualGroup; note: string; model: string; empty: boolean } {
  const { group } = aiAnswer(current, r.group);
  return { group, note: r.note ?? "", model: r.model ?? "", empty: group.processes.length === 0 };
}

/** Whether the old content has to be confirmed before the new replaces it. */
export const needsReplaceConfirm = (hasOld: boolean, incomingEmpty: boolean) => hasOld && !incomingEmpty;

export const replaceProcessText = (oldCount: number) => `Replace the ${oldCount} existing processes with the AI's Process table?`;
export const replaceJigText = (oldCount: number) => `Replace the ${oldCount} rows of the JIG/MC List with the AI's result?`;

export const processSummary = (r: ProcessResult) => `Process: ${r.processes} processes, ${r.materials} materials`;
export const jigSummary = (r: JigResult) => `JIG/MC List: ${r.processes} processes, ${r.jigKinds} JIG kinds`;
export const jigCounts = (rows: ProcessRow[]) => ({ processes: rows.length, jigKinds: jigTotals(rows).kinds });

export const errorText = (e: unknown) => (e instanceof Error ? e.message : String(e));
export const NOTHING_TO_DO = "Upload the Work Instructions, or choose a JIG / Machine / Tester workbook (.xlsx / .xlsm), first.";
export const isWorkbook = (file: { name: string } | null | undefined) => !!file && /\.(xlsx|xlsm)$/i.test(file.name);

type ProcessSteps = {
  hasInstructions: boolean;
  currentGroup: ManualGroup;
  ask: (group: ManualGroup) => Promise<ProcessAiResponse>;
  confirm: (text: string) => boolean;
  save: (group: ManualGroup) => Promise<void>;
};

/** Process: ask the AI, confirm replacing a non-empty table, then save; the order is always ask → confirm → save. */
export async function runProcess(s: ProcessSteps): Promise<ProcessResult> {
  const a = readProcessAnswer(s.currentGroup, await s.ask(s.currentGroup));
  const result = { processes: a.group.processes.length, materials: a.group.materials.length, note: a.note, model: a.model, saved: false, group: a.group };
  if (a.empty) return result;
  if (needsReplaceConfirm(s.currentGroup.processes.length > 0, a.empty) && !s.confirm(replaceProcessText(s.currentGroup.processes.length))) return result;
  await s.save(a.group);
  return { ...result, saved: true };
}

type JigSteps = {
  match: () => Promise<{ rows: ProcessRow[]; note: string; unmatched: string[] }>;
  current: () => Promise<{ rows: ProcessRow[]; revision: number }>;
  confirm: (text: string) => boolean;
  save: (rows: ProcessRow[], revision: number) => Promise<void>;
};

/** JIG/MC List: AI match, read the current revision, confirm replacing existing rows, then save. */
export async function runJig(s: JigSteps): Promise<JigResult> {
  const m = await s.match();
  const counts = jigCounts(m.rows);
  const result = { ...counts, note: m.note ?? "", unmatched: m.unmatched?.length ?? 0, saved: false, rows: m.rows };
  if (m.rows.length === 0) return result;
  const cur = await s.current();
  if (cur.rows.length > 0 && !s.confirm(replaceJigText(cur.rows.length))) return result;
  await s.save(m.rows, cur.revision);
  return { ...result, saved: true };
}

// ---- several documents picked at once ----

export type DocKind = "instructions" | "jig-list" | "unknown";
export const KIND_TEXT: Record<DocKind, string> = { instructions: "Work Instructions", "jig-list": "JIG / Machine / Tester list", unknown: "not recognized" };

/** The largest workbook the server reads (as the Work Instructions upload: 200 MB). */
export const MAX_DOC_BYTES = 200 * 1024 * 1024;

export type FoundKind = { kind: DocKind; reason?: string };

/** What the server finds a workbook to be (nothing is saved); a workbook too large is not sent. */
export async function docKind(name: string, file: File): Promise<FoundKind> {
  if (file.size > MAX_DOC_BYTES) return { kind: "unknown", reason: `larger than ${MAX_DOC_BYTES / 1024 / 1024} MB` };
  const form = new FormData();
  form.append("file", file, file.name);
  let response: Response;
  try {
    response = await fetch(toolPath(`working/models/${encodeURIComponent(name)}/documents/kind`), {
      method: "POST", credentials: "same-origin", headers: { Accept: "application/json", "X-Requested-With": "PortalGNLAB" }, body: form,
    });
  } catch (e) {
    // the browser says only "Failed to fetch" when the connection is cut (the server stopped, or refused the upload)
    throw new Error(`${file.name} could not be sent to the server (${errorText(e)}). Check that the server is running, then try again.`);
  }
  if (response.status === 401) window.location.assign(loginUrl(location.pathname));
  const body = await response.json().catch(() => null) as { kind?: string; reason?: string; error?: string } | null;
  if (!response.ok) throw new Error(`${file.name}: ${body?.error ?? `Request failed (HTTP ${response.status})`}`);
  const kind = body?.kind === "instructions" || body?.kind === "jig-list" ? body.kind : "unknown";
  return body?.reason ? { kind, reason: body.reason } : { kind };
}

/**
 * The files picked, by what they are: the Work Instructions and the JIG / Machine / Tester list used (the first of each),
 * and the names of the others (a second file of a kind, or one that is neither) with why they are not used.
 */
export function sortDocs<F extends { name: string }>(found: ({ file: F } & FoundKind)[]): { instructions?: F; jig?: F; skipped: string[] } {
  const out: { instructions?: F; jig?: F; skipped: string[] } = { skipped: [] };
  for (const { file, kind, reason } of found) {
    if (kind === "instructions" && !out.instructions) out.instructions = file;
    else if (kind === "jig-list" && !out.jig) out.jig = file;
    else out.skipped.push(`${file.name} (${kind === "unknown" ? reason || KIND_TEXT.unknown : `another ${KIND_TEXT[kind]}`})`);
  }
  return out;
}

/** A process as the documents tie it together: its materials (Process) and its machines and JIGs (JIG/MC List). */
export type DocLink = { no: string; process: string; materials: string[]; machine: string; jigs: string[]; inProcess: boolean; inJig: boolean };

const nameKey = (s: string) => s.trim().toLowerCase().replace(/\s+/g, " ");

/**
 * How the documents relate: the JIG/MC List rows (in the line order of the Work Instructions) with the materials of the
 * Process of the same name, then the processes of the Process table that no JIG/MC row has.
 */
export function docLinks(group: ManualGroup | undefined, rows: ProcessRow[] | undefined): DocLink[] {
  const procs = group?.processes ?? [];
  const byName = new Map(procs.map((p) => [nameKey(p.name), p]));
  const used = new Set<string>();
  const links: DocLink[] = (rows ?? []).map((r) => {
    const p = byName.get(nameKey(r.process));
    if (p) used.add(p.id);
    return {
      no: r.no, process: r.process, materials: p ? p.materials.map((u) => u.material) : [], machine: r.machine,
      jigs: (r.jigs ?? []).map((j) => j.name), inProcess: !!p, inJig: true,
    };
  });
  for (const p of procs) {
    if (!used.has(p.id)) links.push({ no: "", process: p.name, materials: p.materials.map((u) => u.material), machine: "", jigs: [], inProcess: true, inJig: false });
  }
  return links;
}

export { readManual };
