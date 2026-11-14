// Model Info › Documents: the AI makes the Process table (processes and materials) from the Work Instructions and fills the
// JIG / MC List from a JIG / Machine / Tester workbook. Pure helpers (summary text, replace confirmation, run order) live here.
import { toolSend } from "../../toolsApi";
import { aiAnswer, readManual, type ManualGroup } from "./manual";
import { jigTotals, type ProcessRow } from "./processList";

export type ProcessAiResponse = { group: unknown; note: string; model: string };

/** AI makes the process table from the model's uploaded Work Instructions (not saved here). */
export const manualAiFromInstructions = (name: string, group: ManualGroup, signal?: AbortSignal) =>
  toolSend<ProcessAiResponse>(`working/models/${encodeURIComponent(name)}/manual/ai-from-instructions`, { group }, "POST", signal);

export type ProcessResult = { processes: number; materials: number; note: string; model: string; saved: boolean };
export type JigResult = { processes: number; jigKinds: number; note: string; unmatched: number; saved: boolean };
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
  const result = { processes: a.group.processes.length, materials: a.group.materials.length, note: a.note, model: a.model, saved: false };
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
  const result = { ...counts, note: m.note ?? "", unmatched: m.unmatched?.length ?? 0, saved: false };
  if (m.rows.length === 0) return result;
  const cur = await s.current();
  if (cur.rows.length > 0 && !s.confirm(replaceJigText(cur.rows.length))) return result;
  await s.save(m.rows, cur.revision);
  return { ...result, saved: true };
}

export { readManual };
