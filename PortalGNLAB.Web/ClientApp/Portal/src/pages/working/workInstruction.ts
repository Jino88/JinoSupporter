import { loginUrl } from "../../api";
import { toolPath, toolSend } from "../../toolsApi";

/** The Work Instructions workbook of a model as the server read it: one process per sheet. */
export interface InstructionDoc {
  fileName: string;
  uploadedAt: string;
  uploadedBy: string;
  size: number;
  revisions: InstructionRevision[];
  processes: InstructionProcess[];
}
export interface InstructionRevision { date: string; page: string; rev: string; content: string; author: string }
export interface InstructionPart { no: string; name: string; spec: string; qty: string }
export interface InstructionTool { no: string; name: string; note: string }
export interface InstructionStep { no: number; text: string }
export interface InstructionProcess {
  sheet: string;
  code: string;
  name: string;
  standardTime: number | null;
  spareTime: number | null;
  page: string;
  ctq: boolean;
  parts: InstructionPart[];
  tools: InstructionTool[];
  steps: InstructionStep[];
  keyPoints: string[];
}

/** Tool names (the first column of Equipment & tools) that are the machine or the jig of the process; the other rows are supplies and settings. */
export const EQUIPMENT_NAMES = ["Thiết bị", "Máy", "장비", "설비", "Equipment", "Machine", "JIG", "지그"];

const base = (name: string) => `working/models/${encodeURIComponent(name)}/instructions`;

export const instructionFileUrl = (name: string) => toolPath(`${base(name)}/file`);

/** The model's workbook, or null when none is uploaded. */
export async function loadInstructions(name: string): Promise<InstructionDoc | null> {
  const response = await fetch(toolPath(base(name)), { credentials: "same-origin", headers: { Accept: "application/json" } });
  if (response.status === 404) return null;
  if (response.status === 401) window.location.assign(loginUrl(location.pathname));
  if (!response.ok) throw new Error((await response.json().catch(() => null))?.error ?? `Request failed (HTTP ${response.status})`);
  return (await response.json()) as InstructionDoc;
}

/** Sends the workbook (multipart "file"); the server reads it and answers with what it found. */
export async function uploadInstructions(name: string, file: File): Promise<InstructionDoc> {
  const form = new FormData();
  form.append("file", file, file.name);
  const response = await fetch(toolPath(base(name)), {
    method: "POST",
    credentials: "same-origin",
    headers: { Accept: "application/json", "X-Requested-With": "PortalGNLAB" },
    body: form,
  });
  if (response.status === 401) window.location.assign(loginUrl(location.pathname));
  if (!response.ok)
    throw new Error((await response.json().catch(() => null))?.error ?? (response.status === 413 ? "The workbook is too large (200 MB at most)." : `Request failed (HTTP ${response.status})`));
  return (await response.json()) as InstructionDoc;
}

export const deleteInstructions = (name: string) => toolSend<void>(base(name), undefined, "DELETE");

/** A time in seconds as written in the workbook: up to 3 decimals without trailing zeros; "" when there is none. */
export const timeText = (value: number | null) => (value === null ? "" : String(Math.round(value * 1000) / 1000));

/** The machine and jig of a process: the notes of its Equipment & tools rows named like a machine or a jig (the name itself when there is no note). */
export function equipmentText(p: InstructionProcess): string {
  const names = p.tools
    .filter((t) => EQUIPMENT_NAMES.some((n) => t.name.trim().toLowerCase().startsWith(n.toLowerCase())))
    .map((t) => (t.note.trim() || t.name.trim()));
  return [...new Set(names)].join(" / ");
}

export const sizeText = (bytes: number) => (bytes >= 1048576 ? `${(bytes / 1048576).toFixed(1)} MB` : `${Math.max(1, Math.round(bytes / 1024))} KB`);

/** yyyy-MM-dd HH:mm of an ISO time (as local time), "" when it is not a time. */
export function whenText(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  const p = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}
