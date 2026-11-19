// Standard names of the Dev Plan: the stages an event can be (ER, DV1, PV ...) and the tasks every model does (material in,
// line ready, sample shipment ...), each with the other names people have written for it, and the tasks each stage starts
// with. Kept on the server (one list for every model, Setting › Dev Plan standard names); the default below until saved.
import { getJson, toolPath, toolSend } from "../../toolsApi";

export interface VocabItem {
  /** Fixed key kept in the plan (stage or task code), e.g. "material-in". */
  code: string;
  /** Name shown and written into the plan. */
  name: string;
  /** Other names written for the same thing, matched case-insensitively inside a name (for tidying up old plans). */
  aliases: string[];
}

export interface PlanVocab {
  stages: VocabItem[];
  tasks: VocabItem[];
  /** Stage code -> the task codes an event of that stage starts with, in order. */
  templates: Record<string, string[]>;
}

const item = (code: string, name: string, aliases: string[] = []): VocabItem => ({ code, name, aliases });

export const DEFAULT_VOCAB: PlanVocab = {
  stages: [item("ER", "ER"), item("DV1", "DV1"), item("DV2", "DV2"), item("PV", "PV"), item("LPR", "LPR"), item("MP", "MP")],
  tasks: [
    item("design-confirm", "Design Confirm", ["design confirm", "설계 확정", "설계확정"]),
    item("mold-order", "Mold order", ["금형발주", "금형 발주", "request mold", "mold order"]),
    item("jig-request", "JIG/MC request", ["jig 제작의뢰", "jig 제작 의뢰", "jig 의뢰", "request jig", "jig/mc request"]),
    item("jig-receive", "JIG/MC receive", ["jig/mc receive", "jig receive", "jig 입고", "지그 입고"]),
    item("material-in", "Material in", ["자재 도착", "자재도착", "자재입고완료", "자재 입고 완료", "자재입고", "자재 입고", "material in", "material arrival"]),
    item("line-ready", "Line ready", ["개발라인 준비", "개발라인준비", "라인준비", "라인 준비", "line ready"]),
    item("setup", "Setup", ["setup", "셋팅시작", "셋팅 시작", "셋팅", "세팅"]),
    item("sample-make", "Sample making", ["make sample", "시료 제작", "시료제작", "샘플 제작", "제작"]),
    item("performance-check", "Performance check", ["성능검토", "성능 검토", "performance check"]),
    item("sample-send", "Sample shipment", ["send sample", "샘플발송", "샘플 발송", "시료 발송", "시료발송"]),
    item("approval-sample", "Approval sample shipment", ["승인시료 발송", "승인 시료 발송", "승인시료", "승인 시료", "approval sample"]),
    item("business-trip", "Business trip", ["business trip", "출장"]),
    item("audit", "Audit", ["audit"]),
  ],
  templates: {
    ER: ["design-confirm", "mold-order", "jig-request", "material-in", "line-ready", "setup", "sample-make", "sample-send"],
    DV1: ["material-in", "line-ready", "sample-make", "sample-send"],
    DV2: ["material-in", "line-ready", "sample-make", "sample-send"],
    PV: ["jig-receive", "setup", "material-in", "line-ready", "performance-check", "sample-make", "sample-send", "approval-sample"],
    LPR: ["audit"],
    MP: [],
  },
};

/** A name compared without case, spaces and punctuation ("Make Sample" = "make-sample"). */
const nameKey = (s: string) => s.toLowerCase().replace(/[^\p{L}\p{N}]+/gu, "");

/** More names a built-in task goes by in a list saved with codes of its own. */
const BUILTIN_NAMES: Record<string, string[]> = {
  "material-in": ["material", "materials", "자재"],
  "sample-make": ["sample make", "sample making"],
  "sample-send": ["sample send", "sample shipment"],
};

/**
 * The built-in task (a code of the default list) a task code of `items` stands for: the code itself when it is one,
 * else the built-in task one of its code, name or other names is ("Make Sample" / "샘플 제작" is sample-make); none
 * when it is no built-in task.
 */
export function builtinTask(code: string | undefined, items: VocabItem[]): string | undefined {
  if (!code) return undefined;
  if (DEFAULT_VOCAB.tasks.some((d) => d.code === code)) return code;
  const it = items.find((i) => i.code === code);
  const own = new Set([code, ...(it ? [it.name, ...it.aliases] : [])].map(nameKey).filter(Boolean));
  return DEFAULT_VOCAB.tasks.find((d) => [d.code, d.name, ...d.aliases, ...(BUILTIN_NAMES[d.code] ?? [])].some((n) => own.has(nameKey(n))))?.code;
}

/** A saved list read leniently; nothing saved (null) or nothing usable = the default. */
export function readVocab(v: unknown): PlanVocab {
  const o = (v && typeof v === "object" ? v : {}) as Record<string, unknown>;
  const str = (x: unknown) => (typeof x === "string" ? x.trim() : "");
  const items = (x: unknown): VocabItem[] => {
    const seen = new Set<string>();
    return (Array.isArray(x) ? x : []).flatMap((i) => {
      const r = (i && typeof i === "object" ? i : {}) as Record<string, unknown>;
      const code = str(r.code), name = str(r.name);
      if (!code || !name || seen.has(code.toLowerCase())) return [];
      seen.add(code.toLowerCase());
      const aliases = [...new Set((Array.isArray(r.aliases) ? r.aliases : []).map(str).filter(Boolean))];
      return [{ code, name, aliases }];
    });
  };
  const stages = items(o.stages), tasks = items(o.tasks);
  if (!stages.length && !tasks.length) return DEFAULT_VOCAB;
  const taskCodes = new Set(tasks.map((t) => t.code));
  const raw = (o.templates && typeof o.templates === "object" ? o.templates : {}) as Record<string, unknown>;
  const templates = Object.fromEntries(stages.map((s) => [s.code,
    [...new Set((Array.isArray(raw[s.code]) ? raw[s.code] as unknown[] : []).map(str).filter((c) => taskCodes.has(c)))]]));
  return { stages, tasks, templates };
}

/**
 * Server: GET {vocab: PlanVocab | null} (null until saved) for every Working user;
 * PUT {vocab} (Admin only) saves it cleaned up (blank or repeated codes dropped, templates of known codes only) and returns it.
 */
export const planVocabApi = {
  get: () => getJson<{ vocab: PlanVocab | null }>(toolPath("working/plan-vocab")).then((r) => readVocab(r.vocab)),
  save: (vocab: PlanVocab) => toolSend<{ vocab: PlanVocab }>("working/plan-vocab", { vocab }, "PUT").then((r) => readVocab(r.vocab)),
};
