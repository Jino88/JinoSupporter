import { useEffect, useState } from "react";
import { DEFAULT_VOCAB, planVocabApi, type PlanVocab } from "./working/planVocab";

/** Setting › Dev Plan standard names: the stages, the tasks (each with the other names written for it) and the tasks each stage starts with. */

type Msg = { text: string; error: boolean } | null;

/** One table row: aliases and the stage template are edited as comma-separated text. */
export interface VocabRow {
  code: string;
  name: string;
  aliases: string;
  /** Stages only: the task codes the stage starts with, in order. */
  template: string;
}

const errText = (e: unknown) => (e instanceof Error ? e.message : String(e));

export const splitList = (text: string): string[] => text.split(/[,;\n]/).map((s) => s.trim()).filter(Boolean);

export function toRows(vocab: PlanVocab): { stages: VocabRow[]; tasks: VocabRow[] } {
  return {
    stages: vocab.stages.map((s) => ({ code: s.code, name: s.name, aliases: s.aliases.join(", "), template: (vocab.templates[s.code] ?? []).join(", ") })),
    tasks: vocab.tasks.map((t) => ({ code: t.code, name: t.name, aliases: t.aliases.join(", "), template: "" })),
  };
}

/** The rows with one (dragged) moved to the place of another (dropped on): after it going down, before it going up. */
export function moveRow<T>(rows: T[], from: number, to: number): T[] {
  if (from === to || from < 0 || to < 0 || from >= rows.length || to >= rows.length) return rows;
  const out = [...rows];
  const [row] = out.splice(from, 1);
  out.splice(to, 0, row);
  return out;
}

export function toVocab(stages: VocabRow[], tasks: VocabRow[]): PlanVocab {
  const item = (r: VocabRow) => ({ code: r.code.trim(), name: r.name.trim(), aliases: splitList(r.aliases) });
  return {
    stages: stages.map(item),
    tasks: tasks.map(item),
    templates: Object.fromEntries(stages.filter((s) => s.code.trim()).map((s) => [s.code.trim(), splitList(s.template)])),
  };
}

export default function PlanVocabCard() {
  const [vocab, setVocab] = useState<PlanVocab>();
  const [loadError, setLoadError] = useState("");

  useEffect(() => {
    planVocabApi.get().then(setVocab).catch((e) => setLoadError(errText(e)));
  }, []);

  return (
    <section className="card">
      <h2>Dev Plan standard names</h2>
      {loadError ? <div className="alert">{loadError}</div> : vocab ? <PlanVocabEditor initial={vocab} /> : <div className="muted">Loading...</div>}
    </section>
  );
}

export function PlanVocabEditor({ initial }: { initial: PlanVocab }) {
  const start = toRows(initial);
  const [stages, setStages] = useState(start.stages);
  const [tasks, setTasks] = useState(start.tasks);
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState<Msg>(null);
  /** The row being dragged by its grip, and the row it is over, within one table. */
  const [drag, setDrag] = useState<{ kind: string; from: number; over: number } | null>(null);

  const load = (v: PlanVocab) => {
    const r = toRows(v);
    setStages(r.stages);
    setTasks(r.tasks);
  };

  async function save() {
    setBusy(true);
    setMsg(null);
    try {
      // the server drops blank / repeated codes and unknown template codes; show what it actually kept
      const saved = await planVocabApi.save(toVocab(stages, tasks));
      load(saved);
      setMsg({ text: `${saved.stages.length} stage(s) and ${saved.tasks.length} task(s) saved.`, error: false });
    } catch (e) {
      setMsg({ text: errText(e), error: true });
    } finally {
      setBusy(false);
    }
  }

  const table = (kind: "stage" | "task", rows: VocabRow[], setRows: (r: VocabRow[]) => void) => {
    const update = (i: number, patch: Partial<VocabRow>) => setRows(rows.map((r, j) => (j === i ? { ...r, ...patch } : r)));
    const isStage = kind === "stage";
    return (
      <table className="set-table" style={{ maxWidth: "none" }} aria-label={isStage ? "Stages" : "Tasks"}>
        <thead>
          <tr>
            <th style={{ width: "1rem" }}></th>
            <th style={{ width: "10rem" }}>Code</th>
            <th style={{ width: "12rem" }}>Name</th>
            <th>Other names (comma separated)</th>
            {isStage && <th>Starts with (task codes, in order)</th>}
            <th style={{ width: "6rem" }}></th>
          </tr>
        </thead>
        <tbody>
          {rows.length === 0 && (
            <tr>
              <td colSpan={isStage ? 6 : 5} className="muted small">None.</td>
            </tr>
          )}
          {rows.map((row, i) => (
            <tr key={i}
              className={drag?.kind === kind && drag.over === i && drag.from !== i ? (drag.from < i ? "drop-below" : "drop-above") : undefined}
              onDragOver={(x) => {
                if (drag?.kind !== kind) return;
                x.preventDefault();
                if (drag.over !== i) setDrag({ ...drag, over: i });
              }}
              onDrop={(x) => { x.preventDefault(); if (drag?.kind === kind) setRows(moveRow(rows, drag.from, i)); setDrag(null); }}>
              {/* only the grip drags, so text in the boxes can still be selected */}
              <td className="set-grip" title="Drag to move" aria-label={`Move ${kind} ${row.code || i + 1}`} draggable
                onDragStart={(x) => {
                  x.dataTransfer.effectAllowed = "move";
                  x.dataTransfer.setData("text/plain", row.code);
                  const tr = x.currentTarget.parentElement;
                  if (tr) x.dataTransfer.setDragImage(tr, 8, 8);
                  setDrag({ kind, from: i, over: i });
                }}
                onDragEnd={() => setDrag(null)}>⠿</td>
              <td><input className="inp sm mono" value={row.code} aria-label={`${kind} code`} onChange={(e) => update(i, { code: e.target.value })} /></td>
              <td><input className="inp sm" value={row.name} aria-label={`${kind} name`} onChange={(e) => update(i, { name: e.target.value })} /></td>
              <td><input className="inp sm" value={row.aliases} aria-label={`${kind} other names`} onChange={(e) => update(i, { aliases: e.target.value })} /></td>
              {isStage && (
                <td><input className="inp sm mono" value={row.template} aria-label="stage template" onChange={(e) => update(i, { template: e.target.value })} /></td>
              )}
              <td>
                <button type="button" className="btn sm danger" onClick={() => setRows(rows.filter((_, j) => j !== i))}>Remove</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    );
  };

  const blank: VocabRow = { code: "", name: "", aliases: "", template: "" };
  return (
    <>
      <p className="muted small">
        One list for every model. A code is the fixed key kept in a plan; the name is what is shown and written. Other names are matched inside a name
        (any case) to tidy up old plans. Save to apply; a blank or repeated code and a template code that is not a task are dropped.
      </p>
      <h3>Stages</h3>
      {table("stage", stages, setStages)}
      <div className="set-actions">
        <button type="button" className="btn sm" onClick={() => setStages([...stages, { ...blank }])}>Add stage</button>
      </div>
      <h3>Tasks</h3>
      {table("task", tasks, setTasks)}
      <div className="set-actions">
        <button type="button" className="btn sm" onClick={() => setTasks([...tasks, { ...blank }])}>Add task</button>
      </div>
      <div className="set-actions">
        <button type="button" className="btn sm primary" disabled={busy} onClick={() => void save()}>{busy ? "Saving..." : "Save"}</button>
        <button type="button" className="btn sm" disabled={busy} onClick={() => { load(DEFAULT_VOCAB); setMsg({ text: "Default filled in; press Save to keep it.", error: false }); }}>
          Reset to default
        </button>
        {msg?.text && <span className={"set-note " + (msg.error ? "err" : "ok")}>{msg.text}</span>}
      </div>
    </>
  );
}
