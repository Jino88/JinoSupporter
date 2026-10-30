import { Fragment, useMemo, type ReactNode } from "react";
import { NumInput } from "./fields";
import {
  SEGMENTS,
  WEEKDAYS,
  clearOverride,
  entryRows,
  firstOf,
  fmt,
  isWeekend,
  ngRate,
  ngTotals,
  pct,
  setInput,
  setPair,
  totalRows,
  viewPairs,
  weekday,
  type DayEntry,
  type Rows,
  type TotalRow,
} from "./logic";
import type { DayData, DayStatus, ModelSetup, Slot } from "./types";

const rateCell = (v: number) => (
  <span className={"pi-rate " + (v === 0 ? "z" : v >= 0.02 ? "hi" : "")}>{v === 0 ? "0" : pct(v)}</span>
);

export type DayView = "date" | "total";

/** Left: the dates of the window and their Total. Right: the chosen day's input, or the read-only Total. */
export default function DailyTab(props: {
  setup: ModelSetup;
  data: DayData;
  rows: Rows;
  slot: Slot;
  onSlot: (slot: Slot) => void;
  onData: (data: DayData) => void;
  view: DayView;
  date: string;
  dates: string[];
  /** Entered days of the window; the day on screen carries its unsaved input. */
  entries: DayEntry[];
  statusOf: Record<string, DayStatus>;
  unsaved: boolean;
  onPick: (date: string) => void;
  onTotal: () => void;
  onShift: (days: number) => void;
}) {
  const { setup, dates, entries, view } = props;
  const dayRates = useMemo(() => {
    const map: Record<string, ReturnType<typeof ngTotals>> = {};
    for (const e of entries) map[e.date] = ngTotals(setup, entryRows(setup, e));
    return map;
  }, [setup, entries]);
  const total = useMemo(() => totalRows(setup, entries), [setup, entries]);
  const totalRates = ngTotals(setup, total);
  const rateLine = (t: ReturnType<typeof ngTotals>) => `P ${pct(t.processAfter)} · F ${pct(t.functionAfter)}`;

  return (
    <div className="pi-daily">
      <aside className="card pi-days">
        <div className="pi-days-h">
          <button className="btn sm" title="Previous week" aria-label="Previous week" onClick={() => props.onShift(-7)}>←</button>
          <span>{dates[0].slice(5)} – {dates.at(-1)?.slice(5)}</span>
          <button className="btn sm" title="Next week" aria-label="Next week" onClick={() => props.onShift(7)}>→</button>
        </div>
        <button className={"pi-di total" + (view === "total" ? " on" : "")} onClick={props.onTotal}>
          <b>Total</b>
          <span className="muted small">{entries.length} entered day{entries.length === 1 ? "" : "s"}</span>
          <span className="pi-di-r">{entries.length ? rateLine(totalRates) : "—"}</span>
        </button>
        {dates.map((d) => {
          const st = props.statusOf[d];
          const r = dayRates[d];
          const here = d === props.date;
          return (
            <button key={d} className={"pi-di" + (view === "date" && here ? " on" : "") + (isWeekend(d) ? " we" : "")}
              onClick={() => props.onPick(d)}>
              <b>{d.slice(5)} <span>{WEEKDAYS[weekday(d)]}</span></b>
              <span className={"pi-di-st " + (st === "Submitted" ? "ok" : st ? "draft" : "none")}>
                {here && props.unsaved ? "Unsaved" : st === "Submitted" ? "Submitted" : st ? "Draft" : "Not entered"}
              </span>
              <span className="pi-di-r">{r ? rateLine(r) : ""}</span>
            </button>
          );
        })}
        <div className="pi-hint">P / F = process / function NG rate after rework.</div>
      </aside>
      {view === "total"
        ? <TotalTable setup={setup} total={total} days={entries.length} from={dates[0]} to={dates.at(-1) ?? ""} />
        : <DayTable {...props} />}
    </div>
  );
}

function TotalTable(props: { setup: ModelSetup; total: Record<string, TotalRow>; days: number; from: string; to: string }) {
  const { setup, total } = props;
  let lastSeg = "";
  return (
    <section className="card pi-sec">
      <div className="pi-sec-h">
        <h2>Total {props.from} – {props.to} <span className="cnt">{props.days} entered day{props.days === 1 ? "" : "s"}</span></h2>
        <span className="muted small">Σ of the entered days · NG % = Σ NG / Σ input · read only — pick a date on the left to edit.</span>
      </div>
      <div className="pi-wrap">
        <table className="pi-pe">
          <thead>
            <tr><th>#</th><th>Process</th><th>NG item</th><th className="n">Input</th><th className="n">Before</th><th className="n">After</th>
              <th className="n">NG % Before</th><th className="n">NG % After</th><th className="n">Output</th></tr>
          </thead>
          <tbody>
            {setup.processes.map((p, idx) => {
              const segHead = p.segment !== lastSeg;
              lastSeg = p.segment;
              const t = total[p.id];
              const withTotal = p.ngItems.length > 1;
              const n = Math.max(1, p.ngItems.length) + (withTotal ? 1 : 0);
              const rates = (b: number, a: number) => (
                <>
                  <td className="n">{rateCell(ngRate(b, t.input))}</td>
                  <td className="n">{rateCell(ngRate(a, t.input))}</td>
                </>
              );
              const head = (
                <>
                  <td className="pi-pno rs" rowSpan={n}>{idx + 1}</td>
                  <td className="pi-pname rs" rowSpan={n}>{p.name}{p.localName && <small>{p.localName}</small>}</td>
                </>
              );
              const inputTd = <td className="n rs" rowSpan={n}><b>{fmt(t.input)}</b></td>;
              const outTd = <td className="n rs" rowSpan={n}><b>{fmt(t.out)}</b></td>;
              const items = p.ngItems.length ? p.ngItems : [{ id: "", name: "—" }];
              return (
                <Fragment key={p.id}>
                  {segHead && (
                    <tr className="seg-row">
                      <td colSpan={9}>
                        <span className="bar" style={{ background: SEGMENTS[p.segment].color }} />
                        {SEGMENTS[p.segment].name}
                      </td>
                    </tr>
                  )}
                  {items.map((item, i) => {
                    const [b, a] = t.pairs[i] ?? [t.before, t.after];
                    return (
                      <tr key={item.id || "none"} className={i === n - 1 ? "plast" : "sub"} data-pid={p.id}>
                        {i === 0 && head}
                        <td className="dn">{item.name}</td>
                        {i === 0 && inputTd}
                        <td className="n">{fmt(b)}</td>
                        <td className="n">{fmt(a)}</td>
                        {rates(b, a)}
                        {i === 0 && outTd}
                      </tr>
                    );
                  })}
                  {withTotal && (
                    <tr className="plast pi-total" data-pid={p.id}>
                      <td className="dn">Total</td>
                      <td className="n">{fmt(t.before)}</td>
                      <td className="n">{fmt(t.after)}</td>
                      {rates(t.before, t.after)}
                    </tr>
                  )}
                </Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function DayTable(props: {
  setup: ModelSetup;
  data: DayData;
  rows: Rows;
  slot: Slot;
  onSlot: (slot: Slot) => void;
  onData: (data: DayData) => void;
}) {
  const { setup, data, rows, slot, onData } = props;
  const firstE = firstOf(setup, "E");
  let lastSeg = "";

  return (
    <section className="card pi-sec">
      <div className="pi-sec-h">
        <h2>Daily NG — enter only the NG counts <span className="cnt">Before rework / After rework</span></h2>
        <div className="pi-legend">
          <span><button className="pi-tag-a" disabled>AUTO</button> input carried from the previous process · type over it when units are taken out</span>
          <span><button className="pi-tag-o" disabled>SET</button> typed by hand · click to go back to AUTO</span>
          <span>Function slot
            <span className="pi-seg">
              {(["day", "am", "pm"] as Slot[]).map((s) => (
                <button key={s} className={slot === s ? "on" : ""} onClick={() => props.onSlot(s)}>
                  {s === "day" ? "Day total" : s === "am" ? "08–12h" : "12–17h"}
                </button>
              ))}
            </span>
          </span>
        </div>
      </div>
      <div className="pi-wrap">
        <table className="pi-pe">
          <thead>
            <tr><th>#</th><th>Process</th><th>NG item</th><th className="n">Input</th><th className="n">Before</th><th className="n">After</th>
              <th className="n">NG % Before</th><th className="n">NG % After</th><th className="n">Output</th></tr>
          </thead>
          <tbody>
            {setup.processes.map((p, idx) => {
              const segHead = p.segment !== lastSeg;
              lastSeg = p.segment;
              const r = rows[p.id];
              const v = slot !== "day" && p.segment === "E" ? r[slot] : null;
              const input = v ? v.input : r.input;
              const before = v ? v.before : r.before, after = v ? v.after : r.after, out = v ? v.out : r.out;
              const pairs = viewPairs(p, data, slot);
              const readOnly = p.segment === "E" && slot === "day";
              // Every input is typed in place: units taken out between processes (e.g. for review) make it
              // smaller than the previous output. Typing the automatic value back returns to AUTO.
              const overridden = data.override[p.id] != null;
              const slotView = p.segment === "E" && slot !== "day";
              const firstSet = firstOf(setup, p.segment) === p && SEGMENTS[p.segment].start === "set";
              let inCell: ReactNode;
              if (slotView && !(p === firstE && slot === "am")) {
                inCell = <span className="pi-auto"><b>{fmt(input)}</b><button className="pi-tag-a" disabled
                  title={p === firstE ? "Day input − 08–12h input" : "Carried within the time slot · type the day input in Day total"}>AUTO</button></span>;
              } else {
                const tag = overridden
                  ? <button className="pi-tag-o" title="Typed by hand. Click to go back to AUTO" onClick={() => onData(clearOverride(data, p.id))}>SET</button>
                  : slotView || firstSet
                    ? <button className="pi-tag-o" disabled title={slotView ? "Input tested in 08–12h" : "Segment start count"}>SET</button>
                    : r.mode === "prev"
                      ? <button className="pi-tag-a" disabled title="Previous day's output of the previous segment · type to change">PREV</button>
                      : <button className="pi-tag-a" disabled title="Previous process output · type to change">AUTO</button>;
                const taken = overridden && !slotView ? r.autoInput - r.input : 0;
                inCell = (
                  <span className="pi-auto">
                    <NumInput className={"wide" + (overridden ? " typed" : "")} value={input} label={`Input of ${p.name}`}
                      onCommit={(val) => onData(setInput(setup, data, p, slot, val ?? 0, slotView ? undefined : r.autoInput))} />
                    {tag}
                    {taken !== 0 && (
                      <small className="pi-taken" title={`Automatic input ${fmt(r.autoInput)}`}>
                        {taken > 0 ? `−${fmt(taken)} out` : `+${fmt(-taken)} in`}
                      </small>
                    )}
                  </span>
                );
              }
              // Each NG item shows its own rate over the process input; a process with several items
              // adds a total row with the summed counts and the process NG rate.
              const withTotal = p.ngItems.length > 1;
              const n = Math.max(1, p.ngItems.length) + (withTotal ? 1 : 0);
              const head = (
                <>
                  <td className="pi-pno rs" rowSpan={n}>{idx + 1}</td>
                  <td className="pi-pname rs" rowSpan={n}>
                    {p.name}
                    {p.localName && <small>{p.localName}</small>}
                    {readOnly && <small>Day = 08–12h + 12–17h · switch slot to edit</small>}
                  </td>
                </>
              );
              const inputTd = <td className="n rs" rowSpan={n}>{inCell}</td>;
              const outTd = <td className="n rs" rowSpan={n}><b>{fmt(out)}</b></td>;
              const rates = (b: number, a: number) => (
                <>
                  <td className="n">{rateCell(ngRate(b, input))}</td>
                  <td className="n">{rateCell(ngRate(a, input))}</td>
                </>
              );
              return (
                <Fragment key={p.id}>
                  {segHead && (
                    <tr className="seg-row">
                      <td colSpan={9}>
                        <span className="bar" style={{ background: SEGMENTS[p.segment].color }} />
                        {SEGMENTS[p.segment].name}<span className="why">{SEGMENTS[p.segment].why}</span>
                      </td>
                    </tr>
                  )}
                  {p.ngItems.length === 0 ? (
                    <tr className="plast" data-pid={p.id}>
                      {head}<td className="muted small">No NG item — add one in Setup</td>{inputTd}<td colSpan={2}></td>{rates(before, after)}{outTd}
                    </tr>
                  ) : p.ngItems.map((item, i) => {
                    const [b, a] = pairs[i];
                    return (
                      <tr key={item.id} className={i === n - 1 ? "plast" : "sub"} data-pid={p.id}>
                        {i === 0 && head}
                        <td className="dn">{item.name}</td>
                        {i === 0 && inputTd}
                        <td className="n">
                          <NumInput value={b} disabled={readOnly} label={`${p.name} ${item.name} before`}
                            onCommit={(val) => onData(setPair(data, p, slot, item.id, 0, val ?? 0))} />
                        </td>
                        <td className="n">
                          <NumInput value={a} disabled={readOnly} className={a > b ? "err" : ""} label={`${p.name} ${item.name} after`}
                            title={a > b ? "After cannot be larger than Before" : undefined}
                            onCommit={(val) => onData(setPair(data, p, slot, item.id, 1, val ?? 0))} />
                        </td>
                        {rates(b, a)}
                        {i === 0 && outTd}
                      </tr>
                    );
                  })}
                  {withTotal && (
                    <tr className="plast pi-total" data-pid={p.id}>
                      <td className="dn">Total</td>
                      <td className="n">{fmt(before)}</td>
                      <td className="n">{fmt(after)}</td>
                      {rates(before, after)}
                    </tr>
                  )}
                </Fragment>
              );
            })}
            <tr className="plast">
              <td className="pi-pno">—</td>
              <td className="pi-pname">Send to OQC</td>
              <td className="muted">—</td>
              <td className="n">
                <span className="pi-auto">
                  <NumInput className="wide" value={data.sendOqc} label="Send to OQC"
                    onCommit={(val) => onData({ ...data, sendOqc: val ?? 0 })} />
                  <button className="pi-tag-o" disabled>SET</button>
                </span>
              </td>
              <td className="muted" colSpan={5}>Moved quantity (was "OQC" column of NG Rate + "Send OQC" of By hour)</td>
            </tr>
          </tbody>
        </table>
      </div>
    </section>
  );
}
