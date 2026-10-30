import { Fragment, type ReactNode } from "react";
import { NumInput } from "./fields";
import { SEGMENTS, clearOverride, firstOf, fmt, pct, setInput, setPair, viewPairs, type Rows } from "./logic";
import type { DayData, ModelSetup, Slot } from "./types";

const rateCell = (v: number) => (
  <span className={"pi-rate " + (v === 0 ? "z" : v >= 0.02 ? "hi" : "")}>{v === 0 ? "0" : pct(v)}</span>
);

export default function DailyTab(props: {
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
              const n = Math.max(1, p.ngItems.length);
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
              const tail = (
                <>
                  <td className="n rs" rowSpan={n}>{rateCell(input ? before / input : 0)}</td>
                  <td className="n rs" rowSpan={n}>{rateCell(input ? after / input : 0)}</td>
                  <td className="n rs" rowSpan={n}><b>{fmt(out)}</b></td>
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
                      {head}<td className="muted small">No NG item — add one in Setup</td>{inputTd}<td colSpan={2}></td>{tail}
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
                        {i === 0 && tail}
                      </tr>
                    );
                  })}
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
