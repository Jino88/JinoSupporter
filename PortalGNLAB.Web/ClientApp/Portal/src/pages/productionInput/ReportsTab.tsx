import type { Sheet } from "./types";

/** Renders a sheet with its merges, the same grid POST /export writes to .xlsx. */
export function SheetTable({ sheet }: { sheet: Sheet }) {
  const covered = new Set<string>();
  const spans = new Map<string, { rowSpan: number; colSpan: number }>();
  for (const m of sheet.merges) {
    spans.set(`${m.row},${m.col}`, { rowSpan: m.rowSpan, colSpan: m.colSpan });
    for (let r = m.row; r < m.row + m.rowSpan; r++)
      for (let c = m.col; c < m.col + m.colSpan; c++)
        if (r !== m.row || c !== m.col) covered.add(`${r},${c}`);
  }
  return (
    <table className="pi-xl">
      <tbody>
        {sheet.rows.map((row, r) => (
          <tr key={r}>
            {row.map((cell, c) => {
              if (covered.has(`${r},${c}`)) return null;
              const span = spans.get(`${r},${c}`);
              const Tag = r < sheet.headerRows ? "th" : "td";
              return (
                <Tag key={c} rowSpan={span?.rowSpan} colSpan={span?.colSpan}>
                  {typeof cell === "number" ? cell.toLocaleString("en-US") : cell ?? ""}
                </Tag>
              );
            })}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

export default function ReportsTab(props: {
  planSheet: Sheet;
  ngSheet: Sheet;
  ngDays: number;
  busy: boolean;
  onExport: (sheets: Sheet[]) => void;
}) {
  const { planSheet, ngSheet } = props;
  return (
    <>
      <section className="card pi-sec">
        <div className="pi-sec-h">
          <h2>Output daily — Plan vs Actual <span className="cnt">plan window</span></h2>
          <button className="btn sm" disabled={props.busy} onClick={() => props.onExport([planSheet])}>Download .xlsx</button>
        </div>
        <div className="pi-sec-b pi-wrap"><SheetTable sheet={planSheet} /></div>
      </section>
      <section className="card pi-sec">
        <div className="pi-sec-h">
          <h2>Daily NG Rate — original wide layout <span className="cnt">{props.ngDays} saved day(s) in the plan window</span></h2>
          <div className="pi-legend">
            <button className="btn sm" disabled={props.busy || props.ngDays === 0} onClick={() => props.onExport([ngSheet])}>Download .xlsx</button>
            <button className="btn sm primary" disabled={props.busy} onClick={() => props.onExport([planSheet, ngSheet])}>Download both</button>
          </div>
        </div>
        <div className="pi-sec-b pi-wrap">
          {props.ngDays === 0 ? <div className="muted">No saved day in this window yet. Save the daily input first.</div> : <SheetTable sheet={ngSheet} />}
        </div>
      </section>
    </>
  );
}
