import type { DailyReportModelRow } from "../contract";
import { ModelRow } from "./ModelRow";

interface BandRowProps {
  kind: "worse" | "better";
  rows: DailyReportModelRow[];
  recentHeader: string;
  noun: "week" | "day";
  stepUnit: "w" | "d";
}

export function BandRow({ kind, rows, recentHeader, noun, stepUnit }: BandRowProps) {
  const label = kind === "worse" ? "WORSENED" : "IMPROVED";
  return (
    <>
      <tr className={`dr-band dr-band-${kind}`}>
        <td colSpan={7}>
          {label}
          <span className="dr-band-count">{rows.length === 0 ? "none" : `${rows.length} model${rows.length === 1 ? "" : "s"}`}</span>
        </td>
      </tr>
      {rows.map((row, index) => (
        <ModelRow key={`${kind}-${row.groupName}-${row.modelName}`} row={row} rank={index + 1} recentHeader={recentHeader} noun={noun} stepUnit={stepUnit} />
      ))}
    </>
  );
}
