import type { DailyReportModelRow } from "../contract";
import type { GranularityData, SortKey } from "../logic";
import { BandRow } from "./BandRow";

interface RecordGridProps {
  data: GranularityData;
  worsened: DailyReportModelRow[];
  improved: DailyReportModelRow[];
  onSortChange: (value: SortKey) => void;
}

export function RecordGrid({ data, worsened, improved, onSortChange }: RecordGridProps) {
  return (
    <div className="dr-grid-scroll">
      <table className="dr-react-grid">
        <colgroup>
          <col className="dr-col-rank" />
          <col className="dr-col-model" />
          <col className="dr-col-latest" />
          <col className="dr-col-baseline" />
          <col className="dr-col-change" />
          <col className="dr-col-trend" />
          <col className="dr-col-causes" />
        </colgroup>
        <thead>
          <tr>
            <th>#</th>
            <th><button className="dr-sort-button" type="button" onClick={() => onSortChange("name")}>MODEL</button></th>
            <th className="dr-head-number"><button className="dr-sort-button" type="button" onClick={() => onSortChange("latest")}>LATEST</button><span>{data.recentHeader} PPM</span></th>
            <th className="dr-head-number">BASELINE<span>last {data.noun} with data</span></th>
            <th className="dr-head-number"><button className="dr-sort-button" type="button" onClick={() => onSortChange("delta")}>CHANGE</button><span>latest − baseline</span></th>
            <th>TREND<span>LATEST ▶ OLDER</span></th>
            <th>TOP 5 DEFECT CAUSES<span>[TYPE] NG · PROCESS · ORIGINAL TYPE · CURRENT · BASELINE · CHANGE (CURRENT−BASELINE)</span></th>
          </tr>
        </thead>
        <tbody>
          <BandRow kind="worse" rows={worsened} recentHeader={data.recentHeader} noun={data.noun} stepUnit={data.stepUnit} />
          <BandRow kind="better" rows={improved} recentHeader={data.recentHeader} noun={data.noun} stepUnit={data.stepUnit} />
        </tbody>
      </table>
    </div>
  );
}
