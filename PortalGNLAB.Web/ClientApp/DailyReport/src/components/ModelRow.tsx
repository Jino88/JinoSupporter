import type { DailyReportModelRow } from "../contract";
import { directionClass, formatPercent, formatPpm, formatSignedPpm } from "../format";
import { baselineGapLabel } from "../logic";
import { CauseList } from "./CauseList";
import { TrendSpark } from "./TrendSpark";

interface ModelRowProps {
  row: DailyReportModelRow;
  rank: number;
  recentHeader: string;
  noun: "week" | "day";
  stepUnit: "w" | "d";
}

export function ModelRow({ row, rank, recentHeader, noun, stepUnit }: ModelRowProps) {
  const gap = baselineGapLabel(row.baselineStepsBack, stepUnit);
  return (
    <tr className="dr-model-row">
      <td className="dr-rank"><span className={`dr-rank-badge${rank === 1 ? " dr-rank-first" : ""}`}>{rank}</span></td>
      <td>
        <span className="dr-model-name">{row.modelName}</span>
        <span className="dr-model-path"><span>{row.groupName}</span> · {row.parentModelName}</span>
      </td>
      <td className="dr-number"><strong className="dr-latest-value">{formatPpm(row.recentPpm)}</strong></td>
      <td className="dr-number">
        <span className="dr-baseline-value">{row.baselineStepsBack > 0 ? formatPpm(row.previousPpm) : "—"}</span>
        <span className="dr-baseline-period">{row.baselineHeader || "no earlier data"}{gap ? <span className="dr-gap-badge">{gap}</span> : null}</span>
      </td>
      <td className="dr-number">
        <strong className={`dr-change ${directionClass(row.deltaPpm)}`}>{formatSignedPpm(row.deltaPpm)}</strong>
        <span className="dr-change-percent">{formatPercent(row.deltaPercent)}</span>
      </td>
      <td>
        <TrendSpark
          modelName={row.modelName}
          noun={noun}
          values={row.trendPpm}
          baselinePpm={row.previousPpm}
          latestPpm={row.recentPpm}
          deltaPpm={row.deltaPpm}
        />
      </td>
      <td>
        <CauseList causes={row.topCauses} recentHeader={recentHeader} baselineHeader={row.baselineHeader} />
        <div className="dr-cause-total"><strong>{row.recentNgQty.toLocaleString("en-US")}</strong> defects in the latest {noun}</div>
      </td>
    </tr>
  );
}
