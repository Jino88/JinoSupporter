import type { DailyReportCause } from "../contract";
import {
  directionClass,
  formatCauseDelta,
  formatPpm,
  isMainProcessType,
  processTypeLabel,
} from "../format";

interface CauseListProps {
  causes: DailyReportCause[];
  recentHeader: string;
  baselineHeader: string;
}

export function CauseList({ causes, recentHeader, baselineHeader }: CauseListProps) {
  const comparisonHeader = baselineHeader || "—";
  return (
    <div className="dr-cause-block">
      <div
        className="dr-cause-column-head"
        aria-label={`Cause PPM: current ${recentHeader}; baseline ${comparisonHeader}; change equals current minus baseline`}
      >
        <span className="dr-cause-column-spacer" aria-hidden="true" />
        <span>CURRENT<strong>{recentHeader}</strong></span>
        <span>BASELINE<strong>{comparisonHeader}</strong></span>
        <span>CHANGE<strong>{recentHeader}−{comparisonHeader}</strong></span>
      </div>
      <ol className="dr-cause-list">
        {causes.slice(0, 5).map((cause, index) => (
          <li className="dr-cause-row" key={`${cause.ngName}-${cause.processName}-${index}`}>
            <span className={`dr-process-tag ${isMainProcessType(cause.processType) ? "dr-process-main" : ""}`}>
              {processTypeLabel(cause.processType)}
            </span>
            <span className="dr-cause-name">
              <span className="dr-cause-ng">{cause.ngName}</span>
              <span className="dr-cause-process"> · {cause.processName}</span>
              <span className="dr-cause-type"> · {cause.processType}</span>
            </span>
            <span className="dr-cause-value">{formatPpm(cause.ppm)}</span>
            <span className="dr-cause-previous">{cause.isNew ? "—" : formatPpm(cause.previousPpm)}</span>
            <span className={`dr-cause-delta ${directionClass(cause.deltaPpm, cause.isNew)}`}>
              {cause.isNew ? <span className="dr-new-badge">NEW</span> : formatCauseDelta(cause)}
            </span>
          </li>
        ))}
      </ol>
    </div>
  );
}
