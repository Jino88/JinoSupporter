import { Fragment, type ReactNode } from "react";
import type { GroupRow, NgRow, ProcessRow, ReasonRow, Report, SummaryRow } from "../../ngRateApi";
import { FlatPeriodCells, FlatPeriodHeads } from "./parts";
import { processTypeClass, type ColumnSet } from "./logic";

// The four report tables of By Model (NgRatePage) and By Section (NgRateAllPage): Summary, Process 10 NG,
// Worst 10 NG and Reason. Each top row expands into one "↳ group" row per group of the report.

const PtBadge = ({ type }: { type: string }) => <span className={"pt-badge " + processTypeClass(type)}>{type}</span>;

const Toggle = ({ open }: { open: boolean }) => <span className="toggle-icon">{open ? "▼" : "▶"}</span>;

const Blank = ({ n }: { n: number }) => <>{Array.from({ length: n }, (_, i) => <td key={i} />)}</>;

function GroupRows({ groups, cols, lead, trail = 0, shown }: {
  groups: GroupRow[];
  cols: ColumnSet;
  /** Empty cells before the "↳ group" cell. */
  lead: number;
  trail?: number;
  shown: (group: string) => boolean;
}) {
  return (
    <>
      {groups.filter((g) => shown(g.groupName)).map((g) => (
        <tr key={g.groupName} className="group-row">
          <Blank n={lead} />
          <td className="label-td group-name-td">↳ {g.groupName}</td>
          <Blank n={trail} />
          <FlatPeriodCells cols={cols} ppm={g.ppm} plain />
        </tr>
      ))}
    </>
  );
}

export interface SummaryTableProps {
  report: Report;
  cols: ColumnSet;
  expanded: ReadonlySet<string>;
  onToggle: (processType: string) => void;
  /** Whether the "↳ group" row of the group is shown. */
  shown: (group: string) => boolean;
  /** First cell of a "↳ group" row (By Model: hide button). */
  groupLead?: (group: string) => ReactNode;
}

export function SummaryTable({ report, cols, expanded, onToggle, shown, groupLead }: SummaryTableProps) {
  const hasGroups = Object.keys(report.groupSummary).length > 0;
  const groupRow = (row: SummaryRow) => report.groupNames.flatMap((grp) => {
    if (!shown(grp)) return [];
    const match = report.groupSummary[grp]?.find((r) => r.processType === row.processType);
    if (!match) return [];
    return [
      <tr key={grp} className="group-row">
        {groupLead ? groupLead(grp) : <td />}
        <td className="label-td group-name-td">↳ {grp}</td>
        <FlatPeriodCells cols={cols} ppm={match.ppm} plain />
      </tr>,
    ];
  });
  return (
    <div className="pivot-wrap">
      <table className="pivot-table text-fit-table">
        <thead>
          <tr>
            <th className="label-th" style={{ minWidth: 20, textAlign: "center" }} />
            <th className="label-th">Process Type</th>
            <FlatPeriodHeads cols={cols} />
          </tr>
        </thead>
        <tbody>
          {report.summary.map((row) => {
            const open = expanded.has(row.processType);
            return (
              <Fragment key={row.processType}>
                <tr className={row.isTotal ? "total-row" : "rank-row"}>
                  <td className="toggle-cell" onClick={() => hasGroups && onToggle(row.processType)}>
                    {hasGroups && <Toggle open={open} />}
                  </td>
                  <td className="label-td">{row.isTotal ? row.processType : <PtBadge type={row.processType} />}</td>
                  <FlatPeriodCells cols={cols} ppm={row.ppm} />
                </tr>
                {open && groupRow(row)}
              </Fragment>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

interface RankTableProps<T> {
  rows: T[];
  cols: ColumnSet;
  expanded: ReadonlySet<number>;
  onToggle: (rank: number) => void;
  shown: (group: string) => boolean;
}

export function ProcessTable({ rows, cols, expanded, onToggle, shown }: RankTableProps<ProcessRow>) {
  return (
    <div className="pivot-wrap">
      <table className="pivot-table text-fit-table">
        <thead>
          <tr>
            <th className="label-th" style={{ minWidth: 28, textAlign: "center" }}>#</th>
            <th className="label-th" style={{ minWidth: 72 }}>Type</th>
            <th className="label-th process-name-col" style={{ minWidth: 150 }}>Process Name</th>
            <FlatPeriodHeads cols={cols} />
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => {
            const hasGroups = row.groups.length > 0;
            const open = hasGroups && expanded.has(row.rank);
            return (
              <Fragment key={row.rank}>
                <tr className="rank-row">
                  <td className="toggle-cell" onClick={() => hasGroups && onToggle(row.rank)}>
                    {hasGroups && <Toggle open={open} />}
                    {row.rank}
                  </td>
                  <td className="label-td"><PtBadge type={row.processType} /></td>
                  <td className="label-td process-name-col">{row.processName}</td>
                  <FlatPeriodCells cols={cols} ppm={row.ppm} />
                </tr>
                {open && <GroupRows groups={row.groups} cols={cols} lead={2} shown={shown} />}
              </Fragment>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

export function NgTable({ rows, cols, expanded, onToggle, shown }: RankTableProps<NgRow>) {
  return (
    <div className="pivot-wrap">
      <table className="pivot-table text-fit-table">
        <thead>
          <tr>
            <th className="label-th" style={{ minWidth: 28, textAlign: "center" }}>#</th>
            <th className="label-th" style={{ minWidth: 72 }}>Type</th>
            <th className="label-th process-name-col" style={{ minWidth: 150 }}>Process Name</th>
            <th className="label-th ng-name-col" style={{ minWidth: 130 }}>NG Name</th>
            <FlatPeriodHeads cols={cols} />
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => {
            const hasGroups = row.groups.length > 0;
            const open = hasGroups && expanded.has(row.rank);
            return (
              <Fragment key={row.rank}>
                <tr className="rank-row">
                  <td className="toggle-cell" onClick={() => hasGroups && onToggle(row.rank)}>
                    {hasGroups && <Toggle open={open} />}
                    {row.rank}
                  </td>
                  <td className="label-td"><PtBadge type={row.processType} /></td>
                  <td className="label-td process-name-col">{row.processName}</td>
                  <td className="label-td ng-name-col">{row.ngName}</td>
                  <FlatPeriodCells cols={cols} ppm={row.ppm} />
                </tr>
                {open && <GroupRows groups={row.groups} cols={cols} lead={2} trail={1} shown={shown} />}
              </Fragment>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

/** Key of a reason combo row in the expanded set. */
export const reasonKey = (row: ReasonRow): string => `${row.reason}|${row.no}`;

export function ReasonTable({ rows, cols, expanded, onToggle, collapsed, onToggleCategory, shown }: {
  /** Rows already filtered for display (logic.displayReasonRows). */
  rows: ReasonRow[];
  cols: ColumnSet;
  expanded: ReadonlySet<string>;
  onToggle: (key: string) => void;
  /** Reason categories whose combo rows are folded under the Total row. */
  collapsed: ReadonlySet<string>;
  onToggleCategory: (reason: string) => void;
  shown: (group: string) => boolean;
}) {
  let lastReason: string | null = null;
  return (
    <div className="pivot-wrap">
      <table className="pivot-table text-fit-table">
        <thead>
          <tr>
            <th className="label-th" style={{ minWidth: 90 }}>Reason</th>
            <th className="label-th" style={{ minWidth: 28, textAlign: "center" }}>No</th>
            <th className="label-th" style={{ minWidth: 72 }}>Type</th>
            <th className="label-th process-name-col" style={{ minWidth: 160 }}>Process Name</th>
            <th className="label-th ng-name-col" style={{ minWidth: 120 }}>NG Name</th>
            <FlatPeriodHeads cols={cols} />
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => {
            const reasonChanged = row.reason !== lastReason;
            lastReason = row.reason;
            const folded = collapsed.has(row.reason);
            if (row.isTotal) {
              return (
                <tr key={i} className="total-row">
                  <td className="toggle-cell label-td fw-bold" onClick={() => onToggleCategory(row.reason)}>
                    <Toggle open={!folded} />
                    {row.reason}
                  </td>
                  <td />
                  <td />
                  <td className="label-td text-muted" style={{ fontSize: 10 }}>Total</td>
                  <td />
                  <FlatPeriodCells cols={cols} ppm={row.ppm} />
                </tr>
              );
            }
            if (folded) return null;
            const key = reasonKey(row);
            const hasGroups = row.groups.length > 0;
            const open = hasGroups && expanded.has(key);
            return (
              <Fragment key={i}>
                <tr className="rank-row">
                  <td className="label-td" style={{ color: "#94a3b8", fontSize: 10 }}>{reasonChanged ? row.reason : ""}</td>
                  <td className="toggle-cell" onClick={() => hasGroups && onToggle(key)}>
                    {hasGroups && <Toggle open={open} />}
                    {row.no}
                  </td>
                  <td className="label-td"><PtBadge type={row.processType} /></td>
                  <td className="label-td process-name-col">{row.processName}</td>
                  <td className="label-td ng-name-col">{row.ngName}</td>
                  <FlatPeriodCells cols={cols} ppm={row.ppm} />
                </tr>
                {open && <GroupRows groups={row.groups} cols={cols} lead={3} trail={1} shown={shown} />}
              </Fragment>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

/** Expand All / Collapse All links of a card header. */
export function ExpandLinks({ onExpand, onCollapse }: { onExpand: () => void; onCollapse: () => void }) {
  return (
    <div className="d-flex gap-2 align-items-center">
      <button type="button" className="expand-btn" onClick={onExpand}>Expand All</button>
      <span className="text-muted" style={{ fontSize: 10 }}>|</span>
      <button type="button" className="expand-btn" onClick={onCollapse}>Collapse All</button>
    </div>
  );
}
