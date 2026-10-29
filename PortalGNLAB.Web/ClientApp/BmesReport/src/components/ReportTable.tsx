import type { PropsWithChildren, ReactNode } from "react";

export interface HeaderGroup {
  label: ReactNode;
  span: number;
}

interface ReportTableProps extends PropsWithChildren {
  caption: string;
  headers: ReactNode[];
  /** Optional upper header row. `lead` columns stay blank, each group spans `span` of the lower headers. */
  groups?: { lead: number; items: HeaderGroup[] };
  /** Columns from this index on are numeric (right-aligned header, like their cells). */
  numericFrom?: number;
  /** Additional numeric columns before `numericFrom`. */
  numericCols?: number[];
  className?: string;
}

export function ReportTable({ caption, headers, groups, numericFrom = Infinity, numericCols = [], children, className = "" }: ReportTableProps) {
  return (
    <div className="table-scroll" tabIndex={0} role="region" aria-label={`${caption} scroll area`}>
      <table className={`report-table ${className}`.trim()}>
        <caption>{caption}</caption>
        <thead>
          {groups && (
            <tr className="group-row">
              {groups.lead > 0 && <th colSpan={groups.lead} />}
              {groups.items.map((group, index) => <th key={index} colSpan={group.span} scope="colgroup" className="period-group">{group.label}</th>)}
            </tr>
          )}
          <tr>{headers.map((header, index) => <th key={index} scope="col" className={index >= numericFrom || numericCols.includes(index) ? "numeric" : undefined}>{header}</th>)}</tr>
        </thead>
        <tbody>{children}</tbody>
      </table>
    </div>
  );
}

/** Dim placeholder for empty / zero values. */
export function dim(value: ReactNode): ReactNode {
  if (value === null || value === undefined || value === "-" || value === 0) return <span className="z">-</span>;
  if (typeof value === "string" && /^[$]?0(\.0+)?%?( VND)?$/.test(value)) return <span className="z">-</span>;
  return value;
}

export function EmptySection({ message = "No data to show" }: { message?: string }) {
  return <p className="empty-state" role="status">{message}</p>;
}

export function Section({ title, count, children }: PropsWithChildren<{ title: string; count?: number }>) {
  return (
    <section className="report-section">
      <div className="section-heading">
        <h2>{title}</h2>
        {count !== undefined && <span className="count-badge">{count.toLocaleString("en-US")}</span>}
      </div>
      <div className="section-body">{children}</div>
    </section>
  );
}

export function indentStyle(depth: number) {
  return { paddingInlineStart: `${Math.max(0, depth) * 1.1 + 0.7}rem` };
}
