import type { PropsWithChildren, ReactNode } from "react";

interface ReportTableProps extends PropsWithChildren {
  caption: string;
  headers: ReactNode[];
  className?: string;
}

export function ReportTable({ caption, headers, children, className = "" }: ReportTableProps) {
  return (
    <div className="table-scroll" tabIndex={0} role="region" aria-label={`${caption} 스크롤 영역`}>
      <table className={`report-table ${className}`.trim()}>
        <caption>{caption}</caption>
        <thead>
          <tr>{headers.map((header, index) => <th key={index} scope="col">{header}</th>)}</tr>
        </thead>
        <tbody>{children}</tbody>
      </table>
    </div>
  );
}

export function EmptySection({ message = "표시할 데이터가 없습니다." }: { message?: string }) {
  return <p className="empty-state" role="status">{message}</p>;
}

export function Section({ title, count, children }: PropsWithChildren<{ title: string; count?: number }>) {
  return (
    <section className="report-section">
      <div className="section-heading">
        <h2>{title}</h2>
        {count !== undefined && <span className="count-badge">{count.toLocaleString("en-US")}</span>}
      </div>
      {children}
    </section>
  );
}

export function indentStyle(depth: number) {
  return { paddingInlineStart: `${Math.max(0, depth) * 1.1 + 0.7}rem` };
}
