import type { ClipboardEvent } from "react";
import { columnName, type Table } from "./data";
import type { GraphOptions } from "./options";

interface Props {
  table: Table;
  kind: number;
  options: GraphOptions;
  onPaste: (text: string) => void;
}
export default function DataSheet({ table, kind, options, onPaste }: Props) {
  const hasData = table.headers.length > 0 && table.rows.length > 0;
  const columnCount = hasData ? table.headers.length : 12;
  const rowCount = hasData ? Math.max(table.rows.length, 30) : 30;
  function paste(event: ClipboardEvent<HTMLDivElement>) {
    const text = event.clipboardData.getData("text/plain");
    if (!text.trim()) return;
    event.preventDefault();
    onPaste(text);
  }
  function columnClass(column: string) {
    const classes: string[] = [];
    if (
      kind !== 7 &&
      kind !== 4 &&
      (column === options.xColumn ||
        (kind === 1 && options.xColumns.includes(column)))
    )
      classes.push("graph-col-x");
    if (options.yColumns.includes(column)) classes.push("graph-col-y");
    if (kind === 7) {
      if (column === options.heatMapXColumn) classes.push("graph-col-x");
      if (column === options.heatMapYColumn) classes.push("graph-col-y");
      if (
        column === options.heatMapValueColumn ||
        (options.heatMapUseNgRate &&
          (column === options.heatMapNgColumn ||
            column === options.heatMapInputColumn))
      )
        classes.push("graph-col-value");
    }
    return classes.join(" ");
  }
  function rowLabel(row: number) {
    return [
      options.processUpperRowNumber === row ? "USL" : "",
      options.processSpecRowNumber === row ? "SPEC" : "",
      options.processLowerRowNumber === row ? "LSL" : "",
    ]
      .filter(Boolean)
      .join(" ");
  }
  return (
    <div
      id="graphMakerSheet"
      className="graph-sheet"
      tabIndex={0}
      onPaste={paste}
      role="region"
      aria-label="Excel data table. Focus here and press Ctrl+V."
    >
      <table className="grid">
        <thead>
          <tr>
            <th />
            {Array.from({ length: columnCount }, (_, i) => (
              <th className={columnClass(table.headers[i] ?? "")} key={i}>
                <div className="muted small">{columnName(i)}</div>
                {hasData && table.headers[i]}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {Array.from({ length: rowCount }, (_, row) => (
            <tr
              key={row}
              className={rowLabel(row + 1) ? "graph-limit-row" : ""}
            >
              <th>
                {row + 1} {rowLabel(row + 1)}
              </th>
              {Array.from({ length: columnCount }, (_, column) => (
                <td
                  key={column}
                  className={columnClass(table.headers[column] ?? "")}
                >
                  {table.rows[row]?.[column] ?? ""}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
