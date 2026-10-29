import type { Table } from "./data";
export interface GraphOptions {
  xColumn: string;
  xColumns: string[];
  yColumns: string[];
  heatMapXColumn: string;
  heatMapYColumn: string;
  heatMapValueColumn: string;
  heatMapUseNgRate: boolean;
  heatMapNgColumn: string;
  heatMapInputColumn: string;
  heatMapRateUnit: number;
  graphSpecColumn: string;
  graphUpperColumn: string;
  graphLowerColumn: string;
  graphColumnGroups: Record<string, string>;
  upperLimit: number | null;
  specLimit: number | null;
  lowerLimit: number | null;
  processSpecRowNumber: number | null;
  processUpperRowNumber: number | null;
  processLowerRowNumber: number | null;
  showProcessContour: boolean;
  showProcessTrendLine: boolean;
  showProcessMax: boolean;
  showProcessMin: boolean;
  showProcessAvg: boolean;
  showProcessStd: boolean;
  showProcessCpu: boolean;
  showProcessCpl: boolean;
  showProcessCpk: boolean;
  scatterPointLayout: number;
  scatterLabelPosition: number;
  lineMode: number;
  scatterHeaderNotUse: boolean;
  showScatterMax: boolean;
  showScatterMin: boolean;
  showScatterAvg: boolean;
  showScatterStd: boolean;
  showScatterCpu: boolean;
  showScatterCpl: boolean;
  showScatterCpk: boolean;
}
export const graphFields = [
  { key: "xColumn", label: "XColumn", kind: "column" },
  { key: "xColumns", label: "XColumns", kind: "multi" },
  { key: "yColumns", label: "YColumns", kind: "multi" },
  { key: "heatMapXColumn", label: "Heat Map XColumn", kind: "column" },
  { key: "heatMapYColumn", label: "Heat Map YColumn", kind: "column" },
  { key: "heatMapValueColumn", label: "Heat Map Value Column", kind: "column" },
  { key: "heatMapUseNgRate", label: "Heat Map Use Ng Rate", kind: "check" },
  { key: "heatMapNgColumn", label: "Heat Map Ng Column", kind: "column" },
  { key: "heatMapInputColumn", label: "Heat Map Input Column", kind: "column" },
  {
    key: "heatMapRateUnit",
    label: "Heat Map Rate Unit",
    kind: "GraphMakerRateUnit",
  },
  { key: "graphSpecColumn", label: "Graph Spec Column", kind: "column" },
  { key: "graphUpperColumn", label: "Graph Upper Column", kind: "column" },
  { key: "graphLowerColumn", label: "Graph Lower Column", kind: "column" },
  { key: "graphColumnGroups", label: "Graph Column Groups", kind: "groups" },
  { key: "upperLimit", label: "Upper Limit", kind: "number" },
  { key: "specLimit", label: "Spec Limit", kind: "number" },
  { key: "lowerLimit", label: "Lower Limit", kind: "number" },
  {
    key: "processSpecRowNumber",
    label: "Process Spec Row Number",
    kind: "number",
  },
  {
    key: "processUpperRowNumber",
    label: "Process Upper Row Number",
    kind: "number",
  },
  {
    key: "processLowerRowNumber",
    label: "Process Lower Row Number",
    kind: "number",
  },
  { key: "showProcessContour", label: "Show Process Contour", kind: "check" },
  {
    key: "showProcessTrendLine",
    label: "Show Process Trend Line",
    kind: "check",
  },
  { key: "showProcessMax", label: "Show Process Max", kind: "check" },
  { key: "showProcessMin", label: "Show Process Min", kind: "check" },
  { key: "showProcessAvg", label: "Show Process Avg", kind: "check" },
  { key: "showProcessStd", label: "Show Process Std", kind: "check" },
  { key: "showProcessCpu", label: "Show Process Cpu", kind: "check" },
  { key: "showProcessCpl", label: "Show Process Cpl", kind: "check" },
  { key: "showProcessCpk", label: "Show Process Cpk", kind: "check" },
  {
    key: "scatterPointLayout",
    label: "Scatter Point Layout",
    kind: "GraphMakerScatterPointLayout",
  },
  {
    key: "scatterLabelPosition",
    label: "Scatter Label Position",
    kind: "GraphMakerScatterLabelPosition",
  },
  { key: "lineMode", label: "Line Mode", kind: "GraphMakerLineMode" },
  {
    key: "scatterHeaderNotUse",
    label: "Scatter Header Not Use",
    kind: "check",
  },
  { key: "showScatterMax", label: "Show Scatter Max", kind: "check" },
  { key: "showScatterMin", label: "Show Scatter Min", kind: "check" },
  { key: "showScatterAvg", label: "Show Scatter Avg", kind: "check" },
  { key: "showScatterStd", label: "Show Scatter Std", kind: "check" },
  { key: "showScatterCpu", label: "Show Scatter Cpu", kind: "check" },
  { key: "showScatterCpl", label: "Show Scatter Cpl", kind: "check" },
  { key: "showScatterCpk", label: "Show Scatter Cpk", kind: "check" },
];
export const graphDefaults: GraphOptions = {
  xColumn: "",
  xColumns: [],
  yColumns: [],
  heatMapXColumn: "",
  heatMapYColumn: "",
  heatMapValueColumn: "",
  heatMapUseNgRate: false,
  heatMapNgColumn: "",
  heatMapInputColumn: "",
  heatMapRateUnit: 1,
  graphSpecColumn: "",
  graphUpperColumn: "",
  graphLowerColumn: "",
  graphColumnGroups: {},
  upperLimit: null,
  specLimit: null,
  lowerLimit: null,
  processSpecRowNumber: null,
  processUpperRowNumber: null,
  processLowerRowNumber: null,
  showProcessContour: false,
  showProcessTrendLine: false,
  showProcessMax: false,
  showProcessMin: false,
  showProcessAvg: false,
  showProcessStd: false,
  showProcessCpu: false,
  showProcessCpl: false,
  showProcessCpk: false,
  scatterPointLayout: 0,
  scatterLabelPosition: 0,
  lineMode: 0,
  scatterHeaderNotUse: false,
  showScatterMax: false,
  showScatterMin: false,
  showScatterAvg: false,
  showScatterStd: false,
  showScatterCpu: false,
  showScatterCpl: false,
  showScatterCpk: false,
};
export const graphKinds = [
  "Graph",
  "Scatter",
  "Average",
  "CPK",
  "Normal distribution",
  "No X multi Y",
  "Process trend",
  "Heatmap",
];
export function inputMode(kind: number): number {
  return kind === 7 ? 4 : kind === 6 ? 3 : kind === 5 ? 2 : 0;
}
export function defaultGroup(column: string): string {
  const value = column.trim();
  const marker = value.lastIndexOf("#");
  return marker > 0 && /^\d*$/.test(value.slice(marker + 1))
    ? value.slice(0, marker).trimEnd()
    : value;
}
export function withGroups(options: GraphOptions): GraphOptions {
  return {
    ...options,
    graphColumnGroups: Object.fromEntries(
      options.yColumns.map((column) => [
        column,
        options.graphColumnGroups[column]?.trim() || defaultGroup(column),
      ]),
    ),
  };
}
export function loadedOptions(
  table: Table,
  previous: GraphOptions,
): GraphOptions {
  const headers = table.headers;
  const next = { ...previous, xColumns: [], yColumns: [] } as GraphOptions;
  if (!headers.length || !table.rows.length) return next;
  next.xColumn = headers[0];
  next.yColumns = headers.slice(1, 4);
  next.heatMapXColumn = headers[0];
  next.heatMapYColumn = headers[1] ?? "";
  next.heatMapValueColumn =
    headers
      .slice(2)
      .find((h) =>
        table.rows.some(
          (row) =>
            row[headers.indexOf(h)]?.trim() &&
            Number.isFinite(
              Number(row[headers.indexOf(h)].replaceAll(",", "")),
            ),
        ),
      ) ?? "";
  next.heatMapNgColumn =
    headers.find((h) => /ng/i.test(h)) ?? next.heatMapValueColumn;
  next.heatMapInputColumn =
    headers.find((h) => /input/i.test(h)) ?? next.heatMapValueColumn;
  next.graphSpecColumn = headers.find((h) => /^REF$/i.test(h)) ?? "";
  next.graphUpperColumn = headers.find((h) => /^UP$/i.test(h)) ?? "";
  next.graphLowerColumn = headers.find((h) => /^LOW$/i.test(h)) ?? "";
  return withGroups(next);
}
export function exampleOptions(
  kind: number,
  name: string,
  table: Table,
  previous: GraphOptions,
): GraphOptions {
  const headers = table.headers;
  const next = {
    ...loadedOptions(table, previous),
    xColumns: [],
    yColumns: [],
  } as GraphOptions;
  if (!table.rows.length) return next;
  if (/SCATTER_NOHEADER/i.test(name)) {
    next.xColumn = headers[0];
    next.xColumns = headers.slice(0, 1);
    next.yColumns = headers.slice(1);
    Object.assign(next, {
      scatterPointLayout: 1,
      scatterLabelPosition: 1,
      showScatterMax: true,
      showScatterMin: true,
      showScatterAvg: true,
      showScatterStd: true,
    });
  } else if (/SCATTER_hookdimensiondata/i.test(name)) {
    next.xColumns = headers.slice(0, 4);
    next.xColumn = next.xColumns[0] ?? "";
    next.yColumns = headers.slice(4, 8);
    Object.assign(next, {
      scatterPointLayout: 0,
      scatterLabelPosition: 1,
      showScatterMax: true,
      showScatterMin: true,
      showScatterAvg: true,
      showScatterStd: true,
    });
  } else if (kind === 6) {
    next.xColumn = headers[0];
    next.yColumns = headers
      .slice(2)
      .filter((h) => /_Side/i.test(h))
      .slice(0, 3);
    Object.assign(next, {
      processUpperRowNumber: 2,
      processLowerRowNumber: 3,
      processSpecRowNumber: null,
      showProcessMax: true,
      showProcessMin: true,
      showProcessAvg: true,
      showProcessStd: true,
      showProcessContour: true,
      showProcessTrendLine: true,
    });
  } else if (kind === 0) {
    next.xColumn = headers[0];
    next.yColumns = headers
      .slice(1)
      .filter((h) => !/^(REF|UP|LOW|MASK)$/i.test(h));
    next.lineMode = 0;
  } else if (kind === 7) {
    next.xColumn = "";
    next.heatMapXColumn =
      headers.find((h) => /^IR_Code$/i.test(h)) ?? headers[0] ?? "";
    next.heatMapYColumn =
      headers.find((h) => /^VP_Mold$/i.test(h)) ?? headers[1] ?? "";
    next.heatMapNgColumn =
      headers.find((h) => /^Hearing_Noise$/i.test(h)) ??
      headers.find((h) => /^Hearing_NG$/i.test(h)) ??
      "";
    next.heatMapInputColumn =
      headers.find((h) => /^Hearing_Input$/i.test(h)) ?? "";
    Object.assign(next, {
      heatMapValueColumn: "",
      heatMapUseNgRate: true,
      heatMapRateUnit: 0,
    });
  }
  return withGroups(next);
}
export function buildRequests(kind: number, options: GraphOptions) {
  const request = {
    ...withGroups(options),
    graphType: kind,
    inputMode: inputMode(kind),
    xColumn:
      kind === 7 || kind === 4 || (kind === 1 && options.xColumns.length === 0)
        ? ""
        : options.xColumn,
    scatterHeaderNotUse: kind === 1 && options.scatterHeaderNotUse,
  };
  const columns =
    kind === 6 && options.yColumns.length > 1
      ? options.yColumns.filter((c) => c !== options.xColumn)
      : [];
  return columns.length
    ? columns.map((column) => ({ ...request, yColumns: [column] }))
    : [request];
}

export function visibleField(
  key: string,
  kind: number,
  options: GraphOptions,
): boolean {
  if (key.startsWith("heatMap"))
    return (
      kind === 7 &&
      (key === "heatMapXColumn" ||
        key === "heatMapYColumn" ||
        key === "heatMapUseNgRate" ||
        (options.heatMapUseNgRate
          ? key !== "heatMapValueColumn"
          : key === "heatMapValueColumn"))
    );
  if (kind === 7) return false;
  if (key.startsWith("process") || key.startsWith("showProcess"))
    return kind === 6;
  if (
    key.startsWith("scatter") ||
    key.startsWith("showScatter") ||
    key === "xColumns"
  )
    return kind === 1;
  if (key.startsWith("graph") || key === "lineMode") return kind === 0;
  if (key === "xColumn") return kind !== 1 && kind !== 4;
  if (key.endsWith("Limit")) return kind === 1 || kind === 3;
  return key === "yColumns";
}
