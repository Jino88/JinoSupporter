import {
  graphFields,
  visibleField,
  withGroups,
  type GraphOptions as Options,
} from "./options";

interface Props {
  kind: number;
  headers: string[];
  options: Options;
  onChange: (options: Options) => void;
  onNoHeader: (enabled: boolean) => void;
}
const labels: Record<string, string> = {
  xColumn: "X axis",
  xColumns: "X columns",
  yColumns: "Y columns",
  heatMapXColumn: "X condition",
  heatMapYColumn: "Y condition",
  heatMapValueColumn: "Cell value",
  heatMapUseNgRate: "NG RATE",
  heatMapNgColumn: "NG column",
  heatMapInputColumn: "INPUT column",
  heatMapRateUnit: "Rate unit",
  graphSpecColumn: "REF column",
  graphUpperColumn: "UP column",
  graphLowerColumn: "LOW column",
  graphColumnGroups: "Y groups",
  upperLimit: "USL",
  specLimit: "Spec",
  lowerLimit: "LSL",
  processUpperRowNumber: "USL row",
  processSpecRowNumber: "SPEC row",
  processLowerRowNumber: "LSL row",
  showProcessContour: "Contour boundary",
  showProcessTrendLine: "Trend line",
  scatterPointLayout: "Point layout",
  scatterLabelPosition: "X label position",
  scatterHeaderNotUse: "Header Not use",
  lineMode: "Line mode",
};
function label(key: string, fallback: string) {
  return (
    labels[key] ??
    key
      .replace(/^show(Process|Scatter)/, "")
      .replace(/Cpu$/, "CPU")
      .replace(/Cpl$/, "CPL")
      .replace(/Cpk$/, "CPK") ??
    fallback
  );
}

export default function GraphOptions({
  kind,
  headers,
  options,
  onChange,
  onNoHeader,
}: Props) {
  function change(key: string, value: unknown) {
    let next = { ...options, [key]: value };
    if (key === "xColumns")
      next = { ...next, xColumn: (value as string[])[0] ?? "" };
    onChange(withGroups(next));
  }
  function clearX() {
    onChange({ ...options, xColumn: "", xColumns: [] });
  }
  return (
    <aside className="card graph-options">
      <h2>
        {kind === 7
          ? "Heatmap Options"
          : kind === 6
            ? "Process Trend Options"
            : kind === 4
              ? "Distribution Options"
              : kind === 3
                ? "CPK Options"
                : "Columns"}
      </h2>
      {kind === 7 && (
        <p className="muted small">Display: Max · Min · Avg · Std</p>
      )}
      {graphFields
        .filter((field) => visibleField(field.key, kind, options))
        .map((field) => {
          const key = field.key as keyof Options;
          const value = options[key];
          const caption =
            key === "yColumns" && kind === 6
              ? "Y process columns"
              : key === "yColumns" && kind === 4
                ? "Distribution columns"
                : key === "yColumns" && kind === 3
                  ? "CPK columns"
                  : key === "xColumn" && kind === 6
                    ? "X process"
                    : label(key, field.label);
          if (field.kind === "groups")
            return (
              <section key={key}>
                <h3>Y groups</h3>
                {options.yColumns.map((column) => (
                  <label className="tool-field" key={column}>
                    {column}
                    <input
                      className="inp sm"
                      value={options.graphColumnGroups[column] ?? column}
                      onChange={(e) =>
                        change(key, {
                          ...options.graphColumnGroups,
                          [column]: e.target.value,
                        })
                      }
                    />
                  </label>
                ))}
              </section>
            );
          if (field.kind === "check")
            return (
              <label key={key} className="graph-check">
                <input
                  type="checkbox"
                  checked={Boolean(value)}
                  onChange={(e) =>
                    key === "scatterHeaderNotUse"
                      ? onNoHeader(e.target.checked)
                      : change(key, e.target.checked)
                  }
                />
                {caption}
              </label>
            );
          return (
            <label className="tool-field" key={key}>
              {caption}
              {field.kind === "number" ? (
                <input
                  className="inp"
                  type="number"
                  step={key.startsWith("process") ? "1" : "any"}
                  min={key.startsWith("process") ? 1 : undefined}
                  value={value == null ? "" : String(value)}
                  onChange={(e) =>
                    change(
                      key,
                      e.target.value === "" ? null : Number(e.target.value),
                    )
                  }
                />
              ) : field.kind === "column" || field.kind === "multi" ? (
                <>
                  <select
                    className="inp graph-select"
                    multiple={field.kind === "multi"}
                    size={
                      field.kind === "multi"
                        ? key === "xColumns"
                          ? 6
                          : kind === 6
                            ? 10
                            : 8
                        : undefined
                    }
                    value={value as string | string[]}
                    onChange={(e) =>
                      change(
                        key,
                        field.kind === "multi"
                          ? Array.from(
                              e.target.selectedOptions,
                              (option) => option.value,
                            )
                          : e.target.value,
                      )
                    }
                  >
                    {field.kind === "column" && (
                      <option value="">
                        {key === "xColumn" ? "Row index" : "None"}
                      </option>
                    )}
                    {headers.map((header) => (
                      <option key={header}>{header}</option>
                    ))}
                  </select>
                  {key === "xColumns" && (
                    <>
                      <span className="muted small">
                        Leave empty to use selected Y columns as the X axis.
                      </span>
                      <button type="button" className="btn sm" onClick={clearX}>
                        Clear X
                      </button>
                    </>
                  )}
                </>
              ) : (
                <select
                  className="inp"
                  value={value as number}
                  onChange={(e) => change(key, Number(e.target.value))}
                >
                  {(field.kind === "GraphMakerRateUnit"
                    ? ["Percent", "PPM"]
                    : field.kind === "GraphMakerLineMode"
                      ? ["Line", "Smoothing"]
                      : field.kind === "GraphMakerScatterPointLayout"
                        ? ["Aligned on one X line", "Spread around X line"]
                        : ["On axis tick", "Between grid lines"]
                  ).map((name, index) => (
                    <option key={index} value={index}>
                      {name}
                    </option>
                  ))}
                </select>
              )}
            </label>
          );
        })}
    </aside>
  );
}
