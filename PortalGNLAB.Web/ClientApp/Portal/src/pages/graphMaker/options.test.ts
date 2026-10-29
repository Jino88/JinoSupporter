import { describe, it, expect } from "vitest";
import {
  inputMode,
  graphFields,
  graphDefaults,
  loadedOptions,
  defaultGroup,
  exampleOptions,
  buildRequests,
  visibleField,
} from "./options";
describe("graph request options", () => {
  it("maps specialized graph modes", () => {
    expect([0, 1, 4, 5, 6, 7].map(inputMode)).toEqual([0, 0, 0, 2, 3, 4]);
  });
  it("provides a distinct default for every control", () => {
    expect(new Set(graphFields.map((f) => f.key)).size).toBe(
      graphFields.length,
    );
    for (const f of graphFields)
      expect(Object.hasOwn(graphDefaults, f.key)).toBe(true);
  });
});

describe("Blazor option defaults", () => {
  const table = {
    headers: ["Date", "A#1", "A#2", "NG", "Input", "REF", "UP", "LOW", "MASK"],
    rows: [["2026", "1", "2", "3", "100", "2", "5", "0", "0"]],
  };
  it("selects at most three Y columns and detects limits and heatmap columns", () => {
    const options = loadedOptions(table, { ...graphDefaults });
    expect(options.yColumns).toEqual(["A#1", "A#2", "NG"]);
    expect(options.graphColumnGroups).toEqual({
      "A#1": "A",
      "A#2": "A",
      NG: "NG",
    });
    expect([
      options.graphSpecColumn,
      options.graphUpperColumn,
      options.graphLowerColumn,
    ]).toEqual(["REF", "UP", "LOW"]);
    expect([
      options.heatMapNgColumn,
      options.heatMapInputColumn,
      options.heatMapRateUnit,
    ]).toEqual(["NG", "Input", 1]);
    expect(defaultGroup(" Test#12 ")).toBe("Test");
    expect(defaultGroup("Test#abc")).toBe("Test#abc");
    expect(defaultGroup("#1")).toBe("#1");
    expect(defaultGroup("Test#")).toBe("Test");
  });
  it("excludes line limit columns from the example series", () => {
    expect(
      exampleOptions(0, "GRAPH_spl.txt", table, { ...graphDefaults }).yColumns,
    ).toEqual(["A#1", "A#2", "NG", "Input"]);
  });
  it("applies both scatter example variants", () => {
    const noHeader = exampleOptions(1, "SCATTER_NOHEADER.txt", table, {
      ...graphDefaults,
    });
    expect(noHeader.xColumns).toEqual(["Date"]);
    expect(noHeader.scatterPointLayout).toBe(1);
    expect(noHeader.showScatterStd).toBe(true);
    const hook = exampleOptions(1, "SCATTER_hookdimensiondata.txt", table, {
      ...graphDefaults,
    });
    expect(hook.xColumns).toHaveLength(4);
    expect(hook.yColumns).toHaveLength(4);
    expect(hook.scatterPointLayout).toBe(0);
  });
  it("splits process trend charts and excludes X from series", () => {
    const options = {
      ...graphDefaults,
      xColumn: "X",
      yColumns: ["X", "Y1", "Y2"],
    };
    expect(buildRequests(6, options).map((r) => r.yColumns)).toEqual([
      ["Y1"],
      ["Y2"],
    ]);
    expect(buildRequests(1, options)[0].xColumn).toBe("");
    expect(buildRequests(4, options)[0].xColumn).toBe("");
  });
  it("sets process and heatmap example settings", () => {
    const process = exampleOptions(
      6,
      "process.txt",
      {
        headers: ["Time", "Type", "L_Side", "R_Side"],
        rows: [["1", "x", "1", "2"]],
      },
      { ...graphDefaults },
    );
    expect(process.yColumns).toEqual(["L_Side", "R_Side"]);
    expect([
      process.processUpperRowNumber,
      process.processLowerRowNumber,
      process.showProcessContour,
    ]).toEqual([2, 3, true]);
    const heat = exampleOptions(
      7,
      "heat.json",
      {
        headers: ["IR_Code", "VP_Mold", "Hearing_NG", "Hearing_Input"],
        rows: [["x", "y", "1", "100"]],
      },
      { ...graphDefaults },
    );
    expect([
      heat.heatMapNgColumn,
      heat.heatMapInputColumn,
      heat.heatMapUseNgRate,
      heat.heatMapRateUnit,
    ]).toEqual(["Hearing_NG", "Hearing_Input", true, 0]);
    expect(visibleField("heatMapValueColumn", 7, heat)).toBe(false);
    expect(visibleField("heatMapNgColumn", 7, heat)).toBe(true);
    expect(visibleField("showScatterCpk", 6, heat)).toBe(false);
  });
});
