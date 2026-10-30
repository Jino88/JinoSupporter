import { describe as suite, expect, it } from "vitest";
import type { GaussPoint, QrCell, QrData, QrSummaryRow } from "../../qrApi";
import {
  UNMAPPED_CODE,
  buildDataLines,
  buildModelOptions,
  buildNgBreakdown,
  buildNgItemColumns,
  buildRowGroups,
  buildScale,
  buildScatter,
  buildSummaryLines,
  describe,
  exportDateText,
  extractModelCode,
  filterGroups,
  filterSummaryRows,
  filteredSummary,
  formatDay,
  makeJudge,
  matchesModel,
  niceStep,
  quantile,
  selectedDatesLabel,
  stationCards,
  stationDetailUrl,
  virtualWindow,
} from "./logic";

const cell = (station: string, verdict: QrCell["verdict"], time: string | null = "2026-10-01 08:30:00"): QrCell => ({
  station,
  verdict,
  result: verdict === "pass" ? "OK" : "NG",
  time,
  items: [],
});

const row = (productId: string, airleak: QrCell | null, bako: QrCell | null, gauss: QrCell | null, older: QrSummaryRow[] = []): QrSummaryRow => ({
  productId,
  round: 1,
  latestTime: null,
  finalPass: [airleak, bako, gauss].every((c) => c?.verdict === "pass"),
  finalNote: "",
  hasAnyFail: [airleak, bako, gauss].some((c) => c?.verdict === "fail"),
  complete: airleak !== null && bako !== null && gauss !== null,
  airleak,
  bako,
  gauss,
  older,
});

const point = (value: number, date = "2026-10-01", ng = false): GaussPoint => ({
  productId: "J69NAG1GD03700009",
  time: `${date} 08:00:00`,
  date,
  value,
  ng,
  jitter: 0,
});

suite("model codes", () => {
  it("reads characters 11-12 of the QR, upper-cased", () => {
    expect(extractModelCode("J69NAG1GD0370000")).toBe("37");
    expect(extractModelCode("J69NAG1GD0ab")).toBe("AB");
    expect(extractModelCode("short")).toBe("");
    expect(extractModelCode(null)).toBe("");
  });

  it("offers only models present in the fetched QR, with counts, plus Other", () => {
    const codes = [{ code: "37", name: "MSM-X316B" }, { code: "36", name: "MSM-G778B" }, { code: "99", name: "" }];
    const ids = ["J69NAG1GD0370000A", "J69NAG1GD0370000B", "J69NAG1GD0120000C"];
    const { options, known } = buildModelOptions(codes, ids);
    expect(options.map((o) => o.label)).toEqual(["MSM-X316B (37) · 2 QR", "Other · 1 QR"]);
    expect(options[1].code).toBe(UNMAPPED_CODE);
    expect(known.has("99")).toBe(true);
  });

  it("matches Other against QR the table does not name", () => {
    const known = new Set(["37"]);
    expect(matchesModel("J69NAG1GD0370000A", "37", known)).toBe(true);
    expect(matchesModel("J69NAG1GD0120000A", UNMAPPED_CODE, known)).toBe(true);
    expect(matchesModel("J69NAG1GD0370000A", UNMAPPED_CODE, known)).toBe(false);
    expect(matchesModel("anything", "", known)).toBe(true);
  });
});

suite("QR SUMMARY", () => {
  const older = row("Q1", cell("AIRLEAK", "fail"), null, null);
  const rows = [
    row("Q1", cell("AIRLEAK", "pass"), cell("BAKO", "fail"), cell("GAUSS", "pass"), [older]),
    row("Q2", cell("AIRLEAK", "pass"), cell("BAKO", "pass"), cell("GAUSS", "pass")),
    row("Q3", cell("AIRLEAK", "pass"), null, null),
  ];

  it("filters by NG and by complete set", () => {
    const known = new Set<string>();
    expect(filterSummaryRows(rows, "", known, "ng").map((r) => r.productId)).toEqual(["Q1"]);
    expect(filterSummaryRows(rows, "", known, "complete").map((r) => r.productId)).toEqual(["Q1", "Q2"]);
    expect(filterSummaryRows(rows, "", known, "all")).toHaveLength(3);
  });

  it("unfolds earlier rounds in place", () => {
    expect(buildSummaryLines(rows, new Set())).toHaveLength(3);
    const lines = buildSummaryLines(rows, new Set(["Q1"]));
    expect(lines.map((l) => l.isChild)).toEqual([false, true, false, false]);
    expect(lines[0].expanded).toBe(true);
  });

  it("counts tested / NG per station", () => {
    const [air, bako, gauss] = stationCards(rows);
    expect([air.tested, air.ng]).toEqual([3, 0]);
    expect([bako.tested, bako.ng]).toEqual([2, 1]);
    expect(bako.ngRatePpm).toBe(500000);
    expect(gauss.tested).toBe(2);
  });

  it("builds the station deep link with QR and test date", () => {
    expect(stationDetailUrl("J69 X", cell("BAKO", "pass"))).toBe("/app/qr/bako-data?qr=J69%20X&date=2026-10-01&chrome=off");
    expect(stationDetailUrl("Q", cell("GAUSS", "pass", null))).toBe("/app/qr/gauss-data?qr=Q&chrome=off");
    expect(stationDetailUrl("", cell("GAUSS", "pass"))).toBe("");
    expect(stationDetailUrl("Q", cell("OTHER", "pass"))).toBe("");
  });
});

suite("data browser", () => {
  const data: QrData = {
    table: { key: "BKTD", station: "BAKO", menuLabel: "QR BAKO DATA", displayTable: "dbo.BKTD", passValues: ["PASS", "OK"], failValues: ["FAIL", "NG"] },
    fetchedAt: "",
    maxRows: 100,
    totalRows: 4,
    validTotalRows: 4,
    excludedRows: 0,
    sortColumnName: "TestTime",
    selectedDates: ["2026-10-01"],
    warning: "",
    hitMaxRows: false,
    summary: { inputCount: 3, ngCount: 1, ngRatePpm: 333333 },
    dates: [{ date: "2026-10-01", rowCount: 4 }],
    columns: [
      { name: "Product_ID", dataType: "nvarchar", numeric: false },
      { name: "TestResult", dataType: "nvarchar", numeric: false },
      { name: "SPL1_Result", dataType: "nvarchar", numeric: false },
      { name: "THD_Result", dataType: "nvarchar", numeric: false },
    ],
    layout: [-1, 0, 1, 2, 3],
    rows: [
      { values: ["A", "FAIL", "FAIL", "PASS"], date: "2026-10-01" },
      { values: ["A", "PASS", "PASS", "PASS"], date: "2026-10-01" },
      { values: ["B", "FAIL", "FAIL", "FAIL"], date: "2026-10-01" },
      { values: ["C", "PASS", "PASS", "PASS"], date: "2026-10-01" },
    ],
  };
  const judge = makeJudge(data.table.passValues, data.table.failValues);
  const groups = buildRowGroups(data);

  it("groups rows by ProductID, newest first, loose column match", () => {
    expect(groups.map((g) => [g.key, g.rows.length])).toEqual([["product:A", 2], ["product:B", 1], ["product:C", 1]]);
  });

  it("finds item columns by the _Result suffix and skips TestResult", () => {
    expect(buildNgItemColumns(data.columns)).toEqual([{ index: 2, label: "SPL1" }, { index: 3, label: "THD" }]);
  });

  it("filters row by row and regroups what survives", () => {
    const f = { search: "", modelCode: "", columns: { 1: "fail" } };
    const kept = filterGroups(groups, f, 0, new Set());
    expect(kept.map((g) => [g.key, g.rows.length])).toEqual([["product:A", 1], ["product:B", 1]]);
    expect(filterGroups(groups, { search: "", modelCode: "", columns: {} }, 0, new Set())).toBe(groups);
    expect(filterGroups(groups, { search: "c", modelCode: "", columns: {} }, 0, new Set()).map((g) => g.key)).toEqual(["product:C"]);
    expect(filterGroups(groups, { search: "", modelCode: "", columns: { [-1]: "2026-10-01" } }, 0, new Set())).toHaveLength(3);
  });

  it("counts NG over the newest surviving row per QR", () => {
    expect(filteredSummary(groups, 1, judge)).toMatchObject({ inputCount: 3, ngCount: 2 });
    expect(filteredSummary([], 1, judge).ngRatePpm).toBeNull();
  });

  it("splits NG items into alone / with-others", () => {
    const breakdown = buildNgBreakdown(groups, buildNgItemColumns(data.columns), 1, judge);
    expect(breakdown).toMatchObject({ rowCount: 2, multiCount: 1, comboCount: 2, itemTotalSum: 3, unspecified: 0 });
    const spl1 = breakdown.groups.find((g) => g.label === "SPL1")!;
    expect(spl1.total).toBe(2);
    expect(spl1.children).toEqual([{ label: "alone", count: 1 }, { label: "+ THD", count: 1 }]);
    expect(breakdown.groups[0].label).toBe("SPL1");
  });

  it("flattens expanded groups into one entry per row", () => {
    expect(buildDataLines(groups, new Set())).toHaveLength(3);
    const lines = buildDataLines(groups, new Set(["product:A"]));
    expect(lines.map((l) => l.isChild)).toEqual([false, true, false, false]);
  });

  it("labels and names exports", () => {
    expect(selectedDatesLabel([])).toBe("No date selected");
    expect(selectedDatesLabel(["2026-10-01"])).toBe("2026-10-01 (Thu)");
    expect(selectedDatesLabel(["2026-10-03", "2026-10-01"])).toBe("2 dates · 2026-10-01 ~ 2026-10-03");
    expect(exportDateText(["2026-10-01"])).toBe("20261001");
    expect(exportDateText(["2026-10-03", "2026-10-01"])).toBe("20261001-20261003_2dates");
    expect(formatDay("2026-10-30")).toBe("2026-10-30 (Fri)");
  });
});

suite("GAUSS scatter", () => {
  it("interpolates quantiles like Excel PERCENTILE", () => {
    expect(quantile([1, 2, 3, 4], 0.5)).toBe(2.5);
    expect(quantile([1, 2, 3, 4], 0.25)).toBe(1.75);
    expect(quantile([7], 0.9)).toBe(7);
  });

  it("describes a series with a sample standard deviation", () => {
    const stats = describe([1, 2, 3, 4].map((v) => point(v)));
    expect(stats).toMatchObject({ n: 4, mean: 2.5, median: 2.5, min: 1, max: 4 });
    expect(stats.stdDev).toBeCloseTo(1.2909944, 6);
    expect(describe([]).n).toBe(0);
  });

  it("chooses nice tick steps and pads a flat run", () => {
    expect(niceStep(0.7)).toBe(1);
    expect(niceStep(1.3)).toBe(2);
    expect(niceStep(2.2)).toBe(2.5);
    expect(niceStep(37)).toBe(50);
    const flat = buildScale(10, 10);
    expect(flat.lo).toBeLessThan(10);
    expect(flat.hi).toBeGreaterThan(10);
    expect(buildScale(0, 100).ticks.length).toBeGreaterThan(3);
  });

  it("puts one column per date and keeps the last visible series", () => {
    const value = [point(1, "2026-10-01"), point(2, "2026-10-01"), point(3, "2026-10-03", true)];
    const value2 = [point(10, "2026-10-02")];
    const both = buildScatter({ value, value2, showValue: true, showValue2: true, ngOnly: false });
    expect(both.dates).toEqual(["2026-10-01", "2026-10-02", "2026-10-03"]);
    expect(both.views.map((v) => v.label)).toEqual(["Value", "Value 2"]);
    expect(both.views[0].histogram).toEqual([]);
    const ng = buildScatter({ value, value2, showValue: true, showValue2: false, ngOnly: true });
    expect(ng.totalPoints).toBe(1);
    expect(ng.views[0].ngCount).toBe(1);
    expect(ng.views[0].histogram.length).toBeGreaterThan(0);
  });

  it("spreads box labels at least one pitch apart", () => {
    const model = buildScatter({ value: [1, 2, 3, 4, 5].map((v) => point(v)), value2: [], showValue: true, showValue2: false, ngOnly: false });
    const ys = model.views[0].boxLabels.map((l) => l.textY);
    for (let i = 1; i < ys.length; i++) expect(ys[i] - ys[i - 1]).toBeGreaterThanOrEqual(12 - 1e-9);
  });
});

suite("windowing", () => {
  it("mounts the rows near the viewport, starting on an even row", () => {
    const w = virtualWindow(1000, 32, 3200, 640, 12);
    expect(w.first % 2).toBe(0);
    expect(w.first).toBeLessThanOrEqual(100 - 12);
    expect(w.last).toBeGreaterThanOrEqual(120 + 12);
    expect(w.before + (w.last - w.first) * 32 + w.after).toBe(1000 * 32);
    expect(virtualWindow(5, 32, 0, 640)).toMatchObject({ first: 0, last: 5, after: 0 });
  });
});
