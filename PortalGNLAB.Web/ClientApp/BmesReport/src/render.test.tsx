import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import fixture from "../test/fixtures/report-v1.json";
import { validateReportDocument } from "./api";
import { TAB_KEYS } from "./contract";
import { ReportTabContent } from "./components/ReportTabContent";
import { ReportShell, TAB_LABELS } from "./components/ReportShell";
import { preferencesFromDefaults } from "./logic";

const report = validateReportDocument(fixture);
const preferences = preferencesFromDefaults(report.viewerDefaults);

describe("eight tab render paths", () => {
  it.each(TAB_KEYS)("renders %s from the representative fixture", (tabKey) => {
    const html = renderToStaticMarkup(<ReportTabContent report={report} tabKey={tabKey} preferences={preferences} />);
    expect(html.length).toBeGreaterThan(100);
    expect(html).not.toContain("undefined");
  });

  it("emits accessible tab semantics and all labels", () => {
    const html = renderToStaticMarkup(<ReportShell report={report} />);
    expect(html).toContain('role="tablist"');
    expect(html).toContain('role="tabpanel"');
    for (const tabKey of TAB_KEYS) expect(html).toContain(TAB_LABELS[tabKey]);
  });

  it("renders English UI text and two-row period headers", () => {
    const shell = renderToStaticMarkup(<ReportShell report={report} />);
    for (const text of ["BMES Integrated Report", "Generated", "Calculation", "Periods shown", "Defaults", "Print"]) expect(shell).toContain(text);
    const daily = renderToStaticMarkup(<ReportTabContent report={report} tabKey="daily" preferences={preferences} />);
    expect(daily).toContain("Daily summary");
    expect(daily).toContain('scope="colgroup"');
  });
});
