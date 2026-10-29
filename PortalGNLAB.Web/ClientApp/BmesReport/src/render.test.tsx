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

  it("aligns every header with its column cells (text left, numeric right)", () => {
    const html = TAB_KEYS.map((tabKey) => renderToStaticMarkup(<ReportTabContent report={report} tabKey={tabKey} preferences={preferences} />)).join("");
    const cellsOf = (row: string) => Array.from(row.matchAll(/<(td|th)([^>]*)>(.*?)<\/(?:td|th)>/g), (m) => ({ attrs: m[2], text: m[3] }));
    const tables = html.match(/<table[\s\S]*?<\/table>/g) ?? [];
    expect(tables.length).toBeGreaterThan(0);
    let checked = 0;
    for (const table of tables) {
      const thead = table.match(/<thead>([\s\S]*?)<\/thead>/)![1];
      const headers = cellsOf(thead.match(/<tr[^>]*>(?:(?!<tr)[\s\S])*?<\/tr>\s*$/)![0]);
      const tbody = table.match(/<tbody>([\s\S]*?)<\/tbody>/)![1];
      for (const row of tbody.match(/<tr[^>]*>[\s\S]*?<\/tr>/g) ?? []) {
        const cells = cellsOf(row);
        if (cells.some((cell) => cell.attrs.includes("colSpan"))) continue;
        cells.forEach((cell, index) => {
          if (cell.text === "") return;
          checked += 1;
          expect(headers[index].attrs.includes("numeric"), `${headers[index].text} vs ${cell.text}`).toBe(cell.attrs.includes("numeric"));
        });
      }
    }
    expect(checked).toBeGreaterThan(10);
  });
});
