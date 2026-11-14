import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import ProcessListTab from "./ProcessListTab";
import { blankRow } from "./processList";
import type { ProcessList, ProcessMatch } from "./processList";

const list: ProcessList = {
  rows: [
    { ...blankRow("01"), id: "a", process: "UPPER ARRAY", workers: 1, stdTime: 12.312, machine: "IONIZER", jig: "ENCLOSURE DISPENSOR JIG (V4) ×70" },
    { ...blankRow("02"), id: "b", process: "BONDING", workers: 2, stdTime: 6.156, machine: "NON CONTACT DISPENSOR DT", material: "SL233LF / VZ10-L1" },
  ],
  revision: 3, updatedBy: "jhbyun", updatedAt: "2026-11-14T05:20:00Z", source: "JIG list.xlsx",
};
const render = (l: ProcessList, editing = false) =>
  renderToStaticMarkup(<ProcessListTab modelName="MSM-G778B" onError={() => {}} initial={l} editing={editing} />);

describe("process list tab", () => {
  it("shows the rows with the people and the slowest process", () => {
    const html = render(list);
    expect(html).toContain("<h2>Process List</h2>");
    expect(html).toContain("<b>JIG list.xlsx</b>");
    expect(html).toContain("<b>UPPER ARRAY</b>");
    expect(html).toContain("ENCLOSURE DISPENSOR JIG (V4) ×70");
    expect(html).toContain("SL233LF / VZ10-L1");
    expect(html).toMatch(/Workers <b>3<\/b>/);
    expect(html).toMatch(/Slowest <b>12.312s<\/b> \(01 UPPER ARRAY\)/);
    expect(html).toContain("Import JIG / Machine list");
    expect(html).not.toContain("<input class=\"inp sm pl-proc\"");
  });

  it("asks for the workbook when there are no rows", () => {
    const html = render({ rows: [], revision: 0, updatedBy: "", updatedAt: null, source: "" });
    expect(html).toContain("No processes yet.");
    expect(html).toContain('accept=".xlsx,.xlsm"');
    expect(html).not.toContain("<table");
  });

  it("edits every cell of every row", () => {
    const html = render(list, true);
    expect(html.match(/aria-label="Process"/g)).toHaveLength(2);
    expect(html).toContain('value="BONDING"');
    expect(html).toContain('value="6.156"');
    expect(html).toContain(">Save</button>");
    expect(html).toContain(">+ Row</button>");
    expect(html).not.toContain("Import JIG / Machine list");
  });

  it("has an AI match button with its own workbook input next to Import", () => {
    const html = render(list);
    expect(html).toContain("✦ AI match");
    expect(html).toContain('aria-label="AI match workbook"');
    expect(html.match(/accept="\.xlsx,\.xlsm"/g)).toHaveLength(2);
  });

  describe("AI match result", () => {
    const match: ProcessMatch = {
      rows: [{ ...blankRow("01"), process: "GLUING", machine: "DISPENSOR" }, { ...blankRow("02"), process: "CURING" }],
      note: "Two jigs had no process.", unmatched: ["JIG-77", "TESTER-9"], basis: "instructions", model: "gpt-5.6-luna · high",
    };
    const renderMatch = (m: ProcessMatch, l = list) =>
      renderToStaticMarkup(<ProcessListTab modelName="M" onError={() => {}} initial={l} match={m} />);

    it("puts the rows in the draft and explains them above it", () => {
      const html = renderMatch(match);
      expect(html).toContain('value="GLUING"');
      expect(html).toContain('value="CURING"');
      expect(html).not.toContain('value="BONDING"');
      expect(html).toContain(">Save</button>");
      expect(html).toContain("Process order from the Work Instructions");
      expect(html).toContain("gpt-5.6-luna · high");
      expect(html).toContain("Two jigs had no process.");
      expect(html).toContain("<details");
      expect(html).toContain("JIG-77");
      expect(html.indexOf("pl-ai")).toBeLessThan(html.indexOf("<table"));
    });

    it("words the basis and leaves the list out when nothing was unmatched", () => {
      const html = renderMatch({ ...match, basis: "workbook", unmatched: [] });
      expect(html).toContain("Process order from the workbook&#x27;s process numbers");
      expect(html).not.toContain("<details");
      expect(renderMatch({ ...match, basis: "process-list" })).toContain("Process order from the current Process List");
    });

    it("shows no box without a result", () => {
      expect(render(list, true)).not.toContain("pl-ai");
    });

    it("keeps the Import label while the AI is matching", () => {
      const html = renderToStaticMarkup(<ProcessListTab modelName="M" onError={() => {}} initial={list} matching />);
      expect(html).toContain("AI is matching…");
      expect(html).toContain("Import JIG / Machine list");
      expect(html).not.toContain("Reading the workbook…");
    });
  });
});
