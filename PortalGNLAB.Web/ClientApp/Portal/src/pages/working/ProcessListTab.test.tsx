import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import ProcessListTab from "./ProcessListTab";
import { blankRow } from "./processList";
import type { ProcessList, ProcessMatch } from "./processList";

const list: ProcessList = {
  rows: [
    { ...blankRow("01"), id: "a", process: "UPPER ARRAY", machine: "IONIZER\nDISPENSOR", jig: "ENCLOSURE JIG\nCOVER JIG", jigVer: "V4\nV1", jigQty: "70\n6" },
    { ...blankRow("02"), id: "b", process: "BONDING", machine: "NON CONTACT DISPENSOR DT", material: "SL233LF / VZ10-L1" },
  ],
  revision: 3, updatedBy: "jhbyun", updatedAt: "2026-11-14T05:20:00Z", source: "JIG list.xlsx",
};
const render = (l: ProcessList, editing = false) =>
  renderToStaticMarkup(<ProcessListTab modelName="MSM-G778B" onError={() => {}} initial={l} editing={editing} />);

describe("process list tab", () => {
  it("shows the rows with the process count", () => {
    const html = render(list);
    expect(html).toContain("<h2>Process List</h2>");
    expect(html).toContain("<b>JIG list.xlsx</b>");
    expect(html).toContain("<b>UPPER ARRAY</b>");
    expect(html).toContain("ENCLOSURE JIG\nCOVER JIG");
    expect(html).toContain("SL233LF / VZ10-L1");
    expect(html).toMatch(/Processes <b>2<\/b>/);
    expect(html).toContain("Import JIG / Machine list");
    expect(html).not.toContain("<input class=\"inp sm pl-proc\"");
  });

  it("has no Workers or Std. time column or summary, and orders the columns", () => {
    for (const html of [render(list), render(list, true)]) {
      expect(html).not.toMatch(/Workers|Slowest|Std\. time|Standard time/);
      expect(html).not.toContain('type="number"');
    }
    const head = render(list).match(/<thead>.*<\/thead>/)![0];
    const labels = [...head.matchAll(/<th[^>]*>([^<]*)<\/th>/g)].map((m) => m[1]);
    expect(labels).toEqual(["No.", "Process", "Machine", "JIG", "Ver.", "Qty", "Tester", "Material", "Note"]);
  });

  it("shows one item per line, JIG, Ver. and Qty side by side", () => {
    const html = render(list);
    expect(html).toContain("IONIZER\nDISPENSOR");
    expect(html).toContain("V4\nV1");
    expect(html).toContain("70\n6");
    expect(html).toContain("pl-cell pl-narrow");
    // a long JIG name is not wrapped onto a second line, which would put it beside the next version and quantity
    expect(html).toContain('class="l pl-cell pl-lines">ENCLOSURE JIG\nCOVER JIG</td>');
    expect(html).toContain('class="pl-cell pl-narrow pl-lines">V4\nV1</td>');
    expect(html).toContain('class="l pl-cell">SL233LF / VZ10-L1</td>');
  });

  it("edits JIG, Ver. and Qty in boxes of the same height that do not wrap", () => {
    const html = render({ ...list, rows: [{ ...list.rows[0], jigVer: "V4", jigQty: "" }] }, true);
    for (const label of ["JIG", "Ver\\.", "Qty"]) expect(html).toMatch(new RegExp(`<textarea[^>]*wrap="off"[^>]*rows="2"[^>]*aria-label="${label}"`));
    expect(html).toMatch(/<textarea[^>]*wrap="off"[^>]*rows="2"[^>]*aria-label="Machine"/);
    expect(html).toMatch(/<textarea class="inp sm pl-text" rows="1" aria-label="Note"/);
  });

  it("shows data saved without JIG version and quantity", () => {
    const old = { ...list, rows: list.rows.map(({ jigVer: _v, jigQty: _q, ...r }) => r as unknown as ProcessList["rows"][number]) };
    for (const editing of [false, true]) {
      const html = render(old, editing);
      expect(html).toContain("ENCLOSURE JIG");
      expect(html).not.toContain("undefined");
    }
    expect(render(old, true)).toContain('aria-label="Ver."');
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
    for (const label of ["Machine", "JIG", "Ver.", "Qty", "Tester"]) expect(html.match(new RegExp(`aria-label="${label.replace(".", "[.]")}"`, "g"))).toHaveLength(2);
    // a textarea grows to its lines (min 1) so nothing is cut off
    expect(html).toMatch(/<textarea[^>]*rows="2"[^>]*aria-label="JIG"[^>]*>ENCLOSURE JIG\nCOVER JIG<\/textarea>/);
    expect(html).toMatch(/<textarea[^>]*rows="2"[^>]*aria-label="Ver\."[^>]*>V4\nV1<\/textarea>/);
    expect(html).toMatch(/<textarea[^>]*rows="1"[^>]*aria-label="Note"/);
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
