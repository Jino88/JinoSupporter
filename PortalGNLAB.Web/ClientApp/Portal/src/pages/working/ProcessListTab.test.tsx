import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import ProcessListTab, { tailCols } from "./ProcessListTab";
import { TrContext } from "./i18n";
import { addJig, blankRow, removeJig } from "./processList";
import type { ProcessList, ProcessMatch } from "./processList";

const list: ProcessList = {
  rows: [
    {
      ...blankRow("01"), id: "a", process: "UPPER ARRAY", machine: "IONIZER\nDISPENSOR",
      jigs: [
        { name: "ENCLOSURE JIG", drawingNo: "DWG-100", ver: "V4", qty: "70" },
        { name: "COVER JIG", drawingNo: "DWG-200", ver: "V1", qty: "6" },
      ],
    },
    {
      ...blankRow("02"), id: "b", process: "BONDING", machine: "NON CONTACT DISPENSOR DT", material: "SL233LF / VZ10-L1",
      jigs: [{ name: "COVER JIG", drawingNo: "DWG-200", ver: "V1", qty: "10" }],
    },
    { ...blankRow("03"), id: "c", process: "INSPECTION" },
  ],
  revision: 3, updatedBy: "jhbyun", updatedAt: "2026-11-14T05:20:00Z", source: "JIG list.xlsx",
};
const render = (l: ProcessList, editing = false) =>
  renderToStaticMarkup(<ProcessListTab modelName="MSM-G778B" onError={() => {}} initial={l} editing={editing} />);
const count = (html: string, re: RegExp) => (html.match(re) ?? []).length;

describe("process list tab", () => {
  it("shows the rows as text with the process count, and no input box", () => {
    const html = render(list);
    expect(html).toContain("<h2>JIG/MC List</h2>");
    expect(html).not.toContain("<h2>Process List</h2>");
    expect(html).toContain("<b>JIG list.xlsx</b>");
    expect(html).toContain("<b>UPPER ARRAY</b>");
    expect(html).toContain("IONIZER\nDISPENSOR");
    expect(html).toContain("SL233LF / VZ10-L1");
    expect(html).toMatch(/Processes <b>3<\/b>/);
    expect(html).toContain("Import JIG / Machine list");
    expect(html).not.toContain('class="inp');
    expect(html).not.toContain("<textarea");
  });

  it("has no Workers or Std. time column or summary, and orders the columns", () => {
    for (const html of [render(list), render(list, true)]) {
      expect(html).not.toMatch(/Workers|Slowest|Std\. time|Standard time/);
      expect(html).not.toContain('type="number"');
    }
    const labels = (html: string) => [...html.match(/<table class="wk-it pl-t"><thead>[^]*?<[/]thead>/)![0].matchAll(/<th[^>]*>([^<]*)<\/th>/g)].map((m) => m[1]).filter(Boolean);
    // to read, Tester and Note are left out while no process has one; to edit, every column is there
    expect(labels(render(list))).toEqual(["No.", "Process", "Machine", "JIG", "Drawing No", "Ver.", "Qty", "Material"]);
    expect(labels(render(list, true))).toEqual(["No.", "Process", "Machine", "JIG", "Drawing No", "Ver.", "Qty", "Tester", "Material", "Note"]);
  });

  it("shows only the Tester / Material / Note columns some process fills in, all of them to edit", () => {
    const keys = (rows: typeof list.rows, edit = false) => tailCols(rows, edit).map(([k]) => k);
    expect(keys(list.rows)).toEqual(["material"]);
    expect(keys([{ ...list.rows[2], tester: "AOI", note: " x " }])).toEqual(["tester", "note"]);
    expect(keys([{ ...list.rows[2], note: "  " }])).toEqual([]);
    expect(keys(list.rows, true)).toEqual(["tester", "material", "note"]);
  });

  it("puts every JIG on a sub-row of its process, with the cells of the process spanning them", () => {
    const html = render(list);
    const body = html.match(/<table class="wk-it pl-t"><thead>[^]*<[/]table>/)![0];
    // 2 JIGs + 1 JIG + no JIG (one row) = 4 rows
    expect(count(body, /<tr class="pl-(first|sub) pl-g[01]">/g)).toBe(4);
    expect(count(body, /<tr class="pl-first pl-g[01]">/g)).toBe(3);
    expect(count(body, /<tr class="pl-sub pl-g[01]">/g)).toBe(1);
    // No., Process, Machine and Material of the first process span its 2 JIG rows
    expect(body).toContain('<td rowSpan="2" class="pl-no"><span class="pl-badge">01</span></td>');
    expect(body).toContain('<td rowSpan="2" class="l"><b>UPPER ARRAY</b></td>');
    expect(count(body, /rowSpan="2"/g)).toBe(4);
    // a process without JIG is one row, its JIG cells empty
    expect(body).toContain('<td rowSpan="1" class="pl-no"><span class="pl-badge">03</span></td>');
    // the JIG name, drawing no., version and quantity are cells of their own
    expect(body).toContain('<td class="l pl-jigcell pl-jfirst">ENCLOSURE JIG</td><td class="l pl-jigcell">DWG-100</td><td class="pl-narrow pl-jigcell">V4</td><td class="pl-narrow pl-jigcell pl-qty">70</td>');
    expect(body).toContain('<td class="l pl-jigcell pl-jfirst">COVER JIG</td><td class="l pl-jigcell">DWG-200</td><td class="pl-narrow pl-jigcell">V1</td><td class="pl-narrow pl-jigcell pl-qty">6</td>');
    // JIGs are not joined into a multi-line text
    expect(body).not.toContain("ENCLOSURE JIG\nCOVER JIG");
  });

  it("shows the JIG total at the top as a small table with a head, open: one line per kind of JIG with its biggest quantity, and the sum", () => {
    const html = render(list);
    // ENCLOSURE JIG 70 + COVER JIG max(6, 10) = 80 over 2 kinds
    expect(html).toMatch(/JIG total: <b>2<\/b> kinds · <b>80<\/b> pcs/);
    const sum = html.match(/<details class="pl-jigsum"[^>]*>[^]*?<\/details>/)![0];
    expect(sum).toMatch(/^<details class="pl-jigsum" open/);
    const head = [...sum.matchAll(/<th[^>]*>([^<]*)<\/th>/g)].map((m) => m[1]);
    expect(head).toEqual(["JIG", "Drawing No", "Ver.", "Qty"]);
    const lines = [...sum.matchAll(/<tbody>[^]*?<\/tbody>/g)][0][0].split("</tr>").filter((x) => x.includes("<td")).map((m) => m.replace(/<[^>]+>/g, "|").replace(/\|+/g, "|"));
    expect(lines).toEqual(["|ENCLOSURE JIG|DWG-100|V4|70|", "|COVER JIG|DWG-200|V1|10|"]);
    expect(html.indexOf("pl-jigsum")).toBeLessThan(html.indexOf('<table class="wk-it pl-t">'));
  });

  it("gives every process its own zebra group", () => {
    const body = render(list);
    expect(count(body, /pl-g0/g)).toBe(3);
    expect(count(body, /pl-g1/g)).toBe(1);
  });

  describe("translation", () => {
    const map: Record<string, string> = { "UPPER ARRAY": "상부 배열", "BONDING": "BONDING", "Two jigs": "x" };
    const withTr = (editing = false) => renderToStaticMarkup(
      <TrContext.Provider value={(t) => map[t] ?? t}>
        <ProcessListTab modelName="M" onError={() => {}} initial={{ ...list, rows: [{ ...list.rows[0], note: "UPPER ARRAY" }, list.rows[1]] }} editing={editing} />
      </TrContext.Provider>);

    it("shows the translation under the process name and the note, only when it differs", () => {
      const html = withTr();
      expect(html).toContain('<b>UPPER ARRAY</b><div class="pl-tr">상부 배열</div>');
      expect(html).toContain('UPPER ARRAY<div class="pl-tr">상부 배열</div></td>');
      // BONDING translates to itself: no line
      expect(count(html, /class="pl-tr"/g)).toBe(2);
      expect(html).not.toContain("<b>BONDING</b><div");
    });

    it("shows no translation line in the language as written", () => {
      expect(render(list)).not.toContain("pl-tr");
    });

    it("keeps the edit boxes as written", () => {
      const html = withTr(true);
      expect(html).toContain('value="UPPER ARRAY"');
      expect(html).not.toContain("상부 배열");
    });
  });

  it("edits with boxes only after Edit: JIG, Drawing No, Ver. and Qty per JIG row, + JIG per process", () => {
    const html = render(list, true);
    expect(html.match(/aria-label="Process"/g)).toHaveLength(3);
    expect(html).toContain('value="BONDING"');
    for (const label of ["JIG", "Drawing No", "Ver\\.", "Qty"]) expect(count(html, new RegExp(`<input[^>]*aria-label="${label}"`, "g"))).toBe(3);
    expect(html).toContain('value="DWG-100"');
    expect(count(html, /aria-label="Remove the JIG"/g)).toBe(3);
    expect(count(html, /\+ JIG<\/button>/g)).toBe(3);
    expect(count(html, /aria-label="Remove the row"/g)).toBe(3);
    for (const label of ["Machine", "Tester", "Material", "Note"]) expect(count(html, new RegExp(`aria-label="${label}"`, "g"))).toBe(3);
    // a textarea grows to its lines (min 1) so nothing is cut off
    expect(html).toMatch(/<textarea[^>]*rows="2"[^>]*aria-label="Machine"[^>]*>IONIZER\nDISPENSOR<\/textarea>/);
    expect(html).toMatch(/<textarea[^>]*rows="1"[^>]*aria-label="Note"/);
    expect(html).toContain(">Save</button>");
    expect(html).toContain(">+ Row</button>");
    expect(html).not.toContain(">Edit</button>");
    expect(html).not.toContain("Import JIG / Machine list");
  });

  it("adds and removes JIGs in the draft rows", () => {
    // what the + JIG and ✕ buttons of the edit rows do to the draft
    const added = addJig(list.rows[2]);
    expect(added.jigs).toHaveLength(1);
    const html = render({ ...list, rows: [added, removeJig(list.rows[0], 0)] }, true);
    expect(count(html, /aria-label="Remove the JIG"/g)).toBe(2);
    expect(html).toContain('<td rowSpan="1" class="pl-no"><input class="inp sm pl-no" aria-label="No." value="03"');
    expect(html).not.toContain("ENCLOSURE JIG");
    expect(html).toContain('value="COVER JIG"');
  });

  it("shows data of an older response without jigs", () => {
    const old = { ...list, rows: list.rows.map(({ jigs: _j, ...r }) => r as unknown as ProcessList["rows"][number]) };
    for (const editing of [false, true]) {
      const html = render(old, editing);
      expect(html).toContain("UPPER ARRAY");
      expect(html).not.toContain("undefined");
    }
    expect(render(old)).toMatch(/<b>0<\/b> kinds/);
  });

  it("asks for the workbook when there are no rows", () => {
    const html = render({ rows: [], revision: 0, updatedBy: "", updatedAt: null, source: "" });
    expect(html).toContain("No processes yet.");
    expect(html).toContain('accept=".xlsx,.xlsm"');
    expect(html).not.toContain("<table");
  });

  it("has an AI match button with its own workbook input next to Import", () => {
    const html = render(list);
    expect(html).toContain("✦ AI match");
    expect(html).toContain('aria-label="AI match workbook"');
    expect(html.match(/accept="\.xlsx,\.xlsm"/g)).toHaveLength(2);
  });

  describe("AI match result", () => {
    const match: ProcessMatch = {
      rows: [
        { ...blankRow("01"), process: "GLUING", machine: "DISPENSOR", jigs: [{ name: "GLUE JIG", drawingNo: "DWG-7", ver: "V2", qty: "3" }] },
        { ...blankRow("02"), process: "CURING" },
      ],
      note: "Two jigs had no process.", unmatched: ["JIG-77", "TESTER-9"], basis: "instructions", model: "gpt-5.6-luna · high",
    };
    const renderMatch = (m: ProcessMatch, l = list) =>
      renderToStaticMarkup(<ProcessListTab modelName="M" onError={() => {}} initial={l} match={m} />);

    it("puts the rows in the draft as text with Save, Cancel and Edit, and explains them above it", () => {
      const html = renderMatch(match);
      expect(html).toContain("<b>GLUING</b>");
      expect(html).toContain("GLUE JIG");
      expect(html).toContain("DWG-7");
      expect(html).toContain("<b>CURING</b>");
      expect(html).not.toContain("BONDING");
      expect(html).toContain(">Save</button>");
      expect(html).toContain(">Cancel</button>");
      expect(html).toContain(">Edit</button>");
      expect(html).not.toContain(">+ Row</button>");
      expect(html).not.toContain('class="inp');
      expect(html).not.toContain("<textarea");
      expect(html).toMatch(/<b>1<\/b> kinds · <b>3<\/b> pcs/);
      expect(html).toContain("Process order from the Work Instructions");
      expect(html).toContain("gpt-5.6-luna · high");
      expect(html).toContain("Two jigs had no process.");
      expect(html).toContain("JIG-77");
      expect(html.indexOf("pl-ai")).toBeLessThan(html.indexOf('<table class="wk-it pl-t">'));
    });

    it("reads as information about an unsaved result, not as an error, and says why items have no process", () => {
      const html = renderMatch(match);
      expect(html).toContain("AI match result (not saved yet)");
      expect(html).toMatch(/AI note(<!-- -->)?:<\/span>/);
      expect(html).toContain("Items not put on a process (2)");
      expect(html).toContain("not in this line&#x27;s process list, or have no process number. Add them with Edit if needed.");
      expect(html).not.toContain("Not placed in a process");
      expect(html).not.toMatch(/pl-err|role="alert"/);
    });

    it("words the basis and leaves the list out when nothing was unmatched", () => {
      const html = renderMatch({ ...match, basis: "workbook", unmatched: [] });
      expect(html).toContain("Process order from the workbook&#x27;s process numbers");
      expect(html).not.toContain("pl-ai-un");
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
