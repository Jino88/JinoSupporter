import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { BomPanel, DrawingDates, ProcessTables, TempModelGroup } from "./BmesPanel";
import { blankProcess, emptyManual } from "./manual";
import type { ManualGroup } from "./manual";
import { addSub, emptyTempGroup } from "./tempGroup";

const temp = addSub(addSub(emptyTempGroup(), null, "MSM-X536B-TOP"), null, "MSM-X536B-BOTTOM");
const [top, bottom] = temp.subGroups.map((s) => s.id);
const table = (process: string, material: string): ManualGroup =>
  ({ materials: [material], processes: [{ ...blankProcess(), name: process, materials: [{ material, qtyPerUnit: 1 }] }] });
const subManual = { [top]: table("Top bonding", "UPPER"), [bottom]: table("Bottom AOI", "LOWER") };

const render = (initial: string | undefined, manual: ManualGroup = emptyManual()) => renderToStaticMarkup(
  <ProcessTables modelName="MSM-X536B" manual={manual} onManual={() => {}} temp={temp}
    subManual={subManual} onSubManual={() => {}} onError={() => {}} initial={initial} />,
);
const chips = (html: string) => [...html.matchAll(/<button class="wk-chip( on)?">([^<]*)<\/button>/g)].map((m) => (m[1] ? "*" : "") + m[2]);

describe("processes and materials per sub group of a temporary model group", () => {
  it("shows the first sub group's processes and materials, with a chip per sub group", () => {
    const html = render(undefined);
    expect(chips(html)).toEqual(["*MSM-X536B-TOP", "MSM-X536B-BOTTOM"]);
    expect(html).toContain('value="Top bonding"');
    expect(html).not.toContain('value="Bottom AOI"');
    expect(html).toContain(">UPPER<");
    expect(html).not.toContain(">LOWER<");
  });

  it("shows the sub group picked, its own processes and materials only", () => {
    const html = render(bottom);
    expect(chips(html)).toEqual(["MSM-X536B-TOP", "*MSM-X536B-BOTTOM"]);
    expect(html).toContain('value="Bottom AOI"');
    expect(html).not.toContain('value="Top bonding"');
    expect(html).toContain(">LOWER<");
    expect(html).not.toContain(">UPPER<");
  });

  it("keeps the model's own table, entered before the sub groups, as a chip of its own", () => {
    const html = render(undefined, { materials: ["MSU-L20S15-07"], processes: [] });
    expect(chips(html)).toEqual(["*Model MSM-X536B", "MSM-X536B-TOP", "MSM-X536B-BOTTOM"]);
    expect(html).toContain(">MSU-L20S15-07<");
    expect(html).not.toContain('value="Top bonding"');
  });
});

describe("Model Info temporary model group", () => {
  it("keeps the Model › Sub Group tree but not the process tables, and points to the Process tab", () => {
    const html = renderToStaticMarkup(<TempModelGroup modelName="MSM-X536B" temp={temp} onTemp={() => {}}
      subManual={subManual} onSubManual={() => {}} onError={() => {}} />);
    expect(html).toContain('value="MSM-X536B-TOP"');
    expect(html).toContain("Processes, NG items and materials are on the Process tab");
    expect(html).not.toContain("Processes and materials of");
    expect(html).not.toContain('value="Top bonding"');
    expect(html).not.toContain(">UPPER<");
  });
});

describe("BOM / Drawing tab", () => {
  const group = { group: "MODULE", productGroup: "", material: "MSM-X526B", products: ["MSM-X526B-TOP", "MSM-X526B-BOTTOM"], subGroups: [] };
  it("shows the BOM of each product of the linked BMES model group, the first one picked", () => {
    const html = renderToStaticMarkup(<BomPanel modelName="MSM-X526/X626B" saved="MSM-X526B" onError={() => {}} initial={[group]} />);
    expect(html).toContain("<h2>BOM / Drawing</h2>");
    expect(chips(html)).toEqual(["*MSM-X526B-TOP", "MSM-X526B-BOTTOM"]);
  });

  it("says to link a BMES model group in Model Info when there is none", () => {
    const html = renderToStaticMarkup(<BomPanel modelName="NEW-1" saved="" onError={() => {}} initial={[group]} />);
    expect(html).toContain("Link it in Model Info");
    expect(chips(html)).toEqual([]);
  });

  it("writes each drawing's Create Date and Mod. Date as a line beside its link, the time in its tip", () => {
    const list = [
      { pdmno: "BJ-QD-6-033-538", pdmvr: "0", pdtnoTx: "", pdrtx: "PLATE", erdat: "2025-04-08 10:12:29", ardat: "2025-10-02 09:00:00" },
      { pdmno: "BJ-QD-6-033-538", pdmvr: "1", pdtnoTx: "", pdrtx: "PLATE" },
    ];
    expect(renderToStaticMarkup(<DrawingDates list={list} field="erdat" />))
      .toBe('<div class="wk-dwg-date" title="2025-04-08 10:12:29">2025-04-08</div><div class="wk-dwg-date" title="">—</div>');
    expect(renderToStaticMarkup(<DrawingDates list={list} field="ardat" />)).toContain(">2025-10-02<");
    expect(renderToStaticMarkup(<DrawingDates list={[]} field="ardat" />)).toBe('<span class="muted">—</span>');
  });
});
