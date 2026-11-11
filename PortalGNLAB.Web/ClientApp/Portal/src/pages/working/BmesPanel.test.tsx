import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { TempModelGroup } from "./BmesPanel";
import { blankProcess, emptyManual } from "./manual";
import type { ManualGroup } from "./manual";
import { addSub, emptyTempGroup } from "./tempGroup";

const temp = addSub(addSub(emptyTempGroup(), null, "MSM-X536B-TOP"), null, "MSM-X536B-BOTTOM");
const [top, bottom] = temp.subGroups.map((s) => s.id);
const table = (process: string, material: string): ManualGroup =>
  ({ materials: [material], processes: [{ ...blankProcess(), name: process, materials: [{ material, qtyPerUnit: 1 }] }] });
const subManual = { [top]: table("Top bonding", "UPPER"), [bottom]: table("Bottom AOI", "LOWER") };

const render = (initial: string | undefined, manual: ManualGroup = emptyManual()) => renderToStaticMarkup(
  <TempModelGroup modelName="MSM-X536B" manual={manual} onManual={() => {}} temp={temp} onTemp={() => {}}
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
    expect(html).toMatch(/class="wk-bg-sub on"[^>]*>.*?value="MSM-X536B-TOP"/);
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
