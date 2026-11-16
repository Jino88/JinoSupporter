import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import DrawingPicker from "./DrawingPicker";
import { blankItem } from "./request";

describe("drawing no picker", () => {
  it("lists the drawings the DB suggests for the row's name, each to pick", () => {
    const item = { ...blankItem(4, "MSU-L15S10-20"), name: "DOME LASER GUIDE BASE" };
    const html = renderToStaticMarkup(<DrawingPicker item={item} model="MSU-L15S10-20" onPick={() => {}} onClose={() => {}} initial={[
      { drawingNo: "BJKP18-08-I-0000-00", name: "CD LASER GUIDE BASE", model: "BRS-161242", category: "MC", assy: 1, latest: 1, score: 0.8 },
      { drawingNo: "BJGP26-22-T-0001-00", name: "(SUB1)DOME LASER CUTTING JIG", model: "MSU-L15S10-20", category: "JIG", assy: 0, latest: 0, score: 0.6 },
    ]} />);
    expect(html).toContain("4. DOME LASER GUIDE BASE");
    expect(html).toContain('value="DOME LASER GUIDE BASE"');
    expect(html).toContain("BJKP18-08-I-0000-00");
    expect(html).toContain("(SUB1)DOME LASER CUTTING JIG");
    expect(html.match(/>Pick</g)).toHaveLength(2);
    expect(html).toContain(">old<");
    expect(html).toContain(">ASSY<");
  });

  it("says so when the DB has nothing like the name", () => {
    const html = renderToStaticMarkup(<DrawingPicker item={{ ...blankItem(1), name: "ZZZ" }} model="" onPick={() => {}} onClose={() => {}} initial={[]} />);
    expect(html).toContain("No drawing in the DB looks like this name.");
    const ofModel = renderToStaticMarkup(<DrawingPicker item={{ ...blankItem(1), name: "ZZZ" }} model="MSM-G778B" onPick={() => {}} onClose={() => {}} initial={[]} />);
    expect(ofModel).toContain("No drawing of this model in the DB looks like this name.");
  });

  it("shows the versions of the row's drawing number first, the one in the row not to pick again", () => {
    const v = (no: string, latest = 0) => ({ drawingNo: no, name: "LASER MARKING ARRAY JIG", model: "MSM-G778B", category: "JIG", latest });
    const item = { ...blankItem(11, "MSM-G778B"), name: "LASER MARKING ARRAY JIG", drawingNo: "BJJ2605-MO-05-A00-02" };
    const html = renderToStaticMarkup(<DrawingPicker item={item} model="MSM-G778B" onPick={() => {}} onClose={() => {}} initial={[]}
      versions={[v("BJJ2605-MO-05-A00-00"), v("BJJ2605-MO-05-A00-02"), v("BJJ2605-MO-05-A00-05", 1)]} />);
    expect(html).toContain("Versions of this drawing no in the DB");
    expect(html).toContain('aria-label="Pick BJJ2605-MO-05-A00-00"');
    expect(html).toContain('aria-label="Pick BJJ2605-MO-05-A00-05"');
    expect(html).not.toContain('aria-label="Pick BJJ2605-MO-05-A00-02"');
    expect(html).toMatch(/<tr class="rq-pick-now"><td class="mono">BJJ2605-MO-05-A00-02<\/td>/);
    expect(html).toContain("BJJ2605-MO-05-A00-05<span class=\"rq-tag none\">latest</span>");
    expect(html.indexOf("Versions of this drawing no")).toBeLessThan(html.indexOf("Or find it by name"));
    // with no other version there is no version list
    expect(renderToStaticMarkup(<DrawingPicker item={item} model="" onPick={() => {}} onClose={() => {}} initial={[]} versions={[v("BJJ2605-MO-05-A00-02")]} />))
      .not.toContain("Versions of this drawing no");
  });
});
