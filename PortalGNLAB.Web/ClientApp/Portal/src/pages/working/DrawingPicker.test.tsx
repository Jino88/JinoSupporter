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
  });
});
