import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import InfoTab from "./InfoTab";

describe("Model information", () => {
  it("shows the model and its stage beside the BMES model group, without the other fields", () => {
    const html = renderToStaticMarkup(<InfoTab onDirty={() => {}} onSaved={() => {}} onError={() => {}} model={{ name: "MSM-X526", revision: 1,
      updatedAt: "", updatedBy: "", info: { phase: "MP", build: "MP", customer: "Acme", owner: "Kim", line: "L1", sopTarget: "2027-03-01", note: "n" } }} />);
    const cols = html.slice(html.indexOf('class="wk-info-cols"'));
    expect(cols).toMatch(/^class="wk-info-cols"><div class="wk-form wk-info">/);
    expect(cols).toContain('<label for="wk-build">Stage</label><input id="wk-build" class="inp sm" placeholder="PV2" value="MP"/>');
    expect(cols).toMatch(/<\/div><\/div><div class="wk-bmes"><div class="wk-bmes-h"><h3>BMES model group<\/h3>/);
    for (const gone of ["Customer", "Build", "PM owner", ">Line<", "SOP target", "Note", "Acme", "2027"]) expect(html).not.toContain(gone);
  });
});
