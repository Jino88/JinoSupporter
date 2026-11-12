import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import InfoTab from "./InfoTab";

describe("Model information", () => {
  it("shows only the BMES model group under the title, without the model, stage or other fields", () => {
    const html = renderToStaticMarkup(<InfoTab onDirty={() => {}} onSaved={() => {}} onError={() => {}} model={{ name: "MSM-X526", revision: 1,
      updatedAt: "", updatedBy: "", info: { phase: "MP", build: "PV2", customer: "Acme", owner: "Kim", line: "L1", sopTarget: "2027-03-01", note: "n" } }} />);
    expect(html).toMatch(/<div class="pi-sec-b"><div class="wk-bmes"><div class="wk-bmes-h"><h3>BMES model group<\/h3>/);
    for (const gone of ["wk-info-cols", "wk-build", "Stage", "PV2", ">Model<", "Customer", "Build", "PM owner", ">Line<", "SOP target", "Note", "Acme", "2027"])
      expect(html).not.toContain(gone);
  });
});
