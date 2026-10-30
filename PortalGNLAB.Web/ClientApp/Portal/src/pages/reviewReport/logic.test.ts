import { describe, expect, it } from "vitest";
import {
  addTag,
  dash,
  editLine,
  filterReports,
  groupPictures,
  placePicture,
  removePicture,
  removeTag,
  renameMetric,
  renameTable,
  replaceText,
  setArms,
  setMetricValue,
} from "./logic";
import { freshReport, type Placement, type ReportDoc } from "./types";
describe("review report controls", () => {
  it("renumbers pictures and placements together without mutating the report", () => {
    const report = freshReport();
    report.images = [1, 2, 4].map((no) => ({
      no,
      fileName: "x",
      mediaType: "image/png",
      data: "",
    }));
    report.doc = {
      images: [1, 2, 4].map((no) => ({
        no,
        section: "results",
        caption: "",
        group: "",
      })),
    } as ReportDoc;
    const next = removePicture(report, 2);
    expect(next.images.map((p) => p.no)).toEqual([1, 2]);
    expect(next.doc!.images.map((p) => p.no)).toEqual([1, 2]);
    expect(report.images.map((p) => p.no)).toEqual([1, 2, 4]);
  });
  it("groups named pictures case-insensitively and leaves unnamed pictures separate", () => {
    const placements = ["A", "", "a", ""].map(
      (group, i) => ({ no: i + 1, group }) as Placement,
    );
    expect(groupPictures(placements).map((g) => g.map((p) => p.no))).toEqual([
      [1, 3],
      [2],
      [4],
    ]);
  });
  it("filters title and author without changing order", () => {
    expect(
      filterReports(
        [{ id: "1", title: "Motor", createdBy: "Sam", updatedAt: "" }],
        "SAM",
      ),
    ).toHaveLength(1);
    expect(dash(" test ")).toBe("- test");
    expect(dash("- test")).toBe("- test");
  });
  it("adds and removes person-set tags without duplicates", () => {
    const doc = { tags: ["6.72"] } as ReportDoc;
    expect(addTag(doc, " laser ").tags).toEqual(["6.72", "laser"]);
    expect(addTag(doc, "6.72")).toBe(doc);
    expect(addTag(doc, "   ")).toBe(doc);
    expect(removeTag(addTag(doc, "laser"), 0).tags).toEqual(["laser"]);
    expect(doc.tags).toEqual(["6.72"]);
  });
  it("edits a report line and drops it when cleared", () => {
    expect(editLine(["a", "b"], 1, "B")).toEqual(["a", "B"]);
    expect(editLine(["a", "b"], 0, "  ")).toEqual(["b"]);
  });
  const sample = (): ReportDoc =>
    ({
      title: "Film M-20",
      resultTitle: "",
      tags: ["Film M-20", "SPL"],
      conditions: [
        {
          id: "C1",
          label: "Film M-20",
          role: "test",
          date: "",
          model: "M-20",
          details: "",
          table: "Film M-20",
        },
      ],
      metrics: [
        {
          id: "input",
          name: "Input",
          unit: "pcs",
          kind: "input",
          rateOf: "",
          values: [{ armId: "C1", value: "930" }],
        },
      ],
      images: [{ no: 1, caption: "M-20 note", group: "", section: "results" }],
    }) as unknown as ReportDoc;
  it("writes a result-table cell back to its condition or metric", () => {
    const doc = sample();
    expect(setArms(doc, ["c1"], { model: "11-20" }).conditions[0].model).toBe(
      "11-20",
    );
    // A row made from a metric value with no listed condition gets one.
    expect(setArms(doc, ["C9"], { label: "New" }).conditions[1]).toMatchObject({
      id: "C9",
      label: "New",
    });
    expect(
      renameTable(doc, "film m-20", "FILM-11-20").conditions[0].table,
    ).toBe("FILM-11-20");
    expect(renameTable(doc, "", "Result check").resultTitle).toBe(
      "Result check",
    );
    expect(setMetricValue(doc, "INPUT", "C1", "940").metrics[0].values).toEqual(
      [{ armId: "C1", value: "940" }],
    );
    expect(
      setMetricValue(doc, "input", "C2", "5").metrics[0].values,
    ).toHaveLength(2);
    expect(
      renameMetric(doc, "input", "Input qty (ea)").metrics[0],
    ).toMatchObject({ name: "Input qty", unit: "ea" });
    expect(renameMetric(doc, "input", "Input").metrics[0].unit).toBe("");
    expect(renameMetric(doc, "missing", "X")).toBe(doc);
    expect(doc.conditions[0].model).toBe("M-20");
  });
  it("replaces a word everywhere except the ids that link the report", () => {
    const doc = sample();
    doc.conditions[0].id = "M-20";
    doc.metrics[0].values[0].armId = "M-20";
    const { doc: next, count } = replaceText(doc, "M-20", "11-20");
    expect(count).toBe(6);
    expect(next.title).toBe("Film 11-20");
    expect(next.tags).toEqual(["Film 11-20", "SPL"]);
    expect(next.conditions[0]).toMatchObject({
      id: "M-20",
      label: "Film 11-20",
      model: "11-20",
      table: "Film 11-20",
    });
    expect(next.metrics[0].values[0].armId).toBe("M-20");
    expect(next.images[0]).toMatchObject({ no: 1, caption: "11-20 note" });
    expect(replaceText(doc, "nothing", "x")).toEqual({ doc, count: 0 });
  });
  it("places a picture in a section, adding one not yet in the report", () => {
    const doc = sample();
    expect(placePicture(doc, 1, "purpose").images).toEqual([
      { no: 1, caption: "M-20 note", group: "", section: "purpose" },
    ]);
    expect(placePicture(doc, 2, "conditions").images[1]).toEqual({
      no: 2,
      caption: "",
      group: "",
      section: "conditions",
    });
  });
});
