import { describe, expect, it } from "vitest";
import {
  addTag,
  dash,
  editLine,
  filterReports,
  groupPictures,
  removePicture,
  removeTag,
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
});
