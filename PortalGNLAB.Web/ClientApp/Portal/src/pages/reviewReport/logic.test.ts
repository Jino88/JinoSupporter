import { describe, expect, it } from "vitest";
import { dash, filterReports, groupPictures, removePicture } from "./logic";
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
});
