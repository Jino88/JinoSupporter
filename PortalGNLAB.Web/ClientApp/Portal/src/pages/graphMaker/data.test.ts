import { describe, it, expect } from "vitest";
import {
  normalizeGraphData,
  parseTable,
  columnName,
  exampleData,
} from "./data";
describe("graph records input", () => {
  it("preserves pasted tables", () =>
    expect(normalizeGraphData("A\tB\n1\t2")).toBe("A\tB\n1\t2"));
  it("unions columns and preserves missing, zero and multiline values", () =>
    expect(
      normalizeGraphData(
        JSON.stringify({
          records: [{ a: 0, b: "line\nbreak" }, null, { b: 2, c: 3 }],
        }),
      ),
    ).toBe("a\tb\tc\n0\tline break\t\n\t2\t3"));
  it("rejects unsupported JSON", () =>
    expect(() => normalizeGraphData("{}")).toThrow("records"));
});

describe("Excel data table", () => {
  it("detects delimiters and preserves blank cells and unique headings", () => {
    expect(
      parseTable("skip;skip;skip\r\nA;a;\r\n1;;3\r\n", "auto", 2, false),
    ).toEqual({ headers: ["A", "a_2", "Column"], rows: [["1", "", "3"]] });
  });
  it("keeps all rows in no-header mode and clamps a header beyond the input", () => {
    expect(parseTable("1\t2\n3\t4", "auto", 1, true)).toEqual({
      headers: ["Column 1", "Column 2"],
      rows: [
        ["1", "2"],
        ["3", "4"],
      ],
    });
    expect(parseTable("a,b\n1,2", ",", 99, false).rows).toEqual([]);
  });
  it("labels columns beyond Z", () =>
    expect([0, 25, 26, 51, 52, 701].map(columnName)).toEqual([
      "A",
      "Z",
      "AA",
      "AZ",
      "BA",
      "ZZ",
    ]));
  it("limits only example heatmaps to first five distinct conditions", () => {
    const text = JSON.stringify({
      records: Array.from({ length: 7 }, (_, i) => ({
        IR_Code: `X${i}`,
        VP_Mold: "Y",
        NG: i,
      })),
    });
    expect(parseTable(exampleData(text), "auto", 1, false).rows).toHaveLength(
      5,
    );
    expect(
      parseTable(normalizeGraphData(text), "auto", 1, false).rows,
    ).toHaveLength(7);
  });
});
