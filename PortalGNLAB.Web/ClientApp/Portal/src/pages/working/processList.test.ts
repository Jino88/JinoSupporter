import { describe, expect, it } from "vitest";
import { blankRow, cellNumber, nextNo, rowsProblem, totals } from "./processList";

const row = (no: string, process: string, workers: number | null, stdTime: number | null) => ({ ...blankRow(no), process, workers, stdTime });

describe("process list helpers", () => {
  it("adds up people and finds the slowest process", () => {
    const t = totals([row("01", "UPPER ARRAY", 1, 12.312), row("02", "GLUING", 2, 6.156), row("03", "DRY", null, null)]);
    expect(t).toMatchObject({ processes: 3, workers: 3, time: 18.468 });
    expect(t.slowest?.no).toBe("01");
    expect(totals([]).slowest).toBeNull();
  });

  it("numbers a new row after the last numbered one, in the same width", () => {
    expect(nextNo([row("01", "", null, null), row("09", "", null, null)])).toBe("10");
    expect(nextNo([row("7", "", null, null), row("A", "", null, null)])).toBe("8");
    expect(nextNo([])).toBe("");
  });

  it("reads typed numbers: empty is none, people are whole, nothing negative", () => {
    expect(cellNumber("", true)).toBeNull();
    expect(cellNumber("3", true)).toBe(3);
    expect(cellNumber("2.5", true)).toBeUndefined();
    expect(cellNumber("6,156", false)).toBe(6.156);
    expect(cellNumber("-1", false)).toBeUndefined();
  });

  it("asks for a number or a name on every row", () => {
    expect(rowsProblem([row("01", "", null, null), row("", "GLUING", null, null)])).toBeNull();
    expect(rowsProblem([row("01", "A", null, null), row(" ", " ", 2, null)])).toBe("Row 2 needs a number or a process name.");
  });
});
