import { afterEach, describe, expect, it, vi } from "vitest";
import { BASIS_TEXT, addJig, blankJig, blankRow, jigTotals, lineCount, nextNo, normalizeList, processListApi, removeJig, rowsProblem, setJig } from "./processList";
import type { ProcessRow } from "./processList";

const row = (no: string, process: string): ProcessRow => ({ ...blankRow(no), process });

describe("process list helpers", () => {
  it("starts a blank row without JIGs", () => {
    expect(blankRow("01")).toMatchObject({ no: "01", jigs: [] });
    expect(blankJig()).toEqual({ name: "", drawingNo: "", ver: "", qty: "" });
  });

  it("normalizes an older response without jigs to none, and drops the retired jig strings", () => {
    const old = { rows: [{ ...blankRow("01"), id: "a", jigs: undefined, jig: "X", jigVer: "V1", jigQty: "2" } as unknown as ProcessRow], revision: 1, updatedBy: "", updatedAt: null, source: "" };
    const n = normalizeList(old).rows[0];
    expect(n.jigs).toEqual([]);
    expect(n).not.toHaveProperty("jig");
    expect(n).not.toHaveProperty("jigVer");
    const withJigs = { ...old, rows: [{ ...old.rows[0], jigs: [{ name: "A", qty: "3" }] } as unknown as ProcessRow] };
    expect(normalizeList(withJigs).rows[0].jigs).toEqual([{ name: "A", drawingNo: "", ver: "", qty: "3" }]);
  });

  it("adds, changes and removes the JIGs of a row without touching the original", () => {
    const r0 = blankRow("01");
    const r1 = addJig(addJig(r0));
    expect(r0.jigs).toHaveLength(0);
    expect(r1.jigs).toHaveLength(2);
    const r2 = setJig(r1, 1, { name: "B", drawingNo: "D-2" });
    expect(r2.jigs.map((j) => j.name)).toEqual(["", "B"]);
    expect(r2.jigs[1].drawingNo).toBe("D-2");
    expect(removeJig(r2, 0).jigs.map((j) => j.name)).toEqual(["B"]);
  });

  describe("jigTotals", () => {
    const jig = (name: string, drawingNo: string, ver: string, qty: string) => ({ name, drawingNo, ver, qty });
    const withJigs = (no: string, ...jigs: ReturnType<typeof jig>[]): ProcessRow => ({ ...blankRow(no), jigs });

    it("counts the same JIG (name + ver. + drawing no.) of several processes once, with its biggest quantity, and sums them", () => {
      const t = jigTotals([
        withJigs("01", jig("COVER JIG", "D-1", "V1", "6"), jig("BASE", "D-2", "V4", "70")),
        withJigs("02", jig("COVER JIG", "D-1", "V1", "10"), jig("cover jig ", "D-1", "V1", "8")),
      ]);
      expect(t.kinds).toBe(2);
      expect(t.items.find((i) => i.name === "COVER JIG")).toMatchObject({ qtyNum: 10, qty: "10" });
      expect(t.total).toBe(80);
    });

    it("keeps a JIG with another version or drawing no. apart, reads an unreadable quantity as 0, and skips unnamed JIGs", () => {
      const t = jigTotals([withJigs("01", jig("A", "D-1", "V1", "2"), jig("A", "D-1", "V2", "3"), jig("A", "D-9", "V1", "x"), jig("", "D-5", "", "9"))]);
      expect(t.kinds).toBe(3);
      expect(t.total).toBe(5);
      expect(t.items.find((i) => i.drawingNo === "D-9")!.qtyNum).toBe(0);
    });

    it("is empty without JIGs", () => {
      expect(jigTotals([blankRow("01")])).toMatchObject({ kinds: 0, total: 0, items: [] });
    });
  });

  it("counts the lines of a cell, at least one", () => {
    expect(lineCount("")).toBe(1);
    expect(lineCount("A\nB\nC")).toBe(3);
  });

  it("numbers a new row after the last numbered one, in the same width", () => {
    expect(nextNo([row("01", ""), row("09", "")])).toBe("10");
    expect(nextNo([row("7", ""), row("A", "")])).toBe("8");
    expect(nextNo([])).toBe("");
  });

  it("asks for a number or a name on every row", () => {
    expect(rowsProblem([row("01", ""), row("", "GLUING")])).toBeNull();
    expect(rowsProblem([row("01", "A"), row(" ", " ")])).toBe("Row 2 needs a number or a process name.");
  });
});

describe("process list AI match api", () => {
  class FakeXhr {
    static last: FakeXhr;
    static reply = { status: 200, body: "{}" };
    method = ""; url = ""; headers: Record<string, string> = {}; withCredentials = false; sent: FormData | null = null;
    status = 0; responseText = "";
    onload: () => void = () => {}; onerror: () => void = () => {};
    constructor() { FakeXhr.last = this; }
    open(m: string, u: string) { this.method = m; this.url = u; }
    setRequestHeader(k: string, v: string) { this.headers[k] = v; }
    send(f: FormData) { this.sent = f; this.status = FakeXhr.reply.status; this.responseText = FakeXhr.reply.body; this.onload(); }
  }
  afterEach(() => vi.unstubAllGlobals());
  const file = new File(["x"], "JIG list.xlsx");

  it("posts only the workbook (no revision) and returns the match", async () => {
    vi.stubGlobal("XMLHttpRequest", FakeXhr);
    const match = { rows: [blankRow("01")], note: "n", unmatched: ["JIG-9"], basis: "instructions", model: "m · high" };
    FakeXhr.reply = { status: 200, body: JSON.stringify(match) };
    await expect(processListApi.aiMatch("MSM G/1", file)).resolves.toEqual(match);
    const x = FakeXhr.last;
    expect(x.method).toBe("POST");
    expect(x.url).toContain("working/models/MSM%20G%2F1/process-list/ai-match");
    expect(x.headers).toMatchObject({ Accept: "application/json", "X-Requested-With": "PortalGNLAB" });
    expect(x.withCredentials).toBe(true);
    expect((x.sent as FormData).get("file")).toBeInstanceOf(File);
    expect((x.sent as FormData).has("revision")).toBe(false);
  });

  it("rejects with the error of the server", async () => {
    vi.stubGlobal("XMLHttpRequest", FakeXhr);
    FakeXhr.reply = { status: 502, body: JSON.stringify({ error: "AI is unavailable" }) };
    await expect(processListApi.aiMatch("M", file)).rejects.toThrow("AI is unavailable");
  });

  it("words where the process order came from", () => {
    expect(BASIS_TEXT.instructions).toBe("Process order from the Work Instructions");
    expect(BASIS_TEXT["process-list"]).toBe("Process order from the current Process List");
    expect(BASIS_TEXT.workbook).toBe("Process order from the workbook's process numbers");
  });
});
