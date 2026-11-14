import { afterEach, describe, expect, it, vi } from "vitest";
import { BASIS_TEXT, blankRow, lineCount, nextNo, normalizeList, processListApi, rowsProblem } from "./processList";
import type { ProcessRow } from "./processList";

const row = (no: string, process: string): ProcessRow => ({ ...blankRow(no), process });

describe("process list helpers", () => {
  it("starts a blank row with empty JIG version and quantity", () => {
    expect(blankRow("01")).toMatchObject({ no: "01", jigVer: "", jigQty: "" });
  });

  it("fills the JIG version and quantity of data saved before those columns", () => {
    const old = { rows: [{ ...blankRow("01"), id: "a", jigVer: undefined, jigQty: undefined } as unknown as ProcessRow], revision: 1, updatedBy: "", updatedAt: null, source: "" };
    expect(normalizeList(old).rows[0]).toMatchObject({ jigVer: "", jigQty: "" });
    expect(normalizeList({ ...old, rows: [{ ...old.rows[0], jigVer: "V1", jigQty: "6" }] }).rows[0]).toMatchObject({ jigVer: "V1", jigQty: "6" });
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
