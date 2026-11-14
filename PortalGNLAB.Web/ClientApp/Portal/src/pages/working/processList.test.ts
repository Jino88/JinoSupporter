import { afterEach, describe, expect, it, vi } from "vitest";
import { BASIS_TEXT, blankRow, cellNumber, nextNo, processListApi, rowsProblem, totals } from "./processList";

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
