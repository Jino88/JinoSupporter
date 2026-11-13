import { describe, expect, it } from "vitest";
import type { SetupRow } from "./setup";
import { readSetupAnswer, setupAiItems, setupAiRequest, splitSetupAi } from "./setupAiRows";

const row = (o: Partial<SetupRow>): SetupRow => ({
  id: "r1", modelName: "M", revision: 4, createdBy: "", createdAt: "", updatedBy: "", updatedAt: "", no: 1, layoutNo: "", line: "L1",
  processName: "Soldering", machinePlan: "", machineState: "", machineNote: "", setupPlan: "", setupFinish: "", setupStatus: "X", ...o,
});
const rows = [row({}), row({ id: "r2", no: 2, processName: "Assembly", setupStatus: "" })];

describe("setup AI request", () => {
  it("sends the text, today, the current rows and the pictures", () => {
    const r = setupAiRequest("hi", "2026-11-12", rows, [{ name: "a.png", dataUrl: "data:image/png;base64,AA" }]);
    expect(r.text).toBe("hi");
    expect(r.today).toBe("2026-11-12");
    expect(r.rows.map((x) => x.id)).toEqual(["r1", "r2"]);
    expect(Object.keys(r.rows[0]).sort()).toEqual(["id", "layoutNo", "line", "machineNote", "machinePlan", "machineState", "no", "processName", "setupFinish", "setupPlan", "setupStatus"]);
    expect(r.images).toHaveLength(1);
  });
});

describe("setup AI answer", () => {
  it("reads rows leniently", () => {
    const a = readSetupAnswer({ rows: [{ id: " r1 ", no: "3", processName: " Test ", machineState: "?", setupStatus: "Δ", line: 5 }, 7, null], note: " ok ", model: "m · medium" });
    expect(a.rows).toHaveLength(1);
    expect(a.rows[0]).toMatchObject({ id: "r1", no: 3, processName: "Test", machineState: "", setupStatus: "Δ", line: "5" });
    expect(a.note).toBe("ok");
    expect(readSetupAnswer("nope")).toEqual({ rows: [], note: "", model: "" });
  });

  it("splits new and changed rows, leaving out a change with nothing different", () => {
    const a = readSetupAnswer({
      rows: [
        { id: "r1", no: 1, line: "L1", processName: "Soldering", setupStatus: "O", setupFinish: "2026-11-25" },
        { id: "r2", no: 2, line: "L1", processName: "Assembly", setupStatus: "" },
        { id: "", line: "L3", processName: "Packing", machinePlan: "2026-11-20" },
        { id: "gone", line: "L4", processName: "Test" },
      ],
    });
    const items = setupAiItems(a, rows);
    expect(items.map((i) => [i.title, i.row?.id ?? "new"])).toEqual([["L1 · Soldering", "r1"], ["L3 · Packing", "new"], ["L4 · Test", "new"]]);
    expect(items[0].detail).toBe("Setup finish – → 2026-11-25 · Setup X → O");
    const split = splitSetupAi(items, new Set([items[0].key, items[2].key]));
    expect(split.added.map((f) => f.processName)).toEqual(["Test"]);
    expect(split.changed).toHaveLength(1);
    expect(split.changed[0].row.revision).toBe(4);
    expect(split.changed[0].fields.setupStatus).toBe("O");
  });
});
