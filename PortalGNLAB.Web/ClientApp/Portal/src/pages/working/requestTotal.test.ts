import { describe, expect, it } from "vitest";
import { blankRequest, type RequestItem, type WorkingRequest } from "./request";
import { filtered, isCommon, jigKey, NO_MODEL, requestTotals, searched } from "./requestTotal";

const row = (model: string, name: string, order: number, drawingNo = "", lineSpec = 0, stock = 0): RequestItem =>
  ({ no: 0, model, name, lineSpec, stock, order, drawingNo, picture: "", note: "" });
const req = (id: string, items: RequestItem[], more: Partial<WorkingRequest> = {}): WorkingRequest =>
  ({ ...blankRequest("2026-11-05"), id, createdBy: "", createdAt: "", updatedBy: "", updatedAt: "", partName: "PART", items, ...more });

describe("request totals", () => {
  it("takes a JIG as its name and its drawing number both", () => {
    expect(jigKey(" ring  array jig ", " bjgp26-22-t-0001-00")).toBe(jigKey("RING ARRAY JIG", "BJGP26-22-T-0001-00"));
    expect(jigKey("RING ARRAY JIG", "BJGP26-22-T-0001-00")).not.toBe(jigKey("RING ARRAY JIG", "BJGP26-22-T-0002-00"));
    expect(jigKey("RING ARRAY JIG", "")).not.toBe(jigKey("RING ARRAY JIG", "BJGP26-22-T-0001-00"));
    expect(jigKey("RING ARRAY JIG", "X")).not.toBe(jigKey("VP CARRIER JIG", "X"));
  });

  it("adds the order of each JIG model by model over the requests; the common ones come first", () => {
    const t = requestTotals([
      req("a", [row("MSU-A", "RING ARRAY JIG", 2, "D1"), row("MSU-A", "VP CARRIER JIG", 74), row("MSU-A", "Ring array jig", 3, "D2")]),
      req("b", [row("MSU-B", "ring  array jig", 5, "d1"), row("", "VP CARRIER JIG", 0, "", 10, 4), row("MSU-B", " ", 9)]),
      req("c", [row("MSU-A", "RING ARRAY JIG", 1, "D1")], { partName: "" }),
    ]);
    expect(t.requests).toBe(3);
    // same name, other drawing number: two JIGs; the row without a model takes the part name; the row without a name is left out
    expect(t.jigs.map((j) => [j.name, j.drawingNo, j.byModel, j.total, j.requests])).toEqual([
      ["VP CARRIER JIG", "", { "MSU-A": 74, PART: 6 }, 80, 2],
      ["RING ARRAY JIG", "D1", { "MSU-A": 3, "MSU-B": 5 }, 8, 3],
      ["Ring array jig", "D2", { "MSU-A": 3 }, 3, 1],
    ]);
    expect(t.jigs.map(isCommon)).toEqual([true, true, false]);
    expect(t.models).toEqual([
      { model: "MSU-A", total: 80, jigs: 3, requests: 2 },
      { model: "PART", total: 6, jigs: 1, requests: 1 },
      { model: "MSU-B", total: 5, jigs: 1, requests: 1 },
    ]);
  });

  it("names the model of a row with no model and no part name", () => {
    expect(requestTotals([req("a", [row("", "JIG", 1)], { partName: "" })]).models[0].model).toBe(NO_MODEL);
  });

  it("keeps the requests of the dates and method chosen, and finds JIGs by name or drawing number", () => {
    const list = [req("a", [], { requestDate: "2026-10-31" }), req("b", [], { requestDate: "2026-11-05", method: "repair" }), req("c", [], { requestDate: "" })];
    expect(filtered(list, { from: "2026-11-01" }).map((r) => r.id)).toEqual(["b"]);
    expect(filtered(list, { to: "2026-11-01" }).map((r) => r.id)).toEqual(["a"]);
    expect(filtered(list, { method: "new" }).map((r) => r.id)).toEqual(["a", "c"]);
    expect(filtered(list).length).toBe(3);
    const jigs = requestTotals([req("a", [row("M", "RING ARRAY JIG", 1, "BJ-01"), row("M", "VP JIG", 1)])]).jigs;
    expect(searched(jigs, "ring").map((j) => j.name)).toEqual(["RING ARRAY JIG"]);
    expect(searched(jigs, "bj-0").map((j) => j.name)).toEqual(["RING ARRAY JIG"]);
    expect(searched(jigs, " ").length).toBe(2);
  });
});
