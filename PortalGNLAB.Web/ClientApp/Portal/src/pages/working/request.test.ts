import { describe, expect, it } from "vitest";
import {
  blankRequest, commonModel, count, dotDate, isoDate, orderOf, readRequest, requestFields, requestProblem, requestTitle, toSave, totals,
} from "./request";

const item = (name: string, lineSpec: number, stock: number, order = 0, model = "SM-A386TOP") =>
  ({ no: 0, model, name, lineSpec, stock, order, drawingNo: "", picture: "", note: "" });

describe("processing request", () => {
  it("starts as the paper form does: new, ME GROUP, New request, today", () => {
    const r = blankRequest("2026-11-05");
    expect(r).toMatchObject({ kind: "JIG/MC", method: "new", departmentName: "ME GROUP", purpose: "New request", requestDate: "2026-11-05", items: [] });
  });

  it("reads dates written with dots or slashes and writes them with dots", () => {
    expect(isoDate("2026.11.5")).toBe("2026-11-05");
    expect(isoDate("2026/11/20")).toBe("2026-11-20");
    expect(isoDate("2026-02-30")).toBe("");
    expect(isoDate("11/20")).toBe("");
    expect(dotDate("2026-11-20")).toBe("2026.11.20");
  });

  it("reads counts leniently", () => {
    expect(count("1,200")).toBe(1200);
    expect(count(-3)).toBe(0);
    expect(count("abc")).toBe(0);
  });

  it("orders what the line needs beyond the stock when no order is given", () => {
    expect(orderOf(item("A", 40, 0))).toBe(40);
    expect(orderOf(item("A", 10, 4))).toBe(6);
    expect(orderOf(item("A", 2, 5))).toBe(0);
    expect(orderOf(item("A", 10, 4, 8))).toBe(8);
    expect(totals([item("A", 40, 0), item("B", 1, 0), item("C", 10, 4)])).toEqual({ lineSpec: 51, stock: 4, order: 47 });
  });

  it("reads an AI answer leniently: defaults for fields left out, numbers from text, unknown method as new", () => {
    const r = readRequest({
      partName: "SM-A386TOP", method: "melt", purpose: "", finishDate1: "2026.11.20",
      items: [{ model: "SM-A386TOP", name: " MAIN 1 JIG-Y ", lineSpec: "40", stock: "0", order: "40" }, "junk"],
    }, "2026-11-05");
    expect(r.method).toBe("new");
    expect(r.departmentName).toBe("ME GROUP");
    expect(r.purpose).toBe("");
    expect(r.requestDate).toBe("2026-11-05");
    expect(r.finishDate1).toBe("2026-11-20");
    expect(r.items[0]).toMatchObject({ no: 1, name: "MAIN 1 JIG-Y", lineSpec: 40, stock: 0, order: 40 });
    expect(r.items[1]).toMatchObject({ no: 2, name: "", lineSpec: 0 });
  });

  it("saves only the form fields, rows with a name renumbered and every order filled", () => {
    const f = { ...blankRequest("2026-11-05"), partName: "SM-A386TOP", items: [item("", 1, 0), item("MAIN 1 JIG-Y", 40, 0), item("F-PCB SOLDER CAP JIG", 2, 1)] };
    const saved = toSave({ ...f, id: "x", updatedBy: "me" } as typeof f);
    expect(saved).not.toHaveProperty("id");
    expect(saved).not.toHaveProperty("updatedBy");
    expect(saved.items.map((i) => [i.no, i.name, i.order])).toEqual([[1, "MAIN 1 JIG-Y", 40], [2, "F-PCB SOLDER CAP JIG", 1]]);
    expect(Object.keys(requestFields(f)).sort()).toEqual(Object.keys(blankRequest("")).sort());
  });

  it("asks for a part or equipment and at least one named row before saving", () => {
    const f = blankRequest("2026-11-05");
    expect(requestProblem(f)).toMatch(/part name/);
    expect(requestProblem({ ...f, partName: "SM-A386TOP" })).toMatch(/at least one item/);
    expect(requestProblem({ ...f, partName: "SM-A386TOP", items: [item("MAIN 1 JIG-Y", 40, 0)] })).toBe("");
  });

  it("takes the model most rows are for as the part name", () => {
    expect(commonModel([item("A", 1, 0, 0, "X"), item("B", 1, 0, 0, "SM-A386TOP"), item("C", 1, 0, 0, "SM-A386TOP")])).toBe("SM-A386TOP");
    expect(commonModel([])).toBe("");
    expect(requestTitle({ ...blankRequest(""), deptCode: "ME-2611004", partName: "SM-A386TOP" })).toBe("ME-2611004 · SM-A386TOP");
  });
});
