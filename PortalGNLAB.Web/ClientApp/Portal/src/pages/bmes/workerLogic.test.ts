import { describe, expect, it } from "vitest";
import type { WorkerRecord } from "../../bmesApi";
import { buildDepartments, checkInStyle, dateColumns, isLate, normalizeRange, parseTime, statusClass, summarize, weekday } from "./workerLogic";

const rec = (o: Partial<WorkerRecord>): WorkerRecord => ({
  date: "2026-10-30T00:00:00", empNo: "1", name: "Kim", departmentNo: "10", departmentCode: "A", departmentName: "",
  workStatus: "", dayType: "", schedStart: "08:00", schedEnd: "17:00", checkIn: "", checkOut: "", ...o,
});

describe("parseTime", () => {
  it("reads HH:mm, HH:mm:ss and date-time texts", () => {
    expect(parseTime("08:30")).toBe(510);
    expect(parseTime("08:30:30")).toBe(510.5);
    expect(parseTime("2026-10-30 08:30:00")).toBe(510);
    expect(parseTime("")).toBeNull();
    expect(parseTime("abc")).toBeNull();
  });
});

describe("isLate", () => {
  it("is late only after the scheduled start", () => {
    expect(isLate({ checkIn: "08:01", schedStart: "08:00" })).toBe(true);
    expect(isLate({ checkIn: "08:00", schedStart: "08:00" })).toBe(false);
    expect(isLate({ checkIn: "", schedStart: "08:00" })).toBe(false);
  });
});

describe("checkInStyle", () => {
  it("is plain bold without times or when more than 10 min early", () => {
    expect(checkInStyle({ checkIn: "", schedStart: "08:00" })).toEqual({ fontWeight: 700 });
    expect(checkInStyle({ checkIn: "07:49", schedStart: "08:00" })).toEqual({ fontWeight: 700 });
  });
  it("runs from green (10 min early) to orange (on time or later)", () => {
    expect(checkInStyle({ checkIn: "07:50", schedStart: "08:00" }).backgroundColor).toBe("rgb(220, 252, 231)");
    expect(checkInStyle({ checkIn: "08:00", schedStart: "08:00" }).backgroundColor).toBe("rgb(255, 237, 213)");
    expect(checkInStyle({ checkIn: "09:00", schedStart: "08:00" }).color).toBe("rgb(194, 65, 12)");
  });
});

describe("statusClass", () => {
  it("maps the status text", () => {
    expect(statusClass("In office")).toBe("ws-in");
    expect(statusClass("Checked out")).toBe("ws-out");
    expect(statusClass("Absent")).toBe("ws-absent");
    expect(statusClass("Leave")).toBe("ws-other");
  });
});

describe("weekday / dateColumns / normalizeRange", () => {
  it("names the weekday", () => expect(weekday("2026-10-30")).toBe("Fri"));
  it("sorts dates newest first", () => {
    expect(dateColumns([rec({ date: "2026-10-29T00:00:00" }), rec({}), rec({})])).toEqual(["2026-10-30", "2026-10-29"]);
  });
  it("keeps From <= To", () => {
    expect(normalizeRange("2026-10-05", "2026-10-01", "start")).toEqual({ start: "2026-10-05", end: "2026-10-05" });
    expect(normalizeRange("2026-10-05", "2026-10-01", "end")).toEqual({ start: "2026-10-01", end: "2026-10-01" });
  });
});

describe("summarize / buildDepartments", () => {
  const records = [
    rec({ empNo: "2", name: "Lee", checkIn: "08:10", checkOut: "17:00" }),
    rec({ empNo: "1", name: "Kim", checkIn: "07:55" }),
    rec({ empNo: "1", name: "Kim", date: "2026-10-29T00:00:00", checkIn: "08:00" }),
    rec({ empNo: "3", name: "Park", departmentCode: "B" }),
  ];
  it("counts people, days, check-ins and late", () => {
    expect(summarize(records, 4)).toEqual({ rows: 4, people: 3, days: 2, checkedIn: 3, checkedOut: 1, late: 1 });
  });
  it("groups by department then person, sorted", () => {
    const depts = buildDepartments(records);
    expect(depts.map((d) => [d.code, d.rows, d.people])).toEqual([["A", 3, 2], ["B", 1, 1]]);
    expect(depts[0].persons.map((p) => p.records[0].name)).toEqual(["Kim", "Lee"]);
    expect([...depts[0].persons[0].byDate.keys()]).toEqual(["2026-10-30", "2026-10-29"]);
  });
});
