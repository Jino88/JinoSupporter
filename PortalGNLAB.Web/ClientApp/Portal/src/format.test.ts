import { formatPublished } from "./format";

describe("formatPublished", () => {
  it("formats as local yyyy-MM-dd HH:mm", () => {
    const local = new Date(2026, 9, 29, 13, 5);
    expect(formatPublished(local.toISOString())).toBe("2026-10-29 13:05");
  });

  it("returns unreadable text unchanged", () => {
    expect(formatPublished("soon")).toBe("soon");
  });
});
