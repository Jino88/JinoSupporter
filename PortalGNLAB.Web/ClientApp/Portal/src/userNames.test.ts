import { afterEach, describe, expect, it } from "vitest";
import { setUserNames, userName } from "./userNames";

describe("user names", () => {
  afterEach(() => setUserNames({}));

  it("shows the name of an ID, whatever its case", () => {
    setUserNames({ admin: "Byun Jinho", Trung: "Hoàng Trung" });
    expect(userName("admin")).toBe("Byun Jinho");
    expect(userName("trung")).toBe("Hoàng Trung");
    expect(userName(" TRUNG ")).toBe("Hoàng Trung");
  });

  it("keeps an ID that has no name (a removed user, unknown) and an empty one", () => {
    setUserNames({ admin: "Byun Jinho" });
    expect(userName("gone")).toBe("gone");
    expect(userName("unknown")).toBe("unknown");
    expect(userName("")).toBe("");
    expect(userName(null)).toBe("");
    expect(userName(undefined)).toBe("");
  });
});
