import { isAllowed, menusByGroup, withPermission } from "./permissions";

describe("isAllowed", () => {
  const perms = { Viewer: ["ng-rate"] };

  it("always allows Admin", () => {
    expect(isAllowed(perms, "Admin", "admin-db-query")).toBe(true);
  });

  it("uses the role's list for other roles", () => {
    expect(isAllowed(perms, "Viewer", "ng-rate")).toBe(true);
    expect(isAllowed(perms, "Viewer", "bmes-f-cost")).toBe(false);
    expect(isAllowed(perms, "Editor", "ng-rate")).toBe(false);
  });
});

describe("withPermission", () => {
  it("adds once and removes without touching the input", () => {
    const perms = { Viewer: ["ng-rate"] };
    const added = withPermission(perms, "Viewer", "qr-summary", true);
    expect(added.Viewer).toEqual(["ng-rate", "qr-summary"]);
    expect(withPermission(added, "Viewer", "qr-summary", true).Viewer).toEqual(["ng-rate", "qr-summary"]);
    expect(withPermission(added, "Viewer", "ng-rate", false).Viewer).toEqual(["qr-summary"]);
    expect(perms.Viewer).toEqual(["ng-rate"]);
  });

  it("creates the role's list when missing", () => {
    expect(withPermission({}, "Leader", "ng-rate", true)).toEqual({ Leader: ["ng-rate"] });
  });
});

describe("menusByGroup", () => {
  it("keeps first-seen group order", () => {
    const groups = menusByGroup([
      { id: "a", label: "A", group: "BMES" },
      { id: "b", label: "B", group: "QR" },
      { id: "c", label: "C", group: "BMES" },
    ]);
    expect(groups.map(([g, items]) => [g, items.map((m) => m.id)])).toEqual([
      ["BMES", ["a", "c"]],
      ["QR", ["b"]],
    ]);
  });
});
