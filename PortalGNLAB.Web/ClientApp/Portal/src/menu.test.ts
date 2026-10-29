import { appPath, groupMenu, isActive } from "./menu";
import type { MenuItem } from "./api";

const item = (id: string, group: string, href: string, react = false): MenuItem => ({ id, label: id, group, href, react });

describe("groupMenu", () => {
  it("keeps server order for groups and items", () => {
    const groups = groupMenu([
      item("ng-rate", "BMES", "/ng-rate"),
      item("qr-summary", "QR", "/qr/summary"),
      item("bmes-f-cost", "BMES", "/bmes/f-cost"),
    ]);
    expect(groups.map((g) => g.name)).toEqual(["BMES", "QR"]);
    expect(groups[0].items.map((i) => i.id)).toEqual(["ng-rate", "bmes-f-cost"]);
  });

  it("returns no groups for no items", () => {
    expect(groupMenu([])).toEqual([]);
  });
});

describe("appPath", () => {
  it("strips the /app base", () => {
    expect(appPath("/app")).toBe("/");
    expect(appPath("/app/")).toBe("/");
    expect(appPath("/app/tools/pc-download")).toBe("/tools/pc-download");
    expect(appPath("/app/tools/pc-download/")).toBe("/tools/pc-download");
  });

  it("leaves other paths alone", () => {
    expect(appPath("/application")).toBe("/application");
    expect(appPath("/ng-rate")).toBe("/ng-rate");
  });
});

describe("isActive", () => {
  it("matches only React pages at the current path", () => {
    const pc = item("pc-download", "Tools", "/app/tools/pc-download", true);
    expect(isActive(pc, "/app/tools/pc-download")).toBe(true);
    expect(isActive(pc, "/app")).toBe(false);
    expect(isActive(item("ng-rate", "BMES", "/ng-rate"), "/ng-rate")).toBe(false);
  });
});
