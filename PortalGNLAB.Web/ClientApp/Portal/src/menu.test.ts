import { appPath, buildRail, groupActive, groupMenu, ICON_PATHS, isActive, isChromeHidden, railTile } from "./menu";
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

describe("railTile", () => {
  it("maps menu ids to the Blazor rail icon and short label", () => {
    expect(railTile("ng-rate")).toEqual({ icon: "trend", label: "NG Rate" });
    expect(railTile("bmes-test5")).toEqual({ icon: "cube", label: "BOM" });
    expect(railTile("review-report")).toEqual({ icon: "write", label: "Write" });
    expect(railTile("admin-ai-usages")).toEqual({ icon: "gauge", label: "AI Usage" });
    expect(railTile("bmes-setting")).toEqual({ icon: "gear", label: "Setting" });
  });

  it("returns undefined for menus that are not on the rail", () => {
    expect(railTile("team-memo")).toBeUndefined();
    expect(railTile("nope")).toBeUndefined();
  });

  it("has an SVG path for every icon used", () => {
    for (const id of ["ng-rate", "bmes-f-cost", "report", "bmes-worker-status", "bmes-test5", "bmes-make-model-group",
      "bmes-routing-table", "bmes-reason-table", "qr-summary", "qr-bako-data", "qr-airleak-data", "qr-gauss-data",
      "bmes-daily-report", "production-input", "translate", "review-report", "bmes-setting", "admin-users", "admin-ai-usages",
      "test-excel-converter", "admin-db-query"]) {
      expect(ICON_PATHS[railTile(id)!.icon]).toBeTruthy();
    }
    expect(ICON_PATHS.factory && ICON_PATHS.qr && ICON_PATHS.signout).toBeTruthy();
  });
});

describe("buildRail", () => {
  const all = [
    "ng-rate", "bmes-f-cost", "report", "bmes-worker-status", "bmes-test5", "bmes-make-model-group",
    "bmes-routing-table", "bmes-reason-table", "bmes-daily-report", "production-input", "qr-summary", "qr-bako-data",
    "qr-airleak-data", "qr-gauss-data", "translate", "review-report", "bmes-setting", "admin-users",
    "admin-ai-usages", "test-excel-converter", "admin-db-query",
  ].map((id) => item(id, "G", "/x/" + id));
  const shape = (entries: ReturnType<typeof buildRail>) =>
    entries.map((e) => (e.kind === "sep" ? "-" : e.kind === "group" ? e.key : e.label));

  it("lays out the rail like NavMenu.razor", () => {
    const rail = buildRail(all);
    expect(shape(rail)).toEqual(["bmes", "Daily", "Input", "qr", "-", "Translate", "Write", "-", "Setting", "Users", "AI Usage", "Excel", "DB"]);
    const bmes = rail[0];
    if (bmes.kind !== "group") throw new Error("expected group");
    expect(bmes.sections.map((s) => s.title ?? null)).toEqual([null, "Setting Model"]);
    expect(bmes.sections[0].tiles.map((t) => t.label)).toEqual(["NG Rate", "F-Cost", "Report", "Worker", "BOM"]);
    expect(bmes.sections[1].tiles.map((t) => t.label)).toEqual(["Group", "Routing", "Reason"]);
    const qr = rail[3];
    if (qr.kind !== "group") throw new Error("expected group");
    expect(qr.sections[0].tiles.map((t) => t.label)).toEqual(["Summary", "Bako", "Airleak", "Gauss"]);
  });

  it("shows only permitted menus and drops empty groups and separators", () => {
    const some = all.filter((i) => ["qr-summary", "translate", "admin-db-query"].includes(i.id));
    expect(shape(buildRail(some))).toEqual(["qr", "-", "Translate", "-", "DB"]);
    expect(buildRail([])).toEqual([]);
    const onlySub = buildRail(all.filter((i) => i.id === "bmes-reason-table"));
    expect(onlySub.length).toBe(1);
    const g = onlySub[0];
    if (g.kind !== "group") throw new Error("expected group");
    expect(g.sections.length).toBe(1);
    expect(g.sections[0].title).toBe("Setting Model");
  });

  it("marks a group active when one of its pages is open", () => {
    const rail = buildRail([item("ng-rate", "BMES", "/app/ng-rate", true), item("translate", "T", "/app/tools/translate", true)]);
    expect(groupActive(rail[0], "/app/ng-rate")).toBe(true);
    expect(groupActive(rail[0], "/app/tools/translate")).toBe(false);
  });
});

describe("isChromeHidden", () => {
  it("is true only for chrome=off (any case)", () => {
    expect(isChromeHidden("?qr=Q&chrome=off")).toBe(true);
    expect(isChromeHidden("?chrome=OFF")).toBe(true);
    expect(isChromeHidden("?chrome=on")).toBe(false);
    expect(isChromeHidden("")).toBe(false);
  });
});
