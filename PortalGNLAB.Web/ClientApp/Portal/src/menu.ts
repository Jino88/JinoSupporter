import type { MenuItem } from "./api";

export interface MenuGroup {
  name: string;
  items: MenuItem[];
}

/** Groups menu items by their group, keeping the server's order for groups and items. */
export function groupMenu(items: MenuItem[]): MenuGroup[] {
  const groups: MenuGroup[] = [];
  for (const item of items) {
    let group = groups.find((g) => g.name === item.group);
    if (!group) {
      group = { name: item.group, items: [] };
      groups.push(group);
    }
    group.items.push(item);
  }
  return groups;
}

/** Path inside the React app ("/app/tools/x" → "/tools/x"); "/" for the app root. */
export function appPath(pathname: string, base = "/app"): string {
  if (pathname === base || pathname === base + "/") return "/";
  if (pathname.startsWith(base + "/")) return pathname.slice(base.length).replace(/\/+$/, "") || "/";
  return pathname;
}

/** True when the menu item points at the page currently shown. */
export function isActive(item: MenuItem, pathname: string): boolean {
  return item.react && appPath(item.href) === appPath(pathname);
}

/** 24×24 stroke icons copied from Components/Layout/NavMenu.razor (IconPaths), drawn in currentColor. */
export const ICON_PATHS: Record<string, string> = {
  factory: "<path d='M3 21V10l5 3v-3l5 3V7l5 3h3v11z'/><path d='M7 17h2M12 17h2M17 17h2'/>",
  daily: "<rect x='3' y='4' width='18' height='17' rx='2'/><path d='M3 9h18M8 2v4M16 2v4M7 13h4v4H7z'/>",
  clipboard: "<rect x='5' y='4' width='14' height='17' rx='2'/><path d='M9 2h6v4H9zM8 11h8M8 15h8M8 19h4'/>",
  qr: "<rect x='3' y='3' width='7' height='7'/><rect x='14' y='3' width='7' height='7'/><rect x='3' y='14' width='7' height='7'/><path d='M14 14h3v3h-3zM21 14v.01M14 21h.01M17 21h4v-4'/>",
  translate: "<path d='M4 5h8M8 3v2c0 4-2 7-5 9M6 9c1 2.5 3 4 6 5'/><path d='M13 21l4-9 4 9M14.5 18h5'/>",
  gear: "<circle cx='12' cy='12' r='3'/><path d='M12 2v3M12 19v3M2 12h3M19 12h3M4.9 4.9L7 7M17 17l2.1 2.1M4.9 19.1L7 17M17 7l2.1-2.1'/>",
  users: "<circle cx='9' cy='8' r='3.5'/><path d='M2.5 20c0-3.5 3-6 6.5-6s6.5 2.5 6.5 6'/><path d='M16 4.5a3.5 3.5 0 0 1 0 7M18 14c2 .6 3.5 2.6 3.5 6'/>",
  gauge: "<path d='M3 15a9 9 0 1 1 18 0'/><path d='M12 15l4-5'/><path d='M3 19h18'/>",
  excel: "<path d='M6 3h9l4 4v14H6z'/><path d='M14 3v5h5M9.5 11.5l5 6M14.5 11.5l-5 6'/>",
  db: "<ellipse cx='12' cy='5' rx='8' ry='3'/><path d='M4 5v14c0 1.7 3.6 3 8 3s8-1.3 8-3V5M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3'/>",
  trend: "<path d='M3 3v18h18'/><path d='M7 9l4 4 3-3 5 6'/>",
  coin: "<circle cx='12' cy='12' r='9'/><path d='M15 9.5c-.5-1-1.6-1.5-3-1.5-1.7 0-3 .8-3 2s1.3 1.7 3 2 3 .8 3 2-1.3 2-3 2c-1.4 0-2.6-.6-3-1.6M12 6v2M12 16v2'/>",
  write: "<path d='M14 3H6a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8z'/><path d='M14 3v5h5'/><path d='M9 13h6M9 17h4'/>",
  report: "<path d='M5 3h14v18H5z'/><path d='M9 17v-4M12 17V9M15 17v-6'/>",
  worker: "<circle cx='12' cy='8' r='4'/><path d='M4 21c0-4 3.6-7 8-7s8 3 8 7'/>",
  cube: "<path d='M12 3l8 4.5v9L12 21l-8-4.5v-9z'/><path d='M4 7.5l8 4.5 8-4.5M12 12v9'/>",
  grid: "<rect x='3' y='3' width='7' height='7' rx='1'/><rect x='14' y='3' width='7' height='7' rx='1'/><rect x='3' y='14' width='7' height='7' rx='1'/><rect x='14' y='14' width='7' height='7' rx='1'/>",
  route: "<circle cx='6' cy='6' r='2'/><circle cx='18' cy='18' r='2'/><path d='M8 6h7a3 3 0 0 1 0 6H9a3 3 0 0 0 0 6h7'/>",
  reason: "<path d='M4 6h16M4 12h9M4 18h6'/><circle cx='17.5' cy='16' r='3'/><path d='M19.7 18.2L21.5 20'/>",
  list: "<path d='M9 6h11M9 12h11M9 18h11'/><path d='M4 6h.01M4 12h.01M4 18h.01'/>",
  box: "<path d='M4 8h16v12H4z'/><path d='M4 8l2.5-4h11L20 8M10 12h4'/>",
  drop: "<path d='M12 3s6 6.5 6 11a6 6 0 0 1-12 0c0-4.5 6-11 6-11z'/>",
  bell: "<path d='M3 19c4 0 5-13 9-13s5 13 9 13'/><path d='M3 21h18'/>",
  signout: "<path d='M15 4h4v16h-4M10 8l-4 4 4 4M6 12h10'/>",
};

export interface RailTileInfo {
  icon: string;
  label: string;
}

const RAIL_TILES: Record<string, RailTileInfo> = {
  "ng-rate": { icon: "trend", label: "NG Rate" },
  "bmes-f-cost": { icon: "coin", label: "F-Cost" },
  report: { icon: "report", label: "Report" },
  "bmes-worker-status": { icon: "worker", label: "Worker" },
  "bmes-test5": { icon: "cube", label: "BOM" },
  "bmes-make-model-group": { icon: "grid", label: "Group" },
  "bmes-routing-table": { icon: "route", label: "Routing" },
  "bmes-reason-table": { icon: "reason", label: "Reason" },
  "qr-summary": { icon: "list", label: "Summary" },
  "qr-bako-data": { icon: "box", label: "Bako" },
  "qr-airleak-data": { icon: "drop", label: "Airleak" },
  "qr-gauss-data": { icon: "bell", label: "Gauss" },
  "bmes-daily-report": { icon: "daily", label: "Daily" },
  "production-input": { icon: "clipboard", label: "Input" },
  translate: { icon: "translate", label: "Translate" },
  "review-report": { icon: "write", label: "Write" },
  "bmes-setting": { icon: "gear", label: "Setting" },
  "admin-users": { icon: "users", label: "Users" },
  "admin-ai-usages": { icon: "gauge", label: "AI Usage" },
  "test-excel-converter": { icon: "excel", label: "Excel" },
  "admin-db-query": { icon: "db", label: "DB" },
};

/** Rail icon and short label of a menu id; undefined when the menu has no rail tile. */
export function railTile(id: string): RailTileInfo | undefined {
  return RAIL_TILES[id];
}

export interface RailTile extends RailTileInfo {
  item: MenuItem;
}

export interface RailSection {
  title?: string;
  tiles: RailTile[];
}

export type RailEntry =
  | ({ kind: "tile" } & RailTile)
  | { kind: "group"; key: string; icon: string; label: string; title: string; sections: RailSection[] }
  | { kind: "sep" };

const BMES_MAIN = ["ng-rate", "bmes-f-cost", "report", "bmes-worker-status", "bmes-test5"];
const BMES_SETTING_MODEL = ["bmes-make-model-group", "bmes-routing-table", "bmes-reason-table"];
const QR_MAIN = ["qr-summary", "qr-bako-data", "qr-airleak-data", "qr-gauss-data"];
const TOOLS = ["translate", "review-report"];
const ADMIN = ["admin-users", "admin-ai-usages", "test-excel-converter", "admin-db-query"];

/** Rail layout of NavMenu.razor, limited to the menus the server sent (= the role's permissions). */
export function buildRail(items: MenuItem[]): RailEntry[] {
  const byId = new Map(items.map((i) => [i.id, i]));
  const tiles = (ids: string[]): RailTile[] =>
    ids.flatMap((id) => {
      const item = byId.get(id);
      const info = RAIL_TILES[id];
      return item && info ? [{ ...info, item }] : [];
    });
  const group = (key: string, icon: string, label: string, title: string, sections: RailSection[]): RailEntry[] => {
    const shown = sections.filter((s) => s.tiles.length > 0);
    return shown.length ? [{ kind: "group", key, icon, label, title, sections: shown }] : [];
  };
  const single = (ids: string[]): RailEntry[] => tiles(ids).map((t) => ({ kind: "tile", ...t }));
  const sep = (entries: RailEntry[]): RailEntry[] => (entries.length ? [{ kind: "sep" }, ...entries] : []);

  return [
    ...group("bmes", "factory", "BMES", "BMES", [
      { tiles: tiles(BMES_MAIN) },
      { title: "Setting Model", tiles: tiles(BMES_SETTING_MODEL) },
    ]),
    ...single(["bmes-daily-report", "production-input"]),
    ...group("qr", "qr", "QR", "QR inspection data", [{ tiles: tiles(QR_MAIN) }]),
    ...sep(single(TOOLS)),
    ...sep([...single(["bmes-setting"]), ...single(ADMIN)]),
  ];
}

/** True when a group entry contains the page currently shown. */
export function groupActive(entry: RailEntry, pathname: string): boolean {
  return entry.kind === "group" && entry.sections.some((s) => s.tiles.some((t) => isActive(t.item, pathname)));
}
