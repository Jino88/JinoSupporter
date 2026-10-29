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
