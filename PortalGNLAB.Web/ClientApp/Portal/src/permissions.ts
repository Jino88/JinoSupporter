import type { AdminMenu } from "./api";

export const ADMIN_ROLE = "Admin";

/** Whether the role may open the menu. The Admin role always may (MenuPermissionService). */
export function isAllowed(permissions: Record<string, string[]>, role: string, menuId: string): boolean {
  return role === ADMIN_ROLE || (permissions[role] ?? []).includes(menuId);
}

/** Permissions after one checkbox change; the input is not modified. */
export function withPermission(
  permissions: Record<string, string[]>,
  role: string,
  menuId: string,
  allowed: boolean,
): Record<string, string[]> {
  const current = permissions[role] ?? [];
  const next = allowed ? (current.includes(menuId) ? current : [...current, menuId]) : current.filter((id) => id !== menuId);
  return { ...permissions, [role]: next };
}

/** Menus grouped by their group name, in the given order. */
export function menusByGroup(menus: AdminMenu[]): [string, AdminMenu[]][] {
  const groups = new Map<string, AdminMenu[]>();
  for (const menu of menus) {
    const list = groups.get(menu.group) ?? [];
    list.push(menu);
    groups.set(menu.group, list);
  }
  return [...groups.entries()];
}
