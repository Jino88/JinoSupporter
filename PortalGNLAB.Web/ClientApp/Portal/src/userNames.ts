/**
 * Users are shown by their name, not their ID (Username): records keep the ID of who wrote them and
 * every screen turns it into the name with userName(). The App loads the list once at sign-in
 * (/api/portal/user-names); an ID that is not in it (a removed user, "unknown") is shown as it is.
 */
let names = new Map<string, string>();

export function setUserNames(list: Record<string, string>): void {
  names = new Map(Object.entries(list).map(([id, name]) => [id.trim().toLowerCase(), name]));
}

export function userName(id: string | null | undefined): string {
  const key = (id ?? "").trim();
  return names.get(key.toLowerCase()) || key;
}
