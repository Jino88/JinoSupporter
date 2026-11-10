// Which models the Working tree shows (Endpoints/WorkingEndpoints.cs, /tree): the Company Glossary groups are read only,
// so hidden models, models added here, sub-groups made here and models moved into them are kept in PortalGNLAB.db.
// /bmes/check tells whether a name is a BMES model.
import { getJson, toolPath, toolSend } from "../../toolsApi";
import type { GroupNode } from "./logic";
import type { ModelGroup } from "./types";

export interface AddedModel {
  name: string;
  /** Group it is listed under (a Company Glossary group or a sub-group made here), or "" for the Added group. */
  groupId: string;
  createdBy: string;
  createdAt: string;
}

/** A sub-group made in Manage, under a Company Glossary group or under another sub-group. */
export interface UserGroup {
  id: string;
  name: string;
  parentId: string;
  sort: number;
}

/** A Company Glossary model moved into another group. */
export interface PlacedModel {
  name: string;
  groupId: string;
}

export interface TreeSettings {
  hidden: string[];
  added: AddedModel[];
  /** Missing from an older server. */
  groups?: UserGroup[];
  placed?: PlacedModel[];
}

export interface BmesCandidate {
  productName: string;
  productCode: string;
  /** "bmes" (BMES product master), "catalog" (BOM & Drawing model list) or "group" (Make Model Group). */
  source: string;
}

export interface BmesCheck {
  name: string;
  exact: boolean;
  candidates: BmesCandidate[];
  /** False when the BMES connection is not set up (only the local lists were searched). */
  bmesChecked: boolean;
  bmesError: string;
}

export const treeApi = {
  get: () => getJson<TreeSettings>(toolPath("working/tree")),
  setHidden: (hidden: string[]) => toolSend<TreeSettings>("working/tree/hidden", { hidden }, "PUT"),
  add: (name: string, groupId: string) => toolSend<AddedModel>("working/tree/models", { name, groupId }),
  remove: (name: string) => toolSend<void>(`working/tree/models/${encodeURIComponent(name)}`, undefined, "DELETE"),
  check: (name: string) => getJson<BmesCheck>(toolPath(`working/bmes/check?name=${encodeURIComponent(name)}`)),
  addGroup: (name: string, parentId: string) => toolSend<UserGroup>("working/tree/groups", { name, parentId }),
  renameGroup: (id: string, name: string) => toolSend<UserGroup>(`working/tree/groups/${encodeURIComponent(id)}`, { name }, "PUT"),
  removeGroup: (id: string) => toolSend<void>(`working/tree/groups/${encodeURIComponent(id)}`, undefined, "DELETE"),
  /** groupId "" puts the model back where it came from. */
  move: (name: string, groupId: string) => toolSend<TreeSettings>(`working/tree/models/${encodeURIComponent(name)}/group`, { groupId }, "PUT"),
};

const key = (s: string) => s.trim().toLowerCase();

/** Id of the group the models added without a group are listed under. */
export const ADDED_GROUP = "__added";

/**
 * The Glossary groups with what Manage changed: its sub-groups (after the Glossary groups beside them), the Glossary
 * models moved into another group, and the added models in their group; an added model without a (known) group goes to
 * an "Added" group at the end. A move to a group that is gone, or of a model the groups do not list, is ignored.
 */
export function withAdded(groups: ModelGroup[], added: AddedModel[], userGroups: UserGroup[] = [], placed: PlacedModel[] = []): ModelGroup[] {
  const all: ModelGroup[] = [
    ...groups.map((g) => ({ ...g, models: [...g.models] })),
    ...userGroups.map((u) => ({ id: u.id, name: u.name, parentId: u.parentId || null, models: [] as string[], sort: 100_000 + u.sort, position: 0 })),
  ];
  const byId = new Map(all.map((g) => [g.id, g]));
  for (const pm of placed) {
    const to = byId.get(pm.groupId);
    const k = key(pm.name);
    // a move of a model the groups no longer list (dropped from the Glossary) is ignored too
    if (!to || !all.some((g) => g.models.some((m) => key(m) === k))) continue;
    for (const g of all) g.models = g.models.filter((m) => key(m) !== k);
    to.models.push(pm.name);
  }
  const loose: string[] = [];
  for (const a of added) {
    const to = a.groupId ? byId.get(a.groupId) : undefined;
    if (to) to.models.push(a.name);
    else loose.push(a.name);
  }
  if (loose.length) all.push({ id: ADDED_GROUP, name: "Added", parentId: null, models: loose, sort: 1_000_000, position: 0 });
  return all;
}

/** The tree without the hidden models (case ignored); groups left with no model drop out and the counts follow. */
export function hideModels(nodes: GroupNode[], hidden: string[]): GroupNode[] {
  if (!hidden.length) return nodes;
  const off = new Set(hidden.map(key));
  const walk = (list: GroupNode[]): GroupNode[] => list.flatMap((n) => {
    const models = n.models.filter((m) => !off.has(key(m)));
    const children = walk(n.children);
    const total = models.length + children.reduce((s, c) => s + c.total, 0);
    return total ? [{ ...n, models, children, total }] : [];
  });
  return walk(nodes);
}

/** Every model of the tree, in tree order. */
export const allModels = (nodes: GroupNode[]): string[] => nodes.flatMap((n) => [...n.models, ...allModels(n.children)]);

/** Groups as indented choices for "Add to group": [id, "MODULE › MSM"]. */
export function groupChoices(nodes: GroupNode[], trail: string[] = []): [string, string][] {
  return nodes.flatMap((n) => n.id === ADDED_GROUP ? [] : [[n.id, [...trail, n.name].join(" › ")] as [string, string], ...groupChoices(n.children, [...trail, n.name])]);
}
