// Which models the Working tree shows (Endpoints/WorkingEndpoints.cs, /tree): the Company Glossary groups are read only,
// so hidden models and models added here are kept in PortalGNLAB.db. /bmes/check tells whether a name is a BMES model.
import { getJson, toolPath, toolSend } from "../../toolsApi";
import type { GroupNode } from "./logic";
import type { ModelGroup } from "./types";

export interface AddedModel {
  name: string;
  /** Company Glossary group it is listed under, or "" for the Added group. */
  groupId: string;
  createdBy: string;
  createdAt: string;
}

export interface TreeSettings {
  hidden: string[];
  added: AddedModel[];
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
};

/** Id of the group the models added without a group are listed under. */
export const ADDED_GROUP = "__added";

/** The Glossary groups with the added models put in theirs; the ones without a (known) group go to an "Added" group. */
export function withAdded(groups: ModelGroup[], added: AddedModel[]): ModelGroup[] {
  const ids = new Set(groups.map((g) => g.id));
  const extra = new Map<string, string[]>();
  for (const a of added) {
    const id = a.groupId && ids.has(a.groupId) ? a.groupId : ADDED_GROUP;
    extra.set(id, [...(extra.get(id) ?? []), a.name]);
  }
  const out = groups.map((g) => (extra.has(g.id) ? { ...g, models: [...g.models, ...extra.get(g.id)!] } : g));
  if (extra.has(ADDED_GROUP))
    out.push({ id: ADDED_GROUP, name: "Added", parentId: null, models: extra.get(ADDED_GROUP)!, sort: 1_000_000, position: 0 });
  return out;
}

const key = (s: string) => s.trim().toLowerCase();

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
