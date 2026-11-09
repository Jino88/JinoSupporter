// Working (Endpoints/WorkingEndpoints.cs). Model groups come read-only from the CompanyGlossary DB;
// model info and the dev / issue / review entries are stored in PortalGNLAB.db.

/** A row of CompanyGlossary model_groups. */
export interface ModelGroup {
  id: string;
  name: string;
  parentId: string | null;
  models: string[];
  sort: number;
  position: number;
}

export interface GroupsResponse {
  groups: ModelGroup[];
  source: string;
  error: string | null;
}

export const PHASES = ["Proto", "EVT", "DVT", "PVT", "MP"] as const;

/** Stored verbatim as InfoJson; every field is optional so older rows keep loading. */
export interface ModelInfo {
  customer?: string;
  build?: string;
  phase?: string;
  owner?: string;
  line?: string;
  sopTarget?: string;
  note?: string;
}

export interface WorkingModel {
  name: string;
  info: ModelInfo;
  revision: number;
  updatedAt: string;
  updatedBy: string;
}

export type EntryKind = "dev" | "issue" | "review";

export interface EntryFields {
  kind: EntryKind;
  date: string;
  title: string;
  body: string;
  /** Issue: Open / Closed. */
  status: string;
  /** Issue: Critical / Major / Minor. */
  severity: string;
  /** Review: Pass / Conditional / Fail. */
  result: string;
  owner: string;
  /** Issue due date (yyyy-MM-dd) or "". */
  due: string;
}

export interface Entry extends EntryFields {
  id: string;
  modelName: string;
  revision: number;
  createdBy: string;
  createdAt: string;
  updatedBy: string;
  updatedAt: string;
}

export const KIND_LABEL: Record<EntryKind, string> = { dev: "Dev", issue: "Issue", review: "Review" };
export const SEVERITIES = ["Critical", "Major", "Minor"];
export const ISSUE_STATUSES = ["Open", "Closed"];
export const RESULTS = ["Pass", "Conditional", "Fail"];
