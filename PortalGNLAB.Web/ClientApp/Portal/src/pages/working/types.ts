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
  /** Issue: planned completion date (yyyy-MM-dd) or "". */
  due: string;
  /** Issue: development phase when it was raised (Proto … MP). */
  phase: string;
  /** Issue classification: issue type, process and cause group. */
  category: string;
  process: string;
  cause: string;
  /** Issue: actual completion date (yyyy-MM-dd) or ""; an issue with a done date is closed. */
  doneDate: string;
  finalResult: string;
  /** Issue: the Setup Plan row it is a problem of, or ""; left out on save to keep the row it has. */
  setupRowId?: string;
}

/** A picture pasted into an issue (logId "") or into one of its notes. */
export interface WorkingImage {
  id: string;
  entryId: string;
  logId: string;
  fileName: string;
  contentType: string;
  size: number;
  createdBy: string;
  createdAt: string;
}

/** Editable fields of a progress note. */
export interface LogFields {
  date: string;
  text: string;
  owner: string;
  /** Planned completion date (yyyy-MM-dd) or "". */
  due: string;
  /** Actual completion date (yyyy-MM-dd) or ""; a note with a done date is done. */
  doneDate: string;
}

/** A dated progress note under an issue (one or more per day). */
export interface EntryLog extends LogFields {
  id: string;
  entryId: string;
  images: WorkingImage[];
  revision: number;
  createdBy: string;
  createdAt: string;
  updatedBy: string;
  updatedAt: string;
}

export interface Entry extends EntryFields {
  id: string;
  /** Number within the model and kind; kept when other entries are deleted. */
  seq: number;
  /** Progress notes, oldest first. */
  logs: EntryLog[];
  /** The entry's own pictures (notes carry theirs). */
  images: WorkingImage[];
  modelName: string;
  revision: number;
  createdBy: string;
  createdAt: string;
  updatedBy: string;
  updatedAt: string;
}

/** Pictures the server stores (WorkingEndpoints images upload). */
export const IMAGE_TYPES = ["image/png", "image/jpeg", "image/gif", "image/webp", "image/bmp"];
export const MAX_IMAGE_BYTES = 10 * 1024 * 1024;
export const MAX_IMAGES_PER_UPLOAD = 20;

export const KIND_LABEL: Record<EntryKind, string> = { dev: "Dev", issue: "Issue", review: "Review" };
export const SEVERITIES = ["Critical", "Major", "Minor"];
export const ISSUE_STATUSES = ["Open", "Closed"];
export const RESULTS = ["Pass", "Conditional", "Fail"];
/** Suggestions only; any text can be typed. */
export const ISSUE_TYPES = ["Quality", "Design", "Process", "Material", "Equipment", "Customer"];
export const ISSUE_PROCESSES = ["SUB", "MAIN", "FUNCTION", "VISUAL", "OQC"];
export const ISSUE_CAUSES = ["Man", "Machine", "Material", "Method", "Measurement", "Design", "Environment"];
