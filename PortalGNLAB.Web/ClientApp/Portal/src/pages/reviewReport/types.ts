export interface Picture {
  no: number;
  fileName: string;
  mediaType: string;
  data: string;
}
export interface Placement {
  no: number;
  caption: string;
  group: string;
  section: string;
}
export interface Condition {
  id: string;
  label: string;
  role: string;
  date: string;
  model: string;
  details: string;
  table: string;
}
export interface Factor {
  name: string;
  baseline: string;
  changed: string;
}
export interface ResultTable {
  caption: string;
  columns: string[];
  rows: string[][];
  source: string;
}
export interface Metric {
  id: string;
  name: string;
  unit: string;
  kind: string;
  rateOf: string;
  values: { armId: string; value: string }[];
}
export interface ReportDoc {
  title: string;
  productModel: string;
  reportDate: string;
  /** Department shown in the header block. */
  dept: string;
  /** Who made the report, shown as Marker; blank falls back to the author. */
  maker: string;
  purpose: string;
  background: string;
  content: string[];
  resultTitle: string;
  changedFactors: Factor[];
  conditions: Condition[];
  resultTables: ResultTable[];
  metrics: Metric[];
  findings: string[];
  conclusion: string;
  limitations: string[];
  nextActions: string[];
  tags: string[];
  images: Placement[];
}
export interface Report {
  id: string;
  title: string;
  createdBy: string;
  createdAt: string;
  updatedAt: string;
  language: string;
  inputText: string;
  extraInstruction: string;
  images: Picture[];
  doc: ReportDoc | null;
  revisions: { at: string; by: string; request: string }[];
  model: string;
}
export interface Summary {
  id: string;
  title: string;
  createdBy: string;
  updatedAt: string;
}
export interface ReportIndex {
  reports: Summary[];
  languages: string[];
  configured: boolean;
  /** The Codex API model and effort chosen in Setting, e.g. "gpt-5.6-luna · low". */
  model: string;
}
export interface PreviewTable {
  title: string;
  /** The conditions' "table" value; blank means the title comes from resultTitle. */
  tableKey: string;
  showDate: boolean;
  showModel: boolean;
  dateSpans: number[];
  modelSpans: number[];
  columns: {
    header: string;
    /** Worked out by the server (OK / Total NG / Total NG rate): read-only here. */
    computed: boolean;
    /** The Total NG rate column: red and one cell for both lines of an arm. */
    rateColumn: boolean;
    hasBelow: boolean;
    belowHeader: string;
    /** Metric behind the number; blank for a worked-out column (added Total NG). */
    mainId: string;
    /** Rate metric shown underneath as written; blank when it is a worked-out NG ppm. */
    rateId: string;
  }[];
  rows: {
    armId: string;
    date: string;
    /** The date as shown (d-MMM); edits go to date. */
    dateLabel: string;
    model: string;
    label: string;
    test: boolean;
    cells: { main: string; below: string; hasBelow: boolean; rate: boolean }[];
  }[];
}
/** Drag data type for a picture number dragged from the picture list into the report. */
export const pictureDragType = "application/x-review-report-picture";
export const sections = [
  "purpose",
  "conditions",
  "results",
  "conclusion",
  "appendix",
];
export function freshReport(): Report {
  return {
    id: "",
    title: "",
    createdBy: "",
    createdAt: "",
    updatedAt: "",
    language: "English",
    inputText: "",
    extraInstruction: "",
    images: [],
    doc: null,
    revisions: [],
    model: "",
  };
}
