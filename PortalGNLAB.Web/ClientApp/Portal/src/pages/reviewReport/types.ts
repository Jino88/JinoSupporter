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
}
export interface PreviewTable {
  title: string;
  showDate: boolean;
  showModel: boolean;
  dateSpans: number[];
  modelSpans: number[];
  columns: { header: string; hasBelow: boolean; belowHeader: string }[];
  rows: {
    date: string;
    model: string;
    label: string;
    test: boolean;
    cells: { main: string; below: string; hasBelow: boolean; rate: boolean }[];
  }[];
}
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
    language: "Same as input",
    inputText: "",
    extraInstruction: "",
    images: [],
    doc: null,
    revisions: [],
    model: "",
  };
}
