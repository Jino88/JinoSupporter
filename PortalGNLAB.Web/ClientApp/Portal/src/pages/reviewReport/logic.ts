import type { Condition, Placement, Report, ReportDoc, Summary } from "./types";

const same = (a: string, b: string) =>
  a.trim().toLowerCase() === b.trim().toLowerCase();

/**
 * Changes the conditions behind result-table rows. A row the server made from a metric value
 * with no listed condition gets a condition added, so the edit has somewhere to live.
 */
export function setArms(
  doc: ReportDoc,
  armIds: string[],
  change: Partial<Condition>,
): ReportDoc {
  const conditions = doc.conditions.map((c) =>
    armIds.some((id) => same(id, c.id)) ? { ...c, ...change } : c,
  );
  for (const id of armIds)
    if (!conditions.some((c) => same(c.id, id)))
      conditions.push({
        id,
        label: id,
        role: "",
        date: "",
        model: "",
        details: "",
        table: "",
        ...change,
      });
  return { ...doc, conditions };
}

/** Renames a result table: the conditions' "table" value, or resultTitle when they have none. */
export function renameTable(
  doc: ReportDoc,
  tableKey: string,
  title: string,
): ReportDoc {
  if (!tableKey.trim()) return { ...doc, resultTitle: title };
  return {
    ...doc,
    conditions: doc.conditions.map((c) =>
      same(c.table, tableKey) ? { ...c, table: title } : c,
    ),
  };
}

/** The server reads the first metric with an id, so that is the one an edit goes to. */
function changeMetric(
  doc: ReportDoc,
  metricId: string,
  change: (m: ReportDoc["metrics"][number]) => ReportDoc["metrics"][number],
): ReportDoc {
  const index = doc.metrics.findIndex((m) => same(m.id, metricId));
  if (!metricId.trim() || index < 0) return doc;
  return {
    ...doc,
    metrics: doc.metrics.map((m, i) => (i === index ? change(m) : m)),
  };
}

export function setMetricValue(
  doc: ReportDoc,
  metricId: string,
  armId: string,
  value: string,
): ReportDoc {
  return changeMetric(doc, metricId, (m) => ({
    ...m,
    values: m.values.some((v) => same(v.armId, armId))
      ? m.values.map((v) => (same(v.armId, armId) ? { ...v, value } : v))
      : [...m.values, { armId, value }],
  }));
}

/** A header reads "Name (unit)"; a trailing "(...)" goes back into the unit. */
export function renameMetric(
  doc: ReportDoc,
  metricId: string,
  header: string,
  withUnit = true,
): ReportDoc {
  const match = withUnit ? /^(.*?)\s*\(([^()]*)\)\s*$/.exec(header) : null;
  return changeMetric(doc, metricId, (m) =>
    match && match[1].trim()
      ? { ...m, name: match[1].trim(), unit: match[2].trim() }
      : { ...m, name: header.trim(), unit: withUnit ? "" : m.unit },
  );
}

/** Keys that link parts of the report together, so find-and-replace leaves them alone. */
const linkKeys = new Set([
  "id",
  "armId",
  "rateOf",
  "kind",
  "role",
  "section",
  "no",
]);

/** Replaces every exact occurrence of a text in the report's wording; returns how many changed. */
export function replaceText(
  doc: ReportDoc,
  find: string,
  replacement: string,
): { doc: ReportDoc; count: number } {
  let count = 0;
  if (!find) return { doc, count };
  const walk = (value: unknown): unknown => {
    if (typeof value === "string") {
      const parts = value.split(find);
      count += parts.length - 1;
      return parts.join(replacement);
    }
    if (Array.isArray(value)) return value.map(walk);
    if (value && typeof value === "object")
      return Object.fromEntries(
        Object.entries(value).map(([key, v]) => [
          key,
          linkKeys.has(key) ? v : walk(v),
        ]),
      );
    return value;
  };
  const next = walk(doc) as ReportDoc;
  return { doc: count ? next : doc, count };
}

/** Puts a picture in a section: moves its placement, or adds one for a picture not in the report yet. */
export function placePicture(
  doc: ReportDoc,
  no: number,
  section: string,
): ReportDoc {
  return doc.images.some((p) => p.no === no)
    ? {
        ...doc,
        images: [
          ...doc.images.filter((p) => p.no !== no),
          { ...doc.images.find((p) => p.no === no)!, section },
        ],
      }
    : {
        ...doc,
        images: [...doc.images, { no, caption: "", group: "", section }],
      };
}

/** Adds a tag typed by a person; blank or already present (any case) leaves the doc as is. */
export function addTag(doc: ReportDoc, tag: string): ReportDoc {
  const value = tag.trim();
  if (
    !value ||
    doc.tags.some((t) => t.trim().toLowerCase() === value.toLowerCase())
  )
    return doc;
  return { ...doc, tags: [...doc.tags, value] };
}

export function removeTag(doc: ReportDoc, index: number): ReportDoc {
  return { ...doc, tags: doc.tags.filter((_, i) => i !== index) };
}

/** Replaces one line of a list; an emptied line is dropped so clearing text deletes it. */
export function editLine(
  list: string[],
  index: number,
  value: string,
): string[] {
  return value.trim()
    ? list.map((line, i) => (i === index ? value : line))
    : list.filter((_, i) => i !== index);
}

export function removePicture(report: Report, no: number): Report {
  const renumber = new Map<number, number>();
  const images = report.images
    .filter((image) => image.no !== no)
    .sort((a, b) => a.no - b.no)
    .map((image, index) => {
      renumber.set(image.no, index + 1);
      return { ...image, no: index + 1 };
    });
  const doc = report.doc
    ? {
        ...report.doc,
        images: report.doc.images
          .filter((p) => renumber.has(p.no))
          .map((p) => ({ ...p, no: renumber.get(p.no)! })),
      }
    : null;
  return { ...report, images, doc };
}

export function groupPictures(placements: Placement[]): Placement[][] {
  const groups: Placement[][] = [];
  const named = new Map<string, Placement[]>();
  for (const placement of placements) {
    const name = placement.group.trim().toLowerCase();
    const existing = name ? named.get(name) : undefined;
    if (existing) existing.push(placement);
    else {
      const group = [placement];
      groups.push(group);
      if (name) named.set(name, group);
    }
  }
  return groups;
}

export function filterReports(reports: Summary[], query: string): Summary[] {
  const needle = query.trim().toLowerCase();
  return reports.filter((r) =>
    `${r.title} ${r.createdBy} ${r.updatedAt}`.toLowerCase().includes(needle),
  );
}

export function dash(text: string): string {
  const value = text.trim();
  return !value || /^[-•]/.test(value) ? value : `- ${value}`;
}
