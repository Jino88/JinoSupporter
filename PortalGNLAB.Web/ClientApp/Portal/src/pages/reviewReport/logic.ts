import type { Placement, Report, ReportDoc, Summary } from "./types";

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
export function editLine(list: string[], index: number, value: string): string[] {
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
