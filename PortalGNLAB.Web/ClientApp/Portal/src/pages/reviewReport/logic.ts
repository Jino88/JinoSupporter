import type { Placement, Report, Summary } from "./types";

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
