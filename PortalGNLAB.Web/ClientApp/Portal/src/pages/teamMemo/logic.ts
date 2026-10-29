import type { Entry, Folder, Memo, Summary } from "./types";
const colors = [
  "#2563eb",
  "#db2777",
  "#059669",
  "#d97706",
  "#7c3aed",
  "#0891b2",
  "#dc2626",
  "#4f46e5",
];
export function initials(name: string): string {
  name = name.trim();
  if (!name) return "?";
  const parts = name.split(/[ ._-]+/).filter(Boolean);
  return (
    parts.length >= 2 ? parts[0][0] + parts[1][0] : name.slice(0, 2)
  ).toUpperCase();
}
export function avatarColor(name: string): string {
  let h = 17;
  for (const c of name.toLowerCase().split(""))
    h = (Math.imul(h, 31) + c.charCodeAt(0)) | 0;
  return colors[(h & 0x7fffffff) % colors.length];
}
export function today(date = new Date()): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}
export function ago(value: string | null, now = Date.now()): string {
  if (!value || value.startsWith("0001-")) return "";
  const minutes = (now - new Date(value).getTime()) / 60000;
  return minutes < 1
    ? "just now"
    : minutes < 60
      ? `${Math.trunc(minutes)}m ago`
      : minutes < 1440
        ? `${Math.trunc(minutes / 60)}h ago`
        : minutes < 10080
          ? `${Math.trunc(minutes / 1440)}d ago`
          : value.slice(0, 10);
}
export const stamp = (value: string) => value.slice(5, 16).replace("T", " ");
export const size = (bytes: number) =>
  bytes >= 1024 ** 2
    ? `${(bytes / 1024 ** 2).toFixed(1)} MB`
    : `${Math.max(1, Math.trunc(bytes / 1024))} KB`;
export const closed = (status: string) =>
  status === "Done" || status === "Canceled";
export const overdue = (
  m: Pick<Summary, "status" | "dueDate">,
  day = today(),
) => !closed(m.status) && !!m.dueDate && m.dueDate < day;
export function descendants(folders: Folder[], roots: string[]): Set<string> {
  const ids = new Set(roots);
  let grew = true;
  while (grew) {
    grew = false;
    for (const f of folders)
      if (f.parentId && ids.has(f.parentId) && !ids.has(f.id)) {
        ids.add(f.id);
        grew = true;
      }
  }
  return ids;
}
export function scopedFolders(folders: Folder[], plan: boolean): Folder[] {
  const hidden = descendants(
    folders,
    folders.filter((f) => f.isHidden).map((f) => f.id),
  );
  const plans = descendants(
    folders,
    folders.filter((f) => f.isPlan).map((f) => f.id),
  );
  return folders.filter((f) => !hidden.has(f.id) && plans.has(f.id) === plan);
}
export function canDropFolder(
  folders: Folder[],
  id: string,
  parent: string | null,
): boolean {
  return parent === null || !descendants(folders, [id]).has(parent);
}
export function filterMemos(
  memos: Summary[],
  folders: Folder[],
  filter: string,
  search: string,
  user: string,
  display: (u: string) => string,
  day = today(),
): Summary[] {
  const query = search.trim().toLocaleLowerCase();
  const same = (a: string) => a.toLowerCase() === user.toLowerCase();
  return memos.filter(
    (m) =>
      (filter !== "Mine" || same(m.owner) || m.participants.some(same)) &&
      (filter !== "Active" || !closed(m.status)) &&
      (filter !== "Overdue" || overdue(m, day)) &&
      (!query ||
        [
          m.title,
          m.tags,
          folders.find((f) => f.id === m.folderId)?.path ?? "",
          display(m.owner),
        ].some((s) => s.toLocaleLowerCase().includes(query))),
  );
}
export function countsByFolder(
  memos: Summary[],
  folders: Folder[],
): Map<string, number> {
  const counts = new Map<string, number>();
  for (const m of memos) {
    let id = m.folderId;
    const seen = new Set<string>();
    while (id && !seen.has(id)) {
      seen.add(id);
      const f = folders.find((f) => f.id === id);
      if (!f) break;
      counts.set(id, (counts.get(id) ?? 0) + 1);
      id = f.parentId;
    }
  }
  return counts;
}
export function newestFirst(entries: Entry[]): Entry[] {
  const groups: {
    entries: Entry[];
    index: number;
  }[] = [];
  entries.forEach((e, index) => {
    const last = groups.at(-1);
    if (e.depth > 0 && last && last.entries[0].id !== "lead")
      last.entries.push(e);
    else groups.push({ entries: [e], index });
  });
  return groups
    .sort(
      (a, b) =>
        Number(a.entries[0].id !== "lead") -
          Number(b.entries[0].id !== "lead") ||
        b.entries[0].entryDate.localeCompare(a.entries[0].entryDate) ||
        b.index - a.index,
    )
    .flatMap((g) => g.entries);
}
const headerKeys = [
  "title",
  "kind",
  "folderId",
  "owner",
  "tags",
  "startDate",
  "dueDate",
  "description",
  "repeatWeekdays",
  "repeatUntil",
] as const;
export const headerDiffers = (a: Memo, b: Memo) =>
  headerKeys.some((key) => a[key] !== b[key]);
export function spanOf(
  m: Summary,
  day = today(),
): {
  start: string;
  end: string;
} | null {
  if (!m.startDate && !m.dueDate) return null;
  let start = m.startDate ?? m.dueDate!;
  let end = m.dueDate ?? (closed(m.status) ? start : start > day ? start : day);
  if (overdue(m, day)) end = day;
  if (end < start) [start, end] = [end, start];
  return { start, end };
}
export const weekdays = (csv: string) =>
  [
    ...new Set(
      csv
        .split(",")
        .map((x) => x.trim())
        .filter((x) => /^[0-6]$/.test(x))
        .map(Number),
    ),
  ].sort();
export function addDays(day: string, n: number): string {
  const d = new Date(day + "T12:00:00");
  d.setDate(d.getDate() + n);
  return today(d);
}
export function daysBetween(a: string, b: string): number {
  return Math.round(
    (Date.parse(b + "T00:00:00Z") - Date.parse(a + "T00:00:00Z")) / 86400000,
  );
}
export interface CalendarItem {
  day: string;
  memo: Summary;
  kind: "Repeat" | "Scheduled" | "Span" | "Due";
  time: string;
  plan: boolean;
}
export function calendarItems(
  memos: Summary[],
  from: string,
  to: string,
  includeDue: boolean,
): CalendarItem[] {
  const items: CalendarItem[] = [];
  for (const m of memos) {
    const repeat = weekdays(m.repeatWeekdays);
    const plan = repeat.length > 0;
    const push = (day: string, kind: CalendarItem["kind"], time = "") =>
      items.push({ day, memo: m, kind, time, plan });
    if (repeat.length) {
      const anchor = (m.scheduleStartAt ?? m.startDate ?? m.createdAt).slice(
        0,
        10,
      );
      const time =
        m.scheduleStartAt && !m.scheduleAllDay
          ? m.scheduleStartAt.slice(11, 16)
          : (/^\s*(\d{1,2}):(\d{2})/
              .exec(m.title)
              ?.slice(1)
              .map((x) => x.padStart(2, "0"))
              .join(":") ?? "");
      for (
        let d = anchor > from ? anchor : from;
        d <= to && (!m.repeatUntil || d <= m.repeatUntil);
        d = addDays(d, 1)
      )
        if (repeat.includes(new Date(d + "T12:00:00").getDay()))
          push(d, "Repeat", time);
    } else if (m.scheduleStartAt && m.scheduleEndAt) {
      for (
        let d =
          m.scheduleStartAt.slice(0, 10) > from
            ? m.scheduleStartAt.slice(0, 10)
            : from;
        d <= to && d <= m.scheduleEndAt.slice(0, 10);
        d = addDays(d, 1)
      )
        push(
          d,
          "Scheduled",
          m.scheduleAllDay ? "" : m.scheduleStartAt.slice(11, 16),
        );
    } else if (
      includeDue &&
      m.startDate &&
      m.dueDate &&
      m.startDate < m.dueDate
    ) {
      for (
        let d = m.startDate > from ? m.startDate : from;
        d <= to && d <= m.dueDate;
        d = addDays(d, 1)
      )
        push(d, "Span");
    } else if (includeDue && m.dueDate && m.dueDate >= from && m.dueDate <= to)
      push(m.dueDate, "Due");
  }
  return items.sort(
    (a, b) =>
      a.day.localeCompare(b.day) ||
      Number(a.kind !== "Span") - Number(b.kind !== "Span") ||
      (a.kind === "Span" && b.kind === "Span"
        ? (a.memo.startDate ?? "").localeCompare(b.memo.startDate ?? "") ||
          a.memo.id.localeCompare(b.memo.id)
        : 0) ||
      (a.time || "99").localeCompare(b.time || "99") ||
      a.memo.title.localeCompare(b.memo.title),
  );
}
const escapeHtml = (s: string) =>
  s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
const inline = (s: string) =>
  escapeHtml(s)
    .replace(/`([^`]+)`/g, "<code>$1</code>")
    .replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");
/** Same minimal markdown subset as TeamMemoUi. Encode before generating markup. */
export function markdownToHtml(markdown: string | null): string {
  const lines = (markdown ?? "").replace(/\r\n/g, "\n").split("\n");
  let html = "",
    list = "";
  let para: string[] = [];
  const flush = () => {
    if (para.length) {
      html += "<p>" + para.map(inline).join("<br/>") + "</p>";
      para = [];
    }
  };
  const close = () => {
    if (list) {
      html += `</${list}>`;
      list = "";
    }
  };
  const cells = (s: string) =>
    s
      .trim()
      .replace(/^\|+|\|+$/g, "")
      .split("|")
      .map((x) => x.trim());
  for (let i = 0; i < lines.length; i++) {
    const t = lines[i].trim();
    if (!t) {
      flush();
      close();
      continue;
    }
    const h = /^(#{1,6})\s+(.*)$/.exec(t);
    if (h) {
      flush();
      close();
      const n = Math.min(6, h[1].length + 2);
      html += `<h${n}>${inline(h[2])}</h${n}>`;
      continue;
    }
    if (
      t.startsWith("|") &&
      i + 1 < lines.length &&
      /^\|?\s*:?-{2,}/.test(lines[i + 1].trim())
    ) {
      flush();
      close();
      html +=
        "<table><thead><tr>" +
        cells(t)
          .map((c) => `<th>${inline(c)}</th>`)
          .join("") +
        "</tr></thead><tbody>";
      i += 2;
      for (; i < lines.length && lines[i].trimStart().startsWith("|"); i++)
        html +=
          "<tr>" +
          cells(lines[i])
            .map((c) => `<td>${inline(c)}</td>`)
            .join("") +
          "</tr>";
      i--;
      html += "</tbody></table>";
      continue;
    }
    const b = /^([-*+]|\d+[.)])\s+(.*)$/.exec(t);
    if (b) {
      flush();
      const kind = /^\d/.test(b[1]) ? "ol" : "ul";
      if (list !== kind) {
        close();
        html += `<${kind}>`;
        list = kind;
      }
      html += `<li>${inline(b[2])}</li>`;
      continue;
    }
    close();
    para.push(t);
  }
  flush();
  close();
  return html;
}
