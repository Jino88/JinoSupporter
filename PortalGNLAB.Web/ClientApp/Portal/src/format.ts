/** "2026-10-29T04:05:00+00:00" → local "2026-10-29 13:05" (same as the Blazor page's yyyy-MM-dd HH:mm). */
export function formatPublished(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
