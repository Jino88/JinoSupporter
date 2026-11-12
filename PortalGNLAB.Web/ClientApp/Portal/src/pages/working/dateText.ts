// Dates typed in Working as yyMMdd (261112), the way they are written by hand; kept as yyyy-MM-dd.

/** A day of that year, month and day as yyyy-MM-dd, or null when there is no such day. */
function day(y: number, m: number, d: number): string | null {
  const t = new Date(Date.UTC(y, m - 1, d));
  if (m < 1 || m > 12 || d < 1 || t.getUTCMonth() !== m - 1) return null;
  return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
}

/**
 * A date typed in a date box: yyMMdd ("261112"), yyyyMMdd, or with separators ("2026-11-12", "26/11/12", "2026.11.12").
 * A two-digit year is 20yy. Output: yyyy-MM-dd, "" when nothing is typed, null when it is not a day.
 */
export function readDayText(text: string): string | null {
  const t = text.trim();
  if (!t) return "";
  const m = /^(\d{2}|\d{4})(\d{2})(\d{2})$/.exec(t) ?? /^(\d{2}|\d{4})\s*[-/.]\s*(\d{1,2})\s*[-/.]\s*(\d{1,2})$/.exec(t);
  if (!m) return null;
  const y = Number(m[1]);
  return day(m[1].length === 2 ? 2000 + y : y, Number(m[2]), Number(m[3]));
}

/** A yyyy-MM-dd date as shown in a date box (yyMMdd); anything else as it is. */
export const dayText = (date: string) => (/^\d{4}-\d{2}-\d{2}$/.test(date) ? date.slice(2).replace(/-/g, "") : date);
