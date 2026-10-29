/** Convert records exports to the same delimited input consumed by GraphMakerParser. */
export function normalizeGraphData(text: string): string {
  if (!text.trimStart().startsWith("{")) return text;
  const data = JSON.parse(text) as { records?: unknown[] };
  if (!Array.isArray(data.records))
    throw Error("JSON must contain a records array.");
  const records = data.records.filter(
    (r): r is Record<string, unknown> =>
      !!r && typeof r === "object" && !Array.isArray(r),
  );
  const headers = [...new Set(records.flatMap(Object.keys))];
  const cell = (v: unknown) => String(v ?? "").replace(/[\t\r\n]/g, " ");
  return [
    headers.map(cell).join("\t"),
    ...records.map((r) => headers.map((h) => cell(r[h])).join("\t")),
  ].join("\n");
}

export interface Table {
  headers: string[];
  rows: string[][];
}

/** Match GraphMakerParser: auto delimiter, one-based header row, duplicate headings and no-header mode. */
export function parseTable(
  text: string,
  delimiter: string,
  headerRow: number,
  noHeader: boolean,
): Table {
  const lines = text
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .filter((line) => line.length > 0);
  if (!lines.length) return { headers: [], rows: [] };
  const candidates = ["\t", ",", ";", "|"];
  const separator =
    delimiter === "auto"
      ? candidates
          .map((value) => ({
            value,
            count: lines
              .slice(0, 20)
              .reduce((n, line) => n + line.split(value).length - 1, 0),
          }))
          .sort((a, b) => b.count - a.count)[0].value
      : delimiter;
  const split = (line: string) =>
    line.split(separator).map((cell) => cell.trim().replace(/^"+|"+$/g, ""));
  if (noHeader) {
    const width = Math.max(...lines.map((line) => split(line).length));
    return {
      headers: Array.from({ length: width }, (_, i) => `Column ${i + 1}`),
      rows: lines.map(split).filter((row) => row.some((cell) => cell.trim())),
    };
  }
  const index = Math.max(0, Math.min((headerRow || 1) - 1, lines.length - 1));
  const used = new Map<string, number>();
  const headers = split(lines[index]).map((raw) => {
    const name = raw.trim() || "Column";
    const count = (used.get(name.toLowerCase()) ?? 0) + 1;
    used.set(name.toLowerCase(), count);
    return count === 1 ? name : `${name}_${count}`;
  });
  return {
    headers,
    rows: lines
      .slice(index + 1)
      .map(split)
      .filter((row) => row.some((cell) => cell.trim())),
  };
}

export function columnName(index: number): string {
  let result = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26))
    result = String.fromCharCode(65 + ((n - 1) % 26)) + result;
  return result;
}

/** The original heatmap example limits IR_Code and VP_Mold to their first five distinct values. */
export function exampleData(text: string): string {
  if (!text.trimStart().startsWith("{")) return text;
  const table = parseTable(normalizeGraphData(text), "\t", 1, false);
  const x = table.headers.indexOf("IR_Code");
  const y = table.headers.indexOf("VP_Mold");
  if (x < 0 || y < 0) return normalizeGraphData(text);
  const firstFive = (index: number) =>
    [
      ...new Set(
        table.rows
          .map((row) => row[index]?.trim().toLowerCase())
          .filter(Boolean),
      ),
    ].slice(0, 5);
  const xs = firstFive(x);
  const ys = firstFive(y);
  const rows = table.rows.filter(
    (row) =>
      xs.includes(row[x]?.trim().toLowerCase()) &&
      ys.includes(row[y]?.trim().toLowerCase()),
  );
  return [table.headers, ...(rows.length ? rows : table.rows)]
    .map((row) => row.join("\t"))
    .join("\n");
}
