// BOM & Drawing logic: port of the helpers in Components/Pages/BmesTest5Page.razor
// (model-name suggestions, M-P / C-S hiding, product sections, price formats, HTML download).
import type { BomResult, BomRow, CachedModel, CatalogModel, PdmRow } from "../../bmesReportApi";
import { cmpIgnoreCase, distinctIgnoreCase, fixedGrouped, groupBy, htmlEscape, n0 } from "./format";

const NOISE_PREFIXES = ["M-P", "C-S"];
const SPLIT = /[ \-_/\\.]/;
const INT_MAX = 2147483647;

export interface Suggestion { name: string; productCode: string; source: "catalog" | "cache"; hint: string }

export const displayName = (row: BomRow) => (row.materialName.trim() ? row.materialName : row.materialCode);
export const productLabel = (row: BomRow) => (row.productName.trim() ? row.productName : row.productCode);

export function isNoiseCode(materialCode: string): boolean {
  const code = materialCode.trimStart().toUpperCase();
  return NOISE_PREFIXES.some((p) => code.startsWith(p));
}

/** Letters and digits only, upper-cased (NormalizeModelSearch). */
export function normalizeModelSearch(value: string): string {
  let out = "";
  for (const ch of value ?? "") if (/[\p{L}\p{Nd}]/u.test(ch)) out += ch.toUpperCase();
  return out;
}

export function levenshtein(left: string, right: string): number {
  if (left.length === 0) return right.length;
  if (right.length === 0) return left.length;
  let previous = Array.from({ length: right.length + 1 }, (_, i) => i);
  let current = new Array<number>(right.length + 1).fill(0);
  for (let i = 1; i <= left.length; i++) {
    current[0] = i;
    for (let j = 1; j <= right.length; j++) {
      const substitution = previous[j - 1] + (left[i - 1] === right[j - 1] ? 0 : 1);
      current[j] = Math.min(previous[j] + 1, current[j - 1] + 1, substitution);
    }
    [previous, current] = [current, previous];
  }
  return previous[right.length];
}

export function modelMatchScore(search: string, candidate: string): number {
  const query = normalizeModelSearch(search);
  const value = normalizeModelSearch(candidate);
  if (query.length === 0) return INT_MAX;
  if (query === value) return 0;
  if (value.startsWith(query)) return 10 + value.length - query.length;
  const contains = value.indexOf(query);
  if (contains >= 0) return 30 + contains + value.length - query.length;
  return 100 + levenshtein(query, value);
}

const searchTokens = (search: string) => search.split(SPLIT).filter((t) => t).map(normalizeModelSearch);

export function hasUsefulModelMatch(search: string, candidate: string): boolean {
  const query = normalizeModelSearch(search);
  const value = normalizeModelSearch(candidate);
  if (query.length === 0 || value.length === 0) return false;
  if (query.includes(value) || value.includes(query)) return true;
  return searchTokens(search).some((token) => token.length >= 2 && value.includes(token));
}

export function catalogMatchScore(search: string, candidate: string): number {
  if (!hasUsefulModelMatch(search, candidate)) return INT_MAX;
  const value = normalizeModelSearch(candidate);
  const tokens = [...new Set(searchTokens(search).filter((t) => t.length >= 2))];
  let score = modelMatchScore(search, candidate);
  for (const token of tokens) {
    const matched = value.includes(token);
    const hasDigit = /\p{Nd}/u.test(token);
    score += matched ? -(hasDigit ? 250 : 70) : hasDigit ? 400 : 35;
  }
  return score;
}

/** Up to 10 names from the local catalog, then the cached BOM models (LocalModelSuggestions). */
export function localSuggestions(searchText: string, catalog: CatalogModel[], cached: CachedModel[]): Suggestion[] {
  if (!searchText.trim()) return [];
  const fromCatalog = catalog
    .map((m) => {
      const name = m.productName.trim() ? m.productName : m.productCode;
      return { m, name, score: catalogMatchScore(searchText, name + " " + m.productCode) };
    })
    .filter((x) => x.score < INT_MAX)
    .sort((a, b) => a.score - b.score || cmpIgnoreCase(a.name, b.name))
    .map((x): Suggestion => ({
      name: x.name, productCode: x.m.productCode, source: "catalog", hint: `Locally cached BMES model name: ${x.m.productCode}`,
    }));
  const fromCache = cached
    .filter((m) => hasUsefulModelMatch(searchText, m.name))
    .map((m) => ({ m, score: modelMatchScore(searchText, m.name) }))
    .sort((a, b) => a.score - b.score || cmpIgnoreCase(a.m.name, b.m.name))
    .slice(0, 10)
    .map((x): Suggestion => ({ name: x.m.name, productCode: "", source: "cache", hint: x.m.groups }));

  const seen = new Set<string>();
  const out: Suggestion[] = [];
  for (const s of [...fromCatalog, ...fromCache]) {
    if (!s.name.trim()) continue;
    const key = s.name.toUpperCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(s);
    if (out.length === 10) break;
  }
  return out;
}

/** Cached models whose name contains the filter (FilteredCachedModels). */
export function filterCachedModels(models: CachedModel[], filter: string): CachedModel[] {
  const f = filter.trim().toUpperCase();
  return f ? models.filter((m) => m.name.toUpperCase().includes(f)) : models;
}

const parentPath = (pathText: string) => {
  const cut = pathText.lastIndexOf(" > ");
  return cut > 0 ? pathText.slice(0, cut) : "";
};

/**
 * Drops M-P-… and C-S-… lines; their subtree is kept and pulled up one level per hidden ancestor,
 * with level and path text recomputed (VisibleRows).
 */
export function visibleRows(rows: BomRow[], hideNoise: boolean): BomRow[] {
  if (!hideNoise) return rows;
  const visible: BomRow[] = [];
  for (const row of rows) {
    if (isNoiseCode(row.materialCode)) continue;
    const codes = row.bomPath.split(">").filter((c) => c);
    const names = row.bomPathText.split(" > ");
    const keep = [0];
    for (let i = 1; i < codes.length - 1; i++) if (!isNoiseCode(codes[i])) keep.push(i);
    keep.push(codes.length - 1);
    visible.push({
      ...row,
      bomLevel: keep.length - 1,
      bomPathText: keep.map((i) => (i < names.length ? names[i] : codes[i])).join(" > "),
    });
  }

  // Collapsing C-S levels can show one material twice under a product; keep the shallowest.
  const deduplicated = [...groupBy(visible, (r) => (r.productCode + "\t" + r.materialCode).toUpperCase()).values()].map((g) =>
    g.reduce((best, r) => (r.bomLevel < best.bomLevel ? r : best)),
  );
  const parents = new Set(deduplicated.map((r) => parentPath(r.bomPathText)).filter((p) => p).map((p) => p.toUpperCase()));
  return deduplicated.map((r) => ({ ...r, hasChildren: parents.has(r.bomPathText.toUpperCase()) }));
}

export interface ProductRows { product: string; isShared: boolean; rows: BomRow[] }

/** Sections of a model's rows: per product with the tree intact, or common / side-specific
 *  materials when the flat list covers several products (ByProduct). */
export function byProduct(rows: BomRow[], hideNoise: boolean): ProductRows[] {
  const products = distinctIgnoreCase(rows.map(productLabel)).sort(cmpIgnoreCase);
  if (!hideNoise || products.length < 2) {
    return [...groupBy(rows, productLabel).entries()]
      .sort((a, b) => cmpIgnoreCase(a[0], b[0]))
      .map(([product, list]) => ({
        product,
        isShared: false,
        rows: [...list].sort((a, b) => cmpIgnoreCase(a.bomPathText, b.bomPathText) || cmpIgnoreCase(a.materialCode, b.materialCode)),
      }));
  }

  const byMaterial = [...groupBy(rows, (r) => r.materialCode.toUpperCase()).values()].map((g) => {
    const first = g.reduce((best, r) => (r.bomLevel < best.bomLevel ? r : best));
    return {
      row: { ...first, bomLevel: 1, hasChildren: false },
      usedBy: distinctIgnoreCase(g.map(productLabel)).sort(cmpIgnoreCase),
    };
  });

  return [...groupBy(byMaterial, (m) => m.usedBy.join(" | ")).values()]
    .map((g) => {
      const usedBy = g[0].usedBy;
      const shared = usedBy.length === products.length;
      return {
        product: shared ? `Common — all ${products.length} products` : usedBy.join(" + ") + " only",
        isShared: shared,
        rows: g.map((m) => m.row).sort((a, b) => cmpIgnoreCase(displayName(a), displayName(b)) || cmpIgnoreCase(a.materialCode, b.materialCode)),
      };
    })
    .sort((a, b) => Number(b.isShared) - Number(a.isShared) || cmpIgnoreCase(a.product, b.product));
}

/** Products the model name resolved to in MATE (MatchedProducts). */
export function matchedProducts(result: BomResult): string[] {
  return distinctIgnoreCase(result.rows.map(productLabel).filter((p) => p.trim())).sort(cmpIgnoreCase);
}

/** PDM rows keyed by PDRTX; the first of any duplicate wins. */
export function pdmIndex(rows: PdmRow[]): Map<string, PdmRow> {
  const map = new Map<string, PdmRow>();
  for (const r of rows) {
    const key = r.pdrtx.trim().toUpperCase();
    if (key && !map.has(key)) map.set(key, r);
  }
  return map;
}

export const findPdm = (index: Map<string, PdmRow>, row: BomRow) => index.get(displayName(row).trim().toUpperCase());

/** Prices keep significant digits for currencies like USD (FormatUnitPrice). */
export function formatUnitPrice(value: number | null | undefined): string {
  if (value === null || value === undefined || value === 0) return "";
  const abs = Math.abs(value);
  if (abs >= 1000) return fixedGrouped(value, 0);
  if (abs >= 1) return fixedGrouped(value, 2);
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: 6, useGrouping: false }).format(value);
}

export function formatVnd(value: number | null | undefined): string {
  return value === null || value === undefined || value === 0 ? "" : fixedGrouped(value, 0);
}

export function sanitize(value: string): string {
  let out = "";
  for (const ch of value) out += /[\p{L}\p{Nd}]/u.test(ch) ? ch : "-";
  return out.replace(/^-+|-+$/g, "");
}

export interface HtmlOptions { plant: string; maxDepth: number; workDate: string; hideNoise: boolean; generated: string }

/** Self-contained HTML of the loaded BOMs (BuildStandaloneHtml). */
export function buildStandaloneHtml(results: BomResult[], pdm: Map<string, PdmRow>, o: HtmlOptions): string {
  const label = results.map((r) => r.model).join(", ");
  const lines: string[] = [
    "<!DOCTYPE html>",
    '<html lang="en"><head><meta charset="utf-8">',
    `<title>BOM — ${htmlEscape(label)}</title>`,
    "<style>",
    "body{font-family:Segoe UI,Malgun Gothic,sans-serif;margin:24px;color:#212529;}",
    "h1{font-size:18px;margin:0 0 4px;} .meta{color:#6c757d;font-size:12px;margin-bottom:18px;}",
    "h2{font-size:14px;margin:22px 0 6px;padding-bottom:4px;border-bottom:1px solid #dee2e6;}",
    "table{border-collapse:collapse;width:auto;min-width:850px;table-layout:fixed;font-size:12.5px;}",
    "th,td{border:1px solid #dee2e6;padding:4px 8px;text-align:left;}",
    "th{background:#f8f9fa;} .lv1 td.mat{font-weight:600;}",
    "th:nth-child(1){width:52px;}th:nth-child(2){width:360px;}th:nth-child(3){width:155px;}th:nth-child(4){width:145px;}th:nth-child(5){width:180px;}th:nth-child(6){width:260px;}",
    "tr.prod td{background:#eef2f7;font-weight:600;color:#23415e;}",
    "tr.prod.shared td{background:#e3f0e5;color:#1e5631;}",
    "</style></head><body>",
    `<h1>BOM &amp; Drawing — ${htmlEscape(label)}</h1>`,
    `<div class="meta">Plant ${htmlEscape(o.plant)} · depth ${o.maxDepth} · work date ${o.workDate}${o.hideNoise ? " · M-P / C-S hidden" : ""} · generated ${o.generated}</div>`,
  ];
  for (const result of results) {
    const rows = visibleRows(result.rows, o.hideNoise);
    lines.push(`<h2>${htmlEscape(result.model)} <span style="font-weight:400;color:#6c757d;">(${n0(rows.length)} rows)</span></h2>`);
    if (result.error.trim()) {
      lines.push(`<div style="color:#b02a37;">${htmlEscape(result.error)}</div>`);
      continue;
    }
    if (rows.length === 0) {
      lines.push('<div style="color:#6c757d;">No BOM rows found.</div>');
      continue;
    }
    lines.push('<table><thead><tr><th>Lv</th><th>Material</th><th>Code</th><th>PDMNO</th><th>PDTNO_TX</th><th>Path</th></tr></thead><tbody>');
    for (const product of byProduct(rows, o.hideNoise)) {
      lines.push(`<tr class="prod${product.isShared ? " shared" : ""}"><td colspan="6">${htmlEscape(product.product)} (${n0(product.rows.length)})</td></tr>`);
      for (const row of product.rows) {
        const level = Math.max(row.bomLevel, 1);
        const indent = "&nbsp;&nbsp;&nbsp;&nbsp;".repeat(level - 1);
        const p = findPdm(pdm, row);
        lines.push(
          `<tr class="lv${level}"><td>${level}</td>` +
          `<td class="mat">${indent}${htmlEscape(displayName(row))}</td>` +
          `<td>${htmlEscape(row.materialCode)}</td>` +
          `<td>${htmlEscape(p?.pdmno ?? "")}</td>` +
          `<td>${htmlEscape(p?.pdtnoTx ?? "")}</td>` +
          `<td>${htmlEscape(row.bomPathText)}</td></tr>`,
        );
      }
    }
    lines.push("</tbody></table>");
  }
  lines.push("</body></html>");
  return lines.map((l) => l + "\r\n").join("");
}
