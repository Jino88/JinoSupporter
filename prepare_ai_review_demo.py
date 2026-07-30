from __future__ import annotations

import argparse
import csv
import html
import json
import re
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any


DEMO_QUERIES = [
    "bond not dry UV LED",
    "VP bending",
    "low gauss",
    "weak solder",
    "NG function high",
    "new supplier material",
]


TERM_RULES: list[tuple[re.Pattern[str], str]] = [
    (re.compile(r"\bvp\s*[+/]?\s*cd\b.*\bsepar|separ.*\bvp\s*[+/]?\s*cd\b", re.I), "VP/CD separate"),
    (re.compile(r"\b(coil\s*sp|sp\s*coil)\b.*\bsepar|separ.*\b(coil\s*sp|sp\s*coil)\b", re.I), "Coil SP separate"),
    (re.compile(r"\bbond(?:ing)?\s+not\s+dry\b|\bnot\s+dry\b", re.I), "Bond not dry"),
    (re.compile(r"\bweak\s+solder\b|\bsolder\s+weak\b|\bsoldering\s+weak\b", re.I), "Weak solder"),
    (re.compile(r"\blow\s+gauss\b|\bgauss\s+low\b|\bng\s+gauss\b", re.I), "Low gauss"),
    (re.compile(r"\bng\s+function\s+high\b|\bfunction\s+ng\b|\bng\s+rate\s+function\b|\bfunction\s+check\b", re.I), "Function NG"),
    (re.compile(r"\bvp\s+bending\b|\bbending\s+vp\b", re.I), "VP bending"),
    (re.compile(r"\bcd\s+bending\b|\bbending\s+cd\b", re.I), "CD bending"),
    (re.compile(r"\bvp\s+deform\b|\bdeform\s+vp\b", re.I), "VP deform"),
    (re.compile(r"\bcoil\s+damage\b|\bdamage\s+coil\b", re.I), "Coil damage"),
    (re.compile(r"\bframe\s+damage\b|\bdamage\s+frame\b", re.I), "Frame damage"),
    (re.compile(r"\bvp\s+damage\b|\bdamage\s+vp\b", re.I), "VP damage"),
    (re.compile(r"\bdome\s+damage\b|\bdamage\s+dome\b", re.I), "Dome damage"),
    (re.compile(r"\bdome\s+offset\b|\boffset\s+dome\b", re.I), "Dome offset"),
    (re.compile(r"\bbond(?:ing)?\s+offset\b|\boffset\s+bond", re.I), "Bonding offset"),
    (re.compile(r"\bover\s+bond\b|\bover\s+glue\b|\bglue\s+over\b", re.I), "Over bond/glue"),
    (re.compile(r"\bair\s+leak\b|\bleak\s+air\b", re.I), "Air leak"),
    (re.compile(r"\bparticle\b|\bdust\b", re.I), "Particle"),
    (re.compile(r"\bburr\b", re.I), "Burr"),
    (re.compile(r"\bgap\b", re.I), "Gap"),
    (re.compile(r"\bdimension\b|\bdim\b", re.I), "Dimension"),
    (re.compile(r"\btension\b", re.I), "Tension"),
    (re.compile(r"\bplasma\b", re.I), "Plasma"),
    (re.compile(r"\bsupplier\b|\bvender\b|\bvendor\b", re.I), "Supplier"),
    (re.compile(r"\bmaterial\b|\bfilm\b|\bplate\b|\byoke\b|\bcd\b|\bcm\b|\bsm\b", re.I), "Material"),
    (re.compile(r"\bdry\s+uv\b|\buv\s+led\b|\bled\s+uv\b", re.I), "Dry UV"),
    (re.compile(r"\bmold\b", re.I), "Mold"),
    (re.compile(r"\bjig\b", re.I), "JIG"),
    (re.compile(r"\breliability\b|\bdrop\b|\bload\b|\bshock\b|\btemperature\b|\bhumidity\b", re.I), "Reliability"),
    (re.compile(r"\bng\s+rate\b|\brate\s+ng\b", re.I), "NG rate"),
    (re.compile(r"\bbefore\b.*\bafter\b|\bafter\b.*\bbefore\b", re.I), "Before/After"),
    (re.compile(r"\bdoe\b", re.I), "DOE"),
    (re.compile(r"\bvision\b|\baoi\b", re.I), "Vision/AOI"),
    (re.compile(r"\bnoise\b", re.I), "Noise"),
    (re.compile(r"\btouch\b", re.I), "Touch"),
    (re.compile(r"\bsigma\b", re.I), "Sigma"),
    (re.compile(r"\bspl\s*[+&/ ]\s*thd\b", re.I), "SPL+THD"),
    (re.compile(r"\bthd\b", re.I), "THD"),
    (re.compile(r"\bspl\b", re.I), "SPL"),
]


def norm_key(text: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "", (text or "").upper())


def read_jsonl(path: Path, max_items: int) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if not path.exists():
        return rows
    with path.open("r", encoding="utf-8-sig", errors="replace") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                obj = json.loads(line)
            except Exception:
                continue
            if isinstance(obj, dict) and isinstance(obj.get("classification"), dict):
                rows.append(obj)
                if max_items > 0 and len(rows) >= max_items:
                    break
    return rows


def list_value(value: Any) -> list[str]:
    if isinstance(value, list):
        return [str(x).strip() for x in value if str(x).strip()]
    if isinstance(value, str) and value.strip():
        return [value.strip()]
    return []


def load_model_map(path: Path | None) -> list[tuple[str, str]]:
    if not path or not path.exists():
        return []
    pairs: list[tuple[str, str]] = []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            src = str(row.get("source_model") or row.get("source") or "").strip()
            dst = str(row.get("target_model") or row.get("target") or "").strip()
            if src and dst:
                pairs.append((src, dst))
    return sorted(pairs, key=lambda x: len(norm_key(x[0])), reverse=True)


def map_model(texts: list[str], pairs: list[tuple[str, str]]) -> tuple[str, list[str]]:
    matches: list[tuple[int, int, str, str]] = []
    for raw_text in texts:
        hay = norm_key(raw_text)
        if not hay:
            continue
        for src, dst in pairs:
            needle = norm_key(src)
            if not needle:
                continue
            pos = hay.find(needle)
            if pos >= 0:
                matches.append((pos, -len(needle), src, dst))
    if not matches:
        return "", []
    matches.sort()
    targets: list[str] = []
    sources: list[str] = []
    for _, _, src, dst in matches:
        if dst not in targets:
            targets.append(dst)
        if src not in sources:
            sources.append(src)
    return " / ".join(targets[:3]), sources[:8]


def canonical_term(raw: str) -> str:
    text = re.sub(r"\s+", " ", str(raw or "").strip())
    if not text:
        return ""
    for pattern, replacement in TERM_RULES:
        if pattern.search(text):
            return replacement
    cleaned = text.strip(" -_/.,;:")
    replacements = {
        "ng": "NG",
        "uv": "UV",
        "led": "LED",
        "vp": "VP",
        "cd": "CD",
        "sp": "SP",
        "bp": "BP",
        "sm": "SM",
        "cm": "CM",
        "dt": "DT",
        "gmi": "GMI",
        "aoi": "AOI",
        "ir": "IR",
    }
    parts = []
    for part in cleaned.split():
        key = part.lower().strip("()[]")
        parts.append(replacements.get(key, part))
    return " ".join(parts)


def canonical_list(values: list[str]) -> tuple[list[str], list[tuple[str, str]]]:
    out: list[str] = []
    mappings: list[tuple[str, str]] = []
    for value in values:
        canonical = canonical_term(value)
        if not canonical:
            continue
        mappings.append((value, canonical))
        if canonical not in out:
            out.append(canonical)
    return out, mappings


def flatten(item: dict[str, Any], model_pairs: list[tuple[str, str]]) -> tuple[dict[str, Any], list[dict[str, str]]]:
    cls = item.get("classification") or {}
    ai_model = str(cls.get("model") or "").strip()
    dataset = str(item.get("datasetName") or "")
    files = str(item.get("fileNames") or "")
    db_model = str(item.get("dbProductType") or "")
    mapped_model, model_sources = map_model([dataset, files], model_pairs)
    if not mapped_model:
        mapped_model, model_sources = map_model([ai_model, db_model], model_pairs)

    term_maps: list[dict[str, str]] = []
    fields: dict[str, list[str]] = {}
    for field in ("targetDefects", "reviewItems", "tags"):
        values, mappings = canonical_list(list_value(cls.get(field)))
        fields[field] = values
        for original, canonical in mappings:
            if original != canonical:
                term_maps.append({"field": field, "original": original, "canonical": canonical})

    row = {
        "datasetName": dataset,
        "fileNames": files,
        "dbProductType": db_model,
        "dbReportDate": str(item.get("dbReportDate") or ""),
        "aiModel": ai_model,
        "model": mapped_model or ai_model or db_model,
        "modelMappingSource": " | ".join(model_sources),
        "date": str(cls.get("date") or ""),
        "purposeCode": str(cls.get("purposeCode") or ""),
        "reviewPurpose": str(cls.get("reviewPurpose") or ""),
        "purpose": str(cls.get("purpose") or ""),
        "targetDefects": fields["targetDefects"],
        "reviewItems": fields["reviewItems"],
        "tags": fields["tags"],
        "confidence": cls.get("confidence") if isinstance(cls.get("confidence"), (int, float)) else 0,
        "needsDetailedAnalysis": bool(cls.get("needsDetailedAnalysis")),
        "evidenceSummary": str(cls.get("evidenceSummary") or ""),
        "evidenceCells": list_value(cls.get("evidenceCells")),
        "uncertainty": str(cls.get("uncertainty") or ""),
    }
    return row, term_maps


def joined(values: list[str]) -> str:
    return " | ".join(values)


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    cols = [
        "datasetName",
        "model",
        "aiModel",
        "modelMappingSource",
        "date",
        "purposeCode",
        "reviewPurpose",
        "purpose",
        "targetDefects",
        "reviewItems",
        "tags",
        "confidence",
        "needsDetailedAnalysis",
        "uncertainty",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=cols)
        writer.writeheader()
        for row in rows:
            out = dict(row)
            for key in ("targetDefects", "reviewItems", "tags"):
                out[key] = joined(out.get(key) or [])
            writer.writerow({key: out.get(key, "") for key in cols})


def write_term_mapping(path: Path, mappings: list[dict[str, str]]) -> None:
    counts = Counter((m["field"], m["original"], m["canonical"]) for m in mappings)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["field", "original", "canonical", "count"])
        writer.writeheader()
        for (field, original, canonical), count in sorted(counts.items()):
            writer.writerow({"field": field, "original": original, "canonical": canonical, "count": count})


def esc(value: Any) -> str:
    return html.escape(str(value or ""), quote=True)


def pills(values: list[str]) -> str:
    if not values:
        return '<span class="muted">-</span>'
    return "".join(f'<span class="pill">{esc(x)}</span>' for x in values)


def purpose_label(code: str) -> str:
    return {
        "1": "Validation",
        "2": "Defect Cause",
        "3": "Improvement",
        "4": "Summary",
    }.get(str(code or ""), "Unclassified")


def search_text(row: dict[str, Any]) -> str:
    parts = [
        row.get("datasetName", ""),
        row.get("model", ""),
        row.get("aiModel", ""),
        row.get("date", ""),
        row.get("reviewPurpose", ""),
        row.get("purpose", ""),
        row.get("uncertainty", ""),
        " ".join(row.get("targetDefects") or []),
        " ".join(row.get("reviewItems") or []),
        " ".join(row.get("tags") or []),
    ]
    return " ".join(str(x) for x in parts).casefold()


def search(rows: list[dict[str, Any]], query: str, top: int) -> list[dict[str, Any]]:
    terms = [x for x in re.findall(r"[\w.+/-]+", query.casefold()) if len(x) >= 2]
    scored: list[tuple[float, dict[str, Any]]] = []
    for row in rows:
        hay = search_text(row)
        score = 0.0
        if query.casefold() in hay:
            score += 10
        for term in terms:
            score += min(6, hay.count(term) * 2)
        if score > 0:
            scored.append((score + float(row.get("confidence") or 0), row))
    scored.sort(key=lambda x: x[0], reverse=True)
    return [
        {
            "score": round(score, 2),
            "datasetName": row["datasetName"],
            "model": row["model"],
            "reviewPurpose": row["reviewPurpose"],
            "targetDefects": row["targetDefects"],
            "reviewItems": row["reviewItems"],
            "tags": row["tags"],
            "confidence": row["confidence"],
        }
        for score, row in scored[:top]
    ]


def grouped(rows: list[dict[str, Any]], key: str) -> list[tuple[str, int]]:
    counts: dict[str, int] = defaultdict(int)
    for row in rows:
        value = str(row.get(key) or "").strip() or "(unknown)"
        counts[value] += 1
    return sorted(counts.items(), key=lambda x: (-x[1], x[0]))


def write_html(path: Path, rows: list[dict[str, Any]], demo: dict[str, Any]) -> None:
    model_counts = grouped(rows, "model")[:14]
    max_model = max([count for _, count in model_counts] or [1])
    options = "\n".join(f'<option value="{esc(m)}">{esc(m)} ({c})</option>' for m, c in grouped(rows, "model"))
    tr = []
    for i, row in enumerate(rows, 1):
        conf = float(row.get("confidence") or 0)
        tr.append(
            f"""
            <tr data-model="{esc(row.get('model'))}" data-search="{esc(search_text(row))}">
              <td class="no">{i}</td>
              <td><strong>{esc(row.get('datasetName'))}</strong><div class="muted">{esc(row.get('fileNames'))}</div></td>
              <td><b>{esc(row.get('model'))}</b><div class="muted">AI: {esc(row.get('aiModel'))}</div><div class="muted">match: {esc(row.get('modelMappingSource'))}</div></td>
              <td>{esc(row.get('date'))}</td>
              <td>{esc(purpose_label(row.get('purposeCode')))}</td>
              <td><strong>{esc(row.get('reviewPurpose'))}</strong><div class="purpose">{esc(row.get('purpose'))}</div></td>
              <td>{pills(row.get('targetDefects') or [])}</td>
              <td>{pills(row.get('reviewItems') or [])}</td>
              <td>{pills(row.get('tags') or [])}</td>
              <td><span class="conf">{conf:.2f}</span></td>
              <td>{esc(row.get('uncertainty'))}</td>
            </tr>
            """
        )

    demo_html = []
    for query, hits in demo.items():
        demo_html.append("<section class='demo'><h3>" + esc(query) + "</h3><ol>")
        for hit in hits[:5]:
            demo_html.append(
                f"<li><b>{esc(hit.get('model'))}</b> {esc(hit.get('reviewPurpose'))}"
                f"<em>{esc(hit.get('datasetName'))}</em></li>"
            )
        demo_html.append("</ol></section>")

    bars = "".join(
        f"<div class='barrow'><span>{esc(model)}</span><div><i style='width:{count / max_model * 100:.1f}%'></i></div><b>{count}</b></div>"
        for model, count in model_counts
    )

    path.write_text(
        f"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>AI Review Demo Index</title>
<style>
body{{margin:0;background:#f8fafc;color:#111827;font-family:Segoe UI,Arial,sans-serif}}
header{{background:#1f2937;color:#fff;padding:18px 24px}}h1{{font-size:20px;margin:0 0 6px}}header p{{margin:0;color:#d1d5db;font-size:13px}}
main{{padding:18px 24px}}.stats{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:14px}}.stat{{background:#fff;border:1px solid #d1d5db;padding:12px}}.stat b{{display:block;font-size:22px}}
.panel{{background:#fff;border:1px solid #d1d5db;padding:14px;margin-bottom:14px}}.bars{{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:8px}}
.barrow{{display:grid;grid-template-columns:150px 1fr 40px;gap:8px;align-items:center;font-size:12px}}.barrow div{{height:8px;background:#e5e7eb}}.barrow i{{display:block;height:8px;background:#4b5563}}
.demo-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:10px}}.demo{{border:1px solid #d1d5db;padding:10px}}.demo h3{{font-size:13px;margin:0 0 8px}}.demo em{{display:block;color:#6b7280;font-style:normal;font-size:11px}}
.controls{{display:flex;gap:8px;margin-bottom:10px}}input,select{{border:1px solid #d1d5db;padding:8px 10px;background:#fff}}input{{flex:1;min-width:300px}}
.wrap{{overflow:auto}}table{{width:100%;border-collapse:collapse;background:#fff;font-size:12px;min-width:1700px}}th{{position:sticky;top:0;background:#374151;color:#fff;padding:7px;text-align:left}}td{{border:1px solid #d1d5db;padding:7px;vertical-align:top}}tbody tr:nth-child(even){{background:#fafafa}}
.no{{text-align:right;color:#6b7280}}.muted{{color:#6b7280;font-size:11px;margin-top:3px}}.purpose{{color:#374151;margin-top:4px}}.pill{{display:inline-block;border:1px solid #cbd5e1;background:#f8fafc;padding:2px 5px;margin:1px;max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}.conf{{font-weight:700;background:#e5e7eb;padding:3px 6px}}
</style>
</head>
<body>
<header><h1>AI Review Demo Index</h1><p>Model names are mapped from file names; defects, review items, and tags are canonicalized for retrieval.</p></header>
<main>
<section class="stats"><div class="stat"><b>{len(rows):,}</b>Rows</div><div class="stat"><b>{len(grouped(rows,'model')):,}</b>Models</div><div class="stat"><b>{sum(1 for r in rows if float(r.get('confidence') or 0)<0.7):,}</b>Low confidence</div><div class="stat"><b>{sum(1 for r in rows if r.get('needsDetailedAnalysis')):,}</b>Need detail</div></section>
<section class="panel"><h2>Model Distribution</h2><div class="bars">{bars}</div></section>
<section class="panel"><h2>Demo Searches</h2><div class="demo-grid">{''.join(demo_html)}</div></section>
<section class="panel"><h2>Index</h2><div class="controls"><input id="q" placeholder="Search model, defect, review item, tag"><select id="m"><option value="">All models</option>{options}</select><span id="cnt" class="muted"></span></div><div class="wrap"><table><thead><tr><th>No</th><th>Dataset</th><th>Mapped Model</th><th>Date</th><th>Type</th><th>Review</th><th>Target Defects</th><th>Review Items</th><th>Tags</th><th>Conf</th><th>Uncertainty</th></tr></thead><tbody>{''.join(tr)}</tbody></table></div></section>
</main>
<script>
const q=document.getElementById('q'),m=document.getElementById('m'),cnt=document.getElementById('cnt'),rows=[...document.querySelectorAll('tbody tr')];
function f(){{const t=(q.value||'').toLowerCase(),mv=m.value;let n=0;for(const r of rows){{const show=(!t||r.dataset.search.includes(t))&&(!mv||r.dataset.model===mv);r.style.display=show?'':'none';if(show)n++;}}cnt.textContent=n+' / '+rows.length+' rows';}}
q.addEventListener('input',f);m.addEventListener('change',f);f();
</script>
</body></html>""",
        encoding="utf-8",
    )


def write_plan(path: Path) -> None:
    path.write_text(
        """# Demo Apply Plan

## Current Demo Scope
- Use `classification_results.jsonl` as the raw AI first-pass output.
- Do not modify the SQLite DB during demo preparation.
- Generate normalized demo artifacts under `sample_ready`.

## Apply Rules Added
- Normalize model names from file name/dataset name using `model_mapping_conditions.csv`.
- Keep the original AI model as `aiModel` and use the mapped model as display/search `model`.
- Canonicalize similar `Target Defects`, `Review Items`, and `Tags` into shared vocabulary.
- Save `term_canonicalization.csv` so original terms and canonical terms can be audited.

## Files To Use
- `demo_report.html`: user-facing HTML demo.
- `demo_index.csv`: normalized flat index for Excel inspection.
- `demo_index.json`: normalized structured rows.
- `demo_model_index.json`: rows grouped by mapped model.
- `demo_search_results.json`: example retrieval output.
- `term_canonicalization.csv`: original-to-canonical term mapping.

## Later Full Apply After All Rows Finish
Run this after `classification_summary.json` shows all 989 rows completed:

```powershell
python "JinoSupporter\\prepare_ai_review_demo.py" `
  --results "C:\\Users\\jhbyun\\Desktop\\새 폴더 (4)\\classification_results.jsonl" `
  --out-dir "C:\\Users\\jhbyun\\Desktop\\새 폴더 (4)\\sample_ready" `
  --model-map "C:\\Users\\jhbyun\\Desktop\\새 폴더 (4)\\sample_ready\\model_mapping_conditions.csv" `
  --max-items 0
```

## Next Implementation Target
- Use `demo_index.json` as the retrieval source for current-problem search.
- Select related reports by mapped model + canonical defect/review/tag terms.
- Send only the matched subset to detailed AI analysis.
""",
        encoding="utf-8",
    )


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser()
    p.add_argument("--results", required=True)
    p.add_argument("--out-dir", required=True)
    p.add_argument("--model-map", default="")
    p.add_argument("--max-items", type=int, default=120)
    p.add_argument("--top", type=int, default=8)
    return p.parse_args()


def main() -> int:
    args = parse_args()
    results = Path(args.results)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    model_map = Path(args.model_map) if args.model_map else out_dir / "model_mapping_conditions.csv"
    pairs = load_model_map(model_map)

    raw = read_jsonl(results, args.max_items)
    rows: list[dict[str, Any]] = []
    term_maps: list[dict[str, str]] = []
    for item in raw:
        row, maps = flatten(item, pairs)
        rows.append(row)
        term_maps.extend(maps)

    demo = {q: search(rows, q, args.top) for q in DEMO_QUERIES}
    write_csv(out_dir / "demo_index.csv", rows)
    (out_dir / "demo_index.json").write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    write_term_mapping(out_dir / "term_canonicalization.csv", term_maps)
    (out_dir / "demo_model_index.json").write_text(json.dumps({m: [r for r in rows if r["model"] == m] for m, _ in grouped(rows, "model")}, ensure_ascii=False, indent=2), encoding="utf-8")
    (out_dir / "demo_search_results.json").write_text(json.dumps(demo, ensure_ascii=False, indent=2), encoding="utf-8")
    write_html(out_dir / "demo_report.html", rows, demo)
    write_plan(out_dir / "DEMO_APPLY_PLAN.md")
    print(json.dumps({"rows": len(rows), "outDir": str(out_dir), "modelMapRows": len(pairs), "termMappings": len(term_maps)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
