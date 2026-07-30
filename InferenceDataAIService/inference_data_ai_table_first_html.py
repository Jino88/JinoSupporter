"""Static, user-facing HTML reports for table-first batch analysis results."""

from __future__ import annotations

import html
import json
import math
import re
from pathlib import Path
from typing import Any, Iterable


HTML_REPORT_VERSION = "table-first-html-report-v1"


class TableFirstHtmlError(ValueError):
    """Raised when a complete table-first HTML report cannot be built."""


_STYLE = """
:root{--ink:#15212d;--muted:#647483;--line:#dce4e8;--paper:#fff;
--wash:#f3f7f7;--navy:#17324d;--teal:#167d77;--mint:#dff3ef;
--amber:#a86408;--amber-bg:#fff0d5;--red:#a73b35;--red-bg:#fde8e6;
--blue:#2e628d;--blue-bg:#e6f0f8;--shadow:0 12px 34px rgba(22,43,58,.09)}
*{box-sizing:border-box}body{margin:0;background:var(--wash);color:var(--ink);
font-family:Inter,"Noto Sans KR","Segoe UI",sans-serif;line-height:1.55}
a{color:var(--blue);text-decoration:none}a:hover{text-decoration:underline}
.topbar{background:linear-gradient(120deg,#102a43,#174c5d 70%,#187a74);color:#fff;
padding:30px max(24px,calc((100vw - 1440px)/2));box-shadow:var(--shadow)}
.topbar h1{margin:4px 0 8px;font-size:clamp(25px,3vw,40px);line-height:1.16}
.eyebrow{font-size:12px;font-weight:800;letter-spacing:.14em;text-transform:uppercase;
opacity:.75}.subtitle{max-width:960px;color:#d7ebed;margin:0}.wrap{max-width:1440px;
margin:0 auto;padding:24px}.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));
gap:12px;margin-top:-44px;position:relative}.kpi{background:var(--paper);border:1px solid #e6ecef;
border-radius:14px;padding:17px 18px;box-shadow:var(--shadow)}.kpi .label{color:var(--muted);
font-size:12px;font-weight:700}.kpi .value{font-size:26px;font-weight:850;margin-top:3px}
.panel{background:var(--paper);border:1px solid var(--line);border-radius:15px;
padding:20px;margin-top:18px;box-shadow:0 5px 18px rgba(22,43,58,.04)}
.panel h2,.panel h3{margin:0 0 12px}.toolbar{display:grid;grid-template-columns:minmax(220px,1fr) 190px 190px;
gap:10px}.toolbar input,.toolbar select{width:100%;border:1px solid #cbd7dc;border-radius:10px;
padding:11px 12px;background:#fff;color:var(--ink);font:inherit}.countline{color:var(--muted);
font-size:13px;margin:12px 0 0}.table-wrap{overflow:auto;border:1px solid var(--line);border-radius:12px}
table{width:100%;border-collapse:collapse;font-size:13px}th{position:sticky;top:0;background:#edf3f5;
text-align:left;color:#455b69;font-size:11px;letter-spacing:.04em;text-transform:uppercase}
th,td{padding:11px 12px;border-bottom:1px solid #e6ecef;vertical-align:top}tr:last-child td{border-bottom:0}
tbody tr:hover{background:#f8fbfb}.file-link{font-weight:750;color:#173f5f}.muted{color:var(--muted)}
.badge{display:inline-flex;align-items:center;border-radius:999px;padding:3px 8px;font-size:11px;
font-weight:800;white-space:nowrap;background:#edf1f3;color:#52636e;margin:1px 3px 1px 0}
.badge.good{background:var(--mint);color:#11635f}.badge.warn{background:var(--amber-bg);color:var(--amber)}
.badge.bad{background:var(--red-bg);color:var(--red)}.badge.info{background:var(--blue-bg);color:var(--blue)}
.crumb{display:inline-flex;color:#d8e9ed;margin-bottom:8px;font-weight:700}.summary{font-size:16px;
max-width:1100px}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:13px}
.study{border:1px solid var(--line);border-radius:14px;padding:18px;margin-top:14px;background:#fff}
.study-head{display:flex;align-items:flex-start;justify-content:space-between;gap:14px}.study h3{font-size:20px}
.chips{display:flex;gap:5px;flex-wrap:wrap;margin:8px 0}.chip{background:#eef4f5;border-radius:7px;
padding:3px 7px;font-size:12px;color:#405866}.group-card{border-left:4px solid var(--teal);background:#f5fbfa;
padding:10px 12px;border-radius:7px}.group-card strong{display:block}.group-card p{margin:3px 0 0;color:#50636d;font-size:12px}
.section-title{display:flex;justify-content:space-between;align-items:baseline;gap:10px;margin-top:18px}
.section-title h4{margin:0;font-size:15px}.empty{border:1px dashed #bdcbd1;color:var(--muted);
padding:18px;border-radius:10px;text-align:center;background:#fafcfc}.list{margin:6px 0;padding-left:20px}
.list li{margin:4px 0}.path{display:block;overflow-wrap:anywhere;background:#edf2f4;padding:8px 10px;
border-radius:8px;font-size:12px}.study-caption{display:flex;align-items:center;gap:8px;flex-wrap:wrap;
color:var(--muted);font-size:12px;margin:15px 0 8px}.study-matrix{min-width:max-content}
.study-matrix th:first-child{left:0;z-index:3;min-width:230px}.study-matrix td:first-child{position:sticky;
left:0;z-index:1;min-width:230px;max-width:320px;background:#fff;border-right:1px solid var(--line)}
.study-matrix tbody tr:hover td:first-child{background:#f8fbfb}.metric-head{min-width:150px;max-width:230px}
.metric-head small{display:block;color:var(--muted);font-size:10px;font-weight:650;margin-top:2px}
.group-meta,.metric-detail{display:block;color:var(--muted);font-size:11px;margin-top:4px;font-weight:400}
.comparison-note{display:block;color:var(--blue);font-size:11px;margin-top:5px}.metric-value strong{font-size:14px}
.metric-series+.metric-series{border-top:1px dashed var(--line);margin-top:6px;padding-top:6px}
details{border:1px solid var(--line);border-radius:10px;margin-top:10px;background:#fbfdfd}
summary{cursor:pointer;padding:10px 12px;font-weight:750}details>div{padding:0 12px 12px}
.footer{color:var(--muted);font-size:12px;text-align:center;padding:28px}
@media(max-width:760px){.toolbar{grid-template-columns:1fr}.wrap{padding:14px}.kpis{margin-top:-24px}
.topbar{padding:24px 18px}th,td{padding:9px 8px}.hide-mobile{display:none}}
"""


_INDEX_SCRIPT = """
(() => {
  const search = document.getElementById('search');
  const status = document.getElementById('status-filter');
  const review = document.getElementById('review-filter');
  const rows = [...document.querySelectorAll('.workbook-row')];
  const visible = document.getElementById('visible-count');
  function apply() {
    const q = search.value.trim().toLocaleLowerCase();
    let count = 0;
    rows.forEach(row => {
      const okText = !q || row.dataset.search.toLocaleLowerCase().includes(q);
      const okStatus = !status.value || row.dataset.status === status.value;
      const okReview = !review.value || row.dataset.review === review.value;
      const show = okText && okStatus && okReview;
      row.hidden = !show;
      if (show) count += 1;
    });
    visible.textContent = count.toLocaleString('ko-KR');
  }
  [search, status, review].forEach(el => el.addEventListener('input', apply));
  apply();
})();
"""


def _e(value: object) -> str:
    return html.escape(str(value if value is not None else ""), quote=True)


def _load_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise TableFirstHtmlError(f"Cannot read JSON artifact: {path}") from exc
    if not isinstance(value, dict):
        raise TableFirstHtmlError(f"Expected a JSON object: {path}")
    return value


def _artifact_path(
    batch_dir: Path,
    item: dict[str, Any],
    field: str,
    folder: str,
) -> Path:
    raw = str(item.get(field) or "").strip()
    if not raw:
        raise TableFirstHtmlError(
            f"Batch item {item.get('index')} has no {field} artifact"
        )
    supplied = Path(raw)
    candidates = [supplied]
    if not supplied.is_absolute():
        candidates.append(batch_dir / supplied)
    candidates.append(batch_dir / folder / supplied.name)
    for candidate in candidates:
        if candidate.is_file():
            return candidate.resolve()
    raise TableFirstHtmlError(
        f"Batch item {item.get('index')} {field} artifact was not found: {raw}"
    )


def _status_class(value: object) -> str:
    status = str(value or "").upper()
    if status in {"ANALYZED", "HIGH", "MATCH", "DERIVED", "LOADED"}:
        return "good"
    if status in {
        "NEEDS_REVIEW",
        "MEDIUM",
        "PARTIALLY_DERIVED",
        "CLASSIFIED_NON_NUMERIC",
    }:
        return "warn"
    if status in {"LOW", "MISMATCH", "ERROR", "SKIPPED_UNSUPPORTED"}:
        return "bad"
    return "info"


def _badge(value: object) -> str:
    text = str(value or "-")
    return f'<span class="badge {_status_class(text)}">{_e(text)}</span>'


def _number(value: object) -> str:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return "-"
    number = float(value)
    if not math.isfinite(number):
        return "-"
    if number.is_integer():
        return f"{int(number):,}"
    return f"{number:,.6g}"


def _percent_scaled(meta: dict[str, Any]) -> bool:
    if any("%" in str(value) for value in meta.get("numberFormats") or []):
        return True
    return any(
        str(sample.get("displayScale") or "").startswith("PERCENT")
        for sample in meta.get("displaySamples") or []
        if isinstance(sample, dict)
    )


def _stat(value: object, unit: str, meta: dict[str, Any]) -> str:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return "-"
    number = float(value)
    if unit == "%" and _percent_scaled(meta):
        return f"{number * 100:,.4g}%"
    suffix = f" {unit}" if unit else ""
    return f"{_number(number)}{_e(suffix)}"


def _document(title: str, body: str, *, script: str = "") -> str:
    script_tag = f"<script>{script}</script>" if script else ""
    return (
        "<!doctype html><html lang=\"ko\"><head><meta charset=\"utf-8\">"
        "<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">"
        f"<title>{_e(title)}</title><style>{_STYLE}</style></head><body>"
        f"{body}{script_tag}</body></html>\n"
    )


def _write_if_changed(path: Path, text: str) -> str:
    data = text.encode("utf-8")
    if path.is_file() and path.read_bytes() == data:
        return "REUSED"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    return "WRITTEN"


def _table_and_axis_maps(
    request: dict[str, Any],
) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, Any]]]:
    tables: dict[str, dict[str, Any]] = {}
    axes: dict[str, dict[str, Any]] = {}
    for table in request.get("tables") or []:
        table_id = str(table.get("tableId") or "")
        if not table_id:
            continue
        tables[table_id] = table
        for column in table.get("numericColumns") or []:
            axis_id = str(column.get("columnId") or "")
            if axis_id:
                axes[axis_id] = {**column, "tableId": table_id}
        for series in table.get("numericSeries") or []:
            axis_id = str(series.get("seriesId") or "")
            if axis_id:
                axes[axis_id] = {**series, "tableId": table_id}
    return tables, axes


_COORDINATE_PATTERN = re.compile(r"^([A-Z]+)([0-9]+)$", re.IGNORECASE)


def _coordinate_parts(value: object) -> tuple[str, int] | None:
    match = _COORDINATE_PATTERN.fullmatch(str(value or "").replace("$", ""))
    if not match:
        return None
    return match.group(1).upper(), int(match.group(2))


def _column_number(value: object) -> int | None:
    text = str(value or "").strip().upper()
    if not text or not text.isalpha():
        return None
    number = 0
    for character in text:
        number = number * 26 + ord(character) - ord("A") + 1
    return number


def _normalize_label(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def _label_match_score(value: object, group_label: object) -> int:
    text = _normalize_label(value)
    group = _normalize_label(group_label)
    if not text or not group:
        return 0
    if text == group:
        return 100
    if text.startswith(group):
        suffix = text[len(group) :].lstrip()
        if not suffix or suffix[0].isdigit() or suffix[0] in "#(:/-_":
            return 90
    if len(group) >= 5 and group in text:
        return 60
    return 0


def _unique_groups(study: dict[str, Any]) -> list[dict[str, Any]]:
    groups: list[dict[str, Any]] = []
    positions: dict[str, int] = {}
    for source in study.get("groups") or []:
        group = dict(source)
        key = _normalize_label(group.get("label"))
        if not key:
            continue
        if key not in positions:
            positions[key] = len(groups)
            groups.append(group)
            continue
        current = groups[positions[key]]
        if str(current.get("role") or "").upper() in {"", "UNASSESSED"}:
            current["role"] = group.get("role") or current.get("role")
        if not current.get("basis") and group.get("basis"):
            current["basis"] = group.get("basis")
    return groups


def _preview_maps(
    tables: dict[str, dict[str, Any]],
) -> dict[str, dict[int, dict[str, dict[str, Any]]]]:
    previews: dict[str, dict[int, dict[str, dict[str, Any]]]] = {}
    for table_id, table in tables.items():
        rows: dict[int, dict[str, dict[str, Any]]] = {}
        for preview_row in table.get("previewRows") or []:
            for cell in preview_row.get("cells") or []:
                parts = _coordinate_parts(cell.get("coordinate"))
                if not parts:
                    continue
                column, row = parts
                rows.setdefault(row, {})[column] = cell
        previews[table_id] = rows
    return previews


def _axis_column(meta: dict[str, Any]) -> str:
    column = str(meta.get("column") or "").strip().upper()
    if column.isalpha():
        return column
    source_range = str(meta.get("sourceRange") or "").split(":", 1)[0]
    parts = _coordinate_parts(source_range)
    return parts[0] if parts else ""


def _fact_map(study: dict[str, Any]) -> dict[str, dict[str, Any]]:
    facts: dict[str, dict[str, Any]] = {}
    for fact in [
        *(study.get("deterministicNumericFacts") or []),
        *(study.get("deterministicNumericSeries") or []),
    ]:
        axis_id = str(fact.get("columnId") or fact.get("seriesId") or "")
        if axis_id:
            facts[axis_id] = fact
    return facts


def _render_fact_summary(
    fact: dict[str, Any], unit: str, meta: dict[str, Any]
) -> str:
    if not fact:
        return '<span class="muted">-</span>'
    average = _stat(fact.get("average"), unit, meta)
    minimum = _stat(fact.get("min"), unit, meta)
    maximum = _stat(fact.get("max"), unit, meta)
    count = fact.get("numericCount")
    if count is None:
        count = fact.get("numericCellCount")
    details = []
    if minimum != "-" or maximum != "-":
        details.append(f"범위 {minimum}–{maximum}")
    if isinstance(count, (int, float)) and not isinstance(count, bool):
        details.append(f"n={int(count):,}")
    return (
        '<div class="metric-value">'
        f'<strong>{average}</strong>'
        f'<span class="metric-detail">{" · ".join(details)}</span></div>'
    )


def _cell_display(
    cell: dict[str, Any], unit: str, meta: dict[str, Any]
) -> str:
    value = cell.get("value")
    if isinstance(value, bool):
        return ""
    if isinstance(value, str):
        try:
            value = float(value.replace(",", "").strip())
        except ValueError:
            return ""
    if not isinstance(value, (int, float)) or not math.isfinite(float(value)):
        return ""
    number_format = str(cell.get("numberFormat") or "")
    if "%" in number_format:
        local_meta = {**meta, "numberFormats": [number_format]}
        return _stat(value, "%", local_meta)
    if unit == "%" and _percent_scaled(meta):
        return _stat(value, unit, meta)
    return _number(value)


def _cell_row_span(cell: dict[str, Any], fallback_row: int) -> range:
    merge_range = str(cell.get("mergeRange") or "")
    if ":" not in merge_range:
        return range(fallback_row, fallback_row + 1)
    start, end = (_coordinate_parts(value) for value in merge_range.split(":", 1))
    if not start or not end:
        return range(fallback_row, fallback_row + 1)
    return range(min(start[1], end[1]), max(start[1], end[1]) + 1)


def _group_anchors(
    group_label: str,
    table_ids: set[str],
    previews: dict[str, dict[int, dict[str, dict[str, Any]]]],
) -> list[dict[str, Any]]:
    anchors: list[dict[str, Any]] = []
    for table_id in table_ids:
        for row, cells in previews.get(table_id, {}).items():
            for column, cell in cells.items():
                score = _label_match_score(cell.get("value"), group_label)
                if score:
                    anchors.append(
                        {
                            "tableId": table_id,
                            "row": row,
                            "column": column,
                            "cell": cell,
                            "score": score,
                        }
                    )
    if not anchors:
        return []
    best = max(int(anchor["score"]) for anchor in anchors)
    return [anchor for anchor in anchors if int(anchor["score"]) == best]


def _axis_mentions_group(meta: dict[str, Any], group_label: str) -> bool:
    labels = [*(meta.get("headerTexts") or [])]
    for sample in meta.get("headerSamples") or []:
        labels.extend(sample.get("headerTexts") or [])
    return any(_label_match_score(value, group_label) >= 90 for value in labels)


def _render_direct_metric_values(
    cells: list[dict[str, Any]], unit: str, meta: dict[str, Any]
) -> str:
    displays: list[str] = []
    seen: set[str] = set()
    for cell in cells:
        coordinate = str(cell.get("coordinate") or "")
        if coordinate in seen:
            continue
        seen.add(coordinate)
        display = _cell_display(cell, unit, meta)
        if display and display not in displays:
            displays.append(display)
    if not displays:
        return '<span class="muted">-</span>'
    return f'<div class="metric-value"><strong>{" / ".join(displays)}</strong></div>'


def _render_group_metric_value(
    *,
    group_label: str | None,
    group_index: int,
    group_count: int,
    metric: dict[str, Any],
    facts: dict[str, dict[str, Any]],
    axes: dict[str, dict[str, Any]],
    previews: dict[str, dict[int, dict[str, dict[str, Any]]]],
) -> str:
    unit = str(metric.get("unit") or "")
    axis_refs = [str(value) for value in metric.get("axisRefs") or [] if value]
    candidates = [
        (position, axis_id, axes.get(axis_id, {}), facts.get(axis_id, {}))
        for position, axis_id in enumerate(axis_refs)
    ]
    if not candidates:
        return '<span class="muted">-</span>'
    if group_label is None:
        summaries = [
            _render_fact_summary(fact, unit, meta)
            for _, _, meta, fact in candidates
            if fact
        ]
        return "".join(
            f'<div class="metric-series">{summary}</div>' for summary in summaries
        ) or '<span class="muted">-</span>'

    table_ids = {
        str(meta.get("tableId") or "") for _, _, meta, _ in candidates if meta
    }
    anchors = _group_anchors(group_label, table_ids, previews)
    scored_axes: list[tuple[int, int, str, dict[str, Any], dict[str, Any]]] = []
    for position, axis_id, meta, fact in candidates:
        axis_column = _axis_column(meta)
        axis_number = _column_number(axis_column)
        if axis_number is None:
            continue
        distances = []
        for anchor in anchors:
            if anchor["tableId"] != meta.get("tableId"):
                continue
            anchor_number = _column_number(anchor["column"])
            if anchor_number is not None and axis_number > anchor_number:
                distances.append(axis_number - anchor_number)
        if distances:
            scored_axes.append((min(distances), position, axis_id, meta, fact))

    if scored_axes:
        _, _, _, meta, fact = min(scored_axes, key=lambda value: (value[0], value[1]))
        table_id = str(meta.get("tableId") or "")
        axis_column = _axis_column(meta)
        related = [anchor for anchor in anchors if anchor["tableId"] == table_id]
        rows = {int(anchor["row"]) for anchor in related}
        if len(rows) > 1 and fact:
            return _render_fact_summary(fact, unit, meta)
        direct_cells: list[dict[str, Any]] = []
        for anchor in related:
            for row in _cell_row_span(anchor["cell"], int(anchor["row"])):
                cell = previews.get(table_id, {}).get(row, {}).get(axis_column)
                if cell:
                    direct_cells.append(cell)
        if direct_cells:
            return _render_direct_metric_values(direct_cells, unit, meta)
        if fact and (len(candidates) > 1 or group_count == 1):
            return _render_fact_summary(fact, unit, meta)

    for _, _, meta, fact in candidates:
        if fact and _axis_mentions_group(meta, group_label):
            return _render_fact_summary(fact, unit, meta)
    if group_count == 1:
        _, _, meta, fact = candidates[0]
        return _render_fact_summary(fact, unit, meta)
    if len(candidates) == group_count and group_index < len(candidates):
        _, _, meta, fact = candidates[group_index]
        return _render_fact_summary(fact, unit, meta)
    return '<span class="muted">-</span>'


def _comparison_targets(study: dict[str, Any]) -> dict[str, list[str]]:
    targets: dict[str, list[str]] = {}
    for relation in study.get("comparisonRelations") or []:
        left = str(relation.get("leftGroup") or "").strip()
        right = str(relation.get("rightGroup") or "").strip()
        if not left or not right:
            continue
        targets.setdefault(_normalize_label(left), []).append(right)
        targets.setdefault(_normalize_label(right), []).append(left)
    return targets


def _render_study_matrix(
    study: dict[str, Any],
    tables: dict[str, dict[str, Any]],
    axes: dict[str, dict[str, Any]],
) -> str:
    groups = _unique_groups(study)
    metrics = list(study.get("metrics") or [])
    facts = _fact_map(study)
    previews = _preview_maps(tables)
    comparisons = _comparison_targets(study)
    display_groups: list[dict[str, Any] | None] = groups or [None]

    metric_headers = "".join(
        '<th class="metric-head">'
        f'{_e(metric.get("name") or "미지정 지표")}'
        f'<small>{_e(metric.get("unit") or "단위 미지정")}</small></th>'
        for metric in metrics
    )
    if not metrics:
        metric_headers = '<th class="metric-head">분석 지표</th>'

    rows = []
    for group_index, group in enumerate(display_groups):
        if group is None:
            group_label = None
            group_cell = (
                '<td><strong>전체 / 공통</strong>'
                '<span class="group-meta">시험군이 구분되지 않은 Study</span></td>'
            )
        else:
            group_label = str(group.get("label") or "-")
            related = list(dict.fromkeys(comparisons.get(_normalize_label(group_label), [])))
            comparison_html = (
                f'<span class="comparison-note">비교: {_e(", ".join(related))}</span>'
                if related
                else ""
            )
            basis = str(group.get("basis") or "").strip()
            basis_html = (
                f'<span class="group-meta">{_e(basis)}</span>' if basis else ""
            )
            group_cell = (
                f'<td><strong>{_e(group_label)}</strong> {_badge(group.get("role"))}'
                f'{basis_html}'
                f'{comparison_html}</td>'
            )
        values = "".join(
            "<td>"
            + _render_group_metric_value(
                group_label=group_label,
                group_index=group_index,
                group_count=len(display_groups),
                metric=metric,
                facts=facts,
                axes=axes,
                previews=previews,
            )
            + "</td>"
            for metric in metrics
        )
        if not metrics:
            values = '<td><span class="muted">분류된 측정 지표 없음</span></td>'
        rows.append(f"<tr>{group_cell}{values}</tr>")

    checks = list(study.get("deterministicAggregateChecks") or [])
    check_badges = "".join(_badge(check.get("status")) for check in checks)
    caption = (
        '<div class="study-caption">'
        f'<span>시험군 {len(groups):,}개</span><span>·</span>'
        f'<span>지표 {len(metrics):,}개</span><span>·</span>'
        f'<span>비교 {len(study.get("comparisonRelations") or []):,}건</span>'
        + (f'<span>· 코드 검산 {len(checks):,}건</span>{check_badges}' if checks else "")
        + "</div>"
    )
    return (
        f'{caption}<div class="table-wrap"><table class="study-matrix">'
        f'<thead><tr><th>시험군 / 비교</th>{metric_headers}</tr></thead>'
        f'<tbody>{"".join(rows)}</tbody></table></div>'
    )


def _render_study(
    study: dict[str, Any],
    tables: dict[str, dict[str, Any]],
    axes: dict[str, dict[str, Any]],
    index: int,
) -> str:
    titles = list(study.get("titles") or [])
    table_types = list(study.get("tableTypes") or [])
    limitations = list(study.get("limitations") or [])
    title_chips = "".join(
        f'<span class="chip">{_e(value)}</span>' for value in titles
    )
    limitation_html = (
        '<ul class="list">'
        + "".join(f"<li>{_e(value)}</li>" for value in limitations)
        + "</ul>"
        if limitations
        else '<div class="empty">기록된 제한 사항이 없습니다.</div>'
    )
    return (
        f'<article class="study" id="study-{index}"><div class="study-head"><div>'
        f'<div class="eyebrow">Study {index}</div><h3>{_e(study.get("studyGroup") or "미지정 연구")}</h3>'
        f'<div class="chips">{title_chips}'
        f'{"".join(_badge(v) for v in table_types)}</div></div>{_badge(study.get("verificationStatus"))}</div>'
        f'{_render_study_matrix(study, tables, axes)}'
        '<div class="section-title"><h4>제한 사항</h4></div>'
        f'{limitation_html}</article>'
    )


def _render_source_tables(analysis: dict[str, Any]) -> str:
    rows = []
    for table in analysis.get("tables") or []:
        table_id = str(table.get("tableId") or "")
        groups = ", ".join(
            str(group.get("label") or "") for group in table.get("groups") or []
        )
        metrics = ", ".join(
            str(metric.get("name") or "") for metric in table.get("metrics") or []
        )
        rows.append(
            "<tr>"
            f"<td><strong>{_e(table.get('title') or table_id)}</strong><br><span class=\"muted\">{_e(table.get('studyGroup') or '')}</span></td>"
            f"<td>{_badge(table.get('type'))} {_badge(table.get('confidence'))}</td>"
            f"<td>{_e(groups or '-')}</td><td>{_e(metrics or '-')}</td>"
            "</tr>"
        )
    if not rows:
        return '<div class="empty">AI 분석 대상 표가 없습니다.</div>'
    return (
        '<div class="table-wrap"><table><thead><tr><th>표 / 연구 묶음</th><th>유형 / 신뢰도</th>'
        '<th>시험군</th><th>지표</th></tr></thead>'
        f"<tbody>{''.join(rows)}</tbody></table></div>"
    )


def _render_detail(record: dict[str, Any]) -> str:
    item = record["item"]
    request = record["request"]
    analysis = record["analysis"]
    projection = record["projection"]
    tables, axes = _table_and_axis_maps(request)
    studies = list(projection.get("studies") or [])
    metric_count = sum(len(study.get("metrics") or []) for study in studies)
    comparison_count = sum(
        len(study.get("comparisonRelations") or []) for study in studies
    )
    source = projection.get("source") or request.get("source") or {}
    formula = request.get("formulaDerivation") or {}
    notes = list(analysis.get("notes") or [])
    summary = str(analysis.get("workbookSummary") or "분석 요약이 없습니다.")
    studies_html = "".join(
        _render_study(study, tables, axes, index)
        for index, study in enumerate(studies, start=1)
    ) or '<div class="panel empty">분석 가능한 연구 또는 표가 없습니다.</div>'
    note_html = (
        '<ul class="list">' + "".join(f"<li>{_e(note)}</li>" for note in notes) + "</ul>"
        if notes
        else '<span class="muted">추가 메모 없음</span>'
    )
    review_badges = " ".join(
        _badge(value) for value in item.get("reviewReasons") or []
    ) or '<span class="muted">없음</span>'
    body = (
        '<header class="topbar"><a class="crumb" href="../index.html">← 전체 workbook</a>'
        '<div class="eyebrow">Table-first workbook report</div>'
        f'<h1>{_e(source.get("fileName") or item.get("fileName") or "Workbook")}</h1>'
        f'<p class="subtitle">{_e(summary)}</p></header><main class="wrap">'
        '<section class="kpis">'
        f'<div class="kpi"><div class="label">분석 상태</div><div class="value">{_badge(projection.get("analysisStatus"))}</div></div>'
        f'<div class="kpi"><div class="label">연구 묶음</div><div class="value">{len(studies):,}</div></div>'
        f'<div class="kpi"><div class="label">분석 표</div><div class="value">{len(analysis.get("tables") or []):,}</div></div>'
        f'<div class="kpi"><div class="label">지표</div><div class="value">{metric_count:,}</div></div>'
        f'<div class="kpi"><div class="label">비교 관계</div><div class="value">{comparison_count:,}</div></div>'
        f'<div class="kpi"><div class="label">수식 오류</div><div class="value">{int(formula.get("errorCount") or 0):,}</div></div>'
        '</section><section class="panel"><h2>원본과 판정 정보</h2><div class="grid">'
        f'<div><strong>검증 상태</strong><div>{_badge(projection.get("verificationStatus"))}</div></div>'
        f'<div><strong>질의 사용 가능 여부</strong><div>{_badge(projection.get("queryEligibility"))}</div></div>'
        f'<div><strong>수식 파생</strong><div>{_badge(formula.get("status"))}</div></div>'
        f'<div><strong>검수 사유</strong><div>{review_badges}</div></div>'
        '</div><h3 style="margin-top:18px">원본 파일</h3>'
        f'<code class="path">{_e(source.get("sourcePath") or "-")}</code>'
        f'<p class="muted">Revision: {_e(source.get("revisionUid") or "-")} · SHA-256: {_e(source.get("contentSha256") or "-")}</p>'
        f'<h3>AI 메모</h3>{note_html}</section>'
        f'<section class="panel"><h2>연구·비교·지표 분석</h2>{studies_html}</section>'
        f'<section class="panel"><h2>원본 표 분류</h2>{_render_source_tables(analysis)}</section>'
        '</main><footer class="footer">코드 계산값은 Capture 원본에서 결정적으로 산출되며 AI가 재계산하지 않습니다.</footer>'
    )
    return _document(str(source.get("fileName") or "Workbook report"), body)


def _render_index(
    report: dict[str, Any],
    records: list[dict[str, Any]],
) -> str:
    workbook_count = len(records)
    study_count = sum(len(record["projection"].get("studies") or []) for record in records)
    table_count = sum(len(record["analysis"].get("tables") or []) for record in records)
    metric_count = sum(int(record["item"].get("metricCount") or 0) for record in records)
    comparison_count = sum(
        int(record["item"].get("comparisonRelationCount") or 0) for record in records
    )
    review_count = sum(bool(record["item"].get("reviewRecommended")) for record in records)
    statuses = sorted(
        {str(record["projection"].get("analysisStatus") or "UNKNOWN") for record in records}
    )
    rows = []
    for record in records:
        item = record["item"]
        projection = record["projection"]
        file_name = str(item.get("fileName") or "")
        status = str(projection.get("analysisStatus") or "UNKNOWN")
        review = "true" if item.get("reviewRecommended") else "false"
        reasons = list(item.get("reviewReasons") or [])
        reason_badges = " ".join(_badge(reason) for reason in reasons)
        if not reason_badges:
            reason_badges = '<span class="muted">-</span>'
        confidence = item.get("confidenceCounts") or {}
        search = " ".join(
            [file_name, status, *reasons, *(str(k) for k in confidence)]
        )
        rows.append(
            f'<tr class="workbook-row" data-search="{_e(search)}" data-status="{_e(status)}" data-review="{review}">'
            f'<td><a class="file-link" href="{_e(record["detailFile"])}">{_e(file_name)}</a>'
            f'<br><span class="muted">#{int(item.get("index") or 0):04d}</span></td>'
            f'<td>{_badge(status)}</td>'
            f'<td>{"".join(_badge(f"{key} {value}") for key, value in sorted(confidence.items())) or "-"}</td>'
            f'<td>{int(item.get("studyCount") or len(projection.get("studies") or [])):,}</td>'
            f'<td>{int(item.get("tableCount") or 0):,}</td>'
            f'<td>{int(item.get("metricCount") or 0):,}</td>'
            f'<td>{int(item.get("comparisonRelationCount") or 0):,}</td>'
            f'<td>{reason_badges}</td>'
            "</tr>"
        )
    status_options = "".join(
        f'<option value="{_e(status)}">{_e(status)}</option>' for status in statuses
    )
    body = (
        '<header class="topbar"><div class="eyebrow">Inference Data · Table-first</div>'
        '<h1>Workbook 분석 리포트</h1>'
        f'<p class="subtitle">{_e(report.get("completedAt") or "")} 기준. AI 의미 분류와 코드 계산 통계를 사용자 중심으로 제공합니다.</p></header>'
        '<main class="wrap"><section class="kpis">'
        f'<div class="kpi"><div class="label">Workbook</div><div class="value">{workbook_count:,}</div></div>'
        f'<div class="kpi"><div class="label">연구 묶음</div><div class="value">{study_count:,}</div></div>'
        f'<div class="kpi"><div class="label">분석 표</div><div class="value">{table_count:,}</div></div>'
        f'<div class="kpi"><div class="label">지표</div><div class="value">{metric_count:,}</div></div>'
        f'<div class="kpi"><div class="label">비교 관계</div><div class="value">{comparison_count:,}</div></div>'
        f'<div class="kpi"><div class="label">검수 권장</div><div class="value">{review_count:,}</div></div>'
        '</section><section class="panel"><h2>Workbook 찾기</h2><div class="toolbar">'
        '<input id="search" type="search" placeholder="파일명, 상태, 검수 사유 검색">'
        f'<select id="status-filter"><option value="">모든 분석 상태</option>{status_options}</select>'
        '<select id="review-filter"><option value="">모든 검수 상태</option><option value="true">검수 권장</option><option value="false">일반</option></select>'
        f'</div><p class="countline"><span id="visible-count">{workbook_count:,}</span> / {workbook_count:,} workbook 표시</p></section>'
        '<section class="panel"><div class="table-wrap"><table><thead><tr><th>Workbook</th><th>상태</th>'
        '<th>신뢰도</th><th>연구</th><th>표</th><th>지표</th><th>비교</th><th>검수 사유</th>'
        f'</tr></thead><tbody>{"".join(rows)}</tbody></table></div></section></main>'
        '<footer class="footer">각 workbook을 선택하면 연구 묶음, 시험군, 비교 관계와 지표 통계를 확인할 수 있습니다.</footer>'
    )
    return _document("Workbook 분석 리포트", body, script=_INDEX_SCRIPT)


def _safe_slug(value: object) -> str:
    slug = re.sub(r"[^A-Za-z0-9_-]+", "-", str(value or "")).strip("-")
    return slug[:100] or "workbook"


def build_table_first_html_report(
    *,
    batch_dir: str | Path,
    output_dir: str | Path,
) -> dict[str, Any]:
    """Render a deterministic static report from a completed table-first batch."""

    source_dir = Path(batch_dir).expanduser().resolve()
    report_path = source_dir / "batch-report.json"
    report = _load_json(report_path)
    if report.get("schemaVersion") != "table-first-batch-report-v1":
        raise TableFirstHtmlError(
            "HTML rendering requires a table-first-batch-report-v1 analysis batch"
        )
    items = sorted(
        list(report.get("items") or []),
        key=lambda item: int(item.get("index") or 0),
    )
    target_dir = Path(output_dir).expanduser().resolve()
    workbook_dir = target_dir / "workbooks"
    records: list[dict[str, Any]] = []
    for item in items:
        request_path = _artifact_path(source_dir, item, "request", "requests")
        analysis_path = _artifact_path(source_dir, item, "analysis", "analyses")
        projection_path = _artifact_path(
            source_dir, item, "projection", "projections"
        )
        request = _load_json(request_path)
        analysis = _load_json(analysis_path)
        projection = _load_json(projection_path)
        request_ids = {
            str(request.get("requestId") or ""),
            str(analysis.get("requestId") or ""),
            str(projection.get("requestId") or ""),
        }
        if len(request_ids) != 1 or not next(iter(request_ids)):
            raise TableFirstHtmlError(
                f"Mismatched request identities for batch item {item.get('index')}"
            )
        revision_uid = str(
            (projection.get("source") or {}).get("revisionUid")
            or analysis.get("revisionUid")
            or Path(projection_path).stem
        )
        detail_name = (
            f"{int(item.get('index') or 0):04d}-{_safe_slug(revision_uid)}.html"
        )
        records.append(
            {
                "item": item,
                "request": request,
                "analysis": analysis,
                "projection": projection,
                "detailPath": workbook_dir / detail_name,
                "detailFile": f"workbooks/{detail_name}",
            }
        )

    written = 0
    reused = 0
    for record in records:
        action = _write_if_changed(
            record["detailPath"],
            _render_detail(record),
        )
        written += action == "WRITTEN"
        reused += action == "REUSED"
    index_path = target_dir / "index.html"
    action = _write_if_changed(index_path, _render_index(report, records))
    written += action == "WRITTEN"
    reused += action == "REUSED"

    manifest = {
        "schemaVersion": HTML_REPORT_VERSION,
        "sourceBatchReport": str(report_path),
        "sourceStatus": str(report.get("status") or ""),
        "sourceCompletedAt": report.get("completedAt"),
        "workbookCount": len(records),
        "detailPageCount": len(records),
        "index": str(index_path),
    }
    manifest_path = target_dir / "html-report.json"
    manifest_action = _write_if_changed(
        manifest_path,
        json.dumps(manifest, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
    )
    written += manifest_action == "WRITTEN"
    reused += manifest_action == "REUSED"
    return {
        "status": "ok",
        "schemaVersion": HTML_REPORT_VERSION,
        "batchDir": str(source_dir),
        "outputDir": str(target_dir),
        "index": str(index_path),
        "manifest": str(manifest_path),
        "workbookCount": len(records),
        "detailPageCount": len(records),
        "written": written,
        "reused": reused,
    }


__all__ = [
    "HTML_REPORT_VERSION",
    "TableFirstHtmlError",
    "build_table_first_html_report",
]
