from __future__ import annotations

import csv
import datetime as dt
import hashlib
import json
import math
import os
import re
import sqlite3
import tempfile
import unicodedata
from pathlib import Path
from typing import Any, Iterable


INDEX_SCHEMA_VERSION = "table-first-history-index-v1"
PACK_SCHEMA_VERSION = "table-first-history-pack-v1"
ANSWER_SCHEMA_VERSION = "table-first-history-answer-v1"
DETAIL_SCHEMA_VERSION = "table-first-history-detail-v1"

TOKEN_PATTERN = re.compile(r"[\w]+(?:[+./%:-][\w]+)*", re.UNICODE)
DATE_YMD_PATTERN = re.compile(
    r"(?<!\d)(20\d{2})[.\-/](0?[1-9]|1[0-2])(?:[.\-/](0?[1-9]|[12]\d|3[01]))?(?!\d)"
)
DATE_DMY_PATTERN = re.compile(
    r"(?<!\d)(0?[1-9]|[12]\d|3[01])[.\-/](0?[1-9]|1[0-2])[.\-/](20\d{2})(?!\d)"
)
CELL_PATTERN = re.compile(r"^([A-Z]+)([1-9]\d*)$")
RANGE_PATTERN = re.compile(r"^([A-Z]+)([1-9]\d*):([A-Z]+)([1-9]\d*)$")

QUESTION_STOP_WORDS = {
    "ai",
    "data",
    "excel",
    "history",
    "report",
    "결과",
    "과거",
    "관련",
    "기반",
    "변경",
    "변화",
    "내역",
    "데이터",
    "리포트",
    "무엇",
    "문서",
    "보고서",
    "설명",
    "알려줘",
    "어떻게",
    "어떤",
    "엑셀",
    "이력",
    "이전",
    "자료",
    "정리",
    "조건",
    "질문",
    "따른",
    "추이",
    "찾아줘",
    "시험",
    "해줘",
}

KOREAN_SUFFIXES = (
    "으로부터",
    "에게서",
    "에서는",
    "으로는",
    "까지는",
    "부터는",
    "이라고",
    "이라는",
    "에서",
    "에게",
    "으로",
    "와는",
    "과는",
    "에는",
    "들을",
    "들의",
    "이랑",
    "랑은",
    "하고",
    "보다",
    "부터",
    "까지",
    "에서의",
    "으로의",
    "은",
    "는",
    "이",
    "가",
    "을",
    "를",
    "의",
    "에",
    "와",
    "과",
    "도",
    "로",
    "만",
)

QUERY_ALIASES = {
    "전후": ("before", "after"),
    "공급처": ("supplier", "vendor"),
    "협력사": ("supplier", "vendor"),
    "불량": ("ng", "defect", "failure"),
    "불량률": ("ng", "defect", "failure", "rate"),
    "히어링": ("hearing",),
    "조립": ("assembly", "assy", "assemble"),
    "본딩": ("bond", "bonding"),
    "본드량": ("bond", "bonding", "amount"),
    "비교": ("compare", "comparison", "versus"),
    "개선": ("improve", "improved", "improvement"),
    "기준군": ("normal", "reference", "control"),
    "대조군": ("normal", "reference", "control"),
    "공정별": ("process",),
    "이미지": ("image", "picture"),
    "숫자표": ("numeric", "table"),
    "수치": ("numeric", "value"),
    "신뢰성": ("reliability",),
}


class TableFirstHistoryError(RuntimeError):
    pass


def _canonical_bytes(value: object) -> bytes:
    return (
        json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        )
        + "\n"
    ).encode("utf-8")


def history_json_bytes(value: object) -> bytes:
    return (
        json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True) + "\n"
    ).encode("utf-8")


def _sha_id(prefix: str, *parts: object) -> str:
    digest = hashlib.sha256(
        "\x1f".join(str(part) for part in parts).encode("utf-8")
    ).hexdigest()[:20]
    return f"{prefix}-{digest.upper()}"


def normalize_text(value: object) -> str:
    return " ".join(
        unicodedata.normalize("NFKC", str(value or "")).casefold().split()
    )


def _strip_korean_suffix(token: str) -> str:
    for suffix in KOREAN_SUFFIXES:
        if len(token) > len(suffix) + 1 and token.endswith(suffix):
            return token[: -len(suffix)]
    return token


def question_tokens(value: object) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for match in TOKEN_PATTERN.finditer(normalize_text(value)):
        token = _strip_korean_suffix(match.group(0).strip("._-/"))
        if (
            len(token) < 2
            or token in QUESTION_STOP_WORDS
            or token.endswith(("설명해줘", "알려줘", "해주세요", "해줘"))
            or token in seen
        ):
            continue
        seen.add(token)
        result.append(token)
    return result


def _load_json(path: Path, schema: str, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise TableFirstHistoryError(f"{label} not found: {path}") from exc
    except json.JSONDecodeError as exc:
        raise TableFirstHistoryError(f"{label} is invalid JSON: {path}") from exc
    if not isinstance(value, dict) or value.get("schemaVersion") != schema:
        raise TableFirstHistoryError(f"Unexpected {label} schema: {path}")
    return value


def _artifact_path(batch_dir: Path, item: dict[str, Any], kind: str) -> Path:
    configured = Path(str(item.get(kind) or ""))
    local = batch_dir / f"{kind}s" / configured.name
    if local.is_file():
        return local
    if configured.is_file():
        return configured
    raise TableFirstHistoryError(
        f"Missing {kind} artifact for {item.get('fileName')}: {local}"
    )


def _string_values(value: object) -> list[str]:
    result: list[str] = []
    if isinstance(value, str):
        if value.strip():
            result.append(value.strip())
    elif isinstance(value, dict):
        for item in value.values():
            result.extend(_string_values(item))
    elif isinstance(value, list):
        for item in value:
            result.extend(_string_values(item))
    return result


def _date_candidates(entries: Iterable[tuple[str, str]]) -> list[dict[str, str]]:
    result: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for origin, text in entries:
        for match in DATE_YMD_PATTERN.finditer(text):
            year, month, day = match.groups()
            try:
                normalized = (
                    dt.date(int(year), int(month), int(day)).isoformat()
                    if day
                    else f"{int(year):04d}-{int(month):02d}"
                )
            except ValueError:
                continue
            key = (normalized, match.group(0))
            if key not in seen:
                seen.add(key)
                result.append(
                    {
                        "normalized": normalized,
                        "sourceText": match.group(0),
                        "origin": origin,
                    }
                )
        for match in DATE_DMY_PATTERN.finditer(text):
            day, month, year = match.groups()
            try:
                normalized = dt.date(
                    int(year), int(month), int(day)
                ).isoformat()
            except ValueError:
                continue
            key = (normalized, match.group(0))
            if key not in seen:
                seen.add(key)
                result.append(
                    {
                        "normalized": normalized,
                        "sourceText": match.group(0),
                        "origin": origin,
                    }
                )
    return result


def _term_rows(request: dict[str, Any]) -> list[dict[str, str]]:
    adapter = request.get("codeOwnedTermDictionary") or {}
    source = Path(str(adapter.get("sourcePath") or ""))
    result: list[dict[str, str]] = []
    if source.is_file():
        try:
            with source.open("r", encoding="utf-8-sig", newline="") as stream:
                for row in csv.DictReader(stream):
                    if str(row.get("definition_status") or "").upper() != "DEFINED":
                        continue
                    raw = str(row.get("term_raw") or "").strip()
                    normalized = str(row.get("normalized_name") or "").strip()
                    if not raw or not normalized:
                        continue
                    result.append(
                        {
                            "term": raw,
                            "canonical": normalized,
                            "koreanDescription": str(
                                row.get("korean_desc") or ""
                            ).strip(),
                            "notes": str(row.get("notes") or "").strip(),
                        }
                    )
        except (OSError, csv.Error):
            result = []
    if result:
        return result
    for group in adapter.get("aliasGroups") or []:
        canonical = str(group.get("normalizedName") or "").strip()
        for term in group.get("terms") or []:
            if canonical and str(term).strip():
                result.append(
                    {
                        "term": str(term).strip(),
                        "canonical": canonical,
                        "koreanDescription": "",
                        "notes": "",
                    }
                )
    return result


DDL = """
PRAGMA foreign_keys=ON;
CREATE TABLE history_meta (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
CREATE TABLE history_terms (
    term TEXT NOT NULL,
    canonical TEXT NOT NULL,
    korean_description TEXT NOT NULL,
    notes TEXT NOT NULL,
    PRIMARY KEY(term, canonical)
);
CREATE TABLE history_workbooks (
    workbook_id INTEGER PRIMARY KEY,
    public_workbook_id TEXT NOT NULL UNIQUE,
    request_id TEXT NOT NULL UNIQUE,
    revision_uid TEXT NOT NULL,
    content_sha256 TEXT NOT NULL,
    file_name TEXT NOT NULL,
    source_path TEXT NOT NULL,
    analysis_status TEXT NOT NULL,
    verification_status TEXT NOT NULL,
    query_eligibility TEXT NOT NULL,
    workbook_summary TEXT NOT NULL,
    primary_date TEXT,
    date_candidates_json TEXT NOT NULL,
    request_path TEXT NOT NULL,
    analysis_path TEXT NOT NULL,
    projection_path TEXT NOT NULL,
    text_blocks_json TEXT NOT NULL,
    search_text TEXT NOT NULL
);
CREATE TABLE history_studies (
    study_id INTEGER PRIMARY KEY,
    workbook_id INTEGER NOT NULL REFERENCES history_workbooks(workbook_id)
        ON DELETE CASCADE,
    ordinal INTEGER NOT NULL,
    public_study_id TEXT NOT NULL UNIQUE,
    study_group TEXT NOT NULL,
    titles_json TEXT NOT NULL,
    table_types_json TEXT NOT NULL,
    groups_json TEXT NOT NULL,
    metrics_json TEXT NOT NULL,
    comparison_relations_json TEXT NOT NULL,
    numeric_facts_json TEXT NOT NULL,
    numeric_series_json TEXT NOT NULL,
    limitations_json TEXT NOT NULL,
    verification_status TEXT NOT NULL,
    payload_json TEXT NOT NULL,
    search_text TEXT NOT NULL,
    UNIQUE(workbook_id, ordinal)
);
CREATE TABLE history_evidence (
    evidence_id INTEGER PRIMARY KEY,
    study_id INTEGER NOT NULL REFERENCES history_studies(study_id)
        ON DELETE CASCADE,
    public_evidence_id TEXT NOT NULL UNIQUE,
    table_id TEXT NOT NULL,
    sheet TEXT NOT NULL,
    range_address TEXT NOT NULL,
    request_preview_json TEXT NOT NULL
);
CREATE INDEX history_workbooks_date_idx ON history_workbooks(primary_date);
CREATE INDEX history_studies_workbook_idx ON history_studies(workbook_id);
CREATE INDEX history_evidence_study_idx ON history_evidence(study_id);
"""


def _table_preview(request: dict[str, Any], table_id: str) -> dict[str, Any]:
    for table in request.get("tables") or []:
        if str(table.get("tableId") or "") == table_id:
            return {
                "tableId": table_id,
                "sheet": str(table.get("sheet") or ""),
                "range": str(table.get("range") or ""),
                "bounds": table.get("bounds") or {},
                "previewRows": table.get("previewRows") or [],
            }
    return {}


def _enriched_search_text(
    texts: Iterable[str],
    terms: list[dict[str, str]],
) -> str:
    base = " \n ".join(text for text in texts if text)
    normalized = normalize_text(base)
    additions: list[str] = []
    for row in terms:
        term = normalize_text(row["term"])
        if not term or term not in normalized:
            continue
        additions.extend(
            [
                row["canonical"],
                row["koreanDescription"],
                row["notes"],
            ]
        )
    return normalize_text(" \n ".join([base, *additions]))


def build_history_index(
    batch_dir: str | Path,
    database_path: str | Path,
    *,
    require_complete: bool = True,
) -> dict[str, Any]:
    batch = Path(batch_dir).expanduser().resolve()
    report = _load_json(
        batch / "batch-report.json",
        "table-first-batch-report-v1",
        "table-first batch report",
    )
    if require_complete and report.get("status") != "ok":
        raise TableFirstHistoryError(
            f"Batch is not complete: {report.get('status') or 'UNKNOWN'}"
        )
    target = Path(database_path).expanduser().resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(
        prefix=f".{target.name}.", suffix=".tmp", dir=target.parent
    )
    os.close(fd)
    temp_path = Path(temp_name)
    workbook_count = study_count = evidence_count = 0
    try:
        with sqlite3.connect(temp_path) as connection:
            connection.executescript(DDL)
            items = sorted(
                report.get("items") or [],
                key=lambda item: int(item.get("index") or 0),
            )
            terms: list[dict[str, str]] = []
            for item in items:
                request_path = _artifact_path(batch, item, "request")
                analysis_path = _artifact_path(batch, item, "analysis")
                projection_path = _artifact_path(batch, item, "projection")
                request = _load_json(
                    request_path, "table-first-request-v1", "request"
                )
                analysis = _load_json(
                    analysis_path, "table-first-analysis-v1", "analysis"
                )
                projection = _load_json(
                    projection_path, "table-first-projection-v1", "projection"
                )
                request_id = str(request.get("requestId") or "")
                if not request_id or {
                    str(analysis.get("requestId") or ""),
                    str(projection.get("requestId") or ""),
                } != {request_id}:
                    raise TableFirstHistoryError(
                        f"Artifact requestId mismatch: {projection_path}"
                    )
                source = projection.get("source") or request.get("source") or {}
                file_name = str(source.get("fileName") or item.get("fileName") or "")
                source_path = str(source.get("sourcePath") or "")
                content_sha256 = str(source.get("contentSha256") or "")
                revision_uid = str(source.get("revisionUid") or "")
                summary = str(analysis.get("workbookSummary") or "")
                text_blocks = projection.get("textBlocks") or []
                dates = _date_candidates(
                    [
                        ("FILE_NAME", file_name),
                        ("WORKBOOK_SUMMARY", summary),
                        ("TEXT_BLOCK", " ".join(_string_values(text_blocks))),
                    ]
                )
                primary_date = dates[0]["normalized"] if dates else None
                if not terms:
                    terms = _term_rows(request)
                    connection.executemany(
                        """
                        INSERT OR IGNORE INTO history_terms(
                            term, canonical, korean_description, notes
                        ) VALUES (?, ?, ?, ?)
                        """,
                        [
                            (
                                row["term"],
                                row["canonical"],
                                row["koreanDescription"],
                                row["notes"],
                            )
                            for row in terms
                        ],
                    )
                workbook_search = _enriched_search_text(
                    [file_name, source_path, summary, *_string_values(text_blocks)],
                    terms,
                )
                public_workbook_id = _sha_id(
                    "TF-WBK", content_sha256, revision_uid
                )
                cursor = connection.execute(
                    """
                    INSERT INTO history_workbooks(
                        public_workbook_id, request_id, revision_uid,
                        content_sha256, file_name, source_path,
                        analysis_status, verification_status,
                        query_eligibility, workbook_summary, primary_date,
                        date_candidates_json, request_path, analysis_path,
                        projection_path, text_blocks_json, search_text
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        public_workbook_id,
                        request_id,
                        revision_uid,
                        content_sha256,
                        file_name,
                        source_path,
                        str(projection.get("analysisStatus") or analysis.get("status") or ""),
                        str(projection.get("verificationStatus") or "NEEDS_REVIEW"),
                        str(projection.get("queryEligibility") or ""),
                        summary,
                        primary_date,
                        json.dumps(dates, ensure_ascii=False, sort_keys=True),
                        str(request_path),
                        str(analysis_path),
                        str(projection_path),
                        json.dumps(text_blocks, ensure_ascii=False, sort_keys=True),
                        workbook_search,
                    ),
                )
                workbook_id = int(cursor.lastrowid)
                workbook_count += 1
                for ordinal, study in enumerate(projection.get("studies") or [], start=1):
                    study_group = str(study.get("studyGroup") or f"Study {ordinal}")
                    public_study_id = _sha_id(
                        "TF-STU", request_id, ordinal, study_group
                    )
                    titles = study.get("titles") or []
                    groups = study.get("groups") or []
                    metrics = study.get("metrics") or []
                    relations = study.get("comparisonRelations") or []
                    limitations = study.get("limitations") or []
                    study_search = _enriched_search_text(
                        [
                            file_name,
                            summary,
                            study_group,
                            *_string_values(titles),
                            *_string_values(groups),
                            *_string_values(metrics),
                            *_string_values(relations),
                            *_string_values(limitations),
                        ],
                        terms,
                    )
                    study_cursor = connection.execute(
                        """
                        INSERT INTO history_studies(
                            workbook_id, ordinal, public_study_id, study_group,
                            titles_json, table_types_json, groups_json,
                            metrics_json, comparison_relations_json,
                            numeric_facts_json, numeric_series_json,
                            limitations_json, verification_status,
                            payload_json, search_text
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            workbook_id,
                            ordinal,
                            public_study_id,
                            study_group,
                            json.dumps(titles, ensure_ascii=False, sort_keys=True),
                            json.dumps(study.get("tableTypes") or [], ensure_ascii=False, sort_keys=True),
                            json.dumps(groups, ensure_ascii=False, sort_keys=True),
                            json.dumps(metrics, ensure_ascii=False, sort_keys=True),
                            json.dumps(relations, ensure_ascii=False, sort_keys=True),
                            json.dumps(study.get("deterministicNumericFacts") or [], ensure_ascii=False, sort_keys=True),
                            json.dumps(study.get("deterministicNumericSeries") or [], ensure_ascii=False, sort_keys=True),
                            json.dumps(limitations, ensure_ascii=False, sort_keys=True),
                            str(study.get("verificationStatus") or "NEEDS_REVIEW"),
                            json.dumps(study, ensure_ascii=False, sort_keys=True),
                            study_search,
                        ),
                    )
                    study_id = int(study_cursor.lastrowid)
                    study_count += 1
                    for evidence in study.get("evidence") or []:
                        table_id = str(evidence.get("tableId") or "")
                        sheet = str(evidence.get("sheet") or "")
                        address = str(evidence.get("range") or "")
                        public_evidence_id = _sha_id(
                            "TF-EVD", public_study_id, table_id, sheet, address
                        )
                        connection.execute(
                            """
                            INSERT INTO history_evidence(
                                study_id, public_evidence_id, table_id,
                                sheet, range_address, request_preview_json
                            ) VALUES (?, ?, ?, ?, ?, ?)
                            """,
                            (
                                study_id,
                                public_evidence_id,
                                table_id,
                                sheet,
                                address,
                                json.dumps(
                                    _table_preview(request, table_id),
                                    ensure_ascii=False,
                                    sort_keys=True,
                                ),
                            ),
                        )
                        evidence_count += 1
            metadata = {
                "schemaVersion": INDEX_SCHEMA_VERSION,
                "batchDir": str(batch),
                "batchReportSha256": hashlib.sha256(
                    (batch / "batch-report.json").read_bytes()
                ).hexdigest(),
                "builderVersion": str(report.get("builderVersion") or ""),
                "promptVersion": str(report.get("promptVersion") or ""),
                "workbookCount": workbook_count,
                "studyCount": study_count,
                "evidenceCount": evidence_count,
                "termCount": len(terms),
            }
            connection.executemany(
                "INSERT INTO history_meta(key, value) VALUES (?, ?)",
                [
                    (key, json.dumps(value, ensure_ascii=False, sort_keys=True))
                    for key, value in metadata.items()
                ],
            )
            connection.execute("PRAGMA user_version=1")
            connection.commit()
            quick_check = connection.execute("PRAGMA quick_check").fetchone()[0]
            if quick_check != "ok":
                raise TableFirstHistoryError(
                    f"History index quick_check failed: {quick_check}"
                )
        connection.close()
        os.replace(temp_path, target)
    except Exception:
        if "connection" in locals():
            connection.close()
        temp_path.unlink(missing_ok=True)
        raise
    return {
        "schemaVersion": INDEX_SCHEMA_VERSION,
        "status": "ok",
        "database": str(target),
        "batchDir": str(batch),
        "workbookCount": workbook_count,
        "studyCount": study_count,
        "evidenceCount": evidence_count,
        "termCount": len(terms),
    }


def _json(text: str) -> Any:
    return json.loads(text)


def _term_groups(
    connection: sqlite3.Connection, tokens: list[str]
) -> list[dict[str, Any]]:
    rows = connection.execute(
        "SELECT term, canonical, korean_description, notes FROM history_terms"
    ).fetchall()
    result: list[dict[str, Any]] = []
    for token in tokens:
        variants = {token}
        variants.update(QUERY_ALIASES.get(token, ()))
        concepts: set[str] = set()
        lookup_tokens = {token}
        rate_suffix = token.endswith(("율", "률")) and len(token) > 2
        if rate_suffix:
            lookup_tokens.add(token[:-1])
            variants.update({"rate", "ratio"})
        for term, canonical, korean_description, notes in rows:
            normalized_term = normalize_text(term)
            normalized_canonical = normalize_text(canonical)
            korean_words = set(
                TOKEN_PATTERN.findall(normalize_text(korean_description))
            )
            matched = any(
                lookup == normalized_term
                or lookup == normalized_canonical
                or (
                    any(ord(character) > 127 for character in lookup)
                    and lookup in korean_words
                )
                for lookup in lookup_tokens
            )
            if not matched:
                continue
            concepts.add(str(canonical))
        if concepts:
            for term, canonical, korean_description, notes in rows:
                if str(canonical) not in concepts:
                    continue
                variants.update(question_tokens(term))
                variants.update(question_tokens(canonical))
                variants.update(question_tokens(korean_description))
        result.append(
            {
                "queryToken": token,
                "variants": sorted(value for value in variants if value),
                "concepts": sorted(concepts),
            }
        )
    return result


def _field_score(
    text: str,
    term_groups: list[dict[str, Any]],
    weight: float,
) -> tuple[float, list[str]]:
    normalized = normalize_text(text)
    searchable_tokens = set(
        TOKEN_PATTERN.findall(normalized.replace("_", " "))
    )
    matched: list[str] = []
    for group in term_groups:
        if any(
            _variant_matches(variant, normalized, searchable_tokens)
            for variant in group["variants"]
        ):
            matched.append(str(group["queryToken"]))
    importance = sum(
        float(group.get("importance") or 1.0)
        for group in term_groups
        if str(group["queryToken"]) in matched
    )
    return weight * importance, matched


def _variant_matches(
    variant: object,
    normalized_text: str,
    searchable_tokens: set[str],
) -> bool:
    normalized_variant = normalize_text(variant).strip("._-/")
    if not normalized_variant:
        return False
    if normalized_variant.isascii():
        if normalized_variant in searchable_tokens:
            return True
        if any(character in "+/:%-." for character in normalized_variant):
            compact_variant = re.sub(r"[^a-z0-9]", "", normalized_variant)
            if len(compact_variant) >= 6:
                compact_text = re.sub(r"[^a-z0-9]", "", normalized_text)
                return compact_variant in compact_text
        return False
    return normalized_variant in normalized_text


def _assign_group_importance(
    groups: list[dict[str, Any]],
    workbook_search_texts: list[str],
) -> None:
    workbook_count = max(len(workbook_search_texts), 1)
    for group in groups:
        document_frequency = 0
        for text in workbook_search_texts:
            normalized = normalize_text(text)
            searchable_tokens = set(
                TOKEN_PATTERN.findall(normalized.replace("_", " "))
            )
            if any(
                _variant_matches(variant, normalized, searchable_tokens)
                for variant in group["variants"]
            ):
                document_frequency += 1
        importance = 1.0 + math.log(
            (workbook_count + 1.0) / (document_frequency + 1.0)
        )
        query_token = str(group["queryToken"])
        if any(character.isdigit() for character in query_token):
            importance *= 1.25
        if any(character in "+/:%-" for character in query_token):
            importance *= 1.2
        group["documentFrequency"] = document_frequency
        group["importance"] = round(importance, 6)


def _study_row_payload(row: sqlite3.Row) -> dict[str, Any]:
    fields = {
        "fileName": row["file_name"],
        "sourcePath": row["source_path"],
        "workbookSummary": row["workbook_summary"],
        "studyGroup": row["study_group"],
        "titles": _json(row["titles_json"]),
        "groups": _json(row["groups_json"]),
        "metrics": _json(row["metrics_json"]),
        "comparisonRelations": _json(row["comparison_relations_json"]),
        "limitations": _json(row["limitations_json"]),
    }
    return fields


def build_history_pack(
    database_path: str | Path,
    question: str,
    *,
    limit: int = 30,
) -> dict[str, Any]:
    db = Path(database_path).expanduser().resolve()
    if not db.is_file():
        raise TableFirstHistoryError(f"History index not found: {db}")
    tokens = question_tokens(question)
    if not tokens:
        raise TableFirstHistoryError("Question has no searchable terms.")
    connection = sqlite3.connect(f"file:{db.as_posix()}?mode=ro", uri=True)
    connection.row_factory = sqlite3.Row
    try:
        schema_row = connection.execute(
            "SELECT value FROM history_meta WHERE key='schemaVersion'"
        ).fetchone()
        if not schema_row or _json(schema_row[0]) != INDEX_SCHEMA_VERSION:
            raise TableFirstHistoryError("Unsupported history index schema.")
        groups = _term_groups(connection, tokens)
        workbook_search_texts = [
            str(row[0])
            for row in connection.execute(
                "SELECT search_text FROM history_workbooks"
            ).fetchall()
        ]
        _assign_group_importance(groups, workbook_search_texts)
        total_importance = sum(
            float(group.get("importance") or 1.0) for group in groups
        )
        rows = connection.execute(
            """
            SELECT s.*, w.public_workbook_id, w.request_id, w.file_name,
                   w.source_path, w.analysis_status,
                   w.verification_status AS workbook_verification_status,
                   w.query_eligibility, w.workbook_summary, w.primary_date,
                   w.date_candidates_json, w.search_text AS workbook_search_text
            FROM history_studies s
            JOIN history_workbooks w ON w.workbook_id=s.workbook_id
            """
        ).fetchall()
        ranked: list[tuple[float, int, sqlite3.Row, list[str]]] = []
        for row in rows:
            payload = _study_row_payload(row)
            score = 0.0
            matched: set[str] = set()
            weighted_fields = [
                (payload["fileName"], 5.0),
                (payload["workbookSummary"], 4.0),
                (payload["studyGroup"], 7.0),
                (" ".join(_string_values(payload["titles"])), 6.0),
                (" ".join(_string_values(payload["groups"])), 5.0),
                (" ".join(_string_values(payload["metrics"])), 5.0),
                (" ".join(_string_values(payload["comparisonRelations"])), 4.0),
                (" ".join(_string_values(payload["limitations"])), 1.0),
            ]
            for text, weight in weighted_fields:
                value, field_matches = _field_score(text, groups, weight)
                score += value
                matched.update(field_matches)
            if not matched:
                continue
            coverage = len(matched)
            if coverage < min(2, len(groups)):
                continue
            matched_importance = sum(
                float(group.get("importance") or 1.0)
                for group in groups
                if str(group["queryToken"]) in matched
            )
            score += 20.0 * matched_importance / total_importance
            if coverage == len(groups):
                score += 50.0
            ranked.append((score, coverage, row, sorted(matched)))
        ranked.sort(
            key=lambda item: (
                -item[0],
                -item[1],
                str(item[2]["primary_date"] or "9999-99-99"),
                str(item[2]["file_name"]),
                int(item[2]["ordinal"]),
            )
        )
        selected: list[tuple[float, int, sqlite3.Row, list[str]]] = []
        selected_workbooks: set[str] = set()
        bounded_limit = max(1, min(int(limit), 200))
        for candidate in ranked:
            workbook_key = str(candidate[2]["public_workbook_id"])
            if workbook_key in selected_workbooks:
                continue
            selected.append(candidate)
            selected_workbooks.add(workbook_key)
            if len(selected) >= bounded_limit:
                break
        studies: list[dict[str, Any]] = []
        citations: dict[str, dict[str, Any]] = {}
        for score, coverage, row, matched in selected:
            evidence_rows = connection.execute(
                """
                SELECT public_evidence_id, table_id, sheet, range_address
                FROM history_evidence WHERE study_id=?
                ORDER BY evidence_id
                """,
                (row["study_id"],),
            ).fetchall()
            study_citations: list[str] = []
            for evidence in evidence_rows:
                evidence_id = str(evidence["public_evidence_id"])
                study_citations.append(evidence_id)
                citations[evidence_id] = {
                    "evidenceId": evidence_id,
                    "sourcePath": str(row["source_path"]),
                    "sheet": str(evidence["sheet"]),
                    "range": str(evidence["range_address"]),
                    "verificationStatus": str(row["verification_status"]),
                    "tableId": str(evidence["table_id"]),
                    "studyId": str(row["public_study_id"]),
                }
            studies.append(
                {
                    "studyId": str(row["public_study_id"]),
                    "workbookId": str(row["public_workbook_id"]),
                    "requestId": str(row["request_id"]),
                    "fileName": str(row["file_name"]),
                    "sourcePath": str(row["source_path"]),
                    "date": row["primary_date"],
                    "dateCandidates": _json(row["date_candidates_json"]),
                    "workbookSummary": str(row["workbook_summary"]),
                    "studyGroup": str(row["study_group"]),
                    "titles": _json(row["titles_json"]),
                    "groups": _json(row["groups_json"]),
                    "metrics": _json(row["metrics_json"]),
                    "comparisonRelations": _json(row["comparison_relations_json"]),
                    "limitations": _json(row["limitations_json"]),
                    "analysisStatus": str(row["analysis_status"]),
                    "verificationStatus": str(row["verification_status"]),
                    "queryEligibility": str(row["query_eligibility"]),
                    "score": round(score, 3),
                    "matchedQueryTerms": matched,
                    "queryTermCoverage": coverage,
                    "citationIds": study_citations,
                }
            )
        source_exclusions: list[dict[str, Any]] = []
        terminal_rows = connection.execute(
            """
            SELECT w.*
            FROM history_workbooks w
            WHERE w.analysis_status='NO_TABLES'
               OR NOT EXISTS (
                    SELECT 1 FROM history_studies s
                    WHERE s.workbook_id=w.workbook_id
               )
            """
        ).fetchall()
        terminal_ranked: list[
            tuple[float, int, sqlite3.Row, list[str]]
        ] = []
        for row in terminal_rows:
            score = 0.0
            matched: set[str] = set()
            for text, weight in (
                (str(row["file_name"]), 7.0),
                (str(row["workbook_summary"]), 5.0),
                (str(row["search_text"]), 1.0),
            ):
                value, field_matches = _field_score(text, groups, weight)
                score += value
                matched.update(field_matches)
            if not matched:
                continue
            coverage = len(matched)
            if coverage < min(2, len(groups)):
                continue
            matched_importance = sum(
                float(group.get("importance") or 1.0)
                for group in groups
                if str(group["queryToken"]) in matched
            )
            score += 20.0 * matched_importance / total_importance
            terminal_ranked.append((score, coverage, row, sorted(matched)))
        terminal_ranked.sort(
            key=lambda item: (
                -item[0],
                -item[1],
                str(item[2]["primary_date"] or "9999-99-99"),
                str(item[2]["file_name"]),
            )
        )
        for score, coverage, row, matched in terminal_ranked[:20]:
            source_exclusions.append(
                {
                    "workbookId": str(row["public_workbook_id"]),
                    "requestId": str(row["request_id"]),
                    "fileName": str(row["file_name"]),
                    "sourcePath": str(row["source_path"]),
                    "date": row["primary_date"],
                    "dateCandidates": _json(row["date_candidates_json"]),
                    "workbookSummary": str(row["workbook_summary"]),
                    "analysisStatus": str(row["analysis_status"]),
                    "verificationStatus": str(row["verification_status"]),
                    "queryEligibility": str(row["query_eligibility"]),
                    "reason": (
                        "No table-first Study was produced; no numeric effect "
                        "is available from this source."
                    ),
                    "score": round(score, 3),
                    "matchedQueryTerms": matched,
                    "queryTermCoverage": coverage,
                }
            )
        workbook_ids = {item["workbookId"] for item in studies}
        workbook_ids.update(
            item["workbookId"] for item in source_exclusions
        )
        workbook_count = len(workbook_ids)
        return {
            "schemaVersion": PACK_SCHEMA_VERSION,
            "question": question,
            "database": str(db),
            "queryTerms": groups,
            "summary": {
                "relevantWorkbookCount": workbook_count,
                "relevantStudyCount": len(studies),
                "eligibleEffectCount": 0,
                "citationCount": len(citations),
                "totalIndexedStudyCount": len(rows),
            },
            "studies": studies,
            "sourceExclusions": source_exclusions,
            "citations": [citations[key] for key in sorted(citations)],
            "trust": {
                "quantitativeClaimsAllowed": False,
                "reason": (
                    "Table-first histories are searchable semantic projections. "
                    "They remain review-gated and do not create approved effects."
                ),
            },
        }
    finally:
        connection.close()


def _display_groups(groups: list[dict[str, Any]]) -> str:
    values = []
    for group in groups[:12]:
        label = str(group.get("label") or "").strip()
        role = str(group.get("role") or "UNASSESSED").strip()
        if label:
            values.append(f"{label}({role})")
    return ", ".join(values) if values else "시험군 정보 없음"


def _display_metrics(metrics: list[dict[str, Any]]) -> str:
    values = []
    for metric in metrics[:18]:
        name = str(metric.get("name") or "").strip()
        unit = str(metric.get("unit") or "").strip()
        if name:
            values.append(f"{name}{f' [{unit}]' if unit else ''}")
    suffix = " 외" if len(metrics) > 18 else ""
    return ", ".join(values) + suffix if values else "명시 지표 없음"


def _answer_markdown(pack: dict[str, Any]) -> str:
    studies = pack.get("studies") or []
    detailed_studies = studies[:12]
    compact_studies = studies[12:]
    source_exclusions = pack.get("sourceExclusions") or []
    if not studies and not source_exclusions:
        return (
            f"질문: {pack['question']}\n\n"
            "현재 인덱스에서 질문과 연결되는 시험 이력을 찾지 못했습니다. "
            "제품명, 공정명, 부품명 또는 불량 항목을 더 구체적으로 입력하세요."
        )
    lines = [
        f"질문: {pack['question']}",
        "",
        (
            f"관련 workbook {pack['summary']['relevantWorkbookCount']}건, "
            f"Study {pack['summary']['relevantStudyCount']}건을 찾았습니다."
        ),
        (
            "아래 내용은 원본 표를 AI가 구조화한 검토 대기 이력입니다. "
            "승인된 효과 계산이나 인과 결론으로 해석하면 안 됩니다."
        ),
        "",
        "## 관련 시험 이력",
    ]
    for item in detailed_studies:
        date = item.get("date") or "날짜 미확정"
        citation_ids = item.get("citationIds") or []
        citation_text = ", ".join(citation_ids) if citation_ids else "직접 표 근거 없음"
        lines.extend(
            [
                "",
                f"### {date} · {item['studyGroup']}",
                f"- 원본: {item['fileName']}",
                f"- workbook 요약: {item['workbookSummary'] or '요약 없음'}",
                f"- 시험군/기준군: {_display_groups(item['groups'])}",
                f"- 확인 지표: {_display_metrics(item['metrics'])}",
            ]
        )
        relations = item.get("comparisonRelations") or []
        if relations:
            rendered = []
            for relation in relations[:12]:
                left = str(relation.get("leftGroup") or "?")
                right = str(relation.get("rightGroup") or "?")
                rendered.append(f"{left} → {right}")
            lines.append(f"- 비교 관계: {', '.join(rendered)}")
        if item.get("limitations"):
            lines.append(
                "- 제한: " + "; ".join(str(value) for value in item["limitations"][:4])
            )
        lines.append(f"- 원본 근거: {citation_text}")
    if compact_studies:
        lines.extend(
            [
                "",
                "## 추가 관련 이력",
                "",
                (
                    "아래 항목은 검색 범위를 보존하되 읽기 부담을 줄이기 위해 "
                    "요약 표시합니다."
                ),
            ]
        )
        for item in compact_studies:
            date = item.get("date") or "날짜 미확정"
            citation_ids = item.get("citationIds") or []
            citation_text = ", ".join(citation_ids) if citation_ids else "표 근거 없음"
            lines.append(
                f"- {date} · {item['studyGroup']} · {item['fileName']} "
                f"(근거: {citation_text})"
            )
    if source_exclusions:
        lines.extend(["", "## 표 기반 분석에서 제외된 관련 원본"])
        for item in source_exclusions:
            date = item.get("date") or "날짜 미확정"
            lines.extend(
                [
                    "",
                    f"### {date} · {item['fileName']}",
                    f"- 상태: {item['analysisStatus']}",
                    f"- 요약: {item['workbookSummary'] or '표 기반 Study 없음'}",
                    f"- 제외 사유: {item['reason']}",
                    f"- 원본: {item['sourcePath']}",
                ]
            )
    lines.extend(
        [
            "",
            "## 해석 제한",
            "",
            "- NEEDS_REVIEW/ANALYZED 기록은 검색과 이력 설명에는 표시하지만 검증된 효과로 승격하지 않습니다.",
            "- 시험군별 수치가 안전하게 연결되지 않은 경우 값을 추정하거나 다른 열의 통계를 복제하지 않습니다.",
            "- 정량 차이와 원인 결론은 원본 표 검토 및 비교 적격성 승인 후에만 생성할 수 있습니다.",
        ]
    )
    return "\n".join(lines)


def build_history_answer(pack: dict[str, Any]) -> dict[str, Any]:
    if pack.get("schemaVersion") != PACK_SCHEMA_VERSION:
        raise TableFirstHistoryError("Expected a table-first history pack.")
    pack_sha256 = hashlib.sha256(_canonical_bytes(pack)).hexdigest()
    has_results = bool(pack.get("studies") or pack.get("sourceExclusions"))
    answer = {
        "schemaVersion": ANSWER_SCHEMA_VERSION,
        "answerStatus": (
            "REVIEW_GATED_HISTORY_FOUND" if has_results else "NO_RELEVANT_HISTORY"
        ),
        "question": str(pack.get("question") or ""),
        "evidencePackSha256": pack_sha256,
        "coverage": {
            "relevantStudyCount": int(pack["summary"]["relevantStudyCount"]),
            "relevantWorkbookCount": int(
                pack["summary"]["relevantWorkbookCount"]
            ),
            "eligibleEffectCount": 0,
            "citationCount": int(pack["summary"]["citationCount"]),
        },
        "markdown": _answer_markdown(pack),
        "citations": pack.get("citations") or [],
        "limitations": [
            "검색된 table-first 의미 결과는 사람 검토 전 이력입니다.",
            "승인되지 않은 수치 효과와 인과관계는 답변에 생성하지 않습니다.",
        ],
    }
    validate_history_answer(answer, pack)
    return answer


def validate_history_answer(
    answer: dict[str, Any], pack: dict[str, Any]
) -> None:
    if answer.get("schemaVersion") != ANSWER_SCHEMA_VERSION:
        raise TableFirstHistoryError("Unexpected history answer schema.")
    expected_sha = hashlib.sha256(_canonical_bytes(pack)).hexdigest()
    if answer.get("evidencePackSha256") != expected_sha:
        raise TableFirstHistoryError("History answer does not match its pack.")
    pack_ids = {
        str(item.get("evidenceId") or "") for item in pack.get("citations") or []
    }
    answer_ids = {
        str(item.get("evidenceId") or "") for item in answer.get("citations") or []
    }
    if answer_ids != pack_ids:
        raise TableFirstHistoryError("History answer citation set changed.")
    if int((answer.get("coverage") or {}).get("eligibleEffectCount") or 0) != 0:
        raise TableFirstHistoryError("History answer cannot expose approved effects.")
    expected = _answer_markdown(pack)
    if answer.get("markdown") != expected:
        raise TableFirstHistoryError("History answer wording is not deterministic.")


def render_history_answer_markdown(answer: dict[str, Any]) -> str:
    return str(answer.get("markdown") or "").rstrip() + "\n"


def run_history_acceptance(
    database_path: str | Path,
    manifest_path: str | Path,
    *,
    query_limit: int = 30,
) -> dict[str, Any]:
    manifest_file = Path(manifest_path).expanduser().resolve()
    try:
        manifest = json.loads(manifest_file.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise TableFirstHistoryError(
            f"Invalid acceptance manifest: {manifest_file}"
        ) from exc
    workbook_names = {
        str(item.get("id") or ""): Path(
            str(item.get("relativePath") or "")
        ).name
        for item in manifest.get("workbooks") or []
    }
    golden_questions = manifest.get("goldenQuestions") or []
    if not golden_questions:
        raise TableFirstHistoryError(
            f"Acceptance manifest has no golden questions: {manifest_file}"
        )
    database_file = Path(database_path).expanduser().resolve()
    try:
        connection = sqlite3.connect(
            f"file:{database_file.as_posix()}?mode=ro", uri=True
        )
        try:
            indexed_file_keys = {
                normalize_text(Path(str(row[0])).name)
                for row in connection.execute(
                    "SELECT file_name FROM history_workbooks"
                ).fetchall()
            }
        finally:
            connection.close()
    except sqlite3.Error as exc:
        raise TableFirstHistoryError(
            f"Could not inspect acceptance database: {database_file}"
        ) from exc
    results: list[dict[str, Any]] = []
    for item in golden_questions:
        question_id = str(item.get("id") or "")
        question = str(item.get("question") or "")
        if not question_id or not question:
            raise TableFirstHistoryError(
                "Every acceptance question requires a non-empty id and question."
            )
        primary_ids = [
            str(value) for value in item.get("primaryPilotIds") or []
        ]
        unknown_primary_ids = sorted(
            pilot_id for pilot_id in primary_ids if pilot_id not in workbook_names
        )
        pack = build_history_pack(
            database_path,
            question,
            limit=query_limit,
        )
        answer = build_history_answer(pack)
        validate_history_answer(answer, pack)
        expected_files = sorted(
            workbook_names[pilot_id]
            for pilot_id in primary_ids
            if pilot_id in workbook_names
        )
        actual_files = {
            Path(str(study.get("sourcePath") or "")).name
            for study in pack.get("studies") or []
        }
        actual_files.update(
            Path(str(source.get("sourcePath") or "")).name
            for source in pack.get("sourceExclusions") or []
        )
        actual_file_keys = {normalize_text(name) for name in actual_files}
        not_indexed = [
            name
            for name in expected_files
            if normalize_text(name) not in indexed_file_keys
        ]
        not_retrieved = [
            name
            for name in expected_files
            if normalize_text(name) in indexed_file_keys
            and normalize_text(name) not in actual_file_keys
        ]
        missing = sorted([*not_indexed, *not_retrieved])
        invalid_citations = [
            str(citation.get("evidenceId") or "")
            for citation in answer.get("citations") or []
            if not str(citation.get("evidenceId") or "").startswith(
                "TF-EVD-"
            )
        ]
        failures: list[str] = []
        if unknown_primary_ids:
            failures.append("UNKNOWN_PRIMARY_PILOT_ID")
        if not_indexed:
            failures.append("PRIMARY_SOURCE_NOT_INDEXED")
        if not_retrieved:
            failures.append("MISSING_PRIMARY_SOURCES")
        if invalid_citations:
            failures.append("INVALID_CITATION_NAMESPACE")
        if answer["coverage"]["eligibleEffectCount"] != 0:
            failures.append("UNAPPROVED_EFFECT_EXPOSED")
        results.append(
            {
                "id": question_id,
                "question": question,
                "status": "PASS" if not failures else "FAIL",
                "failures": failures,
                "unknownPrimaryPilotIds": unknown_primary_ids,
                "expectedPrimaryFiles": expected_files,
                "retrievedPrimaryFiles": sorted(
                    name
                    for name in expected_files
                    if normalize_text(name) in actual_file_keys
                ),
                "missingPrimaryFiles": missing,
                "notIndexedPrimaryFiles": not_indexed,
                "notRetrievedPrimaryFiles": not_retrieved,
                "relevantWorkbookCount": pack["summary"][
                    "relevantWorkbookCount"
                ],
                "relevantStudyCount": pack["summary"][
                    "relevantStudyCount"
                ],
                "sourceExclusionCount": len(
                    pack.get("sourceExclusions") or []
                ),
                "citationCount": answer["coverage"]["citationCount"],
                "eligibleEffectCount": 0,
            }
        )
    passed = sum(item["status"] == "PASS" for item in results)
    failed = len(results) - passed
    return {
        "schemaVersion": "table-first-history-acceptance-v1",
        "status": "PASS" if failed == 0 else "FAIL",
        "database": str(Path(database_path).expanduser().resolve()),
        "manifest": str(manifest_file),
        "summary": {
            "questionCount": len(results),
            "passed": passed,
            "failed": failed,
        },
        "questions": results,
    }


def _column_number(label: str) -> int:
    value = 0
    for character in label:
        value = value * 26 + ord(character) - ord("A") + 1
    return value


def _range_bounds(address: str) -> tuple[int, int, int, int]:
    match = RANGE_PATTERN.fullmatch(address.upper())
    if match:
        start_column, start_row, end_column, end_row = match.groups()
        return (
            int(start_row),
            _column_number(start_column),
            int(end_row),
            _column_number(end_column),
        )
    cell = CELL_PATTERN.fullmatch(address.upper())
    if not cell:
        return (1, 1, 1, 1)
    column, row = cell.groups()
    number = _column_number(column)
    return (int(row), number, int(row), number)


def build_history_detail(
    database_path: str | Path, evidence_id: str
) -> dict[str, Any]:
    db = Path(database_path).expanduser().resolve()
    connection = sqlite3.connect(f"file:{db.as_posix()}?mode=ro", uri=True)
    connection.row_factory = sqlite3.Row
    try:
        row = connection.execute(
            """
            SELECT e.*, s.public_study_id, s.verification_status,
                   w.source_path, w.file_name, w.request_id
            FROM history_evidence e
            JOIN history_studies s ON s.study_id=e.study_id
            JOIN history_workbooks w ON w.workbook_id=s.workbook_id
            WHERE e.public_evidence_id=?
            """,
            (evidence_id,),
        ).fetchone()
        if row is None:
            raise TableFirstHistoryError(
                f"History evidence ID not found: {evidence_id}"
            )
        preview = _json(row["request_preview_json"])
        address = str(row["range_address"])
        min_row, min_column, max_row, max_column = _range_bounds(address)
        cells: list[dict[str, Any]] = []
        for preview_row in preview.get("previewRows") or []:
            for cell in preview_row.get("cells") or []:
                coordinate = str(cell.get("coordinate") or "").upper()
                match = CELL_PATTERN.fullmatch(coordinate)
                if not match:
                    continue
                column, row_number = match.groups()
                value = cell.get("value")
                cells.append(
                    {
                        "row": int(row_number),
                        "column": _column_number(column),
                        "coordinate": coordinate,
                        "displayValue": "" if value is None else str(value),
                        "rawValue": "" if value is None else str(value),
                        "cachedValue": "",
                        "formula": "",
                        "numberFormat": "",
                        "mergeRange": "",
                    }
                )
        return {
            "schemaVersion": DETAIL_SCHEMA_VERSION,
            "publicEvidenceId": evidence_id,
            "trust": {
                "status": str(row["verification_status"]),
                "trusted": False,
                "reason": "Request preview only; exact source remains authoritative.",
            },
            "source": {
                "sourcePath": str(row["source_path"]),
                "fileName": str(row["file_name"]),
                "requestId": str(row["request_id"]),
            },
            "evidence": {
                "studyId": str(row["public_study_id"]),
                "tableId": str(row["table_id"]),
                "sheet": str(row["sheet"]),
                "range": address,
            },
            "preview": {
                "range": {
                    "start": {"row": min_row, "column": min_column},
                    "end": {"row": max_row, "column": max_column},
                },
                "capturedCellCountInRange": len(cells),
                "cells": cells,
                "mergedRanges": [],
                "rowDimensions": [],
                "columnDimensions": [],
            },
        }
    finally:
        connection.close()
