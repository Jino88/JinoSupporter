from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import sqlite3
import subprocess
import tempfile
from pathlib import Path
from typing import Any, Callable, Sequence

from inference_data_ai_table_first_history import build_history_pack


REQUEST_SCHEMA_VERSION = "table-first-relevance-request-v1"
AI_SCHEMA_VERSION = "table-first-relevance-ai-v1"
RESULT_SCHEMA_VERSION = "table-first-relevance-result-v1"
PROMPT_VERSION = "table-first-relevance-prompt-v1"

CELL_PATTERN = re.compile(r"^([A-Z]+)([1-9]\d*)$")
RANGE_PATTERN = re.compile(
    r"^([A-Z]+)([1-9]\d*):([A-Z]+)([1-9]\d*)$"
)
RESULT_HEADER_PATTERN = re.compile(
    r"(?:q'?ty|qty|quantity|input|ok|ng|rate|ratio|ppm|yield|noise|touch|"
    r"hearing|sigma|total|count|amount|수량|불량|검사|투입|양품|비율)",
    re.IGNORECASE,
)


class RelevanceQueryError(RuntimeError):
    pass


RunCommand = Callable[..., subprocess.CompletedProcess[str]]


def relevance_json_bytes(value: object) -> bytes:
    return json.dumps(value, ensure_ascii=False, indent=2).encode("utf-8") + b"\n"


def _canonical_bytes(value: object) -> bytes:
    return json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")


def _bounded_strings(values: Any, *, limit: int, length: int) -> list[str]:
    if not isinstance(values, list):
        return []
    result: list[str] = []
    for value in values:
        text = str(value or "").strip()
        if text:
            result.append(text[:length])
        if len(result) >= limit:
            break
    return result


def _compact_groups(values: Any) -> list[dict[str, str]]:
    if not isinstance(values, list):
        return []
    return [
        {
            "label": str(value.get("label") or "")[:240],
            "role": str(value.get("role") or "UNASSESSED")[:40],
        }
        for value in values[:12]
        if isinstance(value, dict)
    ]


def _compact_metrics(values: Any) -> list[dict[str, str]]:
    if not isinstance(values, list):
        return []
    return [
        {
            "name": str(value.get("name") or "")[:240],
            "unit": str(value.get("unit") or "")[:80],
        }
        for value in values[:24]
        if isinstance(value, dict)
    ]


def _compact_relations(values: Any) -> list[dict[str, str]]:
    if not isinstance(values, list):
        return []
    return [
        {
            "leftGroup": str(value.get("leftGroup") or "")[:240],
            "rightGroup": str(value.get("rightGroup") or "")[:240],
        }
        for value in values[:12]
        if isinstance(value, dict)
    ]


def build_relevance_query_request(
    database_path: str | Path,
    question: str,
    *,
    candidate_limit: int = 200,
) -> dict[str, Any]:
    database = Path(database_path).expanduser().resolve()
    clean_question = str(question or "").strip()
    if not clean_question:
        raise RelevanceQueryError("질문이 비어 있습니다.")
    if not database.is_file():
        raise RelevanceQueryError(f"전체 이력 DB를 찾을 수 없습니다: {database}")
    bounded_limit = max(5, min(int(candidate_limit), 200))
    pack = build_history_pack(database, clean_question, limit=bounded_limit)
    candidates: list[dict[str, Any]] = []
    candidate_ids: set[str] = set()
    for rank, study in enumerate(pack.get("studies") or [], start=1):
        study_id = str(study.get("studyId") or "")
        if not study_id:
            continue
        candidate_ids.add(study_id)
        candidates.append(
            {
                "retrievalRank": rank,
                "studyId": study_id,
                "workbookId": str(study.get("workbookId") or ""),
                "date": str(study.get("date") or ""),
                "fileName": str(study.get("fileName") or "")[:500],
                "sourcePath": str(study.get("sourcePath") or "")[:1200],
                "workbookSummary": str(
                    study.get("workbookSummary") or ""
                )[:900],
                "studyGroup": str(study.get("studyGroup") or "")[:400],
                "titles": _bounded_strings(
                    study.get("titles"), limit=8, length=320
                ),
                "groups": _compact_groups(study.get("groups")),
                "metrics": _compact_metrics(study.get("metrics")),
                "comparisonRelations": _compact_relations(
                    study.get("comparisonRelations")
                ),
                "limitations": _bounded_strings(
                    study.get("limitations"), limit=4, length=500
                ),
                "matchedQueryTerms": _bounded_strings(
                    study.get("matchedQueryTerms"), limit=20, length=100
                ),
                "retrievalScore": study.get("score"),
                "verificationStatus": str(
                    study.get("verificationStatus") or "NEEDS_REVIEW"
                ),
                "evidenceIds": _bounded_strings(
                    study.get("citationIds"), limit=100, length=80
                ),
            }
        )
    evidence_registry = [
        {
            "evidenceId": str(item.get("evidenceId") or ""),
            "studyId": str(item.get("studyId") or ""),
            "sourcePath": str(item.get("sourcePath") or ""),
            "sheet": str(item.get("sheet") or ""),
            "range": str(item.get("range") or ""),
            "tableId": str(item.get("tableId") or ""),
            "verificationStatus": str(
                item.get("verificationStatus") or "NEEDS_REVIEW"
            ),
        }
        for item in pack.get("citations") or []
        if str(item.get("studyId") or "") in candidate_ids
    ]
    request = {
        "schemaVersion": REQUEST_SCHEMA_VERSION,
        "promptVersion": PROMPT_VERSION,
        "question": clean_question,
        "database": str(database),
        "retrieval": {
            "candidateStudyCount": len(candidates),
            "candidateWorkbookCount": len(
                {item["workbookId"] for item in candidates}
            ),
            "indexedStudyCount": int(
                (pack.get("summary") or {}).get("totalIndexedStudyCount") or 0
            ),
            "candidateLimit": bounded_limit,
            "warning": (
                "후보는 DB 검색 결과이며 AI는 질문에 필요한 문서인지 여부만 판정합니다."
            ),
        },
        "candidates": candidates,
        "evidenceRegistry": evidence_registry,
    }
    request["requestSha256"] = hashlib.sha256(
        _canonical_bytes(request)
    ).hexdigest()
    return request


def relevance_output_schema() -> dict[str, Any]:
    string_array = {
        "type": "array",
        "maxItems": 20,
        "items": {"type": "string", "maxLength": 240},
    }
    return {
        "type": "object",
        "additionalProperties": False,
        "required": [
            "schemaVersion",
            "promptVersion",
            "question",
            "queryInterpretation",
            "selectedStudies",
        ],
        "properties": {
            "schemaVersion": {
                "type": "string",
                "enum": [AI_SCHEMA_VERSION],
            },
            "promptVersion": {
                "type": "string",
                "enum": [PROMPT_VERSION],
            },
            "question": {"type": "string"},
            "queryInterpretation": {
                "type": "object",
                "additionalProperties": False,
                "required": [
                    "documentNeed",
                    "subjects",
                    "conditions",
                    "metrics",
                ],
                "properties": {
                    "documentNeed": {"type": "string", "maxLength": 600},
                    "subjects": string_array,
                    "conditions": string_array,
                    "metrics": string_array,
                },
            },
            "selectedStudies": {
                "type": "array",
                "maxItems": 200,
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "required": [
                        "studyId",
                        "relevanceReason",
                        "matchedAspects",
                    ],
                    "properties": {
                        "studyId": {"type": "string"},
                        "relevanceReason": {
                            "type": "string",
                            "maxLength": 500,
                        },
                        "matchedAspects": string_array,
                    },
                },
            },
        },
    }


def build_relevance_prompt(request: dict[str, Any]) -> str:
    if request.get("schemaVersion") != REQUEST_SCHEMA_VERSION:
        raise RelevanceQueryError("지원하지 않는 관련성 요청 스키마입니다.")
    return f"""당신은 Excel 시험 보고서 검색을 돕는 문서 관련성 판정자입니다.
사용자 질문을 이해하고 candidates 중 질문에 필요한 Study를 모두 선택하십시오.

당신의 역할은 오직 문서 관련성 판정입니다.
- 보고서의 수치가 좋다/나쁘다, 증가/감소했다, 효과가 있다/없다를 판단하지 마십시오.
- 원인, 영향 방향, 우열, 개선 여부, 결론을 만들거나 사용자 질문에 답하지 마십시오.
- 수치를 계산·비교·요약하지 마십시오.
- 질문의 대상, 조건, 부품, 공정, 지표가 Study의 목적·그룹·지표·비교 구조에 실제로 연결되면 선택하십시오.
- 파일명에 단어만 우연히 겹치고 Study 내용이 질문과 연결되지 않으면 제외하십시오.
- 관련 Study가 많더라도 임의로 상위 몇 개만 고르지 말고 후보 전체를 판정해 빠짐없이 반환하십시오.
- relevanceReason은 왜 이 문서가 질문 확인에 필요한지만 한 문장으로 쓰고 결과 해석은 쓰지 마십시오.
- matchedAspects에는 질문과 연결된 대상·조건·지표 이름만 짧게 쓰십시오.
- 모든 설명은 한국어로 작성하고 JSON 하나만 반환하십시오.

schemaVersion은 {AI_SCHEMA_VERSION}, promptVersion은 {PROMPT_VERSION}, question은 입력 질문과 정확히 같아야 합니다.

REQUEST_JSON:
{json.dumps(request, ensure_ascii=False, separators=(",", ":"))}
"""


def _require_string(value: Any, path: str, *, allow_empty: bool = False) -> str:
    if not isinstance(value, str):
        raise RelevanceQueryError(f"{path}는 문자열이어야 합니다.")
    text = value.strip()
    if not allow_empty and not text:
        raise RelevanceQueryError(f"{path}는 비어 있을 수 없습니다.")
    return text


def _require_string_list(value: Any, path: str, *, limit: int) -> list[str]:
    if not isinstance(value, list) or len(value) > limit:
        raise RelevanceQueryError(f"{path}는 최대 {limit}개의 문자열 배열이어야 합니다.")
    result = [
        _require_string(item, f"{path}[{index}]")
        for index, item in enumerate(value)
    ]
    if len(result) != len(set(result)):
        raise RelevanceQueryError(f"{path}에는 중복 값을 둘 수 없습니다.")
    return result


def validate_relevance_ai_response(
    response: dict[str, Any], request: dict[str, Any]
) -> dict[str, Any]:
    if response.get("schemaVersion") != AI_SCHEMA_VERSION:
        raise RelevanceQueryError("AI 응답 스키마 버전이 올바르지 않습니다.")
    if response.get("promptVersion") != PROMPT_VERSION:
        raise RelevanceQueryError("AI 응답 프롬프트 버전이 올바르지 않습니다.")
    if response.get("question") != request.get("question"):
        raise RelevanceQueryError("AI 응답 질문이 요청과 일치하지 않습니다.")
    interpretation = response.get("queryInterpretation")
    if not isinstance(interpretation, dict):
        raise RelevanceQueryError("queryInterpretation은 객체여야 합니다.")
    _require_string(interpretation.get("documentNeed"), "queryInterpretation.documentNeed")
    for key in ("subjects", "conditions", "metrics"):
        _require_string_list(
            interpretation.get(key), f"queryInterpretation.{key}", limit=20
        )
    candidate_ids = {
        str(item.get("studyId") or "")
        for item in request.get("candidates") or []
    }
    selections = response.get("selectedStudies")
    if not isinstance(selections, list) or len(selections) > len(candidate_ids):
        raise RelevanceQueryError("selectedStudies의 크기가 후보 범위를 벗어났습니다.")
    seen: set[str] = set()
    for index, selection in enumerate(selections):
        if not isinstance(selection, dict):
            raise RelevanceQueryError(f"selectedStudies[{index}]는 객체여야 합니다.")
        study_id = _require_string(
            selection.get("studyId"), f"selectedStudies[{index}].studyId"
        )
        if study_id not in candidate_ids:
            raise RelevanceQueryError(f"후보에 없는 Study ID입니다: {study_id}")
        if study_id in seen:
            raise RelevanceQueryError(f"중복 Study ID입니다: {study_id}")
        seen.add(study_id)
        reason = _require_string(
            selection.get("relevanceReason"),
            f"selectedStudies[{index}].relevanceReason",
        )
        if len(reason) > 500:
            raise RelevanceQueryError("relevanceReason은 500자를 넘을 수 없습니다.")
        _require_string_list(
            selection.get("matchedAspects"),
            f"selectedStudies[{index}].matchedAspects",
            limit=20,
        )
    return response


def render_relevance_result_markdown(result: dict[str, Any]) -> str:
    coverage = result.get("coverage") or {}
    lines = [
        f"# {result['question']}",
        "",
        "AI는 질문에 필요한 문서인지 여부만 판정했습니다. 수치 결과, 효과, 원인, 우열은 판단하지 않았습니다.",
        "",
        f"- DB 후보 Study: {coverage.get('candidateStudyCount', 0)}건",
        f"- 관련 Study: {coverage.get('relevantStudyCount', 0)}건",
        f"- 원본 근거 범위: {coverage.get('citationCount', 0)}건",
        "",
        "| 날짜 | 원본 문서 | Study | 관련성 |",
        "|---|---|---|---|",
    ]
    for item in result.get("studies") or []:
        values = [
            str(item.get("date") or "미상"),
            str(item.get("fileName") or "").replace("|", "\\|"),
            str(item.get("studyGroup") or "").replace("|", "\\|"),
            str(item.get("relevanceReason") or "").replace("|", "\\|"),
        ]
        lines.append("| " + " | ".join(values) + " |")
    return "\n".join(lines).rstrip() + "\n"


def _format_captured_value(
    cell: dict[str, Any], display_samples: dict[str, str]
) -> str:
    coordinate = str(cell.get("coordinate") or "").upper()
    if coordinate in display_samples:
        return display_samples[coordinate]
    raw_value = cell.get("value")
    number_format = str(cell.get("numberFormat") or "")
    if "%" in number_format:
        try:
            number = float(raw_value)
        except (TypeError, ValueError):
            return str(raw_value or "")
        decimal_pattern = number_format.split("%", 1)[0]
        decimals = (
            len(decimal_pattern.rsplit(".", 1)[1])
            if "." in decimal_pattern
            else 0
        )
        return f"{number * 100:.{decimals}f}%"
    return str(raw_value if raw_value is not None else "")


def _raw_data_for_studies(
    request: dict[str, Any], studies: list[dict[str, Any]]
) -> None:
    database = Path(str(request.get("database") or ""))
    if not database.is_file() or not studies:
        for study in studies:
            study["rawDataPoints"] = []
            study["rawDataPointCount"] = 0
            study["rawDataTruncated"] = False
        return
    study_ids = [str(study.get("studyId") or "") for study in studies]
    connection = sqlite3.connect(
        f"file:{database.resolve().as_posix()}?mode=ro", uri=True
    )
    connection.row_factory = sqlite3.Row
    try:
        placeholders = ",".join("?" for _ in study_ids)
        rows = connection.execute(
            f"""
            SELECT s.public_study_id, s.payload_json, w.request_path
            FROM history_studies s
            JOIN history_workbooks w ON w.workbook_id=s.workbook_id
            WHERE s.public_study_id IN ({placeholders})
            """,
            study_ids,
        ).fetchall()
    finally:
        connection.close()
    source_by_study = {str(row["public_study_id"]): row for row in rows}
    evidence_registry = {
        str(item.get("evidenceId") or ""): item
        for item in request.get("evidenceRegistry") or []
    }
    request_cache: dict[str, dict[str, Any]] = {}
    max_points_per_study = 240
    for study in studies:
        study_id = str(study.get("studyId") or "")
        source = source_by_study.get(study_id)
        points: list[dict[str, Any]] = []
        truncated = False
        if source is None:
            study["rawDataPoints"] = points
            study["rawDataPointCount"] = 0
            study["rawDataTruncated"] = False
            continue
        payload = json.loads(str(source["payload_json"]))
        metric_by_axis: dict[str, tuple[str, str]] = {}
        for metric in payload.get("metrics") or []:
            if not isinstance(metric, dict):
                continue
            name = str(metric.get("name") or "")
            unit = str(metric.get("unit") or "")
            for axis_ref in metric.get("axisRefs") or []:
                metric_by_axis[str(axis_ref)] = (name, unit)
        request_path = str(source["request_path"])
        captured_request = request_cache.get(request_path)
        if captured_request is None:
            try:
                captured_request = json.loads(
                    Path(request_path).read_text(encoding="utf-8")
                )
            except (OSError, json.JSONDecodeError):
                captured_request = {}
            request_cache[request_path] = captured_request
        table_by_id = {
            str(table.get("tableId") or ""): table
            for table in captured_request.get("tables") or []
            if isinstance(table, dict)
        }
        evidence_by_table: dict[str, str] = {}
        for evidence_id in study.get("evidenceIds") or []:
            evidence = evidence_registry.get(str(evidence_id))
            if evidence is None:
                continue
            evidence_by_table.setdefault(
                str(evidence.get("tableId") or ""), str(evidence_id)
            )
        for table_id, evidence_id in evidence_by_table.items():
            table = table_by_id.get(table_id)
            if table is None:
                continue
            included_columns: dict[str, dict[str, str]] = {}
            display_samples: dict[str, str] = {}
            for column in table.get("numericColumns") or []:
                if not isinstance(column, dict):
                    continue
                column_id = str(column.get("columnId") or "")
                header = " / ".join(
                    _bounded_strings(
                        column.get("headerTexts"), limit=8, length=160
                    )
                )
                role = str(column.get("columnRole") or "")
                metric = metric_by_axis.get(column_id)
                if metric is None and role != "MEASURE_VALUE" and not RESULT_HEADER_PATTERN.search(header):
                    continue
                metric_name = metric[0] if metric and metric[0] else header
                metric_unit = metric[1] if metric else ""
                column_label = str(column.get("column") or "").upper()
                if not column_label:
                    continue
                included_columns[column_label] = {
                    "metric": metric_name or header or column_label,
                    "unit": metric_unit,
                    "columnRole": role,
                }
                for sample in column.get("displaySamples") or []:
                    if not isinstance(sample, dict):
                        continue
                    coordinate = str(sample.get("coordinate") or "").upper()
                    display = str(
                        sample.get("normalizedDisplay")
                        or sample.get("sourceDisplay")
                        or ""
                    ).strip()
                    if coordinate and display:
                        display_samples[coordinate] = display
            context_by_row: dict[int, list[str]] = {}
            for preview_row in table.get("previewRows") or []:
                if not isinstance(preview_row, dict):
                    continue
                row_number = int(preview_row.get("row") or 0)
                for cell in preview_row.get("cells") or []:
                    if not isinstance(cell, dict):
                        continue
                    kind = str(cell.get("kind") or "").upper()
                    value = str(cell.get("value") or "").strip()
                    coordinate = str(cell.get("coordinate") or "").upper()
                    if kind not in {"TEXT", "DATE"} or not value or not coordinate:
                        continue
                    label = f"{coordinate}={value}"
                    target_rows = [row_number]
                    merge_match = RANGE_PATTERN.fullmatch(
                        str(cell.get("mergeRange") or "").upper()
                    )
                    if merge_match:
                        target_rows = list(
                            range(
                                int(merge_match.group(2)),
                                int(merge_match.group(4)) + 1,
                            )
                        )
                    for target_row in target_rows:
                        values = context_by_row.setdefault(target_row, [])
                        if label not in values:
                            values.append(label)
            for preview_row in table.get("previewRows") or []:
                if not isinstance(preview_row, dict):
                    continue
                row_number = int(preview_row.get("row") or 0)
                row_context = " | ".join(context_by_row.get(row_number) or [])
                for cell in preview_row.get("cells") or []:
                    if not isinstance(cell, dict):
                        continue
                    coordinate = str(cell.get("coordinate") or "").upper()
                    coordinate_match = CELL_PATTERN.fullmatch(coordinate)
                    if (
                        str(cell.get("kind") or "").upper() != "NUMBER"
                        or coordinate_match is None
                    ):
                        continue
                    column_info = included_columns.get(coordinate_match.group(1))
                    if column_info is None:
                        continue
                    points.append(
                        {
                            "evidenceId": evidence_id,
                            "tableId": table_id,
                            "sheet": str(table.get("sheet") or ""),
                            "range": str(table.get("range") or ""),
                            "row": row_number,
                            "condition": row_context or f"Row {row_number}",
                            "metric": column_info["metric"],
                            "unit": column_info["unit"],
                            "displayValue": _format_captured_value(
                                cell, display_samples
                            ),
                            "coordinate": coordinate,
                        }
                    )
                    if len(points) >= max_points_per_study:
                        truncated = True
                        break
                if truncated:
                    break
            if truncated:
                break
        study["rawDataPoints"] = points
        study["rawDataPointCount"] = len(points)
        study["rawDataTruncated"] = truncated


def build_relevance_result(
    response: dict[str, Any], request: dict[str, Any]
) -> dict[str, Any]:
    validated = validate_relevance_ai_response(response, request)
    candidates = {
        str(item["studyId"]): item for item in request.get("candidates") or []
    }
    evidence = {
        str(item["evidenceId"]): item
        for item in request.get("evidenceRegistry") or []
    }
    studies: list[dict[str, Any]] = []
    selected_ids: list[str] = []
    for selection in validated["selectedStudies"]:
        study_id = str(selection["studyId"])
        selected_ids.append(study_id)
        item = dict(candidates[study_id])
        item["relevanceReason"] = str(selection["relevanceReason"])
        item["matchedAspects"] = list(selection["matchedAspects"])
        studies.append(item)
    _raw_data_for_studies(request, studies)
    citation_ids: list[str] = []
    citation_seen: set[str] = set()
    for study in studies:
        for raw_evidence_id in study.get("evidenceIds") or []:
            evidence_id = str(raw_evidence_id)
            if evidence_id not in evidence or evidence_id in citation_seen:
                continue
            citation_seen.add(evidence_id)
            citation_ids.append(evidence_id)
    result = {
        "schemaVersion": RESULT_SCHEMA_VERSION,
        "promptVersion": PROMPT_VERSION,
        "question": str(request["question"]),
        "requestSha256": str(request["requestSha256"]),
        "answerStatus": (
            "RELEVANT_DOCUMENTS_FOUND" if studies else "NO_RELEVANT_DOCUMENTS"
        ),
        "queryInterpretation": validated["queryInterpretation"],
        "coverage": {
            "candidateStudyCount": len(request.get("candidates") or []),
            "relevantStudyCount": len(studies),
            "relevantWorkbookCount": len(
                {str(item.get("workbookId") or "") for item in studies}
            ),
            "citationCount": len(citation_ids),
            "rawDataPointCount": sum(
                int(item.get("rawDataPointCount") or 0) for item in studies
            ),
            "studiesWithRawData": sum(
                bool(item.get("rawDataPoints")) for item in studies
            ),
            "eligibleEffectCount": 0,
            "indexedStudyCount": int(
                (request.get("retrieval") or {}).get("indexedStudyCount") or 0
            ),
        },
        "studies": studies,
        "citations": [evidence[value] for value in citation_ids],
        "limitations": [
            "AI는 문서 관련성만 판정했으며 보고서 결과를 해석하지 않았습니다.",
            "표시된 수치와 결론은 원본 Excel을 열어 사람이 확인해야 합니다.",
        ],
    }
    result["markdown"] = render_relevance_result_markdown(result)
    return result


def _codex_command(command: Sequence[str] | None) -> list[str]:
    if command:
        return list(command)
    executable = shutil.which("codex.cmd" if os.name == "nt" else "codex")
    if not executable:
        raise RelevanceQueryError("Codex CLI 실행 파일을 PATH에서 찾지 못했습니다.")
    return [executable]


def run_codex_relevance_query(
    *,
    request: dict[str, Any],
    output_path: str | Path,
    model: str | None = None,
    reasoning_effort: str | None = "medium",
    codex_command: Sequence[str] | None = None,
    timeout_seconds: int = 600,
    run_command: RunCommand = subprocess.run,
    max_validation_attempts: int = 2,
) -> dict[str, Any]:
    if max_validation_attempts < 1:
        raise ValueError("max_validation_attempts는 1 이상이어야 합니다.")
    target = Path(output_path).expanduser().resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    prompt = build_relevance_prompt(request)
    validation_error: RelevanceQueryError | None = None
    with tempfile.TemporaryDirectory(prefix="relevance-document-query-") as temp_dir:
        schema_path = Path(temp_dir) / "relevance-query.schema.json"
        response_path = Path(temp_dir) / "last-message.json"
        schema_path.write_text(
            json.dumps(relevance_output_schema(), ensure_ascii=False),
            encoding="utf-8",
        )
        for attempt in range(1, max_validation_attempts + 1):
            response_path.unlink(missing_ok=True)
            command = [
                *_codex_command(codex_command),
                "exec",
                "--ephemeral",
                "--skip-git-repo-check",
                "--sandbox",
                "read-only",
                "--output-schema",
                str(schema_path),
                "--output-last-message",
                str(response_path),
            ]
            if reasoning_effort:
                command.extend(["-c", f'model_reasoning_effort="{reasoning_effort}"'])
            if model:
                command.extend(["--model", model])
            command.append("-")
            retry_note = ""
            if validation_error is not None:
                retry_note = (
                    "\n이전 문서 관련성 응답이 검증에서 거절됐습니다.\n"
                    f"거절 사유: {validation_error}\n"
                    "후보에 있는 Study ID만 중복 없이 반환하십시오. 결과 해석은 하지 마십시오.\n"
                )
            completed = run_command(
                command,
                input=prompt + retry_note,
                text=True,
                encoding="utf-8",
                errors="replace",
                capture_output=True,
                timeout=timeout_seconds,
                check=False,
                cwd=temp_dir,
            )
            if completed.returncode != 0:
                detail = (completed.stderr or completed.stdout or "").strip()
                raise RelevanceQueryError(
                    "문서 관련성 AI 호출이 실패했습니다 "
                    f"(exit {completed.returncode}): {detail[-2000:]}"
                )
            if not response_path.is_file():
                raise RelevanceQueryError("문서 관련성 AI가 최종 JSON을 생성하지 않았습니다.")
            try:
                response = json.loads(response_path.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError) as exc:
                raise RelevanceQueryError("문서 관련성 AI JSON을 읽을 수 없습니다.") from exc
            if not isinstance(response, dict):
                raise RelevanceQueryError("문서 관련성 AI 응답은 객체여야 합니다.")
            try:
                result = build_relevance_result(response, request)
            except RelevanceQueryError as error:
                validation_error = error
                rejected_path = target.with_name(
                    f"{target.stem}.ai-response.attempt-{attempt}.json"
                )
                rejected_path.write_bytes(relevance_json_bytes(response))
                continue
            raw_path = target.with_name(f"{target.stem}.ai-response.json")
            raw_path.write_bytes(relevance_json_bytes(response))
            target.write_bytes(relevance_json_bytes(result))
            return result
    raise RelevanceQueryError(
        "문서 관련성 AI 응답이 반복 검증에 실패했습니다: "
        + str(validation_error or "unknown validation error")
    )


__all__ = [
    "AI_SCHEMA_VERSION",
    "PROMPT_VERSION",
    "REQUEST_SCHEMA_VERSION",
    "RESULT_SCHEMA_VERSION",
    "RelevanceQueryError",
    "build_relevance_prompt",
    "build_relevance_query_request",
    "build_relevance_result",
    "relevance_json_bytes",
    "render_relevance_result_markdown",
    "run_codex_relevance_query",
    "validate_relevance_ai_response",
]
