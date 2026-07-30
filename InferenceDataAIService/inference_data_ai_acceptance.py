"""Deterministic golden-question acceptance reporting for the pilot corpus.

This module deliberately limits automation to structural acceptance:

* every manifest question is executed through the canonical query and answer
  builders;
* the deterministic answer is validated against its evidence pack;
* every primary workbook is matched by its exact source path, never by a
  filename or a fuzzy title;
* missing ingestion is distinguished from a retrieval miss by inspecting the
  current canonical workbook analyses through a read-only SQLite connection.

Natural-language ``requiredBehavior`` entries remain manual-review items unless
the manifest supplies an aligned declarative assertion.  Assertions inspect
only canonical pack/answer structures; they never ask an AI to approve its own
semantic output.
"""

from __future__ import annotations

import json
import os
import re
import sqlite3
import tempfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

from inference_data_ai_answer import (
    build_evidence_answer,
    validate_evidence_answer,
)
from inference_data_ai_query import (
    build_evidence_pack_from_db,
    connect_knowledge_readonly,
)


ACCEPTANCE_SCHEMA_VERSION = "canonical-golden-acceptance-v1"
SAFE_ID_PATTERN = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]*$")

PackBuilder = Callable[[str | Path, str], dict[str, Any]]
AnswerBuilder = Callable[[dict[str, Any]], dict[str, Any]]
AnswerValidator = Callable[
    [dict[str, Any], dict[str, Any]],
    dict[str, Any],
]


class AcceptanceReportError(RuntimeError):
    """Raised when the acceptance inputs or canonical DB are malformed."""


def _canonical_json_bytes(value: Any) -> bytes:
    return (
        json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            indent=2,
            allow_nan=False,
        )
        + "\n"
    ).encode("utf-8")


def _atomic_write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_name: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb",
            prefix=f".{path.name}.",
            suffix=".tmp",
            dir=path.parent,
            delete=False,
        ) as handle:
            temporary_name = handle.name
            handle.write(_canonical_json_bytes(value))
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary_name, path)
        temporary_name = None
    finally:
        if temporary_name is not None:
            Path(temporary_name).unlink(missing_ok=True)


def _load_manifest(manifest_path: str | Path) -> dict[str, Any]:
    path = Path(manifest_path).expanduser().resolve()
    try:
        manifest = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise AcceptanceReportError(
            f"Cannot read representative pilot manifest: {path}"
        ) from error
    if not isinstance(manifest, dict):
        raise AcceptanceReportError("Pilot manifest must be a JSON object.")
    if manifest.get("schemaVersion") != "representative-pilot-v1":
        raise AcceptanceReportError(
            "Pilot manifest schemaVersion must be representative-pilot-v1."
        )
    if not isinstance(manifest.get("sourceRoot"), str) or not str(
        manifest["sourceRoot"]
    ).strip():
        raise AcceptanceReportError("Pilot manifest sourceRoot is required.")
    if not isinstance(manifest.get("workbooks"), list):
        raise AcceptanceReportError("Pilot manifest workbooks must be a list.")
    if not isinstance(manifest.get("goldenQuestions"), list):
        raise AcceptanceReportError(
            "Pilot manifest goldenQuestions must be a list."
        )
    return manifest


def _validated_id(value: object, kind: str) -> str:
    identifier = str(value or "").strip()
    if (
        not identifier
        or not SAFE_ID_PATTERN.fullmatch(identifier)
        or identifier in {".", ".."}
    ):
        raise AcceptanceReportError(f"Invalid {kind} id: {identifier!r}.")
    return identifier


def _exact_path_key(value: str | Path) -> str:
    resolved = Path(value).expanduser().resolve(strict=False)
    return os.path.normcase(os.path.normpath(str(resolved)))


def _workbook_map(manifest: dict[str, Any]) -> dict[str, dict[str, str]]:
    source_root = Path(str(manifest["sourceRoot"])).expanduser().resolve(
        strict=False
    )
    workbooks: dict[str, dict[str, str]] = {}
    for raw in manifest["workbooks"]:
        if not isinstance(raw, dict):
            raise AcceptanceReportError(
                "Every pilot workbook entry must be a JSON object."
            )
        pilot_id = _validated_id(raw.get("id"), "pilot workbook")
        if pilot_id in workbooks:
            raise AcceptanceReportError(
                f"Duplicate pilot workbook id: {pilot_id}."
            )
        relative_path = str(raw.get("relativePath") or "").strip()
        if not relative_path:
            raise AcceptanceReportError(
                f"Pilot workbook {pilot_id} has no relativePath."
            )
        source_path = (source_root / relative_path).resolve(strict=False)
        try:
            source_path.relative_to(source_root)
        except ValueError as error:
            raise AcceptanceReportError(
                f"Pilot workbook {pilot_id} escapes sourceRoot."
            ) from error
        workbooks[pilot_id] = {
            "pilotId": pilot_id,
            "relativePath": relative_path,
            "expectedSourcePath": str(source_path),
        }
    return workbooks


def _current_analysis_index(
    database_path: str | Path,
) -> dict[str, list[dict[str, Any]]]:
    try:
        with connect_knowledge_readonly(database_path) as connection:
            rows = connection.execute(
                """
                SELECT
                    d.source_path,
                    d.original_file_name,
                    r.revision_uid,
                    r.content_sha256,
                    a.analysis_uid,
                    a.public_analysis_id,
                    a.analysis_status,
                    a.verification_status,
                    s.public_data_id
                FROM source_documents d
                JOIN source_revisions r
                  ON r.document_id=d.document_id
                 AND r.is_current=1
                JOIN workbook_analyses a
                  ON a.document_id=d.document_id
                 AND a.revision_id=r.revision_id
                LEFT JOIN knowledge_studies s
                  ON s.workbook_analysis_id=a.workbook_analysis_id
                ORDER BY
                    d.source_path,
                    a.public_analysis_id,
                    s.public_data_id
                """
            ).fetchall()
    except (FileNotFoundError, sqlite3.Error) as error:
        raise AcceptanceReportError(
            "Cannot inspect current canonical workbook analyses."
        ) from error

    grouped: dict[
        tuple[str, str, str, str, str, str, str, str],
        set[str],
    ] = {}
    for row in rows:
        group_key = tuple(str(value or "") for value in row[:8])
        grouped.setdefault(group_key, set())
        if row[8]:
            grouped[group_key].add(str(row[8]))

    index: dict[str, list[dict[str, Any]]] = {}
    for group_key, public_data_ids in grouped.items():
        (
            source_path,
            file_name,
            revision_uid,
            content_sha256,
            analysis_uid,
            public_analysis_id,
            analysis_status,
            verification_status,
        ) = group_key
        index.setdefault(_exact_path_key(source_path), []).append(
            {
                "sourcePath": source_path,
                "fileName": file_name,
                "revisionUid": revision_uid,
                "contentSha256": content_sha256,
                "analysisUid": analysis_uid,
                "publicAnalysisId": public_analysis_id,
                "analysisStatus": analysis_status,
                "verificationStatus": verification_status,
                "publicDataIds": sorted(public_data_ids),
            }
        )
    for analyses in index.values():
        analyses.sort(
            key=lambda item: (
                item["publicAnalysisId"],
                item["revisionUid"],
            )
        )
    return index


def _pack_source_representations(
    pack: dict[str, Any],
) -> dict[str, dict[str, dict[str, list[str]]]]:
    represented: dict[str, dict[str, dict[str, list[str]]]] = {}
    for candidate in pack.get("studyCandidates", []):
        if not isinstance(candidate, dict):
            continue
        source = candidate.get("source")
        if not isinstance(source, dict) or not source.get("sourcePath"):
            continue
        key = _exact_path_key(str(source["sourcePath"]))
        item = represented.setdefault(
            key,
            {
                "studyCandidates": {
                    "publicDataIds": [],
                    "publicAnalysisIds": [],
                },
                "sourceExclusions": {"publicAnalysisIds": []},
            },
        )
        public_data_id = str(candidate.get("publicDataId") or "")
        if public_data_id:
            item["studyCandidates"]["publicDataIds"].append(
                public_data_id
            )
        analysis = candidate.get("analysis")
        if isinstance(analysis, dict):
            public_analysis_id = str(
                analysis.get("publicAnalysisId") or ""
            )
            if public_analysis_id:
                item["studyCandidates"]["publicAnalysisIds"].append(
                    public_analysis_id
                )

    for exclusion in pack.get("sourceExclusions", []):
        if not isinstance(exclusion, dict) or not exclusion.get("sourcePath"):
            continue
        key = _exact_path_key(str(exclusion["sourcePath"]))
        item = represented.setdefault(
            key,
            {
                "studyCandidates": {
                    "publicDataIds": [],
                    "publicAnalysisIds": [],
                },
                "sourceExclusions": {"publicAnalysisIds": []},
            },
        )
        public_analysis_id = str(
            exclusion.get("publicAnalysisId") or ""
        )
        if public_analysis_id:
            item["sourceExclusions"]["publicAnalysisIds"].append(
                public_analysis_id
            )

    for item in represented.values():
        item["studyCandidates"]["publicDataIds"] = sorted(
            set(item["studyCandidates"]["publicDataIds"])
        )
        item["studyCandidates"]["publicAnalysisIds"] = sorted(
            set(item["studyCandidates"]["publicAnalysisIds"])
        )
        item["sourceExclusions"]["publicAnalysisIds"] = sorted(
            set(item["sourceExclusions"]["publicAnalysisIds"])
        )
    return represented


def _primary_source_result(
    workbook: dict[str, str],
    pack_representations: dict[
        str,
        dict[str, dict[str, list[str]]],
    ],
    current_analyses: dict[str, list[dict[str, Any]]],
) -> dict[str, Any]:
    path_key = _exact_path_key(workbook["expectedSourcePath"])
    canonical = current_analyses.get(path_key, [])
    representation = pack_representations.get(path_key)
    if not canonical:
        status = "PENDING_INGEST"
    elif representation is None:
        status = "RETRIEVAL_MISS"
    else:
        status = "REPRESENTED"
    return {
        **workbook,
        "status": status,
        "representedThrough": (
            representation
            if representation is not None
            else {
                "studyCandidates": {
                    "publicDataIds": [],
                    "publicAnalysisIds": [],
                },
                "sourceExclusions": {"publicAnalysisIds": []},
            }
        ),
        "currentCanonicalAnalyses": canonical,
        "imagesAnalyzed": False,
    }


def _eligible_counts(pack: dict[str, Any]) -> dict[str, int]:
    effects = [
        item
        for item in pack.get("answerEligibleEffects", [])
        if isinstance(item, dict)
    ]
    data_ids = {
        str(item["publicDataId"])
        for item in effects
        if item.get("publicDataId")
    }
    comparison_ids = {
        str(item["publicComparisonId"])
        for item in effects
        if item.get("publicComparisonId")
    }
    evidence_ids = {
        str(evidence_id)
        for item in effects
        for evidence_id in item.get("publicEvidenceIds", [])
        if evidence_id
    }
    return {
        "eligibleEffectCount": len(effects),
        "eligibleDataCount": len(data_ids),
        "eligibleComparisonCount": len(comparison_ids),
        "eligibleEvidenceCount": len(evidence_ids),
    }


def _question_status(
    *,
    build_errors: list[dict[str, str]],
    validation_status: str,
    primary_sources: list[dict[str, Any]],
) -> str:
    statuses = {item["status"] for item in primary_sources}
    if (
        build_errors
        or validation_status != "PASS"
        or "RETRIEVAL_MISS" in statuses
    ):
        return "FAIL"
    if "PENDING_INGEST" in statuses:
        return "BLOCKED_PENDING_INGEST"
    return "PASS"


def _factor_difference_signature(item: dict[str, Any]) -> tuple[Any, ...]:
    return (
        str(item.get("factorLabel") or ""),
        str(item.get("controlValue") or ""),
        str(item.get("comparedValue") or ""),
        bool(item.get("controlValueRecorded")),
        bool(item.get("comparedValueRecorded")),
    )


def _confounded_factor_maps(
    pack: dict[str, Any],
    answer: dict[str, Any],
    minimum_factor_count: int,
) -> tuple[
    dict[tuple[str, str], set[tuple[Any, ...]]],
    dict[tuple[str, str], set[tuple[Any, ...]]],
]:
    pack_map: dict[tuple[str, str], set[tuple[Any, ...]]] = {}
    for record in pack.get("excludedCandidates", []):
        comparison = record.get("comparison") or {}
        if (
            str(comparison.get("confoundingStatus") or "").upper()
            != "CONFOUNDED"
        ):
            continue
        differences = {
            _factor_difference_signature(item)
            for item in comparison.get("factorDifferences", [])
            if isinstance(item, dict)
        }
        if len(differences) < minimum_factor_count:
            continue
        comparison_id = str(record.get("publicComparisonId") or "")
        if not comparison_id:
            continue
        pack_map[
            (str(record.get("publicDataId") or ""), comparison_id)
        ] = differences

    answer_map: dict[tuple[str, str], set[tuple[Any, ...]]] = {}
    for record in answer.get("excludedRecords", []):
        assessment = record.get("comparisonAssessment") or {}
        if (
            str(assessment.get("code") or "")
            != "CONFOUNDED_MULTI_FACTOR"
        ):
            continue
        comparison_id = str(record.get("comparisonId") or "")
        if not comparison_id:
            continue
        answer_map[
            (str(record.get("dataId") or ""), comparison_id)
        ] = {
            _factor_difference_signature(item)
            for item in assessment.get("factorDifferences", [])
            if isinstance(item, dict)
        }
    return pack_map, answer_map


def _answer_codes(answer: dict[str, Any]) -> set[str]:
    codes = {
        str(item.get("code") or "")
        for item in answer.get("limitations", [])
        if isinstance(item, dict)
    }
    codes.update(
        str(code)
        for record in answer.get("excludedRecords", [])
        if isinstance(record, dict)
        for code in record.get("reasonCodes", [])
    )
    template_code = str(
        answer.get("directAnswer", {}).get("templateCode") or ""
    )
    if template_code:
        codes.add(template_code)
    return {code for code in codes if code}


def _evaluate_behavior_assertion(
    config: dict[str, Any],
    pack: dict[str, Any] | None,
    answer: dict[str, Any] | None,
) -> dict[str, Any]:
    assertion_type = str(config.get("type") or "").upper()
    if pack is None or answer is None:
        return {
            "status": "NOT_RUN",
            "assertion": config,
            "details": {"reason": "PACK_OR_ANSWER_NOT_AVAILABLE"},
        }
    if assertion_type == "CONFOUNDED_FACTORS_COMPLETE":
        minimum_factor_count = int(config.get("minimumFactorCount", 2))
        minimum_comparison_count = int(
            config.get("minimumComparisonCount", 1)
        )
        if minimum_factor_count < 2 or minimum_comparison_count < 1:
            raise AcceptanceReportError(
                "CONFOUNDED_FACTORS_COMPLETE minima are invalid."
            )
        pack_map, answer_map = _confounded_factor_maps(
            pack,
            answer,
            minimum_factor_count,
        )
        missing = sorted(
            f"{data_id}/{comparison_id}"
            for data_id, comparison_id in set(pack_map) - set(answer_map)
        )
        mismatched = sorted(
            f"{data_id}/{comparison_id}"
            for data_id, comparison_id in set(pack_map) & set(answer_map)
            if pack_map[(data_id, comparison_id)]
            != answer_map[(data_id, comparison_id)]
        )
        passed = (
            len(pack_map) >= minimum_comparison_count
            and not missing
            and not mismatched
        )
        return {
            "status": "PASS" if passed else "FAIL",
            "assertion": config,
            "details": {
                "qualifyingPackComparisonCount": len(pack_map),
                "answerComparisonCount": len(answer_map),
                "missingComparisons": missing,
                "mismatchedComparisons": mismatched,
            },
        }
    if assertion_type == "REQUIRED_ANSWER_CODE":
        required_code = str(config.get("code") or "").strip()
        if not required_code:
            raise AcceptanceReportError(
                "REQUIRED_ANSWER_CODE requires code."
            )
        codes = _answer_codes(answer)
        return {
            "status": "PASS" if required_code in codes else "FAIL",
            "assertion": config,
            "details": {
                "requiredCode": required_code,
                "answerCodes": sorted(codes),
            },
        }
    if assertion_type == "MAX_ELIGIBLE_EFFECT_COUNT":
        maximum = int(config.get("maximum", 0))
        if maximum < 0:
            raise AcceptanceReportError(
                "MAX_ELIGIBLE_EFFECT_COUNT maximum must be non-negative."
            )
        pack_count = len(pack.get("answerEligibleEffects", []))
        answer_count = sum(
            len(group.get("effects", []))
            for group in answer.get("quantitativeGroups", [])
            if isinstance(group, dict)
        )
        return {
            "status": (
                "PASS"
                if pack_count <= maximum and answer_count <= maximum
                else "FAIL"
            ),
            "assertion": config,
            "details": {
                "maximum": maximum,
                "packEligibleEffectCount": pack_count,
                "answerQuantitativeEffectCount": answer_count,
            },
        }
    raise AcceptanceReportError(
        f"Unsupported requiredBehavior assertion type: {assertion_type}."
    )


def run_golden_question_acceptance(
    database_path: str | Path,
    manifest_path: str | Path,
    output_directory: str | Path,
    *,
    pack_builder: PackBuilder = build_evidence_pack_from_db,
    answer_builder: AnswerBuilder = build_evidence_answer,
    answer_validator: AnswerValidator = validate_evidence_answer,
) -> dict[str, Any]:
    """Execute and persist one structural acceptance result per golden question."""

    database = Path(database_path).expanduser().resolve()
    manifest_file = Path(manifest_path).expanduser().resolve()
    output = Path(output_directory).expanduser().resolve()
    manifest = _load_manifest(manifest_file)
    workbooks = _workbook_map(manifest)
    current_analyses = _current_analysis_index(database)

    questions: list[dict[str, Any]] = []
    seen_question_ids: set[str] = set()
    question_output = output / "questions"
    for raw in manifest["goldenQuestions"]:
        if not isinstance(raw, dict):
            raise AcceptanceReportError(
                "Every golden question entry must be a JSON object."
            )
        question_id = _validated_id(raw.get("id"), "golden question")
        if question_id in seen_question_ids:
            raise AcceptanceReportError(
                f"Duplicate golden question id: {question_id}."
            )
        seen_question_ids.add(question_id)
        question_text = str(raw.get("question") or "").strip()
        if not question_text:
            raise AcceptanceReportError(
                f"Golden question {question_id} has no question text."
            )
        raw_primary_ids = raw.get("primaryPilotIds")
        if not isinstance(raw_primary_ids, list) or not raw_primary_ids:
            raise AcceptanceReportError(
                f"Golden question {question_id} has no primaryPilotIds."
            )
        primary_ids = [
            _validated_id(value, "primary pilot")
            for value in raw_primary_ids
        ]
        if len(set(primary_ids)) != len(primary_ids):
            raise AcceptanceReportError(
                f"Golden question {question_id} repeats a primaryPilotId."
            )
        missing_ids = [
            pilot_id for pilot_id in primary_ids if pilot_id not in workbooks
        ]
        if missing_ids:
            raise AcceptanceReportError(
                f"Golden question {question_id} refers to unknown pilot ids: "
                + ", ".join(missing_ids)
            )
        required_behavior = raw.get("requiredBehavior", [])
        if not isinstance(required_behavior, list):
            raise AcceptanceReportError(
                f"Golden question {question_id} requiredBehavior must be a list."
            )
        raw_behavior_assertions = raw.get("requiredBehaviorAssertions")
        if raw_behavior_assertions is None:
            behavior_assertions: list[dict[str, Any] | None] = [
                None
                for _ in required_behavior
            ]
        else:
            if (
                not isinstance(raw_behavior_assertions, list)
                or len(raw_behavior_assertions) != len(required_behavior)
            ):
                raise AcceptanceReportError(
                    f"Golden question {question_id} "
                    "requiredBehaviorAssertions must align one-to-one with "
                    "requiredBehavior."
                )
            behavior_assertions = []
            for assertion in raw_behavior_assertions:
                if assertion is not None and not isinstance(assertion, dict):
                    raise AcceptanceReportError(
                        f"Golden question {question_id} behavior assertion "
                        "must be an object or null."
                    )
                behavior_assertions.append(assertion)

        pack: dict[str, Any] | None = None
        answer: dict[str, Any] | None = None
        build_errors: list[dict[str, str]] = []
        validation = {"status": "NOT_RUN", "error": None}
        pack_relative = f"questions/{question_id}.pack.json"
        answer_relative = f"questions/{question_id}.answer.json"
        try:
            pack = pack_builder(database, question_text)
            if not isinstance(pack, dict):
                raise TypeError("Pack builder did not return a JSON object.")
            _atomic_write_json(output / pack_relative, pack)
        except Exception as error:  # continue so the final report remains useful
            build_errors.append(
                {
                    "stage": "EVIDENCE_PACK",
                    "errorType": type(error).__name__,
                    "message": str(error),
                }
            )

        if pack is not None:
            try:
                answer = answer_builder(pack)
                if not isinstance(answer, dict):
                    raise TypeError(
                        "Answer builder did not return a JSON object."
                    )
                _atomic_write_json(output / answer_relative, answer)
            except Exception as error:
                build_errors.append(
                    {
                        "stage": "EVIDENCE_ANSWER",
                        "errorType": type(error).__name__,
                        "message": str(error),
                    }
                )

        if pack is not None and answer is not None:
            try:
                answer_validator(answer, pack)
                validation = {"status": "PASS", "error": None}
            except Exception as error:
                validation = {
                    "status": "FAIL",
                    "error": {
                        "errorType": type(error).__name__,
                        "message": str(error),
                    },
                }

        pack_representations = (
            _pack_source_representations(pack) if pack is not None else {}
        )
        primary_sources = [
            _primary_source_result(
                workbooks[pilot_id],
                pack_representations,
                current_analyses,
            )
            for pilot_id in primary_ids
        ]
        counts = (
            _eligible_counts(pack)
            if pack is not None
            else {
                "eligibleEffectCount": 0,
                "eligibleDataCount": 0,
                "eligibleComparisonCount": 0,
                "eligibleEvidenceCount": 0,
            }
        )
        status = _question_status(
            build_errors=build_errors,
            validation_status=str(validation["status"]),
            primary_sources=primary_sources,
        )
        behavior_results: list[dict[str, Any]] = []
        for behavior, assertion in zip(
            required_behavior,
            behavior_assertions,
            strict=True,
        ):
            if assertion is None:
                behavior_results.append(
                    {
                        "text": str(behavior),
                        "status": "MANUAL_REVIEW_REQUIRED",
                        "assertion": None,
                        "details": None,
                    }
                )
                continue
            assertion_result = _evaluate_behavior_assertion(
                assertion,
                pack,
                answer,
            )
            behavior_results.append(
                {
                    "text": str(behavior),
                    **assertion_result,
                }
            )
        if status == "PASS" and any(
            item["status"] in {"FAIL", "NOT_RUN"}
            for item in behavior_results
            if item.get("assertion") is not None
        ):
            status = "FAIL"
        questions.append(
            {
                "id": question_id,
                "question": question_text,
                "status": status,
                "primaryPilotIds": primary_ids,
                "primarySources": primary_sources,
                "sourceRepresentationCounts": {
                    "represented": sum(
                        item["status"] == "REPRESENTED"
                        for item in primary_sources
                    ),
                    "pendingIngest": sum(
                        item["status"] == "PENDING_INGEST"
                        for item in primary_sources
                    ),
                    "retrievalMiss": sum(
                        item["status"] == "RETRIEVAL_MISS"
                        for item in primary_sources
                    ),
                },
                "answerValidation": validation,
                "counts": counts,
                "requiredBehavior": behavior_results,
                "artifacts": {
                    "packJson": pack_relative if pack is not None else None,
                    "answerJson": (
                        answer_relative if answer is not None else None
                    ),
                },
                "errors": build_errors,
                "imagesAnalyzed": False,
            }
        )

    if any(question["status"] == "FAIL" for question in questions):
        overall_status = "FAIL"
    elif any(
        question["status"] == "BLOCKED_PENDING_INGEST"
        for question in questions
    ):
        overall_status = "BLOCKED_PENDING_INGEST"
    else:
        overall_status = "PASS"

    report = {
        "schemaVersion": ACCEPTANCE_SCHEMA_VERSION,
        "manifestSchemaVersion": str(manifest["schemaVersion"]),
        "manifestPath": str(manifest_file),
        "databasePath": str(database),
        "sourceRoot": str(
            Path(str(manifest["sourceRoot"]))
            .expanduser()
            .resolve(strict=False)
        ),
        "overallStatus": overall_status,
        "summary": {
            "goldenQuestionCount": len(questions),
            "passedQuestionCount": sum(
                question["status"] == "PASS" for question in questions
            ),
            "blockedPendingIngestQuestionCount": sum(
                question["status"] == "BLOCKED_PENDING_INGEST"
                for question in questions
            ),
            "failedQuestionCount": sum(
                question["status"] == "FAIL" for question in questions
            ),
            "primarySourceCount": sum(
                len(question["primarySources"]) for question in questions
            ),
            "representedPrimarySourceCount": sum(
                source["status"] == "REPRESENTED"
                for question in questions
                for source in question["primarySources"]
            ),
            "pendingIngestPrimarySourceCount": sum(
                source["status"] == "PENDING_INGEST"
                for question in questions
                for source in question["primarySources"]
            ),
            "retrievalMissPrimarySourceCount": sum(
                source["status"] == "RETRIEVAL_MISS"
                for question in questions
                for source in question["primarySources"]
            ),
            "eligibleEffectCount": sum(
                question["counts"]["eligibleEffectCount"]
                for question in questions
            ),
            "eligibleDataCountByQuestion": sum(
                question["counts"]["eligibleDataCount"]
                for question in questions
            ),
            "eligibleEvidenceCountByQuestion": sum(
                question["counts"]["eligibleEvidenceCount"]
                for question in questions
            ),
            "automatedRequiredBehaviorCount": sum(
                item.get("assertion") is not None
                for question in questions
                for item in question["requiredBehavior"]
            ),
            "passedAutomatedRequiredBehaviorCount": sum(
                item.get("assertion") is not None
                and item["status"] == "PASS"
                for question in questions
                for item in question["requiredBehavior"]
            ),
            "failedAutomatedRequiredBehaviorCount": sum(
                item.get("assertion") is not None
                and item["status"] in {"FAIL", "NOT_RUN"}
                for question in questions
                for item in question["requiredBehavior"]
            ),
            "manualRequiredBehaviorCount": sum(
                item["status"] == "MANUAL_REVIEW_REQUIRED"
                for question in questions
                for item in question["requiredBehavior"]
            ),
        },
        "questions": questions,
        "requiredBehaviorPolicy": (
            "DECLARATIVE_WHEN_CONFIGURED_ELSE_MANUAL_REVIEW_REQUIRED"
        ),
        "imagesAnalyzed": False,
    }
    _atomic_write_json(output / "acceptance-report.json", report)
    return report


def run_acceptance(
    database_path: str | Path,
    manifest_path: str | Path,
    output_directory: str | Path,
    **kwargs: Any,
) -> dict[str, Any]:
    """Short alias retained for callers that do not need the longer name."""

    return run_golden_question_acceptance(
        database_path,
        manifest_path,
        output_directory,
        **kwargs,
    )


__all__ = [
    "ACCEPTANCE_SCHEMA_VERSION",
    "AcceptanceReportError",
    "run_acceptance",
    "run_golden_question_acceptance",
]
