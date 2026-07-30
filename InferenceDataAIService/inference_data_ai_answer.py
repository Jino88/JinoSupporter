"""Deterministic Korean answers constrained to a canonical evidence pack."""

from __future__ import annotations

import hashlib
import json
import re
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from typing import Any

from inference_data_ai_query import EVIDENCE_PACK_SCHEMA_VERSION, normalize_text


ANSWER_SCHEMA_VERSION = "canonical-evidence-answer-v1"


class EvidenceAnswerError(RuntimeError):
    """Raised when an evidence pack cannot support a trustworthy answer."""


def _canonical_bytes(value: object) -> bytes:
    try:
        text = json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
            allow_nan=False,
        )
    except (TypeError, ValueError) as exc:
        raise EvidenceAnswerError(
            "Evidence answer input is not canonical JSON data."
        ) from exc
    return text.encode("utf-8")


def evidence_pack_sha256(pack: dict[str, Any]) -> str:
    return hashlib.sha256(_canonical_bytes(pack)).hexdigest()


def _decimal(value: object, field: str) -> Decimal:
    if value is None or isinstance(value, bool):
        raise EvidenceAnswerError(f"{field} requires a finite effect estimate.")
    try:
        number = Decimal(str(value))
    except (InvalidOperation, ValueError) as exc:
        raise EvidenceAnswerError(f"{field} must be numeric.") from exc
    if not number.is_finite():
        raise EvidenceAnswerError(f"{field} must be finite.")
    return number


def _json_number(value: Decimal) -> int | float:
    if value == value.to_integral_value():
        return int(value)
    return float(value)


def _format_number(value: object) -> str:
    number = _decimal(value, "display number")
    text = format(number.normalize(), "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


def _concept_identity(item: dict[str, Any]) -> str:
    concept = item.get("concept")
    if isinstance(concept, dict) and concept.get("conceptId") is not None:
        return f"CONCEPT:{concept['conceptId']}"
    return "LABEL:" + normalize_text(
        item.get("originalLabel")
        or item.get("originalValue")
        or item.get("normalizedValue")
        or ""
    )


def _normalized_term_matches(term: str, text: str) -> bool:
    """Match Latin/digit terms on token boundaries; retain CJK substring use."""

    if re.fullmatch(r"[a-z0-9]+", term):
        return term in set(re.findall(r"[a-z0-9]+", text))
    return term in text


def _sorted_signature(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(
        items,
        key=lambda item: _canonical_bytes(item),
    )


def _compatibility_signature(
    effect: dict[str, Any],
    study: dict[str, Any],
) -> dict[str, Any]:
    outcome = effect["outcome"]
    compared_arm = effect["comparison"]["comparedArm"]
    control_arm = effect["comparison"]["controlArm"]
    contexts = _sorted_signature(
        [
            {
                "kind": str(context.get("kind") or ""),
                "identity": _concept_identity(context),
                "originalValue": str(context.get("originalValue") or ""),
                "normalizedValue": str(context.get("normalizedValue") or ""),
                "valueNumber": context.get("valueNumber"),
                "unit": str(context.get("unit") or ""),
                "startValue": str(context.get("startValue") or ""),
                "endValue": str(context.get("endValue") or ""),
            }
            for context in study.get("contexts", [])
        ]
    )
    factors = _sorted_signature(
        [
            {
                "identity": _concept_identity(factor),
                "originalLabel": str(factor.get("originalLabel") or ""),
                "baselineCondition": str(
                    factor.get("baselineCondition") or ""
                ),
                "changedCondition": str(
                    factor.get("changedCondition") or ""
                ),
                "changeDirection": str(
                    factor.get("changeDirection") or ""
                ),
                "isolationStatus": str(
                    factor.get("isolationStatus") or ""
                ),
            }
            for factor in study.get("factors", [])
        ]
    )

    def arm_signature(arm: dict[str, Any]) -> dict[str, Any]:
        return {
            "role": str(arm.get("role") or ""),
            "condition": str(arm.get("condition") or ""),
            "sampleBasis": str(arm.get("sampleBasis") or ""),
            "matchingBasis": str(arm.get("matchingBasis") or ""),
            "factorValues": _sorted_signature(
                [
                    {
                        "factorUid": str(value.get("factorUid") or ""),
                        "factorLabel": str(value.get("factorLabel") or ""),
                        "originalValue": str(
                            value.get("originalValue") or ""
                        ),
                        "valueNumber": value.get("valueNumber"),
                        "unit": str(value.get("unit") or ""),
                        "isBaseline": bool(value.get("isBaseline")),
                        "heldConstant": bool(value.get("heldConstant")),
                    }
                    for value in arm.get("factorValues", [])
                ]
            ),
        }

    observation_identity = {
        side: sorted(
            {
                (
                    str(observation.get("stratumKey") or ""),
                    str(observation.get("replicateKey") or ""),
                )
                for observation in effect.get("observations", {}).get(
                    side,
                    [],
                )
            }
        )
        for side in ("comparedArm", "controlArm")
    }
    return {
        "outcome": {
            "identity": _concept_identity(outcome),
            "metricType": str(outcome.get("metricType") or ""),
            "unit": str(
                effect["effect"].get("unit")
                or outcome.get("unit")
                or ""
            ),
            "denominatorBasis": str(
                outcome.get("denominatorBasis") or ""
            ),
        },
        "effectType": str(effect["effect"].get("effectType") or ""),
        "formulaVersion": str(
            effect["effect"].get("formulaVersion") or ""
        ),
        "contexts": contexts,
        "factors": factors,
        "controlArm": arm_signature(control_arm),
        "comparedArm": arm_signature(compared_arm),
        "designType": str(
            effect["comparison"].get("designType") or ""
        ),
        "matchingBasis": str(
            effect["comparison"].get("matchingBasis") or ""
        ),
        "observationIdentity": observation_identity,
    }


def _direct_effect_citations(
    effect: dict[str, Any],
) -> list[dict[str, Any]]:
    source = effect.get("source") or {}
    result: list[dict[str, Any]] = []
    for citation in effect.get("evidence", []):
        direct = any(
            str(link.get("entityType") or "").upper() == "EFFECT"
            for link in citation.get("linkedEntities", [])
        )
        if not direct:
            continue
        if str(citation.get("verificationStatus") or "").upper() != "VERIFIED":
            continue
        if int(citation.get("revisionId") or -1) != int(
            source.get("revisionId") or -2
        ):
            continue
        if (
            str(citation.get("contentSha256") or "").lower()
            != str(source.get("contentSha256") or "").lower()
        ):
            continue
        if str(citation.get("sourcePath") or "") != str(
            source.get("sourcePath") or ""
        ):
            continue
        result.append(citation)
    return sorted(
        result,
        key=lambda citation: str(citation["publicEvidenceId"]),
    )


def _citation_payload(
    citation: dict[str, Any],
    data_id: str,
) -> dict[str, Any]:
    return {
        "evidenceId": str(citation["publicEvidenceId"]),
        "dataIds": [data_id],
        "sourcePath": str(citation["sourcePath"]),
        "sheet": str(citation["sheet"]),
        "range": str(citation["range"]),
        "contentSha256": str(citation["contentSha256"]),
        "verificationStatus": str(citation["verificationStatus"]),
    }


def _add_citation(
    citations: dict[str, dict[str, Any]],
    citation: dict[str, Any],
    data_id: str,
) -> None:
    payload = _citation_payload(citation, data_id)
    evidence_id = payload["evidenceId"]
    if evidence_id not in citations:
        citations[evidence_id] = payload
    elif data_id not in citations[evidence_id]["dataIds"]:
        citations[evidence_id]["dataIds"].append(data_id)
        citations[evidence_id]["dataIds"].sort()


def _observation_citations(
    evidence: list[dict[str, Any]],
    observation_uid: str,
    source: dict[str, Any],
    outcome_uid: str = "",
) -> list[dict[str, Any]]:
    return sorted(
        [
            citation
            for citation in evidence
            if str(citation.get("verificationStatus") or "").upper()
            == "VERIFIED"
            and int(citation.get("revisionId") or -1)
            == int(source.get("revisionId") or -2)
            and str(citation.get("contentSha256") or "").lower()
            == str(source.get("contentSha256") or "").lower()
            and str(citation.get("sourcePath") or "")
            == str(source.get("sourcePath") or "")
            and any(
                (
                    str(link.get("entityType") or "").upper()
                    == "OBSERVATION"
                    and str(link.get("entityUid") or "")
                    == observation_uid
                )
                or (
                    bool(outcome_uid)
                    and str(link.get("entityType") or "").upper()
                    == "OUTCOME"
                    and str(link.get("entityUid") or "") == outcome_uid
                )
                for link in citation.get("linkedEntities", [])
            )
        ],
        key=lambda citation: str(citation["publicEvidenceId"]),
    )


def _measurement_series_citations(
    evidence: list[dict[str, Any]],
    series_uid: str,
    source: dict[str, Any],
) -> list[dict[str, Any]]:
    direct = sorted(
        [
            citation
            for citation in evidence
            if str(citation.get("verificationStatus") or "").upper()
            == "VERIFIED"
            and int(citation.get("revisionId") or -1)
            == int(source.get("revisionId") or -2)
            and str(citation.get("contentSha256") or "").lower()
            == str(source.get("contentSha256") or "").lower()
            and str(citation.get("sourcePath") or "")
            == str(source.get("sourcePath") or "")
            and any(
                str(link.get("entityType") or "").upper()
                == "MEASUREMENT_SERIES"
                and str(link.get("entityUid") or "") == series_uid
                for link in citation.get("linkedEntities", [])
            )
        ],
        key=lambda citation: str(citation["publicEvidenceId"]),
    )
    # Summary statistics are calculated from the value matrix.  Header-only
    # or row-identity-only links cannot justify displaying those numbers.
    if not any(
        "MEASUREMENT_VALUES"
        in {
            str(citation.get("role") or "").upper(),
            str(citation.get("linkRole") or "").upper(),
        }
        for citation in direct
    ):
        return []
    return direct


def _comparison_citations(
    evidence: list[dict[str, Any]],
    comparison_uid: str,
    source: dict[str, Any],
) -> list[dict[str, Any]]:
    """Return only direct, verified citations for one comparison."""

    if not comparison_uid:
        return []
    return sorted(
        [
            citation
            for citation in evidence
            if str(citation.get("verificationStatus") or "").upper()
            == "VERIFIED"
            and int(citation.get("revisionId") or -1)
            == int(source.get("revisionId") or -2)
            and str(citation.get("contentSha256") or "").lower()
            == str(source.get("contentSha256") or "").lower()
            and str(citation.get("sourcePath") or "")
            == str(source.get("sourcePath") or "")
            and any(
                str(link.get("entityType") or "").upper()
                == "COMPARISON"
                and str(link.get("entityUid") or "") == comparison_uid
                for link in citation.get("linkedEntities", [])
            )
        ],
        key=lambda citation: str(citation["publicEvidenceId"]),
    )


def _descriptive_studies(
    pack: dict[str, Any],
    study_index: dict[str, dict[str, Any]],
    citations: dict[str, dict[str, Any]],
) -> tuple[list[dict[str, Any]], int]:
    result: list[dict[str, Any]] = []
    omitted_uncited_observations = 0
    seen_records: set[tuple[str, str | None]] = set()
    outcome_terms = [
        normalize_text(term)
        for term in (
            pack.get("queryRoleHints", {}).get("outcomeTerms", [])
            if isinstance(pack.get("queryRoleHints"), dict)
            else []
        )
        if normalize_text(term)
    ]
    for record in pack.get("excludedCandidates", []):
        descriptive = record.get("descriptiveOutcomes") or []
        if not descriptive:
            continue
        if outcome_terms:
            matching_outcomes = [
                outcome_record
                for outcome_record in descriptive
                if any(
                    _normalized_term_matches(
                        term,
                        normalize_text(
                        " ".join(
                            str(outcome_record.get("outcome", {}).get(field) or "")
                            for field in (
                                "originalLabel",
                                "outcomeKey",
                                "domain",
                                "metricType",
                                "definition",
                            )
                        )
                        ),
                    )
                    for term in outcome_terms
                )
            ]
            title_proxy_match = any(
                str(field.get("field") or "")
                in {
                    "source.fileName",
                    "analysis.title",
                    "study.title",
                }
                and set(field.get("terms") or []) & set(outcome_terms)
                for field in study_index[
                    str(record["publicDataId"])
                ].get("relevance", {}).get("matchedFields", [])
            )
            # Keep the complete source-backed descriptive set when the
            # question's inferred outcome words do not occur in any stored
            # outcome label/definition but a title carries that broad outcome.
            # The detailed submetrics are the source-backed result set.
            if matching_outcomes:
                descriptive = matching_outcomes
            elif not title_proxy_match and bool(
                pack.get("queryRoleHints", {}).get(
                    "relationGateApplied"
                )
            ):
                # A relationship question has an explicit outcome side.
                # Unrelated outcomes must not be substituted merely because
                # the workbook title matched the broad question.
                descriptive = []
        data_id = str(record["publicDataId"])
        if data_id not in study_index:
            raise EvidenceAnswerError(
                f"Descriptive record references unknown Study {data_id}."
            )
        source = study_index[data_id]["source"]
        record_key = (
            data_id,
            None
            if str(record.get("descriptiveScope") or "").upper() == "STUDY"
            else record.get("publicComparisonId"),
        )
        if record_key in seen_records:
            continue
        seen_records.add(record_key)
        outcomes: list[dict[str, Any]] = []
        for outcome_record in descriptive:
            outcome = outcome_record["outcome"]
            arms: list[dict[str, Any]] = []
            for arm_record in outcome_record.get("armObservations", []):
                arm = arm_record["arm"]
                observations: list[dict[str, Any]] = []
                for observation in arm_record.get("observations", []):
                    observation_uid = str(observation["observationUid"])
                    observation_evidence = _observation_citations(
                        record.get("evidence", []),
                        observation_uid,
                        source,
                        str(outcome.get("outcomeUid") or ""),
                    )
                    has_value = any(
                        observation.get(field) not in (None, "")
                        for field in (
                            "valueNumber",
                            "valueText",
                            "numerator",
                            "denominator",
                            "ratePpm",
                            "min",
                            "max",
                            "average",
                        )
                    )
                    if not has_value:
                        continue
                    if not observation_evidence:
                        # Legacy rows may predate exact observation-to-cell links.
                        # Never display a value that cannot be opened at its
                        # exact source cell. The record remains represented by
                        # excludedRecords instead of failing the whole answer.
                        omitted_uncited_observations += 1
                        continue
                    evidence_ids = [
                        str(item["publicEvidenceId"])
                        for item in observation_evidence
                    ]
                    for item in observation_evidence:
                        _add_citation(citations, item, data_id)
                    observations.append(
                        {
                            "observationUid": observation_uid,
                            "valueNumber": observation.get("valueNumber"),
                            "valueText": str(
                                observation.get("valueText") or ""
                            ),
                            "numerator": observation.get("numerator"),
                            "denominator": observation.get("denominator"),
                            "ratePpm": observation.get("ratePpm"),
                            "min": observation.get("min"),
                            "max": observation.get("max"),
                            "average": observation.get("average"),
                            "sampleSize": observation.get("sampleSize"),
                            "verificationStatus": str(
                                observation.get("verificationStatus") or ""
                            ),
                            "evidenceIds": evidence_ids,
                        }
                    )
                if observations:
                    arms.append(
                        {
                            "armLabel": str(arm.get("label") or ""),
                            "condition": str(arm.get("condition") or ""),
                            "observations": observations,
                        }
                    )
            if arms:
                outcomes.append(
                    {
                        "label": str(outcome.get("originalLabel") or ""),
                        "unit": str(outcome.get("unit") or ""),
                        "arms": arms,
                    }
                )
        if outcomes:
            result.append(
                {
                    "dataId": data_id,
                    "status": "DESCRIPTIVE_ONLY",
                    "reasonCodes": sorted(
                        {
                            str(reason["code"])
                            for reason in record.get("exclusionReasons", [])
                        }
                    ),
                    "outcomes": outcomes,
                }
            )
    return (
        sorted(result, key=lambda item: item["dataId"]),
        omitted_uncited_observations,
    )


def _series_matches_outcome_terms(
    series: dict[str, Any],
    outcome_terms: list[str],
) -> bool:
    if not outcome_terms:
        return True
    outcome = series.get("outcome", {})
    searchable = normalize_text(
        " ".join(
            str(outcome.get(field) or "")
            for field in (
                "originalLabel",
                "outcomeKey",
                "domain",
                "metricType",
            )
        )
    )
    return any(
        _normalized_term_matches(term, searchable)
        for term in outcome_terms
    )


def _descriptive_measurement_studies(
    pack: dict[str, Any],
    study_index: dict[str, dict[str, Any]],
    citations: dict[str, dict[str, Any]],
) -> tuple[list[dict[str, Any]], int]:
    result: dict[str, dict[str, Any]] = {}
    omitted_uncited_series = 0
    seen_series: set[tuple[str, str]] = set()
    outcome_terms = [
        normalize_text(term)
        for term in (
            pack.get("queryRoleHints", {}).get("outcomeTerms", [])
            if isinstance(pack.get("queryRoleHints"), dict)
            else []
        )
        if normalize_text(term)
    ]
    for record in pack.get("excludedCandidates", []):
        raw_series = record.get("descriptiveMeasurementSeries") or []
        if not raw_series:
            continue
        matching_series = [
            series
            for series in raw_series
            if _series_matches_outcome_terms(series, outcome_terms)
        ]
        if matching_series:
            raw_series = matching_series
        elif outcome_terms and bool(
            pack.get("queryRoleHints", {}).get("relationGateApplied")
        ):
            # Fail closed for relationship questions: an acoustic/profile
            # matrix is not descriptive evidence for an unrelated NG outcome.
            raw_series = []
        data_id = str(record["publicDataId"])
        if data_id not in study_index:
            raise EvidenceAnswerError(
                f"Measurement series references unknown Study {data_id}."
            )
        source = study_index[data_id]["source"]
        study_payload = result.setdefault(
            data_id,
            {
                "dataId": data_id,
                "status": "DESCRIPTIVE_ONLY",
                "reasonCodes": set(),
                "outcomes": [],
                "measurementSeries": [],
            },
        )
        study_payload["reasonCodes"].update(
            str(reason["code"])
            for reason in record.get("exclusionReasons", [])
        )
        for series in raw_series:
            series_uid = str(series.get("seriesUid") or "")
            key = (data_id, series_uid)
            if not series_uid or key in seen_series:
                continue
            series_evidence = _measurement_series_citations(
                record.get("evidence", []),
                series_uid,
                source,
            )
            if not series_evidence:
                omitted_uncited_series += 1
                continue
            seen_series.add(key)
            for citation in series_evidence:
                _add_citation(citations, citation, data_id)
            summary = series.get("pointSummary") or {}
            point_count = int(summary.get("pointCount") or 0)
            if point_count <= 0:
                continue
            study_payload["measurementSeries"].append(
                {
                    "seriesUid": series_uid,
                    "publicSeriesId": str(
                        series.get("publicSeriesId") or ""
                    ),
                    "seriesKey": str(series.get("seriesKey") or ""),
                    "outcomeLabel": str(
                        series.get("outcome", {}).get("originalLabel")
                        or ""
                    ),
                    "armLabel": str(
                        series.get("arm", {}).get("label") or ""
                    ),
                    "condition": str(
                        series.get("arm", {}).get("condition") or ""
                    ),
                    "axisLabel": str(series.get("axisLabel") or ""),
                    "axisSource": str(series.get("axisSource") or ""),
                    "axisUnit": str(series.get("axisUnit") or ""),
                    "valueUnit": str(series.get("valueUnit") or ""),
                    "stratumKey": str(series.get("stratumKey") or ""),
                    "replicateKeys": [
                        str(value)
                        for value in series.get("replicateKeys", [])
                    ],
                    "aggregateReplicateKeys": [
                        str(value)
                        for value in series.get(
                            "aggregateReplicateKeys",
                            [],
                        )
                    ],
                    "verificationStatus": str(
                        series.get("verificationStatus") or ""
                    ),
                    "interpretationStatus": "DESCRIPTIVE_ONLY",
                    "pointSummary": {
                        "pointCount": point_count,
                        "rawPointCount": int(
                            summary.get("rawPointCount", point_count)
                            or 0
                        ),
                        "aggregatePointCount": int(
                            summary.get("aggregatePointCount") or 0
                        ),
                        "minimum": summary.get("minimum"),
                        "maximum": summary.get("maximum"),
                        "average": summary.get("average"),
                        "distinctAxisCount": int(
                            summary.get("distinctAxisCount") or 0
                        ),
                        "distinctReplicateCount": int(
                            summary.get("distinctReplicateCount") or 0
                        ),
                        "aggregateReplicateCount": int(
                            summary.get("aggregateReplicateCount") or 0
                        ),
                    },
                    "evidenceIds": [
                        str(item["publicEvidenceId"])
                        for item in series_evidence
                    ],
                }
            )
    payloads: list[dict[str, Any]] = []
    for data_id in sorted(result):
        item = result[data_id]
        if not item["measurementSeries"]:
            continue
        item["reasonCodes"] = sorted(item["reasonCodes"])
        item["measurementSeries"].sort(
            key=lambda series: (
                series["publicSeriesId"],
                series["seriesUid"],
            )
        )
        payloads.append(item)
    return payloads, omitted_uncited_series


def _merge_descriptive_studies(
    observation_studies: list[dict[str, Any]],
    measurement_studies: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for study in [*observation_studies, *measurement_studies]:
        data_id = str(study["dataId"])
        target = result.setdefault(
            data_id,
            {
                "dataId": data_id,
                "status": "DESCRIPTIVE_ONLY",
                "reasonCodes": [],
                "outcomes": [],
                "measurementSeries": [],
            },
        )
        target["reasonCodes"] = sorted(
            {
                *target["reasonCodes"],
                *study.get("reasonCodes", []),
            }
        )
        target["outcomes"].extend(study.get("outcomes", []))
        target["measurementSeries"].extend(
            study.get("measurementSeries", [])
        )
    return [result[data_id] for data_id in sorted(result)]


def _quantitative_groups(
    pack: dict[str, Any],
    study_index: dict[str, dict[str, Any]],
    citations: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    groups: dict[str, dict[str, Any]] = {}
    for item in pack.get("answerEligibleEffects", []):
        data_id = str(item["publicDataId"])
        if data_id not in study_index:
            raise EvidenceAnswerError(
                f"Eligible effect references unknown Study {data_id}."
            )
        estimate = _decimal(
            item.get("effect", {}).get("estimate"),
            f"{item.get('publicEffectId')}.estimate",
        )
        direct_citations = _direct_effect_citations(item)
        if not direct_citations:
            raise EvidenceAnswerError(
                "Eligible effect lacks direct verified current-revision evidence: "
                + str(item.get("publicEffectId"))
            )
        for citation in direct_citations:
            _add_citation(citations, citation, data_id)
        signature = _compatibility_signature(item, study_index[data_id])
        signature_bytes = _canonical_bytes(signature)
        group_key = "GROUP-" + hashlib.sha256(signature_bytes).hexdigest()[
            :12
        ].upper()
        group = groups.setdefault(
            group_key,
            {
                "groupKey": group_key,
                "compatibility": "EXACT",
                "outcome": {
                    "identity": signature["outcome"]["identity"],
                    "label": str(
                        item["outcome"].get("originalLabel") or ""
                    ),
                    "metricType": str(
                        item["outcome"].get("metricType") or ""
                    ),
                    "denominatorBasis": str(
                        item["outcome"].get("denominatorBasis") or ""
                    ),
                },
                "effectType": str(
                    item["effect"].get("effectType") or ""
                ),
                "unit": str(item["effect"].get("unit") or ""),
                "formulaVersion": str(
                    item["effect"].get("formulaVersion") or ""
                ),
                "contextSignature": signature["contexts"],
                "factorTransitionSignature": signature["factors"],
                "designType": signature["designType"],
                "matchingBasis": signature["matchingBasis"],
                "_estimates": [],
                "effects": [],
            },
        )
        group["_estimates"].append(estimate)
        comparison = item.get("comparison") or {}
        factor_differences = [
            {
                "factorUid": str(value.get("factorUid") or ""),
                "factorLabel": str(value.get("factorLabel") or ""),
                "controlValue": str(value.get("controlValue") or ""),
                "comparedValue": str(value.get("comparedValue") or ""),
                "controlValueRecorded": bool(
                    value.get("controlValueRecorded")
                ),
                "comparedValueRecorded": bool(
                    value.get("comparedValueRecorded")
                ),
            }
            for value in comparison.get("factorDifferences", [])
            if isinstance(value, dict)
        ]
        group["effects"].append(
            {
                "dataId": data_id,
                "comparisonId": str(item["publicComparisonId"]),
                "effectId": str(item["publicEffectId"]),
                "estimate": _json_number(estimate),
                "displayEstimate": _format_number(estimate),
                "unit": str(item["effect"].get("unit") or ""),
                "direction": str(
                    item["effect"].get("direction") or ""
                ),
                "comparedArmLabel": str(
                    item["comparison"]["comparedArm"].get("label") or ""
                ),
                "controlArmLabel": str(
                    item["comparison"]["controlArm"].get("label") or ""
                ),
                "controlCondition": str(
                    item["comparison"]["controlArm"].get("condition") or ""
                ),
                "comparedCondition": str(
                    item["comparison"]["comparedArm"].get("condition") or ""
                ),
                "factorDifferences": factor_differences,
                "evidenceIds": [
                    str(citation["publicEvidenceId"])
                    for citation in direct_citations
                ],
            }
        )
    result: list[dict[str, Any]] = []
    for group_key in sorted(groups):
        group = groups[group_key]
        estimates: list[Decimal] = group.pop("_estimates")
        signs = {
            1 if estimate > 0 else -1 if estimate < 0 else 0
            for estimate in estimates
        }
        if 1 in signs and -1 in signs:
            direction_status = "CONFLICTING"
        elif signs == {0}:
            direction_status = "NO_CHANGE"
        elif signs:
            direction_status = "CONSISTENT"
        else:
            direction_status = "UNKNOWN"
        group["directionStatus"] = direction_status
        group["statistics"] = {
            "effectCount": len(estimates),
            "uniqueDataCount": len(
                {effect["dataId"] for effect in group["effects"]}
            ),
            "uniqueComparisonCount": len(
                {
                    effect["comparisonId"]
                    for effect in group["effects"]
                }
            ),
            "minimum": _json_number(min(estimates)),
            "maximum": _json_number(max(estimates)),
            "mean": _json_number(sum(estimates) / Decimal(len(estimates))),
            "meanRounding": "NONE_DECIMAL_EXACT",
        }
        group["effects"].sort(
            key=lambda effect: (
                effect["dataId"],
                effect["comparisonId"],
                effect["effectId"],
            )
        )
        result.append(group)
    return result


def _confounding_assessment(
    comparison: dict[str, Any] | None,
) -> dict[str, Any] | None:
    if not comparison:
        return None
    if str(comparison.get("confoundingStatus") or "").upper() != "CONFOUNDED":
        return None
    factor_differences = [
        {
            "factorLabel": str(item.get("factorLabel") or ""),
            "comparedValue": str(item.get("comparedValue") or ""),
            "controlValue": str(item.get("controlValue") or ""),
            "comparedValueRecorded": bool(
                item.get("comparedValueRecorded")
            ),
            "controlValueRecorded": bool(item.get("controlValueRecorded")),
        }
        for item in comparison.get("factorDifferences", [])
        if isinstance(item, dict)
    ]
    compared_arm = comparison.get("comparedArm") or {}
    control_arm = comparison.get("controlArm") or {}
    compared_condition = str(compared_arm.get("condition") or "")
    control_condition = str(control_arm.get("condition") or "")
    conditions_differ = (
        bool(compared_condition or control_condition)
        and normalize_text(compared_condition)
        != normalize_text(control_condition)
    )
    if len(factor_differences) >= 2:
        code = "CONFOUNDED_MULTI_FACTOR"
        labels = [
            item["factorLabel"]
            for item in factor_differences
            if item["factorLabel"]
        ]
        label_text = ", ".join(labels) if labels else "여러 조건"
        text_ko = (
            f"{label_text} 등 {len(factor_differences)}개 요인이 동시에 "
            "다르거나 한쪽 값이 기록되지 않아 단일 요인의 효과로 "
            "분리할 수 없습니다."
        )
    else:
        code = "CONFOUNDED_COMPARISON"
        text_ko = (
            "비교군과 대조군 사이에 통제되지 않았거나 값이 명시되지 "
            "않은 조건이 있어 단일 요인의 효과로 분리할 수 없습니다."
        )
    return {
        "code": code,
        "textKo": text_ko,
        "factorDifferences": factor_differences,
        "conditionsDiffer": conditions_differ,
        "comparedArmLabel": str(compared_arm.get("label") or ""),
        "controlArmLabel": str(control_arm.get("label") or ""),
        "comparedCondition": compared_condition,
        "controlCondition": control_condition,
        "designType": str(comparison.get("designType") or ""),
        "matchingBasis": str(comparison.get("matchingBasis") or ""),
        "summary": str(comparison.get("summary") or ""),
        "exclusionReason": str(comparison.get("exclusionReason") or ""),
    }


def _excluded_records(
    pack: dict[str, Any],
    study_index: dict[str, dict[str, Any]],
    citations: dict[str, dict[str, Any]],
) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for record in pack.get("excludedCandidates", []):
        data_id = str(record["publicDataId"])
        comparison = record.get("comparison")
        assessment = _confounding_assessment(comparison)
        direct_citations: list[dict[str, Any]] = []
        if assessment and comparison and data_id in study_index:
            direct_citations = _comparison_citations(
                record.get("evidence", []),
                str(comparison.get("comparisonUid") or ""),
                study_index[data_id]["source"],
            )
            for citation in direct_citations:
                _add_citation(citations, citation, data_id)
        reason_codes = {
            str(reason["code"])
            for reason in record.get("exclusionReasons", [])
        }
        if assessment:
            reason_codes.add(str(assessment["code"]))
        records.append(
            {
                "dataId": data_id,
                "comparisonId": (
                    str(record["publicComparisonId"])
                    if record.get("publicComparisonId")
                    else None
                ),
                "effectId": (
                    str(record["publicEffectId"])
                    if record.get("publicEffectId")
                    else None
                ),
                "reasonCodes": sorted(reason_codes),
                "evidenceIds": [
                    str(citation["publicEvidenceId"])
                    for citation in direct_citations
                ],
                "comparisonAssessment": assessment,
            }
        )
    return sorted(
        records,
        key=lambda record: (
            record["dataId"],
            record["comparisonId"] or "",
            record["effectId"] or "",
        ),
    )


def _source_exclusions(pack: dict[str, Any]) -> list[dict[str, Any]]:
    records = [
        {
            "publicAnalysisId": str(record["publicAnalysisId"]),
            "revisionUid": str(record["revisionUid"]),
            "sourcePath": str(record["sourcePath"]),
            "fileName": str(record["fileName"]),
            "sourceContentStatus": str(record["sourceContentStatus"]),
            "reasonCodes": sorted(
                {
                    str(reason["code"])
                    for reason in record.get("exclusionReasons", [])
                }
            ),
            "imagesAnalyzed": False,
        }
        for record in pack.get("sourceExclusions", [])
    ]
    return sorted(
        records,
        key=lambda record: (
            record["publicAnalysisId"],
            record["revisionUid"],
        ),
    )


LIMITATION_TEXT = {
    "NO_VALID_COMPARISON": (
        "검증된 대조군·비교군 효과가 없어 관계의 크기나 방향을 "
        "수치로 판단하지 않았습니다."
    ),
    "EXCLUDED_RECORDS_PRESENT": (
        "검토 필요·교란·무효·비교 없음 자료는 정량 결론에서 제외했습니다."
    ),
    "DESCRIPTIVE_ONLY": (
        "비교가 없는 자료의 관측값은 그대로 제시하되 조건 간 차이를 새로 계산하지 않았습니다."
    ),
    "DESCRIPTIVE_EVIDENCE_MISSING": (
        "직접 검증된 셀 근거가 없는 과거 설명값은 표시하지 않았습니다."
    ),
    "CONFLICTING_DIRECTION": (
        "완전히 같은 비교 가능 조건 안에서도 효과 방향이 일치하지 않아 단일 결론으로 합치지 않았습니다."
    ),
    "INCOMPATIBLE_GROUPS": (
        "모델·LOT·조건·요인 전환·지표 정의가 다른 효과는 별도 그룹으로 유지했습니다."
    ),
    "CONFOUNDED_MULTI_FACTOR": (
        "둘 이상의 조건이 함께 달라진 비교는 단일 요인의 정량 효과에서 제외했습니다."
    ),
    "CONFOUNDED_COMPARISON": (
        "통제되지 않았거나 값이 명시되지 않은 조건이 있는 비교는 단일 요인의 "
        "정량 효과에서 제외했습니다."
    ),
    "NO_RELEVANT_DATA": "질문과 양쪽 의미가 일치하는 검토 자료를 찾지 못했습니다.",
}


def _limitations(
    *,
    status: str,
    groups: list[dict[str, Any]],
    descriptive: list[dict[str, Any]],
    excluded: list[dict[str, Any]],
    source_exclusions: list[dict[str, Any]],
    omitted_uncited_observations: int,
    omitted_uncited_measurement_series: int,
) -> list[dict[str, Any]]:
    codes: list[str] = []
    if status == "NO_RELEVANT_DATA":
        codes.append("NO_RELEVANT_DATA")
    if status == "INSUFFICIENT_COMPARISON":
        codes.append("NO_VALID_COMPARISON")
    if excluded or source_exclusions:
        codes.append("EXCLUDED_RECORDS_PRESENT")
    if descriptive:
        codes.append("DESCRIPTIVE_ONLY")
    if (
        omitted_uncited_observations
        or omitted_uncited_measurement_series
    ):
        codes.append("DESCRIPTIVE_EVIDENCE_MISSING")
    if any(
        group["directionStatus"] == "CONFLICTING"
        for group in groups
    ):
        codes.append("CONFLICTING_DIRECTION")
    if len(groups) > 1:
        codes.append("INCOMPATIBLE_GROUPS")
    for code in ("CONFOUNDED_MULTI_FACTOR", "CONFOUNDED_COMPARISON"):
        if any(
            (record.get("comparisonAssessment") or {}).get("code")
            == code
            for record in excluded
        ):
            codes.append(code)
    return [
        {
            "code": code,
            "textKo": LIMITATION_TEXT[code],
            "relatedIds": sorted(
                {
                    str(record["comparisonId"])
                    for record in excluded
                    if record.get("comparisonId")
                    and (
                        record.get("comparisonAssessment") or {}
                    ).get("code")
                    == code
                }
            ),
        }
        for code in codes
    ]


def _direct_answer_text(
    status: str,
    relevant_count: int,
    groups: list[dict[str, Any]],
    eligible_count: int,
    excluded: list[dict[str, Any]],
) -> tuple[str, str]:
    confounded_comparison_count = len(
        {
            (record["dataId"], record["comparisonId"])
            for record in excluded
            if record.get("comparisonId")
            and record.get("comparisonAssessment")
        }
    )
    confounded_suffix = (
        f" 이 중 교란 비교 {confounded_comparison_count}건은 조건이 "
        "통제되지 않았거나 여러 요인이 함께 달라 단일 요인의 효과로 "
        "계산하지 않았습니다."
        if confounded_comparison_count
        else ""
    )
    if status == "NO_RELEVANT_DATA":
        return "NO_DATA", LIMITATION_TEXT["NO_RELEVANT_DATA"]
    if status == "INSUFFICIENT_COMPARISON":
        return (
            "NO_VALID_COMPARISON",
            f"관련 검토 자료는 {relevant_count}건 검색되었지만, "
            "검증된 대조군·비교군 효과가 없어 관계의 크기나 방향을 "
            f"수치로 판단할 수 없습니다.{confounded_suffix}",
        )
    if status == "CONFLICTING":
        return (
            "CONFLICT",
            f"검증된 효과 {eligible_count}건이 있으나 동일한 비교 가능 "
            "조건 안에서 방향이 충돌하여 하나의 평균 효과나 인과 결론으로 "
            "합칠 수 없습니다.",
        )
    if status == "PARTIAL":
        return (
            "MIXED_HISTORY",
            f"검증된 비교 효과 {eligible_count}건을 {len(groups)}개 호환 "
            "조건 그룹으로 확인했습니다. 제외 또는 서술 전용 자료는 정량 "
            "결론에 섞지 않았으며, 관찰 이력은 인과관계를 뜻하지 않습니다.",
        )
    return (
        "VERIFIED_HISTORY",
        f"검증된 비교 효과 {eligible_count}건을 {len(groups)}개 호환 조건 "
        "그룹에서 확인했습니다. 이는 비교 이력에서 관찰된 차이이며 "
        "그 자체로 인과관계를 확정하지 않습니다.",
    )


def _observation_display(observation: dict[str, Any]) -> str:
    if observation.get("valueText"):
        return str(observation["valueText"])
    for field, suffix in (
        ("valueNumber", ""),
        ("ratePpm", " ppm"),
        ("average", ""),
        ("numerator", ""),
    ):
        if observation.get(field) is not None:
            if field == "numerator" and observation.get("denominator") is not None:
                return (
                    f"{_format_number(observation['numerator'])}/"
                    f"{_format_number(observation['denominator'])}"
                )
            return _format_number(observation[field]) + suffix
    return "(텍스트 값 없음)"


def _effect_transition_text(effect: dict[str, Any]) -> str:
    differences = [
        value
        for value in effect.get("factorDifferences", [])
        if isinstance(value, dict)
    ]
    if differences:
        return "; ".join(
            (
                f"{str(value.get('factorLabel') or '(요인명 미기록)')}: "
                f"{str(value.get('controlValue') or '(미기록)')} → "
                f"{str(value.get('comparedValue') or '(미기록)')}"
            )
            for value in differences
        )
    control_condition = str(effect.get("controlCondition") or "")
    compared_condition = str(effect.get("comparedCondition") or "")
    if control_condition or compared_condition:
        return (
            f"조건: {control_condition or '(미기록)'} → "
            f"{compared_condition or '(미기록)'}"
        )
    return ""


def render_answer_markdown(answer: dict[str, Any]) -> str:
    lines = [
        "# 근거 기반 답변",
        "",
        str(answer["directAnswer"]["textKo"]),
    ]
    groups = answer["quantitativeGroups"]
    if groups:
        lines.extend(["", "## 검증된 비교 이력", ""])
        for group in groups:
            stats = group["statistics"]
            label = group["outcome"]["label"]
            if group["directionStatus"] == "CONFLICTING":
                lines.append(
                    f"- `{label}`: 동일 호환 조건 내 {stats['effectCount']}개 "
                    "효과의 방향이 충돌하여 평균 결론을 사용하지 않았습니다."
                )
            elif stats["effectCount"] == 1:
                effect = group["effects"][0]
                transition = _effect_transition_text(effect)
                ids = " / ".join(
                    [
                        effect["dataId"],
                        effect["comparisonId"],
                        effect["effectId"],
                        *effect["evidenceIds"],
                    ]
                )
                lines.append(
                    f"- `{label}`: "
                    f"{transition + ' 조건에서 ' if transition else ''}"
                    "비교군은 대조군 대비 "
                    f"{effect['displayEstimate']} {effect['unit']}의 저장된 "
                    f"{group['effectType']} 이력이 있습니다. [{ids}]"
                )
            else:
                lines.append(
                    f"- `{label}`: 완전히 같은 호환 조건의 효과 "
                    f"{stats['effectCount']}개에서 범위 "
                    f"{_format_number(stats['minimum'])}~"
                    f"{_format_number(stats['maximum'])} {group['unit']}, "
                    f"산술평균 {_format_number(stats['mean'])} "
                    f"{group['unit']}였습니다."
                )
                for effect in group["effects"]:
                    transition = _effect_transition_text(effect)
                    ids = " / ".join(
                        [
                            effect["dataId"],
                            effect["comparisonId"],
                            effect["effectId"],
                            *effect["evidenceIds"],
                        ]
                    )
                    lines.append(
                        f"  - "
                        f"{transition + ': ' if transition else ''}"
                        f"{effect['displayEstimate']} {effect['unit']} [{ids}]"
                    )
    if answer["descriptiveStudies"]:
        lines.extend(["", "## 비교 미확인 서술 자료", ""])
        for study in answer["descriptiveStudies"]:
            lines.append(f"- `{study['dataId']}`")
            for outcome in study["outcomes"]:
                for arm in outcome["arms"]:
                    for observation in arm["observations"]:
                        evidence = " / ".join(observation["evidenceIds"])
                        review_note = (
                            " (검토 필요 자료)"
                            if observation["verificationStatus"] != "VERIFIED"
                            else ""
                        )
                        lines.append(
                            f"  - {arm['armLabel']} / {outcome['label']}: "
                            f"{_observation_display(observation)}"
                            f"{review_note} [{study['dataId']}"
                            f"{' / ' + evidence if evidence else ''}]"
                        )
            for series in study.get("measurementSeries", []):
                summary = series["pointSummary"]
                evidence = " / ".join(series["evidenceIds"])
                review_note = (
                    " (검토 필요 자료)"
                    if series["verificationStatus"] != "VERIFIED"
                    else ""
                )
                value_unit = (
                    f" {series['valueUnit']}"
                    if series["valueUnit"]
                    else ""
                )
                series_id = (
                    series["publicSeriesId"]
                    or series["seriesUid"]
                )
                aggregate_count = int(
                    summary.get("aggregatePointCount") or 0
                )
                point_breakdown = (
                    f" ({summary['rawPointCount']} raw + "
                    f"{aggregate_count} aggregate)"
                    if aggregate_count
                    else ""
                )
                aggregate_replicates = int(
                    summary.get("aggregateReplicateCount") or 0
                )
                aggregate_replicate_text = (
                    f", aggregate columns/rows "
                    f"{aggregate_replicates}"
                    if aggregate_replicates
                    else ""
                )
                lines.append(
                    f"  - {series['armLabel']} / "
                    f"{series['outcomeLabel']} / {series_id}: "
                    f"{summary['pointCount']} points{point_breakdown}, "
                    f"range {_format_number(summary['minimum'])}~"
                    f"{_format_number(summary['maximum'])}{value_unit}, "
                    f"descriptive average "
                    f"{_format_number(summary['average'])}{value_unit}; "
                    f"axes {summary['distinctAxisCount']}, "
                    f"raw replicates "
                    f"{summary['distinctReplicateCount']}"
                    f"{aggregate_replicate_text}"
                    f"{review_note} [{study['dataId']}"
                    f"{' / ' + evidence if evidence else ''}]"
                )
    if answer["excludedRecords"]:
        lines.extend(["", "## 정량 결론에서 제외된 기록", ""])
        for record in answer["excludedRecords"]:
            ids = " / ".join(
                value
                for value in (
                    record["dataId"],
                    record["comparisonId"],
                    record["effectId"],
                )
                if value
            )
            reasons = ", ".join(record["reasonCodes"]) or "UNSPECIFIED"
            assessment = record.get("comparisonAssessment")
            if not assessment:
                lines.append(f"- [{ids}] {reasons}")
                continue
            evidence = " / ".join(record["evidenceIds"])
            comparison_label = (
                f"{assessment['comparedArmLabel']} vs "
                f"{assessment['controlArmLabel']}"
            ).strip()
            lines.append(
                f"- [{ids}] {reasons}: {comparison_label} — "
                f"{assessment['textKo']}"
                f"{' [' + evidence + ']' if evidence else ''}"
            )
            if assessment["comparedCondition"] or assessment["controlCondition"]:
                lines.append(
                    "  - 기록 조건: "
                    f"{assessment['comparedCondition'] or '(미기록)'} ↔ "
                    f"{assessment['controlCondition'] or '(미기록)'}"
                )
            if assessment["factorDifferences"]:
                factor_text = "; ".join(
                    (
                        f"{item['factorLabel'] or '(요인명 미기록)'}: "
                        f"{item['controlValue'] or '(미기록)'} → "
                        f"{item['comparedValue'] or '(미기록)'}"
                    )
                    for item in assessment["factorDifferences"]
                )
                lines.append(f"  - 동시 차이 요인: {factor_text}")
            if assessment["matchingBasis"]:
                lines.append(
                    f"  - 비교 정렬 근거: {assessment['matchingBasis']}"
                )
    if answer["excludedSources"]:
        lines.extend(["", "## 표 데이터 없음으로 제외된 원본", ""])
        for record in answer["excludedSources"]:
            reasons = ", ".join(record["reasonCodes"]) or "UNSPECIFIED"
            lines.append(
                f"- [{record['publicAnalysisId']} / "
                f"{record['revisionUid']}] "
                f"{record['fileName']}: {reasons}"
            )
    if answer["citations"]:
        lines.extend(["", "## 원본 근거", ""])
        for citation in answer["citations"]:
            lines.append(
                f"- `{citation['evidenceId']}` — "
                f"`{citation['sourcePath']}` / "
                f"`{citation['sheet']}!{citation['range']}`"
            )
    if answer["limitations"]:
        lines.extend(["", "## 제한 사항", ""])
        for limitation in answer["limitations"]:
            lines.append(
                f"- `{limitation['code']}`: {limitation['textKo']}"
            )
    return "\n".join(lines).rstrip() + "\n"


def build_evidence_answer(
    pack: dict[str, Any],
    *,
    language: str = "ko-KR",
) -> dict[str, Any]:
    """Build a deterministic answer; no AI-authored numeric prose is accepted."""

    if not isinstance(pack, dict):
        raise EvidenceAnswerError("Evidence pack must be a JSON object.")
    if pack.get("schemaVersion") != EVIDENCE_PACK_SCHEMA_VERSION:
        raise EvidenceAnswerError("Unsupported evidence pack schemaVersion.")
    if language != "ko-KR":
        raise EvidenceAnswerError("Only deterministic ko-KR rendering is supported.")
    question = str(pack.get("question") or "").strip()
    if not question:
        raise EvidenceAnswerError("Evidence pack question is required.")
    study_candidates = pack.get("studyCandidates")
    eligible_effects = pack.get("answerEligibleEffects")
    excluded_candidates = pack.get("excludedCandidates")
    source_exclusion_records = pack.get("sourceExclusions", [])
    if not all(
        isinstance(value, list)
        for value in (
            study_candidates,
            eligible_effects,
            excluded_candidates,
            source_exclusion_records,
        )
    ):
        raise EvidenceAnswerError("Evidence pack collections are invalid.")
    study_index = {
        str(candidate["publicDataId"]): candidate
        for candidate in study_candidates
    }
    if len(study_index) != len(study_candidates):
        raise EvidenceAnswerError("Evidence pack has duplicate publicDataId values.")
    citations: dict[str, dict[str, Any]] = {}
    quantitative = _quantitative_groups(pack, study_index, citations)
    observation_descriptive, omitted_uncited_observations = _descriptive_studies(
        pack,
        study_index,
        citations,
    )
    (
        measurement_descriptive,
        omitted_uncited_measurement_series,
    ) = _descriptive_measurement_studies(
        pack,
        study_index,
        citations,
    )
    descriptive = _merge_descriptive_studies(
        observation_descriptive,
        measurement_descriptive,
    )
    excluded = _excluded_records(pack, study_index, citations)
    excluded_sources = _source_exclusions(pack)
    relevant_count = len(study_candidates) + len(excluded_sources)
    eligible_count = len(eligible_effects)
    if relevant_count == 0:
        status = "NO_RELEVANT_DATA"
    elif eligible_count == 0:
        status = "INSUFFICIENT_COMPARISON"
    elif any(
        group["directionStatus"] == "CONFLICTING"
        for group in quantitative
    ):
        status = "CONFLICTING"
    elif excluded or excluded_sources:
        status = "PARTIAL"
    else:
        status = "SUPPORTED"
    represented_ids = {
        effect["dataId"]
        for group in quantitative
        for effect in group["effects"]
    }
    represented_ids.update(study["dataId"] for study in descriptive)
    represented_ids.update(record["dataId"] for record in excluded)
    all_relevant_ids = set(study_index)
    if represented_ids != all_relevant_ids:
        missing = sorted(all_relevant_ids - represented_ids)
        extra = sorted(represented_ids - all_relevant_ids)
        raise EvidenceAnswerError(
            "Evidence answer coverage mismatch; "
            f"missing={missing}, extra={extra}."
        )
    template_code, direct_text = _direct_answer_text(
        status,
        relevant_count,
        quantitative,
        eligible_count,
        excluded,
    )
    answer = {
        "schemaVersion": ANSWER_SCHEMA_VERSION,
        "evidencePackSchemaVersion": EVIDENCE_PACK_SCHEMA_VERSION,
        "evidencePackSha256": evidence_pack_sha256(pack),
        "question": question,
        "language": language,
        "answerStatus": status,
        "directAnswer": {
            "templateCode": template_code,
            "textKo": direct_text,
        },
        "coverage": {
            "relevantStudyCount": len(study_candidates),
            "relevantRecordCount": relevant_count,
            "uniqueEligibleDataCount": len(
                {
                    effect["dataId"]
                    for group in quantitative
                    for effect in group["effects"]
                }
            ),
            "eligibleComparisonCount": len(
                {
                    effect["comparisonId"]
                    for group in quantitative
                    for effect in group["effects"]
                }
            ),
            "eligibleEffectCount": eligible_count,
            "excludedRecordCount": len(excluded),
            "relevantSourceExclusionCount": len(excluded_sources),
            "allRelevantSourceIds": [
                record["publicAnalysisId"]
                for record in excluded_sources
            ],
            "uncitedDescriptiveObservationCount": (
                omitted_uncited_observations
            ),
            "uncitedDescriptiveMeasurementSeriesCount": (
                omitted_uncited_measurement_series
            ),
            "allRelevantDataIds": sorted(all_relevant_ids),
            "representedDataIds": sorted(represented_ids),
        },
        "quantitativeGroups": quantitative,
        "descriptiveStudies": descriptive,
        "excludedRecords": excluded,
        "excludedSources": excluded_sources,
        "limitations": _limitations(
            status=status,
            groups=quantitative,
            descriptive=descriptive,
            excluded=excluded,
            source_exclusions=excluded_sources,
            omitted_uncited_observations=omitted_uncited_observations,
            omitted_uncited_measurement_series=(
                omitted_uncited_measurement_series
            ),
        ),
        "citations": sorted(
            citations.values(),
            key=lambda citation: citation["evidenceId"],
        ),
        "renderedAnswer": {
            "format": "markdown",
            "textKo": "",
        },
    }
    answer["renderedAnswer"]["textKo"] = render_answer_markdown(answer)
    return answer


def validate_evidence_answer(
    answer: dict[str, Any],
    pack: dict[str, Any],
) -> dict[str, Any]:
    """Reject any answer whose wording, values, IDs, or citations were changed."""

    expected = build_evidence_answer(
        pack,
        language=str(answer.get("language") or "ko-KR"),
    )
    if _canonical_bytes(answer) != _canonical_bytes(expected):
        raise EvidenceAnswerError(
            "Evidence answer differs from the deterministic evidence-pack rendering."
        )
    return answer


def answer_json_bytes(answer: dict[str, Any]) -> bytes:
    return (
        json.dumps(
            answer,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
            allow_nan=False,
        )
        + "\n"
    ).encode("utf-8")


__all__ = [
    "ANSWER_SCHEMA_VERSION",
    "EvidenceAnswerError",
    "answer_json_bytes",
    "build_evidence_answer",
    "evidence_pack_sha256",
    "render_answer_markdown",
    "validate_evidence_answer",
]
