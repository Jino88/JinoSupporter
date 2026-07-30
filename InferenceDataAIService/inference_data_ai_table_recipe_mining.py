"""Mine and replay one executable table recipe from historical consensus."""

from __future__ import annotations

import argparse
import json
import os
import re
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable


TABLE_RECIPE_SCHEMA_VERSION = "excel-table-extraction-recipe-v1"
TABLE_REPLAY_REPORT_SCHEMA_VERSION = "excel-table-recipe-replay-v1"
TABLE_RECIPE_MINER_VERSION = "historical-table-consensus-v1.0"

_COLUMN_PATTERN = re.compile(r"^[A-Z]{1,3}$", re.IGNORECASE)
_OPTIONAL_METRIC_PREFIX = re.compile(
    r"^(?:SIGMA|HEARING(?:\s+VOLTAGE\s+UP\s+\d+%?)?)\s*(?:-\s*)?",
    re.IGNORECASE,
)


class TableRecipeMiningError(RuntimeError):
    """Raised when historical members do not support one safe recipe."""


def _read_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as stream:
        value = json.load(stream)
    if not isinstance(value, dict):
        raise TableRecipeMiningError(f"JSON root must be an object: {path}")
    return value


def _write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    temporary.write_text(
        json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        )
        + "\n",
        encoding="utf-8",
    )
    os.replace(temporary, path)


def _column_number(value: str) -> int:
    if _COLUMN_PATTERN.fullmatch(value) is None:
        raise TableRecipeMiningError(f"Invalid column label: {value}")
    result = 0
    for character in value.upper():
        result = result * 26 + ord(character) - ord("A") + 1
    return result


def _metric_core(value: Any) -> str:
    text = " ".join(str(value or "").upper().replace("–", "-").split())
    text = _OPTIONAL_METRIC_PREFIX.sub("", text)
    return text.strip(" -:")


def _table_by_id(request: dict[str, Any], table_id: str) -> dict[str, Any]:
    matches = [
        item
        for item in request.get("tables") or []
        if str(item.get("tableId") or "") == table_id
    ]
    if len(matches) != 1:
        raise TableRecipeMiningError(
            f"Request expected one table {table_id}, found {len(matches)}."
        )
    return matches[0]


def _analysis_table_by_id(
    analysis: dict[str, Any],
    table_id: str,
) -> dict[str, Any]:
    matches = [
        item
        for item in analysis.get("tables") or []
        if str(item.get("tableId") or "") == table_id
    ]
    if len(matches) != 1:
        raise TableRecipeMiningError(
            f"Analysis expected one table {table_id}, found {len(matches)}."
        )
    return matches[0]


def _projection_study(
    projection: dict[str, Any],
    table_id: str,
) -> dict[str, Any]:
    matches = [
        study
        for study in projection.get("studies") or []
        if any(
            str(evidence.get("tableId") or "") == table_id
            for evidence in study.get("evidence") or []
        )
    ]
    if len(matches) != 1:
        raise TableRecipeMiningError(
            f"Projection expected one study for {table_id}, found {len(matches)}."
        )
    return matches[0]


def _axis_columns(
    table: dict[str, Any],
) -> dict[str, tuple[int, dict[str, Any]]]:
    minimum = int(table["bounds"]["minColumn"])
    result: dict[str, tuple[int, dict[str, Any]]] = {}
    for column in table.get("numericColumns") or []:
        column_id = str(column.get("columnId") or "")
        absolute = _column_number(str(column.get("column") or ""))
        result[column_id] = (absolute - minimum, column)
    return result


def _facts_by_column(
    study: dict[str, Any],
) -> dict[str, dict[str, Any]]:
    return {
        str(fact.get("columnId") or ""): fact
        for fact in study.get("deterministicNumericFacts") or []
        if fact.get("columnId")
    }


def _same_number(left: Any, right: Any) -> bool:
    if left is None or right is None:
        return left is right
    try:
        return abs(float(left) - float(right)) <= 1e-12
    except (TypeError, ValueError):
        return left == right


def _stable_group_rules(
    group_examples: list[dict[str, str]],
) -> dict[str, Any]:
    exact: dict[str, Counter[str]] = defaultdict(Counter)
    tokens: dict[str, Counter[str]] = defaultdict(Counter)
    for example in group_examples:
        label = " ".join(example["label"].upper().split())
        role = example["role"]
        exact[label][role] += 1
        for token in re.findall(r"[0-9A-Z가-힣]+", label):
            if len(token) >= 2:
                tokens[token][role] += 1
    exact_rules = [
        {
            "label": label,
            "role": counts.most_common(1)[0][0],
            "supportCount": sum(counts.values()),
        }
        for label, counts in exact.items()
        if len(counts) == 1
    ]
    token_rules = [
        {
            "token": token,
            "role": counts.most_common(1)[0][0],
            "supportCount": sum(counts.values()),
        }
        for token, counts in tokens.items()
        if len(counts) == 1 and sum(counts.values()) >= 2
    ]
    return {
        "exactLabelRules": sorted(
            exact_rules,
            key=lambda item: (
                -int(item["supportCount"]),
                item["label"],
            ),
        ),
        "stableTokenRules": sorted(
            token_rules,
            key=lambda item: (
                -int(item["supportCount"]),
                item["token"],
            ),
        ),
        "unmatchedRole": "UNASSESSED",
    }


def mine_and_replay_table_recipe(
    *,
    table_first_batch_root: str | Path,
    table_structure_catalog_path: str | Path,
    table_structure_id: str,
    generated_at: str | None = None,
) -> dict[str, Any]:
    batch_root = Path(table_first_batch_root).expanduser().resolve()
    catalog_path = Path(table_structure_catalog_path).expanduser().resolve()
    catalog = _read_json(catalog_path)
    matches = [
        item
        for item in catalog.get("structures") or []
        if str(item.get("tableStructureId") or "") == table_structure_id
    ]
    if len(matches) != 1:
        raise TableRecipeMiningError(
            f"Expected one structure {table_structure_id}, found {len(matches)}."
        )
    structure = matches[0]
    members = list(structure.get("members") or [])
    if len(members) < 2:
        raise TableRecipeMiningError("Recipe mining requires repeated history.")

    observations: dict[int, list[dict[str, Any]]] = defaultdict(list)
    semantic_types: Counter[str] = Counter()
    group_examples: list[dict[str, str]] = []
    loaded: list[dict[str, Any]] = []
    for member in members:
        request_file = str(member["requestFile"])
        request = _read_json(batch_root / "requests" / request_file)
        analysis = _read_json(batch_root / "analyses" / request_file)
        projection = _read_json(batch_root / "projections" / request_file)
        table_id = str(member["tableId"])
        source_table = _table_by_id(request, table_id)
        semantic_table = _analysis_table_by_id(analysis, table_id)
        study = _projection_study(projection, table_id)
        axes = _axis_columns(source_table)
        semantic_types[str(semantic_table.get("type") or "UNASSESSED")] += 1
        for group in semantic_table.get("groups") or []:
            label = str(group.get("label") or "").strip()
            role = str(group.get("role") or "UNASSESSED")
            if label:
                group_examples.append({"label": label, "role": role})
        for metric in semantic_table.get("metrics") or []:
            axis_refs = list(metric.get("axisRefs") or [])
            if len(axis_refs) != 1 or axis_refs[0] not in axes:
                raise TableRecipeMiningError(
                    f"Metric axis is not one executable column: {table_id}"
                )
            relative, source_column = axes[axis_refs[0]]
            observations[relative].append(
                {
                    "name": str(metric.get("name") or ""),
                    "coreName": _metric_core(metric.get("name")),
                    "unit": str(metric.get("unit") or ""),
                    "columnRole": str(
                        source_column.get("columnRole") or ""
                    ),
                    "headerTexts": list(
                        source_column.get("headerTexts") or []
                    ),
                }
            )
        loaded.append(
            {
                "member": member,
                "sourceTable": source_table,
                "semanticTable": semantic_table,
                "study": study,
                "axes": axes,
            }
        )

    metric_columns: list[dict[str, Any]] = []
    for relative, values in sorted(observations.items()):
        if len(values) != len(members):
            raise TableRecipeMiningError(
                f"Relative column {relative} has only "
                f"{len(values)}/{len(members)} semantic observations."
            )
        cores = Counter(item["coreName"] for item in values)
        units = Counter(item["unit"] for item in values)
        roles = Counter(item["columnRole"] for item in values)
        if len(cores) != 1 or len(roles) != 1:
            raise TableRecipeMiningError(
                f"Relative column {relative} has conflicting metric semantics."
            )
        metric_columns.append(
            {
                "relativeColumn": relative,
                "canonicalName": cores.most_common(1)[0][0],
                "nameVariants": sorted({item["name"] for item in values}),
                "unit": units.most_common(1)[0][0],
                "unitVariants": sorted(units),
                "columnRole": roles.most_common(1)[0][0],
                "headerTokenVariants": sorted(
                    {
                        " | ".join(str(value) for value in item["headerTexts"])
                        for item in values
                    }
                ),
                "supportCount": len(values),
            }
        )

    dominant_type, dominant_count = semantic_types.most_common(1)[0]
    recipe = {
        "schemaVersion": TABLE_RECIPE_SCHEMA_VERSION,
        "minerVersion": TABLE_RECIPE_MINER_VERSION,
        "recipeId": "table-recipe-" + table_structure_id.removeprefix(
            "table-structure-"
        ),
        "recipeVersion": 1,
        "status": "REPLAY_PENDING",
        "match": {
            "tableStructureId": table_structure_id,
            "fingerprintSha256": structure["fingerprintSha256"],
            "matchMode": "EXACT_TABLE_STRUCTURE",
        },
        "semanticContract": {
            "tableType": dominant_type,
            "tableTypeSupport": dominant_count,
            "historicalMemberCount": len(members),
            "metricColumns": metric_columns,
            "groupRoleRules": _stable_group_rules(group_examples),
        },
        "valueOwnership": {
            "values": "CODE_FROM_CAPTURED_RAW_VALUES",
            "statistics": "CODE_FROM_TABLE_FIRST_REQUEST",
            "evidence": "EXACT_SOURCE_TABLE_AND_COLUMN_RANGE",
            "aiMayWriteValues": False,
        },
    }

    replay_items: list[dict[str, Any]] = []
    expected_by_relative = {
        int(item["relativeColumn"]): item for item in metric_columns
    }
    for item in loaded:
        member = item["member"]
        semantic_table = item["semanticTable"]
        source_table = item["sourceTable"]
        axes = item["axes"]
        study = item["study"]
        actual_semantics: dict[int, str] = {}
        for metric in semantic_table.get("metrics") or []:
            axis_refs = list(metric.get("axisRefs") or [])
            if len(axis_refs) == 1 and axis_refs[0] in axes:
                actual_semantics[axes[axis_refs[0]][0]] = _metric_core(
                    metric.get("name")
                )
        semantic_pass = (
            set(actual_semantics) == set(expected_by_relative)
            and all(
                actual_semantics[relative] == expected["canonicalName"]
                for relative, expected in expected_by_relative.items()
            )
        )

        facts = _facts_by_column(study)
        fact_failures: list[str] = []
        for axis_id, (relative, source_column) in axes.items():
            if relative not in expected_by_relative:
                continue
            fact = facts.get(axis_id)
            if fact is None:
                fact_failures.append(f"missing fact for relative column {relative}")
                continue
            for key in ("numericCount", "min", "max", "average", "sourceRange"):
                source_key = key
                if key == "sourceRange":
                    equal = str(fact.get(key) or "") == str(
                        source_column.get(source_key) or ""
                    )
                else:
                    equal = _same_number(
                        fact.get(key),
                        source_column.get(source_key),
                    )
                if not equal:
                    fact_failures.append(
                        f"relative column {relative} {key} mismatch"
                    )
            if (
                fact.get("calculationAuthority")
                != "CODE_FROM_CAPTURED_RAW_VALUES"
            ):
                fact_failures.append(
                    f"relative column {relative} is not code-owned"
                )
        evidence_pass = any(
            str(evidence.get("tableId") or "") == str(member["tableId"])
            and str(evidence.get("range") or "") == str(source_table.get("range") or "")
            for evidence in study.get("evidence") or []
        )
        replay_items.append(
            {
                "fileName": member["fileName"],
                "requestFile": member["requestFile"],
                "tableId": member["tableId"],
                "sheet": member["sheet"],
                "range": member["range"],
                "semanticPass": semantic_pass,
                "factPass": not fact_failures,
                "evidencePass": evidence_pass,
                "factFailures": fact_failures,
            }
        )

    passed = sum(
        item["semanticPass"] and item["factPass"] and item["evidencePass"]
        for item in replay_items
    )
    replay = {
        "schemaVersion": TABLE_REPLAY_REPORT_SCHEMA_VERSION,
        "minerVersion": TABLE_RECIPE_MINER_VERSION,
        "generatedAt": generated_at
        or datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "recipeId": recipe["recipeId"],
        "tableStructureId": table_structure_id,
        "summary": {
            "historicalMemberCount": len(replay_items),
            "passed": passed,
            "failed": len(replay_items) - passed,
            "semanticPassed": sum(
                bool(item["semanticPass"]) for item in replay_items
            ),
            "factPassed": sum(bool(item["factPass"]) for item in replay_items),
            "evidencePassed": sum(
                bool(item["evidencePass"]) for item in replay_items
            ),
            "aiCalls": 0,
        },
        "items": replay_items,
    }
    if passed == len(replay_items):
        recipe["status"] = "VERIFIED_HISTORICAL_REPLAY"
    return {"recipe": recipe, "replay": replay}


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Mine and replay one repeated table structure recipe."
    )
    parser.add_argument("--table-first-batch-root", required=True)
    parser.add_argument("--table-structure-catalog", required=True)
    parser.add_argument("--table-structure-id", required=True)
    parser.add_argument("--recipe-out", required=True)
    parser.add_argument("--replay-out", required=True)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    arguments = _parser().parse_args(list(argv) if argv is not None else None)
    result = mine_and_replay_table_recipe(
        table_first_batch_root=arguments.table_first_batch_root,
        table_structure_catalog_path=arguments.table_structure_catalog,
        table_structure_id=arguments.table_structure_id,
    )
    recipe_output = Path(arguments.recipe_out).expanduser().resolve()
    replay_output = Path(arguments.replay_out).expanduser().resolve()
    _write_json(recipe_output, result["recipe"])
    _write_json(replay_output, result["replay"])
    print(
        json.dumps(
            {
                "recipeOutput": str(recipe_output),
                "replayOutput": str(replay_output),
                "status": result["recipe"]["status"],
                "summary": result["replay"]["summary"],
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


__all__ = [
    "TABLE_RECIPE_MINER_VERSION",
    "TABLE_RECIPE_SCHEMA_VERSION",
    "TABLE_REPLAY_REPORT_SCHEMA_VERSION",
    "TableRecipeMiningError",
    "mine_and_replay_table_recipe",
]
