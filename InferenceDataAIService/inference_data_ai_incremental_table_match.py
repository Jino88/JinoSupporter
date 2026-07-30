"""Generate AI-free incremental table requests and match historical blocks."""

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

from inference_data_ai_semantic_packets import (
    build_semantic_source_packets,
    connect_capture_v2_readonly,
)
from inference_data_ai_table_first import (
    BUILDER_VERSION,
    build_table_first_request,
    table_first_json_bytes,
)
from inference_data_ai_table_structure_catalog import (
    table_structure_fingerprint,
)


INCREMENTAL_TABLE_MATCH_SCHEMA_VERSION = "excel-incremental-table-match-v1"
INCREMENTAL_TABLE_MATCH_ENGINE_VERSION = "deterministic-table-match-v1.0"


class IncrementalTableMatchError(RuntimeError):
    """Raised when historical or incremental inputs are incompatible."""


def _read_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as stream:
        value = json.load(stream)
    if not isinstance(value, dict):
        raise IncrementalTableMatchError(f"JSON root must be an object: {path}")
    return value


def _atomic_write_bytes(path: Path, value: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    temporary.write_bytes(value)
    os.replace(temporary, path)


def _atomic_write_json(path: Path, value: Any) -> None:
    _atomic_write_bytes(
        path,
        (
            json.dumps(
                value,
                ensure_ascii=False,
                sort_keys=True,
                separators=(",", ":"),
            )
            + "\n"
        ).encode("utf-8"),
    )


def _historical_index(
    catalog: dict[str, Any],
) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for structure in catalog.get("structures") or []:
        digest = str(structure.get("fingerprintSha256") or "")
        if len(digest) != 64:
            raise IncrementalTableMatchError(
                "Historical table structure has an invalid digest."
            )
        if digest in result:
            raise IncrementalTableMatchError(
                f"Duplicate historical table structure digest: {digest}"
            )
        result[digest] = structure
    if not result:
        raise IncrementalTableMatchError("Historical table catalog is empty.")
    return result


def _verified_recipe_index(recipe_root: Path | None) -> dict[str, dict[str, Any]]:
    if recipe_root is None or not recipe_root.is_dir():
        return {}
    result: dict[str, dict[str, Any]] = {}
    for path in sorted(recipe_root.glob("*.json")):
        recipe = _read_json(path)
        if recipe.get("status") != "VERIFIED_HISTORICAL_REPLAY":
            continue
        match = recipe.get("match") or {}
        structure_id = str(match.get("tableStructureId") or "")
        if not structure_id:
            continue
        result[structure_id] = {
            "recipeId": str(recipe.get("recipeId") or ""),
            "recipeVersion": int(recipe.get("recipeVersion") or 0),
            "path": str(path.resolve()),
        }
    return result


def match_table_request(
    request: dict[str, Any],
    historical_structures: dict[str, dict[str, Any]],
    *,
    verified_recipes: dict[str, dict[str, Any]] | None = None,
) -> dict[str, Any]:
    recipes = verified_recipes or {}
    matches: list[dict[str, Any]] = []
    for table in request.get("tables") or []:
        fingerprint = table_structure_fingerprint(table)
        historical = historical_structures.get(
            fingerprint["fingerprintSha256"]
        )
        if historical is None:
            status = "UNMATCHED"
            structure_id = None
            recipe = None
        else:
            structure_id = str(historical["tableStructureId"])
            recipe = recipes.get(structure_id)
            status = (
                "VERIFIED_RECIPE_MATCH"
                if recipe is not None
                else "HISTORICAL_STRUCTURE_MATCH"
            )
        matches.append(
            {
                "tableId": str(table.get("tableId") or ""),
                "sheetIndex": int(table.get("sheetIndex") or 0),
                "sheet": str(table.get("sheet") or ""),
                "range": str(table.get("range") or ""),
                "numericCellCount": int(table.get("numericCellCount") or 0),
                "fingerprintSha256": fingerprint["fingerprintSha256"],
                "status": status,
                "historicalTableStructureId": structure_id,
                "historicalTableCount": (
                    int(historical.get("tableCount") or 0)
                    if historical is not None
                    else 0
                ),
                "historicalWorkbookCount": (
                    int(historical.get("workbookCount") or 0)
                    if historical is not None
                    else 0
                ),
                "dominantSemanticType": (
                    str(historical.get("dominantSemanticType") or "")
                    if historical is not None
                    else ""
                ),
                "semanticConsistency": (
                    float(historical.get("semanticConsistency") or 0.0)
                    if historical is not None
                    else 0.0
                ),
                "verifiedRecipe": recipe,
            }
        )
    return {
        "tableCount": len(matches),
        "exactMatchedTableCount": sum(
            item["status"] != "UNMATCHED" for item in matches
        ),
        "exactMatchedQuantitativeTableCount": sum(
            item["status"] != "UNMATCHED"
            and int(item["numericCellCount"]) > 0
            for item in matches
        ),
        "verifiedRecipeMatchCount": sum(
            item["status"] == "VERIFIED_RECIPE_MATCH"
            for item in matches
        ),
        "tables": matches,
    }


def build_incremental_table_match_report(
    *,
    database_path: str | Path,
    preflight_manifest_path: str | Path,
    historical_table_catalog_path: str | Path,
    output_root: str | Path,
    verified_recipe_root: str | Path | None = None,
    limit: int | None = None,
    generated_at: str | None = None,
) -> dict[str, Any]:
    database = Path(database_path).expanduser().resolve()
    manifest_path = Path(preflight_manifest_path).expanduser().resolve()
    catalog_path = Path(historical_table_catalog_path).expanduser().resolve()
    output = Path(output_root).expanduser().resolve()
    requests_root = output / "requests"
    recipe_root = (
        Path(verified_recipe_root).expanduser().resolve()
        if verified_recipe_root is not None
        else None
    )
    manifest = _read_json(manifest_path)
    catalog = _read_json(catalog_path)
    historical = _historical_index(catalog)
    recipes = _verified_recipe_index(recipe_root)
    items = [
        item
        for item in manifest.get("items") or []
        if int(item.get("captureRevisionId") or 0) > 0
        and str(item.get("status") or "")
        not in {"CAPTURE_FAILED", "EXCLUDED_FORM"}
    ]
    items.sort(
        key=lambda item: (
            str(item.get("relativePath") or ""),
            str(item.get("contentSha256") or ""),
        )
    )
    if limit is not None:
        items = items[: max(int(limit), 0)]

    results: list[dict[str, Any]] = []
    failures: list[dict[str, Any]] = []
    with connect_capture_v2_readonly(database) as connection:
        total = len(items)
        for index, item in enumerate(items, start=1):
            revision_id = int(item["captureRevisionId"])
            try:
                packet = build_semantic_source_packets(
                    connection,
                    revision_id=revision_id,
                )
                revision = packet["inventory"]["sourceRevision"]
                request_path = requests_root / (
                    str(revision["revisionUid"]) + ".json"
                )
                request: dict[str, Any]
                action = "BUILT"
                if request_path.is_file():
                    cached = _read_json(request_path)
                    if (
                        cached.get("builderVersion") == BUILDER_VERSION
                        and str(
                            (cached.get("source") or {}).get(
                                "contentSha256"
                            )
                            or ""
                        )
                        == str(revision["contentSha256"])
                    ):
                        request = cached
                        action = "REUSED"
                    else:
                        request = build_table_first_request(packet)
                        _atomic_write_bytes(
                            request_path,
                            table_first_json_bytes(request),
                        )
                else:
                    request = build_table_first_request(packet)
                    _atomic_write_bytes(
                        request_path,
                        table_first_json_bytes(request),
                    )
                matched = match_table_request(
                    request,
                    historical,
                    verified_recipes=recipes,
                )
                results.append(
                    {
                        "relativePath": str(item.get("relativePath") or ""),
                        "fileName": str(item.get("fileName") or ""),
                        "contentSha256": str(
                            item.get("contentSha256") or ""
                        ),
                        "captureRevisionId": revision_id,
                        "requestPath": str(request_path),
                        "requestAction": action,
                        **matched,
                    }
                )
            except Exception as error:
                failures.append(
                    {
                        "relativePath": str(
                            item.get("relativePath") or ""
                        ),
                        "captureRevisionId": revision_id,
                        "errorType": type(error).__name__,
                        "message": str(error),
                    }
                )
            if index == total or index % 10 == 0:
                print(
                    f"[incremental-table-match] {index}/{total} "
                    f"completed={len(results)} failed={len(failures)}",
                    file=sys.stderr,
                    flush=True,
                )

    matched_workbooks = sum(
        int(item["exactMatchedTableCount"]) > 0 for item in results
    )
    matched_quantitative_workbooks = sum(
        int(item["exactMatchedQuantitativeTableCount"]) > 0
        for item in results
    )
    recipe_workbooks = sum(
        int(item["verifiedRecipeMatchCount"]) > 0 for item in results
    )
    report = {
        "schemaVersion": INCREMENTAL_TABLE_MATCH_SCHEMA_VERSION,
        "engineVersion": INCREMENTAL_TABLE_MATCH_ENGINE_VERSION,
        "generatedAt": generated_at
        or datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "inputs": {
            "databasePath": str(database),
            "preflightManifestPath": str(manifest_path),
            "historicalTableCatalogPath": str(catalog_path),
            "verifiedRecipeRoot": str(recipe_root) if recipe_root else None,
            "eligibleWorkbookCount": len(items),
            "historicalStructureCount": len(historical),
            "verifiedRecipeCount": len(recipes),
            "databaseMode": "READ_ONLY",
            "aiCalls": 0,
            "limit": limit,
        },
        "summary": {
            "eligibleWorkbookCount": len(items),
            "completedWorkbookCount": len(results),
            "failedWorkbookCount": len(failures),
            "tableCount": sum(int(item["tableCount"]) for item in results),
            "exactMatchedTableCount": sum(
                int(item["exactMatchedTableCount"]) for item in results
            ),
            "exactMatchedQuantitativeTableCount": sum(
                int(item["exactMatchedQuantitativeTableCount"])
                for item in results
            ),
            "workbooksWithExactTableMatch": matched_workbooks,
            "workbooksWithExactQuantitativeTableMatch": (
                matched_quantitative_workbooks
            ),
            "verifiedRecipeMatchCount": sum(
                int(item["verifiedRecipeMatchCount"]) for item in results
            ),
            "workbooksWithVerifiedRecipeMatch": recipe_workbooks,
            "requestBuiltCount": sum(
                item["requestAction"] == "BUILT" for item in results
            ),
            "requestReusedCount": sum(
                item["requestAction"] == "REUSED" for item in results
            ),
            "aiCalls": 0,
        },
        "failures": failures,
        "workbooks": results,
    }
    _atomic_write_json(output / "report.json", report)
    return report


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Build and match incremental table-first requests without AI."
    )
    parser.add_argument("--db", required=True)
    parser.add_argument("--preflight-manifest", required=True)
    parser.add_argument("--historical-table-catalog", required=True)
    parser.add_argument("--output-root", required=True)
    parser.add_argument("--verified-recipe-root")
    parser.add_argument("--limit", type=int)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    arguments = _parser().parse_args(list(argv) if argv is not None else None)
    report = build_incremental_table_match_report(
        database_path=arguments.db,
        preflight_manifest_path=arguments.preflight_manifest,
        historical_table_catalog_path=arguments.historical_table_catalog,
        output_root=arguments.output_root,
        verified_recipe_root=arguments.verified_recipe_root,
        limit=arguments.limit,
    )
    print(
        json.dumps(
            {
                "summary": report["summary"],
                "aiCalls": 0,
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


__all__ = [
    "INCREMENTAL_TABLE_MATCH_ENGINE_VERSION",
    "INCREMENTAL_TABLE_MATCH_SCHEMA_VERSION",
    "IncrementalTableMatchError",
    "build_incremental_table_match_report",
    "match_table_request",
]
