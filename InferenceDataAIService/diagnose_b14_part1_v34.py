"""Run only B14 staged-draft part 1 and preserve validation failures."""

from __future__ import annotations

import json
import time
from pathlib import Path

from inference_data_ai_staged_draft_v2 import (
    build_fragment_envelope,
    chunks_for_part_v2,
    finalize_fragment_envelope,
    locators_for_part_v2,
    registry_for_part,
    select_draft_universe,
)
from inference_data_ai_staged_runner_v2 import (
    run_codex_study_fragment_v2,
)
from inference_data_ai_workflow import (
    _source_identity,
    _workbook_summary,
)


ROOT = Path(__file__).resolve().parent
RUN_DIR = (
    ROOT
    / "outputs"
    / "corpus-ingest"
    / "full-989-v1"
    / "workbooks"
    / "ingest-run_892a52ed759774ad19544fe4"
)
OUTPUT_PATH = (
    RUN_DIR / "draft-parts-v2" / "diagnostic-part1.fragment.json"
)
REPORT_PATH = RUN_DIR / "diagnostic-part1.report.json"


def load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def main() -> int:
    started = time.monotonic()
    packet_set = load_json(RUN_DIR / "semantic-source-packet.json")
    locator_results = [
        load_json(path)
        for path in sorted((RUN_DIR / "locators").glob("*.locator.json"))
    ]
    plan = load_json(RUN_DIR / "study-draft-plan.json")
    registry = load_json(RUN_DIR / "study-registry-v2.json")
    universe = select_draft_universe(
        packet_set=packet_set,
        locator_results=locator_results,
    )
    part = plan["parts"][0]
    focused_chunks = chunks_for_part_v2(universe, part)
    envelope = finalize_fragment_envelope(
        build_fragment_envelope(
            source=_source_identity(packet_set, "InputDataFinish"),
            workbook=_workbook_summary(packet_set),
            plan=plan,
            part=part,
            focused_chunks=focused_chunks,
            locator_results=locators_for_part_v2(universe, part),
            registry_slice=registry_for_part(registry, part),
        )
    )
    report = {
        "partId": part["partId"],
        "promptBytes": envelope["promptBytes"],
        "outputPath": str(OUTPUT_PATH),
        "rejectedPath": str(
            OUTPUT_PATH.with_name(
                OUTPUT_PATH.stem + ".rejected" + OUTPUT_PATH.suffix
            )
        ),
    }
    try:
        result = run_codex_study_fragment_v2(
            envelope=envelope,
            all_selected_chunks=focused_chunks,
            output_path=OUTPUT_PATH,
            reasoning_effort="medium",
            timeout_seconds=1800,
        )
    except Exception as exc:
        report.update(
            {
                "status": "FAILED",
                "errorType": type(exc).__name__,
                "error": str(exc),
            }
        )
        exit_code = 1
    else:
        report.update(
            {
                "status": "SUCCEEDED",
                "recordCount": len(result["records"]),
                "coverageDispositionCount": len(
                    result["coverageDispositions"]
                ),
            }
        )
        exit_code = 0
    report["elapsedSeconds"] = round(time.monotonic() - started, 3)
    REPORT_PATH.write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(report, ensure_ascii=False))
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
