from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

import inference_data_ai_cli as cli
from inference_data_ai_table_first_html import build_table_first_html_report


class TableFirstHtmlTests(unittest.TestCase):
    def test_renders_searchable_index_and_deterministic_workbook_detail(self) -> None:
        with tempfile.TemporaryDirectory(dir=cli.SERVICE_DIR) as temp:
            batch_dir = Path(temp) / "batch"
            output_dir = Path(temp) / "html"
            for folder in ("requests", "analyses", "projections"):
                (batch_dir / folder).mkdir(parents=True, exist_ok=True)

            request_id = "table_request_html_test"
            file_name = "Danger <Book>.xlsx"
            request = {
                "requestId": request_id,
                "source": {
                    "revisionUid": "capture_revision_html_test",
                    "contentSha256": "a" * 64,
                    "fileName": file_name,
                    "sourcePath": r"D:\source\Danger <Book>.xlsx",
                },
                "formulaDerivation": {
                    "status": "DERIVED",
                    "errorCount": 0,
                },
                "tables": [
                    {
                        "tableId": "table_1",
                        "sheet": "Data",
                        "range": "A1:B4",
                        "numericColumns": [
                            {
                                "columnId": "table_1_col_B",
                                "column": "B",
                                "columnRole": "MEASURE_VALUE",
                                "sourceRange": "B2:B4",
                                "numberFormats": ["0.0%"],
                                "displaySamples": [
                                    {
                                        "coordinate": "B2",
                                        "rawNumber": 0.1,
                                        "normalizedDisplay": "10.0%",
                                        "displayScale": "PERCENT",
                                    }
                                ],
                            }
                        ],
                        "numericSeries": [],
                        "previewRows": [
                            {
                                "row": 2,
                                "cells": [
                                    {
                                        "coordinate": "A2",
                                        "kind": "TEXT",
                                        "value": "Before",
                                    },
                                    {
                                        "coordinate": "B2",
                                        "kind": "NUMBER",
                                        "numberFormat": "0.0%",
                                        "value": 0.1,
                                    },
                                ],
                            },
                            {
                                "row": 3,
                                "cells": [
                                    {
                                        "coordinate": "A3",
                                        "kind": "TEXT",
                                        "value": "After",
                                    },
                                    {
                                        "coordinate": "B3",
                                        "kind": "NUMBER",
                                        "numberFormat": "0.0%",
                                        "value": 0.2,
                                    },
                                ],
                            },
                        ],
                    }
                ],
            }
            analysis = {
                "requestId": request_id,
                "revisionUid": "capture_revision_html_test",
                "status": "ANALYZED",
                "workbookSummary": "<script>alert('unsafe')</script> NG comparison",
                "notes": ["Review source wording"],
                "tables": [
                    {
                        "tableId": "table_1",
                        "title": "NG rate comparison",
                        "type": "COMPARISON",
                        "studyGroup": "Before vs After",
                        "confidence": "HIGH",
                        "groups": [
                            {"label": "Before"},
                            {"label": "After"},
                        ],
                        "metrics": [{"name": "NG rate"}],
                        "comparisonRelations": [],
                        "limitations": [],
                        "relatedTableIds": [],
                        "textLinks": [],
                    }
                ],
            }
            projection = {
                "requestId": request_id,
                "analysisStatus": "ANALYZED",
                "verificationStatus": "NEEDS_REVIEW",
                "queryEligibility": "NOT_ELIGIBLE_UNTIL_CANONICAL_REVIEW",
                "source": request["source"],
                "textBlocks": [],
                "studies": [
                    {
                        "studyGroup": "Before vs After",
                        "titles": ["NG rate comparison"],
                        "tableTypes": ["COMPARISON"],
                        "verificationStatus": "NEEDS_REVIEW",
                        "groups": [
                            {
                                "label": "Before",
                                "role": "REFERENCE",
                                "basis": "Source label",
                            },
                            {
                                "label": "After",
                                "role": "TEST",
                                "basis": "Source label",
                            },
                        ],
                        "comparisonRelations": [
                            {
                                "leftGroup": "Before",
                                "rightGroup": "After",
                                "basis": "Explicit comparison",
                            }
                        ],
                        "metrics": [
                            {
                                "name": "NG rate",
                                "unit": "%",
                                "axisRefs": ["table_1_col_B"],
                            }
                        ],
                        "deterministicNumericFacts": [
                            {
                                "columnId": "table_1_col_B",
                                "tableId": "table_1",
                                "sourceRange": "B2:B4",
                                "numericCount": 3,
                                "min": 0.1,
                                "max": 0.2,
                                "average": 0.15,
                            }
                        ],
                        "deterministicNumericSeries": [],
                        "deterministicAggregateChecks": [],
                        "evidence": [
                            {
                                "tableId": "table_1",
                                "sheet": "Data",
                                "range": "A1:B4",
                            }
                        ],
                        "limitations": ["Needs canonical review"],
                    }
                ],
            }
            artifacts = {
                "requests/request.json": request,
                "analyses/analysis.json": analysis,
                "projections/projection.json": projection,
            }
            for relative, value in artifacts.items():
                (batch_dir / relative).write_text(
                    json.dumps(value, ensure_ascii=False),
                    encoding="utf-8",
                )
            report = {
                "schemaVersion": "table-first-batch-report-v1",
                "status": "ok",
                "completedAt": "2026-07-21T00:00:00Z",
                "items": [
                    {
                        "index": 1,
                        "fileName": file_name,
                        "request": "old/location/request.json",
                        "analysis": "old/location/analysis.json",
                        "projection": "old/location/projection.json",
                        "analysisStatus": "ANALYZED",
                        "confidenceCounts": {"HIGH": 1},
                        "studyCount": 1,
                        "tableCount": 1,
                        "metricCount": 1,
                        "comparisonRelationCount": 1,
                        "reviewRecommended": True,
                        "reviewReasons": ["AGGREGATE_MISMATCH"],
                    }
                ],
            }
            (batch_dir / "batch-report.json").write_text(
                json.dumps(report, ensure_ascii=False),
                encoding="utf-8",
            )

            first = build_table_first_html_report(
                batch_dir=batch_dir,
                output_dir=output_dir,
            )
            index_html = (output_dir / "index.html").read_text(encoding="utf-8")
            detail_path = next((output_dir / "workbooks").glob("*.html"))
            detail_html = detail_path.read_text(encoding="utf-8")

            self.assertEqual(1, first["workbookCount"])
            self.assertEqual(3, first["written"])
            self.assertIn("Danger &lt;Book&gt;.xlsx", index_html)
            self.assertIn('id="search"', index_html)
            self.assertNotIn("<script>alert('unsafe')</script>", detail_html)
            self.assertIn("&lt;script&gt;alert", detail_html)
            self.assertIn("Before", detail_html)
            self.assertIn("After", detail_html)
            self.assertIn('class="study-matrix"', detail_html)
            self.assertIn("시험군 / 비교", detail_html)
            self.assertIn("비교: After", detail_html)
            self.assertNotIn("Data!B2:B4", detail_html)
            self.assertNotIn("A1:B4", detail_html)
            self.assertNotIn("Excel 위치", detail_html)
            self.assertNotIn("Excel 근거", detail_html)
            self.assertIn("10%", detail_html)
            self.assertIn("20%", detail_html)
            self.assertNotIn("<h4>비교 관계</h4>", detail_html)
            self.assertNotIn("<h4>지표 및 코드 계산 통계</h4>", detail_html)
            self.assertIn("AGGREGATE_MISMATCH", detail_html)

            second = build_table_first_html_report(
                batch_dir=batch_dir,
                output_dir=output_dir,
            )
            self.assertEqual(0, second["written"])
            self.assertEqual(3, second["reused"])

    def test_cli_exposes_html_renderer(self) -> None:
        args = cli.build_parser().parse_args(
            ["table-first-html", "--batch-dir", "batch"]
        )
        self.assertIs(args.func, cli.cmd_table_first_html)


if __name__ == "__main__":
    unittest.main()
