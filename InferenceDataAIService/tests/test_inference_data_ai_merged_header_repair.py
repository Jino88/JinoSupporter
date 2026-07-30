from __future__ import annotations

import unittest

import inference_data_ai_merged_header_repair as repair


class MergedHeaderRepairTests(unittest.TestCase):
    def test_splits_aligned_raw_series_and_rewrites_aggregate_reference(
        self,
    ) -> None:
        manifest = {
            "studies": [
                {
                    "measurementSeries": [
                        {
                            "key": "raw",
                            "seriesRole": "RAW",
                            "aggregationFunction": "",
                            "aggregateOfSeries": [],
                            "headerRange": "C9:F9",
                            "valueRange": "C10:F12",
                            "rowIdentityRange": "B10:B12",
                            "stratumKey": "Samples",
                        },
                        {
                            "key": "aggregate",
                            "seriesRole": "AGGREGATE",
                            "aggregationFunction": "AVERAGE",
                            "aggregateOfSeries": ["raw"],
                            "headerRange": "C20:F20",
                            "valueRange": "C21:F21",
                            "rowIdentityRange": "B21",
                        },
                    ]
                }
            ]
        }
        validation_error = (
            "ValueError: studies[0].measurementSeries[0].headerRange has "
            "multiple logical header cells that resolve to the same merged "
            "anchor; cite distinct lower-level header identities"
        )
        target = repair.merged_header_series_repair_target(
            validation_error
        )
        self.assertIsNotNone(target)

        repaired = repair.apply_merged_header_series_repair(
            manifest,
            target,
        )
        series = repaired["studies"][0]["measurementSeries"]

        self.assertEqual(5, len(series))
        self.assertEqual(
            ["C9", "D9", "E9", "F9"],
            [item["headerRange"] for item in series[:4]],
        )
        self.assertEqual(
            ["C10:C12", "D10:D12", "E10:E12", "F10:F12"],
            [item["valueRange"] for item in series[:4]],
        )
        self.assertEqual(
            [item["key"] for item in series[:4]],
            series[4]["aggregateOfSeries"],
        )
        self.assertEqual(
            "C9:F9",
            manifest["studies"][0]["measurementSeries"][0][
                "headerRange"
            ],
        )

    def test_rejects_aggregate_source_series(self) -> None:
        manifest = {
            "studies": [
                {
                    "measurementSeries": [
                        {
                            "key": "aggregate",
                            "seriesRole": "AGGREGATE",
                            "aggregationFunction": "AVERAGE",
                            "aggregateOfSeries": ["raw"],
                            "headerRange": "C9:F9",
                            "valueRange": "C10:F10",
                        }
                    ]
                }
            ]
        }
        target = repair.MergedHeaderRepairTarget(0, 0)

        with self.assertRaisesRegex(
            repair.MergedHeaderRepairError,
            "only non-aggregate RAW",
        ):
            repair.apply_merged_header_series_repair(manifest, target)


if __name__ == "__main__":
    unittest.main()
