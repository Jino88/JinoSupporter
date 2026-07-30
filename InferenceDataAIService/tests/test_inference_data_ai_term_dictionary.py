from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from inference_data_ai_term_dictionary import default_term_dictionary_path


class DefaultTermDictionaryPathTests(unittest.TestCase):
    def test_explicit_override_has_highest_priority(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            override = Path(directory) / "explicit.csv"
            with patch.dict(
                os.environ,
                {
                    "INFERENCE_DATA_AI_TERM_DICTIONARY": str(override),
                    "LOCALAPPDATA": str(Path(directory) / "LocalAppData"),
                },
                clear=False,
            ):
                self.assertEqual(override.resolve(), default_term_dictionary_path())

    def test_existing_workhub_company_glossary_is_preferred(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            local_app_data = Path(directory) / "LocalAppData"
            glossary = (
                local_app_data
                / "WorkHub"
                / "CompanyGlossary"
                / "term_dictionary.csv"
            )
            glossary.parent.mkdir(parents=True)
            glossary.write_text("term_raw,definition_status\nVP,DEFINED\n", encoding="utf-8")

            environment = os.environ.copy()
            environment.pop("INFERENCE_DATA_AI_TERM_DICTIONARY", None)
            environment["LOCALAPPDATA"] = str(local_app_data)
            with patch.dict(os.environ, environment, clear=True):
                self.assertEqual(glossary.resolve(), default_term_dictionary_path())

    def test_missing_workhub_export_keeps_legacy_repository_fallback(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            environment = os.environ.copy()
            environment.pop("INFERENCE_DATA_AI_TERM_DICTIONARY", None)
            environment["LOCALAPPDATA"] = str(Path(directory) / "LocalAppData")
            with patch.dict(os.environ, environment, clear=True):
                resolved = default_term_dictionary_path()

            self.assertEqual("term_dictionary.csv", resolved.name)
            self.assertEqual("db", resolved.parent.name)
            self.assertEqual("MicroSpeaker_ProductTech_DB", resolved.parent.parent.name)


if __name__ == "__main__":
    unittest.main()
