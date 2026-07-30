"""Read-only adapter for the reviewed micro-speaker term dictionary."""

from __future__ import annotations

import csv
import hashlib
import os
import re
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Any, Iterable, Mapping


ADAPTER_VERSION = "term-dictionary-adapter-v1"


def _term_key(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().casefold()


def default_term_dictionary_path() -> Path:
    override = os.environ.get("INFERENCE_DATA_AI_TERM_DICTIONARY")
    if override:
        return Path(override).expanduser().resolve()
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        workhub_dictionary = (
            Path(local_app_data).expanduser()
            / "WorkHub"
            / "CompanyGlossary"
            / "term_dictionary.csv"
        ).resolve()
        if workhub_dictionary.is_file():
            return workhub_dictionary
    repository_root = Path(__file__).resolve().parents[2]
    return (
        repository_root
        / "MicroSpeaker_ProductTech_DB"
        / "db"
        / "term_dictionary.csv"
    )


@dataclass(frozen=True)
class TermDictionaryAdapter:
    status: str
    source_path: str
    content_sha256: str
    defined_term_count: int
    ignored_terms: tuple[str, ...]
    alias_groups: tuple[dict[str, Any], ...]

    @classmethod
    def from_rows(
        cls,
        rows: Iterable[Mapping[str, object]],
        *,
        source_path: str = "<memory>",
        content_sha256: str = "",
    ) -> "TermDictionaryAdapter":
        defined: list[dict[str, Any]] = []
        ignored: list[str] = []
        for row in rows:
            term = re.sub(r"\s+", " ", str(row.get("term_raw") or "")).strip()
            status = str(row.get("definition_status") or "").strip().upper()
            if not term:
                continue
            if status == "IGNORE":
                ignored.append(term)
                continue
            if status != "DEFINED":
                continue
            normalized_name = re.sub(
                r"\s+",
                " ",
                str(row.get("normalized_name") or ""),
            ).strip()
            if not normalized_name:
                continue
            try:
                source_count = int(row.get("source_count") or 0)
            except (TypeError, ValueError):
                source_count = 0
            defined.append(
                {
                    "term": term,
                    "normalizedName": normalized_name,
                    "sourceCount": source_count,
                }
            )

        by_normalized_name: dict[str, list[dict[str, Any]]] = {}
        for entry in defined:
            by_normalized_name.setdefault(
                _term_key(entry["normalizedName"]),
                [],
            ).append(entry)
        alias_groups: list[dict[str, Any]] = []
        for entries in by_normalized_name.values():
            unique_terms = {
                _term_key(entry["term"]): entry for entry in entries
            }
            if len(unique_terms) < 2:
                continue
            ordered = sorted(
                unique_terms.values(),
                key=lambda entry: (
                    -int(entry["sourceCount"]),
                    _term_key(entry["term"]),
                ),
            )
            alias_groups.append(
                {
                    "canonicalTerm": ordered[0]["term"],
                    "normalizedName": ordered[0]["normalizedName"],
                    "terms": sorted(
                        (entry["term"] for entry in ordered),
                        key=_term_key,
                    ),
                }
            )
        alias_groups.sort(key=lambda group: _term_key(group["normalizedName"]))
        return cls(
            status="LOADED",
            source_path=source_path,
            content_sha256=content_sha256,
            defined_term_count=len(defined),
            ignored_terms=tuple(sorted(set(ignored), key=_term_key)),
            alias_groups=tuple(alias_groups),
        )

    @classmethod
    def missing(cls, path: Path) -> "TermDictionaryAdapter":
        return cls(
            status="MISSING",
            source_path=str(path),
            content_sha256="",
            defined_term_count=0,
            ignored_terms=(),
            alias_groups=(),
        )

    @classmethod
    def from_snapshot(cls, snapshot: object) -> "TermDictionaryAdapter":
        if not isinstance(snapshot, dict):
            return cls.missing(Path("<request>"))
        alias_groups = tuple(
            {
                "canonicalTerm": str(group.get("canonicalTerm") or ""),
                "normalizedName": str(group.get("normalizedName") or ""),
                "terms": [str(term) for term in group.get("terms") or []],
            }
            for group in snapshot.get("aliasGroups") or []
            if isinstance(group, dict)
        )
        return cls(
            status=str(snapshot.get("status") or "MISSING"),
            source_path=str(snapshot.get("sourcePath") or ""),
            content_sha256=str(snapshot.get("contentSha256") or ""),
            defined_term_count=int(snapshot.get("definedTermCount") or 0),
            ignored_terms=tuple(
                str(term) for term in snapshot.get("ignoredTerms") or []
            ),
            alias_groups=alias_groups,
        )

    def snapshot(self) -> dict[str, Any]:
        return {
            "adapterVersion": ADAPTER_VERSION,
            "status": self.status,
            "sourcePath": self.source_path,
            "contentSha256": self.content_sha256,
            "definedTermCount": self.defined_term_count,
            "ignoreTermCount": len(self.ignored_terms),
            "aliasGroupCount": len(self.alias_groups),
            "ignoredTerms": list(self.ignored_terms),
            "aliasGroups": [dict(group) for group in self.alias_groups],
        }

    def is_ignored(self, value: object) -> bool:
        key = _term_key(value)
        return bool(key) and key in {_term_key(term) for term in self.ignored_terms}

    def semantic_key(self, value: object) -> str:
        key = _term_key(value)
        if not key:
            return ""
        for group in self.alias_groups:
            aliases = {
                _term_key(group.get("normalizedName")),
                *(_term_key(term) for term in group.get("terms") or []),
            }
            if key in aliases:
                return "dictionary:" + _term_key(group.get("normalizedName"))
        return key


@lru_cache(maxsize=8)
def _load_cached(
    path_text: str,
    modified_ns: int,
    size: int,
) -> TermDictionaryAdapter:
    del modified_ns, size
    path = Path(path_text)
    payload = path.read_bytes()
    rows = list(csv.DictReader(payload.decode("utf-8-sig").splitlines()))
    return TermDictionaryAdapter.from_rows(
        rows,
        source_path=str(path),
        content_sha256=hashlib.sha256(payload).hexdigest(),
    )


def load_term_dictionary_adapter(
    path: str | Path | None = None,
) -> TermDictionaryAdapter:
    resolved = (
        Path(path).expanduser().resolve()
        if path is not None
        else default_term_dictionary_path()
    )
    if not resolved.is_file():
        return TermDictionaryAdapter.missing(resolved)
    stat = resolved.stat()
    return _load_cached(str(resolved), stat.st_mtime_ns, stat.st_size)


__all__ = [
    "ADAPTER_VERSION",
    "TermDictionaryAdapter",
    "default_term_dictionary_path",
    "load_term_dictionary_adapter",
]
