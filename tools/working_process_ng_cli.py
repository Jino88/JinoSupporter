"""Put the process names and NG names of a Daily NG Rate sheet into a Working model's Model Info.

Copy the header rows of the sheet in Excel (the process row, the Input / NG items / NG rate (%) row and, if wanted,
the Before / After row) and run:

    python tools/working_process_ng_cli.py MSM-G778B --clipboard            # show what would be saved
    python tools/working_process_ng_cli.py MSM-G778B --clipboard --apply    # save it

or save the copied rows to a text file and pass --file header.txt. Each "N. NAME" cell of the process row becomes a
process (its second line is the local name); the cells between an "Input" and the next one are its NG items, without
"NG rate (%)". The processes go into the manual model group of Model Info (info.manualGroup, the table under
"BMES model group"); the material list is kept, and a process with the same name keeps its type and materials.

Only PortalGNLAB.db WorkingModels is written, the way Save in Model Info does (revision + 1).
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import getpass
import io
import json
import os
import random
import re
import sqlite3
import string
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

DEFAULT_DATA_DIR = Path(r"D:\000. MyWorks\002. DB")
MANUAL = "__manual"
TYPES = ("SUB", "MAIN", "FUNCTION", "VISUAL")

_NUMBER = re.compile(r"^\s*\d+\s*[.)]\s*")
_RATE = re.compile(r"^ng\s*rate\b", re.IGNORECASE)


def _flat(text: str) -> str:
    """Line breaks and runs of spaces as one space."""
    return " ".join(text.split())


def _key(text: str) -> str:
    return _flat(text).lower()


def _rows(text: str) -> list[list[str]]:
    """Tab separated rows as Excel copies them: a cell with a line break comes quoted."""
    rows = list(csv.reader(io.StringIO(text.replace("\r\n", "\n").replace("\r", "\n")), delimiter="\t"))
    return [r for r in rows if any(c.strip() for c in r)]


def _process(cell: str) -> dict[str, str]:
    lines = [_flat(x) for x in cell.split("\n") if x.strip()]
    name = _NUMBER.sub("", lines[0]).strip() if lines else ""
    return {"name": name or lines[0], "localName": " ".join(lines[1:])}


def _items(cells: list[str]) -> list[str]:
    """NG items of a process: every cell but Input and NG rate (%), once each (any case)."""
    out: list[str] = []
    for c in cells:
        t = _flat(c)
        if not t or t.lower() == "input" or _RATE.match(t):
            continue
        if not any(_key(x) == t.lower() for x in out):
            out.append(t)
    return out


def parse_header(text: str) -> tuple[list[dict[str, Any]], list[str]]:
    """The processes (name, localName, ngItems) of the copied header rows, and warnings."""
    rows = _rows(text)
    item_at = next((i for i, r in enumerate(rows) if any(c.strip().lower() == "input" for c in r)), -1)
    if item_at < 1:
        raise ValueError("No row with Input cells under a process row was found. Copy the process row and the "
                         "Input / NG / NG rate (%) row together (tab separated, as Excel copies them).")
    head, items = rows[item_at - 1], rows[item_at]
    first = next(i for i, c in enumerate(items) if c.strip().lower() == "input")
    procs = [(i, c) for i, c in enumerate(head) if c.strip() and i >= first]
    if not procs:
        raise ValueError("The process row has no process names above the Input cells.")

    starts = [i for i, c in enumerate(items) if c.strip().lower() == "input"]
    warnings: list[str] = []
    if len(starts) == len(procs):
        # one Input per process: the NG items are the cells up to the next Input
        spans = [(s, starts[k + 1] if k + 1 < len(starts) else len(items)) for k, s in enumerate(starts)]
    else:
        # some process has no Input (or one more): fall back to the columns under each process cell
        warnings.append(f"{len(procs)} processes but {len(starts)} Input cells; NG items were taken from the "
                        "columns under each process name.")
        cols = [i for i, _ in procs]
        spans = [(c, cols[k + 1] if k + 1 < len(cols) else max(len(items), c + 1)) for k, c in enumerate(cols)]

    out: list[dict[str, Any]] = []
    for (_, cell), (a, b) in zip(procs, spans):
        p = _process(cell)
        p["ngItems"] = _items(items[a:b])
        if not p["ngItems"]:
            warnings.append(f"{p['name']} has no NG item.")
        out.append(p)
    seen: dict[str, int] = {}
    for p in out:
        seen[_key(p["name"])] = seen.get(_key(p["name"]), 0) + 1
    for p in out:
        if seen.get(_key(p["name"]), 0) > 1:
            warnings.append(f"{p['name']} appears {seen.pop(_key(p['name']))} times.")
    return out, warnings


def _new_id() -> str:
    """Like newId() in manual.ts: 8 random characters and the end of the time in base 36."""
    digits = string.digits + string.ascii_lowercase
    t, tail = int(time.time() * 1000), ""
    while t:
        t, r = divmod(t, 36)
        tail = digits[r] + tail
    return "".join(random.choice(digits) for _ in range(8)) + tail[-4:]


def merge_manual(current: Any, parsed: list[dict[str, Any]]) -> tuple[dict[str, Any], list[str]]:
    """
    The manual group with the parsed processes in their order. The material list stays; a process with the same name
    as one already there keeps its id, type and materials. Returns the group and the names of processes dropped.
    """
    g = current if isinstance(current, dict) else {}
    materials = [m for m in g.get("materials", []) if isinstance(m, str)] if isinstance(g.get("materials"), list) else []
    old = [p for p in g.get("processes", []) if isinstance(p, dict)] if isinstance(g.get("processes"), list) else []
    used: set[int] = set()
    processes = []
    for p in parsed:
        at = next((i for i, o in enumerate(old) if i not in used and _key(str(o.get("name", ""))) == _key(p["name"])), -1)
        o = old[at] if at >= 0 else {}
        if at >= 0:
            used.add(at)
        processes.append({
            "id": str(o.get("id") or _new_id()),
            "type": o.get("type") if o.get("type") in TYPES else "SUB",
            "name": p["name"],
            "localName": p["localName"] or str(o.get("localName", "")),
            "ngItems": p["ngItems"],
            "materials": o.get("materials") if isinstance(o.get("materials"), list) else [],
        })
    dropped = [str(o.get("name", "")) for i, o in enumerate(old) if i not in used]
    return {"materials": materials, "processes": processes}, dropped


def read_input(args: argparse.Namespace) -> str:
    if args.clipboard:
        ps = "[Console]::OutputEncoding = [Text.Encoding]::UTF8; Get-Clipboard -Raw"
        r = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True)
        if r.returncode != 0:
            raise SystemExit("Could not read the clipboard: " + r.stderr.decode("utf-8", "replace"))
        return r.stdout.decode("utf-8", "replace")
    if args.file == "-":
        return sys.stdin.read()
    raw = Path(args.file).read_bytes()
    for enc in ("utf-8-sig", "utf-16"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("cp1258", "replace")


def db_path(arg: str | None) -> Path:
    if arg:
        return Path(arg)
    return Path(os.environ.get("PORTALGNLAB_DATA_DIR") or DEFAULT_DATA_DIR) / "PortalGNLAB.db"


def now() -> str:
    """UTC time as .NET writes it with "o" (7 decimals)."""
    return dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%f") + "0Z"


def save(db: Path, model: str, parsed: list[dict[str, Any]], create: bool, user: str) -> dict[str, Any]:
    """Writes the processes into the model's info.manualGroup (one transaction); returns what changed."""
    conn = sqlite3.connect(db, timeout=30)
    try:
        conn.execute("BEGIN IMMEDIATE")
        row = conn.execute("SELECT InfoJson, Revision FROM WorkingModels WHERE ModelName=?", (model,)).fetchone()
        if row is None and not create:
            names = [r[0] for r in conn.execute("SELECT ModelName FROM WorkingModels ORDER BY ModelName")]
            raise SystemExit(f"{model} has no Model Info saved yet. Check the name (saved: {', '.join(names) or 'none'}) "
                             "or pass --create.")
        info = json.loads(row[0]) if row else {}
        if not isinstance(info, dict):
            info = {}
        group, dropped = merge_manual(info.get("manualGroup"), parsed)
        info["manualGroup"] = group
        linked = str(info.get("bmesMaterial") or "")
        if not linked:
            info["bmesMaterial"] = MANUAL
        text = json.dumps(info, ensure_ascii=False, separators=(",", ":"))
        if row is None:
            conn.execute("INSERT INTO WorkingModels (ModelName, InfoJson, Revision, UpdatedAt, UpdatedBy) VALUES (?,?,1,?,?)",
                         (model, text, now(), user))
        else:
            conn.execute("UPDATE WorkingModels SET InfoJson=?, Revision=Revision+1, UpdatedAt=?, UpdatedBy=? WHERE ModelName=?",
                         (text, now(), user, model))
        revision = conn.execute("SELECT Revision FROM WorkingModels WHERE ModelName=?", (model,)).fetchone()[0]
        conn.commit()
        return {"revision": revision, "dropped": dropped, "linked": linked}
    except BaseException:
        conn.rollback()
        raise
    finally:
        conn.close()


def show(model: str, parsed: list[dict[str, Any]], warnings: list[str]) -> None:
    print(f"{model}: {len(parsed)} processes")
    for i, p in enumerate(parsed, 1):
        local = f" ({p['localName']})" if p["localName"] else ""
        print(f"{i:>3}. {p['name']}{local}")
        print(f"       NG: {' | '.join(p['ngItems']) or '-'}")
    for w in warnings:
        print("Warning: " + w)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Put process and NG names of a Daily NG Rate sheet header into Working Model Info.")
    ap.add_argument("model", help="Working model name, as in the model tree")
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--file", help="text file with the copied header rows (tab separated); - reads stdin")
    src.add_argument("--clipboard", action="store_true", help="read the rows copied in Excel")
    ap.add_argument("--apply", action="store_true", help="save into Model Info (without it, only shows the result)")
    ap.add_argument("--create", action="store_true", help="save even if the model has no Model Info row yet")
    ap.add_argument("--db", help="PortalGNLAB.db path (default: PORTALGNLAB_DATA_DIR or D:\\000. MyWorks\\002. DB)")
    ap.add_argument("--json", action="store_true", help="print the parsed processes as JSON")
    args = ap.parse_args(argv)
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    try:
        parsed, warnings = parse_header(read_input(args))
    except ValueError as e:
        print("Error: " + str(e), file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(parsed, ensure_ascii=False, indent=2))
    else:
        show(args.model, parsed, warnings)
    if not args.apply:
        print("Not saved (dry run). Add --apply to save it into Model Info.")
        return 0

    result = save(db_path(args.db), args.model.strip(), parsed, args.create, getpass.getuser())
    print(f"Saved into Model Info of {args.model} (revision {result['revision']}).")
    if result["linked"] and result["linked"] != MANUAL:
        print(f"Note: the model is linked to BMES group {result['linked']}; the processes show after choosing "
              "\"Change link\" > manual entry in Model Info.")
    if result["dropped"]:
        print("Removed processes not in the sheet: " + ", ".join(result["dropped"]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
