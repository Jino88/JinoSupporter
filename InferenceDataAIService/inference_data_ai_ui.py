from __future__ import annotations

"""Drag-and-drop operator UI for the existing Inference Data AI CLI."""

import json
import os
import queue
import shutil
import sqlite3
import subprocess
import sys
import threading
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD


FROZEN = bool(getattr(sys, "frozen", False))
SERVICE_DIR = Path(os.environ["INFERENCE_DATA_AI_SERVICE_DIR"]).resolve() if os.environ.get("INFERENCE_DATA_AI_SERVICE_DIR") else (Path(sys.executable).resolve().parent if FROZEN else Path(__file__).resolve().parent)
CLI_PATH = SERVICE_DIR / "inference_data_ai_cli.py"
ANALYSIS_RUNNER_PATH = SERVICE_DIR / "inference_data_ai_analysis_runner.py"
OUTPUT_DIR = SERVICE_DIR / "outputs"
STATE_PATH = OUTPUT_DIR / "ui-state.json"
HISTORY_PATH = OUTPUT_DIR / "ui-run-history.jsonl"
EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xls", ".xlsb"}
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


@dataclass(frozen=True)
class PipelineOptions:
    input_paths: tuple[str, ...]
    dataset: str
    run_universal: bool = True
    run_quick_index: bool = True
    include_hidden: bool = True
    run_ai_analysis: bool = True
    replace_ai_drafts: bool = True


def safe_name(value: str) -> str:
    return "".join(char if char.isalnum() or char in "-_." else "_" for char in value).strip("._") or "dataset"


def output_paths(dataset: str) -> dict[str, Path]:
    name = safe_name(dataset)
    return {"universal_db": OUTPUT_DIR / "universal-grid" / f"{name}.sqlite", "quick_db": OUTPUT_DIR / "quick-index" / f"{name}.sqlite", "quick_html": OUTPUT_DIR / "quick-index" / f"{name}_dashboard.html"}


def cli_python_executable() -> str:
    configured = os.environ.get("INFERENCE_DATA_AI_PYTHON", "").strip()
    if configured:
        return configured
    if FROZEN:
        runner = shutil.which("python") or shutil.which("py")
        if not runner:
            raise RuntimeError("The packaged UI needs Python to run the existing CLI. Set INFERENCE_DATA_AI_PYTHON to python.exe.")
        return runner
    return sys.executable


def expand_excel_paths(paths: list[str] | tuple[str, ...]) -> list[str]:
    found: dict[str, None] = {}
    for raw in paths:
        path = Path(raw).resolve()
        candidates = [path] if path.is_file() else path.rglob("*") if path.is_dir() else []
        for candidate in candidates:
            if candidate.is_file() and candidate.suffix.lower() in EXCEL_SUFFIXES and not candidate.name.startswith("~$"):
                found[str(candidate.resolve())] = None
    return sorted(found)


def build_commands(options: PipelineOptions) -> list[list[str]]:
    runner, paths = cli_python_executable(), output_paths(options.dataset)
    commands: list[list[str]] = []
    for source in options.input_paths:
        if options.run_universal:
            command = [runner, str(CLI_PATH), "com-index", "--input", source, "--dataset", options.dataset, "--db", str(paths["universal_db"].resolve()), "--covered-cell-mode", "blank", "--verify-after-import"]
            if options.include_hidden:
                command.append("--include-hidden")
            commands.append(command)
        if options.run_quick_index:
            commands.append([runner, str(CLI_PATH), "quick-index", "--input", source, "--dataset", options.dataset, "--db", str(paths["quick_db"].resolve()), "--html", str(paths["quick_html"].resolve())])
        if options.run_ai_analysis:
            command = [runner, str(ANALYSIS_RUNNER_PATH), "--service-dir", str(SERVICE_DIR), "--db", str(paths["universal_db"].resolve()), "--source", source, "--dataset", options.dataset]
            if options.replace_ai_drafts:
                command.append("--replace-auto-draft")
            commands.append(command)
    return commands


def universal_counts(db_path: Path) -> dict[str, int]:
    if not db_path.exists():
        return {}
    try:
        with sqlite3.connect(db_path) as conn:
            tables = {row[0] for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            return {name: int(conn.execute(f"SELECT COUNT(*) FROM {name}").fetchone()[0]) for name in ("workbooks", "analysis_reports") if name in tables}
    except sqlite3.Error:
        return {}


def load_state() -> dict[str, object]:
    try:
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


class InferenceDataAiUi(TkinterDnD.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Inference Data AI Service - Drag Excel Files Here")
        self.geometry("1120x780")
        self.minsize(900, 650)
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.scan_worker: threading.Thread | None = None
        self.process: subprocess.Popen[str] | None = None
        saved = load_state(); options = saved.get("options") if isinstance(saved.get("options"), dict) else {}
        self.dataset_var = tk.StringVar(value=str(options.get("dataset") or "InputDataFinish"))
        self.universal_var = tk.BooleanVar(value=bool(options.get("run_universal", True)))
        self.quick_var = tk.BooleanVar(value=bool(options.get("run_quick_index", True)))
        self.hidden_var = tk.BooleanVar(value=bool(options.get("include_hidden", True)))
        self.analysis_var = tk.BooleanVar(value=bool(options.get("run_ai_analysis", True)))
        self.replace_ai_var = tk.BooleanVar(value=bool(options.get("replace_ai_drafts", True)))
        self.status_var = tk.StringVar(value="Drag Excel files or folders into the list.")
        self.db_var = tk.StringVar(value="Universal DB: not created")
        self.selected_paths: list[str] = expand_excel_paths(list(options.get("input_paths") or []))
        self._build(); self._render_paths(); self._refresh_status(); self.after(100, self._drain_events)

    def _build(self) -> None:
        root = ttk.Frame(self, padding=14); root.pack(fill=tk.BOTH, expand=True); root.columnconfigure(0, weight=1); root.rowconfigure(4, weight=1)
        ttk.Label(root, text="Common Excel Pipeline", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(root, text="Drop Excel files or folders below. The original file paths are used directly; nothing is copied.", foreground="#9a6700").grid(row=1, column=0, sticky="w", pady=(2, 9))
        top = ttk.Frame(root); top.grid(row=2, column=0, sticky="ew"); ttk.Label(top, text="Dataset").pack(side=tk.LEFT); ttk.Entry(top, textvariable=self.dataset_var, width=28).pack(side=tk.LEFT, padx=6)
        ttk.Button(top, text="Add files…", command=self._add_files).pack(side=tk.LEFT, padx=3); ttk.Button(top, text="Add folder…", command=self._add_folder).pack(side=tk.LEFT, padx=3); ttk.Button(top, text="Remove selected", command=self._remove_selected).pack(side=tk.LEFT, padx=12); ttk.Button(top, text="Clear list", command=self._clear).pack(side=tk.LEFT)
        self.drop = tk.Label(root, text="DROP EXCEL FILES OR FOLDERS HERE", relief=tk.GROOVE, bd=2, pady=12, bg="#edf4ff", fg="#175cd3", font=("Segoe UI", 11, "bold"))
        self.drop.grid(row=3, column=0, sticky="ew", pady=(8, 4)); self.drop.drop_target_register(DND_FILES); self.drop.dnd_bind("<<Drop>>", self._drop)
        table = ttk.Frame(root); table.grid(row=4, column=0, sticky="nsew"); table.columnconfigure(0, weight=1); table.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(table, columns=("name", "folder", "size"), show="headings", selectmode="extended"); self.tree.heading("name", text="Excel file"); self.tree.heading("folder", text="Source folder"); self.tree.heading("size", text="Bytes"); self.tree.column("name", width=430); self.tree.column("folder", width=510); self.tree.column("size", width=110, anchor="e")
        scroll = ttk.Scrollbar(table, orient=tk.VERTICAL, command=self.tree.yview); self.tree.configure(yscrollcommand=scroll.set); self.tree.grid(row=0, column=0, sticky="nsew"); scroll.grid(row=0, column=1, sticky="ns")
        controls = ttk.Frame(root); controls.grid(row=5, column=0, sticky="ew", pady=(9, 2)); self.run_button = ttk.Button(controls, text="Run selected list / Resume", command=self._start_run); self.run_button.pack(side=tk.LEFT); self.stop_button = ttk.Button(controls, text="Stop", command=self._stop, state=tk.DISABLED); self.stop_button.pack(side=tk.LEFT, padx=6); ttk.Checkbutton(controls, text="Universal raw DB + COM JSON", variable=self.universal_var).pack(side=tk.LEFT, padx=12); ttk.Checkbutton(controls, text="Quick-index DB + HTML", variable=self.quick_var).pack(side=tk.LEFT); ttk.Checkbutton(controls, text="AI draft + verified HTML", variable=self.analysis_var).pack(side=tk.LEFT, padx=8); ttk.Checkbutton(controls, text="Replace old AI drafts", variable=self.replace_ai_var).pack(side=tk.LEFT); ttk.Checkbutton(controls, text="Include hidden sheets", variable=self.hidden_var).pack(side=tk.LEFT, padx=8); ttk.Button(controls, text="Open outputs", command=lambda: os.startfile(str(OUTPUT_DIR))).pack(side=tk.RIGHT); ttk.Button(controls, text="Open candidate dashboard", command=self._open_html).pack(side=tk.RIGHT, padx=5)
        ttk.Label(root, textvariable=self.db_var).grid(row=6, column=0, sticky="w"); ttk.Label(root, textvariable=self.status_var, foreground="#175cd3").grid(row=7, column=0, sticky="w", pady=(1, 4))
        self.progress = ttk.Progressbar(root, mode="indeterminate"); self.progress.grid(row=8, column=0, sticky="ew", pady=(0, 4))
        self.log = tk.Text(root, height=12, state=tk.DISABLED, wrap=tk.WORD, font=("Consolas", 9)); self.log.grid(row=9, column=0, sticky="ew")

    def _add_files(self) -> None:
        self._add_paths(list(filedialog.askopenfilenames(filetypes=[("Excel", "*.xlsx *.xlsm *.xls *.xlsb")])) )

    def _add_folder(self) -> None:
        folder = filedialog.askdirectory()
        if folder: self._add_paths([folder])

    def _drop(self, event: tk.Event) -> None:
        self._add_paths(list(self.tk.splitlist(event.data)))

    def _add_paths(self, paths: list[str]) -> None:
        if self.scan_worker and self.scan_worker.is_alive():
            self.status_var.set("Excel file scan is already running."); return
        self.status_var.set("Scanning Excel files… the window remains usable.")
        self.scan_worker = threading.Thread(target=self._scan_paths_worker, args=(paths,), daemon=True); self.scan_worker.start()

    def _scan_paths_worker(self, paths: list[str]) -> None:
        try: self.events.put(("paths", expand_excel_paths(paths)))
        except OSError as exc: self.events.put(("scan_error", str(exc)))

    def _remove_selected(self) -> None:
        remove = set(self.tree.selection()); self.selected_paths = [path for path in self.selected_paths if path not in remove]; self._render_paths(); self._refresh_status()

    def _clear(self) -> None:
        self.selected_paths.clear(); self._render_paths(); self._refresh_status()

    def _render_paths(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for path_text in self.selected_paths:
            path = Path(path_text); self.tree.insert("", tk.END, iid=path_text, values=(path.name, str(path.parent), f"{path.stat().st_size:,}"))

    def _refresh_status(self) -> None:
        counts = universal_counts(output_paths(self.dataset_var.get().strip() or "dataset")["universal_db"]); self.db_var.set(f"Selected Excel files: {len(self.selected_paths):,} | Universal DB: {counts.get('workbooks', 0):,} workbooks | Analysis reports: {counts.get('analysis_reports', 0):,}")

    def _options(self) -> PipelineOptions | None:
        if not self.selected_paths: messagebox.showerror("Excel files", "Drag Excel files or folders into the list first."); return None
        if not self.dataset_var.get().strip(): messagebox.showerror("Dataset", "Dataset is required."); return None
        if self.analysis_var.get() and not self.universal_var.get(): messagebox.showerror("AI analysis", "AI analysis requires Universal raw DB + COM JSON."); return None
        if not self.universal_var.get() and not self.quick_var.get(): messagebox.showerror("Outputs", "Select at least one output."); return None
        return PipelineOptions(tuple(self.selected_paths), self.dataset_var.get().strip(), self.universal_var.get(), self.quick_var.get(), self.hidden_var.get(), self.analysis_var.get(), self.replace_ai_var.get())

    def _start_run(self) -> None:
        if self.worker and self.worker.is_alive(): return
        options = self._options()
        if not options: return
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True); STATE_PATH.write_text(json.dumps({"options": asdict(options)}, ensure_ascii=False, indent=2), encoding="utf-8")
        self.run_button.configure(state=tk.DISABLED); self.stop_button.configure(state=tk.NORMAL); self.progress.start(12); self.status_var.set("Pipeline started. UI stays responsive while each step runs."); self.worker = threading.Thread(target=self._run_worker, args=(options,), daemon=True); self.worker.start()

    def _run_worker(self, options: PipelineOptions) -> None:
        started = datetime.now(timezone.utc).isoformat(); commands = build_commands(options); ok = True
        try:
            for command in commands:
                self.events.put(("command", "Running: " + Path(command[1]).name + (" " + command[2] if len(command) > 2 else ""))); self.events.put(("log", "$ " + subprocess.list2cmdline(command) + "\n")); self.process = subprocess.Popen(command, cwd=SERVICE_DIR, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace", creationflags=CREATE_NO_WINDOW)
                assert self.process.stdout is not None
                for line in self.process.stdout: self.events.put(("log", line))
                if self.process.wait() != 0: ok = False; break
        except (OSError, RuntimeError) as exc: ok = False; self.events.put(("log", f"Failed: {exc}\n"))
        finally:
            self.process = None; HISTORY_PATH.parent.mkdir(parents=True, exist_ok=True); HISTORY_PATH.open("a", encoding="utf-8").write(json.dumps({"startedAt": started, "finishedAt": datetime.now(timezone.utc).isoformat(), "ok": ok, "options": asdict(options)}, ensure_ascii=False) + "\n"); self.events.put(("done", ok))

    def _stop(self) -> None:
        if self.process and self.process.poll() is None: self.process.terminate()

    def _drain_events(self) -> None:
        for _ in range(80):
            if self.events.empty(): break
            kind, value = self.events.get_nowait()
            if kind == "log":
                self.log.configure(state=tk.NORMAL); self.log.insert(tk.END, str(value));
                if int(self.log.index("end-1c").split(".")[0]) > 3000: self.log.delete("1.0", "1000.0")
                self.log.see(tk.END); self.log.configure(state=tk.DISABLED)
            elif kind == "command": self.status_var.set(str(value))
            elif kind == "paths":
                existing = set(self.selected_paths); self.selected_paths = sorted(existing | set(value)); self._render_paths(); self._refresh_status(); self.status_var.set(f"Added {len(value):,} Excel file(s).")
            elif kind == "scan_error": self.status_var.set(f"Excel scan failed: {value}")
            else:
                self.run_button.configure(state=tk.NORMAL); self.stop_button.configure(state=tk.DISABLED); self.progress.stop(); self.status_var.set("Pipeline completed." if value else "Pipeline failed or stopped; see log."); self._refresh_status()
        self.after(20 if not self.events.empty() else 100, self._drain_events)

    def _open_html(self) -> None:
        html = output_paths(self.dataset_var.get().strip() or "dataset")["quick_html"]
        if html.exists(): os.startfile(str(html))
        else: messagebox.showinfo("Quick HTML", "Run Quick-index first.")


def main() -> int:
    app = InferenceDataAiUi(); app.mainloop(); return 0


if __name__ == "__main__": raise SystemExit(main())
