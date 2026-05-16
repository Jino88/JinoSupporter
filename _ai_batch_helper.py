"""Helper for AI Batch processing per AI_EXCEL_PROC.md schema.

Provides:
 - get_excel_paste(name) -> str
 - get_excel_file(name) -> str  # text-only workbook copy, images removed
 - get_excel_text(name) -> str  # rendered text for every workbook/sheet
 - commit_dataset(name, result, tr_ko, tr_en, tr_vi) -> bool
 - log_failed(name, reason)

`result` is the normalized dict matching AI_EXCEL_PROC.md JSON spec.
Translations tr_xx are dicts containing the narrative fields in that language:
  document: { title, purpose, content }
  conclusions: { conclusion_id: { topic, statement_from_report, normalized_interpretation } }
  hints:       { hint_id: { check_item, reason } }
  log:         { assumptions, warnings, decision_rationale }
"""
from __future__ import annotations
import sqlite3, json, uuid, hashlib, sys, os, io, re, zipfile, xml.etree.ElementTree as ET, datetime as dt

if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DEFAULT_DB_PATH = r'D:\000. MyWorks\002. DB\process-review.db'
DEFAULT_TARGETS_FILE = r'D:\000. MyWorks\005. Program\Repository\JinoSupporter\_batch_selected.txt'
DEFAULT_PROGRESS_FILE = r'D:\000. MyWorks\005. Program\Repository\JinoSupporter\_ai_batch_progress.jsonl'
DEFAULT_FAILED_FILE = r'D:\000. MyWorks\005. Program\Repository\JinoSupporter\_ai_batch_failed.txt'

DB_PATH = os.environ.get('AI_BATCH_DB_PATH') or DEFAULT_DB_PATH
TARGETS_FILE = os.environ.get('AI_BATCH_TARGETS_FILE') or DEFAULT_TARGETS_FILE
PROGRESS_FILE = os.environ.get('AI_BATCH_PROGRESS_FILE') or DEFAULT_PROGRESS_FILE
FAILED_FILE = os.environ.get('AI_BATCH_FAILED_FILE') or DEFAULT_FAILED_FILE


def configure(db_path: str | None = None, targets_file: str | None = None,
              progress_file: str | None = None, failed_file: str | None = None) -> None:
    """Override runtime paths for one batch session."""
    global DB_PATH, TARGETS_FILE, PROGRESS_FILE, FAILED_FILE
    if db_path:
        DB_PATH = db_path
    if targets_file:
        TARGETS_FILE = targets_file
    if progress_file:
        PROGRESS_FILE = progress_file
    if failed_file:
        FAILED_FILE = failed_file


def now_iso() -> str:
    return dt.datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%S')


def open_db() -> sqlite3.Connection:
    con = sqlite3.connect(DB_PATH, timeout=60)
    con.execute('PRAGMA busy_timeout=60000')
    return con


def ensure_schema(con: sqlite3.Connection) -> None:
    cols = {row[1].lower() for row in con.execute('PRAGMA table_info(AiDocuments)').fetchall()}
    if 'generatedreportmarkdown' not in cols:
        con.execute("ALTER TABLE AiDocuments ADD COLUMN GeneratedReportMarkdown TEXT NOT NULL DEFAULT ''")
    tr_cols = {row[1].lower() for row in con.execute('PRAGMA table_info(AiDocumentTranslations)').fetchall()}
    if 'generatedreportmarkdown' not in tr_cols:
        con.execute("ALTER TABLE AiDocumentTranslations ADD COLUMN GeneratedReportMarkdown TEXT NOT NULL DEFAULT ''")


def _generated_report_markdown(result: dict) -> str:
    """Return the AI-authored report markdown from the accepted payload shape."""
    doc = result.get('document') or {}
    generated_report = (
        result.get('generated_report_markdown')
        or result.get('report_markdown')
        or result.get('generated_report')
        or doc.get('generated_report_markdown')
        or ''
    )
    if isinstance(generated_report, dict):
        generated_report = (
            generated_report.get('markdown')
            or generated_report.get('ko')
            or generated_report.get('content')
            or ''
        )
    return str(generated_report or '').strip()


def _translated_report_markdown(tr: dict, lang: str, fallback: str = '') -> str:
    """Return one translated AI-authored report markdown from a translation dict."""
    if not tr:
        return fallback.strip() if lang == 'ko' else ''
    doc = tr.get('document') or {}
    generated_report = (
        tr.get('generated_report_markdown')
        or tr.get('report_markdown')
        or tr.get('generated_report')
        or doc.get('generated_report_markdown')
        or doc.get('report_markdown')
        or doc.get('generated_report')
        or ''
    )
    if isinstance(generated_report, dict):
        generated_report = (
            generated_report.get('markdown')
            or generated_report.get(lang)
            or generated_report.get('content')
            or ''
        )
    text = str(generated_report or '').strip()
    return text or (fallback.strip() if lang == 'ko' else '')


_REPORT_VISUAL_BLOCK_RE = re.compile(
    r"```[ \t]*(report-heatmap|report-bars|report-heatmap-matrix)\b[^\r\n]*\r?\n(.*?)```",
    re.IGNORECASE | re.DOTALL,
)
_REPORT_HEADING_RE = re.compile(r"(?m)^\s*#{2,3}\s+\S+")
_REPORT_TABLE_LINE_RE = re.compile(r"(?m)^\s*\|.+\|\s*$")
_SOURCE_CELL_RE = re.compile(
    r"(?:'[^'\r\n]+'|[A-Za-z0-9_. ()\-]+)![A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?",
    re.IGNORECASE,
)
_BROKEN_TEXT_MARKERS = (
    '\ufffd',
    '?쒗',
    '?먮',
    '?덈',
    '?쇱',
    '寃',
    '紐',
    '怨',
    '沼',
    '梳',
    '휂',
)
_PLACEHOLDER_REPORT_MARKERS = (
    'batch inventory',
    'inventory-only',
    'parameter-only',
    'ng (auto-extracted)',
    'see workbook title/purpose',
    'workbook stored but extraction surfaced narrative only',
)


def _has_result_rows(result: dict) -> bool:
    rows = result.get('results') or []
    return any(isinstance(row, dict) for row in rows)


def _is_valid_visual_rows(value) -> bool:
    return isinstance(value, list) and any(
        isinstance(item, dict)
        and any(str(item.get(key) or '').strip() for key in ('label', 'value', 'detail'))
        for item in value
    )


def _is_valid_heat_matrix(value) -> bool:
    return (
        isinstance(value, dict)
        and isinstance(value.get('columns'), list)
        and len(value.get('columns') or []) > 0
        and isinstance(value.get('rows'), list)
        and len(value.get('rows') or []) > 0
    )


def _has_valid_report_visual(markdown: str) -> bool:
    """Return true only when the AI report includes a renderable visual block."""
    if not markdown:
        return False

    for kind, body in _REPORT_VISUAL_BLOCK_RE.findall(str(markdown)):
        try:
            parsed = json.loads(body.strip())
        except Exception:
            continue

        normalized = kind.strip().lower()
        if normalized == 'report-heatmap' and (_is_valid_visual_rows(parsed) or _is_valid_heat_matrix(parsed)):
            return True
        if normalized == 'report-bars' and _is_valid_visual_rows(parsed):
            return True
        if normalized == 'report-heatmap-matrix' and _is_valid_heat_matrix(parsed):
            return True

    return False


def _report_quality_issues(markdown: str, has_results: bool) -> list[str]:
    """Detect reports that are too thin to be shown as AI-authored analysis."""
    text = str(markdown or '').strip()
    issues: list[str] = []
    min_chars = 1400 if has_results else 700
    min_headings = 6 if has_results else 4

    if len(text) < min_chars:
        issues.append(f'too short ({len(text)} chars; minimum {min_chars})')

    headings = _REPORT_HEADING_RE.findall(text)
    if len(headings) < min_headings:
        issues.append(f'too few report sections ({len(headings)}; minimum {min_headings})')

    if has_results:
        table_lines = _REPORT_TABLE_LINE_RE.findall(text)
        if len(table_lines) < 3:
            issues.append('missing compact markdown evidence/result table')
        if not _has_valid_report_visual(text):
            issues.append('missing valid AI-authored visual block')

    if not _SOURCE_CELL_RE.search(text):
        issues.append('missing auditable worksheet/cell evidence')

    lower = text.lower()
    for marker in _PLACEHOLDER_REPORT_MARKERS:
        if marker in lower:
            issues.append(f'placeholder report marker: {marker}')
            break

    for marker in _BROKEN_TEXT_MARKERS:
        if marker in text:
            issues.append('broken encoding/mojibake detected')
            break

    return issues


def load_targets(path: str | None = None) -> list[str]:
    target_path = path or os.environ.get('AI_BATCH_TARGETS_FILE') or TARGETS_FILE
    with open(target_path, 'r', encoding='utf-8-sig') as f:
        return [l.rstrip('\r\n') for l in f if l.strip()]


def get_excel_paste(con: sqlite3.Connection, name: str) -> str | None:
    row = con.execute(
        "SELECT ExtractedText FROM RawReportText WHERE DatasetName=? AND Kind='excel_paste'",
        (name,),
    ).fetchone()
    return row[0] if row and row[0] else None


def _short_id(seed: str) -> str:
    return hashlib.sha1(seed.encode('utf-8')).hexdigest()[:16]


def _safe_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', '_', name).strip(' .')
    return name[:120] or 'workbook'


_TEXT_ONLY_SKIP_PREFIXES = (
    'xl/media/',
    'xl/drawings/',
    'xl/charts/',
    'xl/embeddings/',
    'xl/activeX/',
    'xl/ctrlProps/',
)
_TEXT_ONLY_SKIP_TARGET_HINTS = (
    '/media/', '../media/',
    '/drawings/', '../drawings/',
    '/charts/', '../charts/',
    '/embeddings/', '../embeddings/',
    '/activeX/', '../activeX/',
    '/ctrlProps/', '../ctrlProps/',
)
_TEXT_ONLY_SKIP_REL_TYPES = (
    '/image',
    '/drawing',
    '/chart',
    '/chartUserShapes',
    '/vmlDrawing',
    '/oleObject',
    '/control',
)
_WORKSHEET_VISUAL_NODES = {
    'drawing',
    'legacyDrawing',
    'legacyDrawingHF',
    'picture',
    'oleObjects',
    'controls',
}


def _local_name(tag: str) -> str:
    return tag.rsplit('}', 1)[-1] if '}' in tag else tag


def _is_text_only_skipped_part(name: str) -> bool:
    normalized = name.replace('\\', '/')
    if normalized.startswith(_TEXT_ONLY_SKIP_PREFIXES):
        return True
    if normalized.startswith('docProps/thumbnail.'):
        return True
    return False


def _xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)


def _strip_visual_nodes(xml_data: bytes) -> bytes:
    root = ET.fromstring(xml_data)
    for parent in list(root.iter()):
        for child in list(parent):
            if _local_name(child.tag) in _WORKSHEET_VISUAL_NODES:
                parent.remove(child)
    return _xml_bytes(root)


def _strip_visual_relationships(xml_data: bytes) -> bytes:
    root = ET.fromstring(xml_data)
    for child in list(root):
        rel_type = child.attrib.get('Type', '')
        target = child.attrib.get('Target', '')
        if any(rel_type.endswith(t) or t in rel_type for t in _TEXT_ONLY_SKIP_REL_TYPES) \
                or any(h in target for h in _TEXT_ONLY_SKIP_TARGET_HINTS):
            root.remove(child)
    return _xml_bytes(root)


def _strip_visual_content_types(xml_data: bytes) -> bytes:
    root = ET.fromstring(xml_data)
    for child in list(root):
        part_name = child.attrib.get('PartName', '').lstrip('/').replace('\\', '/')
        extension = child.attrib.get('Extension', '').lower()
        if _is_text_only_skipped_part(part_name) or extension in {
            'png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'emf', 'wmf', 'tif', 'tiff'
        }:
            root.remove(child)
    return _xml_bytes(root)


def _write_text_only_workbook(file_data: bytes, out_path: str) -> None:
    """Write a workbook copy with images/drawings removed.

    This intentionally does not flatten cells with openpyxl. Worksheet XML,
    shared strings, styles, formulas, dimensions, tables, and mergeCells remain
    in the workbook package so AI sees the Excel cell structure without heavy
    embedded media.
    """
    try:
        with zipfile.ZipFile(io.BytesIO(file_data), 'r') as zin:
            with zipfile.ZipFile(out_path, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zout:
                for info in zin.infolist():
                    name = info.filename.replace('\\', '/')
                    if _is_text_only_skipped_part(name):
                        continue

                    data = zin.read(info.filename)
                    try:
                        if name == '[Content_Types].xml':
                            data = _strip_visual_content_types(data)
                        elif name.startswith('xl/worksheets/') and name.endswith('.xml'):
                            data = _strip_visual_nodes(data)
                        elif name.endswith('.rels'):
                            data = _strip_visual_relationships(data)
                    except Exception:
                        # Keep the original XML if a cleanup step cannot parse a
                        # vendor-specific extension. It is better to keep a valid
                        # workbook than to drop sheet content.
                        pass

                    zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                    zi.compress_type = zipfile.ZIP_DEFLATED
                    zi.external_attr = info.external_attr
                    zi.comment = info.comment
                    zout.writestr(zi, data)
    except zipfile.BadZipFile:
        with open(out_path, 'wb') as f:
            f.write(file_data)


def get_excel_file(con: sqlite3.Connection, name: str, out_dir: str | None = None) -> str | None:
    """Materialize a text-only workbook copy for a dataset and return its path.

    The returned XLSX/XLSM keeps cell XML, sheet structure, styles, formulas and
    merged cells, but removes embedded media/drawings/charts to keep AI payloads
    small. Fall back to get_excel_paste only when this returns None.
    """
    rows = con.execute(
        """
        SELECT FileName, MediaType, FileData
        FROM RawReportFiles
        WHERE DatasetName=?
        ORDER BY Id DESC
        """,
        (name,),
    ).fetchall()
    if out_dir is None:
        out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_batch_excel_files')
    os.makedirs(out_dir, exist_ok=True)

    for file_name, media_type, file_data in rows:
        file_name = file_name or name
        media_type = (media_type or '').lower()
        lower_name = file_name.lower()
        is_workbook = (
            lower_name.endswith(('.xlsx', '.xlsm', '.xls'))
            or 'spreadsheet' in media_type
            or 'excel' in media_type
        )
        if not is_workbook or not file_data:
            continue

        base, ext = os.path.splitext(file_name)
        if not ext:
            ext = '.xlsx'
        safe = _safe_filename(base)
        path = os.path.join(out_dir, f'{_short_id(name + "|" + file_name)}_{safe}_textonly{ext.lower()}')
        _write_text_only_workbook(bytes(file_data), path)
        return path

    return None


def get_excel_files(con: sqlite3.Connection, name: str, out_dir: str | None = None) -> list[str]:
    """Materialize every workbook attached to a dataset and return all paths."""
    rows = con.execute(
        """
        SELECT FileName, MediaType, FileData
        FROM RawReportFiles
        WHERE DatasetName=?
        ORDER BY Id
        """,
        (name,),
    ).fetchall()
    if out_dir is None:
        out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_batch_excel_files')
    os.makedirs(out_dir, exist_ok=True)

    paths: list[str] = []
    for file_name, media_type, file_data in rows:
        file_name = file_name or name
        media_type = (media_type or '').lower()
        lower_name = file_name.lower()
        is_workbook = (
            lower_name.endswith(('.xlsx', '.xlsm', '.xls'))
            or 'spreadsheet' in media_type
            or 'excel' in media_type
        )
        if not is_workbook or not file_data:
            continue

        base, ext = os.path.splitext(file_name)
        if not ext:
            ext = '.xlsx'
        safe = _safe_filename(base)
        path = os.path.join(out_dir, f'{_short_id(name + "|" + file_name)}_{safe}_textonly{ext.lower()}')
        _write_text_only_workbook(bytes(file_data), path)
        paths.append(path)
    return paths


def get_excel_text(con: sqlite3.Connection, name: str) -> str | None:
    """Render all attached workbooks and all sheets into compact cell-coordinate text.

    This is the preferred AI input for Dataset Results. It uses the text-only
    workbook copy, then renders every worksheet in workbook order. No workbook-
    level character cap is applied here; if a file is too large for one prompt,
    process the returned text sheet-by-sheet instead of dropping later sheets.
    """
    paths = get_excel_files(con, name)
    if not paths:
        return None

    try:
        import _xlsx_render
    except Exception:
        _xlsx_render = None

    parts: list[str] = []
    for path in paths:
        parts.append(f'### WORKBOOK: {os.path.basename(path)}')
        try:
            if _xlsx_render is None:
                raise RuntimeError('_xlsx_render import failed')
            parts.append(_xlsx_render.render_workbook(path))
        except Exception as exc:
            parts.append(f'[WORKBOOK_RENDER_FAILED] {exc}')
    return '\n\n'.join(parts)


def _is_inventory_only_result(result: dict) -> bool:
    """Reject fake batch results that only summarize workbook/sheet inventory.

    The batch output must contain actual report extraction. A previous runner
    generated "batch inventory" placeholders and overwrote real analysis rows;
    block that class of payload before any existing DB rows are deleted.
    """
    doc = result.get('document') or {}
    purpose = str(doc.get('purpose') or '').lower()
    if 'capture sheet inventory' in purpose or 'batch inventory' in purpose:
        return True

    log = result.get('ai_extraction_log') or {}
    joined_log = json.dumps(log, ensure_ascii=False).lower()
    if 'inventory-only' in joined_log or 'inventory only' in joined_log:
        return True

    conditions = result.get('test_conditions') or []
    if any('batch inventory' in str(c.get('changed_factor') or '').lower() for c in conditions if isinstance(c, dict)):
        return True

    results = [r for r in (result.get('results') or []) if isinstance(r, dict)]
    if results and all(str(r.get('measurement_type') or '').lower() == 'inventory' for r in results):
        return True

    troubleshooting = result.get('troubleshooting_index') or {}
    joined_trouble = json.dumps(troubleshooting, ensure_ascii=False).lower()
    return 'batch inventory only' in joined_trouble


def _auto_placeholder_reason(result: dict) -> str | None:
    """Reject heuristic fallback payloads before they overwrite real rows."""
    doc = result.get('document') or {}
    primary = doc.get('primary_defect')
    if isinstance(primary, dict):
        primary_text = str(primary.get('canonical_name') or '')
    else:
        primary_text = str(primary or '')
    if 'auto-extracted' in primary_text.lower():
        return 'rejected auto-extracted placeholder primary defect'

    text = json.dumps(result, ensure_ascii=False).lower()
    markers = (
        '_batch_auto.py',
        'ng (auto-extracted)',
        'auto_extracted',
        'see workbook title/purpose',
        'workbook stored but extraction surfaced narrative only',
        'extraction surfaced narrative only',
    )
    for marker in markers:
        if marker in text:
            return f'rejected fallback placeholder marker: {marker}'

    return None


def commit_dataset(name: str, result: dict, tr_ko: dict, tr_en: dict, tr_vi: dict) -> bool:
    """Insert normalized result + 3-lang translations for one dataset.

    Wipes any prior rows for this SourceDataset (idempotent), then INSERTs.
    Returns True on success, False on failure.
    """
    generated_report = _generated_report_markdown(result)
    if not generated_report:
        reason = 'rejected parameter-only result: generated_report_markdown is required'
        try:
            log_failed(name, reason)
        except Exception:
            pass
        print(f'[REJECTED] {name}: {reason}', file=sys.stderr)
        return False

    translated_reports = {
        'ko': _translated_report_markdown(tr_ko, 'ko', ''),
        'en': _translated_report_markdown(tr_en, 'en', ''),
        'vi': _translated_report_markdown(tr_vi, 'vi', ''),
    }
    missing_report_langs = [lang for lang, text in translated_reports.items() if not text]
    if missing_report_langs:
        reason = 'rejected incomplete report translations: missing generated_report_markdown for ' + ','.join(missing_report_langs)
        try:
            log_failed(name, reason)
        except Exception:
            pass
        print(f'[REJECTED] {name}: {reason}', file=sys.stderr)
        return False

    has_results = _has_result_rows(result)
    quality_targets = [('result', generated_report)] + list(translated_reports.items())
    quality_failures = []
    for label, report_text in quality_targets:
        issues = _report_quality_issues(report_text, has_results)
        if issues:
            quality_failures.append(f'{label}: ' + '; '.join(issues[:4]))
    if quality_failures:
        reason = 'rejected low-quality generated report: ' + ' | '.join(quality_failures)
        try:
            log_failed(name, reason)
        except Exception:
            pass
        print(f'[REJECTED] {name}: {reason}', file=sys.stderr)
        return False

    if has_results:
        missing_visual_reports = []
        if not _has_valid_report_visual(generated_report):
            missing_visual_reports.append('result')
        missing_visual_reports.extend(
            lang for lang, text in translated_reports.items()
            if not _has_valid_report_visual(text)
        )
        if missing_visual_reports:
            reason = (
                'rejected report without AI-authored visual block: missing valid '
                'report-heatmap/report-bars/report-heatmap-matrix for '
                + ','.join(missing_visual_reports)
            )
            try:
                log_failed(name, reason)
            except Exception:
                pass
            print(f'[REJECTED] {name}: {reason}', file=sys.stderr)
            return False

    if _is_inventory_only_result(result):
        reason = 'rejected inventory-only batch result; run real AI extraction instead'
        try:
            log_failed(name, reason)
        except Exception:
            pass
        print(f'[REJECTED] {name}: {reason}', file=sys.stderr)
        return False
    placeholder_reason = _auto_placeholder_reason(result)
    if placeholder_reason:
        try:
            log_failed(name, placeholder_reason)
        except Exception:
            pass
        print(f'[REJECTED] {name}: {placeholder_reason}', file=sys.stderr)
        return False

    con = open_db()
    try:
        ensure_schema(con)
        cur = con.cursor()
        cur.execute('BEGIN')

        # 1. Wipe prior rows for this dataset.
        prior = cur.execute('SELECT DocumentId FROM AiDocuments WHERE SourceDataset=?', (name,)).fetchall()
        for (doc_id,) in prior:
            cond_ids = [r[0] for r in cur.execute('SELECT ConditionId FROM AiTestConditions WHERE DocumentId=?', (doc_id,)).fetchall()]
            res_ids = [r[0] for r in cur.execute('SELECT ResultId FROM AiResults WHERE DocumentId=?', (doc_id,)).fetchall()]
            concl_ids = [r[0] for r in cur.execute('SELECT ConclusionId FROM AiConclusions WHERE DocumentId=?', (doc_id,)).fetchall()]
            hint_ids = [r[0] for r in cur.execute('SELECT HintId FROM AiTroubleshootingHints WHERE DocumentId=?', (doc_id,)).fetchall()]
            log_ids = [r[0] for r in cur.execute('SELECT LogId FROM AiExtractionLogs WHERE DocumentId=?', (doc_id,)).fetchall()]
            for rid in res_ids:
                cur.execute('DELETE FROM AiNgBreakdowns WHERE ResultId=?', (rid,))
            cur.execute('DELETE FROM AiResults WHERE DocumentId=?', (doc_id,))
            cur.execute('DELETE FROM AiTestConditions WHERE DocumentId=?', (doc_id,))
            for cid in concl_ids:
                cur.execute('DELETE FROM AiConclusionTranslations WHERE ConclusionId=?', (cid,))
            cur.execute('DELETE FROM AiConclusions WHERE DocumentId=?', (doc_id,))
            for hid in hint_ids:
                cur.execute('DELETE FROM AiHintTranslations WHERE HintId=?', (hid,))
            cur.execute('DELETE FROM AiTroubleshootingHints WHERE DocumentId=?', (doc_id,))
            for lid in log_ids:
                cur.execute('DELETE FROM AiLogTranslations WHERE LogId=?', (lid,))
            cur.execute('DELETE FROM AiExtractionLogs WHERE DocumentId=?', (doc_id,))
            cur.execute('DELETE FROM AiDocumentTranslations WHERE DocumentId=?', (doc_id,))
            cur.execute('DELETE FROM AiDocuments WHERE DocumentId=?', (doc_id,))

        # 2. Generate stable ids based on dataset name.
        doc_id = 'doc_' + _short_id(name)
        doc = result.get('document', {})
        primary = doc.get('primary_defect') or {}
        if isinstance(primary, dict):
            primary_name = primary.get('canonical_name') or ''
            primary_json = json.dumps(primary, ensure_ascii=False)
        else:
            primary_name = str(primary)
            primary_json = json.dumps({'canonical_name': primary_name, 'aliases_in_document': []}, ensure_ascii=False)

        now = now_iso()
        cur.execute("""
            INSERT INTO AiDocuments
              (DocumentId, SourceDataset, SourceFile, Title, Model, ReportDate,
               Department, Marker, Line, ReportType, PrimaryDefect, PrimaryDefectJson,
               RelatedDefectsJson, PartsJson, ProcessesJson, Purpose, ContentJson,
               GeneratedReportMarkdown, SourceCellsJson, Confidence, SchemaVersion, RawJson, CreatedAt, UpdatedAt)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            doc_id, name, doc.get('source_file', name), doc.get('title', '') or '',
            doc.get('model', '') or '', doc.get('report_date', '') or '',
            doc.get('department', '') or '', doc.get('marker', '') or '',
            doc.get('line', '') or '', doc.get('report_type', '') or '',
            primary_name, primary_json,
            json.dumps(doc.get('related_defects') or [], ensure_ascii=False),
            json.dumps(doc.get('parts') or [], ensure_ascii=False),
            json.dumps(doc.get('processes') or [], ensure_ascii=False),
            doc.get('purpose', '') or '',
            json.dumps(doc.get('content') or [], ensure_ascii=False),
            generated_report,
            json.dumps(doc.get('source_cells') or {}, ensure_ascii=False),
            float(result.get('ai_extraction_log', {}).get('confidence') or 0.0),
            result.get('schema_version', '0.1'),
            json.dumps(result, ensure_ascii=False),
            now, now,
        ))

        # 3. Test conditions.
        cond_id_map = {}
        for idx, c in enumerate(result.get('test_conditions') or []):
            cid_raw = c.get('condition_id') or f'cond_{idx}'
            cid = 'cnd_' + _short_id(name + '|' + cid_raw + '|' + str(idx))
            cond_id_map[cid_raw] = cid
            cur.execute("""
                INSERT INTO AiTestConditions
                  (ConditionId, DocumentId, ConditionGroup, Line, Process, ChangedFactor,
                   BeforeValue, AfterValue, Unit, Machine, Jig, MaterialLot, Supplier,
                   DryTimeSec, Temperature, Pressure, BondAmount, UvEnergy,
                   SourceFile, SheetName, SourceCellsJson)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                cid, doc_id, c.get('condition_group', '') or '', c.get('line', '') or '',
                c.get('process', '') or '', c.get('changed_factor', '') or '',
                str(c.get('before_value')) if c.get('before_value') is not None else None,
                str(c.get('after_value')) if c.get('after_value') is not None else None,
                c.get('unit'), c.get('machine'), c.get('jig'), c.get('material_lot'),
                c.get('supplier'),
                float(c['dry_time_sec']) if c.get('dry_time_sec') is not None else None,
                str(c.get('temperature')) if c.get('temperature') is not None else None,
                str(c.get('pressure')) if c.get('pressure') is not None else None,
                str(c.get('bond_amount')) if c.get('bond_amount') is not None else None,
                str(c.get('uv_energy')) if c.get('uv_energy') is not None else None,
                c.get('source_file', name) or name, c.get('sheet_name', '') or '',
                json.dumps(c.get('source_cells') or [], ensure_ascii=False),
            ))

        # 4. Results + breakdowns.
        for idx, r in enumerate(result.get('results') or []):
            rid_raw = r.get('result_id') or f'res_{idx}'
            rid = 'rid_' + _short_id(name + '|' + rid_raw + '|' + str(idx))
            cond_raw = r.get('condition_id') or ''
            cond_id = cond_id_map.get(cond_raw)
            cur.execute("""
                INSERT INTO AiResults
                  (ResultId, DocumentId, ConditionId, MeasurementType, ConditionGroup,
                   ResultDate, Line, InputCount, OkCount, NgCount, NgRateDecimal,
                   NgRatePercent, MetricName, MetricValue, Unit, Judgement,
                   SourceFile, SheetName, SourceCellsJson)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                rid, doc_id, cond_id, r.get('measurement_type', '') or '',
                r.get('condition_group', '') or '', r.get('date', '') or '',
                r.get('line', '') or '',
                float(r['input_count']) if r.get('input_count') is not None else None,
                float(r['ok_count']) if r.get('ok_count') is not None else None,
                float(r['ng_count']) if r.get('ng_count') is not None else None,
                float(r['ng_rate_decimal']) if r.get('ng_rate_decimal') is not None else None,
                float(r['ng_rate_percent']) if r.get('ng_rate_percent') is not None else None,
                r.get('metric_name', '') or '',
                float(r['metric_value']) if r.get('metric_value') is not None else None,
                r.get('unit'), r.get('judgement'),
                r.get('source_file', name) or name, r.get('sheet_name', '') or '',
                json.dumps(r.get('source_cells') or [], ensure_ascii=False),
            ))
            ng_breakdown = r.get('ng_breakdown') or {}
            if isinstance(ng_breakdown, dict):
                for bidx, (defect, val) in enumerate(ng_breakdown.items()):
                    bid = 'bd_' + _short_id(rid + '|' + defect + '|' + str(bidx))
                    if isinstance(val, dict):
                        cnt = val.get('count')
                        rate = val.get('rate')
                    else:
                        cnt = val
                        rate = None
                    cur.execute("""
                        INSERT INTO AiNgBreakdowns (BreakdownId, ResultId, DefectName, DefectCount, DefectRate)
                        VALUES (?,?,?,?,?)
                    """, (bid, rid, str(defect),
                          float(cnt) if cnt is not None else None,
                          float(rate) if rate is not None else None))

        # 5. Conclusions + translations.
        concl_id_map = {}
        for idx, c in enumerate(result.get('conclusions') or []):
            cid_raw = c.get('conclusion_id') or f'concl_{idx}'
            cid = 'cncl_' + _short_id(name + '|' + cid_raw + '|' + str(idx))
            concl_id_map[cid_raw] = cid
            cur.execute("""
                INSERT INTO AiConclusions
                  (ConclusionId, DocumentId, Topic, StatementFromReport,
                   NormalizedInterpretation, SourceFile, SheetName, SourceCellsJson)
                VALUES (?,?,?,?,?,?,?,?)
            """, (cid, doc_id, c.get('topic', '') or '',
                  c.get('statement_from_report', '') or '',
                  c.get('normalized_interpretation', '') or '',
                  c.get('source_file', name) or name,
                  c.get('sheet_name', '') or '',
                  json.dumps(c.get('source_cells') or [], ensure_ascii=False)))

        # 6. Troubleshooting hints + translations.
        ts = result.get('troubleshooting_index') or {}
        defect_name = ts.get('defect_name', '') or primary_name
        hint_id_map = {}
        for idx, h in enumerate(ts.get('suggested_checks') or []):
            hid_raw = h.get('hint_id') or f'hint_{idx}'
            hid = 'hnt_' + _short_id(name + '|' + hid_raw + '|' + str(idx))
            hint_id_map[hid_raw] = hid
            cur.execute("""
                INSERT INTO AiTroubleshootingHints
                  (HintId, DocumentId, DefectName, CheckItem, Reason, EvidenceStrength,
                   RelatedProcess, RelatedPart, SourceFile, SheetName, SourceCellsJson)
                VALUES (?,?,?,?,?,?,?,?,?,?,?)
            """, (hid, doc_id, defect_name,
                  h.get('check_item', '') or '', h.get('reason', '') or '',
                  h.get('evidence_strength', '') or '',
                  h.get('related_process', '') or '',
                  h.get('related_part', '') or '',
                  h.get('source_file', name) or name,
                  h.get('sheet_name', '') or '',
                  json.dumps(h.get('source_cells') or [], ensure_ascii=False)))

        # 7. Extraction log + translation.
        log = result.get('ai_extraction_log') or {}
        log_id = 'log_' + _short_id(name)
        cur.execute("""
            INSERT INTO AiExtractionLogs
              (LogId, DocumentId, Confidence, AssumptionsJson, WarningsJson,
               DecisionRationale, CreatedAt)
            VALUES (?,?,?,?,?,?,?)
        """, (log_id, doc_id, float(log.get('confidence') or 0.0),
              json.dumps(log.get('assumptions') or [], ensure_ascii=False),
              json.dumps(log.get('warnings') or [], ensure_ascii=False),
              log.get('decision_rationale', '') or '', now))

        # 8. Translations for document, conclusions, hints, log (ko/en/vi).
        for lang, tr in (('ko', tr_ko), ('en', tr_en), ('vi', tr_vi)):
            if not tr:
                continue
            tdoc = (tr.get('document') or {})
            translated_report = translated_reports.get(lang) or _translated_report_markdown(tr, lang, generated_report)
            cur.execute("""
                INSERT OR REPLACE INTO AiDocumentTranslations
                  (DocumentId, Lang, Title, Purpose, ContentJson, GeneratedReportMarkdown, UpdatedAt)
                VALUES (?,?,?,?,?,?,?)
            """, (doc_id, lang, tdoc.get('title', '') or '',
                  tdoc.get('purpose', '') or '',
                  json.dumps(tdoc.get('content') or [], ensure_ascii=False),
                  translated_report, now))

            tconcl = tr.get('conclusions') or {}
            for cid_raw, tc in tconcl.items():
                cid = concl_id_map.get(cid_raw)
                if not cid:
                    continue
                cur.execute("""
                    INSERT OR REPLACE INTO AiConclusionTranslations
                      (ConclusionId, Lang, Topic, StatementFromReport,
                       NormalizedInterpretation, UpdatedAt)
                    VALUES (?,?,?,?,?,?)
                """, (cid, lang, tc.get('topic', '') or '',
                      tc.get('statement_from_report', '') or '',
                      tc.get('normalized_interpretation', '') or '', now))

            thints = tr.get('hints') or {}
            for hid_raw, th in thints.items():
                hid = hint_id_map.get(hid_raw)
                if not hid:
                    continue
                cur.execute("""
                    INSERT OR REPLACE INTO AiHintTranslations
                      (HintId, Lang, CheckItem, Reason, UpdatedAt)
                    VALUES (?,?,?,?,?)
                """, (hid, lang, th.get('check_item', '') or '',
                      th.get('reason', '') or '', now))

            tlog = tr.get('log') or {}
            cur.execute("""
                INSERT OR REPLACE INTO AiLogTranslations
                  (LogId, Lang, AssumptionsJson, WarningsJson,
                   DecisionRationale, UpdatedAt)
                VALUES (?,?,?,?,?,?)
            """, (log_id, lang,
                  json.dumps(tlog.get('assumptions') or [], ensure_ascii=False),
                  json.dumps(tlog.get('warnings') or [], ensure_ascii=False),
                  tlog.get('decision_rationale', '') or '', now))

        con.commit()
        _append_progress({'name': name, 'status': 'ok', 'at': now})
        return True
    except Exception as e:
        con.rollback()
        _append_progress({'name': name, 'status': 'error', 'error': repr(e), 'at': now_iso()})
        with open(FAILED_FILE, 'a', encoding='utf-8') as f:
            f.write(f'{name}\t{repr(e)}\n')
        return False
    finally:
        con.close()


def _append_progress(rec: dict) -> None:
    with open(PROGRESS_FILE, 'a', encoding='utf-8') as f:
        f.write(json.dumps(rec, ensure_ascii=False) + '\n')


def log_failed(name: str, reason: str) -> None:
    with open(FAILED_FILE, 'a', encoding='utf-8') as f:
        f.write(f'{name}\t{reason}\n')
    _append_progress({'name': name, 'status': 'failed', 'reason': reason, 'at': now_iso()})


def commit_payload(payload: dict) -> bool:
    """Commit one AI_EXCEL_PROC result payload without generated Python code.

    Expected JSON shape:
      {
        "name": "DatasetName",
        "result": {...},
        "translations": {"ko": {...}, "en": {...}, "vi": {...}}
      }
    """
    name = payload.get('name') or payload.get('dataset') or payload.get('source_dataset')
    if not name:
        raise ValueError('commit payload missing name')
    result = payload.get('result') or payload.get('ai_result') or payload
    translations = payload.get('translations') or {}
    return commit_dataset(
        name,
        result,
        translations.get('ko') or payload.get('tr_ko') or {},
        translations.get('en') or payload.get('tr_en') or {},
        translations.get('vi') or payload.get('tr_vi') or {},
    )


def commit_json_file(path: str) -> bool:
    if path == '-':
        payload = json.load(sys.stdin)
    else:
        with open(path, 'r', encoding='utf-8-sig') as f:
            payload = json.load(f)
    return commit_payload(payload)


def verify_counts() -> tuple[int, int, int]:
    targets = load_targets()
    con = open_db()
    try:
        ok = 0
        for t in targets:
            r = con.execute('SELECT COUNT(*) FROM AiDocuments WHERE SourceDataset=?', (t,)).fetchone()[0]
            if r > 0:
                ok += 1
        failed = 0
        if os.path.exists(FAILED_FILE):
            with open(FAILED_FILE, 'r', encoding='utf-8') as f:
                failed = sum(1 for ln in f if ln.strip())
        return len(targets), ok, failed
    finally:
        con.close()


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == 'verify':
        total, ok, failed = verify_counts()
        print(f'targets={total} processed={ok} failed_log_lines={failed}')
    elif len(sys.argv) > 2 and sys.argv[1] == 'commit-json':
        ok = commit_json_file(sys.argv[2])
        print('ok' if ok else 'failed')
        sys.exit(0 if ok else 1)
    elif len(sys.argv) > 2 and sys.argv[1] == 'excel-file':
        con = open_db()
        try:
            path = get_excel_file(con, sys.argv[2])
            if path:
                print(path)
        finally:
            con.close()
    elif len(sys.argv) > 2 and sys.argv[1] == 'excel-text':
        con = open_db()
        try:
            text = get_excel_text(con, sys.argv[2])
            if text:
                print(text)
        finally:
            con.close()
