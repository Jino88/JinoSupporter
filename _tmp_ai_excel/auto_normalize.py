# -*- coding: utf-8 -*-
"""Programmatic AI normalizer for excel_paste TSVs.

Approach:
- Parse TSV by `## Sheet` header → list of (sheet_name, rows[]).
- Pull report_date/model/process hints from dataset name.
- For each sheet, find rows that look like NG-rate or numeric measurements.
- Detect baseline/normal/before/old/OK rows by keyword (case-insensitive).
- Build minimum-valid payload that satisfies AI_EXCEL_PROC.md schema.
- Confidence is set low (0.3) and a 'auto-extracted' warning is emitted.

This does NOT replace human/agent-quality narrative — it gives schema-correct
coverage. Datasets flagged in AiExtractionLogs.WarningsJson should be re-run by
a deeper pass when the user requests.
"""
import re, os, sys, hashlib, json
from datetime import datetime

sys.path.insert(0, os.path.dirname(__file__))
from commit_lib import commit_payload, fetch_paste, _id

BASELINE_KEYS = re.compile(r'\b(normal|baseline|control|reference|before|old|기존|대조|^ok$)\b', re.I)
TEST_KEYS = re.compile(r'\b(test|after|new|change|改善|개선|trial|sample)\b', re.I)
NG_RATE_HINT = re.compile(r'(ng[\s_]*rate|불량률|불량|defect|fail|fpy)', re.I)
DATE_RX = re.compile(r'(20\d{2})[.\-_/](\d{1,2})[.\-_/](\d{1,2})')
DATE_DMY = re.compile(r'(\d{1,2})[.\-_/](\d{1,2})[.\-_/](20\d{2})')
MODEL_RX = re.compile(r'((?:BRS|MSU|TIU|GMI|DT)[\s-]*[A-Z0-9-]+)', re.I)
PROCESS_HINT = re.compile(r'(VP|CD|coil|frame|bond|UV|plasma|tension|gauss|noise|hearing|SPL|THD|impedance|dimension|height|offset|deform|burr|damage|weak\s*solder|mag|magnet)', re.I)


def parse_sheets(text: str):
    sheets = []
    cur_name = None
    cur_rows = []
    for line in text.splitlines():
        if line.startswith('## '):
            if cur_name is not None:
                sheets.append((cur_name, cur_rows))
            cur_name = line[3:].strip()
            cur_rows = []
        else:
            cur_rows.append(line.split('\t'))
    if cur_name is not None:
        sheets.append((cur_name, cur_rows))
    return sheets


def extract_date(name: str) -> str | None:
    m = DATE_RX.search(name)
    if m:
        y, mo, d = m.groups()
        return f'{int(y):04d}-{int(mo):02d}-{int(d):02d}'
    m = DATE_DMY.search(name)
    if m:
        d, mo, y = m.groups()
        return f'{int(y):04d}-{int(mo):02d}-{int(d):02d}'
    return None


def extract_model(name: str) -> str | None:
    m = MODEL_RX.search(name)
    return m.group(1).upper().replace(' ', '') if m else None


def extract_processes(name: str, body: str) -> list[str]:
    out = set()
    for s in (name, body[:5000]):
        for m in PROCESS_HINT.finditer(s):
            out.add(m.group(1))
        if len(out) > 12:
            break
    return sorted(out, key=str.lower)[:12]


def looks_numeric(s: str) -> bool:
    s = s.strip().rstrip('%')
    if not s:
        return False
    try:
        float(s)
        return True
    except ValueError:
        return False


def find_numeric_rows(rows: list[list[str]]) -> list[tuple[int, list[str]]]:
    out = []
    for i, r in enumerate(rows):
        if not r or all(not c.strip() for c in r):
            continue
        nums = sum(1 for c in r if looks_numeric(c))
        if nums >= 2:
            out.append((i, r))
    return out


def classify(name: str, sheets) -> tuple[str, str]:
    """Return (report_type, primary_defect_canonical)."""
    nl = name.lower()
    body = '\n'.join('\t'.join(r) for _, rows in sheets for r in rows[:80]).lower()
    text = nl + '\n' + body[:8000]

    has_baseline = bool(BASELINE_KEYS.search(text))
    has_ngrate = bool(NG_RATE_HINT.search(text))
    has_dimension = re.search(r'(dimension|height|offset|gap|size|thickness)', text)
    has_reliability = re.search(r'(spl|thd|impedance|reliability|aging|drop test|temperature.+humidity|spec)', text)
    has_doe = re.search(r'(doe|matrix|combination|factor)', text)
    has_image = re.search(r'(picture|photo|x-?ray|image|caption)', text)
    is_blank = sum(len(rows) for _, rows in sheets) < 4

    primary = 'Unknown'
    if re.search(r'vp.*cd.*sep|sep.*vp.*cd', text):
        primary = 'VP+CD Separation'
    elif re.search(r'noise', text): primary = 'NG Hearing Noise'
    elif re.search(r'touch', text): primary = 'NG Hearing Touch'
    elif re.search(r'hearing', text): primary = 'NG Hearing'
    elif re.search(r'spl', text): primary = 'NG Sigma SPL'
    elif re.search(r'thd', text): primary = 'NG Sigma THD'
    elif re.search(r'tension', text): primary = 'Tension NG'
    elif re.search(r'gauss', text): primary = 'Low Gauss'
    elif re.search(r'over\s*bond', text): primary = 'Over Bond'
    elif re.search(r'weak.*solder', text): primary = 'Weak Solder'
    elif re.search(r'deform', text): primary = 'Deform NG'
    elif re.search(r'damage', text): primary = 'Damage NG'
    elif re.search(r'burr', text): primary = 'Burr NG'
    elif re.search(r'offset', text): primary = 'Offset NG'
    elif has_dimension: primary = 'Dimension NG'

    if is_blank:
        return ('image_dependent', primary)
    if has_doe:
        return ('doe_matrix', primary)
    if has_reliability and not has_ngrate:
        return ('reliability_spec', primary)
    if has_dimension and not has_ngrate:
        return ('before_after_dimension', primary)
    if has_image and not has_ngrate:
        return ('image_dependent', primary)
    if has_ngrate and has_baseline:
        return ('normal_comparison', primary)
    if has_ngrate:
        return ('ng_without_baseline', primary)
    return ('mixed', primary)


def find_ng_rows(sheets):
    """Find rows with an NG-rate-looking column ≤ 100. Scans multiple header
    candidates per sheet, not just first row.
    """
    out = []
    for sheet_name, rows in sheets:
        # collect every (row, col) header candidate that mentions ng rate
        headers = []
        for ri, r in enumerate(rows):
            for ci, c in enumerate(r):
                if NG_RATE_HINT.search((c or '').strip()):
                    headers.append((ri, ci))
        # for each header, walk the rows below until blank-band of 3
        for header_row, col in headers:
            blank_streak = 0
            for ri in range(header_row + 1, min(len(rows), header_row + 80)):
                r = rows[ri]
                if col >= len(r):
                    blank_streak += 1
                    if blank_streak >= 3:
                        break
                    continue
                cell = r[col].strip()
                if not cell:
                    blank_streak += 1
                    if blank_streak >= 3:
                        break
                    continue
                v = cell.rstrip('%')
                if not looks_numeric(v):
                    blank_streak += 1
                    if blank_streak >= 3:
                        break
                    continue
                try:
                    fv = float(v)
                except ValueError:
                    continue
                if fv < 0 or fv > 100:
                    continue
                # decimal fraction (≤1) → percent
                pct = fv * 100 if fv <= 1 else fv
                # 100% is almost always a portion/count column, not a real NG rate; skip
                if pct >= 99.5:
                    continue
                label_cells = [c for c in r[:col] if c.strip()]
                label = ' / '.join(label_cells)[:100] if label_cells else f'row{ri}'
                out.append({'sheet': sheet_name, 'row_idx': ri,
                            'label': label, 'value_percent': round(pct, 4)})
                blank_streak = 0
        if len(out) > 60:
            break
    return out[:60]


def build_payload(dataset_name: str, paste: str) -> dict:
    sheets = parse_sheets(paste)
    report_type, primary = classify(dataset_name, sheets)
    date = extract_date(dataset_name)
    model = extract_model(dataset_name)
    processes = extract_processes(dataset_name, paste)
    ng_rows = find_ng_rows(sheets)

    # Baseline detection
    baseline = None
    tests = []
    for nr in ng_rows:
        if BASELINE_KEYS.search(nr['label']):
            if baseline is None or nr['value_percent'] < baseline['value_percent']:
                baseline = nr
        else:
            tests.append(nr)

    # Build results: one row per detected ng entry (cap)
    results = []
    for i, nr in enumerate(ng_rows[:30]):
        row_label = nr['label']
        is_base = baseline and nr is baseline
        results.append({
            'measurement_type': 'function' if 'function' in row_label.lower() else 'ng_rate',
            'condition_group': 'baseline' if is_base else ('test' if BASELINE_KEYS.search(row_label) is None else 'baseline'),
            'date': date or '',
            'line': '',
            'metric_name': f'NG Rate — {row_label}'[:120],
            'metric_value': nr['value_percent'],
            'unit': '%',
            'ng_rate_percent': nr['value_percent'],
            'ng_rate_decimal': nr['value_percent'] / 100.0,
            'judgement': None,
            'sheet_name': nr['sheet'],
            'source_cells': [f"{nr['sheet']}!row{nr['row_idx']}"],
        })

    # Conclusion
    if report_type == 'normal_comparison' and baseline and tests:
        worst = max(tests, key=lambda x: x['value_percent'])
        if baseline['value_percent'] > 0:
            ratio = worst['value_percent'] / baseline['value_percent']
            delta = (ratio - 1) * 100
            verdict_en = (f"Worst test '{worst['label']}' NG {worst['value_percent']:.2f}% vs baseline "
                          f"'{baseline['label']}' {baseline['value_percent']:.2f}% = {ratio:.2f}x, "
                          f"{abs(delta):.1f}% {'worse' if delta>0 else 'improved'} vs same-event baseline.")
        else:
            verdict_en = (f"Baseline '{baseline['label']}' NG 0% — any non-zero test value indicates worsening. "
                          f"Worst test {worst['label']} = {worst['value_percent']:.2f}%.")
    elif report_type == 'ng_without_baseline' and ng_rows:
        worst = max(ng_rows, key=lambda x: x['value_percent'])
        best = min(ng_rows, key=lambda x: x['value_percent'])
        verdict_en = (f"No same-event baseline detected. NG range {best['value_percent']:.2f}%-"
                      f"{worst['value_percent']:.2f}% across {len(ng_rows)} rows. "
                      f"Worst: '{worst['label']}'. Cannot claim improvement/worsening — see raw rows.")
    elif report_type == 'before_after_dimension':
        verdict_en = "Dimension/measurement workbook. NG-rate judgement does not apply — see metric rows for delta vs spec."
    elif report_type == 'reliability_spec':
        verdict_en = "Reliability/spec gate workbook. Rows store metric vs spec; PASS/FAIL where workbook stated."
    elif report_type == 'doe_matrix':
        verdict_en = "DOE/factor matrix workbook. Multiple condition combinations stored as rows; pick best/worst from condition labels."
    elif report_type == 'image_dependent':
        verdict_en = "Workbook is image-dependent or near-empty. Image/OCR review required to extract evidence."
    else:
        verdict_en = "Mixed evidence. Numeric rows stored as-is; no single dominant judgement."

    verdict_ko = {
        'normal_comparison': '같은 이벤트 baseline 대비 상대 변화율로 판정 (위 영문 참고).',
        'ng_without_baseline': '대조군 부재 — 절대 NG 순위만 제공, 개선/악화 단정 불가.',
        'before_after_dimension': '치수/측정 리포트 — NG rate 판정 비적용. metric 행으로 spec 대비 delta 확인.',
        'reliability_spec': '신뢰성/스펙 리포트 — 각 metric 행의 PASS/FAIL 보존.',
        'doe_matrix': 'DOE/조건 매트릭스 — 조건 조합별 best/worst를 라벨로 확인.',
        'image_dependent': '이미지 의존 리포트 — 텍스트 부재, 이미지/OCR 재검토 필요.',
        'mixed': '복합 리포트 — 단일 판정 없음, 개별 metric 행 참고.'
    }[report_type]
    verdict_vi = {
        'normal_comparison': 'So với baseline cùng sự kiện, tỉ lệ thay đổi tương đối (xem dòng tiếng Anh ở trên).',
        'ng_without_baseline': 'Không có baseline — chỉ xếp hạng NG tuyệt đối, không kết luận cải thiện/xấu đi.',
        'before_after_dimension': 'Báo cáo kích thước/đo lường — không áp dụng phán định NG rate. Dùng dòng metric để xem delta so với spec.',
        'reliability_spec': 'Báo cáo độ tin cậy/spec — giữ PASS/FAIL của từng metric.',
        'doe_matrix': 'Ma trận DOE/yếu tố — xem nhãn điều kiện để chọn best/worst.',
        'image_dependent': 'Báo cáo phụ thuộc hình ảnh — thiếu text, cần review hình/OCR.',
        'mixed': 'Báo cáo hỗn hợp — không có phán định duy nhất, xem từng dòng metric.'
    }[report_type]

    purpose_en = (f"Process review report covering {primary}." +
                  (f" Model {model}." if model else '') +
                  (f" Date {date}." if date else ''))
    purpose_ko = (f"{primary} 관련 공정 검토 리포트." +
                  (f" 모델 {model}." if model else '') +
                  (f" 일자 {date}." if date else ''))
    purpose_vi = (f"Báo cáo xem xét quá trình về {primary}." +
                  (f" Model {model}." if model else '') +
                  (f" Ngày {date}." if date else ''))

    sheet_names = [s for s, _ in sheets][:8]
    content_en = [f"Sheets: {', '.join(sheet_names) or 'n/a'}",
                  f"Detected NG-rate rows: {len(ng_rows)} (baseline found: {bool(baseline)})"]
    content_ko = [f"시트: {', '.join(sheet_names) or '없음'}",
                  f"감지된 NG-rate 행: {len(ng_rows)}개 (대조군 존재: {'예' if baseline else '아니오'})"]
    content_vi = [f"Sheet: {', '.join(sheet_names) or 'n/a'}",
                  f"Số dòng NG-rate phát hiện: {len(ng_rows)} (có baseline: {'có' if baseline else 'không'})"]

    payload = {
        'schema_version': '0.1',
        'dataset_name': dataset_name,
        'document': {
            'source_file': dataset_name,
            'source_sheet': sheet_names[0] if sheet_names else '',
            'title': dataset_name,
            'model': model or '',
            'report_date': date or '',
            'department': '',
            'marker': '',
            'line': '',
            'report_type': report_type,
            'primary_defect': {'canonical_name': primary, 'aliases_in_document': []},
            'related_defects': [],
            'parts': [],
            'processes': processes,
            'purpose': purpose_en,
            'content': content_en,
            'source_cells': {'title': [], 'date': [], 'purpose': [], 'content': []},
            'translations': {
                'ko': {'title': dataset_name, 'purpose': purpose_ko, 'content': content_ko},
                'en': {'title': dataset_name, 'purpose': purpose_en, 'content': content_en},
                'vi': {'title': dataset_name, 'purpose': purpose_vi, 'content': content_vi},
            },
        },
        'test_conditions': [],
        'results': results,
        'conclusions': [{
            'topic': f'{report_type} verdict',
            'statement_from_report': '',
            'normalized_interpretation': verdict_en,
            'sheet_name': sheet_names[0] if sheet_names else '',
            'source_cells': [],
            'translations': {
                'ko': {'topic': f'{report_type} 결론', 'statement_from_report': '',
                       'normalized_interpretation': verdict_ko},
                'en': {'topic': f'{report_type} verdict', 'statement_from_report': '',
                       'normalized_interpretation': verdict_en},
                'vi': {'topic': f'Kết luận {report_type}', 'statement_from_report': '',
                       'normalized_interpretation': verdict_vi},
            },
        }],
        'troubleshooting': {
            'defect_name': primary,
            'when_user_asks': [primary, model or ''],
            'suggested_checks': [{
                'check_item': f'Inspect rows of report_type={report_type} for {primary}; auto-extracted, verify with original workbook.',
                'reason': 'Auto-extraction (low confidence) — baseline, factor, and judgement may be incomplete.',
                'evidence_strength': 'low',
                'related_process': processes[0] if processes else '',
                'related_part': '',
                'sheet_name': sheet_names[0] if sheet_names else '',
                'source_cells': [],
                'translations': {
                    'ko': {'check_item': f'{report_type} 리포트의 행을 검토하여 {primary} 확인 — 자동 추출이므로 원본 확인 필요.',
                           'reason': '자동 추출 (낮은 신뢰도) — baseline, 인자, 판정이 불완전할 수 있음.'},
                    'en': {'check_item': f'Inspect rows of report_type={report_type} for {primary}; auto-extracted, verify with original workbook.',
                           'reason': 'Auto-extraction (low confidence) — baseline, factor, and judgement may be incomplete.'},
                    'vi': {'check_item': f'Kiểm tra các dòng của report_type={report_type} cho {primary}; tự động trích xuất, cần đối chiếu file gốc.',
                           'reason': 'Trích xuất tự động (độ tin cậy thấp) — baseline, yếu tố và phán định có thể chưa đầy đủ.'},
                },
            }],
            'limitations': ['Auto-extracted — narrative depth and baseline detection limited.'],
        },
        'ai_extraction_log': {
            'confidence': 0.3,
            'assumptions': [
                f'report_type heuristically classified as {report_type}',
                f'primary_defect heuristically assigned as {primary}',
                'Baseline detected by keyword scan' if baseline else 'No baseline keyword detected',
            ],
            'warnings': [
                'AUTO_EXTRACTED: Low-confidence batch normalization. Re-run a deeper agent pass for production-grade narratives.',
                'Numeric values copied from TSV without unit/spec validation.',
            ],
            'decision_rationale': (
                f'Programmatic normalization. Detected {len(ng_rows)} NG-rate rows across '
                f'{len(sheets)} sheets. Baseline present: {bool(baseline)}. '
                f'report_type={report_type}. primary_defect={primary}.'
            ),
            'translations': {
                'ko': {
                    'assumptions': [f'report_type 휴리스틱: {report_type}', f'primary_defect 휴리스틱: {primary}',
                                    '키워드 스캔으로 baseline 감지' if baseline else 'baseline 키워드 없음'],
                    'warnings': ['AUTO_EXTRACTED: 저신뢰 배치 정규화. 프로덕션 narrative는 심층 에이전트 재처리 필요.',
                                 'TSV 숫자값을 단위/스펙 검증 없이 복사함.'],
                    'decision_rationale': (
                        f'프로그램 정규화. {len(sheets)}개 시트에서 NG-rate 행 {len(ng_rows)}개 감지. '
                        f'Baseline 존재: {bool(baseline)}. report_type={report_type}. primary_defect={primary}.'
                    ),
                },
                'en': {
                    'assumptions': [
                        f'report_type heuristically classified as {report_type}',
                        f'primary_defect heuristically assigned as {primary}',
                        'Baseline detected by keyword scan' if baseline else 'No baseline keyword detected',
                    ],
                    'warnings': ['AUTO_EXTRACTED: Low-confidence batch normalization. Re-run a deeper agent pass for production-grade narratives.',
                                 'Numeric values copied from TSV without unit/spec validation.'],
                    'decision_rationale': (
                        f'Programmatic normalization. Detected {len(ng_rows)} NG-rate rows across '
                        f'{len(sheets)} sheets. Baseline present: {bool(baseline)}. '
                        f'report_type={report_type}. primary_defect={primary}.'
                    ),
                },
                'vi': {
                    'assumptions': [f'report_type phân loại heuristic: {report_type}',
                                    f'primary_defect heuristic: {primary}',
                                    'Phát hiện baseline qua quét từ khóa' if baseline else 'Không có từ khóa baseline'],
                    'warnings': ['AUTO_EXTRACTED: Chuẩn hóa batch độ tin cậy thấp. Cần chạy lại sâu hơn để có narrative chất lượng sản xuất.',
                                 'Các giá trị số được sao chép từ TSV mà không xác minh đơn vị/spec.'],
                    'decision_rationale': (
                        f'Chuẩn hóa lập trình. Phát hiện {len(ng_rows)} dòng NG-rate trên '
                        f'{len(sheets)} sheet. Có baseline: {bool(baseline)}. '
                        f'report_type={report_type}. primary_defect={primary}.'
                    ),
                },
            },
        },
    }
    return payload


def normalize_and_commit(dataset_name: str) -> str | None:
    paste = fetch_paste(dataset_name)
    if not paste:
        return None
    payload = build_payload(dataset_name, paste)
    return commit_payload(payload)


if __name__ == '__main__':
    name = sys.argv[1]
    print(normalize_and_commit(name))
