"""Chunk 04 — AI Batch normalization per AI_EXCEL_PROC.md.

Processes 9 datasets from `_chunk_04.txt`, classifies report_type, builds
normalized JSON, ko/en/vi translations, and commits via `_ai_batch_helper`.
"""
from __future__ import annotations
import sys, io, os
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

sys.path.insert(0, r'D:\000. MyWorks\005. Program\Repository\JinoSupporter')
import _ai_batch_helper as h

# =====================================================================
# Per-dataset normalized result + translations.
# Each entry: (dataset_name, build_fn) where build_fn() -> (result, tr_ko, tr_en, tr_vi)
# =====================================================================

def _doc(name, title, model, report_date, dept, marker, line, report_type,
         primary_canonical, primary_aliases, related, parts, processes,
         purpose, content):
    return {
        'document_id': '',
        'source_file': name,
        'source_sheet': '',
        'title': title,
        'model': model,
        'report_date': report_date,
        'department': dept,
        'marker': marker,
        'line': line,
        'report_type': report_type,
        'primary_defect': {
            'canonical_name': primary_canonical,
            'aliases_in_document': primary_aliases,
        },
        'related_defects': related,
        'parts': parts,
        'processes': processes,
        'purpose': purpose,
        'content': content,
        'source_cells': {'title': [], 'date': [], 'purpose': [], 'content': []},
    }


# ---------------------------------------------------------------------
# 1. TIU C11-20 — Frame Test Mold #1-B improve damage mesh
# ---------------------------------------------------------------------
def build_tiu_c11_20_frame():
    name = '28. TIU C11-20  Report Test Frame #1-B improve damage mesh 2026.03.31. - RAW soundcheck'
    sheet = 'Test'

    doc = _doc(
        name=name,
        title='TIU C11-20 Report Test Frame NG Damage Mesh Improve',
        model='TIU C11-20R',
        report_date='2026-03-31',
        dept='ME', marker='Trung', line='',
        report_type='normal_comparison',
        primary_canonical='Damage Mesh (Frame)',
        primary_aliases=['NG damage mesh', 'Damage Frame mesh'],
        related=['NG Hearing Touch', 'NG Hearing Noise', 'NG AUDIOBUS SPL'],
        parts=['Frame', 'Mold #1-B'],
        processes=['Forming Frame Mold', 'Final Function Check', 'SPL Soundcheck'],
        purpose='Validate that the improved Frame Mold #1-B reduces damage mesh and is acceptable on function and SPL.',
        content=[
            "Make final samples with improved Frame Mold #1-B (98 pcs) and compare function and SPL against same-event Normal lot (280 pcs).",
            "Audiobus (SPL/SPL+RB/No sound) and Hearing (Noise/Touch) checked at final; SPL_FREQ raw measurement done on day shift 31-Mar.",
        ],
    )

    test_conditions = [
        {
            'condition_id': 'cond_1',
            'condition_group': 'Frame mold change',
            'line': '',
            'process': 'Forming Frame',
            'changed_factor': 'Frame mold',
            'before_value': 'Normal Frame mold',
            'after_value': 'Frame Test Mold #1-B (improved)',
            'unit': None, 'machine': 'Mold #1-B', 'jig': None,
            'material_lot': None, 'supplier': None,
            'dry_time_sec': None, 'temperature': None, 'pressure': None,
            'bond_amount': None, 'uv_energy': None,
            'source_file': name, 'sheet_name': sheet,
            'source_cells': ['II. Content', 'Mold #1-B Q\'ty 100pcs'],
        },
    ]

    # 31-Mar: Test 98/96, 0 SPL/0 SPL+RB/0 No sound; Hearing Noise 1, Touch 1 => total 2, NG 2.0%.
    # Normal 280/276, 0/0/0, Noise 4, Touch 0 => total 4, NG 1.4%.
    # (2.04% / 1.43% - 1) * 100 = +42.9% (worse than normal).
    results = [
        {
            'result_id': 'res_1', 'condition_id': 'cond_1',
            'measurement_type': 'Function',
            'condition_group': 'Frame Test Mold #1-B',
            'date': '2026-03-31', 'line': 'Day shift',
            'input_count': 98, 'ok_count': 96, 'ng_count': 2,
            'ng_rate_decimal': 0.0204, 'ng_rate_percent': 2.04,
            'metric_name': 'Total NG rate', 'metric_value': 2.04,
            'unit': '%', 'judgement': None,
            'ng_breakdown': {
                'NG AUDIOBUS SPL': {'count': 0, 'rate': 0.0},
                'NG AUDIOBUS SPL+RB': {'count': 0, 'rate': 0.0},
                'NG AUDIOBUS No sound': {'count': 0, 'rate': 0.0},
                'NG Hearing Noise': {'count': 1, 'rate': 0.0102},
                'NG Hearing Touch': {'count': 1, 'rate': 0.0102},
            },
            'source_file': name, 'sheet_name': sheet,
            'source_cells': ['Row 1 Frame Test Mold #1-B'],
        },
        {
            'result_id': 'res_2', 'condition_id': None,
            'measurement_type': 'Function',
            'condition_group': 'Normal',
            'date': '2026-03-31', 'line': '',
            'input_count': 280, 'ok_count': 276, 'ng_count': 4,
            'ng_rate_decimal': 0.0143, 'ng_rate_percent': 1.43,
            'metric_name': 'Total NG rate', 'metric_value': 1.43,
            'unit': '%', 'judgement': None,
            'ng_breakdown': {
                'NG AUDIOBUS SPL': {'count': 0, 'rate': 0.0},
                'NG AUDIOBUS SPL+RB': {'count': 0, 'rate': 0.0},
                'NG AUDIOBUS No sound': {'count': 0, 'rate': 0.0},
                'NG Hearing Noise': {'count': 4, 'rate': 0.0143},
                'NG Hearing Touch': {'count': 0, 'rate': 0.0},
            },
            'source_file': name, 'sheet_name': sheet,
            'source_cells': ['Row 2 Normal'],
        },
        {
            'result_id': 'res_3', 'condition_id': 'cond_1',
            'measurement_type': 'SPL Acoustic',
            'condition_group': 'Frame Test Mold #1-B vs Normal L (SPL raw)',
            'date': '2026-03-31', 'line': '',
            'input_count': 10, 'ok_count': None, 'ng_count': None,
            'ng_rate_decimal': None, 'ng_rate_percent': None,
            'metric_name': 'SPL avg comparison (Frame Test vs Normal L, 10 samples each)',
            'metric_value': None, 'unit': 'dB', 'judgement': 'CHECK',
            'ng_breakdown': {},
            'source_file': name, 'sheet_name': 'RAW DATA',
            'source_cells': ['SPL_FREQ1/2/3 STD vs Frame Test vs Normal L avgs'],
        },
    ]

    conclusions = [
        {
            'conclusion_id': 'concl_1',
            'topic': 'Frame Test Mold #1-B vs Normal — function NG rate',
            'statement_from_report': 'Frame Test Mold #1-B Check SPL: There is an improvement compared to before the improvement. => Frame after improve can use!',
            'normalized_interpretation': 'Frame Test Mold #1-B (98 pcs) NG rate 2.04% vs same-event Normal (280 pcs) NG rate 1.43% = 1.43x, 42.9% worse than normal on Function. SPL acoustic check is stated as improved by the report (no numeric SPL judgement table shown). Sample size on test arm (98 pcs) is small; conclusion to release the mold is based on report decision, not on a statistical NG-rate improvement.',
            'source_file': name, 'sheet_name': sheet,
            'source_cells': ['IV. Decision'],
        },
    ]

    troubleshooting_index = {
        'defect_name': 'Damage Mesh (Frame)',
        'when_user_asks': ['frame damage mesh', 'frame mold change', 'mold #1-B'],
        'suggested_checks': [
            {
                'hint_id': 'hint_1',
                'check_item': 'Compare Frame mold #1-B function NG vs same-event Normal (not just SPL).',
                'reason': 'On 31-Mar function check, Test Mold #1-B 2.04% (2/98) is 42.9% worse than Normal 1.43% (4/280); decision relies on SPL improvement only. Confirm with larger sample.',
                'evidence_strength': 'low',
                'related_process': 'Forming Frame',
                'related_part': 'Frame Mold #1-B',
                'source_file': name, 'sheet_name': sheet,
                'source_cells': ['III. Result row 1+2'],
            },
            {
                'hint_id': 'hint_2',
                'check_item': 'Run SPL_FREQ1/2/3 comparison test vs Normal L on a larger sample.',
                'reason': 'RAW DATA contains only 10 samples per arm; SPL improvement claim should be replicated.',
                'evidence_strength': 'medium',
                'related_process': 'Final Function (SPL)',
                'related_part': 'Frame Mold #1-B',
                'source_file': name, 'sheet_name': 'RAW DATA',
                'source_cells': ['SPL_FREQ1/2/3 10-sample blocks'],
            },
        ],
        'limitations': ['Test arm sample size is only 98 pcs vs Normal 280 pcs.', 'No numeric SPL spec gate; only verbal "improved" judgement.'],
    }

    ai_log = {
        'confidence': 0.6,
        'assumptions': [
            'Same-event Normal row 2 is the correct baseline for the Test Mold #1-B row 1.',
            'NG rate is computed by report as 2/98=2.04% and 4/280=1.43%.',
        ],
        'warnings': [
            'Function NG rate is worse than Normal even though the report decision is "can use" based on SPL.',
            'Frame Test sample size (98 pcs) is much smaller than Normal (280 pcs).',
        ],
        'decision_rationale': 'Report classifies as normal_comparison: same-event Normal exists. Relative change on Function = (2.04/1.43 - 1)*100 = 42.9% worse. SPL section is qualitative only.',
    }

    result = {
        'schema_version': '0.1',
        'document': doc,
        'test_conditions': test_conditions,
        'results': results,
        'conclusions': conclusions,
        'troubleshooting_index': troubleshooting_index,
        'ai_extraction_log': ai_log,
    }

    tr_en = {
        'document': {
            'title': doc['title'],
            'purpose': doc['purpose'],
            'content': doc['content'],
        },
        'conclusions': {
            'concl_1': {
                'topic': conclusions[0]['topic'],
                'statement_from_report': conclusions[0]['statement_from_report'],
                'normalized_interpretation': conclusions[0]['normalized_interpretation'],
            },
        },
        'hints': {
            'hint_1': {'check_item': troubleshooting_index['suggested_checks'][0]['check_item'],
                       'reason': troubleshooting_index['suggested_checks'][0]['reason']},
            'hint_2': {'check_item': troubleshooting_index['suggested_checks'][1]['check_item'],
                       'reason': troubleshooting_index['suggested_checks'][1]['reason']},
        },
        'log': {
            'assumptions': ai_log['assumptions'],
            'warnings': ai_log['warnings'],
            'decision_rationale': ai_log['decision_rationale'],
        },
    }

    tr_ko = {
        'document': {
            'title': 'TIU C11-20 Frame NG Damage Mesh 개선 시험 리포트',
            'purpose': '개선된 Frame Mold #1-B 가 damage mesh 를 줄이고 Function 및 SPL 기준을 만족하는지 검증.',
            'content': [
                "Frame Test Mold #1-B 로 최종 샘플(98 pcs)을 만들고, 동일 이벤트 Normal(280 pcs) 과 Function/SPL 비교.",
                "Audiobus(SPL/SPL+RB/No sound) 및 Hearing(Noise/Touch) 최종 검사, SPL_FREQ 원시 측정은 31-Mar Day shift 에서 실시.",
            ],
        },
        'conclusions': {
            'concl_1': {
                'topic': 'Frame Test Mold #1-B vs Normal — Function NG rate',
                'statement_from_report': 'Frame Test Mold #1-B SPL 점검 결과 개선 전 대비 향상됨 => Frame 개선 후 사용 가능!',
                'normalized_interpretation': 'Frame Test Mold #1-B (98 pcs) NG rate 2.04% vs 동일 이벤트 Normal (280 pcs) 1.43% = 1.43배, Function 기준으로 normal 대비 42.9% 악화. SPL acoustic 은 리포트가 정량 표 없이 "개선됨" 으로 기술. Test arm 표본(98 pcs)이 작아 NG rate 통계적 개선이 아니라 리포트 판정에 의존한 결정.',
            },
        },
        'hints': {
            'hint_1': {
                'check_item': 'Frame Mold #1-B Function NG 를 같은 이벤트 Normal 과 비교 (SPL 만 보지 말 것).',
                'reason': '31-Mar Function 결과: Test Mold #1-B 2.04% (2/98) 가 Normal 1.43% (4/280) 대비 42.9% 악화. 결정은 SPL 개선에만 근거. 표본 확대 필요.',
            },
            'hint_2': {
                'check_item': 'Normal L 대비 SPL_FREQ1/2/3 비교 시험을 더 큰 표본으로 재실시.',
                'reason': 'RAW DATA 는 arm 당 10 샘플 뿐. SPL 개선 주장 재현 필요.',
            },
        },
        'log': {
            'assumptions': [
                '동일 이벤트 Normal row 2 가 Test Mold #1-B row 1 의 baseline.',
                'NG rate 는 리포트 계산값 2/98=2.04%, 4/280=1.43% 사용.',
            ],
            'warnings': [
                'Function NG rate 는 Normal 대비 악화임에도 리포트는 SPL 근거로 "사용 가능" 판정.',
                'Frame Test 표본(98 pcs)이 Normal(280 pcs) 보다 훨씬 작음.',
            ],
            'decision_rationale': 'normal_comparison 분류: 동일 이벤트 Normal 존재. Function 상대 변화율 = (2.04/1.43 - 1)*100 = 42.9% 악화. SPL 섹션은 정성 표현만 존재.',
        },
    }

    tr_vi = {
        'document': {
            'title': 'Báo cáo thử nghiệm cải tiến NG damage mesh Frame TIU C11-20',
            'purpose': 'Xác nhận Frame Mold #1-B cải tiến giảm damage mesh và đạt Function/SPL.',
            'content': [
                "Làm mẫu final với Frame Test Mold #1-B (98 pcs), so sánh Function/SPL với Normal cùng đợt (280 pcs).",
                "Audiobus (SPL/SPL+RB/No sound) và Hearing (Noise/Touch) kiểm tra cuối; đo SPL_FREQ raw vào ca ngày 31-Mar.",
            ],
        },
        'conclusions': {
            'concl_1': {
                'topic': 'Frame Test Mold #1-B vs Normal — NG rate Function',
                'statement_from_report': 'Frame Test Mold #1-B kiểm SPL: có cải thiện so với trước cải tiến => Frame sau cải tiến có thể dùng!',
                'normalized_interpretation': 'Frame Test Mold #1-B (98 pcs) NG rate 2.04% so với Normal cùng đợt (280 pcs) 1.43% = 1.43x, xấu hơn Normal 42.9% về Function. SPL acoustic được báo cáo mô tả là "cải thiện" nhưng không có bảng định lượng. Mẫu test (98 pcs) nhỏ; kết luận dựa vào quyết định của báo cáo, không phải cải thiện NG rate.',
            },
        },
        'hints': {
            'hint_1': {
                'check_item': 'So sánh NG Function của Frame Mold #1-B với Normal cùng đợt (đừng chỉ nhìn SPL).',
                'reason': 'Function 31-Mar: Test Mold #1-B 2.04% (2/98) xấu hơn Normal 1.43% (4/280) 42.9%; quyết định chỉ dựa vào SPL. Cần mẫu lớn hơn.',
            },
            'hint_2': {
                'check_item': 'Lặp lại so sánh SPL_FREQ1/2/3 với Normal L trên mẫu lớn hơn.',
                'reason': 'RAW DATA chỉ có 10 mẫu mỗi arm; cần xác nhận lại tuyên bố cải thiện SPL.',
            },
        },
        'log': {
            'assumptions': [
                'Normal cùng đợt (row 2) là baseline đúng cho Test Mold #1-B (row 1).',
                'NG rate dùng theo báo cáo: 2/98=2.04% và 4/280=1.43%.',
            ],
            'warnings': [
                'NG rate Function xấu hơn Normal nhưng báo cáo kết luận "có thể dùng" dựa trên SPL.',
                'Mẫu Frame Test (98 pcs) nhỏ hơn nhiều so với Normal (280 pcs).',
            ],
            'decision_rationale': 'Phân loại normal_comparison: có Normal cùng đợt. Thay đổi tương đối Function = (2.04/1.43 - 1)*100 = 42.9% xấu hơn. Phần SPL chỉ định tính.',
        },
    }

    return name, result, tr_ko, tr_en, tr_vi


# ---------------------------------------------------------------------
# 2. TIU L5S3-01 BAKO high — no baseline (ng_without_baseline)
# ---------------------------------------------------------------------
def build_tiu_l5s3_bako():
    name = '28. TIU L5S3-01 R Report test find reason NG BAKO high 2025.12.15'
    sheet = 'Test'

    doc = _doc(
        name=name,
        title='TIU L5S3-01 [R] Report Test Find Reason Improve NG Function (BAKO)',
        model='TIU L5S3-01 [R]',
        report_date='2025-12-16',
        dept='ME', marker='Thao', line='',
        report_type='ng_without_baseline',
        primary_canonical='NG BAKO (Function)',
        primary_aliases=['NG function Bako', 'BAKO FRF', 'BAKO FRF+SPL'],
        related=['BAKO FRF', 'BAKO FRF+SPL', 'BAKO THD', 'BAKO No sound'],
        parts=['SPK Function machine AWF'],
        processes=['Function Check AWF #1/#2/#3'],
        purpose='Find reason NG Function BAKO is very high (~40%) by separating function machine AWF #1/#2/#3 and re-checking.',
        content=[
            "Function NG BAKO is at ~40% baseline level (stated in purpose).",
            "Separate function machine AWF #1, #2, #3 and re-check function on 16-Dec to see if NG rate depends on machine.",
        ],
    )

    test_conditions = [
        {'condition_id': 'cond_1', 'condition_group': 'AWF function machine split',
         'line': '', 'process': 'Function AWF', 'changed_factor': 'AWF machine',
         'before_value': 'All AWF mixed', 'after_value': 'AWF #1 / #2 / #3 separated',
         'unit': None, 'machine': 'AWF', 'jig': None, 'material_lot': None,
         'supplier': None, 'dry_time_sec': None, 'temperature': None,
         'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': name, 'sheet_name': sheet,
         'source_cells': ['II. Content', 'Test separate machine AWF 1-2-3']},
    ]

    results = [
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Machine AWF #1', 'date': '2025-12-16', 'line': '',
         'input_count': 110, 'ok_count': 71, 'ng_count': 39,
         'ng_rate_decimal': 0.3545, 'ng_rate_percent': 35.5,
         'metric_name': 'NG rate', 'metric_value': 35.5, 'unit': '%',
         'judgement': None,
         'ng_breakdown': {
             'BAKO FRF': {'count': 33, 'rate': 0.300},
             'BAKO FRF+SPL': {'count': 6, 'rate': 0.055},
             'BAKO THD': {'count': 0, 'rate': 0.0},
             'BAKO No sound': {'count': 0, 'rate': 0.0},
         },
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Row 1 Machine AWF #1']},
        {'result_id': 'res_2', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Machine AWF #2', 'date': '2025-12-16', 'line': '',
         'input_count': 116, 'ok_count': 68, 'ng_count': 48,
         'ng_rate_decimal': 0.4138, 'ng_rate_percent': 41.4,
         'metric_name': 'NG rate', 'metric_value': 41.4, 'unit': '%',
         'judgement': None,
         'ng_breakdown': {
             'BAKO FRF': {'count': 39, 'rate': 0.336},
             'BAKO FRF+SPL': {'count': 9, 'rate': 0.078},
             'BAKO THD': {'count': 0, 'rate': 0.0},
             'BAKO No sound': {'count': 0, 'rate': 0.0},
         },
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Row 2 Machine AWF #2']},
        {'result_id': 'res_3', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Machine AWF #3', 'date': '2025-12-16', 'line': '',
         'input_count': 70, 'ok_count': 44, 'ng_count': 26,
         'ng_rate_decimal': 0.3714, 'ng_rate_percent': 37.1,
         'metric_name': 'NG rate', 'metric_value': 37.1, 'unit': '%',
         'judgement': None,
         'ng_breakdown': {
             'BAKO FRF': {'count': 18, 'rate': 0.257},
             'BAKO FRF+SPL': {'count': 8, 'rate': 0.114},
             'BAKO THD': {'count': 0, 'rate': 0.0},
             'BAKO No sound': {'count': 0, 'rate': 0.0},
         },
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Row 3 Machine AWF #3']},
    ]

    conclusions = [
        {'conclusion_id': 'concl_1', 'topic': 'AWF machine split — all three high',
         'statement_from_report': 'Follow result test machine AWF #1, #2, #3 very high.',
         'normalized_interpretation': 'No same-event Normal/Baseline in this report. Absolute NG rate by machine: AWF #2 41.4% (48/116) > AWF #3 37.1% (26/70) > AWF #1 35.5% (39/110). Defect mix dominated by BAKO FRF (25.7%~33.6%); BAKO FRF+SPL secondary (5.5%~11.4%). No machine is clearly OK — issue is not isolated to one AWF.',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['III. Result note', 'IV. Decision (empty)']},
    ]

    troubleshooting_index = {
        'defect_name': 'NG BAKO (Function)',
        'when_user_asks': ['NG Function BAKO high', 'BAKO FRF', 'AWF machine'],
        'suggested_checks': [
            {'hint_id': 'hint_1',
             'check_item': 'Investigate BAKO FRF root cause upstream of AWF (sub-assembly/Frame/SP), since all three AWF machines show high rate.',
             'reason': 'AWF #1 35.5%, #2 41.4%, #3 37.1% — close range, dominated by FRF (~25–34%). Splitting AWF did not isolate the defect.',
             'evidence_strength': 'medium', 'related_process': 'Function AWF / upstream sub-assembly',
             'related_part': 'SPK', 'source_file': name, 'sheet_name': sheet,
             'source_cells': ['Result rows 1-3']},
            {'hint_id': 'hint_2',
             'check_item': 'Add a same-event Normal/baseline lot in next AWF separation test.',
             'reason': 'No baseline present, so improvement/worsening vs Normal cannot be measured; only absolute high level confirmed.',
             'evidence_strength': 'high', 'related_process': 'Test design',
             'related_part': '', 'source_file': name, 'sheet_name': sheet,
             'source_cells': ['IV. Decision (empty)']},
        ],
        'limitations': ['IV. Decision section is empty.', 'No baseline NG rate provided in the same event.'],
    }

    ai_log = {
        'confidence': 0.7,
        'assumptions': ['Reported NG rates are accepted as-is (35.5%, 41.4%, 37.1%).'],
        'warnings': [
            'No same-event Normal/Baseline row; cannot compute relative change.',
            'IV. Decision section is empty; AWF separation alone did not isolate root cause.',
        ],
        'decision_rationale': 'Classified as ng_without_baseline: NG rates exist but no baseline row. Rank: AWF #2 41.4% > AWF #3 37.1% > AWF #1 35.5%; dominant defect BAKO FRF.',
    }

    result = {'schema_version': '0.1', 'document': doc, 'test_conditions': test_conditions,
              'results': results, 'conclusions': conclusions,
              'troubleshooting_index': troubleshooting_index, 'ai_extraction_log': ai_log}

    tr_en = {
        'document': {'title': doc['title'], 'purpose': doc['purpose'], 'content': doc['content']},
        'conclusions': {'concl_1': {'topic': conclusions[0]['topic'],
                                     'statement_from_report': conclusions[0]['statement_from_report'],
                                     'normalized_interpretation': conclusions[0]['normalized_interpretation']}},
        'hints': {'hint_1': {'check_item': troubleshooting_index['suggested_checks'][0]['check_item'],
                              'reason': troubleshooting_index['suggested_checks'][0]['reason']},
                  'hint_2': {'check_item': troubleshooting_index['suggested_checks'][1]['check_item'],
                              'reason': troubleshooting_index['suggested_checks'][1]['reason']}},
        'log': {'assumptions': ai_log['assumptions'], 'warnings': ai_log['warnings'],
                'decision_rationale': ai_log['decision_rationale']},
    }
    tr_ko = {
        'document': {
            'title': 'TIU L5S3-01 [R] NG Function 원인 찾기 시험 리포트 (BAKO)',
            'purpose': 'NG Function BAKO 가 ~40% 로 매우 높음. AWF #1/#2/#3 분리해 기능 재검으로 원인 파악.',
            'content': [
                "Function NG BAKO 기저가 ~40% (Purpose 명시).",
                "16-Dec, AWF #1/#2/#3 분리해 Function 재검; 머신에 따른 NG 변화 확인.",
            ],
        },
        'conclusions': {'concl_1': {
            'topic': 'AWF 머신 분리 — 세 대 모두 높음',
            'statement_from_report': '시험 결과: AWF #1, #2, #3 모두 NG 매우 높음.',
            'normalized_interpretation': '동일 이벤트 Normal/Baseline 없음. 머신별 절대 NG rate: AWF #2 41.4% (48/116) > AWF #3 37.1% (26/70) > AWF #1 35.5% (39/110). 주 결함 BAKO FRF (25.7~33.6%), 부차 BAKO FRF+SPL (5.5~11.4%). 어느 AWF 도 정상 수준 아님 — AWF 한 대 문제 아님.',
        }},
        'hints': {
            'hint_1': {'check_item': 'AWF 상류(sub-assembly/Frame/SP) 에서 BAKO FRF 근본 원인 조사.',
                       'reason': 'AWF #1 35.5%, #2 41.4%, #3 37.1% 로 차이 작고 FRF(25~34%) 위주. AWF 분리만으로 결함 격리 안 됨.'},
            'hint_2': {'check_item': '다음 AWF 분리 시험에 동일 이벤트 Normal/baseline lot 포함.',
                       'reason': 'Baseline 부재로 Normal 대비 개선/악화 측정 불가; 절대값 높다는 사실만 확인 가능.'},
        },
        'log': {
            'assumptions': ['리포트 NG rate (35.5%, 41.4%, 37.1%) 수용.'],
            'warnings': ['동일 이벤트 Normal/Baseline row 없음 — 상대 변화율 계산 불가.',
                         'IV. Decision 비어 있음. AWF 분리만으로 근본 원인 격리 안 됨.'],
            'decision_rationale': 'ng_without_baseline 분류: NG rate 만 있고 baseline 없음. 순위: AWF #2 > #3 > #1; 주 결함 BAKO FRF.',
        },
    }
    tr_vi = {
        'document': {
            'title': 'Báo cáo TIU L5S3-01 [R] tìm nguyên nhân NG Function (BAKO) cao',
            'purpose': 'NG Function BAKO rất cao (~40%). Tách máy AWF #1/#2/#3 và kiểm lại để tìm nguyên nhân.',
            'content': [
                "Mức nền NG BAKO ~40% (nêu trong Purpose).",
                "Ngày 16-Dec tách máy AWF #1/#2/#3 và kiểm Function lại để xem NG có thay đổi theo máy không.",
            ],
        },
        'conclusions': {'concl_1': {
            'topic': 'Tách máy AWF — cả ba đều cao',
            'statement_from_report': 'Theo kết quả: AWF #1, #2, #3 đều NG rất cao.',
            'normalized_interpretation': 'Không có Normal/Baseline cùng đợt. NG rate tuyệt đối theo máy: AWF #2 41.4% (48/116) > AWF #3 37.1% (26/70) > AWF #1 35.5% (39/110). Defect chủ đạo BAKO FRF (25.7~33.6%), thứ nhì BAKO FRF+SPL (5.5~11.4%). Không có AWF nào sạch — không phải lỗi của một máy.',
        }},
        'hints': {
            'hint_1': {'check_item': 'Điều tra nguyên nhân gốc BAKO FRF ở thượng nguồn AWF (sub-assembly/Frame/SP).',
                       'reason': 'AWF #1 35.5%, #2 41.4%, #3 37.1% — gần nhau, chủ yếu FRF (~25–34%). Tách AWF không cô lập được defect.'},
            'hint_2': {'check_item': 'Lần tách AWF tới nên có Normal/baseline cùng đợt.',
                       'reason': 'Thiếu baseline nên không tính được thay đổi tương đối; chỉ xác nhận mức tuyệt đối cao.'},
        },
        'log': {
            'assumptions': ['Chấp nhận NG rate báo cáo (35.5%, 41.4%, 37.1%).'],
            'warnings': ['Không có Normal/Baseline cùng đợt — không tính được tỉ lệ thay đổi.',
                         'Phần IV. Decision trống; tách AWF chưa cô lập được nguyên nhân.'],
            'decision_rationale': 'Phân loại ng_without_baseline: có NG rate nhưng không có baseline. Xếp: AWF #2 > #3 > #1; defect chủ đạo BAKO FRF.',
        },
    }
    return name, result, tr_ko, tr_en, tr_vi


# ---------------------------------------------------------------------
# 3. MSU-L20S15-07GMI — NTI lot test new bond (reliability_spec / multi-arm)
# ---------------------------------------------------------------------
def build_msu_nti_bond():
    name = '28.1 MSU-L20S15-07GMI Result check NTI lot test New Bond PW 1470SX-N1 and A-3424B- 2025.07.08'
    sheet = 'GRAPH'

    doc = _doc(
        name=name,
        title='MSU-L20S15-07GMI NTI Lot Test New Bond PW 1470SX-N1 vs A-3424B',
        model='MSU-L20S15-07GMI',
        report_date='2025-07-08',
        dept='ME', marker='', line='',
        report_type='reliability_spec',
        primary_canonical='Bond Material Evaluation',
        primary_aliases=['New Bond PW 1470SX-N1', 'New Bond A-3424B'],
        related=['SPL deviation', 'Fo deviation', 'THD deviation'],
        parts=['VP+CD bond', 'Speaker'],
        processes=['Bond application', 'NTI acoustic measurement'],
        purpose='Compare new bond candidates PW 1470SX-N1 and A-3424B against STD and Normal on NTI acoustic spec (SPL 100~750Hz / 800~1.5kHz / 1.6~14kHz, Fo, THD 200/400/1000Hz).',
        content=[
            "NTI measurement under JIG=BRS-201506 Sub Baffle, drive=3.58V(1.6W), distance=1cm.",
            "Four arms reported: STD (n=2), Normal (n=32), PW 1470SX-N1 (n=54), A-3424B (n=76).",
            "Spec centers: SPL 109.2/118.3/115.8 dB, Fo 675 Hz, THD 200/400/1000Hz = 45/25/8.",
        ],
    )

    test_conditions = [
        {'condition_id': 'cond_1', 'condition_group': 'New bond candidate evaluation',
         'line': '', 'process': 'Bond application', 'changed_factor': 'Bond material',
         'before_value': 'STD / Normal bond', 'after_value': 'PW 1470SX-N1 / A-3424B',
         'unit': None, 'machine': None, 'jig': 'BRS-201506 Sub Baffle JIG',
         'material_lot': 'NTI lot', 'supplier': None, 'dry_time_sec': None,
         'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': name, 'sheet_name': sheet,
         'source_cells': ['GRAPH header bond labels', 'NTI DATA jig/voltage/distance']},
    ]

    # Spec center deltas (averages from GRAPH summary).
    # SPL_100-750Hz center 109.2: STD 107.47 (-1.73), Normal 106.87 (-2.33),
    #   PW 107.26 (-1.94), A-3424B 107.17 (-2.03)
    # SPL_800-1.5k center 118.3: STD 116.23 (-2.07), Normal 115.63 (-2.67),
    #   PW 115.31 (-2.99), A-3424B 115.46 (-2.84)
    # SPL_1.6-14k center 115.8: STD 114.07 (-1.73), Normal 113.65 (-2.15),
    #   PW 113.41 (-2.39), A-3424B 113.46 (-2.34)
    # Fo center 675: STD 654.74 (-20.26), Normal 654.36 (-20.64),
    #   PW 630.36 (-44.64), A-3424B 637.95 (-37.05)
    # THD 200 center 45: STD 30.14 (-14.86), Normal 32.33 (-12.67),
    #   PW 35.30 (-9.70), A-3424B 32.46 (-12.54)
    # THD 400 center 25: STD 9.98 (-15.02), Normal 11.70 (-13.30),
    #   PW 12.88 (-12.12), A-3424B 12.30 (-12.70)
    # THD 1000 center 8: STD 3.02 (-4.98), Normal 4.00 (-4.00),
    #   PW 4.04 (-3.96), A-3424B 3.55 (-4.45)
    arms = [
        ('STD', 'res_std_'),
        ('Normal', 'res_nrm_'),
        ('PW 1470SX-N1', 'res_pw_'),
        ('A-3424B', 'res_a_'),
    ]
    spl1 = {'STD': 107.47, 'Normal': 106.87, 'PW 1470SX-N1': 107.26, 'A-3424B': 107.17}
    spl2 = {'STD': 116.23, 'Normal': 115.63, 'PW 1470SX-N1': 115.31, 'A-3424B': 115.46}
    spl3 = {'STD': 114.07, 'Normal': 113.65, 'PW 1470SX-N1': 113.41, 'A-3424B': 113.46}
    fo = {'STD': 654.74, 'Normal': 654.36, 'PW 1470SX-N1': 630.36, 'A-3424B': 637.95}
    thd200 = {'STD': 30.14, 'Normal': 32.33, 'PW 1470SX-N1': 35.30, 'A-3424B': 32.46}
    thd400 = {'STD': 9.98, 'Normal': 11.70, 'PW 1470SX-N1': 12.88, 'A-3424B': 12.30}
    thd1k = {'STD': 3.02, 'Normal': 4.00, 'PW 1470SX-N1': 4.04, 'A-3424B': 3.55}

    results = []
    rid = 1
    for arm, _ in arms:
        for metric_name, mp, unit, center in [
            ('SPL 100~750Hz', spl1[arm], 'dB', 109.2),
            ('SPL 800~1.5kHz', spl2[arm], 'dB', 118.3),
            ('SPL 1.6~14kHz', spl3[arm], 'dB', 115.8),
            ('Fo', fo[arm], 'Hz', 675.0),
            ('THD 200Hz', thd200[arm], '%', 45.0),
            ('THD 400Hz', thd400[arm], '%', 25.0),
            ('THD 1000Hz', thd1k[arm], '%', 8.0),
        ]:
            results.append({
                'result_id': f'res_{rid}', 'condition_id': 'cond_1',
                'measurement_type': 'Acoustic spec', 'condition_group': arm,
                'date': '2025-07-08', 'line': '',
                'input_count': None, 'ok_count': None, 'ng_count': None,
                'ng_rate_decimal': None, 'ng_rate_percent': None,
                'metric_name': metric_name, 'metric_value': mp, 'unit': unit,
                'judgement': None, 'ng_breakdown': {},
                'source_file': name, 'sheet_name': sheet,
                'source_cells': [f'GRAPH avg row {arm}'],
            })
            rid += 1

    conclusions = [
        {'conclusion_id': 'concl_1',
         'topic': 'PW 1470SX-N1 vs Normal — Fo drift',
         'statement_from_report': 'PW 1470SX-N1 avg Fo 630.36 Hz (center 675, delta -44.64) vs Normal 654.36 Hz (delta -20.64).',
         'normalized_interpretation': 'PW 1470SX-N1 Fo deviation -44.64 Hz is roughly 2.16x larger than Normal -20.64 Hz; both are below center but PW is the worst arm on Fo. THD 200Hz on PW 1470SX-N1 is 35.30% vs Normal 32.33% (i.e. higher distortion).',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['GRAPH Fo/THD rows']},
        {'conclusion_id': 'concl_2',
         'topic': 'A-3424B vs Normal — closer to Normal than PW',
         'statement_from_report': 'A-3424B avg Fo 637.95 Hz (delta -37.05) vs Normal 654.36 Hz (delta -20.64).',
         'normalized_interpretation': 'A-3424B Fo deviation -37.05 Hz is 1.79x worse than Normal, but better than PW 1470SX-N1 (-44.64). On SPL bands A-3424B 107.17/115.46/113.46 dB is within ~0.3 dB of Normal. THD profile of A-3424B is similar to Normal.',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['GRAPH avg rows']},
        {'conclusion_id': 'concl_3',
         'topic': 'Multi-arm ranking by spec adherence',
         'statement_from_report': '(Derived from GRAPH summary deltas)',
         'normalized_interpretation': 'Ranking from closest-to-spec to worst on Fo: Normal (-20.64) ≈ STD (-20.26) > A-3424B (-37.05) > PW 1470SX-N1 (-44.64). On SPL the four arms are within ~0.5 dB of each other.',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['GRAPH delta column']},
    ]

    troubleshooting_index = {
        'defect_name': 'Bond Material Evaluation',
        'when_user_asks': ['new bond candidate', 'PW 1470SX-N1', 'A-3424B', 'Fo drop'],
        'suggested_checks': [
            {'hint_id': 'hint_1',
             'check_item': 'Re-evaluate PW 1470SX-N1 Fo behaviour before adoption.',
             'reason': 'PW 1470SX-N1 Fo delta -44.64 Hz is 2.16x worse than Normal -20.64 Hz; THD 200Hz also 35.3% vs Normal 32.3%.',
             'evidence_strength': 'medium', 'related_process': 'Bond application',
             'related_part': 'VP+CD bond', 'source_file': name, 'sheet_name': sheet,
             'source_cells': ['GRAPH PW 1470SX-N1 row']},
            {'hint_id': 'hint_2',
             'check_item': 'A-3424B is a candidate but verify Fo on larger lot (n=76) is stable.',
             'reason': 'A-3424B Fo -37.05 Hz is closer to Normal than PW but still drifts more than STD/Normal.',
             'evidence_strength': 'medium', 'related_process': 'Bond application',
             'related_part': 'VP+CD bond', 'source_file': name, 'sheet_name': sheet,
             'source_cells': ['GRAPH A-3424B row']},
        ],
        'limitations': ['Workbook has no IV. Decision / verbal conclusion text; classification is from acoustic deltas only.', 'Sample sizes differ across arms (STD 2 / Normal 32 / PW 54 / A-3424B 76).'],
    }

    ai_log = {
        'confidence': 0.65,
        'assumptions': ['Spec centers SPL 109.2/118.3/115.8, Fo 675, THD 45/25/8 are taken from sheet header.',
                        'Arm sizes assumed from the leading index columns (2/32/54/76).'],
        'warnings': ['Report has no narrative conclusion in the workbook text; interpretation is from delta numbers only.',
                     'Small STD sample (n=2) makes STD a weak baseline.'],
        'decision_rationale': 'Classified as reliability_spec (multi-arm bond evaluation against spec gates). Ranking by Fo deviation: Normal/STD better than A-3424B better than PW 1470SX-N1. SPL bands tightly clustered.',
    }

    result = {'schema_version': '0.1', 'document': doc, 'test_conditions': test_conditions,
              'results': results, 'conclusions': conclusions,
              'troubleshooting_index': troubleshooting_index, 'ai_extraction_log': ai_log}

    tr_en = {
        'document': {'title': doc['title'], 'purpose': doc['purpose'], 'content': doc['content']},
        'conclusions': {c['conclusion_id']: {'topic': c['topic'],
                                              'statement_from_report': c['statement_from_report'],
                                              'normalized_interpretation': c['normalized_interpretation']}
                        for c in conclusions},
        'hints': {h['hint_id']: {'check_item': h['check_item'], 'reason': h['reason']}
                  for h in troubleshooting_index['suggested_checks']},
        'log': {'assumptions': ai_log['assumptions'], 'warnings': ai_log['warnings'],
                'decision_rationale': ai_log['decision_rationale']},
    }
    tr_ko = {
        'document': {
            'title': 'MSU-L20S15-07GMI NTI lot 신규 본드 PW 1470SX-N1 vs A-3424B 평가',
            'purpose': '신규 본드 후보 PW 1470SX-N1 과 A-3424B 를 STD/Normal 과 NTI 음향 스펙(SPL 100~750Hz / 800~1.5kHz / 1.6~14kHz, Fo, THD 200/400/1000Hz) 으로 비교.',
            'content': [
                "JIG=BRS-201506 Sub Baffle, 인가=3.58V(1.6W), 거리=1cm 의 NTI 측정.",
                "4 arm: STD (n=2), Normal (n=32), PW 1470SX-N1 (n=54), A-3424B (n=76).",
                "스펙 중심: SPL 109.2/118.3/115.8 dB, Fo 675 Hz, THD 200/400/1000Hz = 45/25/8.",
            ],
        },
        'conclusions': {
            'concl_1': {'topic': 'PW 1470SX-N1 vs Normal — Fo 변동',
                        'statement_from_report': 'PW 1470SX-N1 평균 Fo 630.36 Hz (중심 675, 편차 -44.64) vs Normal 654.36 Hz (편차 -20.64).',
                        'normalized_interpretation': 'PW 1470SX-N1 의 Fo 편차 -44.64 Hz 는 Normal -20.64 Hz 대비 약 2.16배 큼; 둘 다 중심보다 낮지만 PW 가 Fo 최악 arm. THD 200Hz 도 PW 35.30% vs Normal 32.33% 로 더 큼.'},
            'concl_2': {'topic': 'A-3424B vs Normal — PW 보다 Normal 에 근접',
                        'statement_from_report': 'A-3424B 평균 Fo 637.95 Hz (편차 -37.05) vs Normal 654.36 Hz (편차 -20.64).',
                        'normalized_interpretation': 'A-3424B Fo 편차 -37.05 Hz 는 Normal 대비 1.79배 악화이나 PW (-44.64) 보다 양호. SPL 대역에서는 A-3424B (107.17/115.46/113.46 dB) 가 Normal 과 ~0.3 dB 이내. THD 분포도 Normal 과 유사.'},
            'concl_3': {'topic': 'Multi-arm 스펙 적합도 순위',
                        'statement_from_report': '(GRAPH 요약 편차에서 도출)',
                        'normalized_interpretation': 'Fo 기준 스펙 근접 순위: Normal (-20.64) ≈ STD (-20.26) > A-3424B (-37.05) > PW 1470SX-N1 (-44.64). SPL 은 4 arm 모두 ~0.5 dB 이내.'},
        },
        'hints': {
            'hint_1': {'check_item': 'PW 1470SX-N1 Fo 거동을 채택 전 재평가.',
                       'reason': 'PW 1470SX-N1 Fo 편차 -44.64 Hz 가 Normal -20.64 Hz 대비 2.16배 악화; THD 200Hz 도 35.3% vs Normal 32.3%.'},
            'hint_2': {'check_item': 'A-3424B 는 후보지만 더 큰 lot 으로 Fo 안정성 확인 (n=76).',
                       'reason': 'A-3424B Fo -37.05 Hz 가 PW 보다는 Normal 에 가깝지만 여전히 STD/Normal 보다 더 흔들림.'},
        },
        'log': {
            'assumptions': ['스펙 중심 SPL 109.2/118.3/115.8, Fo 675, THD 45/25/8 은 시트 헤더 사용.',
                            'Arm 크기는 선두 인덱스 컬럼 (2/32/54/76) 으로 추정.'],
            'warnings': ['워크북에 IV. Decision/서술 결론 없음; 해석은 편차 수치 기반.',
                         'STD 표본(n=2)이 작아 STD 는 약한 baseline.'],
            'decision_rationale': 'reliability_spec (스펙 게이트 기반 multi-arm 본드 평가) 분류. Fo 편차 순위: Normal/STD > A-3424B > PW 1470SX-N1; SPL 대역은 매우 좁게 클러스터링.',
        },
    }
    tr_vi = {
        'document': {
            'title': 'MSU-L20S15-07GMI Đánh giá keo mới PW 1470SX-N1 vs A-3424B (NTI lot)',
            'purpose': 'So sánh keo mới PW 1470SX-N1 và A-3424B với STD/Normal theo spec acoustic NTI (SPL 100~750Hz / 800~1.5kHz / 1.6~14kHz, Fo, THD 200/400/1000Hz).',
            'content': [
                "Đo NTI dưới JIG=BRS-201506 Sub Baffle, drive=3.58V(1.6W), khoảng cách 1cm.",
                "4 arm: STD (n=2), Normal (n=32), PW 1470SX-N1 (n=54), A-3424B (n=76).",
                "Tâm spec: SPL 109.2/118.3/115.8 dB, Fo 675 Hz, THD 200/400/1000Hz = 45/25/8.",
            ],
        },
        'conclusions': {
            'concl_1': {'topic': 'PW 1470SX-N1 vs Normal — Fo lệch',
                        'statement_from_report': 'PW 1470SX-N1 Fo trung bình 630.36 Hz (tâm 675, lệch -44.64) so với Normal 654.36 Hz (lệch -20.64).',
                        'normalized_interpretation': 'Độ lệch Fo của PW 1470SX-N1 -44.64 Hz lớn gấp ~2.16 lần Normal -20.64 Hz; cả hai đều thấp hơn tâm nhưng PW là arm xấu nhất về Fo. THD 200Hz của PW 35.30% so với Normal 32.33%.'},
            'concl_2': {'topic': 'A-3424B vs Normal — gần Normal hơn PW',
                        'statement_from_report': 'A-3424B Fo trung bình 637.95 Hz (lệch -37.05) so với Normal 654.36 Hz (lệch -20.64).',
                        'normalized_interpretation': 'Lệch Fo A-3424B -37.05 Hz xấu hơn Normal 1.79x nhưng tốt hơn PW 1470SX-N1 (-44.64). Ở dải SPL, A-3424B (107.17/115.46/113.46 dB) chỉ lệch ~0.3 dB so với Normal. THD tương tự Normal.'},
            'concl_3': {'topic': 'Xếp hạng multi-arm theo độ bám spec',
                        'statement_from_report': '(Suy ra từ delta tổng hợp GRAPH)',
                        'normalized_interpretation': 'Xếp hạng bám spec theo Fo: Normal (-20.64) ≈ STD (-20.26) > A-3424B (-37.05) > PW 1470SX-N1 (-44.64). Ở SPL, 4 arm nằm trong ~0.5 dB.'},
        },
        'hints': {
            'hint_1': {'check_item': 'Đánh giá lại Fo của PW 1470SX-N1 trước khi đưa vào sử dụng.',
                       'reason': 'Lệch Fo của PW 1470SX-N1 -44.64 Hz xấu gấp 2.16x Normal -20.64 Hz; THD 200Hz 35.3% so với Normal 32.3%.'},
            'hint_2': {'check_item': 'A-3424B là ứng viên nhưng cần kiểm Fo ổn định trên lot lớn hơn (n=76).',
                       'reason': 'Lệch Fo A-3424B -37.05 Hz gần Normal hơn PW nhưng vẫn lệch nhiều hơn STD/Normal.'},
        },
        'log': {
            'assumptions': ['Tâm spec SPL 109.2/118.3/115.8, Fo 675, THD 45/25/8 lấy từ header sheet.',
                            'Cỡ mẫu các arm suy từ cột chỉ số (2/32/54/76).'],
            'warnings': ['Báo cáo không có phần IV. Decision/diễn giải; phân tích dựa trên delta số.',
                         'Mẫu STD nhỏ (n=2) nên STD là baseline yếu.'],
            'decision_rationale': 'Phân loại reliability_spec (đánh giá multi-arm keo so với spec gate). Xếp theo độ lệch Fo: Normal/STD > A-3424B > PW 1470SX-N1; dải SPL cụm rất gần.',
        },
    }
    return name, result, tr_ko, tr_en, tr_vi


# ---------------------------------------------------------------------
# Shared helpers for the next set of build_fn's.
# ---------------------------------------------------------------------
def _trio_from_en(doc_en, conclusions, hints, log):
    """Build the en mirror dict from the base lists."""
    return {
        'document': doc_en,
        'conclusions': {c['conclusion_id']: {'topic': c['topic'],
                                              'statement_from_report': c['statement_from_report'],
                                              'normalized_interpretation': c['normalized_interpretation']}
                        for c in conclusions},
        'hints': {h['hint_id']: {'check_item': h['check_item'], 'reason': h['reason']}
                  for h in hints},
        'log': {'assumptions': log['assumptions'], 'warnings': log['warnings'],
                'decision_rationale': log['decision_rationale']},
    }


# ---------------------------------------------------------------------
# 4 + 7: BRS-201506 DOE Forming sub VP (two duplicate datasets — same content).
# ---------------------------------------------------------------------
def _brs_201506_payload(name, sheet_label):
    doc = _doc(
        name=name,
        title='BRS-201506 Report Test DOE Forming sub VP',
        model='BRS-201506',
        report_date='2024-03-14',
        dept='ME', marker='Nhung', line='',
        report_type='doe_matrix',
        primary_canonical='NG SPK/Module high (Forming sub VP DOE)',
        primary_aliases=['Module NG high', 'VP Forming defect'],
        related=['Broken VP', 'VP deform', 'Particle', 'Laser cutting burr', 'Separate VP',
                 'NG Hearing Noise', 'NG Hearing Touch'],
        parts=['VP sub5', 'Forming machine'],
        processes=['Forming VP', 'SPK Function'],
        purpose='NG SPK and Module high — find reason via DOE on VP forming Temp_forming × Temp_cooling combinations.',
        content=[
            "DOE on sub5 VP line: 48 pcs/type; if 1 type breaks => stop.",
            "If VP forming OK, move to SPK line Function check then module.",
            "Forming factor1=Temp forming (130~150 / 150~170 / 170~190 / 190~210 / 210~230 °C); factor2=Temp cooling (30~50 / 50~70 / 70~90 / 90~110 / 110~130 °C). Normal=170~190 / 30~50.",
        ],
    )

    test_conditions = [
        {'condition_id': 'cond_1', 'condition_group': 'DOE Forming VP',
         'line': 'VP line', 'process': 'Forming VP', 'changed_factor': 'Temp forming × Temp cooling',
         'before_value': 'Normal 170~190°C / 30~50°C', 'after_value': 'Various combos (see grid)',
         'unit': '°C', 'machine': 'Forming machine (No14-2H + all machines)', 'jig': None,
         'material_lot': None, 'supplier': None,
         'dry_time_sec': None, 'temperature': '170~190 / 190~210 / 210~230 °C forming; 30~50 / 50~70 / 70~90 / 90~110 / 110~130 °C cooling',
         'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': name, 'sheet_name': sheet_label,
         'source_cells': ['II. Content', 'III.1 Result DOE rows']},
    ]

    # VP line DOE NG (3/14/2024, 48 pcs/cell, total NG, NG Rate)
    # Normal 170~190 / 30~50: 48/44, total 4, 8.3%.
    # Cells (Forming/Cooling) → NG Rate%:
    cells_3_14 = [
        ('190~210', '30~50', 48, 45, 3, 6.2),
        ('210~230', '30~50', 48, 46, 2, 4.2),
        ('170~190', '50~70', 48, 45, 3, 6.2),
        ('170~190', '70~90', 48, 45, 3, 6.2),
        ('170~190', '90~110', 48, 43, 5, 10.4),  # marked yellow (worst)
        ('170~190', '110~130', 48, 45, 3, 6.2),
    ]
    results = []
    rid = 1
    # Normal first
    results.append({
        'result_id': f'res_{rid}', 'condition_id': 'cond_1',
        'measurement_type': 'VP forming NG',
        'condition_group': 'Normal 170~190°C / 30~50°C',
        'date': '2024-03-14', 'line': 'VP line',
        'input_count': 48, 'ok_count': 44, 'ng_count': 4,
        'ng_rate_decimal': 0.083, 'ng_rate_percent': 8.3,
        'metric_name': 'NG Rate', 'metric_value': 8.3, 'unit': '%',
        'judgement': None,
        'ng_breakdown': {'Broken': 0, 'VP deform': 1, 'Particle': 2,
                         'Laser cutting burr': 1, 'Separate VP': 0},
        'source_file': name, 'sheet_name': sheet_label,
        'source_cells': ['DOE row Normal 3/14']})
    rid += 1
    for tf, tc, inp, ok, ng, rate in cells_3_14:
        results.append({
            'result_id': f'res_{rid}', 'condition_id': 'cond_1',
            'measurement_type': 'VP forming NG',
            'condition_group': f'VP Test {tf}°C / {tc}°C',
            'date': '2024-03-14', 'line': 'VP line',
            'input_count': inp, 'ok_count': ok, 'ng_count': ng,
            'ng_rate_decimal': rate/100.0, 'ng_rate_percent': rate,
            'metric_name': 'NG Rate', 'metric_value': rate, 'unit': '%',
            'judgement': 'CHECK',
            'ng_breakdown': {},
            'source_file': name, 'sheet_name': sheet_label,
            'source_cells': [f'DOE row {tf}/{tc} 3/14']})
        rid += 1
    # 3/18 (No14-2H 210~230 / 30~50 — 336/328 NG 8, 2.4%; 170~190/90~110 — 312/295 NG 17, 5.4%)
    results.append({
        'result_id': f'res_{rid}', 'condition_id': 'cond_1',
        'measurement_type': 'VP forming NG',
        'condition_group': 'VP Test 210~230°C / 30~50°C (No14-2H)',
        'date': '2024-03-18', 'line': 'VP line',
        'input_count': 336, 'ok_count': 328, 'ng_count': 8,
        'ng_rate_decimal': 0.024, 'ng_rate_percent': 2.4,
        'metric_name': 'NG Rate', 'metric_value': 2.4, 'unit': '%',
        'judgement': 'PASS',
        'ng_breakdown': {'VP deform': 1, 'Particle': 6, 'Laser cutting burr': 1, 'Separate VP': 0},
        'source_file': name, 'sheet_name': sheet_label,
        'source_cells': ['DOE row 3/18 No14-2H 210~230/30~50']})
    rid += 1
    results.append({
        'result_id': f'res_{rid}', 'condition_id': 'cond_1',
        'measurement_type': 'VP forming NG',
        'condition_group': 'VP Test 170~190°C / 90~110°C',
        'date': '2024-03-18', 'line': 'VP line',
        'input_count': 312, 'ok_count': 295, 'ng_count': 17,
        'ng_rate_decimal': 0.054, 'ng_rate_percent': 5.4,
        'metric_name': 'NG Rate', 'metric_value': 5.4, 'unit': '%',
        'judgement': 'CHECK',
        'ng_breakdown': {'VP deform': 5, 'Particle': 7, 'Laser cutting burr': 1, 'Separate VP': 4},
        'source_file': name, 'sheet_name': sheet_label,
        'source_cells': ['DOE row 3/18 170~190/90~110']})
    rid += 1
    results.append({
        'result_id': f'res_{rid}', 'condition_id': 'cond_1',
        'measurement_type': 'VP forming NG',
        'condition_group': 'Normal 170~190°C / 30~50°C',
        'date': '2024-03-18', 'line': 'VP line',
        'input_count': 312, 'ok_count': 303, 'ng_count': 9,
        'ng_rate_decimal': 0.029, 'ng_rate_percent': 2.9,
        'metric_name': 'NG Rate', 'metric_value': 2.9, 'unit': '%',
        'judgement': None,
        'ng_breakdown': {'VP deform': 2, 'Particle': 5, 'Laser cutting burr': 2, 'Separate VP': 0},
        'source_file': name, 'sheet_name': sheet_label,
        'source_cells': ['DOE row 3/18 Normal']})
    rid += 1
    # 3/21 (All machine 170~190/90~110 4476/4360 NG 116 2.6%; Normal 170~190/30~50 5142/5035 NG 107 2.1%)
    results.append({
        'result_id': f'res_{rid}', 'condition_id': 'cond_1',
        'measurement_type': 'VP forming NG',
        'condition_group': 'VP Test 170~190°C / 90~110°C (All machines)',
        'date': '2024-03-21', 'line': 'VP line',
        'input_count': 4476, 'ok_count': 4360, 'ng_count': 116,
        'ng_rate_decimal': 0.026, 'ng_rate_percent': 2.6,
        'metric_name': 'NG Rate', 'metric_value': 2.6, 'unit': '%',
        'judgement': 'PASS',
        'ng_breakdown': {'VP deform': 10, 'Particle': 69, 'Laser cutting burr': 27, 'Separate VP': 10},
        'source_file': name, 'sheet_name': sheet_label,
        'source_cells': ['DOE row 3/21 all machines']})
    rid += 1
    results.append({
        'result_id': f'res_{rid}', 'condition_id': 'cond_1',
        'measurement_type': 'VP forming NG',
        'condition_group': 'Normal 170~190°C / 30~50°C',
        'date': '2024-03-21', 'line': 'VP line',
        'input_count': 5142, 'ok_count': 5035, 'ng_count': 107,
        'ng_rate_decimal': 0.021, 'ng_rate_percent': 2.1,
        'metric_name': 'NG Rate', 'metric_value': 2.1, 'unit': '%',
        'judgement': None,
        'ng_breakdown': {'VP deform': 12, 'Particle': 72, 'Laser cutting burr': 12, 'Separate VP': 11},
        'source_file': name, 'sheet_name': sheet_label,
        'source_cells': ['DOE row 3/21 Normal']})
    rid += 1
    # SPK function (16-Mar VP Test 170~190/90~110 43/42 NG 1, 2.3%; Normal 50/48 NG 2, 4.0%)
    results.append({
        'result_id': f'res_{rid}', 'condition_id': 'cond_1',
        'measurement_type': 'SPK Function',
        'condition_group': 'VP Test 170~190°C / 90~110°C',
        'date': '2024-03-16', 'line': 'SPK line',
        'input_count': 43, 'ok_count': 42, 'ng_count': 1,
        'ng_rate_decimal': 0.023, 'ng_rate_percent': 2.3,
        'metric_name': 'Total NG rate', 'metric_value': 2.3, 'unit': '%',
        'judgement': None,
        'ng_breakdown': {'Sigma SPL': 0, 'Sigma THD': 0, 'Sigma SPL+THD': 0,
                         'Sigma SPL+THD+F0': 0, 'Hearing Noise': 1, 'Hearing Touch': 0},
        'source_file': name, 'sheet_name': sheet_label,
        'source_cells': ['SPK row 3/16 VP Test']})
    rid += 1
    results.append({
        'result_id': f'res_{rid}', 'condition_id': 'cond_1',
        'measurement_type': 'SPK Function',
        'condition_group': 'Normal 170~190°C / 30~50°C',
        'date': '2024-03-16', 'line': 'SPK line',
        'input_count': 50, 'ok_count': 48, 'ng_count': 2,
        'ng_rate_decimal': 0.040, 'ng_rate_percent': 4.0,
        'metric_name': 'Total NG rate', 'metric_value': 4.0, 'unit': '%',
        'judgement': None,
        'ng_breakdown': {'Sigma SPL': 0, 'Sigma THD': 0, 'Sigma SPL+THD': 0,
                         'Sigma SPL+THD+F0': 0, 'Hearing Noise': 0, 'Hearing Touch': 2},
        'source_file': name, 'sheet_name': sheet_label,
        'source_cells': ['SPK row 3/16 Normal']})
    rid += 1
    # 20-Mar VP Test 170~190/90~110 293/285 NG 8 2.7%; Normal 800/765 NG 35 4.4%.
    results.append({
        'result_id': f'res_{rid}', 'condition_id': 'cond_1',
        'measurement_type': 'SPK Function',
        'condition_group': 'VP Test 170~190°C / 90~110°C',
        'date': '2024-03-20', 'line': 'SPK line',
        'input_count': 293, 'ok_count': 285, 'ng_count': 8,
        'ng_rate_decimal': 0.027, 'ng_rate_percent': 2.7,
        'metric_name': 'Total NG rate', 'metric_value': 2.7, 'unit': '%',
        'judgement': None,
        'ng_breakdown': {'Sigma SPL': 1, 'Sigma THD': 0, 'Sigma SPL+THD': 0,
                         'Sigma SPL+THD+F0': 0, 'Hearing Noise': 3, 'Hearing Touch': 4},
        'source_file': name, 'sheet_name': sheet_label,
        'source_cells': ['SPK row 3/20 VP Test']})
    rid += 1
    results.append({
        'result_id': f'res_{rid}', 'condition_id': 'cond_1',
        'measurement_type': 'SPK Function',
        'condition_group': 'Normal 170~190°C / 30~50°C',
        'date': '2024-03-20', 'line': 'SPK line',
        'input_count': 800, 'ok_count': 765, 'ng_count': 35,
        'ng_rate_decimal': 0.044, 'ng_rate_percent': 4.4,
        'metric_name': 'Total NG rate', 'metric_value': 4.4, 'unit': '%',
        'judgement': None,
        'ng_breakdown': {'Sigma SPL': 0, 'Sigma THD': 0, 'Sigma SPL+THD': 0,
                         'Sigma SPL+THD+F0': 0, 'Hearing Noise': 13, 'Hearing Touch': 22},
        'source_file': name, 'sheet_name': sheet_label,
        'source_cells': ['SPK row 3/20 Normal']})
    rid += 1
    # 23-Mar VP Test 170~190/90~110 3880/3744 NG 136 3.5%
    results.append({
        'result_id': f'res_{rid}', 'condition_id': 'cond_1',
        'measurement_type': 'SPK Function',
        'condition_group': 'VP Test 170~190°C / 90~110°C',
        'date': '2024-03-23', 'line': 'SPK line',
        'input_count': 3880, 'ok_count': 3744, 'ng_count': 136,
        'ng_rate_decimal': 0.035, 'ng_rate_percent': 3.5,
        'metric_name': 'Total NG rate', 'metric_value': 3.5, 'unit': '%',
        'judgement': None,
        'ng_breakdown': {'Hearing Noise': 57, 'Hearing Touch': 79},
        'source_file': name, 'sheet_name': sheet_label,
        'source_cells': ['SPK row 3/23 VP Test']})
    rid += 1

    conclusions = [
        {'conclusion_id': 'concl_1',
         'topic': 'VP forming DOE — no broken VP across cells',
         'statement_from_report': "All type test don't happen NG broken VP => NG same normal.",
         'normalized_interpretation': '3/14 DOE on 48 pcs/cell did not produce Broken VP in any of the 6 tested forming×cooling combinations. The worst NG Rate cell (10.4%, 170~190/90~110) vs same-day Normal 8.3% = (10.4/8.3 - 1)*100 = +25.3% worse than normal, but defect mix is non-broken (deform/particle/burr).',
         'source_file': name, 'sheet_name': sheet_label,
         'source_cells': ['Note column 3/14 DOE rows']},
        {'conclusion_id': 'concl_2',
         'topic': 'Scaled DOE — 170~190/90~110 vs Normal 170~190/30~50',
         'statement_from_report': 'NG same normal.',
         'normalized_interpretation': '3/18 VP forming NG: Test 170~190/90~110 5.4% (17/312) vs Normal 2.9% (9/312) = (5.4/2.9 - 1)*100 = +86% worse. 3/21 All-machine Test 2.6% (116/4476) vs Normal 2.1% (107/5142) = +23.8% worse. Report wording "same as normal" understates the relative worsening on the high-cooling cell.',
         'source_file': name, 'sheet_name': sheet_label,
         'source_cells': ['3/18, 3/21 DOE rows']},
        {'conclusion_id': 'concl_3',
         'topic': 'SPK Function vs Normal',
         'statement_from_report': '(No verbal decision text in IV. Decision.)',
         'normalized_interpretation': 'SPK Function on 170~190/90~110 vs same-day Normal (170~190/30~50): 3/16 Test 2.3% (1/43) vs Normal 4.0% (2/50) = -42.5% improved. 3/20 Test 2.7% (8/293) vs Normal 4.4% (35/800) = -38.6% improved. Dominant SPK NG = Hearing Noise + Touch.',
         'source_file': name, 'sheet_name': sheet_label,
         'source_cells': ['SPK rows 3/16, 3/20']},
    ]

    hints = [
        {'hint_id': 'hint_1',
         'check_item': 'Run a confirmation lot at Test forming 170~190°C / cooling 90~110°C and compare Function vs same-event Normal on equal sample size.',
         'reason': 'VP forming NG is +86% worse on 3/18 but SPK Function is -38.6% improved on 3/20. Different metric directions; need same-size baseline to decide.',
         'evidence_strength': 'medium', 'related_process': 'Forming VP / SPK Function',
         'related_part': 'VP sub5', 'source_file': name, 'sheet_name': sheet_label,
         'source_cells': ['3/18 + 3/20 rows']},
        {'hint_id': 'hint_2',
         'check_item': 'Investigate Particle and Laser-cutting-burr as the dominant VP forming defects (not Broken VP).',
         'reason': '3/14–3/21 VP defect mix is mostly Particle and Burr; Broken VP=0 across cells; deform secondary.',
         'evidence_strength': 'high', 'related_process': 'Forming VP / Laser cutting',
         'related_part': 'VP sub5', 'source_file': name, 'sheet_name': sheet_label,
         'source_cells': ['DOE NG breakdown columns']},
    ]
    troubleshooting_index = {
        'defect_name': 'NG SPK/Module high (Forming sub VP DOE)',
        'when_user_asks': ['VP forming DOE', 'NG SPK high', 'temp forming', 'temp cooling'],
        'suggested_checks': hints,
        'limitations': ['IV. Decision section is empty.', 'Function comparison uses Normal-30~50 baseline at different sample sizes from Test 90~110.'],
    }

    ai_log = {
        'confidence': 0.7,
        'assumptions': ['Same-day same-line Normal row (170~190 / 30~50) is the baseline for VP Test rows.',
                        'Cooling combos act as factor2, forming as factor1.'],
        'warnings': ['Report writes "NG same normal" but the worst forming×cooling cell (90~110) is +25–86% worse than normal.',
                     'IV. Decision is empty.'],
        'decision_rationale': 'Classified as doe_matrix: 2-D factor grid Temp_forming × Temp_cooling with Normal cell as reference. Worst VP-forming cell on 3/14 is 170~190/90~110 (10.4%); but at SPK Function on the same combo, Test is -38.6% improved vs Normal.',
    }

    result = {'schema_version': '0.1', 'document': doc, 'test_conditions': test_conditions,
              'results': results, 'conclusions': conclusions,
              'troubleshooting_index': troubleshooting_index, 'ai_extraction_log': ai_log}
    return result, conclusions, hints, ai_log, doc


def _brs_201506_translations(doc, conclusions, hints, log, ko_title, ko_purpose, ko_content,
                              vi_title, vi_purpose, vi_content,
                              ko_concls, ko_hints, ko_log,
                              vi_concls, vi_hints, vi_log):
    tr_en = _trio_from_en({'title': doc['title'], 'purpose': doc['purpose'], 'content': doc['content']},
                          conclusions, hints, log)
    tr_ko = {
        'document': {'title': ko_title, 'purpose': ko_purpose, 'content': ko_content},
        'conclusions': ko_concls, 'hints': ko_hints, 'log': ko_log,
    }
    tr_vi = {
        'document': {'title': vi_title, 'purpose': vi_purpose, 'content': vi_content},
        'conclusions': vi_concls, 'hints': vi_hints, 'log': vi_log,
    }
    return tr_ko, tr_en, tr_vi


def build_brs_201506_v1():
    name = '29-1. BRS-201506 Report test DOE Forming sub VP- date 14.3.2024'
    result, conclusions, hints, log, doc = _brs_201506_payload(name, 'DOE FM')

    ko_concls = {
        'concl_1': {'topic': 'VP forming DOE — 셀별 broken VP 없음',
                    'statement_from_report': '모든 타입에서 broken VP NG 미발생 => NG normal 과 동일.',
                    'normalized_interpretation': '3/14 DOE 48 pcs/cell 6 가지 forming×cooling 조합에서 broken VP 0. 최악 NG Rate cell (10.4%, 170~190/90~110) vs 같은 날 Normal 8.3% = (10.4/8.3 - 1)*100 = +25.3% 악화이나, defect 구성은 broken 아닌 deform/particle/burr.'},
        'concl_2': {'topic': '확대 DOE — 170~190/90~110 vs Normal 170~190/30~50',
                    'statement_from_report': 'NG normal 과 동일.',
                    'normalized_interpretation': '3/18 VP forming NG: Test 170~190/90~110 5.4% (17/312) vs Normal 2.9% (9/312) = (5.4/2.9 - 1)*100 = +86% 악화. 3/21 전 machine Test 2.6% (116/4476) vs Normal 2.1% (107/5142) = +23.8% 악화. 리포트 "same as normal" 은 고냉각 셀의 상대 악화를 과소표현.'},
        'concl_3': {'topic': 'SPK Function vs Normal',
                    'statement_from_report': '(IV. Decision 비어 있음.)',
                    'normalized_interpretation': 'SPK Function 170~190/90~110 vs 같은 날 Normal (170~190/30~50): 3/16 Test 2.3% (1/43) vs Normal 4.0% (2/50) = -42.5% 개선. 3/20 Test 2.7% (8/293) vs Normal 4.4% (35/800) = -38.6% 개선. SPK 주 NG = Hearing Noise + Touch.'},
    }
    ko_hints = {
        'hint_1': {'check_item': 'Test forming 170~190°C / cooling 90~110°C 확인 lot 으로 같은 표본 크기 Normal 과 Function 비교.',
                   'reason': 'VP forming NG 는 3/18 +86% 악화이나 SPK Function 은 3/20 -38.6% 개선 — 지표 방향 상이, 표본 동등 baseline 필요.'},
        'hint_2': {'check_item': 'VP forming 의 주 결함이 Particle/Laser-cut burr 임을 (Broken VP 아님) 확인하고 그 라인 점검.',
                   'reason': '3/14–3/21 VP defect 구성은 Particle/Burr 중심, Broken VP=0, deform 부차.'},
    }
    ko_log = {
        'assumptions': ['같은 날, 같은 라인 Normal row (170~190 / 30~50) 가 VP Test row 의 baseline.',
                        'Cooling 조합이 factor2, forming 이 factor1.'],
        'warnings': ['리포트는 "NG same normal" 이지만 최악 forming×cooling 셀(90~110) 은 normal 대비 +25~86% 악화.',
                     'IV. Decision 비어 있음.'],
        'decision_rationale': 'doe_matrix 분류: 2-D factor grid Temp_forming × Temp_cooling, Normal 셀이 기준. 3/14 VP forming 최악 셀 170~190/90~110 (10.4%); 그러나 같은 조합의 SPK Function 은 Normal 대비 -38.6% 개선.',
    }
    ko_title = 'BRS-201506 DOE Forming sub VP 시험 리포트'
    ko_purpose = 'NG SPK 및 Module 높음 — VP forming Temp_forming × Temp_cooling 조합 DOE 로 원인 추적.'
    ko_content = [
        "Sub5 VP line DOE: 48 pcs/type; 1 type 라도 broken 발생 시 중단.",
        "VP forming OK 면 SPK line Function → module 진행.",
        "Factor1=Temp forming(130~150/150~170/170~190/190~210/210~230°C); factor2=Temp cooling(30~50/50~70/70~90/90~110/110~130°C). Normal=170~190 / 30~50.",
    ]

    vi_concls = {
        'concl_1': {'topic': 'DOE Forming VP — không có VP broken ở các cell',
                    'statement_from_report': 'Không xảy ra NG VP broken => NG giống Normal.',
                    'normalized_interpretation': 'DOE 3/14 48 pcs/cell trên 6 tổ hợp forming×cooling không có VP broken nào. Cell xấu nhất theo NG Rate (10.4%, 170~190/90~110) so với Normal cùng ngày 8.3% = (10.4/8.3 - 1)*100 = +25.3% xấu hơn, nhưng defect là deform/particle/burr chứ không phải broken.'},
        'concl_2': {'topic': 'DOE quy mô lớn — 170~190/90~110 vs Normal 170~190/30~50',
                    'statement_from_report': 'NG giống Normal.',
                    'normalized_interpretation': 'VP forming NG ngày 3/18: Test 170~190/90~110 5.4% (17/312) so với Normal 2.9% (9/312) = (5.4/2.9 - 1)*100 = +86% xấu. 3/21 toàn bộ máy Test 2.6% (116/4476) so với Normal 2.1% (107/5142) = +23.8% xấu. Cách diễn đạt "same as normal" của báo cáo nhẹ hơn thực tế.'},
        'concl_3': {'topic': 'SPK Function vs Normal',
                    'statement_from_report': '(IV. Decision trống.)',
                    'normalized_interpretation': 'SPK Function 170~190/90~110 vs Normal cùng ngày (170~190/30~50): 3/16 Test 2.3% (1/43) so với Normal 4.0% (2/50) = -42.5% cải thiện. 3/20 Test 2.7% (8/293) so với Normal 4.4% (35/800) = -38.6% cải thiện. NG SPK chủ đạo = Hearing Noise + Touch.'},
    }
    vi_hints = {
        'hint_1': {'check_item': 'Chạy lot xác nhận Test forming 170~190°C / cooling 90~110°C, so sánh Function với Normal cùng đợt cỡ mẫu bằng nhau.',
                   'reason': 'VP forming NG 3/18 +86% xấu nhưng SPK Function 3/20 -38.6% cải thiện — hai chỉ số ngược chiều, cần baseline đồng quy mô.'},
        'hint_2': {'check_item': 'Tập trung điều tra Particle và Laser-cutting-burr (không phải Broken VP) là defect chính của VP forming.',
                   'reason': '3/14–3/21 defect VP chủ yếu Particle/Burr; Broken VP=0; deform thứ yếu.'},
    }
    vi_log = {
        'assumptions': ['Normal cùng ngày, cùng line (170~190 / 30~50) là baseline cho các row VP Test.',
                        'Tổ hợp cooling là factor2, forming là factor1.'],
        'warnings': ['Báo cáo viết "NG same normal" nhưng cell forming×cooling xấu nhất (90~110) xấu hơn Normal +25~86%.',
                     'IV. Decision trống.'],
        'decision_rationale': 'Phân loại doe_matrix: lưới 2D Temp_forming × Temp_cooling, cell Normal làm tham chiếu. Cell VP forming xấu nhất ngày 3/14 là 170~190/90~110 (10.4%); cùng tổ hợp đó ở SPK Function lại -38.6% cải thiện so với Normal.',
    }
    vi_title = 'Báo cáo BRS-201506 thử nghiệm DOE Forming sub VP'
    vi_purpose = 'NG SPK và Module cao — tìm nguyên nhân qua DOE forming với tổ hợp Temp_forming × Temp_cooling.'
    vi_content = [
        "DOE trên VP line sub5: 48 pcs/loại; nếu 1 loại có broken => dừng.",
        "Nếu VP forming OK chuyển sang SPK line Function rồi module.",
        "Factor1=Temp forming (130~150/150~170/170~190/190~210/210~230°C); factor2=Temp cooling (30~50/50~70/70~90/90~110/110~130°C). Normal=170~190 / 30~50.",
    ]

    tr_ko, tr_en, tr_vi = _brs_201506_translations(doc, conclusions, hints, log,
                                                    ko_title, ko_purpose, ko_content,
                                                    vi_title, vi_purpose, vi_content,
                                                    ko_concls, ko_hints, ko_log,
                                                    vi_concls, vi_hints, vi_log)
    return name, result, tr_ko, tr_en, tr_vi


def build_brs_201506_v2():
    name = '29. BRS-201506 Report test DOE Forming sub VP- date 14.3.2024'
    result, conclusions, hints, log, doc = _brs_201506_payload(name, 'DOE FM')
    # Reuse same KO/VI from v1 (workbook content is essentially identical;
    # v2 just has more grid cells exposed and same conclusions).
    _, _, tr_ko_v1, tr_en_v1, tr_vi_v1 = build_brs_201506_v1()
    # Re-target the dict to v2's name by rebuilding through the helper.
    # build_brs_201506_v1 already returns translations keyed by conclusion_id
    # which match v2's payload (same conclusion list). So we just return them.
    return name, result, tr_ko_v1, tr_en_v1, tr_vi_v1


# ---------------------------------------------------------------------
# 5. BRS-161014 — Over glue reinforcement
# ---------------------------------------------------------------------
def build_brs_161014_over_glue():
    name = '29. BRS-161014 Report TEST sample over glue reinforcement'
    sheet = 'Report (2)'

    doc = _doc(
        name=name,
        title='BRS-161014 Report Test Sample Over Glue Reinforcement',
        model='BRS-161014',
        report_date='2023-09-15',
        dept='ME', marker='Thao/Thuy', line='',
        report_type='normal_comparison',
        primary_canonical='Over Glue (reinforcement)',
        primary_aliases=['over glue', 'over glue reinforcement'],
        related=['NG Hearing Noise', 'NG Hearing Touch'],
        parts=['SPK', 'Glue'],
        processes=['Glue application (reinforcement)', 'Function check'],
        purpose='Decide whether over-glue reinforcement samples can be used or must be scrapped.',
        content=[
            "Separate over-glue samples, make final, then compare Function NG rate vs same-event Normal sample.",
            "Test: 98 pcs final 9/15/2023. Normal: 794 pcs same day.",
        ],
    )

    test_conditions = [
        {'condition_id': 'cond_1', 'condition_group': 'Over glue reinforcement',
         'line': '', 'process': 'Glue (reinforcement)', 'changed_factor': 'Glue amount status',
         'before_value': 'Normal glue', 'after_value': 'Over glue reinforcement',
         'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None,
         'dry_time_sec': None, 'temperature': None, 'pressure': None,
         'bond_amount': 'over (reinforcement)', 'uv_energy': None,
         'source_file': name, 'sheet_name': sheet,
         'source_cells': ['I.Purpose, II.Content']},
    ]

    results = [
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Test (over glue)', 'date': '2023-09-15', 'line': '',
         'input_count': 98, 'ok_count': 35, 'ng_count': 63,
         'ng_rate_decimal': 0.643, 'ng_rate_percent': 64.3,
         'metric_name': 'Total NG rate', 'metric_value': 64.3, 'unit': '%',
         'judgement': 'FAIL',
         'ng_breakdown': {'Sigma SPL': 0, 'Sigma THD': 0, 'Sigma SPL+THD': 0, 'Sigma SPL+THD+F0': 0,
                          'Hearing Noise': 45, 'Hearing Touch': 18, 'HOHD': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Result row Test']},
        {'result_id': 'res_2', 'condition_id': None, 'measurement_type': 'Function',
         'condition_group': 'Normal', 'date': '2023-09-15', 'line': '',
         'input_count': 794, 'ok_count': 457, 'ng_count': 337,
         'ng_rate_decimal': 0.424, 'ng_rate_percent': 42.4,
         'metric_name': 'Total NG rate', 'metric_value': 42.4, 'unit': '%',
         'judgement': None,
         'ng_breakdown': {'Sigma SPL': 1, 'Sigma THD': 1, 'Sigma SPL+THD': 0, 'Sigma SPL+THD+F0': 0,
                          'Hearing Noise': 256, 'Hearing Touch': 79, 'HOHD': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Result row Normal']},
    ]

    conclusions = [
        {'conclusion_id': 'concl_1',
         'topic': 'Over glue vs Normal — Function NG',
         'statement_from_report': 'NG rate of test lot is higher than normal lot.',
         'normalized_interpretation': 'Over glue Test 64.3% (63/98) vs same-event Normal 42.4% (337/794) = (64.3/42.4 - 1)*100 = +51.7% worse than normal. Defect mix dominated by Hearing Noise (Test 71.4% of NG vs Normal 76.0% of NG); Hearing Touch secondary. Over-glue reinforcement should NOT be reused.',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['III.Result note (blue text)']},
    ]
    hints = [
        {'hint_id': 'hint_1',
         'check_item': 'Reject over-glue reinforcement parts and do not reuse.',
         'reason': 'Function NG 64.3% (63/98) is 51.7% worse than same-event Normal 42.4% (337/794); same defect mix (Hearing Noise + Touch).',
         'evidence_strength': 'high', 'related_process': 'Glue (reinforcement)',
         'related_part': 'SPK', 'source_file': name, 'sheet_name': sheet,
         'source_cells': ['Test vs Normal rows']},
        {'hint_id': 'hint_2',
         'check_item': 'When Hearing Noise dominates NG, investigate dome/glue conditions before testing.',
         'reason': 'Both Test and Normal show Hearing Noise as 71.4%/76.0% of total NG, indicating a process-wide hearing issue independent of over-glue.',
         'evidence_strength': 'medium', 'related_process': 'Hearing',
         'related_part': 'SPK', 'source_file': name, 'sheet_name': sheet,
         'source_cells': ['Defect breakdown rows']},
    ]
    troubleshooting_index = {
        'defect_name': 'Over Glue (reinforcement)',
        'when_user_asks': ['over glue reuse', 'reinforcement glue acceptance'],
        'suggested_checks': hints,
        'limitations': ['IV. Decision section is empty; only III. note line.', 'Normal lot is much larger than test lot.'],
    }
    log = {
        'confidence': 0.85,
        'assumptions': ['Same-day Normal row is the baseline for over-glue Test row.'],
        'warnings': ['Both arms show high Hearing Noise (~30%+ of input).'],
        'decision_rationale': 'Classified normal_comparison: same-event Normal exists. (64.3/42.4 - 1)*100 = +51.7% worse, defect mix similar => over-glue makes the Hearing failure worse; reject over-glue reinforcement.',
    }

    result = {'schema_version': '0.1', 'document': doc, 'test_conditions': test_conditions,
              'results': results, 'conclusions': conclusions,
              'troubleshooting_index': troubleshooting_index, 'ai_extraction_log': log}

    tr_en = _trio_from_en({'title': doc['title'], 'purpose': doc['purpose'], 'content': doc['content']},
                          conclusions, hints, log)
    tr_ko = {
        'document': {
            'title': 'BRS-161014 Over Glue 보강 샘플 시험 리포트',
            'purpose': '오버 글루로 보강된 샘플을 사용할 수 있는지(또는 폐기해야 하는지) 결정.',
            'content': [
                "오버 글루 샘플을 분리해 최종 조립 후, 같은 이벤트 Normal 샘플과 Function NG rate 비교.",
                "Test: 9/15/2023 최종 98 pcs. Normal: 같은 날 794 pcs.",
            ],
        },
        'conclusions': {'concl_1': {
            'topic': 'Over glue vs Normal — Function NG',
            'statement_from_report': 'Test lot 의 NG rate 가 Normal lot 보다 높음.',
            'normalized_interpretation': 'Over glue Test 64.3% (63/98) vs 동일 이벤트 Normal 42.4% (337/794) = (64.3/42.4 - 1)*100 = +51.7% 악화. 결함 구성은 Hearing Noise 위주 (Test NG 의 71.4%, Normal NG 의 76.0%). Over glue 재사용 금지 권고.',
        }},
        'hints': {
            'hint_1': {'check_item': 'Over glue 보강품은 폐기, 재사용 금지.',
                       'reason': 'Function NG 64.3% (63/98) 는 동일 이벤트 Normal 42.4% (337/794) 대비 51.7% 악화; 동일 defect 구성.'},
            'hint_2': {'check_item': 'Hearing Noise 가 NG 의 다수일 때 dome/glue 조건을 사전 점검.',
                       'reason': 'Test/Normal 모두 Hearing Noise 비중 71~76% — over-glue 와 무관한 hearing 공정 이슈 시사.'},
        },
        'log': {
            'assumptions': ['같은 날 Normal row 가 over-glue Test row 의 baseline.'],
            'warnings': ['두 arm 모두 Hearing Noise 절대치가 높음 (>30%).'],
            'decision_rationale': 'normal_comparison: 동일 이벤트 Normal 존재. (64.3/42.4 - 1)*100 = +51.7% 악화, defect 구성 유사 => over glue 는 hearing 불량을 악화. 보강품 폐기.',
        },
    }
    tr_vi = {
        'document': {
            'title': 'Báo cáo BRS-161014 thử mẫu Over Glue gia cố',
            'purpose': 'Quyết định xem mẫu over-glue gia cố có dùng được không hay phải loại.',
            'content': [
                "Tách mẫu over-glue làm final rồi so sánh NG rate Function với Normal cùng đợt.",
                "Test: 98 pcs final 9/15/2023. Normal: 794 pcs cùng ngày.",
            ],
        },
        'conclusions': {'concl_1': {
            'topic': 'Over glue vs Normal — NG Function',
            'statement_from_report': 'NG rate của lot test cao hơn lot normal.',
            'normalized_interpretation': 'Over glue Test 64.3% (63/98) so với Normal cùng đợt 42.4% (337/794) = (64.3/42.4 - 1)*100 = +51.7% xấu hơn Normal. Defect chủ yếu Hearing Noise (Test 71.4% NG, Normal 76.0% NG); Hearing Touch thứ yếu. Không nên tái sử dụng over-glue gia cố.',
        }},
        'hints': {
            'hint_1': {'check_item': 'Loại bỏ mẫu over-glue gia cố, không tái sử dụng.',
                       'reason': 'NG Function 64.3% (63/98) xấu hơn Normal cùng đợt 42.4% (337/794) 51.7%; defect cùng dạng (Hearing Noise + Touch).'},
            'hint_2': {'check_item': 'Khi Hearing Noise chiếm phần lớn NG, kiểm tra điều kiện dome/glue trước khi thử.',
                       'reason': 'Cả Test và Normal đều có Hearing Noise chiếm 71.4%/76.0% tổng NG — gợi ý vấn đề Hearing chung của quy trình.'},
        },
        'log': {
            'assumptions': ['Row Normal cùng ngày là baseline cho row Test over-glue.'],
            'warnings': ['Cả hai arm đều có Hearing Noise tuyệt đối cao (>30%).'],
            'decision_rationale': 'Phân loại normal_comparison: có Normal cùng đợt. (64.3/42.4 - 1)*100 = +51.7% xấu hơn, defect cùng dạng => over-glue làm xấu thêm lỗi Hearing; loại bỏ over-glue gia cố.',
        },
    }
    return name, result, tr_ko, tr_en, tr_vi


# ---------------------------------------------------------------------
# 6. BRS-161016 — Vision Frame damage (quality_log style, all-zero NG)
# ---------------------------------------------------------------------
def build_brs_161016_vision():
    name = '29. BRS-161016 Report check problem damage Frame  13.4.2024'
    sheet = '5.3'

    doc = _doc(
        name=name,
        title='BRS-161016 Report Check Problem Damage Frame (Vision)',
        model='BRS-161016',
        report_date='2024-04-13',
        dept='ME', marker='Thuy', line='Main 1 / Main 2',
        report_type='ng_without_baseline',
        primary_canonical='Damage Frame (Vision Sub 4)',
        primary_aliases=['NG damage Frame'],
        related=[],
        parts=['Frame'],
        processes=['Vision Sub 4', 'Semi Main 1', 'Final Main 2'],
        purpose='Find the reason for NG damage Frame at Sub 4 Vision by 2-hour sampled vision check on Main 1 (Semi) and Main 2 (Final).',
        content=[
            "Sample every 2 hours / time on Main 1 Semi and Main 2 Final, count damage-Frame NG.",
            "4/13: 1500 + 1500 (5 cycles × 300 pcs) Semi+Final, NG = 0; 4/14: 800 + 800 (4 cycles × 200 pcs), NG = 0.",
        ],
    )

    test_conditions = [
        {'condition_id': 'cond_1', 'condition_group': '2-hourly vision sampling',
         'line': 'Main 1 / Main 2', 'process': 'Vision Sub 4', 'changed_factor': 'Time-of-day sampling',
         'before_value': None, 'after_value': None,
         'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None,
         'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': name, 'sheet_name': sheet,
         'source_cells': ['II. Content']},
    ]

    results = [
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Vision damage',
         'condition_group': 'Semi Main 1 (4/13 total)', 'date': '2024-04-13', 'line': 'Main 1',
         'input_count': 1500, 'ok_count': 1500, 'ng_count': 0,
         'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'NG rate damage Frame', 'metric_value': 0.0, 'unit': '%',
         'judgement': 'PASS', 'ng_breakdown': {'NG damage Frame': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['4/13 Total Semi row']},
        {'result_id': 'res_2', 'condition_id': 'cond_1', 'measurement_type': 'Vision damage',
         'condition_group': 'Final Main 2 (4/13 total)', 'date': '2024-04-13', 'line': 'Main 2',
         'input_count': 1500, 'ok_count': 1500, 'ng_count': 0,
         'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'NG rate damage Frame', 'metric_value': 0.0, 'unit': '%',
         'judgement': 'PASS', 'ng_breakdown': {'NG damage Frame': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['4/13 Total Final row']},
        {'result_id': 'res_3', 'condition_id': 'cond_1', 'measurement_type': 'Vision damage',
         'condition_group': 'Semi Main 1 (4/14 total)', 'date': '2024-04-14', 'line': 'Main 1',
         'input_count': 800, 'ok_count': 800, 'ng_count': 0,
         'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'NG rate damage Frame', 'metric_value': 0.0, 'unit': '%',
         'judgement': 'PASS', 'ng_breakdown': {'NG damage Frame': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['4/14 Total Semi row']},
        {'result_id': 'res_4', 'condition_id': 'cond_1', 'measurement_type': 'Vision damage',
         'condition_group': 'Final Main 2 (4/14 total)', 'date': '2024-04-14', 'line': 'Main 2',
         'input_count': 800, 'ok_count': 800, 'ng_count': 0,
         'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'NG rate damage Frame', 'metric_value': 0.0, 'unit': '%',
         'judgement': 'PASS', 'ng_breakdown': {'NG damage Frame': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['4/14 Total Final row']},
    ]

    conclusions = [
        {'conclusion_id': 'concl_1',
         'topic': 'Vision sampling — no damage Frame observed',
         'statement_from_report': "Checking sample 2 day but don't happen NG Frame damage.",
         'normalized_interpretation': 'Across 2 days × 4600 pcs sampled (Semi 2300 + Final 2300), 0 NG damage-Frame recorded at Sub 4 vision. The current sampling did not reproduce the prior damage-Frame issue.',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['III note line']},
    ]
    hints = [
        {'hint_id': 'hint_1',
         'check_item': 'Vision sampling alone could not reproduce damage Frame — investigate upstream Sub 4 forming/handling rather than vision detection.',
         'reason': '2-day sampling 4600 pcs returned 0 NG; problem is likely intermittent or upstream.',
         'evidence_strength': 'medium', 'related_process': 'Sub 4 / upstream forming',
         'related_part': 'Frame', 'source_file': name, 'sheet_name': sheet,
         'source_cells': ['III rows + note']},
    ]
    troubleshooting_index = {
        'defect_name': 'Damage Frame (Vision Sub 4)',
        'when_user_asks': ['damage Frame', 'Sub 4 vision', 'intermittent frame damage'],
        'suggested_checks': hints,
        'limitations': ['No baseline (prior NG rate) cited; only an absence of NG during 2-day sampling.', 'IV. Decision section is empty.'],
    }
    log = {
        'confidence': 0.7,
        'assumptions': ['Sampling cycles 1st–5th (4/13) and 1st–4th (4/14) sum to the reported totals.'],
        'warnings': ['Sampling resulted in 0 NG; cannot confirm root cause from this run.', 'IV. Decision is empty.'],
        'decision_rationale': 'Classified ng_without_baseline (closest to quality_log): metric=damage-Frame count, all zero, no comparison baseline. Per AI_EXCEL_PROC.md the absence of baseline is recorded in warnings and the conclusion does not claim improvement.',
    }

    result = {'schema_version': '0.1', 'document': doc, 'test_conditions': test_conditions,
              'results': results, 'conclusions': conclusions,
              'troubleshooting_index': troubleshooting_index, 'ai_extraction_log': log}

    tr_en = _trio_from_en({'title': doc['title'], 'purpose': doc['purpose'], 'content': doc['content']},
                          conclusions, hints, log)
    tr_ko = {
        'document': {
            'title': 'BRS-161016 Damage Frame 문제 점검 (Vision)',
            'purpose': 'Sub 4 Vision 의 damage Frame NG 원인 파악 — Main 1(Semi)/Main 2(Final) 에서 2시간 단위 표본 vision 점검.',
            'content': [
                "2시간/회 단위로 Main 1 Semi 및 Main 2 Final 표본 점검, damage Frame NG count.",
                "4/13: 1500+1500 (5 cycles × 300 pcs) Semi+Final NG=0; 4/14: 800+800 (4 cycles × 200 pcs) NG=0.",
            ],
        },
        'conclusions': {'concl_1': {
            'topic': 'Vision 표본 — damage Frame 미발생',
            'statement_from_report': '2일 표본 점검에서 NG Frame damage 미발생.',
            'normalized_interpretation': '2일 × 4600 pcs (Semi 2300 + Final 2300) 에서 Sub 4 vision 의 damage Frame NG 0. 현재 표본으로는 이전 damage Frame 문제 재현 안 됨.',
        }},
        'hints': {
            'hint_1': {'check_item': 'Vision 표본만으로는 재현 불가 — Sub 4 상류(forming/handling) 점검.',
                       'reason': '2일 4600 pcs 표본 NG=0; 간헐적 또는 상류 문제 가능성.'},
        },
        'log': {
            'assumptions': ['4/13 1st~5th, 4/14 1st~4th cycle 합계가 리포트 합계와 일치.'],
            'warnings': ['표본 NG=0 — 본 시험으로 근본 원인 확인 불가.', 'IV. Decision 비어 있음.'],
            'decision_rationale': 'ng_without_baseline (quality_log 에 가까움) 분류: damage Frame count 가 전부 0, 비교 baseline 없음. baseline 부재는 warnings 에 기록, 결론에서 개선 주장 금지.',
        },
    }
    tr_vi = {
        'document': {
            'title': 'Báo cáo BRS-161016 kiểm tra vấn đề Damage Frame (Vision)',
            'purpose': 'Tìm nguyên nhân NG damage Frame ở Sub 4 Vision bằng cách kiểm vision 2 giờ/lần ở Main 1 Semi và Main 2 Final.',
            'content': [
                "Lấy mẫu mỗi 2 giờ/lần ở Main 1 Semi và Main 2 Final, đếm NG damage Frame.",
                "4/13: 1500+1500 (5 lượt × 300 pcs) Semi+Final NG=0; 4/14: 800+800 (4 lượt × 200 pcs) NG=0.",
            ],
        },
        'conclusions': {'concl_1': {
            'topic': 'Lấy mẫu vision — không phát hiện damage Frame',
            'statement_from_report': 'Kiểm mẫu 2 ngày nhưng không xảy ra NG Frame damage.',
            'normalized_interpretation': 'Trên 2 ngày × 4600 pcs (Semi 2300 + Final 2300), NG damage Frame ở Sub 4 vision = 0. Lượt lấy mẫu này không tái hiện được lỗi damage Frame trước đó.',
        }},
        'hints': {
            'hint_1': {'check_item': 'Lấy mẫu Vision không tái hiện được — điều tra thượng nguồn Sub 4 (forming/handling).',
                       'reason': 'Lấy mẫu 2 ngày 4600 pcs NG=0; lỗi có thể gián đoạn hoặc nằm ở thượng nguồn.'},
        },
        'log': {
            'assumptions': ['Số lượt 4/13 1st~5th, 4/14 1st~4th cộng đúng tổng báo cáo.'],
            'warnings': ['Lấy mẫu NG=0 — không xác nhận được nguyên nhân gốc qua lần chạy này.', 'IV. Decision trống.'],
            'decision_rationale': 'Phân loại ng_without_baseline (gần quality_log): chỉ số damage-Frame count đều 0, không có baseline so sánh. Không tuyên bố cải thiện; ghi baseline thiếu vào warnings.',
        },
    }
    return name, result, tr_ko, tr_en, tr_vi


# ---------------------------------------------------------------------
# 8. MSU-L20S15-07 LIST TEST IMPROVE NG FUNCTION SPK (multi-arm DOE-like log)
# ---------------------------------------------------------------------
def build_msu_list_test():
    name = '29. MSU-L20S15-07 LIST TEST IMPROVE NG FUNCTION SPK 20225.03.05'
    sheet = 'Sheet1'

    doc = _doc(
        name=name,
        title='MSU-L20S15-07 List Test Improve NG Function SPK',
        model='MSU-L20S15-07',
        report_date='2025-03-05',
        dept='ME', marker='Kim/Byun/Yang/Tu/IQC', line='C2 (AWF #1/#2/#3, UC #1/#2)',
        report_type='mixed',
        primary_canonical='NG Function SPK (Hearing dominated)',
        primary_aliases=['NG Hearing Noise', 'NG Hearing Touch', 'Sigma SPL/THD'],
        related=['Sigma SPL', 'Sigma THD', 'Sigma SPL+THD', 'Hearing Noise', 'Hearing Touch'],
        parts=['Frame', 'SP', 'VP', 'CD', 'Coil', 'CMG', 'SMG'],
        processes=['Array Frame+SP', 'Led UC (VP+CD)', 'Bonding line sub 1',
                   'Dry box / Dry VP-Frame', 'Press UC', 'Led UV clamp', 'CMG/SMG combinations'],
        purpose='Run a long list of intervention tests over 3/4–3/19/2025 to find which change reduces NG Function SPK (Hearing-dominated).',
        content=[
            "Each entry compares a Test condition vs an immediate Normal row on the same day for Sigma, Hearing (+1V) and Hearing (+0V) NG.",
            "Categories tested: array method, UC time (3s/5s), bond line offset, bond amount sweep, AEM 70/75/80/85A, speed bond VP+CD (40/38/36/34), Coil bond amount, Dry VP/Frame and Dry Box, Press UC #1/#2, Led UV clamp 1/2/3/4, Sub 1 OK vs clamp tilted, Frame ring plating, CMG/SMG (Ruijin/Baotou).",
        ],
    )

    test_conditions = [
        {'condition_id': 'cond_grp', 'condition_group': 'Multi-intervention list',
         'line': 'C2', 'process': 'Various (Array/UC/Bond/Dry/Press/Led UV/CMG)',
         'changed_factor': 'Many (see content)',
         'before_value': 'Normal of same day', 'after_value': 'Various Test conditions',
         'unit': None, 'machine': 'AWF #1/2/3, UC #1/#2, Forming machines',
         'jig': None, 'material_lot': None, 'supplier': 'CMG/SMG (Ruijin/Baotou)',
         'dry_time_sec': None, 'temperature': '65°C dry', 'pressure': None,
         'bond_amount': '1.8~2.0 / 2.6~2.8 / 3.3~3.7 / 3.6~3.7 / 3.7~4.0 / 4.1~4.18 mg',
         'uv_energy': None,
         'source_file': name, 'sheet_name': sheet,
         'source_cells': ['Sheet1 rows 1..18']},
    ]

    # Build a flat results list of Hearing(+1V) NG rate per (date, label).
    # All rows are read off the LIST TEST table. Hearing (+1V) is the dominant
    # failure column. Below: (label, date, inp, ok, ng_total, rate_pct, hint_kind)
    rows = [
        ('Test array Frame+SP by hand', '2025-03-04', 400, 274, 123, 30.8, 'test'),
        ('Normal (Array Frame/Sus)', '2025-03-04', 799, 553, 244, 30.5, 'normal'),
        ('Test Led UC (VP+CD) 5s', '2025-03-04', 500, 379, 119, 23.8, 'test'),
        ('Normal Led UC (VP+CD) 3s', '2025-03-04', 1120, 898, 219, 19.6, 'normal'),
        ('Test bonding line sub 1 offset 0.03 + bond 3.3~3.7mg (3/5)', '2025-03-05', 280, 277, 3, 1.1, 'test'),
        ('Test bonding line sub 1 offset 0.03 + bond 3.3~3.7mg (3/6)', '2025-03-06', 249, 245, 4, 1.6, 'test'),
        ('Normal (3/5)', '2025-03-05', 1119, 930, 187, 16.7, 'normal'),
        ('Dry VP/Frame 5min+65°C', '2025-03-05', 800, 695, 105, 13.1, 'test'),
        ('Dry Box final 30min+65°C', '2025-03-05', 800, 685, 115, 14.4, 'test'),
        ('Normal (Dry test 3/5)', '2025-03-05', 800, 726, 73, 9.1, 'normal'),
        ('Test CMG-R1', '2025-03-05', 303, 258, 45, 14.9, 'test'),
        ('Test new Frame+SP load tray', '2025-03-05', 346, 312, 33, 9.5, 'test'),
        ('Normal Frame+SP load tray', '2025-03-05', 350, 284, 66, 18.9, 'normal'),
        ('Test VP AEM 75A', '2025-03-08', 145, 136, 9, 6.2, 'test'),
        ('Test VP AEM 80A', '2025-03-08', 149, 135, 14, 9.4, 'test'),
        ('Test VP AEM 85A', '2025-03-08', 150, 135, 15, 10.0, 'test'),
        ('Normal VP AEM 70A', '2025-03-08', 400, 320, 79, 19.8, 'normal'),
        ('Test bond VP+CD max 3.6~3.7mg', '2025-03-08', 298, 248, 49, 16.4, 'test'),
        ('Test bond VP+CD min 3.3~3.4mg', '2025-03-08', 300, 218, 82, 27.3, 'test'),
        ('Test bond VP+CD 3.3~3.7mg', '2025-03-08', 300, 281, 19, 6.3, 'test'),
        ('Test speed bond VP+CD 34, bond 4.1~4.18mg', '2025-03-09', 300, 255, 45, 15.0, 'test'),
        ('Test speed bond VP+CD 36, bond 3.7~4.0mg', '2025-03-09', 300, 273, 24, 8.0, 'test'),
        ('Test speed bond VP+CD 38, bond 3.72~3.74mg', '2025-03-09', 300, 269, 30, 10.0, 'test'),
        ('Normal speed bond VP+CD 40, bond 3.36~3.6mg', '2025-03-09', 300, 287, 13, 4.3, 'normal'),
        ('Test sample VP+CD OK (3/11)', '2025-03-11', 200, 177, 23, 11.5, 'test'),
        ('Test sample VP+CD offset (3/11)', '2025-03-11', 99, 78, 21, 21.2, 'test'),
        ('Test bond Coil+CD 2.6~2.8mg', '2025-03-11', 500, 413, 86, 17.2, 'test'),
        ('Normal bond Coil+CD 1.8~2.0mg', '2025-03-11', 800, 704, 91, 11.4, 'normal'),
        ('Test semi VP+CD of C2 Coil 0.096 (3/12)', '2025-03-12', 475, 454, 20, 4.2, 'test'),
        ('Normal semi VP+CD of E2 Coil 0.096', '2025-03-12', 348, 335, 13, 3.7, 'normal'),
        ('Test Frame change Damping (3/12)', '2025-03-12', 341, 319, 21, 6.2, 'test'),
        ('Test Coil 0.096 AWF #3', '2025-03-12', 297, 262, 35, 11.8, 'test'),
        ('Normal (Coil 0.096 3/12)', '2025-03-12', 280, 259, 21, 7.5, 'normal'),
        ('Test press UC1 (Sub 1 VP+CD)', '2025-03-13', 199, 189, 8, 4.0, 'test'),
        ('Test press UC2 (Sub 1 VP+CD)', '2025-03-13', 199, 170, 29, 14.6, 'test'),
        ('Test Led UV clamp 1', '2025-03-13', 148, 145, 3, 2.0, 'test'),
        ('Test Led UV clamp 2', '2025-03-13', 141, 134, 7, 5.0, 'test'),
        ('Test Led UV clamp 3', '2025-03-13', 138, 136, 2, 1.4, 'test'),
        ('Test Led UV clamp 4', '2025-03-13', 143, 141, 2, 1.4, 'test'),
        ('Test Sub 1 (VP+CD) clamp tilted', '2025-03-14', 200, 186, 14, 7.0, 'test'),
        ('Test Sub 1 (VP+CD) OK', '2025-03-14', 200, 189, 9, 4.5, 'test'),
        ('Normal Sub 1 (3/15)', '2025-03-15', 400, 384, 15, 3.8, 'normal'),
        ('Test Frame change Damping (3/18)', '2025-03-18', 600, 584, 16, 2.7, 'test'),
        ('Normal lot (3/18)', '2025-03-18', 600, 582, 17, 2.8, 'normal'),
        ('Test Frame ring different plating color', '2025-03-18', 1000, 968, 31, 3.1, 'test'),
        ('Test CMG-Ruijin / SMG-Ruijin', '2025-03-18', 104, 104, 0, 0.0, 'test'),
        ('Test CMG-Baotou / SMG-Ruijin', '2025-03-18', 106, 102, 4, 3.8, 'test'),
        ('Normal CMG/SMG (3/18)', '2025-03-18', 399, 387, 12, 3.0, 'normal'),
        ('Test CMG-Ruijin / SMG-Baotou', '2025-03-19', 100, 97, 3, 3.0, 'test'),
        ('Test CMG-Baotou / SMG-Baotou', '2025-03-19', 100, 97, 3, 3.0, 'test'),
        ('Normal CMG/SMG (3/19)', '2025-03-19', 400, 387, 13, 3.2, 'normal'),
    ]
    results = []
    for idx, (label, date, inp, ok, ng, rate, kind) in enumerate(rows, start=1):
        results.append({
            'result_id': f'res_{idx}', 'condition_id': 'cond_grp',
            'measurement_type': 'Function (Hearing +1V dominant)',
            'condition_group': label, 'date': date, 'line': 'C2',
            'input_count': inp, 'ok_count': ok, 'ng_count': ng,
            'ng_rate_decimal': rate/100.0, 'ng_rate_percent': rate,
            'metric_name': 'Hearing (+1V) Total NG Rate',
            'metric_value': rate, 'unit': '%',
            'judgement': None, 'ng_breakdown': {},
            'source_file': name, 'sheet_name': sheet,
            'source_cells': [f'Row {idx}: {label}']})

    # Headline relative-change conclusions vs same-day Normal where pairs exist.
    conclusions = [
        {'conclusion_id': 'concl_1',
         'topic': 'Bonding line sub 1 offset 0.03 + bond 3.3~3.7mg vs Normal 3/5',
         'statement_from_report': 'NG rate Hearing 1.1% / 1.6% vs Normal 16.7% (highlighted yellow in workbook).',
         'normalized_interpretation': 'Test 3/5 1.1% (3/280) vs same-day Normal 16.7% (187/1119) = (1.1/16.7 - 1)*100 = -93.4% improved. 3/6 1.6% (4/249) vs same Normal 16.7% = -90.4% improved. Strongest improvement in the report.',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Row 3 highlighted']},
        {'conclusion_id': 'concl_2',
         'topic': 'Led UV clamp 3 and clamp 4 vs Normal',
         'statement_from_report': 'Hearing 1.4% (clamp3) and 1.4% (clamp4).',
         'normalized_interpretation': 'No same-day Normal row paired with Led UV clamp on 3/13. Nearest Normal lots earlier in the week ranged 7.5%–18.9%; relative comparison cannot be tight. Absolute Hearing NG 1.4% × 2 conditions is the second-best block after bonding-line offset.',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Rows 14 clamp 3/4']},
        {'conclusion_id': 'concl_3',
         'topic': 'AEM current sweep vs Normal 70A',
         'statement_from_report': 'AEM 75A 6.2%, 80A 9.4%, 85A 10.0% vs Normal 70A 19.8%.',
         'normalized_interpretation': 'All three test currents improve vs Normal 70A: (6.2/19.8-1)*100 = -68.7%, (9.4/19.8-1)*100 = -52.5%, (10.0/19.8-1)*100 = -49.5%. AEM 75A is best.',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Row 7 AEM rows']},
        {'conclusion_id': 'concl_4',
         'topic': 'Speed bond VP+CD vs Normal 40 (3/9)',
         'statement_from_report': 'Speed 34/36/38 vs Normal 40.',
         'normalized_interpretation': 'Test 34 = 15.0% (3.49x = +249% worse than Normal 4.3%); 36 = 8.0% (+86%); 38 = 10.0% (+133%). All slower speeds worsen Hearing; keep speed 40.',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Row 9 speed bond rows']},
        {'conclusion_id': 'concl_5',
         'topic': 'CMG/SMG (Ruijin/Baotou) combinations vs Normal',
         'statement_from_report': 'CMG-Ruijin/SMG-Ruijin = 0%; Baotou/Ruijin = 3.8%; Ruijin/Baotou = 3.0%; Baotou/Baotou = 3.0%; Normal = 3.0~3.2%.',
         'normalized_interpretation': 'CMG-Ruijin/SMG-Ruijin 0% (0/104) is the only combo improved vs Normal; others are within ±20% of Normal. Sample size 100 pcs per combo is small.',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Rows 17-18 CMG combos']},
    ]
    hints = [
        {'hint_id': 'hint_1',
         'check_item': 'Adopt bonding line sub 1 offset 0.03 + bond 3.3~3.7mg.',
         'reason': 'Hearing NG 1.1%/1.6% vs Normal 16.7% = -93.4%/-90.4% improved; 2-day reproduction; highlighted in workbook.',
         'evidence_strength': 'high', 'related_process': 'Bonding line sub 1',
         'related_part': 'VP+CD', 'source_file': name, 'sheet_name': sheet,
         'source_cells': ['Row 3 highlighted yellow']},
        {'hint_id': 'hint_2',
         'check_item': 'Switch VP AEM current 70A → 75A for VP coil bonding.',
         'reason': 'AEM 75A 6.2% is -68.7% vs Normal AEM 70A 19.8%; AEM 80/85A also improve but less.',
         'evidence_strength': 'medium', 'related_process': 'VP AEM bond',
         'related_part': 'VP coil', 'source_file': name, 'sheet_name': sheet,
         'source_cells': ['Row 7 AEM rows']},
        {'hint_id': 'hint_3',
         'check_item': 'Do NOT change speed bond VP+CD from 40 (Normal); all slower speeds worsen Hearing.',
         'reason': '34/36/38 rpm vs Normal 40 → Hearing rates 15.0% / 8.0% / 10.0% vs 4.3% (≥+86% worse).',
         'evidence_strength': 'high', 'related_process': 'Speed bond VP+CD',
         'related_part': 'VP+CD', 'source_file': name, 'sheet_name': sheet,
         'source_cells': ['Row 9 speed bond rows']},
        {'hint_id': 'hint_4',
         'check_item': 'Pilot CMG-Ruijin/SMG-Ruijin (0/104) on a larger sample before adopting.',
         'reason': '0% vs Normal 3.0% is good but n=104 is small; other CMG/SMG combos are within ±20% of Normal.',
         'evidence_strength': 'low', 'related_process': 'Magnet (CMG/SMG)',
         'related_part': 'Magnet', 'source_file': name, 'sheet_name': sheet,
         'source_cells': ['Rows 17-18 CMG combos']},
    ]
    troubleshooting_index = {
        'defect_name': 'NG Function SPK (Hearing dominated)',
        'when_user_asks': ['NG Hearing Noise', 'NG Hearing Touch', 'bond amount', 'AEM', 'CMG/SMG', 'Led UV clamp'],
        'suggested_checks': hints,
        'limitations': ['Long list; many test rows have only #DIV/0! columns in Sigma — Hearing(+1V) is the only consistent metric here.',
                        'Some test rows have no same-day Normal pair.',
                        'IV. Decision section text not present in workbook.'],
    }
    log = {
        'confidence': 0.7,
        'assumptions': ['Reported rates are accepted as-is; relative changes use same-day or nearest-day Normal.',
                        'Hearing (+1V) total NG is the primary metric column (Sigma SPL/THD has many #DIV/0! and small counts).'],
        'warnings': ['Some Sigma cells are #DIV/0! due to 0 OK.', 'Some Test rows have no same-day Normal.', 'Workbook date string "20225.03.05" looks like a typo of 2025.03.05.'],
        'decision_rationale': 'Classified mixed: many small same-day Normal-vs-Test comparisons covering bond, AEM, speed, CMG/SMG, Led UV, Press UC. Strongest improvement is bonding line sub 1 offset (-93%). Slower speed bond worsens; AEM 75A improves; CMG-Ruijin/Ruijin needs larger sample.',
    }

    result = {'schema_version': '0.1', 'document': doc, 'test_conditions': test_conditions,
              'results': results, 'conclusions': conclusions,
              'troubleshooting_index': troubleshooting_index, 'ai_extraction_log': log}

    tr_en = _trio_from_en({'title': doc['title'], 'purpose': doc['purpose'], 'content': doc['content']},
                          conclusions, hints, log)
    tr_ko = {
        'document': {
            'title': 'MSU-L20S15-07 NG Function SPK 개선 리스트 테스트',
            'purpose': '3/4~3/19 기간 동안 다수 개입 시험으로 NG Function SPK (Hearing 위주) 를 줄이는 조건을 탐색.',
            'content': [
                "각 시험 행은 같은 날 Normal row 와 Sigma/Hearing(+1V)/Hearing(+0V) NG 비교.",
                "검토 카테고리: array 방식, UC 시간(3s/5s), bond line offset, bond 양(1.8~4.18mg), AEM 70/75/80/85A, speed bond VP+CD(40/38/36/34), Coil bond, Dry VP/Frame & Dry Box, Press UC1/UC2, Led UV clamp 1/2/3/4, Sub 1 OK vs clamp tilted, Frame ring plating, CMG/SMG (Ruijin/Baotou).",
            ],
        },
        'conclusions': {
            'concl_1': {'topic': 'Bonding line sub 1 offset 0.03 + bond 3.3~3.7mg vs Normal 3/5',
                        'statement_from_report': 'Hearing NG 1.1% / 1.6% vs Normal 16.7% (워크북에서 노란 강조).',
                        'normalized_interpretation': '3/5 Test 1.1% (3/280) vs 같은 날 Normal 16.7% (187/1119) = (1.1/16.7 - 1)*100 = -93.4% 개선. 3/6 1.6% (4/249) vs 같은 Normal = -90.4% 개선. 리스트 중 최강 개선.'},
            'concl_2': {'topic': 'Led UV clamp 3, 4 vs Normal',
                        'statement_from_report': 'Hearing 1.4% (clamp 3), 1.4% (clamp 4).',
                        'normalized_interpretation': '3/13 Led UV clamp 시험에는 같은 날 Normal row 가 페어링 되지 않음. 주변 일자 Normal 은 7.5~18.9% 범위라 엄밀한 상대 비교 어려움. 절대값 1.4% × 2 조건은 bonding-line offset 다음으로 우수한 블록.'},
            'concl_3': {'topic': 'AEM 전류 sweep vs Normal 70A',
                        'statement_from_report': 'AEM 75A 6.2%, 80A 9.4%, 85A 10.0% vs Normal 70A 19.8%.',
                        'normalized_interpretation': '세 전류 모두 Normal 70A 대비 개선: (6.2/19.8-1)*100=-68.7%, (9.4/19.8-1)*100=-52.5%, (10.0/19.8-1)*100=-49.5%. AEM 75A 최선.'},
            'concl_4': {'topic': 'Speed bond VP+CD vs Normal 40 (3/9)',
                        'statement_from_report': 'Speed 34/36/38 vs Normal 40.',
                        'normalized_interpretation': 'Test 34 = 15.0% (Normal 4.3% 대비 3.49배, +249% 악화); 36 = 8.0% (+86% 악화); 38 = 10.0% (+133% 악화). 속도 낮추면 Hearing 악화 — Normal 40 유지.'},
            'concl_5': {'topic': 'CMG/SMG (Ruijin/Baotou) 조합 vs Normal',
                        'statement_from_report': 'CMG-Ruijin/SMG-Ruijin = 0%; Baotou/Ruijin = 3.8%; Ruijin/Baotou = 3.0%; Baotou/Baotou = 3.0%; Normal = 3.0~3.2%.',
                        'normalized_interpretation': 'CMG-Ruijin/SMG-Ruijin 0% (0/104) 만 Normal 대비 개선, 나머지는 Normal ±20% 이내. 표본 100 pcs 라 신뢰성 제한.'},
        },
        'hints': {
            'hint_1': {'check_item': 'Bonding line sub 1 offset 0.03 + bond 3.3~3.7mg 채택.',
                       'reason': 'Hearing 1.1%/1.6% vs Normal 16.7% = -93.4%/-90.4% 개선; 2일 재현; 워크북 노란색 강조.'},
            'hint_2': {'check_item': 'VP AEM 전류 70A → 75A 로 변경.',
                       'reason': 'AEM 75A 6.2% 가 Normal 70A 19.8% 대비 -68.7% 개선; 80/85A 도 개선이지만 폭이 작음.'},
            'hint_3': {'check_item': 'Speed bond VP+CD 는 40 유지 — 40 미만은 모두 Hearing 악화.',
                       'reason': '34/36/38 rpm 대 Normal 40 에서 Hearing rate 15.0%/8.0%/10.0% vs 4.3% (≥+86% 악화).'},
            'hint_4': {'check_item': 'CMG-Ruijin/SMG-Ruijin (0/104) 은 추가 표본 후 채택 결정.',
                       'reason': '0% vs Normal 3.0% 양호하나 n=104 작음; 타 조합은 Normal ±20% 이내.'},
        },
        'log': {
            'assumptions': ['리포트 NG rate 그대로 수용; 상대 변화는 같은 날 또는 인접 일자 Normal 사용.',
                            'Hearing (+1V) total NG 가 주 지표 (Sigma SPL/THD 는 #DIV/0! 많음).'],
            'warnings': ['일부 Sigma 셀이 #DIV/0! (분모 0).', '일부 Test row 는 같은 날 Normal 없음.', '날짜 문자열 "20225.03.05" 는 2025.03.05 오타.'],
            'decision_rationale': 'mixed 분류: 같은 날 Test-vs-Normal 다수 비교 (bond, AEM, speed, CMG/SMG, Led UV, Press UC). 최강 개선: bonding line sub 1 offset (-93%). Speed bond 느림화는 악화; AEM 75A 개선; CMG-Ruijin/Ruijin 은 추가 표본 필요.',
        },
    }
    tr_vi = {
        'document': {
            'title': 'Báo cáo danh sách thử nghiệm cải thiện NG Function SPK — MSU-L20S15-07',
            'purpose': 'Khảo sát các thay đổi trong 3/4–3/19/2025 để giảm NG Function SPK (chủ yếu Hearing).',
            'content': [
                "Mỗi dòng so sánh Test với Normal cùng ngày trên Sigma, Hearing (+1V), Hearing (+0V).",
                "Hạng mục: array, thời gian UC (3s/5s), offset bond line, lượng keo (1.8~4.18mg), AEM 70/75/80/85A, speed bond VP+CD (40/38/36/34), bond Coil, Dry VP/Frame & Dry Box, Press UC1/UC2, Led UV clamp 1/2/3/4, Sub 1 OK vs clamp lệch, Frame ring mạ khác, CMG/SMG (Ruijin/Baotou).",
            ],
        },
        'conclusions': {
            'concl_1': {'topic': 'Bonding line sub 1 offset 0.03 + bond 3.3~3.7mg vs Normal 3/5',
                        'statement_from_report': 'Hearing 1.1% / 1.6% vs Normal 16.7% (đánh dấu vàng trong workbook).',
                        'normalized_interpretation': 'Test 3/5 1.1% (3/280) vs Normal cùng ngày 16.7% (187/1119) = (1.1/16.7 - 1)*100 = -93.4% cải thiện. 3/6 1.6% (4/249) vs Normal đó = -90.4% cải thiện. Cải thiện mạnh nhất trong danh sách.'},
            'concl_2': {'topic': 'Led UV clamp 3, 4 vs Normal',
                        'statement_from_report': 'Hearing 1.4% (clamp 3), 1.4% (clamp 4).',
                        'normalized_interpretation': 'Không có Normal cùng ngày ghép với Led UV clamp 3/13. Normal các ngày khác 7.5~18.9% nên so sánh tương đối không chặt. Mức tuyệt đối 1.4% × 2 điều kiện là khối tốt thứ nhì sau bonding-line offset.'},
            'concl_3': {'topic': 'AEM sweep vs Normal 70A',
                        'statement_from_report': 'AEM 75A 6.2%, 80A 9.4%, 85A 10.0% so với Normal 70A 19.8%.',
                        'normalized_interpretation': 'Cả ba dòng test đều cải thiện so với Normal 70A: (6.2/19.8-1)*100=-68.7%, (9.4/19.8-1)*100=-52.5%, (10.0/19.8-1)*100=-49.5%. AEM 75A tốt nhất.'},
            'concl_4': {'topic': 'Speed bond VP+CD vs Normal 40 (3/9)',
                        'statement_from_report': 'Speed 34/36/38 so với Normal 40.',
                        'normalized_interpretation': 'Test 34 = 15.0% (xấu hơn Normal 4.3% gấp 3.49 lần, +249%); 36 = 8.0% (+86%); 38 = 10.0% (+133%). Giảm tốc độ làm xấu Hearing; giữ Normal 40.'},
            'concl_5': {'topic': 'Tổ hợp CMG/SMG (Ruijin/Baotou) vs Normal',
                        'statement_from_report': 'CMG-Ruijin/SMG-Ruijin = 0%; Baotou/Ruijin = 3.8%; Ruijin/Baotou = 3.0%; Baotou/Baotou = 3.0%; Normal = 3.0~3.2%.',
                        'normalized_interpretation': 'Chỉ CMG-Ruijin/SMG-Ruijin 0% (0/104) cải thiện so với Normal; còn lại nằm trong ±20% Normal. n=104 nhỏ.'},
        },
        'hints': {
            'hint_1': {'check_item': 'Áp dụng bonding line sub 1 offset 0.03 + bond 3.3~3.7mg.',
                       'reason': 'Hearing 1.1%/1.6% vs Normal 16.7% = -93.4%/-90.4% cải thiện; tái hiện 2 ngày; đánh dấu vàng.'},
            'hint_2': {'check_item': 'Đổi dòng VP AEM 70A → 75A khi bond coil VP.',
                       'reason': 'AEM 75A 6.2% cải thiện -68.7% so với Normal 70A 19.8%; 80/85A cải thiện nhỏ hơn.'},
            'hint_3': {'check_item': 'Không thay đổi speed bond VP+CD khỏi 40 (Normal); mọi tốc độ thấp hơn đều làm xấu Hearing.',
                       'reason': '34/36/38 rpm so với Normal 40 → Hearing 15.0%/8.0%/10.0% vs 4.3% (≥+86% xấu).'},
            'hint_4': {'check_item': 'Chạy pilot CMG-Ruijin/SMG-Ruijin (0/104) trên cỡ mẫu lớn hơn trước khi áp dụng.',
                       'reason': '0% so với Normal 3.0% rất tốt nhưng n=104 nhỏ; các tổ hợp khác nằm trong ±20% Normal.'},
        },
        'log': {
            'assumptions': ['Chấp nhận NG rate báo cáo; so sánh tương đối dùng Normal cùng ngày hoặc gần nhất.',
                            'Hearing (+1V) total NG là cột chỉ số chính (Sigma SPL/THD có nhiều #DIV/0!).'],
            'warnings': ['Một số cell Sigma là #DIV/0! do OK=0.', 'Một số row Test không có Normal cùng ngày.', 'Chuỗi ngày "20225.03.05" có vẻ gõ sai từ 2025.03.05.'],
            'decision_rationale': 'Phân loại mixed: nhiều so sánh Test-vs-Normal cùng ngày (bond, AEM, speed, CMG/SMG, Led UV, Press UC). Cải thiện mạnh nhất: bonding line sub 1 offset (-93%). Speed bond chậm hơn → xấu hơn; AEM 75A cải thiện; CMG-Ruijin/Ruijin cần mẫu lớn hơn.',
        },
    }
    return name, result, tr_ko, tr_en, tr_vi


# ---------------------------------------------------------------------
# 9. MSU-L20S15-07 New Dry UV 395nm
# ---------------------------------------------------------------------
def build_msu_uv_395nm():
    name = '29. MSU-L20S15-07 Report test New Dry UV Machine 395nm ( Normal 365nm) date 2.05.2025'
    sheet = '161016 + 201507'

    doc = _doc(
        name=name,
        title='MSU-L20S15-07 Report Test New Dry UV Machine 395nm vs Normal 365nm',
        model='MSU-L20S15-07 (161016 + 201507)',
        report_date='2025-05-02',
        dept='ME', marker='Le', line='C2-3B / C2-2A',
        report_type='reliability_spec',
        primary_canonical='Dry UV Machine Change (395nm vs 365nm)',
        primary_aliases=['New Dry UV 395nm', 'Normal Dry UV 365nm'],
        related=['VP/CD Tension'],
        parts=['VP Film', 'CD bond'],
        processes=['Dry UV', 'VP+CD Tension test', 'UC Press'],
        purpose='Decide whether the new 395nm Dry UV machine can replace the Normal 365nm Dry UV machine, based on LED UV Peak/Total and VP+CD Tension test.',
        content=[
            "BRS-161016 (C2-3B): Peak/Total measured with film and without film; Tension spec 0.5 kgf; 10 samples per arm.",
            "MSU-L20S15-07 201507 (C2-2A): same comparison, Tension spec 1.2 kgf; 10 samples per arm.",
            "Normal sequence: UC PRESS → 365nm Dry1 → 365nm Dry2. Test sequence: UC PRESS → 395nm → 395nm.",
        ],
    )

    test_conditions = [
        {'condition_id': 'cond_1', 'condition_group': 'Dry UV machine change',
         'line': 'C2-3B / C2-2A', 'process': 'Dry UV', 'changed_factor': 'Dry UV wavelength',
         'before_value': '365nm (Normal)', 'after_value': '395nm (Test)',
         'unit': 'nm', 'machine': 'Dry UV machine 395nm / 365nm', 'jig': None,
         'material_lot': '161016 VP Film / 201507 VP Film', 'supplier': None,
         'dry_time_sec': None, 'temperature': None, 'pressure': None,
         'bond_amount': None, 'uv_energy': '395nm Peak / Total',
         'source_file': name, 'sheet_name': sheet,
         'source_cells': ['I.Purpose, II.Content']},
    ]

    results = [
        # 161016 LED UV (with film) — Test 395nm Position 1/2/3 Peak 11/11/33, Total 50/45/84
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'LED UV Peak',
         'condition_group': '161016 Test 395nm with film P1',
         'date': '2025-05-02', 'line': 'C2-3B',
         'input_count': None, 'ok_count': None, 'ng_count': None,
         'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Peak (P1)', 'metric_value': 11.0, 'unit': 'mW/cm2',
         'judgement': None, 'ng_breakdown': {},
         'source_file': name, 'sheet_name': '161016',
         'source_cells': ['161016 Test 395nm Peak P1']},
        {'result_id': 'res_2', 'condition_id': 'cond_1', 'measurement_type': 'LED UV Total',
         'condition_group': '161016 Test 395nm with film P1',
         'date': '2025-05-02', 'line': 'C2-3B',
         'input_count': None, 'ok_count': None, 'ng_count': None,
         'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Total (P1)', 'metric_value': 50.0, 'unit': 'mW/cm2',
         'judgement': None, 'ng_breakdown': {},
         'source_file': name, 'sheet_name': '161016',
         'source_cells': ['161016 Test 395nm Total P1']},
        # 161016 Tension Test 395nm 10 samples min/max/avg
        {'result_id': 'res_3', 'condition_id': 'cond_1', 'measurement_type': 'Tension VP+CD',
         'condition_group': '161016 Test Dry UV 395nm', 'date': '2025-05-02', 'line': 'C2-3B',
         'input_count': 10, 'ok_count': None, 'ng_count': None,
         'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Avg Tension', 'metric_value': 1.310, 'unit': 'kgf',
         'judgement': 'PASS', 'ng_breakdown': {},
         'source_file': name, 'sheet_name': '161016',
         'source_cells': ['Tension Test 395nm avg']},
        {'result_id': 'res_4', 'condition_id': None, 'measurement_type': 'Tension VP+CD',
         'condition_group': '161016 Normal Dry UV 365nm', 'date': '2025-05-02', 'line': 'C2-3B',
         'input_count': 10, 'ok_count': None, 'ng_count': None,
         'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Avg Tension', 'metric_value': 1.374, 'unit': 'kgf',
         'judgement': 'PASS', 'ng_breakdown': {},
         'source_file': name, 'sheet_name': '161016',
         'source_cells': ['Tension Normal 365nm avg']},
        # 201507 Tension Test 395nm vs Normal 365nm 10 samples each
        {'result_id': 'res_5', 'condition_id': 'cond_1', 'measurement_type': 'Tension VP+CD',
         'condition_group': '201507 Test Dry UV 395nm', 'date': '2025-05-02', 'line': 'C2-2A',
         'input_count': 10, 'ok_count': None, 'ng_count': None,
         'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Avg Tension', 'metric_value': 2.875, 'unit': 'kgf',
         'judgement': 'PASS', 'ng_breakdown': {},
         'source_file': name, 'sheet_name': '201507',
         'source_cells': ['Tension Test 395nm avg']},
        {'result_id': 'res_6', 'condition_id': None, 'measurement_type': 'Tension VP+CD',
         'condition_group': '201507 Normal Dry UV 365nm', 'date': '2025-05-02', 'line': 'C2-2A',
         'input_count': 10, 'ok_count': None, 'ng_count': None,
         'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Avg Tension', 'metric_value': 2.503, 'unit': 'kgf',
         'judgement': 'PASS', 'ng_breakdown': {},
         'source_file': name, 'sheet_name': '201507',
         'source_cells': ['Tension Normal 365nm avg']},
    ]

    conclusions = [
        {'conclusion_id': 'concl_1',
         'topic': 'Tension test 161016 (spec 0.5 kgf)',
         'statement_from_report': 'Tension same with normal.',
         'normalized_interpretation': '395nm avg Tension 1.310 kgf vs Normal 365nm 1.374 kgf = (1.310/1.374 - 1)*100 = -4.7% (slightly lower). Both Pass against 0.5 kgf spec. 10-sample basis only.',
         'source_file': name, 'sheet_name': '161016',
         'source_cells': ['IV. Decision 161016']},
        {'conclusion_id': 'concl_2',
         'topic': 'Tension test 201507 (spec 1.2 kgf)',
         'statement_from_report': 'Tension more good.',
         'normalized_interpretation': '395nm avg Tension 2.875 kgf vs Normal 365nm 2.503 kgf = (2.875/2.503 - 1)*100 = +14.9% (higher). Both Pass against 1.2 kgf spec; 395nm is higher on average. 10-sample basis only.',
         'source_file': name, 'sheet_name': '201507',
         'source_cells': ['IV. Decision 201507']},
        {'conclusion_id': 'concl_3',
         'topic': 'LED UV intensity 395nm vs 365nm',
         'statement_from_report': '(Numeric LED UV table provided without verbal judgement.)',
         'normalized_interpretation': '395nm machine readings (Peak/Total per position) are much lower than 365nm machine raw readings, but spec windows differ; 395nm "Peak: 600~900 mW/cm2 / Total 2500~3800" is the documented spec for 2nd lamp at 365nm. Tension is the actual gate and both lots Pass.',
         'source_file': name, 'sheet_name': '161016',
         'source_cells': ['LED UV table rows']},
    ]
    hints = [
        {'hint_id': 'hint_1',
         'check_item': 'New Dry UV 395nm can replace 365nm based on Tension test (both spec lots pass).',
         'reason': '161016: 1.310 vs 1.374 kgf (-4.7%, both > 0.5 kgf spec). 201507: 2.875 vs 2.503 kgf (+14.9%, both > 1.2 kgf spec).',
         'evidence_strength': 'medium', 'related_process': 'Dry UV',
         'related_part': 'VP+CD', 'source_file': name, 'sheet_name': sheet,
         'source_cells': ['Tension tables 161016, 201507']},
        {'hint_id': 'hint_2',
         'check_item': 'Confirm 395nm cures both 161016 and 201507 VP Film by larger lot before full rollout.',
         'reason': 'Tension test n=10 per arm only; no Function NG numbers in this report.',
         'evidence_strength': 'medium', 'related_process': 'Dry UV / Function',
         'related_part': 'VP Film', 'source_file': name, 'sheet_name': sheet,
         'source_cells': ['Tension n=10']},
    ]
    troubleshooting_index = {
        'defect_name': 'Dry UV Machine Change (395nm vs 365nm)',
        'when_user_asks': ['395nm Dry UV', 'Dry UV change wavelength', 'Tension VP CD 365 vs 395'],
        'suggested_checks': hints,
        'limitations': ['Sample n=10 per arm only.', 'No Function/Sigma/Hearing NG data in this report.'],
    }
    log = {
        'confidence': 0.75,
        'assumptions': ['Tension spec 0.5 kgf (161016) and 1.2 kgf (201507) from sheet header.',
                        'Min/Max/Avg quoted are reported by the workbook.'],
        'warnings': ['LED UV peak/total numbers between 395nm and 365nm are not directly comparable because spec windows reflect different lamps.', 'n=10 per arm.'],
        'decision_rationale': 'Classified reliability_spec: Tension PASS/FAIL vs spec for both lots. 395nm passes spec on both 161016 (avg 1.310 kgf vs spec 0.5 kgf) and 201507 (avg 2.875 kgf vs spec 1.2 kgf). 201507 395nm even improves Tension vs 365nm (+14.9%).',
    }

    result = {'schema_version': '0.1', 'document': doc, 'test_conditions': test_conditions,
              'results': results, 'conclusions': conclusions,
              'troubleshooting_index': troubleshooting_index, 'ai_extraction_log': log}

    tr_en = _trio_from_en({'title': doc['title'], 'purpose': doc['purpose'], 'content': doc['content']},
                          conclusions, hints, log)
    tr_ko = {
        'document': {
            'title': 'MSU-L20S15-07 신규 Dry UV 395nm vs Normal 365nm 시험 리포트',
            'purpose': '신규 395nm Dry UV 가 기존 365nm Dry UV 를 대체 가능한지 LED UV Peak/Total 및 VP+CD Tension 으로 판단.',
            'content': [
                "BRS-161016 (C2-3B): Peak/Total with/without film 측정, Tension spec 0.5 kgf, arm 당 10 sample.",
                "MSU-L20S15-07 201507 (C2-2A): 같은 비교, Tension spec 1.2 kgf, arm 당 10 sample.",
                "Normal 순서: UC PRESS → 365nm Dry1 → 365nm Dry2. Test 순서: UC PRESS → 395nm → 395nm.",
            ],
        },
        'conclusions': {
            'concl_1': {'topic': 'Tension 시험 161016 (spec 0.5 kgf)',
                        'statement_from_report': 'Tension normal 과 동일.',
                        'normalized_interpretation': '395nm 평균 Tension 1.310 kgf vs Normal 365nm 1.374 kgf = (1.310/1.374 - 1)*100 = -4.7% (소폭 낮음). 양쪽 모두 0.5 kgf spec Pass. 10 sample 한정.'},
            'concl_2': {'topic': 'Tension 시험 201507 (spec 1.2 kgf)',
                        'statement_from_report': 'Tension 더 양호.',
                        'normalized_interpretation': '395nm 평균 Tension 2.875 kgf vs Normal 365nm 2.503 kgf = (2.875/2.503 - 1)*100 = +14.9% (더 높음). 양쪽 모두 1.2 kgf spec Pass; 395nm 평균이 높음. 10 sample 한정.'},
            'concl_3': {'topic': 'LED UV 강도 395nm vs 365nm',
                        'statement_from_report': '(LED UV 수치 표 — 서술 판정 없음.)',
                        'normalized_interpretation': '395nm 측정치(Position 별 Peak/Total)는 365nm 원시 수치보다 훨씬 낮지만 spec window 가 다름; 문서화된 spec "Peak 600~900 / Total 2500~3800" 은 365nm 2nd 램프 기준. Gate 는 Tension 이며 두 lot 모두 Pass.'},
        },
        'hints': {
            'hint_1': {'check_item': '신규 Dry UV 395nm 는 Tension 시험 기준으로 365nm 대체 가능 (두 lot 모두 spec Pass).',
                       'reason': '161016: 1.310 vs 1.374 kgf (-4.7%, 둘 다 > 0.5 kgf). 201507: 2.875 vs 2.503 kgf (+14.9%, 둘 다 > 1.2 kgf).'},
            'hint_2': {'check_item': '본격 적용 전 161016/201507 VP Film 모두 더 큰 lot 으로 395nm 경화 검증.',
                       'reason': 'Tension 시험 arm 당 10 sample; 본 리포트에 Function NG 수치 없음.'},
        },
        'log': {
            'assumptions': ['Tension spec 0.5 kgf (161016), 1.2 kgf (201507) 은 시트 헤더 기반.',
                            'Min/Max/Avg 는 워크북이 보고한 값.'],
            'warnings': ['395nm 과 365nm 의 Peak/Total 절대값은 spec window 가 달라 직접 비교 불가.', 'arm 당 n=10.'],
            'decision_rationale': 'reliability_spec 분류: 두 lot 의 Tension PASS/FAIL spec 게이트. 395nm 가 161016 (avg 1.310 > 0.5) 와 201507 (avg 2.875 > 1.2) 모두 spec Pass; 201507 에서는 +14.9% 개선.',
        },
    }
    tr_vi = {
        'document': {
            'title': 'Báo cáo MSU-L20S15-07 thử máy Dry UV 395nm mới (Normal 365nm)',
            'purpose': 'Xác định máy Dry UV 395nm có thay được máy 365nm hay không, dựa trên LED UV Peak/Total và Tension VP+CD.',
            'content': [
                "BRS-161016 (C2-3B): đo Peak/Total có/không film; spec Tension 0.5 kgf; 10 mẫu/arm.",
                "MSU-L20S15-07 201507 (C2-2A): so sánh tương tự, spec Tension 1.2 kgf; 10 mẫu/arm.",
                "Trình tự Normal: UC PRESS → 365nm Dry1 → 365nm Dry2. Trình tự Test: UC PRESS → 395nm → 395nm.",
            ],
        },
        'conclusions': {
            'concl_1': {'topic': 'Thử Tension 161016 (spec 0.5 kgf)',
                        'statement_from_report': 'Tension giống Normal.',
                        'normalized_interpretation': 'Tension trung bình 395nm 1.310 kgf so với Normal 365nm 1.374 kgf = (1.310/1.374 - 1)*100 = -4.7% (thấp hơn chút). Cả hai đều Pass spec 0.5 kgf. Chỉ trên 10 mẫu.'},
            'concl_2': {'topic': 'Thử Tension 201507 (spec 1.2 kgf)',
                        'statement_from_report': 'Tension tốt hơn.',
                        'normalized_interpretation': 'Tension trung bình 395nm 2.875 kgf so với Normal 365nm 2.503 kgf = (2.875/2.503 - 1)*100 = +14.9% (cao hơn). Cả hai đều Pass spec 1.2 kgf; 395nm trung bình cao hơn. Chỉ trên 10 mẫu.'},
            'concl_3': {'topic': 'Cường độ LED UV 395nm vs 365nm',
                        'statement_from_report': '(Bảng số LED UV — không có nhận xét bằng lời.)',
                        'normalized_interpretation': 'Số đọc 395nm (Peak/Total mỗi Position) thấp hơn nhiều số đọc 365nm, nhưng cửa sổ spec khác nhau; spec "Peak 600~900 / Total 2500~3800" là cho lamp thứ 2 của 365nm. Cổng quyết định là Tension và cả hai lot đều Pass.'},
        },
        'hints': {
            'hint_1': {'check_item': 'Dry UV 395nm có thể thay 365nm dựa trên Tension (cả hai lot đều Pass).',
                       'reason': '161016: 1.310 vs 1.374 kgf (-4.7%, đều > 0.5 kgf). 201507: 2.875 vs 2.503 kgf (+14.9%, đều > 1.2 kgf).'},
            'hint_2': {'check_item': 'Xác nhận 395nm curing cho cả VP Film 161016 và 201507 trên lot lớn hơn trước khi nhân rộng.',
                       'reason': 'Tension chỉ n=10 mỗi arm; không có dữ liệu Function NG trong báo cáo này.'},
        },
        'log': {
            'assumptions': ['Spec Tension 0.5 kgf (161016) và 1.2 kgf (201507) lấy từ header sheet.',
                            'Min/Max/Avg trích từ workbook.'],
            'warnings': ['Số Peak/Total giữa 395nm và 365nm không so sánh trực tiếp được do cửa sổ spec khác nhau.', 'n=10 mỗi arm.'],
            'decision_rationale': 'Phân loại reliability_spec: cổng Tension PASS/FAIL theo spec. 395nm Pass spec cả 161016 (avg 1.310 > 0.5) và 201507 (avg 2.875 > 1.2); 201507 còn cải thiện +14.9%.',
        },
    }
    return name, result, tr_ko, tr_en, tr_vi


# =====================================================================
# Driver
# =====================================================================
BUILDERS = [
    build_tiu_c11_20_frame,
    build_tiu_l5s3_bako,
    build_msu_nti_bond,
    build_brs_201506_v1,
    build_brs_161014_over_glue,
    build_brs_161016_vision,
    build_brs_201506_v2,
    build_msu_list_test,
    build_msu_uv_395nm,
]


def main() -> int:
    targets = []
    with open(r'D:\000. MyWorks\005. Program\Repository\JinoSupporter\_chunk_04.txt', 'r', encoding='utf-8-sig') as f:
        targets = [l.rstrip('\r\n') for l in f if l.strip()]
    assert len(targets) == len(BUILDERS), f'targets={len(targets)} vs builders={len(BUILDERS)}'

    processed = 0
    failed = 0
    for fn, expected_name in zip(BUILDERS, targets):
        name, result, tr_ko, tr_en, tr_vi = fn()
        if name != expected_name:
            print(f'[NAME MISMATCH] builder produced {name!r} but chunk expects {expected_name!r}')
            h.log_failed(expected_name, f'builder name mismatch: {name}')
            failed += 1
            continue
        ok = h.commit_dataset(name, result, tr_ko, tr_en, tr_vi)
        if ok:
            print(f'[OK] {name}')
            processed += 1
        else:
            print(f'[FAIL] {name}')
            failed += 1

    print(f'chunk 04: processed={processed} failed={failed}')

    # verify_counts uses helper's TARGETS_FILE which points elsewhere; verify chunk-04 directly.
    import sqlite3
    con = sqlite3.connect(h.DB_PATH)
    try:
        ok_in_db = 0
        for t in targets:
            r = con.execute('SELECT COUNT(*) FROM AiDocuments WHERE SourceDataset=?', (t,)).fetchone()[0]
            if r > 0:
                ok_in_db += 1
        print(f'verify_counts(chunk04): targets={len(targets)} present_in_AiDocuments={ok_in_db}')
    finally:
        con.close()

    return 0 if failed == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
