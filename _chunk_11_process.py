"""Process chunk 11 (8 datasets, final) — AI Batch normalization."""
from __future__ import annotations
import sys, os, io
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import _ai_batch_helper as h


# =====================================================================
# DATASET 1: 32. TIU C11-20  Report test Plate  difference colour 100% 2026.04.16
# normal_comparison: Test PT difference color vs Normal across multiple metrics
# =====================================================================
ds1_name = '32. TIU C11-20  Report test Plate  difference colour 100% 2026.04.16'

ds1_result = {
    'schema_version': '0.1',
    'document': {
        'document_id': 'doc_1',
        'source_file': ds1_name,
        'source_sheet': 'Test',
        'title': 'REPORT TEST PLATE DIFFERENCE COLOR 100% C11-20',
        'model': 'TIU C11-20',
        'report_date': '2026-04-16',
        'department': 'ME',
        'marker': 'Thao',
        'line': '',
        'report_type': 'normal_comparison',
        'primary_defect': {
            'canonical_name': 'VP+CD Separation',
            'aliases_in_document': ['MG+PT separate', 'NG bonding not spread']
        },
        'related_defects': ['NG Hearing Noise', 'Over Glue', 'Offset NG', 'Not Dry Glue'],
        'parts': ['PT', 'MG', 'YOKE'],
        'processes': ['Array PT', 'AI Bonding PT', 'Visual YOKE', 'Decap Bonding', 'Drop Test', 'Function'],
        'purpose': 'Test whether PT plate of different colour can be used.',
        'content': [
            'Make semi sub 2 and check NG process, Decap PT+MG, Drop test, AI bonding detection, Function.'
        ],
        'source_cells': {'title': ['Test!B1'], 'date': ['Test!Q2'], 'purpose': ['Test!A4'], 'content': ['Test!A6']}
    },
    'test_conditions': [
        {'condition_id': 'cond_1', 'condition_group': 'plate_color', 'line': '', 'process': 'Bonding',
         'changed_factor': 'PT plate color (different color lot)', 'before_value': 'Normal color',
         'after_value': 'Test different color', 'unit': None, 'machine': None, 'jig': None,
         'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None,
         'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': ds1_name, 'sheet_name': 'Test', 'source_cells': ['Test!A4']}
    ],
    'results': [
        {'result_id': 'res_array_test', 'condition_id': 'cond_1', 'measurement_type': 'Array PT',
         'condition_group': 'Test PT diff color', 'date': '2026-04-16', 'line': '',
         'input_count': 100, 'ok_count': 100, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Array PT NG Rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'pick_up': {'count': 0, 'rate': 0.0}},
         'source_file': ds1_name, 'sheet_name': 'Test', 'source_cells': ['Test!E16:J16']},
        {'result_id': 'res_array_normal', 'condition_id': None, 'measurement_type': 'Array PT',
         'condition_group': 'Normal', 'date': '2026-04-16', 'line': '',
         'input_count': 200, 'ok_count': 200, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Array PT NG Rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'pick_up': {'count': 0, 'rate': 0.0}},
         'source_file': ds1_name, 'sheet_name': 'Test', 'source_cells': ['Test!E18:J18']},
        {'result_id': 'res_aibond_test', 'condition_id': 'cond_1', 'measurement_type': 'AI Bonding PT',
         'condition_group': 'Test PT diff color', 'date': '2026-04-16', 'line': '',
         'input_count': 100, 'ok_count': 100, 'ng_count': 1, 'ng_rate_decimal': 0.01, 'ng_rate_percent': 1.0,
         'metric_name': 'AI Bonding NG Rate', 'metric_value': 1.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'Over Glue': {'count': 0, 'rate': 0.0}, 'Offset NG': {'count': 0, 'rate': 0.0},
                          'Not Dry Glue': {'count': 1, 'rate': 1.0}},
         'source_file': ds1_name, 'sheet_name': 'Test', 'source_cells': ['Test!E24:K24']},
        {'result_id': 'res_aibond_normal', 'condition_id': None, 'measurement_type': 'AI Bonding PT',
         'condition_group': 'Normal', 'date': '2026-04-16', 'line': '',
         'input_count': 200, 'ok_count': 200, 'ng_count': 3, 'ng_rate_decimal': 0.015, 'ng_rate_percent': 1.5,
         'metric_name': 'AI Bonding NG Rate', 'metric_value': 1.5, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'Over Glue': {'count': 1, 'rate': 0.5}, 'Offset NG': {'count': 0, 'rate': 0.0},
                          'Not Dry Glue': {'count': 2, 'rate': 1.0}},
         'source_file': ds1_name, 'sheet_name': 'Test', 'source_cells': ['Test!E26:K26']},
        {'result_id': 'res_visual_test', 'condition_id': 'cond_1', 'measurement_type': 'Visual YOKE',
         'condition_group': 'Test PT diff color', 'date': '2026-04-16', 'line': '',
         'input_count': 93, 'ok_count': 93, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Visual YOKE NG Rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {},
         'source_file': ds1_name, 'sheet_name': 'Test', 'source_cells': ['Test!E32:K32']},
        {'result_id': 'res_visual_normal', 'condition_id': None, 'measurement_type': 'Visual YOKE',
         'condition_group': 'Normal', 'date': '2026-04-16', 'line': '',
         'input_count': 100, 'ok_count': 100, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Visual YOKE NG Rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {},
         'source_file': ds1_name, 'sheet_name': 'Test', 'source_cells': ['Test!E34:K34']},
        {'result_id': 'res_decap_test', 'condition_id': 'cond_1', 'measurement_type': 'Decap Bonding',
         'condition_group': 'Test PT diff color', 'date': '2026-04-16', 'line': '',
         'input_count': 53, 'ok_count': 52, 'ng_count': 1, 'ng_rate_decimal': 0.019, 'ng_rate_percent': 1.9,
         'metric_name': 'Decap NG Rate', 'metric_value': 1.9, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'MG+PT separate (bonding not spread)': {'count': 1, 'rate': 1.9}},
         'source_file': ds1_name, 'sheet_name': 'Test', 'source_cells': ['Test!E38:H38']},
        {'result_id': 'res_decap_normal', 'condition_id': None, 'measurement_type': 'Decap Bonding',
         'condition_group': 'Normal', 'date': '2026-04-16', 'line': '',
         'input_count': 8, 'ok_count': 8, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Decap NG Rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {},
         'source_file': ds1_name, 'sheet_name': 'Test', 'source_cells': ['Test!E39:H39']},
        {'result_id': 'res_drop_test', 'condition_id': 'cond_1', 'measurement_type': 'Drop Test',
         'condition_group': 'Test PT diff color', 'date': '2026-04-16', 'line': '',
         'input_count': 8, 'ok_count': 8, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Drop Test', 'metric_value': None, 'unit': None, 'judgement': 'PASS',
         'ng_breakdown': {},
         'source_file': ds1_name, 'sheet_name': 'Test', 'source_cells': ['Test!E42:M44']},
        {'result_id': 'res_func_test', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Test PT diff color', 'date': '2026-04-16', 'line': '',
         'input_count': 40, 'ok_count': 39, 'ng_count': 1, 'ng_rate_decimal': 0.025, 'ng_rate_percent': 2.5,
         'metric_name': 'Function NG Rate', 'metric_value': 2.5, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'NG Sigma SPL': {'count': 0, 'rate': 0.0}, 'NG Hearing Noise': {'count': 1, 'rate': 2.5}, 'NG Hearing Touch': {'count': 0, 'rate': 0.0}},
         'source_file': ds1_name, 'sheet_name': 'Test', 'source_cells': ['Test!E52:K52']},
        {'result_id': 'res_func_normal', 'condition_id': None, 'measurement_type': 'Function',
         'condition_group': 'Normal', 'date': '2026-04-16', 'line': '',
         'input_count': 100, 'ok_count': 98, 'ng_count': 2, 'ng_rate_decimal': 0.02, 'ng_rate_percent': 2.0,
         'metric_name': 'Function NG Rate', 'metric_value': 2.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'NG Sigma SPL': {'count': 0, 'rate': 0.0}, 'NG Hearing Noise': {'count': 2, 'rate': 2.0}, 'NG Hearing Touch': {'count': 0, 'rate': 0.0}},
         'source_file': ds1_name, 'sheet_name': 'Test', 'source_cells': ['Test!E54:K54']},
    ],
    'conclusions': [
        {'conclusion_id': 'concl_1', 'topic': 'Plate different color usage decision',
         'statement_from_report': 'Result Decap Time 1 NG 1/8 => when use lot PT different color can happen NG bonding not spread => MG+PT separate. Other tests (Array PT, AI Bonding, Visual YOKE, Drop, Function) OK same Normal. Function NG Noise reason by particle, not by PT+MG separate.',
         'normalized_interpretation': 'Test PT different color: AI Bonding 1.0% vs Normal 1.5% (0.67x, 33.3% improved); Function 2.5% vs Normal 2.0% (1.25x, 25.0% worse, marginal); Array/Visual/Drop equal at 0%. Decap shows 1/8 (Time 1) MG+PT separate vs 0/8 Normal — this is the only signal of risk when using different-color PT lot.',
         'source_file': ds1_name, 'sheet_name': 'Test', 'source_cells': ['Test!A40', 'Test!A56-A60']}
    ],
    'troubleshooting_index': {
        'defect_name': 'VP+CD Separation',
        'when_user_asks': ['Plate color change', 'MG+PT separate', 'Bonding not spread', 'Decap NG'],
        'suggested_checks': [
            {'hint_id': 'hint_1', 'check_item': 'Run Decap test on multiple time steps when changing PT plate color lot',
             'reason': 'In this report, Decap Time 1 produced 1/8 MG+PT separate vs Normal 0/8 with different-color PT lot; the deviation only appeared in Decap, not in Array/Visual/AI-Bonding.',
             'evidence_strength': 'medium', 'related_process': 'Decap Bonding', 'related_part': 'PT/MG',
             'source_file': ds1_name, 'sheet_name': 'Test', 'source_cells': ['Test!A40']}
        ],
        'limitations': ['Decap sample size very small (8 per normal, 53 test).']
    },
    'ai_extraction_log': {
        'confidence': 0.8,
        'assumptions': ['Date 26/Feb in cell appears to be a misprint; title and filename indicate 2026-04-16 (16-Apr).'],
        'warnings': ['Decap normal sample size only 8 — limited statistical strength.'],
        'decision_rationale': 'Same-event Normal exists for every process (Array, AI Bonding, Visual, Decap, Drop, Function). Used multiplicative relative change. Only Decap shows a degradation signal (1/53 = 1.9% vs 0/8 = 0.0%), and the workbook itself flags MG+PT separate risk.'
    }
}

ds1_tr_en = {
    'document': {'title': 'Report — Test PT plate different color 100%, TIU C11-20',
                 'purpose': 'Verify whether a PT plate of different colour lot can be used.',
                 'content': ['Make semi sub 2 and verify NG process, Decap PT+MG, Drop test, AI bonding rate, Function.']},
    'conclusions': {'concl_1': {'topic': 'Plate different color usage decision',
                                'statement_from_report': 'Decap Time 1 NG 1/8 — different-color PT lot can cause MG+PT separate (bonding not spread). Array, AI Bonding, Visual YOKE, Drop, and Function OK same Normal. Function NG Noise reason by particle, not PT+MG separate.',
                                'normalized_interpretation': 'Function 2.5% vs Normal 2.0% (1.25x, 25.0% worse, marginal). AI Bonding 1.0% vs Normal 1.5% (0.67x, improved). Decap 1.9% vs Normal 0.0% — small-sample but the only failure signal. The different-color PT lot poses a localized MG+PT separate risk.'}},
    'hints': {'hint_1': {'check_item': 'Run Decap test on multiple time steps when changing PT plate color lot',
                         'reason': 'Decap Time 1 produced 1/8 MG+PT separate vs Normal 0/8 with different-color PT lot — signal appears only in Decap, not in upstream Array/Visual/AI-Bonding.'}},
    'log': {'assumptions': ['Date 26/Feb in cell appears to be a misprint; title and filename indicate 2026-04-16.'],
            'warnings': ['Decap normal sample size only 8 — limited statistical strength.'],
            'decision_rationale': 'Same-event Normal exists for every process. Multiplicative relative change used. Only Decap shows degradation; workbook itself flags MG+PT separate risk.'}
}

ds1_tr_ko = {
    'document': {'title': '리포트 — PT 플레이트 다른 색상 100% 테스트, TIU C11-20',
                 'purpose': 'PT 다른 색상 lot 사용 가능 여부 검증.',
                 'content': ['Semi sub 2 제작 후 NG process, Decap PT+MG, Drop test, AI bonding rate, Function 검증.']},
    'conclusions': {'concl_1': {'topic': '다른 색상 PT 플레이트 사용 결정',
                                'statement_from_report': 'Decap Time 1 NG 1/8 — 다른 색상 PT lot은 MG+PT separate(bonding not spread) 유발 가능. Array, AI Bonding, Visual YOKE, Drop, Function은 Normal과 동일. Function NG Noise는 particle 원인이며 PT+MG separate 아님.',
                                'normalized_interpretation': 'Function 2.5% vs Normal 2.0% (1.25배, 25.0% 악화, 한계점). AI Bonding 1.0% vs Normal 1.5% (0.67배, 개선). Decap 1.9% vs Normal 0.0% — 표본은 작지만 유일한 결함 신호. 다른 색상 PT lot은 국소적 MG+PT separate 위험 있음.'}},
    'hints': {'hint_1': {'check_item': 'PT 플레이트 색상 lot 변경 시 여러 Decap time 조건에서 시험',
                         'reason': '다른 색상 PT lot에서 Decap Time 1 1/8 MG+PT separate, Normal 0/8 — 신호가 Array/Visual/AI-Bonding이 아닌 Decap 단계에서만 나타남.'}},
    'log': {'assumptions': ['셀의 26/Feb 날짜는 오기로 보임; 제목과 파일명은 2026-04-16 지시.'],
            'warnings': ['Decap normal 표본 8개 — 통계적 신뢰도 제한.'],
            'decision_rationale': '모든 공정에 same-event Normal 존재. 곱셈 상대변화율 사용. Decap만 악화 신호; 리포트가 MG+PT separate 위험을 명시.'}
}

ds1_tr_vi = {
    'document': {'title': 'Báo cáo — Test PT plate khác màu 100%, TIU C11-20',
                 'purpose': 'Kiểm tra xem PT plate lot khác màu có thể dùng được hay không.',
                 'content': ['Làm semi sub 2 và kiểm tra NG process, Decap PT+MG, Drop test, AI bonding rate, Function.']},
    'conclusions': {'concl_1': {'topic': 'Quyết định sử dụng PT plate khác màu',
                                'statement_from_report': 'Decap Time 1 NG 1/8 — lot PT khác màu có thể gây MG+PT separate (bonding không lan đều). Array, AI Bonding, Visual YOKE, Drop, Function OK same Normal. Function NG Noise do particle, không phải PT+MG separate.',
                                'normalized_interpretation': 'Function 2.5% vs Normal 2.0% (1.25x, xấu hơn 25.0%, biên độ nhỏ). AI Bonding 1.0% vs Normal 1.5% (0.67x, cải thiện). Decap 1.9% vs Normal 0.0% — mẫu nhỏ nhưng là tín hiệu lỗi duy nhất. PT lot khác màu có rủi ro MG+PT separate cục bộ.'}},
    'hints': {'hint_1': {'check_item': 'Chạy Decap nhiều mốc thời gian khi đổi lot PT khác màu',
                         'reason': 'Decap Time 1 1/8 MG+PT separate vs Normal 0/8 với lot PT khác màu — chỉ Decap phát hiện, Array/Visual/AI-Bonding không thấy.'}},
    'log': {'assumptions': ['Ngày 26/Feb trong ô có vẻ là lỗi in; tiêu đề và tên file chỉ 2026-04-16.'],
            'warnings': ['Mẫu Decap Normal chỉ 8 — độ tin cậy hạn chế.'],
            'decision_rationale': 'Mỗi process có same-event Normal. Dùng thay đổi tương đối nhân. Chỉ Decap xấu đi; bản báo cáo cũng cảnh báo MG+PT separate.'}
}


# =====================================================================
# DATASET 2: 32. TIU L5S3-01 R Report test find reason NG BAKO high 2025.12.19y
# normal_comparison: AWF machines + new bonding line vs Normal/Old bonding line
# =====================================================================
ds2_name = '32. TIU L5S3-01 R Report test find reason NG BAKO high 2025.12.19y'

ds2_result = {
    'schema_version': '0.1',
    'document': {
        'document_id': 'doc_2',
        'source_file': ds2_name,
        'source_sheet': 'Test',
        'title': 'TIU L5S3-01 [R] REPORT TEST FIND REASON IMPROVE NG FUNCTION (NG BAKO high)',
        'model': 'TIU L5S3-01 R',
        'report_date': '2025-12-19',
        'department': 'ME',
        'marker': 'Thao',
        'line': '',
        'report_type': 'normal_comparison',
        'primary_defect': {'canonical_name': 'NG Function BAKO FRF',
                           'aliases_in_document': ['NG BAKO', 'FRF', 'NG function Bako']},
        'related_defects': ['FRF+SPL', 'THD', 'No sound'],
        'parts': ['Frame', 'VP'],
        'processes': ['AWF machine', 'Frame/VP bonding line', 'Standard sample setting'],
        'purpose': 'NG function BAKO rate is very high (~50%); identify root cause and improvement.',
        'content': ['Type 1: separate AWF machines #1-#4; Type 2: run lot through new Frame/VP bonding line; also re-set standard sample.'],
        'source_cells': {'title': ['Test!B1'], 'date': ['Test!N2'], 'purpose': ['Test!A4'], 'content': ['Test!A6']}
    },
    'test_conditions': [
        {'condition_id': 'cond_1', 'condition_group': 'awf_machine', 'line': '', 'process': 'AWF',
         'changed_factor': 'Separate AWF machine #1/#2/#3/#4', 'before_value': None,
         'after_value': None, 'unit': None, 'machine': 'AWF #1-#4', 'jig': None,
         'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None,
         'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': ds2_name, 'sheet_name': 'Test', 'source_cells': ['Test!A6']},
        {'condition_id': 'cond_2', 'condition_group': 'bonding_line', 'line': '', 'process': 'Frame/VP bonding',
         'changed_factor': 'Change to new Frame/VP bonding line', 'before_value': 'Old bonding line',
         'after_value': 'New bonding line', 'unit': None, 'machine': None, 'jig': None,
         'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None,
         'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': ds2_name, 'sheet_name': 'Test', 'source_cells': ['Test!A7']},
        {'condition_id': 'cond_3', 'condition_group': 'standard_sample', 'line': '', 'process': 'Function inspection',
         'changed_factor': 'Re-set standard sample (Room 12)', 'before_value': 'Before reset',
         'after_value': 'After reset', 'unit': None, 'machine': None, 'jig': None,
         'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None,
         'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': ds2_name, 'sheet_name': 'Test', 'source_cells': ['Test!A_standard']},
    ],
    'results': [
        {'result_id': 'awf1', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'AWF #1', 'date': '2025-12-19', 'line': '',
         'input_count': 50, 'ok_count': 19, 'ng_count': 31, 'ng_rate_decimal': 0.62, 'ng_rate_percent': 62.0,
         'metric_name': 'NG BAKO rate', 'metric_value': 62.0, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'FRF': {'count': 29, 'rate': 58.0}, 'FRF+SPL': {'count': 2, 'rate': 4.0},
                          'THD': {'count': 0, 'rate': 0.0}, 'No sound': {'count': 0, 'rate': 0.0}},
         'source_file': ds2_name, 'sheet_name': 'Test', 'source_cells': ['Test_awf1']},
        {'result_id': 'awf2', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'AWF #2', 'date': '2025-12-19', 'line': '',
         'input_count': 50, 'ok_count': 19, 'ng_count': 31, 'ng_rate_decimal': 0.62, 'ng_rate_percent': 62.0,
         'metric_name': 'NG BAKO rate', 'metric_value': 62.0, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'FRF': {'count': 31, 'rate': 62.0}, 'FRF+SPL': {'count': 0, 'rate': 0.0},
                          'THD': {'count': 0, 'rate': 0.0}, 'No sound': {'count': 0, 'rate': 0.0}},
         'source_file': ds2_name, 'sheet_name': 'Test', 'source_cells': ['Test_awf2']},
        {'result_id': 'awf3', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'AWF #3', 'date': '2025-12-19', 'line': '',
         'input_count': 45, 'ok_count': 17, 'ng_count': 28, 'ng_rate_decimal': 0.622, 'ng_rate_percent': 62.2,
         'metric_name': 'NG BAKO rate', 'metric_value': 62.2, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'FRF': {'count': 28, 'rate': 62.2}, 'FRF+SPL': {'count': 0, 'rate': 0.0},
                          'THD': {'count': 0, 'rate': 0.0}, 'No sound': {'count': 0, 'rate': 0.0}},
         'source_file': ds2_name, 'sheet_name': 'Test', 'source_cells': ['Test_awf3']},
        {'result_id': 'awf4', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'AWF #4', 'date': '2025-12-19', 'line': '',
         'input_count': 57, 'ok_count': 24, 'ng_count': 33, 'ng_rate_decimal': 0.579, 'ng_rate_percent': 57.9,
         'metric_name': 'NG BAKO rate', 'metric_value': 57.9, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'FRF': {'count': 25, 'rate': 43.9}, 'FRF+SPL': {'count': 8, 'rate': 14.0},
                          'THD': {'count': 0, 'rate': 0.0}, 'No sound': {'count': 0, 'rate': 0.0}},
         'source_file': ds2_name, 'sheet_name': 'Test', 'source_cells': ['Test_awf4']},
        {'result_id': 'newline_test', 'condition_id': 'cond_2', 'measurement_type': 'Function',
         'condition_group': 'Test new Frame/VP bonding line', 'date': '2025-12-19', 'line': '',
         'input_count': 195, 'ok_count': 179, 'ng_count': 16, 'ng_rate_decimal': 0.082, 'ng_rate_percent': 8.2,
         'metric_name': 'NG BAKO rate', 'metric_value': 8.2, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'FRF': {'count': 8, 'rate': 4.1}, 'FRF+SPL': {'count': 8, 'rate': 4.1},
                          'THD': {'count': 0, 'rate': 0.0}, 'No sound': {'count': 0, 'rate': 0.0}},
         'source_file': ds2_name, 'sheet_name': 'Test', 'source_cells': ['Test_newline']},
        {'result_id': 'oldline_normal', 'condition_id': None, 'measurement_type': 'Function',
         'condition_group': 'Normal (Old bonding line Frame/VP)', 'date': '2025-12-19', 'line': '',
         'input_count': 660, 'ok_count': 291, 'ng_count': 369, 'ng_rate_decimal': 0.559, 'ng_rate_percent': 55.9,
         'metric_name': 'NG BAKO rate', 'metric_value': 55.9, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'FRF': {'count': 361, 'rate': 54.7}, 'FRF+SPL': {'count': 8, 'rate': 1.2},
                          'THD': {'count': 0, 'rate': 0.0}, 'No sound': {'count': 0, 'rate': 0.0}},
         'source_file': ds2_name, 'sheet_name': 'Test', 'source_cells': ['Test_oldline']},
        {'result_id': 'std_before', 'condition_id': 'cond_3', 'measurement_type': 'Function',
         'condition_group': 'Before setting again sample standard (Room 12)', 'date': '2025-12-19', 'line': '',
         'input_count': 40, 'ok_count': 12, 'ng_count': 28, 'ng_rate_decimal': 0.70, 'ng_rate_percent': 70.0,
         'metric_name': 'NG BAKO rate', 'metric_value': 70.0, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'FRF': {'count': 28, 'rate': 70.0}},
         'source_file': ds2_name, 'sheet_name': 'Test', 'source_cells': ['Test_stdbefore']},
        {'result_id': 'std_after', 'condition_id': 'cond_3', 'measurement_type': 'Function',
         'condition_group': 'After setting again sample standard (Room 12)', 'date': '2025-12-19', 'line': '',
         'input_count': 40, 'ok_count': 27, 'ng_count': 13, 'ng_rate_decimal': 0.325, 'ng_rate_percent': 32.5,
         'metric_name': 'NG BAKO rate', 'metric_value': 32.5, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'FRF': {'count': 13, 'rate': 32.5}},
         'source_file': ds2_name, 'sheet_name': 'Test', 'source_cells': ['Test_stdafter']},
    ],
    'conclusions': [
        {'conclusion_id': 'concl_1', 'topic': 'NG BAKO high root cause and counter-measure',
         'statement_from_report': 'AWF #1, #2, #3 NG BAKO is very high (~62%). Test running lot through new Frame/VP bonding line reduces NG significantly. Resetting standard sample also halves NG.',
         'normalized_interpretation': 'AWF separation shows all four machines uniformly high (57.9-62.2%) — not machine-specific. New bonding line 8.2% vs Old (Normal) 55.9% = 0.147x, 85.3% improved. Standard sample reset 32.5% vs 70.0% = 0.464x, 53.6% improved. Two strong improvement levers identified: bonding line and inspection standard. Dominant defect is FRF.',
         'source_file': ds2_name, 'sheet_name': 'Test', 'source_cells': ['Test!A_decision']}
    ],
    'troubleshooting_index': {
        'defect_name': 'NG Function BAKO FRF',
        'when_user_asks': ['NG BAKO high', 'FRF NG', 'Function NG rate high', 'Frame VP bonding'],
        'suggested_checks': [
            {'hint_id': 'hint_1', 'check_item': 'Switch lot to the new Frame/VP bonding line and re-verify NG BAKO; also revisit inspection standard sample',
             'reason': 'In this test, new bonding line dropped NG from 55.9% to 8.2% (85.3% improvement). After standard sample reset, NG fell from 70.0% to 32.5% (53.6% improvement). Both indicate the high NG is process+gauge issue, not AWF-machine specific.',
             'evidence_strength': 'high', 'related_process': 'Frame/VP bonding + Function inspection',
             'related_part': 'Frame, VP',
             'source_file': ds2_name, 'sheet_name': 'Test', 'source_cells': ['Test!_newline', 'Test!_stdafter']}
        ],
        'limitations': ['Decision (IV) section is empty in the workbook — no final action noted.']
    },
    'ai_extraction_log': {
        'confidence': 0.85,
        'assumptions': ['"Normal" labelled on Old bonding line row (660 input) is treated as baseline for the bonding-line comparison.'],
        'warnings': ['AWF #1-#4 table has no same-event Normal row, only inter-machine comparison.'],
        'decision_rationale': 'Two same-event baselines available: Old bonding line (Normal) vs new bonding line, and Before-reset vs After-reset. Both deltas computed multiplicatively. AWF separation table judged as DOE without baseline; all four very high indicates root cause not machine-specific.'
    }
}

ds2_tr_en = {
    'document': {'title': 'TIU L5S3-01 R — Find reason NG Function BAKO high',
                 'purpose': 'NG Function BAKO is very high (~50%); identify root cause and counter-measure.',
                 'content': ['Test 1: separate AWF #1-#4. Test 2: switch lot to new Frame/VP bonding line. Also reset inspection standard sample.']},
    'conclusions': {'concl_1': {'topic': 'NG BAKO high root cause and counter-measure',
                                'statement_from_report': 'AWF #1, #2, #3 NG BAKO is very high (~62%). Running lot through new Frame/VP bonding line reduces NG significantly. Resetting standard sample also lowers NG.',
                                'normalized_interpretation': 'AWF separation shows all four AWF machines uniformly high (57.9-62.2%) — not machine-specific. New bonding line 8.2% vs Old (Normal) 55.9% = 0.147x, 85.3% improved. Standard sample reset 32.5% vs 70.0% = 0.464x, 53.6% improved. Two strong levers: bonding line and inspection standard. Dominant defect is FRF.'}},
    'hints': {'hint_1': {'check_item': 'Switch lot to new Frame/VP bonding line and reset standard sample',
                         'reason': 'New bonding line dropped NG from 55.9% to 8.2% (85.3% improvement). After standard sample reset, NG dropped from 70.0% to 32.5% (53.6% improvement). NG is process+gauge issue, not AWF-machine specific.'}},
    'log': {'assumptions': ['Old bonding line row labelled "Normal" used as baseline for bonding-line comparison.'],
            'warnings': ['AWF #1-#4 table has no same-event Normal row.'],
            'decision_rationale': 'Two same-event baselines (bonding line; standard sample) available. Multiplicative deltas computed. AWF separation evaluated as DOE without baseline.'}
}

ds2_tr_ko = {
    'document': {'title': 'TIU L5S3-01 R — NG Function BAKO 높은 원인 파악',
                 'purpose': 'NG Function BAKO 매우 높음(~50%); 근본 원인 및 대책 파악.',
                 'content': ['Test 1: AWF #1-#4 분리. Test 2: 새 Frame/VP bonding line으로 lot 진행. 또한 검사 standard sample 재설정.']},
    'conclusions': {'concl_1': {'topic': 'NG BAKO 높은 원인 및 대책',
                                'statement_from_report': 'AWF #1, #2, #3 NG BAKO 매우 높음(~62%). 새 Frame/VP bonding line 사용 시 NG 크게 감소. Standard sample 재설정도 NG 감소.',
                                'normalized_interpretation': 'AWF 분리 결과 4대 모두 57.9-62.2%로 균일 — 기기별 문제 아님. 새 bonding line 8.2% vs Old(Normal) 55.9% = 0.147배, 85.3% 개선. Standard sample 재설정 32.5% vs 70.0% = 0.464배, 53.6% 개선. 두 핵심 레버 확인: bonding line, 검사 기준. 주요 결함은 FRF.'}},
    'hints': {'hint_1': {'check_item': '새 Frame/VP bonding line 전환 및 standard sample 재설정',
                         'reason': '새 bonding line으로 NG 55.9% → 8.2%(85.3% 개선). Standard sample 재설정 후 70.0% → 32.5%(53.6% 개선). NG는 공정+게이지 문제이며 AWF 기기 특정 아님.'}},
    'log': {'assumptions': ['"Normal"로 표시된 Old bonding line row를 bonding line 비교의 baseline으로 사용.'],
            'warnings': ['AWF #1-#4 표에는 same-event Normal row 없음.'],
            'decision_rationale': '두 종의 same-event baseline 존재(bonding line; standard sample). 곱셈 변화율 계산. AWF 분리는 baseline 없는 DOE로 평가.'}
}

ds2_tr_vi = {
    'document': {'title': 'TIU L5S3-01 R — Tìm nguyên nhân NG Function BAKO cao',
                 'purpose': 'NG Function BAKO rất cao (~50%); xác định nguyên nhân và biện pháp.',
                 'content': ['Test 1: tách máy AWF #1-#4. Test 2: chạy lot qua bonding line Frame/VP mới. Đồng thời reset standard sample kiểm tra.']},
    'conclusions': {'concl_1': {'topic': 'Nguyên nhân và biện pháp NG BAKO cao',
                                'statement_from_report': 'AWF #1, #2, #3 NG BAKO rất cao (~62%). Chạy lot qua bonding line Frame/VP mới làm giảm NG đáng kể. Reset standard sample cũng giảm NG.',
                                'normalized_interpretation': 'Tách AWF cho thấy cả 4 máy đồng đều 57.9-62.2% — không do máy. Bonding line mới 8.2% vs Old(Normal) 55.9% = 0.147x, cải thiện 85.3%. Reset standard sample 32.5% vs 70.0% = 0.464x, cải thiện 53.6%. Hai đòn bẩy mạnh: bonding line và chuẩn kiểm. Lỗi chủ đạo là FRF.'}},
    'hints': {'hint_1': {'check_item': 'Chuyển lot sang bonding line Frame/VP mới và reset standard sample',
                         'reason': 'Bonding line mới giảm NG 55.9% → 8.2% (cải thiện 85.3%). Sau reset standard sample, NG giảm 70.0% → 32.5% (53.6%). NG do process+chuẩn đo, không do máy AWF.'}},
    'log': {'assumptions': ['Dòng "Normal" trên Old bonding line dùng làm baseline cho so sánh bonding line.'],
            'warnings': ['Bảng AWF #1-#4 không có dòng Normal same-event.'],
            'decision_rationale': 'Có hai baseline same-event (bonding line; standard sample). Tính delta nhân. AWF tách đánh giá như DOE không có baseline.'}
}


# =====================================================================
# DATASET 3: 32.1 BRS-161014 Report TEST AWF #1, #2, #3, #4 2023.09.20
# doe_matrix: AWF machine factor with jig size/pole settings; no same-event Normal -> per-machine comparison
# =====================================================================
ds3_name = '32.1 BRS-161014 Report TEST AWF #1, #2, #3, #4 2023.09.20'

ds3_result = {
    'schema_version': '0.1',
    'document': {
        'document_id': 'doc_3',
        'source_file': ds3_name,
        'source_sheet': 'Report (2)',
        'title': 'REPORT TEST AWF MACHINE — BRS-161014',
        'model': 'BRS-161014',
        'report_date': '2023-09-20',
        'department': 'ME',
        'marker': 'Thao/Thuy',
        'line': '',
        'report_type': 'doe_matrix',
        'primary_defect': {'canonical_name': 'NG Hearing Noise',
                           'aliases_in_document': ['Noise', 'Hearing Noise']},
        'related_defects': ['NG Sigma SPL', 'NG Hearing Touch'],
        'parts': ['Diaphragm/AWF coil'],
        'processes': ['AWF (Winding/Stretching/Cooling)'],
        'purpose': 'Improve high NG Function by separating AWF machines #1-#4 and comparing.',
        'content': ['Test each AWF machine #1-#4 with specified Winding Jig size, Stretching Pole, Cooler, Tension.'],
        'source_cells': {'title': ['Report (2)!B1'], 'date': ['Report (2)!Q3'], 'purpose': ['Report (2)!A4'],
                         'content': ['Report (2)!A6']}
    },
    'test_conditions': [
        {'condition_id': 'cond_1', 'condition_group': 'awf_machine', 'line': '', 'process': 'AWF',
         'changed_factor': 'AWF machine + Winding Jig size + Stretching Pole',
         'before_value': None, 'after_value': None, 'unit': None, 'machine': 'AWF #1',
         'jig': 'Winding 9.34 / Stretch 5.065', 'material_lot': None, 'supplier': None,
         'dry_time_sec': 8.0, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': ds3_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!A8']},
        {'condition_id': 'cond_2', 'condition_group': 'awf_machine', 'line': '', 'process': 'AWF',
         'changed_factor': 'AWF machine + Winding Jig size + Stretching Pole',
         'before_value': None, 'after_value': None, 'unit': None, 'machine': 'AWF #2',
         'jig': 'Winding 9.42 / Stretch 5.065', 'material_lot': None, 'supplier': None,
         'dry_time_sec': 8.0, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': ds3_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!A9']},
        {'condition_id': 'cond_3', 'condition_group': 'awf_machine', 'line': '', 'process': 'AWF',
         'changed_factor': 'AWF machine + Winding Jig size + Stretching Pole',
         'before_value': None, 'after_value': None, 'unit': None, 'machine': 'AWF #3',
         'jig': 'Winding 9.42 / Stretch 5.08', 'material_lot': None, 'supplier': None,
         'dry_time_sec': 8.0, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': ds3_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!A10']},
        {'condition_id': 'cond_4', 'condition_group': 'awf_machine', 'line': '', 'process': 'AWF',
         'changed_factor': 'AWF machine + Winding Jig size + Stretching Pole',
         'before_value': None, 'after_value': None, 'unit': None, 'machine': 'AWF #4',
         'jig': 'Winding 9.42 / Stretch 5.065', 'material_lot': None, 'supplier': None,
         'dry_time_sec': 8.0, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': ds3_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!A11']},
    ],
    'results': [
        # 9/20/2023
        {'result_id': 'awf1_d1', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'AWF #1', 'date': '2023-09-20', 'line': '',
         'input_count': 106, 'ok_count': 79, 'ng_count': 27, 'ng_rate_decimal': 0.255, 'ng_rate_percent': 25.5,
         'metric_name': 'Total NG rate', 'metric_value': 25.5, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG Sigma SPL': {'count': 0, 'rate': 0.0}, 'NG Hearing Noise': {'count': 24, 'rate': 66.7},
                          'NG Hearing Touch': {'count': 3, 'rate': 8.3}, 'HOHD': {'count': 0, 'rate': 0.0}},
         'source_file': ds3_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!_awf1_d1']},
        {'result_id': 'awf2_d1', 'condition_id': 'cond_2', 'measurement_type': 'Function',
         'condition_group': 'AWF #2', 'date': '2023-09-20', 'line': '',
         'input_count': 108, 'ok_count': 76, 'ng_count': 32, 'ng_rate_decimal': 0.296, 'ng_rate_percent': 29.6,
         'metric_name': 'Total NG rate', 'metric_value': 29.6, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG Sigma SPL': {'count': 1, 'rate': 3.1}, 'NG Hearing Noise': {'count': 29, 'rate': 90.6},
                          'NG Hearing Touch': {'count': 2, 'rate': 6.2}, 'HOHD': {'count': 0, 'rate': 0.0}},
         'source_file': ds3_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!_awf2_d1']},
        {'result_id': 'awf3_d1', 'condition_id': 'cond_3', 'measurement_type': 'Function',
         'condition_group': 'AWF #3', 'date': '2023-09-20', 'line': '',
         'input_count': 116, 'ok_count': 80, 'ng_count': 36, 'ng_rate_decimal': 0.31, 'ng_rate_percent': 31.0,
         'metric_name': 'Total NG rate', 'metric_value': 31.0, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG Sigma SPL': {'count': 1, 'rate': 2.8}, 'NG Hearing Noise': {'count': 26, 'rate': 72.2},
                          'NG Hearing Touch': {'count': 9, 'rate': 25.0}, 'HOHD': {'count': 0, 'rate': 0.0}},
         'source_file': ds3_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!_awf3_d1']},
        {'result_id': 'awf4_d1', 'condition_id': 'cond_4', 'measurement_type': 'Function',
         'condition_group': 'AWF #4', 'date': '2023-09-20', 'line': '',
         'input_count': 116, 'ok_count': 68, 'ng_count': 48, 'ng_rate_decimal': 0.414, 'ng_rate_percent': 41.4,
         'metric_name': 'Total NG rate', 'metric_value': 41.4, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG Sigma SPL': {'count': 0, 'rate': 0.0}, 'NG Hearing Noise': {'count': 46, 'rate': 95.8},
                          'NG Hearing Touch': {'count': 2, 'rate': 4.2}, 'HOHD': {'count': 0, 'rate': 0.0}},
         'source_file': ds3_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!_awf4_d1']},
        # 9/21/2023
        {'result_id': 'awf1_d2', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'AWF #1', 'date': '2023-09-21', 'line': '',
         'input_count': 46, 'ok_count': 37, 'ng_count': 9, 'ng_rate_decimal': 0.196, 'ng_rate_percent': 19.6,
         'metric_name': 'Total NG rate', 'metric_value': 19.6, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG Hearing Noise': {'count': 9, 'rate': 25.0}},
         'source_file': ds3_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!_awf1_d2']},
        {'result_id': 'awf2_d2', 'condition_id': 'cond_2', 'measurement_type': 'Function',
         'condition_group': 'AWF #2', 'date': '2023-09-21', 'line': '',
         'input_count': 44, 'ok_count': 34, 'ng_count': 10, 'ng_rate_decimal': 0.227, 'ng_rate_percent': 22.7,
         'metric_name': 'Total NG rate', 'metric_value': 22.7, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG Hearing Noise': {'count': 10, 'rate': 100.0}},
         'source_file': ds3_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!_awf2_d2']},
        {'result_id': 'awf3_d2', 'condition_id': 'cond_3', 'measurement_type': 'Function',
         'condition_group': 'AWF #3', 'date': '2023-09-21', 'line': '',
         'input_count': 48, 'ok_count': 43, 'ng_count': 5, 'ng_rate_decimal': 0.104, 'ng_rate_percent': 10.4,
         'metric_name': 'Total NG rate', 'metric_value': 10.4, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'NG Hearing Noise': {'count': 5, 'rate': 13.9}},
         'source_file': ds3_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!_awf3_d2']},
        {'result_id': 'awf4_d2', 'condition_id': 'cond_4', 'measurement_type': 'Function',
         'condition_group': 'AWF #4', 'date': '2023-09-21', 'line': '',
         'input_count': 46, 'ok_count': 42, 'ng_count': 4, 'ng_rate_decimal': 0.087, 'ng_rate_percent': 8.7,
         'metric_name': 'Total NG rate', 'metric_value': 8.7, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'NG Sigma THD': {'count': 1, 'rate': 25.0}, 'NG Hearing Noise': {'count': 3, 'rate': 75.0}},
         'source_file': ds3_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!_awf4_d2']},
    ],
    'conclusions': [
        {'conclusion_id': 'concl_1', 'topic': 'AWF machine comparison for NG Function',
         'statement_from_report': 'Decision section blank. Tested each AWF #1-#4 with given jig/pole; AWF #3 has Stretching Pole 5.08 (others 5.065) and AWF #1 has Winding Jig 9.34 (others 9.42).',
         'normalized_interpretation': 'No same-event Normal/Baseline row. Day-1 (9/20): #1 25.5%, #2 29.6%, #3 31.0%, #4 41.4% — Noise dominant. Day-2 (9/21): drops markedly across all machines (#1 19.6%, #2 22.7%, #3 10.4%, #4 8.7%) suggesting setup/material variability dominates over machine identity. AWF #1 (Winding 9.34) does not show clear improvement vs Winding 9.42 group; AWF #3 (Pole 5.08) is best on Day-2 but middle on Day-1. Effects are confounded by day.',
         'source_file': ds3_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!A41']}
    ],
    'troubleshooting_index': {
        'defect_name': 'NG Hearing Noise',
        'when_user_asks': ['AWF NG high', 'Hearing Noise', 'Function NG', 'Compare AWF machines'],
        'suggested_checks': [
            {'hint_id': 'hint_1', 'check_item': 'Stabilize day-to-day AWF setup (Winding Jig / Stretching Pole / Cooler / Tension) before drawing per-machine conclusions',
             'reason': 'Day-1 vs Day-2 NG rates differ by 2-3x for the same machine and same nominal jig — indicating day/lot effects dominate. AWF #4 went 41.4% → 8.7%, AWF #3 31.0% → 10.4%. Without a same-event Normal, individual machine ranking is unreliable.',
             'evidence_strength': 'medium', 'related_process': 'AWF', 'related_part': 'Coil/Diaphragm',
             'source_file': ds3_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!_awf_table']}
        ],
        'limitations': ['No same-event Normal/Baseline. Decision section empty. AWF#5 not yet run.']
    },
    'ai_extraction_log': {
        'confidence': 0.7,
        'assumptions': ['Day-1 (9/20) and Day-2 (9/21) treated as separate events because of strong inter-day variation.'],
        'warnings': ['No same-event Normal/Baseline row exists — cannot apply multiplicative relative change against a control.', 'Decision section empty.'],
        'decision_rationale': 'Classified as doe_matrix because four AWF machines with different Winding Jig and Stretching Pole values are compared. No control row. Used per-day cross-machine comparison; flagged the day effect.'
    }
}

ds3_tr_en = {
    'document': {'title': 'BRS-161014 — AWF #1-#4 machine comparison test',
                 'purpose': 'Improve high NG Function by separating AWF machines #1-#4.',
                 'content': ['Test each AWF #1-#4 with assigned Winding Jig and Stretching Pole; AWF#5 not yet run.']},
    'conclusions': {'concl_1': {'topic': 'AWF machine comparison for NG Function',
                                'statement_from_report': 'Tested each AWF #1-#4 with assigned jig/pole. AWF #3 has Pole 5.08 vs 5.065; AWF #1 has Winding Jig 9.34 vs 9.42. Decision section blank.',
                                'normalized_interpretation': 'No same-event Normal exists. Day-1 (9/20): #1 25.5%, #2 29.6%, #3 31.0%, #4 41.4%, Noise dominant. Day-2 (9/21): all machines drop substantially (#1 19.6%, #2 22.7%, #3 10.4%, #4 8.7%). Day/lot effect dominates over machine identity; machine ranking inconsistent across days.'}},
    'hints': {'hint_1': {'check_item': 'Stabilize day-to-day AWF setup before per-machine ranking',
                         'reason': 'Same machine and same nominal jig show 2-3x NG variation between Day-1 and Day-2 (e.g., AWF #4 41.4% → 8.7%). Day/lot effect dominates; without same-event Normal, per-machine ranking is unreliable.'}},
    'log': {'assumptions': ['Day-1 (9/20) and Day-2 (9/21) treated as separate events.'],
            'warnings': ['No same-event Normal/Baseline row. Decision section empty. AWF#5 not yet run.'],
            'decision_rationale': 'doe_matrix because four AWF machines with different jig/pole compared. No control row. Per-day cross-machine comparison used.'}
}

ds3_tr_ko = {
    'document': {'title': 'BRS-161014 — AWF #1-#4 기기 비교 시험',
                 'purpose': 'NG Function 높음을 AWF 기기 #1-#4 분리 시험으로 개선.',
                 'content': ['지정된 Winding Jig 및 Stretching Pole로 AWF #1-#4 각각 시험; AWF#5는 미실시.']},
    'conclusions': {'concl_1': {'topic': 'AWF 기기 NG Function 비교',
                                'statement_from_report': '각 AWF #1-#4 지정 jig/pole로 시험. AWF #3는 Pole 5.08(타 5.065), AWF #1은 Winding Jig 9.34(타 9.42). Decision 비어있음.',
                                'normalized_interpretation': 'Same-event Normal 없음. Day-1(9/20): #1 25.5%, #2 29.6%, #3 31.0%, #4 41.4%, Noise 지배적. Day-2(9/21): 모든 기기 큰 폭 감소(#1 19.6%, #2 22.7%, #3 10.4%, #4 8.7%). 기기보다 일자/lot 효과가 지배적이며 기기 순위는 일자별 일관성 없음.'}},
    'hints': {'hint_1': {'check_item': '기기별 순위 매기기 전 AWF 일자간 셋업 안정화',
                         'reason': '동일 기기·동일 jig에서 Day-1과 Day-2 NG가 2-3배 차이(AWF #4 41.4% → 8.7%). 일자/lot 효과 지배; same-event Normal 없으면 기기 순위 신뢰 불가.'}},
    'log': {'assumptions': ['Day-1(9/20)과 Day-2(9/21)를 별도 이벤트로 처리.'],
            'warnings': ['Same-event Normal row 없음. Decision 비어있음. AWF#5 미실시.'],
            'decision_rationale': 'jig/pole 다른 4대 AWF 비교이므로 doe_matrix. Control row 없음. 일자별 기기간 비교 사용.'}
}

ds3_tr_vi = {
    'document': {'title': 'BRS-161014 — So sánh máy AWF #1-#4',
                 'purpose': 'Cải thiện NG Function cao bằng cách tách máy AWF #1-#4.',
                 'content': ['Thử mỗi AWF #1-#4 với Winding Jig và Stretching Pole đã gán; AWF#5 chưa chạy.']},
    'conclusions': {'concl_1': {'topic': 'So sánh máy AWF cho NG Function',
                                'statement_from_report': 'Thử mỗi AWF #1-#4. AWF #3 có Pole 5.08 (#còn lại 5.065); AWF #1 có Winding Jig 9.34 (#còn lại 9.42). Phần Decision để trống.',
                                'normalized_interpretation': 'Không có Normal same-event. Day-1 (9/20): #1 25.5%, #2 29.6%, #3 31.0%, #4 41.4%, Noise chiếm ưu thế. Day-2 (9/21): tất cả máy giảm mạnh (#1 19.6%, #2 22.7%, #3 10.4%, #4 8.7%). Hiệu ứng ngày/lot mạnh hơn máy; xếp hạng máy không nhất quán giữa các ngày.'}},
    'hints': {'hint_1': {'check_item': 'Ổn định setup AWF giữa các ngày trước khi xếp hạng máy',
                         'reason': 'Cùng máy và cùng jig nhưng NG chênh 2-3x giữa Day-1 và Day-2 (AWF #4 41.4% → 8.7%). Hiệu ứng ngày/lot vượt trội; không có Normal same-event thì xếp hạng máy không đáng tin.'}},
    'log': {'assumptions': ['Day-1 (9/20) và Day-2 (9/21) coi như sự kiện riêng.'],
            'warnings': ['Không có dòng Normal/Baseline same-event. Decision trống. AWF#5 chưa chạy.'],
            'decision_rationale': 'doe_matrix vì so sánh 4 AWF với jig/pole khác nhau. Không có dòng control. Dùng so sánh giữa các máy theo từng ngày.'}
}


# =====================================================================
# DATASET 4: 33. BRS-161014 DT Report test Suspension NG dimension - date 2024.05.02
# normal_comparison: SP NG dimension vs Normal
# =====================================================================
ds4_name = '33. BRS-161014 DT Report test Suspension NG dimension -  date 2024.05.02'

ds4_result = {
    'schema_version': '0.1',
    'document': {
        'document_id': 'doc_4',
        'source_file': ds4_name,
        'source_sheet': 'Test',
        'title': 'BRS-161016 DT — REPORT TEST SUSPENSION NG DIMENSION LASER CUTTING BY VENDOR NANOSYS',
        'model': 'BRS-161014 DT',
        'report_date': '2024-05-02',
        'department': 'ME',
        'marker': 'Le',
        'line': 'C2-3A',
        'report_type': 'normal_comparison',
        'primary_defect': {'canonical_name': 'Dimension NG',
                           'aliases_in_document': ['Suspension NG dimension', 'NG Suspension Gap']},
        'related_defects': ['NG Hearing Noise', 'NG Hearing Touch'],
        'parts': ['Suspension (SP)', 'Frame (Fr)'],
        'processes': ["Ass'y Fr+SP", 'Function'],
        'purpose': 'Verify if Suspension with NG dimension (laser cutting, vendor Nanosys) can be used.',
        'content': ['Standard 5.17-5.27 (5.22 nom); actual spec 5.15-5.22. Make semi and check Ass\'y Fr+SP NG rate; make final and check function NG; compare with Normal.'],
        'source_cells': {'title': ['Test!B1'], 'date': ['Test!M2'], 'purpose': ['Test!A4'], 'content': ['Test!A6']}
    },
    'test_conditions': [
        {'condition_id': 'cond_1', 'condition_group': 'sp_dimension', 'line': 'C2-3A', 'process': "Ass'y Fr+SP",
         'changed_factor': 'Use Suspension with NG dimension (Nanosys laser cutting)', 'before_value': 'Spec 5.17-5.27',
         'after_value': 'Actual 5.15-5.22', 'unit': 'mm', 'machine': None, 'jig': None,
         'material_lot': 'Nanosys NG dim lot', 'supplier': 'Nanosys', 'dry_time_sec': None, 'temperature': None,
         'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': ds4_name, 'sheet_name': 'Test', 'source_cells': ['Test!A5', 'Test!A6']}
    ],
    'results': [
        {'result_id': 'assy_test', 'condition_id': 'cond_1', 'measurement_type': "Ass'y Fr+SP",
         'condition_group': 'Test SP NG dimension', 'date': '2024-05-02', 'line': 'C2-3A',
         'input_count': 384, 'ok_count': 384, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Suspension Gap NG Rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'NG Suspension Gap': {'count': 0, 'rate': 0.0}},
         'source_file': ds4_name, 'sheet_name': 'Test', 'source_cells': ['Test!assy_test']},
        {'result_id': 'assy_normal', 'condition_id': None, 'measurement_type': "Ass'y Fr+SP",
         'condition_group': 'Normal', 'date': '2024-05-02', 'line': 'C2-3A',
         'input_count': 500, 'ok_count': 500, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Suspension Gap NG Rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'NG Suspension Gap': {'count': 0, 'rate': 0.0}},
         'source_file': ds4_name, 'sheet_name': 'Test', 'source_cells': ['Test!assy_normal']},
        {'result_id': 'func_test', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Test SP NG dimension', 'date': '2024-05-02', 'line': 'C2-3A',
         'input_count': 384, 'ok_count': 374, 'ng_count': 10, 'ng_rate_decimal': 0.026, 'ng_rate_percent': 2.6,
         'metric_name': 'Function NG Rate', 'metric_value': 2.6, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'NG Sigma SPL': {'count': 0, 'rate': 0.0}, 'NG Sigma SPL+THD': {'count': 0, 'rate': 0.0},
                          'NG Sigma SPL+THD+F0': {'count': 0, 'rate': 0.0},
                          'NG Hearing Noise': {'count': 10, 'rate': 2.6}, 'NG Hearing Touch': {'count': 0, 'rate': 0.0}},
         'source_file': ds4_name, 'sheet_name': 'Test', 'source_cells': ['Test!func_test']},
        {'result_id': 'func_normal', 'condition_id': None, 'measurement_type': 'Function',
         'condition_group': 'Normal', 'date': '2024-05-02', 'line': 'C2-3A',
         'input_count': 560, 'ok_count': 547, 'ng_count': 13, 'ng_rate_decimal': 0.0232, 'ng_rate_percent': 2.32,
         'metric_name': 'Function NG Rate', 'metric_value': 2.32, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'NG Sigma SPL': {'count': 0, 'rate': 0.0}, 'NG Sigma SPL+THD': {'count': 0, 'rate': 0.0},
                          'NG Sigma SPL+THD+F0': {'count': 0, 'rate': 0.0},
                          'NG Hearing Noise': {'count': 11, 'rate': 2.0}, 'NG Hearing Touch': {'count': 2, 'rate': 0.4}},
         'source_file': ds4_name, 'sheet_name': 'Test', 'source_cells': ['Test!func_normal']},
    ],
    'conclusions': [
        {'conclusion_id': 'concl_1', 'topic': 'Suspension NG dimension usability',
         'statement_from_report': "Result check process Ass'y Fr+SP all OK same Normal. Function test SP NG dimension NG rate 10/384 ~ 2.6% same normal => Can use it.",
         'normalized_interpretation': "Ass'y Fr+SP: Test 0.00% vs Normal 0.00% — equal. Function: Test 2.60% vs Normal 2.32% = 1.121x, 12.1% worse than same-event Normal. Difference is small (~0.28 percentage points) and well within typical noise; workbook decision is approve-for-use.",
         'source_file': ds4_name, 'sheet_name': 'Test', 'source_cells': ['Test!_decision']}
    ],
    'troubleshooting_index': {
        'defect_name': 'Dimension NG',
        'when_user_asks': ['SP/Suspension NG dimension', 'Laser cutting NG', 'Nanosys vendor', 'Dimension out of spec'],
        'suggested_checks': [
            {'hint_id': 'hint_1', 'check_item': 'Run both Ass\'y Fr+SP NG rate and Function NG rate vs same-event Normal when accepting out-of-spec laser-cut suspensions',
             'reason': "Test showed Ass'y 0% (same as Normal) and Function 2.6% vs 2.32% Normal (12% worse but acceptable). The two-stage check (Ass'y + Function) is the evidence basis for accept/reject.",
             'evidence_strength': 'medium', 'related_process': "Ass'y Fr+SP, Function",
             'related_part': 'Suspension',
             'source_file': ds4_name, 'sheet_name': 'Test', 'source_cells': ['Test!_decision']}
        ],
        'limitations': ['Actual spec 5.15-5.22 narrower than standard 5.17-5.27; sample is only one lot from Nanosys.']
    },
    'ai_extraction_log': {
        'confidence': 0.85,
        'assumptions': ['Title says BRS-161016 DT but filename and dataset use BRS-161014 DT; preserved both.'],
        'warnings': ['Function NG slightly higher than Normal (12.1% relative); acceptable given absolute level is low.'],
        'decision_rationale': "Same-event Normal exists for both Ass'y Fr+SP and Function. Multiplicative relative change used. Ass'y 0% vs 0% identical. Function 2.6% vs 2.32% = 1.121x (12.1% worse) — small enough to approve per workbook decision."
    }
}

ds4_tr_en = {
    'document': {'title': 'BRS-161014 DT — Test Suspension NG dimension (Nanosys laser cutting)',
                 'purpose': 'Verify whether Suspension with NG dimension can be used.',
                 'content': ['Standard 5.17-5.27 (nom 5.22); actual 5.15-5.22. Check Ass\'y Fr+SP and Function NG vs Normal.']},
    'conclusions': {'concl_1': {'topic': 'Suspension NG dimension usability',
                                'statement_from_report': 'Ass\'y Fr+SP all OK same Normal. Function 10/384 = 2.6% same Normal => Can use it.',
                                'normalized_interpretation': 'Ass\'y Fr+SP Test 0.00% vs Normal 0.00% — equal. Function Test 2.60% vs Normal 2.32% = 1.121x, 12.1% worse. Small absolute gap (~0.28 pp); workbook approves.'}},
    'hints': {'hint_1': {'check_item': 'Run both Ass\'y and Function vs same-event Normal when accepting out-of-spec laser-cut suspensions',
                         'reason': "Test showed Ass'y 0% (same as Normal) and Function 2.6% vs 2.32% (12% worse, acceptable). Two-stage check is the accept/reject basis."}},
    'log': {'assumptions': ['Title says BRS-161016 DT but dataset/filename uses BRS-161014 DT; both preserved.'],
            'warnings': ['Function NG slightly higher than Normal (12.1% relative); acceptable given low absolute.'],
            'decision_rationale': 'Same-event Normal exists for both metrics. Multiplicative change used. Ass\'y equal, Function +12.1% — small, workbook approves.'}
}

ds4_tr_ko = {
    'document': {'title': 'BRS-161014 DT — Suspension NG dimension(Nanosys 레이저 컷팅) 시험',
                 'purpose': 'NG dimension Suspension 사용 가능 여부 검증.',
                 'content': ['Standard 5.17-5.27 (5.22 nom); actual 5.15-5.22. Ass\'y Fr+SP 및 Function NG를 Normal과 비교.']},
    'conclusions': {'concl_1': {'topic': 'Suspension NG dimension 사용 가능성',
                                'statement_from_report': 'Ass\'y Fr+SP 전부 OK, Normal과 동일. Function 10/384 = 2.6% Normal과 동일 => 사용 가능.',
                                'normalized_interpretation': 'Ass\'y Fr+SP Test 0.00% vs Normal 0.00% — 동일. Function Test 2.60% vs Normal 2.32% = 1.121배, 12.1% 악화. 절대 차 작음(~0.28 pp); 리포트 승인.'}},
    'hints': {'hint_1': {'check_item': '규격 이탈 레이저컷 Suspension 수용 시 Ass\'y와 Function 모두 same-event Normal 대비 확인',
                         'reason': "Ass'y 0%(Normal 동일), Function 2.6% vs 2.32%(12% 악화, 수용 가능). 2단계 검사가 수용/불수용 근거."}},
    'log': {'assumptions': ['제목은 BRS-161016 DT이지만 dataset/filename은 BRS-161014 DT; 둘 다 보존.'],
            'warnings': ['Function NG가 Normal보다 약간 높음(12.1% 상대); 절대 수준 낮아 수용 가능.'],
            'decision_rationale': '두 지표 모두 same-event Normal 존재. 곱셈 변화율 사용. Ass\'y 동일, Function +12.1% — 작음, 리포트 승인.'}
}

ds4_tr_vi = {
    'document': {'title': 'BRS-161014 DT — Test Suspension NG dimension (Nanosys laser cutting)',
                 'purpose': 'Kiểm tra Suspension NG dimension có dùng được không.',
                 'content': ['Standard 5.17-5.27 (nom 5.22); actual 5.15-5.22. So sánh Ass\'y Fr+SP và Function NG với Normal.']},
    'conclusions': {'concl_1': {'topic': 'Khả năng sử dụng Suspension NG dimension',
                                'statement_from_report': 'Ass\'y Fr+SP OK same Normal. Function 10/384 = 2.6% same Normal => Dùng được.',
                                'normalized_interpretation': 'Ass\'y Fr+SP Test 0.00% vs Normal 0.00% — bằng nhau. Function Test 2.60% vs Normal 2.32% = 1.121x, xấu hơn 12.1%. Chênh tuyệt đối nhỏ (~0.28 pp); báo cáo chấp nhận.'}},
    'hints': {'hint_1': {'check_item': 'Chạy cả Ass\'y và Function so với Normal same-event khi chấp nhận Suspension cắt laser ngoài spec',
                         'reason': "Ass'y 0% (như Normal) và Function 2.6% vs 2.32% (xấu hơn 12%, chấp nhận được). Kiểm tra hai bước là cơ sở chấp nhận/từ chối."}},
    'log': {'assumptions': ['Tiêu đề ghi BRS-161016 DT nhưng dataset/filename là BRS-161014 DT; lưu cả hai.'],
            'warnings': ['Function NG cao hơn Normal chút (12.1% tương đối); chấp nhận do mức tuyệt đối thấp.'],
            'decision_rationale': 'Hai chỉ số đều có Normal same-event. Dùng tỷ lệ thay đổi nhân. Ass\'y bằng, Function +12.1% — nhỏ, báo cáo duyệt.'}
}


# =====================================================================
# DATASET 5: 33. BRS-161014 Report TEST VP NG burr, deform 20.9.2023
# normal_comparison: VP burr / VP deform vs Normal
# =====================================================================
ds5_name = '33. BRS-161014 Report TEST VP NG burr, deform 20.9.2023'

ds5_result = {
    'schema_version': '0.1',
    'document': {
        'document_id': 'doc_5',
        'source_file': ds5_name,
        'source_sheet': 'Report (2)',
        'title': 'REPORT TEST VP NG BURR, DEFORM — BRS-161016/161014',
        'model': 'BRS-161014 / 161016',
        'report_date': '2023-09-21',
        'department': 'ME',
        'marker': 'Thuy',
        'line': '',
        'report_type': 'normal_comparison',
        'primary_defect': {'canonical_name': 'Burr NG',
                           'aliases_in_document': ['VP burr', 'VP NG burn', 'VP NG burr']},
        'related_defects': ['Deform NG', 'NG Hearing Noise', 'Insert Jig NG (VP float)'],
        'parts': ['VP'],
        'processes': ['Sub 1 forming', 'VP+CD Ass\'y', 'Function'],
        'purpose': 'Check if VP with NG burr / NG deform can be used.',
        'content': ["Use VP NG burr/deform to make samples at Sub 1 (check Ass'y VP+CD), then move to Main 2 final (check function)."],
        'source_cells': {'title': ['Report (2)!B1'], 'date': ['Report (2)!N2'], 'purpose': ['Report (2)!A4'],
                         'content': ['Report (2)!A6']}
    },
    'test_conditions': [
        {'condition_id': 'cond_1', 'condition_group': 'vp_burr', 'line': '', 'process': 'Sub 1 / Final',
         'changed_factor': 'Use VP with NG burr', 'before_value': 'Normal VP',
         'after_value': 'VP burr', 'unit': None, 'machine': None, 'jig': None,
         'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None,
         'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': ds5_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!A6']},
        {'condition_id': 'cond_2', 'condition_group': 'vp_deform', 'line': '', 'process': 'Sub 1 / Final',
         'changed_factor': 'Use VP with NG deform', 'before_value': 'Normal VP',
         'after_value': 'VP deform', 'unit': None, 'machine': None, 'jig': None,
         'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None,
         'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': ds5_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!A7']},
    ],
    'results': [
        {'result_id': 'sub1_burr', 'condition_id': 'cond_1', 'measurement_type': "Sub 1 Ass'y VP+CD",
         'condition_group': 'VP burr', 'date': '2023-09-20', 'line': '',
         'input_count': 30, 'ok_count': 30, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Sub 1 NG Rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'Cutting NG': {'count': 0, 'rate': 0.0}, 'Insert Jig NG': {'count': 0, 'rate': 0.0},
                          'VP+CD Separation': {'count': 0, 'rate': 0.0}},
         'source_file': ds5_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!_sub1_burr']},
        {'result_id': 'sub1_deform', 'condition_id': 'cond_2', 'measurement_type': "Sub 1 Ass'y VP+CD",
         'condition_group': 'VP deform', 'date': '2023-09-20', 'line': '',
         'input_count': 30, 'ok_count': 0, 'ng_count': 30, 'ng_rate_decimal': 1.0, 'ng_rate_percent': 100.0,
         'metric_name': 'Sub 1 NG Rate', 'metric_value': 100.0, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'Cutting NG': {'count': 0, 'rate': 0.0}, 'Insert Jig NG (VP float)': {'count': 30, 'rate': 100.0},
                          'VP+CD Separation': {'count': 0, 'rate': 0.0}},
         'source_file': ds5_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!_sub1_deform']},
        {'result_id': 'func_burr', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'VP burr', 'date': '2023-09-19', 'line': '',
         'input_count': 30, 'ok_count': 24, 'ng_count': 6, 'ng_rate_decimal': 0.20, 'ng_rate_percent': 20.0,
         'metric_name': 'Function NG Rate', 'metric_value': 20.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'NG Sigma SPL': {'count': 0, 'rate': 0.0}, 'NG Hearing Noise': {'count': 6, 'rate': 100.0},
                          'NG Hearing Touch': {'count': 0, 'rate': 0.0}, 'HOHD': {'count': 0, 'rate': 0.0}},
         'source_file': ds5_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!_func_burr']},
        {'result_id': 'func_deform', 'condition_id': 'cond_2', 'measurement_type': 'Function',
         'condition_group': 'VP deform', 'date': '2023-09-19', 'line': '',
         'input_count': 29, 'ok_count': 23, 'ng_count': 6, 'ng_rate_decimal': 0.207, 'ng_rate_percent': 20.7,
         'metric_name': 'Function NG Rate', 'metric_value': 20.7, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'NG Sigma SPL': {'count': 0, 'rate': 0.0}, 'NG Hearing Noise': {'count': 5, 'rate': 83.3},
                          'NG Hearing Touch': {'count': 1, 'rate': 16.7}, 'HOHD': {'count': 0, 'rate': 0.0}},
         'source_file': ds5_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!_func_deform']},
        {'result_id': 'func_normal', 'condition_id': None, 'measurement_type': 'Function',
         'condition_group': 'Normal', 'date': '2023-09-19', 'line': '',
         'input_count': 100, 'ok_count': 78, 'ng_count': 22, 'ng_rate_decimal': 0.22, 'ng_rate_percent': 22.0,
         'metric_name': 'Function NG Rate', 'metric_value': 22.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'NG Sigma SPL': {'count': 0, 'rate': 0.0}, 'NG Hearing Noise': {'count': 20, 'rate': 90.9},
                          'NG Hearing Touch': {'count': 2, 'rate': 9.1}, 'HOHD': {'count': 0, 'rate': 0.0}},
         'source_file': ds5_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!_func_normal']},
    ],
    'conclusions': [
        {'conclusion_id': 'concl_1', 'topic': 'VP NG burr / NG deform usability',
         'statement_from_report': 'VP NG burr have NG rate same Normal line — can use. VP NG deform: function same Normal but at Sub 1 100% Insert Jig NG (VP float) — can use VP deform but need supplier improvement.',
         'normalized_interpretation': 'Function VP burr 20.0% vs Normal 22.0% = 0.909x, 9.1% improved (statistically negligible). VP deform 20.7% vs Normal 22.0% = 0.941x, 5.9% improved. However Sub 1 Ass\'y VP+CD: VP deform 100% Insert Jig NG (VP float) vs VP burr 0% — VP deform fails at jig insertion. So functional output, when made, is fine; but VP deform is unusable on current jig without supplier fix.',
         'source_file': ds5_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!_decision']}
    ],
    'troubleshooting_index': {
        'defect_name': 'Burr NG / Deform NG',
        'when_user_asks': ['VP burr', 'VP deform', 'VP float in jig', 'Use NG material decision'],
        'suggested_checks': [
            {'hint_id': 'hint_1', 'check_item': 'When considering NG-burr or NG-deform VP for use, check Sub 1 Insert-Jig float rate before relying on final Function comparison',
             'reason': 'VP deform passed Function (20.7% vs Normal 22.0%) but failed at Sub 1 with 100% insert-jig float (30/30). Function-only comparison would have falsely cleared an unusable material.',
             'evidence_strength': 'high', 'related_process': 'Sub 1 forming', 'related_part': 'VP',
             'source_file': ds5_name, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!_sub1_deform']}
        ],
        'limitations': ['No same-event Normal at Sub 1 (only burr vs deform within test). Function sample size 30 each.']
    },
    'ai_extraction_log': {
        'confidence': 0.85,
        'assumptions': ['"VP burr"/"VP burn" treated as the same canonical Burr NG defect.'],
        'warnings': ['Sub 1 table has no Normal baseline row — only burr/deform vs each other.'],
        'decision_rationale': 'Function table has same-event Normal — multiplicative change applied (both NG types ~5-9% improved, but within noise). Sub 1 must be flagged separately because VP deform shows catastrophic 100% jig-float that Function alone would miss.'
    }
}

ds5_tr_en = {
    'document': {'title': 'BRS-161014 / 161016 — Test VP NG burr / NG deform',
                 'purpose': 'Check whether VP with NG burr / NG deform can be used.',
                 'content': ['Sub 1: make samples and check Ass\'y VP+CD. Main 2 final: check function. Compare with Normal.']},
    'conclusions': {'concl_1': {'topic': 'VP NG burr / NG deform usability',
                                'statement_from_report': 'VP NG burr NG rate same Normal — can use. VP NG deform: function same Normal but Sub 1 100% Insert Jig NG (VP float) — can use but supplier must improve.',
                                'normalized_interpretation': 'Function VP burr 20.0% vs Normal 22.0% = 0.909x (9.1% improved). VP deform 20.7% vs Normal 22.0% = 0.941x (5.9% improved). Both within noise. Critical: Sub 1 Ass\'y VP+CD shows VP deform 100% Insert Jig NG (VP float) vs VP burr 0%. Function comparison alone would miss this jig-incompatibility.'}},
    'hints': {'hint_1': {'check_item': 'For NG-burr/NG-deform VP usage decision, include Sub 1 insert-jig float check, not just final Function',
                         'reason': 'VP deform passed Function (20.7% vs Normal 22.0%) but failed Sub 1 with 100% insert-jig float (30/30). Function-only would have falsely cleared an unusable material.'}},
    'log': {'assumptions': ['"VP burr"/"VP burn" treated as same canonical Burr NG.'],
            'warnings': ['Sub 1 has no Normal baseline — only burr/deform compared with each other.'],
            'decision_rationale': 'Function has same-event Normal; multiplicative change ~5-9% improved (within noise). Sub 1 flagged separately because 100% jig float invalidates VP deform use.'}
}

ds5_tr_ko = {
    'document': {'title': 'BRS-161014 / 161016 — VP NG burr / NG deform 시험',
                 'purpose': 'NG burr / NG deform VP 사용 가능 여부 검증.',
                 'content': ['Sub 1: 샘플 제작 후 Ass\'y VP+CD 검사. Main 2 final: function 검사. Normal과 비교.']},
    'conclusions': {'concl_1': {'topic': 'VP NG burr / NG deform 사용 가능성',
                                'statement_from_report': 'VP NG burr는 NG rate가 Normal과 동일 — 사용 가능. VP NG deform: function은 Normal과 동일하지만 Sub 1에서 Insert Jig NG (VP float) 100% — 사용 가능하나 supplier 개선 필요.',
                                'normalized_interpretation': 'Function VP burr 20.0% vs Normal 22.0% = 0.909배(9.1% 개선). VP deform 20.7% vs Normal 22.0% = 0.941배(5.9% 개선). 둘 다 노이즈 범위. 핵심: Sub 1 Ass\'y VP+CD에서 VP deform 100% Insert Jig NG (VP float) vs VP burr 0%. Function만으로는 이 jig 부적합 발견 불가.'}},
    'hints': {'hint_1': {'check_item': 'NG-burr/NG-deform VP 수용 결정 시 final Function만이 아니라 Sub 1 insert-jig float 검사 포함',
                         'reason': 'VP deform Function 통과(20.7% vs Normal 22.0%)했지만 Sub 1에서 100% insert-jig float (30/30) 실패. Function만 보면 사용 불가 자재를 잘못 승인할 수 있었음.'}},
    'log': {'assumptions': ['"VP burr"/"VP burn"은 동일 표준 Burr NG로 처리.'],
            'warnings': ['Sub 1에는 Normal baseline 없음 — burr/deform 상호 비교만 가능.'],
            'decision_rationale': 'Function에는 same-event Normal 존재; 곱셈 변화율 ~5-9% 개선(노이즈 범위). Sub 1은 별도 표시: VP deform 100% jig float은 사용 불가.'}
}

ds5_tr_vi = {
    'document': {'title': 'BRS-161014 / 161016 — Test VP NG burr / NG deform',
                 'purpose': 'Kiểm tra VP NG burr / NG deform có dùng được không.',
                 'content': ['Sub 1: làm mẫu và kiểm Ass\'y VP+CD. Main 2 final: kiểm function. So với Normal.']},
    'conclusions': {'concl_1': {'topic': 'Khả năng dùng VP NG burr / NG deform',
                                'statement_from_report': 'VP NG burr NG rate same Normal — dùng được. VP NG deform: function same Normal nhưng Sub 1 100% Insert Jig NG (VP float) — dùng được nhưng supplier phải cải thiện.',
                                'normalized_interpretation': 'Function VP burr 20.0% vs Normal 22.0% = 0.909x (cải thiện 9.1%). VP deform 20.7% vs Normal 22.0% = 0.941x (cải thiện 5.9%). Cả hai trong khoảng nhiễu. Quan trọng: Sub 1 Ass\'y VP+CD VP deform 100% Insert Jig NG (VP float) vs VP burr 0%. Chỉ so Function sẽ bỏ sót lỗi jig.'}},
    'hints': {'hint_1': {'check_item': 'Quyết định dùng VP NG-burr/NG-deform phải bao gồm kiểm insert-jig float ở Sub 1, không chỉ Function cuối',
                         'reason': 'VP deform đạt Function (20.7% vs Normal 22.0%) nhưng fail Sub 1 với 100% insert-jig float (30/30). Chỉ xét Function sẽ duyệt nhầm vật liệu không dùng được.'}},
    'log': {'assumptions': ['"VP burr"/"VP burn" coi là cùng Burr NG chuẩn.'],
            'warnings': ['Sub 1 không có baseline Normal — chỉ so burr/deform.'],
            'decision_rationale': 'Function có Normal same-event; thay đổi nhân ~5-9% cải thiện (trong nhiễu). Sub 1 đánh dấu riêng: 100% jig float khiến VP deform không dùng được.'}
}


# =====================================================================
# DATASET 6: 33. BRS-161016 Report test material YOKE 161016 occurred dirty plating after clean Ethanol- 24.4.2024
# reliability_spec + mixed: Tension test (spec >=50kgf) + Final SPK visual NG
# =====================================================================
ds6_name = '33. BRS-161016 Report test material YOKE 161016 occurred dirty plating after clean Ethanol- 24.4.2024'

ds6_result = {
    'schema_version': '0.1',
    'document': {
        'document_id': 'doc_6',
        'source_file': ds6_name,
        'source_sheet': '5.3',
        'title': 'REPORT TEST MATERIAL YOKE OCCURRED DIRTY PLATING AFTER CLEAN BY ETHANOL — BRS-161016',
        'model': 'BRS-161016',
        'report_date': '2024-04-24',
        'department': 'ME',
        'marker': 'Thuy',
        'line': '',
        'report_type': 'mixed',
        'primary_defect': {'canonical_name': 'Dirty Plating (YOKE)',
                           'aliases_in_document': ['dirty plating', 'NG visual SPK']},
        'related_defects': ['NG Visual SPK'],
        'parts': ['YOKE'],
        'processes': ['Tension test CMG', 'Final SPK marking', 'Final SPK visual'],
        'purpose': 'YOKE 161016 dirty plating after Ethanol clean — verify material to approve a temporary limit.',
        'content': ['Tension test CMG (spec >= 50kgf): 30 pcs Test YOKE + 30 pcs Normal YOKE.',
                    'Make final SPK and check marking + visual: 20 pcs (actually 18 measured visual).'],
        'source_cells': {'title': ['5.3!B1'], 'date': ['5.3!L2'], 'purpose': ['5.3!A4'], 'content': ['5.3!A6']}
    },
    'test_conditions': [
        {'condition_id': 'cond_1', 'condition_group': 'yoke_dirty_plating', 'line': '', 'process': 'YOKE clean',
         'changed_factor': 'YOKE with dirty plating after Ethanol clean', 'before_value': 'Normal YOKE',
         'after_value': 'Dirty-plating YOKE (Ethanol cleaned)', 'unit': None, 'machine': None, 'jig': None,
         'material_lot': 'YOKE 161016 dirty plating lot', 'supplier': None, 'dry_time_sec': None, 'temperature': None,
         'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': ds6_name, 'sheet_name': '5.3', 'source_cells': ['5.3!A4']}
    ],
    'results': [
        {'result_id': 'tension_test', 'condition_id': 'cond_1', 'measurement_type': 'Tension test CMG',
         'condition_group': 'Test YOKE (dirty plating)', 'date': '2024-04-24', 'line': '',
         'input_count': 30, 'ok_count': 30, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Tension avg', 'metric_value': 128.41, 'unit': 'kgf', 'judgement': 'PASS',
         'ng_breakdown': {},
         'source_file': ds6_name, 'sheet_name': '5.3', 'source_cells': ['5.3!_tension_test_avg']},
        {'result_id': 'tension_test_min', 'condition_id': 'cond_1', 'measurement_type': 'Tension test CMG',
         'condition_group': 'Test YOKE (dirty plating)', 'date': '2024-04-24', 'line': '',
         'input_count': None, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Tension min', 'metric_value': 106.10, 'unit': 'kgf', 'judgement': 'PASS',
         'ng_breakdown': {},
         'source_file': ds6_name, 'sheet_name': '5.3', 'source_cells': ['5.3!_tension_test_min']},
        {'result_id': 'tension_test_max', 'condition_id': 'cond_1', 'measurement_type': 'Tension test CMG',
         'condition_group': 'Test YOKE (dirty plating)', 'date': '2024-04-24', 'line': '',
         'input_count': None, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Tension max', 'metric_value': 152.70, 'unit': 'kgf', 'judgement': 'PASS',
         'ng_breakdown': {},
         'source_file': ds6_name, 'sheet_name': '5.3', 'source_cells': ['5.3!_tension_test_max']},
        {'result_id': 'tension_normal', 'condition_id': None, 'measurement_type': 'Tension test CMG',
         'condition_group': 'Normal YOKE', 'date': '2024-04-24', 'line': '',
         'input_count': 30, 'ok_count': 30, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Tension avg', 'metric_value': 105.56, 'unit': 'kgf', 'judgement': 'PASS',
         'ng_breakdown': {},
         'source_file': ds6_name, 'sheet_name': '5.3', 'source_cells': ['5.3!_tension_normal_avg']},
        {'result_id': 'tension_normal_min', 'condition_id': None, 'measurement_type': 'Tension test CMG',
         'condition_group': 'Normal YOKE', 'date': '2024-04-24', 'line': '',
         'input_count': None, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Tension min', 'metric_value': 93.60, 'unit': 'kgf', 'judgement': 'PASS',
         'ng_breakdown': {},
         'source_file': ds6_name, 'sheet_name': '5.3', 'source_cells': ['5.3!_tension_normal_min']},
        {'result_id': 'tension_normal_max', 'condition_id': None, 'measurement_type': 'Tension test CMG',
         'condition_group': 'Normal YOKE', 'date': '2024-04-24', 'line': '',
         'input_count': None, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Tension max', 'metric_value': 111.70, 'unit': 'kgf', 'judgement': 'PASS',
         'ng_breakdown': {},
         'source_file': ds6_name, 'sheet_name': '5.3', 'source_cells': ['5.3!_tension_normal_max']},
        {'result_id': 'spk_visual', 'condition_id': 'cond_1', 'measurement_type': 'SPK Visual',
         'condition_group': 'Test YOKE (dirty plating)', 'date': '2024-04-24', 'line': '',
         'input_count': 18, 'ok_count': 11, 'ng_count': 7, 'ng_rate_decimal': 0.389, 'ng_rate_percent': 38.9,
         'metric_name': 'SPK Visual NG Rate', 'metric_value': 38.9, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG Visual SPK': {'count': 7, 'rate': 38.9}},
         'source_file': ds6_name, 'sheet_name': '5.3', 'source_cells': ['5.3!_visual_test']},
        {'result_id': 'spk_marking', 'condition_id': 'cond_1', 'measurement_type': 'SPK Marking',
         'condition_group': 'Test YOKE (dirty plating)', 'date': '2024-04-24', 'line': '',
         'input_count': None, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Marking visible', 'metric_value': None, 'unit': None, 'judgement': 'PASS',
         'ng_breakdown': {},
         'source_file': ds6_name, 'sheet_name': '5.3', 'source_cells': ['5.3!_marking']},
    ],
    'conclusions': [
        {'conclusion_id': 'concl_1', 'topic': 'YOKE dirty plating material disposition',
         'statement_from_report': 'Tension test OK, marking OK, but NG visual SPK high (7/18, 38.9%).',
         'normalized_interpretation': "Tension PASS spec >= 50kgf: Test YOKE avg 128.41 kgf (range 106.10-152.70) vs Normal YOKE avg 105.56 kgf (range 93.60-111.70) — Test is actually 21.6% higher than Normal in mean tension. SPK Visual: Test 7/18 = 38.9% NG vs no Normal control row at this stage — high absolute rate that gates the dirty-plating YOKE.",
         'source_file': ds6_name, 'sheet_name': '5.3', 'source_cells': ['5.3!_decide']}
    ],
    'troubleshooting_index': {
        'defect_name': 'Dirty Plating (YOKE)',
        'when_user_asks': ['YOKE dirty plating', 'Ethanol clean', 'SPK visual NG', 'YOKE material disposition'],
        'suggested_checks': [
            {'hint_id': 'hint_1', 'check_item': 'For YOKE with dirty plating, run both Tension (spec gate) and final SPK Visual NG rate; do not approve on Tension alone',
             'reason': 'Tension test passed comfortably (Test avg 128.41 kgf > 50 kgf spec) and was even higher than Normal, but final SPK Visual NG hit 38.9% (7/18). Mechanical pass does not predict visual outcome.',
             'evidence_strength': 'high', 'related_process': 'Final SPK Visual',
             'related_part': 'YOKE',
             'source_file': ds6_name, 'sheet_name': '5.3', 'source_cells': ['5.3!_visual_test']}
        ],
        'limitations': ['SPK Visual has no same-event Normal row. Sample size 18 small. Marking is qualitative only.']
    },
    'ai_extraction_log': {
        'confidence': 0.85,
        'assumptions': ['Tension spec >=50kgf (sheet text: "50kgf trở lên") taken as PASS gate.', 'Visual NG count 7/18 used; the introduction mentioned 20 pcs planned.'],
        'warnings': ['SPK Visual table lacks a same-event Normal row; absolute 38.9% interpreted as fail signal.'],
        'decision_rationale': 'Classified as mixed because Tension (reliability spec) and SPK Visual (NG-rate) are both co-primary evidence. Tension comparison uses same-event Normal (Test +21.6% higher mean). Visual interpreted by absolute NG rate against the implicit zero-NG expectation.'
    }
}

ds6_tr_en = {
    'document': {'title': 'BRS-161016 — YOKE dirty plating after Ethanol clean material test',
                 'purpose': 'YOKE 161016 dirty plating after Ethanol clean — verify whether material can be approved temporarily.',
                 'content': ['Tension test CMG (spec >=50kgf): 30 pcs each Test and Normal YOKE. SPK final marking + visual: 18 pcs visual.']},
    'conclusions': {'concl_1': {'topic': 'YOKE dirty plating material disposition',
                                'statement_from_report': 'Tension OK, marking OK, but NG visual SPK high 7/18 = 38.9%.',
                                'normalized_interpretation': 'Tension PASS: Test avg 128.41 kgf (106.10-152.70) vs Normal 105.56 kgf (93.60-111.70) — Test +21.6% higher mean. SPK Visual NG: 7/18 = 38.9%, no Normal control at this stage. High absolute rate gates dirty-plating YOKE.'}},
    'hints': {'hint_1': {'check_item': 'For dirty-plating YOKE, run both Tension and final SPK Visual; do not approve on Tension alone',
                         'reason': 'Tension passed (avg 128.41 kgf > 50 kgf spec, higher than Normal) but SPK Visual NG hit 38.9% (7/18). Mechanical PASS does not predict visual outcome.'}},
    'log': {'assumptions': ['Tension spec >=50kgf from sheet text.', 'Visual sample 18 used; intro mentioned 20 pcs.'],
            'warnings': ['SPK Visual lacks same-event Normal; 38.9% read as fail signal.'],
            'decision_rationale': 'mixed: Tension (reliability spec) + SPK Visual (NG-rate) co-primary. Tension Test +21.6% above Normal mean. Visual judged by absolute NG vs zero expectation.'}
}

ds6_tr_ko = {
    'document': {'title': 'BRS-161016 — Ethanol 세척 후 YOKE 도금 불량(dirty plating) 자재 시험',
                 'purpose': 'YOKE 161016 Ethanol 세척 후 도금 불량 — 임시 한도 승인 가능 여부 검증.',
                 'content': ['Tension test CMG (스펙 >=50kgf): Test/Normal YOKE 각 30 pcs. SPK final marking + visual: 18 pcs.']},
    'conclusions': {'concl_1': {'topic': 'YOKE dirty plating 자재 판정',
                                'statement_from_report': 'Tension OK, marking OK, 그러나 NG visual SPK 매우 높음 7/18 = 38.9%.',
                                'normalized_interpretation': 'Tension PASS: Test 평균 128.41 kgf (106.10-152.70) vs Normal 105.56 kgf (93.60-111.70) — Test 평균이 21.6% 더 높음. SPK Visual NG 7/18 = 38.9%, 이 단계엔 Normal 없음. 절대치 높아 dirty plating YOKE 사용 불가 신호.'}},
    'hints': {'hint_1': {'check_item': 'Dirty plating YOKE는 Tension만이 아니라 final SPK Visual도 동반 검사; Tension 통과만으로 승인 금지',
                         'reason': 'Tension 통과(평균 128.41 kgf > 50 kgf 스펙, Normal보다 높음)했으나 SPK Visual NG 38.9%(7/18). 기계적 PASS가 외관 결과를 보장 못함.'}},
    'log': {'assumptions': ['스펙 >=50kgf는 시트 텍스트에서 채택.', 'Visual 표본 18 사용; 도입부는 20pcs 계획.'],
            'warnings': ['SPK Visual에 same-event Normal 없음; 38.9%를 fail 신호로 해석.'],
            'decision_rationale': 'mixed: Tension(reliability spec) + SPK Visual(NG-rate) 공동 주요 근거. Tension Test가 Normal 평균보다 21.6% 높음. Visual은 절대 NG로 판단.'}
}

ds6_tr_vi = {
    'document': {'title': 'BRS-161016 — Test vật liệu YOKE bị dirty plating sau khi clean Ethanol',
                 'purpose': 'YOKE 161016 dirty plating sau clean Ethanol — kiểm tra để duyệt giới hạn tạm thời.',
                 'content': ['Tension test CMG (spec >=50kgf): 30 pcs Test và 30 pcs Normal YOKE. Final SPK marking + visual: 18 pcs.']},
    'conclusions': {'concl_1': {'topic': 'Xử lý vật liệu YOKE dirty plating',
                                'statement_from_report': 'Tension OK, marking OK, nhưng NG visual SPK cao 7/18 = 38.9%.',
                                'normalized_interpretation': 'Tension PASS: Test trung bình 128.41 kgf (106.10-152.70) vs Normal 105.56 kgf (93.60-111.70) — Test cao hơn 21.6%. SPK Visual NG 7/18 = 38.9%, không có Normal cùng sự kiện ở bước này. Tỷ lệ tuyệt đối cao chặn việc duyệt YOKE dirty plating.'}},
    'hints': {'hint_1': {'check_item': 'Với YOKE dirty plating, chạy cả Tension và final SPK Visual; không duyệt chỉ dựa Tension',
                         'reason': 'Tension đạt (trung bình 128.41 kgf > 50 kgf spec, cao hơn Normal) nhưng SPK Visual NG 38.9% (7/18). PASS cơ học không dự đoán được kết quả visual.'}},
    'log': {'assumptions': ['Spec Tension >=50kgf lấy từ chữ trong sheet.', 'Mẫu visual 18 dùng; intro nói 20pcs.'],
            'warnings': ['SPK Visual thiếu Normal same-event; coi 38.9% là tín hiệu fail.'],
            'decision_rationale': 'mixed: Tension (reliability spec) + SPK Visual (NG-rate) đồng chủ đạo. Tension Test cao hơn Normal trung bình 21.6%. Visual dựa trên NG tuyệt đối.'}
}


# =====================================================================
# DATASET 7: 33. BRS-201506 Report test material VP happen a little abnormal date 25.3.2024
# normal_comparison: VP AME Film abnormal vs Normal across Air Leak, Vision, Function
# =====================================================================
ds7_name = '33. BRS-201506 Report test material VP happen a little abnormal date 25.3.2024'

ds7_result = {
    'schema_version': '0.1',
    'document': {
        'document_id': 'doc_7',
        'source_file': ds7_name,
        'source_sheet': 'Test VP',
        'title': 'BRS-201506 REPORT TEST MATERIAL VP AEM FILM HAPPEN A LITTLE ABNORMAL',
        'model': 'BRS-201506',
        'report_date': '2024-03-25',
        'department': 'ME',
        'marker': 'Thao',
        'line': '',
        'report_type': 'normal_comparison',
        'primary_defect': {'canonical_name': 'VP+CD Separation',
                           'aliases_in_document': ['Short VP separate', 'NG Vision', 'AEM Film abnormal']},
        'related_defects': ['Long VP separate', 'Short VP damage', 'Short VP not enough glue',
                            'NG Sigma SPL', 'NG Hearing Noise', 'NG Hearing Touch'],
        'parts': ['VP', 'AEM Film'],
        'processes': ['Air Leak', 'VP Vision (Sub 1)', 'Function'],
        'purpose': 'Check whether VP with AEM Film abnormal can be used.',
        'content': ["Check visual after forming. Check NG rate at Sub 1. Make sample and check function. Q'ty test 924 pcs."],
        'source_cells': {'title': ['Test VP!B1'], 'date': ['Test VP!T2'], 'purpose': ['Test VP!A4'],
                         'content': ['Test VP!A6']}
    },
    'test_conditions': [
        {'condition_id': 'cond_1', 'condition_group': 'vp_aem_abnormal', 'line': '', 'process': 'VP forming',
         'changed_factor': 'Use VP with AEM Film a little abnormal', 'before_value': 'VP normal',
         'after_value': 'VP Test (AEM Film abnormal)', 'unit': None, 'machine': None, 'jig': None,
         'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None,
         'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': ds7_name, 'sheet_name': 'Test VP', 'source_cells': ['Test VP!A4']}
    ],
    'results': [
        {'result_id': 'airleak_test', 'condition_id': 'cond_1', 'measurement_type': 'Air Leak',
         'condition_group': 'VP Test', 'date': '2024-03-26', 'line': '',
         'input_count': 920, 'ok_count': 920, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Air Leak NG Rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {},
         'source_file': ds7_name, 'sheet_name': 'Test VP', 'source_cells': ['Test VP!_airleak_test']},
        {'result_id': 'airleak_normal', 'condition_id': None, 'measurement_type': 'Air Leak',
         'condition_group': 'VP normal', 'date': '2024-03-26', 'line': '',
         'input_count': 870, 'ok_count': 870, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Air Leak NG Rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {},
         'source_file': ds7_name, 'sheet_name': 'Test VP', 'source_cells': ['Test VP!_airleak_normal']},
        {'result_id': 'vision_test', 'condition_id': 'cond_1', 'measurement_type': 'VP Vision',
         'condition_group': 'VP Test', 'date': '2024-03-26', 'line': '',
         'input_count': 920, 'ok_count': 914, 'ng_count': 6, 'ng_rate_decimal': 0.0065, 'ng_rate_percent': 0.7,
         'metric_name': 'Vision NG Rate', 'metric_value': 0.7, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'Short VP separate': {'count': 5, 'rate': 0.54},
                          'Short VP damage': {'count': 0, 'rate': 0.0},
                          'Short VP not enough glue': {'count': 1, 'rate': 0.11},
                          'Long VP separate': {'count': 0, 'rate': 0.0},
                          'Long VP damage': {'count': 0, 'rate': 0.0},
                          'Long VP not enough glue': {'count': 0, 'rate': 0.0}},
         'source_file': ds7_name, 'sheet_name': 'Test VP', 'source_cells': ['Test VP!_vision_test']},
        {'result_id': 'vision_normal', 'condition_id': None, 'measurement_type': 'VP Vision',
         'condition_group': 'VP normal', 'date': '2024-03-26', 'line': '',
         'input_count': 870, 'ok_count': 869, 'ng_count': 1, 'ng_rate_decimal': 0.00115, 'ng_rate_percent': 0.1,
         'metric_name': 'Vision NG Rate', 'metric_value': 0.1, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'Short VP separate': {'count': 0, 'rate': 0.0},
                          'Short VP damage': {'count': 0, 'rate': 0.0},
                          'Short VP not enough glue': {'count': 0, 'rate': 0.0},
                          'Long VP separate': {'count': 1, 'rate': 0.115},
                          'Long VP damage': {'count': 0, 'rate': 0.0},
                          'Long VP not enough glue': {'count': 0, 'rate': 0.0}},
         'source_file': ds7_name, 'sheet_name': 'Test VP', 'source_cells': ['Test VP!_vision_normal']},
        {'result_id': 'func_test', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'VP Test', 'date': '2024-03-26', 'line': '',
         'input_count': 914, 'ok_count': 879, 'ng_count': 35, 'ng_rate_decimal': 0.038, 'ng_rate_percent': 3.8,
         'metric_name': 'Function NG Rate', 'metric_value': 3.8, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'NG Sigma SPL': {'count': 5, 'rate': 14.3},
                          'NG Sigma THD': {'count': 0, 'rate': 0.0},
                          'NG Hearing Noise': {'count': 12, 'rate': 34.3},
                          'NG Hearing Touch': {'count': 18, 'rate': 51.4},
                          'HOHD': {'count': 0, 'rate': 0.0}},
         'source_file': ds7_name, 'sheet_name': 'Test VP', 'source_cells': ['Test VP!_func_test']},
        {'result_id': 'func_normal', 'condition_id': None, 'measurement_type': 'Function',
         'condition_group': 'VP normal', 'date': '2024-03-26', 'line': '',
         'input_count': 869, 'ok_count': 829, 'ng_count': 40, 'ng_rate_decimal': 0.046, 'ng_rate_percent': 4.6,
         'metric_name': 'Function NG Rate', 'metric_value': 4.6, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'NG Sigma SPL': {'count': 4, 'rate': 10.0},
                          'NG Sigma THD': {'count': 0, 'rate': 0.0},
                          'NG Hearing Noise': {'count': 16, 'rate': 40.0},
                          'NG Hearing Touch': {'count': 20, 'rate': 50.0},
                          'HOHD': {'count': 0, 'rate': 0.0}},
         'source_file': ds7_name, 'sheet_name': 'Test VP', 'source_cells': ['Test VP!_func_normal']},
    ],
    'conclusions': [
        {'conclusion_id': 'concl_1', 'topic': 'VP AEM Film abnormal — CAN NOT USE',
         'statement_from_report': 'Process VP vision Short VP separate high 6/920 = 0.7% vs Normal 1/870 = 0.1%. Function NG 35/914 = 3.8% less than Normal 4.6%. => CAN NOT USE.',
         'normalized_interpretation': 'Air Leak Test 0.0% vs Normal 0.0% — equal. VP Vision Test 0.7% vs Normal 0.1% = 7.0x, 600% worse (Short VP separate is the dominant new defect: 5/920 vs 0/870). Function Test 3.8% vs Normal 4.6% = 0.826x, 17.4% improved (within noise). Decision is driven by the 6x increase in upstream Vision defects, not by Function.',
         'source_file': ds7_name, 'sheet_name': 'Test VP', 'source_cells': ['Test VP!_decision']}
    ],
    'troubleshooting_index': {
        'defect_name': 'VP+CD Separation',
        'when_user_asks': ['VP AEM Film abnormal', 'Short VP separate', 'VP material disposition', 'Vision NG up'],
        'suggested_checks': [
            {'hint_id': 'hint_1', 'check_item': 'When VP material shows AEM-film abnormality, treat VP Vision Short-VP-separate rate (not Function) as the gating metric',
             'reason': 'Vision Test 0.7% vs Normal 0.1% = 7x degradation, Short VP separate 5/920 vs 0/870. Function actually improved 17% (within noise) and would have falsely cleared the material.',
             'evidence_strength': 'high', 'related_process': 'VP Vision / Sub 1', 'related_part': 'VP / AEM Film',
             'source_file': ds7_name, 'sheet_name': 'Test VP', 'source_cells': ['Test VP!_vision_test']}
        ],
        'limitations': ['Function comparison has low statistical power for small subdefect shifts; main signal is concentrated in Vision Short-VP-separate.']
    },
    'ai_extraction_log': {
        'confidence': 0.9,
        'assumptions': ['"Short VP separate" mapped to canonical VP+CD Separation family.'],
        'warnings': ['Function metric is misleading on its own (Test improved vs Normal); rejection is correctly based on Vision.'],
        'decision_rationale': 'Same-event Normal exists for all three metrics. Multiplicative changes: Air Leak 1.0x (equal), Vision 7.0x worse, Function 0.83x improved. Workbook rejects material based on Vision; AI mirrors this and warns against using Function alone.'
    }
}

ds7_tr_en = {
    'document': {'title': 'BRS-201506 — Test VP AEM Film a little abnormal',
                 'purpose': 'Check whether VP with AEM Film abnormal can be used.',
                 'content': ['Check visual after forming, NG rate at Sub 1, function on samples; Q\'ty 924 pcs.']},
    'conclusions': {'concl_1': {'topic': 'VP AEM Film abnormal — CAN NOT USE',
                                'statement_from_report': 'Vision Short VP separate high 6/920 = 0.7% vs Normal 1/870 = 0.1%. Function NG 35/914 = 3.8% < Normal 4.6%. => CAN NOT USE.',
                                'normalized_interpretation': 'Air Leak 0% vs 0% — equal. Vision Test 0.7% vs Normal 0.1% = 7.0x, 600% worse; Short VP separate is the new defect (5/920 vs 0/870). Function 3.8% vs 4.6% = 0.826x, 17.4% improved (within noise). Decision driven by Vision degradation.'}},
    'hints': {'hint_1': {'check_item': 'For VP AEM-film abnormality, gate on VP Vision Short-VP-separate rate, not Function',
                         'reason': 'Vision 0.7% vs Normal 0.1% = 7x worse, Short VP separate 5/920 vs 0/870. Function actually improved 17% (within noise) and would have falsely cleared the material.'}},
    'log': {'assumptions': ['"Short VP separate" mapped to VP+CD Separation family.'],
            'warnings': ['Function alone misleading (Test better than Normal); rejection correctly based on Vision.'],
            'decision_rationale': 'Same-event Normal exists for all three. Multiplicative changes Air Leak 1.0x, Vision 7.0x worse, Function 0.83x improved. Vision gates the decision.'}
}

ds7_tr_ko = {
    'document': {'title': 'BRS-201506 — VP AEM Film 다소 비정상 자재 시험',
                 'purpose': 'AEM Film 비정상 VP 사용 가능 여부 확인.',
                 'content': ['Forming 후 visual, Sub 1 NG rate, sample function 검사; Q\'ty 924 pcs.']},
    'conclusions': {'concl_1': {'topic': 'VP AEM Film abnormal — 사용 불가',
                                'statement_from_report': 'Vision Short VP separate 6/920 = 0.7% vs Normal 1/870 = 0.1%. Function NG 35/914 = 3.8% < Normal 4.6%. => 사용 불가.',
                                'normalized_interpretation': 'Air Leak 0% vs 0% — 동일. Vision Test 0.7% vs Normal 0.1% = 7.0배, 600% 악화; Short VP separate가 신규 결함(5/920 vs 0/870). Function 3.8% vs 4.6% = 0.826배, 17.4% 개선(노이즈 범위). 판정은 Vision 악화에 근거.'}},
    'hints': {'hint_1': {'check_item': 'VP AEM-film 비정상 자재는 Function이 아닌 VP Vision Short-VP-separate 비율로 게이트',
                         'reason': 'Vision 0.7% vs Normal 0.1% = 7배 악화, Short VP separate 5/920 vs 0/870. Function은 오히려 17% 개선(노이즈 범위)되어 잘못 승인 가능.'}},
    'log': {'assumptions': ['"Short VP separate"를 VP+CD Separation 표준명으로 매핑.'],
            'warnings': ['Function 단독은 오해 소지(Test가 Normal보다 양호); Vision 근거로 정확히 기각.'],
            'decision_rationale': '세 지표 모두 same-event Normal 존재. 곱셈 변화 Air Leak 1.0배, Vision 7.0배 악화, Function 0.83배 개선. Vision이 결정 근거.'}
}

ds7_tr_vi = {
    'document': {'title': 'BRS-201506 — Test vật liệu VP AEM Film hơi bất thường',
                 'purpose': 'Kiểm tra xem VP có AEM Film bất thường có dùng được không.',
                 'content': ['Kiểm visual sau forming, NG rate ở Sub 1, function trên mẫu; Q\'ty 924 pcs.']},
    'conclusions': {'concl_1': {'topic': 'VP AEM Film bất thường — KHÔNG DÙNG ĐƯỢC',
                                'statement_from_report': 'Vision Short VP separate 6/920 = 0.7% vs Normal 1/870 = 0.1%. Function NG 35/914 = 3.8% < Normal 4.6%. => KHÔNG DÙNG ĐƯỢC.',
                                'normalized_interpretation': 'Air Leak 0% vs 0% — bằng. Vision Test 0.7% vs Normal 0.1% = 7.0x, xấu hơn 600%; Short VP separate là lỗi mới (5/920 vs 0/870). Function 3.8% vs 4.6% = 0.826x, cải thiện 17.4% (trong nhiễu). Quyết định dựa trên Vision xấu đi.'}},
    'hints': {'hint_1': {'check_item': 'VP AEM-film bất thường thì xét theo VP Vision Short-VP-separate, không xét Function',
                         'reason': 'Vision 0.7% vs Normal 0.1% = 7x xấu hơn, Short VP separate 5/920 vs 0/870. Function lại cải thiện 17% (trong nhiễu) và có thể duyệt nhầm.'}},
    'log': {'assumptions': ['"Short VP separate" map về VP+CD Separation chuẩn.'],
            'warnings': ['Chỉ xét Function dễ nhầm (Test tốt hơn Normal); việc từ chối dựa Vision là đúng.'],
            'decision_rationale': 'Cả ba có Normal same-event. Tỷ số nhân: Air Leak 1.0x, Vision 7.0x xấu, Function 0.83x tốt. Vision quyết định.'}
}


# =====================================================================
# DATASET 8: 33. MSU-L20S15-07GMI Report test Find reason NG frequency SPL increase high - date 09.10.2025
# normal_comparison (two events): bond amount tests vs Normal + DOE condition matrix
# =====================================================================
ds8_name = '33. MSU-L20S15-07GMI Report test Find reason NG frequency SPL increase high -  date 09.10.2025'

ds8_result = {
    'schema_version': '0.1',
    'document': {
        'document_id': 'doc_8',
        'source_file': ds8_name,
        'source_sheet': 'Test  (2) + Test',
        'title': 'MSU-L20S15-07 GMI — REPORT TEST FIND REASON NG FREQUENCY SPL INCREASE',
        'model': 'MSU-L20S15-07 GMI',
        'report_date': '2025-10-09',
        'department': 'ME',
        'marker': 'Nhung',
        'line': 'E2-4B',
        'report_type': 'normal_comparison',
        'primary_defect': {'canonical_name': 'NG Sigma SPL',
                           'aliases_in_document': ['NG SPL', 'NG Sigma SPL', 'NG frequency SPL increase']},
        'related_defects': ['NG Sigma THD', 'NG Sigma SPL+THD', 'NG Hearing Noise', 'NG Hearing Touch'],
        'parts': ['VP', 'CD', 'Dome', 'Frame', 'Cushion', 'Jacket'],
        'processes': ['VP bonding inside', 'VP+CD bonding (Sub 1)', 'VP+Frame bonding', 'UV LED dry', 'Function'],
        'purpose': 'NG frequency SPL lower point 10~14k increased; find root cause.',
        'content': ['1) Change all cushion and Jacket new; 2) Check bond amount VP+CD/VP+Frame; 3) Check bonding line VP+CD offset and adjust; 4) Compare CD Ralon vs CD GES dimension/weight; 5) UV dry bond VP+CD unstable -> request PE repair; 6) Check bond amount CD min/max spec; plus earlier test on 10/25 increasing bond amount and dry time Semi VP+CD.'],
        'source_cells': {'title': ['Test  (2)!B1', 'Test!B1'], 'date': ['Test  (2)!N3', 'Test!N3'],
                         'purpose': ['Test  (2)!A4', 'Test!A4'], 'content': ['Test  (2)!A6', 'Test!A6']}
    },
    'test_conditions': [
        {'condition_id': 'cond_1', 'condition_group': 'bond_amount_vp_max', 'line': 'E2-4B', 'process': 'VP bonding inside',
         'changed_factor': 'Increase Bond amount to max spec (1.8~2.0mg, used ~1.96mg both nozzles)', 'before_value': None,
         'after_value': '1.96 mg (Nozzle1=Nozzle2)', 'unit': 'mg', 'machine': None, 'jig': None,
         'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None,
         'pressure': None, 'bond_amount': 'Nozzle1 1.96 / Nozzle2 1.96', 'uv_energy': None,
         'source_file': ds8_name, 'sheet_name': 'Test  (2)', 'source_cells': ['Test  (2)!_bond_amount']},
        {'condition_id': 'cond_2', 'condition_group': 'repair_offset', 'line': 'E2-4B', 'process': 'VP bonding inside',
         'changed_factor': 'After Repair bond offset', 'before_value': None,
         'after_value': 'Repaired', 'unit': None, 'machine': None, 'jig': None,
         'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None,
         'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': ds8_name, 'sheet_name': 'Test  (2)', 'source_cells': ['Test  (2)!_repair']},
        {'condition_id': 'cond_3', 'condition_group': 'dry_time_semi_vpcd', 'line': 'E2-4B', 'process': 'Sub 1 Semi VP+CD dry',
         'changed_factor': 'Dry time Semi VP+CD', 'before_value': '5 min',
         'after_value': '10 min', 'unit': 'min', 'machine': None, 'jig': None,
         'material_lot': None, 'supplier': None, 'dry_time_sec': 600.0, 'temperature': None,
         'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': ds8_name, 'sheet_name': 'Test  (2)', 'source_cells': ['Test  (2)!_drytime']},
        {'condition_id': 'cond_4', 'condition_group': 'cushion_jacket_new', 'line': 'E2-4B', 'process': 'Function',
         'changed_factor': 'Change all cushion and Jacket new', 'before_value': None,
         'after_value': None, 'unit': None, 'machine': None, 'jig': None,
         'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None,
         'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': ds8_name, 'sheet_name': 'Test', 'source_cells': ['Test!_cushion']},
        {'condition_id': 'cond_5', 'condition_group': 'bondline_offset_adjust', 'line': 'E2-4B', 'process': 'VP+CD bonding',
         'changed_factor': 'Adjust bonding line offset Y by -0.05', 'before_value': 'Original',
         'after_value': 'Y-0.05', 'unit': 'mm', 'machine': None, 'jig': None,
         'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None,
         'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': ds8_name, 'sheet_name': 'Test', 'source_cells': ['Test!_offset_y']},
        {'condition_id': 'cond_6', 'condition_group': 'uv_led_adjust', 'line': 'E2-4B', 'process': 'UV LED dry',
         'changed_factor': 'Adjust UV LED (unstable, sometimes not dry)', 'before_value': 'Unstable',
         'after_value': 'Adjusted', 'unit': None, 'machine': None, 'jig': None,
         'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None,
         'pressure': None, 'bond_amount': None, 'uv_energy': 'Adjusted',
         'source_file': ds8_name, 'sheet_name': 'Test', 'source_cells': ['Test!_uvled']},
        {'condition_id': 'cond_7', 'condition_group': 'dome_vendor', 'line': 'E2-4B', 'process': 'Sub 1',
         'changed_factor': 'Use Dome of vendor Ralon (stock) instead of GES (after UV LED adjust)', 'before_value': 'GES',
         'after_value': 'Ralon', 'unit': None, 'machine': None, 'jig': None,
         'material_lot': 'Ralon stock', 'supplier': 'Ralon', 'dry_time_sec': None, 'temperature': None,
         'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': ds8_name, 'sheet_name': 'Test', 'source_cells': ['Test!_ralon']},
        {'condition_id': 'cond_8', 'condition_group': 'bond_vpcd_minmax', 'line': 'E2-4B', 'process': 'VP+CD bonding',
         'changed_factor': 'Bond amount VP+CD at spec min (3.34~3.36) and max (3.67~3.68), standard 3.3~3.7mg', 'before_value': None,
         'after_value': None, 'unit': 'mg', 'machine': None, 'jig': None,
         'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None,
         'pressure': None, 'bond_amount': '3.34~3.68 spec', 'uv_energy': None,
         'source_file': ds8_name, 'sheet_name': 'Test', 'source_cells': ['Test!_minmax']},
    ],
    'results': [
        # Test (2): 10/26 increase bond max spec
        {'result_id': 't2_bondmax', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Increase Bond max spec', 'date': '2025-10-26', 'line': 'E2-4B',
         'input_count': 280, 'ok_count': 270, 'ng_count': 10, 'ng_rate_decimal': 0.036, 'ng_rate_percent': 3.6,
         'metric_name': 'Function NG Rate', 'metric_value': 3.6, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'NG Sigma SPL': {'count': 4, 'rate': 1.4}, 'NG Sigma THD': {'count': 0, 'rate': 0.0},
                          'NG Sigma SPL+THD': {'count': 1, 'rate': 0.4}, 'NG Sigma SPL+THD+F0': {'count': 0, 'rate': 0.0},
                          'NG Hearing Noise': {'count': 3, 'rate': 1.1}, 'NG Hearing Touch': {'count': 2, 'rate': 0.7}},
         'source_file': ds8_name, 'sheet_name': 'Test  (2)', 'source_cells': ['Test  (2)!_bondmax']},
        {'result_id': 't2_repair', 'condition_id': 'cond_2', 'measurement_type': 'Function',
         'condition_group': 'After Repair bond offset', 'date': '2025-10-26', 'line': 'E2-4B',
         'input_count': 300, 'ok_count': 279, 'ng_count': 21, 'ng_rate_decimal': 0.07, 'ng_rate_percent': 7.0,
         'metric_name': 'Function NG Rate', 'metric_value': 7.0, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG Sigma SPL': {'count': 3, 'rate': 1.0}, 'NG Sigma THD': {'count': 0, 'rate': 0.0},
                          'NG Sigma SPL+THD': {'count': 1, 'rate': 0.3}, 'NG Sigma SPL+THD+F0': {'count': 1, 'rate': 0.3},
                          'NG Hearing Noise': {'count': 9, 'rate': 3.0}, 'NG Hearing Touch': {'count': 7, 'rate': 2.3}},
         'source_file': ds8_name, 'sheet_name': 'Test  (2)', 'source_cells': ['Test  (2)!_repair']},
        {'result_id': 't2_row3', 'condition_id': None, 'measurement_type': 'Function',
         'condition_group': 'Reference row (no label in workbook)', 'date': '2025-10-26', 'line': 'E2-4B',
         'input_count': 189, 'ok_count': 185, 'ng_count': 4, 'ng_rate_decimal': 0.021, 'ng_rate_percent': 2.1,
         'metric_name': 'Function NG Rate', 'metric_value': 2.1, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'NG Hearing Noise': {'count': 3, 'rate': 1.6}, 'NG Hearing Touch': {'count': 1, 'rate': 0.5}},
         'source_file': ds8_name, 'sheet_name': 'Test  (2)', 'source_cells': ['Test  (2)!_t2row3']},
        {'result_id': 't2_dry5', 'condition_id': 'cond_3', 'measurement_type': 'Function',
         'condition_group': 'Dry time Semi VP+CD 5min', 'date': '2025-10-28', 'line': 'E2-4B',
         'input_count': 280, 'ok_count': 270, 'ng_count': 10, 'ng_rate_decimal': 0.036, 'ng_rate_percent': 3.6,
         'metric_name': 'Function NG Rate', 'metric_value': 3.6, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'NG Sigma SPL': {'count': 4, 'rate': 1.4}, 'NG Sigma THD': {'count': 0, 'rate': 0.0},
                          'NG Sigma SPL+THD': {'count': 1, 'rate': 0.4}, 'NG Sigma SPL+THD+F0': {'count': 0, 'rate': 0.0},
                          'NG Hearing Noise': {'count': 3, 'rate': 1.1}, 'NG Hearing Touch': {'count': 2, 'rate': 0.7}},
         'source_file': ds8_name, 'sheet_name': 'Test  (2)', 'source_cells': ['Test  (2)!_dry5']},
        {'result_id': 't2_dry7', 'condition_id': 'cond_3', 'measurement_type': 'Function',
         'condition_group': 'Dry time Semi VP+CD 7min', 'date': '2025-10-28', 'line': 'E2-4B',
         'input_count': 300, 'ok_count': 279, 'ng_count': 21, 'ng_rate_decimal': 0.07, 'ng_rate_percent': 7.0,
         'metric_name': 'Function NG Rate', 'metric_value': 7.0, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG Sigma SPL': {'count': 3, 'rate': 1.0}, 'NG Sigma THD': {'count': 0, 'rate': 0.0},
                          'NG Sigma SPL+THD': {'count': 1, 'rate': 0.3}, 'NG Sigma SPL+THD+F0': {'count': 1, 'rate': 0.3},
                          'NG Hearing Noise': {'count': 9, 'rate': 3.0}, 'NG Hearing Touch': {'count': 7, 'rate': 2.3}},
         'source_file': ds8_name, 'sheet_name': 'Test  (2)', 'source_cells': ['Test  (2)!_dry7']},
        {'result_id': 't2_dry10', 'condition_id': 'cond_3', 'measurement_type': 'Function',
         'condition_group': 'Dry time Semi VP+CD 10min', 'date': '2025-10-28', 'line': 'E2-4B',
         'input_count': 199, 'ok_count': 195, 'ng_count': 4, 'ng_rate_decimal': 0.020, 'ng_rate_percent': 2.0,
         'metric_name': 'Function NG Rate', 'metric_value': 2.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'NG Hearing Noise': {'count': 3, 'rate': 1.5}, 'NG Hearing Touch': {'count': 1, 'rate': 0.5}},
         'source_file': ds8_name, 'sheet_name': 'Test  (2)', 'source_cells': ['Test  (2)!_dry10']},
        # Test sheet 10/9 series
        {'result_id': 'sigspl_normal', 'condition_id': None, 'measurement_type': 'Function (Sigma)',
         'condition_group': 'Normal', 'date': '2025-10-09', 'line': 'E2-4B',
         'input_count': 963, 'ok_count': 910, 'ng_count': 53, 'ng_rate_decimal': 0.055, 'ng_rate_percent': 5.5,
         'metric_name': 'Sigma NG Rate', 'metric_value': 5.5, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG Sigma SPL': {'count': 53, 'rate': 5.5}, 'NG Sigma THD': {'count': 0, 'rate': 0.0},
                          'NG Sigma SPL+THD': {'count': 0, 'rate': 0.0}, 'NG Sigma SPL+THD+F0': {'count': 0, 'rate': 0.0}},
         'source_file': ds8_name, 'sheet_name': 'Test', 'source_cells': ['Test!_normal']},
        {'result_id': 'sigspl_cushion', 'condition_id': 'cond_4', 'measurement_type': 'Function (Sigma)',
         'condition_group': 'Change all cushion and Jacket new', 'date': '2025-10-09', 'line': 'E2-4B',
         'input_count': 965, 'ok_count': 912, 'ng_count': 53, 'ng_rate_decimal': 0.055, 'ng_rate_percent': 5.5,
         'metric_name': 'Sigma NG Rate', 'metric_value': 5.5, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG Sigma SPL': {'count': 53, 'rate': 5.5}},
         'source_file': ds8_name, 'sheet_name': 'Test', 'source_cells': ['Test!_cushion']},
        {'result_id': 'sigspl_offset', 'condition_id': 'cond_5', 'measurement_type': 'Function (Sigma)',
         'condition_group': 'After adjust bonding line offset Y-0.05', 'date': '2025-10-09', 'line': 'E2-4B',
         'input_count': 210, 'ok_count': 207, 'ng_count': 3, 'ng_rate_decimal': 0.014, 'ng_rate_percent': 1.4,
         'metric_name': 'Sigma NG Rate', 'metric_value': 1.4, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'NG Sigma SPL': {'count': 3, 'rate': 1.4}},
         'source_file': ds8_name, 'sheet_name': 'Test', 'source_cells': ['Test!_offset']},
        {'result_id': 'sigspl_uvled', 'condition_id': 'cond_6', 'measurement_type': 'Function (Sigma)',
         'condition_group': 'After adjust UV LED', 'date': '2025-10-09', 'line': 'E2-4B',
         'input_count': 300, 'ok_count': 300, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Sigma NG Rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'NG Sigma SPL': {'count': 0, 'rate': 0.0}},
         'source_file': ds8_name, 'sheet_name': 'Test', 'source_cells': ['Test!_uvled']},
        {'result_id': 'sigspl_ralon', 'condition_id': 'cond_7', 'measurement_type': 'Function (Sigma)',
         'condition_group': 'Use Dome Ralon (after UV LED adjust)', 'date': '2025-10-09', 'line': 'E2-4B',
         'input_count': 232, 'ok_count': 231, 'ng_count': 1, 'ng_rate_decimal': 0.004, 'ng_rate_percent': 0.4,
         'metric_name': 'Sigma NG Rate', 'metric_value': 0.4, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'NG Sigma SPL': {'count': 1, 'rate': 0.4}},
         'source_file': ds8_name, 'sheet_name': 'Test', 'source_cells': ['Test!_ralon']},
        {'result_id': 'sigspl_bondmin', 'condition_id': 'cond_8', 'measurement_type': 'Function (Sigma)',
         'condition_group': 'Bond amount VP+CD min spec (3.34~3.36)', 'date': '2025-10-09', 'line': 'E2-4B',
         'input_count': 225, 'ok_count': 225, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Sigma NG Rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'NG Sigma SPL': {'count': 0, 'rate': 0.0}},
         'source_file': ds8_name, 'sheet_name': 'Test', 'source_cells': ['Test!_bondmin']},
        {'result_id': 'sigspl_bondmax', 'condition_id': 'cond_8', 'measurement_type': 'Function (Sigma)',
         'condition_group': 'Bond amount VP+CD max spec (3.67~3.68)', 'date': '2025-10-09', 'line': 'E2-4B',
         'input_count': 249, 'ok_count': 248, 'ng_count': 1, 'ng_rate_decimal': 0.004, 'ng_rate_percent': 0.4,
         'metric_name': 'Sigma NG Rate', 'metric_value': 0.4, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'NG Sigma SPL': {'count': 1, 'rate': 0.4}},
         'source_file': ds8_name, 'sheet_name': 'Test', 'source_cells': ['Test!_bondmax']},
    ],
    'conclusions': [
        {'conclusion_id': 'concl_1', 'topic': 'NG Sigma SPL root cause and effective counter-measure',
         'statement_from_report': 'Bond amount VP+CD/VP+Frame within standard. CD Ralon and CD GES dimensions/weights similar (GES slightly bigger). Cushion/Jacket change did not improve. After adjust bonding line offset Y-0.05, NG dropped. After adjust UV LED (UV LED unstable causes incomplete dry), NG went to 0%. Bond min/max also OK after UV LED adjust.',
         'normalized_interpretation': 'Same-event Normal NG Sigma SPL = 5.5% (53/963, 10/9). Lever effects: cushion/jacket change 5.5% vs 5.5% = 1.0x (no effect); bonding line offset Y-0.05 1.4% vs 5.5% = 0.255x, 74.5% improved; UV LED adjust 0.0% vs 5.5% = 0.0x, 100% improved; Dome Ralon (post UV adjust) 0.4% vs 5.5% = 0.073x, 92.7% improved; Bond VP+CD min spec 0.0% (100% improved); Bond VP+CD max spec 0.4% (92.7% improved). UV LED instability is the root cause; bonding-line offset is a secondary contributor.',
         'source_file': ds8_name, 'sheet_name': 'Test', 'source_cells': ['Test!_decision_block']}
    ],
    'troubleshooting_index': {
        'defect_name': 'NG Sigma SPL',
        'when_user_asks': ['NG SPL increase', 'NG frequency SPL low-band 10-14k', 'UV LED dry', 'VP+CD bonding', 'Dome vendor'],
        'suggested_checks': [
            {'hint_id': 'hint_1', 'check_item': 'Check UV LED stability for VP+CD bond dry first when NG Sigma SPL rises; verify by separate-lot pilot before changing other parameters',
             'reason': 'In this report, changing cushion/jacket gave 0% improvement (5.5% vs 5.5%). Adjusting bonding-line offset Y-0.05 improved to 1.4% (74.5% vs 5.5%). Adjusting UV LED (sometimes not dry) brought NG to 0%/300 and 0.4%/232 with Ralon Dome — UV LED is the dominant lever; Dome vendor / bond amount min-max all PASS once UV is fixed.',
             'evidence_strength': 'high', 'related_process': 'UV LED dry, VP+CD bonding',
             'related_part': 'VP, CD, Dome',
             'source_file': ds8_name, 'sheet_name': 'Test', 'source_cells': ['Test!_uvled', 'Test!_ralon']}
        ],
        'limitations': ['Cushion/Jacket change is only one of many possible NG-source hypotheses; ruling it out has 5.5% wide-CI uncertainty.']
    },
    'ai_extraction_log': {
        'confidence': 0.9,
        'assumptions': ['Test (2) sheet rows for 10/28 are interpreted as a separate dry-time DOE; the 5min/7min/10min rates 3.6%/7.0%/2.0% are within the workbook table.', 'Same-event Normal (5.5%, 53/963 on 10/9) used as baseline for all 10/9 Test-sheet conditions.'],
        'warnings': ['Test (2) sheet has an unlabeled middle row (189/185/4) that resembles a reference; not used as baseline for the 10/26 series.'],
        'decision_rationale': 'Same-event Normal exists on the 10/9 Test sheet (5.5%). Multiplicative changes computed for each lever — UV LED adjust dominates (1.0x → 0.0x), offset Y-0.05 secondary (74.5% improved), and bond-amount/Dome-vendor become non-issues after UV LED is fixed.'
    }
}

ds8_tr_en = {
    'document': {'title': 'MSU-L20S15-07 GMI — Find reason NG frequency SPL increase',
                 'purpose': 'NG frequency SPL low-band 10~14k increased; find root cause.',
                 'content': ['Change cushion/jacket; check bond amount VP+CD and VP+Frame; adjust bonding-line offset; compare Dome Ralon vs GES; adjust UV LED dry; bond VP+CD min/max spec check; earlier dry-time Semi VP+CD 5/7/10 min DOE.']},
    'conclusions': {'concl_1': {'topic': 'NG Sigma SPL root cause and effective counter-measure',
                                'statement_from_report': 'Cushion/Jacket change did not improve. Bonding-line offset Y-0.05 reduced NG. UV LED adjust (UV LED unstable) brought NG to 0%. Bond min/max also OK after UV LED adjust.',
                                'normalized_interpretation': 'Same-event Normal NG Sigma SPL = 5.5% (10/9). Cushion/jacket 5.5% vs 5.5% = 1.0x. Offset Y-0.05 1.4% vs 5.5% = 0.255x, 74.5% improved. UV LED adjust 0.0% vs 5.5% = 100% improved. Dome Ralon (post UV) 0.4% vs 5.5% = 92.7% improved. Bond min 0.0%, Bond max 0.4%. UV LED instability is root cause; bond-line offset secondary.'}},
    'hints': {'hint_1': {'check_item': 'On NG Sigma SPL rise, check UV LED dry stability first; verify by separate-lot pilot before adjusting other parameters',
                         'reason': 'Cushion/jacket change gave 0% improvement (5.5% vs 5.5%). Offset Y-0.05 improved to 1.4% (74.5%). UV LED adjust → 0%/300 and 0.4%/232 (Ralon Dome). Dome vendor and bond amount become non-issues after UV LED fix.'}},
    'log': {'assumptions': ['Test (2) sheet rows for 10/28 are a separate dry-time DOE (5/7/10 min).',
                            'Same-event Normal (5.5%, 53/963 on 10/9) used as baseline for 10/9 Test-sheet conditions.'],
            'warnings': ['Test (2) sheet has unlabeled middle row (189/185/4) resembling reference; not used as 10/26 baseline.'],
            'decision_rationale': 'Same-event Normal exists on 10/9 sheet (5.5%). Multiplicative changes per lever: UV LED adjust dominates (5.5%→0%), offset secondary (74.5% improved), bond-amount/Dome non-issue post-UV.'}
}

ds8_tr_ko = {
    'document': {'title': 'MSU-L20S15-07 GMI — NG frequency SPL 상승 원인 규명',
                 'purpose': 'NG frequency SPL 저음대(10~14k) 상승; 근본 원인 파악.',
                 'content': ['Cushion/Jacket 교체; bond amount VP+CD 및 VP+Frame 점검; bonding line offset 조정; Dome Ralon vs GES 비교; UV LED dry 조정; bond VP+CD min/max 스펙 점검; 추가로 dry time Semi VP+CD 5/7/10분 DOE.']},
    'conclusions': {'concl_1': {'topic': 'NG Sigma SPL 근본 원인 및 효과적 대책',
                                'statement_from_report': 'Cushion/Jacket 교체로 개선 없음. Bonding line offset Y-0.05로 NG 감소. UV LED 조정(UV LED 불안정)으로 NG 0% 달성. UV LED 조정 후 bond min/max도 OK.',
                                'normalized_interpretation': 'Same-event Normal NG Sigma SPL = 5.5%(10/9). Cushion/jacket 5.5% vs 5.5% = 1.0배. Offset Y-0.05 1.4% vs 5.5% = 0.255배, 74.5% 개선. UV LED 조정 0.0% vs 5.5% = 100% 개선. Dome Ralon(UV 조정 후) 0.4% vs 5.5% = 92.7% 개선. Bond min 0.0%, Bond max 0.4%. UV LED 불안정이 근본 원인; bond line offset 2차적.'}},
    'hints': {'hint_1': {'check_item': 'NG Sigma SPL 상승 시 UV LED dry 안정성 먼저 점검; 다른 파라미터 조정 전 별도 lot pilot으로 검증',
                         'reason': 'Cushion/Jacket 교체는 개선 0%(5.5% vs 5.5%). Offset Y-0.05로 1.4%(74.5% 개선). UV LED 조정 → 0%/300, Ralon Dome 0.4%/232. UV 수정 후엔 Dome vendor와 bond amount는 비쟁점.'}},
    'log': {'assumptions': ['Test (2) sheet 10/28 행은 별도 dry-time DOE (5/7/10분).',
                            'Same-event Normal(5.5%, 10/9 53/963)을 10/9 Test sheet 조건의 baseline으로 사용.'],
            'warnings': ['Test (2) sheet에 라벨 없는 중간 행(189/185/4) 존재; 10/26 baseline으로 사용하지 않음.'],
            'decision_rationale': '10/9 sheet에 same-event Normal(5.5%) 존재. 레버별 곱셈 변화: UV LED 조정 지배(5.5%→0%), offset 2차(74.5% 개선), UV 수정 후 bond/Dome 비쟁점.'}
}

ds8_tr_vi = {
    'document': {'title': 'MSU-L20S15-07 GMI — Tìm nguyên nhân NG frequency SPL tăng',
                 'purpose': 'NG frequency SPL điểm thấp 10~14k tăng; tìm nguyên nhân.',
                 'content': ['Thay cushion/jacket; kiểm bond amount VP+CD và VP+Frame; chỉnh offset bonding line; so sánh Dome Ralon vs GES; chỉnh UV LED dry; kiểm bond VP+CD min/max spec; thêm DOE dry time Semi VP+CD 5/7/10 phút.']},
    'conclusions': {'concl_1': {'topic': 'Nguyên nhân và biện pháp hiệu quả cho NG Sigma SPL',
                                'statement_from_report': 'Thay cushion/jacket không cải thiện. Offset Y-0.05 giảm NG. Chỉnh UV LED (UV LED không ổn) đưa NG về 0%. Sau khi chỉnh UV LED, bond min/max cũng OK.',
                                'normalized_interpretation': 'Normal same-event NG Sigma SPL = 5.5% (10/9). Cushion/jacket 5.5% vs 5.5% = 1.0x. Offset Y-0.05 1.4% vs 5.5% = 0.255x, cải thiện 74.5%. UV LED chỉnh 0.0% vs 5.5% = 100% cải thiện. Dome Ralon (sau UV) 0.4% vs 5.5% = 92.7% cải thiện. Bond min 0.0%, Bond max 0.4%. UV LED không ổn là nguyên nhân gốc; offset bonding line thứ yếu.'}},
    'hints': {'hint_1': {'check_item': 'Khi NG Sigma SPL tăng, kiểm tra độ ổn định UV LED dry trước; xác minh bằng pilot lot riêng trước khi chỉnh thông số khác',
                         'reason': 'Thay cushion/jacket cải thiện 0% (5.5% vs 5.5%). Offset Y-0.05 còn 1.4% (74.5%). Chỉnh UV LED → 0%/300 và Ralon Dome 0.4%/232. Dome vendor và bond amount không còn là vấn đề sau khi sửa UV LED.'}},
    'log': {'assumptions': ['Các dòng 10/28 trên sheet Test (2) là DOE dry-time riêng (5/7/10 phút).',
                            'Normal same-event (5.5%, 53/963 ngày 10/9) làm baseline cho các điều kiện trên sheet Test 10/9.'],
            'warnings': ['Sheet Test (2) có dòng giữa không gán nhãn (189/185/4); không dùng làm baseline cho chuỗi 10/26.'],
            'decision_rationale': 'Có Normal same-event trên sheet 10/9 (5.5%). Tỷ lệ nhân theo từng đòn bẩy: UV LED quyết định (5.5%→0%), offset thứ cấp (74.5% cải thiện), bond/Dome không còn vấn đề sau khi sửa UV.'}
}


# =====================================================================
# Commit all 8 datasets
# =====================================================================
DATASETS = [
    (ds1_name, ds1_result, ds1_tr_ko, ds1_tr_en, ds1_tr_vi),
    (ds2_name, ds2_result, ds2_tr_ko, ds2_tr_en, ds2_tr_vi),
    (ds3_name, ds3_result, ds3_tr_ko, ds3_tr_en, ds3_tr_vi),
    (ds4_name, ds4_result, ds4_tr_ko, ds4_tr_en, ds4_tr_vi),
    (ds5_name, ds5_result, ds5_tr_ko, ds5_tr_en, ds5_tr_vi),
    (ds6_name, ds6_result, ds6_tr_ko, ds6_tr_en, ds6_tr_vi),
    (ds7_name, ds7_result, ds7_tr_ko, ds7_tr_en, ds7_tr_vi),
    (ds8_name, ds8_result, ds8_tr_ko, ds8_tr_en, ds8_tr_vi),
]

processed = 0
failed = 0
for name, result, tr_ko, tr_en, tr_vi in DATASETS:
    ok = h.commit_dataset(name, result, tr_ko, tr_en, tr_vi)
    if ok:
        processed += 1
        print(f'OK: {name}')
    else:
        failed += 1
        print(f'FAIL: {name}')

print(f'chunk 11: processed={processed} failed={failed}')
total, ok_count, failed_log = h.verify_counts()
print(f'verify_counts: targets={total} processed={ok_count} failed_log_lines={failed_log}')
