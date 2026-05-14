"""Process chunk 02 datasets: build normalized results + 3-lang translations and commit."""
import _ai_batch_helper as h

results = {}

# =====================================================================
# DS01: 26. TIU C11-20 Report test Plate supplier MYUNGJIN difference colour 2026.02.26 - Copy
# =====================================================================
name01 = '26. TIU C11-20  Report test Plate supplier MYUNGJIN difference colour 2026.02.26 - Copy'

result01 = {
  'schema_version': '0.1',
  'document': {
    'document_id': '', 'source_file': name01, 'source_sheet': 'Test',
    'title': 'TIU C11-20 Report Test Plate Supplier MYUNGJIN Difference Color 100% C11-20',
    'model': 'TIU C11-20', 'report_date': '2026-02-26',
    'department': 'ME', 'marker': 'Trung', 'line': '',
    'report_type': 'normal_comparison',
    'primary_defect': {'canonical_name': 'NG Hearing Noise', 'aliases_in_document': ['NG hearing noise', 'NG Drop test FINAL auto by hearing noise']},
    'related_defects': ['Over Glue', 'NG MG+PT Separation', 'NG Not enough glue'],
    'parts': ['Plate', 'YOKE', 'MG', 'PT'], 'processes': ['Sub2', 'Bonding', 'Drop test', 'Function'],
    'purpose': 'Test whether plate from supplier MYUNGJIN with different color can be used.',
    'content': ['Make semi sub2 and check NG process', 'Decap check bond PT+MG', 'Drop test Auto/Manual', 'Check AI bonding detection rate', 'Input main line and check function'],
    'source_cells': {'title': ['Test!B1'], 'date': ['Test!N2'], 'purpose': ['Test!A4'], 'content': ['Test!A6:A10']}
  },
  'test_conditions': [
    {'condition_id': 'cond_1', 'condition_group': 'Test level 1', 'line': '', 'process': 'Sub2 + Drop test + Function', 'changed_factor': 'Plate supplier (MYUNGJIN different color, Test level 1)', 'before_value': 'Normal plate', 'after_value': 'MYUNGJIN different color Test level 1', 'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': 'MYUNGJIN', 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!D22']},
    {'condition_id': 'cond_2', 'condition_group': 'Test level 2', 'line': '', 'process': 'Sub2 + Drop test + Function', 'changed_factor': 'Plate supplier (MYUNGJIN different color, Test level 2)', 'before_value': 'Normal plate', 'after_value': 'MYUNGJIN different color Test level 2', 'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': 'MYUNGJIN', 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!D24']},
    {'condition_id': 'cond_3', 'condition_group': 'Test level 3', 'line': '', 'process': 'Sub2 + Drop test + Function', 'changed_factor': 'Plate supplier (MYUNGJIN different color, Test level 3)', 'before_value': 'Normal plate', 'after_value': 'MYUNGJIN different color Test level 3', 'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': 'MYUNGJIN', 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!D26']},
    {'condition_id': 'cond_4', 'condition_group': 'Normal', 'line': '', 'process': 'Sub2 + Drop test + Function', 'changed_factor': 'Baseline (Normal plate)', 'before_value': None, 'after_value': 'Normal', 'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!D28']},
  ],
  'results': [
    {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'NG Process Visual', 'condition_group': 'Test level 1', 'date': '2026-02-26', 'line': '', 'input_count': 92, 'ok_count': 92, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'NG Visual rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS', 'ng_breakdown': {'Over Glue': {'count': 0, 'rate': 0.0}, 'NG MG+PT Separation': {'count': 0, 'rate': 0.0}, 'NG Not enough glue': {'count': 0, 'rate': 0.0}}, 'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!F22:L22']},
    {'result_id': 'res_2', 'condition_id': 'cond_2', 'measurement_type': 'NG Process Visual', 'condition_group': 'Test level 2', 'date': '2026-02-26', 'line': '', 'input_count': 26, 'ok_count': 26, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'NG Visual rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS', 'ng_breakdown': {'Over Glue': {'count': 0, 'rate': 0.0}, 'NG MG+PT Separation': {'count': 0, 'rate': 0.0}, 'NG Not enough glue': {'count': 0, 'rate': 0.0}}, 'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!F24:L24']},
    {'result_id': 'res_3', 'condition_id': 'cond_3', 'measurement_type': 'NG Process Visual', 'condition_group': 'Test level 3', 'date': '2026-02-26', 'line': '', 'input_count': 28, 'ok_count': 28, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'NG Visual rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS', 'ng_breakdown': {}, 'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!F26:L26']},
    {'result_id': 'res_4', 'condition_id': 'cond_4', 'measurement_type': 'NG Process Visual', 'condition_group': 'Normal', 'date': '2026-02-26', 'line': '', 'input_count': 100, 'ok_count': 100, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'NG Visual rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS', 'ng_breakdown': {}, 'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!F28:L28']},
    {'result_id': 'res_5', 'condition_id': 'cond_1', 'measurement_type': 'Function', 'condition_group': 'Test level 1', 'date': '2026-02-11', 'line': '', 'input_count': 84, 'ok_count': 82, 'ng_count': 2, 'ng_rate_decimal': 0.024, 'ng_rate_percent': 2.4, 'metric_name': 'Function NG rate', 'metric_value': 2.4, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG Sigma SPL': {'count': 1, 'rate': 1.2}, 'NG Hearing Noise': {'count': 1, 'rate': 1.2}, 'NG Hearing Touch': {'count': 0, 'rate': 0.0}}, 'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!F55:L55']},
    {'result_id': 'res_6', 'condition_id': 'cond_2', 'measurement_type': 'Function', 'condition_group': 'Test level 2', 'date': '2026-02-26', 'line': '', 'input_count': 18, 'ok_count': 18, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Function NG rate', 'metric_value': 0.0, 'unit': '%', 'judgement': None, 'ng_breakdown': {}, 'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!F57:L57']},
    {'result_id': 'res_7', 'condition_id': 'cond_3', 'measurement_type': 'Function', 'condition_group': 'Test level 3', 'date': '2026-02-26', 'line': '', 'input_count': 20, 'ok_count': 20, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Function NG rate', 'metric_value': 0.0, 'unit': '%', 'judgement': None, 'ng_breakdown': {}, 'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!F59:L59']},
    {'result_id': 'res_8', 'condition_id': 'cond_4', 'measurement_type': 'Function', 'condition_group': 'Normal', 'date': '2026-02-26', 'line': '', 'input_count': 100, 'ok_count': 96, 'ng_count': 4, 'ng_rate_decimal': 0.04, 'ng_rate_percent': 4.0, 'metric_name': 'Function NG rate', 'metric_value': 4.0, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG Sigma SPL': {'count': 0, 'rate': 0.0}, 'NG Hearing Noise': {'count': 4, 'rate': 4.0}, 'NG Hearing Touch': {'count': 0, 'rate': 0.0}}, 'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!F61:L61']},
    {'result_id': 'res_9', 'condition_id': 'cond_1', 'measurement_type': 'Drop test FINAL auto', 'condition_group': 'Test level 1', 'date': '2026-02-26', 'line': '', 'input_count': 8, 'ok_count': 7, 'ng_count': 1, 'ng_rate_decimal': None, 'ng_rate_percent': None, 'metric_name': 'Drop test FINAL auto', 'metric_value': None, 'unit': None, 'judgement': 'FAIL', 'ng_breakdown': {'NG by hearing noise (no PT+MG separate)': {'count': 1}}, 'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!E47:M47']},
    {'result_id': 'res_10', 'condition_id': 'cond_2', 'measurement_type': 'Drop test FINAL auto', 'condition_group': 'Test level 2', 'date': '2026-02-26', 'line': '', 'input_count': 8, 'ok_count': 6, 'ng_count': 2, 'ng_rate_decimal': None, 'ng_rate_percent': None, 'metric_name': 'Drop test FINAL auto', 'metric_value': None, 'unit': None, 'judgement': 'FAIL', 'ng_breakdown': {'NG by hearing noise (no PT+MG separate)': {'count': 2}}, 'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!E48:M48']},
    {'result_id': 'res_11', 'condition_id': 'cond_3', 'measurement_type': 'Drop test FINAL auto', 'condition_group': 'Test level 3', 'date': '2026-02-26', 'line': '', 'input_count': 8, 'ok_count': 7, 'ng_count': 1, 'ng_rate_decimal': None, 'ng_rate_percent': None, 'metric_name': 'Drop test FINAL auto', 'metric_value': None, 'unit': None, 'judgement': 'FAIL', 'ng_breakdown': {'NG by hearing noise (no PT+MG separate)': {'count': 1}}, 'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!E49:M49']},
    {'result_id': 'res_12', 'condition_id': 'cond_4', 'measurement_type': 'Drop test FINAL auto', 'condition_group': 'Normal', 'date': '2026-02-26', 'line': '', 'input_count': 8, 'ok_count': 8, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Drop test FINAL auto', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS', 'ng_breakdown': {}, 'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!E50:M50']},
  ],
  'conclusions': [
    {'conclusion_id': 'concl_1', 'topic': 'NG visual at sub2',
     'statement_from_report': 'Result NG visual YOKE lever1-lever2-lever3 all OK same normal, no NG MG+PT separate.',
     'normalized_interpretation': 'Test level 1/2/3 NG Visual rate 0.0% matches Normal 0.0% (no MG+PT separation observed); plate from MYUNGJIN different color does not introduce visible bonding defects at sub2.',
     'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!A64']},
    {'conclusion_id': 'concl_2', 'topic': 'Decap bonding',
     'statement_from_report': 'Result decap bonding type lever1-lever2-lever3 all OK same normal.',
     'normalized_interpretation': 'Decap bond NG 0/8 at every test level and Normal: bond strength of MYUNGJIN plate sample is equivalent to baseline.',
     'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!A65']},
    {'conclusion_id': 'concl_3', 'topic': 'Drop test FINAL auto',
     'statement_from_report': 'Drop test FINAL auto type lever1/2/3 happen NG hearing noise. But reason not by PT.',
     'normalized_interpretation': 'Drop test FINAL auto failed for Test L1 (1/8), L2 (2/8), L3 (1/8) while Normal passed 8/8; NG mode is hearing noise without PT+MG separation, suggesting particle/contamination rather than plate adhesion.',
     'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!A67', 'Test!C51']},
    {'conclusion_id': 'concl_4', 'topic': 'Function',
     'statement_from_report': 'Result check function all 3 type test NG same normal.',
     'normalized_interpretation': 'Function NG rate Test L1 2.4%, L2 0.0%, L3 0.0% vs Normal 4.0%: every test level is equal or better than baseline (L1 = 2.4/4.0 - 1 = -40.0% improved, L2/L3 0.0% improved 100% vs normal). Dominant NG mode at L1 is split between SPL (1.2%) and Noise (1.2%); Normal is dominated by Noise (4.0%).',
     'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!A68', 'Test!F55:L61']},
  ],
  'troubleshooting_index': {
    'defect_name': 'NG Hearing Noise',
    'when_user_asks': ['plate supplier change', 'NG hearing noise on drop test', 'MYUNGJIN plate qualification'],
    'suggested_checks': [
      {'hint_id': 'hint_1', 'check_item': 'Check particle/contamination on plate surface before assembly', 'reason': 'Drop test FINAL auto produced hearing noise on every test level (1/8, 2/8, 1/8) while Normal passed 8/8; NG samples did not show PT+MG separation, so particle-induced noise is the likely cause.', 'evidence_strength': 'medium', 'related_process': 'Sub2 / Drop test', 'related_part': 'Plate', 'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!C51']},
      {'hint_id': 'hint_2', 'check_item': 'Compare semi-yoke manual drop test vs final auto drop test consistency', 'reason': 'Drop test semi-yoke manual passed 8/8 for all test levels while final auto failed; the difference points to handling, jig, or assembly stage rather than plate material.', 'evidence_strength': 'medium', 'related_process': 'Drop test', 'related_part': 'YOKE', 'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!E39:M42', 'Test!E47:M50']},
      {'hint_id': 'hint_3', 'check_item': 'Re-test function NG rate vs Normal baseline on larger sample', 'reason': 'Test L1 function NG rate 2.4% (n=84) vs Normal 4.0% (n=100) suggests equivalence or slight improvement, but L2/L3 sample size (n=18, n=20) is small.', 'evidence_strength': 'low', 'related_process': 'Function test', 'related_part': 'Plate', 'source_file': name01, 'sheet_name': 'Test', 'source_cells': ['Test!F55:L61']},
    ],
    'limitations': ['Test L2/L3 sample size is small (18-28 units); confidence in function equivalence is limited.', 'Drop test FINAL auto root cause is described as "not PT" but no quantitative particle/contamination data is recorded.']
  },
  'ai_extraction_log': {
    'confidence': 0.7,
    'assumptions': ['"Test level 1/2/3" represents three lots or coating shades of the MYUNGJIN different-color plate.', 'Normal row in each table is the same-event baseline.'],
    'warnings': ['Drop test FINAL auto root cause is not numerically confirmed in the report.', 'No final decision section is filled in.'],
    'decision_rationale': 'Visual NG, decap, and semi-manual drop test all match Normal. Function NG rate at L1 2.4% vs Normal 4.0% (-40.0% improved). Drop test FINAL auto shows hearing noise on every test level while Normal passes, but the report explicitly states it is not PT-related; treat as a process/particle alarm, not a plate-material rejection.'
  }
}

tr_ko_01 = {
  'document': {'title': 'TIU C11-20 MYUNGJIN 색상 차이 Plate 100% 적용 시험 리포트', 'purpose': '공급업체 MYUNGJIN의 색상이 다른 Plate를 사용할 수 있는지 확인한다.',
               'content': ['Sub2에서 반제품 제작 후 NG 공정 확인', 'PT+MG 본드 디캡 확인', '자동/수동 Drop test', 'AI bonding 검출률 확인', '메인 라인 투입 후 Function 확인']},
  'conclusions': {
    'concl_1': {'topic': 'Sub2 NG visual', 'statement_from_report': 'YOKE lever1~3 NG visual 모두 normal과 같고 MG+PT separate 없음.', 'normalized_interpretation': 'Test L1/L2/L3 NG Visual 0.0%, Normal도 0.0%로 동일. MYUNGJIN 다른 색상 Plate에서도 Sub2에서 외관 본드 결함은 발생하지 않음.'},
    'concl_2': {'topic': 'Decap bonding', 'statement_from_report': 'Decap bonding L1~3 모두 normal과 동일.', 'normalized_interpretation': '각 레벨/Normal 모두 8/8 OK. MYUNGJIN Plate 본드 강도는 baseline과 동일.'},
    'concl_3': {'topic': 'Drop test FINAL auto', 'statement_from_report': 'Drop test FINAL auto L1/2/3에서 NG hearing noise 발생, 단 PT 원인 아님.', 'normalized_interpretation': 'L1 1/8, L2 2/8, L3 1/8 fail, Normal 8/8 pass. NG mode는 hearing noise이며 PT+MG separation 없음 → particle/오염 의심.'},
    'concl_4': {'topic': 'Function', 'statement_from_report': '3개 type 모두 NG 같음.', 'normalized_interpretation': 'Function NG rate L1 2.4%, L2 0.0%, L3 0.0% vs Normal 4.0% (L1: -40.0% 개선). L1 dominant NG: SPL 1.2%, Noise 1.2%; Normal dominant: Noise 4.0%.'},
  },
  'hints': {
    'hint_1': {'check_item': 'Plate 표면 particle/오염 점검', 'reason': 'Drop test FINAL auto L1/2/3 모두 hearing noise NG, Normal은 OK. PT+MG separation 없음 → particle 의심.'},
    'hint_2': {'check_item': 'Semi YOKE manual drop vs FINAL auto drop 비교', 'reason': 'Semi manual은 전 레벨 8/8 통과, FINAL auto만 NG. 조립 단계·jig·취급 요인 점검.'},
    'hint_3': {'check_item': '더 큰 sample로 Function NG rate 재확인', 'reason': 'L1 84개, L2 18개, L3 20개로 표본이 작음.'},
  },
  'log': {'assumptions': ['Test level 1~3은 MYUNGJIN 다른 색상 Plate의 lot/색상 단계로 가정', 'Normal 행이 같은 이벤트의 baseline'],
          'warnings': ['Drop FINAL auto NG의 정량적 원인 데이터 없음', 'Decision 본문 미완'],
          'decision_rationale': 'Visual·Decap·Semi manual drop 모두 Normal 수준. Function L1 2.4% < Normal 4.0% (-40.0% 개선). FINAL auto NG는 PT 원인이 아니라고 명시 → particle/공정 alarm으로 다루며 Plate 자체는 fail하지 않음.'}
}
tr_en_01 = {
  'document': {'title': 'TIU C11-20 Report Test Plate Supplier MYUNGJIN Different Color 100%', 'purpose': 'Test whether plate from supplier MYUNGJIN with different color can be used.',
               'content': ['Make semi at sub2 and check NG process', 'Decap check bond PT+MG', 'Drop test Auto/Manual', 'Check AI bonding detection rate', 'Input main line and check function']},
  'conclusions': {
    'concl_1': {'topic': 'Sub2 NG visual', 'statement_from_report': 'YOKE lever1-3 NG visual all OK same normal, no MG+PT separate.', 'normalized_interpretation': 'Test L1/L2/L3 NG Visual 0.0% matches Normal 0.0%; MYUNGJIN different color plate produces no visible bonding defects at sub2.'},
    'concl_2': {'topic': 'Decap bonding', 'statement_from_report': 'Decap bonding lever1-3 all OK same normal.', 'normalized_interpretation': '8/8 OK at every level and Normal; bond strength matches baseline.'},
    'concl_3': {'topic': 'Drop test FINAL auto', 'statement_from_report': 'Drop test FINAL auto type lever1/2/3 happen NG hearing noise, reason not by PT.', 'normalized_interpretation': 'L1 1/8, L2 2/8, L3 1/8 fail, Normal 8/8 pass; NG mode hearing noise without PT+MG separation suggests particle/contamination.'},
    'concl_4': {'topic': 'Function', 'statement_from_report': 'Function all 3 types same as normal.', 'normalized_interpretation': 'L1 2.4%, L2 0.0%, L3 0.0% vs Normal 4.0% (L1: -40.0% improved). L1 dominant NG: SPL 1.2%, Noise 1.2%; Normal dominant: Noise 4.0%.'},
  },
  'hints': {
    'hint_1': {'check_item': 'Check particle/contamination on plate surface before assembly', 'reason': 'Drop test FINAL auto L1/2/3 all show hearing noise; Normal passes. No PT+MG separation -> particle suspected.'},
    'hint_2': {'check_item': 'Compare semi-yoke manual drop vs final auto drop', 'reason': 'Semi manual passes 8/8 at all levels; final auto fails only -> assembly stage / jig / handling.'},
    'hint_3': {'check_item': 'Re-confirm Function NG rate on larger sample', 'reason': 'L1 n=84, L2 n=18, L3 n=20 sample size is small.'},
  },
  'log': {'assumptions': ['Test level 1-3 represent lots or color shades of MYUNGJIN different color plate.', 'Normal row is same-event baseline.'],
          'warnings': ['Drop FINAL auto NG root cause is not numerically confirmed.', 'Decision section is incomplete.'],
          'decision_rationale': 'Visual, decap and semi-manual drop all match Normal. Function L1 2.4% vs Normal 4.0% = -40.0% improved. Drop FINAL auto failures are explicitly described as not-PT, treat as process particle alarm rather than plate rejection.'}
}
tr_vi_01 = {
  'document': {'title': 'TIU C11-20 Báo cáo test Plate nhà cung cấp MYUNGJIN khác màu 100%', 'purpose': 'Kiểm tra Plate khác màu của nhà cung cấp MYUNGJIN có thể dùng được hay không.',
               'content': ['Làm semi tại sub2 và kiểm tra NG process', 'Decap kiểm tra bond PT+MG', 'Drop test Auto/Manual', 'Kiểm tra tỷ lệ phát hiện AI bonding', 'Cho vào main line và kiểm tra Function']},
  'conclusions': {
    'concl_1': {'topic': 'Sub2 NG visual', 'statement_from_report': 'NG visual YOKE lever1-3 đều OK giống normal, không có MG+PT separate.', 'normalized_interpretation': 'NG Visual Test L1/L2/L3 0.0% bằng Normal 0.0%; Plate MYUNGJIN khác màu không gây NG bond ngoại quan tại sub2.'},
    'concl_2': {'topic': 'Decap bonding', 'statement_from_report': 'Decap bonding lever1-3 đều OK giống normal.', 'normalized_interpretation': '8/8 OK ở mọi level và Normal; lực bond bằng baseline.'},
    'concl_3': {'topic': 'Drop test FINAL auto', 'statement_from_report': 'Drop test FINAL auto lever1/2/3 phát sinh NG hearing noise, lý do không phải PT.', 'normalized_interpretation': 'L1 1/8, L2 2/8, L3 1/8 fail, Normal 8/8 pass; NG mode hearing noise nhưng không có PT+MG separate → nghi particle/nhiễm bẩn.'},
    'concl_4': {'topic': 'Function', 'statement_from_report': 'Function 3 type test NG giống normal.', 'normalized_interpretation': 'L1 2.4%, L2 0.0%, L3 0.0% so Normal 4.0% (L1: -40.0% cải thiện). NG chính của L1: SPL 1.2%, Noise 1.2%; Normal: Noise 4.0%.'},
  },
  'hints': {
    'hint_1': {'check_item': 'Kiểm tra particle/nhiễm bẩn trên bề mặt Plate trước khi assembly', 'reason': 'Drop test FINAL auto L1/2/3 đều NG hearing noise, Normal OK. Không có PT+MG separate → nghi particle.'},
    'hint_2': {'check_item': 'So sánh semi YOKE manual drop vs FINAL auto drop', 'reason': 'Semi manual 8/8 OK cho mọi level; FINAL auto mới NG → vấn đề ở assembly / jig / thao tác.'},
    'hint_3': {'check_item': 'Tái xác nhận Function NG rate trên sample lớn hơn', 'reason': 'L1 n=84, L2 n=18, L3 n=20 sample nhỏ.'},
  },
  'log': {'assumptions': ['Test level 1-3 là các lot/màu của Plate MYUNGJIN khác màu.', 'Hàng Normal là baseline cùng sự kiện.'],
          'warnings': ['Không có dữ liệu định lượng cho NG Drop FINAL auto.', 'Phần Decision chưa hoàn thành.'],
          'decision_rationale': 'Visual, decap, semi manual drop đều giống Normal. Function L1 2.4% vs Normal 4.0% = -40.0% cải thiện. NG Drop FINAL auto được ghi không do PT → coi như cảnh báo particle/quá trình, không loại Plate.'}
}
results[name01] = (result01, tr_ko_01, tr_en_01, tr_vi_01)

# =====================================================================
# DS02: 26. TIU L5S3-01 R Report test find reason NG BAKO high 2025.12.14
# =====================================================================
name02 = '26. TIU L5S3-01 R Report test find reason NG BAKO high 2025.12.14'

result02 = {
  'schema_version': '0.1',
  'document': {
    'document_id': '', 'source_file': name02, 'source_sheet': 'Test',
    'title': 'TIU L5S3-01 [R] Report Test Find Reason Improve NG Function',
    'model': 'TIU L5S3-01 [R]', 'report_date': '2025-12-14',
    'department': 'ME', 'marker': 'Thao', 'line': 'L',
    'report_type': 'ng_without_baseline',
    'primary_defect': {'canonical_name': 'NG BAKO Function', 'aliases_in_document': ['NG function Bako', 'NG BAKO FRF', 'NG BAKO FRF+SPL']},
    'related_defects': ['NG BAKO FRF', 'NG BAKO FRF+SPL', 'NG BAKO THD', 'NG BAKO No sound'],
    'parts': ['VP', 'Coil'], 'processes': ['Dry UV VP/Coil', 'Bonding'],
    'purpose': 'NG BAKO function rate is around 50%; test conditions to reduce NG and find root cause.',
    'content': ['Test dry UV temperature for VP/Coil at MAX and MIN', 'Test running on bonding line of same L line', 'Check Function NG rate'],
    'source_cells': {'title': ['Test!B1'], 'date': ['Test!N2'], 'purpose': ['Test!A4'], 'content': ['Test!A6:A8']}
  },
  'test_conditions': [
    {'condition_id': 'cond_1', 'condition_group': 'Test dry UV VP/Coil MAX (Peak 1350 / Total 6482)', 'line': '', 'process': 'Dry UV VP/Coil', 'changed_factor': 'Dry UV temperature MAX', 'before_value': None, 'after_value': 'MAX (Peak 1350, Total 6482)', 'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': 'MAX', 'pressure': None, 'bond_amount': None, 'uv_energy': 'Peak 1350 / Total 6482', 'source_file': name02, 'sheet_name': 'Test', 'source_cells': ['Test!D17']},
    {'condition_id': 'cond_2', 'condition_group': 'Test dry UV VP/Coil MAX (Peak 1252 / Total 5537)', 'line': '', 'process': 'Dry UV VP/Coil', 'changed_factor': 'Dry UV temperature MAX (lower energy)', 'before_value': None, 'after_value': 'MAX (Peak 1252, Total 5537)', 'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': 'MAX', 'pressure': None, 'bond_amount': None, 'uv_energy': 'Peak 1252 / Total 5537', 'source_file': name02, 'sheet_name': 'Test', 'source_cells': ['Test!D19']},
    {'condition_id': 'cond_3', 'condition_group': 'Test using bonding line VP same L line', 'line': 'L', 'process': 'Bonding', 'changed_factor': 'Bonding line (VP same L line)', 'before_value': None, 'after_value': 'L line VP bonding', 'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name02, 'sheet_name': 'Test', 'source_cells': ['Test!D21']},
  ],
  'results': [
    {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Function', 'condition_group': 'Test dry UV VP/Coil MAX (Peak 1350 / Total 6482)', 'date': '2025-12-14', 'line': '', 'input_count': 393, 'ok_count': 179, 'ng_count': 214, 'ng_rate_decimal': 0.545, 'ng_rate_percent': 54.5, 'metric_name': 'NG BAKO Function rate', 'metric_value': 54.5, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG BAKO FRF': {'count': 188, 'rate': 47.8}, 'NG BAKO FRF+SPL': {'count': 26, 'rate': 6.6}, 'NG BAKO THD': {'count': 0, 'rate': 0.0}, 'NG BAKO No sound': {'count': 0, 'rate': 0.0}}, 'source_file': name02, 'sheet_name': 'Test', 'source_cells': ['Test!F17:N17']},
    {'result_id': 'res_2', 'condition_id': 'cond_2', 'measurement_type': 'Function', 'condition_group': 'Test dry UV VP/Coil MAX (Peak 1252 / Total 5537)', 'date': '2025-12-14', 'line': '', 'input_count': 517, 'ok_count': 170, 'ng_count': 347, 'ng_rate_decimal': 0.671, 'ng_rate_percent': 67.1, 'metric_name': 'NG BAKO Function rate', 'metric_value': 67.1, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG BAKO FRF': {'count': 283, 'rate': 54.7}, 'NG BAKO FRF+SPL': {'count': 64, 'rate': 12.4}, 'NG BAKO THD': {'count': 0, 'rate': 0.0}, 'NG BAKO No sound': {'count': 0, 'rate': 0.0}}, 'source_file': name02, 'sheet_name': 'Test', 'source_cells': ['Test!F19:N19']},
    {'result_id': 'res_3', 'condition_id': 'cond_3', 'measurement_type': 'Function', 'condition_group': 'Test using bonding line VP same L line', 'date': '2025-12-14', 'line': 'L', 'input_count': 417, 'ok_count': 159, 'ng_count': 258, 'ng_rate_decimal': 0.619, 'ng_rate_percent': 61.9, 'metric_name': 'NG BAKO Function rate', 'metric_value': 61.9, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG BAKO FRF': {'count': 242, 'rate': 58.0}, 'NG BAKO FRF+SPL': {'count': 16, 'rate': 3.8}, 'NG BAKO THD': {'count': 0, 'rate': 0.0}, 'NG BAKO No sound': {'count': 0, 'rate': 0.0}}, 'source_file': name02, 'sheet_name': 'Test', 'source_cells': ['Test!F21:N21']},
  ],
  'conclusions': [
    {'conclusion_id': 'concl_1', 'topic': 'NG BAKO not reduced',
     'statement_from_report': 'NG bako not reduce.',
     'normalized_interpretation': 'All three tested conditions still produced NG BAKO at >=54%: UV MAX(Peak1350/Total6482) 54.5%, UV MAX(Peak1252/Total5537) 67.1%, L line VP bonding 61.9%. FRF dominates every condition (47.8% / 54.7% / 58.0%). No same-event baseline row was provided; rank by absolute NG rate.',
     'source_file': name02, 'sheet_name': 'Test', 'source_cells': ['Test!A24']},
  ],
  'troubleshooting_index': {
    'defect_name': 'NG BAKO Function (FRF dominant)',
    'when_user_asks': ['NG BAKO high rate', 'FRF NG dominant', 'UV temperature vs function NG'],
    'suggested_checks': [
      {'hint_id': 'hint_1', 'check_item': 'Inspect VP/Coil drying profile and check UV peak/total energy vs spec', 'reason': 'Both UV MAX conditions (Peak1350/Total6482 and Peak1252/Total5537) failed at 54.5% and 67.1% respectively; raising UV did not reduce NG BAKO.', 'evidence_strength': 'medium', 'related_process': 'Dry UV VP/Coil', 'related_part': 'VP/Coil', 'source_file': name02, 'sheet_name': 'Test', 'source_cells': ['Test!F17:N19']},
      {'hint_id': 'hint_2', 'check_item': 'Inspect FRF tuning/measurement chain (jig, mic position, calibration)', 'reason': 'FRF alone is 47.8-58.0% of failures in every condition; the dominant failure mode does not change with UV or bonding-line change, suggesting it is a measurement/structural issue rather than UV curing.', 'evidence_strength': 'high', 'related_process': 'Function test', 'related_part': 'VP/Coil/Yoke', 'source_file': name02, 'sheet_name': 'Test', 'source_cells': ['Test!G17:G21']},
      {'hint_id': 'hint_3', 'check_item': 'Verify L-line bonding line equivalence to original line', 'reason': 'Switching to L-line bonding gave 61.9% NG vs UV MAX 54.5% in same report; line swap did not improve NG.', 'evidence_strength': 'low', 'related_process': 'Bonding', 'related_part': 'VP', 'source_file': name02, 'sheet_name': 'Test', 'source_cells': ['Test!F21:N21']},
    ],
    'limitations': ['No Normal/Baseline row in this report; comparisons are between test conditions only.', 'Decision section only states "NG bako not reduce" without root-cause statement.']
  },
  'ai_extraction_log': {
    'confidence': 0.65,
    'assumptions': ['"BAKO" interpreted as the model-specific function test station / criterion.'],
    'warnings': ['No same-event baseline row -> classified as ng_without_baseline; no improvement/worsening claim.', 'Two rows are both labelled "Test dry UV VP/COI MAX" with different UV energy values; treated as two distinct conditions.'],
    'decision_rationale': 'All three conditions remain at very high NG BAKO rate (54.5%-67.1%). FRF dominates every condition. Without baseline, ranking is absolute: condition cond_1 has lowest NG BAKO 54.5%, cond_2 worst 67.1%. Root cause is most likely FRF-side (measurement or structural).'
  }
}

tr_ko_02 = {
  'document': {'title': 'TIU L5S3-01 [R] NG Function 개선/원인 찾기 시험 리포트', 'purpose': 'NG BAKO function rate ~50%로 매우 높음. NG 감소 조건 탐색 및 원인 파악.',
               'content': ['VP/Coil Dry UV 온도 MAX/MIN 조건 시험', '동일 L 라인 bonding line 사용', 'Function NG rate 확인']},
  'conclusions': {'concl_1': {'topic': 'NG BAKO 미감소', 'statement_from_report': 'NG bako 감소하지 않음.', 'normalized_interpretation': '3개 조건 모두 NG BAKO 54% 이상: UV MAX(P1350/T6482) 54.5%, UV MAX(P1252/T5537) 67.1%, L 라인 VP bonding 61.9%. 모든 조건에서 FRF가 우세(47.8%/54.7%/58.0%). 같은 이벤트 baseline 없음, 절대값 순위로만 평가.'}},
  'hints': {
    'hint_1': {'check_item': 'VP/Coil dry UV peak/total energy 점검 (스펙 대비)', 'reason': 'UV MAX 두 조건 모두 54.5%, 67.1% NG로 UV 증가가 효과 없음.'},
    'hint_2': {'check_item': 'FRF 측정 chain (jig, mic, calibration) 점검', 'reason': 'FRF 단독이 47.8~58.0%로 모든 조건에서 우세; UV/라인 변경에 영향받지 않음 → 측정/구조 문제 가능.'},
    'hint_3': {'check_item': 'L-line bonding line 등가성 확인', 'reason': 'L 라인 61.9%로 UV MAX 54.5%보다도 높음.'},
  },
  'log': {'assumptions': ['"BAKO"는 해당 모델 function test 단계/기준으로 가정.'],
          'warnings': ['Normal/baseline 행 없음 → ng_without_baseline 분류, 개선/악화 단정 없음.', '두 행이 같은 "Test dry UV VP/COI MAX" 라벨이나 UV 값이 달라 별개 조건으로 처리.'],
          'decision_rationale': '3 조건 모두 54.5~67.1%로 매우 높음. FRF가 모든 조건에서 우세. Baseline 없이 절대값 비교만 가능 → cond_1 54.5%로 최저, cond_2 67.1%로 최고. 원인은 FRF 측면(측정 또는 구조) 가능성 높음.'}
}
tr_en_02 = {
  'document': {'title': 'TIU L5S3-01 [R] Report Test Find Reason Improve NG Function', 'purpose': 'NG BAKO function rate ~50% is very high; test conditions and find root cause.',
               'content': ['Test dry UV temperature for VP/Coil at MAX and MIN', 'Run on bonding line of same L line', 'Check Function NG rate']},
  'conclusions': {'concl_1': {'topic': 'NG BAKO not reduced', 'statement_from_report': 'NG bako not reduce.', 'normalized_interpretation': 'All three conditions remain at >=54% NG BAKO: UV MAX(P1350/T6482) 54.5%, UV MAX(P1252/T5537) 67.1%, L line VP bonding 61.9%. FRF dominates each condition (47.8%/54.7%/58.0%). No same-event baseline; rank by absolute NG.'}},
  'hints': {
    'hint_1': {'check_item': 'Inspect VP/Coil dry UV peak/total energy vs spec', 'reason': 'Both UV MAX conditions failed (54.5% and 67.1%); higher UV does not reduce NG BAKO.'},
    'hint_2': {'check_item': 'Inspect FRF measurement chain (jig, mic, calibration)', 'reason': 'FRF alone is 47.8-58.0% of NG in every condition; unaffected by UV or bonding-line change -> measurement/structural issue likely.'},
    'hint_3': {'check_item': 'Verify L-line bonding line equivalence', 'reason': 'L-line gave 61.9% vs UV MAX 54.5%; line swap did not help.'},
  },
  'log': {'assumptions': ['"BAKO" interpreted as the model-specific function test station/criterion.'],
          'warnings': ['No same-event baseline row -> classified ng_without_baseline; no improvement/worsening claim.', 'Two rows share label "Test dry UV VP/COI MAX" but UV energy differs; treated as two distinct conditions.'],
          'decision_rationale': 'All conditions remain at 54.5-67.1% NG. FRF dominates every condition. Without baseline, absolute rank: cond_1 lowest 54.5%, cond_2 worst 67.1%. Root cause is likely on FRF measurement/structural side.'}
}
tr_vi_02 = {
  'document': {'title': 'TIU L5S3-01 [R] Báo cáo test tìm nguyên nhân và cải thiện NG Function', 'purpose': 'NG BAKO function ~50% rất cao; test các điều kiện để giảm NG và tìm nguyên nhân.',
               'content': ['Test nhiệt độ dry UV VP/Coil MAX & MIN', 'Chạy bonding line cùng L line', 'Kiểm tra NG rate Function']},
  'conclusions': {'concl_1': {'topic': 'NG BAKO không giảm', 'statement_from_report': 'NG bako not reduce.', 'normalized_interpretation': 'Cả 3 điều kiện vẫn >=54%: UV MAX(P1350/T6482) 54.5%, UV MAX(P1252/T5537) 67.1%, L line VP bonding 61.9%. FRF dominant ở mọi điều kiện (47.8%/54.7%/58.0%). Không có baseline cùng sự kiện; xếp hạng tuyệt đối.'}},
  'hints': {
    'hint_1': {'check_item': 'Kiểm tra peak/total UV của VP/Coil so với spec', 'reason': 'Cả hai UV MAX đều fail 54.5% và 67.1%; tăng UV không giảm NG.'},
    'hint_2': {'check_item': 'Kiểm tra chuỗi đo FRF (jig, mic, hiệu chuẩn)', 'reason': 'FRF đơn lẻ chiếm 47.8-58.0% NG ở mọi điều kiện; không bị ảnh hưởng bởi UV hay bonding line → vấn đề đo lường/cấu trúc.'},
    'hint_3': {'check_item': 'Xác minh L-line bonding tương đương line gốc', 'reason': 'L line 61.9% so với UV MAX 54.5%; đổi line không cải thiện.'},
  },
  'log': {'assumptions': ['"BAKO" hiểu là tiêu chí/station function test của model này.'],
          'warnings': ['Không có hàng Normal/baseline -> phân loại ng_without_baseline; không kết luận cải thiện/xấu đi.', 'Hai hàng cùng nhãn "Test dry UV VP/COI MAX" nhưng UV khác nhau; coi là 2 điều kiện riêng.'],
          'decision_rationale': 'Cả 3 điều kiện 54.5-67.1% NG. FRF chiếm ưu thế ở mọi điều kiện. Không có baseline, chỉ xếp hạng tuyệt đối: cond_1 thấp nhất 54.5%, cond_2 cao nhất 67.1%. Nguyên nhân nhiều khả năng ở phía FRF (đo lường/cấu trúc).'}
}
results[name02] = (result02, tr_ko_02, tr_en_02, tr_vi_02)

# =====================================================================
# DS03: 27. BRS-161014 DT Report test VP mold #7,9 add 0.05mm  date 12.3.2024
# Two sheets ('7.3 (2)' = 3/12-3/15, '5.3' = 3/5-3/7)
# =====================================================================
name03 = '27. BRS-161014 DT Report test VP mold #7,9 add 0.05mm  date 12.3.2024'

result03 = {
  'schema_version': '0.1',
  'document': {
    'document_id': '', 'source_file': name03, 'source_sheet': '5.3 + 7.3 (2)',
    'title': 'BRS-161014DT Report Test VP Mold #7,9 Add 0.05mm',
    'model': 'BRS-161014DT', 'report_date': '2024-03-12',
    'department': 'ME', 'marker': 'Le / Thuy', 'line': 'C2-3A / C2-3B / E2',
    'report_type': 'normal_comparison',
    'primary_defect': {'canonical_name': 'NG VP Bending', 'aliases_in_document': ['Vision Laze Cutting NG VP Bending']},
    'related_defects': ['VP/CD Offset NG', 'Glue Not Enough', 'NG Hearing Noise', 'NG Hearing Touch'],
    'parts': ['VP', 'CD', 'Mold #7', 'Mold #9'], 'processes': ['Laze cutting', 'VP/CD vision', 'Function'],
    'purpose': 'Test VP mold #7 and #9 with 0.05mm added in center part; verify improvement vs Normal mold (#4, #6, #10).',
    'content': ['Check material and semi VP after laze cutting (VP bending)', 'Make semi and check NG rate of vision VP+CD', 'Make final and check NG rate of function'],
    'source_cells': {'title': ['7.3 (2)!B1'], 'date': ['7.3 (2)!N2', '5.3!N62'], 'purpose': ['7.3 (2)!A4'], 'content': ['7.3 (2)!A6:A8']}
  },
  'test_conditions': [
    {'condition_id': 'cond_1', 'condition_group': 'Test VP mold #7', 'line': 'C2-3A / C2-3B / E2', 'process': 'Laze cutting + Sub1 vision + Function', 'changed_factor': 'VP mold #7 with +0.05mm at center', 'before_value': 'Normal VP mold #4/#6/#10', 'after_value': 'VP mold #7 +0.05mm', 'unit': 'mm', 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name03, 'sheet_name': '5.3 + 7.3 (2)', 'source_cells': ['7.3 (2)!A8']},
    {'condition_id': 'cond_2', 'condition_group': 'Test VP mold #9', 'line': 'C2-3A / C2-3B / E2', 'process': 'Laze cutting + Sub1 vision + Function', 'changed_factor': 'VP mold #9 with +0.05mm at center', 'before_value': 'Normal VP mold #4/#6/#10', 'after_value': 'VP mold #9 +0.05mm', 'unit': 'mm', 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name03, 'sheet_name': '5.3 + 7.3 (2)', 'source_cells': ['7.3 (2)!A8']},
    {'condition_id': 'cond_3', 'condition_group': 'Normal VP mold #4/#6/#10', 'line': 'C2-3B / E2', 'process': 'Laze cutting + Sub1 vision + Function', 'changed_factor': 'Baseline (Normal mold)', 'before_value': None, 'after_value': 'Normal VP mold #4/#6/#10', 'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!C75']},
  ],
  'results': [
    # Sub1 vision laze cutting (5.3 sheet - has baseline)
    {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Vision Laze Cutting (VP bending)', 'condition_group': 'Test VP mold #7 (5.3 sheet total)', 'date': '2024-03-05~07', 'line': 'C2-3B + E2', 'input_count': 4393, 'ok_count': 4037, 'ng_count': 356, 'ng_rate_decimal': 0.081, 'ng_rate_percent': 8.1, 'metric_name': 'NG VP Bending rate', 'metric_value': 8.1, 'unit': '%', 'judgement': 'FAIL', 'ng_breakdown': {'NG VP Bending': {'count': 356, 'rate': 8.1}}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!F85:I85']},
    {'result_id': 'res_2', 'condition_id': 'cond_2', 'measurement_type': 'Vision Laze Cutting (VP bending)', 'condition_group': 'Test VP mold #9 (5.3 sheet total)', 'date': '2024-03-05~07', 'line': 'C2-3B + E2', 'input_count': 4393, 'ok_count': 3829, 'ng_count': 564, 'ng_rate_decimal': 0.128, 'ng_rate_percent': 12.8, 'metric_name': 'NG VP Bending rate', 'metric_value': 12.8, 'unit': '%', 'judgement': 'FAIL', 'ng_breakdown': {'NG VP Bending': {'count': 564, 'rate': 12.8}}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!F86:I86']},
    {'result_id': 'res_3', 'condition_id': 'cond_3', 'measurement_type': 'Vision Laze Cutting (VP bending)', 'condition_group': 'Normal VP mold #4/#6/#10 (5.3 sheet total)', 'date': '2024-03-05~07', 'line': 'C2-3B + E2', 'input_count': 5156, 'ok_count': 4969, 'ng_count': 187, 'ng_rate_decimal': 0.036, 'ng_rate_percent': 3.6, 'metric_name': 'NG VP Bending rate', 'metric_value': 3.6, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG VP Bending': {'count': 187, 'rate': 3.6}}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!F87:I87']},
    # 7.3(2) Sub1 totals (no baseline)
    {'result_id': 'res_4', 'condition_id': 'cond_1', 'measurement_type': 'Vision Laze Cutting (VP bending)', 'condition_group': 'Test VP mold #7 (7.3(2) total)', 'date': '2024-03-12~13', 'line': 'C2-3A', 'input_count': 11279, 'ok_count': 11262, 'ng_count': 17, 'ng_rate_decimal': 0.0015, 'ng_rate_percent': 0.2, 'metric_name': 'NG VP Bending rate', 'metric_value': 0.2, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG VP Bending': {'count': 17, 'rate': 0.2}}, 'source_file': name03, 'sheet_name': '7.3 (2)', 'source_cells': ['7.3 (2)!E21:I21']},
    {'result_id': 'res_5', 'condition_id': 'cond_2', 'measurement_type': 'Vision Laze Cutting (VP bending)', 'condition_group': 'Test VP mold #9 (7.3(2) total)', 'date': '2024-03-12~13', 'line': 'C2-3A', 'input_count': 16260, 'ok_count': 16183, 'ng_count': 77, 'ng_rate_decimal': 0.0047, 'ng_rate_percent': 0.5, 'metric_name': 'NG VP Bending rate', 'metric_value': 0.5, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG VP Bending': {'count': 77, 'rate': 0.5}}, 'source_file': name03, 'sheet_name': '7.3 (2)', 'source_cells': ['7.3 (2)!E22:I22']},
    # VP/CD vision (5.3 totals)
    {'result_id': 'res_6', 'condition_id': 'cond_1', 'measurement_type': 'Vision VP/CD', 'condition_group': 'Test VP mold #7 (5.3 sheet)', 'date': '2024-03-05~06', 'line': 'C2 + E2', 'input_count': 3952, 'ok_count': 3951, 'ng_count': 1, 'ng_rate_decimal': 0.0003, 'ng_rate_percent': 0.0, 'metric_name': 'NG VP/CD rate', 'metric_value': 0.0, 'unit': '%', 'judgement': None, 'ng_breakdown': {'VP/CD Offset NG': {'count': 1, 'rate': 0.0}, 'Glue Not Enough': {'count': 0, 'rate': 0.0}}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!F115:K115']},
    {'result_id': 'res_7', 'condition_id': 'cond_2', 'measurement_type': 'Vision VP/CD', 'condition_group': 'Test VP mold #9 (5.3 sheet)', 'date': '2024-03-05~06', 'line': 'C2 + E2', 'input_count': 3764, 'ok_count': 3764, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'NG VP/CD rate', 'metric_value': 0.0, 'unit': '%', 'judgement': None, 'ng_breakdown': {}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!F117:K117']},
    {'result_id': 'res_8', 'condition_id': 'cond_3', 'measurement_type': 'Vision VP/CD', 'condition_group': 'Normal VP mold #4/#6/#10 (5.3 sheet)', 'date': '2024-03-05~06', 'line': 'C2 + E2', 'input_count': 4968, 'ok_count': 4966, 'ng_count': 2, 'ng_rate_decimal': 0.0004, 'ng_rate_percent': 0.0, 'metric_name': 'NG VP/CD rate', 'metric_value': 0.0, 'unit': '%', 'judgement': None, 'ng_breakdown': {'Glue Not Enough': {'count': 2, 'rate': 0.0}}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!F119:K119']},
    # Function (5.3 totals - has baseline)
    {'result_id': 'res_9', 'condition_id': 'cond_1', 'measurement_type': 'Function', 'condition_group': 'Test VP mold #7 (5.3 sheet total)', 'date': '2024-03-05~07', 'line': 'C2-3B + E2', 'input_count': 3926, 'ok_count': 3783, 'ng_count': 143, 'ng_rate_decimal': 0.036, 'ng_rate_percent': 3.6, 'metric_name': 'Function NG rate', 'metric_value': 3.6, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG Sigma SPL': {'count': 0, 'rate': 0.0}, 'NG Sigma THD': {'count': 5, 'rate': 0.1}, 'NG Hearing Noise': {'count': 108, 'rate': 2.8}, 'NG Hearing Touch': {'count': 30, 'rate': 0.8}}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!F148:N149']},
    {'result_id': 'res_10', 'condition_id': 'cond_2', 'measurement_type': 'Function', 'condition_group': 'Test VP mold #9 (5.3 sheet total)', 'date': '2024-03-05~07', 'line': 'C2-3B + E2', 'input_count': 3735, 'ok_count': 3596, 'ng_count': 139, 'ng_rate_decimal': 0.037, 'ng_rate_percent': 3.7, 'metric_name': 'Function NG rate', 'metric_value': 3.7, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG Sigma SPL': {'count': 0, 'rate': 0.0}, 'NG Sigma THD': {'count': 0, 'rate': 0.0}, 'NG Hearing Noise': {'count': 115, 'rate': 3.1}, 'NG Hearing Touch': {'count': 24, 'rate': 0.6}}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!F150:N151']},
    {'result_id': 'res_11', 'condition_id': 'cond_3', 'measurement_type': 'Function', 'condition_group': 'Normal VP mold #4/#6/#10 (5.3 sheet total)', 'date': '2024-03-05~07', 'line': 'C2-3B + E2', 'input_count': 5870, 'ok_count': 5720, 'ng_count': 150, 'ng_rate_decimal': 0.026, 'ng_rate_percent': 2.6, 'metric_name': 'Function NG rate', 'metric_value': 2.6, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG Sigma SPL': {'count': 1, 'rate': 0.0}, 'NG Sigma THD': {'count': 0, 'rate': 0.0}, 'NG Hearing Noise': {'count': 120, 'rate': 2.0}, 'NG Hearing Touch': {'count': 29, 'rate': 0.5}}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!F152:N153']},
    # Function (7.3(2) totals - no baseline)
    {'result_id': 'res_12', 'condition_id': 'cond_1', 'measurement_type': 'Function', 'condition_group': 'Test VP mold #7 (7.3(2) total)', 'date': '2024-03-13~14', 'line': 'C2-3A', 'input_count': 10548, 'ok_count': 10405, 'ng_count': 143, 'ng_rate_decimal': 0.014, 'ng_rate_percent': 1.4, 'metric_name': 'Function NG rate', 'metric_value': 1.4, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG Hearing Noise': {'count': 123, 'rate': 1.2}, 'NG Hearing Touch': {'count': 20, 'rate': 0.2}}, 'source_file': name03, 'sheet_name': '7.3 (2)', 'source_cells': ['7.3 (2)!E53:N54']},
    {'result_id': 'res_13', 'condition_id': 'cond_2', 'measurement_type': 'Function', 'condition_group': 'Test VP mold #9 (7.3(2) total)', 'date': '2024-03-13~15', 'line': 'C2-3A', 'input_count': 15175, 'ok_count': 14920, 'ng_count': 255, 'ng_rate_decimal': 0.017, 'ng_rate_percent': 1.7, 'metric_name': 'Function NG rate', 'metric_value': 1.7, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG Hearing Noise': {'count': 214, 'rate': 1.4}, 'NG Hearing Touch': {'count': 41, 'rate': 0.3}}, 'source_file': name03, 'sheet_name': '7.3 (2)', 'source_cells': ['7.3 (2)!E55:N56']},
  ],
  'conclusions': [
    {'conclusion_id': 'concl_1', 'topic': 'VP bending after laze cutting',
     'statement_from_report': 'Type VP #7 after laze cutting happen NG VP bending Rate 4.8%; Type VP #9 NG VP bending Rate 4.2%.',
     'normalized_interpretation': '5.3 sheet (same-event baseline available): VP #7 8.1% vs Normal 3.6% = (8.1/3.6-1)*100 = +125.0% worse than Normal; VP #9 12.8% vs Normal 3.6% = +255.6% worse. 7.3(2) sheet without baseline: VP #7 0.2%, VP #9 0.5%. Mold +0.05mm increases VP bending NG rate substantially in the 5.3 trial.',
     'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!A157:A158']},
    {'conclusion_id': 'concl_2', 'topic': 'Function NG rate vs Normal',
     'statement_from_report': 'Function type VP test add 0.05mm NG rate higher more than normal 1~2%.',
     'normalized_interpretation': '5.3 sheet (baseline 2.6%): VP #7 3.6% = (3.6/2.6-1)*100 = +38.5% worse; VP #9 3.7% = +42.3% worse. Hearing Noise dominates every condition (Test #7 2.8%, Test #9 3.1%, Normal 2.0%).',
     'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!A158']},
    {'conclusion_id': 'concl_3', 'topic': 'VP/CD vision result',
     'statement_from_report': '(Total VP/CD vision tables, 5.3 sheet)',
     'normalized_interpretation': 'VP/CD offset NG is essentially zero for all conditions (Test #7 0.0%, Test #9 0.0%, Normal 0.0%); the mold change does not affect VP/CD assembly vision NG.',
     'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!F115:K119']},
  ],
  'troubleshooting_index': {
    'defect_name': 'NG VP Bending after laze cutting',
    'when_user_asks': ['VP mold dimension change', 'VP bending after laze cutting', 'mold +0.05mm impact'],
    'suggested_checks': [
      {'hint_id': 'hint_1', 'check_item': 'Reconsider +0.05mm center addition on VP mold #7/#9', 'reason': 'In 5.3 trial Test #7 VP bending 8.1% vs Normal 3.6% (+125.0% worse); Test #9 12.8% vs Normal 3.6% (+255.6% worse). Adding 0.05mm at the center clearly worsens laze cutting bending NG.', 'evidence_strength': 'high', 'related_process': 'Laze cutting', 'related_part': 'VP', 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!F85:I87']},
      {'hint_id': 'hint_2', 'check_item': 'Check Hearing Noise NG driver at function station (Test #7/#9 vs Normal)', 'reason': 'Function NG Hearing Noise: Test #7 2.8%, Test #9 3.1%, Normal 2.0%. Noise component of NG increased about 40-55% over baseline.', 'evidence_strength': 'medium', 'related_process': 'Function test', 'related_part': 'Speaker module', 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!F148:N153']},
      {'hint_id': 'hint_3', 'check_item': 'Investigate why 7.3(2) sheet VP bending is much lower (0.2-0.5%) than 5.3 sheet (8.1-12.8%)', 'reason': 'Same mold #7/#9 +0.05mm at two events shows very different NG levels; lot, line, or laze setting may have changed.', 'evidence_strength': 'medium', 'related_process': 'Laze cutting', 'related_part': 'VP', 'source_file': name03, 'sheet_name': '7.3 (2)', 'source_cells': ['7.3 (2)!E21:I22']},
    ],
    'limitations': ['7.3(2) sheet does not provide a same-event baseline row.', 'Decision text body is mostly empty; only 5.3 sheet summary states VP #7 4.8% and VP #9 4.2% as aggregated commentary, which differs from the totals computed above.']
  },
  'ai_extraction_log': {
    'confidence': 0.7,
    'assumptions': ['Two sheets are two separate trial events; only 5.3 sheet contains Normal baseline rows.', 'Totals taken from yellow-highlighted Total rows.'],
    'warnings': ['Decision summary number (VP #7 4.8%, VP #9 4.2%) differs from raw Total rows (8.1%, 12.8%); kept raw numbers.', '7.3(2) sheet has no baseline -> store Test rows without improvement claim.'],
    'decision_rationale': '5.3 sheet has same-event Normal baseline -> classified normal_comparison. Mold +0.05mm worsens VP bending (+125% / +255.6%) and worsens Function NG by ~+38-42%. VP/CD vision is unaffected. 7.3(2) sheet has no baseline; lower absolute NG suggests different lot or laze setting.'
  }
}

tr_ko_03 = {
  'document': {'title': 'BRS-161014DT VP mold #7,9 중심부 0.05mm 추가 시험 리포트', 'purpose': 'VP mold #7, #9 중심부에 0.05mm 추가 후 Normal mold(#4/#6/#10) 대비 개선 여부 확인.',
               'content': ['Laze cutting 후 VP bending 확인', 'Sub1에서 VP+CD vision NG rate 확인', 'Final에서 Function NG rate 확인']},
  'conclusions': {
    'concl_1': {'topic': 'Laze cutting VP bending', 'statement_from_report': 'VP #7 4.8%, VP #9 4.2% NG VP bending.', 'normalized_interpretation': '5.3 sheet (baseline 3.6%): VP #7 8.1% = +125.0% 악화, VP #9 12.8% = +255.6% 악화. 7.3(2) sheet (baseline 없음): VP #7 0.2%, VP #9 0.5%. 5.3 시험에서 0.05mm 추가가 VP bending 악화 명확.'},
    'concl_2': {'topic': 'Function NG rate vs Normal', 'statement_from_report': 'Function NG rate가 normal보다 1~2% 더 높음.', 'normalized_interpretation': '5.3 sheet baseline 2.6% 기준: VP #7 3.6% = +38.5% 악화, VP #9 3.7% = +42.3% 악화. Hearing Noise가 dominant.'},
    'concl_3': {'topic': 'VP/CD vision 결과', 'statement_from_report': '(VP/CD vision 표 합계 5.3)', 'normalized_interpretation': '모든 조건에서 VP/CD vision NG rate ~0.0%; mold 변경이 VP/CD 조립 vision에는 영향 없음.'},
  },
  'hints': {
    'hint_1': {'check_item': 'VP mold #7/#9 중심부 +0.05mm 적용 재검토', 'reason': '5.3 시험에서 VP #7 8.1% vs Normal 3.6% (+125% 악화), VP #9 12.8% vs 3.6% (+255.6% 악화). 0.05mm 추가가 laze cutting bending NG를 명백히 증가시킴.'},
    'hint_2': {'check_item': 'Hearing Noise NG 원인 점검 (Function)', 'reason': 'Function Hearing Noise: Test #7 2.8%, Test #9 3.1%, Normal 2.0%. Noise 항목이 baseline 대비 40~55% 증가.'},
    'hint_3': {'check_item': '7.3(2) 시험에서 VP bending이 5.3 대비 매우 낮은 원인 분석', 'reason': '같은 mold +0.05mm인데 두 이벤트의 NG가 크게 다름 → lot, 라인, laze setting 변동 확인.'},
  },
  'log': {'assumptions': ['두 시트는 별개 시험 이벤트로 가정; baseline은 5.3 sheet에만 존재.', 'Total 행(노란색)을 합계로 사용.'],
          'warnings': ['Decision의 4.8%/4.2%와 raw total 8.1%/12.8%가 다름; raw 사용.', '7.3(2) sheet는 baseline 없음.'],
          'decision_rationale': '5.3 sheet에 same-event Normal baseline 존재 → normal_comparison. Mold +0.05mm로 VP bending +125%/+255.6% 악화, Function +38~42% 악화. VP/CD vision은 영향 없음. 7.3(2)는 baseline 없으며 절대 NG가 낮은 이유는 lot/laze 변동 가능.'}
}
tr_en_03 = {
  'document': {'title': 'BRS-161014DT Report Test VP Mold #7,9 Add 0.05mm', 'purpose': 'Test VP mold #7 and #9 with 0.05mm added at center; verify improvement vs Normal mold #4/#6/#10.',
               'content': ['Check VP bending after laze cutting', 'Check vision VP+CD NG rate at Sub1', 'Check function NG rate at final']},
  'conclusions': {
    'concl_1': {'topic': 'Laze cutting VP bending', 'statement_from_report': 'VP #7 4.8%, VP #9 4.2% NG VP bending.', 'normalized_interpretation': '5.3 sheet (baseline 3.6%): VP #7 8.1% = +125.0% worse, VP #9 12.8% = +255.6% worse. 7.3(2) sheet (no baseline): VP #7 0.2%, VP #9 0.5%. 0.05mm addition clearly worsens VP bending in the 5.3 trial.'},
    'concl_2': {'topic': 'Function NG rate vs Normal', 'statement_from_report': 'Function NG rate higher than normal by 1~2%.', 'normalized_interpretation': '5.3 sheet baseline 2.6%: VP #7 3.6% = +38.5% worse, VP #9 3.7% = +42.3% worse. Hearing Noise dominates.'},
    'concl_3': {'topic': 'VP/CD vision result', 'statement_from_report': '(VP/CD vision totals on 5.3 sheet)', 'normalized_interpretation': 'VP/CD vision NG rate is ~0.0% in every condition; mold change does not affect VP/CD assembly vision.'},
  },
  'hints': {
    'hint_1': {'check_item': 'Reconsider +0.05mm center addition on VP mold #7/#9', 'reason': '5.3 trial: VP #7 8.1% vs Normal 3.6% = +125% worse; VP #9 12.8% vs 3.6% = +255.6% worse. 0.05mm addition clearly increases laze cutting bending NG.'},
    'hint_2': {'check_item': 'Investigate Hearing Noise NG driver at function station', 'reason': 'Function NG Hearing Noise: Test #7 2.8%, Test #9 3.1%, Normal 2.0%. Noise component increased 40-55% vs baseline.'},
    'hint_3': {'check_item': 'Investigate why 7.3(2) sheet VP bending (0.2-0.5%) is much lower than 5.3 sheet (8.1-12.8%)', 'reason': 'Same mold setting at two events shows very different NG; lot, line, or laze setting probably differ.'},
  },
  'log': {'assumptions': ['Two sheets are two separate trial events; baseline only present in 5.3 sheet.', 'Totals taken from yellow Total rows.'],
          'warnings': ['Decision narrative 4.8%/4.2% differs from raw 5.3 totals 8.1%/12.8%; kept raw data.', '7.3(2) sheet has no baseline row.'],
          'decision_rationale': '5.3 sheet has same-event Normal baseline -> normal_comparison. Mold +0.05mm worsens VP bending +125%/+255.6% and Function +38~42%. VP/CD vision unaffected. 7.3(2) absence of baseline plus lower absolute NG suggests lot or laze setting differed.'}
}
tr_vi_03 = {
  'document': {'title': 'BRS-161014DT Báo cáo test VP mold #7,9 thêm 0.05mm', 'purpose': 'Test VP mold #7 và #9 thêm 0.05mm ở giữa; xác minh cải thiện so với mold Normal #4/#6/#10.',
               'content': ['Kiểm tra VP bending sau laze cutting', 'Kiểm tra vision VP+CD NG rate tại Sub1', 'Kiểm tra NG rate Function tại final']},
  'conclusions': {
    'concl_1': {'topic': 'VP bending sau laze cutting', 'statement_from_report': 'VP #7 4.8%, VP #9 4.2% NG VP bending.', 'normalized_interpretation': 'Sheet 5.3 (baseline 3.6%): VP #7 8.1% = +125.0% xấu hơn, VP #9 12.8% = +255.6% xấu hơn. Sheet 7.3(2) (không baseline): VP #7 0.2%, VP #9 0.5%. Thêm 0.05mm rõ ràng làm xấu VP bending trong test 5.3.'},
    'concl_2': {'topic': 'Function NG rate vs Normal', 'statement_from_report': 'Function NG rate cao hơn normal 1~2%.', 'normalized_interpretation': 'Sheet 5.3 baseline 2.6%: VP #7 3.6% = +38.5% xấu hơn, VP #9 3.7% = +42.3% xấu hơn. Hearing Noise chiếm ưu thế.'},
    'concl_3': {'topic': 'VP/CD vision', 'statement_from_report': '(Tổng VP/CD vision sheet 5.3)', 'normalized_interpretation': 'VP/CD vision NG ~0.0% ở mọi điều kiện; mold không ảnh hưởng đến VP/CD assembly vision.'},
  },
  'hints': {
    'hint_1': {'check_item': 'Xem xét lại việc thêm 0.05mm ở giữa cho VP mold #7/#9', 'reason': 'Sheet 5.3: VP #7 8.1% vs Normal 3.6% = +125% xấu hơn; VP #9 12.8% vs 3.6% = +255.6% xấu hơn. Thêm 0.05mm rõ ràng tăng NG VP bending.'},
    'hint_2': {'check_item': 'Điều tra nguyên nhân Hearing Noise NG ở function', 'reason': 'Function Hearing Noise: Test #7 2.8%, Test #9 3.1%, Normal 2.0%. Noise tăng 40-55% so baseline.'},
    'hint_3': {'check_item': 'Tìm hiểu vì sao VP bending sheet 7.3(2) (0.2-0.5%) thấp hơn nhiều so sheet 5.3 (8.1-12.8%)', 'reason': 'Cùng mold +0.05mm nhưng hai event NG rất khác → lot, line hoặc laze setting khác.'},
  },
  'log': {'assumptions': ['Hai sheet là hai event test riêng; baseline chỉ ở sheet 5.3.', 'Tổng lấy từ hàng Total màu vàng.'],
          'warnings': ['Decision ghi 4.8%/4.2% khác raw total 5.3 8.1%/12.8%; giữ raw data.', '7.3(2) không có baseline.'],
          'decision_rationale': 'Sheet 5.3 có baseline cùng sự kiện -> normal_comparison. Mold +0.05mm làm xấu VP bending +125%/+255.6% và Function +38~42%. VP/CD vision không bị ảnh hưởng. Sheet 7.3(2) không baseline, NG tuyệt đối thấp có thể do lot/laze khác.'}
}
results[name03] = (result03, tr_ko_03, tr_en_03, tr_vi_03)

# Commit first batch (1-3)
processed = 0
failed = 0
for n, (res, ko, en, vi) in results.items():
    if h.commit_dataset(n, res, ko, en, vi):
        processed += 1
        print(f'OK: {n}')
    else:
        failed += 1
        print(f'FAIL: {n}')

print(f'partial batch 1: processed={processed} failed={failed}')
