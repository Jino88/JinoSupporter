"""Process chunk 02 datasets 4-9."""
import _ai_batch_helper as h

results = {}

# =====================================================================
# DS04: 27. BRS-161014 Report check problem NG Gap frame+Yoke 2023.09.14
# =====================================================================
name04 = '27. BRS-161014 Report check problem NG Gap frame+Yoke 2023.09.14'

result04 = {
  'schema_version': '0.1',
  'document': {
    'document_id': '', 'source_file': name04, 'source_sheet': 'Report',
    'title': 'BRS-161014 Report Check Reason NG Gap Frame/Yoke',
    'model': 'BRS-161014', 'report_date': '2023-09-14',
    'department': 'ME', 'marker': 'Tu', 'line': '',
    'report_type': 'before_after_dimension',
    'primary_defect': {'canonical_name': 'NG Gap Frame+Yoke (Height NG)', 'aliases_in_document': ['NG height check', 'NG Gap Frame-BPT']},
    'related_defects': ['Dimension NG (Frame)', 'Dimension NG (BPT/Yoke)'],
    'parts': ['Frame', 'BPT', 'Yoke'], 'processes': ['Dimension check (3D)', 'Dimension check (caliper)', 'Assembly'],
    'purpose': 'Improve NG of height/gap check at Frame+Yoke assembly.',
    'content': ['Check Frame dimension by 3D', 'Check BPT dimension by caliper'],
    'source_cells': {'title': ['Report!B1'], 'date': ['Report!L2'], 'purpose': ['Report!A4'], 'content': ['Report!A6:A7']}
  },
  'test_conditions': [
    {'condition_id': 'cond_1', 'condition_group': 'Dimension survey 5 samples', 'line': '', 'process': 'Dimension check', 'changed_factor': 'Measure 5 frames + 5 BPTs', 'before_value': None, 'after_value': None, 'unit': 'mm', 'machine': '3D / caliper', 'jig': None, 'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name04, 'sheet_name': 'Report', 'source_cells': ['Report!A9:A10']},
  ],
  'results': [
    {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Dimension Frame #1', 'condition_group': 'Frame dim 1 (spec 5.80~5.85)', 'date': '2023-09-14', 'line': '', 'input_count': 5, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None, 'metric_name': 'Frame dim1 AVG', 'metric_value': 5.8164, 'unit': 'mm', 'judgement': 'CHECK', 'ng_breakdown': {'MIN': 5.7932, 'MAX': 5.8396}, 'source_file': name04, 'sheet_name': 'Report', 'source_cells': ['Report!C21:E23']},
    {'result_id': 'res_2', 'condition_id': 'cond_1', 'measurement_type': 'Dimension Frame #2', 'condition_group': 'Frame dim 2 (spec 14.21~14.26)', 'date': '2023-09-14', 'line': '', 'input_count': 5, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None, 'metric_name': 'Frame dim2 AVG', 'metric_value': 14.3106, 'unit': 'mm', 'judgement': 'FAIL', 'ng_breakdown': {'MIN': 14.2895, 'MAX': 14.3317}, 'source_file': name04, 'sheet_name': 'Report', 'source_cells': ['Report!C21:E23']},
    {'result_id': 'res_3', 'condition_id': 'cond_1', 'measurement_type': 'Dimension BPT #1', 'condition_group': 'BPT dim 1 (spec 5.82~5.85)', 'date': '2023-09-14', 'line': '', 'input_count': 5, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None, 'metric_name': 'BPT dim1 AVG', 'metric_value': 5.835, 'unit': 'mm', 'judgement': 'PASS', 'ng_breakdown': {'MIN': 5.83, 'MAX': 5.84}, 'source_file': name04, 'sheet_name': 'Report', 'source_cells': ['Report!F21:G23']},
    {'result_id': 'res_4', 'condition_id': 'cond_1', 'measurement_type': 'Dimension BPT #2', 'condition_group': 'BPT dim 2 (spec 14.23~14.26)', 'date': '2023-09-14', 'line': '', 'input_count': 5, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None, 'metric_name': 'BPT dim2 AVG', 'metric_value': 14.245, 'unit': 'mm', 'judgement': 'PASS', 'ng_breakdown': {'MIN': 14.24, 'MAX': 14.25}, 'source_file': name04, 'sheet_name': 'Report', 'source_cells': ['Report!F21:G23']},
    {'result_id': 'res_5', 'condition_id': 'cond_1', 'measurement_type': 'Gap Frame-BPT (1)', 'condition_group': 'Gap (1) sample 1-5', 'date': '2023-09-14', 'line': '', 'input_count': 5, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None, 'metric_name': 'Gap1 range', 'metric_value': None, 'unit': 'mm', 'judgement': 'CHECK', 'ng_breakdown': {'sample_values': '-0.02, -0.04, 0.01, -0.03, -0.04'}, 'source_file': name04, 'sheet_name': 'Report', 'source_cells': ['Report!H16:H20']},
    {'result_id': 'res_6', 'condition_id': 'cond_1', 'measurement_type': 'Gap Frame-BPT (2)', 'condition_group': 'Gap (2) sample 1-5', 'date': '2023-09-14', 'line': '', 'input_count': 5, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None, 'metric_name': 'Gap2 range', 'metric_value': None, 'unit': 'mm', 'judgement': 'CHECK', 'ng_breakdown': {'sample_values': '0.08, 0.06, 0.06, 0.04, 0.09'}, 'source_file': name04, 'sheet_name': 'Report', 'source_cells': ['Report!I16:I20']},
  ],
  'conclusions': [
    {'conclusion_id': 'concl_1', 'topic': 'Frame and Yoke drawing limits cause gap NG',
     'statement_from_report': 'Dimension Frame NG + Drawing trouble: Frame (Min=5.80), Yoke (Min=5.82) => so can not assy OK.',
     'normalized_interpretation': 'Frame dimension #2 AVG 14.3106 mm exceeds spec MAX 14.26 (all 5 samples > 14.26) -> Frame dimension is NG. BPT dimensions #1 (AVG 5.835) and #2 (AVG 14.245) sit at the spec MIN edge. Combined drawing limits (Frame min 5.80, Yoke min 5.82) leave essentially no clearance, so the gap-check assembly cannot be OK.',
     'source_file': name04, 'sheet_name': 'Report', 'source_cells': ['Report!A25']},
  ],
  'troubleshooting_index': {
    'defect_name': 'NG Gap Frame+Yoke',
    'when_user_asks': ['Gap Frame Yoke NG', 'height check NG', 'dimension drawing limit issue'],
    'suggested_checks': [
      {'hint_id': 'hint_1', 'check_item': 'Review Frame drawing min spec vs actual machining capability', 'reason': 'Frame dim2 AVG 14.3106 mm vs spec MAX 14.26 (5/5 samples out of spec); current machining is biased above the upper limit.', 'evidence_strength': 'high', 'related_process': 'Frame machining', 'related_part': 'Frame', 'source_file': name04, 'sheet_name': 'Report', 'source_cells': ['Report!C21:C23']},
      {'hint_id': 'hint_2', 'check_item': 'Re-examine BPT/Yoke drawing minimums and combined gap', 'reason': 'BPT dim1 5.835 is at spec MIN 5.82+0.015 only; combined with Frame MIN 5.80, the gap budget collapses.', 'evidence_strength': 'medium', 'related_process': 'BPT/Yoke machining + Frame+Yoke assembly', 'related_part': 'BPT, Yoke', 'source_file': name04, 'sheet_name': 'Report', 'source_cells': ['Report!F21:G23']},
      {'hint_id': 'hint_3', 'check_item': 'Capture Gap Frame-BPT measurement spec and pass criteria', 'reason': 'Gap (1) and Gap (2) values are tabulated but no spec is given in the sheet; without spec, PASS/FAIL is not decidable.', 'evidence_strength': 'low', 'related_process': 'Assembly gap check', 'related_part': 'Frame+BPT+Yoke', 'source_file': name04, 'sheet_name': 'Report', 'source_cells': ['Report!H16:I20']},
    ],
    'limitations': ['Only 5 samples per dimension are measured.', 'No Pass/Fail spec is given for Gap (1)/(2) columns.']
  },
  'ai_extraction_log': {
    'confidence': 0.6,
    'assumptions': ['Spec rows pinned: Frame 5.80~5.85 / 14.21~14.26, BPT 5.82~5.85 / 14.23~14.26.', '"Frame Min=5.80, Yoke Min=5.82" in the decision text refers to drawing limits, not measured values.'],
    'warnings': ['Sample size n=5 per dimension is small.', 'Gap (1)/(2) measurement has no spec column in the sheet.'],
    'decision_rationale': 'Frame dim2 AVG 14.31 (5/5 above spec MAX 14.26) confirms Frame side is the dimension NG driver. BPT measurements sit at spec MIN edge so combined clearance fails. This is a dimensional spec_gate analysis, not an NG rate comparison; report_type = before_after_dimension.'
  }
}

tr_ko_04 = {
  'document': {'title': 'BRS-161014 Frame/Yoke Gap NG 원인 확인 리포트', 'purpose': 'Frame+Yoke 조립 시 height/gap check NG 개선.',
               'content': ['Frame 치수 3D 측정', 'BPT 치수 caliper 측정']},
  'conclusions': {'concl_1': {'topic': 'Frame/Yoke 도면 한계로 인한 gap NG', 'statement_from_report': 'Frame dimension NG + 도면 문제: Frame(Min=5.80), Yoke(Min=5.82) → 조립 OK 불가.', 'normalized_interpretation': 'Frame dim #2 평균 14.3106 mm로 spec MAX 14.26 초과 (5/5 모두 out-of-spec) → Frame 치수 NG. BPT는 spec MIN 경계. 도면 한계 자체가 여유가 없어 조립 시 gap이 맞지 않음.'}},
  'hints': {
    'hint_1': {'check_item': 'Frame 도면 min 스펙과 실제 가공 능력 재검토', 'reason': 'Frame dim2 AVG 14.3106 mm vs spec MAX 14.26 (5/5 OUT); 가공 편향 상한 초과.'},
    'hint_2': {'check_item': 'BPT/Yoke 도면 min 및 결합 gap 검토', 'reason': 'BPT dim1 5.835는 spec MIN 5.82에 +0.015 수준; Frame MIN 5.80과 결합되면 gap 여유 부족.'},
    'hint_3': {'check_item': 'Gap Frame-BPT 측정 spec 및 합격 기준 정립', 'reason': 'Gap (1)/(2) 측정값은 있으나 spec이 없어 PASS/FAIL 판정 불가.'},
  },
  'log': {'assumptions': ['Spec: Frame 5.80~5.85 / 14.21~14.26, BPT 5.82~5.85 / 14.23~14.26.', '"Frame Min=5.80, Yoke Min=5.82"는 도면 한계.'],
          'warnings': ['샘플 수 n=5로 적음.', 'Gap 컬럼에 spec 없음.'],
          'decision_rationale': 'Frame dim2 14.31 (5/5 over spec MAX) → Frame 치수 NG 원인. BPT는 MIN 경계. 차원 spec_gate 분석 → before_after_dimension.'}
}
tr_en_04 = {
  'document': {'title': 'BRS-161014 Report Check Reason NG Gap Frame/Yoke', 'purpose': 'Improve NG of height/gap check at Frame+Yoke assembly.',
               'content': ['Check Frame dimension by 3D', 'Check BPT dimension by caliper']},
  'conclusions': {'concl_1': {'topic': 'Frame and Yoke drawing limits cause gap NG', 'statement_from_report': 'Dimension Frame NG + drawing trouble: Frame(Min=5.80), Yoke(Min=5.82) -> can not assy OK.', 'normalized_interpretation': 'Frame dim2 AVG 14.3106 mm above spec MAX 14.26 (5/5 OUT) -> Frame dimension NG. BPT dimensions sit at spec MIN edge. Drawing limits leave essentially no clearance for assembly.'}},
  'hints': {
    'hint_1': {'check_item': 'Review Frame drawing min spec vs machining capability', 'reason': 'Frame dim2 AVG 14.3106 vs spec MAX 14.26 (5/5 OUT); machining biased above upper limit.'},
    'hint_2': {'check_item': 'Re-examine BPT/Yoke drawing minimum and combined gap', 'reason': 'BPT dim1 5.835 only 0.015 above spec MIN 5.82; combined with Frame MIN 5.80 the clearance budget collapses.'},
    'hint_3': {'check_item': 'Define Gap Frame-BPT measurement spec and PASS criteria', 'reason': 'Gap (1)/(2) values are recorded but no spec column is given.'},
  },
  'log': {'assumptions': ['Spec rows: Frame 5.80~5.85 / 14.21~14.26, BPT 5.82~5.85 / 14.23~14.26.', '"Frame Min=5.80, Yoke Min=5.82" refers to drawing limits.'],
          'warnings': ['n=5 per dimension is small.', 'Gap columns have no spec.'],
          'decision_rationale': 'Frame dim2 AVG 14.31 (5/5 above spec MAX) -> Frame is the dim NG driver. BPT sits at spec MIN edge. Dimension spec_gate analysis -> report_type before_after_dimension.'}
}
tr_vi_04 = {
  'document': {'title': 'BRS-161014 Báo cáo kiểm tra nguyên nhân NG Gap Frame/Yoke', 'purpose': 'Cải thiện NG height/gap check tại assembly Frame+Yoke.',
               'content': ['Đo kích thước Frame bằng 3D', 'Đo kích thước BPT bằng caliper']},
  'conclusions': {'concl_1': {'topic': 'Giới hạn bản vẽ Frame/Yoke gây NG gap', 'statement_from_report': 'Kích thước Frame NG + vấn đề bản vẽ: Frame(Min=5.80), Yoke(Min=5.82) → không lắp OK.', 'normalized_interpretation': 'Frame dim2 AVG 14.3106 mm vượt spec MAX 14.26 (5/5 OUT) → Frame NG. BPT nằm ở mép spec MIN. Bản vẽ không còn dung sai → assembly gap không đạt.'}},
  'hints': {
    'hint_1': {'check_item': 'Xem xét lại spec min của Frame so với năng lực gia công', 'reason': 'Frame dim2 AVG 14.3106 vs spec MAX 14.26 (5/5 OUT); gia công lệch trên.'},
    'hint_2': {'check_item': 'Xem xét lại min bản vẽ BPT/Yoke và gap kết hợp', 'reason': 'BPT dim1 5.835 chỉ trên spec MIN 5.82 +0.015; kết hợp Frame MIN 5.80 thì gap không còn dung sai.'},
    'hint_3': {'check_item': 'Quy định spec và tiêu chí PASS cho Gap Frame-BPT', 'reason': 'Có giá trị Gap (1)/(2) nhưng không có cột spec.'},
  },
  'log': {'assumptions': ['Spec: Frame 5.80~5.85 / 14.21~14.26, BPT 5.82~5.85 / 14.23~14.26.', '"Frame Min=5.80, Yoke Min=5.82" = giới hạn bản vẽ.'],
          'warnings': ['n=5 mỗi kích thước là ít.', 'Cột Gap không có spec.'],
          'decision_rationale': 'Frame dim2 AVG 14.31 (5/5 trên spec MAX) → Frame là nguyên nhân NG. BPT ở mép spec MIN. Phân tích spec_gate → before_after_dimension.'}
}
results[name04] = (result04, tr_ko_04, tr_en_04, tr_vi_04)

# =====================================================================
# DS05: 27. BRS-161014 Report test MTR VP NG Bending 2023.12.30
# =====================================================================
name05 = '27. BRS-161014 Report test MTR VP NG Bending 2023.12.30'

result05 = {
  'schema_version': '0.1',
  'document': {
    'document_id': '', 'source_file': name05, 'source_sheet': 'Report',
    'title': 'BRS-161014 Report Test MTR VP NG Bending Mold K4',
    'model': 'BRS-161014', 'report_date': '2023-12-30',
    'department': 'ME', 'marker': 'Nhung', 'line': '',
    'report_type': 'normal_comparison',
    'primary_defect': {'canonical_name': 'NG VP Bending', 'aliases_in_document': ['VP NG bending mold K4', 'Dome ass\'y bending']},
    'related_defects': ['Bonding offset / Dome separate', "Dome ass'y offset", "Dome ass'y bending (G2)", 'NG Hearing Noise', 'NG Sigma SPL', 'NG Sigma THD'],
    'parts': ['VP mold K4', 'Dome', 'Coil'], 'processes': ['Vision VP dome at Sub2', 'Function test'],
    'purpose': 'Material VP mold K4 has 100% NG bending; run material on main line and check process+function.',
    'content': ['Use VP NG bending material', 'Check process (Vision VP dome)', 'Check function'],
    'source_cells': {'title': ['Report!B1'], 'date': ['Report!N2'], 'purpose': ['Report!A4'], 'content': ['Report!A6:A8']}
  },
  'test_conditions': [
    {'condition_id': 'cond_1', 'condition_group': 'Test (VP mold K4 NG bending lot)', 'line': '', 'process': 'Sub2 vision + Function', 'changed_factor': 'VP material lot (mold K4 with NG bending)', 'before_value': 'Normal VP lot', 'after_value': 'VP mold K4 NG bending lot', 'unit': None, 'machine': None, 'jig': None, 'material_lot': 'mold K4 NG bending', 'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name05, 'sheet_name': 'Report', 'source_cells': ['Report!E16']},
    {'condition_id': 'cond_2', 'condition_group': 'Normal', 'line': '', 'process': 'Sub2 vision + Function', 'changed_factor': 'Baseline (Normal lot)', 'before_value': None, 'after_value': 'Normal', 'unit': None, 'machine': None, 'jig': None, 'material_lot': 'Normal', 'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name05, 'sheet_name': 'Report', 'source_cells': ['Report!E17']},
  ],
  'results': [
    {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Vision VP Dome', 'condition_group': 'Test (VP mold K4 NG bending)', 'date': '2023-12-30', 'line': '', 'input_count': 1010, 'ok_count': 1006, 'ng_count': 4, 'ng_rate_decimal': 0.004, 'ng_rate_percent': 0.4, 'metric_name': 'NG Process rate', 'metric_value': 0.4, 'unit': '%', 'judgement': None, 'ng_breakdown': {'Bonding offset / Dome separate': {'count': 3}, "Dome ass'y offset": {'count': 1}, "Dome ass'y bending (G2)": {'count': 59}}, 'source_file': name05, 'sheet_name': 'Report', 'source_cells': ['Report!F16:N16']},
    {'result_id': 'res_2', 'condition_id': 'cond_2', 'measurement_type': 'Vision VP Dome', 'condition_group': 'Normal', 'date': '2023-12-30', 'line': '', 'input_count': 1000, 'ok_count': 999, 'ng_count': 1, 'ng_rate_decimal': 0.001, 'ng_rate_percent': 0.1, 'metric_name': 'NG Process rate', 'metric_value': 0.1, 'unit': '%', 'judgement': None, 'ng_breakdown': {'Bonding offset / Dome separate': {'count': 1}, "Dome ass'y offset": {'count': 0}, "Dome ass'y bending (G2)": {'count': 0}}, 'source_file': name05, 'sheet_name': 'Report', 'source_cells': ['Report!F17:N17']},
    {'result_id': 'res_3', 'condition_id': 'cond_1', 'measurement_type': 'Function', 'condition_group': 'Test (VP mold K4 NG bending)', 'date': '2023-12-30', 'line': '', 'input_count': 1006, 'ok_count': 979, 'ng_count': 27, 'ng_rate_decimal': 0.027, 'ng_rate_percent': 2.7, 'metric_name': 'Function NG rate', 'metric_value': 2.7, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG Sigma SPL': {'count': 3}, 'NG Sigma THD': {'count': 1}, 'NG Hearing Noise': {'count': 23}, 'NG Hearing Touch': {'count': 0}}, 'source_file': name05, 'sheet_name': 'Report', 'source_cells': ['Report!F21:N21']},
    {'result_id': 'res_4', 'condition_id': 'cond_2', 'measurement_type': 'Function', 'condition_group': 'Normal', 'date': '2023-12-30', 'line': '', 'input_count': 2188, 'ok_count': 2111, 'ng_count': 77, 'ng_rate_decimal': 0.035, 'ng_rate_percent': 3.5, 'metric_name': 'Function NG rate', 'metric_value': 3.5, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG Sigma SPL': {'count': 1}, 'NG Sigma THD': {'count': 0}, 'NG Hearing Noise': {'count': 67}, 'NG Hearing Touch': {'count': 9}}, 'source_file': name05, 'sheet_name': 'Report', 'source_cells': ['Report!F23:N23']},
  ],
  'conclusions': [
    {'conclusion_id': 'concl_1', 'topic': 'Function NG vs Normal',
     'statement_from_report': 'NG of lot test same normal line.',
     'normalized_interpretation': 'Function NG rate Test 2.7% vs Normal 3.5% = (2.7/3.5-1)*100 = -22.9% improved vs same-event Normal. NG mix: Hearing Noise Test 2.3% vs Normal 3.1% (-25.8%), so noise is dominant in both. VP mold K4 NG bending lot does not increase function NG.',
     'source_file': name05, 'sheet_name': 'Report', 'source_cells': ['Report!A25']},
    {'conclusion_id': 'concl_2', 'topic': 'Process NG (Vision VP Dome)',
     'statement_from_report': '(Process NG rate Test 0.4% vs Normal 0.1%)',
     'normalized_interpretation': 'Sub2 vision NG rate Test 0.4% vs Normal 0.1% = +300.0% worse than Normal. Test condition shows 3 Bonding offset / Dome separate (vs 1) and 1 Dome ass\'y offset (vs 0). Process NG worsens with the bending lot, even though function NG is similar.',
     'source_file': name05, 'sheet_name': 'Report', 'source_cells': ['Report!F16:N17']},
  ],
  'troubleshooting_index': {
    'defect_name': 'VP mold K4 NG Bending material usage',
    'when_user_asks': ['VP bending material use or not', 'mold K4 VP bending function impact'],
    'suggested_checks': [
      {'hint_id': 'hint_1', 'check_item': 'Inspect Sub2 vision Dome bonding offset rate for VP bending lots', 'reason': "Vision VP dome NG Test 0.4% vs Normal 0.1% (+300%). Bonding offset / Dome separate and Dome ass'y offset are present only in test lot.", 'evidence_strength': 'medium', 'related_process': 'Sub2 vision', 'related_part': 'VP / Dome', 'source_file': name05, 'sheet_name': 'Report', 'source_cells': ['Report!F16:N17']},
      {'hint_id': 'hint_2', 'check_item': 'Verify function NG mix on VP bending lots over larger sample', 'reason': 'Function Test 2.7% vs Normal 3.5% (-22.9% improved); however Test Hearing Noise 2.3% (n=1006) is still the dominant NG mode like Normal.', 'evidence_strength': 'low', 'related_process': 'Function test', 'related_part': 'VP/Coil/Dome', 'source_file': name05, 'sheet_name': 'Report', 'source_cells': ['Report!F21:N23']},
      {'hint_id': 'hint_3', 'check_item': 'Decide whether VP bending material can be released based on Sub2 vs Function trade-off', 'reason': 'Sub2 vision worsens (+300%) but Function does not worsen (-22.9% improved). Need PE/QA decision on which is more critical.', 'evidence_strength': 'medium', 'related_process': 'Material disposition', 'related_part': 'VP mold K4', 'source_file': name05, 'sheet_name': 'Report', 'source_cells': ['Report!A25']},
    ],
    'limitations': ['Decision section is empty.', 'Function breakdown counts vs percent values in the sheet have obvious errors (percent column lists >1000% values); only count columns are trusted.']
  },
  'ai_extraction_log': {
    'confidence': 0.65,
    'assumptions': ['Test row vs Nomal row are same-event baseline pair.', 'Function percent cells with >1000% values are spreadsheet errors; treat as garbage and rely on count vs Input.'],
    'warnings': ['Function NG sub-percent cells contain obviously incorrect values (e.g., 11177.8%); they were ignored.', 'Sub2 vision NG breakdown shows Dome bending count 59 but Total NG only 4; sheet has internal inconsistency, kept Total NG 4 as authoritative.'],
    'decision_rationale': 'Same-event Normal row exists. Function NG Test 2.7% vs Normal 3.5% (-22.9% improved). Vision NG Test 0.4% vs Normal 0.1% (+300%). VP bending lot does not damage function but does increase vision NG.'}
}

tr_ko_05 = {
  'document': {'title': 'BRS-161014 VP mold K4 NG Bending 자재 시험 리포트', 'purpose': 'VP mold K4 NG bending 100% 자재를 main line에 투입하여 process / function 영향 확인.',
               'content': ['VP NG bending 자재 사용', 'Sub2 vision VP dome 확인', 'Function 확인']},
  'conclusions': {
    'concl_1': {'topic': 'Function NG vs Normal', 'statement_from_report': 'Test lot의 NG가 normal과 동일.', 'normalized_interpretation': 'Function NG Test 2.7% vs Normal 3.5% = -22.9% 개선. NG mix: Hearing Noise Test 2.3% vs Normal 3.1% (-25.8%). VP mold K4 bending 자재가 function을 악화시키지 않음.'},
    'concl_2': {'topic': 'Sub2 vision NG', 'statement_from_report': '(Vision NG Test 0.4% vs Normal 0.1%)', 'normalized_interpretation': 'Sub2 vision NG Test 0.4% vs Normal 0.1% = +300% 악화. Bonding offset/Dome separate 3건, Dome offset 1건이 Test에서만 발생.'},
  },
  'hints': {
    'hint_1': {'check_item': 'VP bending lot의 Sub2 dome bonding offset 비율 점검', 'reason': 'Vision NG Test 0.4% vs Normal 0.1% (+300%); Test에서만 bonding offset/dome separate 및 dome offset 발생.'},
    'hint_2': {'check_item': '더 큰 sample로 VP bending lot의 function NG mix 재확인', 'reason': 'Function Test 2.7% vs Normal 3.5% (-22.9% 개선)이나 Hearing Noise는 여전히 dominant.'},
    'hint_3': {'check_item': 'VP bending 자재 사용 가능 여부 결정 (Sub2 악화 vs Function 동등 trade-off)', 'reason': 'Sub2 +300% 악화, Function -22.9% 개선이라 PE/QA 판단 필요.'},
  },
  'log': {'assumptions': ['Test와 Nomal 행은 same-event baseline pair.', 'Function 퍼센트 셀이 >1000%인 것은 시트 오류로 무시.'],
          'warnings': ['Function 하위 % 셀의 값 오류 다수.', 'Sub2 dome bending count 59 vs Total NG 4 불일치, Total NG 4 사용.'],
          'decision_rationale': 'Same-event baseline 존재. Function -22.9% 개선, Vision +300% 악화. VP bending 자재가 function은 해치지 않지만 process는 악화.'}
}
tr_en_05 = {
  'document': {'title': 'BRS-161014 Report Test MTR VP NG Bending Mold K4', 'purpose': 'Use VP mold K4 NG bending material on main line and check process and function impact.',
               'content': ['Use VP NG bending material', 'Check Sub2 vision VP dome', 'Check function']},
  'conclusions': {
    'concl_1': {'topic': 'Function NG vs Normal', 'statement_from_report': 'NG of lot test same normal line.', 'normalized_interpretation': 'Function NG Test 2.7% vs Normal 3.5% = -22.9% improved. NG mix: Hearing Noise Test 2.3% vs Normal 3.1% (-25.8%). VP bending lot does not worsen function.'},
    'concl_2': {'topic': 'Sub2 vision NG', 'statement_from_report': '(Vision NG Test 0.4% vs Normal 0.1%)', 'normalized_interpretation': 'Sub2 vision NG Test 0.4% vs Normal 0.1% = +300% worse. Bonding offset/Dome separate 3, Dome offset 1 only in Test.'},
  },
  'hints': {
    'hint_1': {'check_item': 'Inspect Sub2 dome bonding offset rate for VP bending lots', 'reason': 'Vision NG Test 0.4% vs Normal 0.1% (+300%); bonding offset/dome separate and dome offset only in Test.'},
    'hint_2': {'check_item': 'Verify function NG mix on larger sample', 'reason': 'Function Test 2.7% vs Normal 3.5% (-22.9%) but Hearing Noise still dominant.'},
    'hint_3': {'check_item': 'Decide whether VP bending material can be released (Sub2 worse vs Function equal trade-off)', 'reason': 'Sub2 +300% worse but Function -22.9% improved; PE/QA judgement needed.'},
  },
  'log': {'assumptions': ['Test row vs Nomal row are same-event baseline pair.', 'Function percent cells >1000% are spreadsheet errors; ignored.'],
          'warnings': ['Function sub-percent cells contain invalid values.', "Sub2 dome bending count 59 inconsistent with Total NG 4; kept Total NG 4."],
          'decision_rationale': 'Same-event baseline exists. Function -22.9% improved, Vision +300% worse. VP bending lot does not hurt function but worsens process.'}
}
tr_vi_05 = {
  'document': {'title': 'BRS-161014 Báo cáo test vật liệu VP NG Bending mold K4', 'purpose': 'Đưa lot VP mold K4 NG bending vào main line và kiểm tra ảnh hưởng đến process và function.',
               'content': ['Sử dụng vật liệu VP NG bending', 'Kiểm tra Sub2 vision VP dome', 'Kiểm tra function']},
  'conclusions': {
    'concl_1': {'topic': 'Function NG vs Normal', 'statement_from_report': 'NG của lot test giống normal.', 'normalized_interpretation': 'Function NG Test 2.7% vs Normal 3.5% = -22.9% cải thiện. Hearing Noise Test 2.3% vs Normal 3.1% (-25.8%). Lot VP bending không làm xấu function.'},
    'concl_2': {'topic': 'Sub2 vision NG', 'statement_from_report': '(Vision NG Test 0.4% vs Normal 0.1%)', 'normalized_interpretation': 'Sub2 vision NG Test 0.4% vs Normal 0.1% = +300% xấu hơn. Bonding offset/Dome separate 3, Dome offset 1 chỉ ở Test.'},
  },
  'hints': {
    'hint_1': {'check_item': 'Kiểm tra tỉ lệ dome bonding offset Sub2 cho lot VP bending', 'reason': 'Vision NG Test 0.4% vs Normal 0.1% (+300%); bonding offset/dome separate và dome offset chỉ ở Test.'},
    'hint_2': {'check_item': 'Xác nhận lại function NG mix trên sample lớn', 'reason': 'Function Test 2.7% vs Normal 3.5% (-22.9%) nhưng Hearing Noise vẫn dominant.'},
    'hint_3': {'check_item': 'Quyết định có dùng VP bending hay không (trade-off Sub2 xấu vs Function bằng)', 'reason': 'Sub2 +300% xấu, Function -22.9% cải thiện → cần PE/QA quyết định.'},
  },
  'log': {'assumptions': ['Test và Nomal là cặp baseline cùng sự kiện.', 'Ô % function >1000% là lỗi sheet, bỏ qua.'],
          'warnings': ['Ô % phụ function có giá trị lỗi.', 'Sub2 dome bending count 59 không khớp Total NG 4; giữ Total NG 4.'],
          'decision_rationale': 'Có baseline cùng sự kiện. Function -22.9% cải thiện, Vision +300% xấu. Lot VP bending không hại function nhưng làm xấu process.'}
}
results[name05] = (result05, tr_ko_05, tr_en_05, tr_vi_05)

# =====================================================================
# DS06: 27. BRS-161016 Report check problem separate Coil+SP 2.4.2024
# =====================================================================
name06 = '27. BRS-161016 Report check problem separate Coil+SP 2.4.2024'

result06 = {
  'schema_version': '0.1',
  'document': {
    'document_id': '', 'source_file': name06, 'source_sheet': '5.3',
    'title': 'BRS-161016 Report Check Problem Separate Coil+Suspension',
    'model': 'BRS-161016', 'report_date': '2024-04-02',
    'department': 'ME', 'marker': 'Thuy', 'line': '',
    'report_type': 'doe_matrix',
    'primary_defect': {'canonical_name': 'Coil+SP Separation', 'aliases_in_document': ['Coil+SP separate', 'Separation NG']},
    'related_defects': ['Coil+SP Press separate (Press 1-4)'],
    'parts': ['Coil', 'Suspension (SP)'], 'processes': ["Ass'y Coil+SP", 'UC press', 'LED UV', 'Bonding line'],
    'purpose': 'Check the reason for Coil+SP separation NG.',
    'content': ['PE adjust UC press down and check', 'PE change new holder and check', 'Test increase time UC press and LED UV', "Test with and without height check at ass'y Coil+SP process", 'Test change bonding line to move inside'],
    'source_cells': {'title': ['5.3!B1'], 'date': ['5.3!N2'], 'purpose': ['5.3!A4'], 'content': ['5.3!A6:A10']}
  },
  'test_conditions': [
    {'condition_id': 'cond_1', 'condition_group': '7S-7S + height check + Old line + Normal press', 'line': '', 'process': "Ass'y Coil+SP", 'changed_factor': 'Time LED UV/UC 7S-7S, Old bonding line', 'before_value': None, 'after_value': '7S-7S / use height check / Old bonding line / Normal press', 'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None, 'dry_time_sec': 7.0, 'temperature': None, 'pressure': 'Normal', 'bond_amount': None, 'uv_energy': None, 'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!B18:E18']},
    {'condition_id': 'cond_2', 'condition_group': '7.5S-7.5S + height check + Old line + Normal press', 'line': '', 'process': "Ass'y Coil+SP", 'changed_factor': 'Time LED UV/UC 7.5S-7.5S', 'before_value': '7S-7S', 'after_value': '7.5S-7.5S / use height check / Old bonding line / Normal press', 'unit': 'sec', 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None, 'dry_time_sec': 7.5, 'temperature': None, 'pressure': 'Normal', 'bond_amount': None, 'uv_energy': None, 'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!B19:E19']},
    {'condition_id': 'cond_3', 'condition_group': "7.5S-7.5S + don't use height check + Old line + Normal press", 'line': '', 'process': "Ass'y Coil+SP", 'changed_factor': "Don't use height check", 'before_value': 'Use height check', 'after_value': "7.5S-7.5S / don't use height check / Old line / Normal press", 'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None, 'dry_time_sec': 7.5, 'temperature': None, 'pressure': 'Normal', 'bond_amount': None, 'uv_energy': None, 'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!B20:E20']},
    {'condition_id': 'cond_4', 'condition_group': '7S-7S + height check + New line + Normal press', 'line': '', 'process': "Ass'y Coil+SP", 'changed_factor': 'New bonding line', 'before_value': 'Old bonding line', 'after_value': '7S-7S / use height check / New bonding line / Normal press', 'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None, 'dry_time_sec': 7.0, 'temperature': None, 'pressure': 'Normal', 'bond_amount': None, 'uv_energy': None, 'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!B21:E21']},
    {'condition_id': 'cond_5', 'condition_group': '7S-7S + height check + New line + Press down', 'line': '', 'process': "Ass'y Coil+SP", 'changed_factor': 'UC Press down', 'before_value': 'Normal press', 'after_value': '7S-7S / use height check / New line / Press down', 'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None, 'dry_time_sec': 7.0, 'temperature': None, 'pressure': 'Press down', 'bond_amount': None, 'uv_energy': None, 'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!B22:E22']},
    {'condition_id': 'cond_6', 'condition_group': "7S-7S + don't use height check + New line + Press down", 'line': '', 'process': "Ass'y Coil+SP", 'changed_factor': "Don't use height check + UC Press down + New line", 'before_value': None, 'after_value': "7S-7S / don't use height check / New line / Press down", 'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None, 'dry_time_sec': 7.0, 'temperature': None, 'pressure': 'Press down', 'bond_amount': None, 'uv_energy': None, 'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!B23:E23']},
    {'condition_id': 'cond_7', 'condition_group': '7S-7S + height check + New line + Press down (re-confirm)', 'line': '', 'process': "Ass'y Coil+SP", 'changed_factor': 'Re-confirm with height check', 'before_value': None, 'after_value': '7S-7S / use height check / New line / Press down', 'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None, 'dry_time_sec': 7.0, 'temperature': None, 'pressure': 'Press down', 'bond_amount': None, 'uv_energy': None, 'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!B24:E24']},
  ],
  'results': [
    {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': "Ass'y Coil+SP", 'condition_group': '7S-7S + height check + Old line + Normal', 'date': '2024-04-02', 'line': '', 'input_count': 100, 'ok_count': 97, 'ng_count': 3, 'ng_rate_decimal': 0.03, 'ng_rate_percent': 3.0, 'metric_name': 'Coil+SP separate rate', 'metric_value': 3.0, 'unit': '%', 'judgement': None, 'ng_breakdown': {'Press 1': 3, 'Press 2': 0, 'Press 3': 0, 'Press 4': 0}, 'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!F18:N18']},
    {'result_id': 'res_2', 'condition_id': 'cond_2', 'measurement_type': "Ass'y Coil+SP", 'condition_group': '7.5S-7.5S + height check + Old + Normal', 'date': '2024-04-02', 'line': '', 'input_count': 24, 'ok_count': 22, 'ng_count': 2, 'ng_rate_decimal': 0.0833, 'ng_rate_percent': 8.33, 'metric_name': 'Coil+SP separate rate', 'metric_value': 8.33, 'unit': '%', 'judgement': None, 'ng_breakdown': {'Press 1': 2, 'Press 2': 0, 'Press 3': 0, 'Press 4': 0}, 'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!F19:N19']},
    {'result_id': 'res_3', 'condition_id': 'cond_3', 'measurement_type': "Ass'y Coil+SP", 'condition_group': "7.5S-7.5S + no height check + Old + Normal", 'date': '2024-04-02', 'line': '', 'input_count': 72, 'ok_count': 72, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Coil+SP separate rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS', 'ng_breakdown': {}, 'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!F20:N20']},
    {'result_id': 'res_4', 'condition_id': 'cond_4', 'measurement_type': "Ass'y Coil+SP", 'condition_group': '7S-7S + height check + New + Normal', 'date': '2024-04-02', 'line': '', 'input_count': 128, 'ok_count': 123, 'ng_count': 5, 'ng_rate_decimal': 0.0391, 'ng_rate_percent': 3.91, 'metric_name': 'Coil+SP separate rate', 'metric_value': 3.91, 'unit': '%', 'judgement': None, 'ng_breakdown': {'Press 1': 5, 'Press 2': 0, 'Press 3': 0, 'Press 4': 0}, 'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!F21:N21']},
    {'result_id': 'res_5', 'condition_id': 'cond_5', 'measurement_type': "Ass'y Coil+SP", 'condition_group': '7S-7S + height check + New + Press down', 'date': '2024-04-02', 'line': '', 'input_count': 136, 'ok_count': 130, 'ng_count': 6, 'ng_rate_decimal': 0.0441, 'ng_rate_percent': 4.41, 'metric_name': 'Coil+SP separate rate', 'metric_value': 4.41, 'unit': '%', 'judgement': None, 'ng_breakdown': {'Press 1': 6, 'Press 2': 0, 'Press 3': 0, 'Press 4': 0}, 'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!F22:N22']},
    {'result_id': 'res_6', 'condition_id': 'cond_6', 'measurement_type': "Ass'y Coil+SP", 'condition_group': "7S-7S + no height check + New + Press down", 'date': '2024-04-02', 'line': '', 'input_count': 222, 'ok_count': 220, 'ng_count': 2, 'ng_rate_decimal': 0.009, 'ng_rate_percent': 0.9, 'metric_name': 'Coil+SP separate rate', 'metric_value': 0.9, 'unit': '%', 'judgement': None, 'ng_breakdown': {'Press 1': 1, 'Press 2': 0, 'Press 3': 1, 'Press 4': 0}, 'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!F23:N23']},
    {'result_id': 'res_7', 'condition_id': 'cond_7', 'measurement_type': "Ass'y Coil+SP", 'condition_group': '7S-7S + height check + New + Press down (recheck)', 'date': '2024-04-02', 'line': '', 'input_count': 220, 'ok_count': 217, 'ng_count': 3, 'ng_rate_decimal': 0.0136, 'ng_rate_percent': 1.36, 'metric_name': 'Coil+SP separate rate', 'metric_value': 1.36, 'unit': '%', 'judgement': None, 'ng_breakdown': {'Press 1': 2, 'Press 2': 0, 'Press 3': 1, 'Press 4': 0}, 'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!F24:N24']},
  ],
  'conclusions': [
    {'conclusion_id': 'concl_1', 'topic': 'Effect of "do not use height check"',
     'statement_from_report': '(Highlighted rows: don\'t use height check)',
     'normalized_interpretation': 'Removing height check sharply reduces separation NG: 7.5S/Old line/Normal 8.33% (use HC) -> 0.0% (no HC); 7S/New line/Press down 4.41% (use HC) -> 0.90% (no HC). Adding height check appears to introduce the separation rather than detect it.',
     'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!B20:N20', '5.3!B23:N23']},
    {'conclusion_id': 'concl_2', 'topic': 'Press down vs Normal press (with height check)',
     'statement_from_report': '(Conditions 4 vs 5)',
     'normalized_interpretation': '7S-7S + height check + New line: Normal press 3.91% vs Press down 4.41% = (4.41/3.91-1)*100 = +12.8% worse with Press down. UC press down does not reduce Coil+SP separation in the same-event comparison.',
     'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!B21:N22']},
    {'conclusion_id': 'concl_3', 'topic': 'Press number where defect occurs',
     'statement_from_report': '(Press 1-4 breakdown)',
     'normalized_interpretation': 'Across all 7 conditions, NG is concentrated at Press 1 (sum 19 of 21 NG); Press 3 has 2 NG, Press 2 and 4 are zero. Press 1 station is the dominant separation source.',
     'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!F18:N24']},
  ],
  'troubleshooting_index': {
    'defect_name': 'Coil+SP Separation',
    'when_user_asks': ['Coil+SP separate', 'UC press time effect on separation', 'height check vs separation NG'],
    'suggested_checks': [
      {'hint_id': 'hint_1', 'check_item': "Audit height check fixture for Coil+SP - confirm it is not the cause of separation", 'reason': "Two paired conditions show: with height check 8.33% and 4.41%, without height check 0.0% and 0.90%; same line / time / press otherwise. Height check appears causal.", 'evidence_strength': 'high', 'related_process': "Ass'y Coil+SP", 'related_part': 'Coil/SP/Height-check jig', 'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!B19:N23']},
      {'hint_id': 'hint_2', 'check_item': 'Inspect Press 1 station for over-press / misalign', 'reason': 'NG location breakdown: Press 1 accounts for 19/21 = 90.5% of all Coil+SP separation NG across the seven conditions.', 'evidence_strength': 'high', 'related_process': 'UC Press', 'related_part': 'Press 1 head', 'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!F18:N24']},
      {'hint_id': 'hint_3', 'check_item': 'Reconsider UC Press down setting', 'reason': '7S/New line/HC: Press down 4.41% vs Normal 3.91% (+12.8% worse).', 'evidence_strength': 'medium', 'related_process': 'UC Press', 'related_part': 'Coil+SP', 'source_file': name06, 'sheet_name': '5.3', 'source_cells': ['5.3!B21:N22']},
    ],
    'limitations': ['Conditions are not balanced (time, line, press, height check all vary); not a clean DOE.', 'No decision text section is filled in.']
  },
  'ai_extraction_log': {
    'confidence': 0.7,
    'assumptions': ['Each row is a different test condition combining 4 factors (Time, Height check use, Bonding line, UC press setting).', 'NG counts at Press 1-4 are mutually exclusive locations.'],
    'warnings': ['Conditions are factorial-style but not orthogonal; cannot isolate single-factor effect with confidence.', 'Sample sizes range 24-222, biasing comparisons.'],
    'decision_rationale': "Two same-condition pairs differing only by height check use show NG drops to 0.0% / 0.9% when height check is removed. Press 1 accounts for >=90% of all separation events. Most plausible cause: height check fixture interferes with Coil+SP at Press 1 station rather than UC press time or line position."}
}

tr_ko_06 = {
  'document': {'title': 'BRS-161016 Coil+Suspension Separation 문제 확인 리포트', 'purpose': 'Coil+SP 분리 NG 원인 파악.',
               'content': ['PE UC press 조정 후 확인', '새 holder 교체 확인', 'UC press와 LED UV 시간 증가', "Ass'y Coil+SP에서 height check 사용/미사용 비교", '본딩 라인 inside 이동']},
  'conclusions': {
    'concl_1': {'topic': 'Height check 사용 여부 효과', 'statement_from_report': '(Height check 미사용 행이 강조됨)', 'normalized_interpretation': 'Height check 제거 시 NG 급감: 7.5S/Old/Normal HC 8.33% → no HC 0.0%; 7S/New/Press down HC 4.41% → no HC 0.90%. Height check가 결함을 감지하기보단 유발하는 것으로 보임.'},
    'concl_2': {'topic': 'Press down vs Normal press', 'statement_from_report': '(조건 4 vs 5)', 'normalized_interpretation': '7S-7S+HC+New: Normal 3.91% vs Press down 4.41% = +12.8% 악화. UC press down은 Coil+SP separation 감소에 도움 안 됨.'},
    'concl_3': {'topic': 'NG 발생 Press 위치', 'statement_from_report': '(Press 1~4 breakdown)', 'normalized_interpretation': '7개 조건 전체 NG 21건 중 Press 1에서 19건(90.5%); Press 3 2건, Press 2/4 0건. Press 1이 주 원인.'},
  },
  'hints': {
    'hint_1': {'check_item': 'Coil+SP height check fixture가 분리의 원인인지 점검', 'reason': '쌍 조건: HC 사용 8.33%/4.41% → HC 미사용 0.0%/0.90%. 다른 조건 동일.'},
    'hint_2': {'check_item': 'Press 1 station 과압/오정렬 점검', 'reason': '전체 NG 21건 중 19건(90.5%)이 Press 1에서 발생.'},
    'hint_3': {'check_item': 'UC Press down 설정 재검토', 'reason': '7S/New/HC: Press down 4.41% vs Normal 3.91% (+12.8% 악화).'},
  },
  'log': {'assumptions': ['각 행은 4개 인자(Time, HC 사용, 라인, UC press) 조합.', 'Press 1~4 위치는 상호 배타적.'],
          'warnings': ['Factorial이지만 직교적이지 않음; 단일 인자 효과 단정 어려움.', '샘플 크기 24~222로 편향 가능.'],
          'decision_rationale': 'HC 사용/미사용 짝 비교에서 HC 미사용 시 NG 0.0%/0.9%로 급감. NG의 90%가 Press 1에서 발생. → 가장 가능성 높은 원인: HC fixture가 Press 1 station에서 Coil+SP에 간섭.'}
}
tr_en_06 = {
  'document': {'title': 'BRS-161016 Report Check Problem Separate Coil+Suspension', 'purpose': 'Check the reason for Coil+SP separation NG.',
               'content': ['PE adjust UC press down and check', 'PE change new holder and check', 'Increase UC press / LED UV time', "Test with/without height check at ass'y Coil+SP", 'Test change bonding line position']},
  'conclusions': {
    'concl_1': {'topic': 'Effect of height check', 'statement_from_report': '(Highlighted rows: do not use height check)', 'normalized_interpretation': "Removing height check sharply reduces NG: 7.5S/Old/Normal HC 8.33% -> no HC 0.0%; 7S/New/Press down HC 4.41% -> no HC 0.90%. Height check looks causal, not detective."},
    'concl_2': {'topic': 'Press down vs Normal press', 'statement_from_report': '(conditions 4 vs 5)', 'normalized_interpretation': '7S-7S+HC+New: Normal 3.91% vs Press down 4.41% = +12.8% worse. UC press down does not help.'},
    'concl_3': {'topic': 'NG press number location', 'statement_from_report': '(Press 1-4 breakdown)', 'normalized_interpretation': 'Across 7 conditions, 19/21 NG (90.5%) occur at Press 1; Press 3 has 2, Press 2/4 zero. Press 1 is dominant.'},
  },
  'hints': {
    'hint_1': {'check_item': 'Audit height check fixture - confirm not the cause of separation', 'reason': "Paired conditions: HC used 8.33%/4.41% vs no HC 0.0%/0.90% with everything else equal."},
    'hint_2': {'check_item': 'Inspect Press 1 station for over-press / misalign', 'reason': '19/21 = 90.5% of separation NG happens at Press 1.'},
    'hint_3': {'check_item': 'Reconsider UC press down setting', 'reason': '7S/New/HC: Press down 4.41% vs Normal 3.91% (+12.8% worse).'},
  },
  'log': {'assumptions': ['Each row is a 4-factor combination (Time, HC use, Line, UC press).', 'Press 1-4 NG locations are mutually exclusive.'],
          'warnings': ['Factor combinations are not orthogonal; cannot isolate single-factor effect with confidence.', 'Sample size varies 24-222.'],
          'decision_rationale': "Paired HC vs no-HC comparisons show NG drops to 0.0%/0.9% without HC. >=90% of all separation NG occurs at Press 1. Most likely root cause: height check fixture interferes with Coil+SP at Press 1 station."}
}
tr_vi_06 = {
  'document': {'title': 'BRS-161016 Báo cáo kiểm tra vấn đề tách Coil+Suspension', 'purpose': 'Kiểm tra nguyên nhân NG tách Coil+SP.',
               'content': ['PE chỉnh UC press down và kiểm tra', 'PE đổi holder mới và kiểm tra', 'Tăng thời gian UC press và LED UV', "Test có/không dùng height check tại ass'y Coil+SP", 'Đổi vị trí bonding line']},
  'conclusions': {
    'concl_1': {'topic': 'Tác dụng của height check', 'statement_from_report': '(Hàng được tô vàng: không dùng height check)', 'normalized_interpretation': 'Bỏ height check NG giảm mạnh: 7.5S/Old/Normal HC 8.33% → no HC 0.0%; 7S/New/Press down HC 4.41% → no HC 0.90%. Height check có vẻ là nguyên nhân, không phải để phát hiện.'},
    'concl_2': {'topic': 'Press down vs Normal press', 'statement_from_report': '(điều kiện 4 vs 5)', 'normalized_interpretation': '7S-7S+HC+New: Normal 3.91% vs Press down 4.41% = +12.8% xấu hơn. UC press down không giúp.'},
    'concl_3': {'topic': 'Press số mấy gây NG', 'statement_from_report': '(Phân tích Press 1-4)', 'normalized_interpretation': '7 điều kiện, 19/21 NG (90.5%) ở Press 1; Press 3 có 2, Press 2/4 0. Press 1 chiếm chính.'},
  },
  'hints': {
    'hint_1': {'check_item': 'Kiểm tra fixture height check có gây tách Coil+SP không', 'reason': 'Cặp đối chiếu: HC dùng 8.33%/4.41% vs không HC 0.0%/0.90%, các yếu tố khác bằng.'},
    'hint_2': {'check_item': 'Kiểm tra Press 1 quá lực / lệch', 'reason': '19/21 = 90.5% NG xảy ra ở Press 1.'},
    'hint_3': {'check_item': 'Xem lại setting UC Press down', 'reason': '7S/New/HC: Press down 4.41% vs Normal 3.91% (+12.8% xấu).'},
  },
  'log': {'assumptions': ['Mỗi hàng là tổ hợp 4 yếu tố (Time, HC, Line, UC press).', 'Press 1-4 là vị trí loại trừ lẫn nhau.'],
          'warnings': ['Các điều kiện không trực giao; khó tách hiệu ứng đơn yếu tố.', 'Cỡ mẫu 24-222 không cân bằng.'],
          'decision_rationale': 'Cặp HC vs không-HC cho NG giảm còn 0.0%/0.9% khi bỏ HC. >=90% NG xảy ra ở Press 1. Nguyên nhân nhiều khả năng: fixture height check va chạm với Coil+SP tại Press 1.'}
}
results[name06] = (result06, tr_ko_06, tr_en_06, tr_vi_06)

# =====================================================================
# DS07: 27. BRS-201506 Report checking and test SUB2  date 12.3.2024
# =====================================================================
name07 = '27. BRS-201506 Report checking and test SUB2  date 12.3.2024'

result07 = {
  'schema_version': '0.1',
  'document': {
    'document_id': '', 'source_file': name07, 'source_sheet': 'Yoke class 1~3',
    'title': 'BRS-201506 Report Test Yoke NG Vision at Sub2',
    'model': 'BRS-201506', 'report_date': '2024-03-12',
    'department': 'ME', 'marker': 'Thao', 'line': '',
    'report_type': 'normal_comparison',
    'primary_defect': {'canonical_name': 'NG Yoke Vision (Class 2/3)', 'aliases_in_document': ['Yoke NG vision sub2', 'Class 2/3 Yoke']},
    'related_defects': ['NG Hearing Noise', 'NG Hearing Touch'],
    'parts': ['Yoke', 'Yoke jig'], 'processes': ['Sub2 vision', 'Function'],
    'purpose': 'Decide whether NG Yoke (vision) of class 1-3 can be used.',
    'content': ['Check Yoke class 1-3 rate', 'Compare function NG rate of Yoke class 1-3'],
    'source_cells': {'title': ['Yoke class 1~3!B1'], 'date': ['Yoke class 1~3!P2'], 'purpose': ['Yoke class 1~3!A4'], 'content': ['Yoke class 1~3!A6:A7']}
  },
  'test_conditions': [
    {'condition_id': 'cond_1', 'condition_group': 'Normal jig (new machining)', 'line': '', 'process': 'Yoke class check (Sub2)', 'changed_factor': 'Jig (Normal jig new machining)', 'before_value': None, 'after_value': 'Normal jig 신규 가공', 'unit': None, 'machine': None, 'jig': 'Normal jig (new machining)', 'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name07, 'sheet_name': 'Yoke class 1~3', 'source_cells': ['Yoke class 1~3!C14']},
    {'condition_id': 'cond_2', 'condition_group': 'Test jig (new machining)', 'line': '', 'process': 'Yoke class check (Sub2)', 'changed_factor': 'Jig (Test jig new machining)', 'before_value': 'Normal jig', 'after_value': 'Test jig 신규 가공', 'unit': None, 'machine': None, 'jig': 'Test jig (new machining)', 'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name07, 'sheet_name': 'Yoke class 1~3', 'source_cells': ['Yoke class 1~3!C19']},
    {'condition_id': 'cond_3', 'condition_group': 'Function 2024-04-09 New JIG', 'line': '', 'process': 'Function', 'changed_factor': 'Function with New JIG', 'before_value': 'Normal line', 'after_value': 'New JIG', 'unit': None, 'machine': None, 'jig': 'New JIG', 'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name07, 'sheet_name': 'Yoke class 1~3', 'source_cells': ['Yoke class 1~3!B26']},
    {'condition_id': 'cond_4', 'condition_group': 'Function 2024-04-09 Normal line', 'line': '', 'process': 'Function', 'changed_factor': 'Baseline (Normal line)', 'before_value': None, 'after_value': 'Normal line', 'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name07, 'sheet_name': 'Yoke class 1~3', 'source_cells': ['Yoke class 1~3!B28']},
    {'condition_id': 'cond_5', 'condition_group': 'Function 2024-04-11 New JIG', 'line': '', 'process': 'Function', 'changed_factor': 'Function with New JIG (different day)', 'before_value': 'Normal line', 'after_value': 'New JIG', 'unit': None, 'machine': None, 'jig': 'New JIG', 'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name07, 'sheet_name': 'Yoke class 1~3', 'source_cells': ['Yoke class 1~3!B30']},
    {'condition_id': 'cond_6', 'condition_group': 'Function 2024-04-11 Normal line', 'line': '', 'process': 'Function', 'changed_factor': 'Baseline (Normal line)', 'before_value': None, 'after_value': 'Normal line', 'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name07, 'sheet_name': 'Yoke class 1~3', 'source_cells': ['Yoke class 1~3!B32']},
  ],
  'results': [
    {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Yoke class distribution', 'condition_group': 'Normal jig 2024-03-13', 'date': '2024-03-13', 'line': '', 'input_count': 215, 'ok_count': 201, 'ng_count': 14, 'ng_rate_decimal': None, 'ng_rate_percent': None, 'metric_name': 'Yoke class distribution', 'metric_value': None, 'unit': '%', 'judgement': None, 'ng_breakdown': {'Class 1 input': 202, 'Class 2 input': 9, 'Class 3 input': 4, 'Class 1 rate %': 95.5, 'Class 2 rate %': 2.5, 'Class 3 rate %': 1.5, 'Class 1 OK rate %': 95.5, 'Class 2 OK rate %': 98.0, 'Class 3 OK rate %': 99.5}, 'source_file': name07, 'sheet_name': 'Yoke class 1~3', 'source_cells': ['Yoke class 1~3!E14:H18']},
    {'result_id': 'res_2', 'condition_id': 'cond_2', 'measurement_type': 'Yoke class distribution', 'condition_group': 'Test jig 2024-03-13', 'date': '2024-03-13', 'line': '', 'input_count': 205, 'ok_count': 203, 'ng_count': 2, 'ng_rate_decimal': None, 'ng_rate_percent': None, 'metric_name': 'Yoke class distribution', 'metric_value': None, 'unit': '%', 'judgement': None, 'ng_breakdown': {'Class 1 input': 203, 'Class 2 input': 1, 'Class 3 input': 1, 'Class 1 rate %': 99.5, 'Class 2 rate %': 0.0, 'Class 3 rate %': 0.5, 'Class 1 OK rate %': 99.5, 'Class 2 OK rate %': 99.5, 'Class 3 OK rate %': 100.0}, 'source_file': name07, 'sheet_name': 'Yoke class 1~3', 'source_cells': ['Yoke class 1~3!E19:H23']},
    {'result_id': 'res_3', 'condition_id': 'cond_3', 'measurement_type': 'Function', 'condition_group': 'New JIG 2024-04-09', 'date': '2024-04-09', 'line': '', 'input_count': 303, 'ok_count': 301, 'ng_count': 2, 'ng_rate_decimal': 0.007, 'ng_rate_percent': 0.7, 'metric_name': 'Function NG rate', 'metric_value': 0.7, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG Hearing Noise': 1, 'NG Hearing Touch': 1, 'NG Sigma SPL': 0, 'NG Sigma THD': 0}, 'source_file': name07, 'sheet_name': 'Yoke class 1~3', 'source_cells': ['Yoke class 1~3!D26:M26']},
    {'result_id': 'res_4', 'condition_id': 'cond_4', 'measurement_type': 'Function', 'condition_group': 'Normal line 2024-04-09', 'date': '2024-04-09', 'line': '', 'input_count': 299, 'ok_count': 294, 'ng_count': 5, 'ng_rate_decimal': 0.017, 'ng_rate_percent': 1.7, 'metric_name': 'Function NG rate', 'metric_value': 1.7, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG Hearing Noise': 2, 'NG Hearing Touch': 3, 'NG Sigma SPL': 0, 'NG Sigma THD': 0}, 'source_file': name07, 'sheet_name': 'Yoke class 1~3', 'source_cells': ['Yoke class 1~3!D28:M28']},
    {'result_id': 'res_5', 'condition_id': 'cond_5', 'measurement_type': 'Function', 'condition_group': 'New JIG 2024-04-11', 'date': '2024-04-11', 'line': '', 'input_count': 299, 'ok_count': 285, 'ng_count': 14, 'ng_rate_decimal': 0.047, 'ng_rate_percent': 4.7, 'metric_name': 'Function NG rate', 'metric_value': 4.7, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG Hearing Noise': 3, 'NG Hearing Touch': 11, 'NG Sigma SPL': 0, 'NG Sigma THD': 0}, 'source_file': name07, 'sheet_name': 'Yoke class 1~3', 'source_cells': ['Yoke class 1~3!D30:M30']},
    {'result_id': 'res_6', 'condition_id': 'cond_6', 'measurement_type': 'Function', 'condition_group': 'Normal line 2024-04-11', 'date': '2024-04-11', 'line': '', 'input_count': 304, 'ok_count': 290, 'ng_count': 14, 'ng_rate_decimal': 0.046, 'ng_rate_percent': 4.6, 'metric_name': 'Function NG rate', 'metric_value': 4.6, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG Hearing Noise': 2, 'NG Hearing Touch': 12, 'NG Sigma SPL': 0, 'NG Sigma THD': 0}, 'source_file': name07, 'sheet_name': 'Yoke class 1~3', 'source_cells': ['Yoke class 1~3!D32:M32']},
  ],
  'conclusions': [
    {'conclusion_id': 'concl_1', 'topic': 'Yoke class distribution by jig',
     'statement_from_report': '(Normal vs Test jig rate)',
     'normalized_interpretation': 'Normal jig: Class1 95.5%, Class2 2.5%, Class3 1.5% (class2/3 = 4.0% of input). Test jig: Class1 99.5%, Class2 0.0%, Class3 0.5% (class2/3 = 0.5%). Test jig (new machining) clearly reduces Class 2/3 Yoke detection rate vs Normal jig.',
     'source_file': name07, 'sheet_name': 'Yoke class 1~3', 'source_cells': ['Yoke class 1~3!E14:H23']},
    {'conclusion_id': 'concl_2', 'topic': 'Function 2024-04-09 New JIG vs Normal line',
     'statement_from_report': '(Function table 2024-04-09)',
     'normalized_interpretation': 'Function NG rate New JIG 0.7% vs Normal line 1.7% = (0.7/1.7-1)*100 = -58.8% improved on 2024-04-09 (small sample n=303 vs 299). Noise/Touch are both improved.',
     'source_file': name07, 'sheet_name': 'Yoke class 1~3', 'source_cells': ['Yoke class 1~3!D26:M28']},
    {'conclusion_id': 'concl_3', 'topic': 'Function 2024-04-11 New JIG vs Normal line',
     'statement_from_report': '(Function table 2024-04-11)',
     'normalized_interpretation': 'Function NG rate New JIG 4.7% vs Normal line 4.6% = (4.7/4.6-1)*100 = +2.2% essentially same on 2024-04-11. Hearing Touch dominates both (3.7% vs 3.9%).',
     'source_file': name07, 'sheet_name': 'Yoke class 1~3', 'source_cells': ['Yoke class 1~3!D30:M32']},
  ],
  'troubleshooting_index': {
    'defect_name': 'NG Yoke Vision (Class 2/3 at Sub2)',
    'when_user_asks': ['Yoke class 1-3 use or not', 'NG Yoke vision sub2', 'Yoke jig change effect'],
    'suggested_checks': [
      {'hint_id': 'hint_1', 'check_item': 'Confirm Yoke class measurement consistency between Normal and Test jig', 'reason': 'Normal jig flags 4.0% of input as Class 2/3, Test jig only 0.5%. Either the new-machined Test jig classifies more strictly or material lot changed; needs cross-confirmation.', 'evidence_strength': 'medium', 'related_process': 'Sub2 vision (Yoke class)', 'related_part': 'Yoke jig', 'source_file': name07, 'sheet_name': 'Yoke class 1~3', 'source_cells': ['Yoke class 1~3!E14:H23']},
      {'hint_id': 'hint_2', 'check_item': 'Re-run function test on larger Yoke class 2/3 sample to confirm equivalence', 'reason': 'Two events compared: 04-09 New JIG 0.7% vs Normal 1.7% (-58.8% improved); 04-11 New JIG 4.7% vs Normal 4.6% (+2.2% same). Results are not stable across days.', 'evidence_strength': 'low', 'related_process': 'Function test', 'related_part': 'Yoke / module', 'source_file': name07, 'sheet_name': 'Yoke class 1~3', 'source_cells': ['Yoke class 1~3!D26:M32']},
      {'hint_id': 'hint_3', 'check_item': 'Track dominant NG mode (Touch) at function station', 'reason': 'On 2024-04-11 both New JIG (3.7%) and Normal (3.9%) are dominated by NG Hearing Touch, not Yoke class problem; root cause may be unrelated to Yoke.', 'evidence_strength': 'medium', 'related_process': 'Function test', 'related_part': 'Speaker module', 'source_file': name07, 'sheet_name': 'Yoke class 1~3', 'source_cells': ['Yoke class 1~3!D30:M32']},
    ],
    'limitations': ['Two function days give inconsistent results.', 'No decision text section.']
  },
  'ai_extraction_log': {
    'confidence': 0.65,
    'assumptions': ['Class rate row sums to 100% for each jig.', 'Class 1 OK rate vs Class rate columns are interpreted as cumulative pass rate.', 'Total Input for Normal jig is Class1+Class2+Class3 = 215; for Test jig = 205.'],
    'warnings': ['Function NG rate ratio New JIG vs Normal line changes from -58.8% (04-09) to +2.2% (04-11); single-day result is unreliable.', 'Day 04-11 has Hearing Touch as dominant NG for both conditions, unrelated to Yoke.'],
    'decision_rationale': 'Test jig (new machining) classifies fewer Yoke as Class 2/3 (0.5% vs 4.0%); could be either improved jig or under-detection. Function comparison is inconsistent across days, so Yoke class change is not yet shown to affect function reliably. Recommend further test before release.'}
}

tr_ko_07 = {
  'document': {'title': 'BRS-201506 Sub2 Yoke NG Vision 사용 가능 여부 시험 리포트', 'purpose': 'NG Yoke (vision) Class 1~3 사용 가능 여부 결정.',
               'content': ['Yoke Class 1~3 비율 확인', 'Yoke Class 1~3 function NG rate 비교']},
  'conclusions': {
    'concl_1': {'topic': 'Jig 별 Yoke class 분포', 'statement_from_report': '(Normal vs Test jig)', 'normalized_interpretation': 'Normal jig: Class1 95.5%, Class2 2.5%, Class3 1.5% (Class2/3 합 4.0%). Test jig: Class1 99.5%, Class2 0.0%, Class3 0.5% (합 0.5%). Test jig가 Class 2/3 검출률을 크게 낮춤.'},
    'concl_2': {'topic': 'Function 04-09 New JIG vs Normal', 'statement_from_report': '(04-09 Function)', 'normalized_interpretation': 'New JIG 0.7% vs Normal 1.7% = -58.8% 개선 (04-09, n=303 vs 299).'},
    'concl_3': {'topic': 'Function 04-11 New JIG vs Normal', 'statement_from_report': '(04-11 Function)', 'normalized_interpretation': 'New JIG 4.7% vs Normal 4.6% = +2.2% 동등. Touch가 dominant (3.7% vs 3.9%).'},
  },
  'hints': {
    'hint_1': {'check_item': 'Normal jig와 Test jig의 Yoke class 측정 일관성 확인', 'reason': 'Normal jig는 4.0%를 Class 2/3로 분류, Test jig는 0.5%만 분류. 자재 변동인지 jig 차이인지 cross-check 필요.'},
    'hint_2': {'check_item': 'Yoke class 2/3 함유 sample로 function NG 재시험', 'reason': '04-09 -58.8% 개선 vs 04-11 +2.2% 동등. 결과가 날짜에 따라 불안정.'},
    'hint_3': {'check_item': 'Function station Touch NG 원인 추적', 'reason': '04-11에는 New JIG와 Normal 모두 Touch dominant (3.7% vs 3.9%); Yoke와 무관할 수 있음.'},
  },
  'log': {'assumptions': ['Class rate 합 100%.', 'Normal jig Input 215, Test jig 205.'],
          'warnings': ['Function NG ratio가 04-09 -58.8% / 04-11 +2.2%로 불안정.', '04-11 dominant NG는 Touch.'],
          'decision_rationale': 'Test jig가 Class 2/3 검출률을 크게 낮춤 → jig 개선인지 under-detection인지 미확정. Function 비교는 일관성 없음 → release 전 추가 시험 필요.'}
}
tr_en_07 = {
  'document': {'title': 'BRS-201506 Report Test Yoke NG Vision at Sub2', 'purpose': 'Decide whether NG Yoke (vision) Class 1-3 can be used.',
               'content': ['Check Yoke Class 1-3 rate', 'Compare function NG rate of Yoke Class 1-3']},
  'conclusions': {
    'concl_1': {'topic': 'Yoke class distribution by jig', 'statement_from_report': '(Normal vs Test jig)', 'normalized_interpretation': 'Normal jig: Class1 95.5%, Class2 2.5%, Class3 1.5% (Class 2/3 sum 4.0%). Test jig: Class1 99.5%, Class2 0.0%, Class3 0.5% (sum 0.5%). Test jig dramatically lowers Class 2/3 detection.'},
    'concl_2': {'topic': 'Function 04-09 New JIG vs Normal', 'statement_from_report': '(04-09 Function)', 'normalized_interpretation': 'New JIG 0.7% vs Normal 1.7% = -58.8% improved (n=303 vs 299).'},
    'concl_3': {'topic': 'Function 04-11 New JIG vs Normal', 'statement_from_report': '(04-11 Function)', 'normalized_interpretation': 'New JIG 4.7% vs Normal 4.6% = +2.2% essentially same. Touch dominates both (3.7% vs 3.9%).'},
  },
  'hints': {
    'hint_1': {'check_item': 'Confirm Yoke class measurement consistency between Normal and Test jig', 'reason': 'Normal jig flags 4.0% as Class 2/3, Test jig only 0.5%; need cross-check whether material or jig.'},
    'hint_2': {'check_item': 'Re-run function with larger Yoke Class 2/3 sample', 'reason': '04-09 -58.8% improved vs 04-11 +2.2% same; results not stable across days.'},
    'hint_3': {'check_item': 'Track dominant NG mode (Touch) at function station', 'reason': '04-11 dominant NG is Hearing Touch for both jigs (3.7% vs 3.9%), unrelated to Yoke.'},
  },
  'log': {'assumptions': ['Class rate sums to 100%.', 'Normal jig Input 215, Test jig 205.'],
          'warnings': ['Function NG ratio swings between -58.8% (04-09) and +2.2% (04-11); single-day result unreliable.', '04-11 dominant NG is Touch, not Yoke.'],
          'decision_rationale': 'Test jig classifies far fewer parts as Class 2/3 - could be improved jig or under-detection. Function comparison is inconsistent. Recommend more tests before release.'}
}
tr_vi_07 = {
  'document': {'title': 'BRS-201506 Báo cáo test Yoke NG Vision tại Sub2', 'purpose': 'Quyết định Yoke NG vision Class 1-3 có dùng được không.',
               'content': ['Kiểm tra tỉ lệ Yoke Class 1-3', 'So sánh function NG rate Yoke Class 1-3']},
  'conclusions': {
    'concl_1': {'topic': 'Phân bố Yoke class theo jig', 'statement_from_report': '(Normal vs Test jig)', 'normalized_interpretation': 'Normal jig: Class1 95.5%, Class2 2.5%, Class3 1.5% (Class2/3 4.0%). Test jig: Class1 99.5%, Class2 0.0%, Class3 0.5% (0.5%). Test jig giảm mạnh Class 2/3.'},
    'concl_2': {'topic': 'Function 04-09 New JIG vs Normal', 'statement_from_report': '(Function 04-09)', 'normalized_interpretation': 'New JIG 0.7% vs Normal 1.7% = -58.8% cải thiện (n=303 vs 299).'},
    'concl_3': {'topic': 'Function 04-11 New JIG vs Normal', 'statement_from_report': '(Function 04-11)', 'normalized_interpretation': 'New JIG 4.7% vs Normal 4.6% = +2.2% bằng. Touch dominant cả hai (3.7% vs 3.9%).'},
  },
  'hints': {
    'hint_1': {'check_item': 'Xác nhận tính nhất quán đo Yoke class giữa Normal và Test jig', 'reason': 'Normal jig phát hiện 4.0% Class 2/3, Test jig chỉ 0.5%; cần cross-check.'},
    'hint_2': {'check_item': 'Test lại function với sample Yoke Class 2/3 lớn hơn', 'reason': '04-09 -58.8% cải thiện vs 04-11 +2.2% bằng; không ổn định.'},
    'hint_3': {'check_item': 'Theo dõi NG Touch chủ đạo ở function', 'reason': '04-11 Touch dominant cả 2 jig (3.7% vs 3.9%), không liên quan Yoke.'},
  },
  'log': {'assumptions': ['Class rate tổng 100%.', 'Normal jig Input 215, Test jig 205.'],
          'warnings': ['Tỉ số function NG dao động -58.8% (04-09) và +2.2% (04-11); 1 ngày không tin được.', '04-11 NG chủ đạo là Touch, không phải Yoke.'],
          'decision_rationale': 'Test jig phân loại Class 2/3 ít hơn nhiều — có thể là jig tốt hơn hoặc under-detection. So sánh function không ổn định. Đề nghị thêm test trước khi phát hành.'}
}
results[name07] = (result07, tr_ko_07, tr_en_07, tr_vi_07)

# =====================================================================
# DS08: 27. MSU-L20S15-07 GMI Report test material YOKE L20A15-07 happen NG dyne test - 2025.06.27
# =====================================================================
name08 = '27. MSU-L20S15-07 GMI  Report test material YOKE L20A15-07 happen NG dyne test - 2025.06.27'

result08 = {
  'schema_version': '0.1',
  'document': {
    'document_id': '', 'source_file': name08, 'source_sheet': 'Sheet',
    'title': 'MSU-L20S15-07 Report Test Material YOKE L20S15-07 Happen NG Dyne Test',
    'model': 'MSU-L20S15-07', 'report_date': '2025-06-27',
    'department': 'ME', 'marker': 'Nhung', 'line': 'E2-4B',
    'report_type': 'reliability_spec',
    'primary_defect': {'canonical_name': 'NG Yoke Dyne Test', 'aliases_in_document': ['NG dyne test 6/10']},
    'related_defects': ['NG Yoke+SM bond not enough', 'NG Yoke+CM bond'],
    'parts': ['YOKE L20S15-07', 'SM', 'CM'], 'processes': ['Sub2 NG process', 'Decap check bond', 'Drop test', 'Tension test', 'Function'],
    'purpose': 'Decide whether Yoke lot that failed dyne test 6/10 (~60%) can be used.',
    'content': ['Make semi at Sub2 and check: Decap Yoke+SM, Decap Yoke+CM, Drop test, Tension test', 'Make sample final and check function vs Normal'],
    'source_cells': {'title': ['Sheet!B1'], 'date': ['Sheet!N2'], 'purpose': ['Sheet!A4'], 'content': ['Sheet!A6:A9']}
  },
  'test_conditions': [
    {'condition_id': 'cond_1', 'condition_group': 'Test (Yoke dyne-fail lot)', 'line': 'E2-4B', 'process': 'Sub2 + Tension', 'changed_factor': 'Yoke lot that failed dyne test 6/10', 'before_value': 'Normal Yoke lot', 'after_value': 'Yoke dyne-fail lot', 'unit': None, 'machine': None, 'jig': None, 'material_lot': 'Yoke dyne fail (6/10)', 'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!E18']},
    {'condition_id': 'cond_2', 'condition_group': 'Normal', 'line': 'E2-4B', 'process': 'Sub2 + Tension', 'changed_factor': 'Baseline (Normal Yoke lot)', 'before_value': None, 'after_value': 'Normal', 'unit': None, 'machine': None, 'jig': None, 'material_lot': 'Normal', 'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!J18']},
  ],
  'results': [
    {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'NG Process', 'condition_group': 'Test - NG Process', 'date': '2025-06-27', 'line': 'E2-4B', 'input_count': 30, 'ok_count': 30, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'NG Process rate', 'metric_value': 0.0, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG Yoke+SM spread bond not good (<80%)': {'count': 0}}, 'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!F18:I18']},
    {'result_id': 'res_2', 'condition_id': 'cond_2', 'measurement_type': 'NG Process', 'condition_group': 'Normal - NG Process', 'date': '2025-06-27', 'line': 'E2-4B', 'input_count': 30, 'ok_count': 30, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'NG Process rate', 'metric_value': 0.0, 'unit': '%', 'judgement': None, 'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!J18:M18']},
    {'result_id': 'res_3', 'condition_id': 'cond_1', 'measurement_type': 'Decap check bond Yoke+SM', 'condition_group': 'Test - Decap Yoke+SM run 1', 'date': '2025-06-27', 'line': 'E2-4B', 'input_count': 6, 'ok_count': 5, 'ng_count': 1, 'ng_rate_decimal': 0.167, 'ng_rate_percent': 16.7, 'metric_name': 'Decap Yoke+SM NG rate', 'metric_value': 16.7, 'unit': '%', 'judgement': 'FAIL', 'ng_breakdown': {'Bond not good (<80%)': {'count': 1}}, 'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!F19:I19']},
    {'result_id': 'res_4', 'condition_id': 'cond_2', 'measurement_type': 'Decap check bond Yoke+SM', 'condition_group': 'Normal - Decap Yoke+SM run 1', 'date': '2025-06-27', 'line': 'E2-4B', 'input_count': 6, 'ok_count': 6, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Decap Yoke+SM NG rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS', 'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!J19:M19']},
    {'result_id': 'res_5', 'condition_id': 'cond_1', 'measurement_type': 'Decap check bond Yoke+SM', 'condition_group': 'Test - Decap Yoke+SM run 2', 'date': '2025-06-27', 'line': 'E2-4B', 'input_count': 6, 'ok_count': 6, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Decap Yoke+SM NG rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS', 'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!F20:I20']},
    {'result_id': 'res_6', 'condition_id': 'cond_1', 'measurement_type': 'Drop test', 'condition_group': 'Test - Drop test', 'date': '2025-06-27', 'line': 'E2-4B', 'input_count': 10, 'ok_count': 10, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Drop test NG rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS', 'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!F21:I21']},
    {'result_id': 'res_7', 'condition_id': 'cond_2', 'measurement_type': 'Drop test', 'condition_group': 'Normal - Drop test', 'date': '2025-06-27', 'line': 'E2-4B', 'input_count': 10, 'ok_count': 10, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Drop test NG rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS', 'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!J21:M21']},
    {'result_id': 'res_8', 'condition_id': 'cond_1', 'measurement_type': 'Tension Yoke+SM long', 'condition_group': 'Test - Yoke+SM long (spec >=8.0 Kgf)', 'date': '2025-06-28', 'line': 'E2-4B', 'input_count': 7, 'ok_count': 7, 'ng_count': 0, 'ng_rate_decimal': None, 'ng_rate_percent': None, 'metric_name': 'Tension AVG', 'metric_value': 57.54, 'unit': 'Kgf', 'judgement': 'PASS', 'ng_breakdown': {'min': 54.37, 'max': 61.32}, 'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!H26:N26']},
    {'result_id': 'res_9', 'condition_id': 'cond_1', 'measurement_type': 'Tension Yoke+SM short', 'condition_group': 'Test - Yoke+SM short (spec >=6.0 Kgf)', 'date': '2025-06-28', 'line': 'E2-4B', 'input_count': 7, 'ok_count': 7, 'ng_count': 0, 'ng_rate_decimal': None, 'ng_rate_percent': None, 'metric_name': 'Tension AVG', 'metric_value': 23.74, 'unit': 'Kgf', 'judgement': 'PASS', 'ng_breakdown': {'min': 20.11, 'max': 29.44}, 'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!H27:N27']},
    {'result_id': 'res_10', 'condition_id': 'cond_1', 'measurement_type': 'Tension Yoke+CM', 'condition_group': 'Test - Yoke+CM (spec >=70.0 Kgf)', 'date': '2025-06-28', 'line': 'E2-4B', 'input_count': 11, 'ok_count': 11, 'ng_count': 0, 'ng_rate_decimal': None, 'ng_rate_percent': None, 'metric_name': 'Tension AVG', 'metric_value': 195.68, 'unit': 'Kgf', 'judgement': 'PASS', 'ng_breakdown': {'min': 172.0, 'max': 218.0}, 'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!H28:R28']},
    {'result_id': 'res_11', 'condition_id': 'cond_2', 'measurement_type': 'Tension Yoke+SM long', 'condition_group': 'Normal - Yoke+SM long (spec >=8.0 Kgf)', 'date': '2025-06-28', 'line': 'E2-4B', 'input_count': 7, 'ok_count': 7, 'ng_count': 0, 'ng_rate_decimal': None, 'ng_rate_percent': None, 'metric_name': 'Tension AVG', 'metric_value': 54.20, 'unit': 'Kgf', 'judgement': 'PASS', 'ng_breakdown': {'min': 50.76, 'max': 56.32}, 'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!H29:N29']},
    {'result_id': 'res_12', 'condition_id': 'cond_2', 'measurement_type': 'Tension Yoke+SM short', 'condition_group': 'Normal - Yoke+SM short (spec >=6.0 Kgf)', 'date': '2025-06-28', 'line': 'E2-4B', 'input_count': 7, 'ok_count': 7, 'ng_count': 0, 'ng_rate_decimal': None, 'ng_rate_percent': None, 'metric_name': 'Tension AVG', 'metric_value': 31.91, 'unit': 'Kgf', 'judgement': 'PASS', 'ng_breakdown': {'min': 25.34, 'max': 38.39}, 'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!H30:N30']},
    {'result_id': 'res_13', 'condition_id': 'cond_2', 'measurement_type': 'Tension Yoke+CM', 'condition_group': 'Normal - Yoke+CM (spec >=70.0 Kgf)', 'date': '2025-06-28', 'line': 'E2-4B', 'input_count': 11, 'ok_count': 11, 'ng_count': 0, 'ng_rate_decimal': None, 'ng_rate_percent': None, 'metric_name': 'Tension AVG', 'metric_value': 199.86, 'unit': 'Kgf', 'judgement': 'PASS', 'ng_breakdown': {'min': 183.5, 'max': 216.3}, 'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!H31:R31']},
  ],
  'conclusions': [
    {'conclusion_id': 'concl_1', 'topic': 'Tension test - Yoke dyne-fail vs Normal',
     'statement_from_report': '(Tension table)',
     'normalized_interpretation': 'All tension specs are PASS in both Test and Normal: Yoke+SM long Test 57.54 / Normal 54.20 Kgf (both >>8.0); Yoke+SM short Test 23.74 / Normal 31.91 (both >>6.0); Yoke+CM Test 195.68 / Normal 199.86 (both >>70.0). Test Yoke+SM long is even +6.2% higher than Normal; Test Yoke+SM short is -25.6% vs Normal but still 4x spec; Yoke+CM is -2.1% vs Normal. Dyne-fail Yoke still meets bonding strength specification.',
     'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!H26:R31']},
    {'conclusion_id': 'concl_2', 'topic': 'Sub2 decap check bond',
     'statement_from_report': '(Decap Yoke+SM run 1: Test 1/6 NG, Normal 0/6)',
     'normalized_interpretation': "Decap Yoke+SM run 1: Test 16.7% NG (1/6) vs Normal 0.0% (0/6). However a 2nd run of Test scored 0/6 OK. Sample size is too small (n=6 each) to confirm a real difference; note says 'NG Yoke+SM spread bond not good (Not enough 80%)' on the Test sample.",
     'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!F19:M20']},
    {'conclusion_id': 'concl_3', 'topic': 'Drop test',
     'statement_from_report': '(Drop test Test 10/10 OK, Normal 10/10 OK)',
     'normalized_interpretation': 'Drop test Test 0.0% NG vs Normal 0.0% NG; same as baseline.',
     'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!F21:M21']},
  ],
  'troubleshooting_index': {
    'defect_name': 'NG Yoke Dyne Test usage',
    'when_user_asks': ['Yoke dyne test fail material use or not', 'Yoke NG dyne tension impact'],
    'suggested_checks': [
      {'hint_id': 'hint_1', 'check_item': 'Validate dyne-fail Yoke lot tension vs spec on larger sample (>20)', 'reason': 'Yoke+SM short Test AVG 23.74 vs Normal 31.91 Kgf = -25.6% lower; though still 4x spec, the gap is large with only n=7.', 'evidence_strength': 'medium', 'related_process': 'Tension test', 'related_part': 'Yoke+SM', 'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!I26:N31']},
      {'hint_id': 'hint_2', 'check_item': 'Investigate "Yoke+SM spread bond not good (<80%)" finding in NG Process note', 'reason': 'Although NG Process count is 0/30 for both Test and Normal, the note column flags spread bond <80% for Test. This is a workmanship/coverage observation independent of count.', 'evidence_strength': 'medium', 'related_process': "Ass'y Yoke+SM bonding", 'related_part': 'Yoke', 'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!N18']},
      {'hint_id': 'hint_3', 'check_item': 'Increase decap sample size for Yoke+SM (Test had 1/6 NG vs 2nd run 0/6)', 'reason': 'Single decap run can swing 16.7% NG <-> 0%. Need >=20 sample to give a stable estimate.', 'evidence_strength': 'medium', 'related_process': 'Sub2 decap check', 'related_part': 'Yoke+SM', 'source_file': name08, 'sheet_name': 'Sheet', 'source_cells': ['Sheet!F19:M20']},
    ],
    'limitations': ['Function test cells in summary table are empty (#DIV/0!).', 'Decision section is empty.', 'Sample sizes per test are small (n=6-30).']
  },
  'ai_extraction_log': {
    'confidence': 0.65,
    'assumptions': ['Test and Normal columns are paired same-day baseline at line E2-4B.', 'Tension spec values 8.0/6.0/70.0 Kgf are minimum (>= spec).'],
    'warnings': ['Function row in summary table shows #DIV/0! (no input recorded); no function comparison possible.', 'Decap Yoke+SM has only 2 runs with n=6 each, results swing widely.', 'Decision text body is empty.'],
    'decision_rationale': 'Dyne-fail Yoke lot passes all tension specs (every measurement >= 4x spec); drop test 10/10 OK same as Normal; decap 1st run Test 1/6 NG vs Normal 0/6 but 2nd run 0/6. Strength evidence supports release; recommend larger sample for confirmation.'}
}

tr_ko_08 = {
  'document': {'title': 'MSU-L20S15-07 Yoke L20S15-07 Dyne test NG 자재 사용 검토 리포트', 'purpose': 'Dyne test 6/10 fail(~60%) Yoke lot 사용 가능 여부 결정.',
               'content': ['Sub2에서 Decap Yoke+SM, Yoke+CM 확인', 'Drop test', 'Tension test', 'Final 후 function vs Normal']},
  'conclusions': {
    'concl_1': {'topic': 'Tension - Yoke dyne fail vs Normal', 'statement_from_report': '(Tension 표)', 'normalized_interpretation': '모든 spec PASS. Yoke+SM long Test 57.54 / Normal 54.20 Kgf (spec 8.0); Yoke+SM short Test 23.74 / Normal 31.91 Kgf (spec 6.0, -25.6%); Yoke+CM Test 195.68 / Normal 199.86 (spec 70.0). Dyne fail Yoke도 본드 강도 spec 충족.'},
    'concl_2': {'topic': 'Sub2 decap check', 'statement_from_report': '(Decap Yoke+SM run1 Test 1/6 NG, Normal 0/6)', 'normalized_interpretation': 'Run1: Test 16.7% NG vs Normal 0.0%. Run2 Test 0/6 OK. n=6로 표본이 너무 작아 결론 어려움. Test sample에 "spread bond not good (<80%)" 코멘트.'},
    'concl_3': {'topic': 'Drop test', 'statement_from_report': 'Test 10/10 OK, Normal 10/10 OK', 'normalized_interpretation': 'Drop NG rate Test 0.0% = Normal 0.0%.'},
  },
  'hints': {
    'hint_1': {'check_item': 'Dyne-fail Yoke의 tension을 큰 sample(>20)로 검증', 'reason': 'Yoke+SM short Test 23.74 vs Normal 31.91 Kgf (-25.6%). spec의 4배지만 차이 큼.'},
    'hint_2': {'check_item': 'NG Process 노트 "Yoke+SM spread bond not good (<80%)" 점검', 'reason': 'NG count는 0/30이나 노트에 bond coverage 부족 표기. 별도 정성 확인 필요.'},
    'hint_3': {'check_item': 'Decap Yoke+SM 샘플 수 확대', 'reason': '6개 시험이 16.7% ↔ 0%로 결과 흔들림.'},
  },
  'log': {'assumptions': ['Test와 Normal은 E2-4B 같은 날짜 baseline pair.', 'Tension spec 8.0/6.0/70.0 Kgf는 최소값 spec.'],
          'warnings': ['Function 셀이 #DIV/0!로 비어 있어 function 비교 불가.', 'Decap n=6로 적음.', 'Decision 비어 있음.'],
          'decision_rationale': 'Dyne fail Yoke가 모든 tension spec 통과 (최소 4배), drop test 10/10 OK, decap 1차 1/6 NG지만 2차 0/6. 추가 확인 후 release 가능.'}
}
tr_en_08 = {
  'document': {'title': 'MSU-L20S15-07 Report Test Material YOKE L20S15-07 Happen NG Dyne Test', 'purpose': 'Decide whether Yoke lot failed dyne test 6/10 (~60%) can be used.',
               'content': ['Sub2 Decap Yoke+SM and Yoke+CM check', 'Drop test', 'Tension test', 'Final function vs Normal']},
  'conclusions': {
    'concl_1': {'topic': 'Tension - dyne-fail vs Normal', 'statement_from_report': '(Tension table)', 'normalized_interpretation': 'All specs PASS. Yoke+SM long Test 57.54 / Normal 54.20 Kgf (spec 8.0); Yoke+SM short Test 23.74 / Normal 31.91 Kgf (-25.6%); Yoke+CM Test 195.68 / Normal 199.86 (-2.1%). Dyne-fail Yoke still meets bond strength spec by 4x+.'},
    'concl_2': {'topic': 'Sub2 decap check', 'statement_from_report': '(Decap Yoke+SM run1 Test 1/6 NG, Normal 0/6)', 'normalized_interpretation': 'Run1 Test 16.7% NG vs Normal 0.0%; run2 Test 0/6 OK. n=6 is too small to conclude. Test sample carries note "spread bond not good (<80%)".'},
    'concl_3': {'topic': 'Drop test', 'statement_from_report': 'Test 10/10 OK, Normal 10/10 OK', 'normalized_interpretation': 'Drop NG rate Test 0.0% = Normal 0.0%.'},
  },
  'hints': {
    'hint_1': {'check_item': 'Validate dyne-fail Yoke tension on >=20 samples', 'reason': 'Yoke+SM short Test 23.74 vs Normal 31.91 Kgf (-25.6%); still 4x spec but gap is large.'},
    'hint_2': {'check_item': 'Investigate Yoke+SM bond coverage <80% note', 'reason': 'NG Process count 0/30 but Note column flags spread bond <80%.'},
    'hint_3': {'check_item': 'Increase Decap Yoke+SM sample size', 'reason': 'Two n=6 runs swing 16.7% <-> 0%.'},
  },
  'log': {'assumptions': ['Test and Normal are paired baseline at E2-4B on same date.', 'Tension specs are minimum thresholds.'],
          'warnings': ['Function summary cell shows #DIV/0!; no function comparison.', 'Decap sample n=6 is small.', 'Decision section empty.'],
          'decision_rationale': 'Dyne-fail Yoke passes all tension specs (every measurement >= 4x spec), drop test equal Normal, decap inconsistent on small sample. Strength evidence supports release; recommend larger confirmation sample.'}
}
tr_vi_08 = {
  'document': {'title': 'MSU-L20S15-07 Báo cáo test vật liệu Yoke L20S15-07 NG dyne test', 'purpose': 'Quyết định lot Yoke fail dyne test 6/10 (~60%) có dùng được không.',
               'content': ['Sub2 decap Yoke+SM và Yoke+CM', 'Drop test', 'Tension test', 'Function final vs Normal']},
  'conclusions': {
    'concl_1': {'topic': 'Tension - Yoke dyne-fail vs Normal', 'statement_from_report': '(Bảng Tension)', 'normalized_interpretation': 'Tất cả spec PASS. Yoke+SM long Test 57.54 / Normal 54.20 Kgf (spec 8.0); Yoke+SM short Test 23.74 / Normal 31.91 (-25.6%); Yoke+CM Test 195.68 / Normal 199.86 (-2.1%). Yoke fail dyne vẫn vượt spec 4x.'},
    'concl_2': {'topic': 'Sub2 decap check', 'statement_from_report': '(Decap Yoke+SM run1 Test 1/6 NG, Normal 0/6)', 'normalized_interpretation': 'Run1 Test 16.7% vs Normal 0.0%; run2 Test 0/6 OK. n=6 quá nhỏ. Ghi chú Test "spread bond not good (<80%)".'},
    'concl_3': {'topic': 'Drop test', 'statement_from_report': 'Test 10/10 OK, Normal 10/10 OK', 'normalized_interpretation': 'Drop NG rate Test 0.0% = Normal 0.0%.'},
  },
  'hints': {
    'hint_1': {'check_item': 'Xác nhận tension Yoke fail dyne với sample >=20', 'reason': 'Yoke+SM short Test 23.74 vs Normal 31.91 Kgf (-25.6%); dù vượt spec 4x nhưng chênh lớn.'},
    'hint_2': {'check_item': 'Điều tra ghi chú Yoke+SM bond <80%', 'reason': 'NG Process 0/30 nhưng ghi chú nói spread bond <80%.'},
    'hint_3': {'check_item': 'Tăng cỡ mẫu Decap Yoke+SM', 'reason': 'Hai run n=6 dao động 16.7% <-> 0%.'},
  },
  'log': {'assumptions': ['Test và Normal là cặp baseline tại E2-4B cùng ngày.', 'Spec tension là giá trị min.'],
          'warnings': ['Ô function bảng tổng kết #DIV/0!; không so sánh được function.', 'n=6 Decap nhỏ.', 'Decision trống.'],
          'decision_rationale': 'Yoke fail dyne vượt mọi spec tension (>=4x), drop test = Normal, decap n=6 dao động. Bằng chứng strength ủng hộ release; cần sample lớn hơn xác nhận.'}
}
results[name08] = (result08, tr_ko_08, tr_en_08, tr_vi_08)

# =====================================================================
# DS09: 27. MSU-L20S15-07 Report test Find reason NG  LOT DOME 3-17  date 28.4.2025
# =====================================================================
name09 = '27. MSU-L20S15-07 Report test Find reason NG  LOT DOME 3-17  date 28.4.2025'

result09 = {
  'schema_version': '0.1',
  'document': {
    'document_id': '', 'source_file': name09, 'source_sheet': 'Test',
    'title': 'BRS-201507DT Report Test Find Reason NG Lot Dome 3/17',
    'model': 'MSU-L20S15-07 / BRS-201507DT', 'report_date': '2025-04-28',
    'department': 'ME', 'marker': 'Le', 'line': 'C2-2A',
    'report_type': 'normal_comparison',
    'primary_defect': {'canonical_name': 'NG VP/CD Separation', 'aliases_in_document': ['NG separate VP/CD', 'Separation NG']},
    'related_defects': ['NG Hearing Noise', 'NG Hearing Touch'],
    'parts': ['CD', 'VP', 'Dome lot 3/17'], 'processes': ['Laser marking', 'Primer cleaning', 'Plasma', "Ass'y VP/CD", 'Tension test', 'Function'],
    'purpose': 'Find reason for NG Lot Dome 3/17 with high VP/CD separation, by comparing two CD pre-treatments.',
    'content': ['Test 1: Use box 27 (100%) NG separate VP/CD high, CD lot 3/17', 'Compare Laser marking CD vs Primer clean CD before plasma'],
    'source_cells': {'title': ['Test!B1'], 'date': ['Test!T2'], 'purpose': ['Test!A4'], 'content': ['Test!A6:A8']}
  },
  'test_conditions': [
    {'condition_id': 'cond_1', 'condition_group': 'Laser marking CD + Plasma', 'line': 'C2-2A', 'process': 'CD pre-treatment + Plasma + Ass\'y VP/CD + Function', 'changed_factor': 'CD surface pre-treatment (Laser marking)', 'before_value': 'Primer clean CD', 'after_value': 'Laser marking CD', 'unit': None, 'machine': None, 'jig': None, 'material_lot': 'CD lot 17/Mar', 'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name09, 'sheet_name': 'Test', 'source_cells': ['Test!C11']},
    {'condition_id': 'cond_2', 'condition_group': 'Primer clean CD + Plasma', 'line': 'C2-2A', 'process': 'CD pre-treatment + Plasma + Ass\'y VP/CD + Function', 'changed_factor': 'CD surface pre-treatment (Primer clean)', 'before_value': 'Laser marking CD', 'after_value': 'Primer clean CD', 'unit': None, 'machine': None, 'jig': None, 'material_lot': 'CD lot 17/Mar', 'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None, 'source_file': name09, 'sheet_name': 'Test', 'source_cells': ['Test!C12']},
  ],
  'results': [
    {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Separate VP/CD', 'condition_group': 'Laser marking CD + Plasma', 'date': '2025-04-28', 'line': 'C2-2A', 'input_count': 43, 'ok_count': 43, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Separate VP/CD NG rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS', 'ng_breakdown': {}, 'source_file': name09, 'sheet_name': 'Test', 'source_cells': ['Test!E16:I16']},
    {'result_id': 'res_2', 'condition_id': 'cond_2', 'measurement_type': 'Separate VP/CD', 'condition_group': 'Primer clean CD + Plasma', 'date': '2025-04-28', 'line': 'C2-2A', 'input_count': 49, 'ok_count': 36, 'ng_count': 13, 'ng_rate_decimal': 0.265, 'ng_rate_percent': 26.5, 'metric_name': 'Separate VP/CD NG rate', 'metric_value': 26.5, 'unit': '%', 'judgement': 'FAIL', 'ng_breakdown': {'VP/CD Separation': {'count': 13, 'rate': 26.5}}, 'source_file': name09, 'sheet_name': 'Test', 'source_cells': ['Test!E17:I17']},
    {'result_id': 'res_3', 'condition_id': 'cond_1', 'measurement_type': 'Tension VP+CD (decap)', 'condition_group': 'Laser marking CD + Plasma (decap)', 'date': '2025-04-28', 'line': 'C2-2A', 'input_count': 5, 'ok_count': 5, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Tension VP+CD NG rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS', 'ng_breakdown': {}, 'source_file': name09, 'sheet_name': 'Test', 'source_cells': ['Test!E20:I20']},
    {'result_id': 'res_4', 'condition_id': 'cond_2', 'measurement_type': 'Tension VP+CD (decap)', 'condition_group': 'Primer clean CD + Plasma (decap)', 'date': '2025-04-28', 'line': 'C2-2A', 'input_count': 5, 'ok_count': 5, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Tension VP+CD NG rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS', 'ng_breakdown': {}, 'source_file': name09, 'sheet_name': 'Test', 'source_cells': ['Test!E21:I21']},
    {'result_id': 'res_5', 'condition_id': 'cond_1', 'measurement_type': "Tension VP+CD ass'y (spec 1.2 Kgf)", 'condition_group': 'Laser marking CD + Plasma (5 samples)', 'date': '2025-04-28', 'line': 'C2-2A', 'input_count': 5, 'ok_count': 5, 'ng_count': 0, 'ng_rate_decimal': None, 'ng_rate_percent': None, 'metric_name': 'Tension VP+CD AVG', 'metric_value': 2.007, 'unit': 'Kgf', 'judgement': 'PASS', 'ng_breakdown': {'min': 1.816, 'max': 2.345}, 'source_file': name09, 'sheet_name': 'Test', 'source_cells': ['Test!H25:O25']},
    {'result_id': 'res_6', 'condition_id': 'cond_2', 'measurement_type': "Tension VP+CD ass'y (spec 1.2 Kgf)", 'condition_group': 'Primer clean CD + Plasma (5 samples)', 'date': '2025-04-28', 'line': 'C2-2A', 'input_count': 5, 'ok_count': 5, 'ng_count': 0, 'ng_rate_decimal': None, 'ng_rate_percent': None, 'metric_name': 'Tension VP+CD AVG', 'metric_value': 1.572, 'unit': 'Kgf', 'judgement': 'PASS', 'ng_breakdown': {'min': 1.428, 'max': 1.864}, 'source_file': name09, 'sheet_name': 'Test', 'source_cells': ['Test!H26:O26']},
    {'result_id': 'res_7', 'condition_id': 'cond_1', 'measurement_type': 'Function', 'condition_group': 'Laser marking CD + Plasma 04-28', 'date': '2025-04-28', 'line': 'C2-2A', 'input_count': 43, 'ok_count': 42, 'ng_count': 1, 'ng_rate_decimal': 0.023, 'ng_rate_percent': 2.3, 'metric_name': 'Function NG rate', 'metric_value': 2.3, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG Hearing Touch': {'count': 1, 'rate': 2.3}, 'NG Hearing Noise': {'count': 0, 'rate': 0.0}, 'NG Sigma SPL': {'count': 0}, 'NG Sigma THD': {'count': 0}}, 'source_file': name09, 'sheet_name': 'Test', 'source_cells': ['Test!E30:N30']},
    {'result_id': 'res_8', 'condition_id': 'cond_2', 'measurement_type': 'Function', 'condition_group': 'Primer clean CD + Plasma 04-28', 'date': '2025-04-28', 'line': 'C2-2A', 'input_count': 36, 'ok_count': 34, 'ng_count': 2, 'ng_rate_decimal': 0.056, 'ng_rate_percent': 5.6, 'metric_name': 'Function NG rate', 'metric_value': 5.6, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG Hearing Noise': {'count': 1, 'rate': 2.8}, 'NG Hearing Touch': {'count': 1, 'rate': 2.8}, 'note': '2/2 NG Separate VP/CD'}, 'source_file': name09, 'sheet_name': 'Test', 'source_cells': ['Test!E32:N32']},
    {'result_id': 'res_9', 'condition_id': 'cond_1', 'measurement_type': 'Function', 'condition_group': 'Laser marking CD + Plasma 05-08', 'date': '2025-05-08', 'line': 'C2-2A', 'input_count': 200, 'ok_count': 198, 'ng_count': 2, 'ng_rate_decimal': 0.01, 'ng_rate_percent': 1.0, 'metric_name': 'Function NG rate', 'metric_value': 1.0, 'unit': '%', 'judgement': None, 'ng_breakdown': {'NG Hearing Noise': {'count': 2, 'rate': 1.0}, 'NG Sigma SPL': {'count': 1, 'rate': 0.5}, 'NG Sigma THD': {'count': 0}, 'NG Hearing Touch': {'count': 0}}, 'source_file': name09, 'sheet_name': 'Test', 'source_cells': ['Test!E34:N34']},
  ],
  'conclusions': [
    {'conclusion_id': 'concl_1', 'topic': 'CD pre-treatment effect on VP/CD separation',
     'statement_from_report': '(Result table 1)',
     'normalized_interpretation': 'Laser marking CD + Plasma: 0/43 = 0.0% VP/CD separation. Primer clean CD + Plasma: 13/49 = 26.5% VP/CD separation. Same Dome lot 3/17 + same plasma. Laser marking treatment essentially eliminates the separation; primer-clean is the failure mode.',
     'source_file': name09, 'sheet_name': 'Test', 'source_cells': ['Test!E16:I17']},
    {'conclusion_id': 'concl_2', 'topic': "Tension VP+CD ass'y - both PASS, level differs",
     'statement_from_report': "(Tension VP+CD ass'y table spec 1.2 Kgf)",
     'normalized_interpretation': "Laser marking CD AVG 2.007 Kgf (1.816-2.345), Primer clean CD AVG 1.572 Kgf (1.428-1.864). Both >= spec 1.2 Kgf, both PASS, but laser marking is +27.7% stronger ((2.007/1.572-1)*100); primer-clean approaches the lower spec margin.",
     'source_file': name09, 'sheet_name': 'Test', 'source_cells': ['Test!H25:O26']},
    {'conclusion_id': 'concl_3', 'topic': 'Function NG rate comparison',
     'statement_from_report': '(Function table 04-28 and 05-08)',
     'normalized_interpretation': '04-28 same-event: Laser marking 2.3% (n=43) vs Primer clean 5.6% (n=36) = (5.6/2.3-1)*100 = +143.5% worse than laser marking. 05-08 (laser marking only, no baseline) 1.0% (n=200). Primer-clean note states 2/2 function NG are also Separate VP/CD.',
     'source_file': name09, 'sheet_name': 'Test', 'source_cells': ['Test!E30:N34']},
  ],
  'troubleshooting_index': {
    'defect_name': 'NG VP/CD Separation (Dome lot 3/17)',
    'when_user_asks': ['CD pre-treatment for Dome lot 3/17', 'NG separate VP/CD root cause', 'Primer clean vs Laser marking on CD'],
    'suggested_checks': [
      {'hint_id': 'hint_1', 'check_item': 'Adopt laser marking CD + plasma for problematic Dome lot 3/17', 'reason': 'Same lot: Laser marking 0/43 = 0.0% separation vs Primer clean 13/49 = 26.5% separation; laser marking eliminates the failure mode.', 'evidence_strength': 'high', 'related_process': 'CD pre-treatment', 'related_part': 'CD', 'source_file': name09, 'sheet_name': 'Test', 'source_cells': ['Test!E16:I17']},
      {'hint_id': 'hint_2', 'check_item': "Investigate why primer clean lowers VP+CD ass'y tension (1.572 vs 2.007 Kgf, -21.7%)", 'reason': 'Both PASS spec 1.2 Kgf but primer-clean has 21.7% lower AVG tension, consistent with the higher separation rate.', 'evidence_strength': 'high', 'related_process': 'CD pre-treatment + plasma', 'related_part': 'CD surface', 'source_file': name09, 'sheet_name': 'Test', 'source_cells': ['Test!H25:O26']},
      {'hint_id': 'hint_3', 'check_item': 'Track Function NG mode for primer-clean batch (Touch/Noise + Separate VP/CD)', 'reason': 'Function NG 04-28 primer-clean 5.6% (2/36 - both separate VP/CD) vs laser marking 2.3% (1/43); cause is again separation rather than acoustic.', 'evidence_strength': 'medium', 'related_process': 'Function test', 'related_part': 'VP/CD', 'source_file': name09, 'sheet_name': 'Test', 'source_cells': ['Test!E30:N32']},
    ],
    'limitations': ['Sample sizes for VP/CD separation are small (43, 49); tension uses n=5.', 'Decision section is empty.', '05-08 Function row has no paired baseline.']
  },
  'ai_extraction_log': {
    'confidence': 0.8,
    'assumptions': ['Both conditions use the same Dome lot 3/17 to isolate CD pre-treatment.', 'Tension spec 1.2 Kgf is minimum.'],
    'warnings': ['n is small (43-49) at the separation table.', '05-08 function row stands alone without baseline (treated as supplementary).'],
    'decision_rationale': "Same-event comparison: Laser marking CD eliminates VP/CD separation (0.0% vs 26.5%) and gives +27.7% higher ass'y tension; function NG also drops from 5.6% to 2.3% on the same day. Root cause of NG Lot Dome 3/17 separation is the primer-clean CD pre-treatment step; switching to laser marking is supported."}
}

tr_ko_09 = {
  'document': {'title': 'BRS-201507DT NG Lot Dome 3/17 원인 분석 리포트', 'purpose': 'NG separate VP/CD가 높은 Dome lot 3/17의 원인 분석 (CD 전처리 방식 비교).',
               'content': ['Test 1: 박스 27 (100%) NG separate VP/CD high, CD lot 3/17', 'Laser marking CD vs Primer clean CD + Plasma 비교']},
  'conclusions': {
    'concl_1': {'topic': 'CD 전처리에 따른 VP/CD separation', 'statement_from_report': '(결과 표 1)', 'normalized_interpretation': 'Laser marking CD+Plasma: 0/43 = 0.0% separation. Primer clean CD+Plasma: 13/49 = 26.5%. 동일 Dome lot 3/17, 동일 plasma. Laser marking이 separation 거의 제거; primer clean이 fail mode.'},
    'concl_2': {'topic': "VP+CD Ass'y tension - 둘 다 PASS, 강도 차이", 'statement_from_report': "(Tension VP+CD ass'y 표 spec 1.2 Kgf)", 'normalized_interpretation': 'Laser marking 2.007 Kgf (1.816~2.345), Primer clean 1.572 Kgf (1.428~1.864). 둘 다 spec 통과. Laser marking이 +27.7% 강함; Primer clean은 spec MIN에 가까움.'},
    'concl_3': {'topic': 'Function NG 비교', 'statement_from_report': '(04-28, 05-08 Function 표)', 'normalized_interpretation': '04-28 same-event: Laser marking 2.3% vs Primer clean 5.6% = +143.5% 악화. 05-08 (laser only, baseline 없음) 1.0%. Primer clean 2건의 function NG는 모두 separate VP/CD.'},
  },
  'hints': {
    'hint_1': {'check_item': '문제 Dome lot 3/17에는 laser marking CD + plasma 채택', 'reason': '동일 lot에서 Laser marking 0/43 (0.0%) vs Primer clean 13/49 (26.5%); 결함 모드 제거됨.'},
    'hint_2': {'check_item': "Primer clean이 VP+CD ass'y tension을 낮추는 원인 분석", 'reason': 'Laser marking 2.007 vs Primer clean 1.572 Kgf (-21.7%). spec MIN에 가까워짐.'},
    'hint_3': {'check_item': 'Primer clean lot의 Function NG mode 모니터링', 'reason': 'Primer clean 04-28 NG 2/36 모두 Separate VP/CD; 음향 원인이 아님.'},
  },
  'log': {'assumptions': ['두 조건 모두 Dome lot 3/17 사용으로 CD 전처리만 변화.', 'Tension spec 1.2 Kgf는 최소값.'],
          'warnings': ['Separation 표 n=43, 49로 작음.', '05-08 Function은 baseline 없음, 보조 자료.'],
          'decision_rationale': "동일 이벤트 비교: Laser marking CD가 VP/CD separation 0.0% vs 26.5%로 제거, ass'y tension +27.7% 증가, 같은 날 function NG도 5.6% → 2.3%. 원인은 primer-clean 전처리; laser marking 채택 권장."}
}
tr_en_09 = {
  'document': {'title': 'BRS-201507DT Report Test Find Reason NG Lot Dome 3/17', 'purpose': 'Find reason for NG Lot Dome 3/17 with high VP/CD separation by comparing two CD pre-treatments.',
               'content': ['Test 1: Box 27 (100%) NG separate VP/CD high, CD lot 3/17', 'Compare Laser marking CD vs Primer clean CD before plasma']},
  'conclusions': {
    'concl_1': {'topic': 'CD pre-treatment vs VP/CD separation', 'statement_from_report': '(Result table 1)', 'normalized_interpretation': 'Laser marking CD+Plasma: 0/43 = 0.0%. Primer clean CD+Plasma: 13/49 = 26.5%. Same Dome lot 3/17, same plasma. Laser marking essentially eliminates separation; primer-clean is the failure mode.'},
    'concl_2': {'topic': "VP+CD ass'y tension - both PASS, level differs", 'statement_from_report': '(Tension table spec 1.2 Kgf)', 'normalized_interpretation': 'Laser marking AVG 2.007 Kgf, Primer clean AVG 1.572 Kgf; both PASS. Laser marking is +27.7% stronger; primer clean approaches the spec MIN margin.'},
    'concl_3': {'topic': 'Function NG comparison', 'statement_from_report': '(04-28 and 05-08 Function tables)', 'normalized_interpretation': '04-28 same-event: Laser marking 2.3% vs Primer clean 5.6% = +143.5% worse. 05-08 (laser marking only, no baseline) 1.0%. Primer clean 2/2 function NG are also Separate VP/CD.'},
  },
  'hints': {
    'hint_1': {'check_item': 'Adopt laser marking CD + plasma for problematic Dome lot 3/17', 'reason': 'Same lot: Laser marking 0/43 (0.0%) vs Primer clean 13/49 (26.5%); failure mode eliminated.'},
    'hint_2': {'check_item': "Investigate why primer clean lowers VP+CD ass'y tension", 'reason': 'Laser marking 2.007 vs Primer clean 1.572 Kgf (-21.7%); primer clean approaches spec MIN.'},
    'hint_3': {'check_item': 'Track Function NG mode of primer-clean batches', 'reason': 'Primer clean 04-28 NG 2/36 are all Separate VP/CD; the function NG is again a separation, not acoustic.'},
  },
  'log': {'assumptions': ['Both conditions share Dome lot 3/17, isolating CD pre-treatment effect.', 'Tension spec 1.2 Kgf is minimum.'],
          'warnings': ['Separation table n is small (43, 49).', '05-08 Function row has no baseline, treated as supplementary.'],
          'decision_rationale': "Same-event comparison shows laser marking CD eliminates VP/CD separation (0.0% vs 26.5%) and produces +27.7% higher ass'y tension; same-day function NG also drops 5.6% -> 2.3%. Primer-clean pre-treatment is the root cause; laser marking is supported."}
}
tr_vi_09 = {
  'document': {'title': 'BRS-201507DT Báo cáo tìm nguyên nhân NG Lot Dome 3/17', 'purpose': 'Tìm nguyên nhân NG Lot Dome 3/17 có VP/CD separate cao bằng cách so sánh hai phương pháp xử lý CD.',
               'content': ['Test 1: box 27 (100%) NG separate VP/CD cao, CD lot 3/17', 'So sánh Laser marking CD vs Primer clean CD trước Plasma']},
  'conclusions': {
    'concl_1': {'topic': 'Xử lý CD vs VP/CD separation', 'statement_from_report': '(Bảng kết quả 1)', 'normalized_interpretation': 'Laser marking CD+Plasma: 0/43 = 0.0%. Primer clean CD+Plasma: 13/49 = 26.5%. Cùng Dome lot 3/17 và plasma. Laser marking gần như loại bỏ separation; primer-clean là chế độ lỗi.'},
    'concl_2': {'topic': "Tension VP+CD ass'y - cả hai PASS, mức khác nhau", 'statement_from_report': '(Bảng Tension spec 1.2 Kgf)', 'normalized_interpretation': 'Laser marking AVG 2.007 Kgf, Primer clean AVG 1.572 Kgf; đều PASS. Laser marking mạnh hơn +27.7%; primer clean tiến sát biên dưới spec.'},
    'concl_3': {'topic': 'So sánh NG Function', 'statement_from_report': '(Function 04-28 và 05-08)', 'normalized_interpretation': '04-28 cùng sự kiện: Laser marking 2.3% vs Primer clean 5.6% = +143.5% xấu hơn. 05-08 (laser only, không baseline) 1.0%. 2/2 NG function của primer clean cũng là Separate VP/CD.'},
  },
  'hints': {
    'hint_1': {'check_item': 'Áp dụng laser marking CD + plasma cho Dome lot 3/17 bị lỗi', 'reason': 'Cùng lot: Laser marking 0/43 (0.0%) vs Primer clean 13/49 (26.5%); loại bỏ chế độ lỗi.'},
    'hint_2': {'check_item': "Tìm hiểu tại sao primer clean làm giảm tension VP+CD ass'y", 'reason': 'Laser marking 2.007 vs Primer clean 1.572 Kgf (-21.7%); tiến sát min spec.'},
    'hint_3': {'check_item': 'Theo dõi NG Function của các batch primer clean', 'reason': 'Primer clean 04-28 NG 2/36 đều là Separate VP/CD; lỗi function cũng do separation.'},
  },
  'log': {'assumptions': ['Cả hai điều kiện cùng Dome lot 3/17, chỉ khác cách xử lý CD.', 'Spec tension 1.2 Kgf là tối thiểu.'],
          'warnings': ['Bảng separation cỡ mẫu nhỏ (43, 49).', 'Function 05-08 không baseline, là dữ liệu phụ.'],
          'decision_rationale': "So sánh cùng sự kiện: laser marking CD loại bỏ VP/CD separation (0.0% vs 26.5%), ass'y tension cao hơn +27.7%; function NG cùng ngày giảm 5.6% → 2.3%. Nguyên nhân: bước primer clean; ủng hộ chuyển sang laser marking."}
}
results[name09] = (result09, tr_ko_09, tr_en_09, tr_vi_09)

# Commit batch 2
processed = 0
failed = 0
for n, (res, ko, en, vi) in results.items():
    if h.commit_dataset(n, res, ko, en, vi):
        processed += 1
        print(f'OK: {n}')
    else:
        failed += 1
        print(f'FAIL: {n}')

print(f'partial batch 2: processed={processed} failed={failed}')
