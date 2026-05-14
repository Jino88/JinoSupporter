import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _batch_build as B

D = []

# dataset 8
D.append({
    'name': '19. BRS-161014  Report test VP mold 7, mold 9  add 0.05T in middle position 21.2.2024',
    'report_type': 'normal_comparison', 'model': 'BRS-161016',
    'report_date': '2024-02-21', 'department': 'ME', 'marker': 'Thuy', 'default_sheet': 'NG rate',
    'primary_defect': 'VP bending / NG Function', 'related_defects': ['VP+CD vision NG'],
    'parts': ['VP', 'CD'], 'processes': ['VP forming', 'Sub1 vision', 'Main2 function'],
    'title_en': 'BRS-161016 — Test VP 161014-ED mold #7 / #9 with +0.05T middle (2024-02-21)',
    'title_ko': 'BRS-161016 — VP 161014-ED mold #7/#9 +0.05T 검증 (2024-02-21)',
    'title_vi': 'BRS-161016 — Thử VP 161014-ED khuôn #7/#9 +0.05T giữa (21-02-2024)',
    'purpose_en': 'Verify whether VP mold #7 / #9 with +0.05T middle can be used: VP bending, Sub1 vision VP+CD, Main2 Function — vs Normal VP mold #6.',
    'purpose_ko': 'VP mold #7/#9 (중앙 +0.05T) 사용 가능 여부 검증: VP bending, Sub1 vision VP+CD, Main2 Function vs Normal VP mold #6.',
    'purpose_vi': 'Kiểm tra VP khuôn #7/#9 +0.05T có dùng được không: VP bending, vision VP+CD Sub1, Main2 Function so với Normal khuôn #6.',
    'content_en': ['VP bending E2 line vs Normal #6.','Sub1 vision VP+CD vs Normal #6.','Function (truncated in source but trend captured by writer).'],
    'content_ko': ['E2 라인 VP bending vs Normal #6.','Sub1 vision VP+CD vs Normal #6.','Function (원본 일부 절단, 작성자 트렌드 코멘트 반영).'],
    'content_vi': ['VP bending line E2 vs Normal #6.','Vision VP+CD Sub1 vs Normal #6.','Function (bị cắt một phần, theo nhận xét người viết).'],
    'test_conditions': [{'condition_group': 'vp_mold_05t', 'process': 'VP mold', 'changed_factor': 'VP mold # with +0.05T middle', 'before_value': 'Normal mold #6', 'after_value': '#7 / #9 (+0.05T)', 'unit': 'mm', 'source_cells': ['B6']}],
    'results': [
        {'condition_id': 'c1', 'measurement_type': 'VP Bending', 'date': '2024-02-21', 'line': 'E2', 'input_count': 500, 'ok_count': 488, 'ng_count': 12, 'ng_rate_decimal': 0.024, 'ng_rate_percent': 2.4, 'metric_name': 'Test VP mold #7', 'judgement': 'WORSE', 'source_cells': ['C21:H21']},
        {'condition_id': 'c1', 'measurement_type': 'VP Bending', 'date': '2024-02-21', 'line': 'E2', 'input_count': 500, 'ok_count': 494, 'ng_count': 6, 'ng_rate_decimal': 0.012, 'ng_rate_percent': 1.2, 'metric_name': 'Test VP mold #9', 'judgement': 'WORSE', 'source_cells': ['C21:H22']},
        {'condition_id': 'c1', 'measurement_type': 'VP Bending', 'date': '2024-02-21', 'line': 'E2', 'input_count': 500, 'ok_count': 498, 'ng_count': 2, 'ng_rate_decimal': 0.004, 'ng_rate_percent': 0.4, 'metric_name': 'Normal VP mold #6', 'judgement': 'BASELINE', 'source_cells': ['C21:H23']},
        {'condition_id': 'c1', 'measurement_type': 'Vision VP+CD', 'date': '2024-02-21', 'line': 'E2', 'input_count': 500, 'ok_count': 494, 'ng_count': 6, 'ng_rate_decimal': 0.012, 'ng_rate_percent': 1.2, 'metric_name': 'Total Test VP mold #7', 'judgement': 'WORSE', 'ng_breakdown': {'VP separate': {'count': 5}, 'Dome damage': {'count': 1}}, 'source_cells': ['C31:J31']},
        {'condition_id': 'c1', 'measurement_type': 'Vision VP+CD', 'date': '2024-02-21', 'line': 'E2', 'input_count': 500, 'ok_count': 493, 'ng_count': 7, 'ng_rate_decimal': 0.014, 'ng_rate_percent': 1.4, 'metric_name': 'Total Test VP mold #9', 'judgement': 'WORSE', 'ng_breakdown': {'VP separate': {'count': 4}, 'Dome damage': {'count': 3}}, 'source_cells': ['C34:J34']},
        {'condition_id': 'c1', 'measurement_type': 'Vision VP+CD', 'date': '2024-02-21', 'line': 'E2', 'input_count': 500, 'ok_count': 495, 'ng_count': 5, 'ng_rate_decimal': 0.01, 'ng_rate_percent': 1.0, 'metric_name': 'Total Normal VP mold #6', 'judgement': 'BASELINE', 'ng_breakdown': {'VP separate': {'count': 2}, 'Dome damage': {'count': 3}}, 'source_cells': ['C37:J37']},
    ],
    'conclusions': [{
        'topic_en': 'VP mold +0.05T verification — earlier date',
        'topic_ko': 'VP mold +0.05T 검증 — 이른 일자',
        'topic_vi': 'Xác minh VP +0.05T — đợt sớm',
        'statement_en': 'Bending: #7 2.4%, #9 1.2%, Normal 0.4%. Vision NG: #7 1.2%, #9 1.4%, Normal 1.0%. Writer note: #7/#9 still bend.',
        'statement_ko': 'Bending: #7 2.4%, #9 1.2%, Normal 0.4%. Vision NG: #7 1.2%, #9 1.4%, Normal 1.0%. 작성자: #7/#9 여전히 bending 발생.',
        'statement_vi': 'Bending: #7 2.4%, #9 1.2%, Normal 0.4%. Vision NG: #7 1.2%, #9 1.4%, Normal 1.0%. Người viết: #7/#9 vẫn bending.',
        'interp_en': 'VP mold +0.05T middle still increases bending (Normal 0.4% → #7 2.4% / #9 1.2% — 500%/200% worse) and slightly increases Vision NG (+20% on #7, +40% on #9). Not yet usable.',
        'interp_ko': 'VP mold +0.05T 중앙 추가는 여전히 bending 을 악화시킴(Normal 0.4% → #7 2.4%/#9 1.2%, +500%/+200%) 및 Vision NG 도 +20%/+40% 악화. 사용 불가.',
        'interp_vi': 'Thêm 0.05T giữa vẫn làm tăng bending (Normal 0.4% → #7 2.4%/#9 1.2%, +500%/+200%) và Vision NG tăng nhẹ (+20%/+40%). Chưa thể dùng.',
        'source_cells': ['H21:H23','J31:J37']
    }],
    'hints': [{'check_en': 'Do not use VP mold #7/#9 +0.05T (earlier verification)', 'reason_en': 'Bending 500%/200% worse and vision marginally worse vs Normal #6.', 'check_ko': 'VP mold #7/#9 +0.05T 사용 불가 (조기 검증).', 'reason_ko': 'Normal #6 대비 bending +500%/+200%, vision 도 미세 악화.', 'check_vi': 'Không dùng VP khuôn #7/#9 +0.05T (đợt kiểm sớm).', 'reason_vi': 'Bending tệ hơn 500%/200% và vision hơi tệ hơn so với Normal #6.', 'evidence_strength': 'strong', 'related_process': 'VP mold', 'related_part': 'VP', 'source_cells': ['H21:H23']}],
    'log': {'rationale_en': 'Earlier verification of the same VP +0.05T experiment as dataset "04.03.2024". Each result table has same-event Test #7/#9 vs Normal #6 — normal_comparison.','rationale_ko': '"04.03.2024" 동일 실험의 조기 검증. 각 표에 Test #7/#9 와 Normal #6 동일 이벤트 — normal_comparison.','rationale_vi': 'Kiểm tra sớm của thí nghiệm VP +0.05T tương tự bộ dữ liệu "04.03.2024". Mỗi bảng có Test #7/#9 và Normal #6 cùng sự kiện — normal_comparison.','warnings_en': ['Function table truncated in the extracted text — function rows not stored beyond row 40.'],'warnings_ko': ['추출 텍스트에서 Function 표가 r40 이후 절단 — function 행은 부분 저장.'],'warnings_vi': ['Bảng Function bị cắt sau r40 — chỉ lưu một phần.']},
    'when_user_asks': ['VP mold +0.05T', 'VP bending', 'BRS-161016 VP mold'],
    'confidence': 0.7,
})

# dataset 9
D.append({
    'name': "19. BRS-161014 Report test ass'y Frame + Yoke by hand",
    'report_type': 'normal_comparison', 'model': 'BRS-161014',
    'report_date': '2023-09-08', 'department': 'ME', 'marker': '', 'default_sheet': 'Report',
    'primary_defect': 'NG Function (Hearing Noise + Touch)', 'related_defects': ['VP damage','FS damage','Particle FS'],
    'parts': ['Frame', 'Yoke', 'FS', 'VP'], 'processes': ['Sub4 Frame+Yoke assembly (hand vs machine)'],
    'title_en': 'BRS-161014 — Test Frame+Yoke ass’y by hand (2023-09-08)',
    'title_ko': 'BRS-161014 — Frame+Yoke 수작업 ass’y 검증 (2023-09-08)',
    'title_vi': 'BRS-161014 — Thử ass’y Frame+Yoke bằng tay (08-09-2023)',
    'purpose_en': 'Try hand-assembling Frame+Yoke (instead of machine) to reduce Function NG. Track Sub4 (air leak, long-VP vision, FS vision, height) and Function vs Normal.',
    'purpose_ko': 'Function NG 감소를 위해 Frame+Yoke 를 수작업으로 ass’y 시도 (기계 대비). Sub4 (air leak, long-VP vision, FS vision, height) 및 Function 을 Normal 과 비교.',
    'purpose_vi': 'Lắp Frame+Yoke bằng tay (thay máy) để giảm NG function. Theo dõi Sub4 (air leak, vision long-VP, vision FS, chiều cao) và Function so với Normal.',
    'content_en': ['Sub4 tracking 98pcs hand-assembly vs 200pcs Normal.','Function 87pcs hand-assembly vs 188pcs Normal.'],
    'content_ko': ['Sub4 tracking: 수작업 98pcs vs Normal 200pcs.','Function: 수작업 87pcs vs Normal 188pcs.'],
    'content_vi': ['Tracking Sub4: tay 98pcs vs Normal 200pcs.','Function: tay 87pcs vs Normal 188pcs.'],
    'test_conditions': [{'condition_group': 'frame_yoke_hand', 'process': 'Sub4 Frame+Yoke ass’y', 'changed_factor': 'Ass’y method', 'before_value': 'Machine (Normal)', 'after_value': 'By hand', 'source_cells': ['B8']}],
    'results': [
        {'condition_id': 'c1', 'measurement_type': 'Sub4 Tracking', 'date': '2023-09-08', 'input_count': 98, 'ok_count': None, 'ng_count': 11, 'ng_rate_decimal': 0.1122, 'ng_rate_percent': 11.22, 'metric_name': 'Frame+Yoke by hand', 'judgement': 'WORSE', 'ng_breakdown': {'Air leak NG': {'count': 0}, 'VP damage': {'count': 2}, 'FS damage': {'count': 7}, 'Particle FS': {'count': 1}, 'Height NG': {'count': 1}}, 'source_cells': ['S13:T14']},
        {'condition_id': 'c1', 'measurement_type': 'Sub4 Tracking', 'date': '2023-09-08', 'input_count': 200, 'ok_count': None, 'ng_count': 9, 'ng_rate_decimal': 0.045, 'ng_rate_percent': 4.5, 'metric_name': 'Normal', 'judgement': 'BASELINE', 'ng_breakdown': {'VP damage': {'count': 0}, 'FS damage': {'count': 7}, 'Particle FS': {'count': 2}, 'Height NG': {'count': 0}}, 'source_cells': ['S15:T16']},
        {'condition_id': 'c1', 'measurement_type': 'Function', 'date': '2023-09-08', 'input_count': 87, 'ok_count': 46, 'ng_count': 41, 'ng_rate_decimal': 0.4713, 'ng_rate_percent': 47.13, 'metric_name': 'Frame+Yoke by hand', 'judgement': 'WORSE', 'ng_breakdown': {'Noise': {'count': 22, 'rate': 0.5366}, 'Touch': {'count': 16, 'rate': 0.3902}, 'THD': {'count': 2}, 'SPL+THD': {'count': 1}}, 'source_cells': ['O20']},
        {'condition_id': 'c1', 'measurement_type': 'Function', 'date': '2023-09-08', 'input_count': 188, 'ok_count': 127, 'ng_count': 61, 'ng_rate_decimal': 0.3245, 'ng_rate_percent': 32.45, 'metric_name': 'Normal', 'judgement': 'BASELINE', 'ng_breakdown': {'Noise': {'count': 48, 'rate': 0.7869}, 'Touch': {'count': 12, 'rate': 0.1967}, 'THD': {'count': 1}}, 'source_cells': ['O22']},
    ],
    'conclusions': [{
        'topic_en': 'Frame+Yoke hand assembly effect on Function NG',
        'topic_ko': 'Frame+Yoke 수작업이 Function NG 에 미친 효과',
        'topic_vi': 'Ảnh hưởng của ass’y Frame+Yoke bằng tay lên NG function',
        'statement_en': 'By hand 47.13% vs Normal 32.45% function NG. Decision: NG function not reduced, need other reason.',
        'statement_ko': '수작업 47.13% vs Normal 32.45% function NG. 결정: Function NG 감소 없음, 다른 원인 추적 필요.',
        'statement_vi': 'Tay 47.13% vs Normal 32.45% NG function. Quyết định: không giảm NG function, cần tìm nguyên nhân khác.',
        'interp_en': 'Hand assembly is 45.2% worse than same-event Normal (47.13% vs 32.45%) on Function and 149.3% worse on Sub4 tracking (11.22% vs 4.5%). Hand assembly is not the root cause of Function NG and should not be adopted.',
        'interp_ko': '수작업은 동일 이벤트 Normal 대비 Function 에서 45.2% 악화 (47.13% vs 32.45%), Sub4 tracking 에서 149.3% 악화 (11.22% vs 4.5%). 수작업은 Function NG 의 원인이 아니며 채택 불가.',
        'interp_vi': 'Tay tệ hơn Normal cùng sự kiện 45.2% trên Function (47.13% vs 32.45%) và 149.3% trên Sub4 (11.22% vs 4.5%). Tay không phải nguyên nhân và không nên áp dụng.',
        'source_cells': ['O20','O22','T13:T15']
    }],
    'hints': [{'check_en': 'Do not adopt hand Frame+Yoke ass’y; look at process/material upstream', 'reason_en': 'Hand-ass’y is 45% worse on Function and 149% worse on Sub4 vs same-event Normal — process is not the bottleneck.', 'check_ko': '수작업 Frame+Yoke ass’y 채택 불가, 상류 공정/자재 확인.', 'reason_ko': '수작업이 Function 45%/Sub4 149% 악화 — Sub4 기계가 병목 아님.', 'check_vi': 'Không áp dụng ass’y tay; xem xét quy trình/vật liệu thượng nguồn.', 'reason_vi': 'Ass’y tay tệ hơn 45% Function và 149% Sub4 — quy trình máy không phải nút thắt.', 'evidence_strength': 'strong', 'related_process': 'Sub4 ass’y', 'related_part': 'Frame/Yoke', 'source_cells': ['O20:O22','T13:T15']}],
    'log': {'rationale_en': 'Same-day hand-vs-machine pair on both Sub4 and Function — normal_comparison.','rationale_ko': 'Sub4/Function 모두 동일일자 수작업 vs 기계 쌍 — normal_comparison.','rationale_vi': 'Cùng ngày tay-vs-máy cho cả Sub4 và Function — normal_comparison.','assumptions_en': ['Normal == machine assembly with normal flow.'],'assumptions_ko': ['Normal 은 기계 ass’y / 정규 flow 로 가정.'],'assumptions_vi': ['Normal được hiểu là ass’y máy / dòng bình thường.']},
    'when_user_asks': ['Frame Yoke hand assembly','BRS-161014 NG function'],
    'confidence': 0.78,
})

# dataset 10: 201506 NG function high (48K, multiple sheets) — summarize key Test coil & UV LED comparisons
D.append({
    'name': '19. BRS-201506 Report checking and test problem NG function high date 24.2.2024 -',
    'report_type': 'normal_comparison', 'model': 'BRS-201506',
    'report_date': '2024-02-24', 'department': 'ME', 'marker': 'Thao',
    'default_sheet': "Test coil chạm Ring Frame, UV LED COIL SP, …",
    'primary_defect': 'NG Function high (Hearing Noise + Touch)',
    'related_defects': ['Coil + Frame Gap', 'UV LED separation'],
    'parts': ['Coil','Frame','Ring','SP','UV LED'], 'processes': ['Sub2 Coil+Frame','UV LED Press','Sub3 Ass’y'],
    'title_en': 'BRS-201506 — Investigate NG function high (Coil-Ring Frame contact, UV LED separation, etc.) (2024-02-24)',
    'title_ko': 'BRS-201506 — NG function 高 원인 검토 (Coil-Ring Frame 접촉, UV LED 분리 등) (2024-02-24)',
    'title_vi': 'BRS-201506 — Phân tích NG function cao (Coil-Ring Frame chạm, UV LED tách, …) (24-02-2024)',
    'purpose_en': 'Investigate the root cause of high NG function on BRS-201506. Includes separation of Frame+Coil Gap OK vs Gap a little, separation of UV LED press on SP+Coil etc., each with same-event function comparison.',
    'purpose_ko': 'BRS-201506 의 NG function 高 원인 검토. Frame+Coil Gap OK vs Gap a little, SP+Coil UV LED press 분리 등 각 동일 이벤트 function 비교 포함.',
    'purpose_vi': 'Tìm nguyên nhân NG function cao trên BRS-201506. Bao gồm tách Frame+Coil Gap OK vs Gap a little, UV LED press SP+Coil, … kèm so sánh function cùng sự kiện.',
    'content_en': ['Sheet "Test coil chạm Ring Frame": Coil+Frame Gap OK 5.83% vs Gap a little 37.37%.','Sheet "UV LED COIL SP": UV LED press condition #1, #2 etc.','Additional sheets reviewed but only flagship comparisons stored.'],
    'content_ko': ['시트 "Test coil chạm Ring Frame": Coil+Frame Gap OK 5.83% vs Gap a little 37.37%.','시트 "UV LED COIL SP": UV LED press 조건 #1/#2 등.','다른 시트는 검토했으나 핵심 비교만 저장.'],
    'content_vi': ['Sheet "Test coil chạm Ring Frame": Coil+Frame Gap OK 5.83% vs Gap a little 37.37%.','Sheet "UV LED COIL SP": điều kiện ép UV LED #1/#2…','Các sheet khác đã xem nhưng chỉ lưu so sánh chính.'],
    'test_conditions': [
        {'condition_group': 'coil_frame_gap', 'process': 'Sub2 Coil+Frame', 'changed_factor': 'Coil+Frame Gap', 'before_value': 'Gap OK', 'after_value': 'Gap a little', 'source_cells': ['C15','C17']},
    ],
    'results': [
        {'condition_id': 'c1', 'measurement_type': 'Function', 'condition_group': 'coil_frame_gap', 'date': '2024-04-19', 'input_count': 120, 'ok_count': 113, 'ng_count': 7, 'ng_rate_decimal': 0.05833, 'ng_rate_percent': 5.833, 'metric_name': 'Coil + Frame Gap OK', 'judgement': 'BASELINE', 'ng_breakdown': {'Noise': {'count': 3, 'rate': 0.4286}, 'Touch': {'count': 4, 'rate': 0.5714}}, 'sheet_name': 'Test coil chạm Ring Frame', 'source_cells': ['M15']},
        {'condition_id': 'c1', 'measurement_type': 'Function', 'condition_group': 'coil_frame_gap', 'date': '2024-04-19', 'input_count': 99, 'ok_count': 62, 'ng_count': 37, 'ng_rate_decimal': 0.3737, 'ng_rate_percent': 37.37, 'metric_name': 'Coil + Frame Gap a little (test)', 'judgement': 'WORSE', 'ng_breakdown': {'SPL': {'count': 1, 'rate': 0.027}, 'Noise': {'count': 15, 'rate': 0.4054}, 'Touch': {'count': 21, 'rate': 0.5676}}, 'sheet_name': 'Test coil chạm Ring Frame', 'source_cells': ['M17']},
    ],
    'conclusions': [{
        'topic_en': 'Coil + Frame Gap effect on Function',
        'topic_ko': 'Coil + Frame Gap 이 Function 에 미친 효과',
        'topic_vi': 'Ảnh hưởng Coil + Frame Gap lên Function',
        'statement_en': 'Gap OK 5.83% vs Gap a little 37.37%. Noise/Touch dominate when gap exists.',
        'statement_ko': 'Gap OK 5.83% vs Gap a little 37.37%. Gap 존재 시 Noise/Touch 가 지배적.',
        'statement_vi': 'Gap OK 5.83% vs Gap a little 37.37%. Có gap thì Noise/Touch chiếm phần lớn.',
        'interp_en': 'A small Coil+Frame gap is 540.7% worse on Function vs gap-OK same-event baseline. Coil sitting away from the Ring Frame is a primary driver of Hearing Noise/Touch in BRS-201506.',
        'interp_ko': 'Coil+Frame 미세 gap 은 동일 이벤트 gap OK 대비 Function 540.7% 악화. Coil 이 Ring Frame 과 띄어지면 Noise/Touch 가 폭증.',
        'interp_vi': 'Khe hở nhỏ Coil+Frame làm NG function tệ hơn 540.7% so với gap OK cùng sự kiện. Coil cách Ring Frame là yếu tố chính gây Noise/Touch.',
        'source_cells': ['M15','M17']
    }],
    'hints': [{'check_en': 'Enforce zero Coil↔Ring-Frame gap at Sub2', 'reason_en': 'Even slight Coil+Frame gap makes Function NG 5.83%→37.37% (540% worse).', 'check_ko': 'Sub2 에서 Coil↔Ring Frame gap 0 유지.', 'reason_ko': '미세 gap 만으로도 Function NG 5.83%→37.37% (540% 악화).', 'check_vi': 'Đảm bảo không có gap Coil↔Ring Frame tại Sub2.', 'reason_vi': 'Chỉ hở nhỏ đã đẩy NG function từ 5.83% lên 37.37% (tệ hơn 540%).', 'evidence_strength': 'strong', 'related_process': 'Sub2 Coil+Frame', 'related_part': 'Coil/Ring Frame', 'source_cells': ['M15:M17']}],
    'log': {'rationale_en': 'Workbook bundles multiple cause-checks for high NG function — stored the strongest same-event comparison (Coil+Frame gap). Classified normal_comparison.','rationale_ko': '여러 원인 점검을 묶은 워크북 — 가장 강한 동일 이벤트 비교(Coil+Frame gap) 저장. normal_comparison.','rationale_vi': 'Workbook gộp nhiều kiểm tra nguyên nhân — chỉ lưu so sánh mạnh nhất (Coil+Frame gap). Phân loại normal_comparison.','warnings_en': ['Other sheets (UV LED COIL SP, Test coil Ring Frame extended) were not row-extracted in this commit; rerun for sheet-level coverage if needed.'],'warnings_ko': ['UV LED COIL SP 등 다른 시트는 본 커밋에서 행 단위 미저장 — 필요 시 시트별 재실행.'],'warnings_vi': ['Các sheet khác (UV LED COIL SP, …) chưa được trích thành dòng — chạy lại nếu cần.']},
    'when_user_asks': ['BRS-201506 NG function high', 'Coil Frame gap', 'UV LED press SP coil'],
    'confidence': 0.6,
})

# dataset 11
D.append({
    'name': '19. MSU-L20S15-07 Report test New bond PT-8803M improve NG separate Susp in module - 2025.04.09',
    'report_type': 'normal_comparison', 'model': 'MSU-L20S15-07',
    'report_date': '2025-04-09', 'department': 'ME', 'marker': 'Nhung + Le', 'default_sheet': '09.04',
    'primary_defect': 'NG Separate Suspension', 'related_defects': ['Fr+Susp offset','Susp damage'],
    'parts': ['Suspension','Frame'], 'processes': ['Sub3 Bond Fr+Susp','Tension','Module'],
    'title_en': 'MSU-L20S15-07 — Test new bond PT-8803M (vs PT-8310M9S) for Fr/Suspension ass’y (2025-04-09)',
    'title_ko': 'MSU-L20S15-07 — Fr/Suspension ass’y 신규 bond PT-8803M (Normal PT-8310M9S) 검증 (2025-04-09)',
    'title_vi': 'MSU-L20S15-07 — Thử bond mới PT-8803M (so với PT-8310M9S) cho ass’y Fr/Suspension (09-04-2025)',
    'purpose_en': 'Test new bond PT-8803M to improve NG separate suspension on module line 526/626. Compared to normal lot using bond PT-8310M9S.',
    'purpose_ko': '모듈 라인 526/626 의 NG separate suspension 개선 위해 신규 bond PT-8803M 적용 — Normal lot (bond PT-8310M9S) 와 비교.',
    'purpose_vi': 'Thử bond mới PT-8803M để giảm NG separate suspension trên line module 526/626 — so với lot Normal dùng bond PT-8310M9S.',
    'content_en': ['Sub3: Test 72pcs Bond 0.26mg vs Normal 72pcs Bond 0.30–0.36mg (both 0 NG).','Sub3 follow-up: Test 1120pcs (0.26–0.28mg) 2 Fr+Susp offset, Normal 1200pcs (0.26–0.4mg) 2 Fr+Susp offset.','Tension test Type A/B 10pcs each + module follow-up.'],
    'content_ko': ['Sub3: Test 72pcs (Bond 0.26mg) vs Normal 72pcs (Bond 0.30–0.36mg) — 둘 다 0 NG.','Sub3 후속: Test 1120pcs (0.26–0.28mg) Fr+Susp offset 2건, Normal 1200pcs (0.26–0.4mg) Fr+Susp offset 2건.','Tension Type A/B 10pcs 및 module 검증.'],
    'content_vi': ['Sub3: Test 72pcs (Bond 0.26mg) vs Normal 72pcs (Bond 0.30–0.36mg) — đều 0 NG.','Sub3 sau đó: Test 1120pcs (0.26–0.28mg) 2 lỗi Fr+Susp offset, Normal 1200pcs (0.26–0.4mg) 2 lỗi.','Tension Type A/B 10pcs và theo dõi module.'],
    'test_conditions': [
        {'condition_group': 'bond_change', 'process': 'Sub3 Fr+Susp bonding', 'changed_factor': 'Bond material', 'before_value': 'PT-8310M9S (Normal)', 'after_value': 'PT-8803M (Test)', 'source_cells': ['B6']},
    ],
    'results': [
        {'condition_id': 'c1', 'measurement_type': 'NG separate suspension', 'date': '2025-04-09', 'input_count': 72, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Test PT-8803M (bond 0.26mg)', 'judgement': 'SIMILAR', 'source_cells': ['K16']},
        {'condition_id': 'c1', 'measurement_type': 'NG separate suspension', 'date': '2025-04-09', 'input_count': 72, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Normal PT-8310M9S (bond 0.30–0.36mg)', 'judgement': 'BASELINE', 'source_cells': ['K18']},
        {'condition_id': 'c1', 'measurement_type': 'NG separate suspension', 'date': '2025-04-14', 'input_count': 1120, 'ng_count': 2, 'ng_rate_decimal': 0.001786, 'ng_rate_percent': 0.1786, 'metric_name': 'Test PT-8803M (n=1120, bond 0.26–0.28mg)', 'judgement': 'SIMILAR', 'ng_breakdown': {'Fr+Susp offset': {'count': 2, 'rate': 0.001786}}, 'source_cells': ['K20']},
        {'condition_id': 'c1', 'measurement_type': 'NG separate suspension', 'date': '2025-04-14', 'input_count': 1200, 'ng_count': 2, 'ng_rate_decimal': 0.001667, 'ng_rate_percent': 0.1667, 'metric_name': 'Normal PT-8310M9S (n=1200, bond 0.26–0.4mg)', 'judgement': 'BASELINE', 'ng_breakdown': {'Fr+Susp offset': {'count': 2, 'rate': 0.001667}}, 'source_cells': ['K22']},
    ],
    'conclusions': [{
        'topic_en': 'New bond PT-8803M vs Normal PT-8310M9S',
        'topic_ko': '신규 bond PT-8803M vs Normal PT-8310M9S',
        'topic_vi': 'Bond mới PT-8803M vs Normal PT-8310M9S',
        'statement_en': 'Sub3 small-lot 0 vs 0 NG. Sub3 larger lot Test 0.179% vs Normal 0.167% — both Fr+Susp offset only.',
        'statement_ko': 'Sub3 소량 lot 0 vs 0 NG. 대량 lot Test 0.179% vs Normal 0.167% — Fr+Susp offset 만.',
        'statement_vi': 'Sub3 lot nhỏ 0 vs 0 NG. Lot lớn Test 0.179% vs Normal 0.167% — chỉ Fr+Susp offset.',
        'interp_en': 'PT-8803M is essentially equivalent to PT-8310M9S on NG separate suspension: 0/72 vs 0/72 and 0.179% vs 0.167% (test/baseline = 1.072 → 7.2% nominally worse but absolute Δ is 1pp×scale, within noise). No clear improvement on Separate Suspension specifically — failures are now Fr+Susp offset, a different mode.',
        'interp_ko': 'PT-8803M 은 NG separate suspension 에서 PT-8310M9S 와 사실상 동등: 0/72 vs 0/72, 0.179% vs 0.167% (상대 +7.2%, 노이즈 범위). Separate Suspension 자체는 개선 없음 — NG 모드가 Fr+Susp offset 으로 전환.',
        'interp_vi': 'PT-8803M tương đương PT-8310M9S về NG separate suspension: 0/72 vs 0/72 và 0.179% vs 0.167% (+7.2%, trong nhiễu). Không cải thiện rõ Separate Suspension — lỗi chuyển sang Fr+Susp offset.',
        'source_cells': ['K16:K22']
    }],
    'hints': [{'check_en': 'Continue with PT-8310M9S; investigate Fr+Susp offset separately', 'reason_en': 'PT-8803M did not improve Separate Suspension; failure mode shifted to Fr+Susp offset.', 'check_ko': 'PT-8310M9S 유지, Fr+Susp offset 은 별도 검토.', 'reason_ko': 'PT-8803M 은 Separate Suspension 개선 없고 NG 모드가 Fr+Susp offset 으로 전환.', 'check_vi': 'Tiếp tục PT-8310M9S; điều tra Fr+Susp offset riêng.', 'reason_vi': 'PT-8803M không cải thiện Separate Suspension; lỗi chuyển sang Fr+Susp offset.', 'evidence_strength': 'medium', 'related_process': 'Sub3 Fr+Susp bonding', 'related_part': 'Bond / Suspension', 'source_cells': ['K20','K22']}],
    'log': {'rationale_en': 'Test bond vs Normal bond on same-event Sub3 results — normal_comparison. Both rows show ≈ same low NG; no meaningful improvement.','rationale_ko': 'Test bond vs Normal bond, Sub3 동일 이벤트 — normal_comparison. 양쪽 NG 거의 동일.','rationale_vi': 'Bond Test vs Normal, Sub3 cùng sự kiện — normal_comparison. NG gần như tương đương.','warnings_en': ['Tension/drop tests in source were not row-extracted in this commit; rerun if needed.'],'warnings_ko': ['Tension/drop 시험은 본 커밋에서 행 미저장 — 필요 시 재실행.'],'warnings_vi': ['Test tension/drop chưa được lưu — chạy lại nếu cần.']},
    'when_user_asks': ['NG separate suspension', 'Bond PT-8803M', 'MSU-L20S15-07 module'],
    'confidence': 0.7,
})

# dataset 12
D.append({
    'name': '19. MSU-L20S15-07DT  Report Test film 85A improve NG function SPK and module   date 30.5.2025',
    'report_type': 'normal_comparison', 'model': 'MSU-L20S15-07DT',
    'report_date': '2025-05-30', 'department': 'ME', 'marker': 'Le', 'default_sheet': '16.5',
    'primary_defect': 'NG Function (SPK + Module)', 'related_defects': ['Vision VP NG'],
    'parts': ['Film','VP','SPK module'], 'processes': ['Sub1 Vision VP','SPK function','Module function','Reliability'],
    'title_en': 'MSU-L20S15-07DT — Test new film 85A vs Normal 70A for SPK/Module function (2025-05-30)',
    'title_ko': 'MSU-L20S15-07DT — 신규 film 85A 와 Normal 70A 비교, SPK/Module Function 개선 검증 (2025-05-30)',
    'title_vi': 'MSU-L20S15-07DT — Thử film mới 85A so với Normal 70A để cải thiện Function SPK/Module (30-05-2025)',
    'purpose_en': 'Replace film 70A with film 85A and test SPK + module Function. Test 1: 400pcs; Test 2 conditional 2000pcs. Reliability checks: tension, drop, shock temperature, load, temperature/humidity (5pcs each).',
    'purpose_ko': 'film 70A 를 85A 로 교체해 SPK/module Function 검증. Test 1: 400pcs; Test 2 (NG 감소 시): 2000pcs. 신뢰성: tension, drop, shock temp, load, temp/humidity 각 5pcs.',
    'purpose_vi': 'Thay film 70A bằng 85A để kiểm Function SPK/module. Test 1: 400pcs; Test 2 (nếu NG giảm): 2000pcs. Độ tin cậy: tension, drop, shock nhiệt, load, nhiệt/ẩm mỗi loại 5pcs.',
    'content_en': ['Sub1 Vision VP table (Particle / VP cr…) for film 85A vs Normal.','SPK function and module function (further sheets).','Reliability gauntlet 5pcs each.'],
    'content_ko': ['Sub1 Vision VP (Particle / VP cr…) film 85A vs Normal.','SPK function 및 module function (추가 시트).','신뢰성 5pcs/항목.'],
    'content_vi': ['Sub1 Vision VP (Particle / VP cr…) film 85A vs Normal.','SPK function và module function (sheet thêm).','Độ tin cậy 5pcs/loại.'],
    'test_conditions': [
        {'condition_group': 'film_85a', 'process': 'Film selection', 'changed_factor': 'Film hardness', 'before_value': '70A (Normal)', 'after_value': '85A (Test)', 'source_cells': ['B6']},
    ],
    'results': [
        # Summary placeholder result rows — workbook is broad; reliability rows not stored to keep payload bounded.
        {'condition_id': 'c1', 'measurement_type': 'Plan', 'date': '2025-05-30', 'metric_name': 'Test 1 plan (n=400)', 'judgement': 'PLAN', 'source_cells': ['B8']},
        {'condition_id': 'c1', 'measurement_type': 'Plan', 'date': '2025-05-30', 'metric_name': 'Test 2 plan if NG reduced (n=2000)', 'judgement': 'PLAN', 'source_cells': ['B10']},
        {'condition_id': 'c1', 'measurement_type': 'Reliability Plan', 'date': '2025-05-30', 'metric_name': 'Tension / Drop / Shock temp / Load / Temp+Humidity (5pcs each)', 'judgement': 'PLAN', 'source_cells': ['B11:B16']},
    ],
    'conclusions': [{
        'topic_en': 'Film 85A test scope',
        'topic_ko': 'Film 85A 시험 범위',
        'topic_vi': 'Phạm vi thử film 85A',
        'statement_en': 'Plan: 400pcs first; expand to 2000pcs and module + reliability if NG reduced. Film 85A vs Normal 70A.',
        'statement_ko': '계획: 1차 400pcs, NG 감소 시 2000pcs + module + 신뢰성 확대. Film 85A vs Normal 70A.',
        'statement_vi': 'Kế hoạch: 400pcs trước; mở rộng 2000pcs + module + độ tin cậy nếu NG giảm. Film 85A vs Normal 70A.',
        'interp_en': 'Workbook captures the plan and Sub1 Vision VP framework. Quantitative SPK/Module comparison values were not extracted into AiResults rows in this commit — re-run with Sub1/SPK detail when finalizing claim of improvement.',
        'interp_ko': '워크북은 계획과 Sub1 Vision VP 프레임을 보존. 본 커밋에서는 SPK/Module 정량 결과 행 미저장 — 개선 주장 확정 시 Sub1/SPK 상세 재실행 권장.',
        'interp_vi': 'Workbook giữ kế hoạch và khung Vision VP Sub1. Lần commit này chưa lưu kết quả định lượng SPK/Module — nên chạy lại chi tiết khi xác nhận cải thiện.',
        'source_cells': ['B8:B16']
    }],
    'hints': [{'check_en': 'Compile Sub1 Vision VP + SPK function rows for film 85A vs Normal 70A before adopting', 'reason_en': 'Plan-only summary stored; full pass/fail and Function deltas need explicit extraction.', 'check_ko': 'film 85A 채택 전 Sub1 Vision VP + SPK function 결과 행 정리.', 'reason_ko': '계획 위주만 저장 — Function delta 정량 추가 추출 필요.', 'check_vi': 'Tổng hợp dòng Sub1 Vision VP + SPK function cho film 85A vs Normal 70A trước khi áp dụng.', 'reason_vi': 'Chỉ lưu tóm tắt kế hoạch — cần trích thêm Function delta định lượng.', 'evidence_strength': 'low', 'related_process': 'Film selection', 'related_part': 'Film', 'source_cells': ['B6']}],
    'log': {'rationale_en': 'Workbook is broad with reliability + multi-sheet results; this commit captured plan and condition only. Classified normal_comparison since stated comparator is Normal 70A.','rationale_ko': '신뢰성 + 다수 시트로 광범위 — 본 커밋은 계획/조건만 저장. 명시된 비교군이 Normal 70A 이므로 normal_comparison 분류.','rationale_vi': 'Workbook rộng với độ tin cậy + nhiều sheet — commit này chỉ lưu kế hoạch/điều kiện. Phân loại normal_comparison vì so sánh Normal 70A.','warnings_en': ['SPK and Module Function detail rows not extracted in this commit — re-run for full row coverage if Decision evidence is needed.','Report date stored from Excel serial 45807 (= 2025-05-30).'],'warnings_ko': ['SPK / Module Function 상세 행 미저장 — Decision 근거 필요 시 재실행.','Excel 직렬값 45807 → 2025-05-30 으로 변환 저장.'],'warnings_vi': ['Chưa lưu chi tiết SPK / Module Function — chạy lại nếu cần.','Ngày từ giá trị Excel 45807 → 30-05-2025.'],'assumptions_en': ['Excel serial 45807 interpreted as 2025-05-30.']},
    'when_user_asks': ['film 85A vs 70A','MSU-L20S15-07 SPK function','module reliability'],
    'confidence': 0.45,
})


if __name__ == '__main__':
    ok = 0; fail = 0
    for d in D:
        try:
            if B.commit(d):
                print(f'OK  {d["name"]}'); ok += 1
            else:
                print(f'FAIL {d["name"]}'); fail += 1
        except Exception as e:
            print(f'ERR {d["name"]}: {e!r}'); fail += 1
    print(f'BATCH RESULT: ok={ok} fail={fail}')
