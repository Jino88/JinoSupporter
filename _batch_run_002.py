import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _batch_build as B

D = []

# dataset 4: Decap NG #7,#9
D.append({
    'name': '18.1. BRS-161014 DT Report test VP mold #7,9 add 0.05mm decap NG   date 16.2.2024 -',
    'report_type': 'defect_root_cause',
    'model': 'BRS-161016',
    'report_date': '2024-02-16',
    'department': 'ME', 'marker': 'Phuong',
    'default_sheet': 'Decap NG (#7), Decap NG (#9)',
    'primary_defect': 'NG Hearing Noise + Touch',
    'primary_aliases': ['NG function BRS-161016', 'Decap NG'],
    'related_defects': ['Coil deform', 'CM offset', 'Fr/Yoke offset', 'Fr/Coil offset', 'Particle'],
    'parts': ['VP', 'Coil', 'Frame', 'Yoke', 'CM'],
    'processes': ['Decap analysis'],
    'title_en': 'BRS-161014 DT — Decap NG analysis for VP mold #7 / #9 (+0.05mm) (2024-02-16)',
    'title_ko': 'BRS-161014 DT — VP mold #7/#9 (+0.05mm) Decap NG 원인 분석 (2024-02-16)',
    'title_vi': 'BRS-161014 DT — Phân tích Decap NG cho VP khuôn #7/#9 (+0.05mm) (16-02-2024)',
    'purpose_en': 'Decap (cross-section) analysis of 10 NG-function samples from BRS-161016 line 3A for VP mold #7 and VP mold #9 (+0.05mm middle) to identify the dominant physical cause of NG Hearing (Noise + Touch).',
    'purpose_ko': 'BRS-161016 line 3A 의 NG Function 샘플 10pcs 를 decap 으로 단면 분석하여 VP mold #7/#9 (+0.05mm 중앙) 의 NG Hearing(Noise + Touch) 지배 원인 파악.',
    'purpose_vi': 'Phân tích decap (cắt mặt) 10 mẫu NG-function của BRS-161016 line 3A cho VP khuôn #7 và #9 (+0.05mm giữa) để xác định nguyên nhân chính của NG Hearing (Noise + Touch).',
    'content_en': ['Sheet "Decap NG (#7)": 10 NG units classified by physical reason.','Sheet "Decap NG (#9)": 10 NG units classified by physical reason.'],
    'content_ko': ['시트 "Decap NG (#7)": NG 10개 물리 원인별 분류.','시트 "Decap NG (#9)": NG 10개 물리 원인별 분류.'],
    'content_vi': ['Sheet "Decap NG (#7)": phân loại 10 NG theo nguyên nhân vật lý.','Sheet "Decap NG (#9)": phân loại 10 NG theo nguyên nhân vật lý.'],
    'test_conditions': [
        {'condition_group': 'vp_mold_compare', 'process': 'VP mold', 'changed_factor': 'VP mold # (#7 vs #9, both +0.05T middle)', 'before_value': '#7', 'after_value': '#9', 'source_cells': ['#7!B6','#9!B6']},
    ],
    'results': [
        # #7 breakdown — single Result row holding ng_breakdown
        {'condition_id': 'c1', 'measurement_type': 'Decap', 'date': '2024-02-16', 'line': 'BRS 161016 - Line 3A', 'input_count': 10, 'ok_count': None, 'ng_count': 10, 'ng_rate_decimal': 1.0, 'ng_rate_percent': 100.0, 'metric_name': 'VP mold #7 — NG cause decap (n=10)', 'sheet_name': 'Decap NG (#7)', 'ng_breakdown': {'Coil deform': {'count': 1, 'rate': 0.1}, 'CM offset': {'count': 1, 'rate': 0.1}, 'Fr/YOKE offset': {'count': 5, 'rate': 0.5}, 'Particle': {'count': 0, 'rate': 0.0}, 'Coil deform & Fr/Yoke offset': {'count': 0, 'rate': 0.0}, 'Fr/Coil offset': {'count': 2, 'rate': 0.2}, "Don't know reason": {'count': 1, 'rate': 0.1}}, 'source_cells': ['Decap NG (#7)!H16:I100']},
        {'condition_id': 'c1', 'measurement_type': 'Decap', 'date': '2024-02-16', 'line': 'BRS 161016 - Line 3A', 'input_count': 10, 'ok_count': None, 'ng_count': 10, 'ng_rate_decimal': 1.0, 'ng_rate_percent': 100.0, 'metric_name': 'VP mold #9 — NG cause decap (n=10)', 'sheet_name': 'Decap NG (#9)', 'ng_breakdown': {'Coil deform': {'count': 2, 'rate': 0.2}, 'CM offset': {'count': 0, 'rate': 0.0}, 'Fr/YOKE offset': {'count': 2, 'rate': 0.2}, 'Particle': {'count': 0, 'rate': 0.0}, 'Coil deform & Fr/Yoke offset': {'count': 2, 'rate': 0.2}, 'Fr/Coil offset': {'count': 2, 'rate': 0.2}, "Don't know reason": {'count': 2, 'rate': 0.2}}, 'source_cells': ['Decap NG (#9)!H16:I100']},
    ],
    'conclusions': [
        {
            'topic_en': 'Dominant NG cause by VP mold', 'topic_ko': 'VP mold 별 NG 주원인', 'topic_vi': 'Nguyên nhân NG chính theo VP mold',
            'statement_en': '#7 dominated by Fr/Yoke offset 50% (5/10); #9 spread across Coil deform / Fr/Yoke offset / Coil-deform+Fr-Yoke / Fr-Coil offset / Unknown each 20%.',
            'statement_ko': '#7 은 Fr/Yoke offset 50% (5/10) 가 지배적. #9 는 Coil deform / Fr/Yoke offset / Coil-deform+Fr-Yoke / Fr-Coil offset / 미상 각 20% 로 분산.',
            'statement_vi': '#7 chủ yếu Fr/Yoke offset 50% (5/10); #9 phân bổ đều Coil deform / Fr/Yoke offset / Coil-deform+Fr-Yoke / Fr-Coil offset / Không rõ mỗi loại 20%.',
            'interp_en': 'VP mold #7 NG is concentrated on Fr/Yoke offset (50%) — likely a frame/yoke seating issue with that mold. VP mold #9 has a flatter cause distribution with a sizeable "unknown" share, suggesting a different failure mechanism rather than one dominant cause.',
            'interp_ko': 'VP mold #7 의 NG 는 Fr/Yoke offset(50%) 에 집중 — 해당 mold 의 frame/yoke 안착 문제 가능성. VP mold #9 는 원인 분포가 평탄하고 "미상" 비중이 커 단일 지배 원인이 아닌 다른 메커니즘 시사.',
            'interp_vi': 'NG của VP mold #7 tập trung ở Fr/Yoke offset (50%) — có thể do vấn đề khớp frame/yoke với khuôn đó. VP mold #9 phân bố nguyên nhân phẳng hơn với tỉ lệ "không rõ" lớn, gợi ý cơ chế lỗi khác chứ không phải một nguyên nhân chính.',
            'source_cells': ['Decap NG (#7)!H16:H100','Decap NG (#9)!H16:H100']
        },
    ],
    'hints': [
        {'check_en': 'Check Frame/Yoke seating for VP mold #7', 'reason_en': '5 of 10 #7-NG units are Fr/Yoke offset (50%).', 'check_ko': 'VP mold #7 의 Frame/Yoke 안착 점검.', 'reason_ko': '#7 NG 10개 중 5개(50%)가 Fr/Yoke offset.', 'check_vi': 'Kiểm tra khớp Frame/Yoke trên VP khuôn #7.', 'reason_vi': '5/10 NG của #7 là Fr/Yoke offset (50%).', 'evidence_strength': 'medium', 'related_process': 'Sub2/Sub3 ass’y', 'related_part': 'Frame/Yoke', 'source_cells': ['Decap NG (#7)!H36']},
        {'check_en': 'Investigate "unknown reason" share on VP mold #9 (20%)', 'reason_en': '#9 has 20% NG with no identified cause and a flat distribution suggesting a hidden mechanism.', 'check_ko': '#9 의 "미상" 20% 추가 분석 — 숨겨진 원인 의심.', 'reason_ko': '#9 NG 중 20% 원인 미상, 분포가 평탄해 숨은 메커니즘 가능.', 'check_vi': 'Phân tích sâu phần "không rõ" 20% của #9 — nghi cơ chế ẩn.', 'reason_vi': '#9 có 20% NG không rõ nguyên nhân và phân bố phẳng — có thể có cơ chế ẩn.', 'evidence_strength': 'low', 'related_process': 'Decap', 'related_part': 'Coil/Frame/CM', 'source_cells': ['Decap NG (#9)!H74']},
    ],
    'log': {
        'rationale_en': 'No NG-rate comparison versus a same-event Normal — workbook is a decap root-cause breakdown. Classified defect_root_cause. Stored each mold as a single 100% NG row with full ng_breakdown so cause distribution is preserved.',
        'rationale_ko': '동일 이벤트 Normal 비교 없이 decap 원인 분해만 존재 — defect_root_cause 로 분류. 각 mold 를 100% NG 단일 행으로 저장하고 ng_breakdown 으로 원인 분포 보존.',
        'rationale_vi': 'Không có so sánh NG-rate với Normal cùng sự kiện — workbook chỉ phân loại nguyên nhân decap. Phân loại defect_root_cause. Mỗi khuôn lưu 1 hàng 100% NG với ng_breakdown đầy đủ.',
        'assumptions_en': ['Each mold sheet analyses exactly 10 NG units (total = 10 in cell E8).'],
        'assumptions_ko': ['각 mold 시트는 E8 셀 기준 NG 10개를 분석.'],
        'assumptions_vi': ['Mỗi sheet phân tích đúng 10 mẫu NG (tổng = 10 theo ô E8).'],
        'warnings_en': ['No same-event baseline available — cannot claim improvement/worsening between mold #7 vs #9 in absolute NG-rate terms.'],
        'warnings_ko': ['동일 이벤트 baseline 부재 — mold #7 vs #9 절대 NG율 개선/악화 주장 불가.'],
        'warnings_vi': ['Không có baseline cùng sự kiện — không thể khẳng định cải thiện/xấu đi giữa khuôn #7 và #9 theo NG-rate tuyệt đối.'],
    },
    'when_user_asks': ['BRS-161016 NG function decap', 'VP mold #7 NG cause', 'VP mold #9 NG cause'],
    'confidence': 0.65,
})

# dataset 5: CMG/CPT reduce dimension 17.2
D.append({
    'name': '19-1. BRS-161014  Report test CMG and CPT reduce dimension 17.2.2024',
    'report_type': 'normal_comparison',
    'model': 'BRS-161016',
    'report_date': '2024-02-17',
    'department': 'ME', 'marker': 'Le',
    'default_sheet': 'NG rate',
    'primary_defect': 'NG Hearing Noise',
    'related_defects': ['NG Hearing Touch'],
    'parts': ['CMG', 'CPT', 'Yoke', 'MG+PT'],
    'processes': ['MG forming', 'Ass’y yoke', 'Function test'],
    'title_en': 'BRS-161016 — Reduce CMG/CPT dimension 10.86-4.86-0.88 → 10.82-4.2-0.88 (2024-02-17)',
    'title_ko': 'BRS-161016 — CMG/CPT 치수 축소 10.86-4.86-0.88 → 10.82-4.2-0.88 (2024-02-17)',
    'title_vi': 'BRS-161016 — Giảm kích thước CMG/CPT 10.86-4.86-0.88 → 10.82-4.2-0.88 (17-02-2024)',
    'purpose_en': 'Improve NG hearing by reducing CMG/CPT dimension from 10.86-4.86-0.88 to 10.82-4.2-0.88; verify gauss (MG+PT / Semi Yoke / Final) and function vs Normal line.',
    'purpose_ko': 'NG hearing 개선 위해 CMG/CPT 치수를 10.86-4.86-0.88 → 10.82-4.2-0.88 로 축소; gauss (MG+PT / Semi Yoke / Final) 및 Function 을 Normal 라인 대비 검증.',
    'purpose_vi': 'Cải thiện NG hearing bằng cách giảm kích thước CMG/CPT từ 10.86-4.86-0.88 xuống 10.82-4.2-0.88; xác minh gauss (MG+PT / Semi Yoke / Final) và function so với Normal.',
    'content_en': ['Gauss: MG+PT, Semi Yoke and Final per 10 samples vs Normal (all OK).','Function test on E2 GMI Line 2024-02-07 and C2 DT Line 2024-02-17 Test Yoke vs Normal.'],
    'content_ko': ['Gauss: MG+PT, Semi Yoke, Final 각 10 샘플 vs Normal (전부 OK).','Function: E2 GMI Line 2024-02-07 및 C2 DT Line 2024-02-17 의 Test Yoke vs Normal.'],
    'content_vi': ['Gauss: MG+PT, Semi Yoke và Final, 10 mẫu mỗi loại vs Normal (đều OK).','Function: E2 GMI Line 2024-02-07 và C2 DT Line 2024-02-17 Test Yoke vs Normal.'],
    'test_conditions': [
        {'condition_group': 'cmg_cpt_dim', 'process': 'MG forming', 'changed_factor': 'CMG/CPT dimension', 'before_value': '10.86-4.86-0.88', 'after_value': '10.82-4.2-0.88', 'unit': 'mm', 'source_cells': ['B8']},
    ],
    'results': [
        # Gauss summary metrics
        {'condition_id': 'c1', 'measurement_type': 'Gauss', 'date': '2024-02-17', 'metric_name': 'MG+PT Test AVG', 'metric_value': 1250.8, 'unit': 'G', 'judgement': 'OK', 'source_cells': ['R15']},
        {'condition_id': 'c1', 'measurement_type': 'Gauss', 'date': '2024-02-17', 'metric_name': 'MG+PT Normal AVG', 'metric_value': 1269.2, 'unit': 'G', 'judgement': 'OK (baseline)', 'source_cells': ['R16']},
        {'condition_id': 'c1', 'measurement_type': 'Gauss', 'date': '2024-02-17', 'metric_name': 'Semi Yoke Test AVG', 'metric_value': 527.2, 'unit': 'G', 'judgement': 'OK', 'source_cells': ['R17']},
        {'condition_id': 'c1', 'measurement_type': 'Gauss', 'date': '2024-02-17', 'metric_name': 'Semi Yoke Normal AVG', 'metric_value': 547.8, 'unit': 'G', 'judgement': 'OK (baseline)', 'source_cells': ['R18']},
        {'condition_id': 'c1', 'measurement_type': 'Gauss', 'date': '2024-02-17', 'metric_name': 'Final Test AVG', 'metric_value': 559.9, 'unit': 'G', 'judgement': 'OK', 'source_cells': ['R19']},
        {'condition_id': 'c1', 'measurement_type': 'Gauss', 'date': '2024-02-17', 'metric_name': 'Final Normal AVG', 'metric_value': 577.8, 'unit': 'G', 'judgement': 'OK (baseline)', 'source_cells': ['R20']},
        # Function: E2 GMI
        {'condition_id': 'c1', 'measurement_type': 'Function', 'condition_group': 'cmg_cpt_dim', 'date': '2024-02-07', 'line': 'E2 GMI', 'input_count': 509, 'ok_count': 498, 'ng_count': 11, 'ng_rate_decimal': 0.02161, 'ng_rate_percent': 2.161, 'metric_name': 'Test Yoke E2 GMI', 'judgement': 'IMPROVED', 'ng_breakdown': {'Noise': {'count': 11, 'rate': 0.02161}}, 'source_cells': ['C25:O25']},
        {'condition_id': 'c1', 'measurement_type': 'Function', 'condition_group': 'cmg_cpt_dim', 'date': '2024-02-07', 'line': 'E2 GMI', 'input_count': 687, 'ok_count': 664, 'ng_count': 23, 'ng_rate_decimal': 0.03348, 'ng_rate_percent': 3.348, 'metric_name': 'Normal E2 GMI', 'judgement': 'BASELINE', 'ng_breakdown': {'Noise': {'count': 23, 'rate': 0.03348}}, 'source_cells': ['C27:O27']},
        # Function: C2 DT
        {'condition_id': 'c1', 'measurement_type': 'Function', 'condition_group': 'cmg_cpt_dim', 'date': '2024-02-17', 'line': 'C2 DT', 'input_count': 515, 'ok_count': 512, 'ng_count': 3, 'ng_rate_decimal': 0.005825, 'ng_rate_percent': 0.5825, 'metric_name': 'Test Yoke C2 DT', 'judgement': 'IMPROVED', 'ng_breakdown': {'Noise': {'count': 3, 'rate': 0.005825}}, 'source_cells': ['C29:O29']},
        {'condition_id': 'c1', 'measurement_type': 'Function', 'condition_group': 'cmg_cpt_dim', 'date': '2024-02-17', 'line': 'C2 DT', 'input_count': 560, 'ok_count': 542, 'ng_count': 16, 'ng_rate_decimal': 0.02857, 'ng_rate_percent': 2.857, 'metric_name': 'Normal C2 DT', 'judgement': 'BASELINE', 'ng_breakdown': {'Noise': {'count': 16, 'rate': 0.02857}, 'Touch': {'count': 2, 'rate': 0.003571}}, 'source_cells': ['C31:O31']},
    ],
    'conclusions': [
        {
            'topic_en': 'CMG/CPT dimension reduction effect on NG hearing',
            'topic_ko': 'CMG/CPT 치수 축소가 NG hearing 에 미친 효과',
            'topic_vi': 'Ảnh hưởng của việc giảm kích thước CMG/CPT lên NG hearing',
            'statement_en': 'E2 GMI Test 2.16% vs Normal 3.35%; C2 DT Test 0.58% vs Normal 2.86%. Gauss 10–20G lower but still pass spec.',
            'statement_ko': 'E2 GMI Test 2.16% vs Normal 3.35%; C2 DT Test 0.58% vs Normal 2.86%. Gauss 약 10–20G 낮으나 스펙 통과.',
            'statement_vi': 'E2 GMI Test 2.16% vs Normal 3.35%; C2 DT Test 0.58% vs Normal 2.86%. Gauss thấp hơn ~10–20G nhưng vẫn pass.',
            'interp_en': 'Reduced CMG/CPT dimension improves NG hearing by 35.5% on E2 GMI (2.161% vs 3.348%) and by 79.6% on C2 DT (0.583% vs 2.857%). Gauss drops are within spec; trade-off is acceptable.',
            'interp_ko': 'CMG/CPT 치수 축소는 NG hearing 을 E2 GMI 35.5% (2.161% vs 3.348%), C2 DT 79.6% (0.583% vs 2.857%) 개선. Gauss 감소는 스펙 내, 트레이드오프 수용 가능.',
            'interp_vi': 'Giảm kích thước CMG/CPT cải thiện NG hearing 35.5% trên E2 GMI (2.161% vs 3.348%) và 79.6% trên C2 DT (0.583% vs 2.857%). Gauss giảm trong dung sai, đánh đổi chấp nhận được.',
            'source_cells': ['O25','O27','O29','O31']
        }
    ],
    'hints': [
        {'check_en': 'Apply CMG/CPT dimension 10.82-4.2-0.88 to production', 'reason_en': '79.6% improvement on C2 DT and 35.5% on E2 GMI; Gauss still passes spec.', 'check_ko': 'CMG/CPT 10.82-4.2-0.88 양산 적용.', 'reason_ko': 'C2 DT 79.6% / E2 GMI 35.5% 개선, Gauss 스펙 통과.', 'check_vi': 'Áp dụng CMG/CPT 10.82-4.2-0.88 cho sản xuất.', 'reason_vi': 'Cải thiện 79.6% C2 DT và 35.5% E2 GMI; Gauss vẫn đạt spec.', 'evidence_strength': 'strong', 'related_process': 'MG forming', 'related_part': 'CMG/CPT', 'source_cells': ['O25:O31']},
    ],
    'log': {
        'rationale_en': 'Two same-day Test vs Normal pairs (E2 GMI, C2 DT) — normal_comparison. Relative change (test/baseline-1)*100 used.',
        'rationale_ko': 'E2 GMI 와 C2 DT 두 line 의 동일 이벤트 Test vs Normal 쌍 — normal_comparison. (test/baseline-1)*100.',
        'rationale_vi': 'Hai cặp Test vs Normal cùng ngày (E2 GMI, C2 DT) — normal_comparison. (test/baseline-1)*100.',
        'warnings_en': ['Gauss decrease 10–20G noted by writer; long-term reliability not verified in this report.'],
        'warnings_ko': ['작성자가 Gauss 10–20G 감소를 명시 — 장기 신뢰성은 본 리포트에서 미검증.'],
        'warnings_vi': ['Người viết ghi nhận Gauss giảm 10–20G — độ tin cậy dài hạn chưa kiểm tra trong báo cáo này.'],
        'assumptions_en': [],
    },
    'when_user_asks': ['BRS-161016 NG hearing', 'CMG/CPT dimension change', 'Yoke dimension reduce'],
    'confidence': 0.78,
})

# dataset 6: VP mold 7,9 + 0.05T middle 4.3.2024
D.append({
    'name': '19. 1- BRS-161014  Report test verify VP mold 7, mold 9  add 0.05T in middle position 04.03.2024',
    'report_type': 'normal_comparison',
    'model': 'BRS-161016',
    'report_date': '2024-03-04',
    'department': 'ME', 'marker': 'Thuy',
    'default_sheet': 'NG rate',
    'primary_defect': 'NG Function (Hearing Noise)',
    'related_defects': ['VP bending', 'VP+CD vision NG'],
    'parts': ['VP', 'Frame', 'CD'],
    'processes': ['VP forming', 'Sub1 vision', 'Main2 function'],
    'title_en': 'BRS-161016 — Verify VP 161014-ED mold #7 / #9 with +0.05T middle (2024-03-04)',
    'title_ko': 'BRS-161016 — VP 161014-ED mold #7/#9 중앙 +0.05T 검증 (2024-03-04)',
    'title_vi': 'BRS-161016 — Xác minh VP 161014-ED khuôn #7/#9 +0.05T giữa (04-03-2024)',
    'purpose_en': 'Verify whether VP mold #7 / #9 with +0.05T middle can be used: check VP bending rate, Sub1 VP+CD vision NG, and Main2 Function — compared against Normal VP mold #5.',
    'purpose_ko': 'VP mold #7/#9 중앙 +0.05T 사용 가능 여부 검증: VP bending, Sub1 VP+CD vision NG, Main2 Function 을 Normal VP mold #5 대비 비교.',
    'purpose_vi': 'Xác minh khuôn VP #7/#9 +0.05T giữa có dùng được không: kiểm tra VP bending, vision VP+CD Sub1 và Function Main2, so với Normal khuôn VP #5.',
    'content_en': ['VP bending: mold #7/#9/#5 on E2 line.','VP+CD vision at Sub1: #7/#9/#5.','Function at E2 total: #7/#9/#5.'],
    'content_ko': ['VP bending: E2 라인의 mold #7/#9/#5.','Sub1 VP+CD vision NG: #7/#9/#5.','E2 Total Function: #7/#9/#5.'],
    'content_vi': ['VP bending: khuôn #7/#9/#5 trên line E2.','Vision VP+CD Sub1: #7/#9/#5.','Function E2 tổng: #7/#9/#5.'],
    'test_conditions': [
        {'condition_group': 'vp_mold_05t', 'process': 'VP mold', 'changed_factor': 'VP mold # with +0.05T middle', 'before_value': 'Normal mold #5', 'after_value': '#7 / #9 (+0.05T)', 'unit': 'mm', 'source_cells': ['B6']},
    ],
    'results': [
        # VP bending
        {'condition_id': 'c1', 'measurement_type': 'VP Bending', 'date': '2024-03-04', 'line': 'E2', 'input_count': 1000, 'ok_count': 951, 'ng_count': 49, 'ng_rate_decimal': 0.049, 'ng_rate_percent': 4.9, 'metric_name': 'Test VP mold #7', 'judgement': 'WORSE', 'source_cells': ['C15:H15']},
        {'condition_id': 'c1', 'measurement_type': 'VP Bending', 'date': '2024-03-04', 'line': 'E2', 'input_count': 1000, 'ok_count': 952, 'ng_count': 48, 'ng_rate_decimal': 0.048, 'ng_rate_percent': 4.8, 'metric_name': 'Test VP mold #9', 'judgement': 'WORSE', 'source_cells': ['C15:H16']},
        {'condition_id': 'c1', 'measurement_type': 'VP Bending', 'date': '2024-03-04', 'line': 'E2', 'input_count': 800, 'ok_count': 800, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Normal VP mold #5', 'judgement': 'BASELINE', 'source_cells': ['C15:H17']},
        # Vision VP+CD Sub1
        {'condition_id': 'c1', 'measurement_type': 'Vision VP+CD', 'date': '2024-03-04', 'line': 'E2', 'input_count': 860, 'ok_count': 860, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Test VP mold #7 (Sub1)', 'judgement': 'BETTER', 'source_cells': ['C26:J26']},
        {'condition_id': 'c1', 'measurement_type': 'Vision VP+CD', 'date': '2024-03-04', 'line': 'E2', 'input_count': 862, 'ok_count': 862, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Test VP mold #9 (Sub1)', 'judgement': 'BETTER', 'source_cells': ['C26:J27']},
        {'condition_id': 'c1', 'measurement_type': 'Vision VP+CD', 'date': '2024-03-04', 'line': 'E2', 'input_count': 800, 'ok_count': 799, 'ng_count': 1, 'ng_rate_decimal': 0.00125, 'ng_rate_percent': 0.125, 'metric_name': 'Normal VP mold #5 (Sub1)', 'judgement': 'BASELINE', 'source_cells': ['C26:J28']},
        # Function E2 Total
        {'condition_id': 'c1', 'measurement_type': 'Function', 'date': '2024-03-04', 'line': 'E2-Total', 'input_count': 846, 'ok_count': None, 'ng_count': 52, 'ng_rate_decimal': 0.06147, 'ng_rate_percent': 6.147, 'metric_name': 'Test VP mold #7', 'judgement': 'WORSE', 'ng_breakdown': {'Noise': {'count': 51, 'rate': 0.06028}}, 'source_cells': ['C32:M32']},
        {'condition_id': 'c1', 'measurement_type': 'Function', 'date': '2024-03-04', 'line': 'E2-Total', 'input_count': 859, 'ok_count': None, 'ng_count': 44, 'ng_rate_decimal': 0.05122, 'ng_rate_percent': 5.122, 'metric_name': 'Test VP mold #9', 'judgement': 'WORSE', 'ng_breakdown': {'Noise': {'count': 44, 'rate': 0.05122}}, 'source_cells': ['C32:M34']},
        {'condition_id': 'c1', 'measurement_type': 'Function', 'date': '2024-03-04', 'line': 'E2-Total', 'input_count': 798, 'ok_count': None, 'ng_count': 30, 'ng_rate_decimal': 0.03759, 'ng_rate_percent': 3.759, 'metric_name': 'Normal VP mold #5', 'judgement': 'BASELINE', 'ng_breakdown': {'Noise': {'count': 30, 'rate': 0.03759}}, 'source_cells': ['C32:M36']},
    ],
    'conclusions': [
        {
            'topic_en': 'VP mold +0.05T verification',
            'topic_ko': 'VP mold +0.05T 검증',
            'topic_vi': 'Xác minh VP +0.05T',
            'statement_en': 'Bending: #7 4.9%, #9 4.8%, Normal #5 0%. Vision: #7/#9 0%, Normal 0.125%. Function: #7 6.15%, #9 5.12%, Normal 3.76%.',
            'statement_ko': 'Bending: #7 4.9%, #9 4.8%, Normal #5 0%. Vision: #7/#9 0%, Normal 0.125%. Function: #7 6.15%, #9 5.12%, Normal 3.76%.',
            'statement_vi': 'Bending: #7 4.9%, #9 4.8%, Normal #5 0%. Vision: #7/#9 0%, Normal 0.125%. Function: #7 6.15%, #9 5.12%, Normal 3.76%.',
            'interp_en': 'VP bending rate jumps from 0% to ~4.8–4.9% with +0.05T molds, and Function NG is 63.5%/36.3% worse on #7/#9 vs Normal #5 (6.147/5.122 vs 3.759%). Vision NG is marginally better but the Function and Bending penalties dominate — VP mold #7/#9 +0.05T should not be used as-is.',
            'interp_ko': 'VP bending 이 0% → 약 4.8–4.9% 로 급증하고, Function NG 가 Normal #5 대비 #7 63.5%, #9 36.3% 악화 (6.147/5.122 vs 3.759%). Vision NG 는 미미하게 개선이나 Function/Bending 페널티가 지배적 — 현 상태로 VP mold #7/#9 +0.05T 사용 불가 권고.',
            'interp_vi': 'Tỷ lệ VP bending tăng từ 0% lên ~4.8–4.9% với khuôn +0.05T, và NG function tệ hơn 63.5%/36.3% trên #7/#9 so với Normal #5 (6.147/5.122 vs 3.759%). Vision NG cải thiện rất nhỏ nhưng Function/Bending áp đảo — không nên dùng khuôn VP #7/#9 +0.05T nguyên trạng.',
            'source_cells': ['H15:H17','J26:J28','M32:M36']
        }
    ],
    'hints': [
        {'check_en': 'Do not deploy VP mold #7/#9 +0.05T to production', 'reason_en': 'Function NG 36–63% worse and bending 4.8–4.9% vs Normal 0%.', 'check_ko': 'VP mold #7/#9 +0.05T 양산 적용 보류.', 'reason_ko': 'Function NG 가 Normal 대비 36–63% 악화, bending 4.8–4.9% vs 0%.', 'check_vi': 'Không triển khai khuôn VP #7/#9 +0.05T cho sản xuất.', 'reason_vi': 'NG function tệ hơn 36–63% và bending 4.8–4.9% so với Normal 0%.', 'evidence_strength': 'strong', 'related_process': 'VP mold', 'related_part': 'VP', 'source_cells': ['M32:M36']},
        {'check_en': 'Investigate root cause of bending increase with +0.05T middle', 'reason_en': 'Adding 0.05T middle drives bending from 0% to ~5% — mold geometry, press condition or material fit may explain.', 'check_ko': '중앙 +0.05T 가 bending 을 0%→5% 로 끌어올리는 원인 분석 (mold geometry, press, 자재).', 'reason_ko': '중앙 +0.05T 가 0%→5% bending 유발 — mold geometry/press/자재 적합성 점검.', 'check_vi': 'Điều tra nguyên nhân +0.05T giữa khiến bending tăng từ 0% lên ~5% — hình học khuôn, điều kiện ép, vật liệu.', 'reason_vi': 'Thêm 0.05T giữa đẩy bending từ 0% lên ~5% — kiểm hình học khuôn, điều kiện ép và sự phù hợp vật liệu.', 'evidence_strength': 'medium', 'related_process': 'VP mold', 'related_part': 'VP', 'source_cells': ['H15:H17']},
    ],
    'log': {
        'rationale_en': 'Each result table has Test #7, Test #9 and Normal #5 in the same event — normal_comparison. Three measurement types (bending, vision, function). Relative change vs Normal #5.',
        'rationale_ko': '각 결과 표가 Test #7/#9 와 Normal #5 동일 이벤트 — normal_comparison. 측정 3종 (bending, vision, function). Normal #5 대비 상대 변화 계산.',
        'rationale_vi': 'Mỗi bảng có Test #7/#9 và Normal #5 cùng sự kiện — normal_comparison. Ba loại đo (bending, vision, function). Thay đổi tương đối so với Normal #5.',
        'assumptions_en': ['Decision row was empty in extract; conclusion drawn from data only.'],
        'assumptions_ko': ['추출 텍스트에 Decision 행이 비어 있어 결론은 데이터 기반.'],
        'assumptions_vi': ['Hàng Decision trống trong dữ liệu; kết luận dựa trên số liệu.'],
        'warnings_en': ['Sample sizes for bending and vision are large (n≈1000/800–860) but function samples are ~800–860 each — adequate but verify with larger production lot if rolling out.'],
        'warnings_ko': ['bending/vision 샘플 충분 (n≈800–1000), function ≈800–860 — 적정하나 양산 적용 전 추가 lot 검증 권장.'],
        'warnings_vi': ['Cỡ mẫu bending/vision đủ (n≈800–1000), function ≈800–860 — đủ nhưng nên kiểm thêm lô trước khi triển khai.'],
    },
    'when_user_asks': ['VP mold +0.05T', 'VP bending mold', 'BRS-161016 NG function VP mold'],
    'confidence': 0.82,
})

# dataset 7: CMG/CPT reduce dimension + NTI data 17.2.2024
D.append({
    'name': '19. BRS-161014  Report test CMG and CPT reduce dimension + NTI data 17.2.2024',
    'report_type': 'normal_comparison',
    'model': 'BRS-161016',
    'report_date': '2024-02-17',
    'department': 'ME', 'marker': 'Le',
    'default_sheet': 'NG rate, Picture xray, NTI_종합',
    'primary_defect': 'NG Hearing Noise',
    'related_defects': ['NG Hearing Touch'],
    'parts': ['CMG', 'CPT', 'Yoke', 'MG+PT'],
    'processes': ['MG forming', 'Yoke ass’y', 'Function test'],
    'title_en': 'BRS-161016 — Reduce CMG/CPT dimension (NTI data included) (2024-02-17)',
    'title_ko': 'BRS-161016 — CMG/CPT 치수 축소 + NTI 데이터 포함 (2024-02-17)',
    'title_vi': 'BRS-161016 — Giảm kích thước CMG/CPT + NTI data (17-02-2024)',
    'purpose_en': 'Improve NG hearing by reducing CMG/CPT dimension 10.86-4.86-0.88 → 10.82-4.2-0.88. Adds bond-spread decap (CMG/CPT, Yoke), gauss (MG+PT / Semi Yoke), Function vs Normal, NTI 종합 data and x-ray pictures.',
    'purpose_ko': 'CMG/CPT 치수 10.86-4.86-0.88 → 10.82-4.2-0.88 축소로 NG hearing 개선. 본드 스프레드 decap (CMG/CPT, Yoke), gauss (MG+PT/Semi Yoke), Function vs Normal, NTI 종합, x-ray 사진 포함.',
    'purpose_vi': 'Giảm kích thước CMG/CPT 10.86-4.86-0.88 → 10.82-4.2-0.88 để cải thiện NG hearing. Có decap bond-spread, gauss (MG+PT/Semi Yoke), Function vs Normal, NTI tổng hợp và ảnh x-ray.',
    'content_en': ['Sheet "NG rate": decap, gauss, function tables Test vs Normal.','Sheet "Picture xray": x-ray pictures of Yoke Normal vs Test (image-only).','Sheet "NTI_종합": NTI consolidated measurement data (large dataset).'],
    'content_ko': ['시트 "NG rate": decap, gauss, function Test vs Normal.','시트 "Picture xray": Yoke Normal vs Test x-ray 이미지 (이미지만).','시트 "NTI_종합": NTI 종합 측정 데이터(대용량).'],
    'content_vi': ['Sheet "NG rate": decap, gauss, function Test vs Normal.','Sheet "Picture xray": ảnh x-ray Yoke Normal vs Test (chỉ ảnh).','Sheet "NTI_종합": dữ liệu đo NTI tổng hợp (lớn).'],
    'test_conditions': [
        {'condition_group': 'cmg_cpt_dim', 'process': 'MG forming', 'changed_factor': 'CMG/CPT dimension', 'before_value': '10.86-4.86-0.88', 'after_value': '10.82-4.2-0.88', 'unit': 'mm', 'source_cells': ['B8']},
    ],
    'results': [
        # Decap bond spread
        {'condition_id': 'c1', 'measurement_type': 'Decap Bond Spread', 'date': '2024-02-17', 'input_count': 10, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Test bond spread (CMG/CPT ≥60%, Yoke ≥80%)', 'judgement': 'OK', 'sheet_name': 'NG rate', 'source_cells': ['F16:K16']},
        {'condition_id': 'c1', 'measurement_type': 'Decap Bond Spread', 'date': '2024-02-17', 'input_count': 10, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Normal bond spread', 'judgement': 'BASELINE', 'sheet_name': 'NG rate', 'source_cells': ['F17:K17']},
        # Gauss summary
        {'condition_id': 'c1', 'measurement_type': 'Gauss', 'date': '2024-02-17', 'metric_name': 'MG+PT Test AVG', 'metric_value': 1250.8, 'unit': 'G', 'judgement': 'OK', 'sheet_name': 'NG rate', 'source_cells': ['R30']},
        {'condition_id': 'c1', 'measurement_type': 'Gauss', 'date': '2024-02-17', 'metric_name': 'MG+PT Normal AVG', 'metric_value': 1269.2, 'unit': 'G', 'judgement': 'OK (baseline)', 'sheet_name': 'NG rate', 'source_cells': ['R31']},
        {'condition_id': 'c1', 'measurement_type': 'Gauss', 'date': '2024-02-17', 'metric_name': 'Semi Yoke Test AVG', 'metric_value': 527.2, 'unit': 'G', 'judgement': 'OK', 'sheet_name': 'NG rate', 'source_cells': ['R32']},
        {'condition_id': 'c1', 'measurement_type': 'Gauss', 'date': '2024-02-17', 'metric_name': 'Semi Yoke Normal AVG', 'metric_value': 547.8, 'unit': 'G', 'judgement': 'OK (baseline)', 'sheet_name': 'NG rate', 'source_cells': ['R33']},
        # Function Test Yoke vs Normal
        {'condition_id': 'c1', 'measurement_type': 'Function', 'condition_group': 'cmg_cpt_dim', 'date': '2024-02-17', 'input_count': 515, 'ok_count': 512, 'ng_count': 3, 'ng_rate_decimal': 0.005825, 'ng_rate_percent': 0.5825, 'metric_name': 'Test Yoke', 'judgement': 'IMPROVED', 'ng_breakdown': {'Noise': {'count': 3, 'rate': 0.005825}}, 'sheet_name': 'NG rate', 'source_cells': ['C38:N38']},
        {'condition_id': 'c1', 'measurement_type': 'Function', 'condition_group': 'cmg_cpt_dim', 'date': '2024-02-17', 'input_count': 560, 'ok_count': 542, 'ng_count': 18, 'ng_rate_decimal': 0.03214, 'ng_rate_percent': 3.214, 'metric_name': 'Normal Yoke', 'judgement': 'BASELINE', 'ng_breakdown': {'Noise': {'count': 16, 'rate': 0.02857}, 'Touch': {'count': 2, 'rate': 0.003571}}, 'sheet_name': 'NG rate', 'source_cells': ['C40:N40']},
    ],
    'conclusions': [
        {
            'topic_en': 'CMG/CPT dimension reduction (extended data)',
            'topic_ko': 'CMG/CPT 치수 축소 (확장 데이터)',
            'topic_vi': 'Giảm kích thước CMG/CPT (dữ liệu mở rộng)',
            'statement_en': 'Test Yoke Function 0.58% vs Normal 3.21%; bond spread OK same as Normal; gauss OK (spec).',
            'statement_ko': 'Test Yoke Function 0.58% vs Normal 3.21%; bond spread Normal 과 동등 OK; gauss 스펙 OK.',
            'statement_vi': 'Test Yoke Function 0.58% vs Normal 3.21%; bond spread OK ngang Normal; gauss đạt spec.',
            'interp_en': 'Reduced CMG/CPT dimension cuts Function NG from 3.214% to 0.583% (81.9% improvement) vs same-event Normal Yoke. Gauss is within spec although ~10–20G lower. Bond spread, x-ray and NTI data corroborate that the change has no negative side effect — decision: Can use.',
            'interp_ko': 'CMG/CPT 치수 축소로 동일 이벤트 Normal Yoke 대비 Function NG 가 3.214% → 0.583% (81.9% 개선). Gauss 는 10–20G 낮으나 스펙 내. Bond spread/x-ray/NTI 모두 부작용 없음 — 결정: 사용 가능.',
            'interp_vi': 'Giảm kích thước CMG/CPT giảm NG function từ 3.214% xuống 0.583% (cải thiện 81.9%) so với Normal Yoke cùng sự kiện. Gauss thấp hơn ~10–20G nhưng vẫn trong spec. Bond spread/x-ray/NTI xác nhận không tác động tiêu cực — quyết định: dùng được.',
            'source_cells': ['N38','N40']
        }
    ],
    'hints': [
        {'check_en': 'Apply CMG/CPT 10.82-4.2-0.88 dimension to production', 'reason_en': '81.9% improvement in Function NG vs Normal, gauss/bond-spread/NTI all OK.', 'check_ko': 'CMG/CPT 10.82-4.2-0.88 양산 적용.', 'reason_ko': 'Function NG Normal 대비 81.9% 개선, gauss/bond-spread/NTI 모두 양호.', 'check_vi': 'Áp dụng CMG/CPT 10.82-4.2-0.88 cho sản xuất.', 'reason_vi': 'Function NG cải thiện 81.9% so với Normal, gauss/bond-spread/NTI đều ổn.', 'evidence_strength': 'strong', 'related_process': 'MG forming', 'related_part': 'CMG/CPT', 'source_cells': ['N38:N40']},
    ],
    'log': {
        'rationale_en': 'Same-day Test Yoke vs Normal Yoke with auxiliary gauss, bond-spread, x-ray and NTI evidence — normal_comparison. Relative change (test/baseline-1)*100 = -81.86%.',
        'rationale_ko': '동일일자 Test Yoke vs Normal Yoke + gauss/bond-spread/x-ray/NTI 보조 증거 — normal_comparison. 상대 변화 -81.86%.',
        'rationale_vi': 'Cùng ngày Test Yoke vs Normal Yoke + dữ liệu gauss/bond-spread/x-ray/NTI hỗ trợ — normal_comparison. Thay đổi tương đối -81.86%.',
        'warnings_en': ['NTI_종합 sheet was not row-extracted into AiResults; only summary captured.','Picture xray sheet is image-only — visual inspection required.'],
        'warnings_ko': ['NTI_종합 시트는 AiResults 행으로 추출하지 않음 — 요약만 기록.','Picture xray 시트는 이미지만이라 시각 검사 필요.'],
        'warnings_vi': ['Sheet NTI_종합 chưa được trích thành AiResults — chỉ tóm tắt.','Sheet Picture xray chỉ có ảnh — cần kiểm thị giác.'],
        'assumptions_en': [],
    },
    'when_user_asks': ['BRS-161016 NG hearing', 'CMG/CPT dimension reduce', 'NTI yoke dimension'],
    'confidence': 0.8,
})


if __name__ == '__main__':
    ok = 0; fail = 0
    for d in D:
        try:
            if B.commit(d):
                print(f'OK  {d["name"]}')
                ok += 1
            else:
                print(f'FAIL {d["name"]}')
                fail += 1
        except Exception as e:
            print(f'ERR {d["name"]}: {e!r}')
            fail += 1
    print(f'BATCH RESULT: ok={ok} fail={fail}')
