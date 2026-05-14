"""Commit a batch of dataset analyses. Each dataset is a compact dict; the
builder fills in defaults so we keep the script terse."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _batch_build as B

DATASETS = [
    # ---------------- dataset 2 ----------------
    {
        'name': '18. TIU C11-20  Report test VP finf reason NG funtion 2026.1.15',
        'report_type': 'normal_comparison',
        'model': 'TIU C11-20',
        'report_date': '2026-01-15',
        'department': 'ME',
        'marker': 'Thao',
        'default_sheet': 'Test',
        'primary_defect': 'NG Function (Hearing RB / Noise)',
        'primary_aliases': ['NG hearing', 'NG function high'],
        'related_defects': ['RB', 'Noise', 'Hearing Touch'],
        'parts': ['VP', 'Frame', 'Coil'],
        'processes': ['VP heat-press', 'AWF', 'Sub3 Ass’y VP+Frame'],
        'title_en': 'TIU C11-20 Report — Test VP find reason NG function (2026-01-15)',
        'title_ko': 'TIU C11-20 리포트 — VP NG Function 원인 검토 (2026-01-15)',
        'title_vi': 'Báo cáo TIU C11-20 — Test VP tìm nguyên nhân NG function (15-01-2026)',
        'purpose_en': 'Identify the root cause of high NG-function (mainly RB/Noise) on TIU C11-20 by varying VP heat-press temperature/pressure, comparing VP lot dated 6/1, comparing line A2/D3 transfer flow and comparing AWF machines #1–4 against the same-event Normal sample.',
        'purpose_ko': 'TIU C11-20 의 NG Function (RB/Noise) 원인 파악을 위해 VP heat-press 온도/압력 변경, 6/1 lot VP 비교, A2/D3 라인 이동 검토, AWF #1~4 머신별 비교를 동일 이벤트의 Normal 대조군과 진행한다.',
        'purpose_vi': 'Tìm nguyên nhân NG function (RB/Noise) cao trên TIU C11-20 bằng cách thay đổi nhiệt độ/áp suất VP heat-press, so sánh lot VP 6/1, so sánh chuyển dòng A2/D3 và so sánh máy AWF #1–4 với mẫu Normal cùng sự kiện.',
        'content_en': [
            'Table 1 — VP heat-press temperature/pressure types (195°C-2.8Kf / 190°C-2.8Kf / 190°C-2Kf) vs Normal.',
            'Table 2 — VP lot date 6/1 vs Normal on L and R lines (same event).',
            'Table 3 — Sub1 D3→A2 movement and L vs R make-line comparison.',
            'Table 4 — AWF machines #1–4 vs Normal.',
            'Table 5 — Frame/VP, VP laser-cut and VP semi-OK transferred from A2 to D3 vs Normal.'
        ],
        'content_ko': [
            '표1 — VP heat-press 온도/압력 (195°C-2.8Kf / 190°C-2.8Kf / 190°C-2Kf) vs Normal.',
            '표2 — VP 6/1 lot vs Normal (L/R 라인, 동일 이벤트).',
            '표3 — Sub1 D3→A2 이동 및 L vs R Make line 비교.',
            '표4 — AWF #1~4 vs Normal.',
            '표5 — Frame/VP, VP laser-cut, VP semi-OK 의 A2→D3 이동 vs Normal.'
        ],
        'content_vi': [
            'Bảng 1 — Nhiệt độ/áp suất VP heat-press (195°C-2.8Kf / 190°C-2.8Kf / 190°C-2Kf) vs Normal.',
            'Bảng 2 — VP lot ngày 6/1 vs Normal trên line L và R (cùng sự kiện).',
            'Bảng 3 — Di chuyển Sub1 D3→A2 và so sánh Make line L vs R.',
            'Bảng 4 — Máy AWF #1–4 vs Normal.',
            'Bảng 5 — Frame/VP, VP laser-cut, VP semi-OK chuyển A2→D3 vs Normal.'
        ],
        'test_conditions': [
            {'condition_group': 'vp_heat_press', 'process': 'VP heat-press', 'changed_factor': 'VP heat-press temperature/pressure', 'before_value': '195°C-2.8Kf', 'after_value': '190°C-2Kf', 'unit': '°C / Kf', 'source_cells': ['B9:B11']},
            {'condition_group': 'vp_lot_date', 'process': 'VP lot', 'changed_factor': 'VP lot date 6/1', 'before_value': 'Normal lot', 'after_value': '6/1 lot', 'source_cells': ['F33','F35','F37','F39']},
            {'condition_group': 'awf_machine', 'process': 'AWF', 'changed_factor': 'AWF machine #', 'before_value': 'Normal', 'after_value': '#1..#4', 'source_cells': ['F55','F57','F59','F61','F63']},
            {'condition_group': 'a2_d3_move', 'process': 'Sub3 ass’y VP+Frame', 'changed_factor': 'Move A2 → D3 semi step', 'before_value': 'Normal', 'after_value': 'A2→D3 move at Frame/VP, laser-cut, semi-OK', 'source_cells': ['F69','F71','F73','F75']},
        ],
        'results': [
            # Table 1 — 195C/190C-2.8Kf/190C-2Kf vs Normal 14%
            {'condition_id': 'c1', 'measurement_type': 'Function', 'condition_group': 'vp_heat_press', 'date': '2026-01-15', 'line': 'TIU C11-20L', 'input_count': 47, 'ok_count': 37, 'ng_count': 10, 'ng_rate_decimal': 0.2128, 'ng_rate_percent': 21.28, 'metric_name': 'Type 1 195°C-2.8Kf', 'judgement': 'WORSE', 'ng_breakdown': {'RB': {'count': 10, 'rate': 0.2128}, 'Noise': {'count': 10, 'rate': 0.2128}}, 'source_cells': ['C21:R21']},
            {'condition_id': 'c1', 'measurement_type': 'Function', 'condition_group': 'vp_heat_press', 'date': '2026-01-15', 'line': 'TIU C11-20L', 'input_count': 42, 'ok_count': 40, 'ng_count': 2, 'ng_rate_decimal': 0.04762, 'ng_rate_percent': 4.762, 'metric_name': 'Type 2 190°C-2.8Kf', 'judgement': 'IMPROVED', 'ng_breakdown': {'RB': {'count': 3, 'rate': 0.0714}, 'Noise': {'count': 2, 'rate': 0.04762}}, 'source_cells': ['C23:R23']},
            {'condition_id': 'c1', 'measurement_type': 'Function', 'condition_group': 'vp_heat_press', 'date': '2026-01-15', 'line': 'TIU C11-20L', 'input_count': 43, 'ok_count': 41, 'ng_count': 2, 'ng_rate_decimal': 0.04651, 'ng_rate_percent': 4.651, 'metric_name': 'Type 3 190°C-2Kf', 'judgement': 'IMPROVED', 'ng_breakdown': {'RB': {'count': 4, 'rate': 0.0930}, 'Noise': {'count': 2, 'rate': 0.04651}}, 'source_cells': ['C25:R25']},
            {'condition_id': 'c1', 'measurement_type': 'Function', 'condition_group': 'vp_heat_press', 'date': '2026-01-15', 'line': 'TIU C11-20L', 'input_count': 50, 'ok_count': 43, 'ng_count': 7, 'ng_rate_decimal': 0.14, 'ng_rate_percent': 14.0, 'metric_name': 'Normal', 'judgement': 'BASELINE', 'ng_breakdown': {'RB': {'count': 7, 'rate': 0.14}, 'Noise': {'count': 7, 'rate': 0.14}}, 'source_cells': ['C27:R27']},
            # Table 2 — Lot 6/1 vs Normal (L line, R line)
            {'condition_id': 'c2', 'measurement_type': 'Function', 'condition_group': 'vp_lot_date', 'date': '2026-01-15', 'line': 'TIU C11-20L', 'input_count': 410, 'ok_count': 362, 'ng_count': 48, 'ng_rate_decimal': 0.1171, 'ng_rate_percent': 11.71, 'metric_name': 'Lot VP 6/1 (L line)', 'judgement': 'IMPROVED', 'ng_breakdown': {'SPL': {'count': 6, 'rate': 0.01463}, 'RB': {'count': 50, 'rate': 0.122}, 'Noise': {'count': 38, 'rate': 0.0927}}, 'source_cells': ['C33:R33']},
            {'condition_id': 'c2', 'measurement_type': 'Function', 'condition_group': 'vp_lot_date', 'date': '2026-01-15', 'line': 'TIU C11-20L', 'input_count': 410, 'ok_count': 324, 'ng_count': 86, 'ng_rate_decimal': 0.20976, 'ng_rate_percent': 20.98, 'metric_name': 'Normal (L line)', 'judgement': 'BASELINE', 'ng_breakdown': {'SPL': {'count': 4, 'rate': 0.00976}, 'RB': {'count': 108, 'rate': 0.2634}, 'Noise': {'count': 75, 'rate': 0.1829}}, 'source_cells': ['C35:R35']},
            {'condition_id': 'c2', 'measurement_type': 'Function', 'condition_group': 'vp_lot_date', 'date': '2026-01-15', 'line': 'TIU C11-20R', 'input_count': 444, 'ok_count': 417, 'ng_count': 27, 'ng_rate_decimal': 0.0608, 'ng_rate_percent': 6.08, 'metric_name': 'Lot VP 6/1 (R line)', 'judgement': 'IMPROVED', 'ng_breakdown': {'RB': {'count': 29, 'rate': 0.0653}, 'Noise': {'count': 26, 'rate': 0.0586}}, 'source_cells': ['C37:R37']},
            {'condition_id': 'c2', 'measurement_type': 'Function', 'condition_group': 'vp_lot_date', 'date': '2026-01-15', 'line': 'TIU C11-20R', 'input_count': 516, 'ok_count': 446, 'ng_count': 70, 'ng_rate_decimal': 0.1357, 'ng_rate_percent': 13.57, 'metric_name': 'Normal (R line)', 'judgement': 'BASELINE', 'ng_breakdown': {'RB': {'count': 92, 'rate': 0.1783}, 'Noise': {'count': 69, 'rate': 0.1337}}, 'source_cells': ['C39:R39']},
            # Table 4 — AWF #1..#4 vs Normal
            {'condition_id': 'c3', 'measurement_type': 'Function', 'condition_group': 'awf_machine', 'date': '2026-01-15', 'line': 'TIU C11-20L', 'input_count': 57, 'ok_count': 48, 'ng_count': 9, 'ng_rate_decimal': 0.1579, 'ng_rate_percent': 15.79, 'metric_name': 'AWF #1', 'judgement': 'WORSE', 'source_cells': ['C55:R55']},
            {'condition_id': 'c3', 'measurement_type': 'Function', 'condition_group': 'awf_machine', 'date': '2026-01-15', 'line': 'TIU C11-20L', 'input_count': 57, 'ok_count': 49, 'ng_count': 8, 'ng_rate_decimal': 0.1404, 'ng_rate_percent': 14.04, 'metric_name': 'AWF #2', 'judgement': 'SIMILAR', 'source_cells': ['C57:R57']},
            {'condition_id': 'c3', 'measurement_type': 'Function', 'condition_group': 'awf_machine', 'date': '2026-01-15', 'line': 'TIU C11-20L', 'input_count': 50, 'ok_count': 42, 'ng_count': 8, 'ng_rate_decimal': 0.16, 'ng_rate_percent': 16.0, 'metric_name': 'AWF #3', 'judgement': 'WORSE', 'source_cells': ['C59:R59']},
            {'condition_id': 'c3', 'measurement_type': 'Function', 'condition_group': 'awf_machine', 'date': '2026-01-15', 'line': 'TIU C11-20L', 'input_count': 70, 'ok_count': 63, 'ng_count': 7, 'ng_rate_decimal': 0.10, 'ng_rate_percent': 10.0, 'metric_name': 'AWF #4', 'judgement': 'IMPROVED', 'source_cells': ['C61:R61']},
            {'condition_id': 'c3', 'measurement_type': 'Function', 'condition_group': 'awf_machine', 'date': '2026-01-15', 'line': 'TIU C11-20L', 'input_count': 50, 'ok_count': 43, 'ng_count': 7, 'ng_rate_decimal': 0.14, 'ng_rate_percent': 14.0, 'metric_name': 'Normal', 'judgement': 'BASELINE', 'source_cells': ['C63:R63']},
            # Table 5 — A2→D3 move
            {'condition_id': 'c4', 'measurement_type': 'Function', 'condition_group': 'a2_d3_move', 'date': '2026-01-15', 'line': 'TIU C11-20L (Night)', 'input_count': 50, 'ok_count': 48, 'ng_count': 2, 'ng_rate_decimal': 0.04, 'ng_rate_percent': 4.0, 'metric_name': 'Semi Frame/VP A2→D3 (ass’y yoke)', 'judgement': 'IMPROVED', 'source_cells': ['C69:R69']},
            {'condition_id': 'c4', 'measurement_type': 'Function', 'condition_group': 'a2_d3_move', 'date': '2026-01-15', 'line': 'TIU C11-20L (Night)', 'input_count': 200, 'ok_count': 177, 'ng_count': 23, 'ng_rate_decimal': 0.115, 'ng_rate_percent': 11.5, 'metric_name': 'VP laser-cut A2→D3', 'judgement': 'IMPROVED', 'source_cells': ['C71:R71']},
            {'condition_id': 'c4', 'measurement_type': 'Function', 'condition_group': 'a2_d3_move', 'date': '2026-01-15', 'line': 'TIU C11-20L (Night)', 'input_count': 187, 'ok_count': 144, 'ng_count': 43, 'ng_rate_decimal': 0.2299, 'ng_rate_percent': 22.99, 'metric_name': 'VP Semi OK A2→D3', 'judgement': 'WORSE', 'source_cells': ['C73:R73']},
            {'condition_id': 'c4', 'measurement_type': 'Function', 'condition_group': 'a2_d3_move', 'date': '2026-01-15', 'line': 'TIU C11-20L (Night)', 'input_count': 200, 'ok_count': 159, 'ng_count': 41, 'ng_rate_decimal': 0.205, 'ng_rate_percent': 20.5, 'metric_name': 'Normal', 'judgement': 'BASELINE', 'source_cells': ['C75:R75']},
        ],
        'conclusions': [
            {
                'topic_en': 'VP heat-press temperature/pressure',
                'topic_ko': 'VP heat-press 온도/압력',
                'topic_vi': 'Nhiệt độ/áp suất VP heat-press',
                'statement_en': 'Type1 195°C-2.8Kf 21.3%, Type2 190°C-2.8Kf 4.8%, Type3 190°C-2Kf 4.7%, Normal 14.0%.',
                'statement_ko': 'Type1 195°C-2.8Kf 21.3%, Type2 190°C-2.8Kf 4.8%, Type3 190°C-2Kf 4.7%, Normal 14.0%.',
                'statement_vi': 'Type1 195°C-2.8Kf 21.3%, Type2 190°C-2.8Kf 4.8%, Type3 190°C-2Kf 4.7%, Normal 14.0%.',
                'interp_en': 'Same-event Normal NG=14.0%. Type1 (195°C/2.8Kf) is 52.0% worse than Normal; Type2 (190°C/2.8Kf) is 66.0% improved; Type3 (190°C/2Kf) is 66.8% improved. Lowering temperature from 195°C to 190°C reduces NG function — pressure change (2.8→2Kf) gives only marginal further effect.',
                'interp_ko': '동일 이벤트 Normal NG=14.0%. Type1 (195°C/2.8Kf) 은 Normal 대비 52.0% 악화, Type2 (190°C/2.8Kf) 66.0% 개선, Type3 (190°C/2Kf) 66.8% 개선. 195°C → 190°C 로 온도를 낮춘 것이 NG Function 감소의 주요인이며 압력(2.8→2Kf) 차이는 미미.',
                'interp_vi': 'Normal cùng sự kiện NG=14.0%. Type1 (195°C/2.8Kf) tệ hơn Normal 52.0%; Type2 (190°C/2.8Kf) cải thiện 66.0%; Type3 (190°C/2Kf) cải thiện 66.8%. Hạ nhiệt từ 195°C xuống 190°C là yếu tố chính giảm NG function; thay đổi áp suất (2.8→2Kf) chỉ tạo khác biệt nhỏ.',
                'source_cells': ['R21','R23','R25','R27']
            },
            {
                'topic_en': 'VP lot 6/1 vs Normal',
                'topic_ko': 'VP 6/1 lot vs Normal',
                'topic_vi': 'Lot VP 6/1 vs Normal',
                'statement_en': 'L line: lot 6/1 11.71% vs Normal 20.98%. R line: lot 6/1 6.08% vs Normal 13.57%.',
                'statement_ko': 'L 라인: 6/1 lot 11.71% vs Normal 20.98%. R 라인: 6/1 lot 6.08% vs Normal 13.57%.',
                'statement_vi': 'Line L: lot 6/1 11.71% vs Normal 20.98%. Line R: lot 6/1 6.08% vs Normal 13.57%.',
                'interp_en': 'L line lot-6/1 is 44.2% improved vs Normal; R line lot-6/1 is 55.2% improved vs Normal. VP lot dated 6/1 gives a clear function improvement on both lines — VP lot quality is a primary driver of NG function.',
                'interp_ko': 'L 라인 6/1 lot 은 Normal 대비 44.2% 개선, R 라인은 55.2% 개선. VP lot 6/1 은 양 라인에서 Function 개선이 명확 — VP lot 품질이 NG Function 의 주요 요인.',
                'interp_vi': 'Line L lot 6/1 cải thiện 44.2% so với Normal; Line R cải thiện 55.2%. Lot VP 6/1 cải thiện rõ rệt cả hai line — chất lượng lot VP là yếu tố chính của NG function.',
                'source_cells': ['R33','R35','R37','R39']
            },
            {
                'topic_en': 'AWF machine comparison',
                'topic_ko': 'AWF 머신 비교',
                'topic_vi': 'So sánh máy AWF',
                'statement_en': 'AWF #1 15.79%, #2 14.04%, #3 16.0%, #4 10.0%, Normal 14.0%.',
                'statement_ko': 'AWF #1 15.79%, #2 14.04%, #3 16.0%, #4 10.0%, Normal 14.0%.',
                'statement_vi': 'AWF #1 15.79%, #2 14.04%, #3 16.0%, #4 10.0%, Normal 14.0%.',
                'interp_en': 'AWF #4 is 28.6% improved vs same-event Normal; #1/#3 are 12.8%/14.3% worse; #2 is similar (+0.3%). AWF #4 should be the preferred machine.',
                'interp_ko': 'AWF #4 는 동일 이벤트 Normal 대비 28.6% 개선, #1/#3 은 12.8%/14.3% 악화, #2 는 유사(+0.3%). AWF #4 사용 권장.',
                'interp_vi': 'AWF #4 cải thiện 28.6% so với Normal cùng sự kiện; #1/#3 tệ hơn 12.8%/14.3%; #2 tương đương (+0.3%). Nên ưu tiên dùng AWF #4.',
                'source_cells': ['R55','R57','R59','R61','R63']
            },
        ],
        'hints': [
            {'check_en': 'Lower VP heat-press temperature to 190°C', 'reason_en': 'Type2/3 at 190°C show 66% improvement vs Normal 14%; 195°C is 52% worse.', 'check_ko': 'VP heat-press 온도를 190°C 로 낮춘다.', 'reason_ko': 'Type2/3 (190°C) 는 Normal 대비 66% 개선, 195°C 는 52% 악화.', 'check_vi': 'Hạ nhiệt VP heat-press xuống 190°C.', 'reason_vi': 'Type2/3 ở 190°C cải thiện 66% so với Normal 14%; 195°C tệ hơn 52%.', 'evidence_strength': 'strong', 'related_process': 'VP heat-press', 'related_part': 'VP', 'source_cells': ['R21:R27']},
            {'check_en': 'Use VP lot dated 6/1 or replicate its conditions', 'reason_en': 'Lot 6/1 improves L line by 44.2% and R line by 55.2% vs Normal.', 'check_ko': 'VP lot 6/1 사용 또는 동일 조건 재현.', 'reason_ko': 'lot 6/1 은 L 라인 44.2%, R 라인 55.2% 개선.', 'check_vi': 'Sử dụng lot VP 6/1 hoặc tái hiện điều kiện của nó.', 'reason_vi': 'Lot 6/1 cải thiện 44.2% trên line L và 55.2% trên line R.', 'evidence_strength': 'strong', 'related_process': 'VP lot', 'related_part': 'VP', 'source_cells': ['R33:R39']},
            {'check_en': 'Prioritise AWF #4 over #1/#3', 'reason_en': 'AWF #4 is 28.6% improved vs Normal; #1/#3 are 12.8%/14.3% worse.', 'check_ko': 'AWF #4 우선 사용, #1/#3 회피.', 'reason_ko': 'AWF #4 는 Normal 대비 28.6% 개선, #1/#3 는 12.8%/14.3% 악화.', 'check_vi': 'Ưu tiên dùng AWF #4 hơn #1/#3.', 'reason_vi': 'AWF #4 cải thiện 28.6% so với Normal; #1/#3 tệ hơn 12.8%/14.3%.', 'evidence_strength': 'medium', 'related_process': 'AWF', 'related_part': 'Coil/SP', 'source_cells': ['R55:R63']},
        ],
        'log': {
            'rationale_en': 'Each Test block has a same-sheet Normal row; classified as normal_comparison. Relative change computed as (test/baseline - 1)*100. Lower VP heat-press temperature + VP lot 6/1 + AWF #4 all clearly reduce NG function vs same-event Normal.',
            'rationale_ko': '각 테스트 블록에 동일 시트 Normal 행이 있어 normal_comparison 으로 분류. 상대 변화율은 (test/baseline-1)*100. VP heat-press 온도 하향, VP lot 6/1, AWF #4 모두 동일 이벤트 Normal 대비 NG Function 명확히 감소.',
            'rationale_vi': 'Mỗi khối Test có hàng Normal trên cùng sheet; phân loại normal_comparison. Thay đổi tương đối tính (test/baseline-1)*100. Hạ nhiệt VP heat-press, lot VP 6/1, máy AWF #4 đều giảm rõ rệt NG function so với Normal cùng sự kiện.',
            'warnings_en': ['Table 3 (Sub1 D3→A2 etc.) and Table 5 transfer flow rows do not have a clean Normal pair inside the same sub-table — kept as descriptive but no relative-change claim.'],
            'warnings_ko': ['표3 (Sub1 D3→A2) 와 표5 일부 행은 동일 서브-표 내 Normal 짝이 명확하지 않아 상대 변화 주장 보류.'],
            'warnings_vi': ['Bảng 3 (Sub1 D3→A2) và một số dòng Bảng 5 không có Normal cùng cụm rõ ràng — chỉ mô tả, không khẳng định thay đổi tương đối.'],
            'assumptions_en': ['Treated the row labelled "Normal" in each result table as that table’s baseline.'],
            'assumptions_ko': ['각 결과 표 내 "Normal" 라벨 행을 해당 표의 baseline 으로 사용.'],
            'assumptions_vi': ['Coi dòng có nhãn "Normal" trong mỗi bảng là baseline của bảng đó.']
        },
        'when_user_asks': ['TIU C11-20 NG function high', 'VP heat-press temperature', 'AWF machine NG'],
        'confidence': 0.78,
    },
    # ---------------- dataset 3 ----------------
    {
        'name': '18. TIU L5S3-01 R Report test Frame improve separate S-MG  date 2025.12.02',
        'report_type': 'process_condition_change',
        'model': 'TIU L5S3-01 R',
        'report_date': '2025-12-02',
        'department': 'ME',
        'marker': 'Nhung',
        'default_sheet': 'Test',
        'primary_defect': 'Separate Side MG',
        'primary_aliases': ['NG S-MG separate', 'Frame separate side MG'],
        'related_defects': [],
        'parts': ['Frame', 'Side MG'],
        'processes': ['Sub3 tracking', 'Magnetization', 'Ass’y frame+vp', 'Ass’y frame+F-PCB'],
        'title_en': 'TIU L5S3-01 [R] — Test Frame new mold to improve Separate Side MG (2025-12-02)',
        'title_ko': 'TIU L5S3-01 [R] — Frame 신규 금형으로 Side MG separate 개선 검증 (2025-12-02)',
        'title_vi': 'TIU L5S3-01 [R] — Thử Frame khuôn mới cải thiện Separate Side MG (02-12-2025)',
        'purpose_en': 'Frame material had Separate Side MG defect (Frame L 1.7%, Frame R 3.75%). Verify whether the new Frame mold removes Separate Side MG through Sub3 tracking from MTR visual to Eject.',
        'purpose_ko': 'Frame 자재에서 Side MG separate 발생 (Frame L 1.7%, Frame R 3.75%). 신규 Frame 금형이 Side MG separate 를 제거하는지 Sub3 tracking (MTR visual → Eject) 으로 검증.',
        'purpose_vi': 'Frame có lỗi Separate Side MG (Frame L 1.7%, Frame R 3.75%). Kiểm tra khuôn Frame mới có loại bỏ Separate Side MG không qua tracking Sub3 (MTR visual → Eject).',
        'content_en': ['Sub3 tracking: MTR visual → Frame loading → Magnetization → Ass’y frame+vp → Reverse check → Ass’y frame+F-PCB → Eject check.','Input 50 per process step.'],
        'content_ko': ['Sub3 tracking: MTR visual → Frame loading → Magnetization → Ass’y frame+vp → Reverse → Ass’y frame+F-PCB → Eject.','각 공정 단계 Input 50.'],
        'content_vi': ['Tracking Sub3: MTR visual → Frame loading → Magnetization → Ass’y frame+vp → Reverse → Ass’y frame+F-PCB → Eject.','Mỗi bước Input 50.'],
        'test_conditions': [
            {'condition_group': 'frame_mold', 'process': 'Frame mold', 'changed_factor': 'Frame mold (new vs current)', 'before_value': 'Current Frame', 'after_value': 'New Frame mold', 'source_cells': ['B6']},
        ],
        'results': [
            {'condition_id': 'c1', 'measurement_type': 'Tracking', 'date': '2025-12-02', 'input_count': 50, 'ok_count': 50, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'MTR visual', 'judgement': 'OK', 'source_cells': ['C15:I15']},
            {'condition_id': 'c1', 'measurement_type': 'Tracking', 'date': '2025-12-02', 'input_count': 50, 'ok_count': 50, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Frame loading', 'judgement': 'OK', 'source_cells': ['C16:I16']},
            {'condition_id': 'c1', 'measurement_type': 'Tracking', 'date': '2025-12-02', 'input_count': 50, 'ok_count': 50, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Magnetization', 'judgement': 'OK', 'source_cells': ['C17:I17']},
            {'condition_id': 'c1', 'measurement_type': 'Tracking', 'date': '2025-12-02', 'input_count': 50, 'ok_count': 50, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Ass’y frame+vp', 'judgement': 'OK', 'source_cells': ['C18:I18']},
            {'condition_id': 'c1', 'measurement_type': 'Tracking', 'date': '2025-12-02', 'input_count': 50, 'ok_count': 50, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Reverse check', 'judgement': 'OK', 'source_cells': ['C19:I19']},
            {'condition_id': 'c1', 'measurement_type': 'Tracking', 'date': '2025-12-02', 'input_count': 50, 'ok_count': 50, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Ass’y frame+F-PCB', 'judgement': 'OK', 'source_cells': ['C20:I20']},
            {'condition_id': 'c1', 'measurement_type': 'Tracking', 'date': '2025-12-02', 'input_count': 50, 'ok_count': 50, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0, 'metric_name': 'Eject check', 'judgement': 'OK', 'source_cells': ['C21:I21']},
        ],
        'conclusions': [
            {
                'topic_en': 'New Frame mold vs prior Frame baseline',
                'topic_ko': '신규 Frame 금형 vs 기존 Frame baseline',
                'topic_vi': 'Khuôn Frame mới vs Frame baseline',
                'statement_en': 'Sub3 tracking lot of improved Frame is OK at every step (0/50 NG). Decision: Can use.',
                'statement_ko': '개선 Frame Sub3 tracking 결과 모든 공정 0/50 NG. 결정: 사용 가능.',
                'statement_vi': 'Tracking Sub3 cho lot Frame cải tiến đạt 0/50 NG ở mọi bước. Quyết định: dùng được.',
                'interp_en': 'New Frame mold reduces Separate Side MG from Frame-L 1.7% / Frame-R 3.75% baseline to 0% across all 7 tracking steps — full elimination on a 50-piece sample per step (cross-event baseline, not same-table).',
                'interp_ko': '신규 Frame 금형은 Frame-L 1.7%/Frame-R 3.75% baseline 을 7개 tracking 공정 모두 0% 로 낮춤 — 각 50개 샘플 기준 완전 제거 (baseline 은 동일 표가 아닌 이전 이벤트).',
                'interp_vi': 'Khuôn Frame mới giảm Separate Side MG từ baseline Frame-L 1.7%/Frame-R 3.75% xuống 0% ở cả 7 bước tracking — loại bỏ hoàn toàn trên 50 mẫu mỗi bước (baseline lấy từ sự kiện trước, không cùng bảng).',
                'source_cells': ['B6','I15:I21']
            }
        ],
        'hints': [
            {'check_en': 'Apply the new Frame mold change company-wide', 'reason_en': 'Sub3 tracking shows 0/50 NG at every step versus prior Frame L 1.7% / R 3.75%.', 'check_ko': '신규 Frame 금형 변경을 양산 전면 적용 검토.', 'reason_ko': 'Sub3 tracking 모든 공정 0/50 NG, 기존 Frame L 1.7%/R 3.75% 대비 완전 개선.', 'check_vi': 'Áp dụng khuôn Frame mới cho toàn dây.', 'reason_vi': 'Tracking Sub3 đạt 0/50 NG mọi bước so với Frame L 1.7%/R 3.75% trước đó.', 'evidence_strength': 'medium', 'related_process': 'Frame mold', 'related_part': 'Frame', 'source_cells': ['I15:I21']},
        ],
        'log': {
            'rationale_en': 'Classified as process_condition_change (Frame mold change). Baseline (Frame L 1.7%, R 3.75%) comes from prior production data referenced in Purpose, not same-table — relative change recorded but flagged as cross-event baseline.',
            'rationale_ko': 'Frame 금형 변경이므로 process_condition_change 로 분류. baseline (Frame L 1.7%, R 3.75%) 은 동일 표가 아닌 Purpose 의 이전 데이터 — cross-event baseline 로 표기.',
            'rationale_vi': 'Phân loại process_condition_change (đổi khuôn Frame). Baseline (Frame L 1.7%, R 3.75%) lấy từ dữ liệu trước đó nêu ở Purpose, không phải cùng bảng — đánh dấu baseline chéo sự kiện.',
            'warnings_en': ['Sample size 50/step is small; verify with larger lot before mass change.'],
            'warnings_ko': ['단계별 샘플 50개는 작음 — 양산 변경 전 대량 lot 추가 검증 필요.'],
            'warnings_vi': ['Cỡ mẫu 50/bước nhỏ — cần xác minh lô lớn trước khi đổi sản xuất.'],
            'assumptions_en': ['"Decision: OK => Can use" interpreted as approval to apply new Frame mold.'],
            'assumptions_ko': ['"Decision: OK => Can use" 를 신규 Frame 금형 적용 승인으로 해석.'],
            'assumptions_vi': ['"Decision: OK => Can use" được hiểu là chấp thuận khuôn Frame mới.']
        },
        'when_user_asks': ['Separate Side MG', 'Frame mold change', 'Side MG separate Frame'],
        'confidence': 0.72,
    },
]


if __name__ == '__main__':
    ok = 0; fail = 0
    for d in DATASETS:
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
