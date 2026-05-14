"""Run chunk 00 datasets through commit_dataset."""
from __future__ import annotations
import sys, io, json
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import _ai_batch_helper as h


def run(name: str, result: dict, tr_ko: dict, tr_en: dict, tr_vi: dict) -> None:
    ok = h.commit_dataset(name, result, tr_ko, tr_en, tr_vi)
    if ok:
        print(f'[OK {name}]')
    else:
        print(f'[FAIL {name} commit_dataset returned False]')


# ===== DS 00 =====
name00 = '25. BRS-161014  GMI Report test VP NG surface abnormal 12.16.2024'
result00 = {
    'schema_version': '0.1',
    'document': {
        'document_id': '', 'source_file': name00, 'source_sheet': '1',
        'title': 'REPORT TEST VP MATERIAL SURFACE ABNORMAL BRS-161016',
        'model': 'BRS-161016', 'report_date': '2024-12-16',
        'department': 'ME', 'marker': 'Huong', 'line': 'E2-3A',
        'report_type': 'normal_comparison',
        'primary_defect': {'canonical_name': 'VP Surface Abnormal', 'aliases_in_document': ['VP material suface abnormal']},
        'related_defects': ['NG Hearing Noise', 'VP+CD Separation', 'Damage VP', 'Particle', 'Cutting offset', 'Burr'],
        'parts': ['VP', 'CD'], 'processes': ['Sub 1 Vision', 'SPK Function', 'Module line', 'Reliability'],
        'purpose': 'Test VP material whose surface is abnormal to confirm if it can be used.',
        'content': [
            'Test SPK line and check vision Sub 1, function.',
            'Test reliability at A1: Tension VP+CD 5pcs, Load test 96h 5pcs, Temperature & humidity 5pcs.',
            'Module line test and check function.'
        ],
        'source_cells': {'title': ['1!B2'], 'date': ['1!P3'], 'purpose': ['1!A6:A7'], 'content': ['1!A8:A13']}
    },
    'test_conditions': [
        {'condition_id': 'cond_1', 'condition_group': 'VP surface abnormal vs Normal VP',
         'line': 'E2-3A', 'process': 'Sub 1 Vision', 'changed_factor': 'VP material surface state',
         'before_value': 'Normal VP', 'after_value': 'VP material surface abnormal',
         'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None,
         'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': name00, 'sheet_name': '1', 'source_cells': ['1!B19:N20']},
        {'condition_id': 'cond_2', 'condition_group': 'SPK function test',
         'line': 'E2-3A', 'process': 'SPK Function', 'changed_factor': 'VP material surface state',
         'before_value': 'Normal VP', 'after_value': 'VP material surface abnormal',
         'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None,
         'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': name00, 'sheet_name': '1', 'source_cells': ['1!B25:N26']},
        {'condition_id': 'cond_3', 'condition_group': 'Module line check',
         'line': 'B2-8A', 'process': 'Module Line', 'changed_factor': 'VP material surface state',
         'before_value': 'Normal', 'after_value': 'VP surface abnormal',
         'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None,
         'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': name00, 'sheet_name': '1', 'source_cells': ['1!B37:N38']},
        {'condition_id': 'cond_4', 'condition_group': 'Reliability',
         'line': 'A1', 'process': 'Reliability', 'changed_factor': 'VP surface abnormal sample',
         'before_value': None, 'after_value': None,
         'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None,
         'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': name00, 'sheet_name': '1', 'source_cells': ['1!B30:G32']},
    ],
    'results': [
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Vision Sub 1',
         'condition_group': 'VP surface abnormal', 'date': '2024-12-16', 'line': 'E2-3A',
         'input_count': 90, 'ok_count': 90, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Total NG rate Sub1', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'Cutting offset': 0, 'Burr': 0, 'Damage VP': 0, 'Particle': 0, 'VP+CD Separation': 0},
         'source_file': name00, 'sheet_name': '1', 'source_cells': ['1!E19:N19']},
        {'result_id': 'res_2', 'condition_id': 'cond_1', 'measurement_type': 'Vision Sub 1',
         'condition_group': 'Normal VP (baseline)', 'date': '2024-12-16', 'line': 'E2-3A',
         'input_count': 100, 'ok_count': 100, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Total NG rate Sub1', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'Cutting offset': 0, 'Burr': 0, 'Damage VP': 0, 'Particle': 0, 'VP+CD Separation': 0},
         'source_file': name00, 'sheet_name': '1', 'source_cells': ['1!E20:N20']},
        {'result_id': 'res_3', 'condition_id': 'cond_2', 'measurement_type': 'SPK Function',
         'condition_group': 'VP surface abnormal', 'date': '2024-12-17', 'line': 'E2-3A',
         'input_count': 89, 'ok_count': 88, 'ng_count': 1, 'ng_rate_decimal': 0.0112, 'ng_rate_percent': 1.1,
         'metric_name': 'Total NG rate Function', 'metric_value': 1.1, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'Air leak': 0, 'NG Sigma SPL': 0, 'NG Sigma THD': 0, 'NG Sigma SPL+THD': 1, 'NG Sigma SPL+THD+F0': 0, 'NG Hearing Noise': 0, 'NG Hearing Touch': 0},
         'source_file': name00, 'sheet_name': '1', 'source_cells': ['1!E25:N25']},
        {'result_id': 'res_4', 'condition_id': 'cond_2', 'measurement_type': 'SPK Function',
         'condition_group': 'Normal VP (baseline)', 'date': '2024-12-17', 'line': 'E2-3A',
         'input_count': 794, 'ok_count': 772, 'ng_count': 22, 'ng_rate_decimal': 0.0277, 'ng_rate_percent': 2.8,
         'metric_name': 'Total NG rate Function', 'metric_value': 2.8, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'Air leak': 0, 'NG Sigma SPL': 0, 'NG Sigma THD': 0, 'NG Sigma SPL+THD': 0, 'NG Sigma SPL+THD+F0': 0, 'NG Hearing Noise': 22, 'NG Hearing Touch': 0},
         'source_file': name00, 'sheet_name': '1', 'source_cells': ['1!E26:N26']},
        {'result_id': 'res_5', 'condition_id': 'cond_3', 'measurement_type': 'Module Function',
         'condition_group': 'VP surface abnormal', 'date': '2024-12-19', 'line': 'B2-8A',
         'input_count': 87, 'ok_count': 87, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Total NG rate Module', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'Bako': 0, 'Hearing1': 0, 'Hearing2': 0, 'Air leak': 0},
         'source_file': name00, 'sheet_name': '1', 'source_cells': ['1!E37:K37']},
        {'result_id': 'res_6', 'condition_id': 'cond_3', 'measurement_type': 'Module Function',
         'condition_group': 'Normal (baseline)', 'date': '2024-12-19', 'line': 'B2-8A',
         'input_count': 100, 'ok_count': 100, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Total NG rate Module', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'Bako': 0, 'Hearing1': 0, 'Hearing2': 0, 'Air leak': 0},
         'source_file': name00, 'sheet_name': '1', 'source_cells': ['1!E38:K38']},
        {'result_id': 'res_7', 'condition_id': 'cond_4', 'measurement_type': 'Reliability',
         'condition_group': 'Tension VP+CD', 'date': '2024-12', 'line': 'A1',
         'input_count': 5, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Tension', 'metric_value': None, 'unit': None, 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name00, 'sheet_name': '1', 'source_cells': ['1!E30']},
        {'result_id': 'res_8', 'condition_id': 'cond_4', 'measurement_type': 'Reliability',
         'condition_group': 'Load test 96h White noise EIA Max 80%', 'date': '2024-12', 'line': 'A1',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Load test 96h', 'metric_value': None, 'unit': None, 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name00, 'sheet_name': '1', 'source_cells': ['1!E31']},
        {'result_id': 'res_9', 'condition_id': 'cond_4', 'measurement_type': 'Reliability',
         'condition_group': 'Temperature & Humidity', 'date': '2024-12', 'line': 'A1',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Temperature & humidity', 'metric_value': None, 'unit': None, 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name00, 'sheet_name': '1', 'source_cells': ['1!E32']},
    ],
    'conclusions': [
        {'conclusion_id': 'concl_1', 'topic': 'Decision',
         'statement_from_report': 'Result test function SPK line and modul line OK => Can use',
         'normalized_interpretation': 'Sub1 vision NG rate equal: VP abnormal 0.0% vs Normal 0.0%. Function: VP abnormal 1.1% (1/89) vs Normal 2.8% (22/794) = 0.39x, 60.8% improved vs same-event Normal but sample size small. Module line: both 0.0%. Reliability tension/load/temp&humidity all PASS. Conclude: VP surface abnormal material is usable.',
         'source_file': name00, 'sheet_name': '1', 'source_cells': ['1!A39:A41']},
    ],
    'troubleshooting_index': {
        'defect_name': 'VP Surface Abnormal',
        'when_user_asks': ['Can VP with surface abnormal be used?', 'How to qualify suspect VP material?'],
        'suggested_checks': [
            {'hint_id': 'hint_1', 'check_item': 'Compare Sub1 vision NG rate of suspect VP lot vs same-day Normal VP',
             'reason': 'Sub1 vision is first defect gate; abnormal-VP=0.0% vs Normal=0.0% gave equivalent vision pass in this report.',
             'evidence_strength': 'medium', 'related_process': 'Sub 1 Vision', 'related_part': 'VP',
             'source_file': name00, 'sheet_name': '1', 'source_cells': ['1!B19:N20']},
            {'hint_id': 'hint_2', 'check_item': 'Run SPK function on suspect VP lot and compare Hearing Noise vs same-event Normal',
             'reason': 'Function NG dominated by Hearing Noise on Normal; abnormal-VP 1.1% (1 SPL+THD) vs Normal 2.8% confirmed no extra hearing risk.',
             'evidence_strength': 'medium', 'related_process': 'SPK Function', 'related_part': 'VP',
             'source_file': name00, 'sheet_name': '1', 'source_cells': ['1!B25:N26']},
            {'hint_id': 'hint_3', 'check_item': 'Run reliability set: Tension VP+CD 5pcs, Load 96h 5pcs, Temp&Humidity 5pcs',
             'reason': 'All three passed for VP surface abnormal lot, supporting release decision.',
             'evidence_strength': 'medium', 'related_process': 'Reliability', 'related_part': 'VP/CD assembly',
             'source_file': name00, 'sheet_name': '1', 'source_cells': ['1!B30:G32']},
        ],
        'limitations': ['Function test size only 89pcs; small sample for low NG rate detection.']
    },
    'ai_extraction_log': {
        'confidence': 0.8,
        'assumptions': ['Title BRS-161016 vs filename BRS-161014: treated title as authoritative model.', 'Report year inferred as 2024 from filename.'],
        'warnings': ['Sub1 and Module NG rates were both 0.0% - no defect-rate spread to draw strong conclusions from.'],
        'decision_rationale': 'Same-event Normal baselines exist for Sub1 vision (0.0% vs 0.0%), Function (1.1% vs 2.8% = 0.39x, 60.8% improved but n=89), and Module line (0.0% vs 0.0%). Reliability all PASS. No worsening signal; VP surface abnormal lot usable.'
    },
}

tr_en_00 = {
    'document': {'title': result00['document']['title'], 'purpose': result00['document']['purpose'], 'content': result00['document']['content']},
    'conclusions': {'concl_1': {'topic': 'Decision', 'statement_from_report': result00['conclusions'][0]['statement_from_report'], 'normalized_interpretation': result00['conclusions'][0]['normalized_interpretation']}},
    'hints': {h2['hint_id']: {'check_item': h2['check_item'], 'reason': h2['reason']} for h2 in result00['troubleshooting_index']['suggested_checks']},
    'log': {'assumptions': result00['ai_extraction_log']['assumptions'], 'warnings': result00['ai_extraction_log']['warnings'], 'decision_rationale': result00['ai_extraction_log']['decision_rationale']},
}

tr_ko_00 = {
    'document': {
        'title': 'VP 자재 표면 이상 시험 리포트 BRS-161016',
        'purpose': '표면이 이상한 VP 자재가 사용 가능한지 시험하여 확인한다.',
        'content': [
            'SPK 라인 시험 및 Sub 1 비전·기능 확인.',
            'A1에서 신뢰성 시험: Tension VP+CD 5pcs, Load test 96h 5pcs, Temperature & humidity 5pcs.',
            '모듈 라인 시험 및 기능 확인.'
        ]
    },
    'conclusions': {'concl_1': {
        'topic': '결정',
        'statement_from_report': 'SPK 라인 및 모듈 라인 기능 시험 결과 OK => 사용 가능',
        'normalized_interpretation': 'Sub1 비전 NG율 동일: VP 이상 0.0% vs Normal 0.0%. Function: VP 이상 1.1%(1/89) vs Normal 2.8%(22/794) = 0.39배, 동일 이벤트 Normal 대비 60.8% 개선이나 샘플 수 작음. 모듈 라인: 양쪽 0.0%. 신뢰성 Tension/Load/온습도 모두 PASS. 결론: VP 표면 이상 자재 사용 가능.'
    }},
    'hints': {
        'hint_1': {'check_item': '의심 VP 로트의 Sub1 비전 NG율을 동일자 Normal VP와 비교',
                   'reason': 'Sub1 비전이 1차 게이트; 이상-VP 0.0% vs Normal 0.0%로 동등 통과.'},
        'hint_2': {'check_item': '의심 VP 로트로 SPK Function 시험 후 Hearing Noise를 동일 이벤트 Normal과 비교',
                   'reason': 'Normal은 Hearing Noise가 주 NG; 이상-VP 1.1%(1 SPL+THD) vs Normal 2.8%로 hearing 위험 추가 없음 확인.'},
        'hint_3': {'check_item': '신뢰성 세트: Tension VP+CD 5pcs, Load 96h 5pcs, 온습도 5pcs 실시',
                   'reason': 'VP 표면 이상 로트에서 모두 PASS, 출하 결정 근거.'}
    },
    'log': {
        'assumptions': ['타이틀 BRS-161016 vs 파일명 BRS-161014: 타이틀을 모델로 채택.', '연도는 파일명에서 2024로 추정.'],
        'warnings': ['Sub1과 Module NG율이 0.0%로 동일 - 결정의 통계적 근거 약함.'],
        'decision_rationale': '동일 이벤트 Normal 비교 존재: Sub1(0.0% vs 0.0%), Function(1.1% vs 2.8% = 0.39배, 60.8% 개선이나 n=89), 모듈(0.0% vs 0.0%). 신뢰성 전부 PASS. 악화 신호 없음; VP 표면 이상 로트 사용 가능.'
    }
}

tr_vi_00 = {
    'document': {
        'title': 'BÁO CÁO TEST VẬT LIỆU VP BỀ MẶT BẤT THƯỜNG BRS-161016',
        'purpose': 'Test vật liệu VP có bề mặt bất thường để xác nhận có thể dùng hay không.',
        'content': [
            'Test line SPK và check vision Sub 1, function.',
            'Test reliability tại A1: Tension VP+CD 5pcs, Load test 96h 5pcs, Temperature & humidity 5pcs.',
            'Test line module và check function.'
        ]
    },
    'conclusions': {'concl_1': {
        'topic': 'Quyết định',
        'statement_from_report': 'Kết quả test function line SPK và line module OK => Có thể dùng',
        'normalized_interpretation': 'NG rate Sub1 vision bằng nhau: VP bất thường 0.0% vs Normal 0.0%. Function: VP bất thường 1.1% (1/89) vs Normal 2.8% (22/794) = 0.39x, cải thiện 60.8% so với Normal cùng sự kiện nhưng size nhỏ. Module line: cả hai 0.0%. Reliability Tension/Load/Temp&Humidity đều PASS. Kết luận: Vật liệu VP bề mặt bất thường có thể dùng.'
    }},
    'hints': {
        'hint_1': {'check_item': 'So sánh NG rate Sub1 vision của lot VP nghi vấn với Normal VP cùng ngày',
                   'reason': 'Sub1 vision là cổng NG đầu tiên; abnormal-VP 0.0% vs Normal 0.0% cho thấy pass tương đương.'},
        'hint_2': {'check_item': 'Chạy function SPK trên lot VP nghi vấn và so sánh Hearing Noise với Normal cùng sự kiện',
                   'reason': 'NG function chủ yếu là Hearing Noise ở Normal; abnormal-VP 1.1% (1 SPL+THD) vs Normal 2.8% xác nhận không tăng risk hearing.'},
        'hint_3': {'check_item': 'Chạy bộ reliability: Tension VP+CD 5pcs, Load 96h 5pcs, Temp&Humidity 5pcs',
                   'reason': 'Cả ba đều PASS với lot VP bề mặt bất thường, làm căn cứ release.'}
    },
    'log': {
        'assumptions': ['Title BRS-161016 vs filename BRS-161014: dùng title làm model.', 'Năm suy ra 2024 từ filename.'],
        'warnings': ['NG rate Sub1 và Module đều 0.0% - cơ sở thống kê yếu.'],
        'decision_rationale': 'Có baseline Normal cùng sự kiện: Sub1 (0.0% vs 0.0%), Function (1.1% vs 2.8% = 0.39x, cải thiện 60.8% nhưng n=89), Module (0.0% vs 0.0%). Reliability đều PASS. Không có signal xấu đi; lot VP bề mặt bất thường có thể dùng.'
    }
}

run(name00, result00, tr_ko_00, tr_en_00, tr_vi_00)


# ===== DS 01 =====
name01 = '25. BRS-161014 Report test Jig CM+CP NG after check by Jig master 12.28.2023'
result01 = {
    'schema_version': '0.1',
    'document': {
        'document_id': '', 'source_file': name01, 'source_sheet': 'Report (2)',
        'title': 'REPORT TEST CM+CP JIG NG AFTER CHECK BY MASTER JIG - BRS-161014',
        'model': 'BRS-161014', 'report_date': '2023-12-29', 'department': 'ME',
        'marker': 'Nhung', 'line': '',
        'report_type': 'normal_comparison',
        'primary_defect': {'canonical_name': 'Yoke Offset', 'aliases_in_document': ['yoke offset', 'NG Yoke offset']},
        'related_defects': ['NG Hearing Noise', 'Coil sus offset', 'Frame + Coil offset', 'Coil Deform'],
        'parts': ['Yoke', 'Jig', 'Coil', 'Frame'],
        'processes': ['CM+CP Jig', 'Decap', 'Function'],
        'purpose': 'Improve NG Yoke offset by sorting jigs and re-testing function with rejected jigs after master jig recheck.',
        'content': [
            'Use jigs marked NG and re-test function.',
            'Decap to confirm yoke offset.',
            'Compare with normal lot.',
            'Worker insertion direction (long-axis side only) suspected as a contributing factor.'
        ],
        'source_cells': {'title': ['Report (2)!B2'], 'date': ['Report (2)!N3'], 'purpose': ['Report (2)!A6'], 'content': ['Report (2)!A8', 'Sheet1!A1:A4']}
    },
    'test_conditions': [
        {'condition_id': 'cond_1', 'condition_group': 'Jig NG (after master jig recheck OK when cool) vs Normal',
         'line': '', 'process': 'CM+CP', 'changed_factor': 'Jig sorted as NG when hot then OK when cool',
         'before_value': 'Normal jig', 'after_value': '4 NG jigs (master OK after cool)',
         'unit': None, 'machine': None, 'jig': '4pcs NG-when-hot, OK after cooling',
         'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None,
         'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': name01, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!B15:N17']},
    ],
    'results': [
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Test (Jig NG sorted, cooling OK)', 'date': '2023-12-29', 'line': '',
         'input_count': 120, 'ok_count': 116, 'ng_count': 4, 'ng_rate_decimal': 0.0333, 'ng_rate_percent': 3.3,
         'metric_name': 'Total NG rate Function', 'metric_value': 3.3, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'NG Sigma SPL': 0, 'NG Sigma THD': 0, 'NG Sigma SPL+THD': 0, 'NG Sigma SPL+THD+F0': 0, 'NG Hearing Noise': 4, 'NG Hearing Touch': 0},
         'source_file': name01, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!C15:M15']},
        {'result_id': 'res_2', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Normal (baseline)', 'date': '2023-12-29', 'line': '',
         'input_count': 560, 'ok_count': 539, 'ng_count': 21, 'ng_rate_decimal': 0.0375, 'ng_rate_percent': 3.8,
         'metric_name': 'Total NG rate Function', 'metric_value': 3.8, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'NG Sigma SPL': 0, 'NG Sigma THD': 0, 'NG Sigma SPL+THD': 0, 'NG Sigma SPL+THD+F0': 0, 'NG Hearing Noise': 21, 'NG Hearing Touch': 0},
         'source_file': name01, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!C17:M17']},
        {'result_id': 'res_3', 'condition_id': 'cond_1', 'measurement_type': 'Decap',
         'condition_group': 'NG Hearing Noise + Touch decap analysis', 'date': '2023-12-30', 'line': '',
         'input_count': 4, 'ok_count': None, 'ng_count': 4, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Decap NG analysis breakdown', 'metric_value': None, 'unit': None, 'judgement': None,
         'ng_breakdown': {'Coil sus offset': {'count': 2, 'rate': 0.5},
                          'Frame + Coil offset': {'count': 1, 'rate': 0.25},
                          'Coil Deform': {'count': 1, 'rate': 0.25},
                          'Yoke Offset': {'count': 0, 'rate': 0.0},
                          'Over Glue': {'count': 0, 'rate': 0.0},
                          'Particle': {'count': 0, 'rate': 0.0},
                          'Gap Coil+SP': {'count': 0, 'rate': 0.0},
                          'Coil Separate': {'count': 0, 'rate': 0.0}},
         'source_file': name01, 'sheet_name': 'TEST YOKE', 'source_cells': ['TEST YOKE!C42:G50']},
    ],
    'conclusions': [
        {'conclusion_id': 'concl_1', 'topic': 'Decision',
         'statement_from_report': 'When test Jig yoke NG. Function don\'t have NG by yoke offset. => Jig check when cooling OK',
         'normalized_interpretation': 'Test (jigs initially NG-when-hot, OK after cooling) Function NG = 3.3% (4/120) vs same-event Normal = 3.8% (21/560), ratio 0.88x, 11.8% improved vs Normal. Both populations only show Hearing Noise. Decap of 4 NG samples shows 0 yoke offset; main decap reasons are Coil sus offset (50%) and Frame+Coil offset (25%). Conclusion supports releasing jigs that pass master check after cooling.',
         'source_file': name01, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!A22']},
        {'conclusion_id': 'concl_2', 'topic': 'Decap observation',
         'statement_from_report': 'Don\'t have NG Yoke offset',
         'normalized_interpretation': 'Decap result: zero yoke offset among the 4 function-NG samples; Hearing-Noise NG appears driven by coil-related offsets (Coil sus + Frame+Coil), not yoke.',
         'source_file': name01, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!A20']},
    ],
    'troubleshooting_index': {
        'defect_name': 'Yoke Offset / Hearing Noise',
        'when_user_asks': ['How to handle jigs flagged NG when hot but OK when cool?', 'What causes Hearing Noise after jig-related NG?'],
        'suggested_checks': [
            {'hint_id': 'hint_1', 'check_item': 'Re-check NG-when-hot jigs against master jig after cooling',
             'reason': 'In this lot, 4 NG-when-hot jigs all passed master after cooling and function NG was 3.3% vs Normal 3.8% (0.88x, 11.8% improved).',
             'evidence_strength': 'medium', 'related_process': 'CM+CP', 'related_part': 'Jig',
             'source_file': name01, 'sheet_name': 'Report (2)', 'source_cells': ['Report (2)!A11', 'Report (2)!B15']},
            {'hint_id': 'hint_2', 'check_item': 'Decap NG-Function samples and tally yoke/coil/frame offset distribution',
             'reason': 'Decap revealed 0/4 yoke offset, 2/4 coil sus offset, 1/4 frame+coil offset, 1/4 coil deform - jig-thermal effect did not produce yoke offset here.',
             'evidence_strength': 'medium', 'related_process': 'Decap', 'related_part': 'Yoke / Coil / Frame',
             'source_file': name01, 'sheet_name': 'TEST YOKE', 'source_cells': ['TEST YOKE!C42:G50']},
            {'hint_id': 'hint_3', 'check_item': 'Audit worker insertion direction - long-axis side guidance',
             'reason': 'Sheet1 notes only long-axis side defects, suspected to come from worker guiding insertion only on long-axis side.',
             'evidence_strength': 'low', 'related_process': 'Coil insertion', 'related_part': 'Coil',
             'source_file': name01, 'sheet_name': 'Sheet1', 'source_cells': ['Sheet1!A2:A4']},
        ],
        'limitations': ['Function NG sample size for test condition was 120, baseline 560, both reporting only Hearing-Noise NG.']
    },
    'ai_extraction_log': {
        'confidence': 0.75,
        'assumptions': ['"NG Yoke offset" treated as primary defect target.', 'Decap percentages stored as decimal rate (0.5 = 50%).'],
        'warnings': ['Only 4 decap samples; small sample for absolute distribution claims.'],
        'decision_rationale': 'Test (jig NG-when-hot, OK after cooling) function NG 3.3% (4/120) vs same-event Normal 3.8% (21/560) = 0.88x, 11.8% improved. Decap shows yoke offset is NOT the driver in this batch; coil-related offsets dominate. Supports the report decision to re-use jigs that pass master after cooling.'
    },
}
tr_en_01 = {
    'document': {'title': result01['document']['title'], 'purpose': result01['document']['purpose'], 'content': result01['document']['content']},
    'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in result01['conclusions']},
    'hints': {h2['hint_id']: {'check_item': h2['check_item'], 'reason': h2['reason']} for h2 in result01['troubleshooting_index']['suggested_checks']},
    'log': {'assumptions': result01['ai_extraction_log']['assumptions'], 'warnings': result01['ai_extraction_log']['warnings'], 'decision_rationale': result01['ai_extraction_log']['decision_rationale']},
}
tr_ko_01 = {
    'document': {
        'title': 'CM+CP Jig NG 마스터 지그 재확인 후 시험 리포트 - BRS-161014',
        'purpose': '마스터 지그 재점검 후 OK가 된 NG 지그로 기능 시험을 재실시하여 Yoke offset 개선 여부 확인.',
        'content': [
            'NG 표기된 지그로 기능 재시험.',
            'Decap으로 yoke offset 여부 확인.',
            'Normal 로트와 비교.',
            '작업자 삽입 방향(장축 가이드)이 영향일 가능성 메모.'
        ]
    },
    'conclusions': {
        'concl_1': {'topic': '결정',
                    'statement_from_report': 'Yoke NG 지그 시험 시 yoke offset에 의한 NG 없음 => 지그가 식은 후 OK면 사용 가능',
                    'normalized_interpretation': 'Test(가열 시 NG, 냉각 후 OK 지그) Function NG = 3.3%(4/120) vs 동일 이벤트 Normal = 3.8%(21/560), 비 0.88배, Normal 대비 11.8% 개선. 두 집단 모두 Hearing Noise만 발생. NG 4건 decap에서 yoke offset 0, 주 원인은 Coil sus offset(50%) 및 Frame+Coil offset(25%). 냉각 후 master OK 지그 재사용 결정 지지.'},
        'concl_2': {'topic': 'Decap 관찰',
                    'statement_from_report': 'Yoke offset NG 없음',
                    'normalized_interpretation': 'Decap 결과 4개 NG 샘플 중 yoke offset 0; Hearing Noise NG는 coil 관련 offset이 주요 원인.'},
    },
    'hints': {
        'hint_1': {'check_item': '가열 시 NG로 분류된 지그를 냉각 후 마스터 지그로 재확인',
                   'reason': '본 로트에서 가열-NG 지그 4개 모두 냉각 후 master OK, 기능 NG 3.3% vs Normal 3.8%(0.88배, 11.8% 개선).'},
        'hint_2': {'check_item': 'Function-NG 샘플 decap 후 yoke/coil/frame offset 분포 집계',
                   'reason': 'Decap 4건 중 yoke offset 0/4, coil sus offset 2/4, frame+coil offset 1/4, coil deform 1/4 - 지그 열영향이 yoke offset로 직결되지 않음.'},
        'hint_3': {'check_item': '작업자 장축 삽입 방향 감사',
                   'reason': 'Sheet1에 장축쪽에만 불량 발생, 작업자가 장축으로 가이드하면서 발생 가능 메모.'}
    },
    'log': {
        'assumptions': ['"NG Yoke offset"을 주 결함으로 간주.', 'Decap 비율은 decimal rate(0.5=50%)로 저장.'],
        'warnings': ['Decap 샘플 4건으로 분포 추론은 작은 표본.'],
        'decision_rationale': '동일 이벤트 Normal 대비 Function NG 3.3% vs 3.8% = 0.88배, 11.8% 개선. Decap에서 yoke offset는 주 원인이 아님(coil 관련). 냉각 후 master OK 지그 재사용 결정 타당.'
    }
}
tr_vi_01 = {
    'document': {
        'title': 'BÁO CÁO TEST JIG CM+CP NG SAU KHI CHECK BẰNG MASTER JIG - BRS-161014',
        'purpose': 'Cải thiện NG Yoke offset bằng cách dùng các jig bị NG khi nóng, OK khi nguội rồi test lại function.',
        'content': [
            'Dùng jig bị đánh NG để test lại function.',
            'Decap để xác nhận yoke offset.',
            'So sánh với lot Normal.',
            'Hướng đưa của công nhân (chỉ phía trục dài) nghi là yếu tố góp phần.'
        ]
    },
    'conclusions': {
        'concl_1': {'topic': 'Quyết định',
                    'statement_from_report': 'Khi test jig yoke NG, function không có NG do yoke offset => Jig sau khi nguội master check OK',
                    'normalized_interpretation': 'Test (jig NG khi nóng, OK khi nguội) function NG = 3.3% (4/120) vs Normal cùng sự kiện = 3.8% (21/560), tỉ lệ 0.88x, cải thiện 11.8% vs Normal. Cả hai chỉ có Hearing Noise. Decap 4 mẫu NG: 0 yoke offset; chủ yếu Coil sus offset (50%) và Frame+Coil offset (25%). Hỗ trợ quyết định dùng lại jig khi nguội master OK.'},
        'concl_2': {'topic': 'Quan sát decap',
                    'statement_from_report': 'Không có NG Yoke offset',
                    'normalized_interpretation': 'Decap 4 mẫu NG function: 0 yoke offset; NG Hearing Noise chủ yếu do offset liên quan coil.'},
    },
    'hints': {
        'hint_1': {'check_item': 'Re-check jig NG-khi-nóng bằng master jig sau khi nguội',
                   'reason': 'Lot này 4 jig NG-khi-nóng đều pass master sau khi nguội, NG function 3.3% vs Normal 3.8% (0.88x, cải thiện 11.8%).'},
        'hint_2': {'check_item': 'Decap mẫu NG-Function và tally phân bố yoke/coil/frame offset',
                   'reason': 'Decap: 0/4 yoke offset, 2/4 coil sus offset, 1/4 frame+coil offset, 1/4 coil deform - tác động nhiệt jig không gây yoke offset trong lô này.'},
        'hint_3': {'check_item': 'Kiểm tra hướng đưa của công nhân - guide phía trục dài',
                   'reason': 'Sheet1 ghi chỉ phía trục dài bị NG, nghi do công nhân guide insertion một phía.'}
    },
    'log': {
        'assumptions': ['"NG Yoke offset" làm defect chính.', 'Tỷ lệ decap lưu dưới dạng decimal rate (0.5 = 50%).'],
        'warnings': ['Decap chỉ 4 mẫu - sample size nhỏ cho phân bố.'],
        'decision_rationale': 'Function NG 3.3% (4/120) vs Normal cùng sự kiện 3.8% (21/560) = 0.88x, cải thiện 11.8%. Decap cho thấy yoke offset không phải driver; chủ yếu coil offset. Quyết định re-use jig khi nguội master OK hợp lý.'
    }
}
run(name01, result01, tr_ko_01, tr_en_01, tr_vi_01)


# ===== DS 02 =====
name02 = '25. BRS-161014 Report test bond 8030 at SUB3 date 12.9.2023'
result02 = {
    'schema_version': '0.1',
    'document': {
        'document_id': '', 'source_file': name02, 'source_sheet': 'Report',
        'title': 'REPORT TEST BOND 8030 AT SUB 3 OF BRS-161014',
        'model': 'BRS-161014', 'report_date': '2023-09-12', 'department': 'ME',
        'marker': 'Thao', 'line': '',
        'report_type': 'normal_comparison',
        'primary_defect': {'canonical_name': 'Suspension Separation', 'aliases_in_document': ['suspension separate', 'NG Suspension separate', 'separate VP CD']},
        'related_defects': ['NG Bonding', 'NG bending'],
        'parts': ['Frame', 'Suspension', 'Bond 8030', 'Bond 0930'],
        'processes': ['Sub 3 Frame bonding', 'Sub 3 Vision', 'Tension'],
        'purpose': 'Improve suspension separation by switching to Bond 8030 at SUB-3.',
        'content': [
            'Change bond to 8030.',
            'Check Frame bonding process.',
            'Check Vision Frame+Suspension.',
            'Check Tension.',
            'Quantity test: 200pcs.'
        ],
        'source_cells': {'title': ['Report!B2'], 'date': ['Report!K3'], 'purpose': ['Report!A6'], 'content': ['Report!A8:A12']}
    },
    'test_conditions': [
        {'condition_id': 'cond_1', 'condition_group': 'Bond 8030 vs Bond 0930 at Sub3 Frame bonding',
         'line': '', 'process': 'Sub 3 Frame bonding', 'changed_factor': 'Bond type',
         'before_value': 'Bond 0930', 'after_value': 'Bond 8030',
         'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None,
         'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': name02, 'sheet_name': 'Report', 'source_cells': ['Report!E17:H18']},
        {'condition_id': 'cond_2', 'condition_group': 'Bond 8030 vs Bond 0930 at Sub3 Vision (Suspension separate)',
         'line': '', 'process': 'Sub 3 Vision', 'changed_factor': 'Bond type & bond amount',
         'before_value': 'Bond 0930 0.3~0.4mg', 'after_value': 'Bond 8030 0.3~0.5mg',
         'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None,
         'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': '0.3~0.5mg vs 0.3~0.4mg', 'uv_energy': None,
         'source_file': name02, 'sheet_name': 'Report', 'source_cells': ['Report!D22:H25']},
        {'condition_id': 'cond_3', 'condition_group': 'Sheet1: 10 Min 130C combo (highlighted)',
         'line': '', 'process': 'Sub 3 Vision + SP+CO Bonding', 'changed_factor': 'Time x Temperature',
         'before_value': None, 'after_value': '10 Min 130C',
         'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None,
         'dry_time_sec': 600.0, 'temperature': '130C', 'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': name02, 'sheet_name': 'Sheet1', 'source_cells': ['Sheet1!A38:I38']},
    ],
    'results': [
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Sub3 Frame bonding',
         'condition_group': 'Bond 8030', 'date': '2023-09-12', 'line': '',
         'input_count': 48, 'ok_count': 48, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'NG Bonding rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'NG Bonding': 0}, 'source_file': name02, 'sheet_name': 'Report', 'source_cells': ['Report!E17:I17']},
        {'result_id': 'res_2', 'condition_id': 'cond_1', 'measurement_type': 'Sub3 Frame bonding',
         'condition_group': 'Bond 0930 (baseline)', 'date': '2023-09-12', 'line': '',
         'input_count': 48, 'ok_count': 48, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'NG Bonding rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'NG Bonding': 0}, 'source_file': name02, 'sheet_name': 'Report', 'source_cells': ['Report!E18:I18']},
        {'result_id': 'res_3', 'condition_id': 'cond_2', 'measurement_type': 'Sub3 Vision',
         'condition_group': 'Bond 8030 0.3~0.5mg', 'date': '2023-09-12', 'line': '',
         'input_count': 113, 'ok_count': 96, 'ng_count': 17, 'ng_rate_decimal': 0.1504, 'ng_rate_percent': 15.04,
         'metric_name': 'NG Suspension separate rate', 'metric_value': 15.04, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'Suspension Separation': 17}, 'source_file': name02, 'sheet_name': 'Report', 'source_cells': ['Report!F22:I22']},
        {'result_id': 'res_4', 'condition_id': 'cond_2', 'measurement_type': 'Sub3 Vision',
         'condition_group': 'Bond 8030 0.3mg', 'date': '2023-09-12', 'line': '',
         'input_count': 175, 'ok_count': 171, 'ng_count': 4, 'ng_rate_decimal': 0.0229, 'ng_rate_percent': 2.29,
         'metric_name': 'NG Suspension separate rate', 'metric_value': 2.29, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'Suspension Separation': 4}, 'source_file': name02, 'sheet_name': 'Report', 'source_cells': ['Report!F23:I23']},
        {'result_id': 'res_5', 'condition_id': 'cond_2', 'measurement_type': 'Sub3 Vision',
         'condition_group': 'Bond 8030 Total', 'date': '2023-09-12', 'line': '',
         'input_count': 288, 'ok_count': 267, 'ng_count': 21, 'ng_rate_decimal': 0.0729, 'ng_rate_percent': 7.29,
         'metric_name': 'NG Suspension separate rate', 'metric_value': 7.29, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'Suspension Separation': 21}, 'source_file': name02, 'sheet_name': 'Report', 'source_cells': ['Report!F24:I24']},
        {'result_id': 'res_6', 'condition_id': 'cond_2', 'measurement_type': 'Sub3 Vision',
         'condition_group': 'Bond 0930 0.3~0.4mg (baseline)', 'date': '2023-09-12', 'line': '',
         'input_count': 273, 'ok_count': 270, 'ng_count': 3, 'ng_rate_decimal': 0.011, 'ng_rate_percent': 1.10,
         'metric_name': 'NG Suspension separate rate', 'metric_value': 1.10, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'Suspension Separation': 3}, 'source_file': name02, 'sheet_name': 'Report', 'source_cells': ['Report!F25:I25']},
        {'result_id': 'res_7', 'condition_id': 'cond_3', 'measurement_type': 'Sub3 Vision',
         'condition_group': '10 Min 130C (highlighted)', 'date': '2023-09-08', 'line': '',
         'input_count': 376, 'ok_count': None, 'ng_count': 69, 'ng_rate_decimal': 0.184, 'ng_rate_percent': 18.4,
         'metric_name': 'NG bending rate', 'metric_value': 18.4, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG bending': 69}, 'source_file': name02, 'sheet_name': 'Sheet1', 'source_cells': ['Sheet1!E38:G38']},
        {'result_id': 'res_8', 'condition_id': 'cond_3', 'measurement_type': 'Main1 SP+CO bonding',
         'condition_group': '10 Min 130C (highlighted)', 'date': '2023-09-08', 'line': '',
         'input_count': 366, 'ok_count': None, 'ng_count': 9, 'ng_rate_decimal': 0.0246, 'ng_rate_percent': 2.46,
         'metric_name': 'NG bonding rate Main1', 'metric_value': 2.46, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'NG Bonding': 9}, 'source_file': name02, 'sheet_name': 'Sheet1', 'source_cells': ['Sheet1!H38:J38']},
    ],
    'conclusions': [
        {'conclusion_id': 'concl_1', 'topic': 'Decision',
         'statement_from_report': 'Decision section in source is empty.',
         'normalized_interpretation': 'Bond 8030 at 0.3~0.5mg suspension-separate NG = 15.04% (17/113); at 0.3mg = 2.29% (4/175); Total 7.29% (21/288). Baseline Bond 0930 0.3~0.4mg = 1.10% (3/273). Total ratio 7.29/1.10 = 6.63x, 562.7% worse than baseline; 0.5mg high-amount split 15.04/1.10 = 13.7x, 1267% worse. Bond 8030 trial worsens suspension separation vs Bond 0930; only the lower-amount 0.3mg slice approaches but does not match baseline (2.08x, 108% worse).',
         'source_file': name02, 'sheet_name': 'Report', 'source_cells': ['Report!A32']},
        {'conclusion_id': 'concl_2', 'topic': 'Time x Temperature matrix observation',
         'statement_from_report': 'Sheet1 highlights 10 Min @ 130C row.',
         'normalized_interpretation': 'At Bond 8030 trial, 10 Min 130C cell shows 18.4% bending NG and 2.46% SP+CO bonding NG (376/366 inputs). Other DOE cells: 5 Min 100C 31.8%/4.9%, 5 Min 160C 35.5%/8.5%, 10 Min 100C/130C combinations between 17.7% and 31.8%. 10 Min 130C is the relatively best cell in this DOE but still high.',
         'source_file': name02, 'sheet_name': 'Sheet1', 'source_cells': ['Sheet1!A35:I41']},
    ],
    'troubleshooting_index': {
        'defect_name': 'Suspension Separation',
        'when_user_asks': ['Does changing bond to 8030 reduce suspension separation?', 'What time/temperature combination minimizes bending NG?'],
        'suggested_checks': [
            {'hint_id': 'hint_1', 'check_item': 'Compare Sub3 vision suspension-separate NG of new bond vs Bond 0930 baseline at matched bond amount',
             'reason': 'Bond 8030 total 7.29% vs Bond 0930 1.10% = 6.63x, 562.7% worse; high-amount 0.3~0.5mg split 13.7x worse, 0.3mg split 2.08x worse - Bond 8030 not yet acceptable.',
             'evidence_strength': 'high', 'related_process': 'Sub 3 Vision', 'related_part': 'Suspension / Bond',
             'source_file': name02, 'sheet_name': 'Report', 'source_cells': ['Report!F22:I25']},
            {'hint_id': 'hint_2', 'check_item': 'Sweep time x temperature DOE for new bond and lock in best cell',
             'reason': 'Sheet1 DOE shows 10 Min 130C achieves 18.4% bending NG, the lowest of recorded cells; 5 Min 160C the worst at 35.5%.',
             'evidence_strength': 'medium', 'related_process': 'Sub 3 Frame bonding', 'related_part': 'Bond',
             'source_file': name02, 'sheet_name': 'Sheet1', 'source_cells': ['Sheet1!A35:I41']},
            {'hint_id': 'hint_3', 'check_item': 'Reduce bond amount and re-evaluate suspension separation',
             'reason': 'Within Bond 8030, lowering amount from 0.3~0.5mg to 0.3mg dropped NG from 15.04% to 2.29% (6.57x improvement within the test arm).',
             'evidence_strength': 'medium', 'related_process': 'Sub 3 Frame bonding', 'related_part': 'Bond amount',
             'source_file': name02, 'sheet_name': 'Report', 'source_cells': ['Report!H22:I23']},
        ],
        'limitations': ['Tension results are #DIV/0!; no usable tension data. Decision section in source is blank.']
    },
    'ai_extraction_log': {
        'confidence': 0.8,
        'assumptions': ['Date "12-Sep" treated as 2023-09-12 based on filename "12.9.2023".', 'Sheet1 time-temperature combo treated as DOE rather than baseline.'],
        'warnings': ['Tension cells empty (#DIV/0!); no actual numeric tension comparison.', 'Decision section in source is blank; conclusion built from data.'],
        'decision_rationale': 'Bond 8030 total Sub3 vision suspension-separate NG 7.29% (21/288) vs same-event Bond 0930 0.3~0.4mg 1.10% (3/273) = 6.63x, 562.7% worse. Bond 8030 not yet a viable replacement at current settings.'
    },
}
tr_en_02 = {
    'document': {'title': result02['document']['title'], 'purpose': result02['document']['purpose'], 'content': result02['document']['content']},
    'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in result02['conclusions']},
    'hints': {h2['hint_id']: {'check_item': h2['check_item'], 'reason': h2['reason']} for h2 in result02['troubleshooting_index']['suggested_checks']},
    'log': {'assumptions': result02['ai_extraction_log']['assumptions'], 'warnings': result02['ai_extraction_log']['warnings'], 'decision_rationale': result02['ai_extraction_log']['decision_rationale']},
}
tr_ko_02 = {
    'document': {
        'title': 'BRS-161014 Sub3 Bond 8030 시험 리포트',
        'purpose': 'Bond 8030으로 변경하여 Sub3 Suspension separate 개선 여부 확인.',
        'content': [
            'Bond 8030으로 교체.', 'Frame bonding 공정 확인.',
            'Frame+Suspension 비전 확인.', 'Tension 확인.', '시험 수량: 200pcs.'
        ]
    },
    'conclusions': {
        'concl_1': {'topic': '결정', 'statement_from_report': '원본 Decision 섹션 비어있음.',
                    'normalized_interpretation': 'Bond 8030 0.3~0.5mg suspension-separate NG = 15.04%(17/113); 0.3mg = 2.29%(4/175); Total 7.29%(21/288). Baseline Bond 0930 0.3~0.4mg = 1.10%(3/273). 총합 비 7.29/1.10 = 6.63배, 베이스라인 대비 562.7% 악화; 0.5mg 고량 13.7배(1267% 악화), 0.3mg 슬라이스 2.08배(108% 악화). Bond 8030은 현재 조건에서 Suspension separation을 악화시킴.'},
        'concl_2': {'topic': '시간×온도 매트릭스 관찰',
                    'statement_from_report': 'Sheet1에서 10 Min @ 130C 행 강조.',
                    'normalized_interpretation': 'Bond 8030 시험에서 10 Min 130C 셀: bending NG 18.4%, SP+CO bonding NG 2.46%(376/366). 다른 DOE 셀: 5분 100C 31.8%/4.9%, 5분 160C 35.5%/8.5%, 10분 100C/130C 17.7~31.8%. 10분 130C가 DOE 중 상대적 최선이나 여전히 높음.'}
    },
    'hints': {
        'hint_1': {'check_item': '새 본드 Sub3 vision suspension-separate NG를 동일 본드량의 Bond 0930과 비교',
                   'reason': 'Bond 8030 총 7.29% vs Bond 0930 1.10% = 6.63배(562.7% 악화); 고량 13.7배, 저량 2.08배 - Bond 8030 미수용.'},
        'hint_2': {'check_item': '신본드 시간×온도 DOE 진행 및 최적 셀 고정',
                   'reason': 'Sheet1 DOE에서 10분 130C가 bending NG 18.4%로 가장 낮음; 5분 160C가 35.5%로 최악.'},
        'hint_3': {'check_item': '본드량 감소 후 suspension separation 재평가',
                   'reason': 'Bond 8030 내에서 0.3~0.5mg→0.3mg 변경 시 NG 15.04%→2.29%(6.57배 개선).'}
    },
    'log': {
        'assumptions': ['"12-Sep"는 파일명 "12.9.2023" 기준 2023-09-12.', 'Sheet1 시간×온도 행은 DOE로 간주.'],
        'warnings': ['Tension 셀이 #DIV/0!로 비어있음; 수치 비교 불가.', '원본 Decision 섹션 비어있음.'],
        'decision_rationale': 'Bond 8030 총 Sub3 vision NG 7.29%(21/288) vs 동일 이벤트 Bond 0930 0.3~0.4mg 1.10%(3/273) = 6.63배, 562.7% 악화. 현 조건의 Bond 8030은 대체 본드로 부적합.'
    }
}
tr_vi_02 = {
    'document': {
        'title': 'BÁO CÁO TEST BOND 8030 TẠI SUB3 - BRS-161014',
        'purpose': 'Chuyển sang Bond 8030 tại SUB-3 để cải thiện NG suspension separate.',
        'content': [
            'Đổi sang Bond 8030.', 'Check công đoạn Frame bonding.',
            'Check vision Frame+Suspension.', 'Check Tension.', 'Số lượng test: 200pcs.'
        ]
    },
    'conclusions': {
        'concl_1': {'topic': 'Quyết định', 'statement_from_report': 'Phần Decision trong file gốc trống.',
                    'normalized_interpretation': 'Bond 8030 0.3~0.5mg NG suspension-separate = 15.04% (17/113); 0.3mg = 2.29% (4/175); Total 7.29% (21/288). Baseline Bond 0930 0.3~0.4mg = 1.10% (3/273). Tổng tỉ lệ 7.29/1.10 = 6.63x, xấu hơn baseline 562.7%; mức cao 0.5mg 13.7x (xấu 1267%), 0.3mg 2.08x (xấu 108%). Bond 8030 ở điều kiện hiện tại làm xấu suspension separation.'},
        'concl_2': {'topic': 'Quan sát ma trận Time×Temp',
                    'statement_from_report': 'Sheet1 highlight hàng 10 Min @ 130C.',
                    'normalized_interpretation': 'Bond 8030, cell 10 Min 130C: NG bending 18.4%, NG SP+CO bonding 2.46% (376/366). Cell khác: 5 Min 100C 31.8%/4.9%, 5 Min 160C 35.5%/8.5%, 10 Min 100C/130C 17.7~31.8%. 10 Min 130C tốt nhất trong DOE nhưng vẫn cao.'}
    },
    'hints': {
        'hint_1': {'check_item': 'So sánh NG suspension-separate Sub3 vision của bond mới với Bond 0930 baseline cùng bond amount',
                   'reason': 'Bond 8030 total 7.29% vs Bond 0930 1.10% = 6.63x (xấu 562.7%); mức cao 13.7x, mức thấp 2.08x - Bond 8030 chưa đạt.'},
        'hint_2': {'check_item': 'Sweep DOE time×temperature cho bond mới và khóa cell tốt nhất',
                   'reason': 'Sheet1 DOE: 10 Min 130C NG bending 18.4% (thấp nhất); 5 Min 160C 35.5% (xấu nhất).'},
        'hint_3': {'check_item': 'Giảm bond amount và đánh giá lại suspension separation',
                   'reason': 'Trong Bond 8030, giảm 0.3~0.5mg về 0.3mg, NG giảm 15.04%→2.29% (cải thiện 6.57x).'}
    },
    'log': {
        'assumptions': ['"12-Sep" tính là 2023-09-12 theo filename "12.9.2023".', 'Sheet1 time×temp coi là DOE.'],
        'warnings': ['Cell Tension trống (#DIV/0!); không có dữ liệu tension.', 'Phần Decision trong file trống.'],
        'decision_rationale': 'Bond 8030 total NG Sub3 vision 7.29% (21/288) vs Bond 0930 0.3~0.4mg 1.10% (3/273) = 6.63x, xấu 562.7%. Bond 8030 chưa thể thay thế ở điều kiện này.'
    }
}
run(name02, result02, tr_ko_02, tr_en_02, tr_vi_02)
