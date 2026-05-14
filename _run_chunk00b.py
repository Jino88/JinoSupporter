"""Run chunk 00 datasets 03-08."""
from __future__ import annotations
import sys, io
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import _ai_batch_helper as h


def run(name, result, tr_ko, tr_en, tr_vi):
    ok = h.commit_dataset(name, result, tr_ko, tr_en, tr_vi)
    if ok:
        print(f'[OK {name}]')
    else:
        print(f'[FAIL {name} commit_dataset returned False]')


def _tr_from_result(r):
    return {
        'document': {'title': r['document']['title'], 'purpose': r['document']['purpose'], 'content': r['document']['content']},
        'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in r['conclusions']},
        'hints': {hh['hint_id']: {'check_item': hh['check_item'], 'reason': hh['reason']} for hh in r['troubleshooting_index']['suggested_checks']},
        'log': {'assumptions': r['ai_extraction_log']['assumptions'], 'warnings': r['ai_extraction_log']['warnings'], 'decision_rationale': r['ai_extraction_log']['decision_rationale']},
    }


# ===== DS 03 =====
name03 = '25. BRS-161016 Report test new machine vision VP+CD at Sub 1 22.03.2024'
result03 = {
    'schema_version': '0.1',
    'document': {
        'document_id': '', 'source_file': name03, 'source_sheet': '5.3',
        'title': 'REPORT TEST NEW MACHINE VISION VP+CD AT SUB 1 BRS-161016',
        'model': 'BRS-161016', 'report_date': '2024-03-22', 'department': 'ME',
        'marker': 'Thuy', 'line': 'Sub 1',
        'report_type': 'before_after_dimension',
        'primary_defect': {'canonical_name': 'VP+CD Pickup/Press/Vision NG', 'aliases_in_document': ['Pickup NG', 'VP+CD float', 'Dome damage', 'VP damage']},
        'related_defects': ['Dome Damage', 'VP Damage', 'VP+CD float (press)'],
        'parts': ['VP', 'CD', 'Dome', 'Pusher', 'AOI'],
        'processes': ['Sub 1 Pickup', 'Sub 1 Press', 'Sub 1 Vision', 'Capa measurement'],
        'purpose': 'Evaluate whether new Sub-1 VP+CD vision machine is usable: pickup OK rate, NG detection ability, press OK rate, vision damage, AOI tuning, and capacity.',
        'content': [
            'Check pickup VP OK/NG rate.',
            'Check whether the machine can detect NG samples.',
            'Check press VP+CD OK rate (no float).',
            'Check vision for Dome damage / VP damage.',
            'Measure CAPA before vs after machine improvement.'
        ],
        'source_cells': {'title': ['5.3!B2'], 'date': ['5.3!P3'], 'purpose': ['5.3!A6'], 'content': ['5.3!A8:A11']}
    },
    'test_conditions': [
        {'condition_id': 'cond_1', 'condition_group': 'Pickup VP+CD machine condition sweep',
         'line': 'Sub 1', 'process': 'Sub 1 Pickup', 'changed_factor': 'Machine repair / pusher setting',
         'before_value': 'Before', 'after_value': 'After repair MC / pusher variants',
         'unit': None, 'machine': 'New Sub1 vision machine', 'jig': 'Pusher', 'material_lot': None,
         'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None,
         'bond_amount': None, 'uv_energy': None,
         'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!B15:I20']},
        {'condition_id': 'cond_2', 'condition_group': 'Press VP+CD machine condition sweep',
         'line': 'Sub 1', 'process': 'Sub 1 Press', 'changed_factor': 'Machine repair state',
         'before_value': 'Before', 'after_value': 'After repair MC',
         'unit': None, 'machine': 'New Sub1 vision machine', 'jig': None, 'material_lot': None,
         'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None,
         'bond_amount': None, 'uv_energy': None,
         'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!B23:I26']},
        {'condition_id': 'cond_3', 'condition_group': 'Vision VP+CD damage detection',
         'line': 'Sub 1', 'process': 'Sub 1 Vision', 'changed_factor': 'Machine condition / pusher',
         'before_value': 'Before', 'after_value': 'After repair MC/pusher',
         'unit': None, 'machine': 'New Sub1 vision machine', 'jig': 'Pusher', 'material_lot': None,
         'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None,
         'bond_amount': None, 'uv_energy': None,
         'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!B29:I33']},
        {'condition_id': 'cond_4', 'condition_group': 'CAPA before vs after MC improve',
         'line': 'Sub 1', 'process': 'Capa measurement', 'changed_factor': 'Machine improvement',
         'before_value': 'Before (3/27)', 'after_value': 'After improve MC (4/9)',
         'unit': 'pcs/10h', 'machine': 'New Sub1 vision machine', 'jig': None, 'material_lot': None,
         'supplier': None, 'dry_time_sec': None, 'temperature': None, 'pressure': None,
         'bond_amount': None, 'uv_energy': None,
         'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!B40:O41']},
    ],
    'results': [
        # Pickup
        {'result_id': 'res_p1', 'condition_id': 'cond_1', 'measurement_type': 'Sub1 Pickup',
         'condition_group': 'Before', 'date': '2024-03-22', 'line': 'Sub 1',
         'input_count': 28, 'ok_count': 26, 'ng_count': 2, 'ng_rate_decimal': 0.071, 'ng_rate_percent': 7.1,
         'metric_name': 'Pickup NG rate', 'metric_value': 7.1, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'Pickup NG': 2}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!E15:I15']},
        {'result_id': 'res_p2', 'condition_id': 'cond_1', 'measurement_type': 'Sub1 Pickup',
         'condition_group': 'After repair MC', 'date': '2024-03-22', 'line': 'Sub 1',
         'input_count': 84, 'ok_count': 81, 'ng_count': 3, 'ng_rate_decimal': 0.036, 'ng_rate_percent': 3.6,
         'metric_name': 'Pickup NG rate', 'metric_value': 3.6, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'Pickup NG': 3}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!E16:I16']},
        {'result_id': 'res_p3', 'condition_id': 'cond_1', 'measurement_type': 'Sub1 Pickup',
         'condition_group': 'Use pusher', 'date': '2024-03-23', 'line': 'Sub 1',
         'input_count': 80, 'ok_count': 78, 'ng_count': 2, 'ng_rate_decimal': 0.025, 'ng_rate_percent': 2.5,
         'metric_name': 'Pickup NG rate', 'metric_value': 2.5, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'Pickup NG': 2}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!E17:I17']},
        {'result_id': 'res_p4', 'condition_id': 'cond_1', 'measurement_type': 'Sub1 Pickup',
         'condition_group': "Don't use pusher", 'date': '2024-03-23', 'line': 'Sub 1',
         'input_count': 80, 'ok_count': 79, 'ng_count': 1, 'ng_rate_decimal': 0.012, 'ng_rate_percent': 1.2,
         'metric_name': 'Pickup NG rate', 'metric_value': 1.2, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'Pickup NG': 1}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!E18:I18']},
        {'result_id': 'res_p5', 'condition_id': 'cond_1', 'measurement_type': 'Sub1 Pickup',
         'condition_group': 'Test more up pusher', 'date': '2024-03-23', 'line': 'Sub 1',
         'input_count': 80, 'ok_count': 80, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Pickup NG rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'Pickup NG': 0}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!E19:I19']},
        {'result_id': 'res_p6', 'condition_id': 'cond_1', 'measurement_type': 'Sub1 Pickup',
         'condition_group': 'After repair pusher', 'date': '2024-03-27', 'line': 'Sub 1',
         'input_count': 132, 'ok_count': 132, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Pickup NG rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'Pickup NG': 0}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!E20:I20']},
        # Press
        {'result_id': 'res_pr1', 'condition_id': 'cond_2', 'measurement_type': 'Sub1 Press',
         'condition_group': 'Before', 'date': '2024-03-22', 'line': 'Sub 1',
         'input_count': 28, 'ok_count': 27, 'ng_count': 1, 'ng_rate_decimal': 0.036, 'ng_rate_percent': 3.6,
         'metric_name': 'Press NG rate', 'metric_value': 3.6, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'VP+CD float': 1}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!E23:I23']},
        {'result_id': 'res_pr2', 'condition_id': 'cond_2', 'measurement_type': 'Sub1 Press',
         'condition_group': 'After repair MC', 'date': '2024-03-22', 'line': 'Sub 1',
         'input_count': 84, 'ok_count': 75, 'ng_count': 9, 'ng_rate_decimal': 0.107, 'ng_rate_percent': 10.7,
         'metric_name': 'Press NG rate', 'metric_value': 10.7, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'VP+CD float': 9}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!E24:I24']},
        {'result_id': 'res_pr3', 'condition_id': 'cond_2', 'measurement_type': 'Sub1 Press',
         'condition_group': 'Before (Mar23)', 'date': '2024-03-23', 'line': 'Sub 1',
         'input_count': 160, 'ok_count': 160, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Press NG rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'VP+CD float': 0}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!E25:I25']},
        {'result_id': 'res_pr4', 'condition_id': 'cond_2', 'measurement_type': 'Sub1 Press',
         'condition_group': 'Before (Mar27)', 'date': '2024-03-27', 'line': 'Sub 1',
         'input_count': 132, 'ok_count': 132, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Press NG rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'VP+CD float': 0}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!E26:I26']},
        # Vision
        {'result_id': 'res_v1', 'condition_id': 'cond_3', 'measurement_type': 'Sub1 Vision',
         'condition_group': 'Before', 'date': '2024-03-22', 'line': 'Sub 1',
         'input_count': 28, 'ok_count': 22, 'ng_count': 6, 'ng_rate_decimal': 0.214, 'ng_rate_percent': 21.4,
         'metric_name': 'Vision NG rate', 'metric_value': 21.4, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'Dome Damage': 6, 'VP Damage': 0}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!E29:J29']},
        {'result_id': 'res_v2', 'condition_id': 'cond_3', 'measurement_type': 'Sub1 Vision',
         'condition_group': 'After repair MC', 'date': '2024-03-22', 'line': 'Sub 1',
         'input_count': 84, 'ok_count': 82, 'ng_count': 2, 'ng_rate_decimal': 0.024, 'ng_rate_percent': 2.4,
         'metric_name': 'Vision NG rate', 'metric_value': 2.4, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'Dome Damage': 2, 'VP Damage': 0}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!E30:J30']},
        {'result_id': 'res_v3', 'condition_id': 'cond_3', 'measurement_type': 'Sub1 Vision',
         'condition_group': 'Before (Mar23)', 'date': '2024-03-23', 'line': 'Sub 1',
         'input_count': 160, 'ok_count': 160, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Vision NG rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'Dome Damage': 0, 'VP Damage': 0}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!E31:J31']},
        {'result_id': 'res_v4', 'condition_id': 'cond_3', 'measurement_type': 'Sub1 Vision',
         'condition_group': 'Test more up pusher', 'date': '2024-03-23', 'line': 'Sub 1',
         'input_count': 80, 'ok_count': 74, 'ng_count': 6, 'ng_rate_decimal': 0.075, 'ng_rate_percent': 7.5,
         'metric_name': 'Vision NG rate', 'metric_value': 7.5, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'Dome Damage': 6, 'VP Damage': 0}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!E32:J32']},
        {'result_id': 'res_v5', 'condition_id': 'cond_3', 'measurement_type': 'Sub1 Vision',
         'condition_group': 'After repair pusher', 'date': '2024-03-27', 'line': 'Sub 1',
         'input_count': 132, 'ok_count': 132, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Vision NG rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'Dome Damage': 0, 'VP Damage': 0}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!E33:J33']},
        # Capa
        {'result_id': 'res_c1', 'condition_id': 'cond_4', 'measurement_type': 'Capa',
         'condition_group': 'Before', 'date': '2024-03-27', 'line': 'Sub 1',
         'input_count': None, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Capa', 'metric_value': 8761, 'unit': 'pcs/10h', 'judgement': 'FAIL',
         'ng_breakdown': {}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!O40']},
        {'result_id': 'res_c2', 'condition_id': 'cond_4', 'measurement_type': 'Capa',
         'condition_group': 'After improve MC', 'date': '2024-04-09', 'line': 'Sub 1',
         'input_count': None, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Capa', 'metric_value': 13018, 'unit': 'pcs/10h', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!O41']},
    ],
    'conclusions': [
        {'conclusion_id': 'concl_1', 'topic': 'Decision',
         'statement_from_report': 'Need improve Capa of new machine (now 8761pcs/10h). After improve MC Capa of new machine is 13K/10h.',
         'normalized_interpretation': 'Pickup converges from 7.1% → 0.0% across After-repair-MC, pusher tuning, and final After-repair-pusher (3/27). Press shows transient spike to 10.7% at After-repair-MC (Mar22) but reaches 0.0% on Mar23/27 - investigate root cause. Vision dropped from 21.4% → 0.0% after MC repair on Mar22, but a regression to 7.5% (6/80 Dome Damage) appeared at "Test more up pusher" (Mar23), then 0.0% after pusher repair (Mar27). AOI brightness issue was fixed on Mar27. Capa improved from 8,761 → 13,018 pcs/10h (+48.6%), exceeding 12K target. Machine usable after improvements.',
         'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!A44:A45']},
        {'conclusion_id': 'concl_2', 'topic': 'AOI tuning',
         'statement_from_report': 'AOI is so bright so can not see NG sample VP separate or VP not enough glue => Need to adjust AOI. 27/3: Already adjust AOI can see NG sample.',
         'normalized_interpretation': 'Initial AOI brightness blocked NG-detection of VP separate and not-enough-glue; adjustment on Mar 27 enabled NG detection.',
         'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!A35:A36']},
    ],
    'troubleshooting_index': {
        'defect_name': 'Sub1 VP+CD pickup/press/vision NG (new machine)',
        'when_user_asks': ['How to validate a new Sub-1 VP+CD vision machine?', 'How to improve pickup/press/vision NG and capa on Sub-1?'],
        'suggested_checks': [
            {'hint_id': 'hint_1', 'check_item': 'Sweep pusher up/down/off across pickup runs',
             'reason': 'Pusher off 1.2%, pusher on 2.5%, more up pusher 0.0% - pusher height is the key knob.',
             'evidence_strength': 'medium', 'related_process': 'Sub 1 Pickup', 'related_part': 'Pusher',
             'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!E17:I19']},
            {'hint_id': 'hint_2', 'check_item': 'Retest press after machine repair with bigger sample',
             'reason': 'Mar22 after-repair-MC press spike 10.7% (9/84) contradicted by Mar23 0/160 and Mar27 0/132 - either MC retune or sampling artifact; needs follow-up.',
             'evidence_strength': 'medium', 'related_process': 'Sub 1 Press', 'related_part': 'Machine',
             'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!E23:I26']},
            {'hint_id': 'hint_3', 'check_item': 'Verify AOI brightness can resolve VP-separate and not-enough-glue',
             'reason': 'Original AOI too bright; Mar 27 adjustment enabled NG detection. Validate on retained NG samples.',
             'evidence_strength': 'high', 'related_process': 'Sub 1 Vision (AOI)', 'related_part': 'AOI',
             'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!A35:A36']},
            {'hint_id': 'hint_4', 'check_item': 'Re-measure capa for 10h after MC improvement',
             'reason': 'Capa moved 8,761 → 13,018 pcs/10h (+48.6%), exceeding 12K threshold.',
             'evidence_strength': 'high', 'related_process': 'Capa measurement', 'related_part': 'Machine',
             'source_file': name03, 'sheet_name': '5.3', 'source_cells': ['5.3!O40:O41']},
        ],
        'limitations': ['No same-event baseline against the legacy machine; comparisons are within-machine before/after.']
    },
    'ai_extraction_log': {
        'confidence': 0.8,
        'assumptions': ['Same-machine before/after runs treated as the comparison axis.', 'Sample sizes vary 28-160 across conditions.'],
        'warnings': ['No baseline against the previous machine, only within-machine before/after.', 'After-repair-MC press 10.7% vs. later 0.0% suggests retune; not a final state.'],
        'decision_rationale': 'Within-machine before/after: pickup 7.1%→0.0%, vision 21.4%→0.0% after final pusher repair; press converged to 0.0% by Mar23/27 despite Mar22 spike. AOI tuned on Mar27. Capa rose 8,761→13,018 pcs/10h (+48.6%). Machine usable after the full improvement set.'
    },
}
run(name03, result03, {}, {}, {})  # placeholder; we will overwrite right below by re-running with proper translations

# Build full translations for ds_03
tr_en_03 = _tr_from_result(result03)
tr_ko_03 = {
    'document': {
        'title': 'Sub1 신규 VP+CD 비전 머신 시험 리포트 BRS-161016',
        'purpose': 'Sub-1 신규 VP+CD 비전 머신의 사용 가능성 검증: pickup OK 율, NG 검출 가능성, press OK 율, 비전 damage, AOI 조정, capa.',
        'content': [
            'Pickup VP OK/NG 율 확인.', '머신의 NG 샘플 검출 가능 여부 확인.',
            'Press VP+CD OK 율(float 없음) 확인.', 'Dome damage / VP damage 비전 확인.',
            '머신 개선 전후 CAPA 측정.'
        ]
    },
    'conclusions': {
        'concl_1': {'topic': '결정',
                    'statement_from_report': '현재 8761pcs/10h - capa 개선 필요. 개선 후 13K/10h 달성.',
                    'normalized_interpretation': 'Pickup 7.1% → 0.0%(After repair MC, pusher 튜닝, 3/27 pusher 수리 종합). Press는 3/22 After-repair-MC에서 10.7%로 일시 악화 후 3/23·3/27에 0.0% 도달 - 원인 추적 필요. Vision 21.4% → 0.0%(3/22 MC 수리), 그러나 3/23 "Test more up pusher"에서 Dome Damage 7.5%(6/80)로 재발, 3/27 pusher 수리 후 0.0%. AOI는 3/27 조정 완료. Capa 8,761 → 13,018 pcs/10h(+48.6%), 12K 목표 초과. 개선 후 머신 사용 가능.'},
        'concl_2': {'topic': 'AOI 조정',
                    'statement_from_report': 'AOI 과밝아 VP separate / 본드 부족 NG 미검출 → AOI 조정 필요. 3/27 조정 완료.',
                    'normalized_interpretation': '초기 AOI 과밝음으로 VP-separate, 본드 부족 NG 미검출; 3/27 조정으로 검출 가능.'},
    },
    'hints': {
        'hint_1': {'check_item': 'Pickup 시 pusher 위치(상/하/Off) 스윕',
                   'reason': 'Pusher Off 1.2%, On 2.5%, More-up 0.0% - pusher 높이가 핵심.'},
        'hint_2': {'check_item': '머신 수리 후 큰 표본으로 press 재시험',
                   'reason': '3/22 After-repair-MC 10.7%(9/84) vs 3/23 0/160·3/27 0/132 - 재튜닝 또는 표본 이슈, 후속 필요.'},
        'hint_3': {'check_item': 'AOI 밝기 조정으로 VP-separate / 본드 부족 검출 가능 여부 검증',
                   'reason': '초기 AOI 과밝음; 3/27 조정 후 NG 검출 가능. 보관 NG 샘플로 검증.'},
        'hint_4': {'check_item': '머신 개선 후 10h capa 재측정',
                   'reason': 'Capa 8,761 → 13,018 pcs/10h(+48.6%), 12K 기준 초과.'},
    },
    'log': {
        'assumptions': ['동일 머신 before/after를 비교 축으로 간주.', '조건별 표본 28-160pcs 범위.'],
        'warnings': ['이전 머신과의 동일 이벤트 baseline 없음, 머신 내부 before/after만.', 'After-repair-MC press 10.7% → 0.0% 추이가 최종 상태가 아닐 수 있음.'],
        'decision_rationale': '머신 내부 before/after: pickup 7.1%→0.0%, vision 21.4%→0.0%(최종 pusher 수리), press는 3/22 스파이크 후 3/23·3/27 0.0% 수렴. AOI 3/27 조정. Capa 8,761→13,018 pcs/10h(+48.6%). 전체 개선 적용 후 사용 가능.'
    }
}
tr_vi_03 = {
    'document': {
        'title': 'BÁO CÁO TEST MÁY VISION VP+CD MỚI TẠI SUB 1 BRS-161016',
        'purpose': 'Đánh giá máy vision VP+CD mới Sub-1: tỉ lệ pickup OK, khả năng phát hiện NG, tỉ lệ press OK, damage trên vision, chỉnh AOI và capa.',
        'content': [
            'Check tỉ lệ pickup VP OK/NG.', 'Check máy có detect được NG sample.',
            'Check tỉ lệ press VP+CD OK (không float).', 'Check vision Dome damage / VP damage.',
            'Đo CAPA trước/sau khi cải tiến máy.'
        ]
    },
    'conclusions': {
        'concl_1': {'topic': 'Quyết định',
                    'statement_from_report': 'Hiện 8761pcs/10h - cần cải thiện capa. Sau cải thiện đạt 13K/10h.',
                    'normalized_interpretation': 'Pickup 7.1% → 0.0% (sau repair MC, pusher tune, 3/27 sửa pusher). Press tăng đột biến 10.7% ngày 3/22 After-repair-MC nhưng về 0.0% ngày 3/23·3/27 - cần truy root cause. Vision 21.4% → 0.0% sau MC repair ngày 3/22, nhưng quay lại 7.5% (6/80 Dome Damage) tại "Test more up pusher" ngày 3/23, rồi 0.0% sau khi sửa pusher 3/27. AOI tune ngày 3/27. Capa 8,761 → 13,018 pcs/10h (+48.6%), vượt mục tiêu 12K. Máy dùng được sau khi cải tiến.'},
        'concl_2': {'topic': 'Chỉnh AOI',
                    'statement_from_report': 'AOI quá sáng nên không thấy NG VP separate / không đủ keo → cần chỉnh. 3/27 đã chỉnh xong.',
                    'normalized_interpretation': 'AOI ban đầu quá sáng nên không detect VP-separate, không-đủ-keo; chỉnh ngày 3/27 phát hiện được.'},
    },
    'hints': {
        'hint_1': {'check_item': 'Sweep vị trí pusher (lên/xuống/tắt) khi pickup',
                   'reason': 'Pusher off 1.2%, on 2.5%, more-up 0.0% - chiều cao pusher là yếu tố chính.'},
        'hint_2': {'check_item': 'Test lại press với sample lớn sau khi sửa máy',
                   'reason': '3/22 After-repair-MC 10.7% (9/84) vs 3/23 0/160 và 3/27 0/132 - re-tune hoặc artifact sampling, cần follow-up.'},
        'hint_3': {'check_item': 'Verify AOI có detect được VP-separate / không-đủ-keo',
                   'reason': 'AOI ban đầu quá sáng; 3/27 chỉnh xong cho phép detect. Verify với NG sample lưu.'},
        'hint_4': {'check_item': 'Đo lại capa 10h sau khi cải tiến máy',
                   'reason': 'Capa 8,761 → 13,018 pcs/10h (+48.6%), vượt ngưỡng 12K.'},
    },
    'log': {
        'assumptions': ['So sánh before/after cùng máy.', 'Sample size khác nhau 28-160pcs.'],
        'warnings': ['Không có baseline máy cũ, chỉ before/after trong máy mới.', '10.7% → 0.0% press chứng tỏ có re-tune; có thể không phải state cuối.'],
        'decision_rationale': 'Before/after cùng máy: pickup 7.1%→0.0%, vision 21.4%→0.0% (cuối cùng pusher repair); press 3/22 spike rồi 3/23·3/27 = 0.0%. AOI tune 3/27. Capa 8,761→13,018 pcs/10h (+48.6%). Máy dùng được sau khi áp dụng đủ cải tiến.'
    }
}
# Re-run with translations to overwrite the placeholder
run(name03, result03, tr_ko_03, tr_en_03, tr_vi_03)


# ===== DS 04 =====
name04 = '25. MSU-L20S15-07 Report test check lot separate VP+CD date 15.4.2025'
result04 = {
    'schema_version': '0.1',
    'document': {
        'document_id': '', 'source_file': name04, 'source_sheet': 'Test MC aging',
        'title': 'REPORT TEST CHECK LOT SEPARATE VP+CD BRS-201507DT (MSU-L20S15-07)',
        'model': 'MSU-L20S15-07 / BRS-201507DT', 'report_date': '2025-04-14', 'department': 'ME',
        'marker': 'Le', 'line': '',
        'report_type': 'normal_comparison',
        'primary_defect': {'canonical_name': 'VP+CD Separation', 'aliases_in_document': ['VP+CD separate', 'NG separate VP + CD']},
        'related_defects': ['NG Hearing Noise', 'NG Hearing Touch', 'NG Sigma SPL+THD', 'NG Sigma THD'],
        'parts': ['VP', 'CD'],
        'processes': ['Aging M/C', 'Sigma', 'Hearing', 'Decap'],
        'purpose': 'Determine whether the aging machine can detect VP/CD separation.',
        'content': [
            'Test 1: OK and separate VP+CD groups (20pcs each) through aging→sigma→hearing→separation check.',
            'Test 2: CD lots dated 17/3, 23/3, 14/4 (50pcs each) through same flow.'
        ],
        'source_cells': {'title': ['Test MC aging!B2'], 'date': ['Test MC aging!Q3'], 'purpose': ['Test MC aging!A6'], 'content': ['Test MC aging!A8:A16']}
    },
    'test_conditions': [
        {'condition_id': 'cond_1', 'condition_group': 'VP+CD OK vs Separate (Test 1)',
         'line': '', 'process': 'Aging + Sigma + Hearing', 'changed_factor': 'Sample VP+CD condition',
         'before_value': 'VP+CD OK', 'after_value': 'VP+CD separate',
         'unit': None, 'machine': 'Aging M/C', 'jig': None, 'material_lot': None, 'supplier': None,
         'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': name04, 'sheet_name': 'Test MC aging', 'source_cells': ['Test MC aging!B21:Q24']},
        {'condition_id': 'cond_2', 'condition_group': 'CD lots 17/3 vs 23/3 vs 14/4',
         'line': '', 'process': 'Aging + Sigma + Hearing', 'changed_factor': 'CD lot date',
         'before_value': '23/3 lot (recent OK lot)', 'after_value': '17/3 + 14/4 lots',
         'unit': None, 'machine': 'Aging M/C', 'jig': None,
         'material_lot': '17/3, 23/3, 14/4', 'supplier': None,
         'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': name04, 'sheet_name': 'Test MC aging', 'source_cells': ['Test MC aging!B25:Q30']},
    ],
    'results': [
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Function (Sigma+Hearing)',
         'condition_group': 'VP+CD OK', 'date': '2025-04-15', 'line': '',
         'input_count': 20, 'ok_count': 19, 'ng_count': 1, 'ng_rate_decimal': 0.05, 'ng_rate_percent': 5.0,
         'metric_name': 'Hearing NG rate', 'metric_value': 5.0, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'NG Sigma SPL': 0, 'NG Sigma THD': 0, 'NG Sigma SPL+THD': 0, 'NG Sigma SPL+THD+F0': 0,
                          'NG Hearing Noise': 1, 'NG Hearing Touch': 0},
         'source_file': name04, 'sheet_name': 'Test MC aging', 'source_cells': ['Test MC aging!E21:P21']},
        {'result_id': 'res_2', 'condition_id': 'cond_1', 'measurement_type': 'Function (Sigma+Hearing)',
         'condition_group': 'VP+CD separate', 'date': '2025-04-15', 'line': '',
         'input_count': 20, 'ok_count': 1, 'ng_count': 19, 'ng_rate_decimal': 0.95, 'ng_rate_percent': 95.0,
         'metric_name': 'Hearing+Sigma NG rate', 'metric_value': 95.0, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG Sigma SPL': 0, 'NG Sigma THD': 0, 'NG Sigma SPL+THD': 1, 'NG Sigma SPL+THD+F0': 0,
                          'NG Hearing Noise': 18, 'NG Hearing Touch': 0},
         'source_file': name04, 'sheet_name': 'Test MC aging', 'source_cells': ['Test MC aging!E23:P23']},
        {'result_id': 'res_3', 'condition_id': 'cond_2', 'measurement_type': 'Function (Sigma+Hearing)',
         'condition_group': 'CD lot 17/3', 'date': '2025-04-15', 'line': '',
         'input_count': 46, 'ok_count': 6, 'ng_count': 40, 'ng_rate_decimal': 0.87, 'ng_rate_percent': 87.0,
         'metric_name': 'Hearing NG rate', 'metric_value': 87.0, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG Hearing Noise': 39, 'NG Hearing Touch': 1},
         'source_file': name04, 'sheet_name': 'Test MC aging', 'source_cells': ['Test MC aging!E25:P25']},
        {'result_id': 'res_4', 'condition_id': 'cond_2', 'measurement_type': 'Function (Sigma+Hearing)',
         'condition_group': 'CD lot 23/3', 'date': '2025-04-15', 'line': '',
         'input_count': 50, 'ok_count': 48, 'ng_count': 2, 'ng_rate_decimal': 0.04, 'ng_rate_percent': 4.0,
         'metric_name': 'Hearing NG rate', 'metric_value': 4.0, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'NG Hearing Noise': 2, 'NG Hearing Touch': 0},
         'source_file': name04, 'sheet_name': 'Test MC aging', 'source_cells': ['Test MC aging!E27:P27']},
        {'result_id': 'res_5', 'condition_id': 'cond_2', 'measurement_type': 'Function (Sigma+Hearing)',
         'condition_group': 'CD lot 14/4', 'date': '2025-04-15', 'line': '',
         'input_count': 51, 'ok_count': 47, 'ng_count': 4, 'ng_rate_decimal': 0.078, 'ng_rate_percent': 7.8,
         'metric_name': 'Hearing NG rate', 'metric_value': 7.8, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'NG Hearing Noise': 4, 'NG Hearing Touch': 0},
         'source_file': name04, 'sheet_name': 'Test MC aging', 'source_cells': ['Test MC aging!E29:P29']},
        # Separation confirmation table
        {'result_id': 'res_sep1', 'condition_id': 'cond_1', 'measurement_type': 'Decap Separate VP+CD',
         'condition_group': 'VP+CD OK arm - NG noise samples', 'date': '2025-04-15', 'line': '',
         'input_count': 1, 'ok_count': 1, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'NG separate VP+CD found', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'VP+CD Separation': 0}, 'source_file': name04, 'sheet_name': 'Test MC aging', 'source_cells': ['Test MC aging!E33:I33']},
        {'result_id': 'res_sep2', 'condition_id': 'cond_1', 'measurement_type': 'Decap Separate VP+CD',
         'condition_group': 'VP+CD separate arm - NG noise samples', 'date': '2025-04-15', 'line': '',
         'input_count': 18, 'ok_count': 0, 'ng_count': 18, 'ng_rate_decimal': 1.0, 'ng_rate_percent': 100.0,
         'metric_name': 'NG separate VP+CD found', 'metric_value': 100.0, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'VP+CD Separation': 18}, 'source_file': name04, 'sheet_name': 'Test MC aging', 'source_cells': ['Test MC aging!E36:I36']},
        {'result_id': 'res_sep3', 'condition_id': 'cond_2', 'measurement_type': 'Decap Separate VP+CD',
         'condition_group': 'Lot 17/3 NG noise', 'date': '2025-04-15', 'line': '',
         'input_count': 39, 'ok_count': 0, 'ng_count': 39, 'ng_rate_decimal': 1.0, 'ng_rate_percent': 100.0,
         'metric_name': 'NG separate VP+CD found', 'metric_value': 100.0, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'VP+CD Separation': 39}, 'source_file': name04, 'sheet_name': 'Test MC aging', 'source_cells': ['Test MC aging!E39:I39']},
        {'result_id': 'res_sep4', 'condition_id': 'cond_2', 'measurement_type': 'Decap Separate VP+CD',
         'condition_group': 'Lot 23/3 NG noise', 'date': '2025-04-15', 'line': '',
         'input_count': 2, 'ok_count': 0, 'ng_count': 2, 'ng_rate_decimal': 1.0, 'ng_rate_percent': 100.0,
         'metric_name': 'NG separate VP+CD found', 'metric_value': 100.0, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'VP+CD Separation': 2}, 'source_file': name04, 'sheet_name': 'Test MC aging', 'source_cells': ['Test MC aging!E41:I41']},
        {'result_id': 'res_sep5', 'condition_id': 'cond_2', 'measurement_type': 'Decap Separate VP+CD',
         'condition_group': 'Lot 14/4 NG noise', 'date': '2025-04-15', 'line': '',
         'input_count': 4, 'ok_count': 0, 'ng_count': 4, 'ng_rate_decimal': 1.0, 'ng_rate_percent': 100.0,
         'metric_name': 'NG separate VP+CD found', 'metric_value': 100.0, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'VP+CD Separation': 4}, 'source_file': name04, 'sheet_name': 'Test MC aging', 'source_cells': ['Test MC aging!E43:I43']},
    ],
    'conclusions': [
        {'conclusion_id': 'concl_1', 'topic': 'Decision',
         'statement_from_report': 'Decision section in source is empty.',
         'normalized_interpretation': 'Aging-then-hearing flow can flag VP+CD separation: VP+CD-separate arm 95.0% (19/20) vs OK arm 5.0% (1/20), ratio 19.0x, 1800% worse. Lot 17/3 87.0% (40/46) vs Lot 23/3 4.0% (2/50) = 21.75x, 2075% worse. All NG-noise samples decap-confirmed VP+CD separation 100% (61/61 across arms) - hearing-noise NG after aging is a strong indicator of VP+CD separation. Lot 14/4 7.8% (4/51) elevated vs 23/3 4.0% (1.95x, 95% worse).',
         'source_file': name04, 'sheet_name': 'Test MC aging', 'source_cells': ['Test MC aging!A45']},
        {'conclusion_id': 'concl_2', 'topic': 'CD lot risk ranking',
         'statement_from_report': 'CD lots compared: 17/3, 23/3, 14/4',
         'normalized_interpretation': 'Risk ranking by hearing NG: lot 17/3 (87.0%) >> lot 14/4 (7.8%) > lot 23/3 (4.0%). 17/3 lot must be quarantined; 14/4 needs further evaluation.',
         'source_file': name04, 'sheet_name': 'Test MC aging', 'source_cells': ['Test MC aging!B25:P29']},
    ],
    'troubleshooting_index': {
        'defect_name': 'VP+CD Separation',
        'when_user_asks': ['Can aging machine reveal VP+CD separation?', 'How to screen CD lots for separation risk?'],
        'suggested_checks': [
            {'hint_id': 'hint_1', 'check_item': 'Run aging→sigma→hearing flow and decap NG-noise samples',
             'reason': 'VP+CD-separate arm hearing NG 95.0% vs OK arm 5.0% = 19x (1800% worse); 100% of NG-noise samples confirmed VP+CD separation on decap.',
             'evidence_strength': 'high', 'related_process': 'Aging + Sigma + Hearing + Decap', 'related_part': 'VP/CD',
             'source_file': name04, 'sheet_name': 'Test MC aging', 'source_cells': ['Test MC aging!E21:P36']},
            {'hint_id': 'hint_2', 'check_item': 'Quarantine high-risk CD lots based on hearing NG rate after aging',
             'reason': 'Lot 17/3 87.0% vs Lot 23/3 4.0% = 21.75x (2075% worse) - clear lot-level signal.',
             'evidence_strength': 'high', 'related_process': 'CD lot screening', 'related_part': 'CD',
             'source_file': name04, 'sheet_name': 'Test MC aging', 'source_cells': ['Test MC aging!B25:P29']},
            {'hint_id': 'hint_3', 'check_item': 'Watch recent CD lots even with low NG rate',
             'reason': 'Lot 14/4 hearing NG 7.8% (1.95x baseline 4.0%); decap shows 100% confirmed separation among NG samples.',
             'evidence_strength': 'medium', 'related_process': 'CD lot screening', 'related_part': 'CD',
             'source_file': name04, 'sheet_name': 'Test MC aging', 'source_cells': ['Test MC aging!B29:P29']},
        ],
        'limitations': ['Test1 OK arm sample only 20pcs; report Decision section blank.']
    },
    'ai_extraction_log': {
        'confidence': 0.85,
        'assumptions': ['Date "15/4/2025" used as report date.', 'Lot 23/3 used as low-risk baseline for lot comparison.'],
        'warnings': ['Decision section in original is blank; conclusions are AI-derived.'],
        'decision_rationale': 'Aging+hearing successfully screens VP+CD separation: separate vs OK = 19x worse (1800%); lot 17/3 vs 23/3 = 21.75x worse (2075%); 100% of NG-noise samples confirmed separation in decap. Lot 14/4 7.8% (1.95x baseline) warrants follow-up.'
    },
}
tr_en_04 = _tr_from_result(result04)
tr_ko_04 = {
    'document': {
        'title': 'MSU-L20S15-07 CD 로트별 VP+CD Separate 검출 시험 리포트',
        'purpose': 'Aging 머신으로 VP/CD separation 검출 가능 여부 확인.',
        'content': [
            'Test 1: VP+CD OK 및 separate 각 20pcs를 aging→sigma→hearing→separation 확인.',
            'Test 2: CD 로트 17/3, 23/3, 14/4 각 50pcs 동일 흐름.'
        ]
    },
    'conclusions': {
        'concl_1': {'topic': '결정', 'statement_from_report': '원본 Decision 섹션 비어있음.',
                    'normalized_interpretation': 'Aging-Hearing 흐름이 VP+CD separation을 검출: VP+CD-separate 95.0%(19/20) vs OK 5.0%(1/20) = 19.0배, 1800% 악화. 로트 17/3 87.0%(40/46) vs 로트 23/3 4.0%(2/50) = 21.75배, 2075% 악화. NG-noise 샘플 decap에서 VP+CD separation 100%(61/61) 확인. 로트 14/4 7.8%(4/51), 23/3 대비 1.95배(95% 악화).'},
        'concl_2': {'topic': 'CD 로트 리스크 순위',
                    'statement_from_report': '비교 CD 로트: 17/3, 23/3, 14/4',
                    'normalized_interpretation': 'Hearing NG 기준 리스크: 17/3(87.0%) >> 14/4(7.8%) > 23/3(4.0%). 17/3 격리 필요, 14/4 추가 평가.'},
    },
    'hints': {
        'hint_1': {'check_item': 'Aging→sigma→hearing 후 NG-noise 샘플 decap',
                   'reason': 'VP+CD-separate hearing NG 95.0% vs OK 5.0% = 19배(1800% 악화); NG-noise 샘플 100% VP+CD separation 확인.'},
        'hint_2': {'check_item': 'Aging 후 hearing NG로 고위험 CD 로트 격리',
                   'reason': '17/3 87.0% vs 23/3 4.0% = 21.75배(2075% 악화) - 로트 단위 시그널 명확.'},
        'hint_3': {'check_item': 'NG율이 낮아도 최근 CD 로트 추가 관찰',
                   'reason': '14/4 NG 7.8%(베이스라인 4.0% 대비 1.95배), NG 샘플 100% separation 확인.'},
    },
    'log': {
        'assumptions': ['리포트 일자 2025-04-15 사용.', '로트 23/3을 저위험 베이스라인으로 설정.'],
        'warnings': ['원본 Decision 섹션 비어있음.'],
        'decision_rationale': 'Aging+Hearing이 VP+CD separation 검출 성공: separate vs OK = 19배(1800%); 17/3 vs 23/3 = 21.75배(2075%); NG-noise 100% separation. 14/4 7.8%(1.95배) 추가 추적 필요.'
    }
}
tr_vi_04 = {
    'document': {
        'title': 'BÁO CÁO TEST CHECK LOT SEPARATE VP+CD MSU-L20S15-07',
        'purpose': 'Dùng máy aging có thể phát hiện separate VP/CD hay không.',
        'content': [
            'Test 1: VP+CD OK và separate, mỗi nhóm 20pcs, qua aging→sigma→hearing→check separate.',
            'Test 2: Lot CD ngày 17/3, 23/3, 14/4 mỗi lot 50pcs, cùng flow.'
        ]
    },
    'conclusions': {
        'concl_1': {'topic': 'Quyết định', 'statement_from_report': 'Phần Decision trong file gốc trống.',
                    'normalized_interpretation': 'Flow Aging-Hearing có thể flag VP+CD separation: nhóm VP+CD-separate 95.0% (19/20) vs OK 5.0% (1/20) = 19.0x, xấu 1800%. Lot 17/3 87.0% (40/46) vs Lot 23/3 4.0% (2/50) = 21.75x, xấu 2075%. Decap mẫu NG-noise xác nhận 100% (61/61) là VP+CD separation. Lot 14/4 7.8% (4/51), so với 23/3 = 1.95x (xấu 95%).'},
        'concl_2': {'topic': 'Xếp hạng rủi ro lot CD',
                    'statement_from_report': 'Lot so sánh: 17/3, 23/3, 14/4',
                    'normalized_interpretation': 'Theo Hearing NG: 17/3 (87.0%) >> 14/4 (7.8%) > 23/3 (4.0%). 17/3 phải cách ly; 14/4 cần đánh giá thêm.'},
    },
    'hints': {
        'hint_1': {'check_item': 'Chạy flow aging→sigma→hearing rồi decap mẫu NG-noise',
                   'reason': 'Hearing NG VP+CD-separate 95.0% vs OK 5.0% = 19x (xấu 1800%); 100% mẫu NG-noise xác nhận VP+CD separation khi decap.'},
        'hint_2': {'check_item': 'Cách ly lot CD rủi ro cao dựa trên Hearing NG sau aging',
                   'reason': 'Lot 17/3 87.0% vs 23/3 4.0% = 21.75x (xấu 2075%) - tín hiệu lot rõ.'},
        'hint_3': {'check_item': 'Theo dõi lot CD gần đây ngay cả khi NG thấp',
                   'reason': 'Lot 14/4 Hearing NG 7.8% (1.95x baseline 4.0%); decap 100% separation.'},
    },
    'log': {
        'assumptions': ['Dùng "15/4/2025" làm ngày báo cáo.', 'Lot 23/3 dùng làm baseline rủi ro thấp.'],
        'warnings': ['Phần Decision trong file trống.'],
        'decision_rationale': 'Aging+Hearing phát hiện thành công VP+CD separation: separate vs OK = 19x xấu (1800%); 17/3 vs 23/3 = 21.75x xấu (2075%); 100% mẫu NG-noise xác nhận separation. Lot 14/4 7.8% (1.95x baseline) cần follow-up.'
    }
}
run(name04, result04, tr_ko_04, tr_en_04, tr_vi_04)


# ===== DS 05 =====
name05 = '25. TIU C11-20  Report test Plate supplier Yousteel NG difference with standard 100% 2026.02.24'
result05 = {
    'schema_version': '0.1',
    'document': {
        'document_id': '', 'source_file': name05, 'source_sheet': 'Test',
        'title': 'REPORT TEST PLATE SUPPLIER YOUSTEEL NG DIFFERENCE WITH STANDARD 100% C11-20',
        'model': 'TIU C11-20', 'report_date': '2026-02-24', 'department': 'ME',
        'marker': 'Trung', 'line': '',
        'report_type': 'normal_comparison',
        'primary_defect': {'canonical_name': 'MG+PT Separation', 'aliases_in_document': ['NG MG+PT separate', 'MG+PT separate']},
        'related_defects': ['Over Glue', 'Not Dry Glue', 'NG bond spread'],
        'parts': ['Plate (PT)', 'MG', 'Bond'],
        'processes': ['Sub 2 process', 'Decap bonding', 'Drop test (Auto/Manual)', 'Tension', 'AI bonding inspection'],
        'purpose': 'Test plates from Yousteel supplier that show NG-difference with standard 100% to confirm usability.',
        'content': [
            'Make semi sub 2 and check NG process.',
            'Decap to check bond PT+MG.',
            'Drop test auto/manual.',
            'Tension.',
            'Check whether AI bonding detection works correctly.'
        ],
        'source_cells': {'title': ['Test!B2'], 'date': ['Test!N3'], 'purpose': ['Test!A6'], 'content': ['Test!A8:A13']}
    },
    'test_conditions': [
        {'condition_id': 'cond_1', 'condition_group': 'Yousteel PT vs Normal PT',
         'line': '', 'process': 'Sub 2 process', 'changed_factor': 'Plate supplier (Yousteel vs Standard)',
         'before_value': 'Normal', 'after_value': 'Test PT (Yousteel)',
         'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': 'Yousteel',
         'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': name05, 'sheet_name': 'Test', 'source_cells': ['Test!B19:N22']},
        {'condition_id': 'cond_2', 'condition_group': 'Decap + Drop test on Yousteel PT',
         'line': '', 'process': 'Decap + Drop test', 'changed_factor': 'Plate supplier',
         'before_value': None, 'after_value': 'Test PT (Yousteel)',
         'unit': None, 'machine': None, 'jig': None, 'material_lot': None, 'supplier': 'Yousteel',
         'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': name05, 'sheet_name': 'Test', 'source_cells': ['Test!B24:N25']},
    ],
    'results': [
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Sub 2 NG Visual',
         'condition_group': 'Test PT (Yousteel)', 'date': '2026-02-24', 'line': '',
         'input_count': 100, 'ok_count': 94, 'ng_count': 6, 'ng_rate_decimal': 0.06, 'ng_rate_percent': 6.0,
         'metric_name': 'NG visual rate', 'metric_value': 6.0, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'Over Glue': 0, 'MG+PT Separation': 6, 'Not Dry Glue': 0},
         'source_file': name05, 'sheet_name': 'Test', 'source_cells': ['Test!E19:K19']},
        {'result_id': 'res_2', 'condition_id': 'cond_1', 'measurement_type': 'Sub 2 NG Visual',
         'condition_group': 'Normal (baseline)', 'date': '2026-02-24', 'line': '',
         'input_count': 200, 'ok_count': 200, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'NG visual rate', 'metric_value': 0.0, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {'Over Glue': 0, 'MG+PT Separation': 0, 'Not Dry Glue': 0},
         'source_file': name05, 'sheet_name': 'Test', 'source_cells': ['Test!E21:K21']},
        {'result_id': 'res_3', 'condition_id': 'cond_2', 'measurement_type': 'Decap Bonding',
         'condition_group': 'Test PT (Yousteel)', 'date': '2026-02-24', 'line': '',
         'input_count': 8, 'ok_count': 0, 'ng_count': 8, 'ng_rate_decimal': 1.0, 'ng_rate_percent': 100.0,
         'metric_name': 'NG decap bonding rate', 'metric_value': 100.0, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'Not Dry Glue': 8, 'NG bond spread': 8},
         'source_file': name05, 'sheet_name': 'Test', 'source_cells': ['Test!E24:I24']},
        {'result_id': 'res_4', 'condition_id': 'cond_2', 'measurement_type': 'Drop test Manual',
         'condition_group': 'Test PT (Yousteel)', 'date': '2026-02-24', 'line': '',
         'input_count': 8, 'ok_count': 7, 'ng_count': 1, 'ng_rate_decimal': 0.125, 'ng_rate_percent': 12.5,
         'metric_name': 'NG drop rate', 'metric_value': 12.5, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'MG+PT Separation': 1},
         'source_file': name05, 'sheet_name': 'Test', 'source_cells': ['Test!E25:I25']},
    ],
    'conclusions': [
        {'conclusion_id': 'concl_1', 'topic': 'Decision (data observations)',
         'statement_from_report': 'Result decap bonding NG 8/8pcs and result drop test manual NG 1/8 type MG+PT separate. Result check NG process 6/100 (6%) NG MG+PT separate.',
         'normalized_interpretation': 'Sub2 process NG with Yousteel PT 6.0% (6/100) vs Normal 0.0% (0/200); baseline = 0, multiplicative ratio undefined - test arm absolute-worse by 6.0 percentage points. All Sub2 NG are MG+PT separation. Decap bonding 8/8 NG (Not Dry Glue + bond spread). Drop test manual 1/8 NG, type MG+PT separation. Yousteel PT not acceptable as-is.',
         'source_file': name05, 'sheet_name': 'Test', 'source_cells': ['Test!A27:A28']},
    ],
    'troubleshooting_index': {
        'defect_name': 'MG+PT Separation (Plate supplier change)',
        'when_user_asks': ['Can Yousteel plate be used vs standard 100% plate?', 'How to verify a new plate supplier?'],
        'suggested_checks': [
            {'hint_id': 'hint_1', 'check_item': 'Run Sub2 NG process compare new-supplier vs normal supplier',
             'reason': 'Yousteel PT Sub2 NG 6.0% (6/100, all MG+PT separation) vs Normal 0.0% (0/200) - baseline 0% means ratio undefined; absolute +6.0pp worse.',
             'evidence_strength': 'high', 'related_process': 'Sub 2 process', 'related_part': 'PT (plate)',
             'source_file': name05, 'sheet_name': 'Test', 'source_cells': ['Test!E19:K21']},
            {'hint_id': 'hint_2', 'check_item': 'Decap bonding 8pcs and inspect bond spread / dry state',
             'reason': 'Yousteel PT decap 8/8 NG: NG bond not dry and bond spread NG.',
             'evidence_strength': 'high', 'related_process': 'Decap bonding', 'related_part': 'Bond / PT',
             'source_file': name05, 'sheet_name': 'Test', 'source_cells': ['Test!E24:I24']},
            {'hint_id': 'hint_3', 'check_item': 'Drop test manual on new-supplier plate',
             'reason': 'Yousteel PT drop test manual 1/8 NG, type MG+PT separation - matches Sub2 failure mode.',
             'evidence_strength': 'medium', 'related_process': 'Drop test Manual', 'related_part': 'MG+PT bond',
             'source_file': name05, 'sheet_name': 'Test', 'source_cells': ['Test!E25:I25']},
        ],
        'limitations': ['Normal baseline 200/200 OK = 0%; multiplicative ratio undefined, must use absolute delta.', 'AI bonding detection part of plan but no detection rate row in TSV.']
    },
    'ai_extraction_log': {
        'confidence': 0.85,
        'assumptions': ['Date 2026-02-24 from filename.', 'Sub2 NG broken into MG+PT separation only.'],
        'warnings': ['Baseline NG rate 0.0% (200/200) - multiplicative ratio undefined; switched to absolute pp delta.', 'AI bonding detection rate not present in TSV.'],
        'decision_rationale': 'Yousteel PT Sub2 NG 6.0% (6/100, all MG+PT separation) vs Normal 0.0% (0/200) - absolute +6.0pp worse; decap 100% NG (8/8 bond-spread/not-dry); drop manual 1/8 NG = MG+PT separation. Yousteel PT not yet usable.'
    },
}
tr_en_05 = _tr_from_result(result05)
tr_ko_05 = {
    'document': {
        'title': 'C11-20 Plate 공급사 Yousteel NG 차이 시험 리포트 (표준 100%)',
        'purpose': '표준 100% 대비 NG 차이가 있는 Yousteel 플레이트 사용 가능 여부 시험.',
        'content': [
            'Semi Sub2 제작 후 NG 공정 확인.',
            'Decap으로 PT+MG 본드 확인.',
            'Drop test (Auto/Manual).',
            'Tension.',
            'AI bonding 검출 정확성 확인.'
        ]
    },
    'conclusions': {
        'concl_1': {'topic': '결정 (데이터 관찰)',
                    'statement_from_report': 'Decap bonding 8/8 NG, drop manual 1/8 NG(MG+PT separate), NG process 6/100(6%) MG+PT separate.',
                    'normalized_interpretation': 'Yousteel PT Sub2 NG 6.0%(6/100) vs Normal 0.0%(0/200); baseline 0%로 곱셈 비율 미정의 - 절대 +6.0pp 악화. Sub2 NG 모두 MG+PT separation. Decap 8/8 NG(Not Dry Glue + bond spread). Drop manual 1/8 NG MG+PT separation. Yousteel PT 현 상태 사용 불가.'}
    },
    'hints': {
        'hint_1': {'check_item': '신공급사 Sub2 NG 공정과 normal 공급사 비교',
                   'reason': 'Yousteel PT Sub2 NG 6.0%(6/100, 전부 MG+PT separation) vs Normal 0.0%(0/200) - 절대 +6.0pp 악화.'},
        'hint_2': {'check_item': '8pcs decap bonding으로 bond spread/건조 상태 확인',
                   'reason': 'Yousteel PT decap 8/8 NG: 본드 미건조 + bond spread NG.'},
        'hint_3': {'check_item': '신공급사 플레이트 drop test manual',
                   'reason': 'Yousteel PT drop manual 1/8 NG MG+PT separation - Sub2 실패 모드와 동일.'}
    },
    'log': {
        'assumptions': ['파일명에서 2026-02-24 사용.', 'Sub2 NG는 MG+PT separation으로만 구성.'],
        'warnings': ['Baseline 0.0%(200/200)로 비율 미정의; 절대 pp 차이 사용.', 'AI bonding 검출률 행 없음.'],
        'decision_rationale': 'Yousteel PT Sub2 NG 6.0%(6/100) vs Normal 0.0%(0/200) - 절대 +6.0pp 악화; decap 8/8 NG; drop manual 1/8 NG = MG+PT separation. Yousteel PT 사용 불가.'
    }
}
tr_vi_05 = {
    'document': {
        'title': 'BÁO CÁO TEST PLATE SUPPLIER YOUSTEEL NG KHÁC STANDARD 100% C11-20',
        'purpose': 'Test plate Yousteel có NG khác standard 100% có thể dùng hay không.',
        'content': [
            'Làm semi sub 2 và check NG process.',
            'Decap check bond PT+MG.',
            'Drop test Auto/Manual.',
            'Tension.',
            'Check AI bonding detect chính xác hay không.'
        ]
    },
    'conclusions': {
        'concl_1': {'topic': 'Quyết định (quan sát dữ liệu)',
                    'statement_from_report': 'Decap bonding NG 8/8, drop manual NG 1/8 (MG+PT separate), NG process 6/100 (6%) MG+PT separate.',
                    'normalized_interpretation': 'Yousteel PT NG Sub2 6.0% (6/100) vs Normal 0.0% (0/200); baseline 0% nên tỉ lệ nhân không xác định - tuyệt đối xấu hơn +6.0pp. Tất cả NG Sub2 là MG+PT separation. Decap 8/8 NG (Not Dry Glue + bond spread). Drop manual 1/8 NG MG+PT separation. Yousteel PT chưa dùng được.'}
    },
    'hints': {
        'hint_1': {'check_item': 'So sánh NG process Sub2 supplier mới với supplier bình thường',
                   'reason': 'Yousteel PT Sub2 NG 6.0% (6/100, 100% MG+PT separation) vs Normal 0.0% (0/200) - tuyệt đối +6.0pp xấu hơn.'},
        'hint_2': {'check_item': 'Decap bonding 8pcs, kiểm tra bond spread / trạng thái khô',
                   'reason': 'Yousteel PT decap 8/8 NG: bond chưa khô + bond spread NG.'},
        'hint_3': {'check_item': 'Drop test manual cho plate supplier mới',
                   'reason': 'Yousteel PT drop manual 1/8 NG MG+PT separation - trùng failure mode Sub2.'}
    },
    'log': {
        'assumptions': ['Dùng 2026-02-24 từ filename.', 'NG Sub2 chỉ là MG+PT separation.'],
        'warnings': ['Baseline 0.0% (200/200), tỉ lệ nhân không xác định; dùng delta tuyệt đối.', 'Không có hàng tỉ lệ phát hiện AI bonding.'],
        'decision_rationale': 'Yousteel PT Sub2 NG 6.0% (6/100) vs Normal 0.0% (0/200) - tuyệt đối +6.0pp xấu; decap 8/8 NG; drop manual 1/8 NG = MG+PT separation. Yousteel PT chưa dùng được.'
    }
}
run(name05, result05, tr_ko_05, tr_en_05, tr_vi_05)


# ===== DS 06 =====
name06 = '25. TIU C11-20  Report test UV Led 3rd date 27.12.2025'
result06 = {
    'schema_version': '0.1',
    'document': {
        'document_id': '', 'source_file': name06, 'source_sheet': 'Test',
        'title': 'REPORT TEST UV LED 3rd TIU-C11-20',
        'model': 'TIU C11-20', 'report_date': '2025-12-27', 'department': 'ME',
        'marker': 'Thao', 'line': '',
        'report_type': 'doe_matrix',
        'primary_defect': {'canonical_name': 'NG Hearing / NG Audiobus', 'aliases_in_document': ['NG AUDIOBUS', 'NG HEARING', 'Noise', 'SPL', 'SPL+RB', 'RB']},
        'related_defects': ['NG Audiobus SPL', 'NG Audiobus SPL+RB', 'NG Audiobus RB', 'NG Hearing Noise'],
        'parts': ['UV LED 3rd', 'VP', 'Coil'],
        'processes': ['UV cure 3rd', 'Function', 'Tension'],
        'purpose': 'Test reducing/increasing UV LED 3rd Peak and Total energy and compare function vs normal lot.',
        'content': [
            'Reduce and increase UV LED 3rd energy.',
            'Make sample, check tension (5-10pcs).',
            'Function test 100pcs per condition.',
            'Compare with normal lot. Spec Peak 300~400 mW/cm, Total 2500~3000 mJ/cm.'
        ],
        'source_cells': {'title': ['Test!B2'], 'date': ['Test!Q3'], 'purpose': ['Test!A6'], 'content': ['Test!A8:A11']}
    },
    'test_conditions': [
        {'condition_id': 'cond_1', 'condition_group': 'Reduce UV LED 3rd Peak/Total (Dec27)',
         'line': '', 'process': 'UV cure 3rd', 'changed_factor': 'UV LED 3rd energy',
         'before_value': 'Normal 322 / 2913', 'after_value': '277 / 2077',
         'unit': 'mW/cm , mJ/cm', 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None,
         'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None,
         'uv_energy': 'Peak 277 / Total 2077',
         'source_file': name06, 'sheet_name': 'Test', 'source_cells': ['Test!D16:G16']},
        {'condition_id': 'cond_2', 'condition_group': 'Increase UV LED 3rd Peak/Total (Dec27)',
         'line': '', 'process': 'UV cure 3rd', 'changed_factor': 'UV LED 3rd energy',
         'before_value': 'Normal 322 / 2913', 'after_value': '689 / 5049',
         'unit': 'mW/cm , mJ/cm', 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None,
         'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None,
         'uv_energy': 'Peak 689 / Total 5049',
         'source_file': name06, 'sheet_name': 'Test', 'source_cells': ['Test!D18:G18']},
        {'condition_id': 'cond_3', 'condition_group': 'Normal UV (Dec27)',
         'line': '', 'process': 'UV cure 3rd', 'changed_factor': 'Baseline',
         'before_value': None, 'after_value': 'Normal',
         'unit': 'mW/cm , mJ/cm', 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None,
         'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None,
         'uv_energy': 'Peak 322 / Total 2913',
         'source_file': name06, 'sheet_name': 'Test', 'source_cells': ['Test!D20:G20']},
        {'condition_id': 'cond_4', 'condition_group': 'Reduce UV LED 3rd Peak/Total (Dec29)',
         'line': '', 'process': 'UV cure 3rd', 'changed_factor': 'UV LED 3rd energy (retest)',
         'before_value': 'Normal 315 / 2976', 'after_value': '280 / 2545',
         'unit': 'mW/cm , mJ/cm', 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None,
         'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None,
         'uv_energy': 'Peak 280 / Total 2545',
         'source_file': name06, 'sheet_name': 'Test', 'source_cells': ['Test!D22:G22']},
        {'condition_id': 'cond_5', 'condition_group': 'Normal UV (Dec29)',
         'line': '', 'process': 'UV cure 3rd', 'changed_factor': 'Baseline retest',
         'before_value': None, 'after_value': 'Normal',
         'unit': 'mW/cm , mJ/cm', 'machine': None, 'jig': None, 'material_lot': None, 'supplier': None,
         'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None,
         'uv_energy': 'Peak 315 / Total 2976',
         'source_file': name06, 'sheet_name': 'Test', 'source_cells': ['Test!D24:G24']},
    ],
    'results': [
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Reduce UV 277/2077', 'date': '2025-12-27', 'line': '',
         'input_count': 100, 'ok_count': 132, 'ng_count': 12, 'ng_rate_decimal': 0.12, 'ng_rate_percent': 12.0,
         'metric_name': 'NG rate Function', 'metric_value': 12.0, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG Audiobus SPL': 7, 'NG Audiobus SPL+RB': 0, 'NG Audiobus RB': 8,
                          'NG Audiobus No sound': 0, 'NG Hearing Noise': 5, 'NG Hearing Touch': 0,
                          'VP+Coil Separation in decap': 0},
         'source_file': name06, 'sheet_name': 'Test', 'source_cells': ['Test!H16:Q16']},
        {'result_id': 'res_2', 'condition_id': 'cond_2', 'measurement_type': 'Function',
         'condition_group': 'Increase UV 689/5049', 'date': '2025-12-27', 'line': '',
         'input_count': 10, 'ok_count': 1, 'ng_count': 9, 'ng_rate_decimal': 0.9, 'ng_rate_percent': 90.0,
         'metric_name': 'NG rate Function', 'metric_value': 90.0, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG Audiobus SPL': 6, 'NG Audiobus SPL+RB': 3, 'NG Audiobus RB': 0,
                          'NG Audiobus No sound': 0, 'NG Hearing Noise': 0, 'NG Hearing Touch': 0},
         'source_file': name06, 'sheet_name': 'Test', 'source_cells': ['Test!H18:Q18']},
        {'result_id': 'res_3', 'condition_id': 'cond_3', 'measurement_type': 'Function',
         'condition_group': 'Normal 322/2913 (baseline)', 'date': '2025-12-27', 'line': '',
         'input_count': 140, 'ok_count': 131, 'ng_count': 9, 'ng_rate_decimal': 0.064, 'ng_rate_percent': 6.4,
         'metric_name': 'NG rate Function', 'metric_value': 6.4, 'unit': '%', 'judgement': 'CHECK',
         'ng_breakdown': {'NG Audiobus SPL': 6, 'NG Audiobus SPL+RB': 1, 'NG Audiobus RB': 2,
                          'NG Audiobus No sound': 0, 'NG Hearing Noise': 2, 'NG Hearing Touch': 0},
         'source_file': name06, 'sheet_name': 'Test', 'source_cells': ['Test!H20:Q20']},
        {'result_id': 'res_4', 'condition_id': 'cond_4', 'measurement_type': 'Function',
         'condition_group': 'Reduce UV 280/2545 (highlighted)', 'date': '2025-12-29', 'line': '',
         'input_count': 178, 'ok_count': 132, 'ng_count': 44, 'ng_rate_decimal': 0.247, 'ng_rate_percent': 24.7,
         'metric_name': 'NG rate Function', 'metric_value': 24.7, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG Audiobus SPL': 7, 'NG Audiobus SPL+RB': 3, 'NG Audiobus RB': 39,
                          'NG Audiobus No sound': 0, 'NG Hearing Noise': 34, 'NG Hearing Touch': 0,
                          'SPL all 10kHz note': None},
         'source_file': name06, 'sheet_name': 'Test', 'source_cells': ['Test!H22:Q22']},
        {'result_id': 'res_5', 'condition_id': 'cond_5', 'measurement_type': 'Function',
         'condition_group': 'Normal 315/2976 (baseline Dec29)', 'date': '2025-12-29', 'line': '',
         'input_count': 100, 'ok_count': 82, 'ng_count': 18, 'ng_rate_decimal': 0.18, 'ng_rate_percent': 18.0,
         'metric_name': 'NG rate Function', 'metric_value': 18.0, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG Audiobus SPL': 2, 'NG Audiobus SPL+RB': 1, 'NG Audiobus RB': 17,
                          'NG Audiobus No sound': 0, 'NG Hearing Noise': 15, 'NG Hearing Touch': 0},
         'source_file': name06, 'sheet_name': 'Test', 'source_cells': ['Test!H24:Q24']},
        # Tension Dec27
        {'result_id': 'res_t1', 'condition_id': 'cond_1', 'measurement_type': 'Tension',
         'condition_group': 'Reduce UV Dec27 VP+COIL', 'date': '2025-12-27', 'line': '',
         'input_count': 10, 'ok_count': 10, 'ng_count': 0, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Tension AVG', 'metric_value': 0.969, 'unit': 'kg', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name06, 'sheet_name': 'Test', 'source_cells': ['Test!H29:S29']},
        {'result_id': 'res_t2', 'condition_id': 'cond_3', 'measurement_type': 'Tension',
         'condition_group': 'Normal VP+COIL', 'date': '2025-12-27', 'line': '',
         'input_count': 10, 'ok_count': 10, 'ng_count': 0, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Tension AVG', 'metric_value': 1.042, 'unit': 'kg', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name06, 'sheet_name': 'Test', 'source_cells': ['Test!H30:S30']},
    ],
    'conclusions': [
        {'conclusion_id': 'concl_1', 'topic': 'Decision',
         'statement_from_report': 'Decision section in source is empty.',
         'normalized_interpretation': 'DOE compares 4 UV settings to Normal. Dec27: Reduce 277/2077 → 12.0% (12/100) vs Normal 6.4% (9/140) = 1.88x, 87.5% worse. Increase 689/5049 → 90.0% (9/10) vs Normal 6.4% = 14.06x, 1306% worse (sample 10). Dec29: Reduce 280/2545 → 24.7% (44/178) vs Normal 18.0% (18/100) = 1.37x, 37.2% worse, dominated by RB (21.9%) and Hearing Noise (19.1%). All reduce/increase conditions worsen function vs same-event Normal. Tension PASS with AVG 0.969 (reduce) vs Normal 1.042 (-7.0% lower but still PASS).',
         'source_file': name06, 'sheet_name': 'Test', 'source_cells': ['Test!A33']},
        {'conclusion_id': 'concl_2', 'topic': 'Failure-mode shift',
         'statement_from_report': 'SPL all 10Khz note on reduce UV (Dec29).',
         'normalized_interpretation': 'Reducing UV shifts dominant NG toward RB and Hearing Noise (24.7% total: 21.9% RB, 19.1% Noise). Increasing UV shifts toward SPL/SPL+RB (60% SPL, 30% SPL+RB).',
         'source_file': name06, 'sheet_name': 'Test', 'source_cells': ['Test!N16:Q24']},
    ],
    'troubleshooting_index': {
        'defect_name': 'UV LED 3rd energy sensitivity (function NG)',
        'when_user_asks': ['What is the function impact of reducing/increasing UV LED 3rd Peak and Total?', 'Which UV setting minimizes RB/SPL NG?'],
        'suggested_checks': [
            {'hint_id': 'hint_1', 'check_item': 'Keep UV LED 3rd within spec 300-400 / 2500-3000',
             'reason': 'Reduce 277/2077 12.0% vs Normal 6.4% = 1.88x worse (87.5%); Increase 689/5049 90.0% vs Normal 6.4% = 14.06x worse (1306%) - both extremes degrade function.',
             'evidence_strength': 'high', 'related_process': 'UV cure 3rd', 'related_part': 'UV LED 3rd',
             'source_file': name06, 'sheet_name': 'Test', 'source_cells': ['Test!H16:Q20']},
            {'hint_id': 'hint_2', 'check_item': 'Watch RB / Hearing Noise when UV is below center',
             'reason': 'Dec29 Reduce 280/2545: NG Audiobus RB 21.9% and Hearing Noise 19.1% dominate; Normal Dec29 also high at 18.0%, reduced UV 24.7% = 1.37x (37.2% worse).',
             'evidence_strength': 'medium', 'related_process': 'UV cure 3rd', 'related_part': 'Audiobus / Hearing',
             'source_file': name06, 'sheet_name': 'Test', 'source_cells': ['Test!N22:Q22']},
            {'hint_id': 'hint_3', 'check_item': 'Watch SPL / SPL+RB when UV is over spec',
             'reason': 'Increase 689/5049 (Dec27, n=10): SPL 60%, SPL+RB 30% - sigma audiobus dominated.',
             'evidence_strength': 'medium', 'related_process': 'UV cure 3rd', 'related_part': 'Audiobus SPL',
             'source_file': name06, 'sheet_name': 'Test', 'source_cells': ['Test!N18:Q18']},
            {'hint_id': 'hint_4', 'check_item': 'Verify VP+Coil tension after UV change',
             'reason': 'Reduce UV tension AVG 0.969 vs Normal 1.042 (-7.0%, both PASS spec 0.4); margin still acceptable.',
             'evidence_strength': 'medium', 'related_process': 'Tension', 'related_part': 'VP/Coil',
             'source_file': name06, 'sheet_name': 'Test', 'source_cells': ['Test!H29:S30']},
        ],
        'limitations': ['Increase-UV row has only 10pcs - statistically weak.', 'Tension table Dec29 is empty (#DIV/0!).']
    },
    'ai_extraction_log': {
        'confidence': 0.8,
        'assumptions': ['Spec Peak 300~400 mW/cm, Total 2500~3000 mJ/cm interpreted as cure spec window.'],
        'warnings': ['Increase-UV n=10 too small for strong claim; Tension Dec29 row empty.', 'Decision section blank in source.'],
        'decision_rationale': 'Reducing or increasing UV LED 3rd energy worsens function vs same-event Normal: Reduce 277/2077 1.88x worse (87.5%), Reduce 280/2545 1.37x worse (37.2%), Increase 689/5049 14.06x worse (1306%). Tension still PASS though slightly lower. Stay within spec window.'
    },
}
tr_en_06 = _tr_from_result(result06)
tr_ko_06 = {
    'document': {
        'title': 'TIU-C11-20 UV LED 3rd 시험 리포트',
        'purpose': 'UV LED 3rd Peak/Total 에너지를 낮추고 높여 function과 Normal 비교.',
        'content': [
            'UV LED 3rd 감소/증가 시험.', '샘플 제작 후 tension 확인.',
            '조건당 100pcs Function 시험.', 'Normal 로트와 비교. Spec Peak 300~400 mW/cm, Total 2500~3000 mJ/cm.'
        ]
    },
    'conclusions': {
        'concl_1': {'topic': '결정', 'statement_from_report': '원본 Decision 섹션 비어있음.',
                    'normalized_interpretation': '4 UV 설정 vs Normal DOE. 12/27 Reduce 277/2077 → 12.0%(12/100) vs Normal 6.4%(9/140) = 1.88배, 87.5% 악화. Increase 689/5049 → 90.0%(9/10) vs Normal 6.4% = 14.06배, 1306% 악화(n=10). 12/29 Reduce 280/2545 → 24.7%(44/178) vs Normal 18.0%(18/100) = 1.37배, 37.2% 악화, RB 21.9%·Hearing Noise 19.1% 주도. 모든 조건이 Normal 대비 악화. Tension PASS, Reduce AVG 0.969 vs Normal 1.042(-7.0% 낮음).'},
        'concl_2': {'topic': '실패 모드 이동',
                    'statement_from_report': 'Reduce UV(12/29) SPL all 10kHz 메모.',
                    'normalized_interpretation': 'UV 감소 시 RB·Hearing Noise 주도(24.7%: RB 21.9%, Noise 19.1%). UV 증가 시 SPL·SPL+RB 주도(60%, 30%).'},
    },
    'hints': {
        'hint_1': {'check_item': 'UV LED 3rd 스펙 300-400 / 2500-3000 유지',
                   'reason': 'Reduce 12.0% vs 6.4% = 1.88배(87.5% 악화); Increase 90.0% vs 6.4% = 14.06배(1306% 악화) - 양극단 모두 악화.'},
        'hint_2': {'check_item': 'UV가 중심 미만일 때 RB / Hearing Noise 주시',
                   'reason': '12/29 Reduce 280/2545 RB 21.9%·Noise 19.1% 주도; Normal 18.0%·Reduce 24.7% = 1.37배(37.2% 악화).'},
        'hint_3': {'check_item': 'UV가 스펙 상한 초과 시 SPL / SPL+RB 주시',
                   'reason': 'Increase 689/5049(n=10): SPL 60%, SPL+RB 30%.'},
        'hint_4': {'check_item': 'UV 변경 후 VP+Coil tension 확인',
                   'reason': 'Reduce tension AVG 0.969 vs Normal 1.042(-7.0%, 둘 다 스펙 0.4 PASS).'},
    },
    'log': {
        'assumptions': ['Spec Peak 300~400 mW/cm, Total 2500~3000 mJ/cm을 cure window로 해석.'],
        'warnings': ['Increase-UV n=10으로 통계적 약함; Tension 12/29 비어있음.', '원본 Decision 섹션 비어있음.'],
        'decision_rationale': 'UV LED 3rd 감소/증가 모두 Normal 대비 function 악화: Reduce 277/2077 1.88배(87.5%), Reduce 280/2545 1.37배(37.2%), Increase 689/5049 14.06배(1306%). Tension PASS이나 약간 낮음. 스펙 window 유지 필요.'
    }
}
tr_vi_06 = {
    'document': {
        'title': 'BÁO CÁO TEST UV LED 3rd TIU-C11-20',
        'purpose': 'Test giảm và tăng UV LED 3rd Peak/Total và so sánh function với lot normal.',
        'content': [
            'Test giảm và tăng UV LED 3rd.', 'Làm mẫu, check tension.',
            'Test function 100pcs / điều kiện.', 'So sánh với normal lot. Spec Peak 300~400 mW/cm, Total 2500~3000 mJ/cm.'
        ]
    },
    'conclusions': {
        'concl_1': {'topic': 'Quyết định', 'statement_from_report': 'Phần Decision trong file gốc trống.',
                    'normalized_interpretation': 'DOE 4 setting UV vs Normal. 27/12 Reduce 277/2077 → 12.0% (12/100) vs Normal 6.4% (9/140) = 1.88x, xấu 87.5%. Increase 689/5049 → 90.0% (9/10) vs Normal 6.4% = 14.06x, xấu 1306% (n=10). 29/12 Reduce 280/2545 → 24.7% (44/178) vs Normal 18.0% (18/100) = 1.37x, xấu 37.2%, chủ yếu RB (21.9%) và Hearing Noise (19.1%). Tất cả Reduce/Increase đều xấu hơn Normal cùng sự kiện. Tension PASS, Reduce AVG 0.969 vs Normal 1.042 (-7.0% nhưng vẫn PASS).'},
        'concl_2': {'topic': 'Failure mode shift',
                    'statement_from_report': 'Reduce UV (29/12) ghi SPL all 10kHz.',
                    'normalized_interpretation': 'Giảm UV thì NG dồn về RB và Hearing Noise (24.7% total: RB 21.9%, Noise 19.1%). Tăng UV dồn về SPL/SPL+RB (60% SPL, 30% SPL+RB).'},
    },
    'hints': {
        'hint_1': {'check_item': 'Giữ UV LED 3rd trong spec 300-400 / 2500-3000',
                   'reason': 'Reduce 12.0% vs 6.4% = 1.88x (xấu 87.5%); Increase 90.0% vs 6.4% = 14.06x (xấu 1306%) - cả hai cực đều xấu.'},
        'hint_2': {'check_item': 'Theo dõi RB / Hearing Noise khi UV dưới center',
                   'reason': '29/12 Reduce 280/2545: RB 21.9% và Noise 19.1% dominate; Normal 18.0%, Reduce 24.7% = 1.37x (xấu 37.2%).'},
        'hint_3': {'check_item': 'Theo dõi SPL / SPL+RB khi UV vượt spec',
                   'reason': 'Increase 689/5049 (n=10): SPL 60%, SPL+RB 30%.'},
        'hint_4': {'check_item': 'Verify tension VP+Coil sau khi đổi UV',
                   'reason': 'Reduce UV tension AVG 0.969 vs Normal 1.042 (-7.0%, cả hai PASS spec 0.4).'},
    },
    'log': {
        'assumptions': ['Spec Peak 300~400 mW/cm, Total 2500~3000 mJ/cm = cure window.'],
        'warnings': ['Increase-UV n=10 nhỏ; Tension 29/12 trống.', 'Decision section file gốc trống.'],
        'decision_rationale': 'Giảm hay tăng UV LED 3rd đều xấu hơn Normal cùng sự kiện: Reduce 277/2077 1.88x (87.5%), Reduce 280/2545 1.37x (37.2%), Increase 689/5049 14.06x (1306%). Tension vẫn PASS dù thấp hơn chút. Giữ trong window spec.'
    }
}
run(name06, result06, tr_ko_06, tr_en_06, tr_vi_06)


# ===== DS 07 =====
name07 = '25. TIU L5S3-01 R Report test machine AWF change design pusher -  13.12.2025'
result07 = {
    'schema_version': '0.1',
    'document': {
        'document_id': '', 'source_file': name07, 'source_sheet': 'Test',
        'title': 'TIU L5S3-01 [R] REPORT TEST MACHINE AWF #2 CHANGE DESIGN PUSHER',
        'model': 'TIU L5S3-01 R', 'report_date': '2025-12-13', 'department': 'ME',
        'marker': 'Thao', 'line': '',
        'report_type': 'normal_comparison',
        'primary_defect': {'canonical_name': 'NG BAKO (FRF/FRF+SPL)', 'aliases_in_document': ['NG BAKO', 'FRF', 'FRF+SPL', 'glue in coil']},
        'related_defects': ['NG BAKO FRF', 'NG BAKO FRF+SPL', 'Glue in coil'],
        'parts': ['AWF Machine', 'Pusher', 'VP', 'Coil'],
        'processes': ['AWF bonding', 'Decap'],
        'purpose': 'Evaluate AWF Machine #2 with redesigned pusher and find the reason for NG.',
        'content': [
            'Test machine AWF #2 with redesigned pusher and compare with AWF #1 and AWF #3.',
            'Check function after AWF separation.'
        ],
        'source_cells': {'title': ['Test!B2'], 'date': ['Test!P3'], 'purpose': ['Test!A6'], 'content': ['Test!A8:A9']}
    },
    'test_conditions': [
        {'condition_id': 'cond_1', 'condition_group': 'AWF #1 vs AWF #2 (new pusher) vs AWF #3',
         'line': '', 'process': 'AWF bonding', 'changed_factor': 'Pusher design + machine identity',
         'before_value': 'AWF #1 / AWF #3 (existing pusher)', 'after_value': 'AWF #2 (new pusher design)',
         'unit': None, 'machine': 'AWF #2', 'jig': 'New design pusher',
         'material_lot': None, 'supplier': None, 'dry_time_sec': None, 'temperature': None,
         'pressure': None, 'bond_amount': '0.11~0.12mg VP/Coil', 'uv_energy': None,
         'source_file': name07, 'sheet_name': 'Test', 'source_cells': ['Test!B15:O22']},
    ],
    'results': [
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'AWF #1 (baseline)', 'date': '2025-12-13', 'line': '',
         'input_count': 140, 'ok_count': 103, 'ng_count': 37, 'ng_rate_decimal': 0.264, 'ng_rate_percent': 26.4,
         'metric_name': 'NG rate Function', 'metric_value': 26.4, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG BAKO FRF': 31, 'NG BAKO FRF+SPL': 6, 'NG BAKO THD': 0, 'NG BAKO No sound': 0},
         'source_file': name07, 'sheet_name': 'Test', 'source_cells': ['Test!H15:N15']},
        {'result_id': 'res_2', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'AWF #2 new pusher (test 1)', 'date': '2025-12-13', 'line': '',
         'input_count': 84, 'ok_count': 52, 'ng_count': 32, 'ng_rate_decimal': 0.381, 'ng_rate_percent': 38.1,
         'metric_name': 'NG rate Function', 'metric_value': 38.1, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG BAKO FRF': 22, 'NG BAKO FRF+SPL': 10, 'NG BAKO THD': 0, 'NG BAKO No sound': 0,
                          'Glue in coil (decap 8/10 ~ 80%)': None},
         'source_file': name07, 'sheet_name': 'Test', 'source_cells': ['Test!H17:N17']},
        {'result_id': 'res_3', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'AWF #2 new pusher (test 2)', 'date': '2025-12-13', 'line': '',
         'input_count': 46, 'ok_count': 30, 'ng_count': 16, 'ng_rate_decimal': 0.348, 'ng_rate_percent': 34.8,
         'metric_name': 'NG rate Function', 'metric_value': 34.8, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG BAKO FRF': 12, 'NG BAKO FRF+SPL': 4, 'NG BAKO THD': 0, 'NG BAKO No sound': 0},
         'source_file': name07, 'sheet_name': 'Test', 'source_cells': ['Test!H19:N19']},
        {'result_id': 'res_4', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'AWF #3 (reference)', 'date': '2025-12-13', 'line': '',
         'input_count': 143, 'ok_count': 108, 'ng_count': 35, 'ng_rate_decimal': 0.245, 'ng_rate_percent': 24.5,
         'metric_name': 'NG rate Function', 'metric_value': 24.5, 'unit': '%', 'judgement': 'FAIL',
         'ng_breakdown': {'NG BAKO FRF': 32, 'NG BAKO FRF+SPL': 3, 'NG BAKO THD': 0, 'NG BAKO No sound': 0},
         'source_file': name07, 'sheet_name': 'Test', 'source_cells': ['Test!H21:N21']},
    ],
    'conclusions': [
        {'conclusion_id': 'concl_1', 'topic': 'Decision (data observations)',
         'statement_from_report': 'Decision section in source is empty.',
         'normalized_interpretation': 'AWF #2 (new pusher) test1 NG 38.1% (32/84) vs AWF #1 baseline 26.4% (37/140) = 1.44x, 44.3% worse. AWF #2 test2 NG 34.8% (16/46) vs AWF #1 = 1.32x, 31.8% worse. AWF #2 test1 vs AWF #3 baseline (24.5%) = 1.55x, 55.5% worse. New pusher design worsens function across both runs and produces glue-in-coil decap evidence (test1: 8/10 FRF decaps showed glue in coil ~ 80%); test2 noted no glue-in-coil. Existing pushers (AWF #1, AWF #3) similar at 24-26%, suggesting baseline FRF issue is platform-wide.',
         'source_file': name07, 'sheet_name': 'Test', 'source_cells': ['Test!A25']},
        {'conclusion_id': 'concl_2', 'topic': 'Decap mode shift',
         'statement_from_report': 'Test 1: 8/10 NG FRF show glue in coil ~80%. Test 2: no glue in coil.',
         'normalized_interpretation': 'New pusher first run produced an unusual glue-in-coil failure (8/10 of decapped FRF NG); the second run did not - process repeatability is an open question.',
         'source_file': name07, 'sheet_name': 'Test', 'source_cells': ['Test!N17:N19']},
    ],
    'troubleshooting_index': {
        'defect_name': 'NG BAKO FRF + glue in coil (AWF pusher change)',
        'when_user_asks': ['Does the redesigned AWF #2 pusher reduce NG BAKO?', 'Why does NG FRF spike with the new pusher?'],
        'suggested_checks': [
            {'hint_id': 'hint_1', 'check_item': 'Compare same-day AWF #2 new pusher vs AWF #1 / AWF #3 baselines',
             'reason': 'AWF #2 test1 NG 38.1% vs AWF #1 26.4% (1.44x, 44.3% worse) and AWF #3 24.5% (1.55x, 55.5% worse) - new pusher worsens function.',
             'evidence_strength': 'high', 'related_process': 'AWF bonding', 'related_part': 'Pusher',
             'source_file': name07, 'sheet_name': 'Test', 'source_cells': ['Test!H15:N21']},
            {'hint_id': 'hint_2', 'check_item': 'Decap NG FRF samples to look for glue in coil',
             'reason': 'AWF #2 test1: 8/10 FRF NG decaps showed glue in coil; test2 did not - new pusher may push glue into coil intermittently.',
             'evidence_strength': 'medium', 'related_process': 'Decap', 'related_part': 'Coil',
             'source_file': name07, 'sheet_name': 'Test', 'source_cells': ['Test!N17:N19']},
            {'hint_id': 'hint_3', 'check_item': 'Reproduce AWF #2 result with more pcs',
             'reason': 'AWF #2 test1 (84pcs) and test2 (46pcs) show 38.1% vs 34.8% - directionally similar but small sample.',
             'evidence_strength': 'low', 'related_process': 'AWF bonding', 'related_part': 'Pusher',
             'source_file': name07, 'sheet_name': 'Test', 'source_cells': ['Test!H17:H19']},
        ],
        'limitations': ['AWF #2 sample sizes 46~84 vs AWF #1/#3 ~140 each.', 'Decision section blank.']
    },
    'ai_extraction_log': {
        'confidence': 0.8,
        'assumptions': ['AWF #1 and #3 treated as separate same-event baselines (existing pusher).', 'Glue-in-coil decap "8/10 ~80%" treated as a qualitative observation, not a count for the 84-pcs population.'],
        'warnings': ['Decision section blank.', 'Glue-in-coil claim relies on a 10-piece decap sub-sample.'],
        'decision_rationale': 'AWF #2 new pusher worsens function vs AWF #1 and AWF #3 same-day baselines: test1 38.1% (1.44x vs #1, 1.55x vs #3); test2 34.8% (1.32x vs #1, 1.42x vs #3). Decap shows glue-in-coil in test1. New pusher design not yet acceptable; revert or further redesign indicated.'
    },
}
tr_en_07 = _tr_from_result(result07)
tr_ko_07 = {
    'document': {
        'title': 'TIU L5S3-01 [R] AWF #2 Pusher 디자인 변경 시험 리포트',
        'purpose': 'AWF #2의 신규 pusher 디자인 평가 및 NG 원인 파악.',
        'content': [
            'AWF #2 신 pusher 시험, AWF #1 / AWF #3과 비교.',
            'AWF 분리 후 function 확인.'
        ]
    },
    'conclusions': {
        'concl_1': {'topic': '결정(데이터 관찰)', 'statement_from_report': '원본 Decision 섹션 비어있음.',
                    'normalized_interpretation': 'AWF #2(신 pusher) test1 NG 38.1%(32/84) vs AWF #1 26.4%(37/140) = 1.44배, 44.3% 악화. AWF #2 test2 NG 34.8%(16/46) vs AWF #1 = 1.32배, 31.8% 악화. AWF #2 test1 vs AWF #3(24.5%) = 1.55배, 55.5% 악화. 신 pusher가 두 차례 모두 악화, test1 decap 8/10 FRF에서 glue-in-coil 약 80% 관찰; test2는 없음. AWF #1·#3은 24-26%로 비슷, 베이스라인 FRF 이슈는 플랫폼 전반적.'},
        'concl_2': {'topic': 'Decap 모드 변화',
                    'statement_from_report': 'Test1: NG FRF 8/10 glue in coil ~80%. Test2: glue in coil 없음.',
                    'normalized_interpretation': '신 pusher 첫 시험에서 glue-in-coil 실패 모드 등장, 두 번째 시험에서는 미관찰 - 공정 재현성 미확정.'},
    },
    'hints': {
        'hint_1': {'check_item': '동일 일 AWF #2 신 pusher vs AWF #1 / AWF #3 비교',
                   'reason': 'AWF #2 test1 38.1% vs AWF #1 26.4%(1.44배, 44.3% 악화), AWF #3 24.5%(1.55배, 55.5% 악화) - 신 pusher 악화.'},
        'hint_2': {'check_item': 'NG FRF 샘플 decap으로 glue in coil 확인',
                   'reason': 'AWF #2 test1: FRF NG decap 8/10 glue-in-coil; test2 없음 - 신 pusher가 간헐적으로 본드를 coil로 밀어내는 가능성.'},
        'hint_3': {'check_item': 'AWF #2 결과를 더 큰 표본으로 재현',
                   'reason': 'test1 84pcs vs test2 46pcs에서 38.1%·34.8%로 방향성 유사하나 표본 작음.'},
    },
    'log': {
        'assumptions': ['AWF #1·#3을 각각 동일 이벤트 baseline(기존 pusher)으로 간주.', 'Glue-in-coil "8/10 ~80%"는 84pcs 카운트 아닌 정성적 관찰로 처리.'],
        'warnings': ['원본 Decision 섹션 비어있음.', 'Glue-in-coil 주장은 10pcs decap 서브샘플 기반.'],
        'decision_rationale': 'AWF #2 신 pusher가 두 시험 모두 baseline AWF #1·#3 대비 악화: test1 38.1%(1.44배 vs #1, 1.55배 vs #3); test2 34.8%(1.32배 vs #1, 1.42배 vs #3). Test1 decap에서 glue-in-coil. 신 pusher 미수용, 원복 또는 재설계 필요.'
    }
}
tr_vi_07 = {
    'document': {
        'title': 'TIU L5S3-01 [R] BÁO CÁO TEST MÁY AWF #2 ĐỔI DESIGN PUSHER',
        'purpose': 'Đánh giá pusher mới của máy AWF #2 và tìm reason NG.',
        'content': [
            'Test máy AWF #2 pusher mới và so sánh với AWF #1, AWF #3.',
            'Check function sau AWF.'
        ]
    },
    'conclusions': {
        'concl_1': {'topic': 'Quyết định (quan sát dữ liệu)', 'statement_from_report': 'Phần Decision trong file gốc trống.',
                    'normalized_interpretation': 'AWF #2 (pusher mới) test1 NG 38.1% (32/84) vs AWF #1 26.4% (37/140) = 1.44x, xấu 44.3%. AWF #2 test2 NG 34.8% (16/46) vs AWF #1 = 1.32x, xấu 31.8%. AWF #2 test1 vs AWF #3 (24.5%) = 1.55x, xấu 55.5%. Pusher mới làm xấu ở cả hai lần chạy, test1 decap 8/10 FRF có glue-in-coil ~80%; test2 không. AWF #1 và #3 ~24-26%, vấn đề FRF nền tảng chung.'},
        'concl_2': {'topic': 'Decap mode shift',
                    'statement_from_report': 'Test1: 8/10 NG FRF có glue in coil ~80%. Test2: không có glue in coil.',
                    'normalized_interpretation': 'Pusher mới lần đầu tạo failure mode glue-in-coil; lần hai không - tái lặp của process còn nghi vấn.'},
    },
    'hints': {
        'hint_1': {'check_item': 'So sánh cùng ngày AWF #2 pusher mới vs AWF #1 / AWF #3',
                   'reason': 'AWF #2 test1 38.1% vs AWF #1 26.4% (1.44x, xấu 44.3%) và AWF #3 24.5% (1.55x, xấu 55.5%) - pusher mới làm xấu.'},
        'hint_2': {'check_item': 'Decap mẫu NG FRF để tìm glue in coil',
                   'reason': 'AWF #2 test1: 8/10 FRF NG decap có glue in coil; test2 không - pusher mới có thể đẩy keo vào coil không liên tục.'},
        'hint_3': {'check_item': 'Tái hiện kết quả AWF #2 với sample lớn hơn',
                   'reason': 'test1 84pcs vs test2 46pcs cho 38.1% và 34.8% - cùng hướng nhưng size nhỏ.'},
    },
    'log': {
        'assumptions': ['AWF #1 và #3 dùng làm baseline cùng sự kiện (pusher hiện tại).', 'Glue-in-coil "8/10 ~80%" coi là quan sát định tính, không phải count cho 84pcs.'],
        'warnings': ['Decision section trống.', 'Glue-in-coil dựa trên 10pcs decap sub-sample.'],
        'decision_rationale': 'Pusher mới AWF #2 xấu hơn baseline AWF #1 và AWF #3 cùng ngày: test1 38.1% (1.44x vs #1, 1.55x vs #3); test2 34.8% (1.32x vs #1, 1.42x vs #3). Decap test1 có glue-in-coil. Pusher mới chưa được, cần revert hoặc redesign.'
    }
}
run(name07, result07, tr_ko_07, tr_en_07, tr_vi_07)


# ===== DS 08 (large reliability_spec, summary only) =====
name08 = '25.1 MSU-L20S15-07 Result check NTI lot test New Bond SJ4774 and VE562850- 2025.06.17'
result08 = {
    'schema_version': '0.1',
    'document': {
        'document_id': '', 'source_file': name08, 'source_sheet': 'GRAPH + RAW DATA',
        'title': 'NTI Lot Test - New Bond SJ4774 vs VE562850 vs Normal (MSU-L20S15-07)',
        'model': 'MSU-L20S15-07', 'report_date': '2025-06-17', 'department': 'ME',
        'marker': '', 'line': '',
        'report_type': 'reliability_spec',
        'primary_defect': {'canonical_name': 'Acoustic spec compliance (SPL/Fo/THD)', 'aliases_in_document': ['SPL', 'Fo', 'THD']},
        'related_defects': ['THD high', 'Fo low', 'SPL low'],
        'parts': ['Bond SJ4774', 'Bond VE562850', 'Speaker (MSU-L20S15-07)'],
        'processes': ['Acoustic measurement', 'Bond change qualification'],
        'purpose': 'Qualify new bonds SJ4774 and VE562850 against Normal lot using acoustic spec metrics (SPL bands, Fo, THD at 200/400/1000Hz).',
        'content': [
            'GRAPH sheet summarizes deltas of each lot AVG vs spec Center for: SPL 100-750Hz (109.2 +/-3dB), SPL 800-1.5k (118.3 +/-2dB), SPL 1.6-14k (115.8 +/-3dB), Fo (675 +/-120Hz), THD% 200/400/1000Hz (45/25/8 upper limits).',
            'RAW DATA sheet holds per-sample SPL/THD/Fo by frequency for Normal, SJ4774, VE562850 (10 samples each).',
            'Lots compared: STD, Normal, SJ4774, VE562850.'
        ],
        'source_cells': {'title': ['GRAPH!A1'], 'date': [], 'purpose': ['GRAPH!A1'], 'content': ['GRAPH!A1:Q10', 'RAW DATA!*']}
    },
    'test_conditions': [
        {'condition_id': 'cond_1', 'condition_group': 'Bond SJ4774 vs Normal',
         'line': '', 'process': 'Bond change qualification', 'changed_factor': 'Bond type',
         'before_value': 'Normal bond', 'after_value': 'SJ4774',
         'unit': None, 'machine': None, 'jig': None, 'material_lot': 'SJ4774', 'supplier': None,
         'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!A7']},
        {'condition_id': 'cond_2', 'condition_group': 'Bond VE562850 vs Normal',
         'line': '', 'process': 'Bond change qualification', 'changed_factor': 'Bond type',
         'before_value': 'Normal bond', 'after_value': 'VE562850',
         'unit': None, 'machine': None, 'jig': None, 'material_lot': 'VE562850', 'supplier': None,
         'dry_time_sec': None, 'temperature': None, 'pressure': None, 'bond_amount': None, 'uv_energy': None,
         'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!A8']},
    ],
    'results': [
        # SPL 100-750 Hz center 109.2 +/-3
        {'result_id': 'r_std_spl1', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'STD (spec reference)', 'date': '2025-06-17', 'line': '',
         'input_count': None, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'SPL 100-750Hz AVG', 'metric_value': 108.57, 'unit': 'dB', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!A5:E5']},
        {'result_id': 'r_norm_spl1', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'Normal (baseline)', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'SPL 100-750Hz AVG', 'metric_value': 108.42, 'unit': 'dB', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!A6:E6']},
        {'result_id': 'r_sj_spl1', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'SJ4774', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'SPL 100-750Hz AVG', 'metric_value': 108.04, 'unit': 'dB', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!A7:E7']},
        {'result_id': 'r_ve_spl1', 'condition_id': 'cond_2', 'measurement_type': 'Acoustic',
         'condition_group': 'VE562850', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'SPL 100-750Hz AVG', 'metric_value': 108.54, 'unit': 'dB', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!A8:E8']},
        # SPL 800-1.5k center 118.3 +/-2
        {'result_id': 'r_std_spl2', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'STD (spec reference)', 'date': '2025-06-17', 'line': '',
         'input_count': None, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'SPL 800Hz-1.5kHz AVG', 'metric_value': 117.35, 'unit': 'dB', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!F5:G5']},
        {'result_id': 'r_norm_spl2', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'Normal (baseline)', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'SPL 800Hz-1.5kHz AVG', 'metric_value': 117.18, 'unit': 'dB', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!F6:G6']},
        {'result_id': 'r_sj_spl2', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'SJ4774', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'SPL 800Hz-1.5kHz AVG', 'metric_value': 117.07, 'unit': 'dB', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!F7:G7']},
        {'result_id': 'r_ve_spl2', 'condition_id': 'cond_2', 'measurement_type': 'Acoustic',
         'condition_group': 'VE562850', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'SPL 800Hz-1.5kHz AVG', 'metric_value': 116.99, 'unit': 'dB', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!F8:G8']},
        # SPL 1.6-14k center 115.8 +/-3
        {'result_id': 'r_std_spl3', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'STD (spec reference)', 'date': '2025-06-17', 'line': '',
         'input_count': None, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'SPL 1.6kHz-14kHz AVG', 'metric_value': 115.13, 'unit': 'dB', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!H5:I5']},
        {'result_id': 'r_norm_spl3', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'Normal (baseline)', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'SPL 1.6kHz-14kHz AVG', 'metric_value': 115.15, 'unit': 'dB', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!H6:I6']},
        {'result_id': 'r_sj_spl3', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'SJ4774', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'SPL 1.6kHz-14kHz AVG', 'metric_value': 115.19, 'unit': 'dB', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!H7:I7']},
        {'result_id': 'r_ve_spl3', 'condition_id': 'cond_2', 'measurement_type': 'Acoustic',
         'condition_group': 'VE562850', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'SPL 1.6kHz-14kHz AVG', 'metric_value': 115.14, 'unit': 'dB', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!H8:I8']},
        # Fo center 675 +/-120
        {'result_id': 'r_std_fo', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'STD (spec reference)', 'date': '2025-06-17', 'line': '',
         'input_count': None, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Fo', 'metric_value': 653.44, 'unit': 'Hz', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!J5:K5']},
        {'result_id': 'r_norm_fo', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'Normal (baseline)', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Fo', 'metric_value': 652.24, 'unit': 'Hz', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!J6:K6']},
        {'result_id': 'r_sj_fo', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'SJ4774', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Fo', 'metric_value': 662.91, 'unit': 'Hz', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!J7:K7']},
        {'result_id': 'r_ve_fo', 'condition_id': 'cond_2', 'measurement_type': 'Acoustic',
         'condition_group': 'VE562850', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'Fo', 'metric_value': 637.46, 'unit': 'Hz', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!J8:K8']},
        # THD 200Hz upper 45
        {'result_id': 'r_std_thd200', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'STD (spec reference)', 'date': '2025-06-17', 'line': '',
         'input_count': None, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'THD% 200Hz', 'metric_value': 25.08, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!L5:M5']},
        {'result_id': 'r_norm_thd200', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'Normal (baseline)', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'THD% 200Hz', 'metric_value': 30.86, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!L6:M6']},
        {'result_id': 'r_sj_thd200', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'SJ4774', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'THD% 200Hz', 'metric_value': 31.73, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!L7:M7']},
        {'result_id': 'r_ve_thd200', 'condition_id': 'cond_2', 'measurement_type': 'Acoustic',
         'condition_group': 'VE562850', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'THD% 200Hz', 'metric_value': 30.25, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!L8:M8']},
        # THD 400Hz upper 25
        {'result_id': 'r_std_thd400', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'STD (spec reference)', 'date': '2025-06-17', 'line': '',
         'input_count': None, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'THD% 400Hz', 'metric_value': 8.83, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!N5:O5']},
        {'result_id': 'r_norm_thd400', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'Normal (baseline)', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'THD% 400Hz', 'metric_value': 12.29, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!N6:O6']},
        {'result_id': 'r_sj_thd400', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'SJ4774', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'THD% 400Hz', 'metric_value': 13.62, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!N7:O7']},
        {'result_id': 'r_ve_thd400', 'condition_id': 'cond_2', 'measurement_type': 'Acoustic',
         'condition_group': 'VE562850', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'THD% 400Hz', 'metric_value': 12.60, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!N8:O8']},
        # THD 1000Hz upper 8
        {'result_id': 'r_std_thd1k', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'STD (spec reference)', 'date': '2025-06-17', 'line': '',
         'input_count': None, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'THD% 1000Hz', 'metric_value': 2.93, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!P5:Q5']},
        {'result_id': 'r_norm_thd1k', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'Normal (baseline)', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'THD% 1000Hz', 'metric_value': 3.91, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!P6:Q6']},
        {'result_id': 'r_sj_thd1k', 'condition_id': 'cond_1', 'measurement_type': 'Acoustic',
         'condition_group': 'SJ4774', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'THD% 1000Hz', 'metric_value': 4.25, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!P7:Q7']},
        {'result_id': 'r_ve_thd1k', 'condition_id': 'cond_2', 'measurement_type': 'Acoustic',
         'condition_group': 'VE562850', 'date': '2025-06-17', 'line': '',
         'input_count': 10, 'ok_count': None, 'ng_count': None, 'ng_rate_decimal': None, 'ng_rate_percent': None,
         'metric_name': 'THD% 1000Hz', 'metric_value': 4.05, 'unit': '%', 'judgement': 'PASS',
         'ng_breakdown': {}, 'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!P8:Q8']},
    ],
    'conclusions': [
        {'conclusion_id': 'concl_1', 'topic': 'Bond SJ4774 vs Normal acoustic delta',
         'statement_from_report': 'No explicit decision text in TSV; GRAPH shows lot AVGs and deltas from spec Center.',
         'normalized_interpretation': 'SJ4774 vs Normal AVGs: SPL 100-750Hz 108.04 vs 108.42 (-0.38dB, -0.35%), SPL 800Hz-1.5k 117.07 vs 117.18 (-0.11dB, -0.09%), SPL 1.6k-14k 115.19 vs 115.15 (+0.04dB, +0.03%), Fo 662.91 vs 652.24 (+10.67Hz, +1.64%), THD200 31.73 vs 30.86 (+0.87pp, +2.82%), THD400 13.62 vs 12.29 (+1.33pp, +10.82%), THD1k 4.25 vs 3.91 (+0.34pp, +8.70%). All metrics still within spec band; THD400 +10.82% worse vs same-event Normal is the biggest deviation.',
         'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!A6:Q7']},
        {'conclusion_id': 'concl_2', 'topic': 'Bond VE562850 vs Normal acoustic delta',
         'statement_from_report': 'No explicit decision text in TSV.',
         'normalized_interpretation': 'VE562850 vs Normal AVGs: SPL 100-750Hz 108.54 vs 108.42 (+0.12dB, +0.11%), SPL 800Hz-1.5k 116.99 vs 117.18 (-0.19dB, -0.16%), SPL 1.6k-14k 115.14 vs 115.15 (-0.01dB, -0.01%), Fo 637.46 vs 652.24 (-14.78Hz, -2.27%), THD200 30.25 vs 30.86 (-0.61pp, -1.98% improved), THD400 12.60 vs 12.29 (+0.31pp, +2.52%), THD1k 4.05 vs 3.91 (+0.14pp, +3.58%). All metrics within spec band; closer to Normal than SJ4774.',
         'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!A6:Q8']},
        {'conclusion_id': 'concl_3', 'topic': 'Bond ranking summary',
         'statement_from_report': 'No explicit decision text.',
         'normalized_interpretation': 'Within spec, VE562850 tracks Normal more closely than SJ4774, particularly on THD400 (+2.52% vs +10.82%) and THD1k (+3.58% vs +8.70%). Fo: SJ4774 +1.64% closer to spec Center 675Hz than VE562850 -2.27%. No fail in spec band for either bond on n=10.',
         'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!A6:Q8']},
    ],
    'troubleshooting_index': {
        'defect_name': 'Acoustic spec compliance with new bond',
        'when_user_asks': ['How does new bond affect SPL/Fo/THD?', 'Which candidate bond tracks Normal better?'],
        'suggested_checks': [
            {'hint_id': 'hint_1', 'check_item': 'Compare candidate bond AVG vs same-event Normal AVG across SPL bands, Fo, THD bands',
             'reason': 'SJ4774 THD400 +10.82% vs Normal (12.29 → 13.62); VE562850 THD400 +2.52% (12.29 → 12.60); both still inside upper 25%.',
             'evidence_strength': 'medium', 'related_process': 'Acoustic measurement', 'related_part': 'Bond',
             'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!N6:O8']},
            {'hint_id': 'hint_2', 'check_item': 'Watch Fo shift after bond change',
             'reason': 'SJ4774 Fo 662.91Hz (Normal +1.64%); VE562850 637.46Hz (Normal -2.27%); both inside 675 +/-120Hz window.',
             'evidence_strength': 'medium', 'related_process': 'Acoustic measurement', 'related_part': 'Bond / suspension',
             'source_file': name08, 'sheet_name': 'GRAPH', 'source_cells': ['GRAPH!J6:K8']},
            {'hint_id': 'hint_3', 'check_item': 'Run n>=10 sample per bond and store per-sample raw values',
             'reason': 'RAW DATA captures 10 samples per bond across many frequencies, enabling sigma comparison after release.',
             'evidence_strength': 'low', 'related_process': 'Acoustic measurement', 'related_part': 'Bond',
             'source_file': name08, 'sheet_name': 'RAW DATA', 'source_cells': ['RAW DATA!*']},
        ],
        'limitations': ['n=10 per bond; small sample for reliability claim.', 'TSV decoded with replacement characters in headers - column letter mapping approximate.', 'No explicit decision text in source.']
    },
    'ai_extraction_log': {
        'confidence': 0.7,
        'assumptions': ['GRAPH row values for STD/Normal/SJ4774/VE562850 read in column order: SPL 100-750, SPL 800-1.5k, SPL 1.6-14k, Fo, THD200, THD400, THD1k.', 'All AVGs within spec band treated as PASS.'],
        'warnings': ['TSV has encoding artifacts (��) - some inline labels could not be read; numbers were read from the row pattern.', 'Decision text absent; conclusion derived purely from numerics.', 'n=10 per bond.'],
        'decision_rationale': 'Both new bonds remain within spec. SJ4774 vs Normal: largest gap THD400 +10.82% worse; THD1k +8.70% worse; Fo +1.64% closer to spec Center. VE562850 vs Normal: largest gap Fo -2.27% lower; THD400 +2.52% worse; all SPL within +/-0.2dB. Overall VE562850 tracks Normal more tightly; SJ4774 trades a closer Fo for higher THD.'
    },
}
tr_en_08 = _tr_from_result(result08)
tr_ko_08 = {
    'document': {
        'title': 'MSU-L20S15-07 신본드 SJ4774·VE562850 음향 사양 시험 (Normal 비교)',
        'purpose': 'SJ4774 및 VE562850 신본드를 Normal과 동일 이벤트로 음향 사양(SPL 대역, Fo, THD 200/400/1000Hz)으로 평가.',
        'content': [
            'GRAPH 시트: 각 lot AVG가 스펙 Center 대비 차이 (SPL 100-750Hz 109.2±3dB, SPL 800-1.5k 118.3±2dB, SPL 1.6-14k 115.8±3dB, Fo 675±120Hz, THD 200/400/1000Hz 상한 45/25/8).',
            'RAW DATA 시트: Normal·SJ4774·VE562850 각 10pcs의 주파수별 SPL/THD/Fo.',
            '비교 lot: STD, Normal, SJ4774, VE562850.'
        ]
    },
    'conclusions': {
        'concl_1': {'topic': 'SJ4774 vs Normal 음향 차이',
                    'statement_from_report': 'TSV에 명시적 Decision 없음; GRAPH 시트 AVG 및 Center 대비 델타 표.',
                    'normalized_interpretation': 'SJ4774 vs Normal AVG: SPL 100-750 108.04 vs 108.42(-0.38dB, -0.35%), SPL 800-1.5k 117.07 vs 117.18(-0.11dB, -0.09%), SPL 1.6-14k 115.19 vs 115.15(+0.04dB, +0.03%), Fo 662.91 vs 652.24(+10.67Hz, +1.64%), THD200 31.73 vs 30.86(+0.87pp, +2.82%), THD400 13.62 vs 12.29(+1.33pp, +10.82%), THD1k 4.25 vs 3.91(+0.34pp, +8.70%). 모두 스펙 내; THD400 +10.82% 악화가 최대 편차.'},
        'concl_2': {'topic': 'VE562850 vs Normal 음향 차이',
                    'statement_from_report': 'TSV에 명시적 Decision 없음.',
                    'normalized_interpretation': 'VE562850 vs Normal AVG: SPL 100-750 108.54 vs 108.42(+0.12dB, +0.11%), SPL 800-1.5k 116.99 vs 117.18(-0.19dB, -0.16%), SPL 1.6-14k 115.14 vs 115.15(-0.01dB, -0.01%), Fo 637.46 vs 652.24(-14.78Hz, -2.27%), THD200 30.25 vs 30.86(-0.61pp, -1.98% 개선), THD400 12.60 vs 12.29(+0.31pp, +2.52%), THD1k 4.05 vs 3.91(+0.14pp, +3.58%). 모두 스펙 내; SJ4774보다 Normal에 가까움.'},
        'concl_3': {'topic': '본드 순위 요약',
                    'statement_from_report': '명시적 Decision 없음.',
                    'normalized_interpretation': '스펙 내 두 본드 모두 통과. VE562850이 SJ4774보다 Normal에 가깝게 추종, 특히 THD400(+2.52% vs +10.82%) 및 THD1k(+3.58% vs +8.70%). Fo는 SJ4774(+1.64%)가 스펙 Center 675Hz에 더 가깝고 VE562850(-2.27%)는 약간 떨어짐. n=10에서 어느 본드도 스펙 fail 없음.'},
    },
    'hints': {
        'hint_1': {'check_item': '신본드 AVG와 동일 이벤트 Normal AVG를 SPL 대역·Fo·THD 대역별로 비교',
                   'reason': 'SJ4774 THD400 +10.82%(12.29→13.62); VE562850 THD400 +2.52%(12.29→12.60); 모두 상한 25% 내.'},
        'hint_2': {'check_item': '본드 변경 후 Fo 이동 추적',
                   'reason': 'SJ4774 Fo 662.91Hz(Normal +1.64%); VE562850 637.46Hz(Normal -2.27%); 모두 675±120Hz 내.'},
        'hint_3': {'check_item': '본드당 n>=10 및 RAW 원본 보존',
                   'reason': 'RAW DATA에 본드별 10pcs의 주파수별 원자료 보존, 출하 후 sigma 비교 가능.'},
    },
    'log': {
        'assumptions': ['GRAPH 행 값을 SPL 100-750, SPL 800-1.5k, SPL 1.6-14k, Fo, THD200, THD400, THD1k 순으로 해석.', '스펙 내는 PASS.'],
        'warnings': ['TSV 인코딩 깨짐(��) - 일부 헤더 라벨 추론.', 'Decision 텍스트 없음; 수치 기반.', '본드당 n=10.'],
        'decision_rationale': '두 신본드 모두 스펙 내. SJ4774: 최대 편차 THD400 +10.82%, THD1k +8.70%, Fo +1.64% Center 근접. VE562850: 최대 편차 Fo -2.27%, THD400 +2.52%, SPL 전 대역 ±0.2dB 내. 전체적으로 VE562850이 Normal에 더 가깝고, SJ4774는 Fo 근접성을 대가로 THD가 다소 상승.'
    }
}
tr_vi_08 = {
    'document': {
        'title': 'MSU-L20S15-07 - Bond mới SJ4774 và VE562850 - kết quả NTI lot test',
        'purpose': 'Qualify bond mới SJ4774 và VE562850 với Normal lot bằng acoustic spec (SPL bands, Fo, THD 200/400/1000Hz).',
        'content': [
            'GRAPH sheet: delta AVG mỗi lot so với spec Center cho SPL 100-750Hz (109.2±3dB), SPL 800-1.5k (118.3±2dB), SPL 1.6-14k (115.8±3dB), Fo (675±120Hz), THD 200/400/1000Hz (upper 45/25/8).',
            'RAW DATA sheet: SPL/THD/Fo theo từng tần số cho Normal, SJ4774, VE562850 (mỗi nhóm 10 sample).',
            'Lot so sánh: STD, Normal, SJ4774, VE562850.'
        ]
    },
    'conclusions': {
        'concl_1': {'topic': 'Delta acoustic SJ4774 vs Normal',
                    'statement_from_report': 'TSV không có Decision rõ ràng; GRAPH thể hiện AVG và delta so với spec Center.',
                    'normalized_interpretation': 'SJ4774 vs Normal AVG: SPL 100-750 108.04 vs 108.42 (-0.38dB, -0.35%), SPL 800-1.5k 117.07 vs 117.18 (-0.11dB, -0.09%), SPL 1.6-14k 115.19 vs 115.15 (+0.04dB, +0.03%), Fo 662.91 vs 652.24 (+10.67Hz, +1.64%), THD200 31.73 vs 30.86 (+0.87pp, +2.82%), THD400 13.62 vs 12.29 (+1.33pp, +10.82%), THD1k 4.25 vs 3.91 (+0.34pp, +8.70%). Tất cả trong spec; sai biệt lớn nhất THD400 +10.82%.'},
        'concl_2': {'topic': 'Delta acoustic VE562850 vs Normal',
                    'statement_from_report': 'TSV không có Decision rõ ràng.',
                    'normalized_interpretation': 'VE562850 vs Normal AVG: SPL 100-750 108.54 vs 108.42 (+0.12dB, +0.11%), SPL 800-1.5k 116.99 vs 117.18 (-0.19dB, -0.16%), SPL 1.6-14k 115.14 vs 115.15 (-0.01dB, -0.01%), Fo 637.46 vs 652.24 (-14.78Hz, -2.27%), THD200 30.25 vs 30.86 (-0.61pp, -1.98% cải thiện), THD400 12.60 vs 12.29 (+0.31pp, +2.52%), THD1k 4.05 vs 3.91 (+0.14pp, +3.58%). Tất cả trong spec; gần Normal hơn SJ4774.'},
        'concl_3': {'topic': 'Tổng kết xếp hạng bond',
                    'statement_from_report': 'Không có Decision rõ ràng.',
                    'normalized_interpretation': 'Cả hai trong spec. VE562850 bám Normal sát hơn SJ4774, đặc biệt THD400 (+2.52% vs +10.82%) và THD1k (+3.58% vs +8.70%). Fo: SJ4774 +1.64% gần spec Center 675Hz hơn VE562850 (-2.27%). Không có fail spec với n=10.'},
    },
    'hints': {
        'hint_1': {'check_item': 'So sánh AVG bond ứng viên với Normal cùng event ở các SPL band, Fo, THD band',
                   'reason': 'SJ4774 THD400 +10.82% (12.29 → 13.62); VE562850 THD400 +2.52% (12.29 → 12.60); cả hai trong upper 25%.'},
        'hint_2': {'check_item': 'Theo dõi Fo shift sau khi đổi bond',
                   'reason': 'SJ4774 Fo 662.91Hz (Normal +1.64%); VE562850 637.46Hz (Normal -2.27%); cả hai trong window 675 ±120Hz.'},
        'hint_3': {'check_item': 'Chạy n>=10 mỗi bond và lưu raw từng sample',
                   'reason': 'RAW DATA lưu 10 sample/bond, các tần số, cho phép so sánh sigma sau release.'},
    },
    'log': {
        'assumptions': ['Đọc giá trị GRAPH theo thứ tự cột: SPL 100-750, SPL 800-1.5k, SPL 1.6-14k, Fo, THD200, THD400, THD1k.', 'AVG trong spec coi như PASS.'],
        'warnings': ['TSV bị artifact mã hóa (��); ánh xạ cột chỉ gần đúng.', 'Không có Decision text.', 'n=10 / bond.'],
        'decision_rationale': 'Cả hai bond mới trong spec. SJ4774 vs Normal: sai biệt lớn nhất THD400 +10.82%; THD1k +8.70%; Fo +1.64% gần spec Center. VE562850 vs Normal: sai biệt lớn nhất Fo -2.27%; THD400 +2.52%; SPL trong ±0.2dB. Tổng thể VE562850 bám Normal sát hơn; SJ4774 đổi Fo gần spec hơn lấy THD cao hơn.'
    }
}
run(name08, result08, tr_ko_08, tr_en_08, tr_vi_08)
