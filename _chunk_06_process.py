"""Process chunk 06 AI Batch normalization."""
from __future__ import annotations
import sys, os, io, importlib.util, json

if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Load helper
spec = importlib.util.spec_from_file_location('h', os.path.join(os.path.dirname(__file__), '_ai_batch_helper.py'))
h = importlib.util.module_from_spec(spec); spec.loader.exec_module(h)  # type: ignore


def ds1():
    """3. BRS-161016 (DT) Report test new machine dry date 2024.04.02."""
    name = '3. BRS-161016 (DT) Report  test new machine dry date 2024.04.02'
    result = {
        'schema_version': '0.1',
        'document': {
            'document_id': '', 'source_file': name, 'source_sheet': 'Sub 3,Yoke,BP.SM',
            'title': "BRS-161016 DT Report test new machine dry for Process Ass'y Frame+SUS (Sub3), BPT/CMG (Yoke), BPT/SMG (BP.SM)",
            'model': 'BRS-161016', 'report_date': '2024-04-02',
            'department': 'ME', 'marker': 'Le/Nhung', 'line': 'C2-3A',
            'report_type': 'mixed',
            'primary_defect': {'canonical_name': 'NG Separate SM', 'aliases_in_document': ['Separate SM', 'Glue not dry', 'Separate', 'Offset']},
            'related_defects': ['Offset', 'Glue not dry', 'NG Hearing Noise', 'NG Hearing Touch'],
            'parts': ['Frame', 'Suspension', 'BPT', 'CMG', 'SMG'],
            'processes': ["Ass'y Frame+SUS", "Ass'y BPT/CMG", "Ass'y BPT/SMG", 'Dry'],
            'purpose': 'Check whether new dry machine can be used for BRS-161016 across Frame+SUS, BPT/CMG, BPT/SMG processes.',
            'content': [
                'Sub3: Press JIG dry test 230C @ 5/7/9 min comparing JIG-with-MG vs 1body JIG, plus tension test.',
                'Yoke (BPT/CMG): semi-product dry test at 230C and 195C with Decap/Tension/Drop and Gauss before-after gap, plus function-line comparison Test vs Normal.',
                'BP.SM (BPT/SMG): dry test at 195C and 230C with 5/7/9/12/14 min, tension test and decision.'
            ],
            'source_cells': {'title': ['Sub 3!B1', 'Yoke!B1', 'BP.SM!B1'], 'date': ['Sub 3', 'Yoke', 'BP.SM'], 'purpose': ['I. Purpose'], 'content': ['II. Content', 'III. Result']}
        },
        'test_conditions': [
            {'condition_id': 'cond_1', 'condition_group': 'Sub3 Frame+SUS dry', 'process': "Ass'y Frame+SUS dry", 'changed_factor': 'Press JIG type and dry time', 'before_value': 'Press JIG 1body (Press 0.05)', 'after_value': 'Press JIG have MG (Press 0.11)', 'temperature': '230C', 'sheet_name': 'Sub 3', 'source_file': name, 'source_cells': ['Sub 3!Items test']},
            {'condition_id': 'cond_2', 'condition_group': 'Yoke BPT/CMG dry', 'process': "Ass'y BPT/CMG dry", 'changed_factor': 'Temp / Time', 'before_value': 'Normal 102C 6min', 'after_value': 'Test 230C/195C, 3-5 min', 'temperature': '230C/195C', 'sheet_name': 'Yoke', 'source_file': name, 'source_cells': ['Yoke!Items test']},
            {'condition_id': 'cond_3', 'condition_group': 'BPT/SMG dry', 'process': "Ass'y BPT/SMG dry", 'changed_factor': 'Time at 195C/230C', 'before_value': "105C Normal 20'55", 'after_value': "195C/230C 5-14 min", 'temperature': '195C/230C', 'sheet_name': 'BP.SM', 'source_file': name, 'source_cells': ['BP.SM!Items test']},
        ],
        'results': [
            # Sub3 Frame+SUS dry results
            {'result_id': 'r1', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'Sub3 230C 5min JIG MG', 'date': '2024-04-02', 'input_count': 4, 'ok_count': 0, 'ng_count': 4, 'ng_rate_percent': 100.0, 'metric_name': 'NG Rate Frame+SUS Vision', 'sheet_name': 'Sub 3', 'source_file': name, 'ng_breakdown': {'Offset': 0, 'Separate': 0, 'Glue not dry': 4}, 'source_cells': ['Sub 3!230C-5 row']},
            {'result_id': 'r2', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'Sub3 230C 5min JIG 1body', 'date': '2024-04-02', 'input_count': 4, 'ok_count': 0, 'ng_count': 4, 'ng_rate_percent': 100.0, 'metric_name': 'NG Rate Frame+SUS Vision', 'sheet_name': 'Sub 3', 'source_file': name, 'ng_breakdown': {'Offset': 0, 'Separate': 1, 'Glue not dry': 3}, 'source_cells': ['Sub 3!230C-5 row 1body']},
            {'result_id': 'r3', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'Sub3 230C 7min JIG MG', 'date': '2024-04-02', 'input_count': 4, 'ok_count': 4, 'ng_count': 0, 'ng_rate_percent': 0.0, 'metric_name': 'NG Rate Frame+SUS Vision', 'sheet_name': 'Sub 3', 'source_file': name, 'source_cells': ['Sub 3!230C-7 MG']},
            {'result_id': 'r4', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'Sub3 230C 7min JIG 1body', 'date': '2024-04-02', 'input_count': 4, 'ok_count': 0, 'ng_count': 4, 'ng_rate_percent': 100.0, 'metric_name': 'NG Rate Frame+SUS Vision', 'sheet_name': 'Sub 3', 'source_file': name, 'ng_breakdown': {'Offset': 0, 'Separate': 0, 'Glue not dry': 4}, 'source_cells': ['Sub 3!230C-7 1body']},
            {'result_id': 'r5', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'Sub3 230C 9min JIG MG', 'date': '2024-04-02', 'input_count': 4, 'ok_count': 4, 'ng_count': 0, 'ng_rate_percent': 0.0, 'metric_name': 'NG Rate Frame+SUS Vision', 'sheet_name': 'Sub 3', 'source_file': name, 'source_cells': ['Sub 3!230C-9 MG']},
            {'result_id': 'r6', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'Sub3 230C 9min JIG 1body', 'date': '2024-04-02', 'input_count': 4, 'ok_count': 4, 'ng_count': 0, 'ng_rate_percent': 0.0, 'metric_name': 'NG Rate Frame+SUS Vision', 'sheet_name': 'Sub 3', 'source_file': name, 'source_cells': ['Sub 3!230C-9 1body']},
            # Sub3 Tension
            {'result_id': 'r7', 'condition_id': 'cond_1', 'measurement_type': 'Tension', 'condition_group': 'Sub3 Tension all conditions', 'date': '2024-04-03', 'metric_name': 'Tension Fr/SP min AVG', 'metric_value': 0.284, 'unit': 'kgf', 'judgement': 'PASS', 'sheet_name': 'Sub 3', 'source_file': name, 'source_cells': ['Sub 3!Tension section']},
            # Yoke function line Test vs Normal C2-3A
            {'result_id': 'r8', 'condition_id': 'cond_2', 'measurement_type': 'Function', 'condition_group': 'Yoke C2-3A Test', 'date': '2024-03-18', 'line': 'C2-3A', 'input_count': 350, 'ok_count': 346, 'ng_count': 4, 'ng_rate_percent': 1.14, 'metric_name': 'Function Total NG Rate Test', 'sheet_name': 'Yoke', 'source_file': name, 'ng_breakdown': {'SPL': 0, 'THD': 0, 'SPL+THD': 0, 'SPL+THD+F0': 0, 'Hearing Noise': 4, 'Hearing Touch': 0}, 'source_cells': ['Yoke!3/18 C2-3A Test']},
            {'result_id': 'r9', 'condition_id': 'cond_2', 'measurement_type': 'Function', 'condition_group': 'Yoke C2-3A Normal', 'date': '2024-03-18', 'line': 'C2-3A', 'input_count': 450, 'ok_count': 445, 'ng_count': 5, 'ng_rate_percent': 1.11, 'metric_name': 'Function Total NG Rate Normal', 'sheet_name': 'Yoke', 'source_file': name, 'ng_breakdown': {'Hearing Noise': 5}, 'source_cells': ['Yoke!3/18 C2-3A Normal']},
            {'result_id': 'r10', 'condition_id': 'cond_2', 'measurement_type': 'Function', 'condition_group': 'Yoke C2-3A Test', 'date': '2024-03-19', 'line': 'C2-3A', 'input_count': 330, 'ok_count': 323, 'ng_count': 7, 'ng_rate_percent': 2.12, 'metric_name': 'Function Total NG Rate Test', 'sheet_name': 'Yoke', 'source_file': name, 'ng_breakdown': {'Hearing Noise': 7}, 'source_cells': ['Yoke!3/19 C2-3A Test']},
            {'result_id': 'r11', 'condition_id': 'cond_2', 'measurement_type': 'Function', 'condition_group': 'Yoke C2-3A Normal', 'date': '2024-03-19', 'line': 'C2-3A', 'input_count': 559, 'ok_count': 536, 'ng_count': 23, 'ng_rate_percent': 4.11, 'metric_name': 'Function Total NG Rate Normal', 'sheet_name': 'Yoke', 'source_file': name, 'ng_breakdown': {'Hearing Noise': 21, 'Hearing Touch': 2}, 'source_cells': ['Yoke!3/19 C2-3A Normal']},
            {'result_id': 'r12', 'condition_id': 'cond_2', 'measurement_type': 'Function', 'condition_group': 'Yoke C2-3A Test', 'date': '2024-03-20', 'line': 'C2-3A', 'input_count': 320, 'ok_count': 305, 'ng_count': 15, 'ng_rate_percent': 4.69, 'metric_name': 'Function Total NG Rate Test', 'sheet_name': 'Yoke', 'source_file': name, 'ng_breakdown': {'Hearing Noise': 15}, 'source_cells': ['Yoke!3/20 C2-3A Test']},
            {'result_id': 'r13', 'condition_id': 'cond_2', 'measurement_type': 'Function', 'condition_group': 'Yoke C2-3A Normal', 'date': '2024-03-20', 'line': 'C2-3A', 'input_count': 559, 'ok_count': 533, 'ng_count': 26, 'ng_rate_percent': 4.65, 'metric_name': 'Function Total NG Rate Normal', 'sheet_name': 'Yoke', 'source_file': name, 'ng_breakdown': {'Hearing Noise': 24, 'Hearing Touch': 2}, 'source_cells': ['Yoke!3/20 C2-3A Normal']},
            # BP.SM dry results
            {'result_id': 'r14', 'condition_id': 'cond_3', 'measurement_type': 'Vision', 'condition_group': 'BP.SM 195C 5min', 'date': '2024-03-09', 'input_count': 8, 'ok_count': 0, 'ng_count': 8, 'ng_rate_percent': 100.0, 'metric_name': 'NG Rate Separate SM', 'sheet_name': 'BP.SM', 'source_file': name, 'source_cells': ['BP.SM!195C 5min']},
            {'result_id': 'r15', 'condition_id': 'cond_3', 'measurement_type': 'Vision', 'condition_group': 'BP.SM 195C 7min', 'date': '2024-03-09', 'input_count': 8, 'ok_count': 2, 'ng_count': 6, 'ng_rate_percent': 75.0, 'metric_name': 'NG Rate Separate SM', 'sheet_name': 'BP.SM', 'source_file': name, 'source_cells': ['BP.SM!195C 7min']},
            {'result_id': 'r16', 'condition_id': 'cond_3', 'measurement_type': 'Vision', 'condition_group': 'BP.SM 195C 9min', 'date': '2024-03-09', 'input_count': 8, 'ok_count': 3, 'ng_count': 5, 'ng_rate_percent': 62.5, 'metric_name': 'NG Rate Separate SM', 'sheet_name': 'BP.SM', 'source_file': name, 'source_cells': ['BP.SM!195C 9min']},
            {'result_id': 'r17', 'condition_id': 'cond_3', 'measurement_type': 'Vision', 'condition_group': 'BP.SM 195C 12min', 'date': '2024-03-09', 'input_count': 112, 'ok_count': 108, 'ng_count': 4, 'ng_rate_percent': 3.6, 'metric_name': 'NG Rate Separate SM', 'sheet_name': 'BP.SM', 'source_file': name, 'source_cells': ['BP.SM!195C 12min']},
            {'result_id': 'r18', 'condition_id': 'cond_3', 'measurement_type': 'Vision', 'condition_group': 'BP.SM 195C 14min', 'date': '2024-03-09', 'input_count': 112, 'ok_count': 112, 'ng_count': 0, 'ng_rate_percent': 0.0, 'metric_name': 'NG Rate Separate SM', 'sheet_name': 'BP.SM', 'judgement': 'OK', 'source_file': name, 'source_cells': ['BP.SM!195C 14min']},
            {'result_id': 'r19', 'condition_id': 'cond_3', 'measurement_type': 'Vision', 'condition_group': 'BP.SM Normal', 'date': '2024-03-09', 'input_count': 500, 'ok_count': 500, 'ng_count': 0, 'ng_rate_percent': 0.0, 'metric_name': 'NG Rate Separate SM', 'sheet_name': 'BP.SM', 'judgement': 'OK', 'source_file': name, 'source_cells': ['BP.SM!Normal 105C']},
            {'result_id': 'r20', 'condition_id': 'cond_3', 'measurement_type': 'Vision', 'condition_group': 'BP.SM 195C 14min 2nd', 'date': '2024-03-11', 'input_count': 216, 'ok_count': 208, 'ng_count': 8, 'ng_rate_percent': 3.7, 'metric_name': 'NG Rate Separate SM', 'sheet_name': 'BP.SM', 'source_file': name, 'source_cells': ['BP.SM!11/Mar 14min']},
            {'result_id': 'r21', 'condition_id': 'cond_3', 'measurement_type': 'Vision', 'condition_group': 'BP.SM 230C 12min', 'date': '2024-03-12', 'input_count': 216, 'ok_count': 206, 'ng_count': 10, 'ng_rate_percent': 4.6, 'metric_name': 'NG Rate Separate SM', 'sheet_name': 'BP.SM', 'source_file': name, 'source_cells': ['BP.SM!12/Mar 230C 12min']},
            {'result_id': 'r22', 'condition_id': 'cond_3', 'measurement_type': 'Vision', 'condition_group': 'BP.SM 230C 14min', 'date': '2024-03-12', 'input_count': 216, 'ok_count': 210, 'ng_count': 6, 'ng_rate_percent': 2.8, 'metric_name': 'NG Rate Separate SM', 'sheet_name': 'BP.SM', 'source_file': name, 'source_cells': ['BP.SM!12/Mar 230C 14min']},
            {'result_id': 'r23', 'condition_id': 'cond_3', 'measurement_type': 'Tension', 'condition_group': 'BP.SM 195C 14min SM Long', 'date': '2024-03-09', 'metric_name': 'Tension SM Long min', 'metric_value': 4.03, 'unit': 'kgf', 'judgement': 'FAIL', 'sheet_name': 'BP.SM', 'source_file': name, 'source_cells': ['BP.SM!Tension 14min SM Long']},
        ],
        'conclusions': [
            {'conclusion_id': 'concl_1', 'topic': "Sub3 Frame+SUS dry", 'statement_from_report': "230C-7' and 230C-9' Press JIG have MG gave 0% NG; 5min and 1body still showed Glue not dry.",
             'normalized_interpretation': 'For Sub3 Frame+SUS dry at 230C, 7-9 min with Press JIG having MG eliminates NG Glue not dry; tension PASS in all conditions (min AVG 0.284 kgf vs spec >=0.2 kgf).',
             'sheet_name': 'Sub 3', 'source_file': name, 'source_cells': ['Sub 3!IV. Decision']},
            {'conclusion_id': 'concl_2', 'topic': 'Yoke BPT/CMG dry', 'statement_from_report': '=> Can use new dry machine with process Dry BPT/CMG',
             'normalized_interpretation': 'Yoke BPT/CMG: new dry machine accepted. Function line C2-3A NG rate: 2024-03-18 Test 1.14% vs Normal 1.11% = 1.03x (2.7% worse than same-event normal); 2024-03-19 Test 2.12% vs Normal 4.11% = 0.52x (48.4% improved); 2024-03-20 Test 4.69% vs Normal 4.65% = 1.01x (0.9% worse). All Tension PASS, Drop test 0 NG.',
             'sheet_name': 'Yoke', 'source_file': name, 'source_cells': ['Yoke!IV. Decision']},
            {'conclusion_id': 'concl_3', 'topic': 'BPT/SMG dry', 'statement_from_report': '=> Can not use new dry machine with process Dry SM/BPT',
             'normalized_interpretation': 'BPT/SMG: new dry machine rejected. SM Separate NG only reaches 0% at 195C 14min initial test but reappears at 195C 14min 2nd run (3.7%) and 230C 12-14min (4.6%/2.8%). Tension SM Long 195C 14min fails at 4.03 kgf vs spec >=5.0 kgf.',
             'sheet_name': 'BP.SM', 'source_file': name, 'source_cells': ['BP.SM!IV. Decision']},
        ],
        'troubleshooting_index': {
            'defect_name': 'NG Separate SM / Glue not dry',
            'when_user_asks': ['How to set dry machine for BPT/SMG?', 'Which dry time eliminates NG Glue not dry on Frame+SUS at 230C?'],
            'suggested_checks': [
                {'hint_id': 'hint_1', 'check_item': 'Dry time at 230C for Frame+SUS', 'reason': 'NG Glue not dry only disappears at 230C 7-9 min when Press JIG has MG; 5 min still shows 100% NG.', 'evidence_strength': 'strong', 'related_process': "Sub3 Ass'y Frame+SUS dry", 'related_part': 'Frame+SUS', 'sheet_name': 'Sub 3', 'source_file': name, 'source_cells': ['Sub 3!Result table']},
                {'hint_id': 'hint_2', 'check_item': 'Tension SM Long after dry', 'reason': 'BPT/SMG 195C 14min SM Long Tension min 4.03 kgf failed vs spec >=5.0 kgf; tension margin must be re-verified before accepting new dry machine.', 'evidence_strength': 'strong', 'related_process': "BPT/SMG dry", 'related_part': 'SM Long', 'sheet_name': 'BP.SM', 'source_file': name, 'source_cells': ['BP.SM!Tension 14min']},
                {'hint_id': 'hint_3', 'check_item': 'Hearing Noise dominance on C2-3A', 'reason': 'On Yoke C2-3A function line, all NG counts are Hearing Noise/Touch; SPL/THD/F0 are zero in same-event Test and Normal.', 'evidence_strength': 'medium', 'related_process': 'Function check main line', 'related_part': 'BPT/CMG', 'sheet_name': 'Yoke', 'source_file': name, 'source_cells': ['Yoke!Function table']},
            ],
            'limitations': ['Sub3 Q ty very small (4 pcs per cell); single-event NG rate not statistically robust.'],
        },
        'ai_extraction_log': {
            'confidence': 0.78,
            'assumptions': ['Sub3 IV. Decision text empty; conclusion derived from numeric table.'],
            'warnings': ['Sub3 sample size per cell is only 4 pcs.', 'BP.SM 195C 12min after-action 2nd run regressed to 3.7% NG Separate SM.'],
            'decision_rationale': 'Three sub-reports treated as one mixed document. Function line uses normal_comparison logic with relative_change_percent = (test/baseline-1)*100. Decisions follow IV. Decision text per sheet.'
        }
    }
    tr_ko = {
        'document': {'title': "BRS-161016 DT 신규 건조 장비 사용성 검증 리포트 (Frame+SUS/BPT+CMG/BPT+SMG)",
                     'purpose': "신규 건조 장비를 BRS-161016 모델의 Frame+SUS, BPT/CMG, BPT/SMG 공정에 사용할 수 있는지 확인.",
                     'content': ['Sub3: 230C에서 5/7/9분 dry, Press JIG MG vs 1body 비교 및 Tension.', 'Yoke: 230C/195C dry, Decap/Tension/Drop/Gauss before-after, Function line C2-3A Test vs Normal.', 'BP.SM: 195C/230C dry 5~14분, Tension SM Short/SM Long.']},
        'conclusions': {
            'concl_1': {'topic': 'Sub3 Frame+SUS dry', 'statement_from_report': "230C 7'/9' Press JIG MG 사용 시 NG 0%, 5'와 1body는 Glue not dry 발생.", 'normalized_interpretation': "Sub3 Frame+SUS dry 230C 7~9분 + Press JIG MG 사용 시 NG Glue not dry 제거됨. Tension 모든 조건 PASS (min AVG 0.284 kgf, spec >=0.2 kgf)."},
            'concl_2': {'topic': 'Yoke BPT/CMG dry', 'statement_from_report': '신규 건조 장비 BPT/CMG 사용 가능', 'normalized_interpretation': 'Yoke BPT/CMG 신규 건조 장비 사용 가능. 동일 이벤트 Function C2-3A NG rate: 2024-03-18 Test 1.14% vs Normal 1.11% = 1.03배(2.7% 악화), 2024-03-19 Test 2.12% vs Normal 4.11% = 0.52배(48.4% 개선), 2024-03-20 Test 4.69% vs Normal 4.65% = 1.01배(0.9% 악화). Tension/Drop 전부 PASS.'},
            'concl_3': {'topic': 'BPT/SMG dry', 'statement_from_report': '신규 건조 장비 BPT/SMG 사용 불가', 'normalized_interpretation': 'BPT/SMG: 신규 건조 장비 불합격. NG Separate SM이 195C 14분 1차에서만 0%, 2차 14분 3.7%, 230C 12~14분 4.6%/2.8%로 재발. Tension SM Long 195C 14분 min 4.03 kgf로 spec >=5.0 kgf 미달.'},
        },
        'hints': {
            'hint_1': {'check_item': '230C에서 Frame+SUS dry 시간', 'reason': '230C 7~9분 + Press JIG MG 일 때만 NG Glue not dry 0%, 5분은 100% NG 발생.'},
            'hint_2': {'check_item': 'dry 후 Tension SM Long', 'reason': 'BPT/SMG 195C 14분 SM Long Tension min 4.03 kgf로 spec >=5.0 kgf fail. 신규 건조 장비 채택 전 Tension margin 재확인 필요.'},
            'hint_3': {'check_item': 'C2-3A NG Hearing Noise 우세', 'reason': 'Yoke C2-3A function 라인의 NG는 동일 이벤트 Test/Normal 모두 Hearing Noise/Touch만 발생, SPL/THD/F0는 0.'},
        },
        'log': {'assumptions': ['Sub3 IV. Decision 빈 칸; 표 데이터로부터 결론 도출.'],
                'warnings': ['Sub3 셀별 4pcs로 표본 작음.', 'BP.SM 195C 12분 2차 시도 시 3.7% NG Separate SM 재발.'],
                'decision_rationale': 'Function 라인은 same-event Normal 대조군 비교(상대 변화율 (test/baseline-1)*100). IV. Decision 텍스트 우선.'}
    }
    tr_en = {
        'document': {'title': result['document']['title'], 'purpose': result['document']['purpose'], 'content': result['document']['content']},
        'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in result['conclusions']},
        'hints': {h['hint_id']: {'check_item': h['check_item'], 'reason': h['reason']} for h in result['troubleshooting_index']['suggested_checks']},
        'log': {'assumptions': result['ai_extraction_log']['assumptions'], 'warnings': result['ai_extraction_log']['warnings'], 'decision_rationale': result['ai_extraction_log']['decision_rationale']},
    }
    tr_vi = {
        'document': {'title': 'BRS-161016 DT Report kiểm tra máy dry mới cho Frame+SUS/BPT+CMG/BPT+SMG',
                     'purpose': 'Kiểm tra máy dry mới có dùng được cho model BRS-161016 ở các process Frame+SUS, BPT/CMG, BPT/SMG hay không.',
                     'content': ['Sub3: dry 230C 5/7/9 phút, Press JIG có MG vs 1body, Tension.', 'Yoke: dry 230C/195C, Decap/Tension/Drop/Gauss before-after, line Function C2-3A Test vs Normal.', 'BP.SM: dry 195C/230C 5~14 phút, Tension SM Short/SM Long.']},
        'conclusions': {
            'concl_1': {'topic': 'Sub3 Frame+SUS dry', 'statement_from_report': "230C 7'/9' Press JIG có MG: NG 0%; 5' và 1body vẫn còn Glue not dry.", 'normalized_interpretation': 'Sub3 Frame+SUS dry 230C 7~9 phút + Press JIG có MG: NG Glue not dry hết. Tension mọi điều kiện PASS (min AVG 0.284 kgf, spec >=0.2 kgf).'},
            'concl_2': {'topic': 'Yoke BPT/CMG dry', 'statement_from_report': 'Có thể dùng máy dry mới cho BPT/CMG', 'normalized_interpretation': 'Yoke BPT/CMG dùng được máy dry mới. Cùng sự kiện Function C2-3A NG rate: 2024-03-18 Test 1.14% vs Normal 1.11% = 1.03x (xấu hơn 2.7%); 2024-03-19 Test 2.12% vs Normal 4.11% = 0.52x (cải thiện 48.4%); 2024-03-20 Test 4.69% vs Normal 4.65% = 1.01x (xấu hơn 0.9%). Tension/Drop đều PASS.'},
            'concl_3': {'topic': 'BPT/SMG dry', 'statement_from_report': 'Không thể dùng máy dry mới cho SM/BPT', 'normalized_interpretation': 'BPT/SMG: từ chối máy dry mới. NG Separate SM chỉ về 0% ở 195C 14 phút lần 1, lần 2 lên 3.7%, 230C 12-14 phút 4.6%/2.8%. Tension SM Long 195C 14 phút min 4.03 kgf không đạt spec >=5.0 kgf.'},
        },
        'hints': {
            'hint_1': {'check_item': 'Thời gian dry 230C cho Frame+SUS', 'reason': 'Chỉ 230C 7~9 phút + Press JIG có MG mới hết NG Glue not dry; 5 phút vẫn 100% NG.'},
            'hint_2': {'check_item': 'Tension SM Long sau dry', 'reason': 'BPT/SMG 195C 14 phút SM Long Tension min 4.03 kgf fail vs spec >=5.0 kgf; cần re-verify margin trước khi accept máy dry mới.'},
            'hint_3': {'check_item': 'C2-3A NG Hearing Noise chiếm chủ đạo', 'reason': 'Trên line C2-3A của Yoke function, NG chủ yếu Hearing Noise/Touch; SPL/THD/F0 đều 0 ở cả Test và Normal cùng event.'},
        },
        'log': {'assumptions': ['Sub3 IV. Decision trống; kết luận lấy từ bảng số liệu.'],
                'warnings': ['Sub3 mẫu 4pcs mỗi ô, cỡ mẫu nhỏ.', 'BP.SM 195C 12 phút lần 2 hồi quy 3.7% NG Separate SM.'],
                'decision_rationale': 'Line Function dùng so sánh same-event Normal (relative_change_percent = (test/baseline-1)*100). Ưu tiên IV. Decision text.'}
    }
    return name, result, tr_ko, tr_en, tr_vi


def ds2():
    """3. BRS-161016 Report test new press coil SP of UC machine 2024.12.05."""
    name = '3. BRS-161016 Report test new press coil SP of UC machine 2024.12.05'
    result = {
        'schema_version': '0.1',
        'document': {
            'document_id': '', 'source_file': name, 'source_sheet': '1',
            'title': 'REPORT TEST NEW PRESS COIL SP OF UC MACHINE - BRS-161016',
            'model': 'BRS-161016', 'report_date': '2024-12-05', 'department': 'ME', 'marker': 'Hương',
            'line': 'E2-3A/E2-3B',
            'report_type': 'normal_comparison',
            'primary_defect': {'canonical_name': 'NG Coil SP Assembly', 'aliases_in_document': ['NG over glue', 'NG gap coil SP', 'NG separate', 'NG SP bending']},
            'related_defects': ['Over Glue', 'NG Gap Coil SP', 'Separation', 'SP Bending', 'NG Hearing Noise'],
            'parts': ['Coil SP'], 'processes': ["Ass'y coil SP", 'Press', 'UV LED'],
            'purpose': "Check whether new UC machine press for coil SP is OK or NG to improve NG at process ass'y coil SP.",
            'content': ["Setting new press coil SP and check NG coil SP after ass'y", 'Check function and compare with normal old press.', 'Check position press NG of new jig.'],
            'source_cells': {'title': ['1!B1'], 'date': ['1!Date'], 'purpose': ['1!I. Purpose'], 'content': ['1!II. Content']},
        },
        'test_conditions': [
            {'condition_id': 'cond_1', 'condition_group': 'New vs Old press', 'process': "Ass'y coil SP press", 'changed_factor': 'Press machine type', 'before_value': 'Use old press coil SP', 'after_value': 'Use new press coil SP', 'machine': 'UC press machine', 'sheet_name': '1', 'source_file': name, 'source_cells': ['1!Test table']},
            {'condition_id': 'cond_2', 'condition_group': 'Press position on new JIG', 'process': "Ass'y coil SP press", 'changed_factor': 'Position press (1/2/3/4)', 'sheet_name': '1', 'source_file': name, 'source_cells': ['1!Position press table']},
        ],
        'results': [
            {'result_id': 'r1', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'E2-3B New press 1st', 'date': '2024-12-05', 'line': 'E2-3B', 'input_count': 300, 'ng_count': 1, 'ng_rate_percent': 0.3, 'metric_name': "NG Rate ass'y coil SP", 'sheet_name': '1', 'source_file': name, 'ng_breakdown': {'Over glue': 0, 'Gap coil SP': 1, 'Separate': 0, 'SP bending': 0}, 'source_cells': ['1!12/5 E2-3B new']},
            {'result_id': 'r2', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'E2-3B Old press 1st', 'date': '2024-12-05', 'line': 'E2-3B', 'input_count': 200, 'ng_count': 1, 'ng_rate_percent': 0.5, 'metric_name': "NG Rate ass'y coil SP", 'sheet_name': '1', 'source_file': name, 'ng_breakdown': {'Over glue': 0, 'Gap coil SP': 1, 'Separate': 0, 'SP bending': 0}, 'source_cells': ['1!12/5 E2-3B old']},
            {'result_id': 'r3', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'E2-3A New press 1st', 'date': '2024-12-05', 'line': 'E2-3A', 'input_count': 100, 'ng_count': 1, 'ng_rate_percent': 1.0, 'metric_name': "NG Rate ass'y coil SP", 'sheet_name': '1', 'source_file': name, 'ng_breakdown': {'Gap coil SP': 1}, 'source_cells': ['1!E2-3A new 1st']},
            {'result_id': 'r4', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'E2-3A New press 2nd', 'date': '2024-12-05', 'line': 'E2-3A', 'input_count': 18, 'ng_count': 6, 'ng_rate_percent': 33.3, 'metric_name': "NG Rate ass'y coil SP", 'sheet_name': '1', 'source_file': name, 'ng_breakdown': {'Over glue': 4, 'Gap coil SP': 2}, 'source_cells': ['1!E2-3A new 2nd']},
            {'result_id': 'r5', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'E2-3A New press 3rd', 'date': '2024-12-05', 'line': 'E2-3A', 'input_count': 30, 'ng_count': 3, 'ng_rate_percent': 10.0, 'metric_name': "NG Rate ass'y coil SP", 'sheet_name': '1', 'source_file': name, 'ng_breakdown': {'Over glue': 3}, 'source_cells': ['1!E2-3A new 3rd']},
            {'result_id': 'r6', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'E2-3A New press 4th', 'date': '2024-12-05', 'line': 'E2-3A', 'input_count': 16, 'ng_count': 1, 'ng_rate_percent': 6.2, 'metric_name': "NG Rate ass'y coil SP", 'sheet_name': '1', 'source_file': name, 'ng_breakdown': {'Gap coil SP': 1}, 'source_cells': ['1!E2-3A new 4th']},
            {'result_id': 'r7', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'E2-3A Old press 1st', 'date': '2024-12-05', 'line': 'E2-3A', 'input_count': 50, 'ng_count': 0, 'ng_rate_percent': 0.0, 'metric_name': "NG Rate ass'y coil SP", 'sheet_name': '1', 'source_file': name, 'source_cells': ['1!E2-3A old 1st']},
            # Position
            {'result_id': 'r8', 'condition_id': 'cond_2', 'measurement_type': 'Vision', 'condition_group': 'Position 1', 'date': '2024-12-06', 'input_count': 164, 'ng_count': 3, 'ng_rate_percent': 1.8, 'metric_name': 'NG Rate Position press', 'sheet_name': '1', 'source_file': name, 'ng_breakdown': {'Gap': 1, 'Over glue': 0, 'SP bending': 3}, 'source_cells': ['1!Pos 1']},
            {'result_id': 'r9', 'condition_id': 'cond_2', 'measurement_type': 'Vision', 'condition_group': 'Position 2', 'date': '2024-12-06', 'input_count': 164, 'ng_count': 4, 'ng_rate_percent': 2.4, 'metric_name': 'NG Rate Position press', 'sheet_name': '1', 'source_file': name, 'ng_breakdown': {'Gap': 1, 'Over glue': 3}, 'source_cells': ['1!Pos 2']},
            {'result_id': 'r10', 'condition_id': 'cond_2', 'measurement_type': 'Vision', 'condition_group': 'Position 3', 'date': '2024-12-06', 'input_count': 164, 'ng_count': 6, 'ng_rate_percent': 3.7, 'metric_name': 'NG Rate Position press', 'sheet_name': '1', 'source_file': name, 'ng_breakdown': {'Gap': 2, 'Over glue': 4}, 'source_cells': ['1!Pos 3']},
            {'result_id': 'r11', 'condition_id': 'cond_2', 'measurement_type': 'Vision', 'condition_group': 'Position 4', 'date': '2024-12-06', 'input_count': 164, 'ng_count': 4, 'ng_rate_percent': 2.4, 'metric_name': 'NG Rate Position press', 'sheet_name': '1', 'source_file': name, 'ng_breakdown': {'Over glue': 4}, 'source_cells': ['1!Pos 4']},
            # Function
            {'result_id': 'r12', 'condition_id': 'cond_1', 'measurement_type': 'Function', 'condition_group': 'E2-3A Function Test (new)', 'date': '2024-12-04', 'line': 'E2-3A', 'input_count': 1200, 'ng_count': 28, 'ng_rate_percent': 2.3, 'metric_name': 'Function NG Rate (new press)', 'sheet_name': '1', 'source_file': name, 'ng_breakdown': {'THD': 1, 'THD+SPL': 1, 'SPL': 0, 'Hearing Noise': 26, 'Hearing Touch': 0}, 'source_cells': ['1!Function new']},
            {'result_id': 'r13', 'condition_id': 'cond_1', 'measurement_type': 'Function', 'condition_group': 'E2-3A Function Normal (old)', 'date': '2024-12-04', 'line': 'E2-3A', 'input_count': 763, 'ng_count': 16, 'ng_rate_percent': 2.1, 'metric_name': 'Function NG Rate (old press)', 'sheet_name': '1', 'source_file': name, 'ng_breakdown': {'THD': 0, 'THD+SPL': 0, 'SPL': 0, 'Hearing Noise': 16, 'Hearing Touch': 0}, 'source_cells': ['1!Function old']},
        ],
        'conclusions': [
            {'conclusion_id': 'concl_1', 'topic': 'New press coil SP UC vs old', 'statement_from_report': "Result check new jig press UC: ass'y coil SP NG over glue and gap 2.56% (auto setting NG), Old JIG: 0.4%.",
             'normalized_interpretation': "On line E2-3B same-event new vs old press NG ass'y coil SP: 0.3% vs 0.5% = 0.6x (40.0% improved vs old). On E2-3A new press auto setting was unstable (1.0%/33.3%/10.0%/6.2%) until changing back to old (0.0%). Function NG: new 2.3% vs old 2.1% = 1.10x (9.5% worse than same-event normal).",
             'sheet_name': '1', 'source_file': name, 'source_cells': ['1!IV. Decision']},
        ],
        'troubleshooting_index': {
            'defect_name': "NG ass'y coil SP (Over glue / Gap)",
            'when_user_asks': ['Should we adopt new UC press for coil SP?', 'Why E2-3A new press unstable?'],
            'suggested_checks': [
                {'hint_id': 'hint_1', 'check_item': 'Auto setting stability of new UC press on E2-3A', 'reason': "E2-3A new press NG rate jumped 1.0%->33.3%->10.0%->6.2% with repeated auto setting; auto setting could not stabilize -> reverted to old press at 0.0%.", 'evidence_strength': 'strong', 'related_process': "Ass'y coil SP press", 'related_part': 'Coil SP', 'sheet_name': '1', 'source_file': name, 'source_cells': ['1!E2-3A new press runs']},
                {'hint_id': 'hint_2', 'check_item': 'Position press 3 of new JIG', 'reason': 'Position 3 has worst NG rate 3.7% (over glue 4 + gap 2) vs Position 1 1.8%; check JIG positioning at slot 3.', 'evidence_strength': 'medium', 'related_process': "Press position", 'related_part': 'Press JIG', 'sheet_name': '1', 'source_file': name, 'source_cells': ['1!Position press table']},
                {'hint_id': 'hint_3', 'check_item': 'Hearing Noise on Function', 'reason': 'Both new (26/28) and old (16/16) function NG are dominated by Hearing Noise; new vs old function NG rate 2.3% vs 2.1% = 9.5% worse than same-event normal.', 'evidence_strength': 'medium', 'related_process': 'Function line', 'related_part': 'Coil SP', 'sheet_name': '1', 'source_file': name, 'source_cells': ['1!Function table']},
            ],
            'limitations': ['Q ty samples on E2-3A new press runs are small (16~30), single-event NG rates not robust.'],
        },
        'ai_extraction_log': {
            'confidence': 0.75,
            'assumptions': ['IV. Decision text only partially captured; figure 2.56% aggregated by report.'],
            'warnings': ['E2-3A new press unstable across 4 retries; old press restored.'],
            'decision_rationale': 'normal_comparison: each test line has same-event old or normal baseline; relative change computed where pairs exist.'
        },
    }
    tr_ko = {
        'document': {'title': 'BRS-161016 UC 신규 Coil SP 프레스 검증 리포트', 'purpose': "UC 신규 프레스가 Coil SP 공정에 사용 가능한지 확인.",
                     'content': ["신규 프레스 설정 후 ass'y Coil SP NG 확인", '기존 프레스와 Function 비교', '신규 JIG 포지션별 NG 확인.']},
        'conclusions': {'concl_1': {'topic': 'UC 신규 vs 기존 Coil SP 프레스', 'statement_from_report': "신규 JIG ass'y over glue+gap 2.56%, 기존 0.4%.",
                                      'normalized_interpretation': "E2-3B 동일 이벤트 신규 vs 기존 0.3% vs 0.5% = 0.6배(40.0% 개선). E2-3A 신규는 1.0%/33.3%/10.0%/6.2%로 auto setting 불안정 → 기존 복귀(0.0%). Function NG 신규 2.3% vs 기존 2.1% = 1.10배(9.5% 악화)."}},
        'hints': {
            'hint_1': {'check_item': 'E2-3A 신규 UC 프레스 auto setting 안정성', 'reason': "E2-3A 신규 NG 1.0%->33.3%->10.0%->6.2% 재시도에도 안정화 실패, 기존 프레스 복귀 시 0%."},
            'hint_2': {'check_item': '신규 JIG Position 3 점검', 'reason': 'Position 3 NG 3.7%(over glue 4 + gap 2)로 Position 1 1.8% 대비 최악.'},
            'hint_3': {'check_item': 'Function Hearing Noise', 'reason': '신규(26/28)와 기존(16/16) Function NG 모두 Hearing Noise 우세. 신규 vs 기존 2.3% vs 2.1% = 9.5% 악화.'},
        },
        'log': {'assumptions': ['IV. Decision 일부만 캡처; 2.56%는 보고서 집계치.'],
                'warnings': ['E2-3A 신규 프레스 4회 재시도에도 불안정, 기존 프레스 복귀.'],
                'decision_rationale': 'normal_comparison: 페어 존재 시 (test/baseline-1)*100 적용.'}
    }
    tr_en = {
        'document': {'title': result['document']['title'], 'purpose': result['document']['purpose'], 'content': result['document']['content']},
        'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in result['conclusions']},
        'hints': {h['hint_id']: {'check_item': h['check_item'], 'reason': h['reason']} for h in result['troubleshooting_index']['suggested_checks']},
        'log': {'assumptions': result['ai_extraction_log']['assumptions'], 'warnings': result['ai_extraction_log']['warnings'], 'decision_rationale': result['ai_extraction_log']['decision_rationale']},
    }
    tr_vi = {
        'document': {'title': 'BRS-161016 Report kiểm tra press coil SP mới của máy UC', 'purpose': "Kiểm tra press mới của máy UC có dùng được cho ass'y coil SP hay không.",
                     'content': ['Setting press mới và check NG sau khi as\'y', 'Compare Function với press cũ', 'Check NG theo position press của JIG mới.']},
        'conclusions': {'concl_1': {'topic': 'Press coil SP UC mới vs cũ', 'statement_from_report': "JIG mới ass'y over glue+gap 2.56%, JIG cũ 0.4%.",
                                      'normalized_interpretation': "Cùng event E2-3B mới vs cũ 0.3% vs 0.5% = 0.6x (cải thiện 40.0%). E2-3A press mới auto setting không ổn định: 1.0%/33.3%/10.0%/6.2%, phải đổi về press cũ (0.0%). Function NG mới 2.3% vs cũ 2.1% = 1.10x (xấu hơn 9.5% so với normal cùng event)."}},
        'hints': {
            'hint_1': {'check_item': 'Ổn định auto setting press UC mới trên E2-3A', 'reason': 'E2-3A press mới NG nhảy 1.0%->33.3%->10.0%->6.2% sau 4 lần auto setting, không ổn định -> quay về press cũ 0%.'},
            'hint_2': {'check_item': 'Position 3 của JIG mới', 'reason': 'Position 3 NG 3.7% (over glue 4 + gap 2) xấu nhất so với Position 1 1.8%.'},
            'hint_3': {'check_item': 'Hearing Noise trên Function', 'reason': 'NG Function ở cả mới (26/28) và cũ (16/16) chủ yếu Hearing Noise; mới 2.3% vs cũ 2.1% = xấu hơn 9.5%.'},
        },
        'log': {'assumptions': ['IV. Decision chỉ chụp được một phần; 2.56% là số tổng của report.'],
                'warnings': ['Press mới E2-3A không ổn định qua 4 retries, phải đổi về press cũ.'],
                'decision_rationale': 'normal_comparison: dùng (test/baseline-1)*100 cho các cặp có baseline cùng event.'}
    }
    return name, result, tr_ko, tr_en, tr_vi


def ds3():
    """3. BRS-161016 Report test process CM press change from machine semi AM to K3 6AXIS AM Date 2025.2.20."""
    name = '3. BRS-161016 Report test process CM press change from machine semi AM to K3 6AXIS AM Date 2025.2.20'
    result = {
        'schema_version': '0.1',
        'document': {
            'document_id': '', 'source_file': name, 'source_sheet': '18.6',
            'title': 'BRS-161016 - REPORT TEST PROCESS CM PRESS CHANGE FROM MACHINE SEMI AM TO K3 6AXIS AM',
            'model': 'BRS-161016', 'report_date': '2025-02-18', 'department': 'ME', 'marker': 'Le', 'line': '',
            'report_type': 'normal_comparison',
            'primary_defect': {'canonical_name': 'NG Bond Spread (YK)', 'aliases_in_document': ['NG Bond Spread not good']},
            'related_defects': ['NG Bond Spread', 'Drop test NG'], 'parts': ['YK', 'MG', 'PT', 'CMG', 'CPT'],
            'processes': ['CM Press', 'Decap', 'Drop test', 'Tension'],
            'purpose': 'Check whether process CM PRESS changed from SEMI AM to K3 6AXIS AM can be used.',
            'content': ['Make semi at Sub and check: Decap bond YK + MG/PT, Drop test, Tension test.'],
            'source_cells': {'title': ['18.6!B1'], 'date': ['18.6'], 'purpose': ['I. Purpose'], 'content': ['II. Content']},
        },
        'test_conditions': [
            {'condition_id': 'cond_1', 'condition_group': 'CM Press machine swap', 'process': 'CM Press', 'changed_factor': 'CM press machine', 'before_value': 'SEMI AM', 'after_value': 'K3 6AXIS AM', 'machine': 'SEMI AM / K3 6AXIS AM', 'sheet_name': '18.6', 'source_file': name, 'source_cells': ['18.6!Items']},
        ],
        'results': [
            {'result_id': 'r1', 'condition_id': 'cond_1', 'measurement_type': 'Decap', 'condition_group': 'SEMI AM', 'date': '2025-02-18', 'input_count': 8, 'ok_count': 8, 'ng_count': 0, 'ng_rate_percent': 0.0, 'metric_name': 'NG Bond Spread (YK >=80%)', 'sheet_name': '18.6', 'source_file': name, 'source_cells': ['18.6!Decap SEMI']},
            {'result_id': 'r2', 'condition_id': 'cond_1', 'measurement_type': 'Decap', 'condition_group': 'K3 6AXIS AM', 'date': '2025-02-18', 'input_count': 8, 'ok_count': 8, 'ng_count': 0, 'ng_rate_percent': 0.0, 'metric_name': 'NG Bond Spread (YK >=80%)', 'sheet_name': '18.6', 'source_file': name, 'source_cells': ['18.6!Decap K3']},
            {'result_id': 'r3', 'condition_id': 'cond_1', 'measurement_type': 'Drop test', 'condition_group': 'SEMI AM', 'date': '2025-02-18', 'input_count': 8, 'ok_count': 8, 'ng_count': 0, 'ng_rate_percent': 0.0, 'metric_name': 'NG Rate Drop test', 'sheet_name': '18.6', 'source_file': name, 'source_cells': ['18.6!Drop SEMI']},
            {'result_id': 'r4', 'condition_id': 'cond_1', 'measurement_type': 'Drop test', 'condition_group': 'K3 6AXIS AM', 'date': '2025-02-18', 'input_count': 8, 'ok_count': 8, 'ng_count': 0, 'ng_rate_percent': 0.0, 'metric_name': 'NG Rate Drop test', 'sheet_name': '18.6', 'source_file': name, 'source_cells': ['18.6!Drop K3']},
            {'result_id': 'r5', 'condition_id': 'cond_1', 'measurement_type': 'Tension', 'condition_group': 'SEMI AM YK+CMG/CPT', 'date': '2025-02-18', 'metric_name': 'Tension AVG', 'metric_value': 73.69, 'unit': 'kgf', 'judgement': 'OK', 'sheet_name': '18.6', 'source_file': name, 'source_cells': ['18.6!Tension SEMI']},
            {'result_id': 'r6', 'condition_id': 'cond_1', 'measurement_type': 'Tension', 'condition_group': 'K3 6AXIS AM YK+CMG/CPT', 'date': '2025-02-18', 'metric_name': 'Tension AVG', 'metric_value': 78.98, 'unit': 'kgf', 'judgement': 'OK', 'sheet_name': '18.6', 'source_file': name, 'source_cells': ['18.6!Tension K3']},
        ],
        'conclusions': [
            {'conclusion_id': 'concl_1', 'topic': 'CM Press K3 6AXIS AM acceptance', 'statement_from_report': 'IV. Decision row is empty in source.',
             'normalized_interpretation': 'K3 6AXIS AM matches SEMI AM on Decap NG Bond Spread (both 0/8) and Drop test (both 0/8). Tension AVG K3 78.98 kgf vs SEMI 73.69 kgf, both >=80kgf internal spec >=50kgf, both OK. Sample size only 8 pcs each.',
             'sheet_name': '18.6', 'source_file': name, 'source_cells': ['18.6!Result tables']},
        ],
        'troubleshooting_index': {
            'defect_name': 'NG Bond Spread (YK >=80%)',
            'when_user_asks': ['Can we use K3 6AXIS AM in place of SEMI AM for CM press?'],
            'suggested_checks': [
                {'hint_id': 'hint_1', 'check_item': 'Verify Decap bond YK >=80% on K3 6AXIS AM in larger sample', 'reason': 'Both machines 0 NG out of 8; small Q ty limits confidence.', 'evidence_strength': 'weak', 'related_process': 'CM Press', 'related_part': 'YK', 'sheet_name': '18.6', 'source_file': name, 'source_cells': ['18.6!Decap']},
                {'hint_id': 'hint_2', 'check_item': 'Tension AVG comparison', 'reason': 'Tension K3 AVG 78.98 vs SEMI 73.69 kgf, both >= internal spec 50 kgf, K3 slightly higher.', 'evidence_strength': 'medium', 'related_process': 'CM Press Tension', 'related_part': 'YK+CMG/CPT', 'sheet_name': '18.6', 'source_file': name, 'source_cells': ['18.6!Tension table']},
            ],
            'limitations': ['Sample size 8 pcs per machine; IV. Decision text empty in source.'],
        },
        'ai_extraction_log': {'confidence': 0.7, 'assumptions': ['IV. Decision empty -> derived from data tables.'],
                              'warnings': ['Sample size only 8 pcs per machine; statistical confidence low.'],
                              'decision_rationale': 'normal_comparison: K3 6AXIS AM compared against SEMI AM same-event same-day; no NG differences observed.'}
    }
    tr_ko = {
        'document': {'title': 'BRS-161016 CM 프레스 SEMI AM → K3 6AXIS AM 변경 검증 리포트', 'purpose': 'CM 프레스를 SEMI AM에서 K3 6AXIS AM으로 변경 가능한지 확인.',
                     'content': ['Sub에서 semi 제작 후 Decap bond YK+MG/PT, Drop test, Tension test.']},
        'conclusions': {'concl_1': {'topic': 'K3 6AXIS AM 수용 여부', 'statement_from_report': '원본 IV. Decision 빈 칸.', 'normalized_interpretation': 'K3는 SEMI와 Decap NG Bond Spread(둘 다 0/8), Drop test(둘 다 0/8) 동등. Tension AVG K3 78.98 kgf vs SEMI 73.69 kgf, 둘 다 OK(내부 spec >=50 kgf, 외부 spec >=80 kgf). 표본 각 8pcs.'}},
        'hints': {
            'hint_1': {'check_item': 'K3 6AXIS AM Decap bond YK >=80% 더 큰 표본으로 확인', 'reason': '두 장비 모두 0/8 NG이나 Q ty 작아 신뢰도 부족.'},
            'hint_2': {'check_item': 'Tension AVG 비교', 'reason': 'K3 AVG 78.98 vs SEMI 73.69 kgf, 둘 다 내부 spec 50 kgf 이상이며 K3가 약간 높음.'},
        },
        'log': {'assumptions': ['IV. Decision 빈 칸 → 데이터 표로부터 결론 도출.'],
                'warnings': ['표본 8pcs로 통계적 신뢰도 낮음.'],
                'decision_rationale': 'normal_comparison: K3와 SEMI를 동일 이벤트 동일 날짜로 비교, NG 차이 없음.'}
    }
    tr_en = {
        'document': {'title': result['document']['title'], 'purpose': result['document']['purpose'], 'content': result['document']['content']},
        'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in result['conclusions']},
        'hints': {h['hint_id']: {'check_item': h['check_item'], 'reason': h['reason']} for h in result['troubleshooting_index']['suggested_checks']},
        'log': {'assumptions': result['ai_extraction_log']['assumptions'], 'warnings': result['ai_extraction_log']['warnings'], 'decision_rationale': result['ai_extraction_log']['decision_rationale']},
    }
    tr_vi = {
        'document': {'title': 'BRS-161016 Report test process CM press đổi từ SEMI AM sang K3 6AXIS AM', 'purpose': 'Kiểm tra CM press đổi từ SEMI AM sang K3 6AXIS AM có dùng được hay không.',
                     'content': ['Làm semi tại Sub: Decap bond YK+MG/PT, Drop test, Tension test.']},
        'conclusions': {'concl_1': {'topic': 'Chấp nhận K3 6AXIS AM', 'statement_from_report': 'IV. Decision trống trong source.', 'normalized_interpretation': 'K3 6AXIS AM tương đương SEMI AM về Decap NG Bond Spread (cả hai 0/8) và Drop test (cả hai 0/8). Tension AVG K3 78.98 kgf vs SEMI 73.69 kgf, đều OK (internal spec >=50 kgf). Cỡ mẫu 8 pcs mỗi máy.'}},
        'hints': {
            'hint_1': {'check_item': 'Verify Decap bond YK >=80% trên K3 6AXIS AM với mẫu lớn hơn', 'reason': 'Cả hai máy 0 NG trên 8 pcs; mẫu nhỏ giới hạn độ tin cậy.'},
            'hint_2': {'check_item': 'So sánh Tension AVG', 'reason': 'Tension K3 AVG 78.98 vs SEMI 73.69 kgf, đều đạt internal spec 50 kgf, K3 cao hơn một chút.'},
        },
        'log': {'assumptions': ['IV. Decision trống -> lấy từ bảng số liệu.'],
                'warnings': ['Mẫu 8 pcs mỗi máy, độ tin cậy thống kê thấp.'],
                'decision_rationale': 'normal_comparison: K3 6AXIS AM so sánh với SEMI AM cùng event cùng ngày; không có khác biệt NG.'}
    }
    return name, result, tr_ko, tr_en, tr_vi


def ds4():
    """3. BRS-2015 Report test material Frame using PT Dimension 0.20 - 0.22 date 28.8.2024."""
    name = '3. BRS-2015 Report test material Frame using PT Dimension 0.20 - 0.22 date 28.8.2024'
    result = {
        'schema_version': '0.1',
        'document': {
            'document_id': '', 'source_file': name, 'source_sheet': '29.8,29.8 Hearing',
            'title': 'REPORT TEST FRAME USING PT DIMENSION 0.20 ~ 0.22 OF MODEL BRS-201506',
            'model': 'BRS-201506', 'report_date': '2024-08-28', 'department': 'ME', 'marker': 'Thao',
            'line': '',
            'report_type': 'mixed',
            'primary_defect': {'canonical_name': 'NG Hearing', 'aliases_in_document': ['NG Hearing Noise', 'NG Hearing Touch', 'NG Function']},
            'related_defects': ['NG Bonding', 'NG SP Gap', 'NG Hearing Noise', 'NG Hearing Touch', 'Particle', 'Glue Clot', 'Don\'t know reason'],
            'parts': ['Frame', 'Suspension', 'Coil'],
            'processes': ['Frame Bonding', 'Frame+Suspension Vision', 'Function check', 'Final dimension', 'Modul line'],
            'purpose': 'Test material Frame using PT dimension 0.20 ~ 0.22 OK or Not? (spec 0.18)',
            'content': ["Test material Frame using PT dimension 0.20 ~ 0.22 (spec 0.18)", "Check process Frame+Sus ass'y at SUB3", "Make sample and check function", "Check dimension final", "Q'ty 299 pcs", "Compare and decide on modul test"],
            'source_cells': {'title': ['29.8!B1'], 'date': ['29.8'], 'purpose': ['29.8!I. Purpose'], 'content': ['29.8!II. Content']},
        },
        'test_conditions': [
            {'condition_id': 'cond_1', 'condition_group': 'Frame material thickness change', 'process': "Frame Bonding/Frame+Suspension/Function", 'changed_factor': 'Frame PT dimension', 'before_value': '0.18 (spec)', 'after_value': '0.20~0.22', 'unit': 'mm', 'sheet_name': '29.8', 'source_file': name, 'source_cells': ['29.8!II. Content']},
        ],
        'results': [
            {'result_id': 'r1', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'Frame Bonding Test', 'date': '2024-08-27', 'input_count': 299, 'ok_count': 299, 'ng_count': 0, 'ng_rate_percent': 0.0, 'metric_name': 'Frame Bonding NG Rate', 'sheet_name': '29.8', 'source_file': name, 'source_cells': ['29.8!Frame Bonding Test']},
            {'result_id': 'r2', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'Frame Bonding Normal', 'date': '2024-08-27', 'input_count': 500, 'ok_count': 499, 'ng_count': 1, 'ng_rate_percent': 0.2, 'metric_name': 'Frame Bonding NG Rate', 'sheet_name': '29.8', 'source_file': name, 'source_cells': ['29.8!Frame Bonding Normal']},
            {'result_id': 'r3', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'Frame+Suspension Vision Test', 'date': '2024-08-27', 'input_count': 299, 'ok_count': 297, 'ng_count': 2, 'ng_rate_percent': 0.7, 'metric_name': 'SP GAP NG Rate', 'sheet_name': '29.8', 'source_file': name, 'ng_breakdown': {'SP GAP': 2}, 'source_cells': ['29.8!Frame+SP Test']},
            {'result_id': 'r4', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'Frame+Suspension Vision Normal', 'date': '2024-08-27', 'input_count': 500, 'ok_count': 498, 'ng_count': 2, 'ng_rate_percent': 0.4, 'metric_name': 'SP GAP NG Rate', 'sheet_name': '29.8', 'source_file': name, 'ng_breakdown': {'SP GAP': 2}, 'source_cells': ['29.8!Frame+SP Normal']},
            {'result_id': 'r5', 'condition_id': 'cond_1', 'measurement_type': 'Function', 'condition_group': 'Function Frame Test', 'date': '2024-08-29', 'input_count': 293, 'ok_count': 274, 'ng_count': 19, 'ng_rate_percent': 6.5, 'metric_name': 'Function Total NG Rate Test', 'sheet_name': '29.8', 'source_file': name, 'ng_breakdown': {'SPL': 0, 'THD': 0, 'SPL+THD': 0, 'SPL+THD+F0': 0, 'Hearing Noise': 3, 'Hearing Touch': 16}, 'source_cells': ['29.8!Function Test']},
            {'result_id': 'r6', 'condition_id': 'cond_1', 'measurement_type': 'Function', 'condition_group': 'Function Frame Normal', 'date': '2024-08-29', 'input_count': 1017, 'ok_count': 966, 'ng_count': 51, 'ng_rate_percent': 5.0, 'metric_name': 'Function Total NG Rate Normal', 'sheet_name': '29.8', 'source_file': name, 'ng_breakdown': {'SPL': 0, 'THD': 1, 'SPL+THD': 0, 'SPL+THD+F0': 0, 'Hearing Noise': 11, 'Hearing Touch': 39}, 'source_cells': ['29.8!Function Normal']},
            {'result_id': 'r7', 'condition_id': 'cond_1', 'measurement_type': 'Dimension', 'condition_group': 'Frame Test dimension', 'date': '2024-08-29', 'input_count': 20, 'ok_count': 20, 'ng_count': 0, 'ng_rate_percent': 0.0, 'metric_name': 'Frame final dimension AVG', 'metric_value': 1.98, 'unit': 'mm', 'judgement': 'PASS', 'sheet_name': '29.8', 'source_file': name, 'source_cells': ['29.8!Dimension table Frame Test']},
            {'result_id': 'r8', 'condition_id': 'cond_1', 'measurement_type': 'Dimension', 'condition_group': 'Frame Normal dimension', 'date': '2024-08-29', 'input_count': 20, 'metric_name': 'Frame final dimension AVG', 'metric_value': 1.98, 'unit': 'mm', 'judgement': 'PASS', 'sheet_name': '29.8', 'source_file': name, 'source_cells': ['29.8!Dimension table Frame Normal']},
            # Hearing analysis
            {'result_id': 'r9', 'condition_id': 'cond_1', 'measurement_type': 'NG Analysis', 'condition_group': 'NG Hearing analysis 10pcs', 'date': '2024-08-29', 'input_count': 10, 'metric_name': 'NG Hearing breakdown analysis (10pcs)', 'sheet_name': '29.8 Hearing', 'source_file': name, 'ng_breakdown': {'Particle': 1, 'Glue clot': 2, "Don't know reason": 7}, 'source_cells': ['29.8 Hearing!Synthetic table']},
        ],
        'conclusions': [
            {'conclusion_id': 'concl_1', 'topic': 'Frame PT 0.20~0.22 acceptance', 'statement_from_report': 'Result checking Function NG rate 6.5% Higher NG rate normal rate 3.8% => Can not use. Result test modul line NG rate 0.73% same NG rate normal 0.75%.',
             'normalized_interpretation': 'Frame Bonding Test 0% vs Normal 0.2%; SP Gap Test 0.7% vs Normal 0.4% = 1.75x (75.0% worse). Function Test 6.5% vs Normal 5.0% = 1.30x (30.0% worse than same-event normal). Dimension AVG identical 1.98 mm. Hearing NG dominant in both Test and Normal (Touch share 84.2% Test vs 76.5% Normal). NG analysis (10pcs) shows 70% Don\'t know reason.',
             'sheet_name': '29.8', 'source_file': name, 'source_cells': ['29.8!4. Decision']},
        ],
        'troubleshooting_index': {
            'defect_name': 'NG Hearing Touch / Noise (Function)',
            'when_user_asks': ['Can Frame PT 0.20~0.22 replace spec 0.18?', 'Why Function NG rate higher when using Frame PT 0.20~0.22?'],
            'suggested_checks': [
                {'hint_id': 'hint_1', 'check_item': 'Function Hearing Touch dominance', 'reason': 'Function Test 6.5% vs Normal 5.0% = 1.30x (30.0% worse). NG mostly Hearing Touch (84.2% of 19 NG in Test).', 'evidence_strength': 'strong', 'related_process': 'Function check', 'related_part': 'Frame', 'sheet_name': '29.8', 'source_file': name, 'source_cells': ['29.8!Function table']},
                {'hint_id': 'hint_2', 'check_item': "NG Hearing root cause analysis", 'reason': "NG Hearing breakdown 10pcs: Don't know reason 70%, Glue clot 20%, Particle 10%; root cause unclear.", 'evidence_strength': 'medium', 'related_process': 'Function NG analysis', 'related_part': 'Module', 'sheet_name': '29.8 Hearing', 'source_file': name, 'source_cells': ['29.8 Hearing!Synthetic analyse']},
                {'hint_id': 'hint_3', 'check_item': 'SP Gap on Frame+Suspension', 'reason': 'SP Gap Test 0.7% vs Normal 0.4% = 1.75x (75.0% worse than same-event normal); check Frame+Suspension vision step.', 'evidence_strength': 'medium', 'related_process': 'Frame+Suspension Vision', 'related_part': 'Suspension', 'sheet_name': '29.8', 'source_file': name, 'source_cells': ['29.8!Frame+SP table']},
            ],
            'limitations': ['NG analysis sample is only 10 pcs; 70% "Don\'t know reason" leaves cause uncertain.'],
        },
        'ai_extraction_log': {'confidence': 0.78, 'assumptions': ["Normal rate cited as 3.8% in source decision text differs from table-derived 5.0%; both retained but interpretation uses table."],
                              'warnings': ['Most-frequent Hearing NG analysis bucket is "Don\'t know reason" 70%.'],
                              'decision_rationale': 'normal_comparison applied to Function/Frame+SP/Frame Bonding. Dimension/final dimension treated as before_after_dimension PASS. Hearing breakdown stored for traceability.'}
    }
    tr_ko = {
        'document': {'title': 'BRS-201506 Frame PT 두께 0.20~0.22 (spec 0.18) 사용 검증 리포트', 'purpose': 'Frame 자재 PT 두께 0.20~0.22 사용 가능 여부 확인 (spec 0.18).',
                     'content': ['SUB3에서 Frame+SP ass\'y, Function, Final dimension, Modul line 비교. Q ty 299.']},
        'conclusions': {'concl_1': {'topic': 'Frame PT 0.20~0.22 채택 여부', 'statement_from_report': 'Function NG 6.5% > Normal 3.8% → 사용 불가. Modul line NG 0.73% ~ Normal 0.75% 동등.',
                                      'normalized_interpretation': 'Frame Bonding Test 0% vs Normal 0.2%. SP Gap Test 0.7% vs Normal 0.4% = 1.75배(75.0% 악화). Function Test 6.5% vs Normal 5.0% = 1.30배(30.0% 악화). Final dimension AVG 동일 1.98 mm. NG는 Hearing Touch 우세 (Touch 84.2% Test vs 76.5% Normal). NG 분석 10pcs는 70% Don\'t know reason.'}},
        'hints': {
            'hint_1': {'check_item': 'Function Hearing Touch 우세 점검', 'reason': 'Function Test 6.5% vs Normal 5.0% = 1.30배(30.0% 악화). Test의 19 NG 중 84.2%가 Hearing Touch.'},
            'hint_2': {'check_item': 'NG Hearing 근본원인 분석 표본 확대', 'reason': 'NG Hearing 분석 10pcs 중 Don\'t know reason 70%, Glue clot 20%, Particle 10%. 원인 불명.'},
            'hint_3': {'check_item': 'Frame+Suspension SP Gap', 'reason': 'SP Gap Test 0.7% vs Normal 0.4% = 1.75배(75.0% 악화); Vision 공정 점검.'},
        },
        'log': {'assumptions': ['보고서 Decision 텍스트의 Normal 3.8%와 표 도출치 5.0% 불일치; 표 기준 해석.'],
                'warnings': ['Hearing NG 분석 표본 10pcs, "Don\'t know reason" 70%.'],
                'decision_rationale': 'normal_comparison 적용. Dimension은 before_after_dimension PASS, Hearing breakdown 별도 저장.'}
    }
    tr_en = {
        'document': {'title': result['document']['title'], 'purpose': result['document']['purpose'], 'content': result['document']['content']},
        'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in result['conclusions']},
        'hints': {h['hint_id']: {'check_item': h['check_item'], 'reason': h['reason']} for h in result['troubleshooting_index']['suggested_checks']},
        'log': {'assumptions': result['ai_extraction_log']['assumptions'], 'warnings': result['ai_extraction_log']['warnings'], 'decision_rationale': result['ai_extraction_log']['decision_rationale']},
    }
    tr_vi = {
        'document': {'title': 'BRS-201506 Report test material Frame PT dimension 0.20~0.22 (spec 0.18)', 'purpose': 'Kiểm tra Frame PT dimension 0.20~0.22 có dùng được hay không (spec 0.18).',
                     'content': ["Tại SUB3 làm Frame+SP ass'y, Function, Final dimension, Modul line. Q'ty 299."]},
        'conclusions': {'concl_1': {'topic': 'Chấp nhận Frame PT 0.20~0.22', 'statement_from_report': 'Function NG 6.5% > Normal 3.8% => Không dùng được. Modul line NG 0.73% ~ Normal 0.75%.',
                                      'normalized_interpretation': 'Frame Bonding Test 0% vs Normal 0.2%. SP Gap Test 0.7% vs Normal 0.4% = 1.75x (xấu hơn 75.0%). Function Test 6.5% vs Normal 5.0% = 1.30x (xấu hơn 30.0% so với normal cùng event). Final dimension AVG cùng 1.98 mm. Hearing Touch chiếm chủ đạo (Touch 84.2% Test vs 76.5% Normal). Phân tích NG 10pcs có 70% Don\'t know reason.'}},
        'hints': {
            'hint_1': {'check_item': 'Kiểm tra Hearing Touch chủ đạo trên Function', 'reason': 'Function Test 6.5% vs Normal 5.0% = 1.30x (xấu hơn 30.0%); 84.2% NG là Hearing Touch.'},
            'hint_2': {'check_item': 'Phân tích nguyên nhân NG Hearing với mẫu lớn hơn', 'reason': 'Phân tích 10pcs có Don\'t know reason 70%, Glue clot 20%, Particle 10%, chưa rõ nguyên nhân.'},
            'hint_3': {'check_item': 'SP Gap trên Frame+Suspension', 'reason': 'SP Gap Test 0.7% vs Normal 0.4% = 1.75x (xấu hơn 75.0%); kiểm tra vision step.'},
        },
        'log': {'assumptions': ['Decision text báo Normal 3.8% nhưng bảng tính ra 5.0%; lấy theo bảng.'],
                'warnings': ['Mẫu phân tích Hearing chỉ 10pcs, "Don\'t know reason" 70%.'],
                'decision_rationale': 'normal_comparison cho Function/Frame+SP/Frame Bonding. Dimension/final là before_after_dimension PASS. Hearing breakdown lưu để truy vết.'}
    }
    return name, result, tr_ko, tr_en, tr_vi


def ds5():
    """3. BRS-201506 DT Report checking and test porblem Frame + VP NG damage Date 14.12.2024."""
    name = '3. BRS-201506 DT Report checking and test porblem Frame + VP NG damage Date 14.12.2024'
    result = {
        'schema_version': '0.1',
        'document': {
            'document_id': '', 'source_file': name, 'source_sheet': '15.8',
            'title': 'REPORT CHECKING AND TEST PROBLEM FRAME + VP NG DAMAGE MODEL BRS-201506',
            'model': 'BRS-201506', 'report_date': '2024-12-14', 'department': 'ME', 'marker': 'Le', 'line': '',
            'report_type': 'normal_comparison',
            'primary_defect': {'canonical_name': 'NG VP Damage', 'aliases_in_document': ['VP damage', 'VP deform', 'VP stick to JIG']},
            'related_defects': ['VP damage', 'VP deform', 'VP stick to JIG'],
            'parts': ['Frame', 'VP'], 'processes': ['Frame + VP drying', 'Guide JIG'],
            'purpose': 'Find reason for NG Frame + VP damage and test improvements.',
            'content': ['Check and test Frame Press JIG.', 'Test inserting tape (0.13mm, 0.19mm) in guide Frame press JIG.', 'Test using JIG VP Array assembling Frame Array A.'],
            'source_cells': {'title': ['15.8!B1'], 'date': ['15.8'], 'purpose': ['15.8!1. Purpose'], 'content': ['15.8!2. Content']},
        },
        'test_conditions': [
            {'condition_id': 'cond_1', 'condition_group': 'Guide JIG with insert tape', 'process': 'Frame+VP drying Guide JIG', 'changed_factor': 'Insert tape thickness in guide JIG', 'before_value': 'Normal JIG', 'after_value': 'Tape 0.13/0.19 mm; JIG VP Array', 'sheet_name': '15.8', 'source_file': name, 'source_cells': ['15.8!Content']},
        ],
        'results': [
            {'result_id': 'r1', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'Guide JIG insert tape 0.13mm', 'date': '2024-12-14', 'input_count': 428, 'ok_count': 413, 'ng_count': 15, 'ng_rate_percent': 3.5, 'metric_name': 'NG Rate Frame+VP drying', 'sheet_name': '15.8', 'source_file': name, 'ng_breakdown': {'VP damage / deform': 8, 'VP stick to JIG': 7}, 'source_cells': ['15.8!Tape 0.13 row']},
            {'result_id': 'r2', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'Guide JIG insert tape 0.19mm', 'date': '2024-12-14', 'input_count': 200, 'ok_count': 198, 'ng_count': 2, 'ng_rate_percent': 1.0, 'metric_name': 'NG Rate Frame+VP drying', 'sheet_name': '15.8', 'source_file': name, 'ng_breakdown': {'VP damage / deform': 2, 'VP stick to JIG': 0}, 'source_cells': ['15.8!Tape 0.19 row']},
            {'result_id': 'r3', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': "Using JIG VP Array ass'y Frame Array A", 'date': '2024-12-14', 'input_count': 500, 'ok_count': 500, 'ng_count': 0, 'ng_rate_percent': 0.0, 'metric_name': 'NG Rate Frame+VP drying', 'sheet_name': '15.8', 'source_file': name, 'source_cells': ['15.8!JIG VP Array row']},
            {'result_id': 'r4', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'Using JIG Normal', 'date': '2024-12-14', 'input_count': 2000, 'ok_count': 1861, 'ng_count': 139, 'ng_rate_percent': 7.0, 'metric_name': 'NG Rate Frame+VP drying', 'sheet_name': '15.8', 'source_file': name, 'ng_breakdown': {'VP damage / deform': 125, 'VP stick to JIG': 14}, 'source_cells': ['15.8!JIG Normal row']},
        ],
        'conclusions': [
            {'conclusion_id': 'concl_1', 'topic': 'Frame+VP drying JIG variants', 'statement_from_report': 'IV. Decision row blank in source.',
             'normalized_interpretation': 'Same-event baseline Normal JIG NG 7.0%. Tape 0.13mm 3.5% = 0.50x (50.0% improved). Tape 0.19mm 1.0% = 0.14x (85.7% improved). JIG VP Array 0.0% = 0.0x (100% improved). VP damage/deform is the dominant defect across all rows.',
             'sheet_name': '15.8', 'source_file': name, 'source_cells': ['15.8!Result table']},
        ],
        'troubleshooting_index': {
            'defect_name': 'NG VP damage / VP stick to JIG',
            'when_user_asks': ['How to reduce Frame+VP damage?', 'Which guide JIG configuration is best?'],
            'suggested_checks': [
                {'hint_id': 'hint_1', 'check_item': "Adopt JIG VP Array for Frame Array A", 'reason': 'JIG VP Array gave 0/500 NG vs Normal JIG 139/2000 (7.0%) - 100% improvement same event.', 'evidence_strength': 'strong', 'related_process': 'Frame+VP drying', 'related_part': 'JIG / Frame Array', 'sheet_name': '15.8', 'source_file': name, 'source_cells': ['15.8!JIG VP Array row']},
                {'hint_id': 'hint_2', 'check_item': 'Increase tape thickness in guide JIG (0.13 -> 0.19 mm)', 'reason': 'Increasing tape thickness from 0.13 to 0.19 mm dropped NG from 3.5% to 1.0% = 0.29x (71.4% improvement).', 'evidence_strength': 'strong', 'related_process': 'Guide Frame press JIG', 'related_part': 'JIG tape', 'sheet_name': '15.8', 'source_file': name, 'source_cells': ['15.8!Tape rows']},
            ],
            'limitations': ['IV. Decision text empty; recommendations derived from NG counts.'],
        },
        'ai_extraction_log': {'confidence': 0.82, 'assumptions': ['IV. Decision empty -> derive from numeric NG counts.'],
                              'warnings': ['DIV/0 errors appear in source for empty breakdown cells; treated as 0.'],
                              'decision_rationale': 'normal_comparison: Normal JIG same-event baseline 7.0% NG. Compared each variant via (test/baseline-1)*100.'}
    }
    tr_ko = {
        'document': {'title': 'BRS-201506 Frame+VP NG damage 분석 및 개선 테스트 리포트', 'purpose': 'Frame+VP damage NG 원인 분석 및 개선 테스트.',
                     'content': ['Frame Press JIG 점검 및 테스트.', '가이드 Frame Press JIG에 Tape 0.13mm / 0.19mm 삽입.', 'JIG VP Array ass\'y Frame Array A 테스트.']},
        'conclusions': {'concl_1': {'topic': 'Frame+VP 건조 JIG 비교', 'statement_from_report': '원본 IV. Decision 빈 칸.', 'normalized_interpretation': '동일 이벤트 Normal JIG NG 7.0% 대조군. Tape 0.13mm 3.5% = 0.50배(50.0% 개선). Tape 0.19mm 1.0% = 0.14배(85.7% 개선). JIG VP Array 0.0% = 0.0배(100% 개선). 전 조건에서 VP damage/deform이 우세 결함.'}},
        'hints': {
            'hint_1': {'check_item': 'JIG VP Array를 Frame Array A에 적용', 'reason': 'JIG VP Array 0/500 NG, Normal JIG 139/2000(7.0%) 대비 100% 개선.'},
            'hint_2': {'check_item': '가이드 JIG Tape 두께 0.13 -> 0.19 mm로 증가', 'reason': 'Tape 0.13->0.19 mm로 NG 3.5%->1.0% = 0.29배(71.4% 개선).'},
        },
        'log': {'assumptions': ['IV. Decision 빈 칸 → 표의 NG 수치로 추론.'],
                'warnings': ['원본 breakdown 일부에 DIV/0 표시; 0으로 처리.'],
                'decision_rationale': 'normal_comparison: Normal JIG 7.0% 대조군 기반 (test/baseline-1)*100.'}
    }
    tr_en = {
        'document': {'title': result['document']['title'], 'purpose': result['document']['purpose'], 'content': result['document']['content']},
        'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in result['conclusions']},
        'hints': {h['hint_id']: {'check_item': h['check_item'], 'reason': h['reason']} for h in result['troubleshooting_index']['suggested_checks']},
        'log': {'assumptions': result['ai_extraction_log']['assumptions'], 'warnings': result['ai_extraction_log']['warnings'], 'decision_rationale': result['ai_extraction_log']['decision_rationale']},
    }
    tr_vi = {
        'document': {'title': 'BRS-201506 Report kiểm tra và test problem Frame+VP NG damage', 'purpose': 'Tìm nguyên nhân NG Frame+VP damage và test cải thiện.',
                     'content': ['Kiểm tra và test Frame Press JIG.', 'Test insert tape (0.13mm, 0.19mm) trong guide Frame press JIG.', 'Dùng JIG VP Array ass\'y Frame Array A.']},
        'conclusions': {'concl_1': {'topic': 'So sánh các phương án JIG dry Frame+VP', 'statement_from_report': 'IV. Decision trống.', 'normalized_interpretation': 'Cùng event Normal JIG NG 7.0% là baseline. Tape 0.13mm 3.5% = 0.50x (cải thiện 50.0%). Tape 0.19mm 1.0% = 0.14x (cải thiện 85.7%). JIG VP Array 0.0% = 0.0x (cải thiện 100%). VP damage/deform là defect chủ đạo ở mọi phương án.'}},
        'hints': {
            'hint_1': {'check_item': 'Áp dụng JIG VP Array cho Frame Array A', 'reason': 'JIG VP Array 0/500 NG so với Normal JIG 139/2000 (7.0%) – cải thiện 100% cùng event.'},
            'hint_2': {'check_item': 'Tăng độ dày tape trong guide JIG (0.13 -> 0.19 mm)', 'reason': 'Tape 0.13->0.19 mm giảm NG 3.5%->1.0% = 0.29x (cải thiện 71.4%).'},
        },
        'log': {'assumptions': ['IV. Decision trống -> suy từ số NG.'],
                'warnings': ['Source có cell DIV/0 ở breakdown; xử lý như 0.'],
                'decision_rationale': 'normal_comparison: Normal JIG 7.0% baseline cùng event, dùng (test/baseline-1)*100.'}
    }
    return name, result, tr_ko, tr_en, tr_vi


def ds6():
    """3. BRS-201506 Report checking and improve problem NG VP vision date 4.1.2024."""
    name = '3. BRS-201506 Report checking and improve problem NG VP vision date 4.1.2024'
    result = {
        'schema_version': '0.1',
        'document': {
            'document_id': '', 'source_file': name, 'source_sheet': 'Report (2),Sheet1',
            'title': 'REPORT CHECKING AND IMPROVE PROBLEM NG VP VISION OF MODEL BRS-201506',
            'model': 'BRS-201506', 'report_date': '2024-01-04', 'department': 'ME', 'marker': 'Thao', 'line': '',
            'report_type': 'normal_comparison',
            'primary_defect': {'canonical_name': 'NG VP Vision', 'aliases_in_document': ['Short VP separate', 'Short VP damage', 'Short VP not enough glue', 'Long VP separate', 'Long VP damage', 'Long VP not enough glue']},
            'related_defects': ['VP Separate', 'VP Damage', 'Not enough glue', 'NG Hearing Noise', 'NG Hearing Touch'],
            'parts': ['VP', 'Coil'], 'processes': ["VP ass'y Press JIG", 'AWF winding'],
            'purpose': 'Improve high NG VP vision rate.',
            'content': ['Test changing bonding line outside to 0.03.', "Test new VP ass'y press JIG (cutting guide).", "Make sample and check function. AWF winding JIG size set for AWF#1/#2/#3/#4."],
            'source_cells': {'title': ['Report (2)!B1'], 'date': ['Report (2)'], 'purpose': ['Report (2)!I. Purpose'], 'content': ['Report (2)!II. Content', 'Sheet1!table']},
        },
        'test_conditions': [
            {'condition_id': 'cond_1', 'condition_group': "VP ass'y Press JIG cutting guide vs Normal", 'process': "VP ass'y press", 'changed_factor': "Press JIG", 'before_value': 'Normal line', 'after_value': "VP ass'y Press JIG cutting guide", 'sheet_name': 'Report (2)', 'source_file': name, 'source_cells': ['Report (2)!II. Content']},
            {'condition_id': 'cond_2', 'condition_group': 'AWF winding JIG size', 'process': 'AWF winding', 'changed_factor': 'Winding JIG size', 'before_value': '9.42', 'after_value': '9.34 (AWF#1)', 'unit': 'mm', 'sheet_name': 'Sheet1', 'source_file': name, 'source_cells': ['Sheet1!AWF#1 row']},
            {'condition_id': 'cond_3', 'condition_group': 'AWF#3 stretching pole', 'process': 'AWF winding', 'changed_factor': 'Stretching pole', 'before_value': '5.065', 'after_value': '5.08 (AWF#3)', 'sheet_name': 'Sheet1', 'source_file': name, 'source_cells': ['Sheet1!AWF#3 row']},
        ],
        'results': [
            {'result_id': 'r1', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'VP Vision Test 1/6', 'date': '2024-01-06', 'input_count': 106, 'ok_count': 105, 'ng_count': 1, 'ng_rate_percent': 0.9, 'metric_name': 'VP Vision NG Rate (Test)', 'sheet_name': 'Report (2)', 'source_file': name, 'ng_breakdown': {'Short VP separate': 1, 'Short VP damage': 0, 'Short VP not enough glue': 0, 'Long VP separate': 0, 'Long VP damage': 0, 'Long VP not enough glue': 0}, 'source_cells': ['Report (2)!1/6 Test']},
            {'result_id': 'r2', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'VP Vision Normal 1/6', 'date': '2024-01-06', 'input_count': 477, 'ok_count': 465, 'ng_count': 12, 'ng_rate_percent': 2.5, 'metric_name': 'VP Vision NG Rate (Normal)', 'sheet_name': 'Report (2)', 'source_file': name, 'ng_breakdown': {'Short VP separate': 1, 'Short VP damage': 0, 'Short VP not enough glue': 8, 'Long VP separate': 1, 'Long VP damage': 2, 'Long VP not enough glue': 0}, 'source_cells': ['Report (2)!1/6 Normal']},
            {'result_id': 'r3', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'VP Vision Test 1/8', 'date': '2024-01-08', 'input_count': 200, 'ok_count': 200, 'ng_count': 0, 'ng_rate_percent': 0.0, 'metric_name': 'VP Vision NG Rate (Test)', 'sheet_name': 'Report (2)', 'source_file': name, 'source_cells': ['Report (2)!1/8 Test']},
            {'result_id': 'r4', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'VP Vision Normal 1/8', 'date': '2024-01-08', 'input_count': 445, 'ok_count': 433, 'ng_count': 12, 'ng_rate_percent': 2.7, 'metric_name': 'VP Vision NG Rate (Normal)', 'sheet_name': 'Report (2)', 'source_file': name, 'ng_breakdown': {'Short VP not enough glue': 9, 'Long VP separate': 2, 'Long VP damage': 1}, 'source_cells': ['Report (2)!1/8 Normal']},
            {'result_id': 'r5', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'VP Vision Test 1/10', 'date': '2024-01-10', 'input_count': 1044, 'ok_count': 1027, 'ng_count': 17, 'ng_rate_percent': 1.6, 'metric_name': 'VP Vision NG Rate (Test)', 'sheet_name': 'Report (2)', 'source_file': name, 'ng_breakdown': {'Short VP not enough glue': 9, 'Long VP separate': 3, 'Long VP damage': 3, 'Long VP not enough glue': 2}, 'source_cells': ['Report (2)!1/10 Test']},
            {'result_id': 'r6', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'VP Vision Normal 1/10', 'date': '2024-01-10', 'input_count': 1600, 'ok_count': 1576, 'ng_count': 24, 'ng_rate_percent': 1.5, 'metric_name': 'VP Vision NG Rate (Normal)', 'sheet_name': 'Report (2)', 'source_file': name, 'ng_breakdown': {'Short VP separate': 5, 'Short VP damage': 1, 'Short VP not enough glue': 12, 'Long VP separate': 3, 'Long VP damage': 2, 'Long VP not enough glue': 1}, 'source_cells': ['Report (2)!1/10 Normal']},
            {'result_id': 'r7', 'condition_id': 'cond_1', 'measurement_type': 'Function', 'condition_group': 'Function Test 1/6', 'date': '2024-01-06', 'input_count': 105, 'ok_count': 103, 'ng_count': 2, 'ng_rate_percent': 1.9, 'metric_name': 'Function Total NG Rate (Test)', 'sheet_name': 'Report (2)', 'source_file': name, 'ng_breakdown': {'Hearing Noise': 1, 'Hearing Touch': 1}, 'source_cells': ['Report (2)!Function 1/6 Test']},
            {'result_id': 'r8', 'condition_id': 'cond_1', 'measurement_type': 'Function', 'condition_group': 'Function Normal 1/6', 'date': '2024-01-06', 'input_count': 465, 'ok_count': 454, 'ng_count': 11, 'ng_rate_percent': 2.4, 'metric_name': 'Function Total NG Rate (Normal)', 'sheet_name': 'Report (2)', 'source_file': name, 'ng_breakdown': {'Hearing Noise': 6, 'Hearing Touch': 5}, 'source_cells': ['Report (2)!Function 1/6 Normal']},
            {'result_id': 'r9', 'condition_id': 'cond_1', 'measurement_type': 'Function', 'condition_group': 'Function Test 1/8', 'date': '2024-01-08', 'input_count': 200, 'ok_count': 195, 'ng_count': 5, 'ng_rate_percent': 2.5, 'metric_name': 'Function Total NG Rate (Test)', 'sheet_name': 'Report (2)', 'source_file': name, 'ng_breakdown': {'Hearing Noise': 5}, 'source_cells': ['Report (2)!Function 1/8 Test']},
            {'result_id': 'r10', 'condition_id': 'cond_1', 'measurement_type': 'Function', 'condition_group': 'Function Normal 1/8', 'date': '2024-01-08', 'input_count': 430, 'ok_count': 415, 'ng_count': 15, 'ng_rate_percent': 3.5, 'metric_name': 'Function Total NG Rate (Normal)', 'sheet_name': 'Report (2)', 'source_file': name, 'ng_breakdown': {'SPL': 1, 'Hearing Noise': 9, 'Hearing Touch': 5}, 'source_cells': ['Report (2)!Function 1/8 Normal']},
            {'result_id': 'r11', 'condition_id': 'cond_1', 'measurement_type': 'Function', 'condition_group': 'Function Test 1/10', 'date': '2024-01-10', 'input_count': 1024, 'ok_count': 988, 'ng_count': 36, 'ng_rate_percent': 3.5, 'metric_name': 'Function Total NG Rate (Test)', 'sheet_name': 'Report (2)', 'source_file': name, 'ng_breakdown': {'SPL': 3, 'THD': 1, 'Hearing Noise': 20, 'Hearing Touch': 12}, 'source_cells': ['Report (2)!Function 1/10 Test']},
            {'result_id': 'r12', 'condition_id': 'cond_1', 'measurement_type': 'Function', 'condition_group': 'Function Normal 1/10', 'date': '2024-01-10', 'input_count': 1584, 'ok_count': 1528, 'ng_count': 56, 'ng_rate_percent': 3.5, 'metric_name': 'Function Total NG Rate (Normal)', 'sheet_name': 'Report (2)', 'source_file': name, 'ng_breakdown': {'SPL': 1, 'Hearing Noise': 38, 'Hearing Touch': 17}, 'source_cells': ['Report (2)!Function 1/10 Normal']},
        ],
        'conclusions': [
            {'conclusion_id': 'concl_1', 'topic': "VP ass'y Press JIG cutting guide", 'statement_from_report': "IV. Decision empty (4. Decision shown blank in source).",
             'normalized_interpretation': "VP Vision Test vs Normal: 1/6 0.9% vs 2.5% = 0.36x (64.0% improved); 1/8 0.0% vs 2.7% = 0.0x (100% improved); 1/10 1.6% vs 1.5% = 1.07x (6.7% worse). Function Test vs Normal: 1/6 1.9% vs 2.4% = 0.79x (21% improved); 1/8 2.5% vs 3.5% = 0.71x (28.6% improved); 1/10 3.5% vs 3.5% = 1.0x (equal).",
             'sheet_name': 'Report (2)', 'source_file': name, 'source_cells': ['Report (2)!Result tables']},
            {'conclusion_id': 'concl_2', 'topic': 'AWF winding JIG variants', 'statement_from_report': 'AWF#1 winding JIG size 9.34 (TEST), AWF#3 stretching pole 5.08 (TEST), AWF#5 not yet run.',
             'normalized_interpretation': 'AWF#1 test winding JIG 9.34mm (vs Normal 9.42), AWF#3 test stretching pole 5.08 (vs 5.065). AWF#5 not yet run. No NG figures attached to AWF rows.',
             'sheet_name': 'Sheet1', 'source_file': name, 'source_cells': ['Sheet1!AWF rows']},
        ],
        'troubleshooting_index': {
            'defect_name': 'NG VP Vision (Short / Long not enough glue, separate, damage)',
            'when_user_asks': ['How to reduce NG VP Vision rate?', "Does new VP ass'y Press JIG cutting guide improve VP Vision NG?"],
            'suggested_checks': [
                {'hint_id': 'hint_1', 'check_item': "Adopt VP ass'y Press JIG cutting guide", 'reason': 'On 1/6 and 1/8 Test 0.9% and 0.0% vs Normal 2.5% and 2.7% (64% and 100% improved), but on 1/10 Test 1.6% vs Normal 1.5% (6.7% worse). Benefit reduces as Q ty grows; confirm with extended run.', 'evidence_strength': 'medium', 'related_process': "VP ass'y press", 'related_part': 'VP', 'sheet_name': 'Report (2)', 'source_file': name, 'source_cells': ['Report (2)!Vision rows']},
                {'hint_id': 'hint_2', 'check_item': 'Dominant defect: Short VP not enough glue', 'reason': 'Across Normal events, Short VP not enough glue accounts for the largest portion (8/12, 9/12, 12/24) of VP Vision NG.', 'evidence_strength': 'strong', 'related_process': 'VP bonding glue', 'related_part': 'Short VP', 'sheet_name': 'Report (2)', 'source_file': name, 'source_cells': ['Report (2)!Vision breakdown']},
                {'hint_id': 'hint_3', 'check_item': 'AWF winding JIG size and stretching pole adjustments', 'reason': 'AWF#1 JIG size changed to 9.34 (vs 9.42), AWF#3 stretching pole 5.08 (vs 5.065) marked TEST; collect downstream NG before approving.', 'evidence_strength': 'weak', 'related_process': 'AWF winding', 'related_part': 'Coil', 'sheet_name': 'Sheet1', 'source_file': name, 'source_cells': ['Sheet1!AWF rows']},
            ],
            'limitations': ['IV. Decision rows are empty; AWF table lacks NG metrics; AWF#5 not yet run.'],
        },
        'ai_extraction_log': {'confidence': 0.8, 'assumptions': ['Decision text not provided; conclusions derived from numerical tables.'],
                              'warnings': ['Function Normal 1/10 breakdown row in source shows >100% portions (253.3% Noise / 113.3% Touch); raw counts (38/17) used instead.', 'AWF table has no NG measurement.'],
                              'decision_rationale': "normal_comparison: same-event Normal line baselines for each date; (test/baseline-1)*100."}
    }
    tr_ko = {
        'document': {'title': 'BRS-201506 VP Vision NG 개선 검증 리포트', 'purpose': '높은 VP Vision NG 개선.',
                     'content': ['Bonding line outside 0.03 변경.', '신규 VP ass\'y Press JIG (cutting guide) 테스트.', 'AWF 권선 JIG 사이즈 AWF#1~#5 변경.']},
        'conclusions': {
            'concl_1': {'topic': "VP ass'y Press JIG cutting guide", 'statement_from_report': 'IV. Decision 빈 칸.', 'normalized_interpretation': 'VP Vision Test vs Normal: 1/6 0.9% vs 2.5% = 0.36배(64.0% 개선), 1/8 0.0% vs 2.7% = 0.0배(100% 개선), 1/10 1.6% vs 1.5% = 1.07배(6.7% 악화). Function Test vs Normal: 1/6 1.9% vs 2.4% = 0.79배(21% 개선), 1/8 2.5% vs 3.5% = 0.71배(28.6% 개선), 1/10 3.5% vs 3.5% = 1.0배 동일.'},
            'concl_2': {'topic': 'AWF 권선 JIG 변경', 'statement_from_report': 'AWF#1 9.34, AWF#3 5.08, AWF#5 미실행.', 'normalized_interpretation': 'AWF#1 권선 JIG 9.34(기준 9.42), AWF#3 스트레칭 폴 5.08(기준 5.065). AWF#5 미실행. AWF 표에 NG 수치 없음.'},
        },
        'hints': {
            'hint_1': {'check_item': "VP ass'y Press JIG cutting guide 채택 검토", 'reason': '1/6, 1/8은 64%, 100% 개선이나 1/10은 6.7% 악화. Q ty 증가시 효과 감소, 추가 검증 필요.'},
            'hint_2': {'check_item': 'Short VP not enough glue가 우세 결함', 'reason': 'Normal 이벤트 NG의 최대 비중이 Short VP not enough glue(8/12, 9/12, 12/24).'},
            'hint_3': {'check_item': 'AWF 권선 JIG 사이즈/스트레칭 폴 조정', 'reason': 'AWF#1 9.34, AWF#3 5.08 TEST 마킹; 다운스트림 NG 확인 후 승인.'},
        },
        'log': {'assumptions': ['Decision 빈 칸 → 표에서 결론.'],
                'warnings': ['Function Normal 1/10 source breakdown 비율이 253.3%/113.3%로 비정상; 카운트(38/17) 사용. AWF 표 NG 정보 없음.'],
                'decision_rationale': 'normal_comparison: 동일 일자 Normal line baseline 기반 (test/baseline-1)*100.'}
    }
    tr_en = {
        'document': {'title': result['document']['title'], 'purpose': result['document']['purpose'], 'content': result['document']['content']},
        'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in result['conclusions']},
        'hints': {h['hint_id']: {'check_item': h['check_item'], 'reason': h['reason']} for h in result['troubleshooting_index']['suggested_checks']},
        'log': {'assumptions': result['ai_extraction_log']['assumptions'], 'warnings': result['ai_extraction_log']['warnings'], 'decision_rationale': result['ai_extraction_log']['decision_rationale']},
    }
    tr_vi = {
        'document': {'title': 'BRS-201506 Report kiểm tra và cải thiện NG VP Vision', 'purpose': 'Cải thiện NG VP Vision cao.',
                     'content': ["Đổi bonding line outside 0.03.", "Test JIG ass'y press VP cutting guide.", 'Đổi winding JIG size AWF#1~#5.']},
        'conclusions': {
            'concl_1': {'topic': "JIG ass'y press VP cutting guide", 'statement_from_report': 'IV. Decision trống.', 'normalized_interpretation': 'VP Vision Test vs Normal: 1/6 0.9% vs 2.5% = 0.36x (cải thiện 64.0%); 1/8 0.0% vs 2.7% = 0.0x (cải thiện 100%); 1/10 1.6% vs 1.5% = 1.07x (xấu hơn 6.7%). Function Test vs Normal: 1/6 1.9% vs 2.4% = 0.79x (cải thiện 21%); 1/8 2.5% vs 3.5% = 0.71x (cải thiện 28.6%); 1/10 3.5% vs 3.5% = 1.0x (bằng).'},
            'concl_2': {'topic': 'AWF winding JIG', 'statement_from_report': 'AWF#1 9.34, AWF#3 5.08, AWF#5 chưa chạy.', 'normalized_interpretation': 'AWF#1 winding JIG 9.34 (vs 9.42), AWF#3 stretching pole 5.08 (vs 5.065). AWF#5 chưa chạy. Bảng AWF không có NG figures.'},
        },
        'hints': {
            'hint_1': {'check_item': "Áp dụng JIG ass'y press VP cutting guide", 'reason': '1/6, 1/8 cải thiện 64%, 100% nhưng 1/10 xấu hơn 6.7%; benefit giảm khi Q ty tăng, cần verify thêm.'},
            'hint_2': {'check_item': 'Defect chủ đạo: Short VP not enough glue', 'reason': 'Trên các event Normal, Short VP not enough glue chiếm phần lớn (8/12, 9/12, 12/24).'},
            'hint_3': {'check_item': 'Điều chỉnh winding JIG size và stretching pole', 'reason': 'AWF#1 9.34, AWF#3 5.08 đang TEST; thu thập NG downstream trước khi approve.'},
        },
        'log': {'assumptions': ['IV. Decision trống → suy từ bảng số.'],
                'warnings': ['Function Normal 1/10 breakdown trong source có % bất thường (253.3%/113.3%); dùng số đếm 38/17. Bảng AWF không có NG.'],
                'decision_rationale': "normal_comparison: dùng Normal line cùng ngày làm baseline, (test/baseline-1)*100."}
    }
    return name, result, tr_ko, tr_en, tr_vi


def _c11_common(name, src_sheet):
    """Both C11-20-R datasets are reliability/SPL/IMP/THD measurement curves."""
    result = {
        'schema_version': '0.1',
        'document': {
            'document_id': '', 'source_file': name, 'source_sheet': src_sheet,
            'title': 'C11-20-R SPL/IMP/THD measurement for ME (Đo cho ME 12.1.2026)',
            'model': 'C11-20-R', 'report_date': '2026-01-12', 'department': 'ME', 'marker': '', 'line': '',
            'report_type': 'reliability_spec',
            'primary_defect': {'canonical_name': 'SPL/IMP/THD measurement (no NG)', 'aliases_in_document': []},
            'related_defects': [], 'parts': ['C11-20-R speaker'], 'processes': ['Acoustic measurement'],
            'purpose': 'Compare SPL, IMP and THD curves between 10 Test (Hàng test) units and 10 Normal (Hàng thường) units across frequency 20 Hz ~ 20 kHz.',
            'content': ['Sheet SPL: frequency vs SPL dB for 10 Test units (#1..#10) and 10 Normal units across standard frequency points 20 Hz ~ 20 kHz.',
                        'Sheet IMP: impedance ohm vs frequency for the same 10+10 units.',
                        'Sheet THD: THD percent or distortion across frequency for the same 10+10 units.'],
            'source_cells': {'title': [src_sheet+'!header'], 'date': [], 'purpose': ['SPL header'], 'content': [src_sheet]},
        },
        'test_conditions': [
            {'condition_id': 'cond_1', 'condition_group': 'Acoustic curve measurement', 'process': 'Final acoustic test', 'changed_factor': 'Unit group', 'before_value': 'Normal group (Hàng thường) 10 units', 'after_value': 'Test group (Hàng test) 10 units', 'sheet_name': src_sheet, 'source_file': name, 'source_cells': [src_sheet+'!header row']},
        ],
        'results': [
            {'result_id': 'r1', 'condition_id': 'cond_1', 'measurement_type': 'SPL', 'condition_group': 'SPL curve 20 Hz~20 kHz', 'metric_name': 'SPL dB vs frequency (10 Test + 10 Normal)', 'sheet_name': 'SPL', 'judgement': 'CHECK', 'source_file': name, 'source_cells': ['SPL!full table']},
            {'result_id': 'r2', 'condition_id': 'cond_1', 'measurement_type': 'IMP', 'condition_group': 'Impedance curve', 'metric_name': 'Impedance Ohm vs frequency (10 Test + 10 Normal)', 'sheet_name': 'IMP', 'judgement': 'CHECK', 'source_file': name, 'source_cells': ['IMP!full table']},
            {'result_id': 'r3', 'condition_id': 'cond_1', 'measurement_type': 'THD', 'condition_group': 'THD curve', 'metric_name': 'THD percent vs frequency (10 Test + 10 Normal)', 'sheet_name': 'THD', 'judgement': 'CHECK', 'source_file': name, 'source_cells': ['THD!full table']},
        ],
        'conclusions': [
            {'conclusion_id': 'concl_1', 'topic': 'C11-20-R acoustic curves Test vs Normal', 'statement_from_report': 'No textual decision; sheets contain only frequency-by-frequency SPL/IMP/THD numbers for 10 Test and 10 Normal units.',
             'normalized_interpretation': 'Numerical SPL/IMP/THD curves stored for ME review. No PASS/FAIL judgement is given in source. Test and Normal groups can be compared per frequency to confirm equivalence; no NG rate computation possible.',
             'sheet_name': src_sheet, 'source_file': name, 'source_cells': [src_sheet+'!data']},
        ],
        'troubleshooting_index': {
            'defect_name': 'C11-20-R acoustic deviation (SPL/IMP/THD)',
            'when_user_asks': ['Are C11-20-R Test units within Normal SPL/IMP/THD distribution?'],
            'suggested_checks': [
                {'hint_id': 'hint_1', 'check_item': 'Frequency-by-frequency SPL delta Test vs Normal', 'reason': 'Sheet SPL provides 10 Test + 10 Normal values for each standard frequency 20 Hz~20 kHz; mean and spread should be compared.', 'evidence_strength': 'medium', 'related_process': 'Acoustic measurement', 'related_part': 'Speaker', 'sheet_name': 'SPL', 'source_file': name, 'source_cells': ['SPL!data']},
                {'hint_id': 'hint_2', 'check_item': 'Impedance resonance peak', 'reason': 'Sheet IMP holds impedance vs frequency for 10 Test + 10 Normal; resonance frequency and peak Ohm should match Normal distribution.', 'evidence_strength': 'medium', 'related_process': 'Acoustic measurement', 'related_part': 'Speaker', 'sheet_name': 'IMP', 'source_file': name, 'source_cells': ['IMP!data']},
                {'hint_id': 'hint_3', 'check_item': 'THD percent above spec bands', 'reason': 'Sheet THD lists distortion vs frequency for 10 Test + 10 Normal; check for any Test unit exceeding Normal envelope above audible bands.', 'evidence_strength': 'medium', 'related_process': 'Acoustic measurement', 'related_part': 'Speaker', 'sheet_name': 'THD', 'source_file': name, 'source_cells': ['THD!data']},
            ],
            'limitations': ['No textual judgement; no NG/OK counts; only numerical measurement curves.'],
        },
        'ai_extraction_log': {'confidence': 0.6, 'assumptions': ['No PASS/FAIL provided; treated as reliability_spec measurement set.'],
                              'warnings': ['Large numeric tables not normalized to individual results rows; only one summary result per metric type stored. Source has only 10+10 acoustic curves, no NG rate.'],
                              'decision_rationale': 'reliability_spec: measurement curves stored per metric type; full grid preserved in RawJson.'}
    }
    tr_ko = {
        'document': {'title': 'C11-20-R 음향(SPL/IMP/THD) 측정 (ME용 측정 2026-01-12)', 'purpose': 'C11-20-R 모델의 SPL, IMP, THD를 Test 10대 vs Normal 10대로 20Hz~20kHz에서 비교 측정.',
                     'content': ['SPL 시트: 20Hz~20kHz의 dB 값을 Test 10대(#1~#10)와 Normal 10대 각각 기록.', 'IMP 시트: 동일 단위의 임피던스(Ohm).', 'THD 시트: 동일 단위의 왜곡율(%).']},
        'conclusions': {'concl_1': {'topic': 'C11-20-R 음향 곡선 Test vs Normal', 'statement_from_report': '원본에 결정/평가 텍스트 없음. SPL/IMP/THD 숫자만 존재.',
                                      'normalized_interpretation': '숫자 데이터만 있으므로 PASS/FAIL 판정 불가. ME 측 검토용으로 곡선 데이터 저장. NG 비율은 계산할 수 없음.'}},
        'hints': {
            'hint_1': {'check_item': '주파수별 SPL Test vs Normal 편차', 'reason': '20Hz~20kHz 각 점에 Test 10/Normal 10 데이터; 평균과 산포 비교 필요.'},
            'hint_2': {'check_item': '임피던스 공진 피크 비교', 'reason': 'IMP 시트의 Test/Normal 곡선 공진 주파수와 피크 Ohm가 정상 분포 안에 있는지 확인.'},
            'hint_3': {'check_item': '특정 대역 THD 초과 여부', 'reason': 'THD 시트에서 Test 유닛이 가청 대역에서 Normal envelope를 초과하는지 점검.'},
        },
        'log': {'assumptions': ['PASS/FAIL 미기재 → reliability_spec 측정 데이터로 처리.'],
                'warnings': ['대형 측정 표를 행단위 result로 분해하지 않고 metric별 1개 요약만 저장; 전체 그리드는 RawJson에 보존.'],
                'decision_rationale': 'reliability_spec: metric별 측정 곡선 요약 저장. 전체 데이터는 RawJson 보존.'}
    }
    tr_en = {
        'document': {'title': result['document']['title'], 'purpose': result['document']['purpose'], 'content': result['document']['content']},
        'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in result['conclusions']},
        'hints': {h['hint_id']: {'check_item': h['check_item'], 'reason': h['reason']} for h in result['troubleshooting_index']['suggested_checks']},
        'log': {'assumptions': result['ai_extraction_log']['assumptions'], 'warnings': result['ai_extraction_log']['warnings'], 'decision_rationale': result['ai_extraction_log']['decision_rationale']},
    }
    tr_vi = {
        'document': {'title': 'C11-20-R đo SPL/IMP/THD cho ME 12.1.2026', 'purpose': 'So sánh SPL, IMP, THD giữa 10 đơn vị Hàng test và 10 đơn vị Hàng thường ở 20Hz~20kHz cho model C11-20-R.',
                     'content': ['Sheet SPL: dB theo tần số 20Hz~20kHz cho 10 Test (#1~#10) và 10 Normal.', 'Sheet IMP: trở kháng Ohm theo tần số cùng 10+10 đơn vị.', 'Sheet THD: phần trăm méo theo tần số cùng 10+10 đơn vị.']},
        'conclusions': {'concl_1': {'topic': 'Đường cong âm thanh C11-20-R Test vs Normal', 'statement_from_report': 'Không có decision text; chỉ có số liệu SPL/IMP/THD.',
                                      'normalized_interpretation': 'Chỉ có số liệu, không có PASS/FAIL trong source. Lưu đường cong để ME review; không tính được NG rate.'}},
        'hints': {
            'hint_1': {'check_item': 'Chênh lệch SPL Test vs Normal theo tần số', 'reason': 'Mỗi điểm 20Hz~20kHz có 10 Test/10 Normal; cần so trung bình và độ phân tán.'},
            'hint_2': {'check_item': 'Đỉnh cộng hưởng impedance', 'reason': 'Sheet IMP có đường cong 10 Test + 10 Normal; cần kiểm tần số cộng hưởng và đỉnh Ohm có nằm trong dải Normal.'},
            'hint_3': {'check_item': 'THD percent vượt spec', 'reason': 'Sheet THD liệt kê méo theo tần số; kiểm tra Test có vượt envelope Normal ở dải nghe được hay không.'},
        },
        'log': {'assumptions': ['Source không có PASS/FAIL → coi là reliability_spec.'],
                'warnings': ['Bảng lớn không normalize từng dòng; chỉ lưu 1 summary mỗi metric. Toàn bộ grid được giữ ở RawJson.'],
                'decision_rationale': 'reliability_spec: lưu summary mỗi metric; raw grid trong RawJson.'}
    }
    return result, tr_ko, tr_en, tr_vi


def ds7():
    name = '3. C11-20-R (Đo cho ME 12.1.2026)'
    result, tr_ko, tr_en, tr_vi = _c11_common(name, 'SPL,IMP,THD')
    return name, result, tr_ko, tr_en, tr_vi


def ds8():
    name = '3. C11-20-R Đo cho ME 12.1.2026'
    result, tr_ko, tr_en, tr_vi = _c11_common(name, 'SPL DATA,IMP,THD')
    return name, result, tr_ko, tr_en, tr_vi


def ds9():
    """3. MSU-201507 DT Report test Check reason NG not dry glue VP-CD date 11.3.2025."""
    name = '3. MSU-201507 DT Report test Check reason NG not dry glue  VP-CD  date 11.3.2025'
    result = {
        'schema_version': '0.1',
        'document': {
            'document_id': '', 'source_file': name, 'source_sheet': 'Change method,CD press JIG reduce MG,CD press JIG insert tape,NG rate (VP deform),Picture',
            'title': 'BRS-201507DT / MSU-L2S15-07 REPORT TEST CHECK REASON NG NOT DRY GLUE VP/CD, REDUCE OVER GLUE VP/CD, INSERT TAPE TEST, VP DEFORM CHECK',
            'model': 'BRS-201507DT / MSU-L2S15-07', 'report_date': '2025-03-11', 'department': 'ME', 'marker': 'Le', 'line': '',
            'report_type': 'mixed',
            'primary_defect': {'canonical_name': 'NG Not Dry Glue VP/CD', 'aliases_in_document': ['Not dry glue', 'Over glue', 'CD offset', 'VP deform', 'VP separate', 'Dome damage', 'Over glue and CD separate']},
            'related_defects': ['Not dry glue', 'Over glue', 'VP deform', 'CD offset', 'VP separate', 'Dome damage', 'CD separate', 'NG Hearing Noise', 'NG Hearing Touch'],
            'parts': ['VP', 'CD', 'Dome'], 'processes': ['UC press VP/CD', 'CD press JIG'],
            'purpose': 'Test reasons for NG not dry glue and over glue on VP/CD process, and propose JIG/method changes.',
            'content': ['Change method UC press VP/CD vs normal dry method; check Vision/Function/Tension.', 'Reduce magnet of CD press JIG (2 -> 1) vs Normal.', 'Insert tape into CD press JIG vs Normal.', 'Check VP deform when UC press controller error.', 'Picture-only sheet for method change explanation.'],
            'source_cells': {'title': ['Change method!B1', 'CD press JIG reduce MG!B1'], 'date': ['Change method'], 'purpose': ['I. Purpose'], 'content': ['II. Content']},
        },
        'test_conditions': [
            {'condition_id': 'cond_1', 'condition_group': 'Change method UC press VP/CD', 'process': 'UC press VP/CD', 'changed_factor': 'Dry method (UV order)', 'before_value': 'Normal dry method (UV PRESS -> UV LED with PRESS JIG -> UV LED)', 'after_value': 'Change method (UV PRESS -> UV LED -> UV LED with PRESS JIG)', 'sheet_name': 'Change method', 'source_file': name, 'source_cells': ['Change method!Picture method sheet']},
            {'condition_id': 'cond_2', 'condition_group': 'CD press JIG magnet reduction', 'process': 'CD press JIG', 'changed_factor': 'Magnet count on CD press JIG', 'before_value': '2 magnets (Normal)', 'after_value': '1 magnet (Reduce magnet)', 'sheet_name': 'CD press JIG reduce MG', 'source_file': name, 'source_cells': ['CD press JIG reduce MG!Content']},
            {'condition_id': 'cond_3', 'condition_group': 'CD press JIG insert tape', 'process': 'CD press JIG', 'changed_factor': 'Insert tape', 'before_value': 'CD press JIG normal', 'after_value': 'CD press JIG insert tape', 'sheet_name': 'CD press JIG insert tape', 'source_file': name, 'source_cells': ['CD press JIG insert tape!Content']},
            {'condition_id': 'cond_4', 'condition_group': 'UC press machine controller', 'process': 'UC press VP/CD', 'changed_factor': 'UC press controller state', 'before_value': 'Controller OK', 'after_value': 'Controller error', 'sheet_name': 'NG rate  (VP deform)', 'source_file': name, 'source_cells': ['NG rate (VP deform)!Content']},
        ],
        'results': [
            # Change method - Vision per day
            {'result_id': 'r1', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'Change method UC press VP/CD', 'date': '2025-03-10', 'input_count': 800, 'ok_count': 800, 'ng_count': 0, 'ng_rate_percent': 0.0, 'metric_name': 'Vision VP/CD NG Rate', 'sheet_name': 'Change method', 'source_file': name, 'source_cells': ['Change method!3/10 row']},
            {'result_id': 'r2', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'Normal dry UC press VP/CD', 'date': '2025-03-10', 'input_count': 960, 'ok_count': 934, 'ng_count': 26, 'ng_rate_percent': 2.7, 'metric_name': 'Vision VP/CD NG Rate (Normal)', 'sheet_name': 'Change method', 'source_file': name, 'ng_breakdown': {'Not enough glue': 26, 'Over glue': 0, 'CD offset': 0, 'VP deform': 0, 'VP Separate': 0, 'Dome damage': 0}, 'source_cells': ['Change method!Normal row']},
            {'result_id': 'r3', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'Change method UC press VP/CD', 'date': '2025-03-11', 'input_count': 10108, 'ok_count': 9883, 'ng_count': 225, 'ng_rate_percent': 2.2, 'metric_name': 'Vision VP/CD NG Rate', 'sheet_name': 'Change method', 'source_file': name, 'ng_breakdown': {'Not enough glue': 17, 'Over glue': 51, 'CD offset': 57, 'VP deform': 28, 'VP Separate': 54, 'Dome damage': 18}, 'source_cells': ['Change method!3/11 row']},
            {'result_id': 'r4', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'Change method UC press VP/CD', 'date': '2025-03-12', 'input_count': 9120, 'ok_count': 9069, 'ng_count': 51, 'ng_rate_percent': 0.6, 'metric_name': 'Vision VP/CD NG Rate', 'sheet_name': 'Change method', 'source_file': name, 'ng_breakdown': {'Not enough glue': 18, 'Over glue': 22, 'CD offset': 10, 'VP deform': 0, 'VP Separate': 1, 'Dome damage': 0}, 'source_cells': ['Change method!3/12 row']},
            {'result_id': 'r5', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'Change method UC press VP/CD', 'date': '2025-03-13', 'input_count': 10145, 'ok_count': 10116, 'ng_count': 29, 'ng_rate_percent': 0.3, 'metric_name': 'Vision VP/CD NG Rate', 'sheet_name': 'Change method', 'source_file': name, 'ng_breakdown': {'Not enough glue': 11, 'Over glue': 15, 'CD offset': 3, 'VP deform': 0, 'VP Separate': 0, 'Dome damage': 0}, 'source_cells': ['Change method!3/13 row']},
            {'result_id': 'r6', 'condition_id': 'cond_1', 'measurement_type': 'Vision', 'condition_group': 'Change method total', 'date': '2025-03-10..18', 'input_count': 80928, 'ok_count': 80320, 'ng_count': 608, 'ng_rate_percent': 0.8, 'metric_name': 'Vision VP/CD NG Rate (Total Change method)', 'sheet_name': 'Change method', 'source_file': name, 'ng_breakdown': {'Not enough glue': 134, 'Over glue': 183, 'CD offset': 137, 'VP deform': 81, 'VP Separate': 55, 'Dome damage': 18}, 'source_cells': ['Change method!Total row']},
            # Function summary Change method
            {'result_id': 'r7', 'condition_id': 'cond_1', 'measurement_type': 'Function', 'condition_group': 'Change method Total Function', 'date': '2025-03-10..17', 'input_count': 67832, 'ok_count': 64369, 'metric_name': 'Function Total NG (Sigma+Hearing+1V/+0V)', 'sheet_name': 'Change method', 'source_file': name, 'ng_breakdown': {'Sigma Total': 94, 'Hearing +1V Total': 3369, 'Hearing +0V Total': 1123}, 'source_cells': ['Change method!Function Total']},
            {'result_id': 'r8', 'condition_id': 'cond_1', 'measurement_type': 'Function', 'condition_group': 'Normal dry Function', 'date': '2025-03-10', 'input_count': 896, 'ok_count': 869, 'metric_name': 'Function Total NG (Normal)', 'sheet_name': 'Change method', 'source_file': name, 'ng_breakdown': {'Sigma Total': 0, 'Hearing +1V Total': 27, 'Hearing +0V Total': 10}, 'source_cells': ['Change method!Normal Function']},
            # Tension Change method
            {'result_id': 'r9', 'condition_id': 'cond_1', 'measurement_type': 'Tension', 'condition_group': 'UC press VP/CD Normal Tension', 'date': '2025-03-05', 'metric_name': 'Tension AVG Normal', 'metric_value': 1.95, 'unit': 'kgf', 'judgement': 'NG', 'sheet_name': 'Change method', 'source_file': name, 'source_cells': ['Change method!Tension 3/5 Normal']},
            {'result_id': 'r10', 'condition_id': 'cond_1', 'measurement_type': 'Tension', 'condition_group': 'UC press VP/CD Normal Tension', 'date': '2025-03-06', 'metric_name': 'Tension AVG Normal', 'metric_value': 2.07, 'unit': 'kgf', 'judgement': 'NG', 'sheet_name': 'Change method', 'source_file': name, 'source_cells': ['Change method!Tension 3/6 Normal']},
            {'result_id': 'r11', 'condition_id': 'cond_1', 'measurement_type': 'Tension', 'condition_group': 'UC press VP/CD Normal Tension', 'date': '2025-03-07', 'metric_name': 'Tension AVG Normal', 'metric_value': 2.65, 'unit': 'kgf', 'judgement': 'OK', 'sheet_name': 'Change method', 'source_file': name, 'source_cells': ['Change method!Tension 3/7 Normal']},
            {'result_id': 'r12', 'condition_id': 'cond_1', 'measurement_type': 'Tension', 'condition_group': 'UC press VP/CD Change method TEST Tension', 'date': '2025-03-10', 'metric_name': 'Tension AVG TEST', 'metric_value': 2.32, 'unit': 'kgf', 'judgement': 'NG', 'sheet_name': 'Change method', 'source_file': name, 'source_cells': ['Change method!Tension 3/10 TEST']},
            # CD press JIG reduce magnet - vision totals
            {'result_id': 'r13', 'condition_id': 'cond_2', 'measurement_type': 'Vision', 'condition_group': 'Reduce magnet CD press JIG Total', 'date': '2025-03-13..15', 'input_count': 18731, 'ok_count': 18571, 'ng_count': 160, 'ng_rate_percent': 0.9, 'metric_name': 'Vision VP/CD NG Rate (Reduce magnet Total)', 'sheet_name': 'CD press JIG reduce MG', 'source_file': name, 'ng_breakdown': {'Not enough bond': 19, 'Over glue and CD separate': 51, 'Dome damage': 10, 'VP deform': 71, 'Particle': 9}, 'source_cells': ['CD press JIG reduce MG!Total']},
            {'result_id': 'r14', 'condition_id': 'cond_2', 'measurement_type': 'Vision', 'condition_group': 'Normal CD press JIG', 'date': '2025-03-13', 'input_count': 80, 'ok_count': 77, 'ng_count': 3, 'ng_rate_percent': 3.8, 'metric_name': 'Vision VP/CD NG Rate (Normal CD press JIG)', 'sheet_name': 'CD press JIG reduce MG', 'source_file': name, 'ng_breakdown': {'Not enough bond': 0, 'Over glue and CD separate': 3, 'Dome damage': 0, 'VP deform': 0, 'Particle': 0}, 'source_cells': ['CD press JIG reduce MG!Normal']},
            {'result_id': 'r15', 'condition_id': 'cond_2', 'measurement_type': 'Function', 'condition_group': 'Reduce magnet Function Total', 'date': '2025-03-13..15', 'input_count': 16584, 'ok_count': 15924, 'metric_name': 'Function Total NG (Reduce magnet)', 'sheet_name': 'CD press JIG reduce MG', 'source_file': name, 'ng_breakdown': {'Sigma Total': 71, 'Hearing +1V Total': 589, 'Hearing +0V Total': 187}, 'source_cells': ['CD press JIG reduce MG!Function Total']},
            {'result_id': 'r16', 'condition_id': 'cond_2', 'measurement_type': 'Function', 'condition_group': 'Normal Function', 'date': '2025-03-13', 'input_count': 77, 'ok_count': 73, 'metric_name': 'Function Total NG (Normal)', 'sheet_name': 'CD press JIG reduce MG', 'source_file': name, 'ng_breakdown': {'Sigma Total': 0, 'Hearing +1V Total': 4, 'Hearing +0V Total': 2}, 'source_cells': ['CD press JIG reduce MG!Normal Function']},
            # CD press JIG insert tape
            {'result_id': 'r17', 'condition_id': 'cond_3', 'measurement_type': 'Vision', 'condition_group': 'JIG CD Press Insert tape', 'date': '2025-03-11', 'input_count': 39, 'ok_count': 35, 'ng_count': 4, 'ng_rate_percent': 10.3, 'metric_name': 'Vision VP/CD NG Rate (Insert tape)', 'sheet_name': 'CD press JIG insert tape', 'source_file': name, 'ng_breakdown': {'Not dry glue': 0, 'VP deform': 3, 'Over glue': 1}, 'source_cells': ['CD press JIG insert tape!Insert tape row']},
            {'result_id': 'r18', 'condition_id': 'cond_3', 'measurement_type': 'Vision', 'condition_group': 'JIG CD Press normal', 'date': '2025-03-11', 'input_count': 40, 'ok_count': 39, 'ng_count': 1, 'ng_rate_percent': 2.5, 'metric_name': 'Vision VP/CD NG Rate (Normal)', 'sheet_name': 'CD press JIG insert tape', 'source_file': name, 'ng_breakdown': {'Over glue': 1}, 'source_cells': ['CD press JIG insert tape!Normal row']},
            # VP deform controller error
            {'result_id': 'r19', 'condition_id': 'cond_4', 'measurement_type': 'Vision', 'condition_group': 'UC press controller error', 'date': '2025-03-11', 'input_count': 24, 'ok_count': 0, 'ng_count': 24, 'ng_rate_percent': 100.0, 'metric_name': 'Vision VP/CD NG Rate (controller error)', 'sheet_name': 'NG rate  (VP deform)', 'source_file': name, 'ng_breakdown': {'VP deform': 24}, 'source_cells': ['NG rate (VP deform)!Error row']},
            {'result_id': 'r20', 'condition_id': 'cond_4', 'measurement_type': 'Vision', 'condition_group': 'UC press controller OK', 'date': '2025-03-11', 'input_count': 40, 'ok_count': 40, 'ng_count': 0, 'ng_rate_percent': 0.0, 'metric_name': 'Vision VP/CD NG Rate (controller OK)', 'sheet_name': 'NG rate  (VP deform)', 'source_file': name, 'source_cells': ['NG rate (VP deform)!OK row']},
        ],
        'conclusions': [
            {'conclusion_id': 'concl_1', 'topic': 'Change method UC press VP/CD', 'statement_from_report': 'Can use new method.',
             'normalized_interpretation': 'Change method Total Vision VP/CD NG 0.8% (608/80928) vs Normal dry 2.7% (26/960) = 0.30x (70.4% improved vs same-event normal). Function NG Hearing +1V Change method 5.0% vs Normal 3.0% = 1.65x (65.0% worse); Hearing +0V 1.7% vs 1.1% = 1.51x (50.9% worse). Tension test runs include 3 NG out of 5 (3/5, 3/6 and 3/10 below 1.9 kgf spec).',
             'sheet_name': 'Change method', 'source_file': name, 'source_cells': ['Change method!IV. Decision']},
            {'conclusion_id': 'concl_2', 'topic': 'CD press JIG reduce 1 magnet', 'statement_from_report': 'When use CD Press JIG reduce 1 magnet, NG over glue and CD separate reduces => Can use JIG repair 1 magnet.',
             'normalized_interpretation': 'Reduce magnet Total Vision NG 0.9% (160/18731) vs Normal CD press JIG 3.8% (3/80) = 0.24x (76.3% improved). Over glue and CD separate count drops from 3.8% to 0.3%. Function Hearing +1V 3.6% vs Normal 5.2% = 0.69x (30.8% improved); +0V 1.1% vs 2.6% = 0.42x (57.7% improved).',
             'sheet_name': 'CD press JIG reduce MG', 'source_file': name, 'source_cells': ['CD press JIG reduce MG!IV. Decision']},
            {'conclusion_id': 'concl_3', 'topic': 'CD press JIG insert tape', 'statement_from_report': 'When use JIG CD press insert tape, when open JIG happen NG VP/CD stickup on press JIG => Result NG increase more than normal => Can not use.',
             'normalized_interpretation': 'Insert tape Vision NG 10.3% (4/39) vs Normal 2.5% (1/40) = 4.12x (312% worse than same-event normal). Reject.',
             'sheet_name': 'CD press JIG insert tape', 'source_file': name, 'source_cells': ['CD press JIG insert tape!Decision text']},
            {'conclusion_id': 'concl_4', 'topic': 'VP deform when UC press controller error', 'statement_from_report': 'When UC press VP/CD machine controller error NG when dry UC happen NG VP deform. PE repair machine controller error and ME training worker when dry UC.',
             'normalized_interpretation': 'Controller error event 24/24 NG (100%) all VP deform vs controller OK 0/40 (0.0%). Equipment failure cause confirmed.',
             'sheet_name': 'NG rate  (VP deform)', 'source_file': name, 'source_cells': ['NG rate (VP deform)!Decision text']},
        ],
        'troubleshooting_index': {
            'defect_name': 'NG Not Dry Glue / Over Glue / VP Deform on VP/CD',
            'when_user_asks': ['How to reduce NG not dry glue VP/CD?', 'Should we reduce CD press JIG magnet?', 'Does inserting tape on CD press JIG help?', 'Why does VP deform appear in bursts?'],
            'suggested_checks': [
                {'hint_id': 'hint_1', 'check_item': 'Adopt Change method UC press VP/CD (UV order: PRESS -> LED -> LED with PRESS JIG)', 'reason': 'Change method Total Vision VP/CD NG 0.8% vs Normal 2.7% = 70.4% improved. Source decision: "Can use new method".', 'evidence_strength': 'strong', 'related_process': 'UC press VP/CD', 'related_part': 'VP/CD', 'sheet_name': 'Change method', 'source_file': name, 'source_cells': ['Change method!Total + Decision']},
                {'hint_id': 'hint_2', 'check_item': 'Function Hearing on Change method must be re-checked', 'reason': 'Function Hearing +1V 5.0% vs Normal 3.0% = 65% worse and +0V 1.7% vs 1.1% = 50.9% worse on Change method despite Vision improvement.', 'evidence_strength': 'medium', 'related_process': 'Function (Hearing)', 'related_part': 'VP/CD', 'sheet_name': 'Change method', 'source_file': name, 'source_cells': ['Change method!Function Total']},
                {'hint_id': 'hint_3', 'check_item': 'Reduce CD press JIG magnet from 2 to 1', 'reason': 'Reduce magnet Total Vision NG 0.9% vs Normal 3.8% = 76.3% improved; Over glue and CD separate drops from 3.8% to 0.3%.', 'evidence_strength': 'strong', 'related_process': 'CD press JIG', 'related_part': 'CD', 'sheet_name': 'CD press JIG reduce MG', 'source_file': name, 'source_cells': ['CD press JIG reduce MG!Total']},
                {'hint_id': 'hint_4', 'check_item': 'Reject inserting tape into CD press JIG', 'reason': 'Insert tape Vision NG 10.3% (4/39) vs Normal 2.5% (1/40) = 312% worse; "VP/CD stickup on press JIG" when opening JIG.', 'evidence_strength': 'strong', 'related_process': 'CD press JIG', 'related_part': 'CD', 'sheet_name': 'CD press JIG insert tape', 'source_file': name, 'source_cells': ['CD press JIG insert tape!Decision']},
                {'hint_id': 'hint_5', 'check_item': 'UC press machine controller maintenance', 'reason': '24/24 (100%) VP deform during controller error; controller-OK reference shows 0/40 (0%).', 'evidence_strength': 'strong', 'related_process': 'UC press VP/CD', 'related_part': 'VP/CD machine controller', 'sheet_name': 'NG rate  (VP deform)', 'source_file': name, 'source_cells': ['NG rate (VP deform)!Error row']},
                {'hint_id': 'hint_6', 'check_item': 'Tension margin verification under Change method', 'reason': 'Tension AVG was NG (<1.9 kgf spec) on 2025-03-05, 03-06 Normal and 03-10 TEST runs; only 03-07 was OK.', 'evidence_strength': 'medium', 'related_process': 'UC press VP/CD Tension', 'related_part': 'VP/CD', 'sheet_name': 'Change method', 'source_file': name, 'source_cells': ['Change method!Tension table']},
            ],
            'limitations': ['Normal sample for Change method is only 960 pcs (one day); larger normal baseline desirable.', 'Tension spec >=1.9 kgf cited from source.'],
        },
        'ai_extraction_log': {'confidence': 0.83, 'assumptions': ['Source has visible "8100.0%" in Change method Total VP deform row (likely division typo); rate replaced with derived 0.1% (81/80928) when computing.'],
                              'warnings': ['Date 3/172025 normalized to 2025-03-17.', 'Some breakdown percentages in source are inconsistent (e.g., 253.3% / 113.3% style); only counts used for breakdown.'],
                              'decision_rationale': 'Mixed (Vision normal_comparison + Tension reliability_spec + controller-error case). Each baseline is the Normal row in the same sheet; relative changes computed per (test/baseline-1)*100.'}
    }
    tr_ko = {
        'document': {'title': 'BRS-201507DT/MSU-L2S15-07 VP/CD Not Dry Glue, Over Glue, VP Deform 원인 분석 및 개선', 'purpose': 'VP/CD 공정의 not dry glue / over glue / VP deform NG 원인 분석 및 JIG/방법 변경 평가.',
                     'content': ['Change method (UV 순서 변경) vs Normal Vision/Function/Tension.', 'CD press JIG 자석 2→1 비교.', 'CD press JIG insert tape vs Normal.', 'UC press controller error 시 VP deform.', 'Picture 시트로 방법 변경 설명.']},
        'conclusions': {
            'concl_1': {'topic': 'Change method UC press VP/CD', 'statement_from_report': '새로운 방법 사용 가능.', 'normalized_interpretation': 'Change method Total Vision NG 0.8%(608/80928) vs Normal 2.7%(26/960) = 0.30배(70.4% 개선). Function Hearing +1V Change 5.0% vs Normal 3.0% = 1.65배(65.0% 악화), +0V 1.7% vs 1.1% = 1.51배(50.9% 악화). Tension은 5회 중 3회 NG(1.9 kgf 미만).'},
            'concl_2': {'topic': 'CD press JIG 자석 1개로 감소', 'statement_from_report': '자석 1개 사용 시 NG over glue 및 CD separate 감소 → JIG 1자석 사용 가능.', 'normalized_interpretation': 'Reduce magnet Total Vision NG 0.9%(160/18731) vs Normal 3.8%(3/80) = 0.24배(76.3% 개선). Over glue and CD separate 3.8%→0.3%. Function Hearing +1V 3.6% vs 5.2% = 0.69배(30.8% 개선), +0V 1.1% vs 2.6% = 0.42배(57.7% 개선).'},
            'concl_3': {'topic': 'CD press JIG insert tape', 'statement_from_report': 'tape 삽입 시 JIG 열 때 VP/CD가 press JIG에 들러붙는 NG 발생 → 사용 불가.', 'normalized_interpretation': 'Insert tape Vision NG 10.3%(4/39) vs Normal 2.5%(1/40) = 4.12배(312% 악화). 채택 거절.'},
            'concl_4': {'topic': 'UC press controller error 시 VP deform', 'statement_from_report': 'Controller error 발생 시 VP deform NG 발생. PE 수리, ME 작업자 교육 필요.', 'normalized_interpretation': 'Controller error 24/24(100%) 모두 VP deform vs Controller OK 0/40(0%). 설비 고장이 원인 확정.'},
        },
        'hints': {
            'hint_1': {'check_item': 'Change method UC press VP/CD 채택', 'reason': 'Vision NG 0.8% vs Normal 2.7% = 70.4% 개선; "Can use new method" 결정.'},
            'hint_2': {'check_item': 'Change method 후 Function Hearing 재검토', 'reason': 'Vision은 개선되었으나 Function Hearing +1V 5.0% vs Normal 3.0%(65% 악화), +0V 1.7% vs 1.1%(50.9% 악화).'},
            'hint_3': {'check_item': 'CD press JIG 자석 2→1로 감소', 'reason': 'Vision NG 0.9% vs Normal 3.8% = 76.3% 개선; Over glue and CD separate 3.8%→0.3%.'},
            'hint_4': {'check_item': 'CD press JIG insert tape 적용 거절', 'reason': 'Insert tape NG 10.3% vs Normal 2.5% = 312% 악화; JIG 개방 시 VP/CD stickup 발생.'},
            'hint_5': {'check_item': 'UC press controller 정비/예방 점검', 'reason': 'Controller error 시 VP deform 24/24(100%); OK 시 0/40(0%).'},
            'hint_6': {'check_item': 'Change method 적용 시 Tension 마진 확인', 'reason': 'Tension AVG 1.9 kgf spec 미달 3회/5회 (3/5, 3/6, 3/10).'},
        },
        'log': {'assumptions': ['원본 Change method Total의 VP deform 표시 8100.0%는 분모 오기; 81/80928 = 약 0.1%로 해석.'],
                'warnings': ['Date 3/172025를 2025-03-17로 정규화.', '일부 breakdown 비율(253.3%/113.3% 등) 비정상 → count만 사용.'],
                'decision_rationale': 'Vision normal_comparison + Tension reliability_spec + controller-error case 혼합. baseline은 동일 시트의 Normal 행; (test/baseline-1)*100 사용.'}
    }
    tr_en = {
        'document': {'title': result['document']['title'], 'purpose': result['document']['purpose'], 'content': result['document']['content']},
        'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in result['conclusions']},
        'hints': {h['hint_id']: {'check_item': h['check_item'], 'reason': h['reason']} for h in result['troubleshooting_index']['suggested_checks']},
        'log': {'assumptions': result['ai_extraction_log']['assumptions'], 'warnings': result['ai_extraction_log']['warnings'], 'decision_rationale': result['ai_extraction_log']['decision_rationale']},
    }
    tr_vi = {
        'document': {'title': 'BRS-201507DT / MSU-L2S15-07 Report kiểm tra NG not dry glue / over glue / VP deform trên VP/CD và đánh giá đổi JIG/method', 'purpose': 'Kiểm tra nguyên nhân NG not dry glue và over glue trên VP/CD, đề xuất đổi JIG/method.',
                     'content': ['Đổi method UC press VP/CD vs dry thường (Vision/Function/Tension).', 'Giảm 1 magnet CD press JIG vs Normal.', 'Insert tape CD press JIG vs Normal.', 'Khi UC press controller error gây VP deform.', 'Sheet Picture giải thích đổi method.']},
        'conclusions': {
            'concl_1': {'topic': 'Đổi method UC press VP/CD', 'statement_from_report': 'Có thể dùng method mới.', 'normalized_interpretation': 'Change method Total Vision NG 0.8% (608/80928) vs Normal 2.7% (26/960) = 0.30x (cải thiện 70.4%). Function Hearing +1V Change 5.0% vs Normal 3.0% = 1.65x (xấu hơn 65.0%); +0V 1.7% vs 1.1% = 1.51x (xấu hơn 50.9%). Tension 3/5 lần NG (<1.9 kgf spec).'},
            'concl_2': {'topic': 'Giảm magnet CD press JIG', 'statement_from_report': 'Dùng JIG giảm 1 magnet, NG over glue và CD separate giảm => Dùng được JIG 1 magnet.', 'normalized_interpretation': 'Reduce magnet Total Vision NG 0.9% (160/18731) vs Normal 3.8% (3/80) = 0.24x (cải thiện 76.3%). Over glue and CD separate 3.8% → 0.3%. Function Hearing +1V 3.6% vs 5.2% = 0.69x (cải thiện 30.8%); +0V 1.1% vs 2.6% = 0.42x (cải thiện 57.7%).'},
            'concl_3': {'topic': 'CD press JIG insert tape', 'statement_from_report': 'Khi mở JIG xảy ra NG VP/CD stickup trên JIG => NG tăng so với normal => Không dùng được.', 'normalized_interpretation': 'Insert tape Vision NG 10.3% (4/39) vs Normal 2.5% (1/40) = 4.12x (xấu hơn 312%). Từ chối.'},
            'concl_4': {'topic': 'VP deform khi controller error UC press', 'statement_from_report': 'Khi controller error UC press, dry UC sinh NG VP deform. PE sửa controller, ME đào tạo công nhân.', 'normalized_interpretation': 'Controller error 24/24 (100%) đều VP deform vs Controller OK 0/40 (0%). Xác nhận do thiết bị.'},
        },
        'hints': {
            'hint_1': {'check_item': 'Áp dụng Change method UC press VP/CD', 'reason': 'Vision NG 0.8% vs Normal 2.7% = cải thiện 70.4%; decision: "Can use new method".'},
            'hint_2': {'check_item': 'Re-check Function Hearing dưới Change method', 'reason': 'Vision tốt hơn nhưng Function Hearing +1V 5.0% vs Normal 3.0% (xấu hơn 65%), +0V 1.7% vs 1.1% (xấu hơn 50.9%).'},
            'hint_3': {'check_item': 'Giảm CD press JIG magnet 2 → 1', 'reason': 'Vision NG 0.9% vs Normal 3.8% = cải thiện 76.3%; Over glue và CD separate 3.8% → 0.3%.'},
            'hint_4': {'check_item': 'Từ chối insert tape vào CD press JIG', 'reason': 'Insert tape NG 10.3% vs Normal 2.5% = xấu hơn 312%; VP/CD stickup khi mở JIG.'},
            'hint_5': {'check_item': 'Bảo trì controller UC press', 'reason': 'Controller error 24/24 VP deform; Controller OK 0/40.'},
            'hint_6': {'check_item': 'Verify Tension dưới Change method', 'reason': 'Tension AVG NG (<1.9 kgf spec) ở 3/5 lần (3/5, 3/6, 3/10).'},
        },
        'log': {'assumptions': ['Source Change method Total có VP deform 8100.0% (chia sai); thực tế 81/80928 ≈ 0.1%.'],
                'warnings': ['Date 3/172025 đã normalize thành 2025-03-17.', 'Một số % breakdown bất thường (253.3%/113.3%) → chỉ dùng count.'],
                'decision_rationale': 'Hỗn hợp Vision normal_comparison + Tension reliability_spec + controller-error. Baseline là hàng Normal trong cùng sheet; relative_change_percent = (test/baseline-1)*100.'}
    }
    return name, result, tr_ko, tr_en, tr_vi


def main():
    datasets = [ds1, ds2, ds3, ds4, ds5, ds6, ds7, ds8, ds9]
    processed = 0
    failed = 0
    for fn in datasets:
        try:
            name, result, tr_ko, tr_en, tr_vi = fn()
            ok = h.commit_dataset(name, result, tr_ko, tr_en, tr_vi)
            if ok:
                processed += 1
                print(f'OK   {name}')
            else:
                failed += 1
                print(f'FAIL {name}')
        except Exception as e:
            failed += 1
            print(f'EXC  {fn.__name__}: {e!r}')
            try:
                h.log_failed(fn.__name__, repr(e))
            except Exception:
                pass

    print(f'chunk 06: processed={processed} failed={failed}')

    # verify_counts uses targets file, but our chunk is independent.
    # Provide chunk-specific verification too.
    import sqlite3
    con = sqlite3.connect(h.DB_PATH)
    try:
        with open(r'D:\000. MyWorks\005. Program\Repository\JinoSupporter\_chunk_06.txt', 'r', encoding='utf-8-sig') as f:
            names = [l.strip() for l in f if l.strip()]
        present = 0
        for n in names:
            r = con.execute('SELECT COUNT(*) FROM AiDocuments WHERE SourceDataset=?', (n,)).fetchone()[0]
            if r > 0:
                present += 1
        print(f'chunk 06 verify: targets={len(names)} in_db={present}')
    finally:
        con.close()

    # Also call helper.verify_counts() for global view
    try:
        total, ok, failed_log = h.verify_counts()
        print(f'helper verify_counts: targets={total} processed={ok} failed_log_lines={failed_log}')
    except Exception as e:
        print(f'verify_counts error: {e!r}')


if __name__ == '__main__':
    main()
