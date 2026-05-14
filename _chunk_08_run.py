"""Process chunk 08 AI Batch normalization."""
from __future__ import annotations
import sys, io, importlib.util, pathlib

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

spec = importlib.util.spec_from_file_location(
    'h', r'D:\000. MyWorks\005. Program\Repository\JinoSupporter\_ai_batch_helper.py')
h = importlib.util.module_from_spec(spec); spec.loader.exec_module(h)


def base_doc(name, title_en, model, report_date, marker, line, report_type,
             primary_canonical, primary_aliases, related, parts, processes,
             purpose_en, content_en, source_cells):
    return {
        'document_id': '', 'source_file': name, 'source_sheet': '',
        'title': title_en, 'model': model, 'report_date': report_date,
        'department': 'ME', 'marker': marker, 'line': line,
        'report_type': report_type,
        'primary_defect': {'canonical_name': primary_canonical, 'aliases_in_document': primary_aliases},
        'related_defects': related, 'parts': parts, 'processes': processes,
        'purpose': purpose_en, 'content': content_en,
        'source_cells': source_cells,
    }


# -------------------- Dataset 1 --------------------
def ds1():
    name = "30. TIU C11-20  Report Test Frame damage mesh 2026.04.10. - RAW soundcheck"
    sheet = 'Test'
    doc = base_doc(
        name,
        "Report test Frame NG damage mesh TIU C11-20",
        "TIU C11-20L", "2026-04-10", "Trung", "",
        "normal_comparison",
        "NG Function", ["NG SPL", "NG Hearing Noise", "Low Gauss", "Particle", "VP+CD Offset"],
        ["NG Function", "NG SPL", "NG Hearing Noise"],
        ["Frame", "Mesh", "VP"], ["Frame assembly", "Soundcheck"],
        "Test improve NG Function on TIU C11-20 by separating frame mesh types.",
        ["Frame TIU C11-20L separate types and test function.",
         "Quantity tested: 100pcs per type.",
         "Same-event Normal baseline (300pcs) included for comparison."],
        {'title': ['Test!A1'], 'date': ['Test!Date'], 'purpose': ['Test!Purpose'], 'content': ['Test!Content']})

    conds = [
        {'condition_id': 'cond_1', 'condition_group': 'Frame test', 'line': 'TIU C11-20L',
         'process': 'Function test', 'changed_factor': 'Frame mesh type (Type1/2/3 vs Normal)',
         'before_value': 'Normal frame', 'after_value': 'Type1/Type2/Type3 mesh',
         'unit': None, 'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test!E18:E24']},
    ]

    results = [
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Type 1', 'date': '2026-04-10', 'line': 'TIU C11-20L',
         'input_count': 85, 'ok_count': 84, 'ng_count': 1,
         'ng_rate_decimal': 0.012, 'ng_rate_percent': 1.2,
         'metric_name': 'Total NG Rate', 'metric_value': 1.2, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG SPL': 1, 'NG SPL+RB': 0, 'NG No sound': 0, 'NG Hearing Noise': 0, 'NG Hearing Touch': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test!18:19']},
        {'result_id': 'res_2', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Type 2', 'date': '2026-04-10', 'line': 'TIU C11-20L',
         'input_count': 88, 'ok_count': 76, 'ng_count': 12,
         'ng_rate_decimal': 0.136, 'ng_rate_percent': 13.6,
         'metric_name': 'Total NG Rate', 'metric_value': 13.6, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG SPL': 4, 'NG SPL+RB': 0, 'NG No sound': 0, 'NG Hearing Noise': 8, 'NG Hearing Touch': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test!20:21']},
        {'result_id': 'res_3', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Type 3', 'date': '2026-04-10', 'line': 'TIU C11-20L',
         'input_count': 99, 'ok_count': 97, 'ng_count': 2,
         'ng_rate_decimal': 0.020, 'ng_rate_percent': 2.0,
         'metric_name': 'Total NG Rate', 'metric_value': 2.0, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG SPL': 0, 'NG SPL+RB': 0, 'NG No sound': 0, 'NG Hearing Noise': 2, 'NG Hearing Touch': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test!22:23']},
        {'result_id': 'res_4', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Normal (baseline)', 'date': '2026-04-10', 'line': 'TIU C11-20L',
         'input_count': 300, 'ok_count': 293, 'ng_count': 7,
         'ng_rate_decimal': 0.0233, 'ng_rate_percent': 2.3,
         'metric_name': 'Total NG Rate', 'metric_value': 2.3, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG SPL': 1, 'NG SPL+RB': 0, 'NG No sound': 0, 'NG Hearing Noise': 4, 'NG Hearing Touch': 2},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test!24:25']},
    ]

    # Relative changes vs Normal 2.3%:
    # T1 1.2/2.3-1 = -47.8% improved
    # T2 13.6/2.3-1 = +491.3% worse
    # T3 2.0/2.3-1 = -13.0% improved
    conclusions = [
        {'conclusion_id': 'concl_1', 'topic': 'Frame Type 2 vs Normal',
         'statement_from_report': "NG SPL type 1 and type 2 by low gauss; NG hearing by particle and ass'y VP+Frame offset.",
         'normalized_interpretation': 'Type 2 13.6% vs Normal 2.3% = 5.91x, 491.3% worse than same-event Normal (NG Hearing Noise 9.1% dominant + NG SPL 4.5%). Type 1 1.2% vs Normal 2.3% = -47.8% improved. Type 3 2.0% vs Normal 2.3% = -13.0% improved.',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test!26:31']},
    ]

    ts = {'defect_name': 'NG Function',
          'when_user_asks': ['Frame damage mesh', 'NG SPL low gauss', 'NG Hearing Noise particle'],
          'suggested_checks': [
              {'hint_id': 'hint_1', 'check_item': 'Check gauss on Frame Type 1 and Type 2 (NG SPL caused by low gauss)',
               'reason': 'Report decision states NG SPL on Type 1 and Type 2 is caused by low gauss; Type 2 worsened 491.3% vs Normal.',
               'evidence_strength': 'medium', 'related_process': 'Magnetization / Frame assembly',
               'related_part': 'Frame mesh', 'source_file': name, 'sheet_name': sheet,
               'source_cells': ['Test!29:31']}],
          'limitations': ['Soundcheck raw data not fully tabulated; only function NG summary used.']}

    log = {'confidence': 0.8,
           'assumptions': ['Normal row (300pcs) treated as same-event baseline for all three Type rows.'],
           'warnings': ['Type 1 sample size 85 is small; relative change may be noisy.'],
           'decision_rationale': 'Type 2 frame mesh significantly worse than Normal driven by NG Hearing Noise; Type 1 and Type 3 are similar to or slightly better than Normal.'}

    result = {'schema_version': '0.1', 'document': doc,
              'test_conditions': conds, 'results': results,
              'conclusions': conclusions, 'troubleshooting_index': ts,
              'ai_extraction_log': log}

    tr_en = {
        'document': {'title': doc['title'], 'purpose': doc['purpose'], 'content': doc['content']},
        'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in conclusions},
        'hints': {h_['hint_id']: {'check_item': h_['check_item'], 'reason': h_['reason']} for h_ in ts['suggested_checks']},
        'log': {'assumptions': log['assumptions'], 'warnings': log['warnings'], 'decision_rationale': log['decision_rationale']},
    }
    tr_ko = {
        'document': {'title': 'Frame NG damage mesh 시험 보고서 TIU C11-20',
                     'purpose': 'TIU C11-20에서 Frame mesh type을 분리하여 NG Function 개선 여부를 확인.',
                     'content': ['Frame TIU C11-20L mesh type 분리 후 기능 시험.', '시험 수량: type별 100pcs.', '동일 이벤트 Normal 기준(300pcs) 비교.']},
        'conclusions': {'concl_1': {'topic': 'Frame Type 2 vs Normal',
                                     'statement_from_report': 'NG SPL type 1·type 2는 low gauss, NG hearing은 particle 및 VP+Frame offset 이슈.',
                                     'normalized_interpretation': 'Type 2 13.6% vs Normal 2.3% = 5.91배, 동일 이벤트 Normal 대비 491.3% 악화 (NG Hearing Noise 9.1% 주도 + NG SPL 4.5%). Type 1 1.2% vs Normal 2.3% = -47.8% 개선. Type 3 2.0% vs Normal 2.3% = -13.0% 개선.'}},
        'hints': {'hint_1': {'check_item': 'Frame Type 1·Type 2 gauss 점검 (NG SPL 원인이 low gauss)',
                              'reason': '리포트 결정문에 Type 1·Type 2 NG SPL이 low gauss로 명시되어 있으며 Type 2는 Normal 대비 491.3% 악화.'}},
        'log': {'assumptions': ['Normal 300pcs 행을 세 Type 모두의 동일 이벤트 baseline 으로 사용.'],
                 'warnings': ['Type 1 표본 85pcs로 작아 상대 변화율 노이즈 가능.'],
                 'decision_rationale': 'Type 2 frame mesh는 NG Hearing Noise 주도로 Normal 대비 크게 악화; Type 1·Type 3은 Normal 과 동등 또는 약간 개선.'},
    }
    tr_vi = {
        'document': {'title': 'Báo cáo test Frame NG damage mesh TIU C11-20',
                     'purpose': 'Test cải thiện NG Function bằng cách tách loại frame mesh trên TIU C11-20.',
                     'content': ['Frame TIU C11-20L tách type và test function.', 'Số lượng test: 100pcs mỗi type.', 'Normal cùng sự kiện (300pcs) làm baseline.']},
        'conclusions': {'concl_1': {'topic': 'Frame Type 2 vs Normal',
                                     'statement_from_report': 'NG SPL type 1 và type 2 do low gauss; NG hearing do particle và VP+Frame offset.',
                                     'normalized_interpretation': 'Type 2 13.6% vs Normal 2.3% = 5.91x, xấu hơn Normal cùng sự kiện 491.3% (NG Hearing Noise 9.1% chủ đạo + NG SPL 4.5%). Type 1 1.2% vs Normal 2.3% = -47.8% cải thiện. Type 3 2.0% vs Normal 2.3% = -13.0% cải thiện.'}},
        'hints': {'hint_1': {'check_item': 'Kiểm tra gauss trên Frame Type 1 và Type 2 (NG SPL do low gauss)',
                              'reason': 'Báo cáo nêu NG SPL Type 1 và Type 2 do low gauss; Type 2 xấu hơn Normal 491.3%.'}},
        'log': {'assumptions': ['Hàng Normal 300pcs được dùng làm baseline cùng sự kiện cho cả ba Type.'],
                 'warnings': ['Type 1 cỡ mẫu 85pcs nhỏ, biến động tỷ lệ có thể nhiễu.'],
                 'decision_rationale': 'Frame mesh Type 2 xấu hơn Normal rõ rệt do NG Hearing Noise; Type 1 và Type 3 tương đương hoặc nhẹ hơn Normal.'},
    }
    return name, result, tr_ko, tr_en, tr_vi


# -------------------- Dataset 2 --------------------
def ds2():
    name = "30. TIU C11-20  Report test find reason NG NG function high 2026.01.02"
    sheet = 'Test (3)'
    doc = base_doc(
        name,
        "Report test find reason NG Function high TIU C11-20",
        "TIU C11-20R", "2026-01-09", "Thao", "TIU C11-20R",
        "normal_comparison",
        "NG Function High Rate", ["NG Function high", "RB", "Noise"],
        ["NG Function", "RB", "NG Hearing Noise"],
        ["CD rubber", "VP", "Heat forming No3"], ["VP forming", "CD rubber clean"],
        "Find reason of NG Function high on TIU C11-20.",
        ["Type 1: Test CD rubber no clean Ethanol (Temp 230°C).",
         "Type 2: Test heat forming No3 inside & outside 185°C.",
         "Compare with same-event Normal lot."],
        {'title': ['Test (3)!A1'], 'date': [], 'purpose': [], 'content': []})

    conds = [
        {'condition_id': 'cond_1', 'condition_group': 'CD rubber clean', 'line': 'TIU C11-20R',
         'process': 'CD rubber clean', 'changed_factor': 'Ethanol clean (Normal LSR 2x clean vs no clean / EPDM no clean)',
         'before_value': '2x Ethanol clean', 'after_value': 'No clean Ethanol',
         'unit': None, 'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!D58']},
        {'condition_id': 'cond_2', 'condition_group': 'Heat forming No3 temp', 'line': 'TIU C11-20R',
         'process': 'VP forming No3', 'changed_factor': 'Forming temperature',
         'before_value': '180', 'after_value': '185', 'unit': '°C',
         'temperature': '185', 'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!D60:D63']},
        {'condition_id': 'cond_3', 'condition_group': 'Separate position', 'line': 'TIU C11-20R',
         'process': 'VP+CD separation machine', 'changed_factor': 'Separation position (No1/No2/No3/No4)',
         'before_value': 'Position 1', 'after_value': 'Position 2/3/4',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!D82:D88']},
    ]

    results = [
        # Event 9-Jan
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Test CD rubber no clean Ethanol (9-Jan)',
         'date': '2026-01-09', 'line': 'TIU C11-20R',
         'input_count': 60, 'ok_count': 57, 'ng_count': 3, 'ng_rate_decimal': 0.05, 'ng_rate_percent': 5.0,
         'metric_name': 'Total NG Rate', 'metric_value': 5.0, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG SPL': 0, 'NG SPL+RB': 0, 'RB': 4, 'No sound': 0, 'Noise': 3, 'Touch': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!58:59']},
        {'result_id': 'res_2', 'condition_id': 'cond_2', 'measurement_type': 'Function',
         'condition_group': 'Heat forming No3 inside 185°C (9-Jan)',
         'date': '2026-01-09', 'line': 'TIU C11-20R',
         'input_count': 41, 'ok_count': 35, 'ng_count': 6, 'ng_rate_decimal': 0.146, 'ng_rate_percent': 14.6,
         'metric_name': 'Total NG Rate', 'metric_value': 14.6, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG SPL': 1, 'NG SPL+RB': 0, 'RB': 6, 'No sound': 0, 'Noise': 5, 'Touch': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!60:61']},
        {'result_id': 'res_3', 'condition_id': 'cond_2', 'measurement_type': 'Function',
         'condition_group': 'Heat forming No3 outside 185°C (9-Jan)',
         'date': '2026-01-09', 'line': 'TIU C11-20R',
         'input_count': 36, 'ok_count': 31, 'ng_count': 5, 'ng_rate_decimal': 0.139, 'ng_rate_percent': 13.9,
         'metric_name': 'Total NG Rate', 'metric_value': 13.9, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG SPL': 0, 'NG SPL+RB': 0, 'RB': 5, 'No sound': 0, 'Noise': 5, 'Touch': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!62:63']},
        {'result_id': 'res_4', 'condition_id': None, 'measurement_type': 'Function',
         'condition_group': 'Normal (Temp 180°C) 9-Jan baseline',
         'date': '2026-01-09', 'line': 'TIU C11-20R',
         'input_count': 50, 'ok_count': 46, 'ng_count': 4, 'ng_rate_decimal': 0.08, 'ng_rate_percent': 8.0,
         'metric_name': 'Total NG Rate', 'metric_value': 8.0, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG SPL': 1, 'NG SPL+RB': 0, 'RB': 3, 'No sound': 0, 'Noise': 3, 'Touch': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!64:65']},
        # 10-Jan
        {'result_id': 'res_5', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Test CD rubber no clean Ethanol 2nd (10-Jan)',
         'date': '2026-01-10', 'line': 'TIU C11-20R',
         'input_count': 68, 'ok_count': 65, 'ng_count': 3, 'ng_rate_decimal': 0.044, 'ng_rate_percent': 4.4,
         'metric_name': 'Total NG Rate', 'metric_value': 4.4, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG SPL': 1, 'NG SPL+RB': 0, 'RB': 4, 'No sound': 0, 'Noise': 2, 'Touch': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!66:67']},
        {'result_id': 'res_6', 'condition_id': None, 'measurement_type': 'Function',
         'condition_group': 'Normal 10-Jan baseline',
         'date': '2026-01-10', 'line': 'TIU C11-20R',
         'input_count': 50, 'ok_count': 47, 'ng_count': 3, 'ng_rate_decimal': 0.06, 'ng_rate_percent': 6.0,
         'metric_name': 'Total NG Rate', 'metric_value': 6.0, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG SPL': 1, 'NG SPL+RB': 0, 'RB': 2, 'No sound': 0, 'Noise': 2, 'Touch': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!68:69']},
        # 13-Jan
        {'result_id': 'res_7', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Test CD rubber no clean Ethanol 3rd (13-Jan)',
         'date': '2026-01-13', 'line': 'TIU C11-20R',
         'input_count': 54, 'ok_count': 50, 'ng_count': 4, 'ng_rate_decimal': 0.074, 'ng_rate_percent': 7.4,
         'metric_name': 'Total NG Rate', 'metric_value': 7.4, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG SPL': 0, 'NG SPL+RB': 0, 'RB': 4, 'No sound': 0, 'Noise': 4, 'Touch': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!70:71']},
        {'result_id': 'res_8', 'condition_id': None, 'measurement_type': 'Function',
         'condition_group': 'Normal 13-Jan baseline',
         'date': '2026-01-13', 'line': 'TIU C11-20R',
         'input_count': 50, 'ok_count': 47, 'ng_count': 3, 'ng_rate_decimal': 0.06, 'ng_rate_percent': 6.0,
         'metric_name': 'Total NG Rate', 'metric_value': 6.0, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG SPL': 1, 'NG SPL+RB': 0, 'RB': 2, 'No sound': 0, 'Noise': 2, 'Touch': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!72:73']},
        # Tension
        {'result_id': 'res_9', 'condition_id': None, 'measurement_type': 'Tension',
         'condition_group': 'Normal LSR Rubber (2x clean) 9-Jan',
         'date': '2026-01-09', 'line': 'TIU C11-20R',
         'metric_name': 'Tension AVG', 'metric_value': 0.892, 'unit': 'Kgf', 'judgement': 'PASS',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!76']},
        {'result_id': 'res_10', 'condition_id': None, 'measurement_type': 'Tension',
         'condition_group': 'EPDM rubber 65A No Clean 9-Jan',
         'date': '2026-01-09', 'line': 'TIU C11-20R',
         'metric_name': 'Tension AVG', 'metric_value': 0.812, 'unit': 'Kgf', 'judgement': 'PASS',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!77']},
        {'result_id': 'res_11', 'condition_id': None, 'measurement_type': 'Tension',
         'condition_group': 'EPDM rubber 65A No Clean 10-Jan',
         'date': '2026-01-10', 'line': 'TIU C11-20R',
         'metric_name': 'Tension AVG', 'metric_value': 0.838, 'unit': 'Kgf', 'judgement': 'PASS',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!78']},
        # Position
        {'result_id': 'res_12', 'condition_id': 'cond_3', 'measurement_type': 'Function',
         'condition_group': 'Separate Position No1', 'date': '2026-01-09', 'line': 'TIU C11-20R',
         'input_count': 58, 'ok_count': 56, 'ng_count': 2, 'ng_rate_decimal': 0.034, 'ng_rate_percent': 3.4,
         'metric_name': 'Total NG Rate', 'metric_value': 3.4, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'RB': 2, 'Noise': 2},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!82:83']},
        {'result_id': 'res_13', 'condition_id': 'cond_3', 'measurement_type': 'Function',
         'condition_group': 'Separate Position No2', 'date': '2026-01-09', 'line': 'TIU C11-20R',
         'input_count': 59, 'ok_count': 54, 'ng_count': 5, 'ng_rate_decimal': 0.085, 'ng_rate_percent': 8.5,
         'metric_name': 'Total NG Rate', 'metric_value': 8.5, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG SPL': 1, 'RB': 5, 'Noise': 4},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!84:85']},
        {'result_id': 'res_14', 'condition_id': 'cond_3', 'measurement_type': 'Function',
         'condition_group': 'Separate Position No3', 'date': '2026-01-09', 'line': 'TIU C11-20R',
         'input_count': 55, 'ok_count': 50, 'ng_count': 5, 'ng_rate_decimal': 0.091, 'ng_rate_percent': 9.1,
         'metric_name': 'Total NG Rate', 'metric_value': 9.1, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG SPL': 2, 'RB': 3, 'Noise': 3},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!86:87']},
        {'result_id': 'res_15', 'condition_id': 'cond_3', 'measurement_type': 'Function',
         'condition_group': 'Separate Position No4', 'date': '2026-01-09', 'line': 'TIU C11-20R',
         'input_count': 59, 'ok_count': 54, 'ng_count': 5, 'ng_rate_decimal': 0.085, 'ng_rate_percent': 8.5,
         'metric_name': 'Total NG Rate', 'metric_value': 8.5, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'SPL+RB': 2, 'RB': 3, 'Noise': 3},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!88:89']},
    ]

    # rel: CD no clean 5.0/8.0-1 = -37.5%; 4.4/6.0-1 = -26.7%; 7.4/6.0-1 = +23.3%
    # Heat forming inside 14.6/8.0-1 = +82.5% worse; outside 13.9/8.0-1 = +73.8% worse
    # Position No1 3.4% vs avg-of-positions baseline; report says "NG same" without explicit normal
    conclusions = [
        {'conclusion_id': 'concl_1', 'topic': 'CD rubber no clean Ethanol vs Normal',
         'statement_from_report': "Test CD rubber no clean Ethanol vs Normal (Temp 230°C / 180°C) trials over 9/10/13-Jan.",
         'normalized_interpretation': "CD rubber no clean: 9-Jan 5.0% vs Normal 8.0% = -37.5% improved; 10-Jan 4.4% vs Normal 6.0% = -26.7% improved; 13-Jan 7.4% vs Normal 6.0% = +23.3% worse. Tension all PASS (0.81-0.89 Kgf >= 0.4 spec).",
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!58:73']},
        {'conclusion_id': 'concl_2', 'topic': 'Heat forming No3 185°C vs Normal 180°C',
         'statement_from_report': 'Heat forming No3 inside and outside increased temp to 185°C; Normal 180°C baseline 8.0%.',
         'normalized_interpretation': 'Heat forming inside 14.6% vs Normal 8.0% = +82.5% worse; outside 13.9% vs Normal 8.0% = +73.8% worse. Both driven by RB and NG Hearing Noise.',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test (3)!60:65']},
    ]

    ts = {'defect_name': 'NG Function High Rate',
          'when_user_asks': ['NG function high', 'CD rubber clean Ethanol', 'Heat forming No3 temperature'],
          'suggested_checks': [
              {'hint_id': 'hint_1', 'check_item': 'Heat forming No3 temperature (185°C vs Normal 180°C)',
               'reason': '185°C inside 14.6% (+82.5% worse vs Normal 8.0%) and outside 13.9% (+73.8% worse) driven by RB and NG Hearing Noise — do not raise above 180°C.',
               'evidence_strength': 'strong', 'related_process': 'VP forming No3',
               'related_part': 'VP', 'source_file': name, 'sheet_name': sheet,
               'source_cells': ['Test (3)!60:65']}],
          'limitations': ['Position separation test "NG same" comment in report does not specify which position is baseline.']}

    log = {'confidence': 0.78,
           'assumptions': ['Same-date Normal rows used as baseline per trial day.',
                            'NG Function High Rate is the cross-event primary defect.'],
           'warnings': ['CD rubber no clean shows mixed direction (improved 9/10-Jan, worse 13-Jan); not a stable improvement.'],
           'decision_rationale': 'Heat forming No3 at 185°C is clearly worse than Normal 180°C on both inside and outside. CD rubber no clean Ethanol is inconclusive (mixed across days). Tension is PASS in all variants and not a driver.'}

    result = {'schema_version': '0.1', 'document': doc,
              'test_conditions': conds, 'results': results,
              'conclusions': conclusions, 'troubleshooting_index': ts,
              'ai_extraction_log': log}

    tr_en = {
        'document': {'title': doc['title'], 'purpose': doc['purpose'], 'content': doc['content']},
        'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in conclusions},
        'hints': {h_['hint_id']: {'check_item': h_['check_item'], 'reason': h_['reason']} for h_ in ts['suggested_checks']},
        'log': {'assumptions': log['assumptions'], 'warnings': log['warnings'], 'decision_rationale': log['decision_rationale']},
    }
    tr_ko = {
        'document': {'title': 'TIU C11-20 NG Function 고율 원인 분석 보고서',
                     'purpose': 'TIU C11-20에서 NG Function 고율 원인 파악.',
                     'content': ['Type 1: CD rubber no clean Ethanol (230°C).', 'Type 2: Heat forming No3 inside·outside 185°C.', '동일 이벤트 Normal lot 비교.']},
        'conclusions': {'concl_1': {'topic': 'CD rubber no clean Ethanol vs Normal',
                                     'statement_from_report': '9·10·13-Jan에 걸쳐 CD rubber no clean Ethanol vs Normal 비교.',
                                     'normalized_interpretation': 'CD rubber no clean: 9-Jan 5.0% vs Normal 8.0% = -37.5% 개선; 10-Jan 4.4% vs Normal 6.0% = -26.7% 개선; 13-Jan 7.4% vs Normal 6.0% = +23.3% 악화. Tension 모두 PASS (0.81-0.89 Kgf, spec ≥0.4).'},
                         'concl_2': {'topic': 'Heat forming No3 185°C vs Normal 180°C',
                                     'statement_from_report': 'Heat forming No3 inside·outside 185°C 상승 시험; Normal 180°C 8.0% baseline.',
                                     'normalized_interpretation': 'Heat forming inside 14.6% vs Normal 8.0% = +82.5% 악화; outside 13.9% vs Normal 8.0% = +73.8% 악화. RB 와 NG Hearing Noise 주도.'}},
        'hints': {'hint_1': {'check_item': 'Heat forming No3 온도 점검 (185°C vs Normal 180°C)',
                              'reason': '185°C inside 14.6% (+82.5% 악화) outside 13.9% (+73.8% 악화), RB·Noise 주도 — 180°C 이상 올리지 말 것.'}},
        'log': {'assumptions': ['동일 일자 Normal 행을 일자별 baseline 으로 사용.', 'NG Function High Rate 를 교차 이벤트 primary defect 로 설정.'],
                 'warnings': ['CD rubber no clean 은 일자별 방향이 혼재 — 안정적 개선 아님.'],
                 'decision_rationale': 'Heat forming No3 185°C 는 Normal 180°C 보다 inside·outside 모두 명확히 악화. CD rubber no clean Ethanol 은 일자 간 혼재로 결론 보류. Tension 은 모두 PASS 로 원인 아님.'},
    }
    tr_vi = {
        'document': {'title': 'Báo cáo tìm nguyên nhân NG Function cao TIU C11-20',
                     'purpose': 'Tìm nguyên nhân NG Function cao trên TIU C11-20.',
                     'content': ['Type 1: Test CD rubber no clean Ethanol (230°C).', 'Type 2: Heat forming No3 inside·outside 185°C.', 'So sánh với Normal lot cùng sự kiện.']},
        'conclusions': {'concl_1': {'topic': 'CD rubber no clean Ethanol vs Normal',
                                     'statement_from_report': 'Test CD rubber no clean Ethanol vs Normal trong 9/10/13-Jan.',
                                     'normalized_interpretation': 'CD rubber no clean: 9-Jan 5.0% vs Normal 8.0% = -37.5% cải thiện; 10-Jan 4.4% vs Normal 6.0% = -26.7% cải thiện; 13-Jan 7.4% vs Normal 6.0% = +23.3% xấu hơn. Tension PASS hết (0.81-0.89 Kgf, spec ≥0.4).'},
                         'concl_2': {'topic': 'Heat forming No3 185°C vs Normal 180°C',
                                     'statement_from_report': 'Heat forming No3 inside·outside tăng lên 185°C; Normal 180°C 8.0% baseline.',
                                     'normalized_interpretation': 'Heat forming inside 14.6% vs Normal 8.0% = +82.5% xấu hơn; outside 13.9% vs Normal 8.0% = +73.8% xấu hơn. Chủ yếu do RB và NG Hearing Noise.'}},
        'hints': {'hint_1': {'check_item': 'Kiểm tra nhiệt độ Heat forming No3 (185°C vs Normal 180°C)',
                              'reason': '185°C inside 14.6% (+82.5% xấu hơn) outside 13.9% (+73.8% xấu hơn), do RB và Noise — không nâng nhiệt độ trên 180°C.'}},
        'log': {'assumptions': ['Hàng Normal cùng ngày dùng làm baseline.', 'NG Function High Rate là primary defect xuyên các thử nghiệm.'],
                 'warnings': ['CD rubber no clean xu hướng không nhất quán giữa các ngày.'],
                 'decision_rationale': 'Heat forming No3 185°C xấu hơn Normal 180°C rõ ràng cả inside lẫn outside. CD rubber no clean Ethanol chưa kết luận được do xu hướng trộn lẫn. Tension PASS không phải nguyên nhân.'},
    }
    return name, result, tr_ko, tr_en, tr_vi


# -------------------- Dataset 3 --------------------
def ds3():
    name = "30. TIU L5S3-01 R Report test Dry final sample and wait dry 12h to check function 2025.12.16"
    sheet = 'Test'
    doc = base_doc(
        name,
        "TIU L5S3-01 [R] Report test Dry final sample and wait 12h to check function",
        "TIU L5S3-01R", "2025-12-16", "Nhung", "",
        "ng_without_baseline",
        "NG Function", ["NG FRF high", "NG FRF low", "FRF+SPL", "THD", "No sound"],
        ["NG Function", "NG FRF"],
        ["VP", "Final sample"], ["Dry oven", "Function check"],
        "Find reason of NG Function and improve via dry condition variation.",
        ["Type 1: 300pcs, Dry oven 60°C / 30min, wait 30min, check function, recheck 3 days.",
         "Type 2: 300pcs, wait 12h, check function, recheck 3 days.",
         "Compare with normal — but no same-event Normal row provided in this sheet."],
        {'title': ['Test!A1'], 'date': [], 'purpose': [], 'content': []})

    conds = [
        {'condition_id': 'cond_1', 'condition_group': 'Type 1 dry oven', 'line': 'TIU L5S3-01R',
         'process': 'Final dry', 'changed_factor': 'Dry oven 60°C / 30min + wait 30min',
         'before_value': None, 'after_value': '60°C 30min + 30min wait',
         'unit': None, 'temperature': '60', 'dry_time_sec': 1800,
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test!Type1']},
        {'condition_id': 'cond_2', 'condition_group': 'Type 2 wait 12h', 'line': 'TIU L5S3-01R',
         'process': 'Final dry', 'changed_factor': 'Wait 12 hour (no oven)',
         'before_value': None, 'after_value': '12 hour wait',
         'unit': 'hour', 'dry_time_sec': 43200,
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test!Type2']},
    ]

    results = [
        # Type 1
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Type 1 Before (16-Dec)', 'date': '2025-12-16', 'line': 'TIU L5S3-01R',
         'input_count': 300, 'ok_count': 234, 'ng_count': 66, 'ng_rate_decimal': 0.22, 'ng_rate_percent': 22.0,
         'metric_name': 'Total NG Rate', 'metric_value': 22.0, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG FRF high': 47, 'NG FRF low': 19, 'FRF+SPL': 0, 'THD': 0, 'No sound': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test!108:109']},
        {'result_id': 'res_2', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Type 1 After 1 day Total (17-Dec)', 'date': '2025-12-17', 'line': 'TIU L5S3-01R',
         'input_count': 300, 'ok_count': 265, 'ng_count': 35, 'ng_rate_decimal': 0.117, 'ng_rate_percent': 11.7,
         'metric_name': 'Total NG Rate', 'metric_value': 11.7, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG FRF high': 29, 'NG FRF low': 6},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test!116:117']},
        {'result_id': 'res_3', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Type 1 After 2 day Total (18-Dec)', 'date': '2025-12-18', 'line': 'TIU L5S3-01R',
         'input_count': 300, 'ok_count': 242, 'ng_count': 58, 'ng_rate_decimal': 0.193, 'ng_rate_percent': 19.3,
         'metric_name': 'Total NG Rate', 'metric_value': 19.3, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG FRF high': 53, 'NG FRF low': 5},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test!124:125']},
        {'result_id': 'res_4', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Type 1 After 3 day Total (19-Dec)', 'date': '2025-12-19', 'line': 'TIU L5S3-01R',
         'input_count': 300, 'ok_count': 245, 'ng_count': 55, 'ng_rate_decimal': 0.183, 'ng_rate_percent': 18.3,
         'metric_name': 'Total NG Rate', 'metric_value': 18.3, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG FRF high': 51, 'NG FRF low': 4},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test!132:133']},
        # Type 2
        {'result_id': 'res_5', 'condition_id': 'cond_2', 'measurement_type': 'Function',
         'condition_group': 'Type 2 Before (17-Dec)', 'date': '2025-12-17', 'line': 'TIU L5S3-01R',
         'input_count': 300, 'ok_count': 246, 'ng_count': 54, 'ng_rate_decimal': 0.18, 'ng_rate_percent': 18.0,
         'metric_name': 'Total NG Rate', 'metric_value': 18.0, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG FRF high': 46, 'NG FRF low': 3, 'FRF+SPL': 5},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test!134:135']},
        {'result_id': 'res_6', 'condition_id': 'cond_2', 'measurement_type': 'Function',
         'condition_group': 'Type 2 After 2 day Total (18-Dec)', 'date': '2025-12-18', 'line': 'TIU L5S3-01R',
         'input_count': 300, 'ok_count': 232, 'ng_count': 68, 'ng_rate_decimal': 0.227, 'ng_rate_percent': 22.7,
         'metric_name': 'Total NG Rate', 'metric_value': 22.7, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG FRF high': 60, 'NG FRF low': 8},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test!142:143']},
        {'result_id': 'res_7', 'condition_id': 'cond_2', 'measurement_type': 'Function',
         'condition_group': 'Type 2 After 3 day Total (19-Dec)', 'date': '2025-12-19', 'line': 'TIU L5S3-01R',
         'input_count': 300, 'ok_count': 244, 'ng_count': 56, 'ng_rate_decimal': 0.187, 'ng_rate_percent': 18.7,
         'metric_name': 'Total NG Rate', 'metric_value': 18.7, 'unit': '%', 'judgement': None,
         'ng_breakdown': {'NG FRF high': 47, 'NG FRF low': 9},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test!150:151']},
    ]

    conclusions = [
        {'conclusion_id': 'concl_1', 'topic': 'Type 1 oven dry vs Type 2 wait 12h (no Normal baseline)',
         'statement_from_report': 'Report compares oven dry 60°C/30min + 30min wait vs 12h wait; no Normal baseline row provided.',
         'normalized_interpretation': 'Without same-event Normal, cannot say improve/worsen. Absolute NG: Type 1 Before 22.0%, settles to 18.3% by Day 3. Type 2 Before 18.0%, settles to 18.7% by Day 3. NG FRF high dominates both types (>=78% within NG mix on later days). Type 2 starts lower but Type 1 catches up after dry; both remain ~18-19%.',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['Test!108:151']},
    ]

    ts = {'defect_name': 'NG Function',
          'when_user_asks': ['NG FRF high', 'Dry oven condition', 'Wait 12h'],
          'suggested_checks': [
              {'hint_id': 'hint_1', 'check_item': 'Verify NG FRF high stability after dry — recheck OK lots after 2-3 days',
               'reason': 'In both Type 1 and Type 2, the "After 1/2/3 day" recheck of previously OK lots produces NG FRF high 75-97% rates, indicating drift after settling.',
               'evidence_strength': 'medium', 'related_process': 'Final dry / aging',
               'related_part': 'VP / FRF resonance', 'source_file': name, 'sheet_name': sheet,
               'source_cells': ['Test!120:121', 'Test!128:129', 'Test!138:139']}],
          'limitations': ['No same-event Normal baseline present — only test variants are compared.']}

    log = {'confidence': 0.65,
           'assumptions': ['No same-event Normal — classified as ng_without_baseline.'],
           'warnings': ['Type 2 Day 3 NG FRF-low row shows -1 OK count vs 8 input, indicating data entry issue.',
                         'IV. Decision section in report is empty.'],
           'decision_rationale': 'Both dry conditions converge to ~18-19% NG with NG FRF high dominant. Neither shows clear advantage over the other; no Normal baseline included to assess overall improvement.'}

    result = {'schema_version': '0.1', 'document': doc,
              'test_conditions': conds, 'results': results,
              'conclusions': conclusions, 'troubleshooting_index': ts,
              'ai_extraction_log': log}

    tr_en = {
        'document': {'title': doc['title'], 'purpose': doc['purpose'], 'content': doc['content']},
        'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in conclusions},
        'hints': {h_['hint_id']: {'check_item': h_['check_item'], 'reason': h_['reason']} for h_ in ts['suggested_checks']},
        'log': {'assumptions': log['assumptions'], 'warnings': log['warnings'], 'decision_rationale': log['decision_rationale']},
    }
    tr_ko = {
        'document': {'title': 'TIU L5S3-01 [R] Final sample dry 및 12h 대기 기능 시험 보고서',
                     'purpose': 'NG Function 원인 파악 및 dry 조건 변경으로 개선 시도.',
                     'content': ['Type 1: 300pcs, 60°C/30min 오븐 dry + 30min 대기 후 검사, 3일간 재검사.', 'Type 2: 300pcs, 12시간 대기 후 검사, 3일간 재검사.', 'Normal 비교 의도였으나 동일 이벤트 Normal 행이 시트에 없음.']},
        'conclusions': {'concl_1': {'topic': 'Type 1 오븐 dry vs Type 2 12h 대기 (Normal baseline 없음)',
                                     'statement_from_report': '오븐 dry 60°C/30min+30min 대기 vs 12h 대기 비교; Normal baseline 행 없음.',
                                     'normalized_interpretation': '동일 이벤트 Normal 부재로 개선/악화 단정 불가. 절대치: Type 1 Before 22.0% → Day 3 18.3%; Type 2 Before 18.0% → Day 3 18.7%. NG FRF high 가 양 type 의 NG 구성 대부분(78% 이상). Type 2 가 초기엔 낮지만 dry 후 두 조건 모두 18-19% 수준 수렴.'}},
        'hints': {'hint_1': {'check_item': 'Dry 후 OK lot 의 NG FRF high drift 재확인 (1/2/3 일 후)',
                              'reason': 'Type 1·Type 2 모두 After 1/2/3 day 재검 시 기존 OK lot 에서 NG FRF high 가 75-97%로 발생 — 안정성 미흡.'}},
        'log': {'assumptions': ['동일 이벤트 Normal 부재로 ng_without_baseline 분류.'],
                 'warnings': ['Type 2 Day 3 NG FRF-low 행에 OK -1 / Input 8 입력 오류 추정.', '리포트 IV. Decision 비어 있음.'],
                 'decision_rationale': '두 dry 조건 모두 ~18-19%로 수렴하며 NG FRF high 가 주도. 어느 한 쪽이 명확히 우세하지 않음; 전체 개선 평가용 Normal baseline 부재.'},
    }
    tr_vi = {
        'document': {'title': 'Báo cáo TIU L5S3-01 [R] test Dry final sample và chờ 12h kiểm function',
                     'purpose': 'Tìm nguyên nhân NG Function và cải thiện qua điều kiện dry.',
                     'content': ['Type 1: 300pcs, Dry oven 60°C/30min + chờ 30min rồi kiểm, kiểm lại 3 ngày.', 'Type 2: 300pcs, chờ 12 giờ rồi kiểm, kiểm lại 3 ngày.', 'Định so với normal nhưng không có hàng Normal cùng sự kiện trong sheet.']},
        'conclusions': {'concl_1': {'topic': 'Type 1 dry oven vs Type 2 chờ 12h (không có Normal baseline)',
                                     'statement_from_report': 'So sánh dry oven 60°C/30min+30min wait vs chờ 12h; không có hàng Normal baseline.',
                                     'normalized_interpretation': 'Không có Normal cùng sự kiện nên không kết luận cải thiện/xấu hơn. Giá trị tuyệt đối: Type 1 Before 22.0% → Day 3 18.3%; Type 2 Before 18.0% → Day 3 18.7%. NG FRF high chiếm chủ đạo trong NG mix (>=78%). Type 2 thấp hơn lúc đầu nhưng sau dry cả hai đều ~18-19%.'}},
        'hints': {'hint_1': {'check_item': 'Kiểm tra drift NG FRF high sau dry — kiểm lại lot OK sau 2-3 ngày',
                              'reason': 'Cả Type 1 và Type 2, kiểm lại lot trước đó OK sau 1/2/3 ngày cho NG FRF high 75-97% — không ổn định.'}},
        'log': {'assumptions': ['Không có Normal cùng sự kiện — phân loại ng_without_baseline.'],
                 'warnings': ['Type 2 Day 3 NG FRF-low hàng có OK -1 / Input 8 — có thể nhập sai.', 'Phần IV. Decision của báo cáo trống.'],
                 'decision_rationale': 'Cả hai điều kiện dry hội tụ ~18-19% với NG FRF high chủ đạo. Không có ưu thế rõ ràng; thiếu Normal baseline để đánh giá cải thiện tổng thể.'},
    }
    return name, result, tr_ko, tr_en, tr_vi


# -------------------- Dataset 4 --------------------
def ds4():
    name = "30.1 MSU-L20S15-07  Report checking  gauss yoke date 10.6.2025"
    sheet = '17.1'
    doc = base_doc(
        name,
        "Report check gauss Yoke MSU-L20S15-07",
        "MSU-L20S15-07", "2025-06-10", "Le", "C2-2A",
        "reliability_spec",
        "Low Gauss", ["SMG short", "SMG long", "CMG/PT low"],
        ["Low Gauss", "NG Sigma"],
        ["Yoke", "SMG", "CMG", "PT"], ["Magnetization", "Gauss check"],
        "Check gauss on Yoke when SMG or CMG is low.",
        ["Make sample with low gauss.", "Measure SMG short, SMG long, MG+PT, Final SPK and Function (Sigma/Hearing)."],
        {'title': ['17.1!A1'], 'date': [], 'purpose': [], 'content': []})

    conds = [
        {'condition_id': 'cond_1', 'condition_group': 'SMG low + CMG OK', 'line': 'C2-2A',
         'process': 'Magnetization', 'changed_factor': 'Intentional SMG low magnetization',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['17.1!E171:E172']},
        {'condition_id': 'cond_2', 'condition_group': 'SMG OK + CMG low', 'line': 'C2-2A',
         'process': 'Magnetization', 'changed_factor': 'Intentional CMG low magnetization',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['17.1!E173:E174']},
    ]

    results = [
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Gauss',
         'condition_group': 'SMG low + CMG OK #1 (Sigma OK / Hearing OK)', 'date': '2025-06-10', 'line': 'C2-2A',
         'metric_name': 'SMG Short', 'metric_value': 982, 'unit': 'G', 'judgement': 'CHECK',
         'ng_breakdown': {'SMG Short (spec 900-1350)': 982, 'SMG Long': 992, 'MG+PT': 1075, 'Final SPK': 1028},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['17.1!171']},
        {'result_id': 'res_2', 'condition_id': 'cond_1', 'measurement_type': 'Gauss',
         'condition_group': 'SMG low + CMG OK #2 (Sigma NG / Hearing OK)', 'date': '2025-06-10', 'line': 'C2-2A',
         'metric_name': 'SMG Short', 'metric_value': 492, 'unit': 'G', 'judgement': 'FAIL',
         'ng_breakdown': {'SMG Short (spec 900-1350)': 492, 'SMG Long': 588, 'MG+PT': 448, 'Final SPK': 932},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['17.1!172']},
        {'result_id': 'res_3', 'condition_id': 'cond_2', 'measurement_type': 'Gauss',
         'condition_group': 'SMG OK + CMG low #1 (Sigma NG / Hearing OK)', 'date': '2025-06-10', 'line': 'C2-2A',
         'metric_name': 'CMG/PT', 'metric_value': 304, 'unit': 'G', 'judgement': 'FAIL',
         'ng_breakdown': {'SMG Short': 1124, 'SMG Long': 1184, 'MG+PT (spec 400-570)': 304, 'Final SPK': 1184},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['17.1!173']},
        {'result_id': 'res_4', 'condition_id': 'cond_2', 'measurement_type': 'Gauss',
         'condition_group': 'SMG OK + CMG low #2 (Sigma NG / Hearing OK)', 'date': '2025-06-10', 'line': 'C2-2A',
         'metric_name': 'CMG/PT', 'metric_value': 305, 'unit': 'G', 'judgement': 'FAIL',
         'ng_breakdown': {'SMG Short': 1221, 'SMG Long': 1168, 'MG+PT (spec 400-570)': 305, 'Final SPK': 1169},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['17.1!174']},
    ]

    conclusions = [
        {'conclusion_id': 'concl_1', 'topic': 'Gauss low samples produce Sigma NG but Hearing OK',
         'statement_from_report': 'IV. Decision section is empty in report.',
         'normalized_interpretation': 'SMG low #2 (SMG short 492G, below spec 900-1350G) -> Sigma NG. CMG low #1/#2 (MG+PT ~304-305G, below spec 400-570G) -> Sigma NG. All four cases keep Hearing OK. SMG low #1 (982G in spec) -> Sigma OK confirms spec threshold drives Sigma NG.',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['17.1!171:174']},
    ]

    ts = {'defect_name': 'Low Gauss',
          'when_user_asks': ['Low Gauss', 'Sigma NG', 'SMG short out of spec', 'CMG/PT out of spec'],
          'suggested_checks': [
              {'hint_id': 'hint_1', 'check_item': 'Check SMG short and MG+PT against spec (SMG 900-1350G, MG+PT 400-570G)',
               'reason': 'Samples below spec produced Sigma NG in all observed cases (#2 SMG 492G, #3/#4 MG+PT 304-305G); in-spec sample #1 (SMG 982G) gave Sigma OK.',
               'evidence_strength': 'strong', 'related_process': 'Magnetization',
               'related_part': 'Yoke / SMG / CMG / PT', 'source_file': name, 'sheet_name': sheet,
               'source_cells': ['17.1!171:174']}],
          'limitations': ['Only 4 samples; report decision section is empty.']}

    log = {'confidence': 0.7,
           'assumptions': ['Spec values read from sheet header (SMG 900-1350G, MG+PT 400-570G, Final SPK 390-590G).'],
           'warnings': ['Final SPK spec 390-590G but samples #1 (1028), #3 (1184), #4 (1169) far exceed upper bound — likely different metric unit/scale; flagged for manual review.'],
           'decision_rationale': 'Low gauss out-of-spec correlates 1:1 with Sigma NG; Hearing remains OK in all cases. This supports spec-based gauss screening to prevent Sigma NG.'}

    result = {'schema_version': '0.1', 'document': doc,
              'test_conditions': conds, 'results': results,
              'conclusions': conclusions, 'troubleshooting_index': ts,
              'ai_extraction_log': log}

    tr_en = {
        'document': {'title': doc['title'], 'purpose': doc['purpose'], 'content': doc['content']},
        'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in conclusions},
        'hints': {h_['hint_id']: {'check_item': h_['check_item'], 'reason': h_['reason']} for h_ in ts['suggested_checks']},
        'log': {'assumptions': log['assumptions'], 'warnings': log['warnings'], 'decision_rationale': log['decision_rationale']},
    }
    tr_ko = {
        'document': {'title': 'MSU-L20S15-07 Yoke gauss 점검 보고서',
                     'purpose': 'SMG·CMG low 조건에서 Yoke gauss 점검.',
                     'content': ['Low gauss 샘플 제작.', 'SMG short/long, MG+PT, Final SPK 및 Sigma/Hearing 측정.']},
        'conclusions': {'concl_1': {'topic': 'Low gauss 샘플은 Sigma NG, Hearing OK',
                                     'statement_from_report': 'IV. Decision 비어 있음.',
                                     'normalized_interpretation': 'SMG low #2 (SMG short 492G, spec 900-1350 미만) → Sigma NG. CMG low #1/#2 (MG+PT 304-305G, spec 400-570 미만) → Sigma NG. 네 케이스 모두 Hearing OK. SMG low #1 (982G, spec 내) → Sigma OK 로 spec 임계점이 Sigma NG 의 트리거임을 확인.'}},
        'hints': {'hint_1': {'check_item': 'SMG short 와 MG+PT 의 spec 부합 점검 (SMG 900-1350G, MG+PT 400-570G)',
                              'reason': 'Spec 미달 샘플은 모두 Sigma NG (#2 SMG 492G, #3/#4 MG+PT 304-305G); spec 내 #1 (SMG 982G) 은 Sigma OK.'}},
        'log': {'assumptions': ['Spec 은 시트 헤더 (SMG 900-1350G, MG+PT 400-570G, Final SPK 390-590G) 기준.'],
                 'warnings': ['Final SPK spec 390-590G 인데 #1(1028)·#3(1184)·#4(1169) 초과 — 단위/스케일 차이 가능, 수동 검토 필요.'],
                 'decision_rationale': 'Spec 미달 low gauss 는 모두 Sigma NG 와 1:1 매칭; Hearing 은 모두 OK. Spec 기반 gauss 스크리닝 정당화됨.'},
    }
    tr_vi = {
        'document': {'title': 'Báo cáo kiểm gauss Yoke MSU-L20S15-07',
                     'purpose': 'Kiểm gauss Yoke khi SMG hoặc CMG thấp.',
                     'content': ['Tạo mẫu low gauss.', 'Đo SMG short/long, MG+PT, Final SPK và Sigma/Hearing.']},
        'conclusions': {'concl_1': {'topic': 'Mẫu low gauss cho Sigma NG nhưng Hearing OK',
                                     'statement_from_report': 'Phần IV. Decision của báo cáo trống.',
                                     'normalized_interpretation': 'SMG low #2 (SMG short 492G, dưới spec 900-1350G) → Sigma NG. CMG low #1/#2 (MG+PT 304-305G, dưới spec 400-570G) → Sigma NG. Cả 4 trường hợp Hearing đều OK. SMG low #1 (982G, trong spec) → Sigma OK xác nhận ngưỡng spec là tác nhân.'}},
        'hints': {'hint_1': {'check_item': 'Kiểm SMG short và MG+PT theo spec (SMG 900-1350G, MG+PT 400-570G)',
                              'reason': 'Mẫu ngoài spec đều cho Sigma NG (#2 SMG 492G, #3/#4 MG+PT 304-305G); mẫu trong spec #1 (SMG 982G) cho Sigma OK.'}},
        'log': {'assumptions': ['Spec lấy từ tiêu đề sheet (SMG 900-1350G, MG+PT 400-570G, Final SPK 390-590G).'],
                 'warnings': ['Final SPK spec 390-590G nhưng #1(1028)·#3(1184)·#4(1169) vượt — có thể đơn vị khác, cần review thủ công.'],
                 'decision_rationale': 'Low gauss ngoài spec tương quan 1:1 với Sigma NG; Hearing vẫn OK. Hỗ trợ sàng lọc gauss theo spec để ngăn Sigma NG.'},
    }
    return name, result, tr_ko, tr_en, tr_vi


# -------------------- Dataset 5 --------------------
def ds5():
    name = "30.1. BRS-161014  Report checking NTI test VP mold 7 add 0.3mm 16.4.2024"
    sheet = 'SPL DATA_(NTI'
    doc = base_doc(
        name,
        "Report checking NTI test VP mold #7 add 0.3mm BRS-161014",
        "BRS-161014", "2024-04-16", "", "",
        "reliability_spec",
        "SPL NTI", ["SPL", "NTI"],
        [], ["VP mold #7", "Normal VP"], ["VP forming", "NTI / SPL test"],
        "NTI/SPL test comparing Test VP (mold #7 add 0.3mm) vs Normal VP and Standard reference.",
        ["Mask spec table by Hz (100/200/300/400/500/1000/2000/7000/10000/14000 Hz).",
         "SPL by frequency for Test VP #1-10 AVG, Normal VP #1-10 AVG, and ST AVG."],
        {'title': ['NTI_종합!A1'], 'date': [], 'purpose': [], 'content': []})

    conds = [
        {'condition_id': 'cond_1', 'condition_group': 'VP mold thickness', 'line': '',
         'process': 'VP forming', 'changed_factor': 'VP mold #7 add 0.3mm thickness',
         'before_value': 'Normal VP', 'after_value': 'VP mold #7 +0.3mm',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['SPL DATA_(NTI!header']},
    ]

    results = [
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'SPL',
         'condition_group': 'Test VP AVG @100Hz', 'date': '2024-04-16',
         'metric_name': 'SPL @100Hz', 'metric_value': 71.74456, 'unit': 'dB', 'judgement': 'CHECK',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['SPL DATA_(NTI!Row100Test']},
        {'result_id': 'res_2', 'condition_id': None, 'measurement_type': 'SPL',
         'condition_group': 'Normal VP AVG @100Hz', 'date': '2024-04-16',
         'metric_name': 'SPL @100Hz', 'metric_value': 71.78439, 'unit': 'dB', 'judgement': 'CHECK',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['SPL DATA_(NTI!Row100Normal']},
        {'result_id': 'res_3', 'condition_id': None, 'measurement_type': 'SPL',
         'condition_group': 'Standard AVG @100Hz', 'date': '2024-04-16',
         'metric_name': 'SPL @100Hz', 'metric_value': 72.14835, 'unit': 'dB', 'judgement': None,
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['SPL DATA_(NTI!Row100ST']},
        {'result_id': 'res_4', 'condition_id': 'cond_1', 'measurement_type': 'SPL',
         'condition_group': 'Test VP AVG @1000Hz (interpolated row not in head excerpt)', 'date': '2024-04-16',
         'metric_name': 'SPL @300Hz row1', 'metric_value': 92.00765, 'unit': 'dB', 'judgement': 'CHECK',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['SPL DATA_(NTI!Row280']},
    ]

    conclusions = [
        {'conclusion_id': 'concl_1', 'topic': 'NTI/SPL Test VP mold #7 vs Normal VP',
         'statement_from_report': 'No textual decision present in extracted block; only numeric SPL/NTI table.',
         'normalized_interpretation': 'At 100Hz: Test VP AVG 71.74dB vs Normal VP AVG 71.78dB - effectively equal (-0.04dB). Test VP follows Normal VP closely across the lower frequency range (100-280Hz) within +/-0.2dB; both stay below Standard AVG by ~0.4-0.5dB. No clear NTI deviation in extracted range.',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['SPL DATA_(NTI!100:280']},
    ]

    ts = {'defect_name': 'SPL NTI',
          'when_user_asks': ['VP mold #7 add 0.3mm', 'SPL drift', 'NTI mask exceedance'],
          'suggested_checks': [
              {'hint_id': 'hint_1', 'check_item': 'Compare full-band SPL of Test VP vs Normal vs ST AVG against mask spec (100Hz:40% / 200Hz:30% etc)',
               'reason': 'Mask spec table is loose at low Hz (40% at 100Hz) but tightens at mid-band (5% at 2000Hz, 11% at 7000-10000Hz); extracted data shows Test VP within Normal +/- 0.2dB but full band not extracted.',
               'evidence_strength': 'weak', 'related_process': 'VP forming',
               'related_part': 'VP mold #7', 'source_file': name, 'sheet_name': sheet,
               'source_cells': ['NTI_종합!A1:B11']}],
          'limitations': ['Only 100-280Hz portion of SPL data fits in extract; high-frequency rows truncated.']}

    log = {'confidence': 0.5,
           'assumptions': ['SPL values are AVG of 10 samples per type as labeled in header.'],
           'warnings': ['Extracted text truncated mid-table at ~300Hz; full conclusion needs raw sheet.'],
           'decision_rationale': 'Test VP mold #7 +0.3mm thickness produces SPL very close to Normal VP in the extracted low-mid range; without full-band, cannot conclude NTI pass/fail.'}

    result = {'schema_version': '0.1', 'document': doc,
              'test_conditions': conds, 'results': results,
              'conclusions': conclusions, 'troubleshooting_index': ts,
              'ai_extraction_log': log}

    tr_en = {
        'document': {'title': doc['title'], 'purpose': doc['purpose'], 'content': doc['content']},
        'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in conclusions},
        'hints': {h_['hint_id']: {'check_item': h_['check_item'], 'reason': h_['reason']} for h_ in ts['suggested_checks']},
        'log': {'assumptions': log['assumptions'], 'warnings': log['warnings'], 'decision_rationale': log['decision_rationale']},
    }
    tr_ko = {
        'document': {'title': 'BRS-161014 VP mold #7 add 0.3mm NTI 점검 보고서',
                     'purpose': 'VP mold #7(+0.3mm) vs Normal VP, Standard 기준 NTI/SPL 비교.',
                     'content': ['Hz별 Mask spec (100/200/300/.../14000 Hz).', 'Test VP #1-10 AVG, Normal VP #1-10 AVG, ST AVG SPL.']},
        'conclusions': {'concl_1': {'topic': 'NTI/SPL Test VP mold #7 vs Normal VP',
                                     'statement_from_report': '추출 영역에 결정 텍스트 없음; SPL/NTI 수치표만 존재.',
                                     'normalized_interpretation': '100Hz: Test VP AVG 71.74dB vs Normal VP AVG 71.78dB (-0.04dB) 사실상 동등. 100-280Hz 범위에서 Test VP 가 Normal VP 와 ±0.2dB 이내로 추종; 두 값 모두 ST AVG 보다 약 0.4-0.5dB 낮음. 추출 구간 내 명확한 NTI 편차 없음.'}},
        'hints': {'hint_1': {'check_item': 'Test VP vs Normal vs ST AVG 전대역 SPL 을 Mask spec(100Hz:40%, 200Hz:30%, ...) 과 대조',
                              'reason': 'Mask 는 저주파 느슨(100Hz 40%) 중대역 타이트(2000Hz 5%, 7000-10000Hz 11%); 추출 데이터 상 Test VP 는 Normal ±0.2dB 이내지만 전대역 미확인.'}},
        'log': {'assumptions': ['SPL 값은 헤더 표기대로 type 별 10샘플 AVG.'],
                 'warnings': ['추출 텍스트가 300Hz 부근에서 잘림; 고주파 행 누락.'],
                 'decision_rationale': 'VP mold #7 +0.3mm 는 저-중대역에서 Normal VP 와 매우 근접; 전대역 부재로 NTI pass/fail 단정 불가.'},
    }
    tr_vi = {
        'document': {'title': 'Báo cáo kiểm NTI VP mold #7 +0.3mm BRS-161014',
                     'purpose': 'So sánh NTI/SPL Test VP (mold #7 +0.3mm) vs Normal VP và Standard.',
                     'content': ['Bảng Mask spec theo Hz (100/200/300/.../14000 Hz).', 'SPL theo tần số: Test VP #1-10 AVG, Normal VP #1-10 AVG, ST AVG.']},
        'conclusions': {'concl_1': {'topic': 'NTI/SPL Test VP mold #7 vs Normal VP',
                                     'statement_from_report': 'Phần extract không có quyết định văn bản; chỉ có bảng SPL/NTI.',
                                     'normalized_interpretation': '@100Hz: Test VP AVG 71.74dB vs Normal VP AVG 71.78dB (-0.04dB) gần như bằng nhau. Test VP bám Normal VP trong khoảng 100-280Hz với ±0.2dB; cả hai thấp hơn ST AVG ~0.4-0.5dB. Không thấy chênh lệch NTI rõ trong phần extract.'}},
        'hints': {'hint_1': {'check_item': 'So sánh SPL toàn dải Test VP vs Normal vs ST AVG với Mask spec',
                              'reason': 'Mask lỏng ở tần số thấp (100Hz 40%) chặt ở giữa (2000Hz 5%, 7000-10000Hz 11%); dữ liệu extract cho thấy Test VP gần Normal ±0.2dB nhưng chưa có toàn dải.'}},
        'log': {'assumptions': ['Giá trị SPL là AVG của 10 mẫu mỗi type như header.'],
                 'warnings': ['Text bị cắt ở ~300Hz; thiếu các hàng tần số cao.'],
                 'decision_rationale': 'VP mold #7 +0.3mm gần Normal VP trong dải thấp-trung; thiếu dữ liệu toàn dải để kết luận NTI pass/fail.'},
    }
    return name, result, tr_ko, tr_en, tr_vi


# -------------------- Dataset 6 --------------------
def ds6():
    name = "30.1. MSU-L20S15-07 ( Đo cho ME 28.8.2025)"
    sheet = 'SPL'
    doc = base_doc(
        name,
        "MSU-L20S15-07 SPL measurement for ME 28.8.2025 (Normal vs new vendor GES TRADING vs CD from KR)",
        "MSU-L20S15-07", "2025-08-28", "", "",
        "reliability_spec",
        "SPL NTI", ["SPL", "Material CD"],
        [], ["Center Dome (CD)"], ["SPL measurement"],
        "Measure SPL for ME comparing Normal vs new vendor GES TRADING CD vs CD material from KR.",
        ["Standard vs Normal #1-10 vs NEW VEDER GES TRADING #1-10 vs CD KR #1-10 across frequency (100Hz upward)."],
        {'title': ['SPL!A1'], 'date': [], 'purpose': [], 'content': []})

    conds = [
        {'condition_id': 'cond_1', 'condition_group': 'CD vendor', 'line': '',
         'process': 'CD material', 'changed_factor': 'CD vendor (Normal vs GES TRADING vs KR)',
         'before_value': 'Normal CD', 'after_value': 'GES TRADING / KR',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['SPL!header']},
    ]

    results = [
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'SPL',
         'condition_group': 'Normal AVG @100Hz', 'date': '2025-08-28',
         'metric_name': 'SPL @100Hz Normal AVG', 'metric_value': 79.86401, 'unit': 'dB', 'judgement': 'CHECK',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['SPL!Row100Normal']},
        {'result_id': 'res_2', 'condition_id': 'cond_1', 'measurement_type': 'SPL',
         'condition_group': 'GES TRADING AVG @100Hz', 'date': '2025-08-28',
         'metric_name': 'SPL @100Hz GES AVG', 'metric_value': 79.97823, 'unit': 'dB', 'judgement': 'CHECK',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['SPL!Row100GES']},
        {'result_id': 'res_3', 'condition_id': 'cond_1', 'measurement_type': 'SPL',
         'condition_group': 'CD KR AVG @100Hz', 'date': '2025-08-28',
         'metric_name': 'SPL @100Hz KR AVG', 'metric_value': 80.04898, 'unit': 'dB', 'judgement': 'CHECK',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['SPL!Row100KR']},
        {'result_id': 'res_4', 'condition_id': 'cond_1', 'measurement_type': 'SPL',
         'condition_group': 'Standard @100Hz', 'date': '2025-08-28',
         'metric_name': 'SPL @100Hz Standard', 'metric_value': 79.71099, 'unit': 'dB', 'judgement': None,
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['SPL!Row100Std']},
    ]

    conclusions = [
        {'conclusion_id': 'concl_1', 'topic': 'SPL Normal vs GES TRADING vs KR CD',
         'statement_from_report': 'No textual decision in extract; SPL numeric table only.',
         'normalized_interpretation': '@100Hz: Normal AVG 79.86dB, GES TRADING AVG 79.98dB (+0.12dB vs Normal), KR AVG 80.05dB (+0.19dB vs Normal). Standard 79.71dB. All three vendors track each other within +/-0.2dB at low frequency; both new vendors slightly higher than Normal but within typical sample variance.',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['SPL!Row100']},
    ]

    ts = {'defect_name': 'SPL NTI',
          'when_user_asks': ['CD vendor change', 'SPL difference Normal vs new vendor'],
          'suggested_checks': [
              {'hint_id': 'hint_1', 'check_item': 'Compare full-band SPL average of new vendor CD against Normal CD for ME release approval',
               'reason': 'At 100Hz the difference is small (+0.12 to +0.19dB) but full-band tracking is needed before vendor swap.',
               'evidence_strength': 'weak', 'related_process': 'Center Dome material',
               'related_part': 'CD', 'source_file': name, 'sheet_name': sheet,
               'source_cells': ['SPL!all']}],
          'limitations': ['Only one frequency row inspected in detail; full band needs sheet review.']}

    log = {'confidence': 0.5,
           'assumptions': ['Standard column is reference, last column is AVG per vendor.'],
           'warnings': ['Title only marker is "Đo cho ME" (measure for ME); no formal purpose/conclusion paragraph in extracted sheet.'],
           'decision_rationale': 'New vendor CDs (GES TRADING, KR) produce SPL within 0.2dB of Normal at 100Hz; tentatively similar but needs full-band sign-off.'}

    result = {'schema_version': '0.1', 'document': doc,
              'test_conditions': conds, 'results': results,
              'conclusions': conclusions, 'troubleshooting_index': ts,
              'ai_extraction_log': log}

    tr_en = {
        'document': {'title': doc['title'], 'purpose': doc['purpose'], 'content': doc['content']},
        'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in conclusions},
        'hints': {h_['hint_id']: {'check_item': h_['check_item'], 'reason': h_['reason']} for h_ in ts['suggested_checks']},
        'log': {'assumptions': log['assumptions'], 'warnings': log['warnings'], 'decision_rationale': log['decision_rationale']},
    }
    tr_ko = {
        'document': {'title': 'MSU-L20S15-07 ME 의뢰 SPL 측정 (Normal vs GES TRADING vs KR CD) 28.8.2025',
                     'purpose': 'ME 요청으로 Normal CD 대비 신규 vendor GES TRADING / KR CD 의 SPL 비교 측정.',
                     'content': ['Standard 대비 Normal #1-10, NEW VEDER GES TRADING #1-10, CD KR #1-10 의 주파수별 SPL.']},
        'conclusions': {'concl_1': {'topic': 'SPL Normal vs GES TRADING vs KR CD',
                                     'statement_from_report': '추출 영역에 결정 텍스트 없음; SPL 수치표만.',
                                     'normalized_interpretation': '@100Hz: Normal AVG 79.86dB, GES TRADING AVG 79.98dB (+0.12dB vs Normal), KR AVG 80.05dB (+0.19dB). Standard 79.71dB. 저주파 ±0.2dB 이내 추종; 신규 vendor 가 Normal 보다 미세 상승하나 통상 산포 범위.'}},
        'hints': {'hint_1': {'check_item': 'ME 승인 전 신규 vendor CD 의 전대역 SPL 평균을 Normal 과 비교',
                              'reason': '@100Hz 차이 +0.12~+0.19dB 로 작으나 vendor 변경 전 전대역 검증 필요.'}},
        'log': {'assumptions': ['Standard 열은 reference, 각 vendor 끝열이 AVG.'],
                 'warnings': ['시트에 별도 purpose/conclusion 문단 없음; 제목만 "Đo cho ME".'],
                 'decision_rationale': '신규 vendor CD 는 @100Hz Normal 대비 ±0.2dB 이내; 잠정 동등하나 전대역 sign-off 필요.'},
    }
    tr_vi = {
        'document': {'title': 'MSU-L20S15-07 Đo SPL cho ME 28.8.2025 (Normal vs GES TRADING vs CD KR)',
                     'purpose': 'Đo SPL theo yêu cầu ME, so sánh Normal vs vendor mới GES TRADING vs CD từ KR.',
                     'content': ['Standard, Normal #1-10, NEW VEDER GES TRADING #1-10, CD KR #1-10 theo tần số.']},
        'conclusions': {'concl_1': {'topic': 'SPL Normal vs GES TRADING vs KR CD',
                                     'statement_from_report': 'Extract không có quyết định văn bản; chỉ bảng SPL.',
                                     'normalized_interpretation': '@100Hz: Normal AVG 79.86dB, GES TRADING AVG 79.98dB (+0.12dB so Normal), KR AVG 80.05dB (+0.19dB). Standard 79.71dB. Cả ba vendor bám nhau trong ±0.2dB ở tần số thấp; vendor mới hơi cao hơn Normal nhưng trong dải biến động thông thường.'}},
        'hints': {'hint_1': {'check_item': 'So sánh trung bình SPL toàn dải của CD vendor mới với Normal trước khi ME phê duyệt',
                              'reason': 'Chênh lệch @100Hz +0.12~+0.19dB nhỏ nhưng cần toàn dải trước khi đổi vendor.'}},
        'log': {'assumptions': ['Cột Standard là tham chiếu; cột cuối mỗi vendor là AVG.'],
                 'warnings': ['Tiêu đề chỉ ghi "Đo cho ME"; không có purpose/conclusion chính thức.'],
                 'decision_rationale': 'CD vendor mới (GES TRADING, KR) cho SPL trong ±0.2dB so Normal @100Hz; tạm tương đương nhưng cần sign-off toàn dải.'},
    }
    return name, result, tr_ko, tr_en, tr_vi


# -------------------- Dataset 7 --------------------
def ds7():
    name = "30.2 BRS-161016  Report test 2nd  VP mold #7 add 0.3mm  date 25.4.2024"
    sheet = '25.4'
    doc = base_doc(
        name,
        "Report test 2nd VP mold #7 improve by add 0.3mm of thickness BRS-161016",
        "BRS-161016", "2024-04-25", "Thuy", "E2-3A",
        "normal_comparison",
        "NG Function", ["NG Vision VP", "VP damage", "VP bending", "NG Hearing Noise"],
        ["NG Vision VP", "VP damage", "NG Function"],
        ["VP", "VP mold #7", "VP mold #8"], ["VP laser cutting", "VP bending", "VP+CD assembly", "Function test"],
        "Test 2nd VP mold #7 with +0.3mm thickness to verify improvement vs Normal VP mold #8.",
        ["Check material and semi VP after laser cutting for VP bending.",
         "Make semi and check NG rate vision VP+CD.",
         "Make final and check NG rate of function.",
         "Check NTI on OK sample."],
        {'title': ['25.4!A1'], 'date': [], 'purpose': [], 'content': []})

    conds = [
        {'condition_id': 'cond_1', 'condition_group': 'VP mold #7 +0.3mm', 'line': 'E2-3A',
         'process': 'VP laser cutting / forming', 'changed_factor': 'VP mold #7 thickness +0.3mm both sides',
         'before_value': 'Normal VP mold #8', 'after_value': 'VP mold #7 +0.3mm',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['25.4!cond']},
    ]

    results = [
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Vision',
         'condition_group': 'Test VP mold #7 laser cutting', 'date': '2024-04-24', 'line': 'E2-3A',
         'input_count': 1970, 'ok_count': 1950, 'ng_count': 20, 'ng_rate_decimal': 0.01, 'ng_rate_percent': 1.0,
         'metric_name': 'NG Rate VP laser cutting', 'metric_value': 1.0, 'unit': '%',
         'ng_breakdown': {'Particle': 3, 'VP damage': 16, 'Burr': 1},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['25.4!264']},
        {'result_id': 'res_2', 'condition_id': None, 'measurement_type': 'Vision',
         'condition_group': 'Normal VP mold #8 laser cutting', 'date': '2024-04-24', 'line': 'E2-3A',
         'input_count': 1000, 'ok_count': 1000, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'NG Rate VP laser cutting', 'metric_value': 0.0, 'unit': '%',
         'ng_breakdown': {'Particle': 0, 'VP damage': 0, 'Burr': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['25.4!265']},
        {'result_id': 'res_3', 'condition_id': 'cond_1', 'measurement_type': 'VP Bending',
         'condition_group': 'Test VP mold #7 bending', 'date': '2024-04-24', 'line': 'E2-3A',
         'input_count': 1950, 'ok_count': 1950, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'NG Rate VP bending', 'metric_value': 0.0, 'unit': '%',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['25.4!268']},
        {'result_id': 'res_4', 'condition_id': None, 'measurement_type': 'VP Bending',
         'condition_group': 'Normal VP mold #8 bending', 'date': '2024-04-24', 'line': 'E2-3A',
         'input_count': 1000, 'ok_count': 1000, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'NG Rate VP bending', 'metric_value': 0.0, 'unit': '%',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['25.4!269']},
        {'result_id': 'res_5', 'condition_id': 'cond_1', 'measurement_type': 'Vision',
         'condition_group': 'Test VP #7 VP/CD vision', 'date': '2024-04-24', 'line': 'E2-3A',
         'input_count': 1950, 'ok_count': 1945, 'ng_count': 5, 'ng_rate_decimal': 0.003, 'ng_rate_percent': 0.3,
         'metric_name': 'NG Rate VP/CD vision', 'metric_value': 0.3, 'unit': '%',
         'ng_breakdown': {'Particle': 0, 'Glue not enough': 4, 'Dome damage': 1},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['25.4!273']},
        {'result_id': 'res_6', 'condition_id': None, 'measurement_type': 'Vision',
         'condition_group': 'Normal VP mold #8 VP/CD vision', 'date': '2024-04-24', 'line': 'E2-3A',
         'input_count': 1000, 'ok_count': 995, 'ng_count': 5, 'ng_rate_decimal': 0.005, 'ng_rate_percent': 0.5,
         'metric_name': 'NG Rate VP/CD vision', 'metric_value': 0.5, 'unit': '%',
         'ng_breakdown': {'Particle': 0, 'Glue not enough': 4, 'Dome damage': 1},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['25.4!275']},
        {'result_id': 'res_7', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Test VP mold #7 function', 'date': '2024-04-25', 'line': 'E2-3A',
         'input_count': 1937, 'ok_count': 1907, 'ng_count': 30, 'ng_rate_decimal': 0.015, 'ng_rate_percent': 1.5,
         'metric_name': 'Total NG Rate function', 'metric_value': 1.5, 'unit': '%',
         'ng_breakdown': {'NG SPL': 0, 'NG THD': 0, 'NG SPL+THD': 1, 'NG SPL+THD+F0': 0, 'NG Hearing Noise': 29, 'NG Hearing Touch': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['25.4!280']},
        {'result_id': 'res_8', 'condition_id': None, 'measurement_type': 'Function',
         'condition_group': 'Normal VP mold #8 function', 'date': '2024-04-25', 'line': 'E2-3A',
         'input_count': 999, 'ok_count': 968, 'ng_count': 31, 'ng_rate_decimal': 0.031, 'ng_rate_percent': 3.1,
         'metric_name': 'Total NG Rate function', 'metric_value': 3.1, 'unit': '%',
         'ng_breakdown': {'NG SPL': 0, 'NG THD': 0, 'NG SPL+THD': 0, 'NG SPL+THD+F0': 0, 'NG Hearing Noise': 31, 'NG Hearing Touch': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['25.4!282']},
    ]

    # Vision laser cut: 1.0/0.0 inf - test increases damage; Function: 1.5/3.1-1 = -51.6% improved
    conclusions = [
        {'conclusion_id': 'concl_1', 'topic': 'VP mold #7 +0.3mm 2nd trial vs Normal',
         'statement_from_report': "VP mold #7 add 0.3mm don't have NG bending but happen NG VP damage 0.8% rate high => Can use but need to improve VP damage from supplier.",
         'normalized_interpretation': "Laser cut: Test 1.0% (VP damage 0.8%) vs Normal 0.0% — Test worse on VP damage. Bending: both 0.0%. VP/CD vision: Test 0.3% vs Normal 0.5% = -40% improved. Function: Test 1.5% vs Normal 3.1% = -51.6% improved (NG Hearing Noise dominant in both). Net: function and vision improved but VP damage from laser cutting got worse — supplier improvement needed.",
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['25.4!264:285']},
    ]

    ts = {'defect_name': 'VP damage',
          'when_user_asks': ['VP mold #7 +0.3mm', 'VP damage after laser cutting', 'VP bending'],
          'suggested_checks': [
              {'hint_id': 'hint_1', 'check_item': 'Request supplier improvement on VP damage (mold #7 +0.3mm thickness)',
               'reason': 'Test VP mold #7 generates VP damage 0.8% (16/1950) vs Normal 0% at laser cutting; report explicitly recommends supplier improvement.',
               'evidence_strength': 'strong', 'related_process': 'VP laser cutting',
               'related_part': 'VP mold #7', 'source_file': name, 'sheet_name': sheet,
               'source_cells': ['25.4!264:265', '25.4!285']}],
          'limitations': ['Only one date trial per condition.']}

    log = {'confidence': 0.85,
           'assumptions': ['Normal VP mold #8 treated as same-event baseline for each measurement_type.'],
           'warnings': [],
           'decision_rationale': 'VP mold #7 +0.3mm shows positive trade-off: function NG -51.6% improved vs Normal but laser-cut VP damage 0.8% appears (Normal 0%). Report decision keeps it usable conditional on supplier VP-damage improvement.'}

    result = {'schema_version': '0.1', 'document': doc,
              'test_conditions': conds, 'results': results,
              'conclusions': conclusions, 'troubleshooting_index': ts,
              'ai_extraction_log': log}

    tr_en = {
        'document': {'title': doc['title'], 'purpose': doc['purpose'], 'content': doc['content']},
        'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in conclusions},
        'hints': {h_['hint_id']: {'check_item': h_['check_item'], 'reason': h_['reason']} for h_ in ts['suggested_checks']},
        'log': {'assumptions': log['assumptions'], 'warnings': log['warnings'], 'decision_rationale': log['decision_rationale']},
    }
    tr_ko = {
        'document': {'title': 'BRS-161016 VP mold #7 +0.3mm 2차 시험 보고서 (25.4.2024)',
                     'purpose': 'VP mold #7 두께 +0.3mm 적용으로 Normal VP mold #8 대비 개선 여부 검증.',
                     'content': ['Laser cutting 후 VP bending 점검.', 'Semi 단계 VP+CD vision NG rate.', 'Final 기능 NG rate.', 'OK 샘플 NTI 점검.']},
        'conclusions': {'concl_1': {'topic': 'VP mold #7 +0.3mm 2차 vs Normal',
                                     'statement_from_report': 'VP mold #7 +0.3mm 는 NG bending 없으나 VP damage 0.8% 발생 → 사용 가능하지만 공급사 VP damage 개선 필요.',
                                     'normalized_interpretation': 'Laser cut: Test 1.0% (VP damage 0.8%) vs Normal 0.0% — Test 가 VP damage 에서 악화. Bending 양쪽 0.0%. VP/CD vision: Test 0.3% vs Normal 0.5% = -40% 개선. Function: Test 1.5% vs Normal 3.1% = -51.6% 개선 (NG Hearing Noise 주도). 종합: function·vision 개선 vs Laser-cut VP damage 악화 — 공급사 개선 필요.'}},
        'hints': {'hint_1': {'check_item': '공급사에 VP damage 개선 요청 (mold #7 +0.3mm)',
                              'reason': 'Test VP mold #7 의 laser cutting VP damage 0.8% (16/1950) vs Normal 0%; 리포트가 공급사 개선을 명시.'}},
        'log': {'assumptions': ['각 measurement_type 별 Normal VP mold #8 행을 동일 이벤트 baseline 으로 사용.'],
                 'warnings': [],
                 'decision_rationale': 'VP mold #7 +0.3mm 는 function NG -51.6% 개선이지만 laser-cut VP damage 0.8% 발생. 공급사 VP damage 개선 조건부 사용 가능.'},
    }
    tr_vi = {
        'document': {'title': 'Báo cáo test lần 2 VP mold #7 +0.3mm BRS-161016 (25.4.2024)',
                     'purpose': 'Test VP mold #7 +0.3mm để xác minh cải thiện so với Normal VP mold #8.',
                     'content': ['Kiểm VP bending sau laser cutting.', 'NG rate VP+CD vision ở bán thành phẩm.', 'NG rate function ở thành phẩm.', 'Kiểm NTI mẫu OK.']},
        'conclusions': {'concl_1': {'topic': 'VP mold #7 +0.3mm lần 2 vs Normal',
                                     'statement_from_report': 'VP mold #7 +0.3mm không có NG bending nhưng có VP damage 0.8% → có thể dùng nhưng cần cải thiện VP damage từ nhà cung cấp.',
                                     'normalized_interpretation': 'Laser cut: Test 1.0% (VP damage 0.8%) vs Normal 0.0% — Test xấu hơn ở VP damage. Bending cả hai 0.0%. VP/CD vision: Test 0.3% vs Normal 0.5% = -40% cải thiện. Function: Test 1.5% vs Normal 3.1% = -51.6% cải thiện (NG Hearing Noise chủ đạo). Tổng kết: function·vision cải thiện nhưng VP damage laser cut xấu hơn — cần cải thiện từ supplier.'}},
        'hints': {'hint_1': {'check_item': 'Yêu cầu supplier cải thiện VP damage (mold #7 +0.3mm)',
                              'reason': 'Test VP mold #7 có VP damage 0.8% (16/1950) sau laser cutting vs Normal 0%; báo cáo nêu cần supplier cải thiện.'}},
        'log': {'assumptions': ['Hàng Normal VP mold #8 dùng làm baseline cùng sự kiện cho từng measurement_type.'],
                 'warnings': [],
                 'decision_rationale': 'VP mold #7 +0.3mm có function NG -51.6% cải thiện nhưng VP damage laser cut 0.8% phát sinh. Báo cáo cho dùng có điều kiện kèm cải thiện supplier.'},
    }
    return name, result, tr_ko, tr_en, tr_vi


# -------------------- Dataset 8 --------------------
def ds8():
    name = "30.2.MSU-20S15-07 Result check NTI test new material center Dome from GES TRADING and KR"
    sheet = '23.8'
    doc = base_doc(
        name,
        "Report test material center dome of new vender GES TRADING and KR MSU-L20L15-07",
        "MSU-L20L15-07", "2025-08-23", "Thao", "E2-4B",
        "mixed",
        "VP+CD Separation", ["NG separate VP+CD", "VP CD separate"],
        ["VP+CD Separation", "NG Function", "NG Hearing Noise"],
        ["Center Dome (CD)", "VP+CD"], ["VP+CD ass'y Sub1", "Tension", "Function", "NTI", "Reliability"],
        "Improve VP+CD separate by testing new center dome vendors GES TRADING and from KR.",
        ["Test material Center Dome new vendor GES TRADING and from KR.",
         "Make sample and check VP+CD assembly NG.",
         "Check tension; make final sample and check function.",
         "Check NTI and Reliability at A1."],
        {'title': ['23.8!A1'], 'date': [], 'purpose': [], 'content': []})

    conds = [
        {'condition_id': 'cond_1', 'condition_group': 'CD vendor + clean/laser combo', 'line': 'E2-4B',
         'process': "VP+CD ass'y Sub1", 'changed_factor': 'CD material vendor (GES TRADING vs KR vs Normal Ralon) and clean/laser combo',
         'before_value': 'Normal Ralon CD', 'after_value': 'GES TRADING / KR / clean+laser variants',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['23.8!cond']},
    ]

    results = [
        # Sub1 VP+CD separation - all 0% (no NG observed at sub-process)
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Vision',
         'condition_group': 'GES TRADING CD VP+CD separate (22-Aug)', 'date': '2025-08-22', 'line': 'E2-4B',
         'input_count': 133, 'ok_count': 133, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'VP+CD Separate NG Rate', 'metric_value': 0.0, 'unit': '%',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['23.8!331']},
        {'result_id': 'res_2', 'condition_id': 'cond_1', 'measurement_type': 'Vision',
         'condition_group': 'KR CD VP+CD separate (22-Aug)', 'date': '2025-08-22', 'line': 'E2-4B',
         'input_count': 251, 'ok_count': 251, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'VP+CD Separate NG Rate', 'metric_value': 0.0, 'unit': '%',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['23.8!332']},
        {'result_id': 'res_3', 'condition_id': None, 'measurement_type': 'Vision',
         'condition_group': 'Normal Ralon laser baseline (22-Aug)', 'date': '2025-08-22', 'line': 'E2-4B',
         'input_count': 100, 'ok_count': 100, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'VP+CD Separate NG Rate', 'metric_value': 0.0, 'unit': '%',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['23.8!333']},
        # Tension (spec >= 1.2Kgf)
        {'result_id': 'res_4', 'condition_id': 'cond_1', 'measurement_type': 'Tension',
         'condition_group': 'GES TRADING tension (23-Aug)', 'date': '2025-08-23', 'line': 'E2-4B',
         'metric_name': 'Tension AVG', 'metric_value': 2.47, 'unit': 'Kgf', 'judgement': 'PASS',
         'ng_breakdown': {'min': 2.20, 'max': 2.82, 'spec_min': 1.2},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['23.8!341']},
        {'result_id': 'res_5', 'condition_id': 'cond_1', 'measurement_type': 'Tension',
         'condition_group': 'KR tension (23-Aug)', 'date': '2025-08-23', 'line': 'E2-4B',
         'metric_name': 'Tension AVG', 'metric_value': 2.92, 'unit': 'Kgf', 'judgement': 'PASS',
         'ng_breakdown': {'min': 2.50, 'max': 3.26, 'spec_min': 1.2},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['23.8!342']},
        {'result_id': 'res_6', 'condition_id': None, 'measurement_type': 'Tension',
         'condition_group': 'Normal Ralon laser tension (23-Aug)', 'date': '2025-08-23', 'line': 'E2-4B',
         'metric_name': 'Tension AVG', 'metric_value': 2.30, 'unit': 'Kgf', 'judgement': 'PASS',
         'ng_breakdown': {'min': 2.04, 'max': 2.67, 'spec_min': 1.2},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['23.8!343']},
        # Function
        {'result_id': 'res_7', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'GES TRADING function (23-Aug)', 'date': '2025-08-23', 'line': 'E2-4B',
         'input_count': 133, 'ok_count': 129, 'ng_count': 4, 'ng_rate_decimal': 0.03, 'ng_rate_percent': 3.0,
         'metric_name': 'Total NG Rate function', 'metric_value': 3.0, 'unit': '%',
         'ng_breakdown': {'NG SPL': 0, 'NG THD': 0, 'NG SPL+THD': 0, 'NG SPL+THD+F0': 0, 'NG Hearing Noise': 3, 'NG Hearing Touch': 1, 'NG VP+CD Separate (decap)': 4},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['23.8!351']},
        {'result_id': 'res_8', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'KR function (23-Aug)', 'date': '2025-08-23', 'line': 'E2-4B',
         'input_count': 251, 'ok_count': 218, 'ng_count': 33, 'ng_rate_decimal': 0.131, 'ng_rate_percent': 13.1,
         'metric_name': 'Total NG Rate function', 'metric_value': 13.1, 'unit': '%',
         'ng_breakdown': {'NG SPL': 30, 'NG THD': 0, 'NG SPL+THD': 0, 'NG SPL+THD+F0': 0, 'NG Hearing Noise': 0, 'NG Hearing Touch': 3, 'NG VP+CD Separate (decap)': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['23.8!353']},
        {'result_id': 'res_9', 'condition_id': None, 'measurement_type': 'Function',
         'condition_group': 'Normal Ralon laser function (23-Aug)', 'date': '2025-08-23', 'line': 'E2-4B',
         'input_count': 1120, 'ok_count': 1095, 'ng_count': 25, 'ng_rate_decimal': 0.022, 'ng_rate_percent': 2.2,
         'metric_name': 'Total NG Rate function', 'metric_value': 2.2, 'unit': '%',
         'ng_breakdown': {'NG SPL': 0, 'NG THD': 0, 'NG SPL+THD': 0, 'NG SPL+THD+F0': 0, 'NG Hearing Noise': 15, 'NG Hearing Touch': 10},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['23.8!355']},
        # 25-Aug GES TRADING variants
        {'result_id': 'res_10', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'GES TRADING Clean MS-200VA+No laser function (25-Aug)', 'date': '2025-08-25', 'line': 'E2-4B',
         'input_count': 87, 'ok_count': 85, 'ng_count': 2, 'ng_rate_decimal': 0.023, 'ng_rate_percent': 2.3,
         'metric_name': 'Total NG Rate function', 'metric_value': 2.3, 'unit': '%',
         'ng_breakdown': {'NG Hearing Noise': 2, 'NG VP+CD Separate (decap)': 1},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['23.8!357']},
        {'result_id': 'res_11', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'GES TRADING No Clean+laser function (25-Aug)', 'date': '2025-08-25', 'line': 'E2-4B',
         'input_count': 134, 'ok_count': 132, 'ng_count': 2, 'ng_rate_decimal': 0.015, 'ng_rate_percent': 1.5,
         'metric_name': 'Total NG Rate function', 'metric_value': 1.5, 'unit': '%',
         'ng_breakdown': {'NG Hearing Noise': 2},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['23.8!359']},
        {'result_id': 'res_12', 'condition_id': None, 'measurement_type': 'Function',
         'condition_group': 'Normal Ralon Clean MS-200VA+No laser function (25-Aug)', 'date': '2025-08-25', 'line': 'E2-4B',
         'input_count': 135, 'ok_count': 135, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Total NG Rate function', 'metric_value': 0.0, 'unit': '%',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['23.8!361']},
        {'result_id': 'res_13', 'condition_id': None, 'measurement_type': 'Function',
         'condition_group': 'Normal No Clean+laser function (25-Aug)', 'date': '2025-08-25', 'line': 'E2-4B',
         'input_count': 137, 'ok_count': 137, 'ng_count': 0, 'ng_rate_decimal': 0.0, 'ng_rate_percent': 0.0,
         'metric_name': 'Total NG Rate function', 'metric_value': 0.0, 'unit': '%',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['23.8!363']},
    ]

    # rel:
    # 23-Aug GES 3.0/2.2-1 = +36.4% worse; KR 13.1/2.2-1 = +495.5% worse
    # 25-Aug Normal Ralon 0.0 makes rel undefined; use absolute comparison
    conclusions = [
        {'conclusion_id': 'concl_1', 'topic': 'GES TRADING and KR CD vs Normal Ralon CD',
         'statement_from_report': 'NG VP+CD separate observed on GES TRADING (4pcs decap on 23-Aug; 1 decap on 25-Aug Clean+No laser). KR shows NG SPL 12.0% (30/251) with no VP+CD separate after decap.',
         'normalized_interpretation': "23-Aug function: GES TRADING 3.0% vs Normal Ralon 2.2% = +36.4% worse (driven by NG Hearing Noise 2.3% and decap-confirmed VP+CD separate 4pcs). KR 13.1% vs Normal Ralon 2.2% = +495.5% worse, dominated by NG SPL 12.0% (likely SPL-side material issue, not VP+CD separate). 25-Aug function: GES TRADING with Clean+No laser 2.3% vs Normal Clean+No laser 0.0% — still worse. Tension all PASS (GES 2.47, KR 2.92, Normal 2.30 Kgf vs spec >=1.2). VP+CD vision sub-process all 0% across vendors.",
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['23.8!331:363']},
    ]

    ts = {'defect_name': 'VP+CD Separation',
          'when_user_asks': ['Center Dome vendor change', 'VP+CD separate', 'NG SPL after CD change'],
          'suggested_checks': [
              {'hint_id': 'hint_1', 'check_item': 'Check NG SPL when using KR CD material (12.0%, much higher than Normal Ralon 0%)',
               'reason': 'KR CD function NG 13.1% vs Normal 2.2% = +495.5% worse; NG SPL alone is 12.0% (30/251). Indicates SPL-side material problem with KR center dome.',
               'evidence_strength': 'strong', 'related_process': "VP+CD ass'y Sub1",
               'related_part': 'CD material (KR vendor)', 'source_file': name, 'sheet_name': sheet,
               'source_cells': ['23.8!353:354']},
              {'hint_id': 'hint_2', 'check_item': 'Check VP+CD separation after decap when using GES TRADING CD',
               'reason': 'GES TRADING shows NG VP+CD separate 4pcs in decap (23-Aug) and 1/2pcs (25-Aug clean+no laser). Sub1 vision is 0% so issue is detectable only after decap.',
               'evidence_strength': 'medium', 'related_process': "VP+CD ass'y Sub1",
               'related_part': 'CD material (GES TRADING vendor)', 'source_file': name, 'sheet_name': sheet,
               'source_cells': ['23.8!351', '23.8!357']}],
          'limitations': ["VP+CD separate at Sub1 vision is 0% across all vendors; only post-decap reveals issue. Need decap on more samples."]}

    log = {'confidence': 0.8,
           'assumptions': ['Normal Ralon laser row treated as same-event baseline for 23-Aug; Normal Clean+No laser for 25-Aug; Normal No Clean+laser for 25-Aug laser variant.'],
           'warnings': ['25-Aug Normal rows are 0.0%, so relative_change is undefined — switched to absolute comparison for that sub-event.'],
           'decision_rationale': 'Both new vendors are worse than Normal Ralon in function: KR driven by NG SPL (likely SPL acoustic issue), GES TRADING driven by NG Hearing Noise + post-decap VP+CD separate. Tension is fine. Neither vendor cleanly improves VP+CD separate vs Normal.'}

    result = {'schema_version': '0.1', 'document': doc,
              'test_conditions': conds, 'results': results,
              'conclusions': conclusions, 'troubleshooting_index': ts,
              'ai_extraction_log': log}

    tr_en = {
        'document': {'title': doc['title'], 'purpose': doc['purpose'], 'content': doc['content']},
        'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in conclusions},
        'hints': {h_['hint_id']: {'check_item': h_['check_item'], 'reason': h_['reason']} for h_ in ts['suggested_checks']},
        'log': {'assumptions': log['assumptions'], 'warnings': log['warnings'], 'decision_rationale': log['decision_rationale']},
    }
    tr_ko = {
        'document': {'title': 'MSU-L20L15-07 신규 vendor GES TRADING·KR Center Dome NTI 시험 보고서',
                     'purpose': 'VP+CD separate 개선 위해 신규 CD vendor GES TRADING·KR 평가.',
                     'content': ['신규 vendor GES TRADING·KR CD 자재 시험.', 'Sample 제작 후 VP+CD ass\'y NG 점검.', 'Tension 측정 및 최종 sample 기능 시험.', 'A1 에서 NTI·Reliability 점검.']},
        'conclusions': {'concl_1': {'topic': 'GES TRADING·KR CD vs Normal Ralon CD',
                                     'statement_from_report': 'GES TRADING: decap 시 VP+CD separate 23-Aug 4pcs, 25-Aug Clean+No laser 1pcs. KR: NG SPL 12.0% (30/251), decap 시 separate 없음.',
                                     'normalized_interpretation': "23-Aug 기능: GES TRADING 3.0% vs Normal Ralon 2.2% = +36.4% 악화 (NG Hearing Noise 2.3% + decap-검출 VP+CD separate 4pcs). KR 13.1% vs Normal Ralon 2.2% = +495.5% 악화, NG SPL 12.0% 주도 (SPL 측 자재 이슈, VP+CD separate 아님). 25-Aug 기능: GES TRADING Clean+No laser 2.3% vs Normal Clean+No laser 0.0% — 여전히 악화. Tension 전 항목 PASS (GES 2.47, KR 2.92, Normal 2.30 Kgf, spec ≥1.2). Sub1 VP+CD separate 0%."}},
        'hints': {'hint_1': {'check_item': 'KR CD 사용 시 NG SPL 점검 (12.0%, Normal Ralon 0% 대비 매우 높음)',
                              'reason': 'KR CD 기능 NG 13.1% vs Normal 2.2% = +495.5% 악화; NG SPL 단독으로 12.0% (30/251). KR center dome 의 SPL 측 자재 문제 추정.'},
                  'hint_2': {'check_item': 'GES TRADING CD 사용 시 decap 후 VP+CD separate 점검',
                              'reason': 'GES TRADING 은 decap 후 VP+CD separate 가 4pcs(23-Aug), 1/2pcs(25-Aug clean+no laser) 발생. Sub1 vision 은 0%로 decap 해야 검출.'}},
        'log': {'assumptions': ['23-Aug 은 Normal Ralon laser, 25-Aug clean+no laser 변형은 Normal Clean+No laser, 25-Aug laser 변형은 Normal No Clean+laser 를 baseline 으로 사용.'],
                 'warnings': ['25-Aug Normal 행이 0.0% 라 상대 변화율 정의 불가 — 절대 비교로 대체.'],
                 'decision_rationale': '신규 두 vendor 모두 기능에서 Normal Ralon 보다 악화: KR 은 NG SPL, GES TRADING 은 NG Hearing Noise + decap VP+CD separate. Tension 은 문제 없음. 어느 한 vendor 도 VP+CD separate 를 명확히 개선하지 못함.'},
    }
    tr_vi = {
        'document': {'title': 'Báo cáo test material Center Dome vendor mới GES TRADING và KR MSU-L20L15-07',
                     'purpose': 'Cải thiện VP+CD separate bằng việc test vendor CD mới GES TRADING và KR.',
                     'content': ['Test material Center Dome vendor mới GES TRADING và từ KR.', "Tạo mẫu và kiểm VP+CD ass'y NG.", 'Đo tension, làm mẫu cuối và test function.', 'Kiểm NTI và Reliability tại A1.']},
        'conclusions': {'concl_1': {'topic': 'GES TRADING và KR CD vs Normal Ralon CD',
                                     'statement_from_report': 'GES TRADING: decap thấy VP+CD separate 4pcs (23-Aug), 1pcs (25-Aug Clean+No laser). KR: NG SPL 12.0% (30/251), không có separate sau decap.',
                                     'normalized_interpretation': "23-Aug function: GES TRADING 3.0% vs Normal Ralon 2.2% = +36.4% xấu hơn (NG Hearing Noise 2.3% + VP+CD separate 4pcs sau decap). KR 13.1% vs Normal Ralon 2.2% = +495.5% xấu hơn, do NG SPL 12.0% (vấn đề material phía SPL, không phải VP+CD separate). 25-Aug function: GES TRADING Clean+No laser 2.3% vs Normal Clean+No laser 0.0% — vẫn xấu hơn. Tension đều PASS (GES 2.47, KR 2.92, Normal 2.30 Kgf, spec ≥1.2). Sub1 vision VP+CD separate đều 0%."}},
        'hints': {'hint_1': {'check_item': 'Kiểm NG SPL khi dùng CD vendor KR (12.0%, cao hơn nhiều Normal Ralon 0%)',
                              'reason': 'KR CD function NG 13.1% vs Normal 2.2% = +495.5% xấu hơn; NG SPL riêng 12.0% (30/251). Cho thấy vấn đề material phía SPL của KR center dome.'},
                  'hint_2': {'check_item': 'Kiểm VP+CD separate sau decap khi dùng GES TRADING CD',
                              'reason': 'GES TRADING có VP+CD separate 4pcs sau decap (23-Aug) và 1/2pcs (25-Aug clean+no laser). Sub1 vision 0% nên chỉ thấy được sau decap.'}},
        'log': {'assumptions': ['23-Aug dùng Normal Ralon laser, 25-Aug variant clean+no laser dùng Normal Clean+No laser, 25-Aug variant laser dùng Normal No Clean+laser làm baseline.'],
                 'warnings': ['Hàng Normal 25-Aug bằng 0.0% nên không tính được tỷ lệ tương đối — chuyển sang so sánh tuyệt đối.'],
                 'decision_rationale': 'Cả hai vendor mới đều xấu hơn Normal Ralon ở function: KR do NG SPL, GES TRADING do NG Hearing Noise + VP+CD separate sau decap. Tension đều ổn. Không vendor nào cải thiện rõ ràng VP+CD separate.'},
    }
    return name, result, tr_ko, tr_en, tr_vi


# -------------------- Dataset 9 --------------------
def ds9():
    name = "30.3 BRS-161016  Report test 3rd  VP mold #7 add 0.3mm  date 13.5.2024"
    sheet = '13.5'
    doc = base_doc(
        name,
        "Report test 3rd VP mold #7 improve by add 0.3mm of thickness BRS-161016",
        "BRS-161016", "2024-05-13", "Thuy", "E2-3A",
        "normal_comparison",
        "VP damage", ["VP bending", "Burr", "Particle", "Dome damage", "NG Hearing Noise"],
        ["VP damage", "VP bending", "NG Function"],
        ["VP mold #7", "VP mold #2"], ["VP laser cutting", "VP bending", "VP+CD vision", "Function"],
        "Test 3rd VP mold #7 with +0.3mm thickness to verify improvement vs Normal VP mold #2.",
        ["Check material and semi VP after laser cutting for VP bending.",
         "Make semi and check NG rate vision VP+CD.",
         "Make final and check NG rate of function."],
        {'title': ['13.5!A1'], 'date': [], 'purpose': [], 'content': []})

    conds = [
        {'condition_id': 'cond_1', 'condition_group': 'VP mold #7 +0.3mm (3rd trial)', 'line': 'E2-3A',
         'process': 'VP laser cutting / forming', 'changed_factor': 'VP mold #7 thickness +0.3mm both sides',
         'before_value': 'Normal VP mold #2', 'after_value': 'VP mold #7 +0.3mm',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['13.5!cond']},
    ]

    results = [
        {'result_id': 'res_1', 'condition_id': 'cond_1', 'measurement_type': 'Vision',
         'condition_group': 'Test VP mold #7 laser cutting (13-May)', 'date': '2024-05-13', 'line': 'E2-3A',
         'input_count': 9960, 'ok_count': 9928, 'ng_count': 32, 'ng_rate_decimal': 0.003, 'ng_rate_percent': 0.3,
         'metric_name': 'NG Rate VP laser cutting', 'metric_value': 0.3, 'unit': '%',
         'ng_breakdown': {'Particle': 4, 'VP damage': 10, 'Burr': 18},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['13.5!402']},
        {'result_id': 'res_2', 'condition_id': None, 'measurement_type': 'Vision',
         'condition_group': 'Normal VP mold #2 laser cutting (13-May)', 'date': '2024-05-13', 'line': 'E2-3A',
         'input_count': 10620, 'ok_count': 10612, 'ng_count': 8, 'ng_rate_decimal': 0.00075, 'ng_rate_percent': 0.075,
         'metric_name': 'NG Rate VP laser cutting', 'metric_value': 0.075, 'unit': '%',
         'ng_breakdown': {'Particle': 5, 'VP damage': 2, 'Burr': 1},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['13.5!403']},
        {'result_id': 'res_3', 'condition_id': 'cond_1', 'measurement_type': 'VP Bending',
         'condition_group': 'Test VP mold #7 bending (13-May)', 'date': '2024-05-13', 'line': 'E2-3A',
         'input_count': 9928, 'ok_count': 9813, 'ng_count': 115, 'ng_rate_decimal': 0.012, 'ng_rate_percent': 1.2,
         'metric_name': 'NG Rate VP bending', 'metric_value': 1.2, 'unit': '%',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['13.5!406']},
        {'result_id': 'res_4', 'condition_id': None, 'measurement_type': 'VP Bending',
         'condition_group': 'Normal VP mold #2 bending (13-May)', 'date': '2024-05-13', 'line': 'E2-3A',
         'input_count': 10612, 'ok_count': 10497, 'ng_count': 115, 'ng_rate_decimal': 0.011, 'ng_rate_percent': 1.1,
         'metric_name': 'NG Rate VP bending', 'metric_value': 1.1, 'unit': '%',
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['13.5!407']},
        {'result_id': 'res_5', 'condition_id': 'cond_1', 'measurement_type': 'Vision',
         'condition_group': 'Test VP #7 VP/CD vision (13-May)', 'date': '2024-05-13', 'line': 'E2-3A',
         'input_count': 9813, 'ok_count': 9788, 'ng_count': 25, 'ng_rate_decimal': 0.003, 'ng_rate_percent': 0.3,
         'metric_name': 'NG Rate VP/CD vision', 'metric_value': 0.3, 'unit': '%',
         'ng_breakdown': {'Particle': 1, 'Glue not enough': 16, 'Dome damage': 8, 'VP damage': 1},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['13.5!411']},
        {'result_id': 'res_6', 'condition_id': None, 'measurement_type': 'Vision',
         'condition_group': 'Normal VP mold #2 VP/CD vision (13-May)', 'date': '2024-05-13', 'line': 'E2-3A',
         'input_count': 10497, 'ok_count': 10476, 'ng_count': 21, 'ng_rate_decimal': 0.002, 'ng_rate_percent': 0.2,
         'metric_name': 'NG Rate VP/CD vision', 'metric_value': 0.2, 'unit': '%',
         'ng_breakdown': {'Particle': 2, 'Glue not enough': 13, 'Dome damage': 6, 'VP damage': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['13.5!413']},
        {'result_id': 'res_7', 'condition_id': 'cond_1', 'measurement_type': 'Function',
         'condition_group': 'Test VP mold #7 function (14-May)', 'date': '2024-05-14', 'line': 'E2-3A',
         'input_count': 9684, 'ok_count': 9444, 'ng_count': 240, 'ng_rate_decimal': 0.025, 'ng_rate_percent': 2.5,
         'metric_name': 'Total NG Rate function', 'metric_value': 2.5, 'unit': '%',
         'ng_breakdown': {'NG SPL': 1, 'NG THD': 2, 'NG SPL+THD': 4, 'NG SPL+THD+F0': 0, 'NG Hearing Noise': 233, 'NG Hearing Touch': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['13.5!418']},
        {'result_id': 'res_8', 'condition_id': None, 'measurement_type': 'Function',
         'condition_group': 'Normal VP mold #2 function (14-May)', 'date': '2024-05-14', 'line': 'E2-3A',
         'input_count': 10460, 'ok_count': 10128, 'ng_count': 332, 'ng_rate_decimal': 0.032, 'ng_rate_percent': 3.2,
         'metric_name': 'Total NG Rate function', 'metric_value': 3.2, 'unit': '%',
         'ng_breakdown': {'NG SPL': 0, 'NG THD': 3, 'NG SPL+THD': 3, 'NG SPL+THD+F0': 0, 'NG Hearing Noise': 326, 'NG Hearing Touch': 0},
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['13.5!420']},
    ]

    # rel:
    # Laser cut: 0.3/0.075-1 = +300% worse
    # Bending: 1.2/1.1-1 = +9.1% worse
    # VP/CD vision: 0.3/0.2-1 = +50% worse
    # Function: 2.5/3.2-1 = -21.9% improved
    conclusions = [
        {'conclusion_id': 'concl_1', 'topic': 'VP mold #7 +0.3mm 3rd trial vs Normal VP mold #2',
         'statement_from_report': 'From result test 3rd: VP mold #7 still happen bending after laser cutting VP 1.1%, happen cutting burr and damage (rate higher than Normal VP) => So cannot use it.',
         'normalized_interpretation': "Laser cut: Test 0.3% (VP damage 0.10%, Burr 0.18%) vs Normal 0.075% = +300% worse. Bending: Test 1.2% vs Normal 1.1% = +9.1% worse. VP/CD vision: Test 0.3% vs Normal 0.2% = +50% worse (extra VP damage 1pcs only on Test). Function: Test 2.5% vs Normal 3.2% = -21.9% improved (NG Hearing Noise dominant in both). Report rejects mold #7 — laser-cut damage/burr is the killer despite function improvement.",
         'source_file': name, 'sheet_name': sheet, 'source_cells': ['13.5!402:423']},
    ]

    ts = {'defect_name': 'VP damage',
          'when_user_asks': ['VP mold #7 +0.3mm 3rd trial', 'VP laser cutting damage', 'VP bending'],
          'suggested_checks': [
              {'hint_id': 'hint_1', 'check_item': 'Reject VP mold #7 +0.3mm — laser-cut damage and burr 4x Normal',
               'reason': "Laser cut NG 0.3% (Test) vs 0.075% (Normal) = +300% worse, driven by VP damage (10 vs 2) and Burr (18 vs 1). Report decision: 'So cannot use it'.",
               'evidence_strength': 'strong', 'related_process': 'VP laser cutting',
               'related_part': 'VP mold #7', 'source_file': name, 'sheet_name': sheet,
               'source_cells': ['13.5!402:403', '13.5!423:424']}],
          'limitations': ['Function improvement is real (-21.9%) but laser-cut damage outweighs benefit per report.']}

    log = {'confidence': 0.9,
           'assumptions': ['Normal VP mold #2 same-event baseline for each measurement_type on the same dates.'],
           'warnings': [],
           'decision_rationale': "Although VP mold #7 +0.3mm gives a function-NG improvement (-21.9% vs Normal), it worsens laser-cut damage +300% and VP bending +9.1%. Report final decision rejects mold #7 because of laser-cut damage/burr."}

    result = {'schema_version': '0.1', 'document': doc,
              'test_conditions': conds, 'results': results,
              'conclusions': conclusions, 'troubleshooting_index': ts,
              'ai_extraction_log': log}

    tr_en = {
        'document': {'title': doc['title'], 'purpose': doc['purpose'], 'content': doc['content']},
        'conclusions': {c['conclusion_id']: {'topic': c['topic'], 'statement_from_report': c['statement_from_report'], 'normalized_interpretation': c['normalized_interpretation']} for c in conclusions},
        'hints': {h_['hint_id']: {'check_item': h_['check_item'], 'reason': h_['reason']} for h_ in ts['suggested_checks']},
        'log': {'assumptions': log['assumptions'], 'warnings': log['warnings'], 'decision_rationale': log['decision_rationale']},
    }
    tr_ko = {
        'document': {'title': 'BRS-161016 VP mold #7 +0.3mm 3차 시험 보고서 (13.5.2024)',
                     'purpose': 'VP mold #7 두께 +0.3mm 적용으로 Normal VP mold #2 대비 개선 여부 검증.',
                     'content': ['Laser cutting 후 VP bending 점검.', 'Semi 단계 VP+CD vision NG rate.', '최종 기능 NG rate.']},
        'conclusions': {'concl_1': {'topic': 'VP mold #7 +0.3mm 3차 vs Normal VP mold #2',
                                     'statement_from_report': '3차 시험 결과: VP mold #7 은 laser cutting 후에도 bending 1.1% 발생, cutting burr 와 damage 가 Normal 보다 높음 → 사용 불가.',
                                     'normalized_interpretation': 'Laser cut: Test 0.3% (VP damage 0.10%, Burr 0.18%) vs Normal 0.075% = +300% 악화. Bending: Test 1.2% vs Normal 1.1% = +9.1% 악화. VP/CD vision: Test 0.3% vs Normal 0.2% = +50% 악화 (VP damage 1pcs Test 만 추가). Function: Test 2.5% vs Normal 3.2% = -21.9% 개선 (NG Hearing Noise 양쪽 주도). 리포트는 mold #7 거절 — 기능 개선에도 laser-cut damage·burr 가 결정적.'}},
        'hints': {'hint_1': {'check_item': 'VP mold #7 +0.3mm 사용 거절 — laser-cut damage·burr 가 Normal 의 4배',
                              'reason': 'Laser cut NG 0.3% (Test) vs 0.075% (Normal) = +300% 악화, VP damage(10 vs 2) 와 Burr(18 vs 1) 주도. 리포트 결정문: "So cannot use it".'}},
        'log': {'assumptions': ['Normal VP mold #2 행을 각 measurement_type 별 동일 이벤트 baseline 으로 사용.'],
                 'warnings': [],
                 'decision_rationale': 'VP mold #7 +0.3mm 는 기능 NG -21.9% 개선이지만 laser-cut damage +300% 와 VP bending +9.1% 악화. 리포트 최종 결정은 laser-cut damage·burr 로 인해 mold #7 거절.'},
    }
    tr_vi = {
        'document': {'title': 'Báo cáo test lần 3 VP mold #7 +0.3mm BRS-161016 (13.5.2024)',
                     'purpose': 'Test VP mold #7 +0.3mm để xác minh cải thiện so với Normal VP mold #2.',
                     'content': ['Kiểm VP bending sau laser cutting.', 'NG rate VP+CD vision ở bán thành phẩm.', 'NG rate function ở thành phẩm.']},
        'conclusions': {'concl_1': {'topic': 'VP mold #7 +0.3mm lần 3 vs Normal VP mold #2',
                                     'statement_from_report': 'Kết quả test lần 3: VP mold #7 vẫn có bending 1.1% sau laser cutting, có burr và damage (cao hơn Normal) → Không dùng được.',
                                     'normalized_interpretation': 'Laser cut: Test 0.3% (VP damage 0.10%, Burr 0.18%) vs Normal 0.075% = +300% xấu hơn. Bending: Test 1.2% vs Normal 1.1% = +9.1% xấu hơn. VP/CD vision: Test 0.3% vs Normal 0.2% = +50% xấu hơn (Test có thêm 1pcs VP damage). Function: Test 2.5% vs Normal 3.2% = -21.9% cải thiện (NG Hearing Noise chủ đạo cả hai). Báo cáo loại bỏ mold #7 — damage·burr laser cutting là yếu tố quyết định dù function cải thiện.'}},
        'hints': {'hint_1': {'check_item': 'Loại bỏ VP mold #7 +0.3mm — damage·burr laser cut gấp 4 lần Normal',
                              'reason': 'Laser cut NG 0.3% (Test) vs 0.075% (Normal) = +300% xấu hơn, do VP damage (10 vs 2) và Burr (18 vs 1). Báo cáo: "So cannot use it".'}},
        'log': {'assumptions': ['Normal VP mold #2 dùng làm baseline cùng sự kiện cho từng measurement_type cùng ngày.'],
                 'warnings': [],
                 'decision_rationale': 'VP mold #7 +0.3mm có function NG -21.9% cải thiện nhưng laser-cut damage +300% và VP bending +9.1% xấu hơn. Báo cáo quyết định loại mold #7 vì damage·burr laser cutting.'},
    }
    return name, result, tr_ko, tr_en, tr_vi


def main():
    datasets = [ds1, ds2, ds3, ds4, ds5, ds6, ds7, ds8, ds9]
    processed = 0
    failed = 0
    for fn in datasets:
        try:
            name, result, tr_ko, tr_en, tr_vi = fn()
        except Exception as e:
            print(f'BUILD ERROR {fn.__name__}: {e!r}')
            failed += 1
            continue
        ok = h.commit_dataset(name, result, tr_ko, tr_en, tr_vi)
        if ok:
            processed += 1
            print(f'OK  {name[:80]}')
        else:
            failed += 1
            print(f'ERR {name[:80]}')
    print(f'chunk 08: processed={processed} failed={failed}')

    total, ok_count, failed_log = h.verify_counts()
    print(f'verify_counts(): targets_in_targets_file={total} processed_in_db={ok_count} failed_log_lines={failed_log}')


if __name__ == '__main__':
    main()
