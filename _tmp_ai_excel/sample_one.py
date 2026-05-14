# -*- coding: utf-8 -*-
"""Validate commit_lib on the smallest dataset (dimension report)."""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from commit_lib import commit_payload

DS = '17 BRS-161014 Report check dimension Frame array JIG+MTR 2023.09.07'

payload = {
    'schema_version': '0.1',
    'dataset_name': DS,
    'document': {
        'source_file': DS,
        'source_sheet': 'Sheet1',
        'title': 'Dimension check — Frame array JIG and JIG+FRAME',
        'model': 'BRS-161014',
        'report_date': '2023-09-07',
        'department': None, 'marker': None, 'line': None,
        'report_type': 'before_after_dimension',
        'primary_defect': {'canonical_name': 'Dimension NG', 'aliases_in_document': []},
        'related_defects': [],
        'parts': ['Frame array JIG', 'JIG+FRAME'],
        'processes': ['Dimension measurement'],
        'purpose': 'Record dimension measurements at 8 positions for FRAME ARRAY JIG and JIG+FRAME assemblies.',
        'content': [
            'FRAME ARRAY JIG positions 1-8: 0, -0.001, 0, -0.002, -0.004, -0.003, -0.003, -0.003',
            'JIG+FRAME positions 1-8: 0.853, 0.854, 0.856, 0.853, 0.854, 0.842, 0.858, 0.861',
        ],
        'source_cells': {'title': ['Sheet1!A1'], 'date': [], 'purpose': [], 'content': ['Sheet1!A3:J4']},
        'translations': {
            'ko': {'title': '치수 측정 — Frame array JIG 및 JIG+FRAME',
                   'purpose': 'FRAME ARRAY JIG 및 JIG+FRAME 조립품의 8개 위치 치수 측정값 기록.',
                   'content': ['FRAME ARRAY JIG 위치 1-8: 0, -0.001, 0, -0.002, -0.004, -0.003, -0.003, -0.003',
                               'JIG+FRAME 위치 1-8: 0.853, 0.854, 0.856, 0.853, 0.854, 0.842, 0.858, 0.861']},
            'en': {'title': 'Dimension check — Frame array JIG and JIG+FRAME',
                   'purpose': 'Record dimension measurements at 8 positions for FRAME ARRAY JIG and JIG+FRAME assemblies.',
                   'content': ['FRAME ARRAY JIG positions 1-8: 0, -0.001, 0, -0.002, -0.004, -0.003, -0.003, -0.003',
                               'JIG+FRAME positions 1-8: 0.853, 0.854, 0.856, 0.853, 0.854, 0.842, 0.858, 0.861']},
            'vi': {'title': 'Kiểm tra kích thước — Frame array JIG và JIG+FRAME',
                   'purpose': 'Ghi lại các giá trị kích thước tại 8 vị trí cho cụm FRAME ARRAY JIG và JIG+FRAME.',
                   'content': ['Vị trí 1-8 của FRAME ARRAY JIG: 0, -0.001, 0, -0.002, -0.004, -0.003, -0.003, -0.003',
                               'Vị trí 1-8 của JIG+FRAME: 0.853, 0.854, 0.856, 0.853, 0.854, 0.842, 0.858, 0.861']},
        },
    },
    'test_conditions': [],
    'results': [
        # FRAME ARRAY JIG offsets
        {'measurement_type': 'dimension', 'metric_name': 'Frame array JIG offset (mm)',
         'metric_value': -0.0025, 'unit': 'mm', 'judgement': None,
         'sheet_name': 'Sheet1', 'source_cells': ['Sheet1!C3:J3']},
        {'measurement_type': 'dimension', 'metric_name': 'JIG+FRAME height (mm)',
         'metric_value': 0.85388, 'unit': 'mm', 'judgement': None,
         'sheet_name': 'Sheet1', 'source_cells': ['Sheet1!C4:J4']},
    ],
    'conclusions': [{
        'topic': 'Dimension distribution',
        'statement_from_report': None,
        'normalized_interpretation': 'No baseline/spec row provided. Frame array JIG offsets cluster at 0 to -0.004 mm; JIG+FRAME stack height ranges 0.842-0.861 mm (mean 0.854 mm). Cannot judge improvement/worsening without spec or normal row.',
        'sheet_name': 'Sheet1',
        'source_cells': ['Sheet1!A3:J4'],
        'translations': {
            'ko': {'topic': '치수 분포',
                   'statement_from_report': None,
                   'normalized_interpretation': '대조군/스펙 행 없음. Frame array JIG 오프셋은 0~-0.004 mm 범위, JIG+FRAME 높이는 0.842-0.861 mm (평균 0.854 mm). 스펙이나 정상 행 없이는 개선/악화 판단 불가.'},
            'en': {'topic': 'Dimension distribution',
                   'statement_from_report': None,
                   'normalized_interpretation': 'No baseline/spec row provided. Frame array JIG offsets cluster at 0 to -0.004 mm; JIG+FRAME stack height ranges 0.842-0.861 mm (mean 0.854 mm). Cannot judge improvement/worsening without spec or normal row.'},
            'vi': {'topic': 'Phân bố kích thước',
                   'statement_from_report': None,
                   'normalized_interpretation': 'Không có hàng baseline/spec. Offset Frame array JIG dao động 0 đến -0.004 mm; chiều cao JIG+FRAME 0.842-0.861 mm (trung bình 0.854 mm). Không thể kết luận cải thiện/xấu đi nếu thiếu spec hoặc hàng normal.'},
        },
    }],
    'troubleshooting': {
        'defect_name': 'Dimension NG',
        'when_user_asks': ['frame dimension drift', 'jig+frame height variation'],
        'suggested_checks': [{
            'check_item': 'Provide tolerance/spec for JIG+FRAME height to enable PASS/FAIL judgement',
            'reason': 'Workbook lists 8 measured points but no upper/lower spec, so no row can be judged PASS/FAIL.',
            'evidence_strength': 'low',
            'related_process': 'Dimension measurement',
            'related_part': 'JIG+FRAME',
            'sheet_name': 'Sheet1',
            'source_cells': ['Sheet1!A4:J4'],
            'translations': {
                'ko': {'check_item': 'PASS/FAIL 판정을 위해 JIG+FRAME 높이 공차/스펙 제공 필요',
                       'reason': '8개 측정값만 있고 상·하한 스펙이 없어 PASS/FAIL을 판정할 수 없음.'},
                'en': {'check_item': 'Provide tolerance/spec for JIG+FRAME height to enable PASS/FAIL judgement',
                       'reason': 'Workbook lists 8 measured points but no upper/lower spec, so no row can be judged PASS/FAIL.'},
                'vi': {'check_item': 'Cung cấp dung sai/spec cho chiều cao JIG+FRAME để phán định PASS/FAIL',
                       'reason': 'Sổ tay chỉ liệt kê 8 điểm đo, không có spec trên/dưới nên không thể phán định PASS/FAIL.'},
            },
        }],
        'limitations': ['No spec or normal/baseline row'],
    },
    'ai_extraction_log': {
        'confidence': 0.4,
        'assumptions': ['Negative offsets in row 1 are mm; positive values in row 2 are mm stack height.'],
        'warnings': ['No baseline/spec — judgement and improvement/worsening cannot be inferred.'],
        'decision_rationale': 'Workbook is a raw dimension table with 8 measurement positions for two parts (FRAME ARRAY JIG, JIG+FRAME). No NG rate, no normal/spec row. Classified as before_after_dimension with judgement=null.',
        'translations': {
            'ko': {'assumptions': ['1행의 음수 오프셋은 mm; 2행의 양수값은 mm 스택 높이로 가정.'],
                   'warnings': ['대조군/스펙 없음 — 개선/악화 또는 판정 추론 불가.'],
                   'decision_rationale': '문서는 두 부품(FRAME ARRAY JIG, JIG+FRAME)에 대한 8개 측정 위치의 원시 치수 표. NG rate 없음, 정상/스펙 행 없음. before_after_dimension으로 분류, judgement=null.'},
            'en': {'assumptions': ['Negative offsets in row 1 are mm; positive values in row 2 are mm stack height.'],
                   'warnings': ['No baseline/spec — judgement and improvement/worsening cannot be inferred.'],
                   'decision_rationale': 'Workbook is a raw dimension table with 8 measurement positions for two parts (FRAME ARRAY JIG, JIG+FRAME). No NG rate, no normal/spec row. Classified as before_after_dimension with judgement=null.'},
            'vi': {'assumptions': ['Các giá trị âm ở hàng 1 là mm; các giá trị dương ở hàng 2 là chiều cao stack mm.'],
                   'warnings': ['Không có baseline/spec — không thể suy ra phán định hay cải thiện/xấu đi.'],
                   'decision_rationale': 'Sổ tay là bảng kích thước thô với 8 vị trí đo cho hai chi tiết (FRAME ARRAY JIG, JIG+FRAME). Không có NG rate, không có hàng normal/spec. Phân loại before_after_dimension, judgement=null.'},
        },
    },
}

doc_id = commit_payload(payload)
print('committed:', doc_id)
