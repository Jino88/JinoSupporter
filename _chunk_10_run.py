# -*- coding: utf-8 -*-
"""Chunk 10 normalization run."""
from __future__ import annotations
import sys, io, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import _ai_batch_helper as h

CHUNK_FILE = r'D:\000. MyWorks\005. Program\Repository\JinoSupporter\_chunk_10.txt'

# ---------- Dataset 1 ----------
# 31.1 TIU C11-20  Report Test new press Jig ( Frame+ VP) 2026.1.5. - RAW soundcheck
# normal_comparison. Function totals: test 1214 input, 101 NG, 8.3% vs Normal 4405 input, 435 NG, 9.9%
# rel = (8.3/9.9-1)*100 = -16.16% improved

DS1_NAME = "31.1 TIU C11-20  Report Test new press Jig ( Frame+ VP) 2026.1.5. - RAW soundcheck"
ds1_result = {
    "schema_version": "0.1",
    "document": {
        "document_id": "doc_ds1",
        "source_file": DS1_NAME,
        "source_sheet": "Test",
        "title": "Report Test Use New VP Press Jig Frame/VP TIU-C11-20",
        "model": "TIU-C11-20",
        "report_date": "2026-01-05",
        "department": "ME",
        "marker": "Thao",
        "line": "",
        "report_type": "normal_comparison",
        "primary_defect": {"canonical_name": "VP+Frame Damage",
                           "aliases_in_document": ["NG VP+Frame damage", "Not enough glue VP/Frame", "VP separate"]},
        "related_defects": ["NG Hearing Noise", "NG Sigma SPL", "NG Sigma RB", "VP+CD Separation"],
        "parts": ["VP", "Frame"],
        "processes": ["Frame+VP Ass'y", "VP Press Jig"],
        "purpose": "Find way to improve NG VP+Frame damage by testing a new VP press jig.",
        "content": [
            "Make 100pcs after ass'y Frame/VP using new VP press jig.",
            "Check NG rate for VP separate, glue issues, and function (SPL/RB/Noise/Touch)."
        ],
        "source_cells": {"title": ["Test!B2"], "date": ["Test!T3"], "purpose": ["Test!A4"], "content": ["Test!A6"]}
    },
    "test_conditions": [
        {"condition_id": "cond_1", "condition_group": "VP Press Jig",
         "line": "", "process": "Frame+VP Ass'y", "changed_factor": "VP press jig",
         "before_value": "Old VP press jig", "after_value": "New VP press jig",
         "unit": None, "machine": None, "jig": "New VP press Jig", "material_lot": None,
         "supplier": None, "dry_time_sec": None, "temperature": None, "pressure": None,
         "bond_amount": None, "uv_energy": None,
         "source_file": DS1_NAME, "sheet_name": "Test", "source_cells": ["Test!E10"]}
    ],
    "results": [
        {"result_id": "res_glue_total_test", "condition_id": "cond_1",
         "measurement_type": "Vision Frame+VP", "condition_group": "Test new VP press Jig",
         "date": "", "line": "",
         "input_count": 448, "ok_count": 438, "ng_count": 10,
         "ng_rate_decimal": 0.0223, "ng_rate_percent": 2.23,
         "metric_name": "Not enough glue VP/Frame NG rate", "metric_value": 2.23, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"Not enough glue VP/Frame": {"count": 10, "rate": 2.23}},
         "source_file": DS1_NAME, "sheet_name": "Test", "source_cells": ["Test!G15:N19"]},
        {"result_id": "res_glue_total_normal", "condition_id": None,
         "measurement_type": "Vision Frame+VP", "condition_group": "Normal",
         "date": "", "line": "",
         "input_count": 300, "ok_count": 295, "ng_count": 5,
         "ng_rate_decimal": 0.0167, "ng_rate_percent": 1.67,
         "metric_name": "Not enough glue VP/Frame NG rate (Normal)", "metric_value": 1.67, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"Not enough glue VP/Frame": {"count": 5, "rate": 1.67}},
         "source_file": DS1_NAME, "sheet_name": "Test", "source_cells": ["Test!G16:N20"]},
        {"result_id": "res_func_total_test", "condition_id": "cond_1",
         "measurement_type": "Function", "condition_group": "Test use new VP press Jig",
         "date": "", "line": "",
         "input_count": 1214, "ok_count": 1113, "ng_count": 101,
         "ng_rate_decimal": 0.083, "ng_rate_percent": 8.3,
         "metric_name": "Function Total NG rate", "metric_value": 8.3, "unit": "%",
         "judgement": None,
         "ng_breakdown": {
             "NG Sigma SPL": {"count": 39, "rate": 3.2},
             "NG Sigma SPL+RB": {"count": 9, "rate": 0.7},
             "NG Sigma RB": {"count": 64, "rate": 5.3},
             "NG Hearing No sound": {"count": 0, "rate": 0.0},
             "NG Hearing Noise": {"count": 51, "rate": 4.2},
             "NG Hearing Touch": {"count": 2, "rate": 0.2}
         },
         "source_file": DS1_NAME, "sheet_name": "Test", "source_cells": ["Test!G30:Q31"]},
        {"result_id": "res_func_total_normal", "condition_id": None,
         "measurement_type": "Function", "condition_group": "Normal",
         "date": "", "line": "",
         "input_count": 4405, "ok_count": 3970, "ng_count": 435,
         "ng_rate_decimal": 0.099, "ng_rate_percent": 9.9,
         "metric_name": "Function Total NG rate (Normal)", "metric_value": 9.9, "unit": "%",
         "judgement": None,
         "ng_breakdown": {
             "NG Sigma SPL": {"count": 38, "rate": 0.9},
             "NG Sigma SPL+RB": {"count": 20, "rate": 0.5},
             "NG Sigma RB": {"count": 568, "rate": 12.9},
             "NG Hearing No sound": {"count": 0, "rate": 0.0},
             "NG Hearing Noise": {"count": 364, "rate": 8.3},
             "NG Hearing Touch": {"count": 13, "rate": 0.3}
         },
         "source_file": DS1_NAME, "sheet_name": "Test", "source_cells": ["Test!G32:Q33"]}
    ],
    "conclusions": [
        {"conclusion_id": "concl_1",
         "topic": "Function NG comparison Test vs Normal",
         "statement_from_report": "Test use new VP press Jig: 8.3% NG vs Normal 9.9% NG.",
         "normalized_interpretation": "Test (new VP press jig) Function NG 8.3% vs Normal 9.9% = 0.84x, 16.2% improved vs same-event Normal. Dominant NG remains NG Sigma RB and NG Hearing Noise in both groups.",
         "source_file": DS1_NAME, "sheet_name": "Test", "source_cells": ["Test!Q30","Test!Q32"]}
    ],
    "troubleshooting_index": {
        "defect_name": "VP+Frame Damage",
        "when_user_asks": ["How to reduce VP+Frame damage and glue issues?"],
        "suggested_checks": [
            {"hint_id": "hint_1",
             "check_item": "Verify VP press jig design and pressing condition.",
             "reason": "Test using new VP press jig showed 8.3% function NG vs Normal 9.9% (16.2% improvement). VP separation and not-enough-glue rates remained low (<3%).",
             "evidence_strength": "medium", "related_process": "Frame+VP Ass'y", "related_part": "VP/Frame",
             "source_file": DS1_NAME, "sheet_name": "Test", "source_cells": ["Test!Q30","Test!Q32"]}
        ],
        "limitations": ["Some daily test cells have small sample size (n<200) that limits resolution."]
    },
    "ai_extraction_log": {
        "confidence": 0.7,
        "assumptions": ["Same-event Normal rows used as baseline for each daily sub-table.",
                        "Aggregated totals are summed across daily rows where the sheet did not provide a total row."],
        "warnings": ["RAW soundcheck SPL frequency tables not normalized (large tables, out of NG-rate scope)."],
        "decision_rationale": "Report belongs to normal_comparison: each daily sub-table has Test and Normal rows side by side. Relative change of function NG = (8.3/9.9 - 1)*100 = -16.2% (improved). NG Sigma RB and NG Hearing Noise dominate in both groups; jig change did not shift the NG mix significantly."
    }
}
ds1_tr_en = {
    "document": {"title": ds1_result["document"]["title"], "purpose": ds1_result["document"]["purpose"], "content": ds1_result["document"]["content"]},
    "conclusions": {"concl_1": {"topic": "Function NG comparison Test vs Normal",
                                 "statement_from_report": "Test use new VP press Jig: 8.3% NG vs Normal 9.9% NG.",
                                 "normalized_interpretation": "Test (new VP press jig) Function NG 8.3% vs Normal 9.9% = 0.84x, 16.2% improved vs same-event Normal. Dominant NG remains NG Sigma RB and NG Hearing Noise in both groups."}},
    "hints": {"hint_1": {"check_item": "Verify VP press jig design and pressing condition.",
                          "reason": "Test using new VP press jig showed 8.3% function NG vs Normal 9.9% (16.2% improvement). VP separation and not-enough-glue rates remained low (<3%)."}},
    "log": {"assumptions": ds1_result["ai_extraction_log"]["assumptions"],
            "warnings": ds1_result["ai_extraction_log"]["warnings"],
            "decision_rationale": ds1_result["ai_extraction_log"]["decision_rationale"]}
}
ds1_tr_ko = {
    "document": {"title": "신규 VP 프레스 지그 (Frame/VP) 시험 리포트 TIU-C11-20",
                  "purpose": "신규 VP 프레스 지그로 NG VP+Frame damage 개선 가능성 시험.",
                  "content": ["Frame/VP 조립 후 신규 VP 프레스 지그로 100pcs 제작.",
                              "VP separate, 글루 불량, Function (SPL/RB/Noise/Touch) NG rate 확인."]},
    "conclusions": {"concl_1": {"topic": "Function NG Test vs Normal 비교",
                                 "statement_from_report": "Test 신규 VP 프레스 지그: 8.3% NG / Normal: 9.9% NG.",
                                 "normalized_interpretation": "Test(신규 VP press jig) Function NG 8.3% vs Normal 9.9% = 0.84배, 같은 이벤트 Normal 대비 16.2% 개선. 주요 NG 항목은 양측 모두 NG Sigma RB, NG Hearing Noise로 동일."}},
    "hints": {"hint_1": {"check_item": "VP 프레스 지그 설계 및 가압 조건 확인.",
                          "reason": "신규 VP 프레스 지그 사용 시 Function NG 8.3% vs Normal 9.9% (16.2% 개선). VP separate / glue 부족 NG rate는 양측 모두 3% 미만."}},
    "log": {"assumptions": ["일자별 sub-table 내 같은 Normal 행을 baseline으로 사용.",
                            "시트에 총계가 없으면 일자별 행 합으로 총계 산출."],
            "warnings": ["RAW soundcheck SPL 주파수표는 NG rate 분석 범위 밖이라 정규화 미수행."],
            "decision_rationale": "각 일자별 sub-table에 Test와 Normal이 동시에 존재하므로 normal_comparison로 분류. 상대 변화율 = (8.3/9.9 - 1)*100 = -16.2% (개선). NG Sigma RB, NG Hearing Noise가 양측에서 지배적이며 지그 변경으로 NG 종류는 크게 변하지 않음."}
}
ds1_tr_vi = {
    "document": {"title": "Báo cáo test jig ép VP mới (Frame/VP) TIU-C11-20",
                  "purpose": "Tìm cách cải thiện NG VP+Frame damage bằng jig ép VP mới.",
                  "content": ["Sản xuất 100pcs sau khi ass'y Frame/VP bằng jig ép VP mới.",
                              "Kiểm tra NG rate VP separate, thiếu keo, function (SPL/RB/Noise/Touch)."]},
    "conclusions": {"concl_1": {"topic": "So sánh NG Function Test vs Normal",
                                 "statement_from_report": "Test dùng jig ép VP mới: 8.3% NG so với Normal 9.9% NG.",
                                 "normalized_interpretation": "Test (jig mới) Function NG 8.3% so với Normal 9.9% = 0.84x, cải thiện 16.2% so với Normal cùng sự kiện. NG Sigma RB và NG Hearing Noise vẫn là NG chính ở cả hai nhóm."}},
    "hints": {"hint_1": {"check_item": "Kiểm tra thiết kế và điều kiện ép của jig VP press.",
                          "reason": "Khi dùng jig mới, NG function 8.3% so với Normal 9.9% (cải thiện 16.2%). NG VP separate và thiếu keo đều dưới 3% ở cả hai nhóm."}},
    "log": {"assumptions": ["Sử dụng dòng Normal trong cùng bảng làm baseline cho từng ngày.",
                            "Tổng các ngày được cộng lại khi sheet không có hàng total."],
            "warnings": ["Bảng tần số SPL RAW soundcheck không nằm trong phạm vi NG rate, không chuẩn hoá."],
            "decision_rationale": "Mỗi sub-table có dòng Test và Normal song song nên phân loại là normal_comparison. Tỷ lệ thay đổi = (8.3/9.9 - 1)*100 = -16.2% (cải thiện). NG Sigma RB, NG Hearing Noise chiếm tỷ trọng cao ở cả hai bên, jig mới không làm thay đổi cấu trúc NG nhiều."}
}

# ---------- Dataset 2 ----------
# 31.MSU-20S15-07 Result checking Problem VP+CD separate date 25.9.2025
# normal_comparison: laser machine settings vs Normal. Multiple sub-events.
DS2_NAME = "31.MSU-20S15-07 Result checking Problem VP+CD separate date 25.9.2025"
ds2_result = {
    "schema_version": "0.1",
    "document": {
        "document_id": "doc_ds2",
        "source_file": DS2_NAME,
        "source_sheet": "Multiple",
        "title": "Report Checking and Test Problem VP+CD Separate of Model MSU-L20S15-07",
        "model": "MSU-L20S15-07",
        "report_date": "2025-09-25",
        "department": "ME",
        "marker": "Thao",
        "line": "E2-4A/E2-4B",
        "report_type": "normal_comparison",
        "primary_defect": {"canonical_name": "VP+CD Separation",
                           "aliases_in_document": ["VP+CD separate", "VP/CD separate"]},
        "related_defects": ["NG Hearing Noise", "NG Hearing Touch", "NG Sigma SPL", "NG Sigma THD"],
        "parts": ["VP", "CD"],
        "processes": ["Sub1 VP+CD Ass'y", "Laser CD", "Plasma"],
        "purpose": "Investigate root cause of VP+CD separate by changing laser CD marking and lowering plasma Z-axis.",
        "content": [
            "Test change laser marking CD (power and mask speed).",
            "Test down Z-axis plasma machine.",
            "Check VP+CD separation at Sub1; tension test; check function and decap NG to inspect VP+CD separate."
        ],
        "source_cells": {"title": ["15.10!B2"], "date": ["15.10!T3"], "purpose": ["15.10!A4"], "content": ["15.10!A6"]}
    },
    "test_conditions": [
        {"condition_id": "cond_laser_100", "condition_group": "Laser power/interval",
         "line": "E2-4A", "process": "Laser CD", "changed_factor": "Power/Interval",
         "before_value": "Power 100% / Interval 0.28", "after_value": "Power 100% / Interval 0.28",
         "unit": None, "machine": "Laser machine 1 (now)", "jig": None, "material_lot": None,
         "supplier": None, "dry_time_sec": None, "temperature": None, "pressure": None,
         "bond_amount": None, "uv_energy": None,
         "source_file": DS2_NAME, "sheet_name": "15.10", "source_cells": ["15.10!E10"]},
        {"condition_id": "cond_laser_95", "condition_group": "Laser power/interval",
         "line": "E2-4A", "process": "Laser CD", "changed_factor": "Power/Interval",
         "before_value": "Power 100% / Interval 0.28", "after_value": "Power 95% / Interval 0.25",
         "unit": None, "machine": "Laser machine setting 2", "jig": None, "material_lot": None,
         "supplier": None, "dry_time_sec": None, "temperature": None, "pressure": None,
         "bond_amount": None, "uv_energy": None,
         "source_file": DS2_NAME, "sheet_name": "15.10", "source_cells": ["15.10!E11"]},
        {"condition_id": "cond_laser_90", "condition_group": "Laser power/interval",
         "line": "E2-4A", "process": "Laser CD", "changed_factor": "Power/Interval",
         "before_value": "Power 100% / Interval 0.28", "after_value": "Power 90% / Interval 0.20",
         "unit": None, "machine": "Laser machine setting 3", "jig": None, "material_lot": None,
         "supplier": None, "dry_time_sec": None, "temperature": None, "pressure": None,
         "bond_amount": None, "uv_energy": None,
         "source_file": DS2_NAME, "sheet_name": "15.10", "source_cells": ["15.10!E12"]}
    ],
    "results": [
        {"result_id": "res_sep_100", "condition_id": "cond_laser_100",
         "measurement_type": "Sub1 VP+CD separate", "condition_group": "Power 100%",
         "date": "2025-10-15", "line": "E2-4A",
         "input_count": 55, "ok_count": 55, "ng_count": 0,
         "ng_rate_decimal": 0.0, "ng_rate_percent": 0.0,
         "metric_name": "VP+CD separate NG rate", "metric_value": 0.0, "unit": "%",
         "judgement": "PASS", "ng_breakdown": {"VP+CD Separation": {"count": 0, "rate": 0.0}},
         "source_file": DS2_NAME, "sheet_name": "15.10", "source_cells": ["15.10!H17"]},
        {"result_id": "res_sep_95", "condition_id": "cond_laser_95",
         "measurement_type": "Sub1 VP+CD separate", "condition_group": "Power 95%",
         "date": "2025-10-15", "line": "E2-4A",
         "input_count": 53, "ok_count": 53, "ng_count": 0,
         "ng_rate_decimal": 0.0, "ng_rate_percent": 0.0,
         "metric_name": "VP+CD separate NG rate", "metric_value": 0.0, "unit": "%",
         "judgement": "PASS", "ng_breakdown": {"VP+CD Separation": {"count": 0, "rate": 0.0}},
         "source_file": DS2_NAME, "sheet_name": "15.10", "source_cells": ["15.10!H18"]},
        {"result_id": "res_sep_90", "condition_id": "cond_laser_90",
         "measurement_type": "Sub1 VP+CD separate", "condition_group": "Power 90%",
         "date": "2025-10-15", "line": "E2-4A",
         "input_count": 54, "ok_count": 54, "ng_count": 0,
         "ng_rate_decimal": 0.0, "ng_rate_percent": 0.0,
         "metric_name": "VP+CD separate NG rate", "metric_value": 0.0, "unit": "%",
         "judgement": "PASS", "ng_breakdown": {"VP+CD Separation": {"count": 0, "rate": 0.0}},
         "source_file": DS2_NAME, "sheet_name": "15.10", "source_cells": ["15.10!H19"]},
        {"result_id": "res_tension_test", "condition_id": None,
         "measurement_type": "Tension", "condition_group": "Test Laser machine",
         "date": "2025-10-02", "line": "",
         "input_count": None, "ok_count": None, "ng_count": None,
         "ng_rate_decimal": None, "ng_rate_percent": None,
         "metric_name": "VP+CD Tension AVG", "metric_value": 2.96, "unit": "Kgf",
         "judgement": "PASS",
         "ng_breakdown": {},
         "source_file": DS2_NAME, "sheet_name": "3.10", "source_cells": ["3.10!T9"]},
        {"result_id": "res_tension_normal", "condition_id": None,
         "measurement_type": "Tension", "condition_group": "Normal",
         "date": "2025-10-02", "line": "",
         "input_count": None, "ok_count": None, "ng_count": None,
         "ng_rate_decimal": None, "ng_rate_percent": None,
         "metric_name": "VP+CD Tension AVG (Normal)", "metric_value": 2.15, "unit": "Kgf",
         "judgement": "PASS",
         "ng_breakdown": {},
         "source_file": DS2_NAME, "sheet_name": "3.10", "source_cells": ["3.10!T10"]},
        {"result_id": "res_func_test", "condition_id": None,
         "measurement_type": "Function", "condition_group": "Test Laser machine",
         "date": "2025-10-03", "line": "",
         "input_count": 1046, "ok_count": 1036, "ng_count": 10,
         "ng_rate_decimal": 0.010, "ng_rate_percent": 1.0,
         "metric_name": "Function Total NG rate", "metric_value": 1.0, "unit": "%",
         "judgement": None,
         "ng_breakdown": {
             "NG Sigma SPL": {"count": 1, "rate": 0.1},
             "NG Sigma THD": {"count": 0, "rate": 0.0},
             "NG Sigma SPL+THD": {"count": 1, "rate": 0.1},
             "NG Hearing Noise": {"count": 3, "rate": 0.3},
             "NG Hearing Touch": {"count": 5, "rate": 0.5}
         },
         "source_file": DS2_NAME, "sheet_name": "3.10", "source_cells": ["3.10!K15"]},
        {"result_id": "res_func_normal", "condition_id": None,
         "measurement_type": "Function", "condition_group": "Normal",
         "date": "2025-10-03", "line": "",
         "input_count": 1116, "ok_count": 1098, "ng_count": 18,
         "ng_rate_decimal": 0.016, "ng_rate_percent": 1.6,
         "metric_name": "Function Total NG rate (Normal)", "metric_value": 1.6, "unit": "%",
         "judgement": None,
         "ng_breakdown": {
             "NG Sigma SPL": {"count": 0, "rate": 0.0},
             "NG Sigma THD": {"count": 0, "rate": 0.0},
             "NG Sigma SPL+THD": {"count": 0, "rate": 0.0},
             "NG Hearing Noise": {"count": 13, "rate": 1.2},
             "NG Hearing Touch": {"count": 5, "rate": 0.4}
         },
         "source_file": DS2_NAME, "sheet_name": "3.10", "source_cells": ["3.10!K17"]}
    ],
    "conclusions": [
        {"conclusion_id": "concl_1",
         "topic": "Laser CD setting change reduces VP+CD separate / function NG",
         "statement_from_report": "Power 90~100% and interval 0.20~0.28 result OK => CAN USE.",
         "normalized_interpretation": "VP+CD separation NG = 0/162pcs across all 3 laser settings (100/95/90%). Function NG Test 1.0% vs Normal 1.6% = 0.625x, 37.5% improved vs same-event Normal. Tension AVG 2.96Kgf (Test) vs 2.15Kgf (Normal), both > SPEC 1.2Kgf.",
         "source_file": DS2_NAME, "sheet_name": "15.10", "source_cells": ["15.10!B25"]}
    ],
    "troubleshooting_index": {
        "defect_name": "VP+CD Separation",
        "when_user_asks": ["What to check to reduce VP+CD separation?"],
        "suggested_checks": [
            {"hint_id": "hint_1",
             "check_item": "Verify laser CD marking power and mask speed/interval window.",
             "reason": "VP+CD separation NG = 0% in all 3 laser settings (Power 90~100%, Interval 0.20~0.28). Tension test on changed-setting samples averages 2.96Kgf vs Normal 2.15Kgf, well above 1.2Kgf spec.",
             "evidence_strength": "medium", "related_process": "Laser CD", "related_part": "CD",
             "source_file": DS2_NAME, "sheet_name": "15.10", "source_cells": ["15.10!H17:H19"]}
        ],
        "limitations": ["VP+CD separate sample size per setting (~54pcs) is small; long-term run validation needed."]
    },
    "ai_extraction_log": {
        "confidence": 0.7,
        "assumptions": ["3 laser settings (100%/95%/90%) compared to Normal where available.",
                        "Tension AVG values taken from Min/Max/AVG summary columns."],
        "warnings": ["Decap function statistics from 29-Sep table are small (n=7) and percentages stored as fractions in source."],
        "decision_rationale": "Multiple sub-tables compare laser-setting Tests against Normal in the same event. VP+CD separation NG rate is 0% across all settings, function NG Test 1.0% vs Normal 1.6% (-37.5% improved), Tension all PASS spec. The 25-Sep sub-table shows Normal-before-change 4.0% sep NG vs change-settings 0~1% which further supports normal_comparison classification."
    }
}
ds2_tr_en = {
    "document": {"title": ds2_result["document"]["title"], "purpose": ds2_result["document"]["purpose"], "content": ds2_result["document"]["content"]},
    "conclusions": {"concl_1": {"topic": ds2_result["conclusions"][0]["topic"],
                                 "statement_from_report": ds2_result["conclusions"][0]["statement_from_report"],
                                 "normalized_interpretation": ds2_result["conclusions"][0]["normalized_interpretation"]}},
    "hints": {"hint_1": {"check_item": ds2_result["troubleshooting_index"]["suggested_checks"][0]["check_item"],
                          "reason": ds2_result["troubleshooting_index"]["suggested_checks"][0]["reason"]}},
    "log": {"assumptions": ds2_result["ai_extraction_log"]["assumptions"],
            "warnings": ds2_result["ai_extraction_log"]["warnings"],
            "decision_rationale": ds2_result["ai_extraction_log"]["decision_rationale"]}
}
ds2_tr_ko = {
    "document": {"title": "MSU-L20S15-07 VP+CD separate 문제 검토 및 테스트 리포트",
                  "purpose": "Laser CD 마킹 설정 변경 및 Plasma Z-axis 인하로 VP+CD separate 원인 검토.",
                  "content": ["Laser CD 마킹 Power 및 Mask speed 변경 테스트.",
                              "Plasma machine Z-axis 인하 테스트.",
                              "Sub1 VP+CD separate 확인, Tension, function, Decap NG 시 VP+CD separate 검사."]},
    "conclusions": {"concl_1": {"topic": "Laser CD 설정 변경에 따른 VP+CD separate / Function NG 감소",
                                 "statement_from_report": "Power 90~100%, Interval 0.20~0.28 모두 OK => 사용 가능.",
                                 "normalized_interpretation": "VP+CD separation NG = 0/162pcs (Power 100/95/90% 세 설정 모두). Function NG Test 1.0% vs Normal 1.6% = 0.625배, 같은 이벤트 Normal 대비 37.5% 개선. Tension 평균 2.96Kgf(Test) vs 2.15Kgf(Normal), 모두 SPEC 1.2Kgf 이상."}},
    "hints": {"hint_1": {"check_item": "Laser CD 마킹 Power 및 Interval/Mask speed 범위 확인.",
                          "reason": "Power 90~100%, Interval 0.20~0.28 모든 설정에서 VP+CD separation NG 0%. Tension Test 평균 2.96Kgf vs Normal 2.15Kgf, SPEC 1.2Kgf 대비 충분."}},
    "log": {"assumptions": ["3가지 Laser 설정(100%/95%/90%)을 Normal과 비교.",
                            "Tension AVG는 시트 Min/Max/AVG 요약 컬럼 사용."],
            "warnings": ["29-Sep decap function 통계는 n=7 수준으로 작고 일부 % 값이 분수로 저장됨."],
            "decision_rationale": "여러 sub-table에서 Laser 설정 Test와 Normal이 동시에 존재. VP+CD separation NG 0%, Function NG Test 1.0% vs Normal 1.6% (-37.5% 개선), Tension 모두 PASS. 25-Sep sub-table에서 변경 전 Normal 4.0% separate NG vs 변경 후 0~1%로 normal_comparison 분류가 타당."}
}
ds2_tr_vi = {
    "document": {"title": "Báo cáo kiểm tra và test sự cố VP+CD separate model MSU-L20S15-07",
                  "purpose": "Tìm nguyên nhân VP+CD separate qua thay đổi Laser CD và hạ Z-axis máy Plasma.",
                  "content": ["Test thay đổi Power/Mask speed laser CD.",
                              "Test hạ Z-axis máy Plasma.",
                              "Kiểm tra VP+CD separate ở Sub1, Tension, function và decap NG để kiểm tra VP+CD separate."]},
    "conclusions": {"concl_1": {"topic": "Thay đổi cài đặt Laser CD giảm VP+CD separate / NG function",
                                 "statement_from_report": "Power 90~100%, Interval 0.20~0.28 đều OK => CÓ THỂ DÙNG.",
                                 "normalized_interpretation": "VP+CD separation NG = 0/162pcs ở cả 3 cài đặt Laser (100/95/90%). NG function Test 1.0% so với Normal 1.6% = 0.625x, cải thiện 37.5% so với Normal cùng sự kiện. Tension TB 2.96Kgf (Test) vs 2.15Kgf (Normal), đều trên SPEC 1.2Kgf."}},
    "hints": {"hint_1": {"check_item": "Kiểm tra dải Power, Interval/Mask speed của máy Laser CD.",
                          "reason": "Tỷ lệ NG VP+CD separate = 0% ở cả 3 cài đặt Power 90~100%, Interval 0.20~0.28. Tension TB 2.96Kgf so với Normal 2.15Kgf, vượt SPEC 1.2Kgf."}},
    "log": {"assumptions": ["So sánh 3 cài đặt Laser (100/95/90%) với Normal khi có.",
                            "AVG Tension lấy từ cột Min/Max/AVG."],
            "warnings": ["Decap function 29-Sep cỡ mẫu nhỏ (n=7), một số % lưu dưới dạng phân số."],
            "decision_rationale": "Nhiều sub-table có Test và Normal song song. VP+CD separation NG 0%, NG function Test 1.0% vs Normal 1.6% (-37.5% cải thiện), Tension đều PASS. Bảng 25-Sep cho Normal trước thay đổi 4.0% sep NG vs sau 0~1%, khẳng định phân loại normal_comparison."}
}

# ---------- Dataset 3 ----------
# 32-1. BRS-201506 Report test New base JIG improvement NG VP laser cutting offset- date 15.3.2024
# normal_comparison. Sub VP line: Test G1/G2/G3 separated, Normal not separated.
DS3_NAME = "32-1. BRS-201506 Report test New base JIG improvement NG VP laser cutting offset- date 15.3.2024"
ds3_result = {
    "schema_version": "0.1",
    "document": {
        "document_id": "doc_ds3",
        "source_file": DS3_NAME,
        "source_sheet": "Base",
        "title": "BRS-201506 Report Test New Base JIG Improvement NG VP Laser Cutting Offset (Vision VP Sub 5)",
        "model": "BRS-201506",
        "report_date": "2024-03-15",
        "department": "ME",
        "marker": "Nhung",
        "line": "",
        "report_type": "normal_comparison",
        "primary_defect": {"canonical_name": "VP Laser Cutting Offset",
                           "aliases_in_document": ["VP laser cutting offset"]},
        "related_defects": ["VP Damage", "VP Deform", "Particle", "Separation NG", "Offset NG"],
        "parts": ["VP"],
        "processes": ["Sub VP Line", "Laser cutting", "VP Vision"],
        "purpose": "Check if new base JIG (with VP vision before pickup) can detect VP laser cutting offset.",
        "content": [
            "Normal flow: Laser cutting -> Pickup -> VP Vision.",
            "Test flow: Laser cutting -> VP Vision -> Pickup.",
            "Separate Limit G1/G2/G3 and move to main line; check function."
        ],
        "source_cells": {"title": ["Base!B2"], "date": ["Base!S3"], "purpose": ["Base!A4"], "content": ["Base!A6"]}
    },
    "test_conditions": [
        {"condition_id": "cond_1", "condition_group": "Base JIG flow change",
         "line": "", "process": "Sub VP Line", "changed_factor": "VP vision position vs pickup",
         "before_value": "Laser cutting -> Pickup -> VP Vision",
         "after_value": "Laser cutting -> VP Vision -> Pickup",
         "unit": None, "machine": "Vision VP Sub 5", "jig": "New base JIG", "material_lot": None,
         "supplier": None, "dry_time_sec": None, "temperature": None, "pressure": None,
         "bond_amount": None, "uv_energy": None,
         "source_file": DS3_NAME, "sheet_name": "Base", "source_cells": ["Base!A7"]}
    ],
    "results": [
        {"result_id": "res_sub_test_315", "condition_id": "cond_1",
         "measurement_type": "Sub VP", "condition_group": "Test",
         "date": "2024-03-15", "line": "",
         "input_count": 1996, "ok_count": 1971, "ng_count": 25,
         "ng_rate_decimal": 0.0125, "ng_rate_percent": 1.3,
         "metric_name": "Sub VP Total NG rate", "metric_value": 1.3, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"VP Damage": {"count": 2}, "VP Deform": {"count": 18}, "Particle": {"count": 2},
                          "Laser Cutting Burr": {"count": 3}, "Separation VP": {"count": 0}},
         "source_file": DS3_NAME, "sheet_name": "Base", "source_cells": ["Base!K15"]},
        {"result_id": "res_sub_normal_315", "condition_id": None,
         "measurement_type": "Sub VP", "condition_group": "Normal",
         "date": "2024-03-15", "line": "",
         "input_count": 2880, "ok_count": 2851, "ng_count": 29,
         "ng_rate_decimal": 0.0101, "ng_rate_percent": 1.0,
         "metric_name": "Sub VP Total NG rate (Normal)", "metric_value": 1.0, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"VP Damage": {"count": 2}, "VP Deform": {"count": 20}, "Particle": {"count": 3},
                          "Laser Cutting Burr": {"count": 4}, "Separation VP": {"count": 0}},
         "source_file": DS3_NAME, "sheet_name": "Base", "source_cells": ["Base!K16"]},
        {"result_id": "res_sub_test_316", "condition_id": "cond_1",
         "measurement_type": "Sub VP", "condition_group": "Test",
         "date": "2024-03-16", "line": "",
         "input_count": 1966, "ok_count": 1938, "ng_count": 28,
         "ng_rate_decimal": 0.014, "ng_rate_percent": 1.4,
         "metric_name": "Sub VP Total NG rate", "metric_value": 1.4, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"VP Damage": {"count": 2}, "VP Deform": {"count": 20}, "Particle": {"count": 3},
                          "Laser Cutting Burr": {"count": 3}, "Separation VP": {"count": 0}},
         "source_file": DS3_NAME, "sheet_name": "Base", "source_cells": ["Base!K17"]},
        {"result_id": "res_sub_normal_316", "condition_id": None,
         "measurement_type": "Sub VP", "condition_group": "Normal",
         "date": "2024-03-16", "line": "",
         "input_count": 2520, "ok_count": 2489, "ng_count": 31,
         "ng_rate_decimal": 0.0123, "ng_rate_percent": 1.2,
         "metric_name": "Sub VP Total NG rate (Normal)", "metric_value": 1.2, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"VP Damage": {"count": 3}, "VP Deform": {"count": 21}, "Particle": {"count": 4},
                          "Laser Cutting Burr": {"count": 3}, "Separation VP": {"count": 0}},
         "source_file": DS3_NAME, "sheet_name": "Base", "source_cells": ["Base!K18"]},
        {"result_id": "res_func_test_total", "condition_id": "cond_1",
         "measurement_type": "Function", "condition_group": "Test (G1+G2)",
         "date": "", "line": "",
         "input_count": 3808, "ok_count": 3637, "ng_count": 171,
         "ng_rate_decimal": 0.045, "ng_rate_percent": 4.5,
         "metric_name": "Function Total NG rate", "metric_value": 4.5, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"NG Sigma SPL": {"count": 3, "rate": 1.8},
                          "NG Sigma THD": {"count": 1, "rate": 0.6},
                          "NG Hearing Noise": {"count": 39, "rate": 22.8},
                          "NG Hearing Touch": {"count": 128, "rate": 74.9}},
         "source_file": DS3_NAME, "sheet_name": "Base", "source_cells": ["Base!K24"]},
        {"result_id": "res_func_test_g3", "condition_id": "cond_1",
         "measurement_type": "Function", "condition_group": "Test G3",
         "date": "", "line": "",
         "input_count": 69, "ok_count": 65, "ng_count": 4,
         "ng_rate_decimal": 0.058, "ng_rate_percent": 5.8,
         "metric_name": "Function Total NG rate (G3)", "metric_value": 5.8, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"NG Hearing Touch": {"count": 4, "rate": 100.0}},
         "source_file": DS3_NAME, "sheet_name": "Base", "source_cells": ["Base!K26"]},
        {"result_id": "res_func_normal", "condition_id": None,
         "measurement_type": "Function", "condition_group": "Normal",
         "date": "", "line": "",
         "input_count": 600, "ok_count": 564, "ng_count": 36,
         "ng_rate_decimal": 0.06, "ng_rate_percent": 6.0,
         "metric_name": "Function Total NG rate (Normal)", "metric_value": 6.0, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"NG Hearing Noise": {"count": 9, "rate": 25.0},
                          "NG Hearing Touch": {"count": 27, "rate": 75.0}},
         "source_file": DS3_NAME, "sheet_name": "Base", "source_cells": ["Base!K28"]}
    ],
    "conclusions": [
        {"conclusion_id": "concl_1",
         "topic": "New base JIG flow effect on Sub VP NG and Function NG",
         "statement_from_report": "New base JIG can separate VP laser cutting offset into G1/G2/G3 groups; sub-line NG rates similar to Normal.",
         "normalized_interpretation": "Sub VP NG rates Test vs Normal: 1.3%/1.0% (15-Mar) and 1.4%/1.2% (16-Mar) - relative change +30% and +16.7% (slightly worse). Function NG Test G1+G2 4.5% vs Normal 6.0% = 0.75x, 25.0% improved vs Normal. G3 (most-offset group) 5.8% NG, dominated by NG Hearing Touch (100% of G3 NG).",
         "source_file": DS3_NAME, "sheet_name": "Base", "source_cells": ["Base!K15:K28"]}
    ],
    "troubleshooting_index": {
        "defect_name": "VP Laser Cutting Offset",
        "when_user_asks": ["How to detect VP laser cutting offset before pickup?"],
        "suggested_checks": [
            {"hint_id": "hint_1",
             "check_item": "Verify VP vision position before pickup and separation limits G1/G2/G3.",
             "reason": "New base JIG allows VP vision before pickup. Sub VP NG comparable to Normal (1.3~1.4% vs 1.0~1.2%). Function NG Test G1+G2 4.5% improved 25% vs Normal 6.0%, but G3 group (high offset) still 5.8% NG (Touch 100% of NG).",
             "evidence_strength": "medium", "related_process": "Sub VP Line", "related_part": "VP",
             "source_file": DS3_NAME, "sheet_name": "Base", "source_cells": ["Base!K24:K28"]}
        ],
        "limitations": ["G3 sample size is small (69pcs); long-run monitoring needed for Touch NG mode."]
    },
    "ai_extraction_log": {
        "confidence": 0.7,
        "assumptions": ["Same-date Normal rows used as baseline for Sub VP. Function Normal row used for total comparison.",
                        "NG breakdown counts taken from per-defect columns."],
        "warnings": ["Sub VP NG rate slightly higher on Test side (small absolute delta)."],
        "decision_rationale": "Each daily Sub VP table has Test and Normal rows with same date/line, so normal_comparison applies. Sub VP NG relative change = (1.3/1.0-1)*100 = +30% (worse) on 15-Mar and +16.7% on 16-Mar, but absolute deltas are 0.2~0.3 ppt - within day-to-day noise. Function NG G1+G2 improved 25% vs Normal; G3 small-n indicates Touch NG persists at high-offset boundary."
    }
}
ds3_tr_en = {
    "document": {"title": ds3_result["document"]["title"], "purpose": ds3_result["document"]["purpose"], "content": ds3_result["document"]["content"]},
    "conclusions": {"concl_1": {"topic": ds3_result["conclusions"][0]["topic"],
                                 "statement_from_report": ds3_result["conclusions"][0]["statement_from_report"],
                                 "normalized_interpretation": ds3_result["conclusions"][0]["normalized_interpretation"]}},
    "hints": {"hint_1": {"check_item": ds3_result["troubleshooting_index"]["suggested_checks"][0]["check_item"],
                          "reason": ds3_result["troubleshooting_index"]["suggested_checks"][0]["reason"]}},
    "log": {"assumptions": ds3_result["ai_extraction_log"]["assumptions"],
            "warnings": ds3_result["ai_extraction_log"]["warnings"],
            "decision_rationale": ds3_result["ai_extraction_log"]["decision_rationale"]}
}
ds3_tr_ko = {
    "document": {"title": "BRS-201506 신규 베이스 지그 VP laser cutting offset 개선 시험 리포트",
                  "purpose": "신규 베이스 지그에서 Vision으로 VP laser cutting offset 검출 가능 여부 확인.",
                  "content": ["Normal flow: Laser cutting -> Pickup -> VP Vision.",
                              "Test flow: Laser cutting -> VP Vision -> Pickup.",
                              "Limit G1/G2/G3로 분리 후 main line 진행, function 확인."]},
    "conclusions": {"concl_1": {"topic": "신규 베이스 지그 흐름의 Sub VP / Function NG 영향",
                                 "statement_from_report": "신규 베이스 지그는 VP laser cutting offset을 G1/G2/G3로 분리 가능, Sub line NG rate는 Normal과 유사.",
                                 "normalized_interpretation": "Sub VP NG Test vs Normal: 15-Mar 1.3%/1.0%, 16-Mar 1.4%/1.2% (상대 변화 +30%, +16.7%, 미약한 악화). Function NG Test G1+G2 4.5% vs Normal 6.0% = 0.75배, 25.0% 개선. G3(최대 offset 군) NG 5.8%로 NG Hearing Touch 비중 100%."}},
    "hints": {"hint_1": {"check_item": "Pickup 전 VP vision 위치 및 G1/G2/G3 분리 한계 검토.",
                          "reason": "신규 베이스 지그로 pickup 전 VP vision 가능. Sub VP NG는 Normal과 유사(1.3~1.4% vs 1.0~1.2%). Function NG Test G1+G2 4.5%로 Normal 6.0% 대비 25% 개선, G3는 5.8% NG에 Touch가 100%."}},
    "log": {"assumptions": ["같은 날짜 Normal 행을 Sub VP baseline으로 사용.",
                            "NG 종류별 컬럼에서 직접 카운트 추출."],
            "warnings": ["Test 측 Sub VP NG가 Normal보다 약간 높지만 절대 차이는 0.2~0.3 ppt 수준."],
            "decision_rationale": "일자별 Sub VP에 Test/Normal 동시 행, normal_comparison. Sub VP 상대 변화 +30%/+16.7%이지만 절대 0.2~0.3 ppt로 일내 변동 범위. Function G1+G2 25% 개선, G3 소수 표본이라 Touch NG가 잔존."}
}
ds3_tr_vi = {
    "document": {"title": "Báo cáo test JIG base mới cải thiện NG VP laser cutting offset BRS-201506",
                  "purpose": "Kiểm tra JIG base mới có phát hiện VP laser cutting offset bằng vision không.",
                  "content": ["Luồng Normal: Laser cutting -> Pickup -> VP Vision.",
                              "Luồng Test: Laser cutting -> VP Vision -> Pickup.",
                              "Phân loại G1/G2/G3 sau đó chạy main line, kiểm tra function."]},
    "conclusions": {"concl_1": {"topic": "Tác động của luồng JIG base mới tới NG Sub VP / Function",
                                 "statement_from_report": "JIG base mới phân loại VP laser cutting offset theo G1/G2/G3, NG Sub line tương đương Normal.",
                                 "normalized_interpretation": "NG Sub VP Test vs Normal: 15-Mar 1.3%/1.0% và 16-Mar 1.4%/1.2% (thay đổi tương đối +30% và +16.7%, hơi xấu hơn). NG function Test G1+G2 4.5% vs Normal 6.0% = 0.75x, cải thiện 25.0%. Nhóm G3 (offset cao nhất) NG 5.8% trong đó Hearing Touch chiếm 100%."}},
    "hints": {"hint_1": {"check_item": "Kiểm tra vị trí VP vision trước pickup và ngưỡng G1/G2/G3.",
                          "reason": "JIG base mới cho phép VP vision trước pickup. NG Sub VP ngang Normal (1.3~1.4% vs 1.0~1.2%). NG function Test G1+G2 4.5% cải thiện 25% so với Normal 6.0%, nhưng G3 vẫn 5.8% NG với Touch chiếm 100%."}},
    "log": {"assumptions": ["Dùng dòng Normal cùng ngày làm baseline cho Sub VP.",
                            "Số lượng NG lấy từ cột chi tiết NG."],
            "warnings": ["NG Sub VP phía Test hơi cao hơn Normal nhưng chênh tuyệt đối chỉ 0.2~0.3 ppt."],
            "decision_rationale": "Mỗi sub-table Sub VP có cả Test và Normal cùng ngày nên normal_comparison. Tỷ lệ thay đổi +30%/+16.7% nhưng chênh tuyệt đối nhỏ. NG function G1+G2 cải thiện 25%; G3 mẫu nhỏ nên NG Touch còn cao."}
}

# ---------- Dataset 4 ----------
# 32. BRS-161014 Report TEST AWF #1 & 3 - normal_comparison
DS4_NAME = "32. BRS-161014 Report TEST AWF #1 & 3"
ds4_result = {
    "schema_version": "0.1",
    "document": {
        "document_id": "doc_ds4",
        "source_file": DS4_NAME,
        "source_sheet": "Report (2)",
        "title": "Report Test AWF #1 & AWF #3 Machine -161014",
        "model": "BRS-161014",
        "report_date": "2023-09-19",
        "department": "ME",
        "marker": "Thao/Thuy",
        "line": "",
        "report_type": "normal_comparison",
        "primary_defect": {"canonical_name": "NG Function High Rate",
                           "aliases_in_document": ["NG function high"]},
        "related_defects": ["NG Hearing Noise", "NG Hearing Touch", "NG Sigma SPL+THD"],
        "parts": ["Coil"],
        "processes": ["AWF Winding"],
        "purpose": "Improve NG function high rate by testing AWF #1 (insert cooling part after circle coil winding) and AWF #3 (change pole).",
        "content": [
            "AWF #1 - insert cooling part after winding circle coil; separate lot to compare with normal AWF.",
            "AWF #3 - change pole; compare function NG with normal AWF."
        ],
        "source_cells": {"title": ["Report (2)!B2"], "date": ["Report (2)!T3"], "purpose": ["Report (2)!A4"], "content": ["Report (2)!A6"]}
    },
    "test_conditions": [
        {"condition_id": "cond_awf1", "condition_group": "AWF #1",
         "line": "", "process": "AWF Winding", "changed_factor": "Insert cooling part after circle coil winding",
         "before_value": "Normal AWF", "after_value": "AWF #1 with cooling part",
         "unit": None, "machine": "AWF #1", "jig": None, "material_lot": None,
         "supplier": None, "dry_time_sec": None, "temperature": None, "pressure": None,
         "bond_amount": None, "uv_energy": None,
         "source_file": DS4_NAME, "sheet_name": "Report (2)", "source_cells": ["Report (2)!A8"]},
        {"condition_id": "cond_awf3", "condition_group": "AWF #3",
         "line": "", "process": "AWF Winding", "changed_factor": "Change pole",
         "before_value": "Normal AWF", "after_value": "AWF #3 changed pole",
         "unit": None, "machine": "AWF #3", "jig": None, "material_lot": None,
         "supplier": None, "dry_time_sec": None, "temperature": None, "pressure": None,
         "bond_amount": None, "uv_energy": None,
         "source_file": DS4_NAME, "sheet_name": "Report (2)", "source_cells": ["Report (2)!A9"]}
    ],
    "results": [
        {"result_id": "res_awf1", "condition_id": "cond_awf1",
         "measurement_type": "Function", "condition_group": "AWF #1",
         "date": "2023-09-19", "line": "",
         "input_count": 133, "ok_count": 82, "ng_count": 51,
         "ng_rate_decimal": 0.383, "ng_rate_percent": 38.3,
         "metric_name": "Function Total NG rate", "metric_value": 38.3, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"NG Sigma SPL+THD": {"count": 1, "rate": 2.0},
                          "NG Hearing Noise": {"count": 35, "rate": 68.6},
                          "NG Hearing Touch": {"count": 15, "rate": 29.4}},
         "source_file": DS4_NAME, "sheet_name": "Report (2)", "source_cells": ["Report (2)!N13"]},
        {"result_id": "res_awf3", "condition_id": "cond_awf3",
         "measurement_type": "Function", "condition_group": "AWF #3",
         "date": "2023-09-19", "line": "",
         "input_count": 148, "ok_count": 114, "ng_count": 34,
         "ng_rate_decimal": 0.23, "ng_rate_percent": 23.0,
         "metric_name": "Function Total NG rate", "metric_value": 23.0, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"NG Sigma SPL+THD": {"count": 1, "rate": 2.9},
                          "NG Hearing Noise": {"count": 24, "rate": 70.6},
                          "NG Hearing Touch": {"count": 9, "rate": 26.5}},
         "source_file": DS4_NAME, "sheet_name": "Report (2)", "source_cells": ["Report (2)!N15"]},
        {"result_id": "res_normal", "condition_id": None,
         "measurement_type": "Function", "condition_group": "Normal",
         "date": "2023-09-19", "line": "",
         "input_count": 991, "ok_count": 507, "ng_count": 484,
         "ng_rate_decimal": 0.488, "ng_rate_percent": 48.8,
         "metric_name": "Function Total NG rate (Normal)", "metric_value": 48.8, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"NG Sigma SPL": {"count": 1, "rate": 0.2},
                          "NG Sigma SPL+THD": {"count": 4, "rate": 0.8},
                          "NG Hearing Noise": {"count": 402, "rate": 83.1},
                          "NG Hearing Touch": {"count": 77, "rate": 15.9}},
         "source_file": DS4_NAME, "sheet_name": "Report (2)", "source_cells": ["Report (2)!N17"]}
    ],
    "conclusions": [
        {"conclusion_id": "concl_1",
         "topic": "AWF #1 and #3 function NG comparison vs Normal AWF",
         "statement_from_report": "NG rate of function of AWF #1 and #3 is reduced.",
         "normalized_interpretation": "AWF #1 38.3% vs Normal 48.8% = 0.785x, 21.5% improved. AWF #3 23.0% vs Normal 48.8% = 0.471x, 52.9% improved. NG Hearing Noise dominates all three groups (~68-83%). Touch NG share increases on test machines (~26-29% vs 15.9% Normal).",
         "source_file": DS4_NAME, "sheet_name": "Report (2)", "source_cells": ["Report (2)!B19"]}
    ],
    "troubleshooting_index": {
        "defect_name": "NG Function High Rate",
        "when_user_asks": ["How to reduce function NG via AWF winding settings?"],
        "suggested_checks": [
            {"hint_id": "hint_1",
             "check_item": "Verify AWF cooling-part insertion (#1) and pole-change (#3) configurations.",
             "reason": "AWF #1 NG 38.3% vs Normal 48.8% (21.5% improved). AWF #3 NG 23.0% vs Normal 48.8% (52.9% improved). NG Hearing Noise is the dominant NG mode across all three groups.",
             "evidence_strength": "medium", "related_process": "AWF Winding", "related_part": "Coil",
             "source_file": DS4_NAME, "sheet_name": "Report (2)", "source_cells": ["Report (2)!N13:N17"]}
        ],
        "limitations": ["Sample size for AWF #1/#3 (~130-150) is much smaller than Normal (~991)."]
    },
    "ai_extraction_log": {
        "confidence": 0.7,
        "assumptions": ["Same-date Normal row used as baseline for both test groups."],
        "warnings": ["NG Hearing Noise rate stays high (~68-83%) even after improvement, suggesting noise root cause not fully resolved."],
        "decision_rationale": "Same-event Normal row exists; normal_comparison classification. Test conditions clearly identified per machine. AWF #3 shows larger relative improvement (52.9%) than AWF #1 (21.5%); Hearing Noise still dominant across all groups."
    }
}
ds4_tr_en = {
    "document": {"title": ds4_result["document"]["title"], "purpose": ds4_result["document"]["purpose"], "content": ds4_result["document"]["content"]},
    "conclusions": {"concl_1": {"topic": ds4_result["conclusions"][0]["topic"],
                                 "statement_from_report": ds4_result["conclusions"][0]["statement_from_report"],
                                 "normalized_interpretation": ds4_result["conclusions"][0]["normalized_interpretation"]}},
    "hints": {"hint_1": {"check_item": ds4_result["troubleshooting_index"]["suggested_checks"][0]["check_item"],
                          "reason": ds4_result["troubleshooting_index"]["suggested_checks"][0]["reason"]}},
    "log": {"assumptions": ds4_result["ai_extraction_log"]["assumptions"],
            "warnings": ds4_result["ai_extraction_log"]["warnings"],
            "decision_rationale": ds4_result["ai_extraction_log"]["decision_rationale"]}
}
ds4_tr_ko = {
    "document": {"title": "BRS-161014 AWF #1 & #3 시험 리포트",
                  "purpose": "AWF #1 (Circle Coil 권선 후 cooling part 삽입)과 AWF #3 (pole 변경)으로 NG function high rate 개선 검토.",
                  "content": ["AWF #1: Circle Coil 권선 후 cooling part 삽입, lot 분리로 normal AWF와 NG rate 비교.",
                              "AWF #3: pole 변경 후 normal AWF와 function NG 비교."]},
    "conclusions": {"concl_1": {"topic": "AWF #1, #3 Function NG Normal AWF 대비 비교",
                                 "statement_from_report": "AWF #1, #3의 function NG rate 감소.",
                                 "normalized_interpretation": "AWF #1 38.3% vs Normal 48.8% = 0.785배, 21.5% 개선. AWF #3 23.0% vs Normal 48.8% = 0.471배, 52.9% 개선. 세 그룹 모두 NG Hearing Noise가 지배적(약 68-83%). Touch NG 비중은 test 기기에서 증가(약 26-29%, Normal 15.9%)."}},
    "hints": {"hint_1": {"check_item": "AWF #1의 cooling part 삽입 및 AWF #3의 pole 변경 조건 확인.",
                          "reason": "AWF #1 NG 38.3% vs Normal 48.8% (21.5% 개선). AWF #3 NG 23.0% vs Normal 48.8% (52.9% 개선). 모든 그룹에서 NG Hearing Noise가 최대 NG mode."}},
    "log": {"assumptions": ["같은 날 Normal 행을 두 test 그룹의 baseline으로 사용."],
            "warnings": ["개선 후에도 NG Hearing Noise rate가 약 68-83%로 높아 noise 근본 원인은 완전 해소되지 않음."],
            "decision_rationale": "같은 이벤트 Normal 행 존재로 normal_comparison. 시험 조건은 기기별로 명확. AWF #3가 #1보다 상대 개선이 크나(52.9% vs 21.5%) NG Hearing Noise는 여전히 지배적."}
}
ds4_tr_vi = {
    "document": {"title": "Báo cáo test AWF #1 & AWF #3 máy -161014",
                  "purpose": "Cải thiện NG function cao bằng AWF #1 (chèn cooling part sau khi quấn circle coil) và AWF #3 (đổi pole).",
                  "content": ["AWF #1: chèn cooling part sau khi quấn circle coil, tách lot để so sánh với AWF thường.",
                              "AWF #3: đổi pole, so sánh NG function với AWF thường."]},
    "conclusions": {"concl_1": {"topic": "So sánh NG function AWF #1, #3 với AWF Normal",
                                 "statement_from_report": "NG function của AWF #1 và #3 giảm.",
                                 "normalized_interpretation": "AWF #1 38.3% so với Normal 48.8% = 0.785x, cải thiện 21.5%. AWF #3 23.0% so với Normal 48.8% = 0.471x, cải thiện 52.9%. Cả ba nhóm đều có NG Hearing Noise chi phối (khoảng 68-83%). Tỷ trọng NG Touch tăng ở máy test (~26-29% so với 15.9% ở Normal)."}},
    "hints": {"hint_1": {"check_item": "Kiểm tra điều kiện chèn cooling part AWF #1 và đổi pole AWF #3.",
                          "reason": "AWF #1 NG 38.3% vs Normal 48.8% (cải thiện 21.5%). AWF #3 NG 23.0% vs Normal 48.8% (cải thiện 52.9%). NG Hearing Noise vẫn là dạng NG chủ đạo ở mọi nhóm."}},
    "log": {"assumptions": ["Dùng dòng Normal cùng ngày làm baseline cho hai nhóm test."],
            "warnings": ["NG Hearing Noise vẫn ở mức 68-83%, nguyên nhân chính của noise chưa được giải quyết hoàn toàn."],
            "decision_rationale": "Có dòng Normal cùng sự kiện nên phân loại normal_comparison. Điều kiện test rõ ràng theo từng máy. AWF #3 cải thiện tương đối lớn hơn AWF #1 nhưng Hearing Noise vẫn chi phối."}
}

# ---------- Dataset 5 ----------
# 32. BRS-161014 Report check dimension Frame 2023.4.23 - before_after_dimension
DS5_NAME = "32. BRS-161014 Report check dimension Frame 2023.4.23"
ds5_result = {
    "schema_version": "0.1",
    "document": {
        "document_id": "doc_ds5",
        "source_file": DS5_NAME,
        "source_sheet": "Multiple",
        "title": "Report Check Dimension Material Frame and Ass'y Frame+Yoke, Gap Check BRS-161014/161016",
        "model": "BRS-161014/161016",
        "report_date": "2023-04-23",
        "department": "ME",
        "marker": "Le",
        "line": "C2-3A",
        "report_type": "before_after_dimension",
        "primary_defect": {"canonical_name": "NG Hearing",
                           "aliases_in_document": ["NG hearing", "FP+YK offset"]},
        "related_defects": ["Dimension NG", "FP+YK Offset", "NG Hearing Noise"],
        "parts": ["Frame", "Yoke"],
        "processes": ["Frame mold", "Frame+Yoke Ass'y"],
        "purpose": "Find reason of NG hearing by checking material frame dimension and Frame+Yoke assembly gap.",
        "content": [
            "Check frame material dimension on 2D machine (Spec width 5.82~5.86, length 14.23~14.27).",
            "Check Frame+Yoke assembly dimension on 2D machine.",
            "Check Gap Frame+Yoke per cavity (G3-D high NG vs G2-A low NG) by position.",
            "Test separate cavity mold frame 3 -> check function (A~G cavity)."
        ],
        "source_cells": {"title": ["Check Dimension!B2"], "date": ["Check Dimension!K3"], "purpose": ["Check Dimension!A4"], "content": ["Check Dimension!A6"]}
    },
    "test_conditions": [
        {"condition_id": "cond_1", "condition_group": "Cavity comparison",
         "line": "C2-3A", "process": "Frame Mold", "changed_factor": "Cavity (A~G)",
         "before_value": "Mixed cavity", "after_value": "Separated cavity A~G",
         "unit": None, "machine": "2D machine / Function test", "jig": None, "material_lot": None,
         "supplier": None, "dry_time_sec": None, "temperature": None, "pressure": None,
         "bond_amount": None, "uv_energy": None,
         "source_file": DS5_NAME, "sheet_name": "MOLD FRAME 3", "source_cells": ["MOLD FRAME 3!A8"]}
    ],
    "results": [
        {"result_id": "res_frame_width", "condition_id": None,
         "measurement_type": "Dimension", "condition_group": "Material Frame width",
         "date": "2023-05-03", "line": "",
         "input_count": 5, "ok_count": None, "ng_count": None,
         "ng_rate_decimal": None, "ng_rate_percent": None,
         "metric_name": "Frame Width Avg", "metric_value": 5.850, "unit": "mm",
         "judgement": "PASS",
         "ng_breakdown": {},
         "source_file": DS5_NAME, "sheet_name": "Check Dimension", "source_cells": ["Check Dimension!E12"]},
        {"result_id": "res_frame_length", "condition_id": None,
         "measurement_type": "Dimension", "condition_group": "Material Frame length",
         "date": "2023-05-03", "line": "",
         "input_count": 5, "ok_count": None, "ng_count": None,
         "ng_rate_decimal": None, "ng_rate_percent": None,
         "metric_name": "Frame Length Avg", "metric_value": 14.254, "unit": "mm",
         "judgement": "PASS",
         "ng_breakdown": {},
         "source_file": DS5_NAME, "sheet_name": "Check Dimension", "source_cells": ["Check Dimension!G12"]},
        {"result_id": "res_yoke_width", "condition_id": None,
         "measurement_type": "Dimension", "condition_group": "Yoke width (Spec 5.82~5.85)",
         "date": "2023-04-27", "line": "",
         "input_count": 5, "ok_count": None, "ng_count": None,
         "ng_rate_decimal": None, "ng_rate_percent": None,
         "metric_name": "Yoke Width Avg", "metric_value": 5.807, "unit": "mm",
         "judgement": "FAIL",
         "ng_breakdown": {},
         "source_file": DS5_NAME, "sheet_name": "Check dimesion Frame  + Yoke", "source_cells": ["Check dimesion Frame  + Yoke!E14"]},
        {"result_id": "res_yoke_length", "condition_id": None,
         "measurement_type": "Dimension", "condition_group": "Yoke length (Spec 14.23~14.26)",
         "date": "2023-04-27", "line": "",
         "input_count": 5, "ok_count": None, "ng_count": None,
         "ng_rate_decimal": None, "ng_rate_percent": None,
         "metric_name": "Yoke Length Avg", "metric_value": 14.224, "unit": "mm",
         "judgement": "FAIL",
         "ng_breakdown": {},
         "source_file": DS5_NAME, "sheet_name": "Check dimesion Frame  + Yoke", "source_cells": ["Check dimesion Frame  + Yoke!G14"]},
        {"result_id": "res_gap_g3d", "condition_id": None,
         "measurement_type": "Gap", "condition_group": "G3 Cavity D (NG high)",
         "date": "2023-05-02", "line": "",
         "input_count": 10, "ok_count": None, "ng_count": None,
         "ng_rate_decimal": None, "ng_rate_percent": None,
         "metric_name": "Frame-Yoke Gap AVG (G3-D)", "metric_value": 0.060, "unit": "mm",
         "judgement": "FAIL",
         "ng_breakdown": {},
         "source_file": DS5_NAME, "sheet_name": "Check Gap Frame  + Yoke", "source_cells": ["Check Gap Frame  + Yoke!H14"]},
        {"result_id": "res_gap_g2a", "condition_id": None,
         "measurement_type": "Gap", "condition_group": "G2 Cavity A (NG low)",
         "date": "2023-05-02", "line": "",
         "input_count": 10, "ok_count": None, "ng_count": None,
         "ng_rate_decimal": None, "ng_rate_percent": None,
         "metric_name": "Frame-Yoke Gap AVG (G2-A)", "metric_value": 0.040, "unit": "mm",
         "judgement": "PASS",
         "ng_breakdown": {},
         "source_file": DS5_NAME, "sheet_name": "Check Gap Frame  + Yoke", "source_cells": ["Check Gap Frame  + Yoke!H22"]},
        {"result_id": "res_func_total", "condition_id": "cond_1",
         "measurement_type": "Function", "condition_group": "Mold Frame 3 total (A~G)",
         "date": "", "line": "C2-3A",
         "input_count": 4676, "ok_count": 4572, "ng_count": 104,
         "ng_rate_decimal": 0.022, "ng_rate_percent": 2.2,
         "metric_name": "Function Total NG rate", "metric_value": 2.2, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"NG Hearing Noise": {"count": 104, "rate": 2.2}},
         "source_file": DS5_NAME, "sheet_name": "MOLD FRAME 3", "source_cells": ["MOLD FRAME 3!L11"]}
    ],
    "conclusions": [
        {"conclusion_id": "concl_1",
         "topic": "Frame/Yoke dimension and gap drive Hearing NG variation",
         "statement_from_report": "Frame within tolerance; some YK lower than spec; high-NG cavity G3-D gap 0.05~0.07 OVER spec; low-NG cavity G2-A gap 0.029~0.055 within spec. Need >=20% material tolerance tightening.",
         "normalized_interpretation": "Frame dimension Avg 5.850/14.254 within spec. Yoke dimension Avg 5.807/14.224 below lower-spec limit (5.82/14.23) - FAIL. G3-D gap AVG 0.060 mm > G2-A 0.040 mm (50% higher). Cavity-separated function NG: A 2.48%, B 2.11%, C 2.84%, D 2.71%, E 2.77%, F 1.53%, G 1.14% - F/G lowest, C/D/E highest. No same-event Normal row for direct relative-change calculation; report is dimension/gap-driven.",
         "source_file": DS5_NAME, "sheet_name": "SUMMARY", "source_cells": ["SUMMARY!A2:A8"]}
    ],
    "troubleshooting_index": {
        "defect_name": "NG Hearing (Frame+Yoke offset)",
        "when_user_asks": ["What dimension to tighten to reduce hearing NG?"],
        "suggested_checks": [
            {"hint_id": "hint_1",
             "check_item": "Tighten Yoke width/length tolerance (>=20%) and re-check Frame+Yoke gap by cavity.",
             "reason": "Yoke width Avg 5.807 mm < lower spec 5.82; Yoke length Avg 14.224 mm < lower spec 14.23 (FAIL). High-NG cavity G3-D gap Avg 0.060 mm vs low-NG G2-A 0.040 mm (50% larger). Cavity NG ranking matches: D/C/E top, F/G bottom.",
             "evidence_strength": "high", "related_process": "Frame Mold / Frame+Yoke Ass'y", "related_part": "Yoke",
             "source_file": DS5_NAME, "sheet_name": "Multiple", "source_cells": ["SUMMARY!A2:A8"]}
        ],
        "limitations": ["Only one yoke sample group measured; Frame mold 3 vs mold 5 not normalized separately."]
    },
    "ai_extraction_log": {
        "confidence": 0.7,
        "assumptions": ["FAIL judgement for Yoke based on Avg vs Spec lower limit per Summary text.",
                        "Cavity NG ratios taken from MOLD FRAME 3 TOTAL row."],
        "warnings": ["No same-event Normal/Baseline row exists for absolute NG comparison; dimension-driven evidence only."],
        "decision_rationale": "Primary evidence is dimensional: Frame within spec, Yoke below spec, gap higher on high-NG cavity. Function table is supplementary cavity-mix breakdown. before_after_dimension fits because the report contrasts material/assembly dimensions against spec and high/low NG cavity gap measurements."
    }
}
ds5_tr_en = {
    "document": {"title": ds5_result["document"]["title"], "purpose": ds5_result["document"]["purpose"], "content": ds5_result["document"]["content"]},
    "conclusions": {"concl_1": {"topic": ds5_result["conclusions"][0]["topic"],
                                 "statement_from_report": ds5_result["conclusions"][0]["statement_from_report"],
                                 "normalized_interpretation": ds5_result["conclusions"][0]["normalized_interpretation"]}},
    "hints": {"hint_1": {"check_item": ds5_result["troubleshooting_index"]["suggested_checks"][0]["check_item"],
                          "reason": ds5_result["troubleshooting_index"]["suggested_checks"][0]["reason"]}},
    "log": {"assumptions": ds5_result["ai_extraction_log"]["assumptions"],
            "warnings": ds5_result["ai_extraction_log"]["warnings"],
            "decision_rationale": ds5_result["ai_extraction_log"]["decision_rationale"]}
}
ds5_tr_ko = {
    "document": {"title": "BRS-161014/161016 Frame 치수 및 Frame+Yoke Gap 검토 리포트",
                  "purpose": "원자재 Frame 치수와 Frame+Yoke 조립 Gap을 확인하여 NG hearing 원인 분석.",
                  "content": ["2D 측정기로 Frame 원자재 치수 확인 (Spec width 5.82~5.86, length 14.23~14.27).",
                              "Frame+Yoke 조립 치수 2D 측정기로 확인.",
                              "Frame+Yoke Gap을 Cavity별 확인 (G3-D 불량 多, G2-A 불량 少).",
                              "Mold Frame 3의 cavity A~G 분리 후 function 확인."]},
    "conclusions": {"concl_1": {"topic": "Frame/Yoke 치수 및 Gap이 Hearing NG 편차의 원인",
                                 "statement_from_report": "Frame은 공차 내, 일부 YK는 공차 하한 OVER; 불량 多 G3-D Gap 0.05~0.07 SPEC 초과; 불량 少 G2-A Gap 0.029~0.055 SPEC 내. 자재 공차 20% 이상 강화 필요.",
                                 "normalized_interpretation": "Frame 평균 5.850/14.254 SPEC 내. Yoke 평균 5.807/14.224 SPEC 하한(5.82/14.23) 미만 - FAIL. G3-D Gap 평균 0.060 mm > G2-A 0.040 mm (50% 큼). Cavity별 NG: A 2.48%, B 2.11%, C 2.84%, D 2.71%, E 2.77%, F 1.53%, G 1.14%로 F/G 최저, C/D/E 최고. 같은 이벤트 Normal 행은 없음 - 치수 중심 보고."}},
    "hints": {"hint_1": {"check_item": "Yoke 폭/길이 공차를 20% 이상 강화하고 cavity별 Frame+Yoke Gap 재검토.",
                          "reason": "Yoke 폭 평균 5.807 < 하한 5.82, 길이 평균 14.224 < 하한 14.23 (FAIL). 불량 多 G3-D Gap 평균 0.060 vs G2-A 0.040 (50% 큼). Cavity NG 순위(D/C/E 상위)와 일치."}},
    "log": {"assumptions": ["Summary 본문 기준 Yoke FAIL 판정.",
                            "Cavity별 NG는 MOLD FRAME 3 TOTAL 행 사용."],
            "warnings": ["같은 이벤트 Normal 행이 없어 절대 NG rate 비교 불가, 치수 근거만."],
            "decision_rationale": "Frame은 SPEC 내, Yoke는 SPEC 미만, 불량 多 cavity의 Gap이 더 큼 - 치수 중심 evidence. Function 표는 cavity-mix 보조 자료. before_after_dimension 분류."}
}
ds5_tr_vi = {
    "document": {"title": "Báo cáo kiểm tra kích thước Frame, lắp Frame+Yoke và Gap BRS-161014/161016",
                  "purpose": "Tìm nguyên nhân NG hearing bằng cách đo kích thước Frame, Yoke và Gap khi lắp ráp.",
                  "content": ["Đo Frame trên máy 2D (Spec width 5.82~5.86, length 14.23~14.27).",
                              "Đo Frame+Yoke sau ass'y trên máy 2D.",
                              "Đo Gap Frame+Yoke theo cavity (G3-D NG cao vs G2-A NG thấp).",
                              "Tách cavity Mold Frame 3 (A~G) và đo function."]},
    "conclusions": {"concl_1": {"topic": "Kích thước Frame/Yoke và Gap dẫn tới biến động NG Hearing",
                                 "statement_from_report": "Frame trong dung sai, một số YK dưới spec; cavity NG cao G3-D Gap 0.05~0.07 vượt spec; cavity NG thấp G2-A Gap 0.029~0.055 trong spec. Cần thắt chặt dung sai vật liệu >=20%.",
                                 "normalized_interpretation": "Kích thước Frame AVG 5.850/14.254 trong spec. Yoke AVG 5.807/14.224 thấp hơn cận dưới (5.82/14.23) - FAIL. Gap G3-D AVG 0.060 mm so với G2-A 0.040 mm (lớn hơn 50%). NG function theo cavity: A 2.48%, B 2.11%, C 2.84%, D 2.71%, E 2.77%, F 1.53%, G 1.14% - F/G thấp nhất, C/D/E cao nhất. Không có dòng Normal cùng sự kiện - báo cáo dựa vào kích thước."}},
    "hints": {"hint_1": {"check_item": "Thắt dung sai Yoke width/length (>=20%) và đo lại Gap Frame+Yoke theo cavity.",
                          "reason": "Yoke width AVG 5.807 < 5.82 (FAIL); length AVG 14.224 < 14.23 (FAIL). Gap cavity G3-D 0.060 vs G2-A 0.040 (lớn hơn 50%). NG cavity theo thứ tự cùng chiều: D/C/E cao, F/G thấp."}},
    "log": {"assumptions": ["Phán định FAIL Yoke dựa trên AVG so với cận dưới.",
                            "Tỷ lệ NG cavity lấy từ hàng TOTAL MOLD FRAME 3."],
            "warnings": ["Không có dòng Normal cùng sự kiện để so sánh tuyệt đối, chỉ dựa vào kích thước."],
            "decision_rationale": "Bằng chứng chính là kích thước: Frame đạt, Yoke không đạt, Gap cao ở cavity NG cao. Bảng function là hỗ trợ. Phân loại before_after_dimension."}
}

# ---------- Dataset 6 ----------
# 32. BRS-161016 Report test increase bonding VP+CD  21.4.2024 - normal_comparison
DS6_NAME = "32. BRS-161016 Report test increase bonding VP+CD  21.4.2024"
ds6_result = {
    "schema_version": "0.1",
    "document": {
        "document_id": "doc_ds6",
        "source_file": DS6_NAME,
        "source_sheet": "5.3",
        "title": "BRS-161016 Report Test Increase Bonding VP+CD",
        "model": "BRS-161016",
        "report_date": "2024-04-21",
        "department": "ME",
        "marker": "Thuy",
        "line": "E2-3B / C2-3A",
        "report_type": "normal_comparison",
        "primary_defect": {"canonical_name": "NG Hearing Noise",
                           "aliases_in_document": ["NG hearing"]},
        "related_defects": ["NG Hearing Noise"],
        "parts": ["VP", "CD"],
        "processes": ["VP+CD Bonding"],
        "purpose": "Improve NG hearing on SPK and Module line by increasing bond amount at the 4 corners of VP+CD.",
        "content": [
            "Reduce bonding corner speed from 30 to 20/25/35/40/55 -> increase bond amount at 4 corners.",
            "Make sample, check function on SPK line, move to Module line."
        ],
        "source_cells": {"title": ["5.3!B2"], "date": ["5.3!K3"], "purpose": ["5.3!A4"], "content": ["5.3!A6"]}
    },
    "test_conditions": [
        {"condition_id": "cond_speed35", "condition_group": "Bond speed change",
         "line": "E2-3B", "process": "VP+CD Bonding", "changed_factor": "Bonding corner speed",
         "before_value": "Speed 30 (Normal)", "after_value": "Speed 35",
         "unit": None, "machine": "Bonding machine", "jig": None, "material_lot": None,
         "supplier": None, "dry_time_sec": None, "temperature": None, "pressure": None,
         "bond_amount": None, "uv_energy": None,
         "source_file": DS6_NAME, "sheet_name": "5.3", "source_cells": ["5.3!A10"]},
        {"condition_id": "cond_speed40", "condition_group": "Bond speed change",
         "line": "E2-3B", "process": "VP+CD Bonding", "changed_factor": "Bonding corner speed",
         "before_value": "Speed 30 (Normal)", "after_value": "Speed 40",
         "unit": None, "machine": "Bonding machine", "jig": None, "material_lot": None,
         "supplier": None, "dry_time_sec": None, "temperature": None, "pressure": None,
         "bond_amount": None, "uv_energy": None,
         "source_file": DS6_NAME, "sheet_name": "5.3", "source_cells": ["5.3!A11"]},
        {"condition_id": "cond_speed55", "condition_group": "Bond speed change",
         "line": "C2-3A", "process": "VP+CD Bonding", "changed_factor": "Bonding corner speed",
         "before_value": "Speed 70 (Normal)", "after_value": "Speed 55",
         "unit": None, "machine": "Bonding machine", "jig": None, "material_lot": None,
         "supplier": None, "dry_time_sec": None, "temperature": None, "pressure": None,
         "bond_amount": None, "uv_energy": None,
         "source_file": DS6_NAME, "sheet_name": "5.3", "source_cells": ["5.3!A13"]}
    ],
    "results": [
        {"result_id": "res_e2_speed35", "condition_id": "cond_speed35",
         "measurement_type": "Function", "condition_group": "Reduce bond corner speed 35",
         "date": "2024-05-07", "line": "E2-3B",
         "input_count": 225, "ok_count": 214, "ng_count": 11,
         "ng_rate_decimal": 0.049, "ng_rate_percent": 4.9,
         "metric_name": "Function Total NG rate", "metric_value": 4.9, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"NG Sigma THD": {"count": 1, "rate": 0.4},
                          "NG Hearing Noise": {"count": 10, "rate": 4.4}},
         "source_file": DS6_NAME, "sheet_name": "5.3", "source_cells": ["5.3!N12"]},
        {"result_id": "res_e2_speed40", "condition_id": "cond_speed40",
         "measurement_type": "Function", "condition_group": "Reduce bond corner speed 40",
         "date": "2024-05-07", "line": "E2-3B",
         "input_count": 212, "ok_count": 205, "ng_count": 7,
         "ng_rate_decimal": 0.033, "ng_rate_percent": 3.3,
         "metric_name": "Function Total NG rate", "metric_value": 3.3, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"NG Hearing Noise": {"count": 7, "rate": 3.3}},
         "source_file": DS6_NAME, "sheet_name": "5.3", "source_cells": ["5.3!N14"]},
        {"result_id": "res_e2_normal", "condition_id": None,
         "measurement_type": "Function", "condition_group": "Normal speed 30 (E2-3B)",
         "date": "2024-05-07", "line": "E2-3B",
         "input_count": 796, "ok_count": 773, "ng_count": 23,
         "ng_rate_decimal": 0.029, "ng_rate_percent": 2.9,
         "metric_name": "Function Total NG rate (Normal)", "metric_value": 2.9, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"NG Sigma SPL+THD": {"count": 1, "rate": 0.1},
                          "NG Hearing Noise": {"count": 22, "rate": 2.8}},
         "source_file": DS6_NAME, "sheet_name": "5.3", "source_cells": ["5.3!N16"]},
        {"result_id": "res_c2_speed55", "condition_id": "cond_speed55",
         "measurement_type": "Function", "condition_group": "Reduce bond corner speed 55",
         "date": "2024-05-07", "line": "C2-3A",
         "input_count": 560, "ok_count": 536, "ng_count": 24,
         "ng_rate_decimal": 0.043, "ng_rate_percent": 4.3,
         "metric_name": "Function Total NG rate", "metric_value": 4.3, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"NG Hearing Noise": {"count": 24, "rate": 4.3}},
         "source_file": DS6_NAME, "sheet_name": "5.3", "source_cells": ["5.3!N18"]},
        {"result_id": "res_c2_normal", "condition_id": None,
         "measurement_type": "Function", "condition_group": "Normal speed 70 (C2-3A)",
         "date": "2024-05-07", "line": "C2-3A",
         "input_count": 560, "ok_count": 542, "ng_count": 18,
         "ng_rate_decimal": 0.032, "ng_rate_percent": 3.2,
         "metric_name": "Function Total NG rate (Normal)", "metric_value": 3.2, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"NG Hearing Noise": {"count": 18, "rate": 3.2}},
         "source_file": DS6_NAME, "sheet_name": "5.3", "source_cells": ["5.3!N20"]}
    ],
    "conclusions": [
        {"conclusion_id": "concl_1",
         "topic": "Increase bond corner amount via reduced speed - effect on Function NG",
         "statement_from_report": "Reduce bonding corner speed -> increase bond amount at 4 corners. Compare NG with Normal.",
         "normalized_interpretation": "E2-3B speed 35 4.9% vs Normal 30 2.9% = 1.69x (69.0% worse). E2-3B speed 40 3.3% vs Normal 2.9% = 1.14x (13.8% worse). C2-3A speed 55 4.3% vs Normal 70 3.2% = 1.34x (34.4% worse). NG Hearing Noise dominates all groups. Increasing bond corner amount did NOT reduce NG Hearing Noise in this run.",
         "source_file": DS6_NAME, "sheet_name": "5.3", "source_cells": ["5.3!N12:N20"]}
    ],
    "troubleshooting_index": {
        "defect_name": "NG Hearing Noise",
        "when_user_asks": ["Does increasing bond corner amount reduce NG hearing noise?"],
        "suggested_checks": [
            {"hint_id": "hint_1",
             "check_item": "Re-evaluate VP+CD bonding speed vs noise correlation; do not increase bond amount blindly.",
             "reason": "All three reduced-speed (higher bond amount) test conditions show worse function NG than same-event Normal: +69.0%, +13.8%, +34.4%. NG Hearing Noise is the dominant NG mode in every group.",
             "evidence_strength": "medium", "related_process": "VP+CD Bonding", "related_part": "VP/CD",
             "source_file": DS6_NAME, "sheet_name": "5.3", "source_cells": ["5.3!N12:N20"]}
        ],
        "limitations": ["Bond amount in mg measured but speed differences alone may not directly equal bond-amount differences when only corner speed changes."]
    },
    "ai_extraction_log": {
        "confidence": 0.7,
        "assumptions": ["Same-day Normal rows used as baseline per line.",
                        "Bond amount values from setup table are per dispenser side, not used for direct NG comparison."],
        "warnings": ["Higher bond amount correlated with WORSE function NG in this test - improvement hypothesis not supported."],
        "decision_rationale": "Both lines have Normal rows in the same date/event. Relative change is positive (worse) on all three test conditions; report classified normal_comparison. Hearing Noise dominates, so root cause likely not at the bond amount only."
    }
}
ds6_tr_en = {
    "document": {"title": ds6_result["document"]["title"], "purpose": ds6_result["document"]["purpose"], "content": ds6_result["document"]["content"]},
    "conclusions": {"concl_1": {"topic": ds6_result["conclusions"][0]["topic"],
                                 "statement_from_report": ds6_result["conclusions"][0]["statement_from_report"],
                                 "normalized_interpretation": ds6_result["conclusions"][0]["normalized_interpretation"]}},
    "hints": {"hint_1": {"check_item": ds6_result["troubleshooting_index"]["suggested_checks"][0]["check_item"],
                          "reason": ds6_result["troubleshooting_index"]["suggested_checks"][0]["reason"]}},
    "log": {"assumptions": ds6_result["ai_extraction_log"]["assumptions"],
            "warnings": ds6_result["ai_extraction_log"]["warnings"],
            "decision_rationale": ds6_result["ai_extraction_log"]["decision_rationale"]}
}
ds6_tr_ko = {
    "document": {"title": "BRS-161016 VP+CD 4 모서리 본드 증량 시험 리포트",
                  "purpose": "SPK / Module 라인 NG hearing 개선을 위해 VP+CD 4 모서리 본드 증량 시험.",
                  "content": ["코너 본드 속도를 30에서 20/25/35/40/55로 낮춰 4 코너 본드 증량.",
                              "샘플 제작 후 SPK 라인 function 점검, Module 라인 이동 테스트."]},
    "conclusions": {"concl_1": {"topic": "코너 본드 증량(속도 감소)의 Function NG 영향",
                                 "statement_from_report": "코너 본드 속도 감소로 4 코너 본드 증량. Normal과 NG 비교.",
                                 "normalized_interpretation": "E2-3B 속도 35 4.9% vs Normal 30 2.9% = 1.69배 (69.0% 악화). 속도 40 3.3% vs Normal 2.9% = 1.14배 (13.8% 악화). C2-3A 속도 55 4.3% vs Normal 70 3.2% = 1.34배 (34.4% 악화). 모든 그룹 NG Hearing Noise 지배. 본드 증량은 이번 시험에서 NG Hearing Noise 감소시키지 못함."}},
    "hints": {"hint_1": {"check_item": "VP+CD bond 속도-Noise 상관을 재평가, 본드 증량은 무조건적 처방이 아님.",
                          "reason": "세 가지 속도 감소(본드 증량) 조건 모두 Normal 대비 +69.0%, +13.8%, +34.4% 악화. 모든 그룹에서 NG Hearing Noise가 지배."}},
    "log": {"assumptions": ["같은 날 Normal 행을 라인별 baseline으로 사용.",
                            "Bond amount(mg)는 dispenser별 값이라 NG 직접 비교에 미사용."],
            "warnings": ["본드 증량이 오히려 Function NG를 악화시킴 - 가설 미지지."],
            "decision_rationale": "양쪽 라인 모두 동일 이벤트에 Normal 행 존재, normal_comparison. 세 test 조건 모두 +상대 변화로 악화. Hearing Noise 지배라 본드 증량이 근본 원인 해소가 아님."}
}
ds6_tr_vi = {
    "document": {"title": "BRS-161016 Báo cáo test tăng keo 4 góc VP+CD",
                  "purpose": "Cải thiện NG hearing trên line SPK và Module bằng cách tăng keo 4 góc VP+CD.",
                  "content": ["Giảm tốc độ bonding corner từ 30 xuống 20/25/35/40/55 để tăng keo 4 góc.",
                              "Sản xuất mẫu, kiểm tra function trên line SPK, chuyển sang line Module."]},
    "conclusions": {"concl_1": {"topic": "Tăng keo góc (giảm tốc) tác động đến NG Function",
                                 "statement_from_report": "Giảm tốc bonding corner để tăng keo 4 góc. So sánh NG với Normal.",
                                 "normalized_interpretation": "E2-3B speed 35 4.9% so với Normal 30 2.9% = 1.69x (xấu hơn 69.0%). E2-3B speed 40 3.3% so với 2.9% = 1.14x (xấu 13.8%). C2-3A speed 55 4.3% so với Normal 70 3.2% = 1.34x (xấu 34.4%). NG Hearing Noise chi phối toàn bộ. Tăng keo không giảm NG Hearing Noise trong lần test này."}},
    "hints": {"hint_1": {"check_item": "Đánh giá lại quan hệ tốc độ bonding VP+CD và Noise, không nên tăng keo mù quáng.",
                          "reason": "Cả ba điều kiện giảm tốc (tăng keo) đều cho NG function xấu hơn Normal: +69.0%, +13.8%, +34.4%. NG Hearing Noise vẫn là loại NG chính."}},
    "log": {"assumptions": ["Dùng dòng Normal cùng ngày làm baseline cho từng line.",
                            "Bond amount mg là theo từng dispenser, không dùng để so trực tiếp."],
            "warnings": ["Tăng keo lại làm NG function xấu hơn, giả thuyết không được hỗ trợ."],
            "decision_rationale": "Cả hai line đều có Normal cùng sự kiện nên normal_comparison. Cả ba test đều thay đổi tương đối dương (xấu hơn). NG Hearing Noise chiếm chính nên tăng keo không giải quyết gốc rễ."}
}

# ---------- Dataset 7 ----------
# 32. BRS-201506 Report Test new wire from KR date 16.03.2024 - normal_comparison (multi)
DS7_NAME = "32. BRS-201506 Report Test new wire from KR date 16.03.2024"
ds7_result = {
    "schema_version": "0.1",
    "document": {
        "document_id": "doc_ds7",
        "source_file": DS7_NAME,
        "source_sheet": "Test Coil 2nd",
        "title": "Report Checking and Test New Wire from Korea 0.095mm & 0.096mm BRS-201506",
        "model": "BRS-201506",
        "report_date": "2024-03-16",
        "department": "ME",
        "marker": "Thao",
        "line": "",
        "report_type": "normal_comparison",
        "primary_defect": {"canonical_name": "Weak Solder",
                           "aliases_in_document": ["NG weak solder", "weak solder"]},
        "related_defects": ["NG Hearing Noise", "NG Hearing Touch"],
        "parts": ["Coil", "Wire"],
        "processes": ["Coil Winding", "Spot Welding", "Function"],
        "purpose": "Test new coil wires 0.095mm and 0.096mm from Korea vs Normal 0.097mm for SPK function and module line NG hearing.",
        "content": [
            "Use 0.095mm/0.096mm wires; reduce 1 turn (58 -> 57).",
            "Check coil DCR, 3D dimension, height.",
            "Make samples; check function; check NTI on OK samples; compare with normal wire.",
            "50 pcs per type."
        ],
        "source_cells": {"title": ["Test Coil 2nd!B2"], "date": ["Test Coil 2nd!T3"], "purpose": ["Test Coil 2nd!A4"], "content": ["Test Coil 2nd!A6"]}
    },
    "test_conditions": [
        {"condition_id": "cond_w095", "condition_group": "Wire diameter",
         "line": "", "process": "Coil Winding", "changed_factor": "Wire diameter / turn count",
         "before_value": "0.097mm, 58 turn", "after_value": "0.095mm, 57 turn",
         "unit": "mm", "machine": None, "jig": None, "material_lot": None,
         "supplier": "Korea (new wire)", "dry_time_sec": None, "temperature": None, "pressure": None,
         "bond_amount": None, "uv_energy": None,
         "source_file": DS7_NAME, "sheet_name": "Test Coil 2nd", "source_cells": ["Test Coil 2nd!A11"]},
        {"condition_id": "cond_w096", "condition_group": "Wire diameter",
         "line": "", "process": "Coil Winding", "changed_factor": "Wire diameter / turn count",
         "before_value": "0.097mm, 58 turn", "after_value": "0.096mm, 57 turn",
         "unit": "mm", "machine": None, "jig": None, "material_lot": None,
         "supplier": "Korea (new wire)", "dry_time_sec": None, "temperature": None, "pressure": None,
         "bond_amount": None, "uv_energy": None,
         "source_file": DS7_NAME, "sheet_name": "Test Coil 2nd", "source_cells": ["Test Coil 2nd!A12"]}
    ],
    "results": [
        {"result_id": "res_solder_095", "condition_id": "cond_w095",
         "measurement_type": "Spot Welding Vision", "condition_group": "Coil 0.095mm",
         "date": "2024-03-16", "line": "",
         "input_count": 92, "ok_count": 66, "ng_count": 26,
         "ng_rate_decimal": 0.283, "ng_rate_percent": 28.3,
         "metric_name": "Weak Solder NG rate", "metric_value": 28.3, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"Weak Solder": {"count": 26, "rate": 28.3}},
         "source_file": DS7_NAME, "sheet_name": "Test Coil 2nd", "source_cells": ["Test Coil 2nd!H52"]},
        {"result_id": "res_solder_096", "condition_id": "cond_w096",
         "measurement_type": "Spot Welding Vision", "condition_group": "Coil 0.096mm",
         "date": "2024-03-16", "line": "",
         "input_count": 55, "ok_count": 34, "ng_count": 21,
         "ng_rate_decimal": 0.382, "ng_rate_percent": 38.2,
         "metric_name": "Weak Solder NG rate", "metric_value": 38.2, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"Weak Solder": {"count": 21, "rate": 38.2}},
         "source_file": DS7_NAME, "sheet_name": "Test Coil 2nd", "source_cells": ["Test Coil 2nd!H53"]},
        {"result_id": "res_solder_normal", "condition_id": None,
         "measurement_type": "Spot Welding Vision", "condition_group": "Normal 0.097mm",
         "date": "2024-03-16", "line": "",
         "input_count": 63, "ok_count": 58, "ng_count": 5,
         "ng_rate_decimal": 0.079, "ng_rate_percent": 7.9,
         "metric_name": "Weak Solder NG rate (Normal)", "metric_value": 7.9, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"Weak Solder": {"count": 5, "rate": 7.9}},
         "source_file": DS7_NAME, "sheet_name": "Test Coil 2nd", "source_cells": ["Test Coil 2nd!H54"]},
        {"result_id": "res_func_095_316", "condition_id": "cond_w095",
         "measurement_type": "Function", "condition_group": "Coil 0.095mm",
         "date": "2024-03-13", "line": "",
         "input_count": 80, "ok_count": 75, "ng_count": 5,
         "ng_rate_decimal": 0.062, "ng_rate_percent": 6.2,
         "metric_name": "Function Total NG rate", "metric_value": 6.2, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"NG Hearing Noise": {"count": 2, "rate": 40.0},
                          "NG Hearing Touch": {"count": 3, "rate": 60.0}},
         "source_file": DS7_NAME, "sheet_name": "Test Coil 2nd", "source_cells": ["Test Coil 2nd!K58"]},
        {"result_id": "res_func_096_316", "condition_id": "cond_w096",
         "measurement_type": "Function", "condition_group": "Coil 0.096mm",
         "date": "2024-03-13", "line": "",
         "input_count": 44, "ok_count": 43, "ng_count": 1,
         "ng_rate_decimal": 0.023, "ng_rate_percent": 2.3,
         "metric_name": "Function Total NG rate", "metric_value": 2.3, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"NG Hearing Touch": {"count": 1, "rate": 100.0}},
         "source_file": DS7_NAME, "sheet_name": "Test Coil 2nd", "source_cells": ["Test Coil 2nd!K60"]},
        {"result_id": "res_func_normal_316", "condition_id": None,
         "measurement_type": "Function", "condition_group": "Normal 0.097mm",
         "date": "2024-03-13", "line": "",
         "input_count": 50, "ok_count": 43, "ng_count": 7,
         "ng_rate_decimal": 0.14, "ng_rate_percent": 14.0,
         "metric_name": "Function Total NG rate (Normal)", "metric_value": 14.0, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"NG Hearing Noise": {"count": 1, "rate": 14.3},
                          "NG Hearing Touch": {"count": 6, "rate": 85.7}},
         "source_file": DS7_NAME, "sheet_name": "Test Coil 2nd", "source_cells": ["Test Coil 2nd!K62"]}
    ],
    "conclusions": [
        {"conclusion_id": "concl_1",
         "topic": "New wire 0.095/0.096mm increases Weak Solder but reduces Function NG",
         "statement_from_report": "When use new wire happen NG weak solder so high.",
         "normalized_interpretation": "Weak Solder NG: 0.095mm 28.3% vs Normal 0.097mm 7.9% = 3.58x (258.2% worse); 0.096mm 38.2% vs 7.9% = 4.84x (383.5% worse). Function NG (13-Mar): 0.095mm 6.2% vs Normal 14.0% = 0.443x (55.7% improved); 0.096mm 2.3% vs 14.0% = 0.164x (83.6% improved). Thinner wire reduces hearing NG but causes severe Weak Solder issue.",
         "source_file": DS7_NAME, "sheet_name": "Test Coil 2nd", "source_cells": ["Test Coil 2nd!H52:K62"]}
    ],
    "troubleshooting_index": {
        "defect_name": "Weak Solder vs NG Hearing trade-off",
        "when_user_asks": ["What is the impact of thinner coil wire (0.095/0.096mm) on hearing and solder?"],
        "suggested_checks": [
            {"hint_id": "hint_1",
             "check_item": "Verify spot welding parameters when changing wire diameter from 0.097mm to 0.095/0.096mm.",
             "reason": "Weak Solder NG rises sharply (28.3% / 38.2% vs Normal 7.9%, +258% / +384%), while hearing NG drops (6.2% / 2.3% vs 14.0%, -55.7% / -83.6%). Spot welding setup must be tuned for thinner wire.",
             "evidence_strength": "high", "related_process": "Spot Welding", "related_part": "Coil/Wire",
             "source_file": DS7_NAME, "sheet_name": "Test Coil 2nd", "source_cells": ["Test Coil 2nd!H52:K62"]}
        ],
        "limitations": ["Function sample size per type is small (44-80); some 19-Mar function rows lack NG percentages (#DIV/0!)."]
    },
    "ai_extraction_log": {
        "confidence": 0.7,
        "assumptions": ["Used 16-Mar weak solder vision row and 13-Mar function row as primary baselines (largest sample).",
                        "Dimensional check tables (length/width/angle) recorded for completeness but not as NG."],
        "warnings": ["19-Mar function table has #DIV/0! for several NG breakdown percentages.",
                     "NTI SPL/THD/IMP tables not normalized (large frequency arrays)."],
        "decision_rationale": "Same-event Normal 0.097mm row exists for both welding and function tables. Relative changes calculated. Thinner wire shows trade-off: solder NG worse, hearing NG improved. normal_comparison classification."
    }
}
ds7_tr_en = {
    "document": {"title": ds7_result["document"]["title"], "purpose": ds7_result["document"]["purpose"], "content": ds7_result["document"]["content"]},
    "conclusions": {"concl_1": {"topic": ds7_result["conclusions"][0]["topic"],
                                 "statement_from_report": ds7_result["conclusions"][0]["statement_from_report"],
                                 "normalized_interpretation": ds7_result["conclusions"][0]["normalized_interpretation"]}},
    "hints": {"hint_1": {"check_item": ds7_result["troubleshooting_index"]["suggested_checks"][0]["check_item"],
                          "reason": ds7_result["troubleshooting_index"]["suggested_checks"][0]["reason"]}},
    "log": {"assumptions": ds7_result["ai_extraction_log"]["assumptions"],
            "warnings": ds7_result["ai_extraction_log"]["warnings"],
            "decision_rationale": ds7_result["ai_extraction_log"]["decision_rationale"]}
}
ds7_tr_ko = {
    "document": {"title": "BRS-201506 한국산 신규 와이어 0.095/0.096mm 시험 리포트",
                 "purpose": "한국산 신규 코일 와이어 0.095mm / 0.096mm 와 Normal 0.097mm 비교, SPK function과 Module 라인 NG hearing 검토.",
                 "content": ["0.095/0.096mm 와이어 사용, 권선수 58 → 57로 감소.",
                             "코일 DCR, 3D 치수, 높이 검사.",
                             "샘플 제작 후 function 검사, OK 샘플 NTI 측정, normal 와이어와 비교.",
                             "타입별 50pcs."]},
    "conclusions": {"concl_1": {"topic": "신규 와이어 0.095/0.096mm: Weak Solder 증가, Function NG 감소",
                                 "statement_from_report": "신규 와이어 사용 시 NG weak solder 매우 높음.",
                                 "normalized_interpretation": "Weak Solder NG: 0.095mm 28.3% vs Normal 0.097mm 7.9% = 3.58배 (258.2% 악화), 0.096mm 38.2% vs 7.9% = 4.84배 (383.5% 악화). Function NG(13-Mar): 0.095mm 6.2% vs Normal 14.0% = 0.443배 (55.7% 개선), 0.096mm 2.3% vs 14.0% = 0.164배 (83.6% 개선). 얇은 와이어는 hearing NG 감소, 그러나 Weak Solder 심각."}},
    "hints": {"hint_1": {"check_item": "와이어 직경을 0.097mm → 0.095/0.096mm로 변경 시 spot welding 파라미터 재검토.",
                          "reason": "Weak Solder NG가 28.3%/38.2%로 Normal 7.9% 대비 +258% / +384% 악화, 반면 hearing NG는 6.2%/2.3%로 14.0% 대비 -55.7% / -83.6% 개선. Spot welding 조건을 얇은 와이어용으로 조정 필요."}},
    "log": {"assumptions": ["16-Mar weak solder 비전 행과 13-Mar function 행을 표본 큰 baseline으로 사용.",
                            "치수 표(길이/폭/각도)는 NG는 아니지만 기록."],
            "warnings": ["19-Mar function 표 일부 NG breakdown은 #DIV/0!.",
                          "NTI SPL/THD/IMP 표는 정규화 미수행 (대량 주파수 데이터)."],
            "decision_rationale": "Spot welding과 function 모두 같은 이벤트 Normal 0.097mm 행 존재. 상대 변화 산정. 얇은 와이어는 trade-off (solder 악화, hearing 개선). normal_comparison."}
}
ds7_tr_vi = {
    "document": {"title": "BRS-201506 Báo cáo test wire mới từ Hàn 0.095mm & 0.096mm",
                 "purpose": "So sánh wire 0.095mm/0.096mm với Normal 0.097mm để cải thiện NG hearing trên SPK và line Module.",
                 "content": ["Dùng wire 0.095/0.096mm, giảm 1 vòng (58 → 57).",
                             "Kiểm tra DCR, kích thước 3D, chiều cao coil.",
                             "Sản xuất mẫu, đo function, đo NTI mẫu OK, so với wire thường.",
                             "Mỗi loại 50pcs."]},
    "conclusions": {"concl_1": {"topic": "Wire mới 0.095/0.096mm: Weak Solder tăng, NG function giảm",
                                 "statement_from_report": "Khi dùng wire mới NG weak solder rất cao.",
                                 "normalized_interpretation": "Weak Solder NG: 0.095mm 28.3% so với Normal 0.097mm 7.9% = 3.58x (xấu 258.2%); 0.096mm 38.2% so với 7.9% = 4.84x (xấu 383.5%). NG function (13-Mar): 0.095mm 6.2% so với Normal 14.0% = 0.443x (cải thiện 55.7%); 0.096mm 2.3% so với 14.0% = 0.164x (cải thiện 83.6%). Wire mỏng giảm NG hearing nhưng gây Weak Solder nghiêm trọng."}},
    "hints": {"hint_1": {"check_item": "Khi đổi đường kính wire từ 0.097mm sang 0.095/0.096mm cần chỉnh thông số spot welding.",
                          "reason": "Weak Solder NG tăng mạnh (28.3% / 38.2% so với Normal 7.9%, +258% / +384%), NG hearing giảm (6.2% / 2.3% so với 14.0%, -55.7% / -83.6%). Cần tinh chỉnh spot welding cho wire mỏng."}},
    "log": {"assumptions": ["Dùng hàng spot welding 16-Mar và hàng function 13-Mar (cỡ mẫu lớn nhất) làm baseline.",
                            "Bảng kích thước (length/width/angle) lưu để tham khảo, không phải NG."],
            "warnings": ["Bảng function 19-Mar có #DIV/0! ở một số NG breakdown.",
                          "Bảng NTI SPL/THD/IMP không chuẩn hoá (dữ liệu tần số lớn)."],
            "decision_rationale": "Cả spot welding và function đều có hàng Normal 0.097mm cùng sự kiện. Có thể tính tỷ lệ tương đối. Wire mỏng có trade-off (solder xấu, hearing tốt). Phân loại normal_comparison."}
}

# ---------- Dataset 8 (duplicate of #3 file name but separate dataset row) ----------
# 32. BRS-201506 Report test New base JIG improvement NG VP laser cutting offset- date 15.3.2024
# Different from DS3: function table includes 3-18 G1/G2/G3 vs Normal, 3-19 G1/G2/G3 vs Normal, then Totals.
DS8_NAME = "32. BRS-201506 Report test New base JIG improvement NG VP laser cutting offset- date 15.3.2024"
ds8_result = {
    "schema_version": "0.1",
    "document": {
        "document_id": "doc_ds8",
        "source_file": DS8_NAME,
        "source_sheet": "Base",
        "title": "BRS-201506 Report Test New Base JIG Improvement NG VP Laser Cutting Offset (Vision VP Sub 5)",
        "model": "BRS-201506",
        "report_date": "2024-03-15",
        "department": "ME",
        "marker": "Nhung",
        "line": "",
        "report_type": "normal_comparison",
        "primary_defect": {"canonical_name": "VP Laser Cutting Offset",
                           "aliases_in_document": ["VP laser cutting offset"]},
        "related_defects": ["VP Damage", "VP Deform", "Particle", "Separation NG", "NG Hearing Noise", "NG Hearing Touch"],
        "parts": ["VP"],
        "processes": ["Sub VP Line", "Laser cutting", "VP Vision"],
        "purpose": "Check if VP vision before pickup with new base JIG can separate VP laser cutting offset by group.",
        "content": [
            "New base JIG test for 2 days, 2k each day.",
            "Normal flow: Laser cutting -> Pickup -> VP Vision.",
            "Test flow: Laser cutting -> VP Vision -> Pickup.",
            "1. Normal and test: check problem.",
            "2. Check VP damage after pickup.",
            "3. Separate Limit G1/G2/G3 -> main line test."
        ],
        "source_cells": {"title": ["Base!B2"], "date": ["Base!S3"], "purpose": ["Base!A4"], "content": ["Base!A6"]}
    },
    "test_conditions": [
        {"condition_id": "cond_1", "condition_group": "Base JIG flow",
         "line": "", "process": "Sub VP Line", "changed_factor": "VP vision position",
         "before_value": "Laser cutting -> Pickup -> VP Vision",
         "after_value": "Laser cutting -> VP Vision -> Pickup",
         "unit": None, "machine": "Vision VP Sub 5", "jig": "New base JIG", "material_lot": None,
         "supplier": None, "dry_time_sec": None, "temperature": None, "pressure": None,
         "bond_amount": None, "uv_energy": None,
         "source_file": DS8_NAME, "sheet_name": "Base", "source_cells": ["Base!A7"]}
    ],
    "results": [
        # Sub VP line
        {"result_id": "res_sub_test_315", "condition_id": "cond_1",
         "measurement_type": "Sub VP", "condition_group": "Test",
         "date": "2024-03-15", "line": "",
         "input_count": 1996, "ok_count": 1971, "ng_count": 25,
         "ng_rate_decimal": 0.0125, "ng_rate_percent": 1.3,
         "metric_name": "Sub VP Total NG rate", "metric_value": 1.3, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"VP Damage": {"count": 2}, "VP Deform": {"count": 18}, "Particle": {"count": 2},
                          "Laser Cutting Burr": {"count": 3}, "Separation VP": {"count": 0}},
         "source_file": DS8_NAME, "sheet_name": "Base", "source_cells": ["Base!K15"]},
        {"result_id": "res_sub_normal_315", "condition_id": None,
         "measurement_type": "Sub VP", "condition_group": "Normal",
         "date": "2024-03-15", "line": "",
         "input_count": 2880, "ok_count": 2851, "ng_count": 29,
         "ng_rate_decimal": 0.0101, "ng_rate_percent": 1.0,
         "metric_name": "Sub VP Total NG rate (Normal)", "metric_value": 1.0, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"VP Damage": {"count": 2}, "VP Deform": {"count": 20}, "Particle": {"count": 3},
                          "Laser Cutting Burr": {"count": 4}, "Separation VP": {"count": 0}},
         "source_file": DS8_NAME, "sheet_name": "Base", "source_cells": ["Base!K16"]},
        # Function totals
        {"result_id": "res_func_g1", "condition_id": "cond_1",
         "measurement_type": "Function", "condition_group": "Test G1 Total",
         "date": "", "line": "",
         "input_count": 1902, "ok_count": 1799, "ng_count": 103,
         "ng_rate_decimal": 0.054, "ng_rate_percent": 5.4,
         "metric_name": "Function Total NG rate (G1)", "metric_value": 5.4, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"NG Sigma SPL": {"count": 2, "rate": 1.9},
                          "NG Sigma THD": {"count": 1, "rate": 1.0},
                          "NG Hearing Noise": {"count": 26, "rate": 25.2},
                          "NG Hearing Touch": {"count": 74, "rate": 71.8}},
         "source_file": DS8_NAME, "sheet_name": "Base", "source_cells": ["Base!K28"]},
        {"result_id": "res_func_g2", "condition_id": "cond_1",
         "measurement_type": "Function", "condition_group": "Test G2 Total",
         "date": "", "line": "",
         "input_count": 1906, "ok_count": 1838, "ng_count": 68,
         "ng_rate_decimal": 0.036, "ng_rate_percent": 3.6,
         "metric_name": "Function Total NG rate (G2)", "metric_value": 3.6, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"NG Sigma SPL": {"count": 1, "rate": 1.5},
                          "NG Hearing Noise": {"count": 13, "rate": 19.1},
                          "NG Hearing Touch": {"count": 54, "rate": 79.4}},
         "source_file": DS8_NAME, "sheet_name": "Base", "source_cells": ["Base!K30"]},
        {"result_id": "res_func_g3", "condition_id": "cond_1",
         "measurement_type": "Function", "condition_group": "Test G3 Total",
         "date": "", "line": "",
         "input_count": 69, "ok_count": 65, "ng_count": 4,
         "ng_rate_decimal": 0.058, "ng_rate_percent": 5.8,
         "metric_name": "Function Total NG rate (G3)", "metric_value": 5.8, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"NG Hearing Touch": {"count": 4, "rate": 100.0}},
         "source_file": DS8_NAME, "sheet_name": "Base", "source_cells": ["Base!K32"]},
        {"result_id": "res_func_normal", "condition_id": None,
         "measurement_type": "Function", "condition_group": "Normal Total",
         "date": "", "line": "",
         "input_count": 1398, "ok_count": 1320, "ng_count": 78,
         "ng_rate_decimal": 0.056, "ng_rate_percent": 5.6,
         "metric_name": "Function Total NG rate (Normal)", "metric_value": 5.6, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"NG Sigma SPL+THD": {"count": 1, "rate": 1.3},
                          "NG Hearing Noise": {"count": 23, "rate": 29.5},
                          "NG Hearing Touch": {"count": 54, "rate": 69.2}},
         "source_file": DS8_NAME, "sheet_name": "Base", "source_cells": ["Base!K34"]}
    ],
    "conclusions": [
        {"conclusion_id": "concl_1",
         "topic": "New base JIG groups: relative change vs Normal by group",
         "statement_from_report": "Group totals: G1 5.4%, G2 3.6%, G3 5.8%, Normal 5.6%.",
         "normalized_interpretation": "Sub VP NG Test vs Normal: 15-Mar 1.3%/1.0% (+30%), 16-Mar 1.4%/1.2% (+16.7%). Function NG: G1 5.4% vs Normal 5.6% = 0.964x (-3.6% improved); G2 3.6% vs 5.6% = 0.643x (-35.7% improved); G3 5.8% vs 5.6% = 1.036x (+3.6% worse). G2 (medium offset) shows the clearest improvement. NG Hearing Touch is dominant in every group.",
         "source_file": DS8_NAME, "sheet_name": "Base", "source_cells": ["Base!K28:K34"]}
    ],
    "troubleshooting_index": {
        "defect_name": "VP Laser Cutting Offset",
        "when_user_asks": ["Which separated group (G1/G2/G3) gives the best function NG with new base JIG?"],
        "suggested_checks": [
            {"hint_id": "hint_1",
             "check_item": "Adopt G2 limit window from new base JIG; review G3 limit threshold for Touch NG.",
             "reason": "G2 5.4% (G1) and 3.6% (G2) vs Normal 5.6% = 35.7% improved at G2; G3 (5.8%) similar to Normal but Touch NG = 100%. Sub-VP NG is similar between Test and Normal (~1%).",
             "evidence_strength": "medium", "related_process": "Sub VP Line", "related_part": "VP",
             "source_file": DS8_NAME, "sheet_name": "Base", "source_cells": ["Base!K28:K34"]}
        ],
        "limitations": ["G3 input only 69pcs; long-run validation needed."]
    },
    "ai_extraction_log": {
        "confidence": 0.7,
        "assumptions": ["G1/G2/G3 Test groups compared to combined Normal Total row.",
                        "Same-date Normal row used as baseline for Sub VP daily rows."],
        "warnings": ["G3 sample size is much smaller than G1/G2; %-share of NG mode based on a tiny denominator."],
        "decision_rationale": "Same-event Normal row exists for both Sub-VP and Function tables. Relative changes computed per group. G2 is the best-performing test group. Classification normal_comparison."
    }
}
ds8_tr_en = {
    "document": {"title": ds8_result["document"]["title"], "purpose": ds8_result["document"]["purpose"], "content": ds8_result["document"]["content"]},
    "conclusions": {"concl_1": {"topic": ds8_result["conclusions"][0]["topic"],
                                 "statement_from_report": ds8_result["conclusions"][0]["statement_from_report"],
                                 "normalized_interpretation": ds8_result["conclusions"][0]["normalized_interpretation"]}},
    "hints": {"hint_1": {"check_item": ds8_result["troubleshooting_index"]["suggested_checks"][0]["check_item"],
                          "reason": ds8_result["troubleshooting_index"]["suggested_checks"][0]["reason"]}},
    "log": {"assumptions": ds8_result["ai_extraction_log"]["assumptions"],
            "warnings": ds8_result["ai_extraction_log"]["warnings"],
            "decision_rationale": ds8_result["ai_extraction_log"]["decision_rationale"]}
}
ds8_tr_ko = {
    "document": {"title": "BRS-201506 신규 베이스 지그 VP laser cutting offset 개선 시험 (Sub VP Vision)",
                 "purpose": "신규 베이스 지그(VP vision -> pickup)로 VP laser cutting offset 그룹 분리 가능성 확인.",
                 "content": ["신규 베이스 지그 2일간 테스트, 일 2k.",
                             "Normal: Laser cutting -> Pickup -> VP Vision.",
                             "Test: Laser cutting -> VP Vision -> Pickup.",
                             "1. Normal/Test 문제 발생 여부 확인.",
                             "2. Pickup 후 VP damage 확인.",
                             "3. Separate Limit G1/G2/G3 -> main line test."]},
    "conclusions": {"concl_1": {"topic": "신규 베이스 지그 그룹별 Normal 대비 상대 변화",
                                 "statement_from_report": "그룹 총계: G1 5.4%, G2 3.6%, G3 5.8%, Normal 5.6%.",
                                 "normalized_interpretation": "Sub VP NG Test vs Normal: 15-Mar 1.3%/1.0% (+30%), 16-Mar 1.4%/1.2% (+16.7%). Function NG: G1 5.4% vs Normal 5.6% = 0.964배 (-3.6% 개선), G2 3.6% vs 5.6% = 0.643배 (-35.7% 개선), G3 5.8% vs 5.6% = 1.036배 (+3.6% 악화). 중간 offset G2가 최대 개선. 모든 그룹에서 NG Hearing Touch 지배."}},
    "hints": {"hint_1": {"check_item": "신규 베이스 지그의 G2 한계창을 채택, G3 한계를 Touch NG 관점에서 재검토.",
                          "reason": "G2 3.6% vs Normal 5.6% = 35.7% 개선. G3 5.8%는 Normal과 유사하나 Touch NG 100%. Sub-VP NG는 Test/Normal 약 1%대로 유사."}},
    "log": {"assumptions": ["G1/G2/G3 Test 그룹을 Normal Total 행과 비교.",
                            "Sub VP 일자별 행은 동일 일 Normal 행을 baseline으로 사용."],
            "warnings": ["G3 표본은 69pcs로 G1/G2 대비 매우 작아 NG 비중 해석은 주의 필요."],
            "decision_rationale": "Sub-VP와 Function 모두 같은 이벤트 Normal 행 존재, 그룹별 상대 변화 산정. G2가 최선 그룹. normal_comparison."}
}
ds8_tr_vi = {
    "document": {"title": "BRS-201506 Báo cáo test JIG base mới VP laser cutting offset (Vision VP Sub 5)",
                 "purpose": "Kiểm tra JIG base mới (VP vision trước pickup) có thể tách VP laser cutting offset theo nhóm hay không.",
                 "content": ["Test JIG base mới trong 2 ngày, mỗi ngày 2k.",
                             "Normal: Laser cutting -> Pickup -> VP Vision.",
                             "Test: Laser cutting -> VP Vision -> Pickup.",
                             "1. Kiểm tra phát sinh sự cố giữa Normal và Test.",
                             "2. Kiểm tra VP damage sau pickup.",
                             "3. Phân tách Limit G1/G2/G3 -> main line test."]},
    "conclusions": {"concl_1": {"topic": "Tỷ lệ thay đổi tương đối theo nhóm so với Normal",
                                 "statement_from_report": "Tổng các nhóm: G1 5.4%, G2 3.6%, G3 5.8%, Normal 5.6%.",
                                 "normalized_interpretation": "NG Sub VP Test vs Normal: 15-Mar 1.3%/1.0% (+30%), 16-Mar 1.4%/1.2% (+16.7%). NG function: G1 5.4% vs Normal 5.6% = 0.964x (cải thiện 3.6%), G2 3.6% vs 5.6% = 0.643x (cải thiện 35.7%), G3 5.8% vs 5.6% = 1.036x (xấu 3.6%). Nhóm G2 (offset trung bình) cải thiện rõ nhất. NG Hearing Touch chi phối toàn bộ nhóm."}},
    "hints": {"hint_1": {"check_item": "Áp dụng ngưỡng nhóm G2 của JIG base mới và xem lại ngưỡng G3 dưới góc độ NG Touch.",
                          "reason": "G2 3.6% so với Normal 5.6% = cải thiện 35.7%. G3 5.8% gần Normal nhưng Touch chiếm 100%. NG Sub-VP Test và Normal cùng ~1%."}},
    "log": {"assumptions": ["So sánh nhóm Test G1/G2/G3 với hàng Normal Total.",
                            "Hàng Normal cùng ngày dùng làm baseline cho Sub VP từng ngày."],
            "warnings": ["G3 chỉ 69pcs nhỏ hơn nhiều so với G1/G2 nên tỷ trọng NG cần cẩn thận."],
            "decision_rationale": "Cả Sub-VP và Function đều có hàng Normal cùng sự kiện, tính được tỷ lệ thay đổi theo nhóm. G2 là nhóm test tốt nhất. Phân loại normal_comparison."}
}

# ---------- Dataset 9 ----------
# 32. TIU C11-20  Report Test  SPL JIG improve stuck sample date 13.1.2026 - normal_comparison
DS9_NAME = "32. TIU C11-20  Report Test  SPL JIG improve stuck sample date 13.1.2026"
ds9_result = {
    "schema_version": "0.1",
    "document": {
        "document_id": "doc_ds9",
        "source_file": DS9_NAME,
        "source_sheet": "Test",
        "title": "Report Test New SPL JIG Improve Stuck Sample on JIG TIU-C11-20",
        "model": "TIU-C11-20",
        "report_date": "2026-01-13",
        "department": "ME",
        "marker": "Thao",
        "line": "",
        "report_type": "normal_comparison",
        "primary_defect": {"canonical_name": "Stuck Sample on JIG",
                           "aliases_in_document": ["NG stuck sample on JIG", "Stuck JIG"]},
        "related_defects": ["NG Sigma SPL", "NG Hearing Noise", "NG Hearing Touch", "NG Hearing RB"],
        "parts": ["Sample"],
        "processes": ["SPL Function Test"],
        "purpose": "Improve NG stuck sample on JIG by using a new SPL JIG.",
        "content": [
            "Test new SPL JIG that should reduce stuck sample on JIG.",
            "Check function NG (SPL/RB/Noise/Touch/Stuck JIG)."
        ],
        "source_cells": {"title": ["Test!B2"], "date": ["Test!U3"], "purpose": ["Test!A4"], "content": ["Test!A6"]}
    },
    "test_conditions": [
        {"condition_id": "cond_1", "condition_group": "SPL JIG change",
         "line": "", "process": "SPL Function Test", "changed_factor": "SPL JIG",
         "before_value": "Old SPL JIG (Normal)", "after_value": "New SPL JIG",
         "unit": None, "machine": "SPL function tester", "jig": "New SPL JIG", "material_lot": None,
         "supplier": None, "dry_time_sec": None, "temperature": None, "pressure": None,
         "bond_amount": None, "uv_energy": None,
         "source_file": DS9_NAME, "sheet_name": "Test", "source_cells": ["Test!E10"]}
    ],
    "results": [
        {"result_id": "res_test_total", "condition_id": "cond_1",
         "measurement_type": "Function", "condition_group": "SPL TEST JIG TOTAL",
         "date": "2026-01-12~15", "line": "TIU C11-20L",
         "input_count": 8945, "ok_count": 7631, "ng_count": 1314,
         "ng_rate_decimal": 0.147, "ng_rate_percent": 14.7,
         "metric_name": "Function Total NG rate", "metric_value": 14.7, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"Stuck JIG": {"count": 0, "rate": 0.0},
                          "NG Sigma SPL": {"count": 125, "rate": 1.4},
                          "NG Sigma SPL+RB": {"count": 88, "rate": 1.0},
                          "NG Sigma RB": {"count": 1196, "rate": 13.4},
                          "NG Hearing No sound": {"count": 0, "rate": 0.0},
                          "NG Hearing Noise": {"count": 1057, "rate": 11.8},
                          "NG Hearing Touch": {"count": 44, "rate": 0.5}},
         "source_file": DS9_NAME, "sheet_name": "Test", "source_cells": ["Test!Q14"]},
        {"result_id": "res_normal", "condition_id": None,
         "measurement_type": "Function", "condition_group": "Normal",
         "date": "2026-01-15", "line": "TIU C11-20L",
         "input_count": 2724, "ok_count": 2065, "ng_count": 659,
         "ng_rate_decimal": 0.242, "ng_rate_percent": 24.2,
         "metric_name": "Function Total NG rate (Normal)", "metric_value": 24.2, "unit": "%",
         "judgement": None,
         "ng_breakdown": {"Stuck JIG": {"count": 5, "rate": 0.2},
                          "NG Sigma SPL": {"count": 23, "rate": 0.8},
                          "NG Sigma SPL+RB": {"count": 21, "rate": 0.8},
                          "NG Sigma RB": {"count": 687, "rate": 25.2},
                          "NG Hearing No sound": {"count": 0, "rate": 0.0},
                          "NG Hearing Noise": {"count": 605, "rate": 22.2},
                          "NG Hearing Touch": {"count": 10, "rate": 0.4}},
         "source_file": DS9_NAME, "sheet_name": "Test", "source_cells": ["Test!Q16"]}
    ],
    "conclusions": [
        {"conclusion_id": "concl_1",
         "topic": "New SPL JIG eliminates Stuck JIG NG and reduces function NG",
         "statement_from_report": "Stuck JIG NG 0/8945pcs OK; function NG 1314/8945 = 14.7% same as Normal => CAN USE.",
         "normalized_interpretation": "Stuck JIG NG Test = 0/8945 (0.0%) vs Normal 5/2724 (0.2%) - improvement, root issue resolved. Function NG Test 14.7% vs Normal 24.2% = 0.607x (39.3% improved vs same-event Normal). Dominant NG remains NG Sigma RB and NG Hearing Noise in both groups (sizes broadly similar).",
         "source_file": DS9_NAME, "sheet_name": "Test", "source_cells": ["Test!B26"]}
    ],
    "troubleshooting_index": {
        "defect_name": "Stuck Sample on JIG",
        "when_user_asks": ["How to eliminate stuck sample on SPL JIG?"],
        "suggested_checks": [
            {"hint_id": "hint_1",
             "check_item": "Adopt the new SPL JIG geometry; verify sample-pocket fit and ejection path.",
             "reason": "Stuck JIG NG = 0/8945pcs with new SPL JIG vs Normal 5/2724pcs (0.2%). Function NG Test 14.7% vs Normal 24.2% = 39.3% improved.",
             "evidence_strength": "high", "related_process": "SPL Function Test", "related_part": "JIG",
             "source_file": DS9_NAME, "sheet_name": "Test", "source_cells": ["Test!Q14","Test!Q16"]}
        ],
        "limitations": ["Test sample = 8945 vs Normal = 2724; daily mix variation; only TIU C11-20L line."]
    },
    "ai_extraction_log": {
        "confidence": 0.8,
        "assumptions": ["Same-day TIU C11-20L Normal row used as baseline (15-Jan day shift)."],
        "warnings": ["Total test span across 12~15 Jan; Normal only 15-Jan day shift, so mix differences possible."],
        "decision_rationale": "Same-event Normal row exists at 15-Jan. Stuck JIG NG fully eliminated (0/8945). Function NG Test 14.7% vs Normal 24.2% = 39.3% improved. Decision report states CAN USE. normal_comparison."
    }
}
ds9_tr_en = {
    "document": {"title": ds9_result["document"]["title"], "purpose": ds9_result["document"]["purpose"], "content": ds9_result["document"]["content"]},
    "conclusions": {"concl_1": {"topic": ds9_result["conclusions"][0]["topic"],
                                 "statement_from_report": ds9_result["conclusions"][0]["statement_from_report"],
                                 "normalized_interpretation": ds9_result["conclusions"][0]["normalized_interpretation"]}},
    "hints": {"hint_1": {"check_item": ds9_result["troubleshooting_index"]["suggested_checks"][0]["check_item"],
                          "reason": ds9_result["troubleshooting_index"]["suggested_checks"][0]["reason"]}},
    "log": {"assumptions": ds9_result["ai_extraction_log"]["assumptions"],
            "warnings": ds9_result["ai_extraction_log"]["warnings"],
            "decision_rationale": ds9_result["ai_extraction_log"]["decision_rationale"]}
}
ds9_tr_ko = {
    "document": {"title": "TIU-C11-20 신규 SPL JIG (Stuck sample 개선) 시험 리포트",
                 "purpose": "신규 SPL JIG로 NG stuck sample on JIG 개선 가능성 확인.",
                 "content": ["신규 SPL JIG를 사용하여 stuck sample 발생 여부 확인.",
                             "Function NG (SPL/RB/Noise/Touch/Stuck JIG) 확인."]},
    "conclusions": {"concl_1": {"topic": "신규 SPL JIG: Stuck JIG NG 제거, Function NG 감소",
                                 "statement_from_report": "Stuck JIG NG 0/8945pcs OK, Function NG 1314/8945 = 14.7%로 Normal과 동일 => 사용 가능.",
                                 "normalized_interpretation": "Stuck JIG NG Test = 0/8945 (0.0%) vs Normal 5/2724 (0.2%) - 근본 문제 해결. Function NG Test 14.7% vs Normal 24.2% = 0.607배, 같은 이벤트 Normal 대비 39.3% 개선. 주요 NG는 양측 모두 NG Sigma RB, NG Hearing Noise로 유사."}},
    "hints": {"hint_1": {"check_item": "신규 SPL JIG 형상 적용, 샘플 포켓 fit과 배출 경로 확인.",
                          "reason": "Stuck JIG NG 0/8945pcs로 Normal 5/2724pcs(0.2%) 대비 해소. Function NG Test 14.7% vs Normal 24.2% = 39.3% 개선."}},
    "log": {"assumptions": ["15-Jan TIU C11-20L Normal 행을 baseline으로 사용."],
            "warnings": ["테스트는 12~15 Jan 누적, Normal은 15-Jan Day shift만으로 시점 mix 차이 존재 가능."],
            "decision_rationale": "15-Jan에 Normal 행 존재. Stuck JIG NG 완전 해소(0/8945). Function NG Test 14.7% vs Normal 24.2% = 39.3% 개선. 보고서 결정: 사용 가능. normal_comparison."}
}
ds9_tr_vi = {
    "document": {"title": "TIU-C11-20 Báo cáo test SPL JIG mới giảm stuck sample",
                 "purpose": "Cải thiện NG stuck sample on JIG bằng SPL JIG mới.",
                 "content": ["Test SPL JIG mới giảm stuck sample.",
                             "Kiểm tra NG function (SPL/RB/Noise/Touch/Stuck JIG)."]},
    "conclusions": {"concl_1": {"topic": "SPL JIG mới: loại bỏ NG Stuck JIG, giảm NG function",
                                 "statement_from_report": "Stuck JIG NG 0/8945pcs OK, NG function 1314/8945 = 14.7% giống Normal => CÓ THỂ DÙNG.",
                                 "normalized_interpretation": "NG Stuck JIG Test = 0/8945 (0.0%) so với Normal 5/2724 (0.2%) - hết stuck. NG function Test 14.7% so với Normal 24.2% = 0.607x, cải thiện 39.3% so với Normal cùng sự kiện. NG chính cả hai bên là NG Sigma RB và NG Hearing Noise."}},
    "hints": {"hint_1": {"check_item": "Áp dụng SPL JIG mới, kiểm tra fit của pocket và đường đẩy mẫu ra.",
                          "reason": "NG Stuck JIG 0/8945pcs với jig mới so với Normal 5/2724pcs (0.2%). NG function Test 14.7% so với Normal 24.2% = cải thiện 39.3%."}},
    "log": {"assumptions": ["Dùng hàng Normal 15-Jan TIU C11-20L làm baseline."],
            "warnings": ["Test trải 12~15 Jan trong khi Normal chỉ 15-Jan, có thể có khác biệt mix sản phẩm."],
            "decision_rationale": "Có hàng Normal cùng sự kiện 15-Jan. Stuck JIG NG hết hoàn toàn (0/8945). NG function Test 14.7% vs Normal 24.2% = cải thiện 39.3%. Báo cáo kết luận CÓ THỂ DÙNG. Phân loại normal_comparison."}
}


DATASETS = [
    (DS1_NAME, ds1_result, ds1_tr_ko, ds1_tr_en, ds1_tr_vi),
    (DS2_NAME, ds2_result, ds2_tr_ko, ds2_tr_en, ds2_tr_vi),
    (DS3_NAME, ds3_result, ds3_tr_ko, ds3_tr_en, ds3_tr_vi),
    (DS4_NAME, ds4_result, ds4_tr_ko, ds4_tr_en, ds4_tr_vi),
    (DS5_NAME, ds5_result, ds5_tr_ko, ds5_tr_en, ds5_tr_vi),
    (DS6_NAME, ds6_result, ds6_tr_ko, ds6_tr_en, ds6_tr_vi),
    (DS7_NAME, ds7_result, ds7_tr_ko, ds7_tr_en, ds7_tr_vi),
    (DS8_NAME, ds8_result, ds8_tr_ko, ds8_tr_en, ds8_tr_vi),
    (DS9_NAME, ds9_result, ds9_tr_ko, ds9_tr_en, ds9_tr_vi),
]


def main():
    processed = 0
    failed = 0
    for name, res, tko, ten, tvi in DATASETS:
        ok = h.commit_dataset(name, res, tko, ten, tvi)
        if ok:
            processed += 1
            print(f'OK: {name}')
        else:
            failed += 1
            print(f'FAIL: {name}')
    print(f'chunk 10: processed={processed} failed={failed}')


if __name__ == '__main__':
    main()
