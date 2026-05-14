# -*- coding: utf-8 -*-
import _ai_batch_helper as h

NAME = "26. BRS-161014 DT Report test CD supplier Ralon  -  date 2024.03.11"

result = {
    "schema_version": "0.1",
    "document": {
        "document_id": "", "source_file": NAME, "source_sheet": "Test",
        "title": "BRS-161016 DT - REPORT TEST CD SUPPLIER RALON",
        "model": "BRS-161016", "report_date": "2024-03-11",
        "department": "ME", "marker": "Nhung", "line": "C2-3A",
        "report_type": "normal_comparison",
        "primary_defect": {"canonical_name": "NG Hearing Noise",
                            "aliases_in_document": ["NG hearing noise + touch", "Noise"]},
        "related_defects": ["VP+CD Separation", "NG Hearing Touch", "Coil deform", "Fr/YOKE offset", "CM offset"],
        "parts": ["CD", "Dome"],
        "processes": ["Vision VP/CD", "Function test", "Decap NG analysis"],
        "purpose": "Test CD supplier Ralon vs Normal CD supplier GES; check NG rate Vision VP/CD and Function; analyze Decap NG.",
        "content": [
            "Make semi and check NG rate of Vision VP/CD",
            "Make final and check NG rate of Function",
            "Compare with Normal supplier GES",
            "Decap analysis of NG Hearing Noise + Touch"
        ],
        "source_cells": {
            "title": ["Test!C2"], "date": ["Test!Q3"],
            "purpose": ["Test!A6"], "content": ["Test!A8:A10"]
        }
    },
    "test_conditions": [
        {"condition_id": "cond_1", "condition_group": "Test Ralon",
         "line": "C2-3A", "process": "CD supply", "changed_factor": "CD supplier",
         "before_value": "GES (Normal)", "after_value": "Ralon",
         "unit": None, "machine": None, "jig": None, "material_lot": None,
         "supplier": "Ralon", "dry_time_sec": None, "temperature": None,
         "pressure": None, "bond_amount": None, "uv_energy": None,
         "source_file": NAME, "sheet_name": "Test",
         "source_cells": ["Test!D15", "Test!D22"]},
        {"condition_id": "cond_2", "condition_group": "Normal GES",
         "line": "C2-3A", "process": "CD supply", "changed_factor": "CD supplier",
         "before_value": None, "after_value": "GES",
         "unit": None, "machine": None, "jig": None, "material_lot": None,
         "supplier": "GES", "dry_time_sec": None, "temperature": None,
         "pressure": None, "bond_amount": None, "uv_energy": None,
         "source_file": NAME, "sheet_name": "Test",
         "source_cells": ["Test!D17", "Test!D24"]}
    ],
    "results": [
        {"result_id": "res_1", "condition_id": "cond_1",
         "measurement_type": "Vision", "condition_group": "Vision VP/CD - Test Ralon",
         "date": "2024-03-11", "line": "C2-3A",
         "input_count": 504, "ok_count": 504, "ng_count": 0,
         "ng_rate_decimal": 0.0, "ng_rate_percent": 0.0,
         "metric_name": "Vision VP/CD NG Rate", "metric_value": 0.0,
         "unit": "%", "judgement": "PASS",
         "ng_breakdown": {"Glue not enough": {"count": 0, "rate": 0.0}},
         "source_file": NAME, "sheet_name": "Test",
         "source_cells": ["Test!B15:I16"]},
        {"result_id": "res_2", "condition_id": "cond_2",
         "measurement_type": "Vision", "condition_group": "Vision VP/CD - Normal GES",
         "date": "2024-03-11", "line": "C2-3A",
         "input_count": 504, "ok_count": 504, "ng_count": 0,
         "ng_rate_decimal": 0.0, "ng_rate_percent": 0.0,
         "metric_name": "Vision VP/CD NG Rate", "metric_value": 0.0,
         "unit": "%", "judgement": "PASS",
         "ng_breakdown": {"Glue not enough": {"count": 0, "rate": 0.0}},
         "source_file": NAME, "sheet_name": "Test",
         "source_cells": ["Test!B17:I18"]},
        {"result_id": "res_3", "condition_id": "cond_1",
         "measurement_type": "Function", "condition_group": "Function - Test Ralon",
         "date": "2024-03-11", "line": "C2-3A",
         "input_count": 504, "ok_count": 495, "ng_count": 9,
         "ng_rate_decimal": 0.0179, "ng_rate_percent": 1.79,
         "metric_name": "Function Total NG Rate", "metric_value": 1.79,
         "unit": "%", "judgement": None,
         "ng_breakdown": {
             "NG Sigma SPL": {"count": 0, "rate": 0.0},
             "NG Sigma SPL+THD": {"count": 0, "rate": 0.0},
             "NG Sigma SPL+THD+F0": {"count": 0, "rate": 0.0},
             "NG Hearing Noise": {"count": 8, "rate": 0.016},
             "NG Hearing Touch": {"count": 1, "rate": 0.002},
             "Air leak": {"count": 0, "rate": 0.0}
         },
         "source_file": NAME, "sheet_name": "Test",
         "source_cells": ["Test!B22:N23"]},
        {"result_id": "res_4", "condition_id": "cond_2",
         "measurement_type": "Function", "condition_group": "Function - Normal GES",
         "date": "2024-03-11", "line": "C2-3A",
         "input_count": 504, "ok_count": 492, "ng_count": 12,
         "ng_rate_decimal": 0.0238, "ng_rate_percent": 2.38,
         "metric_name": "Function Total NG Rate", "metric_value": 2.38,
         "unit": "%", "judgement": None,
         "ng_breakdown": {
             "NG Sigma SPL": {"count": 0, "rate": 0.0},
             "NG Sigma SPL+THD": {"count": 0, "rate": 0.0},
             "NG Sigma SPL+THD+F0": {"count": 0, "rate": 0.0},
             "NG Hearing Noise": {"count": 9, "rate": 0.018},
             "NG Hearing Touch": {"count": 3, "rate": 0.006},
             "Air leak": {"count": 0, "rate": 0.0}
         },
         "source_file": NAME, "sheet_name": "Test",
         "source_cells": ["Test!B24:N25"]},
        {"result_id": "res_5", "condition_id": "cond_1",
         "measurement_type": "Decap", "condition_group": "Decap NG analysis - Coil deform",
         "date": "2024-03-11", "line": "C2-3A",
         "input_count": 9, "ok_count": None, "ng_count": 4,
         "ng_rate_decimal": 0.4444, "ng_rate_percent": 44.4,
         "metric_name": "Decap reason portion - Coil deform",
         "metric_value": 44.4, "unit": "%", "judgement": None,
         "ng_breakdown": {},
         "source_file": NAME, "sheet_name": "Decap NG",
         "source_cells": ["Decap NG!D40:F40", "Decap NG!E51:G51"]},
        {"result_id": "res_6", "condition_id": "cond_1",
         "measurement_type": "Decap", "condition_group": "Decap NG analysis - Fr/YOKE offset",
         "date": "2024-03-11", "line": "C2-3A",
         "input_count": 9, "ok_count": None, "ng_count": 2,
         "ng_rate_decimal": 0.2222, "ng_rate_percent": 22.2,
         "metric_name": "Decap reason portion - Fr/YOKE offset",
         "metric_value": 22.2, "unit": "%", "judgement": None,
         "ng_breakdown": {},
         "source_file": NAME, "sheet_name": "Decap NG",
         "source_cells": ["Decap NG!E52:G52"]},
        {"result_id": "res_7", "condition_id": "cond_1",
         "measurement_type": "Decap", "condition_group": "Decap NG analysis - CM offset",
         "date": "2024-03-11", "line": "C2-3A",
         "input_count": 9, "ok_count": None, "ng_count": 1,
         "ng_rate_decimal": 0.1111, "ng_rate_percent": 11.1,
         "metric_name": "Decap reason portion - CM offset",
         "metric_value": 11.1, "unit": "%", "judgement": None,
         "ng_breakdown": {},
         "source_file": NAME, "sheet_name": "Decap NG",
         "source_cells": ["Decap NG!E53:G53"]},
        {"result_id": "res_8", "condition_id": "cond_1",
         "measurement_type": "Decap", "condition_group": "Decap NG analysis - Don't know",
         "date": "2024-03-11", "line": "C2-3A",
         "input_count": 9, "ok_count": None, "ng_count": 2,
         "ng_rate_decimal": 0.2222, "ng_rate_percent": 22.2,
         "metric_name": "Decap reason portion - Don't know",
         "metric_value": 22.2, "unit": "%", "judgement": None,
         "ng_breakdown": {},
         "source_file": NAME, "sheet_name": "Decap NG",
         "source_cells": ["Decap NG!E58:G58"]}
    ],
    "conclusions": [
        {"conclusion_id": "concl_1", "topic": "CD supplier Ralon vs Normal GES",
         "statement_from_report": "Decision row blank; report compares Vision VP/CD and Function for Ralon vs GES.",
         "normalized_interpretation": "Vision VP/CD: Ralon 0.00% vs GES 0.00% (equal). Function: Ralon 1.79% vs Normal GES 2.38% same event => (1.79/2.38 - 1)*100 = -24.8%, 24.8% improved vs same-event Normal. Dominant defect in both groups is NG Hearing Noise (Ralon 1.6%, GES 1.8%). Decap of Ralon's 9 Function NG units shows Coil deform 44.4%, Fr/YOKE offset 22.2%, CM offset 11.1%, Don't know reason 22.2%.",
         "source_file": NAME, "sheet_name": "Test", "source_cells": ["Test!A26"]}
    ],
    "troubleshooting_index": {
        "defect_name": "NG Hearing Noise on Ralon CD",
        "when_user_asks": ["CD supplier", "Ralon", "BRS-161016", "dome supplier"],
        "suggested_checks": [
            {"hint_id": "hint_1",
             "check_item": "Compare Ralon vs GES CD dimension and coil seating to control Coil deform",
             "reason": "Decap of Ralon NG units shows Coil deform 44.4%, the largest portion.",
             "evidence_strength": "moderate",
             "related_process": "Function NG analysis",
             "related_part": "Coil / Dome",
             "source_file": NAME, "sheet_name": "Decap NG",
             "source_cells": ["Decap NG!E51:G51"]},
            {"hint_id": "hint_2",
             "check_item": "Confirm Frame/YOKE alignment when changing CD supplier",
             "reason": "Decap Fr/YOKE offset 22.2% second largest portion among Ralon NG.",
             "evidence_strength": "moderate",
             "related_process": "Frame/YOKE assembly",
             "related_part": "Frame / YOKE",
             "source_file": NAME, "sheet_name": "Decap NG",
             "source_cells": ["Decap NG!E52:G52"]}
        ],
        "limitations": ["Ralon Function NG count only 9 - small base for decap percentages."]
    },
    "ai_extraction_log": {
        "confidence": 0.8,
        "assumptions": ["Title prints BRS-161016 but dataset name BRS-161014; both kept.",
                         "Date 11/Mar interpreted as 2024-03-11 per dataset name."],
        "warnings": ["Small NG base (9 units) for Decap portions."],
        "decision_rationale": "Same-event Normal GES baseline exists. Function NG Ralon 1.79% vs GES 2.38% => (1.79/2.38-1)*100=-24.8%, 24.8% improved. Vision VP/CD equal at 0%. Decap of Ralon's 9 NGs is dominated by Coil deform (44.4%) and Fr/YOKE offset (22.2%)."
    }
}

tr_en = {
    "document": {"title": result["document"]["title"],
                 "purpose": result["document"]["purpose"],
                 "content": result["document"]["content"]},
    "conclusions": {
        "concl_1": {"topic": "Ralon vs GES",
                    "statement_from_report": "Decision row blank; report compares Vision and Function.",
                    "normalized_interpretation": "Vision VP/CD equal 0%. Function Ralon 1.79% vs Normal GES 2.38% => 24.8% improved. Hearing Noise dominant. Decap: Coil deform 44.4%, Fr/YOKE offset 22.2%."}
    },
    "hints": {
        "hint_1": {"check_item": "Compare Ralon vs GES CD dimension and coil seating",
                   "reason": "Coil deform 44.4% of Ralon NG decap."},
        "hint_2": {"check_item": "Confirm Frame/YOKE alignment when changing CD supplier",
                   "reason": "Fr/YOKE offset 22.2% second largest portion."}
    },
    "log": {"assumptions": ["Title prints BRS-161016 but dataset BRS-161014.",
                            "11/Mar -> 2024-03-11."],
            "warnings": ["Small NG base (9) for decap %."],
            "decision_rationale": "Function 1.79% vs 2.38% baseline -> -24.8% improved. Decap dominated by Coil deform."}
}

tr_ko = {
    "document": {
        "title": "BRS-161014 DT CD 공급사 Ralon 시험 보고서",
        "purpose": "CD 공급사 Ralon vs Normal(GES) Vision VP/CD 및 Function NG Rate 비교, Decap NG 분석.",
        "content": [
            "Semi 제작 후 Vision VP/CD NG Rate 확인",
            "Final 제작 후 Function NG Rate 확인",
            "Normal(GES)와 비교",
            "NG Hearing Noise+Touch Decap 분석"
        ]
    },
    "conclusions": {
        "concl_1": {"topic": "Ralon vs Normal GES",
                    "statement_from_report": "Decision 행은 비어있음. Ralon과 GES의 Vision/Function 결과를 비교.",
                    "normalized_interpretation": "Vision VP/CD 양쪽 0% 동일. Function Ralon 1.79% vs Normal GES 2.38% => 24.8% 개선. Hearing Noise가 주된 NG. Decap: Coil deform 44.4%, Fr/YOKE offset 22.2%."}
    },
    "hints": {
        "hint_1": {"check_item": "Ralon vs GES CD 치수 및 Coil 안착 비교",
                   "reason": "Decap에서 Coil deform 44.4%로 최대 비중."},
        "hint_2": {"check_item": "CD 공급사 변경 시 Frame/YOKE 정렬 재확인",
                   "reason": "Decap Fr/YOKE offset 22.2% 두번째 비중."}
    },
    "log": {"assumptions": ["타이틀은 BRS-161016이나 dataset명은 BRS-161014, 둘 다 보존.",
                            "날짜 11/Mar = 2024-03-11."],
            "warnings": ["Decap base 9pcs로 작음."],
            "decision_rationale": "Function 1.79% vs 2.38% baseline => -24.8% 개선. Decap은 Coil deform 주도."}
}

tr_vi = {
    "document": {
        "title": "Báo cáo test CD nhà cung cấp Ralon BRS-161014 DT",
        "purpose": "So sánh CD Ralon với Normal (GES) về Vision VP/CD, Function và phân tích Decap NG.",
        "content": [
            "Làm semi và kiểm tra NG rate Vision VP/CD",
            "Làm final và kiểm tra Function NG",
            "So sánh với Normal supplier GES",
            "Phân tích Decap NG Hearing Noise + Touch"
        ]
    },
    "conclusions": {
        "concl_1": {"topic": "Ralon vs Normal GES",
                    "statement_from_report": "Phần Decision trống; báo cáo so sánh Vision và Function giữa Ralon và GES.",
                    "normalized_interpretation": "Vision VP/CD bằng nhau 0%. Function Ralon 1.79% vs GES 2.38% => giảm 24.8%. NG chủ yếu Hearing Noise. Decap: Coil deform 44.4%, Fr/YOKE offset 22.2%."}
    },
    "hints": {
        "hint_1": {"check_item": "So sánh kích thước và lắp coil của Ralon và GES",
                   "reason": "Coil deform 44.4% trong các NG Decap của Ralon."},
        "hint_2": {"check_item": "Kiểm tra lại căn chỉnh Frame/YOKE khi đổi nhà cung cấp CD",
                   "reason": "Fr/YOKE offset 22.2% phần lớn thứ hai."}
    },
    "log": {"assumptions": ["Tiêu đề ghi BRS-161016 nhưng dataset BRS-161014 - giữ cả hai.",
                            "11/Mar = 2024-03-11."],
            "warnings": ["Decap dựa trên 9 mẫu, nhỏ."],
            "decision_rationale": "Function 1.79% vs baseline 2.38% -> giảm 24.8%. Decap chủ yếu Coil deform."}
}

ok = h.commit_dataset(NAME, result, tr_ko, tr_en, tr_vi)
print("ds03", ok)
