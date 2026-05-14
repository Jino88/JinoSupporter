# -*- coding: utf-8 -*-
"""Process chunk 07 of AI Batch normalization."""
import sys, io
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import _ai_batch_helper as h

DATASETS = []

# =========== Dataset 1: BRS-161014 MC Bonding+Ass'y VP+Dome 30.11.2023 ===========
# Note: name has "3.BRS" and includes "30.11.2023" but title says 30-Oct; this is before_after_dimension (bonding amount sweep) — actually NG-rate sweep over needle/glue amount. ng_without_baseline (no Normal column).
n1 = "3.BRS-161014  REPORT TEST  MC Bonding+Ass'y VP +Dome  30.11.2023"
r1 = {
    "schema_version": "0.1",
    "document": {
        "document_id": "", "source_file": n1, "source_sheet": "V0",
        "title": "Report Test MC Bond+Ass'y VP+Dome - BRS-161014",
        "model": "BRS-161014", "report_date": "2023-10-30",
        "department": "ME", "marker": "Le", "line": "",
        "report_type": "ng_without_baseline",
        "primary_defect": {"canonical_name": "VP+Dome Vision NG", "aliases_in_document": ["Glue not enough", "Over glue", "Glue discontinue"]},
        "related_defects": ["Glue not enough", "Over glue", "Glue discontinue"],
        "parts": ["VP", "Dome"], "processes": ["MC Bonding", "Assy VP+Dome"],
        "purpose": "Check MC Bonding + Ass'y VP+Dome with different needles and bonding amounts; check Vision VP+Dome NG.",
        "content": [
            "Sweep Needle 27G vs 25G and bonding amount ranges.",
            "Inspect Vision VP+Dome for glue not enough / over glue / glue discontinue."
        ],
        "source_cells": {"title": ["V0!B2"], "date": ["V0!Q3"], "purpose": ["V0!A5"], "content": ["V0!A8"]}
    },
    "test_conditions": [
        {"condition_id": "cond_1", "condition_group": "Needle/Bonding Sweep", "line": "", "process": "MC Bonding VP+Dome",
         "changed_factor": "Needle type and bonding amount", "before_value": None, "after_value": None,
         "unit": None, "machine": None, "jig": None, "material_lot": None, "supplier": None,
         "dry_time_sec": None, "temperature": None, "pressure": None, "bond_amount": "1.44-1.6 / 1.86-2.3 / 1.52-1.62 / 2.9-4.0",
         "uv_energy": None, "source_file": n1, "sheet_name": "V0", "source_cells": ["V0!D14:D18"]}
    ],
    "results": [
        {"result_id": "res_1", "condition_id": "cond_1", "measurement_type": "Vision VP+Dome", "condition_group": "Needle 27G, BA 1.44-1.52",
         "date": "2023-10-30", "line": "", "input_count": 100, "ok_count": 4, "ng_count": 96, "ng_rate_decimal": 0.96, "ng_rate_percent": 96.0,
         "metric_name": "Vision NG Rate", "metric_value": 96.0, "unit": "%", "judgement": "FAIL",
         "ng_breakdown": {"Glue not enough": 68, "Over glue": 0, "Glue discontinue": 28, "Other": 0},
         "source_file": n1, "sheet_name": "V0", "source_cells": ["V0!F14:N14"]},
        {"result_id": "res_2", "condition_id": "cond_1", "measurement_type": "Vision VP+Dome", "condition_group": "Needle 27G, BA 1.86-2.3",
         "date": "2023-10-30", "line": "", "input_count": 100, "ok_count": 70, "ng_count": 30, "ng_rate_decimal": 0.30, "ng_rate_percent": 30.0,
         "metric_name": "Vision NG Rate", "metric_value": 30.0, "unit": "%", "judgement": "FAIL",
         "ng_breakdown": {"Glue not enough": 3, "Over glue": 9, "Glue discontinue": 18, "Other": 0},
         "source_file": n1, "sheet_name": "V0", "source_cells": ["V0!F15:N15"]},
        {"result_id": "res_3", "condition_id": "cond_1", "measurement_type": "Vision VP+Dome", "condition_group": "Needle 25G, BA 1.52-1.62",
         "date": "2023-10-30", "line": "", "input_count": 100, "ok_count": 26, "ng_count": 74, "ng_rate_decimal": 0.74, "ng_rate_percent": 74.0,
         "metric_name": "Vision NG Rate", "metric_value": 74.0, "unit": "%", "judgement": "FAIL",
         "ng_breakdown": {"Glue not enough": 58, "Over glue": 0, "Glue discontinue": 16, "Other": 0},
         "source_file": n1, "sheet_name": "V0", "source_cells": ["V0!F16:N16"]},
        {"result_id": "res_4", "condition_id": "cond_1", "measurement_type": "Vision VP+Dome", "condition_group": "Needle 25G, BA 2.9-4.0",
         "date": "2023-10-30", "line": "", "input_count": 100, "ok_count": 36, "ng_count": 64, "ng_rate_decimal": 0.64, "ng_rate_percent": 64.0,
         "metric_name": "Vision NG Rate", "metric_value": 64.0, "unit": "%", "judgement": "FAIL",
         "ng_breakdown": {"Glue not enough": 0, "Over glue": 55, "Glue discontinue": 9, "Other": 0},
         "source_file": n1, "sheet_name": "V0", "source_cells": ["V0!F17:N17"]},
        {"result_id": "res_5", "condition_id": "cond_1", "measurement_type": "Vision VP+Dome", "condition_group": "Extend bonding line 0.05, BA 1.44-1.6",
         "date": "2023-10-30", "line": "", "input_count": 100, "ok_count": 20, "ng_count": 80, "ng_rate_decimal": 0.80, "ng_rate_percent": 80.0,
         "metric_name": "Vision NG Rate", "metric_value": 80.0, "unit": "%", "judgement": "FAIL",
         "ng_breakdown": {"Glue not enough": 72, "Over glue": 0, "Glue discontinue": 8, "Other": 0},
         "source_file": n1, "sheet_name": "V0", "source_cells": ["V0!F18:N18"]}
    ],
    "conclusions": [
        {"conclusion_id": "concl_1", "topic": "MC Bonding setting needs improvement",
         "statement_from_report": "When use needle 27G, 25G need setting MC bonding VP+Dome happen NG glue not enough and glue discontinue, over glue => Need check MC and improve it.",
         "normalized_interpretation": "All tested needle/bonding-amount combinations show very high Vision VP+Dome NG (30-96%). No same-event Normal/Baseline row is present, so absolute NG rates only — MC bonding setting must be revised before further trials.",
         "source_file": n1, "sheet_name": "V0", "source_cells": ["V0!A19"]}
    ],
    "troubleshooting_index": {
        "defect_name": "VP+Dome Vision NG",
        "when_user_asks": ["VP+Dome bonding NG", "Glue not enough", "Over glue", "Glue discontinue"],
        "suggested_checks": [
            {"hint_id": "hint_1", "check_item": "Re-tune MC Bonding setting: needle gauge, bonding amount range, and bonding line extension",
             "reason": "All four needle/bonding-amount combinations produced 30-96% Vision NG with no baseline reference; the MC bonding equipment setting is the dominant cause.",
             "evidence_strength": "high", "related_process": "MC Bonding VP+Dome", "related_part": "VP, Dome",
             "source_file": n1, "sheet_name": "V0", "source_cells": ["V0!A19"]}
        ],
        "limitations": ["No same-event Normal/Baseline row to compute relative change."]
    },
    "ai_extraction_log": {
        "confidence": 0.65,
        "assumptions": ["Report date interpreted as 2023-10-30 from header (30-Oct) despite dataset name '30.11.2023'."],
        "warnings": ["No same-event baseline row; absolute NG rates only.", "Decision section is empty in source."],
        "decision_rationale": "Classified as ng_without_baseline because no Normal/Baseline row appears in the Vision VP+Dome table — only four test conditions are listed."
    }
}
tr_ko_1 = {
    "document": {"title": "MC Bond+Assy VP+Dome 시험 리포트 - BRS-161014",
                 "purpose": "다양한 니들 및 본딩 양 조건에서 MC Bonding + Ass'y VP+Dome 의 Vision NG 를 확인한다.",
                 "content": ["Needle 27G vs 25G 및 본딩량 범위 스윕.",
                             "Vision VP+Dome 의 Glue not enough / Over glue / Glue discontinue 확인."]},
    "conclusions": {"concl_1": {"topic": "MC Bonding 설정 개선 필요",
                                 "statement_from_report": "Needle 27G, 25G 사용 시 MC bonding VP+Dome 에서 glue not enough, glue discontinue, over glue 가 발생 → MC 확인 및 개선 필요.",
                                 "normalized_interpretation": "모든 니들/본딩량 조합에서 Vision VP+Dome NG 가 30~96% 로 매우 높음. 동일 이벤트의 Normal/Baseline 행이 없어 절대 NG 율만으로 판단. 추가 시험 전에 MC bonding 설정 재조정 필요."}},
    "hints": {"hint_1": {"check_item": "MC Bonding 설정 재조정: 니들 게이지, 본딩량 범위, 본딩 라인 연장",
                          "reason": "4가지 니들/본딩량 조합 모두 baseline 없이 30~96% Vision NG 발생. MC bonding 장비 설정이 주된 원인."}},
    "log": {"assumptions": ["헤더 30-Oct 에 따라 보고 일자를 2023-10-30 으로 해석 (데이터셋 이름의 30.11.2023 와 다름)."],
            "warnings": ["동일 이벤트 baseline 행 없음; 절대 NG 율만 사용.", "원본 Decision 칸이 비어 있음."],
            "decision_rationale": "Vision VP+Dome 표에 Normal/Baseline 행이 없고 4개 시험 조건만 나열되어 ng_without_baseline 으로 분류."}
}
tr_en_1 = {
    "document": {"title": r1["document"]["title"], "purpose": r1["document"]["purpose"], "content": r1["document"]["content"]},
    "conclusions": {"concl_1": {"topic": r1["conclusions"][0]["topic"],
                                 "statement_from_report": r1["conclusions"][0]["statement_from_report"],
                                 "normalized_interpretation": r1["conclusions"][0]["normalized_interpretation"]}},
    "hints": {"hint_1": {"check_item": r1["troubleshooting_index"]["suggested_checks"][0]["check_item"],
                          "reason": r1["troubleshooting_index"]["suggested_checks"][0]["reason"]}},
    "log": {"assumptions": r1["ai_extraction_log"]["assumptions"],
            "warnings": r1["ai_extraction_log"]["warnings"],
            "decision_rationale": r1["ai_extraction_log"]["decision_rationale"]}
}
tr_vi_1 = {
    "document": {"title": "Báo cáo test MC Bond+Assy VP+Dome - BRS-161014",
                 "purpose": "Kiểm tra Vision NG của MC Bonding + Ass'y VP+Dome với các loại kim và lượng bonding khác nhau.",
                 "content": ["Quét Needle 27G và 25G cùng các dải bonding amount.",
                             "Kiểm tra Vision VP+Dome: Glue not enough / Over glue / Glue discontinue."]},
    "conclusions": {"concl_1": {"topic": "Cần cải thiện setting MC Bonding",
                                 "statement_from_report": "Khi dùng Needle 27G, 25G phát sinh glue not enough, glue discontinue, over glue ở MC bonding VP+Dome => cần kiểm tra MC và cải thiện.",
                                 "normalized_interpretation": "Tất cả tổ hợp needle/bonding amount đều cho Vision VP+Dome NG rất cao (30-96%). Không có dòng Normal/Baseline cùng sự kiện nên chỉ dùng NG rate tuyệt đối; cần điều chỉnh MC bonding trước khi test tiếp."}},
    "hints": {"hint_1": {"check_item": "Hiệu chỉnh lại setting MC Bonding: cỡ kim, dải bonding amount, kéo dài bonding line",
                          "reason": "Cả 4 tổ hợp đều có NG 30-96% mà không có baseline; thiết bị MC bonding là nguyên nhân chính."}},
    "log": {"assumptions": ["Ngày báo cáo lấy theo header 30-Oct (2023-10-30), khác với 30.11.2023 trong tên dataset."],
            "warnings": ["Không có dòng baseline cùng sự kiện; chỉ dùng NG rate tuyệt đối.", "Phần Decision trong file gốc để trống."],
            "decision_rationale": "Bảng Vision VP+Dome chỉ có 4 điều kiện test, không có Normal/Baseline -> phân loại ng_without_baseline."}
}
DATASETS.append((n1, r1, tr_ko_1, tr_en_1, tr_vi_1))

# =========== Dataset 2: 30. BRS-161014 new Jig ass'y Frame+Yoke by hand ===========
n2 = "30. BRS-161014 Report TEST new Jig ass'y Frame+Yoke by hand"
# Two events: 9/16 test 61.2% vs normal 63.8%; 9/28 test 10% vs normal 9.0%.
# rel: 9/16: 61.2/63.8 -1 = -4.08% (improved 4.1%); 9/28: 10.0/9.0 -1 = +11.1% (worse 11.1%)
r2 = {
    "schema_version": "0.1",
    "document": {
        "document_id": "", "source_file": n2, "source_sheet": "Report (2)",
        "title": "Report Test New Concept Jig Ass'y Frame+Yoke by Hand - BRS-161014",
        "model": "BRS-161014", "report_date": "2023-09-16", "department": "ME",
        "marker": "Thuy", "line": "", "report_type": "normal_comparison",
        "primary_defect": {"canonical_name": "NG Hearing", "aliases_in_document": ["Hearing Noise", "Hearing Touch", "Hearing high"]},
        "related_defects": ["NG Hearing Noise", "NG Hearing Touch"],
        "parts": ["Frame", "Yoke"], "processes": ["Assy Frame+Yoke"],
        "purpose": "Improve NG rate of hearing high using new-concept Jig for Ass'y Frame+Yoke by hand.",
        "content": ["Use new-concept Jig to ass'y Frame+Yoke by hand and compare Function NG rate with normal line."],
        "source_cells": {"title": ["Report (2)!B2"], "date": ["Report (2)!Q3"], "purpose": ["Report (2)!A5"], "content": ["Report (2)!A7"]}
    },
    "test_conditions": [
        {"condition_id": "cond_1", "condition_group": "New Jig Frame+Yoke", "line": "", "process": "Assy Frame+Yoke",
         "changed_factor": "Jig for Frame+Yoke (new concept, by hand)", "before_value": "Normal jig",
         "after_value": "New concept jig (by hand)", "unit": None, "machine": None, "jig": "New concept Frame+Yoke jig",
         "material_lot": None, "supplier": None, "dry_time_sec": None, "temperature": None, "pressure": None,
         "bond_amount": None, "uv_energy": None, "source_file": n2, "sheet_name": "Report (2)", "source_cells": ["Report (2)!A6:A7"]},
        {"condition_id": "cond_2", "condition_group": "New Jig Frame+Yoke After Improve", "line": "", "process": "Assy Frame+Yoke",
         "changed_factor": "Jig for Frame+Yoke (new concept, after improve)", "before_value": "Normal jig",
         "after_value": "New concept jig after improve", "unit": None, "machine": None, "jig": "New concept Frame+Yoke jig (improved)",
         "material_lot": None, "supplier": None, "dry_time_sec": None, "temperature": None, "pressure": None,
         "bond_amount": None, "uv_energy": None, "source_file": n2, "sheet_name": "Report (2)", "source_cells": ["Report (2)!B12"]}
    ],
    "results": [
        {"result_id": "res_1", "condition_id": "cond_1", "measurement_type": "Function", "condition_group": "Test new jig",
         "date": "2023-09-16", "line": "", "input_count": 98, "ok_count": 38, "ng_count": 60, "ng_rate_decimal": 0.612, "ng_rate_percent": 61.2,
         "metric_name": "Total NG Rate", "metric_value": 61.2, "unit": "%", "judgement": None,
         "ng_breakdown": {"NG Hearing Noise": 42, "NG Hearing Touch": 18, "Sigma SPL": 0, "Sigma THD": 0, "SPL+THD": 0, "SPL+THD+F0": 0, "HOHD": 0},
         "source_file": n2, "sheet_name": "Report (2)", "source_cells": ["Report (2)!A10:M10"]},
        {"result_id": "res_2", "condition_id": "cond_1", "measurement_type": "Function", "condition_group": "Normal",
         "date": "2023-09-16", "line": "", "input_count": 788, "ok_count": 290, "ng_count": 503, "ng_rate_decimal": 0.638, "ng_rate_percent": 63.8,
         "metric_name": "Total NG Rate (Baseline)", "metric_value": 63.8, "unit": "%", "judgement": None,
         "ng_breakdown": {"NG Hearing Noise": 413, "NG Hearing Touch": 84, "Sigma SPL": 0, "Sigma THD": 0, "SPL+THD": 1, "SPL+THD+F0": 5, "HOHD": 0},
         "source_file": n2, "sheet_name": "Report (2)", "source_cells": ["Report (2)!A11:M11"]},
        {"result_id": "res_3", "condition_id": "cond_2", "measurement_type": "Function", "condition_group": "Test new jig after improve",
         "date": "2023-09-28", "line": "", "input_count": 100, "ok_count": 90, "ng_count": 10, "ng_rate_decimal": 0.10, "ng_rate_percent": 10.0,
         "metric_name": "Total NG Rate", "metric_value": 10.0, "unit": "%", "judgement": None,
         "ng_breakdown": {"NG Hearing Noise": 10, "NG Hearing Touch": 0, "Sigma SPL": 0, "Sigma THD": 0, "SPL+THD": 0, "SPL+THD+F0": 0, "HOHD": 0},
         "source_file": n2, "sheet_name": "Report (2)", "source_cells": ["Report (2)!A12:M12"]},
        {"result_id": "res_4", "condition_id": "cond_2", "measurement_type": "Function", "condition_group": "Normal",
         "date": "2023-09-28", "line": "", "input_count": 100, "ok_count": 91, "ng_count": 9, "ng_rate_decimal": 0.09, "ng_rate_percent": 9.0,
         "metric_name": "Total NG Rate (Baseline)", "metric_value": 9.0, "unit": "%", "judgement": None,
         "ng_breakdown": {"NG Hearing Noise": 9, "NG Hearing Touch": 0, "Sigma SPL": 0, "Sigma THD": 0, "SPL+THD": 0, "SPL+THD+F0": 0, "HOHD": 0},
         "source_file": n2, "sheet_name": "Report (2)", "source_cells": ["Report (2)!A13:M13"]}
    ],
    "conclusions": [
        {"conclusion_id": "concl_1", "topic": "New Frame+Yoke jig vs normal — equivalent",
         "statement_from_report": "NG rate of test lot is same as normal lot.",
         "normalized_interpretation": "9/16: Test new jig 61.2% vs Normal 63.8% = 0.96x, 4.1% improved vs same-event normal. 9/28 (after jig improve): Test 10.0% vs Normal 9.0% = 1.11x, 11.1% worse vs normal. Overall new jig is comparable to normal line; Hearing Noise/Touch remain the dominant items.",
         "source_file": n2, "sheet_name": "Report (2)", "source_cells": ["Report (2)!A14"]}
    ],
    "troubleshooting_index": {
        "defect_name": "NG Hearing",
        "when_user_asks": ["Hearing Noise", "Hearing Touch", "Hearing high NG", "Frame+Yoke jig"],
        "suggested_checks": [
            {"hint_id": "hint_1", "check_item": "Compare new-concept Frame+Yoke jig against normal jig in same event using Hearing Noise/Touch breakdown",
             "reason": "Function NG rate for new jig is 0.96x vs normal on 9/16 and 1.11x vs normal on 9/28 — within noise; jig change alone does not move Hearing NG meaningfully.",
             "evidence_strength": "medium", "related_process": "Assy Frame+Yoke", "related_part": "Frame, Yoke",
             "source_file": n2, "sheet_name": "Report (2)", "source_cells": ["Report (2)!A10:M13"]}
        ],
        "limitations": ["Small sample size on test lot (98 and 100 pcs)."]
    },
    "ai_extraction_log": {
        "confidence": 0.78,
        "assumptions": [],
        "warnings": ["Small test-lot sample sizes; 9/28 normal lot also only 100 pcs."],
        "decision_rationale": "Same-event Normal rows present for both 9/16 and 9/28 events, so relative change is computable: (61.2/63.8-1)*100 = -4.1% (improved); (10.0/9.0-1)*100 = +11.1% (worse). New jig is statistically comparable to normal."
    }
}
tr_ko_2 = {"document": {"title": "Frame+Yoke 수작업 신규 지그 시험 리포트 - BRS-161014",
                         "purpose": "신규 컨셉 지그로 Frame+Yoke 수작업 조립 후 Hearing high NG 개선 여부 확인.",
                         "content": ["신규 지그로 Frame+Yoke 조립 후 Function NG 율을 Normal line 과 비교."]},
           "conclusions": {"concl_1": {"topic": "신규 Frame+Yoke 지그 vs Normal — 동등",
                                        "statement_from_report": "시험 lot 의 NG rate 는 normal lot 과 동일.",
                                        "normalized_interpretation": "9/16: Test 61.2% vs Normal 63.8% = 0.96x, normal 대비 4.1% 개선. 9/28 (지그 개선 후): Test 10.0% vs Normal 9.0% = 1.11x, normal 대비 11.1% 악화. 전체적으로 신규 지그는 normal 과 동등 수준이며 Hearing Noise/Touch 가 지배적."}},
           "hints": {"hint_1": {"check_item": "동일 이벤트에서 신규 Frame+Yoke 지그를 normal 지그와 Hearing Noise/Touch 분해 기준으로 비교",
                                 "reason": "Function NG 가 9/16 0.96x, 9/28 1.11x 로 noise 수준; 지그 변경만으로 Hearing NG 가 의미 있게 개선되지 않음."}},
           "log": {"assumptions": [],
                   "warnings": ["시험 lot 의 샘플 수가 적음 (98, 100 pcs). 9/28 normal lot 도 100 pcs."],
                   "decision_rationale": "9/16, 9/28 두 이벤트 모두 동일 이벤트 Normal 행이 존재해 상대변화율 계산 가능: -4.1% (개선), +11.1% (악화). 신규 지그는 normal 과 통계적으로 동등."}}
tr_en_2 = {"document": {"title": r2["document"]["title"], "purpose": r2["document"]["purpose"], "content": r2["document"]["content"]},
           "conclusions": {"concl_1": {"topic": r2["conclusions"][0]["topic"],
                                        "statement_from_report": r2["conclusions"][0]["statement_from_report"],
                                        "normalized_interpretation": r2["conclusions"][0]["normalized_interpretation"]}},
           "hints": {"hint_1": {"check_item": r2["troubleshooting_index"]["suggested_checks"][0]["check_item"],
                                 "reason": r2["troubleshooting_index"]["suggested_checks"][0]["reason"]}},
           "log": {"assumptions": r2["ai_extraction_log"]["assumptions"],
                   "warnings": r2["ai_extraction_log"]["warnings"],
                   "decision_rationale": r2["ai_extraction_log"]["decision_rationale"]}}
tr_vi_2 = {"document": {"title": "Báo cáo test Jig mới ass'y Frame+Yoke bằng tay - BRS-161014",
                         "purpose": "Cải thiện NG hearing high bằng jig kiểu mới khi ass'y Frame+Yoke thủ công.",
                         "content": ["Dùng jig kiểu mới ass'y Frame+Yoke bằng tay và so sánh Function NG rate với normal line."]},
           "conclusions": {"concl_1": {"topic": "Jig Frame+Yoke mới vs Normal — tương đương",
                                        "statement_from_report": "NG rate của lot test giống lot normal.",
                                        "normalized_interpretation": "9/16: Test 61.2% vs Normal 63.8% = 0.96x, cải thiện 4.1% so với normal cùng sự kiện. 9/28 (sau cải tiến jig): Test 10.0% vs Normal 9.0% = 1.11x, xấu hơn 11.1%. Nhìn chung jig mới tương đương normal; Hearing Noise/Touch vẫn là NG chính."}},
           "hints": {"hint_1": {"check_item": "So sánh jig Frame+Yoke mới với jig normal trong cùng sự kiện theo phân loại Hearing Noise/Touch",
                                 "reason": "Function NG 0.96x (9/16) và 1.11x (9/28) so với normal — chỉ trong khoảng nhiễu; chỉ thay jig không cải thiện rõ Hearing NG."}},
           "log": {"assumptions": [],
                   "warnings": ["Số mẫu test lot nhỏ (98, 100 pcs); 9/28 normal lot cũng chỉ 100 pcs."],
                   "decision_rationale": "Cả 9/16 và 9/28 đều có hàng Normal cùng sự kiện -> tính được relative change: -4.1% (cải thiện), +11.1% (xấu hơn). Jig mới tương đương normal."}}
DATASETS.append((n2, r2, tr_ko_2, tr_en_2, tr_vi_2))

if __name__ == '__main__':
    processed = 0; failed = 0
    for name, r, ko, en, vi in DATASETS:
        ok = h.commit_dataset(name, r, ko, en, vi)
        if ok: processed += 1
        else: failed += 1
        print(f'  {name} -> {"ok" if ok else "FAIL"}')
    print(f'chunk 07: processed={processed} failed={failed}')
