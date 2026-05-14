# -*- coding: utf-8 -*-
import _ai_batch_helper as h

NAME = "25.BRS-201506 Report check Forming machine at Sub VP - 2024.03.05"

result = {
    "schema_version": "0.1",
    "document": {
        "document_id": "", "source_file": NAME, "source_sheet": "05.03",
        "title": "BRS-201506 REPORT CHECK FORMING MACHINE",
        "model": "BRS-201506", "report_date": "2024-03-05",
        "department": "ME", "marker": "Nhung", "line": "VP",
        "report_type": "reliability_spec",
        "primary_defect": {"canonical_name": "NG Function High Rate",
                            "aliases_in_document": ["NG function hight", "function high rate"]},
        "related_defects": [],
        "parts": ["Forming machine VP"],
        "processes": ["Sub VP forming"],
        "purpose": "Investigate forming machines feeding main line where Function NG was high; find which machine causes NG.",
        "content": [
            "Check forming machine conditions: temperature, pressure, forming time",
            "Separate by forming machine No13-2H, No14-2H, No14-4H",
            "Spec Temperature 170~190 C, Cooling 40+-10 C, Pressure 1.2+-0.2 kgf/cm2"
        ],
        "source_cells": {
            "title": ["05.03!C2"], "date": ["05.03!K3"],
            "purpose": ["05.03!B6:B7"], "content": ["05.03!B9:B12"]
        }
    },
    "test_conditions": [
        {"condition_id": "cond_1", "condition_group": "FM No13-2H",
         "line": "VP", "process": "Forming machine",
         "changed_factor": "Forming machine = No13-2H",
         "before_value": None, "after_value": None,
         "unit": None, "machine": "No13-2H", "jig": None, "material_lot": None,
         "supplier": None, "dry_time_sec": None,
         "temperature": "165~180 C", "pressure": "1.2~1.47 kgf/cm2",
         "bond_amount": None, "uv_energy": None,
         "source_file": NAME, "sheet_name": "05.03",
         "source_cells": ["05.03!B17:I17"]},
        {"condition_id": "cond_2", "condition_group": "FM No14-2H",
         "line": "VP", "process": "Forming machine",
         "changed_factor": "Forming machine = No14-2H",
         "before_value": None, "after_value": None,
         "unit": None, "machine": "No14-2H", "jig": None, "material_lot": None,
         "supplier": None, "dry_time_sec": None,
         "temperature": "165~194 C", "pressure": "1.32~1.47 kgf/cm2",
         "bond_amount": None, "uv_energy": None,
         "source_file": NAME, "sheet_name": "05.03",
         "source_cells": ["05.03!B18:I18"]},
        {"condition_id": "cond_3", "condition_group": "FM No14-4H",
         "line": "VP", "process": "Forming machine",
         "changed_factor": "Forming machine = No14-4H",
         "before_value": None, "after_value": None,
         "unit": None, "machine": "No14-4H", "jig": None, "material_lot": None,
         "supplier": None, "dry_time_sec": None,
         "temperature": "171~184 C", "pressure": "1.18~1.24 kgf/cm2",
         "bond_amount": None, "uv_energy": None,
         "source_file": NAME, "sheet_name": "05.03",
         "source_cells": ["05.03!B19:I19"]}
    ],
    "results": [
        {"result_id": "res_1", "condition_id": "cond_1",
         "measurement_type": "Function", "condition_group": "FM No13-2H",
         "date": "2024-03-05", "line": "VP",
         "input_count": None, "ok_count": None, "ng_count": None,
         "ng_rate_decimal": 0.096, "ng_rate_percent": 9.6,
         "metric_name": "Main line Function NG Rate", "metric_value": 9.6,
         "unit": "%", "judgement": "CHECK",
         "ng_breakdown": {},
         "source_file": NAME, "sheet_name": "05.03",
         "source_cells": ["05.03!J17"]},
        {"result_id": "res_2", "condition_id": "cond_2",
         "measurement_type": "Function", "condition_group": "FM No14-2H",
         "date": "2024-03-05", "line": "VP",
         "input_count": None, "ok_count": None, "ng_count": None,
         "ng_rate_decimal": 0.036, "ng_rate_percent": 3.6,
         "metric_name": "Main line Function NG Rate", "metric_value": 3.6,
         "unit": "%", "judgement": "PASS",
         "ng_breakdown": {},
         "source_file": NAME, "sheet_name": "05.03",
         "source_cells": ["05.03!J18"]},
        {"result_id": "res_3", "condition_id": "cond_3",
         "measurement_type": "Function", "condition_group": "FM No14-4H",
         "date": "2024-03-05", "line": "VP",
         "input_count": None, "ok_count": None, "ng_count": None,
         "ng_rate_decimal": 0.089, "ng_rate_percent": 8.9,
         "metric_name": "Main line Function NG Rate", "metric_value": 8.9,
         "unit": "%", "judgement": "CHECK",
         "ng_breakdown": {},
         "source_file": NAME, "sheet_name": "05.03",
         "source_cells": ["05.03!J19"]}
    ],
    "conclusions": [
        {"conclusion_id": "concl_1",
         "topic": "Forming machine vs Function NG rate",
         "statement_from_report": "Decision row in report is blank.",
         "normalized_interpretation": "Among three forming machines, No13-2H gives Function NG 9.6% and No14-4H 8.9% vs No14-2H 3.6%. Temperature on No13-2H (165~180 C) and Pressure on No14-2H (1.32~1.47) all stay inside spec; pressure range on No13-2H (1.2~1.47) and No14-4H (1.18~1.24) and cooling time on No14-4H (1'11) differ. No same-event Normal/Baseline forming-machine row exists, so this is reliability/spec-style comparison; flag higher-NG machines for review.",
         "source_file": NAME, "sheet_name": "05.03",
         "source_cells": ["05.03!A20:A21"]}
    ],
    "troubleshooting_index": {
        "defect_name": "NG Function high rate at Sub VP forming",
        "when_user_asks": ["forming machine", "VP forming", "BRS-201506"],
        "suggested_checks": [
            {"hint_id": "hint_1",
             "check_item": "Audit forming machine No13-2H heating profile and pressure stability",
             "reason": "No13-2H Function NG 9.6%, highest of three; temperature top limit 180 C lowest of the set; pressure ranges up to 1.47.",
             "evidence_strength": "moderate",
             "related_process": "VP forming",
             "related_part": "Forming machine",
             "source_file": NAME, "sheet_name": "05.03",
             "source_cells": ["05.03!B17:J17"]},
            {"hint_id": "hint_2",
             "check_item": "Check No14-4H cooling time (1'11) and pressure low side (1.18)",
             "reason": "No14-4H Function NG 8.9%; cooling time shortest and pressure lower bound lowest of three machines.",
             "evidence_strength": "moderate",
             "related_process": "VP forming",
             "related_part": "Forming machine",
             "source_file": NAME, "sheet_name": "05.03",
             "source_cells": ["05.03!B19:J19"]}
        ],
        "limitations": ["No same-event Normal forming machine row; Input/OK/NG counts not reported, only NG rate."]
    },
    "ai_extraction_log": {
        "confidence": 0.6,
        "assumptions": ["Result column interpreted as main-line Function NG rate per machine."],
        "warnings": ["No Normal/Baseline forming machine row for direct multiplicative comparison.",
                     "Input/OK/NG counts missing; only rate stored."],
        "decision_rationale": "Three forming machines tested same date 2024-03-05 with different temperature/pressure ranges; NG rate ranks No13-2H 9.6% > No14-4H 8.9% > No14-2H 3.6%. Without same-event baseline, no improvement/worsening claim; rank by absolute NG rate and flag conditions for review."
    }
}

tr_en = {
    "document": {"title": result["document"]["title"],
                 "purpose": result["document"]["purpose"],
                 "content": result["document"]["content"]},
    "conclusions": {
        "concl_1": {"topic": "Forming machine vs Function NG rate",
                    "statement_from_report": "Decision row blank in report.",
                    "normalized_interpretation": "FM No13-2H 9.6%, No14-4H 8.9%, No14-2H 3.6%. No same-event baseline, rank machines and flag spec drift."}
    },
    "hints": {
        "hint_1": {"check_item": "Audit FM No13-2H heating profile and pressure stability",
                   "reason": "No13-2H Function NG 9.6%, highest; top temperature 180 C lowest."},
        "hint_2": {"check_item": "Check No14-4H cooling time and low pressure",
                   "reason": "No14-4H 8.9%; cooling time 1'11 shortest and pressure 1.18 lowest."}
    },
    "log": {
        "assumptions": ["Result column = main-line Function NG rate per machine."],
        "warnings": ["No same-event Normal baseline.", "Input/OK/NG counts missing."],
        "decision_rationale": "Rank by absolute NG rate without baseline."
    }
}

tr_ko = {
    "document": {
        "title": "BRS-201506 Sub VP Forming Machine 점검 보고서",
        "purpose": "메인라인 Function NG 고율 원인을 찾기 위해 Sub VP Forming Machine별 분리 시험.",
        "content": [
            "Forming Machine 조건: 온도, 압력, Forming 시간 점검",
            "Forming Machine별 분리: No13-2H, No14-2H, No14-4H",
            "Spec 온도 170~190 C, Cooling 40+-10 C, Pressure 1.2+-0.2 kgf/cm2"
        ]
    },
    "conclusions": {
        "concl_1": {"topic": "Forming Machine별 Function NG Rate",
                    "statement_from_report": "보고서 Decision 행 비어있음.",
                    "normalized_interpretation": "FM No13-2H 9.6%, No14-4H 8.9%, No14-2H 3.6%. 동일 이벤트 baseline 없음. 절대 NG rate 순위로 점검 항목 식별."}
    },
    "hints": {
        "hint_1": {"check_item": "FM No13-2H 히팅 프로파일 및 압력 안정성 점검",
                   "reason": "No13-2H Function NG 9.6% 최고; 온도 상한 180 C로 가장 낮음."},
        "hint_2": {"check_item": "FM No14-4H Cooling Time(1'11) 및 압력 하한(1.18) 점검",
                   "reason": "No14-4H NG 8.9%; Cooling Time 최소, 압력 하한 최저."}
    },
    "log": {
        "assumptions": ["결과 컬럼은 머신별 메인라인 Function NG Rate로 해석."],
        "warnings": ["동일 이벤트 Normal baseline 없음.", "Input/OK/NG 수치 미기재."],
        "decision_rationale": "Baseline 부재 → 절대 NG rate 순위로 정리."
    }
}

tr_vi = {
    "document": {
        "title": "Báo cáo kiểm tra máy Forming Sub VP BRS-201506",
        "purpose": "Main line Function NG cao => kiểm tra từng máy Forming ở Sub VP để tìm nguyên nhân.",
        "content": [
            "Kiểm tra điều kiện máy: nhiệt độ, áp suất, thời gian forming",
            "Tách theo máy No13-2H, No14-2H, No14-4H",
            "Spec nhiệt 170~190 C, Cooling 40+-10 C, áp 1.2+-0.2 kgf/cm2"
        ]
    },
    "conclusions": {
        "concl_1": {"topic": "Máy Forming vs Function NG",
                    "statement_from_report": "Phần Decision trống.",
                    "normalized_interpretation": "FM No13-2H 9.6%, No14-4H 8.9%, No14-2H 3.6%. Không có baseline cùng sự kiện, sắp xếp theo NG tuyệt đối."}
    },
    "hints": {
        "hint_1": {"check_item": "Kiểm tra hồ sơ nhiệt và áp suất máy FM No13-2H",
                   "reason": "No13-2H Function NG 9.6% cao nhất; nhiệt độ giới hạn trên 180 C thấp nhất."},
        "hint_2": {"check_item": "Kiểm tra thời gian Cooling và áp suất thấp của FM No14-4H",
                   "reason": "No14-4H NG 8.9%; Cooling 1'11 ngắn nhất, áp 1.18 thấp nhất."}
    },
    "log": {
        "assumptions": ["Cột Result = NG main-line Function theo máy."],
        "warnings": ["Không có baseline Normal cùng sự kiện.", "Thiếu Input/OK/NG."],
        "decision_rationale": "Không có baseline -> xếp hạng theo NG tuyệt đối."
    }
}

ok = h.commit_dataset(NAME, result, tr_ko, tr_en, tr_vi)
print("ds02", ok)
