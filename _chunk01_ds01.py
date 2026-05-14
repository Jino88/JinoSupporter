# -*- coding: utf-8 -*-
import _ai_batch_helper as h

NAME = "25.BRS-161016 GMI Report checking and test Frame bending date 20.4.2026"

result = {
    "schema_version": "0.1",
    "document": {
        "document_id": "",
        "source_file": NAME,
        "source_sheet": "Test",
        "title": "REPORT CHECKING AND TEST MATERIAL FRAME OK/BENDING BRS-161016",
        "model": "BRS-161016",
        "report_date": "2026-04-20",
        "department": "ME",
        "marker": "Thao",
        "line": "",
        "report_type": "normal_comparison",
        "primary_defect": {
            "canonical_name": "Frame Bending",
            "aliases_in_document": ["NG Frame Bending", "Frame Bending"]
        },
        "related_defects": ["NG Hearing Noise", "NG Hearing Touch", "NG Sigma SPL", "NG Sigma THD"],
        "parts": ["Frame"],
        "processes": ["Frame separation", "Function test"],
        "purpose": "Find reason for NG hearing by separating OK vs Bending frames and comparing function test results.",
        "content": [
            "Check NG rate of Frame Bending",
            "Separate OK / Bending frames",
            "Make final samples and check Function",
            "Q'ty 50pcs"
        ],
        "source_cells": {
            "title": ["Test!C2"],
            "date": ["Test!Q3"],
            "purpose": ["Test!A5:A6"],
            "content": ["Test!A7:A11"]
        }
    },
    "test_conditions": [
        {
            "condition_id": "cond_1",
            "condition_group": "Frame Bending",
            "line": "",
            "process": "Material Frame check",
            "changed_factor": "Frame condition (OK vs Bending)",
            "before_value": "Frame Normal",
            "after_value": "Frame Bending",
            "unit": None, "machine": None, "jig": None, "material_lot": None,
            "supplier": None, "dry_time_sec": None, "temperature": None,
            "pressure": None, "bond_amount": None, "uv_energy": None,
            "source_file": NAME, "sheet_name": "Test",
            "source_cells": ["Test!B14:I16"]
        },
        {
            "condition_id": "cond_2",
            "condition_group": "Function Frame OK",
            "line": "",
            "process": "Function test",
            "changed_factor": "Sample = Frame OK",
            "before_value": None, "after_value": None,
            "unit": None, "machine": None, "jig": None, "material_lot": None,
            "supplier": None, "dry_time_sec": None, "temperature": None,
            "pressure": None, "bond_amount": None, "uv_energy": None,
            "source_file": NAME, "sheet_name": "Test",
            "source_cells": ["Test!B20"]
        },
        {
            "condition_id": "cond_3",
            "condition_group": "Function Frame Bending",
            "line": "",
            "process": "Function test",
            "changed_factor": "Sample = Frame Bending",
            "before_value": None, "after_value": None,
            "unit": None, "machine": None, "jig": None, "material_lot": None,
            "supplier": None, "dry_time_sec": None, "temperature": None,
            "pressure": None, "bond_amount": None, "uv_energy": None,
            "source_file": NAME, "sheet_name": "Test",
            "source_cells": ["Test!B22"]
        },
        {
            "condition_id": "cond_4",
            "condition_group": "Function Normal",
            "line": "",
            "process": "Function test",
            "changed_factor": "Sample = Normal (baseline)",
            "before_value": None, "after_value": None,
            "unit": None, "machine": None, "jig": None, "material_lot": None,
            "supplier": None, "dry_time_sec": None, "temperature": None,
            "pressure": None, "bond_amount": None, "uv_energy": None,
            "source_file": NAME, "sheet_name": "Test",
            "source_cells": ["Test!B24"]
        }
    ],
    "results": [
        {
            "result_id": "res_1", "condition_id": "cond_1",
            "measurement_type": "Material", "condition_group": "Frame Bending 4/20",
            "date": "2026-04-20", "line": "",
            "input_count": 349, "ok_count": 322, "ng_count": 27,
            "ng_rate_decimal": 0.077, "ng_rate_percent": 7.7,
            "metric_name": "Frame Bending NG Rate", "metric_value": 7.7,
            "unit": "%", "judgement": None,
            "ng_breakdown": {"NG Frame Bending": {"count": 27, "rate": 0.077}},
            "source_file": NAME, "sheet_name": "Test",
            "source_cells": ["Test!B15:I15"]
        },
        {
            "result_id": "res_2", "condition_id": "cond_1",
            "measurement_type": "Material", "condition_group": "Frame Bending 4/29",
            "date": "2026-04-29", "line": "",
            "input_count": 100, "ok_count": 100, "ng_count": 5,
            "ng_rate_decimal": 0.05, "ng_rate_percent": 5.0,
            "metric_name": "Frame Bending NG Rate", "metric_value": 5.0,
            "unit": "%", "judgement": None,
            "ng_breakdown": {"NG Frame Bending": {"count": 5, "rate": 0.05}},
            "source_file": NAME, "sheet_name": "Test",
            "source_cells": ["Test!B16:I16"]
        },
        {
            "result_id": "res_3", "condition_id": "cond_2",
            "measurement_type": "Function", "condition_group": "Frame OK",
            "date": "2026-04-21", "line": "",
            "input_count": 290, "ok_count": 275, "ng_count": 15,
            "ng_rate_decimal": 0.052, "ng_rate_percent": 5.2,
            "metric_name": "Function Total NG Rate", "metric_value": 5.2,
            "unit": "%", "judgement": None,
            "ng_breakdown": {
                "NG Sigma THD": {"count": 0, "rate": 0.0},
                "NG Sigma SPL": {"count": 0, "rate": 0.0},
                "NG Sigma SPL+THD": {"count": 1, "rate": 0.003},
                "NG Hearing Noise": {"count": 9, "rate": 0.031},
                "NG Hearing Touch": {"count": 5, "rate": 0.017}
            },
            "source_file": NAME, "sheet_name": "Test",
            "source_cells": ["Test!B20:L21"]
        },
        {
            "result_id": "res_4", "condition_id": "cond_3",
            "measurement_type": "Function", "condition_group": "Frame Bending",
            "date": "2026-04-21", "line": "",
            "input_count": 22, "ok_count": 19, "ng_count": 3,
            "ng_rate_decimal": 0.136, "ng_rate_percent": 13.6,
            "metric_name": "Function Total NG Rate", "metric_value": 13.6,
            "unit": "%", "judgement": None,
            "ng_breakdown": {
                "NG Sigma THD": {"count": 0, "rate": 0.0},
                "NG Sigma SPL": {"count": 0, "rate": 0.0},
                "NG Sigma SPL+THD": {"count": 0, "rate": 0.0},
                "NG Hearing Noise": {"count": 3, "rate": 0.136},
                "NG Hearing Touch": {"count": 0, "rate": 0.0}
            },
            "source_file": NAME, "sheet_name": "Test",
            "source_cells": ["Test!B22:L23"]
        },
        {
            "result_id": "res_5", "condition_id": "cond_4",
            "measurement_type": "Function", "condition_group": "Normal",
            "date": "2026-04-21", "line": "",
            "input_count": 792, "ok_count": 742, "ng_count": 50,
            "ng_rate_decimal": 0.063, "ng_rate_percent": 6.3,
            "metric_name": "Function Total NG Rate", "metric_value": 6.3,
            "unit": "%", "judgement": None,
            "ng_breakdown": {
                "NG Sigma THD": {"count": 3, "rate": 0.004},
                "NG Sigma SPL": {"count": 3, "rate": 0.004},
                "NG Sigma SPL+THD": {"count": 3, "rate": 0.004},
                "NG Hearing Noise": {"count": 32, "rate": 0.04},
                "NG Hearing Touch": {"count": 9, "rate": 0.011}
            },
            "source_file": NAME, "sheet_name": "Test",
            "source_cells": ["Test!B24:L25"]
        }
    ],
    "conclusions": [
        {
            "conclusion_id": "concl_1", "topic": "Frame Bending vs Normal function impact",
            "statement_from_report": "Result checking material Frame happen NG bending rate 5~7% => separate checking Function happen NG high rate 13.6%, normal NG rate 6.3%",
            "normalized_interpretation": "Frame Bending samples show Function NG rate 13.6% vs Normal 6.3% same event => (13.6/6.3 - 1)*100 = +115.9%, 115.9% worse than same-event normal. Dominant defect is NG Hearing Noise (13.6% on bending vs 4.0% on normal).",
            "source_file": NAME, "sheet_name": "Test",
            "source_cells": ["Test!A27"]
        }
    ],
    "troubleshooting_index": {
        "defect_name": "NG Hearing Noise (Frame Bending)",
        "when_user_asks": ["frame bending", "hearing noise", "BRS-161016"],
        "suggested_checks": [
            {
                "hint_id": "hint_1",
                "check_item": "Inspect Frame bending rate at incoming material",
                "reason": "Incoming Frame Bending NG was 5.0~7.7%; bending samples show Function NG 13.6% vs Normal 6.3% (+115.9% worse).",
                "evidence_strength": "strong",
                "related_process": "Material Frame inspection",
                "related_part": "Frame",
                "source_file": NAME, "sheet_name": "Test",
                "source_cells": ["Test!B15:L25"]
            },
            {
                "hint_id": "hint_2",
                "check_item": "Focus on Hearing Noise defect mode when frame bending suspected",
                "reason": "On bent frames Hearing Noise = 13.6%; on normal 4.0%; other defect modes near 0 on bent samples.",
                "evidence_strength": "strong",
                "related_process": "Function test",
                "related_part": "Frame",
                "source_file": NAME, "sheet_name": "Test",
                "source_cells": ["Test!H22:H25"]
            }
        ],
        "limitations": ["Bending sample size only 22pcs, small vs Normal 792pcs."]
    },
    "ai_extraction_log": {
        "confidence": 0.85,
        "assumptions": ["report_date 20/Apr interpreted as 2026-04-20 per dataset name", "Marker = Thao"],
        "warnings": ["Bending sample N=22 small; relative change has wide uncertainty."],
        "decision_rationale": "Same-event Normal baseline present (Frame OK 5.2%, Normal 6.3%, Bending 13.6%). Bending vs Normal: (13.6/6.3-1)*100=+115.9% worse. Hearing Noise dominates the NG mix on bending samples."
    }
}

tr_en = {
    "document": {
        "title": result["document"]["title"],
        "purpose": result["document"]["purpose"],
        "content": result["document"]["content"]
    },
    "conclusions": {
        "concl_1": {
            "topic": "Frame Bending vs Normal function impact",
            "statement_from_report": "Material Frame NG bending rate 5~7%; Frame Bending function NG 13.6%, Normal NG 6.3%.",
            "normalized_interpretation": "Frame Bending Function NG 13.6% vs Normal 6.3% => 115.9% worse than same-event normal. NG Hearing Noise dominates."
        }
    },
    "hints": {
        "hint_1": {"check_item": "Inspect Frame bending rate at incoming material",
                   "reason": "Incoming Frame Bending NG 5.0~7.7%; bent frames show Function NG 13.6% vs Normal 6.3% (+115.9% worse)."},
        "hint_2": {"check_item": "Focus on Hearing Noise when frame bending suspected",
                   "reason": "On bent frames Hearing Noise 13.6%; on normal 4.0%; other modes near zero."}
    },
    "log": {
        "assumptions": ["report_date 20/Apr = 2026-04-20", "Marker = Thao"],
        "warnings": ["Bending sample N=22 small."],
        "decision_rationale": "Normal baseline exists in same event. Bending vs Normal (13.6/6.3-1)*100=+115.9% worse, Hearing Noise dominant."
    }
}

tr_ko = {
    "document": {
        "title": "BRS-161016 Frame OK/Bending 자재 검사 및 시험 보고서",
        "purpose": "Frame Bending OK/NG을 분리하여 Function 결과를 비교하고 NG Hearing 원인을 찾는다.",
        "content": [
            "Frame Bending NG Rate 점검",
            "OK / Bending Frame 분리",
            "최종 샘플로 Function 점검",
            "Q'ty 50pcs"
        ]
    },
    "conclusions": {
        "concl_1": {
            "topic": "Frame Bending vs Normal Function 영향",
            "statement_from_report": "Frame 자재 Bending NG 5~7%, Bending 샘플 Function NG 13.6%, Normal NG 6.3%.",
            "normalized_interpretation": "Frame Bending Function NG 13.6% vs Normal 6.3% => 동일 이벤트 normal 대비 115.9% 악화. NG Hearing Noise가 주된 불량."
        }
    },
    "hints": {
        "hint_1": {"check_item": "수입 자재 단계에서 Frame Bending 비율 점검",
                   "reason": "수입 Frame Bending NG 5.0~7.7%, Bending 샘플 Function NG 13.6% vs Normal 6.3%(+115.9% 악화)."},
        "hint_2": {"check_item": "Frame Bending 의심 시 Hearing Noise 항목에 집중",
                   "reason": "Bending 샘플 Hearing Noise 13.6%, Normal 4.0%, 다른 모드는 거의 0."}
    },
    "log": {
        "assumptions": ["report_date 20/Apr = 2026-04-20", "Marker = Thao"],
        "warnings": ["Bending 샘플 수 22pcs로 적음."],
        "decision_rationale": "동일 이벤트 Normal baseline 존재. Bending vs Normal (13.6/6.3-1)*100=+115.9% 악화, Hearing Noise 주도."
    }
}

tr_vi = {
    "document": {
        "title": "Báo cáo kiểm tra và test vật liệu Frame OK/Bending BRS-161016",
        "purpose": "Tách Frame OK và Bending để so sánh Function, tìm nguyên nhân NG Hearing.",
        "content": [
            "Kiểm tra NG rate Frame Bending",
            "Tách OK / Bending Frame",
            "Làm sample final kiểm tra Function",
            "Số lượng 50pcs"
        ]
    },
    "conclusions": {
        "concl_1": {
            "topic": "Ảnh hưởng của Frame Bending so với Normal",
            "statement_from_report": "Frame vật liệu NG bending 5~7%; Frame Bending Function NG 13.6%, Normal 6.3%.",
            "normalized_interpretation": "Frame Bending Function NG 13.6% so với Normal 6.3% => xấu hơn 115.9% so với normal cùng sự kiện. NG Hearing Noise chiếm chủ yếu."
        }
    },
    "hints": {
        "hint_1": {"check_item": "Kiểm tra tỉ lệ Frame Bending ở vật liệu đầu vào",
                   "reason": "Vật liệu Frame Bending NG 5.0~7.7%; mẫu bị bending Function NG 13.6% vs Normal 6.3% (+115.9% xấu hơn)."},
        "hint_2": {"check_item": "Tập trung kiểm tra Hearing Noise khi nghi Frame Bending",
                   "reason": "Trên mẫu bending Hearing Noise 13.6%; Normal 4.0%; các mode khác gần 0."}
    },
    "log": {
        "assumptions": ["report_date 20/Apr = 2026-04-20", "Marker = Thao"],
        "warnings": ["Mẫu Bending chỉ 22pcs, nhỏ."],
        "decision_rationale": "Có Normal baseline cùng sự kiện. Bending vs Normal (13.6/6.3-1)*100=+115.9% xấu hơn, Hearing Noise dẫn dắt."
    }
}

ok = h.commit_dataset(NAME, result, tr_ko, tr_en, tr_vi)
print("ds01", ok)
