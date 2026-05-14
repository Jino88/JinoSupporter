"""
Per-dataset normalized results for chunk_12.
Each entry: {DatasetName: {result, tr_ko, tr_vi}}
Then commit per AI_EXCEL_PROC.md §4.
"""
import sqlite3, json, os
from datetime import datetime, timezone

DB = r"D:\000. MyWorks\002. DB\process-review.db"
CHUNK = r"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_batch_chunks\chunk_12.txt"

RESULTS = {}

# ──────────────────────────────────────────────────────────────────────────
# 0. 17.1 BRS-161014 C2+E2  Report test Center Dome improve bending of Ralon vender 16.2.2024
# Test CD Ralon vs Normal CD Ges on E2+C2; function NG: Ralon 3.8% vs Ges 1.9% → cannot use.
RESULTS["17.1 BRS-161014 C2+E2  Report test Center Dome improve bending of Ralon vender 16.2.2024"] = {
    "result": {
        "measurements": [
            # Section 1 — vision VP+CD Sub1, E2
            {"productType":"BRS-161014","testDate":"2024-02-16","line":"E2","checkType":"visual_inspection",
             "variable":"CD vendor","variableDetail":"Vision VP+CD Sub1","variableGroup":"test","intervention":"Test CD of Ralon",
             "inputQty":112,"okQty":112,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161014","testDate":"2024-02-16","line":"E2","checkType":"visual_inspection",
             "variable":"CD vendor","variableDetail":"Vision VP+CD Sub1","variableGroup":"normal","intervention":"Normal CD of Ges",
             "inputQty":112,"okQty":111,"ngTotal":1,"ngRate":0.9,"defectCategory":"assembly_defect","defectType":"VP separate","defectCount":1},
            # Section 1 — vision VP+CD Sub1, C2
            {"productType":"BRS-161014","testDate":"2024-02-16","line":"C2","checkType":"visual_inspection",
             "variable":"CD vendor","variableDetail":"Vision VP+CD Sub1","variableGroup":"test","intervention":"Test CD supplier Ralon",
             "inputQty":112,"okQty":112,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161014","testDate":"2024-02-16","line":"C2","checkType":"visual_inspection",
             "variable":"CD vendor","variableDetail":"Vision VP+CD Sub1","variableGroup":"normal","intervention":"Normal CD supplier GES",
             "inputQty":112,"okQty":112,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            # Section 2 — Function E2
            {"productType":"BRS-161014","testDate":"2024-02-16","line":"E2","checkType":"function",
             "variable":"CD vendor","variableDetail":"Function","variableGroup":"test","intervention":"Test CD of Ralon",
             "inputQty":103,"okQty":99,"ngTotal":4,"ngRate":3.9,"defectCategory":"function_hearing","defectType":"Noise","defectCount":4},
            {"productType":"BRS-161014","testDate":"2024-02-16","line":"E2","checkType":"function",
             "variable":"CD vendor","variableDetail":"Function","variableGroup":"normal","intervention":"Normal CD of GES",
             "inputQty":105,"okQty":105,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            # Section 2 — Function C2
            {"productType":"BRS-161014","testDate":"2024-02-16","line":"C2","checkType":"function",
             "variable":"CD vendor","variableDetail":"Function","variableGroup":"test","intervention":"Test CD supplier Ralon",
             "inputQty":107,"okQty":103,"ngTotal":4,"ngRate":3.7,"defectCategory":"function_hearing","defectType":"Noise","defectCount":4},
            {"productType":"BRS-161014","testDate":"2024-02-16","line":"C2","checkType":"function",
             "variable":"CD vendor","variableDetail":"Function","variableGroup":"normal","intervention":"Normal CD supplier GES",
             "inputQty":107,"okQty":103,"ngTotal":4,"ngRate":3.7,"defectCategory":"function_hearing","defectType":"Noise","defectCount":4},
            # Section 2 — Function Total C2+E2
            {"productType":"BRS-161014","testDate":"2024-02-16","line":"C2+E2","checkType":"function",
             "variable":"CD vendor","variableDetail":"Function total","variableGroup":"test","intervention":"Test CD supplier Ralon",
             "inputQty":210,"okQty":202,"ngTotal":8,"ngRate":3.8,"defectCategory":"function_hearing","defectType":"Noise","defectCount":8},
            {"productType":"BRS-161014","testDate":"2024-02-16","line":"C2+E2","checkType":"function",
             "variable":"CD vendor","variableDetail":"Function total","variableGroup":"normal","intervention":"Normal CD supplier GES",
             "inputQty":212,"okQty":208,"ngTotal":4,"ngRate":1.9,"defectCategory":"function_hearing","defectType":"Noise","defectCount":4},
        ],
        "tags":["brs-161014","cd-vendor","ralon","ges","center-dome","function-test","hearing-noise","vendor-comparison","comparison-study"],
        "reportType":"comparison_study",
        "verdict":"worsened",
        "headline":"Ralon CD function NG 3.8% vs Ges 1.9% (+1.9pp, worsened)",
        "evidence":[
            {"metric":"Function NG rate (C2+E2)","baselineLabel":"Normal CD Ges","baselineValue":"1.9% (4/212)",
             "variantLabel":"Test CD Ralon","variantValue":"3.8% (8/210)",
             "deltaText":"+1.9pp","deltaSign":"up","note":"Hearing-Noise dominant",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Vision VP+CD (Sub1)","baselineLabel":"Normal Ges","baselineValue":"0.9% (1/112) E2",
             "variantLabel":"Test Ralon","variantValue":"0.0% (0/112) E2",
             "deltaText":"-0.9pp","deltaSign":"down","note":"",
             "comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Reject Ralon CD lot — do not release to production"},
            {"priority":2,"kind":"investigate","text":"Find why Ralon CD raises hearing-noise NG nearly 2x"},
        ],
        "context":{"process":"Center Dome (CD) vendor change verification at Sub1 + main-line function",
                   "stage":"C2 and E2 lines, AWF #1","baselineReason":"same-event Normal CD Ges rows present"},
        "doeGrid":None,"trendPoints":None,
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
        "productType":"BRS-161014",
    },
    "tr_ko": {
        "headline":"Ralon CD 기능 NG 3.8% vs Ges 1.9% (+1.9pp, 악화)",
        "actions":[
            {"priority":1,"kind":"action","text":"Ralon CD 로트 거부 — 양산 투입 금지"},
            {"priority":2,"kind":"investigate","text":"Ralon CD가 hearing-noise NG를 2배 가까이 높이는 원인 조사"},
        ],
        "context":{"process":"Sub1 + 메인라인 기능에서 Center Dome(CD) 벤더 변경 검증",
                   "stage":"C2 및 E2 라인, AWF #1","baselineReason":"동일 이벤트에 Normal CD Ges 행 존재"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
    "tr_vi": {
        "headline":"Ralon CD chức năng NG 3.8% vs Ges 1.9% (+1.9pp, xấu đi)",
        "actions":[
            {"priority":1,"kind":"action","text":"Từ chối lô CD Ralon — không đưa vào sản xuất"},
            {"priority":2,"kind":"investigate","text":"Tìm nguyên nhân CD Ralon tăng NG hearing-noise gấp 2 lần"},
        ],
        "context":{"process":"Xác minh thay đổi nhà cung cấp Center Dome (CD) tại Sub1 + chức năng line chính",
                   "stage":"Line C2 và E2, AWF #1","baselineReason":"có dòng Normal CD Ges cùng sự kiện"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
}

# ──────────────────────────────────────────────────────────────────────────
# 1. 17.2 BRS 161016 Report check dimension compare CD vendor Ralon and GES 2024.02.18
# Pure dimension check — Ralon vs Ges. No NG counts. quality_log style.
RESULTS["17.2 BRS 161016 Report check dimension compare CD vendor Ralon and GES 2024.02.18"] = {
    "result": {
        "measurements": [
            {"productType":"BRS-161016","testDate":"2024-02-18","line":"","checkType":"process",
             "variable":"CD dimension","variableDetail":"Dimension P1-P6 avg","variableGroup":"test","intervention":"Ralon CD",
             "inputQty":5,"okQty":5,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161016","testDate":"2024-02-18","line":"","checkType":"process",
             "variable":"CD dimension","variableDetail":"Dimension P1-P6 avg","variableGroup":"normal","intervention":"Ges CD",
             "inputQty":5,"okQty":5,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        ],
        "tags":["brs-161016","cd-vendor","ralon","ges","dimension-check","vendor-comparison","quality-log"],
        "reportType":"comparison_study",
        "verdict":"no_clear_effect",
        "headline":"Ralon CD avg 0.632mm vs Ges 0.677mm (-0.045mm gap, dimension differs)",
        "evidence":[
            {"metric":"CD avg dimension (P1-P6)","baselineLabel":"Ges","baselineValue":"0.669-0.680 mm avg",
             "variantLabel":"Ralon","variantValue":"0.631-0.657 mm avg",
             "deltaText":"-0.010 to -0.046 mm","deltaSign":"down","note":"Same height but visibly different",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"CD Posi1 (height)","baselineLabel":"Ges","baselineValue":"5.988 mm avg",
             "variantLabel":"Ralon","variantValue":"5.972 mm avg",
             "deltaText":"-0.016 mm","deltaSign":"down","note":"n=5",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"CD Posi2 (height)","baselineLabel":"Ges","baselineValue":"11.959 mm avg",
             "variantLabel":"Ralon","variantValue":"11.971 mm avg",
             "deltaText":"+0.012 mm","deltaSign":"up","note":"n=5",
             "comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"investigate","text":"Confirm whether 0.04mm dimension gap explains function NG difference"},
            {"priority":2,"kind":"action","text":"Re-validate new measurement method (gap -0.06 mm vs old method)"},
        ],
        "context":{"process":"Center Dome dimension comparison Ralon vs Ges, 5 samples each",
                   "stage":"Off-line dimension lab","baselineReason":"Ges treated as standard incumbent vendor"},
        "doeGrid":None,"trendPoints":None,
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
        "productType":"BRS-161016",
    },
    "tr_ko": {
        "headline":"Ralon CD 평균 0.632mm vs Ges 0.677mm (-0.045mm 차이, 치수 상이)",
        "actions":[
            {"priority":1,"kind":"investigate","text":"0.04mm 치수 차이가 기능 NG 차이의 원인인지 확인"},
            {"priority":2,"kind":"action","text":"신규 측정 방법 재검증 (기존 대비 -0.06mm 차이)"},
        ],
        "context":{"process":"Center Dome 치수 비교 Ralon vs Ges, 샘플 각 5개",
                   "stage":"오프라인 치수 측정실","baselineReason":"Ges를 기존 표준 벤더로 간주"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
    "tr_vi": {
        "headline":"Ralon CD trung bình 0.632mm vs Ges 0.677mm (-0.045mm, khác kích thước)",
        "actions":[
            {"priority":1,"kind":"investigate","text":"Kiểm tra chênh lệch 0.04mm có giải thích khác biệt NG chức năng"},
            {"priority":2,"kind":"action","text":"Tái xác nhận phương pháp đo mới (gap -0.06 mm so với phương pháp cũ)"},
        ],
        "context":{"process":"So sánh kích thước Center Dome Ralon vs Ges, 5 mẫu mỗi loại",
                   "stage":"Phòng đo offline","baselineReason":"Ges là nhà cung cấp tiêu chuẩn hiện tại"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
}

# ──────────────────────────────────────────────────────────────────────────
# 2. 17.3 BRS-161014 Summary data TEST compare MTR Ralon and Ges Vendor 19.2.2024
# Aggregated summary: Ralon CD 5.6% vs Ges 2.8% function NG → cannot use
RESULTS["17.3 BRS-161014 Summary data TEST compare MTR Ralon and Ges Vendor 19.2.2024"] = {
    "result": {
        "measurements": [
            # Jan 9 C2 — Normal Ralon Dome
            {"productType":"BRS-161014","testDate":"2024-01-09","line":"C2","checkType":"function",
             "variable":"CD vendor","variableDetail":"Function AWF #3","variableGroup":"test","intervention":"Normal C2 use new Dome Ralon",
             "inputQty":60,"okQty":58,"ngTotal":2,"ngRate":3.3,"defectCategory":"function_hearing","defectType":"Noise","defectCount":2},
            {"productType":"BRS-161014","testDate":"2024-01-09","line":"C2","checkType":"function",
             "variable":"CD vendor","variableDetail":"Function AWF #5","variableGroup":"test","intervention":"Normal C2 use new Dome Ralon",
             "inputQty":60,"okQty":57,"ngTotal":3,"ngRate":5.0,"defectCategory":"function_hearing","defectType":"Noise","defectCount":3},
            # Jan 9 — Test semi VP/Dome E2 Ges Dome
            {"productType":"BRS-161014","testDate":"2024-01-09","line":"E2","checkType":"function",
             "variable":"CD vendor","variableDetail":"Function AWF #1","variableGroup":"normal","intervention":"Test semi VP/Dome E2 use Old Dome Ges",
             "inputQty":70,"okQty":70,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161014","testDate":"2024-01-09","line":"E2","checkType":"function",
             "variable":"CD vendor","variableDetail":"Function AWF #2","variableGroup":"normal","intervention":"Test semi VP/Dome E2 use Old Dome Ges",
             "inputQty":70,"okQty":69,"ngTotal":1,"ngRate":1.4,"defectCategory":"function_hearing","defectType":"Noise","defectCount":1},
            {"productType":"BRS-161014","testDate":"2024-01-09","line":"E2","checkType":"function",
             "variable":"CD vendor","variableDetail":"Function AWF #4","variableGroup":"normal","intervention":"Test semi VP/Dome E2 use Old Dome Ges",
             "inputQty":59,"okQty":58,"ngTotal":1,"ngRate":1.7,"defectCategory":"function_hearing","defectType":"Noise","defectCount":1},
            # Jan 11 — Dome Ges
            {"productType":"BRS-161014","testDate":"2024-01-11","line":"","checkType":"function",
             "variable":"CD vendor","variableDetail":"Function All M/C","variableGroup":"normal","intervention":"Dome vendor Ges",
             "inputQty":300,"okQty":298,"ngTotal":2,"ngRate":0.7,"defectCategory":"function_hearing","defectType":"Touch","defectCount":2},
            # Jan 11 — Dome Ralon
            {"productType":"BRS-161014","testDate":"2024-01-11","line":"","checkType":"function",
             "variable":"CD vendor","variableDetail":"Function","variableGroup":"test","intervention":"Dome vendor Ralon",
             "inputQty":521,"okQty":499,"ngTotal":22,"ngRate":4.2,"defectCategory":"function_hearing","defectType":"Noise","defectCount":9},
            {"productType":"BRS-161014","testDate":"2024-01-11","line":"","checkType":"function",
             "variable":"CD vendor","variableDetail":"Function","variableGroup":"test","intervention":"Dome vendor Ralon",
             "inputQty":521,"okQty":499,"ngTotal":22,"ngRate":4.2,"defectCategory":"function_hearing","defectType":"Touch","defectCount":13},
            # Feb 6 C2 — Test Ralon
            {"productType":"BRS-161014","testDate":"2024-02-06","line":"C2","checkType":"function",
             "variable":"CD vendor","variableDetail":"Function All M/C","variableGroup":"test","intervention":"Test CD supplier Ralon",
             "inputQty":2644,"okQty":2488,"ngTotal":156,"ngRate":5.9,"defectCategory":"function_hearing","defectType":"Noise","defectCount":103},
            {"productType":"BRS-161014","testDate":"2024-02-06","line":"C2","checkType":"function",
             "variable":"CD vendor","variableDetail":"Function All M/C","variableGroup":"test","intervention":"Test CD supplier Ralon",
             "inputQty":2644,"okQty":2488,"ngTotal":156,"ngRate":5.9,"defectCategory":"function_hearing","defectType":"Touch","defectCount":52},
            {"productType":"BRS-161014","testDate":"2024-02-06","line":"C2","checkType":"function",
             "variable":"CD vendor","variableDetail":"Function All M/C","variableGroup":"test","intervention":"Test CD supplier Ralon",
             "inputQty":2644,"okQty":2488,"ngTotal":156,"ngRate":5.9,"defectCategory":"function_spl","defectType":"SPL+THD","defectCount":1},
            # Feb 6 C2 — Normal Ges
            {"productType":"BRS-161014","testDate":"2024-02-06","line":"C2","checkType":"function",
             "variable":"CD vendor","variableDetail":"Function All M/C","variableGroup":"normal","intervention":"Normal CD supplier GES",
             "inputQty":2215,"okQty":2142,"ngTotal":73,"ngRate":3.3,"defectCategory":"function_hearing","defectType":"Noise","defectCount":54},
            {"productType":"BRS-161014","testDate":"2024-02-06","line":"C2","checkType":"function",
             "variable":"CD vendor","variableDetail":"Function All M/C","variableGroup":"normal","intervention":"Normal CD supplier GES",
             "inputQty":2215,"okQty":2142,"ngTotal":73,"ngRate":3.3,"defectCategory":"function_hearing","defectType":"Touch","defectCount":19},
            # Summary all-time
            {"productType":"BRS-161014","testDate":"2024-02-19","line":"All","checkType":"function",
             "variable":"CD vendor","variableDetail":"Summary all-time test","variableGroup":"test","intervention":"Test CD supplier Ralon",
             "inputQty":3435,"okQty":3244,"ngTotal":191,"ngRate":5.6,"defectCategory":"function_hearing","defectType":"Noise","defectCount":125},
            {"productType":"BRS-161014","testDate":"2024-02-19","line":"All","checkType":"function",
             "variable":"CD vendor","variableDetail":"Summary all-time test","variableGroup":"test","intervention":"Test CD supplier Ralon",
             "inputQty":3435,"okQty":3244,"ngTotal":191,"ngRate":5.6,"defectCategory":"function_hearing","defectType":"Touch","defectCount":65},
            {"productType":"BRS-161014","testDate":"2024-02-19","line":"All","checkType":"function",
             "variable":"CD vendor","variableDetail":"Summary all-time normal","variableGroup":"normal","intervention":"Normal CD supplier GES",
             "inputQty":2926,"okQty":2845,"ngTotal":81,"ngRate":2.8,"defectCategory":"function_hearing","defectType":"Noise","defectCount":60},
            {"productType":"BRS-161014","testDate":"2024-02-19","line":"All","checkType":"function",
             "variable":"CD vendor","variableDetail":"Summary all-time normal","variableGroup":"normal","intervention":"Normal CD supplier GES",
             "inputQty":2926,"okQty":2845,"ngTotal":81,"ngRate":2.8,"defectCategory":"function_hearing","defectType":"Touch","defectCount":21},
        ],
        "tags":["brs-161014","cd-vendor","ralon","ges","material-comparison","summary-roll-up","function-test","hearing-noise","comparison-study"],
        "reportType":"comparison_study",
        "verdict":"worsened",
        "headline":"All-time Ralon CD function NG 5.6% vs Ges 2.8% (+2.8pp, worsened)",
        "evidence":[
            {"metric":"Function NG rate (all-time)","baselineLabel":"Normal Ges","baselineValue":"2.8% (81/2926)",
             "variantLabel":"Test Ralon","variantValue":"5.6% (191/3435)",
             "deltaText":"+2.8pp","deltaSign":"up","note":"n>2900 each, high confidence",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Hearing-Noise share (Ralon)","baselineLabel":"Ges","baselineValue":"2.1% (60/2926)",
             "variantLabel":"Ralon","variantValue":"3.6% (125/3435)",
             "deltaText":"+1.5pp","deltaSign":"up","note":"",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Hearing-Touch share (Ralon)","baselineLabel":"Ges","baselineValue":"0.7% (21/2926)",
             "variantLabel":"Ralon","variantValue":"1.9% (65/3435)",
             "deltaText":"+1.2pp","deltaSign":"up","note":"",
             "comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Do not adopt Ralon CD — keep Ges as primary vendor"},
            {"priority":2,"kind":"investigate","text":"Identify CD material/dimension difference driving Ralon hearing-noise"},
        ],
        "context":{"process":"Roll-up of Ralon vs Ges Center Dome trials over Jan-Feb across C2 and E2",
                   "stage":"All M/C, multi-date summary","baselineReason":"same-event Normal Ges rows on each test date"},
        "doeGrid":None,"trendPoints":None,
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
        "productType":"BRS-161014",
    },
    "tr_ko": {
        "headline":"전체 기간 Ralon CD 기능 NG 5.6% vs Ges 2.8% (+2.8pp, 악화)",
        "actions":[
            {"priority":1,"kind":"action","text":"Ralon CD 채택 보류 — Ges를 주 벤더로 유지"},
            {"priority":2,"kind":"investigate","text":"Ralon hearing-noise를 유발하는 CD 재질/치수 차이 파악"},
        ],
        "context":{"process":"1-2월 C2 및 E2에서 진행한 Ralon vs Ges Center Dome 시험 누적 정리",
                   "stage":"전체 M/C, 다일자 요약","baselineReason":"각 시험일에 동일 이벤트 Normal Ges 행 존재"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
    "tr_vi": {
        "headline":"Toàn thời gian Ralon CD NG chức năng 5.6% vs Ges 2.8% (+2.8pp, xấu đi)",
        "actions":[
            {"priority":1,"kind":"action","text":"Không áp dụng CD Ralon — giữ Ges là nhà cung cấp chính"},
            {"priority":2,"kind":"investigate","text":"Xác định khác biệt vật liệu/kích thước CD gây hearing-noise của Ralon"},
        ],
        "context":{"process":"Tổng hợp các thử nghiệm Center Dome Ralon vs Ges trong tháng 1-2 trên C2 và E2",
                   "stage":"Toàn bộ M/C, tóm tắt đa ngày","baselineReason":"có dòng Normal Ges cùng sự kiện ở mỗi ngày thử"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
}

# ──────────────────────────────────────────────────────────────────────────
# 3. 17.BRS-161016 Result checking and test problem VP+CD separate date 20.10.2025
# DOE: Laser CD x Plasma → check VP+CD separate tension. 4 combos × 8-10 samples.
RESULTS["17.BRS-161016 Result checking and test problem VP+CD separate date 20.10.2025"] = {
    "result": {
        "measurements": [
            # Section 15.10 — VP+CD separate combo grid
            {"productType":"BRS-161016","testDate":"2025-10-20","line":"","checkType":"visual_inspection",
             "variable":"Laser+Plasma","variableDetail":"VP+CD separate Laser=X Plasma=X","variableGroup":"test","intervention":"No laser No plasma",
             "inputQty":83,"okQty":83,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161016","testDate":"2025-10-20","line":"","checkType":"visual_inspection",
             "variable":"Laser+Plasma","variableDetail":"VP+CD separate Laser=O Plasma=O","variableGroup":"test","intervention":"Have laser have plasma",
             "inputQty":100,"okQty":100,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161016","testDate":"2025-10-20","line":"","checkType":"visual_inspection",
             "variable":"Laser+Plasma","variableDetail":"VP+CD separate Laser=X Plasma=O","variableGroup":"test","intervention":"No laser have plasma",
             "inputQty":67,"okQty":67,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161016","testDate":"2025-10-20","line":"","checkType":"visual_inspection",
             "variable":"Laser+Plasma","variableDetail":"VP+CD separate Laser=O Plasma=X","variableGroup":"test","intervention":"Have laser no plasma",
             "inputQty":88,"okQty":88,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            # Function 4 combos
            {"productType":"BRS-161016","testDate":"2025-10-21","line":"","checkType":"function",
             "variable":"Laser+Plasma","variableDetail":"Function Laser=X Plasma=X","variableGroup":"test","intervention":"No laser No plasma",
             "inputQty":75,"okQty":75,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161016","testDate":"2025-10-21","line":"","checkType":"function",
             "variable":"Laser+Plasma","variableDetail":"Function Laser=O Plasma=O","variableGroup":"test","intervention":"Have laser have plasma",
             "inputQty":92,"okQty":92,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161016","testDate":"2025-10-21","line":"","checkType":"function",
             "variable":"Laser+Plasma","variableDetail":"Function Laser=X Plasma=O","variableGroup":"test","intervention":"No laser have plasma",
             "inputQty":59,"okQty":58,"ngTotal":1,"ngRate":1.7,"defectCategory":"function_hearing","defectType":"Noise","defectCount":1},
            {"productType":"BRS-161016","testDate":"2025-10-21","line":"","checkType":"function",
             "variable":"Laser+Plasma","variableDetail":"Function Laser=O Plasma=X","variableGroup":"test","intervention":"Have laser no plasma",
             "inputQty":81,"okQty":80,"ngTotal":1,"ngRate":1.2,"defectCategory":"function_hearing","defectType":"Noise","defectCount":1},
            # 23.10 VP/CD separate
            {"productType":"BRS-161016","testDate":"2025-10-23","line":"","checkType":"visual_inspection",
             "variable":"Laser+Plasma","variableDetail":"VP/CD separate Laser=O Plasma=X","variableGroup":"test","intervention":"TEST 1",
             "inputQty":20,"okQty":20,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161016","testDate":"2025-10-23","line":"","checkType":"visual_inspection",
             "variable":"Laser+Plasma","variableDetail":"VP/CD separate Laser=O Plasma=O","variableGroup":"test","intervention":"TEST 2",
             "inputQty":24,"okQty":24,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161016","testDate":"2025-10-23","line":"","checkType":"visual_inspection",
             "variable":"Laser+Plasma","variableDetail":"VP/CD separate Laser=X Plasma=O","variableGroup":"normal","intervention":"NORMAL",
             "inputQty":20,"okQty":20,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        ],
        "tags":["brs-161016","brs-161014","vp-cd-separate","laser-cd","plasma","tension-test","doe-factorial","factor-screening"],
        "reportType":"doe_factorial",
        "verdict":"partial",
        "headline":"VP+CD tension best at Laser=O+Plasma=O avg 2.12 kgf; Laser=O+Plasma=X failed load test",
        "evidence":[
            {"metric":"VP+CD tension avg (best cell, 23/10)","baselineLabel":"Spec","baselineValue":"≥ 0.5 kgf",
             "variantLabel":"Laser=O+Plasma=O","variantValue":"2.12 kgf avg (1.79-2.40)",
             "deltaText":"+1.62 kgf","deltaSign":"up","note":"n=8",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Load test outcome (10pcs each, 21/10)","baselineLabel":"Spec","baselineValue":"Pass",
             "variantLabel":"Laser=X+Plasma=O","variantValue":"Fail (noise — SP+Coil gap)",
             "deltaText":"—","deltaSign":"no_change","note":"3/4 combos pass; 1 fail",
             "comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Adopt Laser=O + Plasma=O recipe — highest tension and load pass"},
            {"priority":2,"kind":"risk","text":"Avoid Laser=X + Plasma=O — load test noise NG (SP+Coil gap)"},
            {"priority":3,"kind":"investigate","text":"Verify Laser=O + Plasma=O on long-run pilot with 200+ pcs"},
        ],
        "context":{"process":"VP+CD separate root cause — Laser CD × Plasma DOE on Sub1 tension/load/function",
                   "stage":"Sub1 + main-line, multiple test dates 20-24 Oct 2025","baselineReason":"DOE — spec ≥0.5 kgf serves as baseline gate"},
        "doeGrid":{
            "factor1Name":"Laser CD","factor2Name":"Plasma",
            "factor1Levels":["X","O"],"factor2Levels":["X","O"],
            "cells":[
                {"f1":"X","f2":"X","status":"ok","value":"VP/CD 0% NG, tension avg 1.38 kgf, load pass"},
                {"f1":"O","f2":"O","status":"ok","value":"VP/CD 0% NG, tension avg 2.12 kgf, load pass"},
                {"f1":"X","f2":"O","status":"borderline","value":"Func NG 1.7%, tension avg 1.37 kgf, load FAIL noise"},
                {"f1":"O","f2":"X","status":"borderline","value":"Func NG 1.2%, tension avg 0.93 kgf, load pass"},
            ],
        },
        "trendPoints":None,
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
        "productType":"BRS-161016",
    },
    "tr_ko": {
        "headline":"VP+CD 인장 최적 조건 Laser=O+Plasma=O 평균 2.12 kgf; Laser=O+Plasma=X는 load test 실패",
        "actions":[
            {"priority":1,"kind":"action","text":"Laser=O + Plasma=O 레시피 채택 — 최고 인장 및 load 통과"},
            {"priority":2,"kind":"risk","text":"Laser=X + Plasma=O 회피 — load test 소음 NG (SP+Coil 갭)"},
            {"priority":3,"kind":"investigate","text":"Laser=O + Plasma=O를 200+ 샘플 장시간 양산 검증"},
        ],
        "context":{"process":"VP+CD 분리 원인 — Sub1 인장/load/기능에서 Laser CD × Plasma DOE",
                   "stage":"Sub1 + 메인라인, 2025-10-20~24 다일자","baselineReason":"DOE — Spec ≥0.5 kgf을 baseline 게이트로 사용"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
    "tr_vi": {
        "headline":"VP+CD lực căng tốt nhất Laser=O+Plasma=O TB 2.12 kgf; Laser=O+Plasma=X thất bại load test",
        "actions":[
            {"priority":1,"kind":"action","text":"Áp dụng công thức Laser=O + Plasma=O — lực căng cao nhất, đạt load test"},
            {"priority":2,"kind":"risk","text":"Tránh Laser=X + Plasma=O — load test NG noise (khe SP+Coil)"},
            {"priority":3,"kind":"investigate","text":"Xác minh Laser=O + Plasma=O với lô 200+ mẫu chạy dài"},
        ],
        "context":{"process":"Nguyên nhân VP+CD tách — DOE Laser CD × Plasma trên lực căng/load/chức năng Sub1",
                   "stage":"Sub1 + line chính, nhiều ngày 20-24/10/2025","baselineReason":"DOE — Spec ≥0.5 kgf làm cổng baseline"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
}

# ──────────────────────────────────────────────────────────────────────────
# 4. 17.BRS-2015 Report Test material Frame NG Bending PT 201506-S date 29.10.2024
# Frame NG bending vs Normal — function NG 5.8% test vs 3.2% normal → cannot use
RESULTS["17.BRS-2015 Report Test material Frame NG Bending PT 201506-S date 29.10.2024"] = {
    "result": {
        "measurements": [
            # 1. Frame vision
            {"productType":"BRS-201506","testDate":"2024-10-29","line":"E2-4B","checkType":"visual_inspection",
             "variable":"Frame material","variableDetail":"Frame vision","variableGroup":"test","intervention":"Frame Test (PT-201506-S bending)",
             "inputQty":100,"okQty":0,"ngTotal":100,"ngRate":100.0,"defectCategory":"other","defectType":"Big Bending","defectCount":46},
            {"productType":"BRS-201506","testDate":"2024-10-29","line":"E2-4B","checkType":"visual_inspection",
             "variable":"Frame material","variableDetail":"Frame vision","variableGroup":"test","intervention":"Frame Test (PT-201506-S bending)",
             "inputQty":100,"okQty":0,"ngTotal":100,"ngRate":100.0,"defectCategory":"other","defectType":"Small Bending","defectCount":54},
            # 2. Frame Array
            {"productType":"BRS-201506","testDate":"2024-10-29","line":"E2-4B","checkType":"process",
             "variable":"Frame material","variableDetail":"Frame Array","variableGroup":"test","intervention":"Frame small Bending",
             "inputQty":54,"okQty":54,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-201506","testDate":"2024-10-29","line":"E2-4B","checkType":"process",
             "variable":"Frame material","variableDetail":"Frame Array","variableGroup":"normal","intervention":"Frame normal",
             "inputQty":500,"okQty":500,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            # 3. FR+SUS Vision
            {"productType":"BRS-201506","testDate":"2024-10-29","line":"E2-4B","checkType":"visual_inspection",
             "variable":"Frame material","variableDetail":"FR+SUS Vision","variableGroup":"test","intervention":"Frame small Bending",
             "inputQty":54,"okQty":54,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-201506","testDate":"2024-10-29","line":"E2-4B","checkType":"visual_inspection",
             "variable":"Frame material","variableDetail":"FR+SUS Vision","variableGroup":"normal","intervention":"Frame normal",
             "inputQty":500,"okQty":500,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            # 4. Function
            {"productType":"BRS-201506","testDate":"2024-10-29","line":"","checkType":"function",
             "variable":"Frame material","variableDetail":"Function","variableGroup":"test","intervention":"Frame small Bending",
             "inputQty":52,"okQty":49,"ngTotal":3,"ngRate":5.8,"defectCategory":"function_hearing","defectType":"Noise","defectCount":3},
            {"productType":"BRS-201506","testDate":"2024-10-29","line":"","checkType":"function",
             "variable":"Frame material","variableDetail":"Function","variableGroup":"normal","intervention":"Frame normal",
             "inputQty":500,"okQty":484,"ngTotal":16,"ngRate":3.2,"defectCategory":"function_hearing","defectType":"Noise","defectCount":16},
            # MSM X516 Bottom module-line test
            {"productType":"MSM-X516BOTTOM","testDate":"2024-10-31","line":"X516 BOTTOM","checkType":"function",
             "variable":"Frame material","variableDetail":"Module-line function","variableGroup":"test","intervention":"Test SPK use Frame NG bending small",
             "inputQty":49,"okQty":47,"ngTotal":2,"ngRate":4.08,"defectCategory":"function_hearing","defectType":"Hearing","defectCount":2},
            {"productType":"MSM-X516BOTTOM","testDate":"2024-10-31","line":"X516 BOTTOM","checkType":"function",
             "variable":"Frame material","variableDetail":"Module-line function","variableGroup":"normal","intervention":"Normal",
             "inputQty":100,"okQty":99,"ngTotal":1,"ngRate":1.0,"defectCategory":"function_hearing","defectType":"Hearing","defectCount":1},
        ],
        "tags":["brs-201506","msm-x516","frame-material","ng-bending","pt-201506-s","function-test","module-line","comparison-study"],
        "reportType":"comparison_study",
        "verdict":"worsened",
        "headline":"NG-bending frame raises SPK function NG 5.8% vs 3.2% normal and module 4.08% vs 1.0%",
        "evidence":[
            {"metric":"SPK function NG","baselineLabel":"Frame normal","baselineValue":"3.2% (16/500)",
             "variantLabel":"Frame NG bending","variantValue":"5.8% (3/52)",
             "deltaText":"+2.6pp","deltaSign":"up","note":"Hearing-Noise 100% of NG",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Module-line function NG (X516)","baselineLabel":"Normal","baselineValue":"1.0% (1/100)",
             "variantLabel":"Frame NG bending","variantValue":"4.08% (2/49)",
             "deltaText":"+3.08pp","deltaSign":"up","note":"Hearing dominant",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Frame vision (incoming)","baselineLabel":"Standard 19.12-19.16 mm","baselineValue":"OK",
             "variantLabel":"PT-201506-S","variantValue":"19.08-19.10 mm 100/100",
             "deltaText":"—","deltaSign":"no_change","note":"Bending defect, dimension out of spec",
             "comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Reject PT-201506-S bending lot — do not consume in production"},
            {"priority":2,"kind":"investigate","text":"Trace supplier process why dimension came 19.08-19.10 mm (out of spec)"},
        ],
        "context":{"process":"Material qualification for bent Frame lot PT-201506-S — SPK + module-line function impact",
                   "stage":"E2-4B SPK line and X516 BOTTOM module line","baselineReason":"same-event Frame normal rows in both lines"},
        "doeGrid":None,"trendPoints":None,
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
        "productType":"BRS-201506",
    },
    "tr_ko": {
        "headline":"NG 벤딩 프레임 SPK 기능 NG 5.8% vs 3.2% 정상, 모듈 4.08% vs 1.0% 악화",
        "actions":[
            {"priority":1,"kind":"action","text":"PT-201506-S 벤딩 로트 거부 — 양산 투입 금지"},
            {"priority":2,"kind":"investigate","text":"공급사 공정에서 치수 19.08-19.10 mm(SPEC 이탈) 발생 원인 추적"},
        ],
        "context":{"process":"벤딩된 프레임 로트 PT-201506-S 자재 자격 — SPK + 모듈라인 기능 영향",
                   "stage":"E2-4B SPK 라인 및 X516 BOTTOM 모듈 라인","baselineReason":"양 라인에 동일 이벤트 Frame normal 행 존재"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
    "tr_vi": {
        "headline":"Frame NG bending nâng NG chức năng SPK 5.8% vs 3.2% normal và module 4.08% vs 1.0%",
        "actions":[
            {"priority":1,"kind":"action","text":"Từ chối lô PT-201506-S bending — không dùng trong sản xuất"},
            {"priority":2,"kind":"investigate","text":"Truy nguồn nhà cung cấp tại sao kích thước 19.08-19.10 mm (ngoài spec)"},
        ],
        "context":{"process":"Đủ điều kiện vật liệu lô Frame bending PT-201506-S — ảnh hưởng SPK + module-line chức năng",
                   "stage":"Line SPK E2-4B và line module X516 BOTTOM","baselineReason":"có dòng Frame normal cùng sự kiện ở cả hai line"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
}

# ──────────────────────────────────────────────────────────────────────────
# 5. 18 . BRS-161014 GMI  Report test YK new vender Doojin date 04.11.2024
# YK new vendor Doojin vs Normal — visual, decap, tension, function, drop all OK → CAN USE
RESULTS["18 . BRS-161014 GMI  Report test YK new vender Doojin date 04.11.2024 -"] = {
    "result": {
        "measurements": [
            # 1. Visual sub line
            {"productType":"BRS-161014","testDate":"2024-11-04","line":"E2-3B","checkType":"visual_inspection",
             "variable":"Yoke vendor","variableDetail":"Visual sub-line","variableGroup":"test","intervention":"YK test Doojin",
             "inputQty":192,"okQty":192,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161014","testDate":"2024-11-04","line":"E2-3B","checkType":"visual_inspection",
             "variable":"Yoke vendor","variableDetail":"Visual sub-line","variableGroup":"normal","intervention":"Normal YK",
             "inputQty":200,"okQty":200,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            # 5. Function
            {"productType":"BRS-161014","testDate":"2024-11-05","line":"E2-3B","checkType":"function",
             "variable":"Yoke vendor","variableDetail":"Function","variableGroup":"test","intervention":"Test YK Doojin",
             "inputQty":157,"okQty":152,"ngTotal":5,"ngRate":3.18,"defectCategory":"function_hearing","defectType":"Noise","defectCount":5},
            {"productType":"BRS-161014","testDate":"2024-11-05","line":"E2-3B","checkType":"function",
             "variable":"Yoke vendor","variableDetail":"Function","variableGroup":"normal","intervention":"Normal YK",
             "inputQty":160,"okQty":154,"ngTotal":6,"ngRate":3.75,"defectCategory":"function_hearing","defectType":"Noise","defectCount":5},
            {"productType":"BRS-161014","testDate":"2024-11-05","line":"E2-3B","checkType":"function",
             "variable":"Yoke vendor","variableDetail":"Function","variableGroup":"normal","intervention":"Normal YK",
             "inputQty":160,"okQty":154,"ngTotal":6,"ngRate":3.75,"defectCategory":"function_hearing","defectType":"Touch","defectCount":1},
            # Drop tests
            {"productType":"BRS-161014","testDate":"2024-11-07","line":"","checkType":"function",
             "variable":"Yoke vendor","variableDetail":"Drop test final","variableGroup":"test","intervention":"YK Doojin Auto+Manual",
             "inputQty":10,"okQty":10,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161014","testDate":"2024-11-07","line":"","checkType":"function",
             "variable":"Yoke vendor","variableDetail":"Drop test semi","variableGroup":"test","intervention":"YK Doojin",
             "inputQty":10,"okQty":10,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161014","testDate":"2024-11-07","line":"","checkType":"function",
             "variable":"Yoke vendor","variableDetail":"Drop test semi","variableGroup":"normal","intervention":"Normal YK",
             "inputQty":10,"okQty":10,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        ],
        "tags":["brs-161014","brs-161016","yoke-vendor","doojin","new-vendor","tension-test","decap-test","drop-test","comparison-study"],
        "reportType":"comparison_study",
        "verdict":"no_clear_effect",
        "headline":"Doojin YK matches normal — function NG 3.18% vs 3.75%, all decap/tension/drop OK",
        "evidence":[
            {"metric":"Function NG rate","baselineLabel":"Normal YK","baselineValue":"3.75% (6/160)",
             "variantLabel":"Test YK Doojin","variantValue":"3.18% (5/157)",
             "deltaText":"-0.57pp","deltaSign":"down","note":"within sample noise",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Tension MG-C (≥80 kgf spec)","baselineLabel":"Normal","baselineValue":"69.8-115.1 kgf",
             "variantLabel":"Test Doojin","variantValue":"81.1-121.8 kgf",
             "deltaText":"—","deltaSign":"up","note":"Doojin no failures (Normal had 69.8)",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Drop test (final + semi)","baselineLabel":"Normal","baselineValue":"0/10 NG",
             "variantLabel":"Test Doojin","variantValue":"0/10 NG both methods",
             "deltaText":"+0pp","deltaSign":"no_change","note":"Auto + Manual",
             "comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Approve Doojin YK new vendor for BRS-161014 production"},
            {"priority":2,"kind":"investigate","text":"Track Doojin lot-to-lot tension stability in next 3 lots"},
        ],
        "context":{"process":"Yoke (YK) new vendor Doojin qualification — visual, decap, tension, drop, function",
                   "stage":"E2-3B sub-line + module-line","baselineReason":"same-event Normal YK rows present"},
        "doeGrid":None,"trendPoints":None,
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
        "productType":"BRS-161014",
    },
    "tr_ko": {
        "headline":"Doojin YK 정상과 동등 — 기능 NG 3.18% vs 3.75%, 디캡/인장/낙하 모두 OK",
        "actions":[
            {"priority":1,"kind":"action","text":"BRS-161014 양산용 Doojin YK 신규 벤더 승인"},
            {"priority":2,"kind":"investigate","text":"이후 3개 로트에서 Doojin 로트간 인장 안정성 추적"},
        ],
        "context":{"process":"Yoke (YK) 신규 벤더 Doojin 자격 — 외관, 디캡, 인장, 낙하, 기능",
                   "stage":"E2-3B 서브라인 + 모듈라인","baselineReason":"동일 이벤트 Normal YK 행 존재"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
    "tr_vi": {
        "headline":"Doojin YK tương đương normal — chức năng NG 3.18% vs 3.75%, decap/tension/drop đều OK",
        "actions":[
            {"priority":1,"kind":"action","text":"Phê duyệt nhà cung cấp mới Doojin YK cho sản xuất BRS-161014"},
            {"priority":2,"kind":"investigate","text":"Theo dõi độ ổn định tension giữa các lô Doojin trong 3 lô tới"},
        ],
        "context":{"process":"Đủ điều kiện nhà cung cấp Yoke (YK) mới Doojin — visual, decap, tension, drop, chức năng",
                   "stage":"Sub-line E2-3B + module-line","baselineReason":"có dòng Normal YK cùng sự kiện"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
}

# ──────────────────────────────────────────────────────────────────────────
# 6. 18. BRS-161014  Report test new bond PT-8310M9S change bonding amount  17.2.2024
# Multi-arm: bonding amount levels 0.3-0.4 / 0.4-0.5 / 0.5-0.6 / Normal 0930 0.5-0.6
RESULTS["18. BRS-161014  Report test new bond PT-8310M9S change bonding amount  17.2.2024"] = {
    "result": {
        "measurements": [
            {"productType":"BRS-161014","testDate":"2024-02-17","line":"Sub 3","checkType":"visual_inspection",
             "variable":"Bond amount","variableDetail":"Vision Frame+SP","variableGroup":"test","intervention":"PT-8310M9S 0.4-0.5 mg",
             "inputQty":104,"okQty":104,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161014","testDate":"2024-02-17","line":"Sub 3","checkType":"visual_inspection",
             "variable":"Bond amount","variableDetail":"Vision Frame+SP","variableGroup":"test","intervention":"PT-8310M9S 0.3-0.4 mg",
             "inputQty":104,"okQty":104,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161014","testDate":"2024-02-17","line":"Sub 3","checkType":"visual_inspection",
             "variable":"Bond amount","variableDetail":"Vision Frame+SP","variableGroup":"test","intervention":"PT-8310M9S 0.5-0.6 mg",
             "inputQty":80,"okQty":71,"ngTotal":9,"ngRate":11.2,"defectCategory":"assembly_defect","defectType":"Over glue","defectCount":9},
            {"productType":"BRS-161014","testDate":"2024-02-17","line":"Sub 3","checkType":"visual_inspection",
             "variable":"Bond amount","variableDetail":"Vision Frame+SP","variableGroup":"normal","intervention":"Normal bond 0930 0.5-0.6 mg",
             "inputQty":200,"okQty":200,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        ],
        "tags":["brs-161014","bond-pt8310m9s","bonding-amount","over-glue","tension-test","multi-arm","intervention-test"],
        "reportType":"multi_arm",
        "verdict":"improved",
        "headline":"PT-8310M9S reduced bonding 0.3-0.5 mg eliminates over-glue (0% vs 11.2% at 0.5-0.6 mg)",
        "evidence":[
            {"metric":"Over-glue NG rate (Sub 3 vision)",
             "baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"",
             "deltaText":"+11.2pp range","deltaSign":"up","note":"",
             "comparisons":[
                {"label":"Normal bond 0930 (0.5-0.6 mg)","value":"0.0% (0/200)","n":200,"isBaseline":True,"isBest":True,"isWorst":False},
                {"label":"PT-8310M9S 0.3-0.4 mg","value":"0.0% (0/104)","n":104,"isBaseline":False,"isBest":False,"isWorst":False},
                {"label":"PT-8310M9S 0.4-0.5 mg","value":"0.0% (0/104)","n":104,"isBaseline":False,"isBest":False,"isWorst":False},
                {"label":"PT-8310M9S 0.5-0.6 mg","value":"11.2% (9/80)","n":80,"isBaseline":False,"isBest":False,"isWorst":True},
             ],
             "bestLabel":"Normal bond 0930 (0.5-0.6 mg)","worstLabel":"PT-8310M9S 0.5-0.6 mg"},
            {"metric":"Suspension C tension avg (kgf)","baselineLabel":"Normal 0930 (0.5-0.6 mg)","baselineValue":"0.31 kgf",
             "variantLabel":"PT-8310M9S 0.4-0.5 mg","variantValue":"0.40 kgf",
             "deltaText":"+0.09 kgf","deltaSign":"up","note":"all three PT-8310M9S levels meet target",
             "comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Set PT-8310M9S bonding amount to 0.4-0.5 mg out-of-spec (vs current 0.5-0.6 mg)"},
            {"priority":2,"kind":"risk","text":"Avoid PT-8310M9S at 0.5-0.6 mg — 11.2% over-glue NG"},
            {"priority":3,"kind":"investigate","text":"Update PT-8310M9S bonding-amount spec and run pilot lot"},
        ],
        "context":{"process":"New bond PT-8310M9S over-glue mitigation by reducing dispense amount",
                   "stage":"Sub 3 Frame+SP bonding station","baselineReason":"Normal old bond 0930 at SPEC dose serves as incumbent baseline"},
        "doeGrid":None,"trendPoints":None,
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
        "productType":"BRS-161014",
    },
    "tr_ko": {
        "headline":"PT-8310M9S 본드량 0.3-0.5 mg로 줄이면 over-glue 제거 (0% vs 0.5-0.6 mg의 11.2%)",
        "actions":[
            {"priority":1,"kind":"action","text":"PT-8310M9S 본드량을 SPEC 외 0.4-0.5 mg로 설정 (기존 0.5-0.6 mg 대신)"},
            {"priority":2,"kind":"risk","text":"PT-8310M9S 0.5-0.6 mg 회피 — over-glue NG 11.2%"},
            {"priority":3,"kind":"investigate","text":"PT-8310M9S 본드량 SPEC 개정 및 파일럿 로트 진행"},
        ],
        "context":{"process":"신규 본드 PT-8310M9S over-glue 저감 — 도포량 감소",
                   "stage":"Sub 3 Frame+SP 본딩 스테이션","baselineReason":"SPEC 도포량의 기존 본드 0930를 기존 기준으로 사용"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
    "tr_vi": {
        "headline":"PT-8310M9S giảm lượng keo 0.3-0.5 mg loại bỏ over-glue (0% vs 11.2% ở 0.5-0.6 mg)",
        "actions":[
            {"priority":1,"kind":"action","text":"Đặt lượng keo PT-8310M9S ở 0.4-0.5 mg ngoài spec (so với 0.5-0.6 mg hiện tại)"},
            {"priority":2,"kind":"risk","text":"Tránh PT-8310M9S ở 0.5-0.6 mg — NG over-glue 11.2%"},
            {"priority":3,"kind":"investigate","text":"Cập nhật spec lượng keo PT-8310M9S và chạy lô thí điểm"},
        ],
        "context":{"process":"Giảm over-glue keo mới PT-8310M9S bằng giảm lượng phun",
                   "stage":"Trạm bonding Frame+SP Sub 3","baselineReason":"keo cũ 0930 ở liều SPEC làm baseline hiện tại"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
}

# ──────────────────────────────────────────────────────────────────────────
# 7. 18. BRS-161014 DT Report test VP mold #7,9 add 0.05mm  date 16.2.2024
# Multi-arm: VP mold #7 (with bending sub-group), #9, vs Normal mold #8
RESULTS["18. BRS-161014 DT Report test VP mold #7,9 add 0.05mm  date 16.2.2024 -"] = {
    "result": {
        "measurements": [
            # VP bending vision Sub1
            {"productType":"BRS-161014DT","testDate":"2024-02-16","line":"C2-2A","checkType":"visual_inspection",
             "variable":"VP mold","variableDetail":"Laze cutting VP bending","variableGroup":"test","intervention":"Test VP mold #7 +0.05mm",
             "inputQty":2400,"okQty":2324,"ngTotal":76,"ngRate":3.2,"defectCategory":"other","defectType":"VP Bending G2","defectCount":70},
            {"productType":"BRS-161014DT","testDate":"2024-02-16","line":"C2-2A","checkType":"visual_inspection",
             "variable":"VP mold","variableDetail":"Laze cutting VP bending","variableGroup":"test","intervention":"Test VP mold #7 +0.05mm",
             "inputQty":2400,"okQty":2324,"ngTotal":76,"ngRate":3.2,"defectCategory":"other","defectType":"Cutting offset","defectCount":6},
            {"productType":"BRS-161014DT","testDate":"2024-02-16","line":"C2-2A","checkType":"visual_inspection",
             "variable":"VP mold","variableDetail":"Laze cutting VP bending","variableGroup":"test","intervention":"Test VP mold #9 +0.05mm",
             "inputQty":2400,"okQty":2398,"ngTotal":2,"ngRate":0.1,"defectCategory":"other","defectType":"Cutting offset","defectCount":2},
            {"productType":"BRS-161014DT","testDate":"2024-02-16","line":"C2-2A","checkType":"visual_inspection",
             "variable":"VP mold","variableDetail":"Laze cutting VP bending","variableGroup":"normal","intervention":"Normal VP mold #8",
             "inputQty":2400,"okQty":2394,"ngTotal":6,"ngRate":0.2,"defectCategory":"other","defectType":"VP Bending G2","defectCount":5},
            # Function
            {"productType":"BRS-161014DT","testDate":"2024-02-16","line":"C2-3A","checkType":"function",
             "variable":"VP mold","variableDetail":"Function","variableGroup":"test","intervention":"Test VP mold #7 OK",
             "inputQty":2321,"okQty":2034,"ngTotal":96,"ngRate":4.1,"defectCategory":"function_hearing","defectType":"Noise","defectCount":83},
            {"productType":"BRS-161014DT","testDate":"2024-02-16","line":"C2-3A","checkType":"function",
             "variable":"VP mold","variableDetail":"Function","variableGroup":"test","intervention":"Test VP mold #7 OK",
             "inputQty":2321,"okQty":2034,"ngTotal":96,"ngRate":4.1,"defectCategory":"function_hearing","defectType":"Touch","defectCount":10},
            {"productType":"BRS-161014DT","testDate":"2024-02-16","line":"C2-3A","checkType":"function",
             "variable":"VP mold","variableDetail":"Function","variableGroup":"test","intervention":"Test VP mold #7 bending",
             "inputQty":70,"okQty":46,"ngTotal":3,"ngRate":4.3,"defectCategory":"function_hearing","defectType":"Noise","defectCount":3},
            {"productType":"BRS-161014DT","testDate":"2024-02-16","line":"C2-3A","checkType":"function",
             "variable":"VP mold","variableDetail":"Function","variableGroup":"test","intervention":"Test VP mold #9 +0.05mm",
             "inputQty":2287,"okQty":1759,"ngTotal":116,"ngRate":5.1,"defectCategory":"function_hearing","defectType":"Noise","defectCount":112},
            {"productType":"BRS-161014DT","testDate":"2024-02-16","line":"C2-3A","checkType":"function",
             "variable":"VP mold","variableDetail":"Function","variableGroup":"normal","intervention":"Normal VP mold #8",
             "inputQty":2224,"okQty":2022,"ngTotal":64,"ngRate":2.9,"defectCategory":"function_hearing","defectType":"Noise","defectCount":43},
            {"productType":"BRS-161014DT","testDate":"2024-02-16","line":"C2-3A","checkType":"function",
             "variable":"VP mold","variableDetail":"Function","variableGroup":"normal","intervention":"Normal VP mold #8",
             "inputQty":2224,"okQty":2022,"ngTotal":64,"ngRate":2.9,"defectCategory":"function_hearing","defectType":"Touch","defectCount":18},
        ],
        "tags":["brs-161014","vp-mold","mold-modification","add-0.05mm","laze-cutting","function-test","multi-arm","vp-bending"],
        "reportType":"multi_arm",
        "verdict":"worsened",
        "headline":"VP mold #7+0.05mm 3.2% bending and mold #9 function 5.1% — both worse than normal #8",
        "evidence":[
            {"metric":"VP bending NG (laze cutting)",
             "baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"",
             "deltaText":"+3.1pp range","deltaSign":"up","note":"",
             "comparisons":[
                {"label":"Normal mold #8","value":"0.2% (6/2400)","n":2400,"isBaseline":True,"isBest":False,"isWorst":False},
                {"label":"Test mold #9 +0.05mm","value":"0.1% (2/2400)","n":2400,"isBaseline":False,"isBest":True,"isWorst":False},
                {"label":"Test mold #7 +0.05mm","value":"3.2% (76/2400)","n":2400,"isBaseline":False,"isBest":False,"isWorst":True},
             ],
             "bestLabel":"Test mold #9 +0.05mm","worstLabel":"Test mold #7 +0.05mm"},
            {"metric":"Function NG rate",
             "baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"",
             "deltaText":"+2.2pp range","deltaSign":"up","note":"",
             "comparisons":[
                {"label":"Normal mold #8","value":"2.9% (64/2224)","n":2224,"isBaseline":True,"isBest":True,"isWorst":False},
                {"label":"Total test mold #7","value":"4.1% (99/2391)","n":2391,"isBaseline":False,"isBest":False,"isWorst":False},
                {"label":"Test mold #9 +0.05mm","value":"5.1% (116/2287)","n":2287,"isBaseline":False,"isBest":False,"isWorst":True},
             ],
             "bestLabel":"Normal mold #8","worstLabel":"Test mold #9 +0.05mm"},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Reject +0.05mm modification on VP molds #7 and #9 — do not use"},
            {"priority":2,"kind":"investigate","text":"Ask supplier why mold #7 +0.05mm produces 3.2% VP bending"},
            {"priority":3,"kind":"investigate","text":"Understand G2 sigma rise (8-18%) on test molds vs normal"},
        ],
        "context":{"process":"VP mold center-part +0.05mm trial — bending and function impact",
                   "stage":"Sub1 laze cutting (C2-2A) and main-line function (C2-3A)","baselineReason":"Normal VP mold #8 present in same event"},
        "doeGrid":None,"trendPoints":None,
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
        "productType":"BRS-161014DT",
    },
    "tr_ko": {
        "headline":"VP 금형 #7+0.05mm 벤딩 3.2%, #9 기능 5.1% — 둘 다 정상 #8보다 악화",
        "actions":[
            {"priority":1,"kind":"action","text":"VP 금형 #7, #9 +0.05mm 개조 거부 — 사용 금지"},
            {"priority":2,"kind":"investigate","text":"#7 +0.05mm가 VP 벤딩 3.2%를 유발하는 원인 공급사에 확인"},
            {"priority":3,"kind":"investigate","text":"테스트 금형의 G2 sigma 상승(8-18%) 원인 분석"},
        ],
        "context":{"process":"VP 금형 중앙부 +0.05mm 시험 — 벤딩 및 기능 영향",
                   "stage":"Sub1 레이저 커팅(C2-2A) 및 메인라인 기능(C2-3A)","baselineReason":"동일 이벤트에 Normal VP 금형 #8 존재"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
    "tr_vi": {
        "headline":"Khuôn VP #7+0.05mm bending 3.2% và #9 chức năng 5.1% — đều xấu hơn normal #8",
        "actions":[
            {"priority":1,"kind":"action","text":"Từ chối sửa đổi +0.05mm trên khuôn VP #7 và #9 — không sử dụng"},
            {"priority":2,"kind":"investigate","text":"Hỏi nhà cung cấp vì sao khuôn #7 +0.05mm gây VP bending 3.2%"},
            {"priority":3,"kind":"investigate","text":"Phân tích nguyên nhân G2 sigma tăng (8-18%) trên khuôn test"},
        ],
        "context":{"process":"Thử nghiệm phần giữa khuôn VP +0.05mm — ảnh hưởng bending và chức năng",
                   "stage":"Laze cutting Sub1 (C2-2A) và chức năng line chính (C2-3A)","baselineReason":"có khuôn VP #8 Normal cùng sự kiện"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
}

# ──────────────────────────────────────────────────────────────────────────
# 8. 18. BRS-161014 Report check reason NG solder weak Date 7.12.2023
# Investigation of NG solder weak across MC AWFs, suspension lots, plasma, dry conditions
RESULTS["18. BRS-161014 Report check reason NG solder weak Date 7.12.2023"] = {
    "result": {
        "measurements": [
            {"productType":"BRS-161014DT","testDate":"2023-12-07","line":"AWF #1","checkType":"process","variable":"Solder weak","variableDetail":"Spot welding AWF","variableGroup":"test","intervention":"Lot laze cutting Nanosys","inputQty":100,"okQty":100,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161014DT","testDate":"2023-12-07","line":"AWF #2","checkType":"process","variable":"Solder weak","variableDetail":"Spot welding AWF","variableGroup":"test","intervention":"Lot laze cutting Nanosys","inputQty":100,"okQty":100,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161014DT","testDate":"2023-12-07","line":"AWF #3","checkType":"process","variable":"Solder weak","variableDetail":"Spot welding AWF","variableGroup":"test","intervention":"Lot laze cutting Nanosys","inputQty":36,"okQty":25,"ngTotal":11,"ngRate":30.6,"defectCategory":"other","defectType":"Solder weak","defectCount":11},
            {"productType":"BRS-161014DT","testDate":"2023-12-07","line":"AWF #3","checkType":"process","variable":"Solder weak","variableDetail":"Spot welding AWF after repair","variableGroup":"after","intervention":"AWF #3 after repair","inputQty":30,"okQty":30,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161014DT","testDate":"2023-12-07","line":"AWF #4","checkType":"process","variable":"Solder weak","variableDetail":"Spot welding AWF","variableGroup":"test","intervention":"Lot laze cutting Nanosys","inputQty":100,"okQty":99,"ngTotal":1,"ngRate":1.0,"defectCategory":"other","defectType":"Solder weak","defectCount":1},
            {"productType":"BRS-161014DT","testDate":"2023-12-07","line":"AWF #5","checkType":"process","variable":"Solder weak","variableDetail":"Spot welding AWF","variableGroup":"test","intervention":"Lot laze cutting Nanosys","inputQty":100,"okQty":100,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161014DT","testDate":"2023-12-07","line":"All M/C","checkType":"process","variable":"Solder weak","variableDetail":"Spot welding all M/C","variableGroup":"test","intervention":"Lot laze cutting Nanosys","inputQty":1500,"okQty":1496,"ngTotal":4,"ngRate":0.3,"defectCategory":"other","defectType":"Solder weak","defectCount":4},
            {"productType":"BRS-161014DT","testDate":"2023-12-07","line":"","checkType":"process","variable":"Suspension lot","variableDetail":"Spot welding","variableGroup":"test","intervention":"Suspension lot A","inputQty":192,"okQty":189,"ngTotal":3,"ngRate":1.6,"defectCategory":"other","defectType":"Solder weak","defectCount":3},
            {"productType":"BRS-161014DT","testDate":"2023-12-07","line":"","checkType":"process","variable":"Suspension lot","variableDetail":"Spot welding","variableGroup":"test","intervention":"Suspension lot A (rerun)","inputQty":240,"okQty":234,"ngTotal":6,"ngRate":2.5,"defectCategory":"other","defectType":"Solder weak","defectCount":6},
            {"productType":"BRS-161014DT","testDate":"2023-12-07","line":"","checkType":"process","variable":"Suspension lot","variableDetail":"Spot welding","variableGroup":"normal","intervention":"Suspension Normal","inputQty":108,"okQty":106,"ngTotal":2,"ngRate":1.9,"defectCategory":"other","defectType":"Solder weak","defectCount":2},
            {"productType":"BRS-161014DT","testDate":"2023-12-07","line":"","checkType":"process","variable":"Clean pad","variableDetail":"Spot welding clean pad alcohol","variableGroup":"test","intervention":"Clean pad alcohol","inputQty":60,"okQty":60,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161014DT","testDate":"2023-12-08","line":"","checkType":"process","variable":"Clean pad","variableDetail":"Spot welding clean pad alcohol","variableGroup":"test","intervention":"Clean pad alcohol","inputQty":90,"okQty":85,"ngTotal":5,"ngRate":5.6,"defectCategory":"other","defectType":"Solder weak","defectCount":5},
            {"productType":"BRS-161014DT","testDate":"2023-12-08","line":"","checkType":"process","variable":"Clean pad","variableDetail":"Spot welding not clean pad","variableGroup":"test","intervention":"Not clean pad alcohol","inputQty":96,"okQty":90,"ngTotal":6,"ngRate":6.2,"defectCategory":"other","defectType":"Solder weak","defectCount":6},
            {"productType":"BRS-161014DT","testDate":"2023-12-09","line":"","checkType":"process","variable":"Dry FR/SP","variableDetail":"Spot welding dry 15min 85C","variableGroup":"normal","intervention":"Dry FR/SP normal","inputQty":108,"okQty":107,"ngTotal":1,"ngRate":0.9,"defectCategory":"other","defectType":"Solder weak","defectCount":1},
            {"productType":"BRS-161014DT","testDate":"2023-12-09","line":"","checkType":"process","variable":"Dry FR/SP","variableDetail":"Clean pad+dry","variableGroup":"test","intervention":"Clean pad alcohol + Dry 15min 85C","inputQty":108,"okQty":101,"ngTotal":7,"ngRate":6.5,"defectCategory":"other","defectType":"Solder weak","defectCount":7},
            {"productType":"BRS-161014DT","testDate":"2023-12-09","line":"","checkType":"process","variable":"Plasma+vendor","variableDetail":"Spot welding plasma Nanosys","variableGroup":"test","intervention":"Plasma pad solder FR/SP Nanosys","inputQty":144,"okQty":139,"ngTotal":5,"ngRate":3.5,"defectCategory":"other","defectType":"Solder weak","defectCount":5},
            {"productType":"BRS-161014DT","testDate":"2023-12-09","line":"","checkType":"process","variable":"Dry+vendor","variableDetail":"Spot welding dry Nanosys","variableGroup":"test","intervention":"Dry FR/SP Nanosys 15min 85C","inputQty":360,"okQty":331,"ngTotal":29,"ngRate":8.1,"defectCategory":"other","defectType":"Solder weak","defectCount":29},
            {"productType":"BRS-161014DT","testDate":"2023-12-12","line":"","checkType":"process","variable":"Plasma+laze","variableDetail":"Spot welding plasma after lazer Nanosys","variableGroup":"test","intervention":"Plasma after cutting lazer FR/SP Nanosys","inputQty":2581,"okQty":2575,"ngTotal":6,"ngRate":0.2,"defectCategory":"other","defectType":"Solder weak","defectCount":6},
            {"productType":"BRS-161014DT","testDate":"2023-12-12","line":"","checkType":"process","variable":"Plasma+laze","variableDetail":"Spot welding FR/SP Nanosys","variableGroup":"normal","intervention":"FR/SP Nanosys Normal","inputQty":2500,"okQty":2495,"ngTotal":5,"ngRate":0.2,"defectCategory":"other","defectType":"Solder weak","defectCount":5},
        ],
        "tags":["brs-161014dt","solder-weak","spot-welding","awf-machine","suspension-lot","plasma","laze-cutting","nanosys","intervention-test","root-cause"],
        "reportType":"intervention_test",
        "verdict":"improved",
        "headline":"Solder weak root: AWF #3 (30.6% NG); plasma+laze cutting drops NG to 0.2% (~normal)",
        "evidence":[
            {"metric":"AWF #3 spot-welding NG (before/after repair)","baselineLabel":"After repair","baselineValue":"0.0% (0/30)",
             "variantLabel":"Before repair","variantValue":"30.6% (11/36)",
             "deltaText":"-30.6pp","deltaSign":"down","note":"machine fault confirmed",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Plasma after laze cutting NG","baselineLabel":"FR/SP Nanosys Normal","baselineValue":"0.2% (5/2500)",
             "variantLabel":"Plasma+laze Nanosys","variantValue":"0.2% (6/2581)",
             "deltaText":"+0.0pp","deltaSign":"no_change","note":"final countermeasure matches normal",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Suspension lot A vs Normal","baselineLabel":"Suspension Normal","baselineValue":"1.9% (2/108)",
             "variantLabel":"Suspension lot A (rerun)","variantValue":"2.5% (6/240)",
             "deltaText":"+0.6pp","deltaSign":"up","note":"smoke during welding observed",
             "comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Apply plasma after laze cutting + Nanosys lot to production (0.2% NG matches normal)"},
            {"priority":2,"kind":"action","text":"Switch from suspension improve lot (smoke) to Nanosys laze-cut suspension"},
            {"priority":3,"kind":"investigate","text":"Monitor AWF #3 after repair to ensure solder-weak stays at 0%"},
        ],
        "context":{"process":"NG solder-weak root-cause hunt — AWF M/C, suspension lots, clean pad, dry, plasma, laze-cut",
                   "stage":"Spot-welding line, multiple AWF machines","baselineReason":"FR/SP Nanosys Normal serves as benchmark; per-AWF before/after used pairwise"},
        "doeGrid":None,"trendPoints":None,
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
        "productType":"BRS-161014DT",
    },
    "tr_ko": {
        "headline":"솔더 약 원인: AWF #3 (NG 30.6%); 플라즈마+레이저 커팅으로 NG 0.2%(정상 수준)로 감소",
        "actions":[
            {"priority":1,"kind":"action","text":"플라즈마+레이저 커팅 + Nanosys 로트를 양산 적용 (NG 0.2%, 정상 동등)"},
            {"priority":2,"kind":"action","text":"smoke 발생하는 서스펜션 개선 로트에서 Nanosys 레이저 커팅 서스펜션으로 전환"},
            {"priority":3,"kind":"investigate","text":"수리 후 AWF #3 솔더-약 0% 유지 모니터링"},
        ],
        "context":{"process":"NG 솔더-약 원인 분석 — AWF M/C, 서스펜션 로트, 패드 세척, 건조, 플라즈마, 레이저 커팅",
                   "stage":"스폿 용접 라인, 다중 AWF 기계","baselineReason":"FR/SP Nanosys Normal을 벤치마크로 사용; AWF별 수리 전/후를 쌍대 비교"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
    "tr_vi": {
        "headline":"Nguyên nhân solder weak: AWF #3 (NG 30.6%); plasma+laze cutting hạ NG xuống 0.2% (gần normal)",
        "actions":[
            {"priority":1,"kind":"action","text":"Áp dụng plasma sau laze cutting + lô Nanosys vào sản xuất (NG 0.2% bằng normal)"},
            {"priority":2,"kind":"action","text":"Chuyển từ lô suspension improve (gây khói) sang suspension Nanosys laze-cut"},
            {"priority":3,"kind":"investigate","text":"Giám sát AWF #3 sau sửa để giữ solder-weak ở 0%"},
        ],
        "context":{"process":"Tìm nguyên nhân gốc NG solder-weak — máy AWF, lô suspension, làm sạch pad, sấy, plasma, laze-cut",
                   "stage":"Line spot welding, nhiều máy AWF","baselineReason":"FR/SP Nanosys Normal làm chuẩn; trước/sau sửa AWF dùng so cặp"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
}

# ──────────────────────────────────────────────────────────────────────────
# 9. 18. BRS-161014 Report test NEW GUIDE BASE AIR LEAK MACHINE 2023.09.07
# Two-arm trial: Guide cap test 1 vs test 2, two test dates
RESULTS["18. BRS-161014 Report test NEW GUIDE BASE AIR LEAK MACHINE 2023.09.07"] = {
    "result": {
        "measurements": [
            # 9/7 1st test
            {"productType":"BRS-161014","testDate":"2023-09-07","line":"","checkType":"visual_inspection",
             "variable":"Guide cap","variableDetail":"Long/Short VP visual 1st","variableGroup":"test","intervention":"Guide cap test 1",
             "inputQty":199,"okQty":198,"ngTotal":1,"ngRate":0.5,"defectCategory":"assembly_defect","defectType":"Long VP separate","defectCount":1},
            {"productType":"BRS-161014","testDate":"2023-09-07","line":"","checkType":"visual_inspection",
             "variable":"Guide cap","variableDetail":"Long/Short VP visual 1st","variableGroup":"test","intervention":"Guide cap test 2",
             "inputQty":200,"okQty":198,"ngTotal":2,"ngRate":1.0,"defectCategory":"assembly_defect","defectType":"Long VP separate","defectCount":1},
            {"productType":"BRS-161014","testDate":"2023-09-07","line":"","checkType":"visual_inspection",
             "variable":"Guide cap","variableDetail":"Long/Short VP visual 1st","variableGroup":"test","intervention":"Guide cap test 2",
             "inputQty":200,"okQty":198,"ngTotal":2,"ngRate":1.0,"defectCategory":"assembly_defect","defectType":"Long VP damage","defectCount":1},
            # 9/8 2nd test
            {"productType":"BRS-161014","testDate":"2023-09-08","line":"","checkType":"visual_inspection",
             "variable":"Guide cap","variableDetail":"Long/Short VP visual 2nd","variableGroup":"test","intervention":"Guide cap test 1",
             "inputQty":200,"okQty":198,"ngTotal":2,"ngRate":1.0,"defectCategory":"assembly_defect","defectType":"Long VP damage","defectCount":2},
            {"productType":"BRS-161014","testDate":"2023-09-08","line":"","checkType":"visual_inspection",
             "variable":"Guide cap","variableDetail":"Long/Short VP visual 2nd","variableGroup":"test","intervention":"Guide cap test 2",
             "inputQty":200,"okQty":200,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        ],
        "tags":["brs-161014","guide-cap","air-leak-machine","vp-separate","jig-improvement","comparison-study"],
        "reportType":"comparison_study",
        "verdict":"improved",
        "headline":"Guide cap type 2 → 0% NG in 2nd test vs type 1 (1.0% damage); adopt type 2",
        "evidence":[
            {"metric":"Long/Short VP NG (2nd test)","baselineLabel":"Guide cap test 1","baselineValue":"1.0% (2/200)",
             "variantLabel":"Guide cap test 2","variantValue":"0.0% (0/200)",
             "deltaText":"-1.0pp","deltaSign":"down","note":"Long VP damage from laser cutting",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Long/Short VP NG (1st test)","baselineLabel":"Guide cap test 1","baselineValue":"0.5% (1/199)",
             "variantLabel":"Guide cap test 2","variantValue":"1.0% (2/200)",
             "deltaText":"+0.5pp","deltaSign":"up","note":"PQC requested separate lot",
             "comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Adopt Guide cap jig type 2 on air-leak machine"},
            {"priority":2,"kind":"risk","text":"Monitor laser-cutting Long VP damage on first production lots"},
        ],
        "context":{"process":"Air-leak machine guide-cap jig improvement to reduce Long/Short VP separate/damage",
                   "stage":"VP visual inspection station, two trial dates 7-8 Sep 2023","baselineReason":"two guide-cap variants compared pairwise — type 2 chosen as best"},
        "doeGrid":None,"trendPoints":None,
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
        "productType":"BRS-161014",
    },
    "tr_ko": {
        "headline":"가이드 캡 type 2가 2차 시험 0% NG (type 1 1.0% 손상 대비); type 2 채택",
        "actions":[
            {"priority":1,"kind":"action","text":"에어 리크 머신에 가이드 캡 jig type 2 적용"},
            {"priority":2,"kind":"risk","text":"초기 양산 로트에서 레이저 커팅에 의한 Long VP 손상 모니터링"},
        ],
        "context":{"process":"Long/Short VP separate/damage 저감 위한 에어 리크 머신 가이드 캡 jig 개선",
                   "stage":"VP 외관 검사 스테이션, 2023-09-07~08 두 차례 시험","baselineReason":"두 가이드 캡 변종을 쌍대 비교 — type 2가 최우수"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
    "tr_vi": {
        "headline":"Nắp dẫn type 2 → 0% NG ở thử lần 2 vs type 1 (1.0% damage); chọn type 2",
        "actions":[
            {"priority":1,"kind":"action","text":"Áp dụng jig nắp dẫn type 2 trên máy air leak"},
            {"priority":2,"kind":"risk","text":"Theo dõi Long VP damage do laser cutting ở các lô sản xuất đầu"},
        ],
        "context":{"process":"Cải tiến jig nắp dẫn máy air leak nhằm giảm Long/Short VP separate/damage",
                   "stage":"Trạm kiểm tra VP, hai lần thử 07-08/09/2023","baselineReason":"so sánh cặp hai biến thể nắp dẫn — chọn type 2 tốt nhất"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
}

# ──────────────────────────────────────────────────────────────────────────
# 10. 18. BRS-161016 Report checking and test laser marking CD date 28.10.2025
# Multi-arm: Day Laser CD 04/11, 06/11; Night Laser CD 06/11 vs Normal each day
RESULTS["18. BRS-161016 Report checking and test laser marking CD date 28.10.2025"] = {
    "result": {
        "measurements": [
            # Laser machine base check
            {"productType":"BRS-161016","testDate":"2025-11-03","line":"E2-3B","checkType":"visual_inspection",
             "variable":"Laser base","variableDetail":"Base check 1st","variableGroup":"test","intervention":"Laser marking CD 1st",
             "inputQty":36,"okQty":35,"ngTotal":1,"ngRate":2.8,"defectCategory":"other","defectType":"Laser weak","defectCount":1},
            {"productType":"BRS-161016","testDate":"2025-11-03","line":"E2-3B","checkType":"visual_inspection",
             "variable":"Laser base","variableDetail":"Base check 2nd JIG reverse","variableGroup":"test","intervention":"Laser marking CD 2nd JIG reverse",
             "inputQty":36,"okQty":31,"ngTotal":5,"ngRate":13.9,"defectCategory":"other","defectType":"Laser offset","defectCount":5},
            # 04/11 Day shift VP+CD separate
            {"productType":"BRS-161016","testDate":"2025-11-04","line":"E2-3B","checkType":"visual_inspection",
             "variable":"Laser CD marking","variableDetail":"VP+CD separate Day shift","variableGroup":"test","intervention":"Laser CD",
             "inputQty":2200,"okQty":2183,"ngTotal":17,"ngRate":0.8,"defectCategory":"assembly_defect","defectType":"Separate VP/CD","defectCount":17},
            {"productType":"BRS-161016","testDate":"2025-11-04","line":"E2-3B","checkType":"visual_inspection",
             "variable":"Laser CD marking","variableDetail":"VP+CD separate Day shift","variableGroup":"normal","intervention":"Normal",
             "inputQty":1650,"okQty":1638,"ngTotal":12,"ngRate":0.7,"defectCategory":"assembly_defect","defectType":"Separate VP/CD","defectCount":12},
            # 06/11 Day E2-3B
            {"productType":"BRS-161016","testDate":"2025-11-06","line":"E2-3B","checkType":"visual_inspection",
             "variable":"Laser CD marking","variableDetail":"VP+CD separate Day shift","variableGroup":"test","intervention":"Laser CD",
             "inputQty":6450,"okQty":6409,"ngTotal":41,"ngRate":0.6,"defectCategory":"assembly_defect","defectType":"Separate VP/CD","defectCount":41},
            {"productType":"BRS-161016","testDate":"2025-11-06","line":"E2-3B","checkType":"visual_inspection",
             "variable":"Laser CD marking","variableDetail":"VP+CD separate Day shift","variableGroup":"normal","intervention":"Normal",
             "inputQty":2400,"okQty":2387,"ngTotal":13,"ngRate":0.5,"defectCategory":"assembly_defect","defectType":"Separate VP/CD","defectCount":13},
            # 06/11 Night E2-3A
            {"productType":"BRS-161016","testDate":"2025-11-06","line":"E2-3A","checkType":"visual_inspection",
             "variable":"Laser CD marking","variableDetail":"VP+CD separate Night shift","variableGroup":"test","intervention":"Laser CD",
             "inputQty":2400,"okQty":2378,"ngTotal":22,"ngRate":0.9,"defectCategory":"assembly_defect","defectType":"Separate VP/CD","defectCount":22},
            {"productType":"BRS-161016","testDate":"2025-11-06","line":"E2-3A","checkType":"visual_inspection",
             "variable":"Laser CD marking","variableDetail":"VP+CD separate Night shift","variableGroup":"normal","intervention":"Normal",
             "inputQty":1200,"okQty":1197,"ngTotal":3,"ngRate":0.2,"defectCategory":"assembly_defect","defectType":"Separate VP/CD","defectCount":3},
            # 05/11 Function
            {"productType":"BRS-161016","testDate":"2025-11-05","line":"","checkType":"function",
             "variable":"Laser CD marking","variableDetail":"Function","variableGroup":"test","intervention":"Laser CD",
             "inputQty":1110,"okQty":1077,"ngTotal":33,"ngRate":3.0,"defectCategory":"function_hearing","defectType":"Noise","defectCount":16},
            {"productType":"BRS-161016","testDate":"2025-11-05","line":"","checkType":"function",
             "variable":"Laser CD marking","variableDetail":"Function","variableGroup":"test","intervention":"Laser CD",
             "inputQty":1110,"okQty":1077,"ngTotal":33,"ngRate":3.0,"defectCategory":"function_hearing","defectType":"Touch","defectCount":15},
            {"productType":"BRS-161016","testDate":"2025-11-05","line":"","checkType":"function",
             "variable":"Laser CD marking","variableDetail":"Function","variableGroup":"normal","intervention":"Normal",
             "inputQty":2028,"okQty":1919,"ngTotal":109,"ngRate":5.4,"defectCategory":"function_hearing","defectType":"Noise","defectCount":50},
            {"productType":"BRS-161016","testDate":"2025-11-05","line":"","checkType":"function",
             "variable":"Laser CD marking","variableDetail":"Function","variableGroup":"normal","intervention":"Normal",
             "inputQty":2028,"okQty":1919,"ngTotal":109,"ngRate":5.4,"defectCategory":"function_hearing","defectType":"Touch","defectCount":51},
        ],
        "tags":["brs-161016","laser-marking-cd","vp-cd-separate","function-test","jig-improvement","intervention-test"],
        "reportType":"intervention_test",
        "verdict":"improved",
        "headline":"Laser marking CD reduces function NG 3.0% vs 5.4% normal (-2.4pp); VP/CD separate similar",
        "evidence":[
            {"metric":"Function NG rate (05/11)","baselineLabel":"Normal","baselineValue":"5.4% (109/2028)",
             "variantLabel":"Laser CD","variantValue":"3.0% (33/1110)",
             "deltaText":"-2.4pp","deltaSign":"down","note":"Hearing-Noise + Touch dominant",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"VP+CD separate NG (Day, 06/11)","baselineLabel":"Normal","baselineValue":"0.5% (13/2400)",
             "variantLabel":"Laser CD","variantValue":"0.6% (41/6450)",
             "deltaText":"+0.1pp","deltaSign":"up","note":"within noise",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"VP+CD separate NG (Night, 06/11)","baselineLabel":"Normal","baselineValue":"0.2% (3/1200)",
             "variantLabel":"Laser CD","variantValue":"0.9% (22/2400)",
             "deltaText":"+0.7pp","deltaSign":"up","note":"night shift gap larger",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Laser base check (jig orientation)","baselineLabel":"Laser 1st","baselineValue":"2.8% (1/36) laser weak",
             "variantLabel":"Laser 2nd JIG reverse","variantValue":"13.9% (5/36) laser offset",
             "deltaText":"+11.1pp","deltaSign":"up","note":"JIG reverse fails setting",
             "comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Adopt laser marking CD on production line — function NG drops 2.4pp"},
            {"priority":2,"kind":"risk","text":"Do not reverse jig on 2nd laser pass — 13.9% offset NG"},
            {"priority":3,"kind":"investigate","text":"Investigate night-shift VP+CD separate gap (0.9% vs 0.2%)"},
        ],
        "context":{"process":"Laser marking on CD to improve VP+CD separate; verify function and visual",
                   "stage":"E2-3B day and E2-3A night across 04-06 Nov 2025","baselineReason":"same-day Normal rows present on every shift"},
        "doeGrid":None,"trendPoints":None,
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
        "productType":"BRS-161016",
    },
    "tr_ko": {
        "headline":"Laser marking CD 적용 시 기능 NG 3.0% vs 정상 5.4% (-2.4pp); VP/CD 분리는 유사",
        "actions":[
            {"priority":1,"kind":"action","text":"양산 라인에 Laser marking CD 적용 — 기능 NG 2.4pp 감소"},
            {"priority":2,"kind":"risk","text":"2nd Laser 통과 시 jig reverse 금지 — offset NG 13.9%"},
            {"priority":3,"kind":"investigate","text":"야간 근무 VP+CD 분리 차이 조사 (0.9% vs 0.2%)"},
        ],
        "context":{"process":"VP+CD 분리 개선용 CD Laser marking; 기능 및 외관 검증",
                   "stage":"2025-11-04~06 E2-3B 주간 및 E2-3A 야간","baselineReason":"각 근무조에 동일 일자 Normal 행 존재"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
    "tr_vi": {
        "headline":"Laser marking CD giảm NG chức năng 3.0% vs 5.4% normal (-2.4pp); VP/CD separate tương đương",
        "actions":[
            {"priority":1,"kind":"action","text":"Áp dụng laser marking CD lên line sản xuất — NG chức năng giảm 2.4pp"},
            {"priority":2,"kind":"risk","text":"Không đảo jig ở lần laser thứ 2 — NG offset 13.9%"},
            {"priority":3,"kind":"investigate","text":"Điều tra khoảng cách VP+CD separate ca đêm (0.9% vs 0.2%)"},
        ],
        "context":{"process":"Laser marking trên CD nhằm cải thiện VP+CD separate; xác minh chức năng và visual",
                   "stage":"E2-3B ca ngày và E2-3A ca đêm trong 04-06/11/2025","baselineReason":"có dòng Normal cùng ngày ở mỗi ca"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
}

# ──────────────────────────────────────────────────────────────────────────
# 11. 18. BRS-2015 Report test Suspension 201506 -A and B  new vender WorldTop  Hanoi Vina
# New vendor suspension WorldTop vs Normal: all comparable → can use SPK line
RESULTS["18. BRS-2015 Report test Suspension 201506 -A and B  new vender WorldTop  Hanoi Vina"] = {
    "result": {
        "measurements": [
            # ass'y sub3
            {"productType":"BRS-201506","testDate":"2024-11-08","line":"E2-4B","checkType":"visual_inspection",
             "variable":"Suspension vendor","variableDetail":"Ass'y Suspension+Frame Sub 3","variableGroup":"test","intervention":"Lot test SP A and B (WorldTop)",
             "inputQty":4860,"okQty":4858,"ngTotal":2,"ngRate":0.04,"defectCategory":"assembly_defect","defectType":"NG separate","defectCount":2},
            {"productType":"BRS-201506","testDate":"2024-11-08","line":"E2-4B","checkType":"visual_inspection",
             "variable":"Suspension vendor","variableDetail":"Ass'y Suspension+Frame Sub 3","variableGroup":"normal","intervention":"Normal",
             "inputQty":3000,"okQty":2999,"ngTotal":1,"ngRate":0.03,"defectCategory":"assembly_defect","defectType":"NG separate","defectCount":1},
            # sport welding
            {"productType":"BRS-201506","testDate":"2024-11-09","line":"E2-4B","checkType":"process",
             "variable":"Suspension vendor","variableDetail":"Sport welding","variableGroup":"test","intervention":"Lot test SP A and B (WorldTop)",
             "inputQty":4860,"okQty":4857,"ngTotal":3,"ngRate":0.06,"defectCategory":"other","defectType":"Weak solder","defectCount":2},
            {"productType":"BRS-201506","testDate":"2024-11-09","line":"E2-4B","checkType":"process",
             "variable":"Suspension vendor","variableDetail":"Sport welding","variableGroup":"test","intervention":"Lot test SP A and B (WorldTop)",
             "inputQty":4860,"okQty":4857,"ngTotal":3,"ngRate":0.06,"defectCategory":"other","defectType":"Offset","defectCount":1},
            {"productType":"BRS-201506","testDate":"2024-11-09","line":"E2-4B","checkType":"process",
             "variable":"Suspension vendor","variableDetail":"Sport welding","variableGroup":"normal","intervention":"Normal",
             "inputQty":3000,"okQty":2998,"ngTotal":2,"ngRate":0.07,"defectCategory":"other","defectType":"Weak solder","defectCount":1},
            {"productType":"BRS-201506","testDate":"2024-11-09","line":"E2-4B","checkType":"process",
             "variable":"Suspension vendor","variableDetail":"Sport welding","variableGroup":"normal","intervention":"Normal",
             "inputQty":3000,"okQty":2998,"ngTotal":2,"ngRate":0.07,"defectCategory":"other","defectType":"Offset","defectCount":1},
            # Function
            {"productType":"BRS-201506","testDate":"2024-10-11","line":"E2-4B","checkType":"function",
             "variable":"Suspension vendor","variableDetail":"Function","variableGroup":"test","intervention":"Lot test SP A and B (WorldTop)",
             "inputQty":4850,"okQty":4725,"ngTotal":125,"ngRate":2.58,"defectCategory":"function_hearing","defectType":"Noise","defectCount":114},
            {"productType":"BRS-201506","testDate":"2024-10-11","line":"E2-4B","checkType":"function",
             "variable":"Suspension vendor","variableDetail":"Function","variableGroup":"normal","intervention":"Normal",
             "inputQty":1120,"okQty":1085,"ngTotal":35,"ngRate":3.12,"defectCategory":"function_hearing","defectType":"Noise","defectCount":16},
            {"productType":"BRS-201506","testDate":"2024-10-11","line":"E2-4B","checkType":"function",
             "variable":"Suspension vendor","variableDetail":"Function","variableGroup":"normal","intervention":"Normal",
             "inputQty":1120,"okQty":1085,"ngTotal":35,"ngRate":3.12,"defectCategory":"function_hearing","defectType":"Touch","defectCount":17},
        ],
        "tags":["brs-201506","suspension-vendor","worldtop","hanoi-vina","new-vendor","tension-test","function-test","comparison-study"],
        "reportType":"comparison_study",
        "verdict":"no_clear_effect",
        "headline":"WorldTop SP A/B function NG 2.58% vs 3.12% normal — comparable across all checks",
        "evidence":[
            {"metric":"SPK function NG","baselineLabel":"Normal","baselineValue":"3.12% (35/1120)",
             "variantLabel":"Lot SP A and B WorldTop","variantValue":"2.58% (125/4850)",
             "deltaText":"-0.54pp","deltaSign":"down","note":"large n WorldTop side",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Sub3 ass'y NG","baselineLabel":"Normal","baselineValue":"0.03% (1/3000)",
             "variantLabel":"Lot SP A and B WorldTop","variantValue":"0.04% (2/4860)",
             "deltaText":"+0.01pp","deltaSign":"up","note":"both <0.05%",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Spot welding NG","baselineLabel":"Normal","baselineValue":"0.07% (2/3000)",
             "variantLabel":"Lot SP A and B WorldTop","variantValue":"0.06% (3/4860)",
             "deltaText":"-0.01pp","deltaSign":"down","note":"",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Tension SP A (≥0.2 kgf spec)","baselineLabel":"Normal","baselineValue":"0.629 kgf avg",
             "variantLabel":"WorldTop","variantValue":"0.594 kgf avg",
             "deltaText":"-0.035 kgf","deltaSign":"down","note":"both above spec; sub-line tension OK",
             "comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Approve WorldTop Hanoi Vina suspension 201506-A/B for SPK line"},
            {"priority":2,"kind":"investigate","text":"Continue module-line qualification — module-line test pending"},
        ],
        "context":{"process":"New suspension vendor WorldTop Hanoi Vina qualification — SPK line ass'y, welding, tension, function",
                   "stage":"E2-4B SPK line, 08-11 Nov 2024","baselineReason":"same-period Normal rows present in every metric"},
        "doeGrid":None,"trendPoints":None,
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
        "productType":"BRS-201506",
    },
    "tr_ko": {
        "headline":"WorldTop SP A/B 기능 NG 2.58% vs 3.12% 정상 — 모든 항목에서 동등",
        "actions":[
            {"priority":1,"kind":"action","text":"SPK 라인에 WorldTop 하노이비나 서스펜션 201506-A/B 승인"},
            {"priority":2,"kind":"investigate","text":"모듈 라인 자격 시험 진행 — 결과 대기"},
        ],
        "context":{"process":"신규 서스펜션 벤더 WorldTop Hanoi Vina 자격 — SPK 라인 조립, 용접, 인장, 기능",
                   "stage":"E2-4B SPK 라인, 2024-11-08~11","baselineReason":"모든 항목에 동일 기간 Normal 행 존재"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
    "tr_vi": {
        "headline":"WorldTop SP A/B chức năng NG 2.58% vs 3.12% normal — tương đương ở mọi hạng mục",
        "actions":[
            {"priority":1,"kind":"action","text":"Phê duyệt suspension WorldTop Hanoi Vina 201506-A/B cho line SPK"},
            {"priority":2,"kind":"investigate","text":"Tiếp tục xét điều kiện cho line module — đang chờ kết quả"},
        ],
        "context":{"process":"Đủ điều kiện nhà cung cấp suspension mới WorldTop Hanoi Vina — line SPK lắp ráp, hàn, tension, chức năng",
                   "stage":"Line SPK E2-4B, 08-11/11/2024","baselineReason":"có dòng Normal cùng giai đoạn ở mọi chỉ số"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
}

# ──────────────────────────────────────────────────────────────────────────
# 12. 18. BRS-201506 Report checking Xray of NG Touch date 24.2.2024
# Xray analysis of NG Touch samples, plus AWF MC settings table — quality_log (pure inspection)
RESULTS["18. BRS-201506 Report checking Xray of NG Touch date 24.2.2024"] = {
    "result": {
        "measurements": [
            {"productType":"BRS-201506","testDate":"2024-02-24","line":"","checkType":"visual_inspection",
             "variable":"NG Touch","variableDetail":"Xray inspection of NG Touch samples","variableGroup":"test","intervention":"Xray 10pcs",
             "inputQty":10,"okQty":0,"ngTotal":10,"ngRate":100.0,"defectCategory":"function_hearing","defectType":"Touch","defectCount":10},
        ],
        "tags":["brs-201506","xray-inspection","ng-touch","hearing-touch","awf-machine","quality-log"],
        "reportType":"quality_log",
        "verdict":"",
        "headline":"Xray of 10 NG-Touch units logged; AWF #1/#3 winding jig/pole differ from #2/#4/#5",
        "evidence":[
            {"metric":"Samples Xray-checked","baselineLabel":"","baselineValue":"",
             "variantLabel":"NG Touch","variantValue":"10 pcs",
             "deltaText":"—","deltaSign":"no_change","note":"Pictures NO1-NO10 captured",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"AWF winding jig setting","baselineLabel":"AWF #2/#4/#5","baselineValue":"9.42 mm",
             "variantLabel":"AWF #1","variantValue":"9.34 mm",
             "deltaText":"-0.08 mm","deltaSign":"down","note":"AWF #1 jig smaller",
             "comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"AWF stretching pole setting","baselineLabel":"AWF #1/#2/#4/#5","baselineValue":"5.065 mm",
             "variantLabel":"AWF #3","variantValue":"5.08 mm",
             "deltaText":"+0.015 mm","deltaSign":"up","note":"AWF #3 pole longer",
             "comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"investigate","text":"Correlate Xray defect type with originating AWF #1 / #3 unique settings"},
            {"priority":2,"kind":"action","text":"Run AWF #5 (not yet operated) with same standard 9.42 / 5.065"},
        ],
        "context":{"process":"Xray analysis of NG Touch samples to find AWF/coil root cause",
                   "stage":"Off-line Xray lab + AWF machine setting comparison","baselineReason":"quality_log — no comparison gate, AWF settings logged as reference"},
        "doeGrid":None,"trendPoints":None,
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
        "productType":"BRS-201506",
    },
    "tr_ko": {
        "headline":"NG-Touch 10개 Xray 기록; AWF #1/#3 권선 jig/pole 설정이 #2/#4/#5와 상이",
        "actions":[
            {"priority":1,"kind":"investigate","text":"Xray 결함 유형과 AWF #1/#3 고유 설정과의 상관성 분석"},
            {"priority":2,"kind":"action","text":"미가동 AWF #5를 표준 9.42 / 5.065로 가동"},
        ],
        "context":{"process":"AWF/coil 원인 분석을 위한 NG Touch 샘플 Xray 분석",
                   "stage":"오프라인 Xray 실험실 + AWF 머신 설정 비교","baselineReason":"quality_log — 비교 게이트 없음, AWF 설정은 참조로 기록"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
    "tr_vi": {
        "headline":"Xray 10 mẫu NG-Touch ghi nhận; cài đặt AWF #1/#3 (jig/pole) khác #2/#4/#5",
        "actions":[
            {"priority":1,"kind":"investigate","text":"Tương quan loại lỗi Xray với cài đặt riêng của AWF #1/#3"},
            {"priority":2,"kind":"action","text":"Chạy AWF #5 (chưa hoạt động) với chuẩn 9.42 / 5.065"},
        ],
        "context":{"process":"Phân tích Xray các mẫu NG Touch để tìm nguyên nhân AWF/coil",
                   "stage":"Phòng Xray offline + so sánh cài đặt máy AWF","baselineReason":"quality_log — không có cổng so sánh, cài đặt AWF ghi nhận tham khảo"},
        "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":"",
    },
}


# ──────────────────────────────────────────────────────────────────────────
# COMMIT
# ──────────────────────────────────────────────────────────────────────────

def commit_all():
    con = sqlite3.connect(DB)
    con.execute("PRAGMA busy_timeout = 30000")
    cur = con.cursor()

    with open(CHUNK, "r", encoding="utf-8") as f:
        targets = [l.rstrip("\n") for l in f if l.strip()]

    ok = 0; skipped = 0; pfail = 0
    for name in targets:
        row = cur.execute(
            "SELECT ExtractedText FROM RawReportText WHERE DatasetName=? AND Kind='excel_paste'",
            (name,)).fetchone()
        if not row or not row[0]:
            print(f"[SKIP {name}] no excel_paste")
            skipped += 1
            continue
        bundle = RESULTS.get(name)
        if not bundle:
            print(f"[PARSE-FAIL {name}] no result built")
            pfail += 1
            continue
        try:
            result = bundle["result"]; tr_ko = bundle["tr_ko"]; tr_vi = bundle["tr_vi"]
            cur.execute("BEGIN")
            now = datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")
            product = result.get("productType", "")

            cur.execute("DELETE FROM NormalizedMeasurements WHERE DatasetName=?", (name,))
            for m in result.get("measurements", []):
                cur.execute("""
                    INSERT INTO NormalizedMeasurements
                      (DatasetName, ProductType, TestDate, Line, CheckType, Variable,
                       VariableDetail, VariableGroup, Intervention, InputQty, OkQty,
                       NgTotal, NgRate, DefectCategory, DefectType, DefectCount, CreatedAt)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (name, m.get("productType") or product, m.get("testDate",""),
                     m.get("line",""), m.get("checkType",""), m.get("variable",""),
                     m.get("variableDetail",""), m.get("variableGroup",""), m.get("intervention",""),
                     int(m.get("inputQty",0)), int(m.get("okQty",0)), int(m.get("ngTotal",0)),
                     float(m.get("ngRate",0)), m.get("defectCategory",""), m.get("defectType",""),
                     int(m.get("defectCount",0)), now))

            tags_json     = json.dumps(result.get("tags") or [], ensure_ascii=False)
            evidence_json = json.dumps(result.get("evidence") or [], ensure_ascii=False)
            actions_json  = json.dumps(result.get("actions") or [], ensure_ascii=False)
            context_json  = json.dumps(result.get("context"), ensure_ascii=False) if result.get("context") else ""
            doe_json      = json.dumps(result.get("doeGrid"), ensure_ascii=False) if result.get("doeGrid") else ""
            trend_json    = json.dumps(result.get("trendPoints"), ensure_ascii=False) if result.get("trendPoints") else ""

            cur.execute("""
                INSERT INTO DatasetSummary
                  (DatasetName, ProductType, Summary, KeyFindings, Tags, CreatedAt,
                   Purpose, TestConditions, RootCause, Decision, RecommendedAction,
                   Verdict, Headline, EvidenceJson, ActionsJson, ContextJson,
                   ReportType, DoeGridJson, TrendJson)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT(DatasetName) DO UPDATE SET
                  ProductType=excluded.ProductType, Summary=excluded.Summary,
                  KeyFindings=excluded.KeyFindings, Tags=excluded.Tags,
                  CreatedAt=excluded.CreatedAt,
                  Purpose=excluded.Purpose, TestConditions=excluded.TestConditions,
                  RootCause=excluded.RootCause, Decision=excluded.Decision,
                  RecommendedAction=excluded.RecommendedAction,
                  Verdict=excluded.Verdict, Headline=excluded.Headline,
                  EvidenceJson=excluded.EvidenceJson,
                  ActionsJson=excluded.ActionsJson,
                  ContextJson=excluded.ContextJson,
                  ReportType=excluded.ReportType,
                  DoeGridJson=excluded.DoeGridJson,
                  TrendJson=excluded.TrendJson""",
                (name, product,
                 result.get("summary",""), result.get("keyFindings",""),
                 tags_json, now,
                 result.get("purpose",""), result.get("testConditions",""),
                 result.get("rootCause",""), result.get("decision",""),
                 result.get("recommendedAction",""),
                 result.get("verdict",""), result.get("headline",""),
                 evidence_json, actions_json, context_json,
                 result.get("reportType",""), doe_json, trend_json))

            for lang, tr in [("ko", tr_ko), ("vi", tr_vi)]:
                if tr is None: continue
                tr_actions_json = json.dumps(tr.get("actions") or [], ensure_ascii=False)
                tr_context_json = json.dumps(tr.get("context"), ensure_ascii=False) if tr.get("context") else ""
                cur.execute("""
                    INSERT INTO DatasetSummaryTranslations
                      (DatasetName, Lang, Summary, KeyFindings, Purpose, TestConditions,
                       RootCause, Decision, RecommendedAction,
                       Headline, ActionsJson, ContextJson, UpdatedAt)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
                    ON CONFLICT(DatasetName, Lang) DO UPDATE SET
                      Summary=excluded.Summary, KeyFindings=excluded.KeyFindings,
                      Purpose=excluded.Purpose, TestConditions=excluded.TestConditions,
                      RootCause=excluded.RootCause, Decision=excluded.Decision,
                      RecommendedAction=excluded.RecommendedAction,
                      Headline=excluded.Headline,
                      ActionsJson=excluded.ActionsJson,
                      ContextJson=excluded.ContextJson,
                      UpdatedAt=excluded.UpdatedAt""",
                    (name, lang,
                     tr.get("summary",""), tr.get("keyFindings",""), tr.get("purpose",""),
                     tr.get("testConditions",""), tr.get("rootCause",""),
                     tr.get("decision",""), tr.get("recommendedAction",""),
                     tr.get("headline",""), tr_actions_json, tr_context_json, now))

            con.commit()
            print(f"[OK {name}]")
            ok += 1
        except Exception as e:
            con.rollback()
            print(f"[PARSE-FAIL {name}] {e}")
            pfail += 1

    con.close()
    print(f"\n=== BATCH DONE === Mode: Reanalyze  OK={ok}  SKIP={skipped}  PARSE-FAIL={pfail}")


if __name__ == "__main__":
    commit_all()

