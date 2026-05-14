# -*- coding: utf-8 -*-
"""CLI AI Batch — chunk_04 commit script.
Agent is the LLM. Normalization JSON + ko/vi translations are embedded
verbatim below. Commits one transaction per dataset.
"""
import sqlite3, json, sys, io, os
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DB = r'D:\000. MyWorks\002. DB\process-review.db'

# ---------------------------------------------------------------------------
# Per-dataset payloads. Each entry has:
#   name, product, result (NormalizeResult), tr_ko, tr_vi
# ---------------------------------------------------------------------------

# Helper for legacy-7 empty strings
LEGACY = {
    "summary":"", "keyFindings":"", "purpose":"", "testConditions":"",
    "rootCause":"", "decision":"", "recommendedAction":""
}

DATASETS = []

# ============================================================================
# 0. 11. L20S15-07GMI Report test CD IR25050828 (130k) and lot replace (140k) Ng dyne pen Tét again
# ============================================================================
DATASETS.append(dict(
    name="11. L20S15-07GMI  Report test CD IR25050828 ( Size 130k ) and lot replace ( Size 140K )Ng dyne pen Tét again",
    product="L20S15-07",
    result={
        "measurements":[
            # Test CD IR25050828 lots 1-6 + Lot replace 140K lots 1-3, plus Normal
            {"productType":"L20S15-07","testDate":"2025-05-23","line":"E2-4A","checkType":"function","variable":"CD lot","variableDetail":"Lot CD IR25050828 (Size 300k) lot 1","variableGroup":"test","intervention":"","inputQty":45,"okQty":41,"ngTotal":4,"ngRate":8.9,"defectCategory":"assembly_defect","defectType":"VP/CD separate","defectCount":3},
            {"productType":"L20S15-07","testDate":"2025-05-23","line":"E2-4A","checkType":"function","variable":"CD lot","variableDetail":"Lot CD IR25050828 (Size 300k) lot 1","variableGroup":"test","intervention":"","inputQty":45,"okQty":41,"ngTotal":4,"ngRate":8.9,"defectCategory":"function_hearing","defectType":"NG Hearing","defectCount":1},
            {"productType":"L20S15-07","testDate":"2025-05-23","line":"E2-4A","checkType":"function","variable":"CD lot","variableDetail":"Lot CD IR25050828 (Size 300k) lot 2","variableGroup":"test","intervention":"","inputQty":45,"okQty":40,"ngTotal":5,"ngRate":11.1,"defectCategory":"assembly_defect","defectType":"VP/CD separate","defectCount":4},
            {"productType":"L20S15-07","testDate":"2025-05-23","line":"E2-4A","checkType":"function","variable":"CD lot","variableDetail":"Lot CD IR25050828 (Size 300k) lot 2","variableGroup":"test","intervention":"","inputQty":45,"okQty":40,"ngTotal":5,"ngRate":11.1,"defectCategory":"function_hearing","defectType":"NG Hearing","defectCount":1},
            {"productType":"L20S15-07","testDate":"2025-05-23","line":"E2-4A","checkType":"function","variable":"CD lot","variableDetail":"Lot CD IR25050828 (Size 300k) lot 3","variableGroup":"test","intervention":"","inputQty":45,"okQty":38,"ngTotal":7,"ngRate":15.6,"defectCategory":"assembly_defect","defectType":"VP/CD separate","defectCount":5},
            {"productType":"L20S15-07","testDate":"2025-05-23","line":"E2-4A","checkType":"function","variable":"CD lot","variableDetail":"Lot CD IR25050828 (Size 300k) lot 3","variableGroup":"test","intervention":"","inputQty":45,"okQty":38,"ngTotal":7,"ngRate":15.6,"defectCategory":"function_hearing","defectType":"NG Hearing","defectCount":2},
            {"productType":"L20S15-07","testDate":"2025-05-23","line":"E2-4A","checkType":"function","variable":"CD lot","variableDetail":"Lot CD IR25050828 (Size 300k) lot 4","variableGroup":"test","intervention":"","inputQty":46,"okQty":43,"ngTotal":3,"ngRate":6.5,"defectCategory":"assembly_defect","defectType":"VP/CD separate","defectCount":1},
            {"productType":"L20S15-07","testDate":"2025-05-23","line":"E2-4A","checkType":"function","variable":"CD lot","variableDetail":"Lot CD IR25050828 (Size 300k) lot 4","variableGroup":"test","intervention":"","inputQty":46,"okQty":43,"ngTotal":3,"ngRate":6.5,"defectCategory":"function_hearing","defectType":"NG Hearing","defectCount":2},
            {"productType":"L20S15-07","testDate":"2025-05-23","line":"E2-4A","checkType":"function","variable":"CD lot","variableDetail":"Lot CD IR25050828 (Size 300k) lot 5","variableGroup":"test","intervention":"","inputQty":45,"okQty":44,"ngTotal":1,"ngRate":2.2,"defectCategory":"assembly_defect","defectType":"VP/CD separate","defectCount":1},
            {"productType":"L20S15-07","testDate":"2025-05-23","line":"E2-4A","checkType":"function","variable":"CD lot","variableDetail":"Lot CD IR25050828 (Size 300k) lot 6","variableGroup":"test","intervention":"","inputQty":45,"okQty":36,"ngTotal":9,"ngRate":20.0,"defectCategory":"assembly_defect","defectType":"VP/CD separate","defectCount":6},
            {"productType":"L20S15-07","testDate":"2025-05-23","line":"E2-4A","checkType":"function","variable":"CD lot","variableDetail":"Lot CD IR25050828 (Size 300k) lot 6","variableGroup":"test","intervention":"","inputQty":45,"okQty":36,"ngTotal":9,"ngRate":20.0,"defectCategory":"function_hearing","defectType":"NG Hearing","defectCount":3},
            {"productType":"L20S15-07","testDate":"2025-05-23","line":"E2-4A","checkType":"function","variable":"CD lot","variableDetail":"Lot replace (Size 140K) lot 1","variableGroup":"new_lot","intervention":"","inputQty":47,"okQty":35,"ngTotal":12,"ngRate":25.5,"defectCategory":"assembly_defect","defectType":"VP/CD separate","defectCount":11},
            {"productType":"L20S15-07","testDate":"2025-05-23","line":"E2-4A","checkType":"function","variable":"CD lot","variableDetail":"Lot replace (Size 140K) lot 1","variableGroup":"new_lot","intervention":"","inputQty":47,"okQty":35,"ngTotal":12,"ngRate":25.5,"defectCategory":"function_hearing","defectType":"NG Hearing","defectCount":1},
            {"productType":"L20S15-07","testDate":"2025-05-23","line":"E2-4A","checkType":"function","variable":"CD lot","variableDetail":"Lot replace (Size 140K) lot 2","variableGroup":"new_lot","intervention":"","inputQty":46,"okQty":40,"ngTotal":6,"ngRate":13.0,"defectCategory":"assembly_defect","defectType":"VP/CD separate","defectCount":6},
            {"productType":"L20S15-07","testDate":"2025-05-23","line":"E2-4A","checkType":"function","variable":"CD lot","variableDetail":"Lot replace (Size 140K) lot 3","variableGroup":"new_lot","intervention":"","inputQty":45,"okQty":37,"ngTotal":8,"ngRate":17.8,"defectCategory":"assembly_defect","defectType":"VP/CD separate","defectCount":7},
            {"productType":"L20S15-07","testDate":"2025-05-23","line":"E2-4A","checkType":"function","variable":"CD lot","variableDetail":"Lot replace (Size 140K) lot 3","variableGroup":"new_lot","intervention":"","inputQty":45,"okQty":37,"ngTotal":8,"ngRate":17.8,"defectCategory":"function_hearing","defectType":"NG Hearing","defectCount":1},
            {"productType":"L20S15-07","testDate":"2025-05-23","line":"E2-4A","checkType":"function","variable":"CD lot","variableDetail":"Normal","variableGroup":"normal","intervention":"","inputQty":50,"okQty":50,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        ],
        "tags":["l20s15-07","cd-lot-test","ir25050828","vp-cd-separate","dyne-pen","tension-test","function-ng","lot-replace"],
        "reportType":"multi_arm",
        "verdict":"worsened",
        "headline":"Test CD lots show VP/CD separate NG 2.2-25.5% vs Normal 0% — cannot use",
        "evidence":[
            {"metric":"NG rate (avg across test lots)","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"+25.5pp range","deltaSign":"up","note":"all test lots NG > 0",
             "comparisons":[
                {"label":"Normal","value":"0.0% (0/50)","n":50,"isBaseline":True,"isBest":True,"isWorst":False},
                {"label":"CD IR25050828 lot 5","value":"2.2% (1/45)","n":45,"isBaseline":False,"isBest":False,"isWorst":False},
                {"label":"CD IR25050828 lot 4","value":"6.5% (3/46)","n":46,"isBaseline":False,"isBest":False,"isWorst":False},
                {"label":"CD IR25050828 lot 1","value":"8.9% (4/45)","n":45,"isBaseline":False,"isBest":False,"isWorst":False},
                {"label":"CD IR25050828 lot 2","value":"11.1% (5/45)","n":45,"isBaseline":False,"isBest":False,"isWorst":False},
                {"label":"Lot replace 140K lot 2","value":"13.0% (6/46)","n":46,"isBaseline":False,"isBest":False,"isWorst":False},
                {"label":"CD IR25050828 lot 6","value":"20.0% (9/45)","n":45,"isBaseline":False,"isBest":False,"isWorst":False},
                {"label":"Lot replace 140K lot 1","value":"25.5% (12/47)","n":47,"isBaseline":False,"isBest":False,"isWorst":True},
             ],
             "bestLabel":"Normal","worstLabel":"Lot replace 140K lot 1"},
            {"metric":"Tension VP+CD (Spec 1.2kgf)","baselineLabel":"Normal","baselineValue":"2.548 kgf avg","variantLabel":"Test lots","variantValue":"1.579–2.931 kgf range","deltaText":"—","deltaSign":"no_change","note":"all Pass spec","comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Do not use CD IR25050828 (130k) and replacement lot (140K) for production"},
            {"priority":2,"kind":"investigate","text":"Investigate root cause of VP/CD separate on test CD lots"},
        ],
        "context":{
            "process":"VP+CD assembly with laser CD and plasma","stage":"E2-4A line — dyne pen check + tension test + separate check + function",
            "baselineReason":"same-event Normal row present (10/Normal lot)"
        },
        "doeGrid":None,"trendPoints":None,
        **LEGACY
    },
    tr_ko={**LEGACY,
        "headline":"테스트 CD lot VP/CD 분리 NG 2.2-25.5% (Normal 0%) — 사용 불가",
        "actions":[
            {"priority":1,"kind":"action","text":"CD IR25050828 (130k) 및 교체 lot (140K) 양산 사용 불가"},
            {"priority":2,"kind":"investigate","text":"테스트 CD lot의 VP/CD 분리 근본원인 조사"},
        ],
        "context":{"process":"레이저 CD + 플라즈마 VP+CD 조립","stage":"E2-4A 라인 — dyne pen 점검 + 인장 시험 + 분리 점검 + function","baselineReason":"동일 이벤트에 Normal 행 존재 (10/Normal lot)"}
    },
    tr_vi={**LEGACY,
        "headline":"Lot CD test NG VP/CD separate 2.2-25.5% (Normal 0%) — không sử dụng được",
        "actions":[
            {"priority":1,"kind":"action","text":"Không sử dụng lot CD IR25050828 (130k) và lot thay thế (140K) cho sản xuất"},
            {"priority":2,"kind":"investigate","text":"Điều tra nguyên nhân gốc VP/CD separate trên các lot CD test"},
        ],
        "context":{"process":"Lắp ráp VP+CD với laser CD và plasma","stage":"Line E2-4A — kiểm tra dyne pen + test lực căng + kiểm tra separate + function","baselineReason":"có dòng Normal cùng sự kiện (lot 10/Normal)"}
    },
))

# ============================================================================
# 1. 11. MSU-L20S15-07 Report Check again Lot SPK Block 50K Date 15.5.2025
# ============================================================================
DATASETS.append(dict(
    name="11. MSU-L20S15-07 Report Check again Lot SPK  Block 50K    Date 15.5.2025",
    product="MSU-L20S15-07",
    result={
        "measurements":[
            # Decap analysis: VP/CD separate 15/15 = 100%
            {"productType":"MSU-L20S15-07","testDate":"2025-05-15","line":"","checkType":"function","variable":"NG analysis","variableDetail":"Decap NG analysis","variableGroup":"test","intervention":"","inputQty":15,"okQty":0,"ngTotal":15,"ngRate":100.0,"defectCategory":"assembly_defect","defectType":"VP/CD separate","defectCount":15},
            # Recheck function lot block 50K
            {"productType":"MSU-L20S15-07","testDate":"2025-05-15","line":"","checkType":"function","variable":"Lot block recheck","variableDetail":"Lot Block (happen VP/CD)","variableGroup":"test","intervention":"Drop test + Aging + Recheck","inputQty":500,"okQty":333,"ngTotal":167,"ngRate":33.4,"defectCategory":"function_hearing","defectType":"Noise","defectCount":156},
            {"productType":"MSU-L20S15-07","testDate":"2025-05-15","line":"","checkType":"function","variable":"Lot block recheck","variableDetail":"Lot Block (happen VP/CD)","variableGroup":"test","intervention":"Drop test + Aging + Recheck","inputQty":500,"okQty":333,"ngTotal":167,"ngRate":33.4,"defectCategory":"function_thd","defectType":"THD","defectCount":8},
            {"productType":"MSU-L20S15-07","testDate":"2025-05-15","line":"","checkType":"function","variable":"Lot block recheck","variableDetail":"Lot Block (happen VP/CD)","variableGroup":"test","intervention":"Drop test + Aging + Recheck","inputQty":500,"okQty":333,"ngTotal":167,"ngRate":33.4,"defectCategory":"function_spl","defectType":"SPL+THD","defectCount":3},
        ],
        "tags":["msu-l20s15-07","ng-function","vp-cd-separate","lot-spk-50k","decap-analysis","hearing-noise"],
        "reportType":"intervention_test",
        "verdict":"worsened",
        "headline":"Lot SPK Block 50K recheck — function NG 33.4%, decap shows VP/CD separate 100%",
        "evidence":[
            {"metric":"Function NG rate","baselineLabel":"","baselineValue":"","variantLabel":"Lot Block (re-aged)","variantValue":"33.4% (167/500)","deltaText":"—","deltaSign":"up","note":"no normal baseline this sheet","comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Decap VP/CD separate","baselineLabel":"","baselineValue":"","variantLabel":"NG decap analysis","variantValue":"100% (15/15)","deltaText":"—","deltaSign":"up","note":"dominant failure mode","comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"investigate","text":"Investigate VP/CD separate root cause on Block 50K lot"},
            {"priority":2,"kind":"risk","text":"Lot SPK Block 50K cannot be released — function fail dominant"},
        ],
        "context":{"process":"VP/CD assembly + function","stage":"Block 50K lot recheck — drop test + aging + 100% function recheck","baselineReason":"no Normal arm; recheck of NG-flagged lot"},
        "doeGrid":None,"trendPoints":None,
        **LEGACY
    },
    tr_ko={**LEGACY,
        "headline":"Lot SPK Block 50K 재검사 — function NG 33.4%, decap에서 VP/CD 분리 100%",
        "actions":[
            {"priority":1,"kind":"investigate","text":"Block 50K lot의 VP/CD 분리 근본원인 조사"},
            {"priority":2,"kind":"risk","text":"Lot SPK Block 50K 출하 불가 — function 불량 지배적"},
        ],
        "context":{"process":"VP/CD 조립 + function","stage":"Block 50K lot 재검사 — drop test + aging + 100% function 재검사","baselineReason":"Normal arm 없음; NG flag된 lot의 재검사"}
    },
    tr_vi={**LEGACY,
        "headline":"Recheck Lot SPK Block 50K — function NG 33.4%, decap thấy VP/CD separate 100%",
        "actions":[
            {"priority":1,"kind":"investigate","text":"Điều tra nguyên nhân gốc VP/CD separate trên lot Block 50K"},
            {"priority":2,"kind":"risk","text":"Lot SPK Block 50K không xuất được — lỗi function chiếm đa số"},
        ],
        "context":{"process":"Lắp ráp VP/CD + function","stage":"Recheck lot Block 50K — drop test + aging + recheck function 100%","baselineReason":"không có Normal; recheck lot đã bị flag NG"}
    },
))

# ============================================================================
# 2. 11. MSU-L20S15-07 Report test Plasma suspension A,B - 2025.03.26
# ============================================================================
DATASETS.append(dict(
    name="11. MSU-L20S15-07 Report test Plasma suspension A,B to check tension Fr+Sus - 2025.03.26",
    product="MSU-L20S15-07",
    result={
        "measurements":[
            {"productType":"MSU-L20S15-07","testDate":"2025-03-26","line":"C1-F-PCB","checkType":"process","variable":"Tension Frame+Susp","variableDetail":"Plasma Suspension A","variableGroup":"test","intervention":"Plasma","inputQty":5,"okQty":5,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"MSU-L20S15-07","testDate":"2025-03-26","line":"C1-F-PCB","checkType":"process","variable":"Tension Frame+Susp","variableDetail":"Plasma Suspension B","variableGroup":"test","intervention":"Plasma","inputQty":5,"okQty":5,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"MSU-L20S15-07","testDate":"2025-03-26","line":"C1-F-PCB","checkType":"process","variable":"Tension Frame+Susp","variableDetail":"Normal Suspension A","variableGroup":"normal","intervention":"","inputQty":5,"okQty":5,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"MSU-L20S15-07","testDate":"2025-03-26","line":"C1-F-PCB","checkType":"process","variable":"Tension Frame+Susp","variableDetail":"Normal Suspension B","variableGroup":"normal","intervention":"","inputQty":5,"okQty":5,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        ],
        "tags":["msu-l20s15-07","plasma-suspension","tension-test","frame-suspension","reliability"],
        "reportType":"comparison_study",
        "verdict":"no_clear_effect",
        "headline":"Plasma suspension A/B tension same as Normal (delta 0.008-0.098 kgf)",
        "evidence":[
            {"metric":"Tension Susp A avg (Spec 0.3 kgf)","baselineLabel":"Normal","baselineValue":"0.623 kgf","variantLabel":"Plasma test","variantValue":"0.587 kgf","deltaText":"-0.036 kgf","deltaSign":"down","note":"both Pass spec","comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Tension Susp B avg (Spec 0.3 kgf)","baselineLabel":"Normal","baselineValue":"0.439 kgf","variantLabel":"Plasma test","variantValue":"0.467 kgf","deltaText":"+0.028 kgf","deltaSign":"up","note":"both Pass spec","comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Plasma suspension A,B safe to apply — tension equivalent to Normal"},
        ],
        "context":{"process":"Plasma suspension at C1-F-PCB then semi Frame+Suspension","stage":"A1 reliability tension test","baselineReason":"same-event Normal Suspension A/B rows present"},
        "doeGrid":None,"trendPoints":None,
        **LEGACY
    },
    tr_ko={**LEGACY,
        "headline":"Plasma suspension A/B 인장 Normal과 동등 (편차 0.008-0.098 kgf)",
        "actions":[
            {"priority":1,"kind":"action","text":"Plasma suspension A,B 적용 안전 — 인장 Normal 동등"},
        ],
        "context":{"process":"C1-F-PCB Plasma suspension 후 Frame+Suspension semi 제작","stage":"A1 신뢰성 인장 시험","baselineReason":"동일 이벤트에 Normal Suspension A/B 행 존재"}
    },
    tr_vi={**LEGACY,
        "headline":"Plasma suspension A/B lực căng tương đương Normal (chênh 0.008-0.098 kgf)",
        "actions":[
            {"priority":1,"kind":"action","text":"Plasma suspension A,B áp dụng an toàn — lực căng tương đương Normal"},
        ],
        "context":{"process":"Plasma suspension tại C1-F-PCB rồi làm semi Frame+Suspension","stage":"Test lực căng tin cậy tại A1","baselineReason":"có dòng Normal Suspension A/B cùng sự kiện"}
    },
))

# ============================================================================
# 3. 11. TIU C11-20 Report check VP damage and VP separate 2025.12.18
# ============================================================================
DATASETS.append(dict(
    name="11. TIU C11-20  Report check VP damage and VP separate 2025.12.18",
    product="TIU C11-20",
    result={
        "measurements":[
            {"productType":"TIU C11-20","testDate":"2025-12-19","line":"D3-1A","checkType":"visual_inspection","variable":"Grill type","variableDetail":"TIU C11-20L Old Grill","variableGroup":"before","intervention":"","inputQty":2266,"okQty":2259,"ngTotal":108,"ngRate":4.8,"defectCategory":"assembly_defect","defectType":"VP separate","defectCount":5},
            {"productType":"TIU C11-20","testDate":"2025-12-19","line":"D3-1A","checkType":"visual_inspection","variable":"Grill type","variableDetail":"TIU C11-20L Old Grill","variableGroup":"before","intervention":"","inputQty":2266,"okQty":2259,"ngTotal":108,"ngRate":4.8,"defectCategory":"cosmetic_defect","defectType":"VP damage / visual final","defectCount":103},
            {"productType":"TIU C11-20","testDate":"2025-12-19","line":"D3-1A","checkType":"visual_inspection","variable":"Grill type","variableDetail":"TIU C11-20L New Grill","variableGroup":"after","intervention":"New Grill","inputQty":1520,"okQty":1513,"ngTotal":59,"ngRate":3.9,"defectCategory":"assembly_defect","defectType":"VP separate","defectCount":7},
            {"productType":"TIU C11-20","testDate":"2025-12-19","line":"D3-1A","checkType":"visual_inspection","variable":"Grill type","variableDetail":"TIU C11-20L New Grill","variableGroup":"after","intervention":"New Grill","inputQty":1520,"okQty":1513,"ngTotal":59,"ngRate":3.9,"defectCategory":"cosmetic_defect","defectType":"VP damage / visual final","defectCount":52},
            {"productType":"TIU C11-20","testDate":"2025-12-19","line":"D3-2A","checkType":"visual_inspection","variable":"Grill type","variableDetail":"TIU C11-20R New Grill (5027)","variableGroup":"after","intervention":"New Grill","inputQty":5027,"okQty":4902,"ngTotal":227,"ngRate":4.5,"defectCategory":"assembly_defect","defectType":"VP separate","defectCount":120},
            {"productType":"TIU C11-20","testDate":"2025-12-19","line":"D3-2A","checkType":"visual_inspection","variable":"Grill type","variableDetail":"TIU C11-20R New Grill (5027)","variableGroup":"after","intervention":"New Grill","inputQty":5027,"okQty":4902,"ngTotal":227,"ngRate":4.5,"defectCategory":"cosmetic_defect","defectType":"VP damage / visual final","defectCount":107},
            {"productType":"TIU C11-20","testDate":"2025-12-19","line":"D3-2A","checkType":"visual_inspection","variable":"Grill type","variableDetail":"TIU C11-20R New Grill (1547)","variableGroup":"after","intervention":"New Grill","inputQty":1547,"okQty":1543,"ngTotal":66,"ngRate":4.3,"defectCategory":"assembly_defect","defectType":"VP separate","defectCount":2},
            {"productType":"TIU C11-20","testDate":"2025-12-19","line":"D3-2A","checkType":"visual_inspection","variable":"Grill type","variableDetail":"TIU C11-20R New Grill (1547)","variableGroup":"after","intervention":"New Grill","inputQty":1547,"okQty":1543,"ngTotal":66,"ngRate":4.3,"defectCategory":"cosmetic_defect","defectType":"VP damage / visual final","defectCount":64},
        ],
        "tags":["tiu-c11-20","vp-separate","vp-damage","grill-change","visual-final","night-shift"],
        "reportType":"comparison_study",
        "verdict":"no_clear_effect",
        "headline":"New Grill VP separate/damage NG 3.9-4.5% — similar to Old Grill 4.8%",
        "evidence":[
            {"metric":"Total NG rate C11-20L","baselineLabel":"Old Grill","baselineValue":"4.8% (108/2266)","variantLabel":"New Grill","variantValue":"3.9% (59/1520)","deltaText":"-0.9pp","deltaSign":"down","note":"D3-1A night shift","comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Total NG rate C11-20R","baselineLabel":"","baselineValue":"","variantLabel":"New Grill","variantValue":"4.3-4.5% (227/5027, 66/1547)","deltaText":"—","deltaSign":"no_change","note":"D3-2A no old-grill arm","comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"investigate","text":"Investigate dominant visual-final NG (VP damage) on D3-2A line"},
            {"priority":2,"kind":"action","text":"Continue New Grill use — VP separate rate 0.13-2.39% acceptable"},
        ],
        "context":{"process":"VP separate + visual final check (TIU C11-20)","stage":"D3-1A and D3-2A night shift","baselineReason":"Old vs New Grill within same model"},
        "doeGrid":None,"trendPoints":None,
        **LEGACY
    },
    tr_ko={**LEGACY,
        "headline":"New Grill VP separate/damage NG 3.9-4.5% — Old Grill 4.8%와 동등",
        "actions":[
            {"priority":1,"kind":"investigate","text":"D3-2A 라인 visual-final NG (VP damage) 지배 원인 조사"},
            {"priority":2,"kind":"action","text":"New Grill 계속 사용 — VP separate 0.13-2.39% 양호"},
        ],
        "context":{"process":"VP separate + 최종 외관 검사 (TIU C11-20)","stage":"D3-1A, D3-2A 야간 근무","baselineReason":"동일 모델 내 Old vs New Grill 비교"}
    },
    tr_vi={**LEGACY,
        "headline":"New Grill VP separate/damage NG 3.9-4.5% — tương đương Old Grill 4.8%",
        "actions":[
            {"priority":1,"kind":"investigate","text":"Điều tra nguyên nhân NG visual-final (VP damage) chiếm đa số trên line D3-2A"},
            {"priority":2,"kind":"action","text":"Tiếp tục sử dụng New Grill — VP separate 0.13-2.39% chấp nhận được"},
        ],
        "context":{"process":"Kiểm tra VP separate + visual final (TIU C11-20)","stage":"Ca đêm D3-1A và D3-2A","baselineReason":"So sánh Old Grill vs New Grill cùng model"}
    },
))

# ============================================================================
# 4. 11. TIU C11-20 Report test VP improve NG hearing high 2025.12.17
# ============================================================================
DATASETS.append(dict(
    name="11. TIU C11-20  Report test VP improve NG hearing high 2025.12.17",
    product="TIU C11-20",
    result={
        "measurements":[
            {"productType":"TIU C11-20","testDate":"2025-12-16","line":"","checkType":"function","variable":"VP height","variableDetail":"VP height 1.4~1.46mm","variableGroup":"test","intervention":"VP height range","inputQty":39,"okQty":34,"ngTotal":5,"ngRate":12.8,"defectCategory":"function_hearing","defectType":"SPL+RB","defectCount":2},
            {"productType":"TIU C11-20","testDate":"2025-12-16","line":"","checkType":"function","variable":"VP height","variableDetail":"VP height 1.4~1.46mm","variableGroup":"test","intervention":"VP height range","inputQty":39,"okQty":34,"ngTotal":5,"ngRate":12.8,"defectCategory":"function_spl","defectType":"RB","defectCount":3},
            {"productType":"TIU C11-20","testDate":"2025-12-16","line":"","checkType":"function","variable":"VP height","variableDetail":"VP height 1.4~1.46mm","variableGroup":"test","intervention":"VP height range","inputQty":39,"okQty":34,"ngTotal":5,"ngRate":12.8,"defectCategory":"function_hearing","defectType":"Noise","defectCount":3},
            {"productType":"TIU C11-20","testDate":"2025-12-16","line":"","checkType":"function","variable":"VP height","variableDetail":"VP height >1.47mm","variableGroup":"test","intervention":"VP height range","inputQty":7,"okQty":7,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"TIU C11-20","testDate":"2025-12-16","line":"","checkType":"function","variable":"VP height","variableDetail":"Normal","variableGroup":"normal","intervention":"","inputQty":30,"okQty":30,"ngTotal":1,"ngRate":3.3,"defectCategory":"function_hearing","defectType":"SPL+RB","defectCount":1},
        ],
        "tags":["tiu-c11-20","vp-height","ng-hearing","rb-defect","function-test","sample-small"],
        "reportType":"multi_arm",
        "verdict":"partial",
        "headline":"VP height 1.4-1.46mm hearing NG 12.8% worse than Normal 3.3%; >1.47mm OK n=7",
        "evidence":[
            {"metric":"Hearing function NG","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"+12.8pp range","deltaSign":"up","note":"small n on >1.47mm arm",
             "comparisons":[
                {"label":"VP >1.47mm","value":"0.0% (0/7)","n":7,"isBaseline":False,"isBest":True,"isWorst":False},
                {"label":"Normal","value":"3.3% (1/30)","n":30,"isBaseline":True,"isBest":False,"isWorst":False},
                {"label":"VP 1.4~1.46mm","value":"12.8% (5/39)","n":39,"isBaseline":False,"isBest":False,"isWorst":True},
             ],
             "bestLabel":"VP >1.47mm","worstLabel":"VP 1.4~1.46mm"},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Raise VP height spec above 1.47mm to suppress hearing-RB NG"},
            {"priority":2,"kind":"investigate","text":"Confirm VP >1.47mm at larger n before mass adoption"},
        ],
        "context":{"process":"VP height variation in final assembly","stage":"Function test final samples","baselineReason":"same-event Normal row present"},
        "doeGrid":None,"trendPoints":None,
        **LEGACY
    },
    tr_ko={**LEGACY,
        "headline":"VP 높이 1.4-1.46mm hearing NG 12.8% (Normal 3.3%보다 악화); >1.47mm OK n=7",
        "actions":[
            {"priority":1,"kind":"action","text":"VP 높이 spec을 1.47mm 이상으로 상향해 hearing-RB NG 억제"},
            {"priority":2,"kind":"investigate","text":"양산 적용 전 VP >1.47mm 추가 n으로 재확인"},
        ],
        "context":{"process":"최종 조립 시 VP 높이 변동","stage":"최종 샘플 function test","baselineReason":"동일 이벤트에 Normal 행 존재"}
    },
    tr_vi={**LEGACY,
        "headline":"VP cao 1.4-1.46mm hearing NG 12.8% (kém hơn Normal 3.3%); >1.47mm OK n=7",
        "actions":[
            {"priority":1,"kind":"action","text":"Nâng spec chiều cao VP trên 1.47mm để giảm NG hearing-RB"},
            {"priority":2,"kind":"investigate","text":"Xác nhận VP >1.47mm với n lớn hơn trước khi áp dụng đại trà"},
        ],
        "context":{"process":"Biến động chiều cao VP trong lắp ráp cuối","stage":"Test function mẫu cuối","baselineReason":"có dòng Normal cùng sự kiện"}
    },
))

# ============================================================================
# 5. 11. TIU L5S3-01 L Frame Magnet side colour 100% - 2026.05.08
# ============================================================================
DATASETS.append(dict(
    name="11. TIU L5S3-01 L - Report test material Frame happen NG Magnet side diffrence colour 100% - date 2026.05.08",
    product="TIU L5S3-01 L",
    result={
        "measurements":[
            {"productType":"TIU L5S3-01 L","testDate":"2026-05-08","line":"Sub3","checkType":"process","variable":"Frame material","variableDetail":"Test (Baotou Magnet side colour diff 100%)","variableGroup":"test","intervention":"","inputQty":148,"okQty":146,"ngTotal":2,"ngRate":1.4,"defectCategory":"assembly_defect","defectType":"Frame Array not pick up","defectCount":2},
            {"productType":"TIU L5S3-01 L","testDate":"2026-05-08","line":"Sub3","checkType":"process","variable":"Frame material","variableDetail":"Test (Baotou Magnet side colour diff 100%) — AI bond","variableGroup":"test","intervention":"","inputQty":148,"okQty":125,"ngTotal":23,"ngRate":15.5,"defectCategory":"assembly_defect","defectType":"Lack glue","defectCount":23},
            {"productType":"TIU L5S3-01 L","testDate":"2026-05-08","line":"Sub3","checkType":"process","variable":"Frame material","variableDetail":"Normal — Frame Array","variableGroup":"normal","intervention":"","inputQty":150,"okQty":150,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"TIU L5S3-01 L","testDate":"2026-05-08","line":"Sub3","checkType":"process","variable":"Frame material","variableDetail":"Normal — AI bond Frame","variableGroup":"normal","intervention":"","inputQty":150,"okQty":150,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        ],
        "tags":["tiu-l5s3-01","frame-material","baotou-vendor","magnet-side-colour","lack-glue","ai-bond","inconclusive"],
        "reportType":"comparison_study",
        "verdict":"inconclusive",
        "headline":"Baotou Frame test AI-bond lack-glue 15.5% vs Normal 0% — function data missing (#DIV/0)",
        "evidence":[
            {"metric":"Frame Array NG","baselineLabel":"Normal","baselineValue":"0.0% (0/150)","variantLabel":"Test","variantValue":"1.4% (2/148)","deltaText":"+1.4pp","deltaSign":"up","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"AI bond Frame NG","baselineLabel":"Normal","baselineValue":"0.0% (0/150)","variantLabel":"Test","variantValue":"15.5% (23/148)","deltaText":"+15.5pp","deltaSign":"up","note":"all lack-glue","comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Function NG","baselineLabel":"Normal","baselineValue":"#DIV/0","variantLabel":"Test","variantValue":"#DIV/0","deltaText":"—","deltaSign":"no_change","note":"no function input recorded","comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Repeat test with completed function and visual-final data before decision"},
            {"priority":2,"kind":"risk","text":"Do not release Baotou colour-diff frames — AI lack-glue 15.5% vs 0%"},
        ],
        "context":{"process":"Frame Array → AI bond Frame → Ass'y Frame+Yoke (Sub3)","stage":"Make semi 150pcs, then final function","baselineReason":"same-event Normal row present"},
        "doeGrid":None,"trendPoints":None,
        **LEGACY
    },
    tr_ko={**LEGACY,
        "headline":"Baotou Frame 테스트 AI-bond lack-glue 15.5% vs Normal 0% — 기능 데이터 결측 (#DIV/0)",
        "actions":[
            {"priority":1,"kind":"action","text":"Function 및 visual-final 데이터 완비 후 재시험 필요"},
            {"priority":2,"kind":"risk","text":"Baotou 색상 차이 Frame 출하 금지 — AI lack-glue 15.5% vs 0%"},
        ],
        "context":{"process":"Frame Array → AI bond Frame → Ass'y Frame+Yoke (Sub3)","stage":"150pcs semi 제작 후 최종 function","baselineReason":"동일 이벤트에 Normal 행 존재"}
    },
    tr_vi={**LEGACY,
        "headline":"Test Frame Baotou AI-bond lack-glue 15.5% vs Normal 0% — thiếu dữ liệu function (#DIV/0)",
        "actions":[
            {"priority":1,"kind":"action","text":"Test lại với đầy đủ dữ liệu function và visual-final trước khi quyết định"},
            {"priority":2,"kind":"risk","text":"Không xuất Frame Baotou khác màu — AI lack-glue 15.5% so với 0%"},
        ],
        "context":{"process":"Frame Array → AI bond Frame → Ass'y Frame+Yoke (Sub3)","stage":"Làm semi 150pcs rồi function cuối","baselineReason":"có dòng Normal cùng sự kiện"}
    },
))

# ============================================================================
# 6. 11.1 BRS-161016 Report test material YK NG over flatness 2025.07.17
# ============================================================================
DATASETS.append(dict(
    name="11.1 BRS-161016 Report test material YK happen  NG over flatness test again 2nd  Date 17.7.2025",
    product="BRS-161016",
    result={
        "measurements":[
            # Decap bond MG-S-A/B
            {"productType":"BRS-161016","testDate":"2025-07-17","line":"","checkType":"process","variable":"Decap bond","variableDetail":"MG-S-A/B Test YK NG over flatness","variableGroup":"test","intervention":"","inputQty":10,"okQty":10,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161016","testDate":"2025-07-17","line":"","checkType":"process","variable":"Decap bond","variableDetail":"MG-S-A/B Normal","variableGroup":"normal","intervention":"","inputQty":8,"okQty":8,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            # Decap bond Yoke
            {"productType":"BRS-161016","testDate":"2025-07-17","line":"","checkType":"process","variable":"Decap bond Yoke","variableDetail":"Test YK NG over flatness","variableGroup":"test","intervention":"","inputQty":10,"okQty":10,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161016","testDate":"2025-07-17","line":"","checkType":"process","variable":"Decap bond Yoke","variableDetail":"Normal","variableGroup":"normal","intervention":"","inputQty":8,"okQty":8,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            # Drop test
            {"productType":"BRS-161016","testDate":"2025-07-17","line":"","checkType":"process","variable":"Drop test","variableDetail":"Test YK NG over flatness","variableGroup":"test","intervention":"","inputQty":5,"okQty":5,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161016","testDate":"2025-07-17","line":"","checkType":"process","variable":"Drop test","variableDetail":"Normal","variableGroup":"normal","intervention":"","inputQty":5,"okQty":5,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            # Tension MG-S-A spec >= 2.5 kgf
            {"productType":"BRS-161016","testDate":"2025-07-17","line":"","checkType":"process","variable":"Tension MG-S-A","variableDetail":"Test (Spec >=2.5 kgf)","variableGroup":"test","intervention":"","inputQty":5,"okQty":5,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161016","testDate":"2025-07-17","line":"","checkType":"process","variable":"Tension MG-S-B","variableDetail":"Test (Spec >=5.0 kgf)","variableGroup":"test","intervention":"","inputQty":5,"okQty":5,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-161016","testDate":"2025-07-17","line":"","checkType":"process","variable":"Tension MG-S-A/B Normal","variableDetail":"Normal","variableGroup":"normal","intervention":"","inputQty":5,"okQty":5,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            # Gauss (Semi Yoke S>=480G)
            {"productType":"BRS-161016","testDate":"2025-07-17","line":"","checkType":"process","variable":"Gauss","variableDetail":"Semi Yoke Test (Spec >=480G)","variableGroup":"test","intervention":"","inputQty":477,"okQty":476,"ngTotal":1,"ngRate":0.2,"defectCategory":"magnetic_defect","defectType":"NG Low Gauss","defectCount":1},
            {"productType":"BRS-161016","testDate":"2025-07-17","line":"","checkType":"process","variable":"Gauss","variableDetail":"Normal","variableGroup":"normal","intervention":"","inputQty":500,"okQty":500,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            # Function C2-3A
            {"productType":"BRS-161016","testDate":"2025-07-17","line":"C2-3A","checkType":"function","variable":"Function","variableDetail":"Test YK NG over flatness","variableGroup":"test","intervention":"","inputQty":467,"okQty":453,"ngTotal":14,"ngRate":3.00,"defectCategory":"function_hearing","defectType":"Noise","defectCount":7},
            {"productType":"BRS-161016","testDate":"2025-07-17","line":"C2-3A","checkType":"function","variable":"Function","variableDetail":"Test YK NG over flatness","variableGroup":"test","intervention":"","inputQty":467,"okQty":453,"ngTotal":14,"ngRate":3.00,"defectCategory":"function_hearing","defectType":"Touch","defectCount":6},
            {"productType":"BRS-161016","testDate":"2025-07-17","line":"C2-3A","checkType":"function","variable":"Function","variableDetail":"Test YK NG over flatness","variableGroup":"test","intervention":"","inputQty":467,"okQty":453,"ngTotal":14,"ngRate":3.00,"defectCategory":"function_spl","defectType":"SPL+THD","defectCount":1},
            {"productType":"BRS-161016","testDate":"2025-07-17","line":"C2-3A","checkType":"function","variable":"Function","variableDetail":"Normal","variableGroup":"normal","intervention":"","inputQty":1680,"okQty":1644,"ngTotal":36,"ngRate":2.14,"defectCategory":"function_hearing","defectType":"Noise","defectCount":14},
            {"productType":"BRS-161016","testDate":"2025-07-17","line":"C2-3A","checkType":"function","variable":"Function","variableDetail":"Normal","variableGroup":"normal","intervention":"","inputQty":1680,"okQty":1644,"ngTotal":36,"ngRate":2.14,"defectCategory":"function_hearing","defectType":"Touch","defectCount":21},
            {"productType":"BRS-161016","testDate":"2025-07-17","line":"C2-3A","checkType":"function","variable":"Function","variableDetail":"Normal","variableGroup":"normal","intervention":"","inputQty":1680,"okQty":1644,"ngTotal":36,"ngRate":2.14,"defectCategory":"function_spl","defectType":"SPL+THD","defectCount":1},
        ],
        "tags":["brs-161016","yoke-material","over-flatness","decap-bond","tension-test","gauss","function-ng","hearing"],
        "reportType":"comparison_study",
        "verdict":"worsened",
        "headline":"YK over-flatness function NG 3.00% > Normal 2.14% with low-gauss 1/477 — cannot use",
        "evidence":[
            {"metric":"Function NG (C2-3A)","baselineLabel":"Normal","baselineValue":"2.14% (36/1680)","variantLabel":"Test YK NG over flatness","variantValue":"3.00% (14/467)","deltaText":"+0.86pp","deltaSign":"up","note":"hearing-noise+touch dominant","comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Gauss low-G","baselineLabel":"Normal","baselineValue":"0.0% (0/500)","variantLabel":"Test","variantValue":"0.2% (1/477)","deltaText":"+0.2pp","deltaSign":"up","note":"common Ruiji MG-C lot","comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Decap/Drop/Tension","baselineLabel":"Normal","baselineValue":"all OK","variantLabel":"Test","variantValue":"all OK","deltaText":"—","deltaSign":"no_change","note":"bond + drop + tension pass","comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Reject YK over-flatness material — function NG and low-gauss exceed Normal"},
            {"priority":2,"kind":"investigate","text":"Investigate FR+coil offset, CM offset, Glue clots, Fr+Yoke offset (each ~15%)"},
            {"priority":3,"kind":"investigate","text":"Trace 38% unknown-reason hearing NG via repeated decap analysis"},
        ],
        "context":{"process":"Yoke material qualification (decap bond, drop, tension, gauss, function)","stage":"C2-3A line + sub-yoke 161016-D2 visual","baselineReason":"same-event Normal arm present in every test"},
        "doeGrid":None,"trendPoints":None,
        **LEGACY
    },
    tr_ko={**LEGACY,
        "headline":"YK over-flatness function NG 3.00% > Normal 2.14%, low-gauss 1/477 — 사용 불가",
        "actions":[
            {"priority":1,"kind":"action","text":"YK over-flatness 재료 부적합 — function NG, low-gauss가 Normal 초과"},
            {"priority":2,"kind":"investigate","text":"FR+coil offset, CM offset, Glue clots, Fr+Yoke offset (각 ~15%) 조사"},
            {"priority":3,"kind":"investigate","text":"38% Unknown-reason hearing NG를 추가 decap 분석으로 추적"},
        ],
        "context":{"process":"Yoke 재료 인증 (decap bond, drop, tension, gauss, function)","stage":"C2-3A 라인 + sub-yoke 161016-D2 visual","baselineReason":"모든 시험에 동일 이벤트 Normal arm 존재"}
    },
    tr_vi={**LEGACY,
        "headline":"YK over-flatness function NG 3.00% > Normal 2.14%, low-gauss 1/477 — không sử dụng được",
        "actions":[
            {"priority":1,"kind":"action","text":"Loại bỏ vật liệu YK over-flatness — function NG và low-gauss vượt Normal"},
            {"priority":2,"kind":"investigate","text":"Điều tra FR+coil offset, CM offset, Glue clots, Fr+Yoke offset (mỗi ~15%)"},
            {"priority":3,"kind":"investigate","text":"Truy nguyên 38% NG hearing không rõ nguyên nhân qua decap lặp"},
        ],
        "context":{"process":"Đánh giá vật liệu Yoke (decap bond, drop, tension, gauss, function)","stage":"Line C2-3A + visual sub-yoke 161016-D2","baselineReason":"có arm Normal cùng sự kiện ở mọi test"}
    },
))

# ============================================================================
# 7. 11.BRS-2015 Ring Fader paint - 2024.09.13
# ============================================================================
DATASETS.append(dict(
    name="11.BRS-2015 Report Test material Ring Fader paint than usual - date 13.9.2024",
    product="BRS-2015",
    result={
        "measurements":[
            {"productType":"BRS-2015","testDate":"2024-09-13","line":"VP line","checkType":"visual_inspection","variable":"Ring material","variableDetail":"Test Ring Fader paint","variableGroup":"test","intervention":"","inputQty":96,"okQty":96,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-2015","testDate":"2024-09-13","line":"VP line","checkType":"visual_inspection","variable":"Ring material","variableDetail":"Normal","variableGroup":"normal","intervention":"","inputQty":96,"okQty":96,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        ],
        "tags":["brs-2015","ring-material","fader-paint","vp-line","separate-ring-vp","first-test"],
        "reportType":"comparison_study",
        "verdict":"no_clear_effect",
        "headline":"Ring Fader paint test 0/96 NG vs Normal 0/96 NG at VP line — can use (recycle pending)",
        "evidence":[
            {"metric":"VP line separate NG","baselineLabel":"Normal","baselineValue":"0.0% (0/96)","variantLabel":"Test Ring Fader","variantValue":"0.0% (0/96)","deltaText":"+0pp","deltaSign":"no_change","note":"recycle Ring not yet checked","comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Release Ring Fader paint material for VP line"},
            {"priority":2,"kind":"investigate","text":"Re-test recycled Ring before full approval"},
        ],
        "context":{"process":"VP line — Ring+VP separation check","stage":"Make semi at sub VP line","baselineReason":"same-event Normal arm present"},
        "doeGrid":None,"trendPoints":None,
        **LEGACY
    },
    tr_ko={**LEGACY,
        "headline":"Ring Fader paint 테스트 0/96 NG vs Normal 0/96 NG (VP 라인) — 사용 가능 (재활용 미확인)",
        "actions":[
            {"priority":1,"kind":"action","text":"VP 라인용 Ring Fader paint 재료 출하 가능"},
            {"priority":2,"kind":"investigate","text":"전체 승인 전 재활용 Ring 재시험"},
        ],
        "context":{"process":"VP 라인 — Ring+VP 분리 점검","stage":"Sub VP 라인에서 semi 제작","baselineReason":"동일 이벤트에 Normal arm 존재"}
    },
    tr_vi={**LEGACY,
        "headline":"Test Ring Fader paint 0/96 NG vs Normal 0/96 NG tại VP line — sử dụng được (chưa kiểm Ring tái chế)",
        "actions":[
            {"priority":1,"kind":"action","text":"Cho phép vật liệu Ring Fader paint cho VP line"},
            {"priority":2,"kind":"investigate","text":"Test lại Ring tái chế trước khi duyệt hoàn toàn"},
        ],
        "context":{"process":"VP line — kiểm tra Ring+VP separate","stage":"Làm semi tại sub VP line","baselineReason":"có arm Normal cùng sự kiện"}
    },
))

# ============================================================================
# 8. 11.BRS-201506 TIP U844 -> U864 - 2024.12.26
# ============================================================================
DATASETS.append(dict(
    name="11.BRS-201506 Report  test change new tip U844- U864 improve NG spot offset date 26.12.2024 -",
    product="BRS-201506",
    result={
        "measurements":[
            # TIP 864 before/after repair (12/26)
            {"productType":"BRS-201506","testDate":"2024-12-26","line":"","checkType":"process","variable":"Spot welding tip","variableDetail":"TIP 864 — before repair","variableGroup":"test","intervention":"TIP U864","inputQty":4930,"okQty":4821,"ngTotal":109,"ngRate":2.2,"defectCategory":"assembly_defect","defectType":"Weak solder","defectCount":90},
            {"productType":"BRS-201506","testDate":"2024-12-26","line":"","checkType":"process","variable":"Spot welding tip","variableDetail":"TIP 864 — before repair","variableGroup":"test","intervention":"TIP U864","inputQty":4930,"okQty":4821,"ngTotal":109,"ngRate":2.2,"defectCategory":"assembly_defect","defectType":"Spot offset","defectCount":19},
            {"productType":"BRS-201506","testDate":"2024-12-26","line":"","checkType":"process","variable":"Spot welding tip","variableDetail":"TIP 864 — after repair","variableGroup":"test","intervention":"TIP U864","inputQty":4930,"okQty":4902,"ngTotal":28,"ngRate":0.57,"defectCategory":"assembly_defect","defectType":"Spot offset","defectCount":19},
            {"productType":"BRS-201506","testDate":"2024-12-26","line":"","checkType":"process","variable":"Spot welding tip","variableDetail":"TIP 864 — after repair","variableGroup":"test","intervention":"TIP U864","inputQty":4930,"okQty":4902,"ngTotal":28,"ngRate":0.57,"defectCategory":"assembly_defect","defectType":"Weak solder","defectCount":9},
            # TIP 844 Normal before/after
            {"productType":"BRS-201506","testDate":"2024-12-26","line":"","checkType":"process","variable":"Spot welding tip","variableDetail":"TIP 844 Normal — before repair","variableGroup":"normal","intervention":"TIP U844","inputQty":3450,"okQty":3370,"ngTotal":80,"ngRate":2.3,"defectCategory":"assembly_defect","defectType":"Weak solder","defectCount":30},
            {"productType":"BRS-201506","testDate":"2024-12-26","line":"","checkType":"process","variable":"Spot welding tip","variableDetail":"TIP 844 Normal — before repair","variableGroup":"normal","intervention":"TIP U844","inputQty":3450,"okQty":3370,"ngTotal":80,"ngRate":2.3,"defectCategory":"assembly_defect","defectType":"Spot offset","defectCount":50},
            {"productType":"BRS-201506","testDate":"2024-12-26","line":"","checkType":"process","variable":"Spot welding tip","variableDetail":"TIP 844 Normal — after repair","variableGroup":"normal","intervention":"TIP U844","inputQty":3450,"okQty":3384,"ngTotal":66,"ngRate":1.91,"defectCategory":"assembly_defect","defectType":"Spot offset","defectCount":50},
            {"productType":"BRS-201506","testDate":"2024-12-26","line":"","checkType":"process","variable":"Spot welding tip","variableDetail":"TIP 844 Normal — after repair","variableGroup":"normal","intervention":"TIP U844","inputQty":3450,"okQty":3384,"ngTotal":66,"ngRate":1.91,"defectCategory":"assembly_defect","defectType":"Weak solder","defectCount":16},
            # TIP 864 12/27
            {"productType":"BRS-201506","testDate":"2024-12-27","line":"","checkType":"process","variable":"Spot welding tip","variableDetail":"TIP 864 — before repair","variableGroup":"test","intervention":"TIP U864","inputQty":10333,"okQty":10040,"ngTotal":293,"ngRate":2.8,"defectCategory":"assembly_defect","defectType":"Weak solder","defectCount":227},
            {"productType":"BRS-201506","testDate":"2024-12-27","line":"","checkType":"process","variable":"Spot welding tip","variableDetail":"TIP 864 — before repair","variableGroup":"test","intervention":"TIP U864","inputQty":10333,"okQty":10040,"ngTotal":293,"ngRate":2.8,"defectCategory":"assembly_defect","defectType":"Spot offset","defectCount":29},
            {"productType":"BRS-201506","testDate":"2024-12-27","line":"","checkType":"process","variable":"Spot welding tip","variableDetail":"TIP 864 — before repair","variableGroup":"test","intervention":"TIP U864","inputQty":10333,"okQty":10040,"ngTotal":293,"ngRate":2.8,"defectCategory":"assembly_defect","defectType":"Suspension damage","defectCount":33},
            {"productType":"BRS-201506","testDate":"2024-12-27","line":"","checkType":"process","variable":"Spot welding tip","variableDetail":"TIP 864 — after repair","variableGroup":"test","intervention":"TIP U864","inputQty":10333,"okQty":10251,"ngTotal":82,"ngRate":0.79,"defectCategory":"assembly_defect","defectType":"Spot offset","defectCount":29},
            {"productType":"BRS-201506","testDate":"2024-12-27","line":"","checkType":"process","variable":"Spot welding tip","variableDetail":"TIP 864 — after repair","variableGroup":"test","intervention":"TIP U864","inputQty":10333,"okQty":10251,"ngTotal":82,"ngRate":0.79,"defectCategory":"assembly_defect","defectType":"Suspension damage","defectCount":33},
            {"productType":"BRS-201506","testDate":"2024-12-27","line":"","checkType":"process","variable":"Spot welding tip","variableDetail":"TIP 864 — after repair","variableGroup":"test","intervention":"TIP U864","inputQty":10333,"okQty":10251,"ngTotal":82,"ngRate":0.79,"defectCategory":"assembly_defect","defectType":"Weak solder","defectCount":16},
            # Function
            {"productType":"BRS-201506","testDate":"2024-12-27","line":"C2-2A","checkType":"function","variable":"Function","variableDetail":"TIP 864","variableGroup":"test","intervention":"TIP U864","inputQty":5464,"okQty":4666,"ngTotal":892,"ngRate":16.3,"defectCategory":"function_hearing","defectType":"Touch","defectCount":615},
            {"productType":"BRS-201506","testDate":"2024-12-27","line":"C2-2A","checkType":"function","variable":"Function","variableDetail":"TIP 864","variableGroup":"test","intervention":"TIP U864","inputQty":5464,"okQty":4666,"ngTotal":892,"ngRate":16.3,"defectCategory":"function_hearing","defectType":"Noise","defectCount":277},
            {"productType":"BRS-201506","testDate":"2024-12-27","line":"C2-2A","checkType":"function","variable":"Function","variableDetail":"TIP 844 Normal","variableGroup":"normal","intervention":"TIP U844","inputQty":1658,"okQty":1346,"ngTotal":312,"ngRate":18.8,"defectCategory":"function_hearing","defectType":"Touch","defectCount":200},
            {"productType":"BRS-201506","testDate":"2024-12-27","line":"C2-2A","checkType":"function","variable":"Function","variableDetail":"TIP 844 Normal","variableGroup":"normal","intervention":"TIP U844","inputQty":1658,"okQty":1346,"ngTotal":312,"ngRate":18.8,"defectCategory":"function_hearing","defectType":"Noise","defectCount":112},
        ],
        "tags":["brs-201506","tip-change","u844-u864","spot-welding","spot-offset","weak-solder","function-test"],
        "reportType":"comparison_study",
        "verdict":"improved",
        "headline":"TIP U864 spot-weld NG 0.72% after repair (vs U844 1.91%); function NG 16.3% vs 18.8% Normal",
        "evidence":[
            {"metric":"Spot welding NG (after repair, total)","baselineLabel":"TIP U844 Normal","baselineValue":"1.91% (66/3450)","variantLabel":"TIP U864","variantValue":"0.72% (110/15263)","deltaText":"-1.19pp","deltaSign":"down","note":"spot offset 0.3% vs 1.4%","comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Function NG (C2-2A)","baselineLabel":"TIP U844 Normal","baselineValue":"18.8% (312/1658)","variantLabel":"TIP U864","variantValue":"16.3% (892/5464)","deltaText":"-2.5pp","deltaSign":"down","note":"hearing-touch dominant both arms","comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Adopt TIP U864 for spot welding — spot offset 0.3% vs 1.4%"},
            {"priority":2,"kind":"investigate","text":"Address hearing-touch dominant function NG independently of tip"},
        ],
        "context":{"process":"Frame+Susp spot welding tip change","stage":"C2-2A line — q'ty test until tip over life","baselineReason":"TIP U844 same-line same-period normal"},
        "doeGrid":None,"trendPoints":None,
        **LEGACY
    },
    tr_ko={**LEGACY,
        "headline":"TIP U864 spot welding NG 수리 후 0.72% (U844 1.91% 대비); function NG 16.3% vs Normal 18.8%",
        "actions":[
            {"priority":1,"kind":"action","text":"Spot welding에 TIP U864 채택 — spot offset 0.3% vs 1.4%"},
            {"priority":2,"kind":"investigate","text":"Hearing-touch 지배 function NG는 tip과 별개로 대책"},
        ],
        "context":{"process":"Frame+Susp spot welding tip 변경","stage":"C2-2A 라인 — tip 수명 종료까지 q'ty test","baselineReason":"TIP U844 동일 라인·동일 기간 Normal"}
    },
    tr_vi={**LEGACY,
        "headline":"TIP U864 spot welding NG 0.72% sau sửa (so U844 1.91%); function NG 16.3% vs Normal 18.8%",
        "actions":[
            {"priority":1,"kind":"action","text":"Áp dụng TIP U864 cho spot welding — spot offset 0.3% so 1.4%"},
            {"priority":2,"kind":"investigate","text":"Xử lý NG function hearing-touch chiếm đa số độc lập với tip"},
        ],
        "context":{"process":"Đổi tip spot welding Frame+Susp","stage":"Line C2-2A — q'ty test đến khi tip hết tuổi thọ","baselineReason":"TIP U844 Normal cùng line cùng thời điểm"}
    },
))

# ============================================================================
# 9. 11.BRS-201506 Report checking problem process NG high - 02.02.2024
# ============================================================================
DATASETS.append(dict(
    name="11.BRS-201506 Report checking problem process NG high date 2.2.2024",
    product="BRS-201506",
    result={
        "measurements":[
            # 25/May solder vision AWF#1-5
            {"productType":"BRS-201506","testDate":"2024-05-25","line":"","checkType":"process","variable":"Solder vision AWF","variableDetail":"AWF#1","variableGroup":"test","intervention":"","inputQty":16,"okQty":11,"ngTotal":5,"ngRate":31.25,"defectCategory":"assembly_defect","defectType":"Weak solder","defectCount":5},
            {"productType":"BRS-201506","testDate":"2024-05-25","line":"","checkType":"process","variable":"Solder vision AWF","variableDetail":"AWF#2","variableGroup":"test","intervention":"","inputQty":16,"okQty":12,"ngTotal":4,"ngRate":25.00,"defectCategory":"assembly_defect","defectType":"Weak solder","defectCount":4},
            {"productType":"BRS-201506","testDate":"2024-05-25","line":"","checkType":"process","variable":"Solder vision AWF","variableDetail":"AWF#3","variableGroup":"test","intervention":"","inputQty":16,"okQty":11,"ngTotal":5,"ngRate":31.25,"defectCategory":"assembly_defect","defectType":"Weak solder","defectCount":5},
            {"productType":"BRS-201506","testDate":"2024-05-25","line":"","checkType":"process","variable":"Solder vision AWF","variableDetail":"AWF#4","variableGroup":"test","intervention":"","inputQty":16,"okQty":14,"ngTotal":2,"ngRate":12.50,"defectCategory":"assembly_defect","defectType":"Weak solder","defectCount":2},
            {"productType":"BRS-201506","testDate":"2024-05-25","line":"","checkType":"process","variable":"Solder vision AWF","variableDetail":"AWF#5","variableGroup":"test","intervention":"","inputQty":16,"okQty":14,"ngTotal":2,"ngRate":12.50,"defectCategory":"assembly_defect","defectType":"Weak solder","defectCount":2},
            # Frame+sus dates
            {"productType":"BRS-201506","testDate":"2024-05-20","line":"","checkType":"process","variable":"Frame+Sus solder","variableDetail":"Frame+sus 5/20","variableGroup":"test","intervention":"","inputQty":144,"okQty":136,"ngTotal":8,"ngRate":5.56,"defectCategory":"assembly_defect","defectType":"Weak solder","defectCount":8},
            {"productType":"BRS-201506","testDate":"2024-05-21","line":"","checkType":"process","variable":"Frame+Sus solder","variableDetail":"Frame+sus 5/21","variableGroup":"test","intervention":"","inputQty":58,"okQty":52,"ngTotal":6,"ngRate":10.34,"defectCategory":"assembly_defect","defectType":"Weak solder","defectCount":6},
            {"productType":"BRS-201506","testDate":"2024-05-22","line":"","checkType":"process","variable":"Frame+Sus solder","variableDetail":"Frame+sus 5/22","variableGroup":"test","intervention":"","inputQty":140,"okQty":137,"ngTotal":3,"ngRate":2.14,"defectCategory":"assembly_defect","defectType":"Weak solder","defectCount":3},
            {"productType":"BRS-201506","testDate":"2024-05-23","line":"","checkType":"process","variable":"Frame+Sus solder","variableDetail":"Frame+sus 5/23","variableGroup":"test","intervention":"","inputQty":140,"okQty":131,"ngTotal":9,"ngRate":6.43,"defectCategory":"assembly_defect","defectType":"Weak solder","defectCount":9},
            {"productType":"BRS-201506","testDate":"2024-05-18","line":"","checkType":"process","variable":"Frame+Sus solder","variableDetail":"Frame+sus 5/18","variableGroup":"test","intervention":"","inputQty":140,"okQty":138,"ngTotal":2,"ngRate":1.43,"defectCategory":"assembly_defect","defectType":"Weak solder","defectCount":2},
            {"productType":"BRS-201506","testDate":"2024-05-24","line":"","checkType":"process","variable":"Frame+Sus solder","variableDetail":"Frame+sus 5/24","variableGroup":"test","intervention":"","inputQty":264,"okQty":264,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
            {"productType":"BRS-201506","testDate":"2024-05-25","line":"4B","checkType":"process","variable":"Frame+Sus solder","variableDetail":"Frame Sus line 4B","variableGroup":"test","intervention":"","inputQty":120,"okQty":95,"ngTotal":25,"ngRate":20.83,"defectCategory":"assembly_defect","defectType":"Weak solder","defectCount":25},
            # 02/Feb Dome process
            {"productType":"BRS-201506","testDate":"2024-02-02","line":"","checkType":"process","variable":"Dome vision","variableDetail":"Dome process","variableGroup":"test","intervention":"","inputQty":6638,"okQty":6615,"ngTotal":23,"ngRate":0.3,"defectCategory":"assembly_defect","defectType":"Dome offset","defectCount":12},
            {"productType":"BRS-201506","testDate":"2024-02-02","line":"","checkType":"process","variable":"Dome vision","variableDetail":"Dome process","variableGroup":"test","intervention":"","inputQty":6638,"okQty":6615,"ngTotal":23,"ngRate":0.3,"defectCategory":"assembly_defect","defectType":"Dome damage","defectCount":1},
            {"productType":"BRS-201506","testDate":"2024-02-02","line":"","checkType":"process","variable":"Dome vision","variableDetail":"Dome process","variableGroup":"test","intervention":"","inputQty":6638,"okQty":6615,"ngTotal":23,"ngRate":0.3,"defectCategory":"assembly_defect","defectType":"Glue discontinuous","defectCount":10},
            # DOE Z-axis vs Delay (Dome don't pick-up)
            {"productType":"BRS-201506","testDate":"2024-02-18","line":"Normal","checkType":"process","variable":"Dome pick-up DOE","variableDetail":"Z=19.050 Delay=0.5","variableGroup":"test","intervention":"","inputQty":200,"okQty":190,"ngTotal":10,"ngRate":5.0,"defectCategory":"assembly_defect","defectType":"Don't pick-up","defectCount":10},
            {"productType":"BRS-201506","testDate":"2024-02-18","line":"Normal","checkType":"process","variable":"Dome pick-up DOE","variableDetail":"Z=19.150 Delay=0.5","variableGroup":"test","intervention":"","inputQty":360,"okQty":358,"ngTotal":2,"ngRate":0.6,"defectCategory":"assembly_defect","defectType":"Don't pick-up","defectCount":2},
            {"productType":"BRS-201506","testDate":"2024-02-18","line":"Normal","checkType":"process","variable":"Dome pick-up DOE","variableDetail":"Z=19.250 Delay=0.5","variableGroup":"test","intervention":"","inputQty":120,"okQty":119,"ngTotal":1,"ngRate":0.8,"defectCategory":"assembly_defect","defectType":"Don't pick-up","defectCount":1},
            {"productType":"BRS-201506","testDate":"2024-02-18","line":"Normal","checkType":"process","variable":"Dome pick-up DOE","variableDetail":"Z=19.250 Delay=0.8","variableGroup":"test","intervention":"","inputQty":272,"okQty":269,"ngTotal":3,"ngRate":1.1,"defectCategory":"assembly_defect","defectType":"Don't pick-up","defectCount":3},
            {"productType":"BRS-201506","testDate":"2024-02-18","line":"Normal","checkType":"process","variable":"Dome pick-up DOE","variableDetail":"Z=19.350 Delay=0.8","variableGroup":"test","intervention":"Z=19.350 Delay=0.8","inputQty":630,"okQty":628,"ngTotal":2,"ngRate":0.3,"defectCategory":"assembly_defect","defectType":"Don't pick-up","defectCount":2},
        ],
        "tags":["brs-201506","ng-process-high","weak-solder","dome-offset","z-axis-doe","frame-sus","tip-machine"],
        "reportType":"doe_factorial",
        "verdict":"improved",
        "headline":"DOE Z=19.350 + Delay=0.8 reduces Dome pick-up NG to 0.3% (vs 5.0% at Z=19.050)",
        "evidence":[
            {"metric":"Dome don't pick-up — best cell","baselineLabel":"Spec","baselineValue":"existing Z=19.050/0.5","variantLabel":"Test cell","variantValue":"0.3% (2/630) at Z=19.350/0.8","deltaText":"-4.7pp","deltaSign":"down","note":"best cell","comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Dome don't pick-up — worst cell","baselineLabel":"Spec","baselineValue":"target near-zero","variantLabel":"Test cell","variantValue":"5.0% (10/200) at Z=19.050/0.5","deltaText":"+5.0pp","deltaSign":"up","note":"worst cell","comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Set Dome pick-up to Z-axis 19.350mm with Delay 0.8s"},
            {"priority":2,"kind":"action","text":"Audit Frame+Sus solder line 4B — Weak solder 20.83%"},
            {"priority":3,"kind":"investigate","text":"Investigate AOI Dome vision 52.2% offset / 43.5% glue-discontinuous"},
        ],
        "context":{"process":"Solder vision (AWF M/C#1/#2) + Dome vision + Dome assy DOE","stage":"BRS-201506 sub1 multi-day audit","baselineReason":"DOE 5x2 grid of Z-axis × Delay; no single Normal baseline"},
        "doeGrid":{
            "factor1Name":"Z-axis (mm)","factor2Name":"Delay (s)",
            "factor1Levels":["19.050","19.150","19.250","19.350"],
            "factor2Levels":["0.5","0.8"],
            "cells":[
                {"f1":"19.050","f2":"0.5","status":"ng","value":"5.0% (10/200)"},
                {"f1":"19.150","f2":"0.5","status":"ok","value":"0.6% (2/360)"},
                {"f1":"19.250","f2":"0.5","status":"ok","value":"0.8% (1/120)"},
                {"f1":"19.250","f2":"0.8","status":"borderline","value":"1.1% (3/272)"},
                {"f1":"19.350","f2":"0.8","status":"ok","value":"0.3% (2/630)"},
            ]
        },
        "trendPoints":None,
        **LEGACY
    },
    tr_ko={**LEGACY,
        "headline":"DOE Z=19.350 + Delay=0.8 적용 시 Dome pick-up NG 0.3% (Z=19.050 5.0% 대비)",
        "actions":[
            {"priority":1,"kind":"action","text":"Dome pick-up Z-축 19.350mm, Delay 0.8s로 설정"},
            {"priority":2,"kind":"action","text":"Frame+Sus 라인 4B 점검 — Weak solder 20.83%"},
            {"priority":3,"kind":"investigate","text":"AOI Dome vision offset 52.2% / glue-discontinuous 43.5% 원인 조사"},
        ],
        "context":{"process":"Solder vision (AWF M/C#1/#2) + Dome vision + Dome assy DOE","stage":"BRS-201506 sub1 다일 점검","baselineReason":"Z-축 × Delay 5x2 DOE; 단일 Normal 기준 없음"}
    },
    tr_vi={**LEGACY,
        "headline":"DOE Z=19.350 + Delay=0.8 giảm NG Dome pick-up xuống 0.3% (so 5.0% tại Z=19.050)",
        "actions":[
            {"priority":1,"kind":"action","text":"Đặt Dome pick-up Z-axis 19.350mm với Delay 0.8s"},
            {"priority":2,"kind":"action","text":"Kiểm tra line Frame+Sus 4B — Weak solder 20.83%"},
            {"priority":3,"kind":"investigate","text":"Điều tra AOI Dome vision offset 52.2% / glue-discontinuous 43.5%"},
        ],
        "context":{"process":"Solder vision (AWF M/C#1/#2) + Dome vision + DOE Dome assy","stage":"BRS-201506 sub1 audit nhiều ngày","baselineReason":"Lưới DOE 5x2 Z-axis × Delay; không có Normal duy nhất"}
    },
))

# ============================================================================
# 10. 12 . BRS-201506 DT Report Picture Spot Welding - 2025.01.17
# ============================================================================
DATASETS.append(dict(
    name="12 . BRS-201506 DT Report Picture Spot Welding  Date 17.1.2025",
    product="BRS-201506 DT",
    result={
        "measurements":[
            {"productType":"BRS-201506 DT","testDate":"2025-01-16","line":"","checkType":"visual_inspection","variable":"Spot welding picture catalog","variableDetail":"10 reference photos","variableGroup":"test","intervention":"","inputQty":10,"okQty":0,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        ],
        "tags":["brs-201506-dt","spot-welding","picture-reference","quality-log","visual-reference"],
        "reportType":"quality_log",
        "verdict":"",
        "headline":"Spot welding picture catalog — 10 reference photos, no measurement data",
        "evidence":[
            {"metric":"Reference photos","baselineLabel":"","baselineValue":"","variantLabel":"Spot welding pictures","variantValue":"10 photos","deltaText":"—","deltaSign":"no_change","note":"visual reference only","comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Use as visual reference standard for spot welding inspection"},
        ],
        "context":{"process":"Spot welding visual reference","stage":"BRS-201506 DT picture catalog","baselineReason":"no comparison — quality log only"},
        "doeGrid":None,"trendPoints":None,
        **LEGACY
    },
    tr_ko={**LEGACY,
        "headline":"Spot welding 사진 카탈로그 — 참조 사진 10장, 측정 데이터 없음",
        "actions":[
            {"priority":1,"kind":"action","text":"Spot welding 검사용 외관 참조 표준으로 활용"},
        ],
        "context":{"process":"Spot welding 외관 참조","stage":"BRS-201506 DT 사진 카탈로그","baselineReason":"비교 없음 — quality log 전용"}
    },
    tr_vi={**LEGACY,
        "headline":"Danh mục ảnh spot welding — 10 ảnh tham chiếu, không có dữ liệu đo",
        "actions":[
            {"priority":1,"kind":"action","text":"Dùng làm chuẩn ảnh tham chiếu cho kiểm tra spot welding"},
        ],
        "context":{"process":"Tham chiếu hình ảnh spot welding","stage":"Danh mục ảnh BRS-201506 DT","baselineReason":"không so sánh — chỉ quality log"}
    },
))

# ============================================================================
# 11. 12. BRS-161014 GMI Report test VP mold #8 old lot - 2024.10.10
# ============================================================================
DATASETS.append(dict(
    name="12. BRS-161014  GMI Report test VP mold #8 old lot NG high already return",
    product="BRS-161014",
    result={
        "measurements":[
            # Sub1 VP bending E2-3B
            {"productType":"BRS-161014","testDate":"2024-10-09","line":"E2-3B","checkType":"process","variable":"VP bending sub1","variableDetail":"VP #8 lot 240809017","variableGroup":"test","intervention":"VP #8","inputQty":600,"okQty":412,"ngTotal":188,"ngRate":31.3,"defectCategory":"assembly_defect","defectType":"VP Bending","defectCount":184},
            {"productType":"BRS-161014","testDate":"2024-10-09","line":"E2-3B","checkType":"process","variable":"VP bending sub1","variableDetail":"VP #8 lot 240809017","variableGroup":"test","intervention":"VP #8","inputQty":600,"okQty":412,"ngTotal":188,"ngRate":31.3,"defectCategory":"assembly_defect","defectType":"VP+CD separate","defectCount":2},
            {"productType":"BRS-161014","testDate":"2024-10-09","line":"E2-3B","checkType":"process","variable":"VP bending sub1","variableDetail":"VP #8 lot 240814027","variableGroup":"test","intervention":"VP #8","inputQty":600,"okQty":321,"ngTotal":279,"ngRate":46.5,"defectCategory":"assembly_defect","defectType":"VP Bending","defectCount":279},
            {"productType":"BRS-161014","testDate":"2024-10-09","line":"E2-3B","checkType":"process","variable":"VP bending sub1","variableDetail":"VP #8 lot 240816006","variableGroup":"test","intervention":"VP #8","inputQty":600,"okQty":376,"ngTotal":224,"ngRate":37.3,"defectCategory":"assembly_defect","defectType":"VP Bending","defectCount":224},
            {"productType":"BRS-161014","testDate":"2024-10-09","line":"E2-3B","checkType":"process","variable":"VP bending sub1","variableDetail":"VP #8 lot 240821007","variableGroup":"test","intervention":"VP #8","inputQty":600,"okQty":535,"ngTotal":65,"ngRate":10.8,"defectCategory":"assembly_defect","defectType":"VP Bending","defectCount":61},
            {"productType":"BRS-161014","testDate":"2024-10-09","line":"E2-3B","checkType":"process","variable":"VP bending sub1","variableDetail":"Normal VP #12 lot 240927005","variableGroup":"normal","intervention":"VP #12","inputQty":2300,"okQty":2293,"ngTotal":7,"ngRate":0.3,"defectCategory":"assembly_defect","defectType":"VP Bending","defectCount":7},
            # Sub1 C2-3B
            {"productType":"BRS-161014","testDate":"2024-10-09","line":"C2-3B","checkType":"process","variable":"VP bending sub1","variableDetail":"VP #8 lot 240809017","variableGroup":"test","intervention":"VP #8","inputQty":600,"okQty":374,"ngTotal":226,"ngRate":37.7,"defectCategory":"assembly_defect","defectType":"VP Bending","defectCount":223},
            {"productType":"BRS-161014","testDate":"2024-10-09","line":"C2-3B","checkType":"process","variable":"VP bending sub1","variableDetail":"VP #8 lot 240816006","variableGroup":"test","intervention":"VP #8","inputQty":600,"okQty":340,"ngTotal":260,"ngRate":43.3,"defectCategory":"assembly_defect","defectType":"VP Bending","defectCount":260},
            {"productType":"BRS-161014","testDate":"2024-10-09","line":"C2-3B","checkType":"process","variable":"VP bending sub1","variableDetail":"VP #8 lot 240821007","variableGroup":"test","intervention":"VP #8","inputQty":600,"okQty":408,"ngTotal":192,"ngRate":32.0,"defectCategory":"assembly_defect","defectType":"VP Bending","defectCount":192},
            {"productType":"BRS-161014","testDate":"2024-10-09","line":"C2-3B","checkType":"process","variable":"VP bending sub1","variableDetail":"VP #8 lot 240814027","variableGroup":"test","intervention":"VP #8","inputQty":600,"okQty":284,"ngTotal":316,"ngRate":52.7,"defectCategory":"assembly_defect","defectType":"VP Bending","defectCount":313},
            {"productType":"BRS-161014","testDate":"2024-10-09","line":"C2-3B","checkType":"process","variable":"VP bending sub1","variableDetail":"Normal VP #4 lot 240927005","variableGroup":"normal","intervention":"VP #4","inputQty":600,"okQty":593,"ngTotal":7,"ngRate":1.2,"defectCategory":"assembly_defect","defectType":"VP Bending","defectCount":7},
            # Function E2-3B
            {"productType":"BRS-161014","testDate":"2024-10-10","line":"E2-3B","checkType":"function","variable":"Function","variableDetail":"VP #8 lot 240809017","variableGroup":"test","intervention":"VP #8","inputQty":401,"okQty":385,"ngTotal":16,"ngRate":4.0,"defectCategory":"function_hearing","defectType":"Noise","defectCount":10},
            {"productType":"BRS-161014","testDate":"2024-10-10","line":"E2-3B","checkType":"function","variable":"Function","variableDetail":"VP #8 lot 240814027","variableGroup":"test","intervention":"VP #8","inputQty":321,"okQty":310,"ngTotal":11,"ngRate":3.4,"defectCategory":"function_hearing","defectType":"Touch","defectCount":7},
            {"productType":"BRS-161014","testDate":"2024-10-10","line":"E2-3B","checkType":"function","variable":"Function","variableDetail":"VP #8 lot 240816006","variableGroup":"test","intervention":"VP #8","inputQty":373,"okQty":361,"ngTotal":12,"ngRate":3.2,"defectCategory":"function_hearing","defectType":"Noise","defectCount":11},
            {"productType":"BRS-161014","testDate":"2024-10-10","line":"E2-3B","checkType":"function","variable":"Function","variableDetail":"VP #8 lot 240821007","variableGroup":"test","intervention":"VP #8","inputQty":532,"okQty":523,"ngTotal":9,"ngRate":1.7,"defectCategory":"function_hearing","defectType":"Noise","defectCount":5},
            {"productType":"BRS-161014","testDate":"2024-10-10","line":"E2-3B","checkType":"function","variable":"Function","variableDetail":"Normal VP #12 lot 240927005","variableGroup":"normal","intervention":"VP #12","inputQty":799,"okQty":779,"ngTotal":20,"ngRate":2.5,"defectCategory":"function_hearing","defectType":"Noise","defectCount":15},
            # Function C2-3A/B
            {"productType":"BRS-161014","testDate":"2024-10-10","line":"C2-3A","checkType":"function","variable":"Function","variableDetail":"VP #8 lot 240809017","variableGroup":"test","intervention":"VP #8","inputQty":376,"okQty":367,"ngTotal":9,"ngRate":2.4,"defectCategory":"function_hearing","defectType":"Noise","defectCount":7},
            {"productType":"BRS-161014","testDate":"2024-10-10","line":"C2-3A","checkType":"function","variable":"Function","variableDetail":"VP #8 lot 240816006","variableGroup":"test","intervention":"VP #8","inputQty":336,"okQty":324,"ngTotal":12,"ngRate":3.6,"defectCategory":"function_hearing","defectType":"Noise","defectCount":10},
            {"productType":"BRS-161014","testDate":"2024-10-10","line":"C2-3B","checkType":"function","variable":"Function","variableDetail":"VP #8 lot 240821007","variableGroup":"test","intervention":"VP #8","inputQty":401,"okQty":380,"ngTotal":21,"ngRate":5.2,"defectCategory":"function_hearing","defectType":"Noise","defectCount":10},
            {"productType":"BRS-161014","testDate":"2024-10-10","line":"C2-3B","checkType":"function","variable":"Function","variableDetail":"VP #8 lot 240814027","variableGroup":"test","intervention":"VP #8","inputQty":279,"okQty":274,"ngTotal":5,"ngRate":1.8,"defectCategory":"function_hearing","defectType":"Noise","defectCount":3},
            {"productType":"BRS-161014","testDate":"2024-10-10","line":"C2-3A","checkType":"function","variable":"Function","variableDetail":"Normal VP #7 lot 240927005","variableGroup":"normal","intervention":"VP #7","inputQty":429,"okQty":418,"ngTotal":11,"ngRate":2.6,"defectCategory":"function_hearing","defectType":"Noise","defectCount":7},
        ],
        "tags":["brs-161014","vp-mold-8","vp-bending","old-lot","laser-cutting","function-hearing","lot-screening"],
        "reportType":"multi_arm",
        "verdict":"partial",
        "headline":"VP #8 sub1 bending 10.8-52.7% vs Normal 0.3-1.2%; function NG split — 2 lots usable, 2 not",
        "evidence":[
            {"metric":"VP bending rate sub1 (all lines)","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"+52.4pp range","deltaSign":"up","note":"all VP#8 lots far above normal",
             "comparisons":[
                {"label":"Normal VP #12 lot 240927005","value":"0.3% (7/2300)","n":2300,"isBaseline":True,"isBest":True,"isWorst":False},
                {"label":"Normal VP #4 lot 240927005","value":"1.2% (7/600)","n":600,"isBaseline":True,"isBest":False,"isWorst":False},
                {"label":"VP #8 lot 240821007 E2","value":"10.8% (65/600)","n":600,"isBaseline":False,"isBest":False,"isWorst":False},
                {"label":"VP #8 lot 240809017 E2","value":"31.3% (188/600)","n":600,"isBaseline":False,"isBest":False,"isWorst":False},
                {"label":"VP #8 lot 240821007 C2","value":"32.0% (192/600)","n":600,"isBaseline":False,"isBest":False,"isWorst":False},
                {"label":"VP #8 lot 240816006 E2","value":"37.3% (224/600)","n":600,"isBaseline":False,"isBest":False,"isWorst":False},
                {"label":"VP #8 lot 240809017 C2","value":"37.7% (226/600)","n":600,"isBaseline":False,"isBest":False,"isWorst":False},
                {"label":"VP #8 lot 240814027 C2","value":"52.7% (316/600)","n":600,"isBaseline":False,"isBest":False,"isWorst":True},
             ],
             "bestLabel":"Normal VP #12 lot 240927005","worstLabel":"VP #8 lot 240814027 C2"},
            {"metric":"Function NG by lot (E2/C2 combined)","baselineLabel":"Normal","baselineValue":"2.5-2.6% (VP#12/#7)","variantLabel":"VP #8 lots","variantValue":"1.7-5.2%","deltaText":"—","deltaSign":"up","note":"return all bending lots; 2/4 function-OK","comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Return all VP #8 old lots to WH — bending sub1 30-52% unacceptable"},
            {"priority":2,"kind":"action","text":"Allow C2 lots 240809017 (2.4%) and 240814027 (1.8%) for limited use"},
            {"priority":3,"kind":"action","text":"Reject E2 lots 240809017 (4.0%) and 240814027 (3.4%) — function above Normal"},
        ],
        "context":{"process":"Laser cutting VP → sub1 bending + function","stage":"E2-3B / C2-3A,3B production lines","baselineReason":"Normal VP #12/#7 same-week available as paired arm"},
        "doeGrid":None,"trendPoints":None,
        **LEGACY
    },
    tr_ko={**LEGACY,
        "headline":"VP #8 sub1 bending 10.8-52.7% vs Normal 0.3-1.2%; function NG는 lot별 — 2 lot 사용 가능, 2 lot 불가",
        "actions":[
            {"priority":1,"kind":"action","text":"VP #8 old lot 전체 창고 반품 — bending sub1 30-52% 불합격"},
            {"priority":2,"kind":"action","text":"C2 lot 240809017 (2.4%), 240814027 (1.8%) 한정 사용 허용"},
            {"priority":3,"kind":"action","text":"E2 lot 240809017 (4.0%), 240814027 (3.4%) Function이 Normal 초과 — 사용 불가"},
        ],
        "context":{"process":"레이저 절단 VP → sub1 bending + function","stage":"E2-3B / C2-3A,3B 양산 라인","baselineReason":"동일 주차 Normal VP #12/#7가 페어 arm으로 존재"}
    },
    tr_vi={**LEGACY,
        "headline":"VP #8 sub1 bending 10.8-52.7% vs Normal 0.3-1.2%; function NG theo lot — 2 lot dùng được, 2 lot không",
        "actions":[
            {"priority":1,"kind":"action","text":"Trả về kho toàn bộ lot cũ VP #8 — bending sub1 30-52% không đạt"},
            {"priority":2,"kind":"action","text":"Cho phép dùng hạn chế lot C2 240809017 (2.4%) và 240814027 (1.8%)"},
            {"priority":3,"kind":"action","text":"Từ chối lot E2 240809017 (4.0%) và 240814027 (3.4%) — function vượt Normal"},
        ],
        "context":{"process":"Cắt laser VP → sub1 bending + function","stage":"Line sản xuất E2-3B / C2-3A,3B","baselineReason":"Normal VP #12/#7 cùng tuần làm arm so sánh"}
    },
))

# ============================================================================
# 12. 12. BRS-161014 Report checking problem sub1 TF line  (big trend dataset)
# ============================================================================
DATASETS.append(dict(
    name="12. BRS-161014  Report checking problem sub1 TF line",
    product="BRS-161014",
    result={
        "measurements":[
            # Aggregate per problem (using totals would be ideal; we capture worst-process rows + key trend points)
            # AOI detect NG not exactly — recurring ~30-50%
            {"productType":"BRS-161014","testDate":"2025-05-13","line":"sub1 TF","checkType":"process","variable":"AOI check VP","variableDetail":"AOI detect NG not exactly","variableGroup":"test","intervention":"","inputQty":10,"okQty":5,"ngTotal":5,"ngRate":50.0,"defectCategory":"assembly_defect","defectType":"AOI miss-detect","defectCount":5},
            # Dome array NG float — recurring 62.1%
            {"productType":"BRS-161014","testDate":"2025-05-13","line":"sub1 TF","checkType":"process","variable":"Dome array","variableDetail":"CD load tray NG float when up locate","variableGroup":"test","intervention":"","inputQty":29,"okQty":11,"ngTotal":18,"ngRate":62.1,"defectCategory":"assembly_defect","defectType":"Dome float","defectCount":18},
            # VP+CD separate — major problem (use representative latest)
            {"productType":"BRS-161014","testDate":"2025-05-13","line":"sub1 TF","checkType":"process","variable":"Bonding VP+CD","variableDetail":"NG separate VP+CD","variableGroup":"test","intervention":"","inputQty":2520,"okQty":2438,"ngTotal":82,"ngRate":3.3,"defectCategory":"assembly_defect","defectType":"VP+CD separate","defectCount":82},
            # AOI cannot detect separate (recurring)
            {"productType":"BRS-161014","testDate":"2025-05-13","line":"sub1 TF","checkType":"process","variable":"AOI cannot detect","variableDetail":"Cannot detect NG separate VP+CD","variableGroup":"test","intervention":"","inputQty":100,"okQty":96,"ngTotal":4,"ngRate":4.0,"defectCategory":"assembly_defect","defectType":"AOI miss-detect","defectCount":4},
            # Function NG hearing aggregate (latest 5/13)
            {"productType":"BRS-161014","testDate":"2025-05-13","line":"sub1 TF","checkType":"function","variable":"Hearing","variableDetail":"NG rate of hearing — 5/13","variableGroup":"test","intervention":"","inputQty":5630,"okQty":5417,"ngTotal":213,"ngRate":3.8,"defectCategory":"function_hearing","defectType":"Hearing","defectCount":213},
            # Function NG comparison trend (sample dates) — VP TF line vs Normal
            {"productType":"BRS-161014","testDate":"2025-03-05","line":"E2-3A","checkType":"function","variable":"VP TF line","variableDetail":"VP TF line","variableGroup":"test","intervention":"","inputQty":3701,"okQty":3288,"ngTotal":413,"ngRate":11.2,"defectCategory":"function_hearing","defectType":"Noise","defectCount":401},
            {"productType":"BRS-161014","testDate":"2025-03-05","line":"E2-3A","checkType":"function","variable":"VP TF line","variableDetail":"Normal line","variableGroup":"normal","intervention":"","inputQty":7157,"okQty":6899,"ngTotal":258,"ngRate":3.6,"defectCategory":"function_hearing","defectType":"Noise","defectCount":247},
            {"productType":"BRS-161014","testDate":"2025-04-26","line":"E2-3A","checkType":"function","variable":"VP TF line","variableDetail":"VP TF line (After UC press more)","variableGroup":"after","intervention":"UC press more","inputQty":218,"okQty":208,"ngTotal":10,"ngRate":4.6,"defectCategory":"function_hearing","defectType":"Noise","defectCount":8},
            {"productType":"BRS-161014","testDate":"2025-04-26","line":"E2-3A","checkType":"function","variable":"VP TF line","variableDetail":"Normal line","variableGroup":"normal","intervention":"","inputQty":800,"okQty":783,"ngTotal":17,"ngRate":2.1,"defectCategory":"function_hearing","defectType":"Noise","defectCount":14},
        ],
        "tags":["brs-161014","sub1-tf-line","trend-analysis","vp-cd-separate","aoi-miss-detect","dome-float","function-hearing","weekly-rate"],
        "reportType":"trend_analysis",
        "verdict":"partial",
        "headline":"Sub1 TF line VP TF hearing trended 11.2% (3/5) → 4.6% (4/26) after UC press more vs Normal 2.1-3.6%",
        "evidence":[
            {"metric":"VP TF line function NG (weekly avg)","baselineLabel":"Normal line","baselineValue":"3.1-3.6% range","variantLabel":"VP TF line","variantValue":"3.2-11.2% range","deltaText":"-6.6pp peak-to-recent","deltaSign":"down","note":"convergence after UC press change","comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"AOI detect NG miss / Dome float","baselineLabel":"Spec","baselineValue":"near-zero","variantLabel":"Sub1 TF","variantValue":"30-62% chronic","deltaText":"+62pp","deltaSign":"up","note":"persistent process problem","comparisons":None,"bestLabel":"","worstLabel":""},
        ],
        "actions":[
            {"priority":1,"kind":"action","text":"Lock in UC press more setting — peak hearing NG down 11.2% to 4.6%"},
            {"priority":2,"kind":"investigate","text":"Address AOI miss-detect 30-62% and Dome float 62% chronic NG"},
            {"priority":3,"kind":"action","text":"Bonding tuning to keep VP+CD separate below 2% weekly"},
        ],
        "context":{"process":"Sub1 TF line: AOI / Dome / Bonding / VP+CD / Function","stage":"Multi-week audit E2-3A/3B March-May 2025","baselineReason":"daily Normal line paired vs VP TF line per shift"},
        "doeGrid":None,
        "trendPoints":[
            {"label":"Wk 3/5","value":"11.2%","note":"VP TF line peak"},
            {"label":"Wk 3/8","value":"3.7%","note":""},
            {"label":"Wk 3/14","value":"5.1%","note":""},
            {"label":"Wk 3/26","value":"5.6%","note":""},
            {"label":"Wk 4/19","value":"2.2%","note":""},
            {"label":"Wk 4/26","value":"4.6%","note":"After UC press more"},
            {"label":"Wk 4/29","value":"5.0%","note":""},
            {"label":"Wk 5/7","value":"4.3%","note":""},
            {"label":"Wk 5/13","value":"3.8%","note":"Hearing NG aggregate"},
        ],
        **LEGACY
    },
    tr_ko={**LEGACY,
        "headline":"Sub1 TF 라인 VP TF hearing 11.2% (3/5) → UC press 강화 후 4.6% (4/26), Normal 2.1-3.6%",
        "actions":[
            {"priority":1,"kind":"action","text":"UC press 강화 조건 표준화 — 최고 hearing NG 11.2%에서 4.6%로 감소"},
            {"priority":2,"kind":"investigate","text":"AOI 미검출 30-62%, Dome float 62% 만성 NG 대책"},
            {"priority":3,"kind":"action","text":"VP+CD 분리율 주간 2% 미만 유지하도록 본딩 튜닝"},
        ],
        "context":{"process":"Sub1 TF 라인: AOI / Dome / Bonding / VP+CD / Function","stage":"2025년 3-5월 E2-3A/3B 다주 점검","baselineReason":"일별 Normal 라인과 시프트별 VP TF 라인 페어링"}
    },
    tr_vi={**LEGACY,
        "headline":"Line Sub1 TF VP TF hearing 11.2% (3/5) → 4.6% (4/26) sau UC press more, Normal 2.1-3.6%",
        "actions":[
            {"priority":1,"kind":"action","text":"Chốt cài đặt UC press more — NG hearing đỉnh từ 11.2% xuống 4.6%"},
            {"priority":2,"kind":"investigate","text":"Xử lý NG kinh niên: AOI miss-detect 30-62% và Dome float 62%"},
            {"priority":3,"kind":"action","text":"Tinh chỉnh bonding để giữ VP+CD separate dưới 2% mỗi tuần"},
        ],
        "context":{"process":"Line Sub1 TF: AOI / Dome / Bonding / VP+CD / Function","stage":"Audit nhiều tuần E2-3A/3B tháng 3-5/2025","baselineReason":"line Normal cùng ngày ghép với VP TF line theo ca"}
    },
))

# ---------------------------------------------------------------------------

def commit_one(con, name, product, result, tr_ko, tr_vi):
    cur = con.cursor()
    cur.execute("BEGIN")
    try:
        now = datetime.utcnow().isoformat() + "Z"
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
        return True, None
    except Exception as e:
        con.rollback()
        return False, str(e)


def main():
    con = sqlite3.connect(DB)
    con.execute("PRAGMA busy_timeout = 30000")
    ok = skip = fail = 0
    for d in DATASETS:
        success, err = commit_one(con, d["name"], d["product"], d["result"], d["tr_ko"], d["tr_vi"])
        if success:
            ok += 1
            print(f"[OK] {d['name']}")
        else:
            fail += 1
            print(f"[PARSE-FAIL] {d['name']}: {err}")
    con.close()
    print(f"\n=== BATCH DONE ===\nMode: Reanalyze\nProcessed: {ok}\nSkipped: {skip}\nParse-fail: {fail}")

if __name__ == "__main__":
    main()
