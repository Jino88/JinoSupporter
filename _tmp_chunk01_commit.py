"""Commit normalized + translated JSON for each dataset in chunk_01."""
import sqlite3
import json
import sys
import os
from datetime import datetime

DB = r"D:\000. MyWorks\002. DB\process-review.db"
sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
# DATASET RESULTS — produced by the agent (me) by applying §3
# rules to each dataset's excel_paste TSV.
# Each entry: { name, en, ko, vi }
# en is the v7 NormalizeFromText output, ko/vi are TranslateAnalysis outputs.
# ============================================================

RESULTS = []

# ----------- DS1: BRS-201506 New Bond G06-0003 -------------
RESULTS.append({
"name": "1. BRS-2015 Report Test New Bond G06-0003 Date 3.5.2024",
"en": {
  "measurements": [
    {"productType":"BRS-201506","testDate":"2024-05-02","line":"E2-4A","checkType":"visual_inspection","variable":"SM VISION","variableDetail":"","variableGroup":"test","intervention":"New Bond G06-0003","inputQty":2540,"okQty":2537,"ngTotal":3,"ngRate":0.1,"defectCategory":"cosmetic_defect","defectType":"","defectCount":0},
    {"productType":"BRS-201506","testDate":"2024-05-02","line":"E2-4A","checkType":"visual_inspection","variable":"SM VISION","variableDetail":"","variableGroup":"test","intervention":"New Bond G06-0003","inputQty":2540,"okQty":2537,"ngTotal":3,"ngRate":0.1,"defectCategory":"cosmetic_defect","defectType":"SM damage","defectCount":2},
    {"productType":"BRS-201506","testDate":"2024-05-02","line":"E2-4A","checkType":"visual_inspection","variable":"SM VISION","variableDetail":"","variableGroup":"test","intervention":"New Bond G06-0003","inputQty":2540,"okQty":2537,"ngTotal":3,"ngRate":0.1,"defectCategory":"cosmetic_defect","defectType":"B-PT damage","defectCount":1},
    {"productType":"BRS-201506","testDate":"2024-05-02","line":"E2-4A","checkType":"visual_inspection","variable":"SM VISION","variableDetail":"","variableGroup":"normal","intervention":"Normal Bond G06-0002","inputQty":3200,"okQty":3194,"ngTotal":6,"ngRate":0.2,"defectCategory":"cosmetic_defect","defectType":"SM damage","defectCount":6},
    {"productType":"BRS-201506","testDate":"2024-05-02","line":"E2-4A","checkType":"visual_inspection","variable":"YOKE VISION","variableDetail":"","variableGroup":"test","intervention":"New Bond G06-0003","inputQty":2513,"okQty":2511,"ngTotal":2,"ngRate":0.1,"defectCategory":"cosmetic_defect","defectType":"PT Gap","defectCount":2},
    {"productType":"BRS-201506","testDate":"2024-05-02","line":"E2-4A","checkType":"visual_inspection","variable":"YOKE VISION","variableDetail":"","variableGroup":"normal","intervention":"Normal Bond G06-0002","inputQty":1720,"okQty":1719,"ngTotal":1,"ngRate":0.1,"defectCategory":"cosmetic_defect","defectType":"PT Gap","defectCount":1},
    {"productType":"BRS-201506","testDate":"2024-05-04","line":"E2-4A","checkType":"function","variable":"Function","variableDetail":"","variableGroup":"test","intervention":"New Bond G06-0003","inputQty":2033,"okQty":1944,"ngTotal":89,"ngRate":4.4,"defectCategory":"function_hearing","defectType":"Noise","defectCount":24},
    {"productType":"BRS-201506","testDate":"2024-05-04","line":"E2-4A","checkType":"function","variable":"Function","variableDetail":"","variableGroup":"test","intervention":"New Bond G06-0003","inputQty":2033,"okQty":1944,"ngTotal":89,"ngRate":4.4,"defectCategory":"function_hearing","defectType":"Touch","defectCount":65},
    {"productType":"BRS-201506","testDate":"2024-05-04","line":"E2-4A","checkType":"function","variable":"Function","variableDetail":"","variableGroup":"normal","intervention":"Normal Bond G06-0002","inputQty":2507,"okQty":2395,"ngTotal":112,"ngRate":4.5,"defectCategory":"function_spl","defectType":"SPL+THD","defectCount":1},
    {"productType":"BRS-201506","testDate":"2024-05-04","line":"E2-4A","checkType":"function","variable":"Function","variableDetail":"","variableGroup":"normal","intervention":"Normal Bond G06-0002","inputQty":2507,"okQty":2395,"ngTotal":112,"ngRate":4.5,"defectCategory":"function_hearing","defectType":"Noise","defectCount":34},
    {"productType":"BRS-201506","testDate":"2024-05-04","line":"E2-4A","checkType":"function","variable":"Function","variableDetail":"","variableGroup":"normal","intervention":"Normal Bond G06-0002","inputQty":2507,"okQty":2395,"ngTotal":112,"ngRate":4.5,"defectCategory":"function_hearing","defectType":"Touch","defectCount":77}
  ],
  "tags":["brs-201506","new-bond-g06-0003","bond-validation","vision-check","function-test","tension-test","drop-test","comparison"],
  "reportType":"comparison_study",
  "verdict":"no_clear_effect",
  "headline":"New Bond G06-0003 — function NG 4.4% vs Normal 4.5% (no change, OK to use)",
  "evidence":[
    {"metric":"Function NG rate","baselineLabel":"Normal Bond G06-0002","baselineValue":"4.5% (112/2507)","variantLabel":"New Bond G06-0003","variantValue":"4.4% (89/2033)","deltaText":"-0.1pp","deltaSign":"no_change","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"SM Vision NG","baselineLabel":"Normal","baselineValue":"0.2% (6/3200)","variantLabel":"New Bond","variantValue":"0.1% (3/2540)","deltaText":"-0.1pp","deltaSign":"down","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"Tension Long S-MG AVG","baselineLabel":"Normal","baselineValue":"50.7","variantLabel":"New Bond","variantValue":"55.0","deltaText":"+4.3","deltaSign":"up","note":"higher is better","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"Drop test NG","baselineLabel":"Normal","baselineValue":"0/8","variantLabel":"New Bond","variantValue":"0/8","deltaText":"+0pp","deltaSign":"no_change","note":"","comparisons":None,"bestLabel":"","worstLabel":""}
  ],
  "actions":[
    {"priority":1,"kind":"action","text":"Approve New Bond G06-0003 for production use"},
    {"priority":2,"kind":"investigate","text":"Confirm function NG breakdown across multiple lots"}
  ],
  "context":{"process":"Sub Yoke bond change (G06-0002 → G06-0003)","stage":"E2-4A sub line vision + tension + function","baselineReason":"same-event Normal Bond G06-0002 row present"},
  "doeGrid":None,"trendPoints":None,
  "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":""
},
"ko":{
  "headline":"신규 본드 G06-0003 — Function NG 4.4% vs Normal 4.5% (변화 없음, 사용 가능)",
  "actionTexts":["New Bond G06-0003 양산 적용 승인","복수 로트에 걸친 Function NG 세부 항목 확인"],
  "contextProcess":"Sub Yoke 본드 변경 (G06-0002 → G06-0003)","contextStage":"E2-4A 서브 라인 비전 + 인장 + Function","contextBaseline":"같은 이벤트에 Normal Bond G06-0002 행 존재"
},
"vi":{
  "headline":"Keo mới G06-0003 — NG chức năng 4.4% so với Normal 4.5% (không đổi, có thể sử dụng)",
  "actionTexts":["Phê duyệt sử dụng keo G06-0003 cho sản xuất","Kiểm tra phân loại NG chức năng trên nhiều lô"],
  "contextProcess":"Thay keo Sub Yoke (G06-0002 → G06-0003)","contextStage":"Vision + tension + chức năng line E2-4A sub","contextBaseline":"Có dòng Normal Bond G06-0002 cùng sự kiện"
}})

# ----------- DS2: JIG Frame Array A improve -------------
RESULTS.append({
"name": "1. BRS-201506 DT Report test JIG Frame Array A improve   date 6.12.2024",
"en":{
  "measurements":[
    {"productType":"BRS-201506DT","testDate":"2024-12-05","line":"C2-2A","checkType":"visual_inspection","variable":"Ass'y Coil+SP","variableDetail":"","variableGroup":"test","intervention":"JIG Frame Array A new version","inputQty":50,"okQty":50,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
    {"productType":"BRS-201506DT","testDate":"2024-12-05","line":"C2-2A","checkType":"visual_inspection","variable":"Ass'y Coil+SP","variableDetail":"","variableGroup":"normal","intervention":"Normal","inputQty":50,"okQty":50,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
    {"productType":"BRS-201506DT","testDate":"2024-12-05","line":"C2-2A","checkType":"visual_inspection","variable":"Spot Welding","variableDetail":"","variableGroup":"test","intervention":"JIG Frame Array A new version","inputQty":50,"okQty":50,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
    {"productType":"BRS-201506DT","testDate":"2024-12-05","line":"C2-2A","checkType":"visual_inspection","variable":"Spot Welding","variableDetail":"","variableGroup":"normal","intervention":"Old JIG","inputQty":20,"okQty":0,"ngTotal":20,"ngRate":100.0,"defectCategory":"assembly_defect","defectType":"Suspension damage","defectCount":20},
    {"productType":"BRS-161016","testDate":"2024-11-21","line":"C2-6B","checkType":"visual_inspection","variable":"Spot Welding","variableDetail":"After SUS material change","variableGroup":"test","intervention":"JIG material SUS","inputQty":60,"okQty":60,"ngTotal":6,"ngRate":10.0,"defectCategory":"assembly_defect","defectType":"Weak solder","defectCount":6},
    {"productType":"BRS-161016","testDate":"2024-11-21","line":"C2-6B","checkType":"visual_inspection","variable":"Spot Welding","variableDetail":"","variableGroup":"normal","intervention":"Normal","inputQty":2433,"okQty":2433,"ngTotal":13,"ngRate":0.5,"defectCategory":"assembly_defect","defectType":"Suspension damage","defectCount":12}
  ],
  "tags":["brs-201506dt","brs-161016","jig-frame-array","spot-welding","suspension-damage","sus-material","aluminum-replacement","jig-improvement"],
  "reportType":"comparison_study",
  "verdict":"improved",
  "headline":"JIG Frame Array A new version eliminates Suspension damage (100% → 0%, OK to use)",
  "evidence":[
    {"metric":"Spot weld NG (Old JIG vs new)","baselineLabel":"Old JIG (normal)","baselineValue":"100% (20/20)","variantLabel":"New JIG","variantValue":"0% (0/50)","deltaText":"-100pp","deltaSign":"down","note":"Suspension damage eliminated","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"Coil+SP NG","baselineLabel":"Normal","baselineValue":"0% (0/50)","variantLabel":"New JIG","variantValue":"0% (0/50)","deltaText":"+0pp","deltaSign":"no_change","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"SUS JIG spot weld NG","baselineLabel":"Normal","baselineValue":"0.5% (13/2433)","variantLabel":"SUS JIG","variantValue":"10.0% (6/60)","deltaText":"+9.5pp","deltaSign":"up","note":"weak solder due to JIG height gap","comparisons":None,"bestLabel":"","worstLabel":""}
  ],
  "actions":[
    {"priority":1,"kind":"action","text":"Adopt JIG Frame Array A new version for spot welding line"},
    {"priority":2,"kind":"action","text":"Change spot pad material from Aluminum to SUS to match GMI line"},
    {"priority":3,"kind":"investigate","text":"Adjust SUS JIG height to close 0.07mm gap causing weak solder"}
  ],
  "context":{"process":"Frame Array JIG redesign for spot welding","stage":"C2-2A / C2-6B sub-line spot welding","baselineReason":"old JIG / Normal lot directly compared on same shift"},
  "doeGrid":None,"trendPoints":None,
  "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":""
},
"ko":{"headline":"JIG Frame Array A 신규 버전, Suspension 손상 제거 (100% → 0%, 사용 승인)",
"actionTexts":["스폿 용접 라인에 JIG Frame Array A 신규 버전 적용","GMI 라인과 동일하게 스폿 패드 재질을 Aluminum → SUS로 변경","SUS JIG의 0.07mm 갭으로 인한 냉땜 해결을 위해 높이 조정 검토"],
"contextProcess":"스폿 용접용 Frame Array JIG 재설계","contextStage":"C2-2A / C2-6B 서브라인 스폿 용접","contextBaseline":"동일 시프트의 Old JIG / Normal 로트와 직접 비교"},
"vi":{"headline":"Phiên bản JIG Frame Array A mới loại bỏ hư hỏng Suspension (100% → 0%, có thể dùng)",
"actionTexts":["Áp dụng JIG Frame Array A phiên bản mới cho line hàn điểm","Đổi vật liệu pad hàn từ Nhôm sang SUS giống line GMI","Điều chỉnh chiều cao JIG SUS để đóng khe 0.07mm gây hàn yếu"],
"contextProcess":"Thiết kế lại JIG Frame Array cho hàn điểm","contextStage":"Line sub hàn điểm C2-2A / C2-6B","contextBaseline":"So sánh trực tiếp với lô JIG cũ / Normal cùng ca"}})

# ----------- DS3: VP die punch V1 -------------
RESULTS.append({
"name": "1. BRS-201506 Report test material VP using die punch cutting inside (V1 ) date 27.7.2024 -",
"en":{
  "measurements":[
    {"productType":"BRS-201506","testDate":"2024-08-31","line":"","checkType":"process","variable":"VP+CD Bonding","variableDetail":"1st batch","variableGroup":"test","intervention":"VP Die Punch V1 new machine","inputQty":500,"okQty":465,"ngTotal":35,"ngRate":7.0,"defectCategory":"assembly_defect","defectType":"NG bonding","defectCount":35},
    {"productType":"BRS-201506","testDate":"2024-08-31","line":"","checkType":"process","variable":"VP+CD Bonding","variableDetail":"","variableGroup":"normal","intervention":"VP normal","inputQty":500,"okQty":500,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
    {"productType":"BRS-201506","testDate":"2024-09-04","line":"","checkType":"process","variable":"VP+CD Bonding","variableDetail":"2nd batch","variableGroup":"test","intervention":"VP Die Punch V1 new machine","inputQty":1000,"okQty":997,"ngTotal":3,"ngRate":0.3,"defectCategory":"assembly_defect","defectType":"NG bonding","defectCount":3},
    {"productType":"BRS-201506","testDate":"2024-09-04","line":"","checkType":"process","variable":"VP+CD Vision","variableDetail":"2nd batch","variableGroup":"test","intervention":"VP Die Punch V1 new machine","inputQty":997,"okQty":997,"ngTotal":77,"ngRate":7.7,"defectCategory":"assembly_defect","defectType":"VP+CD Gap","defectCount":72},
    {"productType":"BRS-201506","testDate":"2024-09-04","line":"","checkType":"process","variable":"VP+CD Vision","variableDetail":"","variableGroup":"normal","intervention":"VP normal","inputQty":2002,"okQty":2002,"ngTotal":2,"ngRate":0.1,"defectCategory":"assembly_defect","defectType":"Not enough glue","defectCount":1},
    {"productType":"BRS-201506","testDate":"2024-09-07","line":"","checkType":"process","variable":"VP+CD Vision","variableDetail":"4th batch","variableGroup":"test","intervention":"VP Die Punch V1 new machine","inputQty":1000,"okQty":999,"ngTotal":103,"ngRate":10.3,"defectCategory":"assembly_defect","defectType":"VP+CD Gap","defectCount":102},
    {"productType":"BRS-201506","testDate":"2024-09-07","line":"","checkType":"process","variable":"VP+CD Vision","variableDetail":"","variableGroup":"normal","intervention":"VP normal","inputQty":1000,"okQty":1000,"ngTotal":2,"ngRate":0.2,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
    {"productType":"BRS-201506","testDate":"2024-09-10","line":"","checkType":"process","variable":"VP Cutting Offset","variableDetail":"","variableGroup":"test","intervention":"VP Die Punch V1 new machine","inputQty":1000,"okQty":953,"ngTotal":47,"ngRate":4.7,"defectCategory":"assembly_defect","defectType":"NG Offset","defectCount":47},
    {"productType":"BRS-201506","testDate":"2024-08-26","line":"","checkType":"function","variable":"Function","variableDetail":"Hearing up 10%","variableGroup":"test","intervention":"VP Die Punch V1","inputQty":2917,"okQty":2784,"ngTotal":193,"ngRate":6.6,"defectCategory":"function_hearing","defectType":"Touch","defectCount":152},
    {"productType":"BRS-201506","testDate":"2024-08-26","line":"","checkType":"function","variable":"Function","variableDetail":"","variableGroup":"normal","intervention":"VP normal","inputQty":1598,"okQty":1497,"ngTotal":141,"ngRate":8.8,"defectCategory":"function_hearing","defectType":"Touch","defectCount":97}
  ],
  "tags":["brs-201506","vp-die-punch-v1","new-machine","vp-cd-vision","vp-cd-gap","cutting-offset","function-test","comparison"],
  "reportType":"comparison_study",
  "verdict":"partial",
  "headline":"VP Die Punch V1 — function OK (6.6% vs 8.8%) but Vision Gap NG ~7-10% needs improvement",
  "evidence":[
    {"metric":"VP+CD Vision NG","baselineLabel":"VP normal","baselineValue":"0.1-0.2%","variantLabel":"V1 new machine","variantValue":"3.2-10.6%","deltaText":"+3-10pp","deltaSign":"up","note":"VP+CD Gap dominant","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"VP Cutting Offset NG","baselineLabel":"—","baselineValue":"—","variantLabel":"V1 new machine","variantValue":"4.7% (47/1000)","deltaText":"—","deltaSign":"no_change","note":"new metric","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"Function NG (Hearing up 10%)","baselineLabel":"VP normal","baselineValue":"8.8% (141/1598)","variantLabel":"V1","variantValue":"6.6% (193/2917)","deltaText":"-2.2pp","deltaSign":"down","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"Tension VP+CD AVG","baselineLabel":"VP normal","baselineValue":"1.762-2.670 kgf","variantLabel":"V1","variantValue":"2.155-2.272 kgf","deltaText":"+0.4","deltaSign":"up","note":"spec 1.2 kgf","comparisons":None,"bestLabel":"","worstLabel":""}
  ],
  "actions":[
    {"priority":1,"kind":"action","text":"Approve VP Die Punch V1 for trial use given function pass"},
    {"priority":2,"kind":"investigate","text":"Reduce VP+CD Gap at sub 1 — root cause die punch new machine"},
    {"priority":3,"kind":"risk","text":"Cutting offset 4.7% may grow under volume — monitor"}
  ],
  "context":{"process":"VP material with die punch cutting inside (V1)","stage":"VP+CD bonding / vision / function lines","baselineReason":"VP normal lot processed in parallel each day"},
  "doeGrid":None,"trendPoints":None,
  "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":""
},
"ko":{"headline":"VP Die Punch V1 — Function OK (6.6% vs 8.8%) 그러나 Vision Gap NG ~7-10% 개선 필요",
"actionTexts":["Function 합격 기반으로 VP Die Punch V1 시험 사용 승인","Sub 1 VP+CD Gap 감소 — 다이 펀치 신규 장비 원인 조사","컷팅 오프셋 4.7%가 양산 확대 시 증가 우려, 모니터링"],
"contextProcess":"VP 자재 다이 펀치 내측 컷팅 (V1)","contextStage":"VP+CD 본딩 / 비전 / Function 라인","contextBaseline":"매일 VP normal 로트가 병행 가공됨"},
"vi":{"headline":"VP Die Punch V1 — Chức năng OK (6.6% vs 8.8%) nhưng NG Gap Vision ~7-10% cần cải tiến",
"actionTexts":["Phê duyệt thử nghiệm VP Die Punch V1 do chức năng đạt","Giảm VP+CD Gap tại sub 1 — điều tra nguyên nhân máy die punch mới","Cutting offset 4.7% có thể tăng khi sản lượng lớn — theo dõi"],
"contextProcess":"Vật liệu VP cắt die punch bên trong (V1)","contextStage":"Line VP+CD bonding / vision / chức năng","contextBaseline":"Lô VP normal chạy song song mỗi ngày"}})

# ----------- DS4: VP improvement bending date 26.1 (3-arm) -------------
RESULTS.append({
"name": "1. BRS-201506DT Report test VP improverment bending date 26.1.2024",
"en":{
  "measurements":[
    {"productType":"BRS-201506DT","testDate":"2025-11-22","line":"C2-3A","checkType":"visual_inspection","variable":"VP Vision Sub 1","variableDetail":"","variableGroup":"new_lot","intervention":"VP new lot","inputQty":228,"okQty":228,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
    {"productType":"BRS-201506DT","testDate":"2025-11-22","line":"C2-3A","checkType":"visual_inspection","variable":"VP Vision Sub 1","variableDetail":"","variableGroup":"old_lot","intervention":"VP Old lot","inputQty":228,"okQty":226,"ngTotal":2,"ngRate":0.9,"defectCategory":"assembly_defect","defectType":"Particle","defectCount":1},
    {"productType":"BRS-201506DT","testDate":"2025-11-22","line":"C2-3A","checkType":"visual_inspection","variable":"VP Vision Sub 1","variableDetail":"","variableGroup":"normal","intervention":"VP Normal","inputQty":230,"okQty":228,"ngTotal":2,"ngRate":0.9,"defectCategory":"assembly_defect","defectType":"Particle","defectCount":1},
    {"productType":"BRS-201506DT","testDate":"2025-11-22","line":"C2-3A","checkType":"process","variable":"VP+CD Vision","variableDetail":"","variableGroup":"new_lot","intervention":"VP new lot","inputQty":228,"okQty":226,"ngTotal":2,"ngRate":0.9,"defectCategory":"assembly_defect","defectType":"VP/CD separate","defectCount":2},
    {"productType":"BRS-201506DT","testDate":"2025-11-22","line":"C2-3A","checkType":"process","variable":"VP+CD Vision","variableDetail":"","variableGroup":"old_lot","intervention":"VP Old lot","inputQty":226,"okQty":226,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
    {"productType":"BRS-201506DT","testDate":"2025-11-22","line":"C2-3A","checkType":"process","variable":"VP+CD Vision","variableDetail":"","variableGroup":"normal","intervention":"VP Normal","inputQty":226,"okQty":226,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
    {"productType":"BRS-201506DT","testDate":"2025-11-22","line":"C2-3A","checkType":"function","variable":"Function","variableDetail":"","variableGroup":"new_lot","intervention":"VP new lot","inputQty":224,"okQty":209,"ngTotal":15,"ngRate":6.7,"defectCategory":"function_hearing","defectType":"Noise","defectCount":11},
    {"productType":"BRS-201506DT","testDate":"2025-11-22","line":"C2-3A","checkType":"function","variable":"Function","variableDetail":"","variableGroup":"old_lot","intervention":"VP old lot","inputQty":218,"okQty":197,"ngTotal":21,"ngRate":9.6,"defectCategory":"function_hearing","defectType":"Noise","defectCount":17},
    {"productType":"BRS-201506DT","testDate":"2025-11-22","line":"C2-3A","checkType":"function","variable":"Function","variableDetail":"","variableGroup":"normal","intervention":"VP normal","inputQty":280,"okQty":253,"ngTotal":27,"ngRate":9.6,"defectCategory":"function_hearing","defectType":"Noise","defectCount":20}
  ],
  "tags":["brs-201506dt","vp-new-lot","aem-115u","material-validation","function-test","vp-cd-vision","multi-arm"],
  "reportType":"multi_arm",
  "verdict":"improved",
  "headline":"VP new lot AEM 115u — Function 6.7% best among VP new/old/normal (all OK)",
  "evidence":[
    {"metric":"Function NG rate","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"+2.9pp range","deltaSign":"down","note":"new lot best",
     "comparisons":[
       {"label":"VP new lot","value":"6.7% (15/224)","n":224,"isBaseline":False,"isBest":True,"isWorst":False},
       {"label":"VP old lot","value":"9.6% (21/218)","n":218,"isBaseline":False,"isBest":False,"isWorst":True},
       {"label":"VP Normal","value":"9.6% (27/280)","n":280,"isBaseline":True,"isBest":False,"isWorst":False}
     ],"bestLabel":"VP new lot","worstLabel":"VP old lot"},
    {"metric":"VP+CD Vision NG","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"+0.9pp range","deltaSign":"up","note":"new lot only NG",
     "comparisons":[
       {"label":"VP new lot","value":"0.9% (2/228)","n":228,"isBaseline":False,"isBest":False,"isWorst":True},
       {"label":"VP old lot","value":"0.0% (0/226)","n":226,"isBaseline":False,"isBest":True,"isWorst":False},
       {"label":"VP Normal","value":"0.0% (0/226)","n":226,"isBaseline":True,"isBest":True,"isWorst":False}
     ],"bestLabel":"VP Normal","worstLabel":"VP new lot"}
  ],
  "actions":[
    {"priority":1,"kind":"action","text":"Approve VP new lot AEM 115u (INV IR250206037) for production"},
    {"priority":2,"kind":"investigate","text":"Investigate VP/CD separate NG on new lot caused by bonding jig"}
  ],
  "context":{"process":"VP material new lot AEM 115u (60A) 113mm validation","stage":"C2-3A sub 1 + VP+CD bonding + function","baselineReason":"3-arm: new lot vs old lot vs Normal — all present same event"},
  "doeGrid":None,"trendPoints":None,
  "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":""
},
"ko":{"headline":"VP 신규 로트 AEM 115u — Function 6.7%로 VP new/old/normal 중 최우수 (모두 합격)",
"actionTexts":["VP 신규 로트 AEM 115u (INV IR250206037) 양산 승인","본딩 지그로 인한 신규 로트의 VP/CD 분리 NG 조사"],
"contextProcess":"VP 자재 신규 로트 AEM 115u (60A) 113mm 검증","contextStage":"C2-3A 서브1 + VP+CD 본딩 + Function","contextBaseline":"3-arm: new lot vs old lot vs Normal — 동일 이벤트 모두 존재"},
"vi":{"headline":"VP lô mới AEM 115u — Chức năng 6.7% tốt nhất trong VP new/old/normal (đều đạt)",
"actionTexts":["Phê duyệt VP lô mới AEM 115u (INV IR250206037) cho sản xuất","Điều tra NG VP/CD tách rời ở lô mới do jig bonding"],
"contextProcess":"Xác nhận vật liệu VP lô mới AEM 115u (60A) 113mm","contextStage":"Sub 1 C2-3A + VP+CD bonding + chức năng","contextBaseline":"3-arm: lô mới vs cũ vs Normal — cùng sự kiện"}})

# ----------- DS5: VP deform reason check (UV dry time multi-arm) -------------
RESULTS.append({
"name": "1. MSU-201507 DT Report test Check reson NG VP deform   date 28.2.2025",
"en":{
  "measurements":[
    {"productType":"BRS-201507DT","testDate":"2025-02-28","line":"","checkType":"process","variable":"VP/CD Vision","variableDetail":"UV Dry 10s, Bonding line normal","variableGroup":"test","intervention":"UV Dry 10s","inputQty":64,"okQty":57,"ngTotal":7,"ngRate":10.9,"defectCategory":"assembly_defect","defectType":"VP deform","defectCount":6},
    {"productType":"BRS-201507DT","testDate":"2025-02-28","line":"","checkType":"process","variable":"VP/CD Vision","variableDetail":"UV Dry 10s, Bonding line normal","variableGroup":"test","intervention":"UV Dry 10s","inputQty":64,"okQty":57,"ngTotal":7,"ngRate":10.9,"defectCategory":"assembly_defect","defectType":"CD damage","defectCount":1},
    {"productType":"BRS-201507DT","testDate":"2025-02-28","line":"","checkType":"process","variable":"VP/CD Vision","variableDetail":"UV Dry 3s, Bonding line normal","variableGroup":"test","intervention":"UV Dry 3s","inputQty":20,"okQty":2,"ngTotal":18,"ngRate":90.0,"defectCategory":"assembly_defect","defectType":"VP deform","defectCount":18},
    {"productType":"BRS-201507DT","testDate":"2025-02-28","line":"","checkType":"process","variable":"VP/CD Vision","variableDetail":"UV Dry 3s + increase UC press height","variableGroup":"test","intervention":"UV Dry 3s + raise UC press","inputQty":12,"okQty":3,"ngTotal":9,"ngRate":75.0,"defectCategory":"assembly_defect","defectType":"VP deform","defectCount":9},
    {"productType":"BRS-201507DT","testDate":"2025-02-28","line":"","checkType":"process","variable":"VP/CD Vision","variableDetail":"UV Dry 4s + Open bonding line 0.06","variableGroup":"test","intervention":"UV Dry 4s + open bonding 0.06","inputQty":248,"okQty":248,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0}
  ],
  "tags":["msu-201507","brs-201507dt","vp-deform","uv-dry-time","bonding-line-gap","root-cause-investigation","multi-arm"],
  "reportType":"multi_arm",
  "verdict":"improved",
  "headline":"VP deform fixed by UV Dry 4s + open bonding 0.06 (NG 90% → 0%, 248pcs OK)",
  "evidence":[
    {"metric":"VP/CD Vision NG rate","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"+90pp range","deltaSign":"down","note":"open bonding fixes deform",
     "comparisons":[
       {"label":"UV 10s normal","value":"10.9% (7/64)","n":64,"isBaseline":True,"isBest":False,"isWorst":False},
       {"label":"UV 3s normal","value":"90.0% (18/20)","n":20,"isBaseline":False,"isBest":False,"isWorst":True},
       {"label":"UV 3s + UC press up","value":"75.0% (9/12)","n":12,"isBaseline":False,"isBest":False,"isWorst":False},
       {"label":"UV 4s + open 0.06","value":"0.0% (0/248)","n":248,"isBaseline":False,"isBest":True,"isWorst":False}
     ],"bestLabel":"UV 4s + open 0.06","worstLabel":"UV 3s normal"}
  ],
  "actions":[
    {"priority":1,"kind":"action","text":"Adopt UV Dry 4s with bonding line gap +0.06mm"},
    {"priority":2,"kind":"risk","text":"Avoid UV Dry 3s — produces 75-90% VP deform"}
  ],
  "context":{"process":"VP/CD UV dry and bonding line gap tuning","stage":"VP/CD vision sub line","baselineReason":"first arm UV 10s is current baseline; deform root-cause search"},
  "doeGrid":None,"trendPoints":None,
  "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":""
},
"ko":{"headline":"VP 변형, UV Dry 4s + 본딩 라인 0.06 개방으로 해결 (NG 90% → 0%, 248pcs 합격)",
"actionTexts":["UV Dry 4s + 본딩 라인 갭 +0.06mm 적용","UV Dry 3s 사용 금지 — VP 변형 75-90% 발생"],
"contextProcess":"VP/CD UV 건조 및 본딩 라인 갭 튜닝","contextStage":"VP/CD 비전 서브 라인","contextBaseline":"첫번째 arm UV 10s를 현행 기준으로 보고 변형 원인 탐색"},
"vi":{"headline":"VP biến dạng được khắc phục bằng UV Dry 4s + mở bonding 0.06 (NG 90% → 0%, 248pcs đạt)",
"actionTexts":["Áp dụng UV Dry 4s với khe bonding +0.06mm","Tránh UV Dry 3s — gây biến dạng VP 75-90%"],
"contextProcess":"Tinh chỉnh UV dry và khe bonding cho VP/CD","contextStage":"Line sub vision VP/CD","contextBaseline":"Arm đầu UV 10s là chuẩn hiện tại; tìm nguyên nhân biến dạng"}})

# ----------- DS6: Lot Dome 3-17 reason check -------------
RESULTS.append({
"name": "1. MSU-L20S15-07DT Report test Find reason NG  LOT DOME 3-17  date 28.4.2025",
"en":{
  "measurements":[
    {"productType":"BRS-201507DT","testDate":"2025-04-28","line":"C2-2A","checkType":"process","variable":"VP/CD Separate","variableDetail":"Laser marking + Plasma","variableGroup":"test","intervention":"Laser marking CD","inputQty":43,"okQty":43,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
    {"productType":"BRS-201507DT","testDate":"2025-04-28","line":"C2-2A","checkType":"process","variable":"VP/CD Separate","variableDetail":"Primer clean + Plasma","variableGroup":"test","intervention":"Primer clean CD","inputQty":49,"okQty":36,"ngTotal":13,"ngRate":26.5,"defectCategory":"assembly_defect","defectType":"VP/CD separate","defectCount":13},
    {"productType":"BRS-201507DT","testDate":"2025-04-28","line":"C2-2A","checkType":"process","variable":"Tension VP+CD Ass'y","variableDetail":"Spec 1.2 kgf — Laser marking","variableGroup":"test","intervention":"Laser marking CD","inputQty":5,"okQty":5,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
    {"productType":"BRS-201507DT","testDate":"2025-04-28","line":"C2-2A","checkType":"process","variable":"Tension VP+CD Ass'y","variableDetail":"Spec 1.2 kgf — Primer clean","variableGroup":"test","intervention":"Primer clean CD","inputQty":5,"okQty":5,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
    {"productType":"BRS-201507DT","testDate":"2025-04-28","line":"C2-2A","checkType":"function","variable":"Function","variableDetail":"Laser marking","variableGroup":"test","intervention":"Laser marking CD","inputQty":43,"okQty":42,"ngTotal":1,"ngRate":2.3,"defectCategory":"function_hearing","defectType":"Touch","defectCount":1},
    {"productType":"BRS-201507DT","testDate":"2025-04-28","line":"C2-2A","checkType":"function","variable":"Function","variableDetail":"Primer clean","variableGroup":"test","intervention":"Primer clean CD","inputQty":36,"okQty":34,"ngTotal":2,"ngRate":5.6,"defectCategory":"function_hearing","defectType":"Touch","defectCount":1}
  ],
  "tags":["msu-l20s15-07dt","lot-dome-3-17","ng-vp-cd-separate","laser-marking","primer-clean","root-cause"],
  "reportType":"comparison_study",
  "verdict":"improved",
  "headline":"Lot Dome 3/17 — Laser marking eliminates VP/CD separate (26.5% → 0%, function 2.3% vs 5.6%)",
  "evidence":[
    {"metric":"VP/CD Separate NG","baselineLabel":"Primer clean","baselineValue":"26.5% (13/49)","variantLabel":"Laser marking","variantValue":"0.0% (0/43)","deltaText":"-26.5pp","deltaSign":"down","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"Function NG","baselineLabel":"Primer clean","baselineValue":"5.6% (2/36)","variantLabel":"Laser marking","variantValue":"2.3% (1/43)","deltaText":"-3.3pp","deltaSign":"down","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"Tension AVG","baselineLabel":"Primer clean","baselineValue":"1.572 kgf","variantLabel":"Laser marking","variantValue":"2.007 kgf","deltaText":"+0.4","deltaSign":"up","note":"both pass spec 1.2 kgf","comparisons":None,"bestLabel":"","worstLabel":""}
  ],
  "actions":[
    {"priority":1,"kind":"action","text":"Apply laser marking CD process for lot dome 3/17 issue"},
    {"priority":2,"kind":"investigate","text":"Identify why primer clean alone fails — surface residue suspect"}
  ],
  "context":{"process":"VP/CD assembly CD pretreatment selection","stage":"C2-2A line, lot Dome date 3/17","baselineReason":"primer clean is the current pretreatment used"},
  "doeGrid":None,"trendPoints":None,
  "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":""
},
"ko":{"headline":"Lot Dome 3/17 — 레이저 마킹으로 VP/CD 분리 NG 제거 (26.5% → 0%, Function 2.3% vs 5.6%)",
"actionTexts":["lot dome 3/17 이슈에 레이저 마킹 CD 공정 적용","프라이머 클린 단독 실패 원인 — 표면 잔류물 가능성 조사"],
"contextProcess":"VP/CD 조립 CD 전처리 방식 선정","contextStage":"C2-2A 라인, Dome 로트 날짜 3/17","contextBaseline":"현재 프라이머 클린이 사용 중인 전처리"},
"vi":{"headline":"Lot Dome 3/17 — Laser marking loại bỏ VP/CD tách (26.5% → 0%, chức năng 2.3% vs 5.6%)",
"actionTexts":["Áp dụng quy trình laser marking CD cho lô dome 3/17","Điều tra vì sao chỉ primer clean thất bại — nghi cặn bề mặt"],
"contextProcess":"Chọn tiền xử lý CD cho lắp ráp VP/CD","contextStage":"Line C2-2A, lô Dome ngày 3/17","contextBaseline":"Primer clean là tiền xử lý đang dùng"}})

# ----------- DS7: VP die punch V0 -------------
RESULTS.append({
"name": "1.1 BRS-201506 Report test material VP using die punch cutting inside ( V0 ) Date 28.8.2024",
"en":{
  "measurements":[
    {"productType":"BRS-201506","testDate":"2024-07-26","line":"","checkType":"process","variable":"VP+CD Vision","variableDetail":"","variableGroup":"test","intervention":"VP Die Punch V0","inputQty":211,"okQty":211,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
    {"productType":"BRS-201506","testDate":"2024-07-26","line":"","checkType":"process","variable":"VP+CD Vision","variableDetail":"","variableGroup":"normal","intervention":"VP normal","inputQty":200,"okQty":199,"ngTotal":1,"ngRate":0.5,"defectCategory":"assembly_defect","defectType":"Dome offset","defectCount":1},
    {"productType":"BRS-201506","testDate":"2024-08-27","line":"","checkType":"process","variable":"VP+CD Gap","variableDetail":"","variableGroup":"test","intervention":"VP Die Punch V0","inputQty":1008,"okQty":1007,"ngTotal":1,"ngRate":0.2,"defectCategory":"assembly_defect","defectType":"VP+CD Gap","defectCount":211},
    {"productType":"BRS-201506","testDate":"2024-08-30","line":"","checkType":"process","variable":"VP+CD Vision","variableDetail":"","variableGroup":"test","intervention":"VP Die Punch V0","inputQty":973,"okQty":728,"ngTotal":245,"ngRate":25.2,"defectCategory":"assembly_defect","defectType":"Glue Clots","defectCount":244},
    {"productType":"BRS-201506","testDate":"2024-08-30","line":"","checkType":"process","variable":"VP+CD Vision","variableDetail":"","variableGroup":"normal","intervention":"VP normal","inputQty":2002,"okQty":2002,"ngTotal":2,"ngRate":0.1,"defectCategory":"assembly_defect","defectType":"Not enough glue","defectCount":1},
    {"productType":"BRS-201506","testDate":"2024-08-28","line":"","checkType":"function","variable":"Function — Hearing up 10%","variableDetail":"All V0 lots aggregated","variableGroup":"test","intervention":"VP Die Punch V0 total","inputQty":1691,"okQty":1464,"ngTotal":335,"ngRate":19.8,"defectCategory":"function_hearing","defectType":"Noise","defectCount":214},
    {"productType":"BRS-201506","testDate":"2024-08-28","line":"","checkType":"function","variable":"Function — Hearing up 10%","variableDetail":"All V0 lots aggregated","variableGroup":"test","intervention":"VP Die Punch V0 total","inputQty":1691,"okQty":1464,"ngTotal":335,"ngRate":19.8,"defectCategory":"function_hearing","defectType":"Touch","defectCount":113},
    {"productType":"BRS-201506","testDate":"2024-08-28","line":"","checkType":"function","variable":"Function — Hearing up 10%","variableDetail":"","variableGroup":"normal","intervention":"VP normal","inputQty":924,"okQty":881,"ngTotal":115,"ngRate":12.4,"defectCategory":"function_hearing","defectType":"Touch","defectCount":100},
    {"productType":"BRS-201506","testDate":"2024-08-28","line":"","checkType":"function","variable":"Air Leak","variableDetail":"VP+CD Gap subset","variableGroup":"test","intervention":"VP Die Punch V0 (VP+CD Gap)","inputQty":201,"okQty":197,"ngTotal":4,"ngRate":2.0,"defectCategory":"function_spl","defectType":"Air leak","defectCount":4}
  ],
  "tags":["brs-201506","vp-die-punch-v0","material-validation","vp-cd-gap","function-hearing","air-leak","comparison"],
  "reportType":"comparison_study",
  "verdict":"worsened",
  "headline":"VP Die Punch V0 — Function NG 19.8% vs Normal 12.4%, vision Gap 25%, NEEDS improvement",
  "evidence":[
    {"metric":"Function NG (Hearing up 10%)","baselineLabel":"VP normal","baselineValue":"12.4% (115/924)","variantLabel":"V0","variantValue":"19.8% (335/1691)","deltaText":"+7.4pp","deltaSign":"up","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"VP+CD Vision NG (8/30)","baselineLabel":"VP normal","baselineValue":"0.1% (2/2002)","variantLabel":"V0","variantValue":"25.2% (245/973)","deltaText":"+25.1pp","deltaSign":"up","note":"Glue Clots dominant","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"Air Leak NG","baselineLabel":"Normal","baselineValue":"0.0% (0/924)","variantLabel":"V0 (VP+CD Gap subset)","variantValue":"2.0% (4/201)","deltaText":"+2.0pp","deltaSign":"up","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"Hearing NG breakdown","baselineLabel":"Normal","baselineValue":"15.0% noise / 166.7% touch","variantLabel":"V0","variantValue":"63.9% noise / 33.7% touch","deltaText":"+49pp noise","deltaSign":"up","note":"VP+Dome separate root cause","comparisons":None,"bestLabel":"","worstLabel":""}
  ],
  "actions":[
    {"priority":1,"kind":"action","text":"Do NOT release VP Die Punch V0 to production"},
    {"priority":2,"kind":"investigate","text":"Reduce VP+Dome separate (83% of hearing NG) — bonding stick jig"},
    {"priority":3,"kind":"investigate","text":"Eliminate Glue Clots root cause in VP+CD vision"}
  ],
  "context":{"process":"VP material with die punch cutting inside (V0)","stage":"Sub 1 VP+CD vision + function lines","baselineReason":"VP normal lot processed in parallel each day"},
  "doeGrid":None,"trendPoints":None,
  "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":""
},
"ko":{"headline":"VP Die Punch V0 — Function NG 19.8% vs Normal 12.4%, Vision Gap 25%, 개선 필요",
"actionTexts":["VP Die Punch V0 양산 적용 보류","Hearing NG의 83%를 차지하는 VP+Dome 분리 — 본딩 스틱 지그 조사","VP+CD 비전의 Glue Clots 원인 제거"],
"contextProcess":"VP 자재 다이 펀치 내측 컷팅 (V0)","contextStage":"Sub 1 VP+CD 비전 + Function 라인","contextBaseline":"매일 VP normal 로트가 병행 가공됨"},
"vi":{"headline":"VP Die Punch V0 — NG chức năng 19.8% vs Normal 12.4%, Vision Gap 25%, cần cải tiến",
"actionTexts":["KHÔNG đưa VP Die Punch V0 vào sản xuất","Giảm VP+Dome tách (83% NG hearing) — kiểm tra jig bonding stick","Loại bỏ nguyên nhân Glue Clots trong VP+CD vision"],
"contextProcess":"Vật liệu VP cắt die punch bên trong (V0)","contextStage":"Sub 1 VP+CD vision + line chức năng","contextBaseline":"Lô VP normal chạy song song mỗi ngày"}})

# ----------- DS8: new JIG coil array + frame array -------------
RESULTS.append({
"name": "1.BRS-161014  REPORT TEST  new JIG coil array and JIG frame array date 29.06.2023",
"en":{
  "measurements":[
    {"productType":"BRS-161014","testDate":"2023-06-29","line":"","checkType":"process","variable":"Ass'y Coil+SP","variableDetail":"","variableGroup":"test","intervention":"New JIG Coil + Frame Array","inputQty":12,"okQty":0,"ngTotal":12,"ngRate":100.0,"defectCategory":"assembly_defect","defectType":"Coil offset","defectCount":9},
    {"productType":"BRS-161014","testDate":"2023-06-29","line":"","checkType":"process","variable":"Ass'y Coil+SP","variableDetail":"","variableGroup":"test","intervention":"New JIG Coil + Frame Array","inputQty":12,"okQty":0,"ngTotal":12,"ngRate":100.0,"defectCategory":"assembly_defect","defectType":"Coil separate","defectCount":3},
    {"productType":"BRS-161014","testDate":"2023-06-29","line":"","checkType":"process","variable":"Frame dimension","variableDetail":"Spec 9.87-9.92, n=20","variableGroup":"test","intervention":"Frame check","inputQty":20,"okQty":20,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
    {"productType":"BRS-161014","testDate":"2023-06-29","line":"","checkType":"process","variable":"JIG Frame Array 2D dimension","variableDetail":"Spec 9.91-9.93, max 9.946 out of spec","variableGroup":"test","intervention":"JIG dimension check","inputQty":24,"okQty":0,"ngTotal":24,"ngRate":100.0,"defectCategory":"assembly_defect","defectType":"Dimension over","defectCount":24}
  ],
  "tags":["brs-161014","new-jig-coil-array","jig-frame-array","jig-dimension","coil-offset","jig-improvement-needed"],
  "reportType":"intervention_test",
  "verdict":"failed",
  "headline":"New JIG Frame Array out-of-spec (max 9.946 > 9.93) → 100% NG coil offset, JIG needs rework",
  "evidence":[
    {"metric":"Ass'y Coil+SP NG","baselineLabel":"Spec","baselineValue":"0% (0/12)","variantLabel":"New JIG","variantValue":"100% (12/12)","deltaText":"+100pp","deltaSign":"up","note":"75% coil offset","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"JIG Frame Array dimension","baselineLabel":"Spec 9.91-9.93","baselineValue":"9.91-9.93","variantLabel":"Measured","variantValue":"9.912-9.946 (AVG 9.929)","deltaText":"+0.016 over","deltaSign":"up","note":"max exceeds upper spec","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"Frame raw dimension","baselineLabel":"Spec 9.87-9.92","baselineValue":"9.87-9.92","variantLabel":"Measured","variantValue":"9.87-9.88 (AVG 9.872)","deltaText":"in spec","deltaSign":"no_change","note":"frame itself OK","comparisons":None,"bestLabel":"","worstLabel":""}
  ],
  "actions":[
    {"priority":1,"kind":"action","text":"Rework JIG Frame Array — bring dimensions inside spec 9.91-9.93"},
    {"priority":2,"kind":"risk","text":"Frame array offset propagates to Coil offset in assembly"}
  ],
  "context":{"process":"New JIG Coil Array + Frame Array first-article validation","stage":"Sub Frame array assembly","baselineReason":"design spec used as baseline; no production baseline yet"},
  "doeGrid":None,"trendPoints":None,
  "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":""
},
"ko":{"headline":"신규 JIG Frame Array 규격 초과 (최대 9.946 > 9.93) → Coil offset 100% NG, JIG 재가공 필요",
"actionTexts":["JIG Frame Array 재가공 — 치수를 9.91-9.93 규격 이내로","Frame array 오프셋이 조립 시 Coil offset으로 전이"],
"contextProcess":"신규 JIG Coil Array + Frame Array 초도 검증","contextStage":"Sub Frame array 조립","contextBaseline":"설계 규격을 기준으로 사용; 양산 기준 없음"},
"vi":{"headline":"JIG Frame Array mới vượt spec (max 9.946 > 9.93) → NG coil offset 100%, cần gia công lại",
"actionTexts":["Gia công lại JIG Frame Array — đưa kích thước về 9.91-9.93","Frame array offset lan sang Coil offset khi lắp ráp"],
"contextProcess":"Xác nhận đầu tiên JIG Coil Array + Frame Array mới","contextStage":"Lắp ráp Sub Frame array","contextBaseline":"Lấy spec thiết kế làm chuẩn; chưa có chuẩn sản xuất"}})

# ----------- DS9: CSY VINA new vendor suspension -------------
RESULTS.append({
"name": "1.BRS-161016, 201506 Report test Suspension of new vendor CSY VINA date 23.07.2024",
"en":{
  "measurements":[
    {"productType":"BRS-161016","testDate":"2024-07-22","line":"E2-3A","checkType":"visual_inspection","variable":"Vision Frame+SP","variableDetail":"","variableGroup":"test","intervention":"CSY VINA suspension","inputQty":5096,"okQty":5088,"ngTotal":8,"ngRate":0.16,"defectCategory":"assembly_defect","defectType":"Separate SP","defectCount":4},
    {"productType":"BRS-161016","testDate":"2024-07-22","line":"E2-3A","checkType":"visual_inspection","variable":"Vision Frame+SP","variableDetail":"","variableGroup":"normal","intervention":"Normal","inputQty":1000,"okQty":998,"ngTotal":2,"ngRate":0.20,"defectCategory":"assembly_defect","defectType":"Separate SP","defectCount":1},
    {"productType":"BRS-161016","testDate":"2024-07-23","line":"E2-3A","checkType":"function","variable":"Function","variableDetail":"","variableGroup":"test","intervention":"CSY VINA suspension","inputQty":3064,"okQty":2933,"ngTotal":131,"ngRate":4.3,"defectCategory":"function_hearing","defectType":"Noise","defectCount":76},
    {"productType":"BRS-161016","testDate":"2024-07-23","line":"E2-3A","checkType":"function","variable":"Function","variableDetail":"","variableGroup":"normal","intervention":"Normal","inputQty":930,"okQty":888,"ngTotal":42,"ngRate":4.5,"defectCategory":"function_hearing","defectType":"Noise","defectCount":22},
    {"productType":"BRS-201506","testDate":"2024-07-23","line":"E2-4B","checkType":"process","variable":"Suspension Array","variableDetail":"WORLD HANOI VINA vendor","variableGroup":"test","intervention":"New vendor","inputQty":7500,"okQty":7500,"ngTotal":974,"ngRate":13.0,"defectCategory":"assembly_defect","defectType":"Suspension bending","defectCount":974},
    {"productType":"BRS-201506","testDate":"2024-07-23","line":"E2-4B","checkType":"process","variable":"Suspension Array","variableDetail":"","variableGroup":"normal","intervention":"Normal","inputQty":1000,"okQty":1000,"ngTotal":10,"ngRate":1.0,"defectCategory":"assembly_defect","defectType":"Suspension bending","defectCount":10},
    {"productType":"BRS-201506","testDate":"2024-07-23","line":"E2-4B","checkType":"process","variable":"Frame+SUS Vision","variableDetail":"","variableGroup":"test","intervention":"New vendor","inputQty":6500,"okQty":6405,"ngTotal":95,"ngRate":1.5,"defectCategory":"assembly_defect","defectType":"Frame+sus GAP","defectCount":95},
    {"productType":"BRS-201506","testDate":"2024-07-23","line":"E2-4B","checkType":"process","variable":"Frame+SUS Vision","variableDetail":"","variableGroup":"normal","intervention":"Normal","inputQty":1000,"okQty":998,"ngTotal":2,"ngRate":0.2,"defectCategory":"assembly_defect","defectType":"Frame+sus GAP","defectCount":2},
    {"productType":"BRS-201506","testDate":"2024-07-25","line":"","checkType":"function","variable":"Function","variableDetail":"Hearing up 10%","variableGroup":"test","intervention":"New vendor WORLD HANOI","inputQty":3914,"okQty":3336,"ngTotal":578,"ngRate":14.8,"defectCategory":"function_hearing","defectType":"Touch","defectCount":403},
    {"productType":"BRS-201506","testDate":"2024-07-25","line":"","checkType":"function","variable":"Function","variableDetail":"Hearing up 10%","variableGroup":"normal","intervention":"Normal","inputQty":2543,"okQty":2104,"ngTotal":310,"ngRate":12.2,"defectCategory":"function_hearing","defectType":"Touch","defectCount":171}
  ],
  "tags":["brs-161016","brs-201506","new-vendor-csy-vina","world-hanoi-vina","suspension-bending","material-qualification"],
  "reportType":"comparison_study",
  "verdict":"partial",
  "headline":"CSY VINA suspension OK on 161016 (function 4.3 vs 4.5%) but 201506 World Hanoi suspension bending 13% — defer",
  "evidence":[
    {"metric":"161016 Vision Frame+SP NG","baselineLabel":"Normal","baselineValue":"0.20% (2/1000)","variantLabel":"CSY VINA","variantValue":"0.16% (8/5096)","deltaText":"-0.04pp","deltaSign":"down","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"161016 Function NG","baselineLabel":"Normal","baselineValue":"4.5% (42/930)","variantLabel":"CSY VINA","variantValue":"4.3% (131/3064)","deltaText":"-0.2pp","deltaSign":"down","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"201506 Suspension Array bending","baselineLabel":"Normal","baselineValue":"1.0% (10/1000)","variantLabel":"World Hanoi","variantValue":"13.0% (974/7500)","deltaText":"+12pp","deltaSign":"up","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"201506 Function NG (Hearing up 10%)","baselineLabel":"Normal","baselineValue":"12.2% (310/2543)","variantLabel":"World Hanoi","variantValue":"14.8% (578/3914)","deltaText":"+2.6pp","deltaSign":"up","note":"","comparisons":None,"bestLabel":"","worstLabel":""}
  ],
  "actions":[
    {"priority":1,"kind":"action","text":"Approve CSY VINA suspension for BRS-161016"},
    {"priority":2,"kind":"risk","text":"Reject WORLD HANOI VINA suspension for BRS-201506 — bending 13x normal"},
    {"priority":3,"kind":"investigate","text":"Vendor process audit at WORLD HANOI for suspension bending root cause"}
  ],
  "context":{"process":"New suspension vendor qualification (two vendors, two models)","stage":"Sub 3 Frame+SP / Suspension Array / function","baselineReason":"Normal lot run in parallel for each model"},
  "doeGrid":None,"trendPoints":None,
  "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":""
},
"ko":{"headline":"CSY VINA 서스펜션 161016 합격 (Function 4.3% vs 4.5%) 그러나 201506 World Hanoi는 굽힘 13% — 보류",
"actionTexts":["BRS-161016용 CSY VINA 서스펜션 승인","BRS-201506용 WORLD HANOI VINA 서스펜션 보류 — 굽힘이 정상의 13배","WORLD HANOI 공장 서스펜션 굽힘 원인 감사"],
"contextProcess":"신규 서스펜션 벤더 인증 (2개 벤더, 2개 모델)","contextStage":"Sub 3 Frame+SP / 서스펜션 어레이 / Function","contextBaseline":"각 모델별로 Normal 로트가 병행 가공됨"},
"vi":{"headline":"Suspension CSY VINA đạt 161016 (chức năng 4.3% vs 4.5%) nhưng 201506 World Hanoi cong 13% — hoãn",
"actionTexts":["Phê duyệt suspension CSY VINA cho BRS-161016","Từ chối suspension WORLD HANOI VINA cho BRS-201506 — cong gấp 13 lần","Kiểm toán quy trình NCC WORLD HANOI tìm nguyên nhân suspension cong"],
"contextProcess":"Đánh giá NCC suspension mới (2 NCC, 2 model)","contextStage":"Sub 3 Frame+SP / Suspension Array / chức năng","contextBaseline":"Lô Normal chạy song song cho từng model"}})

# ----------- DS10: New Mold Frame G4 -------------
RESULTS.append({
"name": "10 . MSU - L20S15-07 REPORT TEST NEW MOLD FRAME G4 24.3.2025 -",
"en":{
  "measurements":[
    {"productType":"MSU-L20S15-07","testDate":"2025-03-24","line":"","checkType":"function","variable":"Function — Sigma","variableDetail":"","variableGroup":"test","intervention":"Frame G4 new mold","inputQty":3801,"okQty":3709,"ngTotal":3,"ngRate":0.1,"defectCategory":"function_thd","defectType":"THD","defectCount":1},
    {"productType":"MSU-L20S15-07","testDate":"2025-03-24","line":"","checkType":"function","variable":"Function — Hearing +1V","variableDetail":"","variableGroup":"test","intervention":"Frame G4 new mold","inputQty":3801,"okQty":3709,"ngTotal":89,"ngRate":2.3,"defectCategory":"function_hearing","defectType":"Noise","defectCount":50},
    {"productType":"MSU-L20S15-07","testDate":"2025-03-24","line":"","checkType":"function","variable":"Function — Hearing +1V","variableDetail":"","variableGroup":"test","intervention":"Frame G4 new mold","inputQty":3801,"okQty":3709,"ngTotal":89,"ngRate":2.3,"defectCategory":"function_hearing","defectType":"Touch","defectCount":39},
    {"productType":"MSU-L20S15-07","testDate":"2025-03-24","line":"","checkType":"function","variable":"Function — Hearing +0V","variableDetail":"","variableGroup":"test","intervention":"Frame G4 new mold","inputQty":3801,"okQty":3709,"ngTotal":29,"ngRate":0.8,"defectCategory":"function_hearing","defectType":"Noise","defectCount":15},
    {"productType":"MSU-L20S15-07","testDate":"2025-03-24","line":"","checkType":"function","variable":"Function — Sigma","variableDetail":"","variableGroup":"normal","intervention":"Normal","inputQty":2260,"okQty":2210,"ngTotal":2,"ngRate":0.1,"defectCategory":"function_thd","defectType":"THD","defectCount":2},
    {"productType":"MSU-L20S15-07","testDate":"2025-03-24","line":"","checkType":"function","variable":"Function — Hearing +1V","variableDetail":"","variableGroup":"normal","intervention":"Normal","inputQty":2260,"okQty":2210,"ngTotal":48,"ngRate":2.1,"defectCategory":"function_hearing","defectType":"Noise","defectCount":33},
    {"productType":"MSU-L20S15-07","testDate":"2025-03-24","line":"","checkType":"function","variable":"Function — Hearing +0V","variableDetail":"","variableGroup":"normal","intervention":"Normal","inputQty":2260,"okQty":2210,"ngTotal":14,"ngRate":0.6,"defectCategory":"function_hearing","defectType":"Noise","defectCount":13}
  ],
  "tags":["msu-l20s15-07","new-mold-frame-g4","function-test","mold-qualification","sigma","hearing"],
  "reportType":"comparison_study",
  "verdict":"no_clear_effect",
  "headline":"New Mold Frame G4 — function NG matches Normal (2.3% vs 2.1% Hearing +1V), OK to use",
  "evidence":[
    {"metric":"Function Sigma NG","baselineLabel":"Normal","baselineValue":"0.1% (2/2260)","variantLabel":"Frame G4","variantValue":"0.1% (3/3801)","deltaText":"+0pp","deltaSign":"no_change","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"Function Hearing +1V NG","baselineLabel":"Normal","baselineValue":"2.1% (48/2260)","variantLabel":"Frame G4","variantValue":"2.3% (89/3801)","deltaText":"+0.2pp","deltaSign":"up","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"Function Hearing +0V NG","baselineLabel":"Normal","baselineValue":"0.6% (14/2260)","variantLabel":"Frame G4","variantValue":"0.8% (29/3801)","deltaText":"+0.2pp","deltaSign":"up","note":"","comparisons":None,"bestLabel":"","worstLabel":""}
  ],
  "actions":[
    {"priority":1,"kind":"action","text":"Approve New Mold Frame G4 for production"}
  ],
  "context":{"process":"Frame G4 new mold introduction","stage":"Final function test","baselineReason":"same-event Normal row present"},
  "doeGrid":None,"trendPoints":None,
  "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":""
},
"ko":{"headline":"신규 Mold Frame G4 — Function NG가 Normal과 동등 (2.3% vs 2.1% Hearing +1V), 사용 승인",
"actionTexts":["신규 Mold Frame G4 양산 승인"],
"contextProcess":"Frame G4 신규 금형 도입","contextStage":"최종 Function 테스트","contextBaseline":"같은 이벤트에 Normal 행 존재"},
"vi":{"headline":"Khuôn Frame G4 mới — NG chức năng tương đương Normal (2.3% vs 2.1% Hearing +1V), được dùng",
"actionTexts":["Phê duyệt khuôn Frame G4 mới cho sản xuất"],
"contextProcess":"Đưa khuôn Frame G4 mới vào","contextStage":"Kiểm tra chức năng cuối","contextBaseline":"Có dòng Normal cùng sự kiện"}})

# ----------- DS11: YK happen different color -------------
RESULTS.append({
"name": "10. BRS-161014  DT Report test YK happen different color  date 7.10.2024",
"en":{
  "measurements":[
    {"productType":"BRS-161016","testDate":"2024-10-07","line":"C2-3B","checkType":"visual_inspection","variable":"Vision Yoke","variableDetail":"","variableGroup":"test","intervention":"YK Clean Color","inputQty":92,"okQty":92,"ngTotal":0,"ngRate":0.0,"defectCategory":"cosmetic_defect","defectType":"","defectCount":0},
    {"productType":"BRS-161016","testDate":"2024-10-07","line":"C2-3B","checkType":"visual_inspection","variable":"Vision Yoke","variableDetail":"","variableGroup":"test","intervention":"YK Not Clean Color","inputQty":102,"okQty":102,"ngTotal":0,"ngRate":0.0,"defectCategory":"cosmetic_defect","defectType":"","defectCount":0},
    {"productType":"BRS-161016","testDate":"2024-10-07","line":"C2-3B","checkType":"visual_inspection","variable":"Vision Yoke","variableDetail":"","variableGroup":"normal","intervention":"Normal","inputQty":100,"okQty":100,"ngTotal":0,"ngRate":0.0,"defectCategory":"cosmetic_defect","defectType":"","defectCount":0},
    {"productType":"BRS-161016","testDate":"2024-10-08","line":"C2-3B","checkType":"function","variable":"Function","variableDetail":"","variableGroup":"test","intervention":"YK Clean Color","inputQty":92,"okQty":91,"ngTotal":1,"ngRate":1.1,"defectCategory":"function_hearing","defectType":"Noise","defectCount":1},
    {"productType":"BRS-161016","testDate":"2024-10-08","line":"C2-3B","checkType":"function","variable":"Function","variableDetail":"","variableGroup":"test","intervention":"YK Not Clean Color","inputQty":102,"okQty":100,"ngTotal":2,"ngRate":2.0,"defectCategory":"function_hearing","defectType":"Noise","defectCount":2},
    {"productType":"BRS-161016","testDate":"2024-10-08","line":"C2-3B","checkType":"function","variable":"Function","variableDetail":"","variableGroup":"normal","intervention":"Normal","inputQty":560,"okQty":548,"ngTotal":12,"ngRate":2.1,"defectCategory":"function_hearing","defectType":"Noise","defectCount":11}
  ],
  "tags":["brs-161016","yoke-color","cosmetic-variation","function-test","tension-test","decap-bond","multi-arm"],
  "reportType":"multi_arm",
  "verdict":"no_clear_effect",
  "headline":"YK color variation (Clean vs Not Clean) — function NG 1.1/2.0/2.1% across arms (all OK)",
  "evidence":[
    {"metric":"Function NG","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"+1.0pp range","deltaSign":"no_change","note":"all within normal range",
     "comparisons":[
       {"label":"YK Clean Color","value":"1.1% (1/92)","n":92,"isBaseline":False,"isBest":True,"isWorst":False},
       {"label":"YK Not Clean Color","value":"2.0% (2/102)","n":102,"isBaseline":False,"isBest":False,"isWorst":False},
       {"label":"Normal","value":"2.1% (12/560)","n":560,"isBaseline":True,"isBest":False,"isWorst":True}
     ],"bestLabel":"YK Clean Color","worstLabel":"Normal"},
    {"metric":"Tension MG-S-C AVG","baselineLabel":"Normal","baselineValue":"71.53","variantLabel":"YK Clean / Not Clean","variantValue":"73.36 / 73.38","deltaText":"+1.8","deltaSign":"up","note":"spec ≥80 kgf — all PASS internal ≥50","comparisons":None,"bestLabel":"","worstLabel":""}
  ],
  "actions":[
    {"priority":1,"kind":"action","text":"Approve YK with cosmetic color variation for production"},
    {"priority":2,"kind":"investigate","text":"Further reduce YK color variation source if cosmetically rejected"}
  ],
  "context":{"process":"Yoke color variation after dry oven","stage":"C2-3B sub line vision + function","baselineReason":"3-arm same-event: Clean/Not Clean/Normal"},
  "doeGrid":None,"trendPoints":None,
  "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":""
},
"ko":{"headline":"YK 색상 편차 (Clean vs Not Clean) — Function NG 1.1/2.0/2.1% (모두 합격)",
"actionTexts":["외관 색상 편차가 있는 YK 양산 승인","외관 거부 시 YK 색상 편차 원인 추가 저감"],
"contextProcess":"건조 오븐 이후 Yoke 색상 편차","contextStage":"C2-3B 서브 라인 비전 + Function","contextBaseline":"3-arm 동일 이벤트: Clean/Not Clean/Normal"},
"vi":{"headline":"Sai khác màu YK (Clean vs Not Clean) — NG chức năng 1.1/2.0/2.1% (đều đạt)",
"actionTexts":["Phê duyệt YK có sai khác màu thẩm mỹ cho sản xuất","Giảm thêm nguyên nhân sai màu YK nếu bị từ chối ngoại quan"],
"contextProcess":"Sai khác màu Yoke sau lò sấy","contextStage":"Line sub C2-3B vision + chức năng","contextBaseline":"3-arm cùng sự kiện: Clean/Not Clean/Normal"}})

# ----------- DS12: Sub 2 CMCP problem -------------
RESULTS.append({
"name": "10. BRS-161014  Report check PROBLEM SUB 2 CMCP 2023.11.18",
"en":{
  "measurements":[
    {"productType":"BRS-161014","testDate":"2023-11-18","line":"","checkType":"process","variable":"Decap bond CMG+CP","variableDetail":"Bonding 0.2-0.4mg in spec","variableGroup":"test","intervention":"Bonding amount 0.2-0.4mg","inputQty":30,"okQty":30,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
    {"productType":"BRS-161014","testDate":"2023-11-18","line":"","checkType":"process","variable":"Decap bond CMG+CP","variableDetail":"Bonding 0.44-0.48mg out of spec","variableGroup":"test","intervention":"Bonding amount 0.44-0.48mg","inputQty":30,"okQty":21,"ngTotal":9,"ngRate":30.0,"defectCategory":"assembly_defect","defectType":"NG over glue","defectCount":9},
    {"productType":"BRS-161014","testDate":"2023-11-16","line":"","checkType":"process","variable":"Decap bond CMG+CP","variableDetail":"","variableGroup":"test","intervention":"PT Abrasition","inputQty":24,"okQty":15,"ngTotal":9,"ngRate":37.5,"defectCategory":"assembly_defect","defectType":"NG Spread glue","defectCount":9},
    {"productType":"BRS-161014","testDate":"2023-11-16","line":"","checkType":"process","variable":"Decap bond CMG+CP","variableDetail":"","variableGroup":"test","intervention":"MG china clean primer","inputQty":16,"okQty":15,"ngTotal":1,"ngRate":6.2,"defectCategory":"assembly_defect","defectType":"NG Spread glue","defectCount":1},
    {"productType":"BRS-161014","testDate":"2023-11-16","line":"","checkType":"process","variable":"Decap bond CMG+CP","variableDetail":"","variableGroup":"test","intervention":"Adjust pressure deep more","inputQty":20,"okQty":17,"ngTotal":3,"ngRate":15.0,"defectCategory":"assembly_defect","defectType":"NG Spread glue","defectCount":3},
    {"productType":"BRS-161014","testDate":"2023-11-16","line":"","checkType":"process","variable":"Decap bond CMG+CP","variableDetail":"","variableGroup":"test","intervention":"Clean MG by alcohol","inputQty":10,"okQty":7,"ngTotal":3,"ngRate":30.0,"defectCategory":"assembly_defect","defectType":"NG Spread glue","defectCount":3},
    {"productType":"BRS-161014","testDate":"2023-11-16","line":"","checkType":"process","variable":"Decap bond CMG+CP","variableDetail":"","variableGroup":"test","intervention":"Lot MG from E2","inputQty":10,"okQty":9,"ngTotal":1,"ngRate":10.0,"defectCategory":"assembly_defect","defectType":"NG Spread glue","defectCount":1},
    {"productType":"BRS-161014","testDate":"2023-11-16","line":"","checkType":"process","variable":"Decap bond CMG+CP","variableDetail":"","variableGroup":"test","intervention":"Bond lot G06-002","inputQty":20,"okQty":20,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
    {"productType":"BRS-161014","testDate":"2023-11-17","line":"","checkType":"process","variable":"Decap bond CMG+CP","variableDetail":"","variableGroup":"test","intervention":"Separate Machine #1","inputQty":32,"okQty":24,"ngTotal":8,"ngRate":25.0,"defectCategory":"assembly_defect","defectType":"NG Spread glue","defectCount":8},
    {"productType":"BRS-161014","testDate":"2023-11-17","line":"","checkType":"process","variable":"Decap bond CMG+CP","variableDetail":"","variableGroup":"test","intervention":"Separate Machine #2","inputQty":40,"okQty":40,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
    {"productType":"BRS-161014","testDate":"2023-11-17","line":"","checkType":"process","variable":"Decap bond CMG+CP","variableDetail":"After dry oven (machine #1)","variableGroup":"test","intervention":"Decap after dry","inputQty":16,"okQty":13,"ngTotal":3,"ngRate":18.8,"defectCategory":"assembly_defect","defectType":"NG Spread glue","defectCount":3},
    {"productType":"BRS-161014","testDate":"2023-11-17","line":"","checkType":"process","variable":"Decap bond CMG+CP","variableDetail":"Loose glue 0.1mm","variableGroup":"test","intervention":"Loose glue 0.1mm","inputQty":24,"okQty":24,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
    {"productType":"BRS-161014","testDate":"2023-11-20","line":"","checkType":"process","variable":"Decap bond CMG+CP","variableDetail":"","variableGroup":"test","intervention":"PT Material sharpen","inputQty":48,"okQty":48,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
    {"productType":"BRS-161014","testDate":"2023-11-22","line":"","checkType":"process","variable":"Decap bond CMG+CP","variableDetail":"PT sharpen + MG Baotou","variableGroup":"test","intervention":"PT sharpen + MG Baotou","inputQty":16,"okQty":10,"ngTotal":6,"ngRate":37.5,"defectCategory":"assembly_defect","defectType":"NG Spread glue","defectCount":6}
  ],
  "tags":["brs-161014","sub-2-cmcp","spread-glue","root-cause-investigation","multi-arm","bond-process"],
  "reportType":"multi_arm",
  "verdict":"partial",
  "headline":"Sub 2 CMCP — Machine #2, Bond G06-002, PT sharpen, Loose glue all 0% NG; Machine #1 & PT Baotou worst (37.5%)",
  "evidence":[
    {"metric":"NG Spread glue rate","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"+37.5pp range","deltaSign":"down","note":"best 0% at multiple conditions",
     "comparisons":[
       {"label":"PT Abrasition","value":"37.5% (9/24)","n":24,"isBaseline":False,"isBest":False,"isWorst":True},
       {"label":"MG china clean primer","value":"6.2% (1/16)","n":16,"isBaseline":False,"isBest":False,"isWorst":False},
       {"label":"Adjust pressure deep more","value":"15.0% (3/20)","n":20,"isBaseline":True,"isBest":False,"isWorst":False},
       {"label":"Bond lot G06-002","value":"0.0% (0/20)","n":20,"isBaseline":False,"isBest":True,"isWorst":False},
       {"label":"Separate Machine #2","value":"0.0% (0/40)","n":40,"isBaseline":False,"isBest":True,"isWorst":False},
       {"label":"Loose glue 0.1mm","value":"0.0% (0/24)","n":24,"isBaseline":False,"isBest":True,"isWorst":False},
       {"label":"PT sharpen + MG Baotou","value":"37.5% (6/16)","n":16,"isBaseline":False,"isBest":False,"isWorst":True}
     ],"bestLabel":"Bond G06-002 / Machine #2 / Loose glue","worstLabel":"PT Abrasition / PT+MG Baotou"}
  ],
  "actions":[
    {"priority":1,"kind":"action","text":"Use Bond lot G06-002, Machine #2 setup, loosen glue 0.1mm"},
    {"priority":2,"kind":"risk","text":"Avoid Machine #1 / PT Abrasition / MG Baotou — NG 25-37%"},
    {"priority":3,"kind":"investigate","text":"Confirm Machine #1 hardware difference vs Machine #2"}
  ],
  "context":{"process":"Sub 2 CMCP spread-glue / dry-glue improvement","stage":"Decap bond CMG+CP","baselineReason":"current pressure-deep adjustment is baseline; many counter-measures screened"},
  "doeGrid":None,"trendPoints":None,
  "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":""
},
"ko":{"headline":"Sub 2 CMCP — Machine #2, Bond G06-002, PT sharpen, Loose glue 모두 NG 0%; Machine #1 및 PT Baotou 최악 (37.5%)",
"actionTexts":["Bond G06-002 로트, Machine #2 셋업, 글루 0.1mm 완화 적용","Machine #1 / PT Abrasition / MG Baotou 사용 금지 — NG 25-37%","Machine #1과 Machine #2 하드웨어 차이 확인"],
"contextProcess":"Sub 2 CMCP Spread glue / dry glue 개선","contextStage":"Decap bond CMG+CP","contextBaseline":"현재 프레셔 딥 조정이 기준; 다수 대책 스크리닝"},
"vi":{"headline":"Sub 2 CMCP — Machine #2, Bond G06-002, PT sharpen, Loose glue đều NG 0%; Machine #1 và PT Baotou tệ nhất (37.5%)",
"actionTexts":["Dùng lô Bond G06-002, setup Machine #2, nới keo 0.1mm","Tránh Machine #1 / PT Abrasition / MG Baotou — NG 25-37%","Kiểm tra khác biệt phần cứng Machine #1 vs Machine #2"],
"contextProcess":"Cải tiến trải keo / khô keo Sub 2 CMCP","contextStage":"Decap bond CMG+CP","contextBaseline":"Điều chỉnh áp lực sâu hiện hành là chuẩn; sàng lọc nhiều đối sách"}})

# ----------- DS13: NTI Compare C2 and E2 -------------
RESULTS.append({
"name": "10. BRS-161014  Report checking NTI Compare C2 and E2  20.01.2024",
"en":{
  "measurements":[
    {"productType":"BRS-161014","testDate":"2024-01-20","line":"E2","checkType":"function","variable":"NTI SPL 1kHz AVG","variableDetail":"NTI mask spec, n=10","variableGroup":"test","intervention":"E2 line samples","inputQty":10,"okQty":10,"ngTotal":0,"ngRate":0.0,"defectCategory":"function_spl","defectType":"","defectCount":0},
    {"productType":"BRS-161014","testDate":"2024-01-20","line":"C2","checkType":"function","variable":"NTI SPL 1kHz AVG","variableDetail":"NTI mask spec, n=10","variableGroup":"test","intervention":"C2 line samples","inputQty":10,"okQty":10,"ngTotal":0,"ngRate":0.0,"defectCategory":"function_spl","defectType":"","defectCount":0}
  ],
  "tags":["brs-161014","nti-acoustic","line-comparison","c2-vs-e2","spl","thd","impedance","quality-baseline"],
  "reportType":"comparison_study",
  "verdict":"no_clear_effect",
  "headline":"NTI C2 vs E2 — SPL/THD/IMP curves overlap within 0.3 dB, both lines pass mask",
  "evidence":[
    {"metric":"SPL 1kHz AVG (dB)","baselineLabel":"E2 line","baselineValue":"117.33","variantLabel":"C2 line","variantValue":"117.18","deltaText":"-0.14 dB","deltaSign":"down","note":"within instrument tolerance","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"SPL 10kHz AVG (dB)","baselineLabel":"E2 line","baselineValue":"118.42","variantLabel":"C2 line","variantValue":"117.76","deltaText":"-0.66 dB","deltaSign":"down","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
    {"metric":"NTI mask pass","baselineLabel":"E2","baselineValue":"PASSED (10/10)","variantLabel":"C2","variantValue":"PASSED (10/10)","deltaText":"+0pp","deltaSign":"no_change","note":"","comparisons":None,"bestLabel":"","worstLabel":""}
  ],
  "actions":[
    {"priority":1,"kind":"action","text":"Treat C2 and E2 as acoustically equivalent for production routing"},
    {"priority":2,"kind":"investigate","text":"Monitor 8-10 kHz region where C2 averages ~0.5 dB below E2"}
  ],
  "context":{"process":"NTI acoustic line comparison (BRS-161014 SPL/THD/IMP)","stage":"Final NTI box, 10 samples per line","baselineReason":"E2 is the reference line in this report"},
  "doeGrid":None,"trendPoints":None,
  "summary":"","keyFindings":"","purpose":"","testConditions":"","rootCause":"","decision":"","recommendedAction":""
},
"ko":{"headline":"NTI C2 vs E2 — SPL/THD/IMP 커브 0.3 dB 이내 중첩, 두 라인 모두 마스크 통과",
"actionTexts":["C2와 E2를 양산 라인 라우팅 시 음향상 동등으로 취급","C2가 E2보다 평균 약 0.5 dB 낮은 8-10 kHz 구간 모니터링"],
"contextProcess":"NTI 음향 라인 비교 (BRS-161014 SPL/THD/IMP)","contextStage":"최종 NTI 박스, 라인별 10대 샘플","contextBaseline":"본 보고서에서는 E2가 기준 라인"},
"vi":{"headline":"NTI C2 vs E2 — đường SPL/THD/IMP trùng trong 0.3 dB, cả hai line đều đạt mask",
"actionTexts":["Xem C2 và E2 tương đương về âm học khi định tuyến sản xuất","Theo dõi vùng 8-10 kHz nơi C2 thấp hơn E2 trung bình ~0.5 dB"],
"contextProcess":"So sánh âm học giữa các line NTI (BRS-161014 SPL/THD/IMP)","contextStage":"Hộp NTI cuối, 10 mẫu mỗi line","contextBaseline":"E2 là line tham chiếu trong báo cáo"}})


# ============================================================
# Commit logic
# ============================================================

def commit_one(con, name, result, tr_ko, tr_vi):
    cur = con.cursor()
    cur.execute("BEGIN")
    try:
        now = datetime.utcnow().isoformat() + "Z"
        product = result.get("productType","")
        # Try take productType from first measurement if not set at top
        if not product and result.get("measurements"):
            product = result["measurements"][0].get("productType","") or ""

        cur.execute("DELETE FROM NormalizedMeasurements WHERE DatasetName=?", (name,))
        for m in result.get("measurements", []):
            cur.execute("""
                INSERT INTO NormalizedMeasurements
                  (DatasetName, ProductType, TestDate, Line, CheckType, Variable,
                   VariableDetail, VariableGroup, Intervention, InputQty, OkQty,
                   NgTotal, NgRate, DefectCategory, DefectType, DefectCount, CreatedAt)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (name, m.get("productType") or product, m.get("testDate", ""),
                 m.get("line",""), m.get("checkType",""), m.get("variable",""),
                 m.get("variableDetail",""), m.get("variableGroup",""), m.get("intervention",""),
                 int(m.get("inputQty",0)), int(m.get("okQty",0)), int(m.get("ngTotal",0)),
                 float(m.get("ngRate",0)), m.get("defectCategory",""), m.get("defectType",""),
                 int(m.get("defectCount",0)), now))

        tags_json     = json.dumps(result.get("tags") or [], ensure_ascii=False)
        evidence_json = json.dumps(result.get("evidence") or [], ensure_ascii=False)
        actions_json  = json.dumps(result.get("actions")  or [], ensure_ascii=False)
        context_json  = json.dumps(result.get("context"), ensure_ascii=False) if result.get("context") else ""
        doe_json   = json.dumps(result.get("doeGrid"), ensure_ascii=False) if result.get("doeGrid") else ""
        trend_json = json.dumps(result.get("trendPoints"), ensure_ascii=False) if result.get("trendPoints") else ""

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
             "", "", tags_json, now,
             "", "", "", "", "",
             result.get("verdict",""), result.get("headline",""),
             evidence_json, actions_json, context_json,
             result.get("reportType",""), doe_json, trend_json))

        # Translations — actions + context
        orig_actions = result.get("actions") or []
        for lang, tr in [("ko", tr_ko), ("vi", tr_vi)]:
            if tr is None: continue
            actionTexts = tr.get("actionTexts") or []
            translated_actions = []
            for i, txt in enumerate(actionTexts):
                a0 = orig_actions[i] if i < len(orig_actions) else {}
                translated_actions.append({
                    "priority": a0.get("priority", i+1),
                    "kind": a0.get("kind","action"),
                    "text": txt
                })
            tr_actions_json = json.dumps(translated_actions, ensure_ascii=False)
            tr_context = {
                "process": tr.get("contextProcess",""),
                "stage":   tr.get("contextStage",""),
                "baselineReason": tr.get("contextBaseline","")
            }
            tr_context_json = json.dumps(tr_context, ensure_ascii=False)
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
                 "", "", "", "", "", "", "",
                 tr.get("headline",""), tr_actions_json, tr_context_json, now))

        con.commit()
        return True, None
    except Exception as e:
        con.rollback()
        return False, str(e)


def main():
    con = sqlite3.connect(DB)
    con.execute("PRAGMA busy_timeout = 30000")

    ok = 0; skip = 0; fail = 0
    fail_names = []
    skip_names = []

    for entry in RESULTS:
        name = entry["name"]
        # double-check paste presence
        row = con.execute(
            "SELECT ExtractedText FROM RawReportText WHERE DatasetName=? AND Kind='excel_paste'",
            (name,)).fetchone()
        if not row or not row[0]:
            print(f"[SKIP {name}] no excel_paste")
            skip += 1; skip_names.append(name)
            continue

        success, err = commit_one(con, name, entry["en"], entry["ko"], entry["vi"])
        if success:
            print(f"[OK {name}]")
            ok += 1
        else:
            print(f"[PARSE-FAIL {name}] {err}")
            fail += 1; fail_names.append(name)

    con.close()
    print("=== BATCH DONE ===")
    print(f"OK: {ok}  SKIP: {skip}  FAIL: {fail}")
    if skip_names: print("SKIP names:", skip_names)
    if fail_names: print("FAIL names:", fail_names)

if __name__ == "__main__":
    main()
