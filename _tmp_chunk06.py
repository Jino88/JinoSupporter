"""Chunk 06 commit script — agent-built analysis results."""
import sqlite3, json, sys
from datetime import datetime

DB = r"D:\000. MyWorks\002. DB\process-review.db"

NOW = datetime.utcnow().isoformat() + "Z"

# Each entry: dataset_name -> dict with measurements + analysis
DATA = {}

# ============================================================
# 1) 13. BRS-161014 TEST AOI Check OK, NG  — quality_log (AOI inspection)
# ============================================================
DATA["13. BRS-161014 TEST AOI Check OK, NG"] = {
    "productType": "BRS-161014",
    "reportType": "quality_log",
    "measurements": [
        {"productType":"BRS-161014","testDate":"","line":"","checkType":"AOI Bonding Frame+VP outside","variable":"AOI bonding","variableDetail":"","variableGroup":"AOI","intervention":"","inputQty":300,"okQty":298,"ngTotal":2,"ngRate":0.67,"defectCategory":"AOI","defectType":"Not enough glue / other","defectCount":2},
        {"productType":"BRS-161014","testDate":"","line":"","checkType":"AOI Laser cutting VP at Main 2","variable":"AOI laser cutting","variableDetail":"","variableGroup":"AOI","intervention":"","inputQty":300,"okQty":297,"ngTotal":3,"ngRate":1.0,"defectCategory":"AOI","defectType":"Cutting offset","defectCount":3},
    ],
    "tags": ["AOI","BRS-161014","laser_cutting","bonding"],
    "verdict": "",
    "headline": "AOI inspection log: bonding 0.67% NG (2/300), laser cutting 1.0% NG (3/300).",
    "evidence": [
        {"metric":"AOI bonding NG","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"—","deltaSign":"flat","note":"2/300 = 0.67%","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"AOI laser cutting NG","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"—","deltaSign":"flat","note":"3/300 = 1.00%","comparisons":None,"bestLabel":"","worstLabel":""},
    ],
    "actions": [],
    "context": {"process":"AOI inspection","stage":"BRS-161014 Main 2 line","baselineReason":"Daily AOI quality log; no comparison arm."},
    "doeGrid": None, "trendPoints": None,
}

# ============================================================
# 2) 12.MSU-20S15-07 DT Result check Height dimension S-MG  — quality_log (distribution)
# ============================================================
DATA["12.MSU-20S15-07 DT Result check Height dimension S-MG - Date 2025.03.28"] = {
    "productType": "MSU-20S15-07",
    "reportType": "quality_log",
    "measurements": [
        {"productType":"MSU-20S15-07","testDate":"2025-03-28","line":"","checkType":"Height S-MG Short","variable":"Height","variableDetail":"S-MG Short, Spec 0.66~0.71mm","variableGroup":"Dimension","intervention":"","inputQty":50,"okQty":50,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"MSU-20S15-07","testDate":"2025-03-28","line":"","checkType":"Height S-MG Long","variable":"Height","variableDetail":"S-MG Long, Spec 0.66~0.71mm","variableGroup":"Dimension","intervention":"","inputQty":50,"okQty":50,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
    ],
    "tags": ["MSU-20S15-07","S-MG","height","dimension"],
    "verdict": "",
    "headline": "S-MG height (Riujin) within spec 0.66~0.71mm; Short peak at 0.69mm (44%), Long peak at 0.68mm (42%).",
    "evidence": [
        {"metric":"S-MG Short distribution","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"—","deltaSign":"flat","note":"0.67=24% / 0.68=26% / 0.69=44% / 0.70=6% (n=50)","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"S-MG Long distribution","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"—","deltaSign":"flat","note":"0.67=32% / 0.68=42% / 0.69=26% (n=50)","comparisons":None,"bestLabel":"","worstLabel":""},
    ],
    "actions": [],
    "context": {"process":"Incoming dimension QA","stage":"Material S-MG (Riujin)","baselineReason":"Lot quality log; all within spec 0.66~0.71mm."},
    "doeGrid": None, "trendPoints": None,
}

# ============================================================
# 3) 13. BRS-161016 Report check VP+CD waiting long time to Led UV — multi_arm (waiting time)
# ============================================================
DATA["13. BRS-161016 Report check VP+CD waiting long time to Led UV afer bonding  28.4.2025"] = {
    "productType": "BRS-161016",
    "reportType": "multi_arm",
    "measurements": [
        {"productType":"BRS-161016","testDate":"2025-04-28","line":"","checkType":"Vision VP+CD","variable":"Wait Led UV","variableDetail":"Normal (17s)","variableGroup":"Wait time","intervention":"Waiting Led UV normal","inputQty":4,"okQty":4,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"VP+CD separate","defectCount":0},
        {"productType":"BRS-161016","testDate":"2025-04-28","line":"","checkType":"Vision VP+CD","variable":"Wait Led UV","variableDetail":"22s","variableGroup":"Wait time","intervention":"Waiting Led UV 22s","inputQty":4,"okQty":3,"ngTotal":1,"ngRate":25.0,"defectCategory":"","defectType":"VP+CD separate","defectCount":1},
        {"productType":"BRS-161016","testDate":"2025-04-28","line":"","checkType":"Vision VP+CD","variable":"Wait Led UV","variableDetail":"24s","variableGroup":"Wait time","intervention":"Waiting Led UV 24s","inputQty":4,"okQty":3,"ngTotal":1,"ngRate":25.0,"defectCategory":"","defectType":"VP+CD separate","defectCount":1},
        {"productType":"BRS-161016","testDate":"2025-04-28","line":"","checkType":"Vision VP+CD","variable":"Wait Led UV","variableDetail":"36s","variableGroup":"Wait time","intervention":"Waiting Led UV 36s","inputQty":4,"okQty":4,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"VP+CD separate","defectCount":0},
        {"productType":"BRS-161016","testDate":"2025-04-28","line":"","checkType":"Vision VP+CD","variable":"Wait Led UV","variableDetail":"60s","variableGroup":"Wait time","intervention":"Waiting Led UV 60s","inputQty":4,"okQty":2,"ngTotal":2,"ngRate":50.0,"defectCategory":"","defectType":"VP+CD separate","defectCount":2},
        {"productType":"BRS-161016","testDate":"2025-04-28","line":"","checkType":"Vision VP+CD","variable":"Wait Led UV","variableDetail":"90s","variableGroup":"Wait time","intervention":"Waiting Led UV 90s","inputQty":4,"okQty":1,"ngTotal":3,"ngRate":75.0,"defectCategory":"","defectType":"VP+CD separate","defectCount":3},
        {"productType":"BRS-161016","testDate":"2025-04-28","line":"","checkType":"Vision VP+CD","variable":"Wait Led UV","variableDetail":"120s","variableGroup":"Wait time","intervention":"Waiting Led UV 120s","inputQty":4,"okQty":2,"ngTotal":2,"ngRate":50.0,"defectCategory":"","defectType":"VP+CD separate","defectCount":2},
        {"productType":"BRS-161016","testDate":"2025-04-28","line":"","checkType":"Vision VP+CD","variable":"Wait Led UV","variableDetail":"240s","variableGroup":"Wait time","intervention":"Waiting Led UV 240s","inputQty":4,"okQty":1,"ngTotal":3,"ngRate":75.0,"defectCategory":"","defectType":"VP+CD separate","defectCount":3},
    ],
    "tags": ["BRS-161016","Led_UV","wait_time","VP_CD_separate","tension"],
    "verdict": "worsened",
    "headline": "VP+CD separation rises sharply when wait-to-LED-UV exceeds 60s, peaking at 75% (90s/240s) vs 0% normal 17s.",
    "evidence": [
        {"metric":"VP+CD separate NG rate","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"+75pp range","deltaSign":"up","note":"4 samples per arm","comparisons":[
            {"label":"17s normal","value":"0.0% (0/4)","n":4,"isBaseline":True,"isBest":True,"isWorst":False},
            {"label":"22s","value":"25.0% (1/4)","n":4,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"24s","value":"25.0% (1/4)","n":4,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"36s","value":"0.0% (0/4)","n":4,"isBaseline":False,"isBest":True,"isWorst":False},
            {"label":"60s","value":"50.0% (2/4)","n":4,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"90s","value":"75.0% (3/4)","n":4,"isBaseline":False,"isBest":False,"isWorst":True},
            {"label":"120s","value":"50.0% (2/4)","n":4,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"240s","value":"75.0% (3/4)","n":4,"isBaseline":False,"isBest":False,"isWorst":True},
        ],"bestLabel":"17s normal","worstLabel":"90s / 240s"},
        {"metric":"Tension avg (kgf)","baselineLabel":"17s normal","baselineValue":"1.354","variantLabel":"240s","variantValue":"0.748","deltaText":"-0.606","deltaSign":"down","note":"Tension drops as wait time grows","comparisons":None,"bestLabel":"","worstLabel":""},
    ],
    "actions": [
        {"priority":1,"kind":"action","text":"Limit VP+CD wait-to-LED-UV window to under 36s after bonding to keep separation NG ~0%."},
        {"priority":2,"kind":"risk","text":"At 60s+ wait, both separation NG and tension drop steeply; avoid line stops/queues at this station."},
    ],
    "context": {"process":"VP+CD bonding + LED UV cure","stage":"BRS-161016 sub-1","baselineReason":"17s 'normal' is the production-line target wait time."},
    "doeGrid": None, "trendPoints": None,
}

# ============================================================
# 4) 13. MSU-L20S15-07DT Report Test DOE improve NG weak solder
# ============================================================
DATA["13. MSU-L20S15-07DT  Report Test DOE improve  NG weak solder  date 27.5.2025"] = {
    "productType": "MSU-L20S15-07",
    "reportType": "multi_arm",
    "measurements": [
        {"productType":"MSU-L20S15-07","testDate":"2025-05-27","line":"","checkType":"Spot welding DOE","variable":"Weld time","variableDetail":"Normal 38ms (use bond pad G05-0001)","variableGroup":"Weld time","intervention":"Normal 38ms","inputQty":100,"okQty":100,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"Weak solder","defectCount":0},
        {"productType":"MSU-L20S15-07","testDate":"2025-05-27","line":"","checkType":"Spot welding DOE","variable":"Weld time","variableDetail":"34ms","variableGroup":"Weld time","intervention":"Test 34ms","inputQty":8,"okQty":4,"ngTotal":4,"ngRate":50.0,"defectCategory":"","defectType":"Weak solder","defectCount":4},
        {"productType":"MSU-L20S15-07","testDate":"2025-05-27","line":"","checkType":"Spot welding DOE","variable":"Weld time","variableDetail":"36ms","variableGroup":"Weld time","intervention":"Test 36ms","inputQty":8,"okQty":4,"ngTotal":4,"ngRate":50.0,"defectCategory":"","defectType":"Weak solder","defectCount":4},
        {"productType":"MSU-L20S15-07","testDate":"2025-05-27","line":"","checkType":"Spot welding DOE","variable":"Weld time","variableDetail":"38ms test","variableGroup":"Weld time","intervention":"Test 38ms","inputQty":8,"okQty":1,"ngTotal":7,"ngRate":87.5,"defectCategory":"","defectType":"Weak solder","defectCount":7},
        {"productType":"MSU-L20S15-07","testDate":"2025-05-27","line":"","checkType":"Spot welding DOE","variable":"Weld time","variableDetail":"40ms","variableGroup":"Weld time","intervention":"Test 40ms","inputQty":16,"okQty":11,"ngTotal":5,"ngRate":31.2,"defectCategory":"","defectType":"Weak solder","defectCount":5},
        {"productType":"MSU-L20S15-07","testDate":"2025-05-27","line":"","checkType":"Spot welding DOE","variable":"Weld time","variableDetail":"42ms","variableGroup":"Weld time","intervention":"Test 42ms","inputQty":8,"okQty":6,"ngTotal":2,"ngRate":25.0,"defectCategory":"","defectType":"Weak solder","defectCount":2},
        {"productType":"MSU-L20S15-07","testDate":"2025-05-27","line":"","checkType":"Spot welding DOE","variable":"Weld time","variableDetail":"44ms","variableGroup":"Weld time","intervention":"Test 44ms","inputQty":8,"okQty":6,"ngTotal":2,"ngRate":25.0,"defectCategory":"","defectType":"Weak solder","defectCount":2},
        {"productType":"MSU-L20S15-07","testDate":"2025-05-27","line":"","checkType":"Spot welding DOE","variable":"Weld time","variableDetail":"46ms","variableGroup":"Weld time","intervention":"Test 46ms","inputQty":8,"okQty":6,"ngTotal":2,"ngRate":25.0,"defectCategory":"","defectType":"Weak solder","defectCount":2},
        {"productType":"MSU-L20S15-07","testDate":"2025-05-27","line":"","checkType":"Spot welding DOE","variable":"Weld time","variableDetail":"48ms","variableGroup":"Weld time","intervention":"Test 48ms","inputQty":8,"okQty":5,"ngTotal":3,"ngRate":37.5,"defectCategory":"","defectType":"Weak solder","defectCount":3},
        {"productType":"MSU-L20S15-07","testDate":"2025-05-13","line":"","checkType":"Sorting C2→E2 vision check","variable":"Pre-DOE check","variableDetail":"OK C2 -> Move E2","variableGroup":"Sorting","intervention":"OK C2 sample","inputQty":50,"okQty":44,"ngTotal":6,"ngRate":12.0,"defectCategory":"","defectType":"Weak solder + Damage","defectCount":6},
        {"productType":"MSU-L20S15-07","testDate":"2025-05-13","line":"","checkType":"Sorting C2→E2 vision check","variable":"Pre-DOE check","variableDetail":"NG C2 -> Move E2","variableGroup":"Sorting","intervention":"NG C2 sample","inputQty":50,"okQty":0,"ngTotal":50,"ngRate":100.0,"defectCategory":"","defectType":"Weak solder","defectCount":50},
    ],
    "tags": ["MSU-L20S15-07","spot_welding","weld_time","weak_solder","DOE"],
    "verdict": "no_clear_effect",
    "headline": "DOE weld-time sweep 34-48ms test arms all show 25-87.5% weak-solder NG vs 0% for the original 38ms+bond-pad-G05-0001 normal.",
    "evidence": [
        {"metric":"Weak-solder NG rate","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"+87.5pp range","deltaSign":"up","note":"all test arms n=8, except 40ms n=16; normal n=100","comparisons":[
            {"label":"Normal 38ms+G05-0001","value":"0.0% (0/100)","n":100,"isBaseline":True,"isBest":True,"isWorst":False},
            {"label":"34ms","value":"50.0% (4/8)","n":8,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"36ms","value":"50.0% (4/8)","n":8,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"38ms test","value":"87.5% (7/8)","n":8,"isBaseline":False,"isBest":False,"isWorst":True},
            {"label":"40ms","value":"31.2% (5/16)","n":16,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"42ms","value":"25.0% (2/8)","n":8,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"44ms","value":"25.0% (2/8)","n":8,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"46ms","value":"25.0% (2/8)","n":8,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"48ms","value":"37.5% (3/8)","n":8,"isBaseline":False,"isBest":False,"isWorst":False},
        ],"bestLabel":"Normal 38ms+G05-0001","worstLabel":"38ms test"},
    ],
    "actions": [
        {"priority":1,"kind":"investigate","text":"Bond-pad spot G05-0001 (used only in normal) is the dominant factor; rerun DOE on weld-time keeping the same pad to isolate time effect."},
        {"priority":2,"kind":"action","text":"Do not roll out any new weld-time setting until a controlled comparison with G05-0001 pad is completed."},
    ],
    "context": {"process":"Spot welding (Weld 1&2 time)","stage":"MSU-L20S15-07 line, E2 verification","baselineReason":"Normal arm uses production bond-pad G05-0001 at 38ms — current working setpoint."},
    "doeGrid": None, "trendPoints": None,
}

# ============================================================
# 5) 13. BRS-161014 Report check PT measure coplanarty  — comparison_study
# ============================================================
DATA["13. BRS-161014  Report check PT measure coplanarty 2023.11.22"] = {
    "productType": "BRS-161014",
    "reportType": "comparison_study",
    "measurements": [
        {"productType":"BRS-161014","testDate":"2023-11-22","line":"","checkType":"Decap bond CMG+CP","variable":"PT coplanarity","variableDetail":"<0.01mm","variableGroup":"PT flatness","intervention":"PT <0.01mm","inputQty":16,"okQty":16,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"Spread glue NG","defectCount":0},
        {"productType":"BRS-161014","testDate":"2023-11-22","line":"","checkType":"Decap bond CMG+CP","variable":"PT coplanarity","variableDetail":"0.01~0.02mm","variableGroup":"PT flatness","intervention":"PT 0.01~0.02mm","inputQty":16,"okQty":6,"ngTotal":10,"ngRate":62.5,"defectCategory":"","defectType":"Spread glue NG","defectCount":10},
        {"productType":"BRS-161014","testDate":"2023-11-23","line":"","checkType":"Decap bond CMG+CP","variable":"PT coplanarity","variableDetail":"<0.01mm","variableGroup":"PT flatness","intervention":"PT <0.01mm","inputQty":16,"okQty":16,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"Spread glue NG","defectCount":0},
        {"productType":"BRS-161014","testDate":"2023-11-23","line":"","checkType":"Decap bond CMG+CP","variable":"PT coplanarity","variableDetail":"0.01~0.02mm","variableGroup":"PT flatness","intervention":"PT 0.01~0.02mm","inputQty":16,"okQty":11,"ngTotal":5,"ngRate":31.2,"defectCategory":"","defectType":"Spread glue NG","defectCount":5},
    ],
    "tags": ["BRS-161014","PT","coplanarity","spread_glue","decap"],
    "verdict": "improved",
    "headline": "PT coplanarity <0.01mm gives 0% spread-glue NG vs 31-62% for 0.01~0.02mm group.",
    "evidence": [
        {"metric":"Decap spread-glue NG (22 Nov)","baselineLabel":"PT 0.01~0.02mm","baselineValue":"62.5% (10/16)","variantLabel":"PT <0.01mm","variantValue":"0.0% (0/16)","deltaText":"-62.5pp","deltaSign":"down","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
            {"metric":"Decap spread-glue NG (23 Nov)","baselineLabel":"PT 0.01~0.02mm","baselineValue":"31.2% (5/16)","variantLabel":"PT <0.01mm","variantValue":"0.0% (0/16)","deltaText":"-31.2pp","deltaSign":"down","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
    ],
    "actions": [
        {"priority":1,"kind":"action","text":"Restrict PT supply to coplanarity <0.01mm for decap bonding; reject 0.01~0.02mm lots."},
    ],
    "context": {"process":"Decap bond CMG+CP","stage":"BRS-161014 sub-2","baselineReason":"PT 0.01~0.02mm represents the wider incoming tolerance; <0.01mm is the tighter target."},
    "doeGrid": None, "trendPoints": None,
}

# ============================================================
# 6) 13. BRS-161014 Report test tracking 1k material  — quality_log (process tracking)
# ============================================================
DATA["13. BRS-161014 Report test tracking 1k material 01.09.2023"] = {
    "productType": "BRS-161014",
    "reportType": "quality_log",
    "measurements": [
        {"productType":"BRS-161014","testDate":"2023-09-01","line":"Sub 1","checkType":"Vision VP","variable":"Tracking 1k","variableDetail":"","variableGroup":"Tracking","intervention":"","inputQty":1000,"okQty":997,"ngTotal":3,"ngRate":0.30,"defectCategory":"","defectType":"VP NG / laser cutting offset","defectCount":3},
        {"productType":"BRS-161014","testDate":"2023-09-01","line":"Sub 1","checkType":"Array CD 2","variable":"Tracking 1k","variableDetail":"","variableGroup":"Tracking","intervention":"","inputQty":1000,"okQty":994,"ngTotal":6,"ngRate":0.60,"defectCategory":"","defectType":"Machine pick-up NG","defectCount":6},
        {"productType":"BRS-161014","testDate":"2023-09-01","line":"Sub 1","checkType":"AOI VP+CD","variable":"Tracking 1k","variableDetail":"","variableGroup":"Tracking","intervention":"","inputQty":994,"okQty":980,"ngTotal":14,"ngRate":1.41,"defectCategory":"","defectType":"Glue/Not enough glue","defectCount":14},
        {"productType":"BRS-161014","testDate":"2023-09-01","line":"Main 2","checkType":"AOI laser cutting","variable":"Tracking 1k","variableDetail":"","variableGroup":"Tracking","intervention":"","inputQty":478,"okQty":464,"ngTotal":14,"ngRate":2.93,"defectCategory":"","defectType":"Cutting offset / damage","defectCount":14},
        {"productType":"BRS-161014","testDate":"2023-09-05","line":"Main 2","checkType":"AOI laser cutting","variable":"Tracking 1k","variableDetail":"","variableGroup":"Tracking","intervention":"","inputQty":494,"okQty":483,"ngTotal":11,"ngRate":2.23,"defectCategory":"","defectType":"Laser cutting NG","defectCount":11},
        {"productType":"BRS-161014","testDate":"2023-09-06","line":"SUB 4","checkType":"Air leak","variable":"Tracking 1k","variableDetail":"","variableGroup":"Tracking","intervention":"","inputQty":971,"okQty":937,"ngTotal":34,"ngRate":3.50,"defectCategory":"","defectType":"VP+CD separate (glue)","defectCount":34},
        {"productType":"BRS-161014","testDate":"2023-09-06","line":"SUB 4","checkType":"Visual Long VP","variable":"Tracking 1k","variableDetail":"","variableGroup":"Tracking","intervention":"","inputQty":937,"okQty":862,"ngTotal":75,"ngRate":8.00,"defectCategory":"","defectType":"VP damage / not enough glue","defectCount":75},
        {"productType":"BRS-161014","testDate":"2023-09-06","line":"SUB 4","checkType":"Visual Short VP","variable":"Tracking 1k","variableDetail":"","variableGroup":"Tracking","intervention":"","inputQty":862,"okQty":776,"ngTotal":86,"ngRate":9.98,"defectCategory":"","defectType":"VP damage / not enough glue","defectCount":86},
        {"productType":"BRS-161014","testDate":"2023-09-06","line":"SUB 4","checkType":"Visual FS","variable":"Tracking 1k","variableDetail":"","variableGroup":"Tracking","intervention":"","inputQty":769,"okQty":727,"ngTotal":42,"ngRate":5.46,"defectCategory":"","defectType":"FS damage (laser cutting)","defectCount":42},
        {"productType":"BRS-161014","testDate":"2023-09-06","line":"Function","checkType":"Function final","variable":"Tracking 1k","variableDetail":"","variableGroup":"Function","intervention":"","inputQty":727,"okQty":381,"ngTotal":409,"ngRate":56.3,"defectCategory":"Function","defectType":"Hearing Noise/Touch/HOHD","defectCount":409},
    ],
    "tags": ["BRS-161014","tracking","1k","process_NG","function"],
    "verdict": "",
    "headline": "1k-material end-to-end tracking: sub-1 NG <1.5%, sub-4 visual stacks 5-10%, function final 56.3% NG (dominated by Hearing Noise 58.7%).",
    "evidence": [
        {"metric":"Sub-4 Visual Short VP NG","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"—","deltaSign":"flat","note":"86/862 = 9.98%","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Sub-4 Air leak (VP+CD separate)","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"—","deltaSign":"flat","note":"34/971 = 3.50% (not enough glue)","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Function final total NG","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"—","deltaSign":"flat","note":"409/727 = 56.3%; Noise 240, Touch 82, HOHD 60","comparisons":None,"bestLabel":"","worstLabel":""},
    ],
    "actions": [
        {"priority":1,"kind":"investigate","text":"Sub-4 Visual VP losses 8-10% from VP damage + not-enough-glue dominate yield drop; review SUB-4 ass'y settings."},
        {"priority":2,"kind":"investigate","text":"Hearing Noise alone causes 240 NG (58.7%); decap study needed at function station."},
    ],
    "context": {"process":"End-to-end 1k-piece process tracking","stage":"Sub-1 → Sub-2 → Main → Sub-3 → Sub-4 → Function","baselineReason":"Single tracked lot to map cumulative yield loss across stations."},
    "doeGrid": None, "trendPoints": None,
}

# ============================================================
# 7) 13. BRS-161014 DT Report test VP bending 15.10.2024  — multi_arm (VP mold)
# ============================================================
DATA["13. BRS-161014  DT Report test VP bending date 15.10.2024"] = {
    "productType": "BRS-161016",
    "reportType": "multi_arm",
    "measurements": [
        {"productType":"BRS-161016","testDate":"2024-10-15","line":"C2-3A","checkType":"Vision VP/CD","variable":"VP mold","variableDetail":"VP #3 bending","variableGroup":"VP mold","intervention":"VP bending #3","inputQty":75,"okQty":74,"ngTotal":1,"ngRate":1.33,"defectCategory":"","defectType":"VP separate","defectCount":1},
        {"productType":"BRS-161016","testDate":"2024-10-15","line":"C2-3A","checkType":"Vision VP/CD","variable":"VP mold","variableDetail":"VP #4 bending","variableGroup":"VP mold","intervention":"VP bending #4","inputQty":96,"okQty":96,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161016","testDate":"2024-10-15","line":"C2-3A","checkType":"Vision VP/CD","variable":"VP mold","variableDetail":"VP #6 bending","variableGroup":"VP mold","intervention":"VP bending #6","inputQty":16,"okQty":16,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161016","testDate":"2024-10-15","line":"C2-3A","checkType":"Vision VP/CD","variable":"VP mold","variableDetail":"VP #9 bending","variableGroup":"VP mold","intervention":"VP bending #9","inputQty":96,"okQty":95,"ngTotal":1,"ngRate":1.0,"defectCategory":"","defectType":"VP separate","defectCount":1},
        {"productType":"BRS-161016","testDate":"2024-10-15","line":"C2-3A","checkType":"Vision VP/CD","variable":"VP mold","variableDetail":"VP #6 normal","variableGroup":"VP mold","intervention":"VP OK normal #6","inputQty":1000,"okQty":999,"ngTotal":1,"ngRate":0.1,"defectCategory":"","defectType":"VP separate","defectCount":1},
        {"productType":"BRS-161016","testDate":"2024-10-15","line":"C2-3A","checkType":"Function","variable":"VP mold","variableDetail":"VP #3","variableGroup":"VP mold","intervention":"VP bending #3 function","inputQty":72,"okQty":70,"ngTotal":2,"ngRate":2.8,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":2},
        {"productType":"BRS-161016","testDate":"2024-10-15","line":"C2-3A","checkType":"Function","variable":"VP mold","variableDetail":"VP #4","variableGroup":"VP mold","intervention":"VP bending #4 function","inputQty":96,"okQty":88,"ngTotal":8,"ngRate":8.3,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":8},
        {"productType":"BRS-161016","testDate":"2024-10-15","line":"C2-3A","checkType":"Function","variable":"VP mold","variableDetail":"VP #6","variableGroup":"VP mold","intervention":"VP bending #6 function","inputQty":16,"okQty":16,"ngTotal":0,"ngRate":0.0,"defectCategory":"Function","defectType":"","defectCount":0},
        {"productType":"BRS-161016","testDate":"2024-10-15","line":"C2-3A","checkType":"Function","variable":"VP mold","variableDetail":"VP #9","variableGroup":"VP mold","intervention":"VP bending #9 function","inputQty":95,"okQty":90,"ngTotal":5,"ngRate":5.3,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":5},
        {"productType":"BRS-161016","testDate":"2024-10-15","line":"C2-3A","checkType":"Function","variable":"VP mold","variableDetail":"VP #6 normal","variableGroup":"VP mold","intervention":"VP OK normal #6 function","inputQty":560,"okQty":546,"ngTotal":14,"ngRate":2.5,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":14},
    ],
    "tags": ["BRS-161016","VP_bending","VP_mold","function","hearing"],
    "verdict": "worsened",
    "headline": "VP-bending function NG 2.8-8.3% vs 2.5% normal (#6); VP #4 worst at 8.3% — bending shows potential VP/CD separation risk.",
    "evidence": [
        {"metric":"Function NG by VP mold","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"+5.8pp range","deltaSign":"up","note":"bending lot vs normal #6","comparisons":[
            {"label":"VP #6 normal","value":"2.5% (14/560)","n":560,"isBaseline":True,"isBest":False,"isWorst":False},
            {"label":"VP #3 bending","value":"2.8% (2/72)","n":72,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"VP #4 bending","value":"8.3% (8/96)","n":96,"isBaseline":False,"isBest":False,"isWorst":True},
            {"label":"VP #6 bending","value":"0.0% (0/16)","n":16,"isBaseline":False,"isBest":True,"isWorst":False},
            {"label":"VP #9 bending","value":"5.3% (5/95)","n":95,"isBaseline":False,"isBest":False,"isWorst":False},
        ],"bestLabel":"VP #6 bending","worstLabel":"VP #4 bending"},
        {"metric":"Vision VP/CD NG (bending lot)","baselineLabel":"VP #6 normal","baselineValue":"0.1% (1/1000)","variantLabel":"All bending","variantValue":"0.7% (2/283)","deltaText":"+0.6pp","deltaSign":"up","note":"VP separate type","comparisons":None,"bestLabel":"","worstLabel":""},
    ],
    "actions": [
        {"priority":1,"kind":"risk","text":"VP #4 bending lot shows 8.3% hearing-noise NG — withhold from mass production."},
        {"priority":2,"kind":"investigate","text":"Decap NG sample: #3=CM offset, #4=Frame+yoke offset 12.5%, #9=Frame+yoke 20% + unknown 60%."},
    ],
    "context": {"process":"VP-bending mold qualification","stage":"BRS-161016 sub-1 + function","baselineReason":"VP #6 normal-bending mold is the current mass-production reference."},
    "doeGrid": None, "trendPoints": None,
}

# ============================================================
# 8) 12. BRS-161016 GMI Test VP all mold (block)  — multi_arm (VP mold)
# ============================================================
DATA["12.BRS-161016 GMI Test VP all mold block date 4.9.2025"] = {
    "productType": "BRS-161016",
    "reportType": "multi_arm",
    "measurements": [
        {"productType":"BRS-161016","testDate":"2025-09-04","line":"E2-3B","checkType":"VP bending check (sub 1)","variable":"VP mold","variableDetail":"VP #9","variableGroup":"VP mold","intervention":"VP #9 block","inputQty":1200,"okQty":943,"ngTotal":257,"ngRate":21.4,"defectCategory":"","defectType":"VP bending","defectCount":257},
        {"productType":"BRS-161016","testDate":"2025-09-04","line":"E2-3B","checkType":"VP bending check (sub 1)","variable":"VP mold","variableDetail":"VP #11","variableGroup":"VP mold","intervention":"VP #11 block","inputQty":1200,"okQty":1085,"ngTotal":115,"ngRate":9.6,"defectCategory":"","defectType":"VP bending","defectCount":115},
        {"productType":"BRS-161016","testDate":"2025-09-04","line":"E2-3B","checkType":"VP bending check (sub 1)","variable":"VP mold","variableDetail":"VP #5","variableGroup":"VP mold","intervention":"VP #5 block","inputQty":1200,"okQty":1134,"ngTotal":66,"ngRate":5.5,"defectCategory":"","defectType":"VP bending","defectCount":66},
        {"productType":"BRS-161016","testDate":"2025-09-04","line":"E2-3B","checkType":"VP bending check (sub 1)","variable":"VP mold","variableDetail":"VP #8","variableGroup":"VP mold","intervention":"VP #8 block","inputQty":1200,"okQty":1008,"ngTotal":192,"ngRate":16.0,"defectCategory":"","defectType":"VP bending","defectCount":192},
        {"productType":"BRS-161016","testDate":"2025-09-04","line":"E2-3B","checkType":"VP bending check (sub 1)","variable":"VP mold","variableDetail":"VP #10","variableGroup":"VP mold","intervention":"VP #10 block","inputQty":1200,"okQty":1125,"ngTotal":75,"ngRate":6.2,"defectCategory":"","defectType":"VP bending","defectCount":75},
        {"productType":"BRS-161016","testDate":"2025-09-04","line":"E2-3B","checkType":"VP bending check (sub 1)","variable":"VP mold","variableDetail":"Normal #6","variableGroup":"VP mold","intervention":"Normal #6","inputQty":800,"okQty":797,"ngTotal":3,"ngRate":0.4,"defectCategory":"","defectType":"VP bending","defectCount":3},
        {"productType":"BRS-161016","testDate":"2025-09-04","line":"E2-3B","checkType":"Function","variable":"VP mold","variableDetail":"VP #9","variableGroup":"VP mold","intervention":"VP #9 function","inputQty":973,"okQty":895,"ngTotal":78,"ngRate":8.0,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":78},
        {"productType":"BRS-161016","testDate":"2025-09-04","line":"E2-3B","checkType":"Function","variable":"VP mold","variableDetail":"VP #11","variableGroup":"VP mold","intervention":"VP #11 function","inputQty":1083,"okQty":1007,"ngTotal":76,"ngRate":7.0,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":76},
        {"productType":"BRS-161016","testDate":"2025-09-04","line":"E2-3B","checkType":"Function","variable":"VP mold","variableDetail":"VP #5","variableGroup":"VP mold","intervention":"VP #5 function","inputQty":1078,"okQty":1014,"ngTotal":64,"ngRate":5.9,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":64},
        {"productType":"BRS-161016","testDate":"2025-09-04","line":"E2-3B","checkType":"Function","variable":"VP mold","variableDetail":"VP #8","variableGroup":"VP mold","intervention":"VP #8 function","inputQty":1003,"okQty":945,"ngTotal":58,"ngRate":5.8,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":58},
        {"productType":"BRS-161016","testDate":"2025-09-04","line":"E2-3B","checkType":"Function","variable":"VP mold","variableDetail":"VP #10","variableGroup":"VP mold","intervention":"VP #10 function","inputQty":1143,"okQty":1080,"ngTotal":63,"ngRate":5.5,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":63},
        {"productType":"BRS-161016","testDate":"2025-09-04","line":"E2-3B","checkType":"Function","variable":"VP mold","variableDetail":"Normal #6","variableGroup":"VP mold","intervention":"Normal #6 function","inputQty":800,"okQty":779,"ngTotal":21,"ngRate":2.6,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":21},
    ],
    "tags": ["BRS-161016","VP_mold","GMI","VP_bending","function"],
    "verdict": "worsened",
    "headline": "All test VP molds (#5/#8/#9/#10/#11) exceed normal #6 in bending NG (5.5-21.4% vs 0.4%) and function NG (5.5-8.0% vs 2.6%).",
    "evidence": [
        {"metric":"VP bending NG (sub 1, 4 Sep)","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"+21pp range","deltaSign":"up","note":"E2-3B day 1","comparisons":[
            {"label":"Normal #6","value":"0.4% (3/800)","n":800,"isBaseline":True,"isBest":True,"isWorst":False},
            {"label":"VP #5","value":"5.5% (66/1200)","n":1200,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"VP #8","value":"16.0% (192/1200)","n":1200,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"VP #9","value":"21.4% (257/1200)","n":1200,"isBaseline":False,"isBest":False,"isWorst":True},
            {"label":"VP #10","value":"6.2% (75/1200)","n":1200,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"VP #11","value":"9.6% (115/1200)","n":1200,"isBaseline":False,"isBest":False,"isWorst":False},
        ],"bestLabel":"Normal #6","worstLabel":"VP #9"},
        {"metric":"Function NG (4 Sep)","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"+5.4pp range","deltaSign":"up","note":"hearing noise/touch dominates","comparisons":[
            {"label":"Normal #6","value":"2.6% (21/800)","n":800,"isBaseline":True,"isBest":True,"isWorst":False},
            {"label":"VP #5","value":"5.9% (64/1078)","n":1078,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"VP #8","value":"5.8% (58/1003)","n":1003,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"VP #9","value":"8.0% (78/973)","n":973,"isBaseline":False,"isBest":False,"isWorst":True},
            {"label":"VP #10","value":"5.5% (63/1143)","n":1143,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"VP #11","value":"7.0% (76/1083)","n":1083,"isBaseline":False,"isBest":False,"isWorst":False},
        ],"bestLabel":"Normal #6","worstLabel":"VP #9"},
    ],
    "actions": [
        {"priority":1,"kind":"action","text":"Continue using normal VP #6 lot; block VP #9 / #8 from mass production."},
        {"priority":2,"kind":"investigate","text":"VP #5 / #10 are closest to normal — recheck mold dimensions vs #6 to find acceptable mold candidates."},
    ],
    "context": {"process":"VP mold qualification (GMI VP-all-mold)","stage":"BRS-161016 line E2-3B / C2-3A","baselineReason":"Normal #6 from lot IR250721007 is the current production reference mold."},
    "doeGrid": None, "trendPoints": None,
}

# ============================================================
# 9) 13. BRS-161016 DT Report check Reason NG Low gauss  — comparison_study (vender lot)
# ============================================================
DATA["13. BRS-161016 DT  Report  check Reason NG Low gauss  date 12.7.2025"] = {
    "productType": "BRS-161016",
    "reportType": "comparison_study",
    "measurements": [
        {"productType":"BRS-161016","testDate":"2025-07-10","line":"C2-3A","checkType":"Low gauss check","variable":"MG-C vender lot","variableDetail":"Beijing (line lot)","variableGroup":"Vender","intervention":"MG-C Beijing C2-3A","inputQty":6149,"okQty":5846,"ngTotal":303,"ngRate":4.9,"defectCategory":"","defectType":"Low gauss","defectCount":303},
        {"productType":"BRS-161016","testDate":"2025-07-10","line":"C2-5A","checkType":"Low gauss check","variable":"MG-C vender lot","variableDetail":"Beijing (line lot)","variableGroup":"Vender","intervention":"MG-C Beijing C2-5A","inputQty":5771,"okQty":5475,"ngTotal":296,"ngRate":5.1,"defectCategory":"","defectType":"Low gauss","defectCount":296},
        {"productType":"BRS-161016","testDate":"2025-07-10","line":"","checkType":"Low gauss check (Test sample)","variable":"MG-C box","variableDetail":"Beijing Box 1","variableGroup":"Vender","intervention":"MG-C Beijing Box 1","inputQty":128,"okQty":87,"ngTotal":41,"ngRate":32.0,"defectCategory":"","defectType":"Low gauss","defectCount":41},
        {"productType":"BRS-161016","testDate":"2025-07-10","line":"","checkType":"Low gauss check (Test sample)","variable":"MG-C box","variableDetail":"Beijing Box 2","variableGroup":"Vender","intervention":"MG-C Beijing Box 2","inputQty":128,"okQty":128,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"Low gauss","defectCount":0},
        {"productType":"BRS-161016","testDate":"2025-07-10","line":"","checkType":"Low gauss check (Test sample)","variable":"MG-C box","variableDetail":"Beijing Box 3","variableGroup":"Vender","intervention":"MG-C Beijing Box 3","inputQty":127,"okQty":127,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"Low gauss","defectCount":0},
        {"productType":"BRS-161016","testDate":"2025-07-10","line":"","checkType":"Low gauss check (Test sample)","variable":"MG-C box","variableDetail":"Ruijin (Normal)","variableGroup":"Vender","intervention":"MG-C Ruijin Normal","inputQty":128,"okQty":128,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"Low gauss","defectCount":0},
        {"productType":"BRS-161016","testDate":"2025-07-16","line":"","checkType":"Low gauss check (2nd lot)","variable":"MG-C box","variableDetail":"Beijing Box 1 (IR250703013)","variableGroup":"Vender","intervention":"MG-C Beijing Box 1 (2nd)","inputQty":128,"okQty":127,"ngTotal":1,"ngRate":0.8,"defectCategory":"","defectType":"Low gauss","defectCount":1},
        {"productType":"BRS-161016","testDate":"2025-07-16","line":"","checkType":"Low gauss check (2nd lot)","variable":"MG-C box","variableDetail":"Beijing Box 2 (IR250703013)","variableGroup":"Vender","intervention":"MG-C Beijing Box 2 (2nd)","inputQty":128,"okQty":118,"ngTotal":10,"ngRate":7.8,"defectCategory":"","defectType":"Low gauss","defectCount":10},
    ],
    "tags": ["BRS-161016","MG-C","low_gauss","vender","Beijing","Ruijin"],
    "verdict": "worsened",
    "headline": "MG-C Beijing Box 1 (1st lot) hits 32% low-gauss NG; Ruijin normal stays 0% — Beijing semi-yoke avg gauss 484~514 vs Ruijin 525.",
    "evidence": [
        {"metric":"Low-gauss NG (test sample, 10 Jul)","baselineLabel":"Ruijin Normal","baselineValue":"0.0% (0/128)","variantLabel":"Beijing Box 1","variantValue":"32.0% (41/128)","deltaText":"+32pp","deltaSign":"up","note":"Box 2/3 = 0%; only Box 1 fails","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Line low-gauss NG (10 Jul)","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"—","deltaSign":"flat","note":"C2-3A 4.9% (303/6149), C2-5A 5.1% (296/5771)","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Semi-yoke avg gauss","baselineLabel":"Ruijin Normal","baselineValue":"525 G","variantLabel":"Beijing Box 1","variantValue":"484 G","deltaText":"-41G","deltaSign":"down","note":"Spec >= 480G; Box 2=514, Box 3=514","comparisons":None,"bestLabel":"","worstLabel":""},
    ],
    "actions": [
        {"priority":1,"kind":"action","text":"Quarantine MG-C Beijing Box 1 (lot IR250623044/IR250703013); only release boxes with avg gauss >= 510."},
        {"priority":2,"kind":"investigate","text":"Investigate Beijing vender process control on semi-yoke gauss — between-box variation 484~514 too wide."},
    ],
    "context": {"process":"MG-C low-gauss inspection","stage":"BRS-161016 lines C2-3A / C2-5A","baselineReason":"Ruijin lot IR250721007 is the production-line normal vender."},
    "doeGrid": None, "trendPoints": None,
}

# ============================================================
# 10) 13. BRS-201506 Report test New BP+SM Ass'y guide JIG  — comparison_study
# ============================================================
DATA["13. BRS-201506 Report test New BP+SM Ass'y guide JIG date 5.2.2024"] = {
    "productType": "BRS-201506",
    "reportType": "comparison_study",
    "measurements": [
        {"productType":"BRS-201506","testDate":"2024-02-03","line":"","checkType":"SM vision","variable":"BP+SM JIG","variableDetail":"New JIG","variableGroup":"JIG","intervention":"New JIG","inputQty":300,"okQty":300,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"SM offset","defectCount":0},
        {"productType":"BRS-201506","testDate":"2024-02-03","line":"","checkType":"SM vision","variable":"BP+SM JIG","variableDetail":"Normal","variableGroup":"JIG","intervention":"Normal line","inputQty":300,"okQty":300,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"SM offset","defectCount":0},
        {"productType":"BRS-201506","testDate":"2024-04-08","line":"","checkType":"SM vision","variable":"BP+SM JIG","variableDetail":"New JIG","variableGroup":"JIG","intervention":"New JIG","inputQty":306,"okQty":306,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"SM offset","defectCount":0},
        {"productType":"BRS-201506","testDate":"2024-04-10","line":"","checkType":"SM vision","variable":"BP+SM JIG","variableDetail":"New JIG","variableGroup":"JIG","intervention":"New JIG","inputQty":300,"okQty":300,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"SM offset","defectCount":0},
        {"productType":"BRS-201506","testDate":"2024-02-05","line":"","checkType":"Function","variable":"BP+SM JIG","variableDetail":"New JIG","variableGroup":"JIG","intervention":"New JIG","inputQty":292,"okQty":283,"ngTotal":9,"ngRate":3.1,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":9},
        {"productType":"BRS-201506","testDate":"2024-02-05","line":"","checkType":"Function","variable":"BP+SM JIG","variableDetail":"Normal","variableGroup":"JIG","intervention":"Normal line","inputQty":795,"okQty":754,"ngTotal":41,"ngRate":5.2,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":41},
        {"productType":"BRS-201506","testDate":"2024-02-15","line":"","checkType":"Function","variable":"BP+SM JIG","variableDetail":"New JIG","variableGroup":"JIG","intervention":"New JIG","inputQty":427,"okQty":414,"ngTotal":13,"ngRate":3.0,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":13},
        {"productType":"BRS-201506","testDate":"2024-02-15","line":"","checkType":"Function","variable":"BP+SM JIG","variableDetail":"Normal","variableGroup":"JIG","intervention":"Normal line","inputQty":794,"okQty":768,"ngTotal":26,"ngRate":3.3,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":26},
        {"productType":"BRS-201506","testDate":"2024-04-09","line":"","checkType":"Function","variable":"BP+SM JIG","variableDetail":"New JIG","variableGroup":"JIG","intervention":"New JIG","inputQty":303,"okQty":301,"ngTotal":2,"ngRate":0.7,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":2},
        {"productType":"BRS-201506","testDate":"2024-04-09","line":"","checkType":"Function","variable":"BP+SM JIG","variableDetail":"Normal","variableGroup":"JIG","intervention":"Normal line","inputQty":299,"okQty":294,"ngTotal":5,"ngRate":1.7,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":5},
        {"productType":"BRS-201506","testDate":"2024-04-11","line":"","checkType":"Function","variable":"BP+SM JIG","variableDetail":"New JIG","variableGroup":"JIG","intervention":"New JIG","inputQty":299,"okQty":285,"ngTotal":14,"ngRate":4.7,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":14},
        {"productType":"BRS-201506","testDate":"2024-04-11","line":"","checkType":"Function","variable":"BP+SM JIG","variableDetail":"Normal","variableGroup":"JIG","intervention":"Normal line","inputQty":304,"okQty":290,"ngTotal":14,"ngRate":4.6,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":14},
    ],
    "tags": ["BRS-201506","BP_SM","JIG","SM_offset","function"],
    "verdict": "no_clear_effect",
    "headline": "New BP+SM ass'y guide JIG (+0.02 out) keeps SM-vision NG at 0% and function NG mixed: 4 dates show new vs normal at 3.0/5.2, 3.0/3.3, 0.7/1.7, 4.7/4.6%.",
    "evidence": [
        {"metric":"SM offset NG (vision)","baselineLabel":"Normal","baselineValue":"0.0%","variantLabel":"New JIG","variantValue":"0.0%","deltaText":"0pp","deltaSign":"flat","note":"3 dates, both arms zero","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Function NG (5 Feb)","baselineLabel":"Normal","baselineValue":"5.2% (41/795)","variantLabel":"New JIG","variantValue":"3.1% (9/292)","deltaText":"-2.1pp","deltaSign":"down","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Function NG (11 Apr)","baselineLabel":"Normal","baselineValue":"4.6% (14/304)","variantLabel":"New JIG","variantValue":"4.7% (14/299)","deltaText":"+0.1pp","deltaSign":"up","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
    ],
    "actions": [
        {"priority":1,"kind":"investigate","text":"Function NG difference inconsistent across 4 trials; collect larger sample before adopting new JIG."},
        {"priority":2,"kind":"action","text":"YK 201506 dimension drawing changed to 19.50/+0.1 and 14.50/+0.1 to absorb the 0.02 JIG-gap effect."},
    ],
    "context": {"process":"BP+SM ass'y","stage":"BRS-201506 sub","baselineReason":"Normal-line JIG is the current setup; new JIG moves out 0.02 to fix SMG spread-glue partial NG."},
    "doeGrid": None, "trendPoints": None,
}

# ============================================================
# 11) 13. BRS-161016 Report Test PT 161014-S of Press line (Doojin coating) NG dimension — comparison_study
# ============================================================
DATA["13. BRS-161016 Report Test PT 161014-S of Press line (Doojin coating) happen  NG dimension 30.9.2025"] = {
    "productType": "BRS-161016",
    "reportType": "comparison_study",
    "measurements": [
        {"productType":"BRS-161016","testDate":"2025-09-29","line":"E2-3B","checkType":"AOI laser cutting","variable":"PT 161014-S Frame","variableDetail":"Test (NG-dimension Doojin)","variableGroup":"Frame source","intervention":"Test Frame (29 Sep)","inputQty":2900,"okQty":2322,"ngTotal":578,"ngRate":19.9,"defectCategory":"","defectType":"AOI laser cutting NG","defectCount":578},
        {"productType":"BRS-161016","testDate":"2025-09-29","line":"E2-3B","checkType":"AOI laser cutting","variable":"PT 161014-S Frame","variableDetail":"Normal","variableGroup":"Frame source","intervention":"Normal Frame (29 Sep)","inputQty":500,"okQty":496,"ngTotal":4,"ngRate":0.8,"defectCategory":"","defectType":"AOI laser cutting NG","defectCount":4},
        {"productType":"BRS-161016","testDate":"2025-10-21","line":"E2-3A","checkType":"AOI laser cutting","variable":"PT 161014-S Frame","variableDetail":"Test (NG-dimension Doojin)","variableGroup":"Frame source","intervention":"Test Frame (21 Oct)","inputQty":855,"okQty":757,"ngTotal":98,"ngRate":11.5,"defectCategory":"","defectType":"AOI laser cutting NG","defectCount":98},
        {"productType":"BRS-161016","testDate":"2025-10-21","line":"E2-3A","checkType":"AOI laser cutting","variable":"PT 161014-S Frame","variableDetail":"Normal","variableGroup":"Frame source","intervention":"Normal Frame (21 Oct)","inputQty":1000,"okQty":985,"ngTotal":15,"ngRate":1.5,"defectCategory":"","defectType":"AOI laser cutting NG","defectCount":15},
        {"productType":"BRS-161016","testDate":"2025-09-30","line":"E2-3B","checkType":"Height dimension","variable":"PT 161014-S Frame","variableDetail":"Test","variableGroup":"Frame source","intervention":"Test Frame height","inputQty":2718,"okQty":2684,"ngTotal":34,"ngRate":1.3,"defectCategory":"","defectType":"Height NG","defectCount":34},
        {"productType":"BRS-161016","testDate":"2025-09-30","line":"E2-3B","checkType":"Height dimension","variable":"PT 161014-S Frame","variableDetail":"Normal","variableGroup":"Frame source","intervention":"Normal Frame height","inputQty":800,"okQty":800,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"Height NG","defectCount":0},
        {"productType":"BRS-161016","testDate":"2025-09-30","line":"E2-3B","checkType":"Function","variable":"PT 161014-S Frame","variableDetail":"Test","variableGroup":"Frame source","intervention":"Test Frame function","inputQty":2684,"okQty":2441,"ngTotal":243,"ngRate":9.1,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":243},
        {"productType":"BRS-161016","testDate":"2025-09-30","line":"E2-3B","checkType":"Function","variable":"PT 161014-S Frame","variableDetail":"Normal","variableGroup":"Frame source","intervention":"Normal Frame function","inputQty":800,"okQty":749,"ngTotal":51,"ngRate":6.4,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":51},
        {"productType":"BRS-161016","testDate":"2025-10-22","line":"E2-3B","checkType":"Function","variable":"PT 161014-S Frame","variableDetail":"Test","variableGroup":"Frame source","intervention":"Test Frame function (22 Oct)","inputQty":864,"okQty":852,"ngTotal":12,"ngRate":1.4,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":12},
        {"productType":"BRS-161016","testDate":"2025-10-22","line":"E2-3B","checkType":"Function","variable":"PT 161014-S Frame","variableDetail":"Normal","variableGroup":"Frame source","intervention":"Normal Frame function (22 Oct)","inputQty":800,"okQty":783,"ngTotal":17,"ngRate":2.1,"defectCategory":"Function","defectType":"Hearing Noise+Touch","defectCount":17},
    ],
    "tags": ["BRS-161016","PT_161014-S","Doojin","frame","laser_cutting","function"],
    "verdict": "worsened",
    "headline": "Frame with NG-dimension PT 161014-S (Doojin coating) drives AOI laser-cutting NG to 19.9% vs 0.8% normal and function NG to 9.1% vs 6.4%.",
    "evidence": [
        {"metric":"AOI laser cutting NG (29 Sep)","baselineLabel":"Normal Frame","baselineValue":"0.8% (4/500)","variantLabel":"Test Frame","variantValue":"19.9% (578/2900)","deltaText":"+19.1pp","deltaSign":"up","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Height dimension NG","baselineLabel":"Normal","baselineValue":"0.0% (0/800)","variantLabel":"Test","variantValue":"1.3% (34/2718)","deltaText":"+1.3pp","deltaSign":"up","note":"Caliper recheck 7/34 still NG (0.3%)","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Function NG (30 Sep)","baselineLabel":"Normal","baselineValue":"6.4% (51/800)","variantLabel":"Test","variantValue":"9.1% (243/2684)","deltaText":"+2.7pp","deltaSign":"up","note":"Decap unknown-reason 55%","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Tension Frame+VP (avg, 30 Sep)","baselineLabel":"Normal","baselineValue":"0.75 kgf","variantLabel":"Test","variantValue":"0.56 kgf","deltaText":"-0.19","deltaSign":"down","note":"Spec 0.4 kgf — both pass","comparisons":None,"bestLabel":"","worstLabel":""},
    ],
    "actions": [
        {"priority":1,"kind":"action","text":"Decision: CAN NOT USE Frame with PT 161014-S NG dimension (Doojin coating) — block from production."},
        {"priority":2,"kind":"risk","text":"Decap shows 55% unknown-reason NG plus CMG/coil offset issues — recovery via rework not viable."},
    ],
    "context": {"process":"Frame qualification with PT 161014-S NG dimension (Doojin coating)","stage":"BRS-161016 line E2-3A/3B","baselineReason":"Normal Frame uses spec-compliant PT 161014-S; test Frame fails 4 dimension callouts."},
    "doeGrid": None, "trendPoints": None,
}

# ============================================================
# 12) 13. BRS-161014 DT Report test VP mold #4 add 0.05mm  — comparison_study
# ============================================================
DATA["13. BRS-161014 DT Report test VP mold #4 add 0.05mm date 27.1.2024"] = {
    "productType": "BRS-161014DT",
    "reportType": "comparison_study",
    "measurements": [
        {"productType":"BRS-161014DT","testDate":"2024-01-27","line":"C2-2A","checkType":"Vision VP+ass'y","variable":"VP mold geom","variableDetail":"VP #4 +0.05mm test","variableGroup":"VP mold","intervention":"VP #4 +0.05mm","inputQty":110,"okQty":110,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"VP bending / VP+CD offset","defectCount":0},
        {"productType":"BRS-161014DT","testDate":"2024-01-27","line":"C2-2A","checkType":"Vision VP+ass'y","variable":"VP mold geom","variableDetail":"Normal VP #2","variableGroup":"VP mold","intervention":"Normal VP #2","inputQty":120,"okQty":120,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"VP bending / VP+CD offset","defectCount":0},
        {"productType":"BRS-161014DT","testDate":"2024-01-27","line":"E2-3A","checkType":"Vision VP+ass'y","variable":"VP mold geom","variableDetail":"VP #4 +0.05mm test","variableGroup":"VP mold","intervention":"VP #4 +0.05mm","inputQty":108,"okQty":108,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161014DT","testDate":"2024-01-27","line":"E2-3A","checkType":"Vision VP+ass'y","variable":"VP mold geom","variableDetail":"Normal VP #9","variableGroup":"VP mold","intervention":"Normal VP #9","inputQty":110,"okQty":110,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161014DT","testDate":"2024-01-27","line":"C2-3A","checkType":"Function","variable":"VP mold geom","variableDetail":"VP #4 +0.05mm test","variableGroup":"VP mold","intervention":"VP #4 +0.05mm function","inputQty":105,"okQty":103,"ngTotal":2,"ngRate":1.9,"defectCategory":"Function","defectType":"Hearing Noise","defectCount":2},
        {"productType":"BRS-161014DT","testDate":"2024-01-27","line":"C2-3A","checkType":"Function","variable":"VP mold geom","variableDetail":"Normal VP #2","variableGroup":"VP mold","intervention":"Normal VP #2 function","inputQty":100,"okQty":99,"ngTotal":1,"ngRate":1.0,"defectCategory":"Function","defectType":"Hearing Noise","defectCount":1},
        {"productType":"BRS-161014DT","testDate":"2024-01-27","line":"E2-3A","checkType":"Function","variable":"VP mold geom","variableDetail":"VP #4 +0.05mm test","variableGroup":"VP mold","intervention":"VP #4 +0.05mm function","inputQty":102,"okQty":100,"ngTotal":2,"ngRate":2.0,"defectCategory":"Function","defectType":"Hearing Noise","defectCount":2},
        {"productType":"BRS-161014DT","testDate":"2024-01-27","line":"E2-3A","checkType":"Function","variable":"VP mold geom","variableDetail":"Normal VP #9","variableGroup":"VP mold","intervention":"Normal VP #9 function","inputQty":109,"okQty":107,"ngTotal":2,"ngRate":1.8,"defectCategory":"Function","defectType":"THD / Noise","defectCount":2},
    ],
    "tags": ["BRS-161014","VP_mold","VP_4","+0.05mm","function","tension"],
    "verdict": "no_clear_effect",
    "headline": "VP #4 +0.05mm (center add) holds Vision NG at 0% on both lines and function NG 1.9-2.0% vs 1.0-1.8% normal — equivalent.",
    "evidence": [
        {"metric":"Vision VP+ass'y NG (C2-2A)","baselineLabel":"Normal VP #2","baselineValue":"0.0% (0/120)","variantLabel":"VP #4 +0.05mm","variantValue":"0.0% (0/110)","deltaText":"0pp","deltaSign":"flat","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Function NG (C2-3A)","baselineLabel":"Normal VP #2","baselineValue":"1.0% (1/100)","variantLabel":"VP #4 +0.05mm","variantValue":"1.9% (2/105)","deltaText":"+0.9pp","deltaSign":"up","note":"Small n; not significant","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Function NG (E2-3A)","baselineLabel":"Normal VP #9","baselineValue":"1.8% (2/109)","variantLabel":"VP #4 +0.05mm","variantValue":"2.0% (2/102)","deltaText":"+0.2pp","deltaSign":"flat","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Tension VP+CD avg (kgf)","baselineLabel":"Normal","baselineValue":"1.64 / 1.51","variantLabel":"VP #4 +0.05mm","variantValue":"1.63 / 1.70","deltaText":"≈ same","deltaSign":"flat","note":"C2-3A / E2-3A; all OK","comparisons":None,"bestLabel":"","worstLabel":""},
    ],
    "actions": [
        {"priority":1,"kind":"action","text":"VP #4 with +0.05mm center add is acceptable — can use it (function NG and tension match normal)."},
    ],
    "context": {"process":"VP-mold geometry tuning","stage":"BRS-161014DT lines C2-2A / E2-3A","baselineReason":"Normal VP #2 (C2) and VP #9 (E2) are the production reference molds; +0.05mm tests whether bending issue can be tuned away."},
    "doeGrid": None, "trendPoints": None,
}

# ============================================================
# 13) 13-1. TIU C11-20 Report test VP find reason NG function high  — comparison_study (acoustic only)
# ============================================================
DATA["13-1. TIU C11-20  Report test VP find reason NG function high 2026.1.12"] = {
    "productType": "TIU C11-20",
    "reportType": "comparison_study",
    "measurements": [],
    "tags": ["TIU","C11-20","VP","SPL","THD","IMP","acoustic"],
    "verdict": "inconclusive",
    "headline": "Acoustic sweep compares Normal 180°, V3.3 200°, V3 210° VP variants; no NG-count summary provided — only SPL/IMP/THD frequency curves.",
    "evidence": [
        {"metric":"Test arms","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"—","deltaSign":"flat","note":"Normal 180° n=10, V3.3 200° n=10, V3 210° n=10","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Data scope","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"","deltaText":"—","deltaSign":"flat","note":"3 sheets: SPL DATA / IMP / THD — frequency curves only","comparisons":None,"bestLabel":"","worstLabel":""},
    ],
    "actions": [
        {"priority":1,"kind":"investigate","text":"Run pass/fail decap or function-NG count on the same arms to convert acoustic curves into a verdict."},
    ],
    "context": {"process":"VP variant acoustic characterisation","stage":"TIU C11-20 function-investigation","baselineReason":"Normal 180° is the current VP angle; V3.3 200° and V3 210° are candidate variants to reduce high-NG function."},
    "doeGrid": None, "trendPoints": None,
}


# ============================================================
# Translations (manual narrative ko/vi for headline, actions, context)
# ============================================================
TRANSLATIONS = {}

# Each: name -> {"ko": {"headline": ..., "actions":[{priority,kind,text}], "context": {process,stage,baselineReason}}, "vi": {...}}

def make_tr(headline_ko, headline_vi, actions, context):
    """Build {ko, vi} dict from per-language strings."""
    ko_actions = [{"priority":a["priority"],"kind":a["kind"],"text":a["ko"]} for a in actions]
    vi_actions = [{"priority":a["priority"],"kind":a["kind"],"text":a["vi"]} for a in actions]
    return {
        "ko": {"headline": headline_ko, "actions": ko_actions,
               "context": {"process": context["process"]["ko"], "stage": context["stage"]["ko"], "baselineReason": context["baselineReason"]["ko"]}},
        "vi": {"headline": headline_vi, "actions": vi_actions,
               "context": {"process": context["process"]["vi"], "stage": context["stage"]["vi"], "baselineReason": context["baselineReason"]["vi"]}},
    }


TRANSLATIONS["13. BRS-161014 TEST AOI Check OK, NG"] = make_tr(
    "AOI 검사 로그: 본딩 NG 0.67% (2/300), 레이저 커팅 NG 1.0% (3/300).",
    "Nhật ký kiểm tra AOI: bonding NG 0.67% (2/300), laser cutting NG 1.0% (3/300).",
    [],
    {"process":{"ko":"AOI 검사","vi":"Kiểm tra AOI"},
     "stage":{"ko":"BRS-161014 Main 2 라인","vi":"BRS-161014 line Main 2"},
     "baselineReason":{"ko":"일일 AOI 품질 로그; 비교군 없음.","vi":"Nhật ký AOI hàng ngày; không có nhóm so sánh."}}
)

TRANSLATIONS["12.MSU-20S15-07 DT Result check Height dimension S-MG - Date 2025.03.28"] = make_tr(
    "S-MG (Riujin) 높이 측정 결과 스펙 0.66~0.71mm 내; Short 0.69mm 44%, Long 0.68mm 42%.",
    "Kết quả đo chiều cao S-MG (Riujin) trong spec 0.66~0.71mm; Short 0.69mm 44%, Long 0.68mm 42%.",
    [],
    {"process":{"ko":"수입 치수 QA","vi":"QA kích thước nguyên liệu nhập"},
     "stage":{"ko":"자재 S-MG (Riujin)","vi":"Vật liệu S-MG (Riujin)"},
     "baselineReason":{"ko":"로트 품질 로그; 전체 스펙 0.66~0.71mm 내.","vi":"Nhật ký chất lượng lô; tất cả trong spec 0.66~0.71mm."}}
)

TRANSLATIONS["13. BRS-161016 Report check VP+CD waiting long time to Led UV afer bonding  28.4.2025"] = make_tr(
    "본딩 후 LED UV 대기시간 60s 초과 시 VP+CD 분리율 급증, 90s/240s 75%; 정상 17s에서는 0%.",
    "Khi thời gian chờ LED UV sau bonding vượt 60s, tỉ lệ tách VP+CD tăng mạnh đến 75% (90s/240s) so với 0% ở normal 17s.",
    [
        {"priority":1,"kind":"action","ko":"본딩 후 VP+CD가 LED UV에 가는 대기시간을 36s 이내로 제한하여 분리 NG를 0% 수준으로 유지.","vi":"Giới hạn thời gian chờ VP+CD đến LED UV dưới 36s sau bonding để giữ NG tách ~0%."},
        {"priority":2,"kind":"risk","ko":"60s 이상 대기 시 분리 NG와 tension 모두 급락; 이 공정에서 라인 정지/대기열 회피 필요.","vi":"Khi chờ trên 60s, NG tách và tension đều giảm mạnh; tránh dừng line/xếp hàng tại công đoạn này."},
    ],
    {"process":{"ko":"VP+CD 본딩 및 LED UV 경화","vi":"Bonding VP+CD và sấy LED UV"},
     "stage":{"ko":"BRS-161016 sub-1","vi":"BRS-161016 sub-1"},
     "baselineReason":{"ko":"17s 'normal'은 양산 라인의 표준 대기시간.","vi":"17s 'normal' là thời gian chờ tiêu chuẩn của line sản xuất."}}
)

TRANSLATIONS["13. MSU-L20S15-07DT  Report Test DOE improve  NG weak solder  date 27.5.2025"] = make_tr(
    "DOE 용접시간 34-48ms 전 테스트 arm은 weak-solder NG 25-87.5%; 원래 38ms + 본드패드 G05-0001 normal은 0%.",
    "DOE quét weld-time 34-48ms tất cả các arm thử nghiệm đều có NG weak solder 25-87.5% so với 0% của normal 38ms + bond pad G05-0001.",
    [
        {"priority":1,"kind":"investigate","ko":"본드패드 G05-0001(normal만 사용)이 주된 요인; 동일 패드로 weld-time을 다시 DOE해야 시간 효과 분리 가능.","vi":"Bond pad G05-0001 (chỉ normal dùng) là yếu tố chính; chạy lại DOE weld-time với cùng pad để tách hiệu ứng thời gian."},
        {"priority":2,"kind":"action","ko":"G05-0001 패드 대조시험 완료 전까지 신규 weld-time 설정 전개 금지.","vi":"Không triển khai cài đặt weld-time mới trước khi hoàn tất so sánh có kiểm soát với pad G05-0001."},
    ],
    {"process":{"ko":"스폿 용접 (Weld 1&2 시간)","vi":"Hàn điểm (thời gian Weld 1&2)"},
     "stage":{"ko":"MSU-L20S15-07 라인, E2 검증","vi":"Line MSU-L20S15-07, kiểm chứng E2"},
     "baselineReason":{"ko":"Normal arm은 양산 본드패드 G05-0001을 38ms로 사용 — 현재 작동 셋포인트.","vi":"Arm normal dùng bond pad sản xuất G05-0001 ở 38ms — setpoint vận hành hiện hành."}}
)

TRANSLATIONS["13. BRS-161014  Report check PT measure coplanarty 2023.11.22"] = make_tr(
    "PT 평탄도 <0.01mm는 spread-glue NG 0%, 0.01~0.02mm 그룹은 31-62%.",
    "PT coplanarity <0.01mm cho NG spread-glue 0% so với 31-62% của nhóm 0.01~0.02mm.",
    [
        {"priority":1,"kind":"action","ko":"디캡 본딩용 PT 공급을 평탄도 <0.01mm로 제한하고 0.01~0.02mm 로트는 거부.","vi":"Hạn chế cấp PT cho decap bonding chỉ coplanarity <0.01mm; từ chối lô 0.01~0.02mm."},
    ],
    {"process":{"ko":"디캡 본딩 CMG+CP","vi":"Decap bond CMG+CP"},
     "stage":{"ko":"BRS-161014 sub-2","vi":"BRS-161014 sub-2"},
     "baselineReason":{"ko":"PT 0.01~0.02mm는 입고 허용 범위 상한; <0.01mm가 타이트한 목표.","vi":"PT 0.01~0.02mm là dung sai nhập rộng hơn; <0.01mm là mục tiêu chặt hơn."}}
)

TRANSLATIONS["13. BRS-161014 Report test tracking 1k material 01.09.2023"] = make_tr(
    "1k 자재 전 공정 추적: sub-1 NG <1.5%, sub-4 visual 5-10%, 최종 function NG 56.3% (Hearing Noise 58.7% 주도).",
    "Theo dõi toàn quy trình 1k vật liệu: sub-1 NG <1.5%, sub-4 visual 5-10%, function cuối NG 56.3% (Hearing Noise 58.7% chiếm chính).",
    [
        {"priority":1,"kind":"investigate","ko":"Sub-4 Visual VP의 VP damage + not-enough-glue 손실 8-10%가 수율 저하 주도; SUB-4 ass'y 세팅 재검토.","vi":"Tổn thất 8-10% từ VP damage + not-enough-glue tại Sub-4 Visual VP là nguyên nhân chính; rà soát thông số SUB-4 ass'y."},
        {"priority":2,"kind":"investigate","ko":"Hearing Noise 단독 NG 240건 (58.7%); function 공정에서 decap 연구 필요.","vi":"Hearing Noise đơn lẻ gây 240 NG (58.7%); cần nghiên cứu decap tại trạm function."},
    ],
    {"process":{"ko":"1k 단위 전 공정 추적","vi":"Theo dõi toàn quy trình 1k"},
     "stage":{"ko":"Sub-1 → Sub-2 → Main → Sub-3 → Sub-4 → Function","vi":"Sub-1 → Sub-2 → Main → Sub-3 → Sub-4 → Function"},
     "baselineReason":{"ko":"단일 추적 로트로 공정별 누적 수율 손실 매핑.","vi":"Một lô theo dõi để vẽ tổn thất sản lượng tích luỹ giữa các trạm."}}
)

TRANSLATIONS["13. BRS-161014  DT Report test VP bending date 15.10.2024"] = make_tr(
    "VP 벤딩 function NG 2.8-8.3% vs 정상 #6 2.5%; VP #4 최악 8.3% — 벤딩 시 VP/CD 분리 잠재 위험 존재.",
    "Function NG của VP bending 2.8-8.3% so với 2.5% của normal #6; VP #4 tệ nhất 8.3% — có nguy cơ tách VP/CD khi bending.",
    [
        {"priority":1,"kind":"risk","ko":"VP #4 벤딩 로트는 hearing-noise NG 8.3%; 양산 투입 보류.","vi":"Lô VP #4 bending có NG hearing-noise 8.3% — tạm ngừng đưa vào sản xuất hàng loạt."},
        {"priority":2,"kind":"investigate","ko":"Decap NG 샘플: #3=CM offset, #4=Frame+yoke offset 12.5%, #9=Frame+yoke 20% + 원인불명 60%.","vi":"Mẫu decap NG: #3=CM offset, #4=Frame+yoke offset 12.5%, #9=Frame+yoke 20% + unknown 60%."},
    ],
    {"process":{"ko":"VP 벤딩 몰드 검증","vi":"Định chuẩn khuôn VP bending"},
     "stage":{"ko":"BRS-161016 sub-1 + function","vi":"BRS-161016 sub-1 + function"},
     "baselineReason":{"ko":"VP #6 정상-벤딩 몰드가 현재 양산 기준.","vi":"Khuôn VP #6 normal-bending là tham chiếu sản xuất hiện hành."}}
)

TRANSLATIONS["12.BRS-161016 GMI Test VP all mold block date 4.9.2025"] = make_tr(
    "테스트 VP 몰드 (#5/#8/#9/#10/#11) 모두 정상 #6 대비 벤딩 NG (0.4% → 5.5-21.4%)와 function NG (2.6% → 5.5-8.0%) 상승.",
    "Tất cả khuôn VP thử (#5/#8/#9/#10/#11) đều cao hơn normal #6 cả về NG bending (5.5-21.4% vs 0.4%) lẫn NG function (5.5-8.0% vs 2.6%).",
    [
        {"priority":1,"kind":"action","ko":"양산은 정상 VP #6 로트 계속 사용; VP #9 / #8은 대량 생산 차단.","vi":"Tiếp tục dùng lô VP #6 normal cho sản xuất; chặn VP #9 / #8 khỏi sản xuất hàng loạt."},
        {"priority":2,"kind":"investigate","ko":"VP #5 / #10이 정상에 가장 근접; #6 대비 몰드 치수 재측정으로 허용 가능한 몰드 후보 탐색.","vi":"VP #5 / #10 gần normal nhất — đo lại kích thước khuôn so với #6 để tìm khuôn ứng viên."},
    ],
    {"process":{"ko":"VP 몰드 검증 (GMI VP-all-mold)","vi":"Định chuẩn khuôn VP (GMI VP-all-mold)"},
     "stage":{"ko":"BRS-161016 라인 E2-3B / C2-3A","vi":"BRS-161016 line E2-3B / C2-3A"},
     "baselineReason":{"ko":"로트 IR250721007의 normal #6이 현재 양산 기준 몰드.","vi":"VP #6 normal từ lô IR250721007 là khuôn tham chiếu sản xuất hiện hành."}}
)

TRANSLATIONS["13. BRS-161016 DT  Report  check Reason NG Low gauss  date 12.7.2025"] = make_tr(
    "MG-C Beijing Box 1 (1차 로트) low-gauss NG 32%; Ruijin normal은 0% — Beijing semi-yoke 평균 484~514G vs Ruijin 525G.",
    "MG-C Beijing Box 1 (lô 1) NG low-gauss 32%; Ruijin normal 0% — semi-yoke trung bình Beijing 484~514G so với Ruijin 525G.",
    [
        {"priority":1,"kind":"action","ko":"MG-C Beijing Box 1 (로트 IR250623044/IR250703013) 격리; 평균 gauss ≥ 510인 박스만 출고.","vi":"Cách ly MG-C Beijing Box 1 (lô IR250623044/IR250703013); chỉ xuất box có gauss trung bình ≥ 510."},
        {"priority":2,"kind":"investigate","ko":"Beijing 벤더의 semi-yoke gauss 공정관리 점검 — 박스 간 편차 484~514 과대.","vi":"Kiểm tra kiểm soát quy trình gauss semi-yoke của vendor Beijing — chênh giữa box 484~514 quá rộng."},
    ],
    {"process":{"ko":"MG-C low-gauss 검사","vi":"Kiểm tra low-gauss MG-C"},
     "stage":{"ko":"BRS-161016 라인 C2-3A / C2-5A","vi":"BRS-161016 line C2-3A / C2-5A"},
     "baselineReason":{"ko":"Ruijin 로트 IR250721007이 양산 라인의 normal 벤더.","vi":"Lô Ruijin IR250721007 là vendor normal trên line sản xuất."}}
)

TRANSLATIONS["13. BRS-201506 Report test New BP+SM Ass'y guide JIG date 5.2.2024"] = make_tr(
    "신규 BP+SM ass'y JIG (+0.02 out)은 SM-vision NG 0% 유지, function NG 4회 비교에서 new vs normal 3.0/5.2, 3.0/3.3, 0.7/1.7, 4.7/4.6%로 혼재.",
    "JIG ass'y BP+SM mới (+0.02 out) giữ NG SM-vision ở 0%, NG function qua 4 lần đo cho new vs normal 3.0/5.2, 3.0/3.3, 0.7/1.7, 4.7/4.6%.",
    [
        {"priority":1,"kind":"investigate","ko":"4회 시험 간 function NG 차이가 일관적이지 않음; 신규 JIG 도입 전 대량 샘플 수집 필요.","vi":"Khác biệt NG function giữa 4 lần không nhất quán; cần thu mẫu lớn hơn trước khi áp dụng JIG mới."},
        {"priority":2,"kind":"action","ko":"YK 201506 치수 도면을 19.50/+0.1, 14.50/+0.1로 변경하여 0.02 JIG-gap 효과 흡수.","vi":"Đổi bản vẽ kích thước YK 201506 thành 19.50/+0.1, 14.50/+0.1 để hấp thụ khe JIG 0.02."},
    ],
    {"process":{"ko":"BP+SM 어셈블리","vi":"Lắp ráp BP+SM"},
     "stage":{"ko":"BRS-201506 sub","vi":"BRS-201506 sub"},
     "baselineReason":{"ko":"기존 라인 JIG이 현재 설정; 신규 JIG은 0.02만큼 바깥으로 이동시켜 SMG spread-glue 부분 NG 해결 목적.","vi":"JIG line hiện hành là cài đặt hiện tại; JIG mới dịch ra 0.02 nhằm xử lý NG partial spread-glue SMG."}}
)

TRANSLATIONS["13. BRS-161016 Report Test PT 161014-S of Press line (Doojin coating) happen  NG dimension 30.9.2025"] = make_tr(
    "NG 치수 PT 161014-S (Doojin 코팅) Frame 사용 시 AOI 레이저 커팅 NG 19.9% (normal 0.8%), function NG 9.1% (normal 6.4%).",
    "Khi dùng Frame có PT 161014-S NG kích thước (Doojin coating): NG AOI laser cutting 19.9% so với normal 0.8%, NG function 9.1% so với 6.4%.",
    [
        {"priority":1,"kind":"action","ko":"결정: PT 161014-S NG 치수 (Doojin 코팅) Frame 사용 불가 — 생산 차단.","vi":"Quyết định: KHÔNG dùng Frame có PT 161014-S NG kích thước (Doojin coating) — chặn khỏi sản xuất."},
        {"priority":2,"kind":"risk","ko":"Decap 결과 원인불명 NG 55% + CMG/coil offset 문제 — 재작업 회수 불가능.","vi":"Decap cho thấy 55% NG không rõ nguyên nhân + lỗi CMG/coil offset — không thể phục hồi bằng rework."},
    ],
    {"process":{"ko":"PT 161014-S NG 치수 (Doojin 코팅) Frame 검증","vi":"Định chuẩn Frame có PT 161014-S NG kích thước (Doojin coating)"},
     "stage":{"ko":"BRS-161016 라인 E2-3A/3B","vi":"BRS-161016 line E2-3A/3B"},
     "baselineReason":{"ko":"Normal Frame은 스펙 적합 PT 161014-S; 테스트 Frame은 4개 치수 항목 불합격.","vi":"Frame normal dùng PT 161014-S đạt spec; Frame test sai 4 chỉ tiêu kích thước."}}
)

TRANSLATIONS["13. BRS-161014 DT Report test VP mold #4 add 0.05mm date 27.1.2024"] = make_tr(
    "VP #4 +0.05mm (센터 추가)은 양 라인 Vision NG 0% 유지, function NG 1.9-2.0% vs normal 1.0-1.8% — 동등.",
    "VP #4 +0.05mm (thêm tâm) giữ NG Vision 0% trên cả hai line, NG function 1.9-2.0% so với 1.0-1.8% của normal — tương đương.",
    [
        {"priority":1,"kind":"action","ko":"VP #4 +0.05mm 센터 추가 사양 사용 가능 (function NG 및 tension이 normal과 일치).","vi":"VP #4 thêm tâm +0.05mm chấp nhận được — có thể dùng (NG function và tension tương đương normal)."},
    ],
    {"process":{"ko":"VP 몰드 형상 튜닝","vi":"Tinh chỉnh hình học khuôn VP"},
     "stage":{"ko":"BRS-161014DT 라인 C2-2A / E2-3A","vi":"BRS-161014DT line C2-2A / E2-3A"},
     "baselineReason":{"ko":"Normal VP #2 (C2)와 VP #9 (E2)가 양산 기준 몰드; +0.05mm는 벤딩 이슈 튜닝 가능성 검증.","vi":"VP #2 (C2) và VP #9 (E2) normal là khuôn tham chiếu; +0.05mm thử xem có thể giảm vấn đề bending."}}
)

TRANSLATIONS["13-1. TIU C11-20  Report test VP find reason NG function high 2026.1.12"] = make_tr(
    "음향 스윕으로 Normal 180°, V3.3 200°, V3 210° VP 변종 비교; NG 카운트 요약 없음 — SPL/IMP/THD 주파수 곡선만 제공.",
    "Quét âm học so sánh các biến thể VP Normal 180°, V3.3 200°, V3 210°; không có tóm tắt số NG — chỉ có đường cong SPL/IMP/THD theo tần số.",
    [
        {"priority":1,"kind":"investigate","ko":"동일한 arm에 대해 decap 또는 function-NG 카운트 시험을 진행해 음향 곡선을 verdict로 전환 필요.","vi":"Chạy đếm pass/fail decap hoặc function-NG trên cùng các arm để chuyển đường cong âm học thành kết luận."},
    ],
    {"process":{"ko":"VP 변종 음향 특성화","vi":"Đặc tính âm học các biến thể VP"},
     "stage":{"ko":"TIU C11-20 function-조사","vi":"TIU C11-20 điều tra function"},
     "baselineReason":{"ko":"Normal 180°가 현재 VP 각도; V3.3 200°와 V3 210°는 높은 function NG 저감용 후보 변종.","vi":"Normal 180° là góc VP hiện hành; V3.3 200° và V3 210° là biến thể ứng viên để giảm NG function cao."}}
)


def commit(con, name, result):
    cur = con.cursor()
    cur.execute("BEGIN")
    try:
        product = result.get("productType","")
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
                 int(m.get("defectCount",0)), NOW))

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
            (name, product, "", "", tags_json, NOW,
             "", "", "", "", "",
             result.get("verdict",""), result.get("headline",""),
             evidence_json, actions_json, context_json,
             result.get("reportType",""), doe_json, trend_json))

        # Translations
        tr = TRANSLATIONS.get(name)
        if tr:
            for lang in ("ko","vi"):
                t = tr[lang]
                tr_actions_json = json.dumps(t.get("actions") or [], ensure_ascii=False)
                tr_context_json = json.dumps(t.get("context"), ensure_ascii=False) if t.get("context") else ""
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
                    (name, lang, "", "", "", "", "", "", "",
                     t.get("headline",""), tr_actions_json, tr_context_json, NOW))

        con.commit()
        return True
    except Exception as e:
        con.rollback()
        print(f"[PARSE-FAIL {name}] {e}")
        return False


def main():
    con = sqlite3.connect(DB)
    con.execute("PRAGMA busy_timeout = 30000")
    ok_count = 0
    skip_count = 0
    fail_count = 0
    targets = [
        "12.BRS-161016 GMI Test VP all mold block date 4.9.2025",
        "12.MSU-20S15-07 DT Result check Height dimension S-MG - Date 2025.03.28",
        "13-1. TIU C11-20  Report test VP find reason NG function high 2026.1.12",
        "13. BRS-161014  DT Report test VP bending date 15.10.2024",
        "13. BRS-161014  Report check PT measure coplanarty 2023.11.22",
        "13. BRS-161014 DT Report test VP mold #4 add 0.05mm date 27.1.2024",
        "13. BRS-161014 Report test tracking 1k material 01.09.2023",
        "13. BRS-161014 TEST AOI Check OK, NG",
        "13. BRS-161016 DT  Report  check Reason NG Low gauss  date 12.7.2025",
        "13. BRS-161016 Report Test PT 161014-S of Press line (Doojin coating) happen  NG dimension 30.9.2025",
        "13. BRS-161016 Report check VP+CD waiting long time to Led UV afer bonding  28.4.2025",
        "13. BRS-201506 Report test New BP+SM Ass'y guide JIG date 5.2.2024",
        "13. MSU-L20S15-07DT  Report Test DOE improve  NG weak solder  date 27.5.2025",
    ]
    for name in targets:
        if name not in DATA:
            print(f"[SKIP {name}] no DATA entry")
            skip_count += 1
            continue
        ok = commit(con, name, DATA[name])
        if ok:
            print(f"[OK {name}]")
            ok_count += 1
        else:
            fail_count += 1
    con.close()
    print(f"\n=== BATCH DONE ===\nMode: Reanalyze\nProcessed: {ok_count}\nSkipped: {skip_count}\nParse-fail: {fail_count}")

if __name__ == "__main__":
    main()
