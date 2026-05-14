"""Commit chunk_03 normalized results to process-review.db."""
import sqlite3, json, os, sys
from datetime import datetime

DB = r"D:\000. MyWorks\002. DB\process-review.db"
con = sqlite3.connect(DB)
con.execute("PRAGMA busy_timeout = 30000")
cur = con.cursor()

def commit(name, result, tr_ko, tr_vi):
    cur.execute("BEGIN")
    try:
        now = datetime.utcnow().isoformat() + "Z"
        product = result.get("productType", "")
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
        doe_json   = json.dumps(result.get("doeGrid"), ensure_ascii=False)     if result.get("doeGrid")     else ""
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
                 "", "", "", "", "", "", "",
                 tr.get("headline",""), tr_actions_json, tr_context_json, now))

        con.commit()
        print(f"[OK] {name}")
        return True
    except Exception as e:
        con.rollback()
        print(f"[PARSE-FAIL] {name}: {e}")
        return False

# ============ Per-dataset analysis ============

DATASETS = []

# 0: BRS-161016 GMI test 3 new bonds (PW1470SX-N1, AC026, A-3424B) vs Normal — multi_arm
DATASETS.append({
"name": "10.BRS-161016GMI- Report test 3 Type new bond PW1470SX-N1 & AC026 & A-3424B Ass'y VP+CD- 2025.07.07",
"result": {
    "productType":"BRS-161016 GMI",
    "measurements":[
        # Day1 07/Jul VP+CD separate + function
        {"productType":"BRS-161016 GMI","testDate":"2025-07-07","line":"E2-3A","checkType":"process","variable":"PW 1470AX-N1","variableDetail":"VP+CD separate","variableGroup":"test","intervention":"bond PW 1470AX-N1","inputQty":76,"okQty":76,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
        {"productType":"BRS-161016 GMI","testDate":"2025-07-07","line":"E2-3A","checkType":"function","variable":"PW 1470AX-N1","variableDetail":"function","variableGroup":"test","intervention":"bond PW 1470AX-N1","inputQty":66,"okQty":63,"ngTotal":3,"ngRate":4.5,"defectCategory":"function_hearing","defectType":"Noise","defectCount":3},
        {"productType":"BRS-161016 GMI","testDate":"2025-07-07","line":"E2-3A","checkType":"process","variable":"AC026","variableDetail":"VP+CD separate","variableGroup":"test","intervention":"bond AC026","inputQty":70,"okQty":69,"ngTotal":1,"ngRate":1.4,"defectCategory":"assembly_defect","defectType":"VP+CD separate","defectCount":1},
        {"productType":"BRS-161016 GMI","testDate":"2025-07-07","line":"E2-3A","checkType":"function","variable":"AC026","variableDetail":"function","variableGroup":"test","intervention":"bond AC026","inputQty":60,"okQty":59,"ngTotal":1,"ngRate":1.7,"defectCategory":"function_hearing","defectType":"Noise","defectCount":1},
        {"productType":"BRS-161016 GMI","testDate":"2025-07-07","line":"E2-3A","checkType":"process","variable":"A-3424B","variableDetail":"VP+CD separate","variableGroup":"test","intervention":"bond A-3424B","inputQty":88,"okQty":88,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
        {"productType":"BRS-161016 GMI","testDate":"2025-07-07","line":"E2-3A","checkType":"function","variable":"A-3424B","variableDetail":"function","variableGroup":"test","intervention":"bond A-3424B","inputQty":78,"okQty":75,"ngTotal":3,"ngRate":3.8,"defectCategory":"function_hearing","defectType":"Noise","defectCount":3},
        {"productType":"BRS-161016 GMI","testDate":"2025-07-07","line":"E2-3A","checkType":"process","variable":"Normal","variableDetail":"VP+CD separate","variableGroup":"normal","intervention":"","inputQty":90,"okQty":90,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
        {"productType":"BRS-161016 GMI","testDate":"2025-07-07","line":"E2-3A","checkType":"function","variable":"Normal","variableDetail":"function","variableGroup":"normal","intervention":"","inputQty":80,"okQty":77,"ngTotal":3,"ngRate":3.8,"defectCategory":"function_hearing","defectType":"Noise","defectCount":3},
        # Day2 09/Jul totals (use Total roll-up)
        {"productType":"BRS-161016 GMI","testDate":"2025-07-09","line":"E2-3A","checkType":"function","variable":"PW 1470AX-N1","variableDetail":"function","variableGroup":"test","intervention":"bond PW 1470AX-N1","inputQty":157,"okQty":152,"ngTotal":5,"ngRate":3.2,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":5},
        {"productType":"BRS-161016 GMI","testDate":"2025-07-09","line":"E2-3A","checkType":"function","variable":"AC026","variableDetail":"function","variableGroup":"test","intervention":"bond AC026","inputQty":153,"okQty":144,"ngTotal":9,"ngRate":5.9,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":9},
        {"productType":"BRS-161016 GMI","testDate":"2025-07-09","line":"E2-3A","checkType":"function","variable":"A-3424B","variableDetail":"function","variableGroup":"test","intervention":"bond A-3424B","inputQty":140,"okQty":136,"ngTotal":4,"ngRate":2.9,"defectCategory":"function_hearing","defectType":"Noise","defectCount":4},
        {"productType":"BRS-161016 GMI","testDate":"2025-07-09","line":"E2-3A","checkType":"function","variable":"Normal","variableDetail":"function","variableGroup":"normal","intervention":"","inputQty":140,"okQty":136,"ngTotal":4,"ngRate":2.9,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":4},
    ],
    "tags":["brs-161016","new-bond","bond-test","vp-cd-ass-y","function-ng","multi-arm","hearing-noise"],
    "reportType":"multi_arm",
    "verdict":"no_clear_effect",
    "headline":"3 new bonds (PW1470SX-N1, AC026, A-3424B) function NG 2.9-4.7% vs Normal 3.2%, within noise.",
    "evidence":[
        {"metric":"Function NG (total)","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"",
         "deltaText":"+1.5pp range","deltaSign":"up","note":"",
         "comparisons":[
            {"label":"Normal (EA6116)","value":"3.2% (7/220)","n":220,"isBaseline":True,"isBest":True,"isWorst":False},
            {"label":"PW 1470AX-N1","value":"3.6% (8/223)","n":223,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"A-3424B","value":"3.2% (7/218)","n":218,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"AC026","value":"4.7% (10/213)","n":213,"isBaseline":False,"isBest":False,"isWorst":True}
         ],
         "bestLabel":"Normal (EA6116)","worstLabel":"AC026"},
        {"metric":"VP+CD separate NG (total)","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"",
         "deltaText":"+0.4pp range","deltaSign":"up","note":"",
         "comparisons":[
            {"label":"Normal","value":"0.0% (0/230)","n":230,"isBaseline":True,"isBest":True,"isWorst":False},
            {"label":"PW 1470AX-N1","value":"0.0% (0/236)","n":236,"isBaseline":False,"isBest":True,"isWorst":False},
            {"label":"AC026","value":"0.4% (1/230)","n":230,"isBaseline":False,"isBest":False,"isWorst":True},
            {"label":"A-3424B","value":"0.4% (1/231)","n":231,"isBaseline":False,"isBest":False,"isWorst":False}
         ],
         "bestLabel":"PW 1470AX-N1","worstLabel":"AC026"},
        {"metric":"Tension VP+CD","baselineLabel":"Normal","baselineValue":">=0.5 kgf (Pass)","variantLabel":"All 3 new bonds","variantValue":"All Pass (AVG 1.17-2.43 kgf)",
         "deltaText":"—","deltaSign":"no_change","note":"all above spec 0.5 kgf","comparisons":None,"bestLabel":"","worstLabel":""}
    ],
    "actions":[
        {"priority":1,"kind":"action","text":"Run 2nd-stage 160pcs/type lot test before mass-production approval"},
        {"priority":2,"kind":"investigate","text":"Re-analyse 25 decap NG samples to confirm bond not root cause"},
        {"priority":3,"kind":"risk","text":"Monitor AC026 hearing-noise rate; slightly above Normal"}
    ],
    "context":{"process":"VP+CD assy bond application + function check","stage":"E2-3A main line 07-09 Jul 2025","baselineReason":"same-event Normal (bond EA6116) row present"}
},
"tr_ko":{
    "headline":"3종 신규 본드(PW1470SX-N1, AC026, A-3424B) 기능 NG 2.9-4.7% vs Normal 3.2%, 편차 범위 내.",
    "actions":[
        {"priority":1,"kind":"action","text":"양산 승인 전 2단계 160pcs/type 로트 테스트 진행"},
        {"priority":2,"kind":"investigate","text":"NG 25 샘플 디캡 재분석으로 본드가 근본원인 아님 확인"},
        {"priority":3,"kind":"risk","text":"AC026 hearing-noise 비율 Normal 대비 약간 높음, 모니터링 필요"}
    ],
    "context":{"process":"VP+CD 조립 본드 도포 + 기능 검사","stage":"E2-3A 메인라인 2025-07-07~09","baselineReason":"동일 시험 내 Normal(EA6116 본드) 행 존재"}
},
"tr_vi":{
    "headline":"3 loại bond mới (PW1470SX-N1, AC026, A-3424B) NG function 2.9-4.7% vs Normal 3.2%, trong dải sai số.",
    "actions":[
        {"priority":1,"kind":"action","text":"Chạy lô test giai đoạn 2 160pcs/type trước khi duyệt sản xuất"},
        {"priority":2,"kind":"investigate","text":"Phân tích lại 25 mẫu NG decap để xác nhận bond không phải nguyên nhân gốc"},
        {"priority":3,"kind":"risk","text":"Theo dõi AC026 hearing-noise, hơi cao hơn Normal"}
    ],
    "context":{"process":"Bôi bond ass'y VP+CD + kiểm tra function","stage":"Line E2-3A main 07-09/07/2025","baselineReason":"có dòng Normal (bond EA6116) cùng sự kiện"}
}
})

# 1: BRS-201506 silicone gasket test (trend over 6 days)
DATASETS.append({
"name": "10.BRS-2015 Report Test material silicone gasket date 13.9.2024",
"result": {
    "productType":"BRS-201506",
    "measurements":[
        {"productType":"BRS-201506","testDate":"2024-09-13","line":"","checkType":"function","variable":"New JIG BAKO + silicone gasket","variableDetail":"function","variableGroup":"test","intervention":"silicone gasket","inputQty":100,"okQty":98,"ngTotal":2,"ngRate":2.0,"defectCategory":"function_hearing","defectType":"Noise+THD","defectCount":2},
        {"productType":"BRS-201506","testDate":"2024-09-13","line":"","checkType":"function","variable":"Normal","variableDetail":"function","variableGroup":"normal","intervention":"","inputQty":100,"okQty":98,"ngTotal":2,"ngRate":2.0,"defectCategory":"function_hearing","defectType":"Noise+THD","defectCount":2},
        {"productType":"BRS-201506","testDate":"2024-09-17","line":"","checkType":"function","variable":"New JIG + silicone + YS 0.6T","variableDetail":"function","variableGroup":"test","intervention":"silicone+YS0.6T","inputQty":100,"okQty":100,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-201506","testDate":"2024-09-17","line":"","checkType":"function","variable":"New JIG + silicone + YS normal 0.45T","variableDetail":"function","variableGroup":"test","intervention":"silicone+YS0.45T","inputQty":100,"okQty":99,"ngTotal":1,"ngRate":1.0,"defectCategory":"function_hearing","defectType":"Noise","defectCount":1},
        {"productType":"BRS-201506","testDate":"2024-09-17","line":"","checkType":"function","variable":"Normal","variableDetail":"function","variableGroup":"normal","intervention":"","inputQty":100,"okQty":99,"ngTotal":1,"ngRate":1.0,"defectCategory":"function_hearing","defectType":"Noise","defectCount":1},
        {"productType":"BRS-201506","testDate":"2024-09-18","line":"","checkType":"function","variable":"New JIG + silicone + YS 0.45T","variableDetail":"function","variableGroup":"test","intervention":"silicone+YS0.45T","inputQty":1000,"okQty":977,"ngTotal":23,"ngRate":2.3,"defectCategory":"function_hearing","defectType":"Noise","defectCount":15},
        {"productType":"BRS-201506","testDate":"2024-09-18","line":"","checkType":"function","variable":"Normal","variableDetail":"function","variableGroup":"normal","intervention":"","inputQty":800,"okQty":772,"ngTotal":28,"ngRate":3.5,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":24},
        {"productType":"BRS-201506","testDate":"2024-09-28","line":"","checkType":"function","variable":"New JIG + silicone + YS 0.55T/0.45T","variableDetail":"function","variableGroup":"test","intervention":"silicone+YS0.55T","inputQty":840,"okQty":812,"ngTotal":28,"ngRate":3.3,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":23},
        {"productType":"BRS-201506","testDate":"2024-09-28","line":"","checkType":"function","variable":"Normal","variableDetail":"function","variableGroup":"normal","intervention":"","inputQty":800,"okQty":773,"ngTotal":27,"ngRate":3.4,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":23},
        {"productType":"BRS-201506","testDate":"2024-09-30","line":"","checkType":"function","variable":"New JIG + silicone + YS 0.55T/0.45T","variableDetail":"function","variableGroup":"test","intervention":"silicone+YS0.55T","inputQty":1755,"okQty":1721,"ngTotal":34,"ngRate":1.9,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":27},
        {"productType":"BRS-201506","testDate":"2024-09-30","line":"","checkType":"function","variable":"Normal","variableDetail":"function","variableGroup":"normal","intervention":"","inputQty":960,"okQty":927,"ngTotal":33,"ngRate":3.4,"defectCategory":"function_hearing","defectType":"Noise","defectCount":28},
        {"productType":"BRS-201506","testDate":"2024-10-01","line":"","checkType":"function","variable":"New JIG + silicone + YS 0.55T/0.45T","variableDetail":"function","variableGroup":"test","intervention":"silicone+YS0.55T","inputQty":957,"okQty":931,"ngTotal":26,"ngRate":2.7,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":21},
        {"productType":"BRS-201506","testDate":"2024-10-01","line":"","checkType":"function","variable":"Normal","variableDetail":"function","variableGroup":"normal","intervention":"","inputQty":800,"okQty":778,"ngTotal":22,"ngRate":2.8,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":22},
        # Total Test vs Total Normal
        {"productType":"BRS-201506","testDate":"","line":"","checkType":"function","variable":"Total Test","variableDetail":"function-total","variableGroup":"test","intervention":"silicone gasket","inputQty":3552,"okQty":3464,"ngTotal":88,"ngRate":2.5,"defectCategory":"function_hearing","defectType":"Noise","defectCount":53},
        {"productType":"BRS-201506","testDate":"","line":"","checkType":"function","variable":"Total Normal","variableDetail":"function-total","variableGroup":"normal","intervention":"","inputQty":2560,"okQty":2478,"ngTotal":82,"ngRate":3.2,"defectCategory":"function_hearing","defectType":"Noise","defectCount":67},
    ],
    "tags":["brs-201506","silicone-gasket","new-jig-bako","material-test","function-ng","hearing-noise","yoke-ys-thickness"],
    "reportType":"comparison_study",
    "verdict":"improved",
    "headline":"New JIG BAKO + silicone gasket function NG 2.5% vs Normal 3.2% (-0.7pp, improved).",
    "evidence":[
        {"metric":"Function NG (total)","baselineLabel":"Normal","baselineValue":"3.2% (82/2560)","variantLabel":"Test (silicone gasket)","variantValue":"2.5% (88/3552)",
         "deltaText":"-0.7pp","deltaSign":"down","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Hearing-Noise share","baselineLabel":"Normal","baselineValue":"81.7% of NG","variantLabel":"Test","variantValue":"60.2% of NG",
         "deltaText":"-21.5pp","deltaSign":"down","note":"noise still dominant","comparisons":None,"bestLabel":"","worstLabel":""}
    ],
    "actions":[
        {"priority":1,"kind":"action","text":"Approve silicone gasket + YS 0.55T/0.45T for production"},
        {"priority":2,"kind":"investigate","text":"Trace noise-NG samples for particle contamination per decap note"}
    ],
    "context":{"process":"SPL + Hearing function check with new BAKO jig + silicone gasket","stage":"Module MSM-X516B function line, 13/9-01/10/2024","baselineReason":"same-period paired Normal rows on each test day"}
},
"tr_ko":{
    "headline":"신규 JIG BAKO + 실리콘 가스켓 기능 NG 2.5% vs Normal 3.2% (-0.7pp, 개선).",
    "actions":[
        {"priority":1,"kind":"action","text":"실리콘 가스켓 + YS 0.55T/0.45T 양산 승인"},
        {"priority":2,"kind":"investigate","text":"디캡 노트의 입자 오염 가능성에 따라 노이즈 NG 샘플 추적"}
    ],
    "context":{"process":"신규 BAKO 지그 + 실리콘 가스켓 SPL/Hearing 기능 검사","stage":"MSM-X516B 모듈 기능 라인, 2024-09-13~10-01","baselineReason":"각 테스트일별 동일 시험 Normal 페어 존재"}
},
"tr_vi":{
    "headline":"JIG BAKO mới + silicone gasket NG function 2.5% vs Normal 3.2% (-0.7pp, cải thiện).",
    "actions":[
        {"priority":1,"kind":"action","text":"Duyệt silicone gasket + YS 0.55T/0.45T cho sản xuất"},
        {"priority":2,"kind":"investigate","text":"Truy vết mẫu NG noise tìm bụi/particle theo ghi chú decap"}
    ],
    "context":{"process":"Kiểm tra function SPL + Hearing với JIG BAKO mới + silicone gasket","stage":"Line function module MSM-X516B, 13/9-01/10/2024","baselineReason":"có cặp Normal cùng kỳ trong mỗi ngày test"}
}
})

# 2: BRS-201506 VP array JIG repair (Test vs Normal)
DATASETS.append({
"name":"10.BRS-201506 Report checking and test VP array JIG repair date 21.12.2024",
"result":{
    "productType":"BRS-201506",
    "measurements":[
        {"productType":"BRS-201506","testDate":"2024-12-21","line":"","checkType":"process","variable":"JIG TEST","variableDetail":"VP+CD corner separate","variableGroup":"test","intervention":"VP array JIG repair","inputQty":151,"okQty":135,"ngTotal":16,"ngRate":10.6,"defectCategory":"assembly_defect","defectType":"VP+CD corner separate","defectCount":16},
        {"productType":"BRS-201506","testDate":"2024-12-21","line":"","checkType":"process","variable":"JIG Normal","variableDetail":"VP+CD corner separate","variableGroup":"normal","intervention":"","inputQty":100,"okQty":5,"ngTotal":95,"ngRate":95.0,"defectCategory":"assembly_defect","defectType":"VP+CD corner separate","defectCount":95},
        {"productType":"BRS-201506","testDate":"2024-12-21","line":"","checkType":"function","variable":"JIG TEST","variableDetail":"function","variableGroup":"test","intervention":"VP array JIG repair","inputQty":150,"okQty":145,"ngTotal":5,"ngRate":3.3,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":5},
        {"productType":"BRS-201506","testDate":"2024-12-21","line":"","checkType":"function","variable":"JIG Normal","variableDetail":"function","variableGroup":"normal","intervention":"","inputQty":296,"okQty":271,"ngTotal":25,"ngRate":8.4,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":25},
    ],
    "tags":["brs-201506","vp-array-jig-repair","jig-repair","function-ng","corner-separate","intervention-test"],
    "reportType":"comparison_study",
    "verdict":"improved",
    "headline":"VP array JIG repair: corner-separate 95.0% to 10.6% (-84.4pp); function NG 8.4% to 3.3% (-5.1pp).",
    "evidence":[
        {"metric":"VP+CD corner separate NG","baselineLabel":"JIG Normal","baselineValue":"95.0% (95/100)","variantLabel":"JIG Test (repaired)","variantValue":"10.6% (16/151)",
         "deltaText":"-84.4pp","deltaSign":"down","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Function NG (total)","baselineLabel":"JIG Normal","baselineValue":"8.4% (25/296)","variantLabel":"JIG Test","variantValue":"3.3% (5/150)",
         "deltaText":"-5.1pp","deltaSign":"down","note":"Touch-noise dominant","comparisons":None,"bestLabel":"","worstLabel":""}
    ],
    "actions":[
        {"priority":1,"kind":"action","text":"Roll out repaired VP array JIG to production"},
        {"priority":2,"kind":"investigate","text":"Confirm residual 10.6% corner-separate root cause"}
    ],
    "context":{"process":"VP array JIG repair targeting VP+CD corner separation","stage":"Sub-line VP assy + function check","baselineReason":"same-day JIG Normal row available as reference"}
},
"tr_ko":{
    "headline":"VP array JIG 수리: 코너 분리 95.0% → 10.6% (-84.4pp); 기능 NG 8.4% → 3.3% (-5.1pp).",
    "actions":[
        {"priority":1,"kind":"action","text":"수리된 VP array JIG 양산 적용"},
        {"priority":2,"kind":"investigate","text":"잔존 10.6% 코너 분리 근본 원인 확인"}
    ],
    "context":{"process":"VP+CD 코너 분리 개선을 위한 VP array JIG 수리","stage":"VP 조립 서브라인 + 기능 검사","baselineReason":"같은 날 JIG Normal 행 기준 가능"}
},
"tr_vi":{
    "headline":"Sửa JIG VP array: tách góc 95.0% → 10.6% (-84.4pp); NG function 8.4% → 3.3% (-5.1pp).",
    "actions":[
        {"priority":1,"kind":"action","text":"Áp dụng JIG VP array đã sửa cho sản xuất"},
        {"priority":2,"kind":"investigate","text":"Xác định nguyên nhân gốc của 10.6% tách góc còn lại"}
    ],
    "context":{"process":"Sửa JIG VP array để giảm tách góc VP+CD","stage":"Sub-line lắp VP + kiểm tra function","baselineReason":"có dòng JIG Normal cùng ngày làm tham chiếu"}
}
})

# 3: L20S15-07 Magnet A/B Ruijin Primer
DATASETS.append({
"name":"10.L20S15-07 Report Test Magnet A and B Of vender Ruijin clean with Primer date 6.9.2025",
"result":{
    "productType":"L20S15-07",
    "measurements":[
        {"productType":"L20S15-07","testDate":"2025-09-06","line":"E2-4B","checkType":"visual_inspection","variable":"Ruijin clean Primer","variableDetail":"SM vision","variableGroup":"test","intervention":"Ruijin Primer","inputQty":100,"okQty":100,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"L20S15-07","testDate":"2025-09-06","line":"E2-4B","checkType":"visual_inspection","variable":"Normal clean alcohol","variableDetail":"SM vision","variableGroup":"normal","intervention":"","inputQty":100,"okQty":100,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        # Decap separate glue
        {"productType":"L20S15-07","testDate":"2025-09-06","line":"E2-4B","checkType":"process","variable":"Ruijin clean Primer","variableDetail":"Decap B-PT+S-MG glue spread","variableGroup":"test","intervention":"Ruijin Primer","inputQty":8,"okQty":4,"ngTotal":4,"ngRate":50.0,"defectCategory":"assembly_defect","defectType":"Bond spread <80%","defectCount":4},
        {"productType":"L20S15-07","testDate":"2025-09-06","line":"E2-4B","checkType":"process","variable":"Normal clean alcohol","variableDetail":"Decap B-PT+S-MG glue spread","variableGroup":"normal","intervention":"","inputQty":8,"okQty":8,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
        # Drop test
        {"productType":"L20S15-07","testDate":"2025-09-06","line":"E2-4B","checkType":"process","variable":"Ruijin clean Primer","variableDetail":"Drop test","variableGroup":"test","intervention":"Ruijin Primer","inputQty":8,"okQty":8,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"L20S15-07","testDate":"2025-09-06","line":"E2-4B","checkType":"process","variable":"Normal clean alcohol","variableDetail":"Drop test","variableGroup":"normal","intervention":"","inputQty":8,"okQty":8,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
    ],
    "tags":["l20s15-07","magnet-ruijin","primer-clean","vender-ruijin","decap-glue-spread","comparison-study"],
    "reportType":"comparison_study",
    "verdict":"worsened",
    "headline":"Ruijin Primer-clean Magnet: decap glue spread NG 50% (4/8) vs Normal alcohol 0% (0/8).",
    "evidence":[
        {"metric":"Decap glue spread B-PT+S-MG (spec >=80%)","baselineLabel":"Normal alcohol","baselineValue":"0.0% (0/8)","variantLabel":"Ruijin Primer","variantValue":"50.0% (4/8)",
         "deltaText":"+50pp","deltaSign":"up","note":"n=8 only — low conf","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"SM vision NG","baselineLabel":"Normal","baselineValue":"0.0% (0/100)","variantLabel":"Primer","variantValue":"0.0% (0/100)",
         "deltaText":"+0pp","deltaSign":"no_change","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Tension Long SM AVG","baselineLabel":"Normal","baselineValue":"33.73 kgf","variantLabel":"Primer","variantValue":"54.49 kgf",
         "deltaText":"+20.76","deltaSign":"up","note":"primer tension higher","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Drop test NG","baselineLabel":"Normal","baselineValue":"0/8","variantLabel":"Primer","variantValue":"0/8",
         "deltaText":"+0pp","deltaSign":"no_change","note":"","comparisons":None,"bestLabel":"","worstLabel":""}
    ],
    "actions":[
        {"priority":1,"kind":"action","text":"Do NOT adopt Ruijin Primer-clean Magnet A/B; decap glue spread fails spec"},
        {"priority":2,"kind":"investigate","text":"Root-cause primer interaction with bond spread <80%"}
    ],
    "context":{"process":"Magnet side A/B clean with Primer at SM/B-PT decap glue check","stage":"E2-4B sub-line 06/09/2025","baselineReason":"same-event Normal alcohol-clean row present"}
},
"tr_ko":{
    "headline":"Ruijin Primer 클린 마그넷: 디캡 본드 스프레드 NG 50% (4/8) vs Normal 알코올 0% (0/8).",
    "actions":[
        {"priority":1,"kind":"action","text":"Ruijin Primer 클린 마그넷 A/B 채택 금지 — 디캡 본드 스프레드 스펙 미달"},
        {"priority":2,"kind":"investigate","text":"Primer가 본드 스프레드 80% 미만으로 떨어뜨리는 원인 규명"}
    ],
    "context":{"process":"마그넷 A/B Primer 클린 후 SM/B-PT 디캡 본드 점검","stage":"E2-4B 서브라인 2025-09-06","baselineReason":"동일 시험 내 알코올 클린 Normal 행 존재"}
},
"tr_vi":{
    "headline":"Magnet Ruijin clean Primer: NG bond spread decap 50% (4/8) vs Normal alcohol 0% (0/8).",
    "actions":[
        {"priority":1,"kind":"action","text":"Không áp dụng Magnet A/B Ruijin clean Primer — bond spread không đạt spec"},
        {"priority":2,"kind":"investigate","text":"Tìm nguyên nhân Primer làm bond spread tụt xuống dưới 80%"}
    ],
    "context":{"process":"Magnet A/B clean Primer kiểm tra bond spread SM/B-PT decap","stage":"Sub-line E2-4B 06/09/2025","baselineReason":"có dòng Normal clean alcohol cùng sự kiện"}
}
})

# 4: BRS-161014 DCR summary DT line — quality_log trend
DATASETS.append({
"name":"11. BRS-161014  Report summary NG of DCR DT line 2023.11.21",
"result":{
    "productType":"BRS-161014 DT",
    "measurements":[
        {"productType":"BRS-161014 DT","testDate":"2023-11-13","line":"C2-3A","checkType":"function","variable":"DCR","variableDetail":"","variableGroup":"","intervention":"","inputQty":1118,"okQty":1118,"ngTotal":0,"ngRate":0.0,"defectCategory":"function_spl","defectType":"","defectCount":0},
        {"productType":"BRS-161014 DT","testDate":"2023-11-14","line":"C2-3A","checkType":"function","variable":"DCR","variableDetail":"","variableGroup":"","intervention":"","inputQty":1544,"okQty":1544,"ngTotal":0,"ngRate":0.0,"defectCategory":"function_spl","defectType":"","defectCount":0},
        {"productType":"BRS-161014 DT","testDate":"2023-11-15","line":"C2-3A","checkType":"function","variable":"DCR","variableDetail":"","variableGroup":"","intervention":"","inputQty":3498,"okQty":3498,"ngTotal":0,"ngRate":0.0,"defectCategory":"function_spl","defectType":"","defectCount":0},
        {"productType":"BRS-161014 DT","testDate":"2023-11-16","line":"C2-3A","checkType":"function","variable":"DCR","variableDetail":"","variableGroup":"","intervention":"","inputQty":3342,"okQty":3342,"ngTotal":0,"ngRate":0.0,"defectCategory":"function_spl","defectType":"","defectCount":0},
        {"productType":"BRS-161014 DT","testDate":"2023-11-17","line":"C2-3A","checkType":"function","variable":"DCR","variableDetail":"","variableGroup":"","intervention":"","inputQty":3451,"okQty":3451,"ngTotal":0,"ngRate":0.0,"defectCategory":"function_spl","defectType":"","defectCount":0},
        {"productType":"BRS-161014 DT","testDate":"2023-11-18","line":"C2-3A","checkType":"function","variable":"DCR","variableDetail":"","variableGroup":"","intervention":"","inputQty":3594,"okQty":3594,"ngTotal":0,"ngRate":0.0,"defectCategory":"function_spl","defectType":"","defectCount":0},
        {"productType":"BRS-161014 DT","testDate":"2023-11-20","line":"C2-3A","checkType":"function","variable":"DCR","variableDetail":"","variableGroup":"","intervention":"","inputQty":2479,"okQty":2479,"ngTotal":0,"ngRate":0.0,"defectCategory":"function_spl","defectType":"","defectCount":0},
        {"productType":"BRS-161014 DT","testDate":"2023-11-21","line":"C2-3A","checkType":"function","variable":"DCR","variableDetail":"","variableGroup":"","intervention":"","inputQty":5771,"okQty":5767,"ngTotal":4,"ngRate":0.1,"defectCategory":"function_spl","defectType":"No sound","defectCount":4},
    ],
    "tags":["brs-161014-dt","dcr-summary","no-sound","weekly-trend","c2-3a","quality-log"],
    "reportType":"trend_analysis",
    "verdict":"no_clear_effect",
    "headline":"DCR weekly summary line C2-3A: 7/8 days 0%, last day 0.1% (4 No-sound).",
    "evidence":[
        {"metric":"DCR NG rate (range)","baselineLabel":"7-day average","baselineValue":"0.0%","variantLabel":"21/Nov peak","variantValue":"0.1% (4/5771)",
         "deltaText":"+0.1pp","deltaSign":"up","note":"No-sound","comparisons":None,"bestLabel":"","worstLabel":""}
    ],
    "actions":[
        {"priority":1,"kind":"investigate","text":"Trace 4 No-sound units on 21/Nov for failure mode"}
    ],
    "context":{"process":"DCR weekly NG summary on DT line","stage":"Line C2-3A 13-21 Nov 2023","baselineReason":"trend report — no comparator, 7-day flat baseline"},
    "trendPoints":[
        {"label":"11/13","value":"0.0%","note":""},
        {"label":"11/14","value":"0.0%","note":""},
        {"label":"11/15","value":"0.0%","note":""},
        {"label":"11/16","value":"0.0%","note":""},
        {"label":"11/17","value":"0.0%","note":""},
        {"label":"11/18","value":"0.0%","note":""},
        {"label":"11/20","value":"0.0%","note":""},
        {"label":"11/21","value":"0.1%","note":"4 No-sound on 5771"}
    ]
},
"tr_ko":{
    "headline":"C2-3A 라인 DCR 주간 요약: 8일 중 7일 0%, 마지막 날 0.1% (No-sound 4건).",
    "actions":[
        {"priority":1,"kind":"investigate","text":"11/21 No-sound 4건 불량 원인 추적"}
    ],
    "context":{"process":"DT 라인 DCR 주간 NG 요약","stage":"라인 C2-3A 2023-11-13~21","baselineReason":"트렌드 보고 — 비교군 없음, 7일 평탄선이 기준"}
},
"tr_vi":{
    "headline":"Tổng kết tuần DCR line C2-3A: 7/8 ngày 0%, ngày cuối 0.1% (4 No-sound).",
    "actions":[
        {"priority":1,"kind":"investigate","text":"Truy vết 4 unit No-sound ngày 21/11 tìm chế độ lỗi"}
    ],
    "context":{"process":"Tổng kết NG tuần DCR trên line DT","stage":"Line C2-3A 13-21/11/2023","baselineReason":"báo cáo trend — không có nhóm so sánh, 7 ngày phẳng làm chuẩn"}
}
})

# 5: BRS-161014 Repair Frame+Coil line 6 — trend over 5-10+ days
DATASETS.append({
"name":"11. BRS-161014 Report  Repair and check damage  Date 24.4.2025 line 6",
"result":{
    "productType":"BRS-161014 DT",
    "measurements":[
        # Vision Frame+Coil totals
        {"productType":"BRS-161014 DT","testDate":"","line":"6","checkType":"visual_inspection","variable":"Total Repair Frame+Coil","variableDetail":"Vision Frame+Coil total","variableGroup":"test","intervention":"Repair Frame+Coil","inputQty":81944,"okQty":63040,"ngTotal":7420,"ngRate":9.1,"defectCategory":"assembly_defect","defectType":"Coil damage+Frame bending+VP separate","defectCount":7420},
        {"productType":"BRS-161014 DT","testDate":"","line":"6","checkType":"visual_inspection","variable":"Total Repair Frame+Coil","variableDetail":"Vision Yoke total","variableGroup":"test","intervention":"Repair Frame+Coil","inputQty":85544,"okQty":79209,"ngTotal":1946,"ngRate":2.3,"defectCategory":"assembly_defect","defectType":"Yoke NG","defectCount":1946},
        {"productType":"BRS-161014 DT","testDate":"","line":"6","checkType":"visual_inspection","variable":"Total Repair Frame+Coil","variableDetail":"Vision VP total","variableGroup":"test","intervention":"Repair Frame+Coil","inputQty":77822,"okQty":76736,"ngTotal":706,"ngRate":0.9,"defectCategory":"assembly_defect","defectType":"VP separate Long+Short","defectCount":706},
        {"productType":"BRS-161014 DT","testDate":"","line":"6","checkType":"function","variable":"Total Repair Frame+Coil","variableDetail":"function total","variableGroup":"test","intervention":"Repair Frame+Coil","inputQty":66120,"okQty":58727,"ngTotal":2271,"ngRate":3.4,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":2040},
    ],
    "tags":["brs-161014-dt","repair-frame-coil","damage-recheck","line-6","frame-bending","coil-damage","vp-separate"],
    "reportType":"quality_log",
    "verdict":"",
    "headline":"Repair Frame+Coil line 6: 81944 vision F+C 9.1% NG; function 3.4% NG over May 2025.",
    "evidence":[
        {"metric":"Vision Frame+Coil NG","baselineLabel":"","baselineValue":"","variantLabel":"Repair lot","variantValue":"9.1% (7420/81944)",
         "deltaText":"—","deltaSign":"no_change","note":"Coil damage 3.4% top","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Vision Yoke NG","baselineLabel":"","baselineValue":"","variantLabel":"Repair lot","variantValue":"2.3% (1946/85544)",
         "deltaText":"—","deltaSign":"no_change","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Vision VP NG","baselineLabel":"","baselineValue":"","variantLabel":"Repair lot","variantValue":"0.9% (706/77822)",
         "deltaText":"—","deltaSign":"no_change","note":"Separate Short 0.8%","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Function NG (total)","baselineLabel":"","baselineValue":"","variantLabel":"Repair lot","variantValue":"3.4% (2271/66120)",
         "deltaText":"—","deltaSign":"no_change","note":"Noise 2.7% dominant","comparisons":None,"bestLabel":"","worstLabel":""}
    ],
    "actions":[
        {"priority":1,"kind":"investigate","text":"Trace Coil damage root cause (3.4% of Frame+Coil vision)"},
        {"priority":2,"kind":"investigate","text":"Drill into Hearing-Noise 2.7% on repaired lots"},
        {"priority":3,"kind":"action","text":"Continue repair separate-Yoke recovery procedure on line 6"}
    ],
    "context":{"process":"Separate Yoke -> clean Yoke/Frame -> vision Coil/Yoke/Frame -> function recovery","stage":"Line 6 May 2025 (lot dates 11-15/Jan)","baselineReason":"quality log of repaired lots — no Normal comparator"}
},
"tr_ko":{
    "headline":"Frame+Coil 수리 라인 6: 비전 F+C 81944 9.1% NG; 5월 누적 기능 3.4% NG.",
    "actions":[
        {"priority":1,"kind":"investigate","text":"Coil damage 근본 원인 추적 (Frame+Coil 비전의 3.4%)"},
        {"priority":2,"kind":"investigate","text":"수리 로트 Hearing-Noise 2.7% 심층 분석"},
        {"priority":3,"kind":"action","text":"라인 6에서 분리-Yoke 회수 절차 유지"}
    ],
    "context":{"process":"분리 Yoke -> Yoke/Frame 클린 -> Coil/Yoke/Frame 비전 -> 기능 회수","stage":"라인 6 2025년 5월 (로트일 1월 11-15)","baselineReason":"수리 로트 품질 로그 — Normal 비교군 없음"}
},
"tr_vi":{
    "headline":"Sửa Frame+Coil line 6: vision F+C 81944 NG 9.1%; tổng function NG 3.4% trong 5/2025.",
    "actions":[
        {"priority":1,"kind":"investigate","text":"Truy nguyên nhân Coil damage (3.4% trong vision Frame+Coil)"},
        {"priority":2,"kind":"investigate","text":"Đào sâu Hearing-Noise 2.7% trên lô sửa"},
        {"priority":3,"kind":"action","text":"Tiếp tục quy trình tách Yoke phục hồi trên line 6"}
    ],
    "context":{"process":"Tách Yoke -> làm sạch Yoke/Frame -> vision Coil/Yoke/Frame -> kiểm function","stage":"Line 6 tháng 5/2025 (lot date 11-15/Jan)","baselineReason":"log chất lượng lô sửa — không có Normal so sánh"}
}
})

# 6: BRS-161014 KR Member function test (high NG 39-81%, quality_log over multi-units)
DATASETS.append({
"name":"11. BRS-161014 Report check Function for TEST of KR Member  2023.08.31",
"result":{
    "productType":"BRS-161014",
    "measurements":[
        {"productType":"BRS-161014","testDate":"2023-08-31","line":"","checkType":"function","variable":"Unit 1","variableDetail":"function","variableGroup":"test","intervention":"KR member","inputQty":27,"okQty":5,"ngTotal":22,"ngRate":81.48,"defectCategory":"function_hearing","defectType":"Noise+Touch+THD","defectCount":15},
        {"productType":"BRS-161014","testDate":"2023-08-31","line":"","checkType":"function","variable":"Unit 3","variableDetail":"function","variableGroup":"test","intervention":"KR member","inputQty":26,"okQty":9,"ngTotal":17,"ngRate":65.38,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":15},
        {"productType":"BRS-161014","testDate":"2023-08-31","line":"","checkType":"function","variable":"Unit 4","variableDetail":"function","variableGroup":"test","intervention":"KR member","inputQty":33,"okQty":20,"ngTotal":13,"ngRate":39.39,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":10},
        {"productType":"BRS-161014","testDate":"2023-08-31","line":"","checkType":"function","variable":"Unit 5","variableDetail":"function","variableGroup":"test","intervention":"KR member","inputQty":30,"okQty":14,"ngTotal":16,"ngRate":53.33,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":13},
        {"productType":"BRS-161014","testDate":"2023-08-31","line":"","checkType":"function","variable":"Unit 6","variableDetail":"function","variableGroup":"test","intervention":"KR member","inputQty":28,"okQty":13,"ngTotal":15,"ngRate":53.57,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":12},
        {"productType":"BRS-161014","testDate":"2023-08-31","line":"","checkType":"function","variable":"Unit 7","variableDetail":"function","variableGroup":"test","intervention":"KR member","inputQty":32,"okQty":12,"ngTotal":20,"ngRate":62.5,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":16},
        {"productType":"BRS-161014","testDate":"2023-08-31","line":"","checkType":"function","variable":"Unit 15","variableDetail":"function","variableGroup":"test","intervention":"KR member","inputQty":24,"okQty":9,"ngTotal":15,"ngRate":62.5,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":12},
    ],
    "tags":["brs-161014","kr-member-test","function-ng","hearing-noise","touch","multi-arm","very-high-ng"],
    "reportType":"multi_arm",
    "verdict":"worsened",
    "headline":"KR Member function test 7 units: NG 39.4-81.5% range; all far above acceptable.",
    "evidence":[
        {"metric":"Function NG rate","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"",
         "deltaText":"+42pp range","deltaSign":"up","note":"all units fail",
         "comparisons":[
            {"label":"Unit 4","value":"39.39% (13/33)","n":33,"isBaseline":False,"isBest":True,"isWorst":False},
            {"label":"Unit 5","value":"53.33% (16/30)","n":30,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"Unit 6","value":"53.57% (15/28)","n":28,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"Unit 7","value":"62.50% (20/32)","n":32,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"Unit 15","value":"62.50% (15/24)","n":24,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"Unit 3","value":"65.38% (17/26)","n":26,"isBaseline":False,"isBest":False,"isWorst":False},
            {"label":"Unit 1","value":"81.48% (22/27)","n":27,"isBaseline":False,"isBest":False,"isWorst":True}
         ],
         "bestLabel":"Unit 4","worstLabel":"Unit 1"}
    ],
    "actions":[
        {"priority":1,"kind":"action","text":"Reject KR Member lot; do NOT release to assy"},
        {"priority":2,"kind":"investigate","text":"Inspect KR Member dimension/material with vendor"},
        {"priority":3,"kind":"risk","text":"High Touch+Noise across all units indicates systemic issue"}
    ],
    "context":{"process":"Function check (Aging/Frequency/Hearing) on KR Member test samples","stage":"31 Aug 2023 — multi-unit pilot","baselineReason":"no Normal — multi-unit comparison only"}
},
"tr_ko":{
    "headline":"KR Member 기능 테스트 7개: NG 39.4-81.5% 범위; 모두 허용 수준 초과.",
    "actions":[
        {"priority":1,"kind":"action","text":"KR Member 로트 거절; 조립 라인 투입 금지"},
        {"priority":2,"kind":"investigate","text":"벤더와 KR Member 치수/재질 검사"},
        {"priority":3,"kind":"risk","text":"모든 유닛 Touch+Noise 동시 발생 — 시스템적 문제 의심"}
    ],
    "context":{"process":"KR Member 테스트 샘플 기능(Aging/Frequency/Hearing) 검사","stage":"2023-08-31 다중 유닛 파일럿","baselineReason":"Normal 없음 — 유닛 간 비교"}
},
"tr_vi":{
    "headline":"Test function KR Member 7 unit: NG 39.4-81.5%; tất cả vượt mức chấp nhận.",
    "actions":[
        {"priority":1,"kind":"action","text":"Loại lô KR Member; không cho vào lắp ráp"},
        {"priority":2,"kind":"investigate","text":"Kiểm tra kích thước/vật liệu KR Member với vendor"},
        {"priority":3,"kind":"risk","text":"Touch+Noise cao đồng đều — vấn đề mang tính hệ thống"}
    ],
    "context":{"process":"Kiểm tra function (Aging/Frequency/Hearing) trên mẫu KR Member","stage":"31/8/2023 thử nghiệm nhiều unit","baselineReason":"không có Normal — chỉ so giữa các unit"}
}
})

# 7: BRS-161014 V0 vs V1 jig Yoke
DATASETS.append({
"name":"11. BRS-161014 Report test Yoke of V0 and V1 jig  date 25.1.2024",
"result":{
    "productType":"BRS-161014",
    "measurements":[
        {"productType":"BRS-161014","testDate":"2024-01-25","line":"E2","checkType":"function","variable":"Test V0 jig","variableDetail":"function","variableGroup":"test","intervention":"Yoke V0 jig","inputQty":500,"okQty":497,"ngTotal":3,"ngRate":0.6,"defectCategory":"function_hearing","defectType":"Noise+THD","defectCount":3},
        {"productType":"BRS-161014","testDate":"2024-01-25","line":"E2","checkType":"function","variable":"Test V1 jig","variableDetail":"function","variableGroup":"test","intervention":"Yoke V1 jig","inputQty":498,"okQty":497,"ngTotal":1,"ngRate":0.2,"defectCategory":"function_hearing","defectType":"Noise","defectCount":1},
        {"productType":"BRS-161014","testDate":"2024-01-25","line":"E2","checkType":"function","variable":"Normal (V0+V1 jig)","variableDetail":"function","variableGroup":"normal","intervention":"","inputQty":800,"okQty":795,"ngTotal":5,"ngRate":0.6,"defectCategory":"function_hearing","defectType":"Noise+SPL+THD","defectCount":5},
    ],
    "tags":["brs-161014","yoke-jig","v0-v1-jig","function-ng","hearing-noise","multi-arm"],
    "reportType":"multi_arm",
    "verdict":"no_clear_effect",
    "headline":"Yoke V0/V1 jig function NG 0.2-0.6% vs Normal 0.6% — both jigs equivalent.",
    "evidence":[
        {"metric":"Function NG","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"",
         "deltaText":"+0.4pp range","deltaSign":"no_change","note":"",
         "comparisons":[
            {"label":"Normal (V0+V1)","value":"0.6% (5/800)","n":800,"isBaseline":True,"isBest":False,"isWorst":True},
            {"label":"V0 jig","value":"0.6% (3/500)","n":500,"isBaseline":False,"isBest":False,"isWorst":True},
            {"label":"V1 jig","value":"0.2% (1/498)","n":498,"isBaseline":False,"isBest":True,"isWorst":False}
         ],
         "bestLabel":"V1 jig","worstLabel":"V0 jig"}
    ],
    "actions":[
        {"priority":1,"kind":"action","text":"Either V0 or V1 jig acceptable — keep V1 if simpler"},
        {"priority":2,"kind":"investigate","text":"Confirm with larger sample if 0.4pp gap V0 vs V1 real"}
    ],
    "context":{"process":"Yoke ass'y adhesion at V0/V1 jig, main 2 final + function NG comparison","stage":"E2 main line 25/1/2024","baselineReason":"Normal pool combines V0+V1 — paired reference"}
},
"tr_ko":{
    "headline":"Yoke V0/V1 지그 기능 NG 0.2-0.6% vs Normal 0.6% — 두 지그 동등.",
    "actions":[
        {"priority":1,"kind":"action","text":"V0 또는 V1 지그 모두 가능 — 단순하면 V1 유지"},
        {"priority":2,"kind":"investigate","text":"표본 확대해 V0-V1 0.4pp 차이 유의성 확인"}
    ],
    "context":{"process":"V0/V1 지그 Yoke 접착, main 2 최종 + 기능 NG 비교","stage":"E2 메인라인 2024-01-25","baselineReason":"Normal이 V0+V1 합산 — 페어 기준"}
},
"tr_vi":{
    "headline":"NG function Yoke jig V0/V1 0.2-0.6% vs Normal 0.6% — hai jig tương đương.",
    "actions":[
        {"priority":1,"kind":"action","text":"V0 hoặc V1 jig đều dùng được — giữ V1 nếu đơn giản hơn"},
        {"priority":2,"kind":"investigate","text":"Tăng cỡ mẫu để xác nhận chênh 0.4pp V0 vs V1"}
    ],
    "context":{"process":"Dán Yoke trên jig V0/V1, main 2 final + so sánh NG function","stage":"Line E2 main 25/01/2024","baselineReason":"Normal gộp V0+V1 — chuẩn so sánh"}
}
})

# 8: BRS-161014 VP mold #2 vs #5 (comparison or quality_log small N)
DATASETS.append({
"name":"11. BRS-161014 Report test material VP mold #2  25.01.2024",
"result":{
    "productType":"BRS-161014",
    "measurements":[
        {"productType":"BRS-161014","testDate":"2024-01-25","line":"","checkType":"process","variable":"VP Mold #2","variableDetail":"Main Vision VP/CD","variableGroup":"test","intervention":"mold #2","inputQty":200,"okQty":199,"ngTotal":1,"ngRate":0.5,"defectCategory":"assembly_defect","defectType":"CD offset","defectCount":1},
        {"productType":"BRS-161014","testDate":"2024-01-25","line":"","checkType":"process","variable":"VP Mold #5","variableDetail":"Main Vision VP/CD","variableGroup":"test","intervention":"mold #5","inputQty":200,"okQty":200,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
    ],
    "tags":["brs-161014","vp-mold","mold-2","mold-5","cd-offset","material-test","multi-arm"],
    "reportType":"multi_arm",
    "verdict":"no_clear_effect",
    "headline":"VP mold #2 NG 0.5% (1/200, CD offset); mold #5 0.0% (0/200) — within sample noise.",
    "evidence":[
        {"metric":"Sub1+Main vision NG","baselineLabel":"","baselineValue":"","variantLabel":"","variantValue":"",
         "deltaText":"+0.5pp range","deltaSign":"up","note":"n=200 each",
         "comparisons":[
            {"label":"VP Mold #2","value":"0.5% (1/200)","n":200,"isBaseline":False,"isBest":False,"isWorst":True},
            {"label":"VP Mold #5","value":"0.0% (0/200)","n":200,"isBaseline":False,"isBest":True,"isWorst":False}
         ],
         "bestLabel":"VP Mold #5","worstLabel":"VP Mold #2"}
    ],
    "actions":[
        {"priority":1,"kind":"action","text":"Both molds acceptable; continue monitor #2 CD offset"},
        {"priority":2,"kind":"investigate","text":"Inspect mold #2 cavity for CD offset risk"}
    ],
    "context":{"process":"Check VP bending defect type per mold # at Sub1 + Main vision","stage":"25/1/2024","baselineReason":"no Normal — paired mold #2 vs #5"}
},
"tr_ko":{
    "headline":"VP 몰드 #2 NG 0.5% (1/200, CD offset); 몰드 #5 0.0% (0/200) — 표본 노이즈 범위.",
    "actions":[
        {"priority":1,"kind":"action","text":"두 몰드 모두 수용 가능; #2 CD offset 지속 모니터링"},
        {"priority":2,"kind":"investigate","text":"몰드 #2 캐비티 CD offset 위험 점검"}
    ],
    "context":{"process":"몰드 별 VP 벤딩 결함 유형 점검 (Sub1 + Main 비전)","stage":"2024-01-25","baselineReason":"Normal 없음 — 몰드 #2 vs #5 비교"}
},
"tr_vi":{
    "headline":"VP mold #2 NG 0.5% (1/200, CD offset); mold #5 0.0% (0/200) — trong dải nhiễu mẫu.",
    "actions":[
        {"priority":1,"kind":"action","text":"Cả hai mold đều chấp nhận; tiếp tục theo dõi CD offset của #2"},
        {"priority":2,"kind":"investigate","text":"Kiểm tra khuôn cavity của mold #2 nguy cơ CD offset"}
    ],
    "context":{"process":"Kiểm tra loại defect VP bending theo từng mold tại Sub1 + Main vision","stage":"25/01/2024","baselineReason":"không có Normal — so sánh mold #2 vs #5"}
}
})

# 9: BRS-161016 DT SUS-ARRAY CSY Tech Vina
DATASETS.append({
"name":"11. BRS-161016 DT Report test  suspension used SUS-ARRAY of new vender CSY Tech Vina date 9.10.2024",
"result":{
    "productType":"BRS-161016 DT",
    "measurements":[
        # Vision suspension
        {"productType":"BRS-161016 DT","testDate":"2024-10-09","line":"","checkType":"visual_inspection","variable":"Test Sus new vender (CSY)","variableDetail":"Vision Suspension at Sub3","variableGroup":"new_lot","intervention":"CSY Tech Vina SUS-ARRAY","inputQty":1125,"okQty":1125,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161016 DT","testDate":"2024-10-09","line":"","checkType":"visual_inspection","variable":"Normal","variableDetail":"Vision Suspension at Sub3","variableGroup":"normal","intervention":"","inputQty":1000,"okQty":1000,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        # Vision spot before repair
        {"productType":"BRS-161016 DT","testDate":"2024-10-09","line":"","checkType":"visual_inspection","variable":"Test Sus new vender (CSY)","variableDetail":"Vision spot before repair","variableGroup":"new_lot","intervention":"CSY Tech Vina SUS-ARRAY","inputQty":1125,"okQty":1114,"ngTotal":11,"ngRate":1.0,"defectCategory":"assembly_defect","defectType":"Solder weak","defectCount":11},
        {"productType":"BRS-161016 DT","testDate":"2024-10-09","line":"","checkType":"visual_inspection","variable":"Normal","variableDetail":"Vision spot before repair","variableGroup":"normal","intervention":"","inputQty":1000,"okQty":994,"ngTotal":6,"ngRate":0.6,"defectCategory":"assembly_defect","defectType":"Solder weak","defectCount":6},
        # Function
        {"productType":"BRS-161016 DT","testDate":"2024-10-09","line":"","checkType":"function","variable":"Test Sus new vender (CSY)","variableDetail":"function","variableGroup":"new_lot","intervention":"CSY Tech Vina SUS-ARRAY","inputQty":1094,"okQty":1070,"ngTotal":24,"ngRate":2.2,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":23},
        {"productType":"BRS-161016 DT","testDate":"2024-10-09","line":"","checkType":"function","variable":"Normal","variableDetail":"function","variableGroup":"normal","intervention":"","inputQty":995,"okQty":980,"ngTotal":15,"ngRate":1.5,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":15},
    ],
    "tags":["brs-161016-dt","suspension","sus-array","new-vender","csy-tech-vina","comparison-study","solder-weak"],
    "reportType":"comparison_study",
    "verdict":"no_clear_effect",
    "headline":"CSY SUS-ARRAY suspension vs Normal: vision OK; spot solder weak 1.0% vs 0.6%; function 2.2% vs 1.5%.",
    "evidence":[
        {"metric":"Vision suspension NG","baselineLabel":"Normal","baselineValue":"0.0% (0/1000)","variantLabel":"CSY","variantValue":"0.0% (0/1125)",
         "deltaText":"+0pp","deltaSign":"no_change","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Vision spot (before repair)","baselineLabel":"Normal","baselineValue":"0.6% (6/1000)","variantLabel":"CSY","variantValue":"1.0% (11/1125)",
         "deltaText":"+0.4pp","deltaSign":"up","note":"all solder weak","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Function NG","baselineLabel":"Normal","baselineValue":"1.5% (15/995)","variantLabel":"CSY","variantValue":"2.2% (24/1094)",
         "deltaText":"+0.7pp","deltaSign":"up","note":"Noise dominant","comparisons":None,"bestLabel":"","worstLabel":""}
    ],
    "actions":[
        {"priority":1,"kind":"action","text":"Approve CSY Tech Vina SUS-ARRAY for production (within tolerance)"},
        {"priority":2,"kind":"investigate","text":"Monitor solder-weak +0.4pp and Noise +0.7pp in next lot"}
    ],
    "context":{"process":"Suspension vision Sub3 + spot vision + function check","stage":"09/10/2024","baselineReason":"same-event paired Normal row available"}
},
"tr_ko":{
    "headline":"CSY SUS-ARRAY 서스펜션 vs Normal: 비전 OK; 스팟 솔더 약 1.0% vs 0.6%; 기능 2.2% vs 1.5%.",
    "actions":[
        {"priority":1,"kind":"action","text":"CSY Tech Vina SUS-ARRAY 양산 승인 (허용 범위 내)"},
        {"priority":2,"kind":"investigate","text":"다음 로트에서 솔더 약 +0.4pp 및 Noise +0.7pp 모니터링"}
    ],
    "context":{"process":"서스펜션 비전(Sub3) + 스팟 비전 + 기능 검사","stage":"2024-10-09","baselineReason":"동일 시험 내 Normal 행 존재"}
},
"tr_vi":{
    "headline":"Suspension CSY SUS-ARRAY vs Normal: vision OK; spot solder weak 1.0% vs 0.6%; function 2.2% vs 1.5%.",
    "actions":[
        {"priority":1,"kind":"action","text":"Duyệt SUS-ARRAY của CSY Tech Vina cho sản xuất (trong dung sai)"},
        {"priority":2,"kind":"investigate","text":"Theo dõi solder weak +0.4pp và Noise +0.7pp ở lô kế"}
    ],
    "context":{"process":"Vision suspension Sub3 + vision spot + kiểm function","stage":"09/10/2024","baselineReason":"có dòng Normal cùng sự kiện"}
}
})

# 10: BRS-161016 GMI DOE Laser VP main2 - DOE
DATASETS.append({
"name":"11. BRS-161016 GMI Report test DOE Laser VP main 2- 2025.06.25",
"result":{
    "productType":"BRS-161016 GMI",
    "measurements":[
        {"productType":"BRS-161016 GMI","testDate":"2025-06-25","line":"E2-3B","checkType":"process","variable":"Standard (ON=0/OFF=50)","variableDetail":"Laser GAP=50","variableGroup":"normal","intervention":"laser delay ON=0 OFF=50","inputQty":20,"okQty":20,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161016 GMI","testDate":"2025-06-25","line":"E2-3B","checkType":"process","variable":"Test No.1 (ON=5/OFF=50)","variableDetail":"GAP=45","variableGroup":"test","intervention":"laser delay ON=5","inputQty":20,"okQty":20,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161016 GMI","testDate":"2025-06-25","line":"E2-3B","checkType":"process","variable":"Test No.2 (ON=10/OFF=50)","variableDetail":"GAP=40","variableGroup":"test","intervention":"laser delay ON=10","inputQty":20,"okQty":19,"ngTotal":1,"ngRate":5.0,"defectCategory":"assembly_defect","defectType":"Not cut all","defectCount":1},
        {"productType":"BRS-161016 GMI","testDate":"2025-06-25","line":"E2-3B","checkType":"process","variable":"Test No.3 (ON=15/OFF=50)","variableDetail":"GAP=35","variableGroup":"test","intervention":"laser delay ON=15","inputQty":20,"okQty":20,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161016 GMI","testDate":"2025-06-25","line":"E2-3B","checkType":"process","variable":"Test No.4 (ON=20/OFF=50)","variableDetail":"GAP=30","variableGroup":"test","intervention":"laser delay ON=20","inputQty":20,"okQty":20,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161016 GMI","testDate":"2025-06-25","line":"E2-3B","checkType":"process","variable":"Test No.5 (ON=25/OFF=50)","variableDetail":"GAP=25","variableGroup":"test","intervention":"laser delay ON=25","inputQty":20,"okQty":20,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161016 GMI","testDate":"2025-06-25","line":"E2-3B","checkType":"process","variable":"Test No.6 (ON=0/OFF=45)","variableDetail":"GAP=45","variableGroup":"test","intervention":"laser delay OFF=45","inputQty":20,"okQty":20,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161016 GMI","testDate":"2025-06-25","line":"E2-3B","checkType":"process","variable":"Test No.7 (ON=0/OFF=40)","variableDetail":"GAP=40","variableGroup":"test","intervention":"laser delay OFF=40","inputQty":20,"okQty":20,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161016 GMI","testDate":"2025-06-25","line":"E2-3B","checkType":"process","variable":"Test No.8 (ON=0/OFF=35)","variableDetail":"GAP=35","variableGroup":"test","intervention":"laser delay OFF=35","inputQty":20,"okQty":20,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161016 GMI","testDate":"2025-06-25","line":"E2-3B","checkType":"process","variable":"Test No.9 (ON=0/OFF=30)","variableDetail":"GAP=30","variableGroup":"test","intervention":"laser delay OFF=30","inputQty":20,"okQty":20,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161016 GMI","testDate":"2025-06-25","line":"E2-3B","checkType":"process","variable":"Test No.10 (ON=0/OFF=25)","variableDetail":"GAP=25","variableGroup":"test","intervention":"laser delay OFF=25","inputQty":20,"okQty":20,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161016 GMI","testDate":"2025-06-25","line":"E2-3B","checkType":"process","variable":"Test No.11 (ON=0/OFF=55)","variableDetail":"GAP=55","variableGroup":"test","intervention":"laser delay OFF=55","inputQty":20,"okQty":20,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161016 GMI","testDate":"2025-06-25","line":"E2-3B","checkType":"process","variable":"Test No.12 (ON=0/OFF=60)","variableDetail":"GAP=60","variableGroup":"test","intervention":"laser delay OFF=60","inputQty":20,"okQty":20,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161016 GMI","testDate":"2025-06-25","line":"E2-3B","checkType":"process","variable":"Test No.13 (ON=0/OFF=65)","variableDetail":"GAP=65","variableGroup":"test","intervention":"laser delay OFF=65","inputQty":20,"okQty":20,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161016 GMI","testDate":"2025-06-25","line":"E2-3B","checkType":"process","variable":"Test No.14 (ON=0/OFF=70)","variableDetail":"GAP=70","variableGroup":"test","intervention":"laser delay OFF=70","inputQty":20,"okQty":20,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161016 GMI","testDate":"2025-06-25","line":"E2-3B","checkType":"process","variable":"Test No.15 (ON=0/OFF=75)","variableDetail":"GAP=75","variableGroup":"test","intervention":"laser delay OFF=75","inputQty":20,"okQty":20,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
    ],
    "tags":["brs-161016-gmi","doe","laser-delay","vp-cutting-offset","laser-on-off","e2-3b","gap-tuning"],
    "reportType":"doe_factorial",
    "verdict":"inconclusive",
    "headline":"Laser DOE 15 cells x 20pcs: only Test No.2 (ON=10) shows 1/20 Not-cut-all; others 0% — too few NG.",
    "evidence":[
        {"metric":"Best DOE cell","baselineLabel":"Standard","baselineValue":"0/20 (ON=0/OFF=50)","variantLabel":"14 of 15 cells","variantValue":"0/20",
         "deltaText":"+0pp","deltaSign":"no_change","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Worst DOE cell","baselineLabel":"Standard","baselineValue":"0/20","variantLabel":"Test No.2 (ON=10)","variantValue":"1/20 (5%) Not cut all",
         "deltaText":"+5pp","deltaSign":"up","note":"n=20 — single NG","comparisons":None,"bestLabel":"","worstLabel":""}
    ],
    "actions":[
        {"priority":1,"kind":"investigate","text":"Repeat DOE with N>=100 per cell — current N=20 insufficient"},
        {"priority":2,"kind":"action","text":"Keep Standard ON=0/OFF=50 until DOE confirms better cell"}
    ],
    "context":{"process":"DOE Laser ON/OFF delay tuning to reduce VP cutting offset","stage":"E2-3B 25/06/2025","baselineReason":"Standard ON=0/OFF=50 as reference cell"},
    "doeGrid":{
        "factor1Name":"Laser ON delay",
        "factor2Name":"Laser OFF delay",
        "factor1Levels":["0","5","10","15","20","25"],
        "factor2Levels":["25","30","35","40","45","50","55","60","65","70","75"],
        "cells":[
            {"f1":"0","f2":"50","status":"ok","value":"0/20 (Standard)"},
            {"f1":"5","f2":"50","status":"ok","value":"0/20"},
            {"f1":"10","f2":"50","status":"ng","value":"1/20 Not cut all"},
            {"f1":"15","f2":"50","status":"ok","value":"0/20"},
            {"f1":"20","f2":"50","status":"ok","value":"0/20"},
            {"f1":"25","f2":"50","status":"ok","value":"0/20"},
            {"f1":"0","f2":"45","status":"ok","value":"0/20"},
            {"f1":"0","f2":"40","status":"ok","value":"0/20"},
            {"f1":"0","f2":"35","status":"ok","value":"0/20"},
            {"f1":"0","f2":"30","status":"ok","value":"0/20"},
            {"f1":"0","f2":"25","status":"ok","value":"0/20"},
            {"f1":"0","f2":"55","status":"ok","value":"0/20"},
            {"f1":"0","f2":"60","status":"ok","value":"0/20"},
            {"f1":"0","f2":"65","status":"ok","value":"0/20"},
            {"f1":"0","f2":"70","status":"ok","value":"0/20"},
            {"f1":"0","f2":"75","status":"ok","value":"0/20"}
        ]
    }
},
"tr_ko":{
    "headline":"레이저 DOE 15 셀 x 20pcs: Test No.2 (ON=10)에서 1/20 Not-cut-all 외 모두 0% — NG 표본 부족.",
    "actions":[
        {"priority":1,"kind":"investigate","text":"셀당 N>=100으로 DOE 재실시 — 현재 N=20 불충분"},
        {"priority":2,"kind":"action","text":"DOE 재실시까지 Standard ON=0/OFF=50 유지"}
    ],
    "context":{"process":"VP 컷팅 오프셋 감소 위한 Laser ON/OFF 딜레이 DOE","stage":"E2-3B 2025-06-25","baselineReason":"Standard ON=0/OFF=50 셀을 기준"}
},
"tr_vi":{
    "headline":"DOE Laser 15 cell x 20pcs: chỉ Test No.2 (ON=10) NG 1/20 Not-cut-all, còn lại 0% — quá ít NG.",
    "actions":[
        {"priority":1,"kind":"investigate","text":"Lặp DOE với N>=100/cell — N=20 hiện tại không đủ"},
        {"priority":2,"kind":"action","text":"Giữ Standard ON=0/OFF=50 đến khi DOE xác nhận cell tốt hơn"}
    ],
    "context":{"process":"DOE delay Laser ON/OFF để giảm cutting offset VP","stage":"E2-3B 25/06/2025","baselineReason":"Cell Standard ON=0/OFF=50 làm tham chiếu"}
}
})

# 11: BRS-161016 YK over flatness - reliability_validation (multiple tests)
DATASETS.append({
"name":"11. BRS-161016 Report test material YK happen  NG over flatness  Date 11.7.2025",
"result":{
    "productType":"BRS-161016",
    "measurements":[
        # 1. Decap MG-S-A/MG-S-B
        {"productType":"BRS-161016","testDate":"2025-07-11","line":"","checkType":"process","variable":"Test YK NG over flatness","variableDetail":"Decap MG-S-A bond spread","variableGroup":"test","intervention":"YK NG over flatness","inputQty":10,"okQty":10,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
        {"productType":"BRS-161016","testDate":"2025-07-11","line":"","checkType":"process","variable":"Test YK NG over flatness","variableDetail":"Decap MG-S-B bond spread","variableGroup":"test","intervention":"YK NG over flatness","inputQty":10,"okQty":10,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
        {"productType":"BRS-161016","testDate":"2025-07-11","line":"","checkType":"process","variable":"Normal","variableDetail":"Decap MG-S-A bond spread","variableGroup":"normal","intervention":"","inputQty":8,"okQty":8,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
        # 2. Decap Yoke
        {"productType":"BRS-161016","testDate":"2025-07-11","line":"","checkType":"process","variable":"Test YK NG over flatness","variableDetail":"Decap Yoke bond spread","variableGroup":"test","intervention":"YK NG over flatness","inputQty":10,"okQty":10,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
        {"productType":"BRS-161016","testDate":"2025-07-11","line":"","checkType":"process","variable":"Normal","variableDetail":"Decap Yoke bond spread","variableGroup":"normal","intervention":"","inputQty":8,"okQty":8,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
        # 3. Drop test
        {"productType":"BRS-161016","testDate":"2025-07-11","line":"","checkType":"process","variable":"Test YK NG over flatness","variableDetail":"Drop test","variableGroup":"test","intervention":"YK NG over flatness","inputQty":5,"okQty":5,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        {"productType":"BRS-161016","testDate":"2025-07-11","line":"","checkType":"process","variable":"Normal","variableDetail":"Drop test","variableGroup":"normal","intervention":"","inputQty":5,"okQty":5,"ngTotal":0,"ngRate":0.0,"defectCategory":"","defectType":"","defectCount":0},
        # 5. Gauss
        {"productType":"BRS-161016","testDate":"2025-07-11","line":"","checkType":"process","variable":"Test YK NG over flatness","variableDetail":"Gauss semi yoke (spec >=480G)","variableGroup":"test","intervention":"YK NG over flatness","inputQty":469,"okQty":462,"ngTotal":7,"ngRate":1.5,"defectCategory":"magnetic_defect","defectType":"Low Gauss","defectCount":7},
        {"productType":"BRS-161016","testDate":"2025-07-11","line":"","checkType":"process","variable":"Normal","variableDetail":"Gauss semi yoke (spec >=480G)","variableGroup":"normal","intervention":"","inputQty":230,"okQty":230,"ngTotal":0,"ngRate":0.0,"defectCategory":"magnetic_defect","defectType":"","defectCount":0},
        # 6. Function
        {"productType":"BRS-161016","testDate":"2025-07-11","line":"C2-3A","checkType":"function","variable":"Test YK NG over flatness","variableDetail":"function","variableGroup":"test","intervention":"YK NG over flatness","inputQty":430,"okQty":420,"ngTotal":10,"ngRate":2.33,"defectCategory":"function_hearing","defectType":"Noise+SPL+THD+Touch","defectCount":10},
        {"productType":"BRS-161016","testDate":"2025-07-11","line":"C2-3A","checkType":"function","variable":"Normal","variableDetail":"function","variableGroup":"normal","intervention":"","inputQty":348,"okQty":342,"ngTotal":6,"ngRate":1.72,"defectCategory":"function_hearing","defectType":"Noise+Touch","defectCount":6},
        # NG decap analyse (Hearing Noise+Touch)
        {"productType":"BRS-161016","testDate":"2025-07-11","line":"3B","checkType":"function","variable":"Decap analyse NG Hearing","variableDetail":"FR+coil offset","variableGroup":"test","intervention":"","inputQty":7,"okQty":4,"ngTotal":3,"ngRate":42.86,"defectCategory":"function_hearing","defectType":"FR+coil offset","defectCount":3},
        {"productType":"BRS-161016","testDate":"2025-07-11","line":"3B","checkType":"function","variable":"Decap analyse NG Air leak","variableDetail":"SMG separate","variableGroup":"test","intervention":"","inputQty":3,"okQty":2,"ngTotal":1,"ngRate":33.33,"defectCategory":"function_thd","defectType":"SMG separate","defectCount":1},
    ],
    "tags":["brs-161016","yoke-over-flatness","material-test","decap","drop-test","tension","gauss","function-ng","reliability-validation"],
    "reportType":"reliability_validation",
    "verdict":"failed",
    "headline":"YK over-flatness reliability: Gauss 1.5% Low (7/469 below 480G); decap/drop/tension/function pass — Gauss fails.",
    "evidence":[
        {"metric":"Decap bond MG-S-A/B/Yoke","baselineLabel":"Spec >=80% spread","baselineValue":"100% (0 NG)","variantLabel":"Test YK","variantValue":"100% (0/10)",
         "deltaText":"—","deltaSign":"no_change","note":"passed","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Gauss semi Yoke","baselineLabel":"Spec >=480G","baselineValue":"Normal 0% Low Gauss (0/230)","variantLabel":"Test YK","variantValue":"1.5% Low Gauss (7/469)",
         "deltaText":"+1.5pp","deltaSign":"up","note":"only test arm fails","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Tension MG-S-A/B/C","baselineLabel":"Spec >=2.5/5.0/80 kgf","baselineValue":"Normal AVG 6.18/11.22/86.36 kgf","variantLabel":"Test YK","variantValue":"AVG 7.21/11.65/76.84 kgf",
         "deltaText":"—","deltaSign":"no_change","note":"all OK","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Function NG","baselineLabel":"Normal","baselineValue":"1.72% (6/348)","variantLabel":"Test YK","variantValue":"2.33% (10/430)",
         "deltaText":"+0.61pp","deltaSign":"up","note":"hearing-noise dominant","comparisons":None,"bestLabel":"","worstLabel":""}
    ],
    "actions":[
        {"priority":1,"kind":"action","text":"Do NOT release YK over-flatness material; Gauss fails spec >=480G"},
        {"priority":2,"kind":"investigate","text":"Cross-check MG-C Ruijin lot causing low Gauss"},
        {"priority":3,"kind":"risk","text":"Common lot material MG-C may affect normal production"}
    ],
    "context":{"process":"Semi Yoke material reliability validation (decap, drop, tension, gauss, function)","stage":"Sub line + C2-3A function 11/7/2025","baselineReason":"same-event Normal arm + Spec gates for each test"}
},
"tr_ko":{
    "headline":"YK 평면도 NG 신뢰성: Gauss 1.5% Low (7/469, 480G 미달); 디캡/낙하/장력/기능 합격 — Gauss 실패.",
    "actions":[
        {"priority":1,"kind":"action","text":"YK 평면도 NG 자재 출하 금지; Gauss 스펙 480G 미충족"},
        {"priority":2,"kind":"investigate","text":"Low Gauss 원인인 MG-C Ruijin 로트 교차 점검"},
        {"priority":3,"kind":"risk","text":"공통 자재 MG-C가 일반 생산에도 영향 가능"}
    ],
    "context":{"process":"Semi Yoke 자재 신뢰성 검증 (디캡, 낙하, 장력, Gauss, 기능)","stage":"서브라인 + C2-3A 기능 라인 2025-07-11","baselineReason":"동일 시험 내 Normal arm + 각 항목 스펙 게이트"}
},
"tr_vi":{
    "headline":"Vật liệu YK NG over flatness: Gauss Low 1.5% (7/469 dưới 480G); decap/drop/tension/function đạt — Gauss fail.",
    "actions":[
        {"priority":1,"kind":"action","text":"Không xuất vật liệu YK over flatness; Gauss không đạt spec >=480G"},
        {"priority":2,"kind":"investigate","text":"Kiểm tra chéo lô MG-C Ruijin gây Low Gauss"},
        {"priority":3,"kind":"risk","text":"Vật liệu MG-C dùng chung có thể ảnh hưởng sản xuất bình thường"}
    ],
    "context":{"process":"Kiểm chứng độ tin cậy vật liệu Semi Yoke (decap, drop, tension, gauss, function)","stage":"Sub line + line function C2-3A 11/7/2025","baselineReason":"có Normal arm cùng sự kiện + cổng spec từng hạng mục"}
}
})

# 12: BRS-161016 TF dry SMG+Yoke DOE temperature x time (Machine)
DATASETS.append({
"name":"11. BRS-161016 TF Report test change condition dry SMG+Yoke",
"result":{
    "productType":"BRS-161016 TF",
    "measurements":[
        {"productType":"BRS-161016 TF","testDate":"2025-01-08","line":"","checkType":"process","variable":"320C / 7min15s / Machine 1","variableDetail":"NG SMG separate","variableGroup":"test","intervention":"dry 320 7m15s M1","inputQty":12,"okQty":11,"ngTotal":1,"ngRate":8.3,"defectCategory":"assembly_defect","defectType":"SMG+Yoke separate","defectCount":1},
        {"productType":"BRS-161016 TF","testDate":"2025-01-08","line":"","checkType":"process","variable":"320C / 7min15s / Machine 2","variableDetail":"NG SMG separate","variableGroup":"test","intervention":"dry 320 7m15s M2","inputQty":11,"okQty":11,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
        {"productType":"BRS-161016 TF","testDate":"2025-01-08","line":"","checkType":"process","variable":"320C / 8min20s / Machine 1","variableDetail":"NG SMG separate","variableGroup":"test","intervention":"dry 320 8m20s M1","inputQty":15,"okQty":15,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
        {"productType":"BRS-161016 TF","testDate":"2025-01-08","line":"","checkType":"process","variable":"320C / 8min20s / Machine 2","variableDetail":"NG SMG separate","variableGroup":"test","intervention":"dry 320 8m20s M2","inputQty":15,"okQty":15,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
        {"productType":"BRS-161016 TF","testDate":"2025-01-08","line":"","checkType":"process","variable":"340C / 7min15s / Machine 1","variableDetail":"NG SMG separate","variableGroup":"test","intervention":"dry 340 7m15s M1","inputQty":24,"okQty":22,"ngTotal":2,"ngRate":8.3,"defectCategory":"assembly_defect","defectType":"SMG+Yoke separate","defectCount":2},
        {"productType":"BRS-161016 TF","testDate":"2025-01-08","line":"","checkType":"process","variable":"340C / 7min15s / Machine 2","variableDetail":"NG SMG separate","variableGroup":"test","intervention":"dry 340 7m15s M2","inputQty":24,"okQty":22,"ngTotal":2,"ngRate":8.3,"defectCategory":"assembly_defect","defectType":"SMG+Yoke separate","defectCount":2},
        {"productType":"BRS-161016 TF","testDate":"2025-01-08","line":"","checkType":"process","variable":"350C / 7min15s / Machine 1","variableDetail":"NG SMG separate","variableGroup":"test","intervention":"dry 350 7m15s M1","inputQty":94,"okQty":94,"ngTotal":0,"ngRate":0.0,"defectCategory":"assembly_defect","defectType":"","defectCount":0},
        {"productType":"BRS-161016 TF","testDate":"2025-01-08","line":"","checkType":"process","variable":"350C / 7min15s / Machine 2","variableDetail":"NG SMG separate","variableGroup":"test","intervention":"dry 350 7m15s M2","inputQty":90,"okQty":89,"ngTotal":1,"ngRate":1.1,"defectCategory":"assembly_defect","defectType":"SMG+Yoke separate","defectCount":1},
    ],
    "tags":["brs-161016-tf","doe-factorial","dry-temperature","dry-time","smg-yoke-separate","heating-press","intervention-test"],
    "reportType":"doe_factorial",
    "verdict":"partial",
    "headline":"Dry temp x time DOE: 320C/8m20s and 350C/7m15s reach 0% NG; 340C still 8.3%.",
    "evidence":[
        {"metric":"Best cell","baselineLabel":"","baselineValue":"","variantLabel":"320C/8m20s (both M)","variantValue":"0/30 (0.0%)",
         "deltaText":"—","deltaSign":"no_change","note":"","comparisons":None,"bestLabel":"","worstLabel":""},
        {"metric":"Worst cell","baselineLabel":"","baselineValue":"","variantLabel":"340C/7m15s","variantValue":"8.3% (4/48)",
         "deltaText":"+8.3pp","deltaSign":"up","note":"both machines","comparisons":None,"bestLabel":"","worstLabel":""}
    ],
    "actions":[
        {"priority":1,"kind":"action","text":"Adopt 320C/8min20s or 350C/7min15s as new dry condition"},
        {"priority":2,"kind":"investigate","text":"Confirm 350C M2 1.1% NG single sample with larger N"}
    ],
    "context":{"process":"Dry SMG+Yoke temperature x time tuning to reduce SMG+Yoke separate","stage":"Heating press machines 1+2 on TF line 08/01/2025","baselineReason":"DOE — no single baseline; 320C/7m15s M1 = starting condition"},
    "doeGrid":{
        "factor1Name":"Dry temperature (C)",
        "factor2Name":"Dry time / Machine",
        "factor1Levels":["320","340","350"],
        "factor2Levels":["7m15s/M1","7m15s/M2","8m20s/M1","8m20s/M2"],
        "cells":[
            {"f1":"320","f2":"7m15s/M1","status":"ng","value":"8.3% (1/12)"},
            {"f1":"320","f2":"7m15s/M2","status":"ok","value":"0.0% (0/11)"},
            {"f1":"320","f2":"8m20s/M1","status":"ok","value":"0.0% (0/15)"},
            {"f1":"320","f2":"8m20s/M2","status":"ok","value":"0.0% (0/15)"},
            {"f1":"340","f2":"7m15s/M1","status":"ng","value":"8.3% (2/24)"},
            {"f1":"340","f2":"7m15s/M2","status":"ng","value":"8.3% (2/24)"},
            {"f1":"350","f2":"7m15s/M1","status":"ok","value":"0.0% (0/94)"},
            {"f1":"350","f2":"7m15s/M2","status":"borderline","value":"1.1% (1/90)"}
        ]
    }
},
"tr_ko":{
    "headline":"건조 온도 x 시간 DOE: 320C/8m20s, 350C/7m15s NG 0%; 340C는 여전히 8.3%.",
    "actions":[
        {"priority":1,"kind":"action","text":"건조 조건을 320C/8min20s 또는 350C/7min15s로 변경"},
        {"priority":2,"kind":"investigate","text":"350C M2 1.1% NG는 표본 확대 후 재확인"}
    ],
    "context":{"process":"SMG+Yoke 분리 개선 위한 건조 온도 x 시간 튜닝","stage":"TF 라인 히팅 프레스 M1+M2, 2025-01-08","baselineReason":"DOE — 단일 기준 없음; 320C/7m15s M1이 출발 조건"}
},
"tr_vi":{
    "headline":"DOE nhiệt sấy x thời gian: 320C/8m20s và 350C/7m15s đạt NG 0%; 340C vẫn 8.3%.",
    "actions":[
        {"priority":1,"kind":"action","text":"Áp dụng điều kiện sấy 320C/8min20s hoặc 350C/7min15s"},
        {"priority":2,"kind":"investigate","text":"Xác nhận lại 350C M2 1.1% NG với mẫu lớn hơn"}
    ],
    "context":{"process":"Tinh chỉnh nhiệt độ x thời gian sấy SMG+Yoke để giảm tách lớp","stage":"Máy heating press 1+2 line TF 08/01/2025","baselineReason":"DOE — không có chuẩn đơn; 320C/7m15s M1 là điều kiện khởi đầu"}
}
})

# ============ Run ============
ok = 0; fail = 0
for d in DATASETS:
    if commit(d["name"], d["result"], d["tr_ko"], d["tr_vi"]):
        ok += 1
    else:
        fail += 1
con.close()
print(f"=== chunk_03 DONE === OK={ok} FAIL={fail} SKIP=0")
