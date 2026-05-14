"""Helper for chunk_07 batch. Reads excel_paste TSVs and writes one per file
so the agent can inspect them. Then provides a commit function."""
import sqlite3
import json
import sys
import os
from datetime import datetime, timezone

DB = r"D:\000. MyWorks\002. DB\process-review.db"

NAMES = [
"13. MSU-L20S15-07DT  Report Test Supension B CSY TECH VINA send some kinds of samples for verify date 24.5.2025",
"13. TIU C11-20  Report test VP find reason NG function high 2026.1.12",
"13. TIU C11-20 Result test Frame load tray 2025.12.18",
"13. TIU L5S3-01 R Result check dimention F-PCB  date 2025.12.01",
"13.1 BRS-161016 Report Test PT 161014-S of Press line (Doojin coating) happen  NG dimension 2nd  22.10.2025 -",
"13.1 BRS-161016 Report checking Reason  low gauss date 14.7.2025",
"13.BRS-201506 Report  test bonding PAD improve quaility more good  date 19.1.2025",
"13.BRS-201506 Report Test clean material CD improve NG tension CD+Coil date 30.9.2024",
"14. BRS-161014  GMI Report test YK happen deform date 21.10.2024",
"14. BRS-161014 DT Report test change MC open test  from manual to MC auto  date 29.1.2024",
"14. BRS-161014 Report check check dimension coil 3D Date 4.12.2023 -",
"14. BRS-161014 Report test sample separate VP Frame",
"14. BRS-161014 Report test tracking process main 2 2023.09.06",
]

OUT_DIR = r"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_tmp_tsv07"

def dump_all():
    os.makedirs(OUT_DIR, exist_ok=True)
    con = sqlite3.connect(DB)
    con.execute("PRAGMA busy_timeout = 30000")
    cur = con.cursor()
    for i, n in enumerate(NAMES):
        r = cur.execute("SELECT ExtractedText FROM RawReportText WHERE DatasetName=? AND Kind='excel_paste'", (n,)).fetchone()
        if not r or not r[0]:
            continue
        path = os.path.join(OUT_DIR, f"{i:02d}.txt")
        with open(path, "w", encoding="utf-8") as f:
            f.write(n + "\n=====\n")
            f.write(r[0])
    con.close()

def commit_one(name, result, tr_ko, tr_vi):
    con = sqlite3.connect(DB)
    con.execute("PRAGMA busy_timeout = 30000")
    cur = con.cursor()
    try:
        cur.execute("BEGIN")
        now = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
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

        tags_json     = json.dumps(result.get("tags")     or [], ensure_ascii=False)
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
        print(f"[OK {name[:50]}]")
    except Exception as e:
        con.rollback()
        print(f"[PARSE-FAIL {name[:50]}] {e}")
        raise
    finally:
        con.close()

if __name__ == "__main__":
    if sys.argv[1] == "dump":
        dump_all()
    elif sys.argv[1] == "commit":
        # commit reads payload json from file path
        payload_path = sys.argv[2]
        with open(payload_path, "r", encoding="utf-8") as f:
            payload = json.load(f)
        commit_one(payload["name"], payload["result"], payload["tr_ko"], payload["tr_vi"])
