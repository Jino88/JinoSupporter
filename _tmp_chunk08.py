"""Helper for chunk_08 batch processing.
Phase 1: dump excel_paste TSV for each dataset name into JSON.
Phase 2: write the agent-produced normalize+translate JSON back to DB.
"""
import sqlite3, json, sys, os, datetime, pathlib

DB = r"D:\000. MyWorks\002. DB\process-review.db"
CHUNK = r"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_batch_chunks\chunk_08.txt"
OUT_DIR = r"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_batch_chunks\pastes_08"
RESULT_DIR = r"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_batch_chunks\results_08"

def names():
    with open(CHUNK, "r", encoding="utf-8") as f:
        return [ln.strip() for ln in f if ln.strip()]

def cmd_dump():
    os.makedirs(OUT_DIR, exist_ok=True)
    con = sqlite3.connect(DB)
    con.execute("PRAGMA busy_timeout=30000")
    cur = con.cursor()
    summary = []
    for i, n in enumerate(names()):
        row = cur.execute(
            "SELECT ExtractedText FROM RawReportText WHERE DatasetName=? AND Kind='excel_paste'",
            (n,)).fetchone()
        if not row or not row[0]:
            summary.append({"idx": i, "name": n, "status": "SKIP_NO_PASTE", "len": 0})
            continue
        txt = row[0]
        path = os.path.join(OUT_DIR, f"{i:02d}.txt")
        with open(path, "w", encoding="utf-8") as f:
            f.write(txt)
        summary.append({"idx": i, "name": n, "status": "OK", "len": len(txt), "path": path})
    con.close()
    print(json.dumps(summary, ensure_ascii=False, indent=2))

def cmd_commit():
    os.makedirs(RESULT_DIR, exist_ok=True)
    con = sqlite3.connect(DB)
    con.execute("PRAGMA busy_timeout=30000")
    cur = con.cursor()
    ok_n = parse_fail_n = 0
    now = datetime.datetime.utcnow().isoformat() + "Z"
    for fn in sorted(os.listdir(RESULT_DIR)):
        if not fn.endswith(".json"): continue
        path = os.path.join(RESULT_DIR, fn)
        try:
            with open(path, "r", encoding="utf-8") as f:
                pkg = json.load(f)
            name = pkg["name"]
            result = pkg["result"]
            tr_ko = pkg.get("tr_ko")
            tr_vi = pkg.get("tr_vi")
        except Exception as e:
            print(f"[PARSE-FAIL load {fn}] {e}")
            parse_fail_n += 1
            continue
        try:
            cur.execute("BEGIN")
            product = result.get("productType", "")
            cur.execute("DELETE FROM NormalizedMeasurements WHERE DatasetName=?", (name,))
            for m in result.get("measurements", []) or []:
                cur.execute("""
                    INSERT INTO NormalizedMeasurements
                      (DatasetName, ProductType, TestDate, Line, CheckType, Variable,
                       VariableDetail, VariableGroup, Intervention, InputQty, OkQty,
                       NgTotal, NgRate, DefectCategory, DefectType, DefectCount, CreatedAt)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (name, m.get("productType") or product, m.get("testDate", ""),
                     m.get("line",""), m.get("checkType",""), m.get("variable",""),
                     m.get("variableDetail",""), m.get("variableGroup",""), m.get("intervention",""),
                     int(m.get("inputQty",0) or 0), int(m.get("okQty",0) or 0), int(m.get("ngTotal",0) or 0),
                     float(m.get("ngRate",0) or 0), m.get("defectCategory",""), m.get("defectType",""),
                     int(m.get("defectCount",0) or 0), now))

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
            ok_n += 1
        except Exception as e:
            con.rollback()
            print(f"[PARSE-FAIL {name}] {e}")
            parse_fail_n += 1
    con.close()
    print(f"=== COMMIT DONE === OK={ok_n} PARSE-FAIL={parse_fail_n}")

if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "dump"
    if cmd == "dump": cmd_dump()
    elif cmd == "commit": cmd_commit()
    else: print("usage: dump|commit")
