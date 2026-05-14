"""Helper for chunk_10 batch — fetch paste, then later commit results."""
import sqlite3, json, sys, os, datetime

DB = r"D:\000. MyWorks\002. DB\process-review.db"
CHUNK = r"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_batch_chunks\chunk_10.txt"

def cmd_fetch():
    with open(CHUNK, "r", encoding="utf-8") as f:
        names = [ln.strip() for ln in f if ln.strip()]
    con = sqlite3.connect(DB)
    con.execute("PRAGMA busy_timeout = 30000")
    cur = con.cursor()
    out = {}
    for n in names:
        row = cur.execute("SELECT ExtractedText FROM RawReportText WHERE DatasetName=? AND Kind='excel_paste'", (n,)).fetchone()
        out[n] = row[0] if row and row[0] else None
    con.close()
    with open(r"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_tmp_pastes.json", "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False)
    for n, v in out.items():
        print(f"{n}\t{'OK' if v else 'MISSING'}\t{len(v) if v else 0}")

def cmd_lens():
    with open(r"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_tmp_pastes.json", "r", encoding="utf-8") as f:
        d = json.load(f)
    for n,v in d.items():
        print(f"=== {n} === len={len(v) if v else 0}")

def cmd_show(idx):
    with open(r"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_tmp_pastes.json", "r", encoding="utf-8") as f:
        d = json.load(f)
    names = list(d.keys())
    name = names[int(idx)]
    print(f"### {name} ###")
    print(d[name] or "(empty)")

def cmd_commit():
    # Reads _tmp_results.json: {name: {result: {...}, tr_ko:{...}, tr_vi:{...}}}
    with open(r"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_tmp_results.json", "r", encoding="utf-8") as f:
        all_data = json.load(f)
    con = sqlite3.connect(DB)
    con.execute("PRAGMA busy_timeout = 30000")
    cur = con.cursor()
    ok, fail = 0, 0
    for name, payload in all_data.items():
        try:
            result = payload["result"]
            tr_ko = payload.get("tr_ko")
            tr_vi = payload.get("tr_vi")
            cur.execute("BEGIN")
            now = datetime.datetime.utcnow().isoformat() + "Z"
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
            ok += 1
            print(f"[OK] {name}")
        except Exception as e:
            con.rollback()
            fail += 1
            print(f"[FAIL] {name}: {e}")
    con.close()
    print(f"=== commit done ok={ok} fail={fail} ===")

if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "fetch"
    if cmd == "fetch": cmd_fetch()
    elif cmd == "lens": cmd_lens()
    elif cmd == "show": cmd_show(sys.argv[2])
    elif cmd == "commit": cmd_commit()
