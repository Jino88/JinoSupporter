"""Helper for chunk 11 batch processing.

Usage:
  python _tmp_chunk11.py list       - list datasets and presence of excel_paste
  python _tmp_chunk11.py dump <idx> - dump the excel_paste for dataset at line index (1-based)
  python _tmp_chunk11.py commit <idx> <json_path>  - commit result + translations
"""
import sys, os, json, sqlite3, datetime, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

DB = r"D:\000. MyWorks\002. DB\process-review.db"
CHUNK = r"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_batch_chunks\chunk_11.txt"

def targets():
    with open(CHUNK, "r", encoding="utf-8") as f:
        return [ln.strip() for ln in f if ln.strip()]

def connect():
    con = sqlite3.connect(DB)
    con.execute("PRAGMA busy_timeout = 30000")
    return con

def cmd_list():
    con = connect()
    cur = con.cursor()
    for i, name in enumerate(targets(), 1):
        row = cur.execute(
            "SELECT length(ExtractedText) FROM RawReportText WHERE DatasetName=? AND Kind='excel_paste'",
            (name,)).fetchone()
        ln = row[0] if row else None
        print(f"{i:2d}\t{ln}\t{name}")
    con.close()

def cmd_dump(idx):
    name = targets()[int(idx)-1]
    con = connect()
    cur = con.cursor()
    row = cur.execute(
        "SELECT ExtractedText FROM RawReportText WHERE DatasetName=? AND Kind='excel_paste'",
        (name,)).fetchone()
    con.close()
    if not row or not row[0]:
        print(f"[NO_PASTE] {name}", file=sys.stderr)
        sys.exit(2)
    out = rf"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_tmp_chunk11_paste_{int(idx)}.txt"
    with open(out, "w", encoding="utf-8") as f:
        f.write(row[0])
    print(out)
    print(f"NAME={name}")

def cmd_commit(idx, json_path):
    name = targets()[int(idx)-1]
    with open(json_path, "r", encoding="utf-8") as f:
        payload = json.load(f)
    result = payload["result"]
    tr_ko = payload.get("ko")
    tr_vi = payload.get("vi")
    product = result.get("productType", "")
    now = datetime.datetime.utcnow().isoformat() + "Z"
    con = connect()
    cur = con.cursor()
    try:
        cur.execute("BEGIN")
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
            if not tr: continue
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
        print(f"[OK] {name}")
    except Exception as e:
        con.rollback()
        print(f"[FAIL] {name}: {e}", file=sys.stderr)
        sys.exit(3)
    finally:
        con.close()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(1)
    cmd = sys.argv[1]
    if cmd == "list": cmd_list()
    elif cmd == "dump": cmd_dump(sys.argv[2])
    elif cmd == "commit": cmd_commit(sys.argv[2], sys.argv[3])
    else: print(__doc__); sys.exit(1)
