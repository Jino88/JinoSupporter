"""Helper for chunk_05 batch. Subcommands: fetch / commit / list."""
import sqlite3, sys, json, os, datetime, io

DB = r"D:\000. MyWorks\002. DB\process-review.db"
CHUNK = r"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_batch_chunks\chunk_05.txt"

def names():
    with open(CHUNK, "r", encoding="utf-8") as f:
        return [ln.strip() for ln in f if ln.strip()]

def fetch(idx):
    n = names()[idx]
    con = sqlite3.connect(DB)
    con.execute("PRAGMA busy_timeout = 30000")
    row = con.execute(
        "SELECT ExtractedText FROM RawReportText WHERE DatasetName=? AND Kind='excel_paste'",
        (n,)).fetchone()
    con.close()
    if not row or not row[0]:
        print(f"[SKIP {n}] no excel_paste")
        return
    text = row[0]
    out = rf"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_tmp_paste_{idx:02d}.txt"
    with open(out, "w", encoding="utf-8") as f:
        f.write(f"### NAME: {n}\n")
        f.write(text)
    print(f"[OK fetched {idx} {n}] -> {out} ({len(text)} chars)")

def fetch_all():
    ns = names()
    con = sqlite3.connect(DB)
    con.execute("PRAGMA busy_timeout = 30000")
    for i, n in enumerate(ns):
        row = con.execute(
            "SELECT ExtractedText FROM RawReportText WHERE DatasetName=? AND Kind='excel_paste'",
            (n,)).fetchone()
        if not row or not row[0]:
            print(f"[{i:02d}] SKIP {n}")
            continue
        out = rf"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_tmp_paste_{i:02d}.txt"
        with open(out, "w", encoding="utf-8") as f:
            f.write(f"### NAME: {n}\n")
            f.write(row[0])
        print(f"[{i:02d}] OK {n} ({len(row[0])} chars) -> _tmp_paste_{i:02d}.txt")
    con.close()

def commit_one(name, payload_path):
    """payload_path: JSON file containing result, tr_ko, tr_vi for one dataset."""
    with open(payload_path, "r", encoding="utf-8") as f:
        bundle = json.load(f)
    result = bundle["result"]
    tr_ko  = bundle.get("tr_ko") or {}
    tr_vi  = bundle.get("tr_vi") or {}

    con = sqlite3.connect(DB)
    con.execute("PRAGMA busy_timeout = 30000")
    cur = con.cursor()
    try:
        cur.execute("BEGIN")
        now = datetime.datetime.utcnow().isoformat() + "Z"
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
                 "", "", "", "", "", "", "",
                 tr.get("headline",""), tr_actions_json, tr_context_json, now))

        con.commit()
        print(f"[OK COMMIT {name}]")
    except Exception as e:
        con.rollback()
        print(f"[FAIL COMMIT {name}] {e}")
        raise
    finally:
        con.close()

if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "fetch_all":
        fetch_all()
    elif cmd == "fetch":
        fetch(int(sys.argv[2]))
    elif cmd == "commit":
        commit_one(sys.argv[2], sys.argv[3])
    elif cmd == "list":
        for i,n in enumerate(names()):
            print(f"{i:02d}\t{n}")
    else:
        print("usage: fetch_all | fetch <idx> | commit <name> <payload.json> | list")
