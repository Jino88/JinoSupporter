"""CLI AI Batch helper — handles DB I/O so the agent focuses on analysis JSON.

Usage:
  python _tmp_ai_batch.py extract       # dump all TSVs to _tmp_tsv/NNN.txt + index.json
  python _tmp_ai_batch.py commit NNN    # commit _tmp_results/NNN.json for that index
  python _tmp_ai_batch.py commit_all    # commit all _tmp_results/*.json
  python _tmp_ai_batch.py status        # show progress
"""
import sqlite3, json, os, sys, glob
from datetime import datetime, timezone

DB = r"D:\000. MyWorks\002. DB\process-review.db"
TARGETS = "_batch_selected_20260512_163659_955aeca_1.txt"
TSV_DIR = "_tmp_tsv"
RES_DIR = "_tmp_results"


def load_targets():
    return [n for n in open(TARGETS, encoding="utf-8-sig").read().splitlines() if n.strip()]


def extract():
    names = load_targets()
    con = sqlite3.connect(DB)
    con.execute("PRAGMA busy_timeout = 30000")
    cur = con.cursor()
    index = {}
    for i, n in enumerate(names):
        row = cur.execute(
            "SELECT ExtractedText FROM RawReportText WHERE DatasetName=? AND Kind='excel_paste'",
            (n,)).fetchone()
        if not row or not row[0]:
            print(f"[SKIP {i:03d}] no excel_paste: {n}")
            continue
        path = os.path.join(TSV_DIR, f"{i:03d}.txt")
        with open(path, "w", encoding="utf-8") as f:
            f.write(row[0])
        index[f"{i:03d}"] = n
    with open(os.path.join(TSV_DIR, "index.json"), "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=2)
    con.close()
    print(f"extracted {len(index)} TSVs")


def _commit_one(con, name, payload):
    """payload = {result: {...v7 fields...}, tr_ko: {...}, tr_vi: {...}}"""
    result = payload["result"]
    tr_ko = payload.get("tr_ko")
    tr_vi = payload.get("tr_vi")
    cur = con.cursor()
    cur.execute("BEGIN")
    try:
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
                 m.get("line", ""), m.get("checkType", ""), m.get("variable", ""),
                 m.get("variableDetail", ""), m.get("variableGroup", ""),
                 m.get("intervention", ""),
                 int(m.get("inputQty", 0) or 0), int(m.get("okQty", 0) or 0),
                 int(m.get("ngTotal", 0) or 0), float(m.get("ngRate", 0) or 0),
                 m.get("defectCategory", ""), m.get("defectType", ""),
                 int(m.get("defectCount", 0) or 0), now))

        tags_json = json.dumps(result.get("tags") or [], ensure_ascii=False)
        evidence_json = json.dumps(result.get("evidence") or [], ensure_ascii=False)
        actions_json = json.dumps(result.get("actions") or [], ensure_ascii=False)
        context_json = json.dumps(result.get("context"), ensure_ascii=False) if result.get("context") else ""
        doe_json = json.dumps(result.get("doeGrid"), ensure_ascii=False) if result.get("doeGrid") else ""
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
             result.get("summary", ""), result.get("keyFindings", ""),
             tags_json, now,
             result.get("purpose", ""), result.get("testConditions", ""),
             result.get("rootCause", ""), result.get("decision", ""),
             result.get("recommendedAction", ""),
             result.get("verdict", ""), result.get("headline", ""),
             evidence_json, actions_json, context_json,
             result.get("reportType", ""), doe_json, trend_json))

        for lang, tr in [("ko", tr_ko), ("vi", tr_vi)]:
            if tr is None:
                continue
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
                 tr.get("summary", ""), tr.get("keyFindings", ""),
                 tr.get("purpose", ""), tr.get("testConditions", ""),
                 tr.get("rootCause", ""), tr.get("decision", ""),
                 tr.get("recommendedAction", ""),
                 tr.get("headline", ""), tr_actions_json, tr_context_json, now))

        con.commit()
        return True
    except Exception as e:
        con.rollback()
        print(f"[FAIL {name}] {type(e).__name__}: {e}")
        return False


def commit_all():
    index = json.load(open(os.path.join(TSV_DIR, "index.json"), encoding="utf-8"))
    con = sqlite3.connect(DB)
    con.execute("PRAGMA busy_timeout = 30000")
    ok = 0
    fail = 0
    for path in sorted(glob.glob(os.path.join(RES_DIR, "*.json"))):
        idx = os.path.splitext(os.path.basename(path))[0]
        name = index.get(idx)
        if not name:
            print(f"[SKIP {idx}] not in index")
            continue
        payload = json.load(open(path, encoding="utf-8"))
        if _commit_one(con, name, payload):
            ok += 1
            print(f"[OK  {idx}] {name[:60]}")
        else:
            fail += 1
    con.close()
    print(f"\ncommitted ok={ok} fail={fail}")


def status():
    if not os.path.exists(os.path.join(TSV_DIR, "index.json")):
        print("no index yet — run extract")
        return
    index = json.load(open(os.path.join(TSV_DIR, "index.json"), encoding="utf-8"))
    total = len(index)
    done = len(glob.glob(os.path.join(RES_DIR, "*.json")))
    print(f"{done}/{total} results written")


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "status"
    if cmd == "extract":
        extract()
    elif cmd == "commit_all":
        commit_all()
    elif cmd == "status":
        status()
    else:
        print(__doc__)
        sys.exit(1)
