"""Fetch ExtractedText for each dataset in chunk_01.txt."""
import sqlite3
import json
import os

DB = r"D:\000. MyWorks\002. DB\process-review.db"
CHUNK = r"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_batch_chunks\chunk_01.txt"
OUT = r"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_batch_chunks\chunk_01_pastes.json"

with open(CHUNK, "r", encoding="utf-8-sig") as f:
    names = [ln.strip() for ln in f if ln.strip()]

con = sqlite3.connect(DB)
con.execute("PRAGMA busy_timeout = 30000")
cur = con.cursor()

data = []
for n in names:
    row = cur.execute(
        "SELECT ExtractedText FROM RawReportText WHERE DatasetName=? AND Kind='excel_paste'",
        (n,)
    ).fetchone()
    if not row or not row[0]:
        data.append({"name": n, "paste": None})
    else:
        data.append({"name": n, "paste": row[0]})

con.close()

with open(OUT, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False)

# Brief summary
for d in data:
    if d["paste"]:
        print(f"OK  {d['name']}: {len(d['paste'])} chars")
    else:
        print(f"SKIP {d['name']}: no excel_paste")
