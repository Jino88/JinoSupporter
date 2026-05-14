"""Helper: dump each dataset's excel_paste TSV to per-dataset files for inspection."""
import sqlite3, os, sys, json

DB = r"D:\000. MyWorks\002. DB\process-review.db"
CHUNK = r"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_batch_chunks\chunk_03.txt"
OUT_DIR = r"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_batch_chunks\_chunk03_tsv"

os.makedirs(OUT_DIR, exist_ok=True)

with open(CHUNK, "r", encoding="utf-8") as f:
    names = [ln.strip() for ln in f if ln.strip()]

con = sqlite3.connect(DB)
con.execute("PRAGMA busy_timeout = 30000")
cur = con.cursor()

results = []
for i, name in enumerate(names):
    row = cur.execute(
        "SELECT ExtractedText FROM RawReportText WHERE DatasetName=? AND Kind='excel_paste'",
        (name,)).fetchone()
    if not row or not row[0]:
        results.append((i, name, 0, "SKIP"))
        continue
    txt = row[0]
    # save to per-dataset file
    safe = f"{i:02d}.txt"
    with open(os.path.join(OUT_DIR, safe), "w", encoding="utf-8") as g:
        g.write(name + "\n" + ("=" * 80) + "\n")
        g.write(txt)
    results.append((i, name, len(txt), "OK"))

con.close()
for r in results:
    print(r)
