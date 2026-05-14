"""Fetch excel_paste for chunk 05 datasets."""
import sys, io, sqlite3, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DB = r'D:\000. MyWorks\002. DB\process-review.db'
CHUNK = r'D:\000. MyWorks\005. Program\Repository\JinoSupporter\_chunk_05.txt'
OUT = r'D:\000. MyWorks\005. Program\Repository\JinoSupporter\_chunk_05_data.json'

with open(CHUNK, 'r', encoding='utf-8') as f:
    targets = [l.rstrip('\r\n') for l in f if l.strip()]

con = sqlite3.connect(DB)
out = {}
for t in targets:
    row = con.execute(
        "SELECT ExtractedText FROM RawReportText WHERE DatasetName=? AND Kind='excel_paste'",
        (t,)
    ).fetchone()
    out[t] = row[0] if row and row[0] else None
    print(f'{t}: {"OK len="+str(len(row[0])) if row and row[0] else "MISSING"}')

with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=2)
con.close()
print(f'\nWrote {OUT}')
