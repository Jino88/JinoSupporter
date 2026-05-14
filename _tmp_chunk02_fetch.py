import sqlite3, os
DB = r'D:\000. MyWorks\002. DB\process-review.db'
CHUNK = r'D:\000. MyWorks\005. Program\Repository\JinoSupporter\_batch_chunks\chunk_02.txt'
OUT = r'D:\000. MyWorks\005. Program\Repository\JinoSupporter\_tmp_pastes'
os.makedirs(OUT, exist_ok=True)
with open(CHUNK, 'r', encoding='utf-8') as f:
    names = [l.strip() for l in f if l.strip()]
con = sqlite3.connect(DB)
con.execute('PRAGMA busy_timeout = 30000')
cur = con.cursor()
for i, n in enumerate(names):
    cur.execute("SELECT ExtractedText FROM RawReportText WHERE DatasetName=? AND Kind='excel_paste'", (n,))
    r = cur.fetchone()
    txt = r[0] if r else None
    fp = os.path.join(OUT, f'{i:02d}.txt')
    with open(fp, 'w', encoding='utf-8') as g:
        g.write(txt if txt else '')
    print(f'{i:02d} {len(txt) if txt else 0} {n}')
con.close()
