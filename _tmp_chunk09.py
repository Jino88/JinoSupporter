import sqlite3, os, json, sys
DB = r'D:\000. MyWorks\002. DB\process-review.db'
ROOT = r'D:\000. MyWorks\005. Program\Repository\JinoSupporter'
names = [l for l in open(os.path.join(ROOT,'_batch_chunks','chunk_09.txt'),'r',encoding='utf-8').read().splitlines() if l.strip()]

def fetch_all():
    con = sqlite3.connect(DB)
    con.execute('PRAGMA busy_timeout=30000')
    cur = con.cursor()
    out = {}
    for n in names:
        r = cur.execute("SELECT ExtractedText FROM RawReportText WHERE DatasetName=? AND Kind='excel_paste'", (n,)).fetchone()
        # Also fetch existing ProductType
        s = cur.execute("SELECT ProductType FROM DatasetSummary WHERE DatasetName=?", (n,)).fetchone()
        out[n] = {'tsv': r[0] if r else None, 'productType': s[0] if s else ''}
    con.close()
    return out

def dump(idx):
    d = fetch_all()
    n = names[idx]
    print('NAME:', n)
    print('PRODUCT:', d[n]['productType'])
    print('==TSV==')
    print(d[n]['tsv'])

if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == 'len':
        d = fetch_all()
        for n in names:
            t = d[n]['tsv']
            print(f"{len(t) if t else 0}  P={d[n]['productType']}  {n}")
    else:
        dump(int(sys.argv[1]))
