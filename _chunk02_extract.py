import _ai_batch_helper as h
import json, os

names = [l.rstrip() for l in open(r'D:\000. MyWorks\005. Program\Repository\JinoSupporter\_chunk_02.txt', 'r', encoding='utf-8-sig') if l.strip()]
con = h.open_db()
out = {}
for n in names:
    txt = h.get_excel_paste(con, n)
    out[n] = txt or ''
    print(f'=== {n} ===')
    print(f'len={len(txt) if txt else 0}')
con.close()

# Write each to a file for inspection
for i, n in enumerate(names, 1):
    fn = rf'D:\000. MyWorks\005. Program\Repository\JinoSupporter\_chunk02_ds{i:02d}.txt'
    with open(fn, 'w', encoding='utf-8') as f:
        f.write(f'# {n}\n')
        f.write(out[n])
print('done')
