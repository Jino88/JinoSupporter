"""Dump each dataset's excel_paste to a separate file."""
import json, os
DATA = r'D:\000. MyWorks\005. Program\Repository\JinoSupporter\_chunk_05_data.json'
OUTDIR = r'D:\000. MyWorks\005. Program\Repository\JinoSupporter\_chunk_05_pastes'
os.makedirs(OUTDIR, exist_ok=True)
with open(DATA, 'r', encoding='utf-8') as f:
    data = json.load(f)
for i, (k, v) in enumerate(data.items()):
    fp = os.path.join(OUTDIR, f'{i:02d}.txt')
    with open(fp, 'w', encoding='utf-8') as f:
        f.write(f'### NAME: {k}\n\n')
        f.write(v or '')
    print(f'{i:02d} {fp} len={len(v) if v else 0}')
