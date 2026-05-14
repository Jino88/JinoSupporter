"""View one dataset's excel_paste at a time."""
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DATA = r'D:\000. MyWorks\005. Program\Repository\JinoSupporter\_chunk_05_data.json'
with open(DATA, 'r', encoding='utf-8') as f:
    data = json.load(f)

idx = int(sys.argv[1])
keys = list(data.keys())
k = keys[idx]
print(f'=== [{idx}] {k} ===')
print(data[k])
