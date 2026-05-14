"""Fetch all paste texts to _tmp_batch_3/pastes/NNN__name.txt and emit a manifest."""
import os, sqlite3, json, re
DB = r'D:\000. MyWorks\002. DB\process-review.db'
ROOT = r'D:\000. MyWorks\005. Program\Repository\JinoSupporter'
TARGETS = os.path.join(ROOT, '_batch_selected_20260512_163659_955aeca_3.txt')
OUTDIR = os.path.join(ROOT, '_tmp_batch_3', 'pastes')
os.makedirs(OUTDIR, exist_ok=True)

with open(TARGETS, encoding='utf-8-sig') as f:
    names = [ln.strip() for ln in f if ln.strip()]

con = sqlite3.connect(DB)
con.execute("PRAGMA busy_timeout = 30000")
cur = con.cursor()
manifest = []
for i, name in enumerate(names, 1):
    row = cur.execute(
        "SELECT ExtractedText FROM RawReportText WHERE DatasetName=? AND Kind='excel_paste'",
        (name,)).fetchone()
    if not row or not row[0]:
        manifest.append({"i": i, "name": name, "file": None, "len": 0})
        continue
    safe = re.sub(r'[^a-zA-Z0-9._-]+', '_', name)[:80]
    fn = f"{i:03d}__{safe}.txt"
    path = os.path.join(OUTDIR, fn)
    with open(path, 'w', encoding='utf-8') as g:
        g.write(row[0])
    manifest.append({"i": i, "name": name, "file": fn, "len": len(row[0])})
con.close()

with open(os.path.join(ROOT, '_tmp_batch_3', 'manifest.json'), 'w', encoding='utf-8') as f:
    json.dump(manifest, f, ensure_ascii=False, indent=1)
print(f"wrote {len([m for m in manifest if m['file']])} paste files; missing {len([m for m in manifest if not m['file']])}")
