"""Split the pastes JSON into individual text files for easy reading."""
import json
import os

OUT_DIR = r"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_batch_chunks\pastes_01"
SRC = r"D:\000. MyWorks\005. Program\Repository\JinoSupporter\_batch_chunks\chunk_01_pastes.json"

os.makedirs(OUT_DIR, exist_ok=True)
with open(SRC, "r", encoding="utf-8") as f:
    data = json.load(f)

for i, d in enumerate(data, 1):
    if not d["paste"]:
        continue
    fn = os.path.join(OUT_DIR, f"ds_{i:02d}.txt")
    with open(fn, "w", encoding="utf-8") as g:
        g.write("=== " + d["name"] + " ===\n\n")
        g.write(d["paste"])
    print(f"wrote {fn}")
