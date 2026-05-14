# -*- coding: utf-8 -*-
"""Drive auto_normalize.normalize_and_commit across batch file 1."""
import sys, io, os, time

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
sys.path.insert(0, os.path.dirname(__file__))

from auto_normalize import normalize_and_commit
from commit_lib import fetch_paste

ROOT = r'D:\000. MyWorks\005. Program\Repository\JinoSupporter'
TARGETS = os.path.join(ROOT, '_batch_selected_20260513_121636_c418cfc_1.txt')

with open(TARGETS, encoding='utf-8-sig') as f:
    names = [l.strip().lstrip('﻿') for l in f if l.strip()]

skipped = []
ok = 0
errs = []
t0 = time.time()
for i, n in enumerate(names, 1):
    try:
        if not fetch_paste(n):
            skipped.append(n)
            continue
        normalize_and_commit(n)
        ok += 1
        if i % 25 == 0:
            print(f'[{i}/{len(names)}] ok={ok} skipped={len(skipped)} err={len(errs)} elapsed={time.time()-t0:.1f}s', flush=True)
    except Exception as e:
        errs.append((n, repr(e)))
        print(f'  ! {n[:60]}: {e!r}', flush=True)

print(f'DONE ok={ok} skipped={len(skipped)} err={len(errs)} elapsed={time.time()-t0:.1f}s')
print('SKIPPED (no excel_paste):')
for n in skipped:
    print('  -', n)
print('ERRORS:')
for n, e in errs:
    print('  -', n, '->', e)
