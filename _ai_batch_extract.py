"""Extract a text-only XLSX workbook into a structured view for AI analysis.

Preserves: sheet names, cell coordinates, merged ranges (and propagates merged
values into every covered cell), formula presence, and basic styles. Output is
designed to be compact enough to fit in AI prompts while keeping enough
structure that merged Date/Model/Line carry-forward rules can be applied.
The row text cap is per sheet, so later worksheets are not dropped just because
earlier sheets are large.
"""
from __future__ import annotations
import zipfile, io, json, sys, re
import xml.etree.ElementTree as ET

NS = {
    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}


def _local(tag):
    return tag.rsplit('}', 1)[-1] if '}' in tag else tag


def col_letter(idx):
    s = ''
    n = idx + 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def coord_to_rc(coord):
    m = re.match(r'^([A-Z]+)(\d+)$', coord)
    if not m:
        return None
    letters, row = m.group(1), int(m.group(2))
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - 64)
    return row, col


def expand_range(rng):
    a, b = rng.split(':')
    r1, c1 = coord_to_rc(a)
    r2, c2 = coord_to_rc(b)
    rows = []
    for r in range(min(r1, r2), max(r1, r2) + 1):
        for c in range(min(c1, c2), max(c1, c2) + 1):
            rows.append((r, c))
    return rows, (min(r1, r2), min(c1, c2))


def extract_workbook(path, max_chars=120000):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        # 1. shared strings
        shared = []
        if 'xl/sharedStrings.xml' in names:
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root:
                # collect text from t children (handles rich text runs)
                parts = []
                for el in si.iter():
                    if _local(el.tag) == 't' and el.text:
                        parts.append(el.text)
                shared.append(''.join(parts))

        # 2. workbook sheet order + relationship map
        wb_root = ET.fromstring(z.read('xl/workbook.xml'))
        rels_root = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rid_to_target = {}
        for rel in rels_root:
            rid_to_target[rel.attrib['Id']] = rel.attrib['Target']
        sheets = []  # (name, target_path)
        for sh in wb_root.iter():
            if _local(sh.tag) == 'sheet':
                rid = sh.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id') \
                      or sh.attrib.get('r:id')
                tgt = rid_to_target.get(rid, '')
                if not tgt:
                    continue
                tgt = tgt.replace('\\', '/').lstrip('/')
                if tgt.startswith('xl/'):
                    target = tgt
                else:
                    target = 'xl/' + tgt
                target = target.replace('xl/xl/', 'xl/')
                sheets.append((sh.attrib.get('name', ''), target))

    out = {'sheets': []}
    truncated_sheets = []
    with zipfile.ZipFile(path) as z:
        for sheet_name, target in sheets:
            if target not in z.namelist():
                # try fallback
                alt = target.replace('xl/', '')
                if 'xl/' + alt in z.namelist():
                    target = 'xl/' + alt
                else:
                    continue
            sx = z.read(target)
            sroot = ET.fromstring(sx)
            merged_ranges = []
            for el in sroot.iter():
                if _local(el.tag) == 'mergeCell':
                    ref = el.attrib.get('ref', '')
                    if ':' in ref:
                        merged_ranges.append(ref)
            # build cell map
            cells = {}
            for c in sroot.iter():
                if _local(c.tag) != 'c':
                    continue
                ref = c.attrib.get('r')
                if not ref:
                    continue
                t = c.attrib.get('t', 'n')
                val = None
                for ch in c:
                    if _local(ch.tag) == 'v':
                        val = ch.text
                    elif _local(ch.tag) == 'is':
                        parts = []
                        for el in ch.iter():
                            if _local(el.tag) == 't' and el.text:
                                parts.append(el.text)
                        val = ''.join(parts)
                if val is None:
                    continue
                if t == 's':
                    try:
                        val = shared[int(val)]
                    except (ValueError, IndexError):
                        pass
                elif t == 'str' or t == 'inlineStr':
                    pass
                cells[ref] = val
            # propagate merged values
            for rng in merged_ranges:
                coords, anchor = expand_range(rng)
                ar, ac = anchor
                anchor_ref = col_letter(ac - 1) + str(ar)
                anchor_val = cells.get(anchor_ref)
                if anchor_val is None:
                    continue
                for (r, c) in coords:
                    ref = col_letter(c - 1) + str(r)
                    if ref not in cells:
                        cells[ref] = anchor_val
            # bucket rows
            rows = {}
            for ref, v in cells.items():
                rc = coord_to_rc(ref)
                if not rc:
                    continue
                r, c = rc
                rows.setdefault(r, {})[c] = (ref, v)
            row_lines = []
            for r in sorted(rows):
                cols = rows[r]
                if not cols:
                    continue
                max_c = max(cols)
                cells_parts = []
                for c in range(1, max_c + 1):
                    if c in cols:
                        ref, v = cols[c]
                        sv = str(v)
                        if len(sv) > 200:
                            sv = sv[:200] + '...'
                        # strip newlines/tabs in cell for line readability
                        sv = sv.replace('\t', ' ').replace('\r', ' ').replace('\n', ' ')
                        cells_parts.append(f'{ref}={sv}')
                if not cells_parts:
                    continue
                row_lines.append('\t'.join(cells_parts))
            sheet_text = '\n'.join(row_lines)
            sheet_obj = {
                'name': sheet_name,
                'merged_ranges': merged_ranges[:200],
                'text': sheet_text,
            }
            if max_chars > 0 and len(sheet_text) > max_chars:
                sheet_obj['text'] = sheet_text[:max_chars] + '\n...[TRUNCATED_SHEET]'
                truncated_sheets.append(sheet_name)
            out['sheets'].append(sheet_obj)
    out['truncated'] = len(truncated_sheets) > 0
    out['truncated_sheets'] = truncated_sheets
    return out


if __name__ == '__main__':
    path = sys.argv[1]
    max_chars = int(sys.argv[2]) if len(sys.argv) > 2 else 120000
    data = extract_workbook(path, max_chars=max_chars)
    print(json.dumps(data, ensure_ascii=False))
