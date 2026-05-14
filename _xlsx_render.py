"""Render a text-only XLSX to a compact text grid for AI ingestion.

Output format:
=== SHEET: <name> ===
DIMENSION: A1:BG40
MERGED: A2:S2, B3:D3, ...
[row 2]
  B2: TITLE | D2: TIU L5S3-01 [R] ... | U2: Dept | V2: ME
[row 3]
  ...

Only non-empty cells appear. Merged-cell ranges are explicit so the model
can carry values down/right. Numeric cells with Excel date styles are rendered
as ISO dates instead of raw Excel serial numbers.
"""
from __future__ import annotations
import sys, io, zipfile, datetime as dt
import xml.etree.ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')


def col_to_idx(col: str) -> int:
    n = 0
    for c in col:
        n = n * 26 + (ord(c.upper()) - ord('A') + 1)
    return n


def ref_split(ref: str) -> tuple[str, int]:
    i = 0
    while i < len(ref) and ref[i].isalpha():
        i += 1
    return ref[:i], int(ref[i:])


def excel_serial_to_iso(v: float) -> str | None:
    try:
        if v < 20000 or v > 80000:
            return None
        base = dt.datetime(1899, 12, 30)
        d = base + dt.timedelta(days=float(v))
        return d.strftime('%Y-%m-%d')
    except Exception:
        return None


def parse_shared_strings(z: zipfile.ZipFile) -> list[str]:
    if 'xl/sharedStrings.xml' not in z.namelist():
        return []
    root = ET.fromstring(z.read('xl/sharedStrings.xml'))
    out = []
    for si in root.findall(NS + 'si'):
        # Could have <t> directly or under <r>
        text = ''.join(t.text or '' for t in si.iter(NS + 't'))
        out.append(text)
    return out


def parse_workbook_sheets(z: zipfile.ZipFile) -> list[tuple[str, str]]:
    """Return workbook sheets in UI order with their actual worksheet XML paths."""
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rid_to_target = {}
    for rel in rels:
        rid = rel.attrib.get('Id')
        target = rel.attrib.get('Target', '')
        if rid:
            target = target.replace('\\', '/')
            if target.startswith('/'):
                target = target.lstrip('/')
            elif target and not target.startswith('xl/'):
                target = 'xl/' + target
            rid_to_target[rid] = target.replace('xl/xl/', 'xl/')

    sheets: list[tuple[str, str]] = []
    for s in wb.iter(NS + 'sheet'):
        rid = s.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id') \
            or s.attrib.get('r:id')
        target = rid_to_target.get(rid or '', '')
        if target:
            sheets.append((s.attrib.get('name', ''), target))
    return sheets


def parse_date_style_indexes(z: zipfile.ZipFile) -> set[int]:
    if 'xl/styles.xml' not in z.namelist():
        return set()
    try:
        root = ET.fromstring(z.read('xl/styles.xml'))
    except Exception:
        return set()

    date_num_fmt_ids = {
        14, 15, 16, 17, 18, 19, 20, 21, 22,
        27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
        45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58,
    }
    custom_date_num_fmt_ids: set[int] = set()
    num_fmts = root.find(NS + 'numFmts')
    if num_fmts is not None:
        for num_fmt in num_fmts.findall(NS + 'numFmt'):
            try:
                fmt_id = int(num_fmt.attrib.get('numFmtId', '0'))
            except Exception:
                continue
            code = num_fmt.attrib.get('formatCode', '').lower()
            # Remove quoted literals and escaped chars before looking for date tokens.
            cleaned = ''
            in_quote = False
            i = 0
            while i < len(code):
                ch = code[i]
                if ch == '"':
                    in_quote = not in_quote
                elif ch == '\\':
                    i += 1
                elif not in_quote:
                    cleaned += ch
                i += 1
            if any(token in cleaned for token in ('yy', 'yyyy', 'dd', 'd/', '/d', 'mm-dd', 'm/d', 'h:mm')):
                custom_date_num_fmt_ids.add(fmt_id)

    date_styles: set[int] = set()
    cell_xfs = root.find(NS + 'cellXfs')
    if cell_xfs is None:
        return date_styles
    for idx, xf in enumerate(cell_xfs.findall(NS + 'xf')):
        try:
            num_fmt_id = int(xf.attrib.get('numFmtId', '0'))
        except Exception:
            continue
        if num_fmt_id in date_num_fmt_ids or num_fmt_id in custom_date_num_fmt_ids:
            date_styles.add(idx)
    return date_styles


def render_sheet(xml_data: bytes, sheet_name: str, shared: list[str], date_styles: set[int]) -> str:
    root = ET.fromstring(xml_data)
    dim = root.find(NS + 'dimension')
    dim_ref = dim.attrib.get('ref', '') if dim is not None else ''

    merges = []
    mc = root.find(NS + 'mergeCells')
    if mc is not None:
        for m in mc.findall(NS + 'mergeCell'):
            merges.append(m.attrib.get('ref', ''))

    rows_out: dict[int, list[tuple[int, str, str]]] = {}

    sd = root.find(NS + 'sheetData')
    if sd is None:
        return f'=== SHEET: {sheet_name} ===\nDIMENSION: {dim_ref}\nMERGED: {", ".join(merges)}\n(empty)\n'

    for row in sd.findall(NS + 'row'):
        r_num = int(row.attrib.get('r', '0'))
        for c in row.findall(NS + 'c'):
            ref = c.attrib.get('r', '')
            t = c.attrib.get('t', 'n')
            val = ''
            if t == 's':
                vnode = c.find(NS + 'v')
                if vnode is not None and vnode.text is not None:
                    idx = int(vnode.text)
                    if 0 <= idx < len(shared):
                        val = shared[idx]
            elif t == 'inlineStr':
                isnode = c.find(NS + 'is')
                if isnode is not None:
                    val = ''.join(t2.text or '' for t2 in isnode.iter(NS + 't'))
            elif t == 'str':
                vnode = c.find(NS + 'v')
                val = vnode.text if vnode is not None and vnode.text else ''
            elif t == 'b':
                vnode = c.find(NS + 'v')
                val = 'TRUE' if (vnode is not None and vnode.text == '1') else 'FALSE'
            else:
                # numeric (or empty)
                vnode = c.find(NS + 'v')
                if vnode is not None and vnode.text is not None:
                    val = vnode.text
                    try:
                        fv = float(val)
                        style_idx = int(c.attrib.get('s', '-1'))
                        iso = excel_serial_to_iso(fv) if style_idx in date_styles else None
                        if iso:
                            val = iso
                    except Exception:
                        pass
            if val is None:
                val = ''
            val = val.strip()
            if not val:
                continue
            col, _ = ref_split(ref)
            rows_out.setdefault(r_num, []).append((col_to_idx(col), ref, val))

    lines = [f'=== SHEET: {sheet_name} ===',
             f'DIMENSION: {dim_ref}',
             f'MERGED: {", ".join(merges) if merges else "(none)"}']
    for r in sorted(rows_out.keys()):
        cells = sorted(rows_out[r], key=lambda x: x[0])
        line = f'[r{r}] ' + ' | '.join(f'{ref}={val}' for _, ref, val in cells)
        lines.append(line)
    return '\n'.join(lines)


def render_workbook(path: str) -> str:
    with zipfile.ZipFile(path) as z:
        shared = parse_shared_strings(z)
        date_styles = parse_date_style_indexes(z)
        out = []
        for name, sxml in parse_workbook_sheets(z):
            if sxml in z.namelist():
                out.append(render_sheet(z.read(sxml), name, shared, date_styles))
        return '\n\n'.join(out)


if __name__ == '__main__':
    print(render_workbook(sys.argv[1]))
