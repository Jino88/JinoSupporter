"""Quick cell dump for a materialized text-only workbook.

Usage: python _dump_xlsx.py "<dataset name>" [--max-rows-per-sheet 250]
Calls _ai_batch_helper.get_excel_file to materialize, then prints sheet/row/cell
content with merged-cell ranges and inferred values carried across merges.
"""
from __future__ import annotations
import io, os, sys, zipfile, argparse, xml.etree.ElementTree as ET, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NSMAP = {'s': NS}


def col_idx(letter: str) -> int:
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch.upper()) - ord('A') + 1)
    return n


def col_letter(idx: int) -> str:
    s = ''
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(ord('A') + r) + s
    return s


def split_ref(ref: str) -> tuple[str, int]:
    m = re.match(r'([A-Z]+)(\d+)', ref)
    return m.group(1), int(m.group(2))


def load_shared_strings(z: zipfile.ZipFile) -> list[str]:
    if 'xl/sharedStrings.xml' not in z.namelist():
        return []
    r = ET.fromstring(z.read('xl/sharedStrings.xml'))
    out = []
    for si in r.findall(f'{{{NS}}}si'):
        out.append(''.join(t.text or '' for t in si.iter(f'{{{NS}}}t')))
    return out


def load_workbook_sheets(z: zipfile.ZipFile) -> list[tuple[str, str]]:
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rmap = {r.attrib['Id']: r.attrib['Target'] for r in rels}
    out = []
    for sh in wb.findall(f'{{{NS}}}sheets/{{{NS}}}sheet'):
        rid = sh.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        target = rmap.get(rid, '')
        if target.startswith('/'):
            target = target.lstrip('/')
        elif target and not target.startswith('xl/'):
            target = 'xl/' + target
        out.append((sh.attrib.get('name', ''), target))
    return out


def dump_sheet(name: str, sheet_path: str, z: zipfile.ZipFile, ss: list[str], max_rows: int) -> None:
    if sheet_path not in z.namelist():
        return
    try:
        root = ET.fromstring(z.read(sheet_path))
    except ET.ParseError as e:
        print(f'## SHEET {name} parse error: {e}')
        return
    merges = []
    mc = root.find(f'{{{NS}}}mergeCells')
    if mc is not None:
        for m in mc.findall(f'{{{NS}}}mergeCell'):
            merges.append(m.attrib.get('ref', ''))
    print(f'## SHEET {name} ({sheet_path})')
    if merges:
        print('MERGES:', '; '.join(merges))
    data = root.find(f'{{{NS}}}sheetData')
    rows = data.findall(f'{{{NS}}}row') if data is not None else []
    shown = 0
    for row in rows:
        rn = row.get('r')
        cells = []
        for c in row.findall(f'{{{NS}}}c'):
            ref = c.get('r')
            t = c.get('t')
            v = c.find(f'{{{NS}}}v')
            is_el = c.find(f'{{{NS}}}is')
            val = ''
            if t == 's' and v is not None and v.text is not None:
                try:
                    val = ss[int(v.text)]
                except Exception:
                    val = ''
            elif t == 'inlineStr' and is_el is not None:
                val = ''.join(x.text or '' for x in is_el.iter(f'{{{NS}}}t'))
            elif v is not None:
                val = v.text or ''
            val = (val or '').replace('\n', ' ').strip()
            if val:
                cells.append(f'{ref}={val}')
        if cells:
            print(f'R{rn}:', ' | '.join(cells))
            shown += 1
            if shown >= max_rows:
                print(f'... (truncated, more rows exist)')
                break


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('name')
    ap.add_argument('--max-rows-per-sheet', type=int, default=300)
    args = ap.parse_args()

    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import _ai_batch_helper as H

    con = H.open_db()
    try:
        path = H.get_excel_file(con, args.name)
    finally:
        con.close()

    if not path or not os.path.exists(path):
        print('NO_EXCEL_FILE')
        # try paste fallback
        con = H.open_db()
        try:
            paste = H.get_excel_paste(con, args.name)
        finally:
            con.close()
        if paste:
            print('## EXCEL_PASTE_FALLBACK')
            print(paste[:60000])
        return

    print(f'PATH={path}')
    with zipfile.ZipFile(path) as z:
        ss = load_shared_strings(z)
        sheets = load_workbook_sheets(z)
        for name, sp in sheets:
            dump_sheet(name, sp, z, ss, args.max_rows_per_sheet)


if __name__ == '__main__':
    main()
