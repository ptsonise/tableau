"""
update_customer_tracker.py - refresh the Customer Daily Production Loss Tracker from the 3 customer files.

Inputs (same folder as this script):
    Sabanci.xlsx, Libra.xlsx, Crimson.xlsx      sheets 'Daily Prodction' + 'Sales' (cached values are read)
    Customer_Daily_Production_Loss_Tracker*.xlsx  template = the NEWEST one (so remarks you typed in
                                                  yesterday's output are carried forward)
Output:
    Customer_Daily_Production_Loss_Tracker_<YYYY-MM-DD>.xlsx

What it does = what 'Data > Refresh All' does in Excel (same 6 Power Query tables, same source sheets):
    Sabanci 'Daily Prodction' -> table Sabanci_Daily_Production  (sheet 'Daily Production_Sabanci')
    Sabanci 'Sales'           -> table Sabanci_Sales             (sheet 'Sales_Sabanci')
    Libra   'Daily Prodction' -> table Libra_Daily_Production    (sheet 'Daily Production_SB')
    Libra   'Sales'           -> table Libra_Sale                (sheet 'Sales_SB')
    Crimson 'Daily Prodction' -> table CS_Daily_Production       (sheet 'Daily Production_CS')
    Crimson 'Sales'           -> table CS_Sales                  (sheet 'Sales_CS')
Only the <sheetData> of these 6 sheets + their table / filter / name ranges are rewritten (XML level), so every
other sheet, formula, colour, conditional format and the Power Query definitions stay exactly as they are.
Excel recalculates all formulas when the file opens (fullCalcOnLoad).

Run:  python update_customer_tracker.py            (Python 3.9+, openpyxl)
"""
import datetime as dt
import glob
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

import openpyxl
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
TRACKER_STEM = 'Customer_Daily_Production_Loss_Tracker'

# Power Query "Changed Type" per source table (columns not listed = 'any', kept as read)
SALES_TYPES = {
    'Plant': 'int', 'Posting Date': 'date', 'Entry Date': 'date', 'Year': 'int', 'FI Docs': 'int', 'SO No.': 'int',
    'Qty.': 'int', 'Qty.Wp': 'int', 'Billing Docs.': 'int', 'Line No.': 'int', 'Delivery No.': 'int',
    'Amt. In Lc': 'num', 'Line No._1': 'int', 'Amt in Doc.': 'num',
    'Type': 'text', 'Ship to Party Name': 'text', 'Name': 'text', 'Material': 'text', 'Billing Type': 'text',
    'Billing Desc.': 'text', 'Desc.': 'text', 'Sales Group': 'text', 'Sales Group Descr.': 'text',
    'Order Reason': 'text', 'Order Reason Desc.': 'text', 'Material Group Desc': 'text', 'Module': 'text',
    'JW or USA': 'text'}

# (source file, source sheet, tracker table name, column types)
JOBS = [
    ('Sabanci.xlsx', 'Daily Prodction', 'Sabanci_Daily_Production', {}),
    ('Sabanci.xlsx', 'Sales', 'Sabanci_Sales', SALES_TYPES),
    ('Libra.xlsx', 'Daily Prodction', 'Libra_Daily_Production', {'Column13': 'text', 'Column14': 'text'}),
    ('Libra.xlsx', 'Sales', 'Libra_Sale', SALES_TYPES),
    ('Crimson.xlsx', 'Daily Prodction', 'CS_Daily_Production', {'Column12': 'int', 'Column18': 'int'}),
    ('Crimson.xlsx', 'Sales', 'CS_Sales', SALES_TYPES),
]

EXCEL_EPOCH = dt.datetime(1899, 12, 30)


# ---------------------------------------------------------------- reading the customer files
def find_file(name):
    """Exact name first, else any file that ends with the name (e.g. an uploaded '4d55-Sabanci.xlsx')."""
    p = os.path.join(HERE, name)
    if os.path.exists(p):
        return p
    hits = [f for f in glob.glob(os.path.join(HERE, '*' + name)) if not os.path.basename(f).startswith('~$')]
    return max(hits, key=os.path.getmtime) if hits else None


def pq_headers(row):
    """Power Query 'Promote headers': blank -> ColumnN, duplicates -> name_1, name_2 ..."""
    out, seen = [], set()
    for i, v in enumerate(row, 1):
        h = 'Column%d' % i if v is None or str(v) == '' else str(v)
        base, k = h, 1
        while h in seen:
            h = '%s_%d' % (base, k)
            k += 1
        seen.add(h)
        out.append(h)
    return out


def to_type(v, kind):
    """Mirror the Power Query column type. A value PQ cannot convert becomes blank (PQ shows Error)."""
    if v is None or kind is None:
        return v
    if kind == 'date':
        if isinstance(v, dt.datetime):
            return v.replace(hour=0, minute=0, second=0, microsecond=0)
        if isinstance(v, dt.date):
            return dt.datetime(v.year, v.month, v.day)
        if isinstance(v, (int, float)):
            return EXCEL_EPOCH + dt.timedelta(days=int(v))
        return None
    if kind == 'int':
        if isinstance(v, bool):
            return int(v)
        if isinstance(v, (int, float)):
            return int(round(v))
        try:
            return int(round(float(str(v).replace(',', ''))))
        except ValueError:
            return None
    if kind == 'num':
        if isinstance(v, (int, float)):
            return v
        try:
            return float(str(v).replace(',', ''))
        except ValueError:
            return None
    if kind == 'text':
        if isinstance(v, dt.datetime):
            return v.strftime('%m/%d/%Y')
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)
    return v


def read_source(path, sheet, types):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if sheet not in wb.sheetnames:
        wb.close()
        raise KeyError("sheet '%s' not found in %s" % (sheet, os.path.basename(path)))
    rows = [list(r) for r in wb[sheet].iter_rows(values_only=True)]
    wb.close()
    while rows and all(v is None for v in rows[-1]):     # PQ drops trailing empty rows
        rows.pop()
    if not rows:
        raise ValueError("sheet '%s' in %s is empty" % (sheet, os.path.basename(path)))
    width = max(len(r) for r in rows)
    rows = [(r + [None] * width)[:width] for r in rows]
    hdr = pq_headers(rows[0])
    kinds = [types.get(h) for h in hdr]
    data = [[to_type(v, k) for v, k in zip(r, kinds)] for r in rows[1:]]
    formulas = sum(1 for r in data for v in r if isinstance(v, str) and v.startswith('='))
    return hdr, data, formulas


# ---------------------------------------------------------------- writing the tracker (XML level)
def norm(p):
    """Resolve a relationship target to a zip path."""
    parts = []
    for seg in p.split('/'):
        if seg == '..':
            parts.pop()
        elif seg and seg != '.':
            parts.append(seg)
    return '/'.join(parts)


def sheet_map(z):
    """sheet name -> (sheet xml path, list of table xml paths)."""
    wb = z.read('xl/workbook.xml').decode('utf-8')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid = {m.group(1): m.group(2) for m in
           re.finditer(r'<Relationship [^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"', rels)}
    rid.update({m.group(2): m.group(1) for m in
                re.finditer(r'<Relationship [^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"', rels)})
    out = {}
    for m in re.finditer(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', wb):
        name = m.group(1).replace('&gt;', '>').replace('&lt;', '<').replace('&amp;', '&').replace('&apos;', "'")
        tgt = rid[m.group(2)]
        path = tgt.lstrip('/') if tgt.startswith('/') else norm('xl/' + tgt)
        rp = path.replace('worksheets/', 'worksheets/_rels/') + '.rels'
        tabs = []
        if rp in z.namelist():
            for t in re.findall(r'Target="([^"]*tables/[^"]+)"', z.read(rp).decode('utf-8')):
                tabs.append(norm('xl/worksheets/' + t))
        out[name] = (path, tabs)
    return out


def cell_xml(ref, v, style):
    s = ' s="%d"' % style if style is not None else ''
    if v is None:
        return '<c r="%s"%s/>' % (ref, s) if style is not None else ''
    if isinstance(v, bool):
        return '<c r="%s"%s t="b"><v>%d</v></c>' % (ref, s, int(v))
    if isinstance(v, dt.datetime):
        d = v - EXCEL_EPOCH
        return '<c r="%s"%s><v>%s</v></c>' % (ref, s, repr(d.days + d.seconds / 86400.0).rstrip('0').rstrip('.'))
    if isinstance(v, (int, float)):
        return '<c r="%s"%s><v>%s</v></c>' % (ref, s, repr(v))
    t = escape(str(v))
    sp = ' xml:space="preserve"' if t != t.strip() else ''
    return '<c r="%s"%s t="inlineStr"><is><t%s>%s</t></is></c>' % (ref, s, sp, t)


def cell_styles(sheet_xml, ncol):
    """Style of every existing cell (row, col) of the tracker sheet + per column: the last data-row style,
    else the column default from <cols>. New rows take the column style, so dates keep their date format."""
    st = {}
    for rm in re.finditer(r'<row [^>]*?r="(\d+)"[^>]*>(.*?)</row>', sheet_xml, re.S):
        r = int(rm.group(1))
        for c in re.finditer(r'<c r="([A-Z]+)\d+"([^>]*?)(/>|>)', rm.group(2)):
            sm = re.search(r' s="(\d+)"', c.group(2))
            if sm:
                st[(r, openpyxl.utils.column_index_from_string(c.group(1)))] = int(sm.group(1))
    col = {}
    for cm in re.finditer(r'<col [^>]*?min="(\d+)"[^>]*?max="(\d+)"[^>]*?/>', sheet_xml):
        sm = re.search(r' style="(\d+)"', cm.group(0))
        if sm:
            for j in range(int(cm.group(1)), int(cm.group(2)) + 1):
                col[j] = int(sm.group(1))
    for j in range(1, ncol + 1):
        rows = sorted(r for (r, c) in st if c == j and r > 1)
        if rows:
            col[j] = st[(rows[-1], j)]
    return st, col


def build_sheet_data(hdr, data, styles):
    st, col = styles
    out = ['<sheetData>']
    rows = [hdr] + data
    for i, row in enumerate(rows, 1):
        cells = ''.join(cell_xml('%s%d' % (get_column_letter(j), i), v, st.get((i, j), col.get(j) if i > 1 else None))
                        for j, v in enumerate(row, 1))
        out.append('<row r="%d" spans="1:%d">%s</row>' % (i, len(hdr), cells))
    out.append('</sheetData>')
    return ''.join(out)


def esc_attr(s):
    return escape(s, {'"': '&quot;'})


def refresh(template, out_path, jobs_data):
    zin = zipfile.ZipFile(template)
    names = zin.namelist()
    smap = sheet_map(zin)
    # table name -> (table xml path, sheet name)
    tables = {}
    for sh, (_, tabs) in smap.items():
        for tp in tabs:
            m = re.search(r'<table [^>]*?\bname="([^"]+)"', zin.read(tp).decode('utf-8'))
            if m:
                tables[m.group(1)] = (tp, sh)
    new_parts = {}
    wb_xml = zin.read('xl/workbook.xml').decode('utf-8')
    report = []
    for tname, (hdr, data, src) in jobs_data.items():
        if tname not in tables:
            report.append('WARNING table %s not in the tracker - skipped' % tname)
            continue
        tp, sh = tables[tname]
        sp = smap[sh][0]
        t_xml = zin.read(tp).decode('utf-8')
        old_cols = [c.replace('_x000a_', '\n') for c in re.findall(r'<tableColumn [^>]*?name="([^"]+)"', t_xml)]
        old_ref = re.search(r'<table [^>]*?\bref="([^"]+)"', t_xml).group(1)
        if len(hdr) < len(old_cols):                  # source narrower: keep the table's extra (empty) columns
            hdr = hdr + old_cols[len(hdr):]
            data = [r + [None] * (len(old_cols) - len(r)) for r in data]
        if len(old_cols) != len(hdr):
            report.append('WARNING %s: source has %d columns, tracker table has %d - NOT refreshed (kept old data)'
                          % (sh, len(hdr), len(old_cols)))
            continue
        renamed = [(o, n) for o, n in zip(old_cols, hdr) if o != n]
        ncol, nrow = len(hdr), len(data) + 1
        if nrow < 2:                                  # an Excel table needs one data row
            data, nrow = [[None] * ncol], 2
        ref = 'A1:%s%d' % (get_column_letter(ncol), nrow)
        s_xml = zin.read(sp).decode('utf-8')
        styles = cell_styles(s_xml, ncol)
        s_xml = re.sub(r'<sheetData>.*?</sheetData>|<sheetData/>', lambda _: build_sheet_data(hdr, data, styles),
                       s_xml, count=1, flags=re.S)
        s_xml = re.sub(r'<dimension ref="[^"]*"/>', '<dimension ref="%s"/>' % ref, s_xml, count=1)
        new_parts[sp] = s_xml
        # table range, filter range, column names (Power Query names = header text)
        t_xml = re.sub(r'(<table [^>]*?\bref=")[^"]+"', lambda m: m.group(1) + ref + '"', t_xml, count=1)
        t_xml = re.sub(r'(<autoFilter [^>]*?\bref=")[^"]+"', lambda m: m.group(1) + ref + '"', t_xml, count=1)
        if renamed:
            it = iter(hdr)
            t_xml = re.sub(r'(<tableColumn [^>]*?name=")[^"]+"',
                           lambda m: m.group(1) + esc_attr(next(it)).replace('\n', '_x000a_') + '"', t_xml)
            report.append('NOTE %s: header changed in the source %s -> table columns renamed: %s'
                          % (sh, src, '; '.join('%r->%r' % rn for rn in renamed)))
        new_parts[tp] = t_xml
        # query table field names follow the header too
        qrels = tp.replace('tables/', 'tables/_rels/') + '.rels'
        if qrels in names and renamed:
            for q in re.findall(r'Target="([^"]*queryTables/[^"]+)"', zin.read(qrels).decode('utf-8')):
                qp = norm('xl/tables/' + q)
                q_xml = zin.read(qp).decode('utf-8')
                it = iter(hdr)
                q_xml = re.sub(r'(<queryTableField [^>]*?name=")[^"]+"',
                               lambda m: m.group(1) + esc_attr(next(it, '')) + '"', q_xml)
                new_parts[qp] = q_xml
        # ExternalData_N / filter names pointing at the old table range on this sheet
        q_sheet = "'%s'" % sh.replace("'", "''") if re.search(r"[^A-Za-z0-9_]", sh) else sh
        old_abs = '$' + re.sub(r'([A-Z]+)(\d+)', r'\1$\2', old_ref).replace(':', ':$')
        new_abs = '$' + re.sub(r'([A-Z]+)(\d+)', r'\1$\2', ref).replace(':', ':$')
        wb_xml = wb_xml.replace('>%s!%s<' % (escape(q_sheet), old_abs), '>%s!%s<' % (escape(q_sheet), new_abs))
        old_n = int(re.search(r'\d+$', old_ref).group(0)) - 1
        report.append('  %-26s %-26s %5d rows (was %d)  <- %s' % (sh, tname, nrow - 1, old_n, src))
    # recalc everything on open; old calc chain would point at removed cells
    if 'calcPr' in wb_xml:
        wb_xml = re.sub(r'<calcPr([^>]*?)/>', lambda m: '<calcPr' + re.sub(r'\s*fullCalcOnLoad="[^"]*"', '', m.group(1))
                        + ' fullCalcOnLoad="1"/>', wb_xml, count=1)
    else:
        wb_xml = wb_xml.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>')
    new_parts['xl/workbook.xml'] = wb_xml
    drop = {'xl/calcChain.xml'}
    ct = zin.read('[Content_Types].xml').decode('utf-8')
    new_parts['[Content_Types].xml'] = re.sub(r'<Override [^>]*?PartName="/xl/calcChain.xml"[^>]*/>', '', ct)
    wrels = zin.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    new_parts['xl/_rels/workbook.xml.rels'] = re.sub(r'<Relationship [^>]*?Target="[^"]*calcChain.xml"[^>]*/>', '', wrels)

    tmp = out_path[:-5] + '.tmp.xlsx'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename in drop:
                continue
            body = new_parts[item.filename].encode('utf-8') if item.filename in new_parts else zin.read(item.filename)
            zout.writestr(item, body, compress_type=zipfile.ZIP_DEFLATED)
    zin.close()
    # read-back check before the output is put in place
    try:
        with zipfile.ZipFile(tmp) as zc:
            for part in new_parts:
                ET.fromstring(zc.read(part))
        chk = openpyxl.load_workbook(tmp, read_only=True)
        chk.close()
    except Exception as e:                                 # noqa: BLE001
        os.remove(tmp)
        raise RuntimeError('ERROR: the new tracker does not read back (%s) - nothing was written' % e)
    os.replace(tmp, out_path)
    return report


def main():
    today = dt.date.today()
    out_path = os.path.join(HERE, '%s_%s.xlsx' % (TRACKER_STEM, today.isoformat()))
    cands = [f for f in glob.glob(os.path.join(HERE, '*%s*.xlsx' % TRACKER_STEM))
             if not os.path.basename(f).startswith('~$') and os.path.abspath(f) != os.path.abspath(out_path)]
    if not cands:
        sys.exit('ERROR: no %s*.xlsx in %s (the tracker is the template)' % (TRACKER_STEM, HERE))
    template = max(cands, key=os.path.getmtime)
    print('Template : %s' % os.path.basename(template))

    jobs_data, missing = {}, []
    for fname, sheet, tname, types in JOBS:
        path = find_file(fname)
        if not path:
            missing.append(fname)
            continue
        hdr, data, nform = read_source(path, sheet, types)
        if nform:
            print('WARNING %s / %s: %d cells hold formula text without a saved value - open and save the file '
                  'in Excel first' % (fname, sheet, nform))
        jobs_data[tname] = (hdr, data, '%s / %s' % (os.path.basename(path), sheet))
    for m in sorted(set(missing)):
        print('WARNING %s not found - its tables keep the old data' % m)
    if not jobs_data:
        sys.exit('ERROR: none of Sabanci.xlsx / Libra.xlsx / Crimson.xlsx found next to the script')

    print('Refreshed tables:')
    for line in refresh(template, out_path, jobs_data):
        print(line)
    print('Output   : %s  (formulas recalculate when Excel opens it)' % os.path.basename(out_path))


if __name__ == '__main__':
    main()
