"""check_align.py - ONE calculation, ONE value in every output (user 26-Nov-2026: "when I want a fix in an area, that fix also needs
to apply to every area with the same calc, e.g. Arrival at plant (Final AtA) in 6_Email = 4_Status = 1_MRP").

Run after `python update_inbound.py`, in the folder with the outputs (or give the folder):  python check_align.py [folder]
For every material it reads the same field from every output that shows it and reports any difference -> PASS / FAIL.
  1_MRP: ALERT, Material Status, Action List | 3_Inbound: PPC In Transit | 4_Status: BOM Material Dashboard - Weekly | 5_Web | 6_Email
Allowed (presentation only, not a difference): whole numbers on Material Status / Action List, line separators (new line / | / · / ;),
Action List row of a material on 2+ WSA sheets (= the whole material, the others = one row per sheet).
Stock Check is not compared: one row per supplier and its own 'Usage / day (next 7 days, roll-forward only)'.
Reads the company outputs only; writes nothing."""
import sys, os, re, glob, json, html as H, datetime as dt, collections
import openpyxl

def latest(folder, pat):
    f = sorted(glob.glob(os.path.join(folder, pat)))
    return f[-1] if f else None

def rows_of(path, sheet_end, key):
    """Data rows of the sheet whose name ends with sheet_end (numbered names, e.g. '01 ALERT'), header row = first row holding key."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sh = next((s for s in wb.sheetnames if s == sheet_end or s.endswith(' ' + sheet_end)), None)
    if sh is None: return []
    rs = list(wb[sh].iter_rows(values_only=True))
    hr = next((i for i, r in enumerate(rs[:15]) if key in [str(x).strip() if x else '' for x in r]), None)
    if hr is None: return []
    hd = [str(x).strip() if x else f'c{j}' for j, x in enumerate(rs[hr])]
    return [dict(zip(hd, r)) for r in rs[hr + 1:] if any(x not in (None, '') for x in r)]

def weekly_detail(path):
    """4_Status 'BOM Material Dashboard - Weekly' has TWO rows per material (user 27-Nov-2026): the line under a material row (UOM filled)
    that has text in column A and nothing in UOM = its detail row 'Stock zero ...  |  next: ...'. Returns {item name: [(zero text, next text), ...]}."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sh = next((x for x in wb.sheetnames if x.endswith('Weekly')), None); out = collections.defaultdict(list)
    if sh is None: return out
    rs = list(wb[sh].iter_rows(values_only=True))
    hr = next((i for i, r in enumerate(rs[:15]) if 'Item name' in [str(x).strip() if x else '' for x in r]), None)
    if hr is None: return out
    iu = [str(x).strip() if x else '' for x in rs[hr]].index('UOM')
    for i in range(hr + 1, len(rs) - 1):
        if rs[i][0] and rs[i][iu] and rs[i + 1][0] and not rs[i + 1][iu]:
            z, _, n = str(rs[i + 1][0]).partition('   |   next'); out[str(rs[i][0]).strip()].append((z.strip(), ('next' + n).strip()))
    return out

CODE = re.compile(r'^[A-Z]{2,}\d{4,}$')
def by_code(rows, key):
    out = collections.defaultdict(list)
    for r in rows:
        c = str(r.get(key) or '').strip()
        if CODE.match(c): out[c].append(r)
    return out

def num(v):
    if v in (None, '', '-', 'na'): return None
    if isinstance(v, (int, float)): return float(v)
    m = re.match(r'^\s*([\d,]+(?:\.\d+)?)', str(v))
    return float(m.group(1).replace(',', '')) if m else str(v)
def day(v):
    if isinstance(v, (dt.datetime, dt.date)): return v.strftime('%Y-%m-%d')
    return None if v in (None, '', '-') else str(v)
def txt(v):
    if isinstance(v, list): v = ' '.join(map(str, v))
    if v in (None, '', '-'): return None
    return re.sub(r'\s+', ' ', re.sub(r'\s*(?:[|·;]|\n)\s*', ' ', str(v))).strip()
def same(a, b):
    if {a, b} <= {None, 0.0}: return True                                              # no demand: 0 on one sheet, blank on another
    if isinstance(a, float) and isinstance(b, float): return abs(a - b) <= 0.51 + abs(a) * 1e-6     # whole numbers on some sheets
    return a == b

def main(folder):
    M, S, W, E, I = (latest(folder, p) for p in ('1_MRP_*.xlsx', '4_Status_*.xlsx', '5_Web_*.html', '6_Email_*.html', '3_Inbound_*.xlsx'))
    miss = [n for n, p in (('1_MRP', M), ('4_Status', S), ('5_Web', W), ('6_Email', E), ('3_Inbound', I)) if not p]
    if miss: print('FAIL: outputs not found in', folder, '-', ', '.join(miss)); return 1
    AL = by_code(rows_of(M, 'ALERT', 'Code'), 'Code')
    ST = by_code(rows_of(M, 'Status', 'Item code'), 'Item code')
    AC = by_code(rows_of(M, 'Actions', 'Component'), 'Component')
    t = open(W, encoding='utf-8').read(); i = t.find('const D=') + 8
    WEB = collections.defaultdict(list)
    for r in json.JSONDecoder().raw_decode(t[i:])[0]['rows']: WEB[r['c']].append(r)
    name2code, desc2code = {}, {}
    for c, rs in WEB.items():
        for r in rs: name2code.setdefault(H.unescape(r['n']), c); desc2code.setdefault(H.unescape(r['d']), c)
    for c, rs in ST.items():
        for r in rs: name2code.setdefault(re.sub(r'\n\(WSA sheet: (?:Glass G12R Topcon )?([^)]+)\)', r' (\1)', str(r['Item name'])), c)
    WK = collections.defaultdict(list)
    for r in rows_of(S, 'Weekly', 'Item name'):
        c = name2code.get(str(r.get('Item name') or '').strip())
        if c and r.get('UOM'): WK[c].append(r)
    DET = weekly_detail(S)
    for rs in WK.values():
        for r in rs:
            z, n = (DET.get(str(r.get('Item name') or '').strip()) or [(None, None)]).pop(0)
            r['Detail zero'], r['Detail next'] = z, n
    for rs in WEB.values():                                     # the 5_Web Summary 'next:' line (qty without the ', X within the lead time' tail)
        for r in rs:
            a = r.get('a21') or []; u = H.unescape(r.get('u') or '').strip()
            r['nxt'] = 'next: none within the lead time' if not a else f"next{f' ({u})' if u else ''}: " + ' · '.join(f"{d} {q:,.0f}" for d, q in a)
    e = open(E, encoding='utf-8').read(); EM = collections.defaultdict(list)
    for tb in re.findall(r'<table.*?</table>', e, re.S):
        trs = re.findall(r'<tr.*?</tr>', tb, re.S)
        hd = [H.unescape(re.sub('<[^>]+>', '', x)).strip() for x in re.findall(r'<t[hd][^>]*>(.*?)</t[hd]>', trs[0], re.S)] if trs else []
        if 'Arrival at plant (Final AtA)' not in hd: continue
        for tr in trs[1:]:
            cs = [re.sub('<[^>]+>', '', H.unescape(re.sub(r'<br\s*/?>', '\n', x))).strip() for x in re.findall(r'<td[^>]*>(.*?)</td>', tr, re.S)]
            r = dict(zip(hd, cs)); c = desc2code.get(r.get('Material')) or name2code.get(r.get('Material'))
            if c: EM[c].append(r)
    # Inbound PPC In Transit: counted quantity per material (= ALERT 'In transit' summed over the material's rows)
    IN = collections.defaultdict(float)
    for r in rows_of(I, 'PPCTransit', 'Supplier'):
        c = str(r.get('MRP code (after code conversion)') or '').strip(); q = r.get('Qty in MRP unit')
        if CODE.match(c) and isinstance(q, (int, float)) and isinstance(r.get('Counted on'), (dt.datetime, dt.date)): IN[c] += q
    multi = {c for c, rs in AL.items() if len(rs) > 1}          # one ALERT row per WSA sheet (back glass 333)
    PER = 'Stock zero from → to (⚠ = starts inside the transit lead time)'
    F = {'Estimated stock today': [('1_MRP ALERT', AL, 'Estimated stock today', num), ('1_MRP Material Status', ST, 'Stock', num), ('1_MRP Action List', AC, 'Stock', num), ('4_Status', WK, 'Estimated stock today', num), ('5_Web', WEB, 'st', num)],
         'SAP stock (MB52)': [('1_MRP ALERT', AL, 'SAP stock (MB52)', num), ('1_MRP Material Status', ST, 'SAP stock (MB52)', num), ('4_Status', WK, 'SAP stock (MB52)', num)],
         'Physical count': [('1_MRP ALERT', AL, 'Physical count', num), ('1_MRP Material Status', ST, 'Physical count', num), ('4_Status', WK, 'Physical count', num)],
         'Usage/Day': [('1_MRP ALERT', AL, 'Usage/Day (module plan)', num), ('1_MRP Material Status', ST, 'Usage/Day', num), ('1_MRP Action List', AC, 'Usage/Day (module plan)', num), ('4_Status', WK, 'Usage/Day (module plan)', num), ('5_Web', WEB, 'nd', num)],
         'Days of cover (stock only)': [('1_MRP ALERT', AL, 'Days of cover (stock only)', num), ('1_MRP Material Status', ST, 'Days of cover (stock only)', num), ('4_Status', WK, 'Days of cover (stock only)', num), ('5_Web', WEB, 'dy', num), ('6_Email', EM, 'Cover today', num)],
         'Days of cover (+ all in transit)': [('1_MRP ALERT', AL, 'Days of cover (+ all in transit)', num), ('1_MRP Material Status', ST, 'Days of cover (+ all in transit)', num), ('5_Web', WEB, 'di', num)],
         'Transit lead time': [('1_MRP ALERT', AL, 'Transit lead time (days)', num), ('1_MRP Action List', AC, 'Material lead time (days)', num), ('5_Web', WEB, 'lt', num)],
         'First stock-zero day': [('1_MRP ALERT', AL, 'First stock-zero day', day), ('1_MRP Material Status', ST, 'Short from', day), ('4_Status', WK, 'First stock-zero day', day)],
         'Stock zero from → to': [('1_MRP ALERT', AL, PER, txt), ('4_Status', WK, PER, txt), ('5_Web', WEB, 'per', txt), ('6_Email', EM, PER, txt)],
         'Weekly detail row: Stock zero from → to': [('1_MRP ALERT', AL, PER, txt), ('4_Status', WK, 'Detail zero', lambda v: None if v == 'no stock-zero day' else txt(v))],
         'Weekly detail row: next arrival': [('5_Web', WEB, 'nxt', txt), ('4_Status', WK, 'Detail next', txt)],
         'Arrival at plant (Final AtA)': [('1_MRP ALERT', AL, 'Arrival at plant (Final AtA)', txt), ('1_MRP Action List', AC, 'Arrival at plant (Final AtA)', txt), ('4_Status', WK, 'Arrival at plant (Final AtA)', txt), ('5_Web', WEB, 'ata', txt), ('6_Email', EM, 'Arrival at plant (Final AtA)', txt)],
         'Next arrival date': [('1_MRP ALERT', AL, 'Next arrival', lambda v: (re.match(r'(\d{2}-\w{3})', str(v)).group(1) if re.match(r'\d{2}-\w{3}', str(v or '')) else None)),
                               ('1_MRP Material Status', ST, 'Next arrival', lambda v: v.strftime('%d-%b') if isinstance(v, (dt.datetime, dt.date)) else None)],
         'In transit': [('1_MRP ALERT', AL, 'In transit', num), ('1_MRP Material Status', ST, 'In transit', num)]}
    bad = 0
    for f, srcs in F.items():
        rows_bad = []
        for c in sorted(set(AL) | set(ST) | set(WK) | set(WEB) | set(EM)):
            n = max(len(D.get(c, [])) for _, D, _, _ in srcs)
            for k in range(n):                                       # k-th row of the material in every source (333: one per WSA sheet)
                vals = {}
                for nm, D, key, fn in srcs:
                    rs = D.get(c, [])
                    if nm == '1_MRP Action List' and c in multi: continue        # Action List = the whole material
                    if nm == '6_Email' and c in multi: continue                  # email: one row per material
                    if k < len(rs): vals[nm] = fn(rs[k].get(key))
                v = list(vals.values())
                if len(v) > 1 and any(not same(v[0], x) for x in v[1:]): rows_bad.append((c, vals))
        if rows_bad: bad += len(rows_bad)
        print(f"{'OK  ' if not rows_bad else 'DIFF'}  {f}: {len(rows_bad)} difference(s)")
        for c, v in rows_bad[:8]: print('        ', c, json.dumps(v, ensure_ascii=False, default=str)[:600])
    d_ = [(c, sum(num(r.get('In transit')) or 0 for r in AL[c]), IN.get(c, 0.0)) for c in sorted(set(AL) | set(IN))]
    d_ = [x for x in d_ if abs(x[1] - x[2]) > 1]
    print(f"{'OK  ' if not d_ else 'DIFF'}  In transit 1_MRP ALERT vs 3_Inbound PPC In Transit (counted lines): {len(d_)} difference(s)")
    for c, a, b in d_[:8]: print('        ', c, f'ALERT {a:,.0f}  Inbound {b:,.0f}')
    bad += len(d_)
    print('PASS: every shared field has the same value in every output.' if not bad else f'FAIL: {bad} difference(s) - fix the RULE at its source so every output follows.')
    return 1 if bad else 0

if __name__ == '__main__':
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else '.'))
