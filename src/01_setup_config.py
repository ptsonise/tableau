"""
=================================================================
  MRP Daily + Inbound  —  one script for the planner
=================================================================

WHAT IT DOES
  Reads your daily WSA file, your MRP.xlsx, and the MIS RM Dashboard.
  Writes two files:
    - 1_MRP_<today>.xlsx   (ALERT, Material Status, Stock Check, MRP Daily, MRP by Line,
                                MRP Weekly, In Transit Status, Substitute Advice, Check, Settings, ...)
    - 3_Inbound_<today>.xlsx     (deliveries, to-chase, changes, ETD alert; no ALERT sheets - the ALERT is on 1_MRP)
    - 4_Status_<today>.xlsx  (sheet 1 Direct Material = the 'Material Status' sheet; sheet 2 Other Materials = PPC Count materials not in the BOM)
    - 6_Email_<today>.eml / .html   (the daily email to top management: double-click the .eml = an Outlook draft with the
                                         4_Status file attached; add the To line, check and Send)

HOW TO RUN
  Put these three files in the same folder:
    1. WSA - Raw Material planning per day.xlsx
    2. MRP.xlsx
    3. MIS RM Dashboard - WSA.xlsx

  Then double-click or run:
    python update_inbound.py

  Or with the file names:
    python update_inbound.py "WSA - Raw Material planning per day.xlsx" MRP.xlsx "MIS RM Dashboard - WSA.xlsx"

  EVERY RUN ALSO REBUILDS THE WSA PLANNING SHEETS (since Nov 2026) - no flag needed:
    YOUR WSA FILE IS ONLY READ, NEVER CHANGED. The result is a new file '2_WSA_<date>.xlsx' next to it (older 2_WSA
    files are deleted, like 1_MRP), and the Inbound and 1_MRP files are built from it. Front and back glass get their own sheets (one Balance
    per glass: Glass Topcon Front / Back, Glass G12R Topcon SB Front / Back; the Subanci back glass goes to SB Back). The sheets
    Total_Plan_Req is left out (not needed); Stock is kept (Summary Dashboard and the Frame sheets read it). Every planning sheet gets ONE layout:
      rows 1-6 material info (stock, supplier & material, unit, codes, totals), row 7 headers, one row per day from the
      1st of the current month (older rows are deleted - their delivery records go to the 'Archive' sheet) to 31-Mar.
      A Date | B Day | C Balance | D Planned modules | E Demand | F-M customer columns (same customers, same order on every
      sheet) | N-S Actual modules, Actual, MW, Actual MW, Plant Days, Plant + Port Days | T PPC In Transit (arrivals) |
      U PPC In Transit (Supplier Name) | V Plant + Port Days (PPC In Transit) | W Delivered | X Scheduled Delivery | Y PO |
      Z Shipper | AA Remarks | ... | AG Status | ...
    PPC In Transit = your PPC file's In_Transit sheet: Final Pending Qty. arriving on Final AtA (late lines = today); each line is
      traced to YOUR material: first your WSA sheet with its PO + BL# (HBLNo. / MBL), then supplier + description + RM Code +
      Final RM Code (Inbound sheet 'PPC In Transit' shows how, and CHECK lines where your WSA sheet has another size / HS);
      1_MRP traces the PPC In_Transit lines (the same shipments) the same way, so all three files count the same material;
      V = (Balance + those arrivals up to the day) / Demand. The last console line says how many In_Transit lines were read
      and on how many sheets they arrive - or, in a WARNING, why column T is empty (the same reason is in T3 of every sheet).
    BOM materials WITHOUT a WSA sheet get the same calculation in memory -> Inbound sheet 'Other materials' + 'Below zero'.
    From today the planning columns are LIVE formulas on the last sheets, PPC_Plan_Req (module plan x BOM) and Archive:
      Balance = stock typed on today's row (yellow), then Balance - Actual + Delivered;
      Plant + Port Days = (Balance + open Scheduled Deliveries up to that day, not cancelled + U2 = open deliveries in the
      Archive) / Demand. When an archived delivery arrives, set its State to 'Delivered' on the Archive sheet.
    Rows before today keep their numbers. Delivered and the columns after it are copied as they are.
    Potting Material and Tapes (two materials side by side) become two sheets each. The Dashboard reads column S (19).
    Each planning sheet opens at today's row. Keep typing in YOUR WSA file; 2_WSA is rebuilt from it on every run.
    To skip the rebuild (old behaviour):  python update_inbound.py --no-align
    Older output files (1_MRP, 2_WSA, 4_Status, 6_Email, 7_Summary, Inbound) are deleted
    once the new ones are saved. To keep them:  python update_inbound.py --keep-old

WHAT YOU CAN CHANGE
  Everything is on the 'Settings' sheet (copy it into your MRP.xlsx once and edit it there). It is read FIRST, before
  anything is calculated, so the Inbound file and 1_MRP always use the same rules. The CONFIG values below are only
  the defaults used when there is no Settings sheet yet.

STATUS RULES (see material_status())
  - Material status: compares Inv. Days with that material's OWN target days (A / B / C / packaging, + Z extra):
      CRITICAL = runs out within its target even with Open PO
      AT RISK  = runs out within its target unless the Open PO ships
      WATCH    = covered only if in transit arrives on time, or falls below target within its lead time
      OK       = everything else
  - Line cover: lines that can run fewer than COVER_CRITICAL days are flagged on Material Status (Settings).

Needs: Python 3.9+, pandas, openpyxl, lxml
"""
import bisect
import sys, os, re, glob, time, math, datetime as dt, warnings
# OUTPUT FILE NAMES - ONE table (user, Nov 2026: short names that start with a number, so the planner can say "fix 3"). Change a name here only.
OUT_NAME={'mrp':'1_MRP','wsa':'2_WSA','inb':'3_Inbound','dms':'4_Status','web':'5_Web','mail':'6_Email','sum':'7_Summary','rev':'8_Review'}
_OUT_OLD={'mrp':'MRP_Daily','wsa':'WSA_Daily','inb':'Inbound','dms':'Daily_Material_Status','web':'Material_Daily_Update','mail':'Daily_Email','sum':'Run_Summary','rev':'Review_Panel'}
def outfile_name(k,ext=''): return f'{OUT_NAME[k]}_{dt.date.today():%Y-%m-%d}'+('.'+ext if ext else '')
def out_rx(k): return '(?:'+OUT_NAME[k]+'|'+_OUT_OLD[k]+')'            # today's name or the old long name (old files are still cleaned up)
def is_output_file(n): return n.startswith(tuple(v+'_' for v in OUT_NAME.values())+tuple(v+'_' for v in _OUT_OLD.values()))
# OUTPUT SHEET NAMES - ONE table (user, Nov 2026: "rename all sheets, you know the priority of each"). (name the code builds, short name) in PRIORITY
# order: the tab order and the number in front follow it ('01 ALERT' = sheet 1 of that file). The files are built with the long names (formulas,
# readers); once a file is finished `rename_sheets` renames it (formulas, print areas, links follow). 2_WSA keeps the planner's own sheet names.
OUT_SHEETS={
 'mrp':[('ALERT','ALERT'),('Material Status','Status'),('ETD alert','ETD'),('Action List','Actions'),('ALERT (formulas)','AlertFormulas'),('Stock Check','StockCheck'),
        ('In Transit Status','Transit'),('Substitute Advice','Substitutes'),('MRP Daily','Daily'),('MRP Weekly','Weekly'),('MRP by Line','ByLine'),('Check','Check'),
        ('No Demand','NoDemand'),('Settings','Settings'),('Movement Log','Moves'),('Module Plan','ModulePlan'),('BOM','BOM'),('Stock Count','Count'),
        ('Stock Data','StockData'),('ALERT Calc','AlertCalc'),('ALERT Plan lines','AlertPlan'),('ALERT Arrival lines','AlertArrivals'),('How it works','HowItWorks'),
        ('MIS Layout','MISLayout'),('Definitions','Definitions')],
 'inb':[('ETD alert','ETD'),('Summary','Summary'),('Below zero','BelowZero'),('To chase','Chase'),('Inbound','Inbound'),('PPC In Transit','PPCTransit'),
        ('Other materials','Other'),('By Week','ByWeek'),('Changes','Changes'),('Read me','ReadMe')],
 'dms':[('BOM Material Dashboard - Weekly','Weekly'),('RM Issues','Issues')]}
def sheet_new(key,old):
    return next((f'{i_+1:02d} {s_}' for i_,(o_,s_) in enumerate(OUT_SHEETS[key]) if o_==old),old)
def rename_sheets(path,key,out=None,back=False):
    """Rename the sheets of a FINISHED output file to the numbered names of OUT_SHEETS (back=True: the other way) and put the tabs in priority
       order. Edits the xlsx XML only, so cached values stay: formulas, print areas, hyperlinks and the quoted sheet names in texts follow.
       Returns (changed, names that are not in the table). Nothing is written when no sheet of the table is in the file."""
    import zipfile, html
    pairs=[(o_,f'{i_+1:02d} {s_}') for i_,(o_,s_) in enumerate(OUT_SHEETS[key])]
    if back: pairs=[(n_,o_) for o_,n_ in pairs]
    mp=dict(pairs); pos={a_:i_ for i_,(a_,b_) in enumerate(pairs)}; out=out or path; tmp=out+'.rename.tmp'
    esc=lambda s_: html.escape(s_,quote=False)
    with zipfile.ZipFile(path) as zi:
        wbx=zi.read('xl/workbook.xml').decode('utf-8')
        sheets=re.findall(r'<sheet\b[^>]*?/>',wbx)
        have=[html.unescape(re.search(r'\bname="([^"]*)"',s_).group(1)) for s_ in sheets]
        if not any(n_ in mp for n_ in have): return False,[]
        gp=[]; rep={}                                    # one regex for every sheet reference: 'Old'!A1  or  Old!A1  ->  'New'!A1
        for i_,o_ in enumerate(n_ for n_ in have if n_ in mp):
            p_="(?:'|&apos;)"+re.escape(esc(o_.replace("'","''")))+"(?:'|&apos;)!"
            if re.fullmatch(r'[A-Za-z_][A-Za-z0-9_.]*',o_): p_+="|(?<![\\w.'$!])"+re.escape(o_)+"!"
            gp.append(f'(?P<g{i_}>{p_})'); rep[f'g{i_}']="'"+esc(mp[o_].replace("'","''"))+"'!"
        rx=re.compile('|'.join(gp)); fx=lambda t_: rx.sub(lambda m_: rep[m_.lastgroup],t_) if '!' in t_ else t_
        tx=[(re.compile(r"(?<![\w])(['\"])"+re.escape(esc(o_))+r"\1(?![\w])"),o_) for o_ in have if o_ in mp]
        order=sorted(range(len(sheets)),key=lambda i_:(pos.get(have[i_],10**6),i_)); newidx={o_:n_ for n_,o_ in enumerate(order)}
        with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED,compresslevel=1) as zo:
            for it in zi.infolist():
                data=zi.read(it.filename); fn=it.filename
                if fn=='xl/workbook.xml':
                    t_=fx(wbx); b_=re.search(r'<sheets>.*?</sheets>',t_,re.S)
                    ns_=''.join(re.sub(r'\bname="[^"]*"','name="'+html.escape(mp.get(have[i_],have[i_]),quote=True)+'"',sheets[i_],count=1) for i_ in order)
                    t_=t_[:b_.start()]+'<sheets>'+ns_+'</sheets>'+t_[b_.end():]
                    data=re.sub(r'localSheetId="(\d+)"',lambda m_:'localSheetId="%d"'%newidx.get(int(m_.group(1)),int(m_.group(1))),t_).encode('utf-8')
                elif fn.startswith(('xl/worksheets/','xl/charts/')) and fn.endswith('.xml'): data=fx(data.decode('utf-8')).encode('utf-8')
                elif fn=='xl/sharedStrings.xml':
                    t_=data.decode('utf-8')
                    for r_,o_ in tx: t_=r_.sub(lambda m_: m_.group(1)+esc(mp[o_])+m_.group(1),t_)
                    data=t_.encode('utf-8')
                zo.writestr(it,data)
    os.replace(tmp,out)
    return True,[n_ for n_ in have if n_ not in mp]
_OLDCOPY={}
def old_names_copy(path):
    """A previous day's output has numbered sheet names; the readers ask for the long names: give them a temp copy with the long names."""
    try:
        bn=os.path.basename(path)
        key=next((k_ for k_ in OUT_SHEETS if bn.startswith(OUT_NAME[k_]+'_') and bn.lower().endswith('.xlsx')),None)
        if not key or not os.path.exists(path): return path
        st=os.stat(path); ck=(os.path.abspath(path),st.st_mtime,st.st_size)
        if ck not in _OLDCOPY:
            if not _OLDCOPY:
                import tempfile, atexit, shutil; _OLDCOPY[None]=tempfile.mkdtemp(prefix='oldnames_'); atexit.register(shutil.rmtree,_OLDCOPY[None],True)
            tp=os.path.join(_OLDCOPY[None],f'{len(_OLDCOPY)}_'+bn)
            _OLDCOPY[ck]=tp if rename_sheets(path,key,out=tp,back=True)[0] else path
        return _OLDCOPY[ck]
    except Exception as e_: print(f'NOTE: sheet names of {os.path.basename(path)} not restored ({type(e_).__name__}: {e_})'); return path
def finalize_sheet_names(*files):
    """Last step of a run, once the files are final (before the email attaches 4_Status): number the sheets. files = (key, path)."""
    for key,p_ in files:
        if not os.path.exists(p_): continue
        try:
            ch_,un_=rename_sheets(p_,key)
            if un_: print(f'NOTE: {os.path.basename(p_)}: sheets not in OUT_SHEETS (left as they are, at the end): '+', '.join(un_))
        except Exception as e_: print(f'NOTE: sheet numbers not set on {os.path.basename(p_)} ({type(e_).__name__}: {e_}) - close it in Excel and run again')
_T0=time.time()      # start of the run - progress lines show seconds since start
SWALLOWED={}      # source line -> [count, exception type, message]: errors caught by a silent handler (listed in the run summary)
def _swallow(line,e):
    # the real line comes from the exception's traceback; the number typed at each call site is only a fallback
    # (21 of the 53 typed numbers had drifted after edits, so the run summary pointed at the wrong lines)
    tb=getattr(e,'__traceback__',None); ln=tb.tb_lineno if tb is not None else line
    x=SWALLOWED.setdefault(ln,[0,type(e).__name__,str(e)[:120]]); x[0]+=1
warnings.filterwarnings('ignore')
import pandas as pd
import openpyxl as _oxl_pkg
# SPEED: openpyxl saves with zip level 6; level 1 writes the same .xlsx content ~2x faster (files ~10% larger, Excel opens them the same)
try:
    import openpyxl.writer.excel as _oxl_writer, zipfile as _zf
    class _FastZip(_zf.ZipFile):
        def __init__(self,*a,**k): k.setdefault('compresslevel',1); super().__init__(*a,**k)
    _oxl_writer.ZipFile=_FastZip
except Exception: pass
if not str(_oxl_pkg.__version__).startswith('3.1'):
    print(f'WARNING: the speed patches below were written for openpyxl 3.1.x (you have {_oxl_pkg.__version__}). If a file fails to open or save, run: pip install "openpyxl==3.1.5"',flush=True)
# SPEED: the same sheet of an input file was parsed 3 times per step (WSA planning sheets: extract_fast, _wsa_read, dashboard code).
# A read-only sheet read in full with values_only (every row, every column) is kept in memory the first time it is read to the end;
# later full reads and row ranges inside it are served from it. Same rows, same values; nothing is kept when a reader stops early,
# and the key holds the file's size + modified time, so a file saved during the run is read again.
import itertools as _it_
import openpyxl.worksheet._read_only as _oxl_ro
_RO_ROWS={}
_ro_orig_cbr=_oxl_ro.ReadOnlyWorksheet._cells_by_row
def _ro_cbr(self, min_col, min_row, max_col, max_row, values_only=False):
    try:
        fn=self.parent._archive.filename; st=os.stat(fn)
        key=(os.path.abspath(fn),st.st_size,st.st_mtime_ns,self.title,bool(self.parent.data_only))
        ok=bool(values_only) and min_col==1 and max_col is not None and max_col==self.max_column and max_row is not None and max_row==self.max_row
    except Exception: ok=False
    if not ok: return _ro_orig_cbr(self,min_col,min_row,max_col,max_row,values_only)
    rows=_RO_ROWS.get(key)
    if rows is not None:
        if (min_row or 1)==1 and max_row==self.max_row: return iter(rows)
        if (min_row or 1)>=1 and max_row<=len(rows): return iter(rows[(min_row or 1)-1:max_row])
        return _ro_orig_cbr(self,min_col,min_row,max_col,max_row,values_only)
    if (min_row or 1)!=1: return _ro_orig_cbr(self,min_col,min_row,max_col,max_row,values_only)
    def gen():
        buf=[]
        for r in _ro_orig_cbr(self,min_col,min_row,max_col,max_row,values_only):
            buf.append(r); yield r
        _RO_ROWS[key]=buf
    return gen()
_oxl_ro.ReadOnlyWorksheet._cells_by_row=_ro_cbr
# SPEED: when a workbook is opened read-only (to read values), only number formats are needed (to recognise dates).
# openpyxl normally rebuilds every font / fill / named style first - the WSA file has ~36,000 named styles, which took
# about 70% of the run time. Read-only loads now parse just the number formats; normal loads (copying formats) are unchanged.
import openpyxl.reader.excel as _oxl_reader
from openpyxl.styles.stylesheet import Stylesheet as _OxlStylesheet
from openpyxl.xml.functions import fromstring as _oxl_fromstring
_oxl_apply_full=_oxl_reader.apply_stylesheet
_STYLE_CACHE={}       # styles.xml parsed once per distinct content (every input file is opened several times per run)
def _oxl_apply_fast(archive, wb):
    if not getattr(wb,'_read_only',False): return _oxl_apply_full(archive, wb)
    try: raw=archive.read('xl/styles.xml')
    except KeyError: return wb
    ss=_STYLE_CACHE.get(raw)
    if ss is None:
        node=_oxl_fromstring(raw)
        for ch in list(node):
            if ch.tag.rsplit('}',1)[-1] not in ('numFmts','cellXfs'): node.remove(ch)
        ss=_STYLE_CACHE[raw]=_OxlStylesheet.from_tree(node)
    wb._cell_styles=ss.cell_styles; wb._number_formats=ss.number_formats; wb._date_formats=ss.date_formats
    wb._timedelta_formats=ss.timedelta_formats; wb._alignments=ss.alignments; wb._protections=ss.protections
    return wb
_oxl_reader.apply_stylesheet=_oxl_apply_fast
# SPEED: assigning a Font / Fill / Border / Alignment to a cell hashes the style object every time (openpyxl serialises all its
# fields) - about 40% of the run time. The same object is assigned thousands of times, so its index in the workbook's style table
# is remembered. The cache is dropped whenever the table is rebuilt; the style table and the saved file stay identical.
from openpyxl.utils.indexed_list import IndexedList as _OxlIndexedList
def _oxl_il_add(self, value):
    d=self._dict; c=self.__dict__.get('_idc')
    if c is None or c[0] is not d: c=self.__dict__['_idc']=(d,{})
    hit=c[1].get(id(value))
    if hit is not None and hit[0] is value: return hit[1]
    self.append(value); idx=self._dict[value]; c[1][id(value)]=(value,idx)
    return idx
_OxlIndexedList.add=_oxl_il_add
# SPEED: the script builds the same Font / PatternFill / Alignment / Border / Side again for every cell. Identical arguments now give
# the SAME object (these objects are never changed after they are made), so the style table lookup above hits its cache and
# openpyxl does not rebuild and hash thousands of identical objects. The saved files are identical.
import openpyxl.styles as _oxl_styles
def _oxl_memo_style(cls):
    cache={}
    def make(*a, **k):
        try:
            key=(a,tuple(sorted(k.items()))); hit=cache.get(key)
        except TypeError: return cls(*a, **k)          # an unhashable argument: build normally
        if hit is None: hit=cache[key]=cls(*a, **k)
        return hit
    make.__name__=cls.__name__
    return make
for _n_ in ('Font','PatternFill','Alignment','Border','Side'): setattr(_oxl_styles,_n_,_oxl_memo_style(getattr(_oxl_styles,_n_)))
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, Reference
from collections import Counter

# =================================================================
#  CONFIG  —  change these numbers if you want different rules
# =================================================================
# Target days, ALERT 'under N days', safety stock and look-ahead = each material's TRANSIT LEAD TIME (`tgt_days`, user rule Nov 2026);
# a production line's threshold = the longest lead time of its materials (LINE_LT). Only the fallback lead times below are numbers.
LEAD_TIME_RM    = 49     # days from order to plant for raw materials when the WSA file has no ETD history (used for planned order release)
LEAD_TIME_CSPM  = 14     # days from order to plant for packaging / consumables
def lt_of(b):
    """Transit lead time (days) of a material: Settings TRANSIT LEAD TIME 'Your lead time', else median ETD -> plant days of its WSA deliveries, else Settings lead time RM / packaging."""
    return max(1,int(b.get('lt_used') or (LEAD_TIME_CSPM if b.get('count_stock') else LEAD_TIME_RM) or 1))
def tgt_days(lt):
    """Days threshold of a material = its transit lead time (user rule, Nov 2026: Settings rows 6-19 'change all values to transit
       lead time'): target days, ALERT 'under N days' / 'below N days from', safety stock days and look-ahead all use it."""
    return max(1,int(lt or LEAD_TIME_RM or 1))
def usage_day(P, lt, k0=0):
    """Usage/Day (user rule, Nov 2026: 'only Usage/Day and Inv. Days'): average plan usage per production day (plan > 0) over the
       material's transit lead time, days k0 .. k0 + lt - 1. None = no production in that window. Inv. Days = stock / Usage/Day."""
    n=max(1,int(lt or LEAD_TIME_RM or 1)); nz=[v for v in list(P or [])[k0:k0+n] if v>0]
    return (sum(nz)/len(nz)) if nz else None
# SAFETY STOCK / REORDER POINT (user 28-Nov-2026, option A 'service level', replaces 'safety stock = lead time of usage' everywhere):
#   SS = z(service level of the ABC x XYZ class) x sqrt(LT x sd_day^2 + Usage/Day^2 x sd_LT^2) ; ROP = Usage/Day x LT + SS
SS_SL={'AX':98,'AY':97,'AZ':95,'BX':95,'BY':95,'BZ':92,'CX':90,'CY':90,'CZ':90,'PM':90}   # service level % per class - Settings table SAFETY STOCK
SS_LT_SD_PCT=25                     # lead-time variation (% of lead time) when the WSA history has < 2 deliveries - Settings row
SS_STATS={}                         # material code -> dict(cv = daily demand variation, sd_lt = lead-time std days, n_lt, cls = 'AX'..'PM')
def demand_cv(P, lt, k0=0):
    """Variation of daily usage (std / mean) over the same window and days as usage_day (production days of the lead time)."""
    n=max(1,int(lt or LEAD_TIME_RM or 1)); nz=[v for v in list(P or [])[k0:k0+n] if v>0]
    if len(nz)<2: return 0.0
    m=sum(nz)/len(nz); return (sum((v-m)**2 for v in nz)/len(nz))**0.5/m
def ss_rop(code, need, lt):
    """Safety stock + reorder point of a material: ONE function for MRP Daily, Material Status, RM Issues, 4_Status, 2_WSA.
       need = Usage/Day of the row, lt = its transit lead time. Returns dict(ss, rop, sl, z, cls, sd_d, sd_lt)."""
    from statistics import NormalDist
    s=SS_STATS.get(str(code or '').upper()) or {}; d=float(need or 0.0); L=tgt_days(lt)
    cls=s.get('cls') or 'CZ'; sl=float(SS_SL.get(cls) or SS_SL.get('CZ') or 90)
    z=NormalDist().inv_cdf(min(99.9,max(50.0,sl))/100)
    sd_d=float(s.get('cv') or 0.0)*d
    sd_lt=s.get('sd_lt') if s.get('sd_lt') is not None else L*SS_LT_SD_PCT/100
    ss=z*(L*sd_d**2+(d*float(sd_lt))**2)**0.5 if d>0 else 0.0
    return dict(ss=ss,rop=d*L+ss,sl=sl,z=z,cls=cls,sd_d=sd_d,sd_lt=float(sd_lt))
def ss_xl(code, lt):
    """The constants of ss_rop for an Excel formula (ALERT (formulas), MRP Daily): z, daily-usage variation, lead-time std days."""
    s=SS_STATS.get(str(code or '').upper()) or {}; q=ss_rop(code,1.0,lt)
    return dict(z=q['z'],cv=float(s.get('cv') or 0.0),sd=q['sd_lt'])

# materials treated as UNLIMITED stock (never short, left out of shortage planning)
UNLIMITED_CODES = {'PMPKM00000338','PMPKM00000523'}   # material codes
UNLIMITED_WORDS = ('plastic bag',)                     # or any material whose description contains these words

# other codes of the SAME material in another pack / unit to count for an MRP material:  {code: (MRP code, factor)}
# factor = how many MRP units one unit of the code is.  Example:  PACK_FACTORS = {'RMBUS00000032': ('RMBUS00000047', 1)}
PACK_FACTORS    = {}
_PACK_FACTORS_SCRIPT=dict(PACK_FACTORS)    # kept: the Settings SUBSTITUTES (all lines) are added on top, they do not wipe these

PLAN_WSA_LINES_ONLY = False   # True = count only MIS plan rows of lines USA* with OEM Category WSA (old rule)
# EASIER: change these in the 'Settings' sheet of 1_MRP (or copy that sheet into your MRP.xlsx) - it overrides the values above.
SETTINGS_SRC=[None]
ABC_A_PCT=80; ABC_B_PCT=95          # ABC cut-offs (% of value)
XYZ_X=0.25; XYZ_Y=0.5               # XYZ cut-offs (coefficient of variation of weekly demand)
ABC_HORIZON=90                      # days of requirement used for the ABC value
SHORT_HORIZON=60                    # days the shortage vs target is measured over
DEMAND_HORIZON=None                 # No Demand rule: days ahead in which a material must have demand (production requirement / BOM demand / daily target); None = the shortage horizon above
SHORT_LEAD=20; LONG_LEAD=75         # Inbound: days between ETD and plant date outside which a delivery is flagged
ALERT_HORIZON=112                   # look-ahead (days) of 'below lead time later' / Why / RM Issues (112 = 16 weeks) - Settings row
RISK_DAYS=5                         # risk order of every dashboard: after the first day below 1, the first day the cover is below this many days - Settings row
DASH_STEP2=14                       # dashboard colours (1_MRP dark): 2nd step ends at this cover (days); 1st step = 1 .. COV_RED_DAYS - Settings row
DASH_BLUE1=14                       # dashboard colours: 1st blue = lead time .. lead time + this many days - Settings row
DASH_BLUE2=39                       # dashboard colours: 2nd blue ends at lead time + this many days - Settings row
DASH_MIN_LT=14                      # dashboard colours: a lead time under this counts as this many days - Settings row
def HZW(): return f'{max(1,round(ALERT_HORIZON/7))} weeks'     # '16 weeks' in the texts follows ALERT_HORIZON
RELEASE_WINDOW=7                    # 'RELEASE ORDER' when an order is due within this many days
RESCHED_OUT_X=3                     # 'RESCHEDULE OUT' when stock is more than this many times the target with more arriving
SET_CUSTSUP=[]; CUSTSUP_ROWS=[]; SUBS_LINES=[]; SUBS_ISSUES=[]         # Customer Supplier / Vendor Priority rules read from the Settings sheet (if filled there)
LINE_RANK={}       # line key -> rank for all materials (Settings 'Line ranking', row 'All materials'); empty = closest to stopping first
LINE_RANK_MAT={}   # material (code / name / description as typed) -> {line key: rank}  (Settings 'Line ranking', one row per material)
ADDED_MATS=[]      # materials added because the plan needs them but the MIS Daily Dashboard does not list them: (code, name, module type)
TGT_CELL={}        # line key -> Settings cell of its daily target (Calc Target formulas point at it)
IT_STATUS_ONLY=False  # in transit = only PPC In_Transit lines whose STATUS is 'IN-TRANSIT' (not at origin / at port / delivered / partly received ...)
def it_ok(r):
    return (not IT_STATUS_ONLY) or (len(r)>22 and it_status_group(r[22])=='In transit')
MIS_PATH_=[None]   # the MIS file of this run (batch split of glass stock)
ALLOC_IT={}        # (material, line) -> in-transit quantity from suppliers that line may use (Customer Supplier)
PROJECT_OF={}      # your customer -> 'Project Name' (Settings MIS LINK table, if you add that column)
PROJ_COL=[None]
def project_of(cu):
    """Project Name of a customer (MIS LINK table): exact name, else via its MIS customer, else the closest spelling
       (e.g. 'Clenera Crimson' in the plan = 'Cleanera Crimson' in the table); else the customer name itself."""
    import difflib
    k=_norm(cu)
    if k in PROJECT_OF: return PROJECT_OF[k]
    L_=next((v for k2,v in LINKS.items() if _norm(k2)==k),None) or {}
    if _norm(L_.get('mis')) in PROJECT_OF: return PROJECT_OF[_norm(L_.get('mis'))]
    m=difflib.get_close_matches(k,list(PROJECT_OF),n=1,cutoff=0.85)
    return PROJECT_OF[m[0]] if m else cu
SUP_CODE={'KB':'KIBING','KN':'KIBING','FL':'FLAT','XY':'XINYI','XI':'XINYI'}   # Stock Count supplier code -> vendor name
STOCK_EST={}       # code -> SAP stock, count, count date, usage / receipts since, estimated stock today (Stock Check sheet)
COUNTED_RM={}      # raw material -> physical count used instead of MIS MB52 (your Stock Count)
SDG={}             # material -> {(type, source, source code): qty, uom, factor} (Action List source columns, RM / PM Status)
AL_VAL={}          # (material, Action List header) -> value written on the Action List (MRP by Line / status sheets copy it)
AL_ORDER_=[]; COLS_=[]   # Action List order / columns of this run (Materials Status)
AL_COPY=[]         # (header, Action List column) copied into MRP by Line
IT_TODAY_HEAD='In transit today (total qty, PPC In_Transit)'   # Action List column; MRP by Line + Action List formulas find it by this name
PLAN_LINE_ROWS={}  # (material, line key) -> Calc Plan row of that line's plan (MRP by Line formulas)
CALC_LAYOUT=[[],{}]  # Calc Target rows / Calc Plan rows of this run (Action List formulas point at them)
PHYS_DATE={}       # material -> 'Physical Count Date' (Stock Count in MRP.xlsx, column with that header)
PHYS_COUNT={}      # material -> 'Physical Count Qty_BOM' (column G of Stock Count in MRP.xlsx, count in the BOM unit), for RM / PM Status
VALUES_ONLY=True   # True = every sheet holds plain values (fast file); False = live formulas (Stock Data -> Action List -> ...)
ADVICE_ISSUES=[]   # Check-sheet issues found while building Substitute Advice (e.g. MIS JB Type vs MIS BOM)
BATCH_SUP={}       # batch supplier code (2nd-3rd character of the SAP batch, e.g. 1JU0369004 -> JU) -> supplier name, from Settings
BATCH_SEEN={}      # code -> dict(mats=set, guess=name) found in MB52 this run (written to Settings for you to name)
MOD_SHORT={}       # MIS BOM name -> short module name (e.g. G12R 3.2, M10R EPEP)
LINE_DISP={}       # line key -> line name as in your Production Plan (without 'USA ')
LINE_ORD={}        # line key -> sort order (Line 1..4 first)
UCMAP={}           # (MIS customer, line) -> {your customer: modules} - customer shares in your own customer names
CUST_TARGET_DEFAULT=[('Pepper Solar Farm LLC','L1',3000,'No'),('Lucky 7 Solar Farm LLC','L2',3000,'No'),('Clenera Crimson','L3',3600,'No'),
                     ('SB Energy G12R','L4',3600,'No'),('Clenera Snowflake','OEM',4000,'Yes')]       # from your PPC 'Usage' sheet
CUSTOMER_TARGETS=True   # True = daily need from the customer daily targets (your PPC Usage way)
CUST_TARGET=[]     # [(customer, line, modules per day, OEM 'Yes'/'No')] - Settings 'CUSTOMER DAILY TARGET'; empty -> CUST_TARGET_DEFAULT
EFF_TARGET={}      # LINE_TARGET + lines without a Settings target at their highest planned day
LINE_TODAY={}      # production line key -> day index of its first production day from the run date (for the daily-target need)
LINE_TARGET={}     # production line key -> (label, modules per day) from Settings 'USA Line N Production Plan Daily Target'
SHARED_MATS=[]     # materials used by 2+ lines this run: (name, description, code, [lines]) - listed on the Settings sheet
LINES_SEEN=[]      # production line names found in this run (to list them on the Settings sheet)
# =================================================================

# =================================================================
#  STATUS RULES  -  the ONLY place statuses are decided. Every sheet (Action List, Today, Planner, MRP by Line,
#  Line Run Days, Supplier Coverage, e-mails) calls these, so a material has one status everywhere.
# =================================================================
STATUS_ORDER  = {'CRITICAL':0,'AT RISK':1,'WATCH':2,'OK':3,'NO DEMAND':4}
STATUS_FILL   = {'CRITICAL':'FFC7CE','AT RISK':'FFEB9C','WATCH':'FFF2CC','OK':'C6EFCE','NO DEMAND':'EDEDED'}
STATUS_FONT   = {'CRITICAL':'9C0006','AT RISK':'9C5700','WATCH':'7F6000','OK':'006100','NO DEMAND':'595959'}
def material_status(cov, td, below=None, lead=None):
    """cov = cover_days() result (days to run-out: stock only, + in transit, + in transit + Open PO);
       td = this material's target days = its transit lead time (`tgt_days`);
       below = first day projected stock falls under the target (day index) ; lead = lead time used for the material."""
    if cov[2]<td: return 'CRITICAL'
    if cov[1]<td: return 'AT RISK'
    if lead and cov[2]<lead: return 'AT RISK'       # runs out inside its supply lead time even with Open PO: only expediting can save it
    if cov[0]<td: return 'WATCH'
    if below is not None and below<=(lead or LEAD_TIME_RM)+14: return 'WATCH'
    return 'OK'

