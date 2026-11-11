"""
=================================================================
  MRP Daily + Inbound  —  one script for the planner
=================================================================

WHAT IT DOES
  Reads your daily WSA file, your MRP.xlsx, and the MIS RM Dashboard.
  Writes two files:
    - MRP_Daily_<today>.xlsx   (ALERT, Material Status, BOM Material Dashboard - Weekly / - Daily, Stock Check, MRP Daily, RM Issues, MRP by Line,
                                MRP Weekly, In Transit Status, Substitute Advice, Check, Settings, ...)
    - Inbound_<today>.xlsx     (deliveries, to-chase, changes)
    - Daily_Material_Status_<today>.xlsx  (sheet 1 Direct Material = the 'Material Status' sheet; sheet 2 Other Materials = PPC Count materials not in the BOM)
    - Daily_Email_<today>.eml / .html   (the daily email to top management: double-click the .eml = an Outlook draft with the
                                         Daily_Material_Status file attached; add the To line, check and Send)

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
    YOUR WSA FILE IS ONLY READ, NEVER CHANGED. The result is a new file 'WSA_Daily_<date>.xlsx' next to it (older WSA_Daily
    files are deleted, like MRP_Daily), and the Inbound and MRP_Daily files are built from it. Front and back glass get their own sheets (one Balance
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
      MRP_Daily traces the MIS In Transit lines (the same shipments) the same way, so all three files count the same material;
      V = (Balance + those arrivals up to the day) / Demand. The last console line says how many In_Transit lines were read
      and on how many sheets they arrive - or, in a WARNING, why column T is empty (the same reason is in T3 of every sheet).
    BOM materials WITHOUT a WSA sheet get the same calculation in memory -> Inbound sheet 'Other materials' + 'Below zero'.
    From today the planning columns are LIVE formulas on the last sheets, PPC_Plan_Req (module plan x BOM) and Archive:
      Balance = stock typed on today's row (yellow), then Balance - Actual + Delivered;
      Plant + Port Days = (Balance + open Scheduled Deliveries up to that day, not cancelled + U2 = open deliveries in the
      Archive) / Demand. When an archived delivery arrives, set its State to 'Delivered' on the Archive sheet.
    Rows before today keep their numbers. Delivered and the columns after it are copied as they are.
    Potting Material and Tapes (two materials side by side) become two sheets each. The Dashboard reads column S (19).
    Each planning sheet opens at today's row. Keep typing in YOUR WSA file; WSA_Daily is rebuilt from it on every run.
    To skip the rebuild (old behaviour):  python update_inbound.py --no-align
    To keep the older MRP_Daily files instead of deleting them:  python update_inbound.py --keep-old

WHAT YOU CAN CHANGE
  Everything is on the 'Settings' sheet (copy it into your MRP.xlsx once and edit it there). It is read FIRST, before
  anything is calculated, so the Inbound file and MRP_Daily always use the same rules. The CONFIG values below are only
  the defaults used when there is no Settings sheet yet.

STATUS RULES (see material_status())
  - Material status: compares days of cover with that material's OWN target days (A / B / C / packaging, + Z extra):
      CRITICAL = runs out within its target even with Open PO
      AT RISK  = runs out within its target unless the Open PO ships
      WATCH    = covered only if in transit arrives on time, or falls below target within its lead time
      OK       = everything else
  - Line cover: lines that can run fewer than COVER_CRITICAL days are flagged on Material Status (Settings).

Needs: Python 3.9+, pandas, openpyxl, lxml
"""
import sys, os, re, glob, time, math, datetime as dt, warnings
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
TARGET_DAYS     = 21     # target days for anything not classified A / B / C / packaging (the default target)
COVER_CRITICAL  = 14     # PRODUCTION LINES: CRITICAL (escalate) if a line can run fewer days than this
COVER_RISK      = 21     # LEGACY: read from old Settings sheets only, no longer used
COVER_WATCH     = 28     # LEGACY: read from old Settings sheets only, no longer used
SHORT_WATCH     = 14     # LEGACY: read from old Settings sheets only, no longer used
TARGET_A        = 21     # target days for A items (top 80% of value): lean, watched closely
TARGET_B        = 25     # target days for B items (next 15% of value)
TARGET_C        = 30     # target days for C items (last 5% of value, or no price): cheap to hold, never stop the line for them
TARGET_CSPM     = 10     # target days for packaging / consumables (local, short lead time)
EXTRA_Z_DAYS    = 5      # extra days for erratic demand (Z)
LEAD_TIME_RM    = 45     # days from order to plant for raw materials when the WSA file has no ETD history (used for planned order release)
LEAD_TIME_CSPM  = 14     # days from order to plant for packaging / consumables

# materials treated as UNLIMITED stock (never short, left out of shortage planning)
UNLIMITED_CODES = {'PMPKM00000338','PMPKM00000523'}   # material codes
UNLIMITED_WORDS = ('plastic bag',)                     # or any material whose description contains these words

# other codes of the SAME material in another pack / unit to count for an MRP material:  {code: (MRP code, factor)}
# factor = how many MRP units one unit of the code is.  Example:  PACK_FACTORS = {'RMBUS00000032': ('RMBUS00000047', 1)}
PACK_FACTORS    = {}
_PACK_FACTORS_SCRIPT=dict(PACK_FACTORS)    # kept: the Settings SUBSTITUTES (all lines) are added on top, they do not wipe these

PLAN_WSA_LINES_ONLY = False   # True = count only MIS plan rows of lines USA* with OEM Category WSA (old rule)
# EASIER: change these in the 'Settings' sheet of MRP_Daily (or copy that sheet into your MRP.xlsx) - it overrides the values above.
SETTINGS_SRC=[None]
ABC_A_PCT=80; ABC_B_PCT=95          # ABC cut-offs (% of value)
XYZ_X=0.25; XYZ_Y=0.5               # XYZ cut-offs (coefficient of variation of weekly demand)
ABC_HORIZON=90                      # days of requirement used for the ABC value
SHORT_HORIZON=60                    # days the shortage vs target is measured over
DEMAND_HORIZON=None                 # No Demand rule: days ahead in which a material must have demand (production requirement / BOM demand / daily target); None = the shortage horizon above
SHORT_LEAD=20; LONG_LEAD=75         # Inbound: days between ETD and plant date outside which a delivery is flagged
RELEASE_WINDOW=7                    # 'RELEASE ORDER' when an order is due within this many days
RESCHED_OUT_X=3                     # 'RESCHEDULE OUT' when stock is more than this many times the target with more arriving
SET_CUSTSUP=[]; CUSTSUP_ROWS=[]; SUBS_LINES=[]; SUBS_ISSUES=[]         # Customer Supplier / Vendor Priority rules read from the Settings sheet (if filled there)
LINE_RANK={}       # line key -> rank for all materials (Settings 'Line ranking', row 'All materials'); empty = closest to stopping first
LINE_RANK_MAT={}   # material (code / name / description as typed) -> {line key: rank}  (Settings 'Line ranking', one row per material)
ADDED_MATS=[]      # materials added because the plan needs them but the MIS Daily Dashboard does not list them: (code, name, module type)
TGT_CELL={}        # line key -> Settings cell of its daily target (Calc Target formulas point at it)
IT_STATUS_ONLY=False  # in transit = only MIS In Transit lines whose STATUS is 'IN-TRANSIT' (not at origin / at port / delivered / partly received ...)
def it_ok(r):
    return (not IT_STATUS_ONLY) or (len(r)>22 and it_status_group(r[22])=='In transit')
MIS_PATH_=[None]   # the MIS file of this run (batch split of glass stock)
ALLOC_IT={}        # (material, line) -> in-transit quantity from suppliers that line may use (Customer Supplier / Vendor Priority)
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
COVER_AVG_DAYS=21  # MRP Daily / Weekly 'Days of cover' = projected stock / average plan usage of the next N days - Settings
SAFETY_DAYS=7      # safety stock = days of need at the module plan (ALERT reorder level) - Settings
LOOKAHEAD_DAYS=49  # a line counts for a material (daily target need, 'Needed now') if it uses it within these days - Settings
AVG_PROD_WEEKS=9    # Averaged production (need per day) = average per production day (plan > 0) over the next N weeks - Settings (was 16 weeks)
REPORT_DAYS=21     # daily report: materials whose stock + in transit lasts fewer days than this at the module plan
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
       td = this material's own target days (Settings: A / B / C / packaging, + Z extra);
       below = first day projected stock falls under the target (day index) ; lead = lead time used for the material."""
    if cov[2]<td: return 'CRITICAL'
    if cov[1]<td: return 'AT RISK'
    if lead and cov[2]<lead: return 'AT RISK'       # runs out inside its supply lead time even with Open PO: only expediting can save it
    if cov[0]<td: return 'WATCH'
    if below is not None and below<=(lead or LEAD_TIME_RM)+14: return 'WATCH'
    return 'OK'

# ---------- read-only workbook cache: each input workbook / sheet is parsed once per run (big speed-up) ----------
def _ref_formula_counts(path):
    """{sheet name: number of formulas containing #REF!} read straight from the xlsx XML (no workbook is opened)."""
    import zipfile, html, posixpath
    out={}
    try:
        with zipfile.ZipFile(path) as z_:
            wbx=z_.read('xl/workbook.xml').decode('utf-8'); rel=z_.read('xl/_rels/workbook.xml.rels').decode('utf-8')
            tgt={}
            for m_ in re.finditer(r'<Relationship\b[^>]*>',rel):
                t_=m_.group(0); i_=re.search(r'\bId="([^"]+)"',t_); g_=re.search(r'\bTarget="([^"]+)"',t_)
                if i_ and g_: tgt[i_.group(1)]=g_.group(1)
            for m_ in re.finditer(r'<(?:\w+:)?sheet\b[^>]*>',wbx):
                t_=m_.group(0); nm_=re.search(r'\bname="([^"]*)"',t_); rid_=re.search(r'\br:id="([^"]+)"',t_)
                if not nm_ or not rid_ or rid_.group(1) not in tgt: continue
                pth_=tgt[rid_.group(1)].lstrip('/'); pth_=pth_ if pth_.startswith('xl/') else posixpath.join('xl',pth_)
                try: x_=z_.read(pth_).decode('utf-8','ignore')
                except KeyError: continue
                k_=len(re.findall(r'<f\b[^>]*>[^<]*#REF!',x_))
                if k_: out[html.unescape(nm_.group(1))]=k_
    except Exception as _e_: _swallow(0,_e_)
    return out
_WB_CACHE={}
class _CachedSheet:
    def __init__(self,rows,title): self._rows=rows; self.title=title; self.max_row=len(rows); self.max_column=max((len(r) for r in rows),default=0)
    def iter_rows(self,min_row=1,max_row=None,min_col=1,max_col=None,values_only=True):
        for r in self._rows[(min_row or 1)-1:max_row]:
            yield r[(min_col or 1)-1:max_col] if (max_col or (min_col or 1)>1) else r

USER_MB52=[None]   # your SAP MB52 export (same folder): the stock quantities; MIS MB52 only gives the extra columns (plant name, final code, key, company, categories)
def _is_mis_file(p):
    try: return {'MB52','Standard BOM'}<=set(_sheetnames(p))
    except Exception as _e_: _swallow(253,_e_); return False
def find_user_mb52(folders):
    """Newest MB52*.xlsx in the folders whose first sheet has Plant / Material / Unrestricted (not the MIS file)."""
    best=[]
    for fo in folders:
        for f in glob.glob(os.path.join(fo,'*.xlsx')):
            bn=os.path.basename(f)
            if bn.startswith('~$') or 'mb52' not in bn.lower() or 'dashboard' in bn.lower(): continue
            try:
                w=load_workbook(f,read_only=True,data_only=True); h=next(w[w.sheetnames[0]].iter_rows(max_row=1,values_only=True))
                hl=[str(x or '').strip().lower() for x in h]
                if {'plant','material','unrestricted'}<=set(hl): best.append((os.path.getmtime(f),f))
            except Exception as _e_: _swallow(265,_e_); continue
    return max(best)[1] if best else None
def user_mb52_rows(mis_rows, wbc):
    """Your MB52 in the MIS MB52 layout (20 columns). Your file gives Plant ... Quality Inspection (columns found by header);
       Plant Master.Plant Name / Company by plant and Final Material Code / Dashboard Key / Comman Name Cat.1 / Cat.2 by material
       come from MIS MB52 (else MIS Common Name / 'plant name-final code'); Total Stock = Unrestricted + Quality Inspection (MIS rule)."""
    hdr=list(mis_rows[0]); body=mis_rows[1:]
    pl_name={}; pl_comp={}; fin={}; key={}; cat={}
    for r in body:
        if not r or len(r)<20: continue
        pc=str(r[0] or '').strip().split('.')[0]
        if r[13]: pl_name.setdefault(pc,r[13])
        if r[17]: pl_comp.setdefault(pc,r[17])
        m=str(r[2] or '').strip()
        if m:
            if r[15]: fin.setdefault(m,str(r[15]).strip())
            if r[16]: key.setdefault(m,r[16])
            if r[18] or r[19]: cat.setdefault(m,(r[18],r[19]))
    try:
        for r in wbc['Common Name'].iter_rows(min_row=2,values_only=True):
            if r and r[0] and len(r)>4: cat.setdefault(str(r[0]).strip(),(r[3],r[4]))
    except Exception as _e_: _swallow(286,_e_); pass
    try:
        for r in wbc['Plant Master'].iter_rows(min_row=2,values_only=True):
            if r and r[0] and len(r)>1 and r[1]: pl_name.setdefault(str(r[0]).strip().split('.')[0],r[1])
    except Exception as _e_: _swallow(290,_e_); pass
    uw=load_workbook(USER_MB52[0],read_only=True,data_only=True); us=uw[uw.sheetnames[0]]
    urows=list(us.iter_rows(values_only=True)); uh=[str(x or '').strip().lower() for x in urows[0]]
    want=['plant','name 1','material','batch','storage location','material description','unrestricted','base unit of measure',
          'descr. of storage loc.','blocked','material group','material type','quality inspection']
    ix=[uh.index(w) if w in uh else None for w in want]
    num=lambda v: float(v) if isinstance(v,(int,float)) and not isinstance(v,bool) else (float(str(v).replace(',','')) if str(v or '').replace(',','').replace('.','',1).lstrip('-').isdigit() else 0.0)
    out=[tuple(hdr)]
    for r in urows[1:]:
        if not r or all(v in (None,'') for v in r): continue
        v=[(r[j] if j is not None and j<len(r) else None) for j in ix]
        if not v[2]: continue
        for k_ in (6,9,12): v[k_]=num(v[k_])
        pc=str(v[0] or '').strip().split('.')[0]; m=str(v[2]).strip()
        pn=pl_name.get(pc); f_=fin.get(m,m); c1,c2=cat.get(f_,cat.get(m,(None,None)))
        try: v[0]=int(pc)
        except ValueError: pass
        row=v+[pn,v[6]+v[12],f_,key.get(m) or (f"{pn}-{f_}" if pn else None),pl_comp.get(pc,pn),c1,c2]
        out.append(tuple(row+[None]*(len(hdr)-len(row))))
    print(f"Stock (MB52) read from your file: {os.path.basename(USER_MB52[0])} - {len(out)-1} lines (MIS MB52 used only for plant name, final code, key, company, categories)",flush=True)
    return out

ROWS_CACHE_DIR=os.environ.get('WSA_ROWS_CACHE')      # set only by the step-1 process for its step-2 child; removed when the run ends
def _rows_cache_file(path,name):
    if not ROWS_CACHE_DIR: return None
    import hashlib
    try: st_=os.stat(path)
    except OSError: return None
    k_=hashlib.md5(f'{os.path.abspath(path)}|{st_.st_size}|{st_.st_mtime_ns}|{name}'.encode('utf-8')).hexdigest()
    return os.path.join(ROWS_CACHE_DIR,k_+'.pkl')
def _rows_cache_get(path,name):
    f_=_rows_cache_file(path,name)
    if f_ and os.path.exists(f_):
        try:
            import pickle
            with open(f_,'rb') as h_: return pickle.load(h_)
        except Exception as _e_: _swallow(326,_e_); return None
    return None
def _rows_cache_put(path,name,rows):
    f_=_rows_cache_file(path,name)
    if f_:
        try:
            import pickle
            with open(f_+'.tmp','wb') as h_: pickle.dump(rows,h_,protocol=pickle.HIGHEST_PROTOCOL)
            os.replace(f_+'.tmp',f_)
        except Exception as _e_: _swallow(335,_e_); pass
class _CachedWB:
    def __init__(self,path):
        self._wb=load_workbook(path,read_only=True,data_only=True); self.sheetnames=list(self._wb.sheetnames); self._s={}; self._path=path
    def __getitem__(self,name):
        if name not in self._s:
            if name not in self.sheetnames:          # a sheet missing from an input file: warn once and carry on with no rows
                print(f"NOTE: sheet '{name}' not found in {os.path.basename(self._path)} - that part is skipped")
                self._s[name]=_CachedSheet([],name); return self._s[name]
            rows=_rows_cache_get(self._path,name)
            if rows is None:
                rows=[tuple(r) for r in self._wb[name].iter_rows(values_only=True)]; _rows_cache_put(self._path,name,rows)
            if name=='In Transit' and rows: rows=normalize_in_transit(rows,self)     # fixed column layout + checked material code
            if name=='MB52' and rows and USER_MB52[0] and _is_mis_file(self._path): rows=user_mb52_rows(rows,self)   # stock from YOUR MB52
            self._s[name]=_CachedSheet(rows,name)
        return self._s[name]
    def __contains__(self,name): return name in self.sheetnames
# ---------- MIS In Transit: column layout by header name, and the material code checked against the shipment description ----------
IT_CANON=['GID','Supplier Name','Invoice Number','Product Name','Description','HSN','PO','Qty','Unit Rate','Invoice value','BOOKING NO','LINE',
          'VESSEL / VOY','inco','OF','Cust','TPT','POL','POD',"40'HQ","20'GP",'Total Cont','STATUS','HBLNo.','MBL','Containers','ETD','ATD',
          'ETA as Per Booking','Revised ETA Port','warehouse','RM Code','Final Qty.','PPC GRN','Final Pending Qty.','Final AtA',
          'PO Wise Material Desc.','PO and Unit Price Key','True/False','Final RM Code','GR. Date']
IT_FIX=[]          # shipments whose MIS Final RM Code was replaced by the traced material (WSA PO + BL#, description, supplier; listed on Check)
IT_MISSING=[]      # In Transit headers not found in the MIS file this run (their column is read as blank - listed on Check)
_FRONT={'HTAR','ARC','FRONT'}; _BACK={'PRINTED','BACK'}
def _desc_keys(t):
    """Distinguishing features of a description: dimensions, thickness, HS frame height, roll length, front / back glass, GSM."""
    u=' '.join(str(t or '').upper().replace('\n',' ').split())
    u=re.sub(r'(\d+)\s+HS\b',r'\1HS',u)
    k={}
    k['hs']=set(re.findall(r'(\d+)HS',u))
    nums=set(re.findall(r'(?<![\d.])(\d{4})(?![\d.])',u)); k['dim']=nums
    th=set(re.findall(r'X\s*(\d+(?:\.\d+)?)\s*MM\b',u)) | set(re.findall(r'\b(\d+\.\d+)\s*MM\b',u))
    k['th']={str(float(x)) for x in th if float(x)<10}
    k['len']=set(re.findall(r'(\d+)\s*M\b(?!M)',u)) if 'TAPE' in u else set()     # roll length only tells tapes apart
    k['gsm']=set(re.findall(r'(\d+)\s*GSM',u))
    k['wd']=set()                                                       # cross-section of busbar / ribbon: '0.36*4 MM', '4 X 0.36MM', '5 X 0.36'
    for a_,b_ in re.findall(r'(\d+(?:\.\d+)?)\s*[X\*]\s*(\d+(?:\.\d+)?)\s*(?:MM|\b)',u):
        if float(a_)<20 and float(b_)<20: k['wd']|={str(float(a_)),str(float(b_))}
    w=set(re.split(r'[^A-Z0-9.]+',u))
    k['side']={'F'} if w&_FRONT else ({'B'} if w&_BACK else set())
    k['words']={x for x in w if len(x)>2 and not x.isdigit()}
    return k
def _desc_score(ship, cand):
    a,b=_desc_keys(ship),_desc_keys(cand); sc=0.0
    for f,wt in (('hs',3),('dim',2),('th',2),('len',2),('gsm',2),('side',3),('wd',3)):
        if a[f] and b[f]: sc+=wt*(len(a[f]&b[f])) - wt*(1 if not (a[f]&b[f]) else 0)*2
    sc+=0.2*len(a['words']&b['words'])
    return sc
def normalize_in_transit(rows, wbc):
    """1) Puts the In Transit columns in the usual order by their header names (MIS sometimes adds / removes a column).
       2) The material of each shipment is traced exactly like the PPC In Transit of the WSA / Inbound files (user rule, Nov 2026:
          'did you do the same with the MRP output'): your WSA sheet with this PO + BL# (HBLNo. / MBL) first, then description +
          supplier + RM Code against your BOM (map_in_transit). MIS 'Final RM Code' comes from the PO; it is replaced when another
          code clearly fits (e.g. EP304 0.50 mm on your EPE 420 sheet, 2272 glass on a 2376 PO, 35HS frame on a 30HS PO).
          Lines that are not BOM materials keep the MIS code. Every replacement is listed on the Check sheet (IT_FIX)."""
    hdr=[str(h).strip() if h is not None else '' for h in rows[0]]
    pos={h:i for i,h in enumerate(hdr) if h}
    IT_MISSING[:]=[h for h in IT_CANON if h not in pos]
    extra=[i for i,h in enumerate(hdr) if h and h not in IT_CANON]
    out=[tuple(IT_CANON+[hdr[i] for i in extra])]
    for r in rows[1:]:
        out.append(tuple([r[pos[h]] if (h in pos and pos[h]<len(r)) else None for h in IT_CANON]+[r[i] if i<len(r) else None for i in extra]))
    I_=IT_CANON.index; iR,iF,iD,iQ,iS,iPW,iPO,iHB,iMB=(I_(h) for h in ('RM Code','Final RM Code','Description','Final Pending Qty.','Supplier Name',
                                                                      'PO Wise Material Desc.','PO','HBLNo.','MBL'))
    up=lambda v: str(v or '').strip().upper()
    lines=[]
    for k,r in enumerate(out[1:],1):
        if not (r[iD] or r[iR] or r[iF]): continue
        lines.append(dict(k=k,final_code=up(r[iF]),rm_code=up(r[iR]),match='',po_codes=re.findall(r'\b(?:RM|PM|CS)[A-Z0-9]{6,}\b',up(r[iPW])),
                          description=' '.join(str(r[iD] or '').split()),qty=r[iQ] if isinstance(r[iQ],(int,float)) and not isinstance(r[iQ],bool) else 0,
                          po=r[iPO],bls=_bl_set(r[iHB])|_bl_set(r[iMB]),supplier=' '.join(str(r[iS] or '').split())))
    cat=in_transit_catalog(PPC_SRC[0],getattr(wbc,'_path',None) or MIS_PATH_[0])
    map_in_transit(lines,cat,dict(idx=wsa_ref_index(WSA_REFS[0]),alias={}) if WSA_REFS[0] else None)
    IT_FIX.clear(); fixed=list(out)
    for x in lines:
        c=x.get('code'); fc=str(out[x['k']][iF] or '').strip()
        if not c or c==x['final_code']: continue
        r=list(out[x['k']]); r[iF]=c; fixed[x['k']]=tuple(r)          # the traced material replaces the MIS code
        IT_FIX.append(dict(desc=x['description'][:80],shipped=c,mis=fc or '(blank)',qty=x['qty'],po=r[iPO],sup=r[iS],
                           cd_s=cat[c]['desc'],cd_m=(cat.get(up(fc)) or {}).get('desc',''),how=x.get('how') or ''))
    return fixed

PPC_SRC=[None]     # PPC workbook (sheets 'Count' + 'BOM' + 'Plan'): the ONLY source of physical stock counts; BOM and plan (else MRP.xlsx)
PP_HDR=('Date','Production Line','Wp','Customer','Qty','MW','Plant','Module Model')
def _cu_clean(x): return ' '.join(str(x or '').split())
def read_revised_plan(ws):
    """PPC 'Revised Plan': blocks side by side, one per production line (line name above the block, then Wp / Customer / Qty / MW),
       one row per date. A month block further down (new header row) is read the same way.
       Returns [(date, line, wp, customer, qty, mw)] for every cell with a Qty (0 included = line not running that day)."""
    rows=[list(r) for r in ws.iter_rows(values_only=True)]
    out=[]; blocks=None; dcol=None
    for i,r in enumerate(rows):
        hl=[' '.join(str(x or '').split()).lower() for x in r]
        if 'customer' in hl and 'qty' in hl:                      # header row: one block per 'Customer' column
            above=rows[i-1] if i>0 else []; blocks=[]
            for j,h in enumerate(hl):
                if h!='customer': continue
                q=next((k for k in range(j+1,len(hl)) if hl[k]=='qty'),None)
                if q is None: continue
                w=j-1 if j>0 and hl[j-1] in ('wp','watt') else None
                m=q+1 if q+1<len(hl) and hl[q+1]=='mw' else None
                st=w if w is not None else j; ln=None
                for k in range(st,-1,-1):                          # line name: the nearest text above the block, at or left of its start
                    v=above[k] if k<len(above) else None
                    if isinstance(v,str) and v.strip(): ln=_cu_clean(v); break
                    if blocks and k<=blocks[-1]['end']: break             # do not take the previous block's name
                blocks.append(dict(line=ln,wp=w,cu=j,qty=q,mw=m,end=(m if m is not None else q)))
            dcol=None; continue
        if not blocks: continue
        if dcol is None:
            dcol=next((k for k,v in enumerate(r) if isinstance(v,(dt.datetime,dt.date))),None)
        d=r[dcol] if dcol is not None and dcol<len(r) else None
        if not isinstance(d,(dt.datetime,dt.date)): continue
        for b in blocks:
            if not b['line']: continue
            q=r[b['qty']] if b['qty']<len(r) else None
            if not isinstance(q,(int,float)) or isinstance(q,bool): continue
            g=lambda k: (r[k] if k is not None and k<len(r) else None)
            out.append((pd.Timestamp(d).normalize(),b['line'],g(b['wp']),_cu_clean(g(b['cu'])),float(q),g(b['mw'])))
    return out
PLAN_NOTE=[None]
def ppc_plan_sheet(ppc):
    """Module plan = PPC 'Plan' with PPC 'Revised Plan' on top: for every date + production line in the Revised Plan, its rows
       replace the 'Plan' rows of that date + line; all other dates and lines stay as on 'Plan'.
       Revised Plan customer names are matched to the 'Plan' names (same line first) so the BOM and MIS links keep working.
       Returned in the Production Plan layout (Date, Production Line, Wp, Customer, Qty, MW, Plant, Module Model)."""
    import difflib
    hit=getattr(ppc,'_pp_merged',None)
    if hit is not None: return hit
    base=[r for r in pp_rows_std(ppc['Plan']) if r and isinstance(r[0],(dt.datetime,dt.date))]
    rev=read_revised_plan(ppc['Revised Plan']) if 'Revised Plan' in ppc.sheetnames else []
    names={}; models={}; plants={}
    for r in base:
        cu=_cu_clean(r[3])
        if not cu or not isinstance(r[4],(int,float)) or not r[4]: continue
        names.setdefault(cu,set()).add(_line_key(r[1]))
        mdl=str(r[7] or '').strip()
        if mdl.upper().startswith('BIN'): models.setdefault(cu,{}); models[cu][mdl]=models[cu].get(mdl,0)+r[4]
        if r[6]: plants.setdefault(_line_key(r[1]),r[6])
    ren={}
    def canon(cu,lk):
        if not cu: return cu
        n=cu.lower()
        for c in names:
            if c.lower()==n: return c
        best,sc=None,0.0
        for c in names:
            s=difflib.SequenceMatcher(None,n,c.lower()).ratio()
            if c.lower().startswith(n) or n.startswith(c.lower()): s=max(s,0.9)
            if lk in names[c]: s+=0.1
            if s>sc: best,sc=c,s
        if best and sc>=0.75: ren[cu]=best; return best
        return cu
    rkeys={(d,_line_key(ln)) for d,ln,*_ in rev}
    merged=[r for r in base if (pd.Timestamp(r[0]).normalize(),_line_key(r[1])) not in rkeys]
    for d,ln,wp,cu,q,mw in rev:
        c=canon(cu,_line_key(ln)); mdl=max(models[c].items(),key=lambda kv:kv[1])[0] if models.get(c) else None
        merged.append((d.to_pydatetime(),ln,wp,c,q,mw,plants.get(_line_key(ln),'USA'),mdl))
    merged.sort(key=lambda r:(pd.Timestamp(r[0]),str(r[1])))
    if rev:
        ds=sorted({d for d,_ in rkeys}); lns=sorted({ln for _,ln,*_ in rev})
        PLAN_NOTE[0]=(f"Module plan from PPC: 'Plan' + 'Revised Plan' ({ds[0]:%d-%b} to {ds[-1]:%d-%b}, {', '.join(lns)} replaced)"
                      +(' - Revised Plan names matched: '+'; '.join(f"'{a}' -> '{b}'" for a,b in sorted(ren.items())) if ren else ''))
    else: PLAN_NOTE[0]="Module plan from PPC 'Plan' (no 'Revised Plan' sheet)"
    print(PLAN_NOTE[0],flush=True)
    sh=_CachedSheet([PP_HDR]+[tuple(r) for r in merged],'Production Plan')
    try: ppc._pp_merged=sh
    except Exception as _e_: _swallow(503,_e_); pass
    return sh

class _RedirWB:
    """Your MRP.xlsx as the script sees it: 'BOM' and 'Stock Count' come from the PPC workbook ('BOM', 'Count'), all else from MRP.xlsx.
       When the PPC workbook has a 'Plan' sheet, 'Production Plan' = PPC 'Plan' + 'Revised Plan' (see ppc_plan_sheet)."""
    MAP={'BOM':'BOM','Stock Count':'Count'}
    def __init__(self,base,ppc): self._b=base; self._p=ppc
    def _has_plan(self): return 'Plan' in self._p.sheetnames
    @property
    def sheetnames(self): return list(dict.fromkeys(list(self._b.sheetnames)+list(self.MAP)+(['Production Plan'] if self._has_plan() else [])))
    def __contains__(self,n): return n in self.sheetnames
    def __getitem__(self,n):
        if n=='Production Plan' and self._has_plan(): return ppc_plan_sheet(self._p)
        return self._p[self.MAP[n]] if n in self.MAP else self._b[n]
    def __getattr__(self,a): return getattr(self._b,a)
def find_ppc(folders):
    """Newest workbook with 'Count' and 'BOM' sheets (a file name with 'PPC' wins), not an MRP_Daily / Inbound output."""
    best=[]
    for fo in folders:
        for f in glob.glob(os.path.join(fo,'*.xlsx')):
            bn=os.path.basename(f)
            if bn.startswith(('~$','MRP_Daily_','Inbound_')): continue
            try: n=set(_sheetnames_raw(f))
            except Exception as _e_: _swallow(527,_e_); continue
            if {'Count','BOM'}<=n: best.append(('ppc' in bn.lower(),os.path.getmtime(f),f))
    return max(best)[2] if best else None
def wb_ro(path):
    """wb_ro(path), parsed once and cached; re-read if the file changes. Your MRP.xlsx: BOM / Stock Count from the PPC workbook."""
    base=_wb_ro_raw(path)
    try:
        um_={os.path.abspath(p_) for p_ in (globals().get('USER_MRP'),(SETTINGS_SRC[0] if SETTINGS_SRC else None)) if p_}
        if PPC_SRC[0] and os.path.abspath(path) in um_:
            return _RedirWB(base,_wb_ro_raw(PPC_SRC[0]))
    except NameError: pass
    return base
def _wb_ro_raw(path):
    key=os.path.abspath(path)
    try: mt=os.path.getmtime(key)
    except OSError: mt=None
    hit=_WB_CACHE.get(key)
    if hit is None or hit[0]!=mt: hit=(mt,_CachedWB(path)); _WB_CACHE[key]=hit
    return hit[1]

RUN_NOTES=[]
class _CaptureDone(Exception): pass    # step 1 of the WSA alignment stops build_new_mrp after align_capture()
_QUIET=[False]                          # step 1 of the WSA alignment: only errors on the console (step 2 prints everything)
import builtins as _bi
def print(*a, **k):                     # every NOTE / WARNING / ERROR line is also kept (RUN_NOTES)
    m=' '.join(str(x) for x in a)
    if m.startswith(('NOTE','WARNING','ERROR')): RUN_NOTES.append(m)
    if _QUIET[0] and not m.lstrip().startswith('ERROR'): return
    _bi.print(*a, **k)
def die(m): print('ERROR:',m); sys.exit(1)
ARGS=sys.argv[1:]; PREV=None; MRP=None; MIS=None
NO_MRP='--no-mrp' in ARGS
if NO_MRP: ARGS.remove('--no-mrp')
KEEP_OLD='--keep-old' in ARGS                  # keep the older MRP_Daily_<date>.xlsx files (default: they are deleted once the new one is saved)
if KEEP_OLD: ARGS.remove('--keep-old')
NO_ALIGN='--no-align' in ARGS                   # default: the WSA planning columns are rebuilt (aligned copy of the WSA file) on every run
if NO_ALIGN: ARGS.remove('--no-align')
if '--align-wsa' in ARGS: ARGS.remove('--align-wsa')  # older command line: aligning is the default now
ALIGN_WSA=not NO_ALIGN
for flag in ('--prev','--mrp','--mrp-out','--mis'):
    if flag in ARGS:
        i=ARGS.index(flag); val=ARGS[i+1] if i+1<len(ARGS) else die(f'{flag} needs a file'); del ARGS[i:i+2]
        if flag=='--prev': PREV=val
        elif flag=='--mrp': MRP=val
        elif flag=='--mis': MIS=val
        else: MRP_OUT=val
def _release(path):
    """Close the cached read-only handles of a workbook (typed_rows / wb_ro) so the file can be deleted or replaced."""
    k=os.path.abspath(path)
    for p_ in [p_ for p_ in list(_PARTS) if os.path.abspath(p_)==k]:
        try: _PARTS.pop(p_)[0].close()
        except Exception as _e_: _swallow(576,_e_); pass
    hit=_WB_CACHE.pop(k,None)
    if hit is not None:
        try: hit[1]._wb.close()
        except Exception as _e_: _swallow(580,_e_); pass
    import gc; gc.collect()
def _sheetnames(p):
    try: return wb_ro(p).sheetnames
    except Exception as _e_: _swallow(584,_e_); return []
def _sheetnames_raw(p):
    """Sheet names straight from xl/workbook.xml (no workbook is opened) - used when scanning a folder for the right file."""
    import zipfile, html
    try:
        with zipfile.ZipFile(p) as z_: x_=z_.read('xl/workbook.xml').decode('utf-8')
        return [html.unescape(n_) for n_ in re.findall(r'<(?:\w+:)?sheet\b[^>]*?\bname="([^"]*)"',x_)]
    except Exception as _e_: _swallow(591,_e_); return []
_WSA_MARK={'Total_Plan_Req','PPC_Plan_Req'}     # a WSA daily file: the planner's (Total_Plan_Req) or an aligned one (PPC_Plan_Req)
if not ARGS:      # double-click / no file names: newest WSA daily file (Total_Plan_Req or PPC_Plan_Req sheet) in this folder or the script's folder
    _cands=[]
    for _f in dict.fromkeys([os.getcwd(),os.path.dirname(os.path.abspath(sys.argv[0]))]):
        for _n in os.listdir(_f):
            _nl=_n.lower()
            if _nl.endswith('.xlsx') and not _n.startswith('~$') and not _nl.endswith(('.sorting.xlsx','.tmp.keep')) and not _nl.startswith(('mrp','inbound_','wsa_daily_')) and 'mis' not in _nl and 'dashboard' not in _nl:
                _cands.append(os.path.join(_f,_n))
    _cands=[p for p in sorted(_cands,key=os.path.getmtime,reverse=True) if _WSA_MARK&set(_sheetnames_raw(p))]
    _own=[p for p in _cands if not re.search(r'\(aligned \d{4}-\d{2}-\d{2}\)',os.path.basename(p))]   # your own WSA file first (the script never writes it)
    if _own: _cands=_own+[p for p in _cands if p not in _own]
    if not _cands: die('No WSA daily file found here. Usage: python update_inbound.py <WSA daily file.xlsx> [MRP.xlsx] [MIS RM Dashboard.xlsx]')
    ARGS=[_cands[0]]; print(f'WSA daily file: {_cands[0]}')
def _is_mrp(p):
    n=_sheetnames(p)            # the MRP workbook has a By Week sheet; an Inbound output also has one, but it has an Inbound sheet too
    return 'By Week' in n and 'Inbound' not in n and not (_WSA_MARK&set(n))
def _is_mis(p):
    n=_sheetnames(p); return 'MB52' in n and ('Daily Dashboard' in n or 'Standard BOM' in n)
for a_ in list(ARGS):                     # an MIS RM Dashboard given in any position
    if MIS is None and a_.lower().endswith('.xlsx') and os.path.exists(a_) and _is_mis(a_): MIS=a_; ARGS.remove(a_)
if not ARGS: die('Usage: python update_inbound.py <daily_file.xlsx> [MRP.xlsx] [MIS_RM_Dashboard.xlsx] [output.xlsx] [--prev ...]')
if len(ARGS)>=2 and os.path.exists(ARGS[0]) and os.path.exists(ARGS[1]):      # accept the daily file and the MRP workbook in either order
    n0,n1=_sheetnames(ARGS[0]),_sheetnames(ARGS[1])
    if not (_WSA_MARK&set(n0)) and _is_mrp(ARGS[0]) and _WSA_MARK&set(n1): ARGS[0],ARGS[1]=ARGS[1],ARGS[0]
SRC=ARGS[0]
if not os.path.exists(SRC): die(f'File not found: {SRC}')
OUT=None
def _is_user_mrp(p):
    n=_sheetnames(p)                     # Nov 2026: the planner's MRP.xlsx may hold only Settings (plan + BOM + counts from the PPC file)
    return ('Settings' in n or ('Production Plan' in n and 'BOM' in n)) and not os.path.basename(p).startswith('MRP_Daily_')
USER_MRP=MRP if (MRP and os.path.exists(MRP)) else None     # your MRP workbook (BOM, Stock Count, Production Plan, Customer Supplier): read only
MRP=None                                                      # the old By Week / MRP_<date> update is retired - MRP_Daily is the output
if len(ARGS)>1:
    if os.path.exists(ARGS[1]) and USER_MRP is None and _is_user_mrp(ARGS[1]): USER_MRP=ARGS[1]
    else: OUT=ARGS[1]
if len(ARGS)>2 and OUT is None: OUT=ARGS[2]
OUTDIR=os.path.dirname(os.path.abspath(SRC))
def _find_user_mrp():
    """MRP.xlsx next to the daily file (or in the current folder); else the newest MRP*.xlsx with BOM + Production Plan sheets."""
    folders=[OUTDIR]+([os.path.abspath(os.getcwd())] if os.path.abspath(os.getcwd())!=OUTDIR else [])
    for f in folders:
        p=os.path.join(f,'MRP.xlsx')
        if os.path.exists(p) and _is_user_mrp(p): return p
    cands=[os.path.join(f,n) for f in folders for n in os.listdir(f)
           if n.lower().startswith('mrp') and n.lower().endswith('.xlsx') and not n.startswith('~$') and not n.startswith('MRP_Daily_')]
    for p in sorted(cands,key=os.path.getmtime,reverse=True):
        if _is_user_mrp(p): return p
    return None
if USER_MRP is None and not NO_MRP: USER_MRP=_find_user_mrp()
if USER_MRP: print(f'Your MRP workbook (read only): {USER_MRP}')
else: print('NOTE: no MRP.xlsx (with BOM and Production Plan sheets) found - MRP_Daily uses the MIS plan only.')
_PPC_DIRS=list(dict.fromkeys([os.path.dirname(os.path.abspath(USER_MRP)) if USER_MRP else os.getcwd(),OUTDIR,os.path.abspath(os.getcwd()),os.path.dirname(os.path.abspath(__file__))]))
PPC_SRC[0]=find_ppc(_PPC_DIRS)
if PPC_SRC[0]: print(f'BOM and Stock Count read from: {os.path.basename(PPC_SRC[0])} (sheets BOM, Count'+(", module plan from Plan + Revised Plan)" if 'Plan' in _sheetnames(PPC_SRC[0]) else ')'))
else: print('WARNING: no PPC workbook (sheets Count + BOM) in the folder - physical stock counts come ONLY from the PPC Count sheet, so this run uses MB52 stock for every material; BOM is read from MRP.xlsx')
USER_MB52[0]=find_user_mb52(list(dict.fromkeys([os.path.dirname(os.path.abspath(USER_MRP)) if USER_MRP else os.getcwd(),OUTDIR,os.path.abspath(os.getcwd()),os.path.dirname(os.path.abspath(__file__))])))
if not USER_MB52[0]: print('NOTE: no MB52.xlsx (your SAP export) found in the folder - stock is read from the MIS MB52 sheet')
_BASE_D0=[None]
def base_d0():
    """The run's day 1 (user rule, Nov 2026: "If you have MB52 Nov 09, then calc everything from Nov 09; the stock count is also
       considered counted on Nov 09"): the date of your MB52.xlsx when it is older than today, else today. Projections, the WSA
       daily's Balance row, the ALERT / dashboards and late shipments all start on this day; a PPC count dated on / after it counts as
       taken on it (no roll-forward), an older count is rolled forward to it. File names keep the run date."""
    if _BASE_D0[0] is None:
        t_=pd.Timestamp(dt.date.today()).normalize(); b_=t_
        try:
            if USER_MB52[0]:
                m_=file_data_date(USER_MB52[0])
                if m_ is not None and m_<t_: b_=pd.Timestamp(m_).normalize()
        except Exception as _e_: _swallow(0,_e_)
        _BASE_D0[0]=b_
        if b_<t_ and not _QUIET[0]: print(f"BASE DATE {b_:%d-%b-%Y}: your MB52.xlsx is dated {b_:%d-%b} - everything is calculated from {b_:%d-%b} (stock, counts, plan, in transit); file names keep today's date")
    return _BASE_D0[0]
try: import lxml.etree
except ImportError: die("This script needs one more add-on. Run this once, then run the script again:\n       pip install lxml")
if MIS is None:   # newest MIS*Dashboard*.xlsx next to the daily file (or in the current folder)
    cands=[]
    for f in {OUTDIR,os.path.abspath(os.getcwd())}:
        for n in os.listdir(f):
            nl=n.lower()
            if nl.endswith('.xlsx') and not n.startswith('~$') and ('mis' in nl or 'dashboard' in nl): cands.append(os.path.join(f,n))
    for p in sorted(cands,key=os.path.getmtime,reverse=True):
        if _is_mis(p): MIS=p; print(f'MIS RM Dashboard found: {p}'); break
if MIS and not os.path.exists(MIS): die(f'MIS RM Dashboard not found: {MIS}')
if not MIS: print('NOTE: no MIS RM Dashboard found (needs MB52 and Standard BOM sheets) - MRP_Daily is not built, only the Inbound file.')
OUT=OUT or os.path.join(OUTDIR,f'Inbound_{dt.date.today():%Y-%m-%d}.xlsx')

CODE=re.compile(r'^(JW)?RM[A-Z]{3,5}\d+$')
LABELS={('Tapes',0):'Tapes - Cell Fixation',('Tapes',1):'Tapes - Perforated 500m',('Potting Material',0):'Potting Material - Part A',('Potting Material',1):'Potting Material - Part B'}
SKIP={'Total_Plan_Req','Stock'}
def _num(x): return float(x) if isinstance(x,(int,float)) and not isinstance(x,bool) else None
def cv(rows,r,c):
    if r<1 or c<1 or r>len(rows): return None
    row=rows[r-1]; return row[c-1] if c<=len(row) else None

_PARTS={}
def typed_rows(path, sheet, col):
    """Rows whose cell in column 'col' holds a typed number (not a formula), read straight from the sheet XML - much faster than
       loading the whole workbook a second time just to see which balances are typed."""
    import zipfile
    from lxml import etree
    M='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    if path not in _PARTS:                 # sheet name -> xml part, read once per file
        z=zipfile.ZipFile(path); wbx=etree.fromstring(z.read('xl/workbook.xml')); rels=etree.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        tg={r.get('Id'):r.get('Target') for r in rels}
        _PARTS[path]=(z,{sh.get('name'):tg.get(sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')) for sh in wbx.find(M+'sheets')})
    z,pm=_PARTS[path]; tgt=pm.get(sheet)
    if tgt is None: return set()
    part=tgt.lstrip('/') if tgt.startswith('/') else 'xl/'+tgt
    letter=L(col); out=set()
    for _,c in etree.iterparse(z.open(part),tag=M+'c'):
        ref=c.get('r') or ''
        if ref.rstrip('0123456789')==letter and c.find(M+'f') is None and c.find(M+'v') is not None and c.get('t') in (None,'n'):
            out.add(int(ref[len(letter):]))
        c.clear()
    return out

def extract_fast(path, plan_end):
    print(f'  [{time.time()-_T0:5.0f} s] reading the WSA planning file ...',flush=True)
    wv=load_workbook(path,data_only=True,read_only=True)
    groups=[]; skipped=[]
    for sh in wv.sheetnames:
        if sh in SKIP or sh==ARCHIVE_SHEET: continue
        ws=wv[sh]
        if ws.sheet_state!='visible': continue
        head=[list(r) for r in ws.iter_rows(min_row=1,max_row=12,values_only=True)]
        starts=[(r+1,c+1) for r,row in enumerate(head) for c,x in enumerate(row[:200]) if str(x).strip().startswith('Required Plant Date')]
        if not starts: continue
        rv=[list(r) for r in ws.iter_rows(values_only=True)]
        maxc=max(len(r) for r in rv)
        starts.sort(key=lambda x:x[1])
        for bi,(hr,c0) in enumerate(starts):
            c1=(starts[bi+1][1]-1) if bi+1<len(starts) else maxc
            hdr={}
            for c in range(c0,c1+1):
                x=str(cv(rv,hr,c) or '').replace('\n',' ').strip().lower()
                if x: hdr.setdefault(x,c)
            def col(*ps):
                for p in ps:
                    for x,c in hdr.items():
                        if x.startswith(p): return c
                return None
            FIELDS=[('balance',lambda h:h.startswith('balance')),('planned',lambda h:h.startswith('planned modules per day')),('demand',lambda h:h.startswith('demand')),
                    ('actmod',lambda h:h.startswith('actual modules per day')),('actual',lambda h:h=='actual' or h.startswith('actual consum')),('mw',lambda h:h.startswith('mw per day')),
                    ('actmw',lambda h:h.startswith('actual mw per day')),('plantdays',lambda h:h.startswith('plant days')),('plantport',lambda h:h.startswith('plant + port') and '(ppc' not in h),('delivered',lambda h:h.startswith('delivered')),
                    ('ppcport',lambda h:h.startswith('plant + port days (ppc')),('ppcit',lambda h:h.startswith('ppc in transit') and 'supplier' not in h),
                    ('vlock',lambda h:h.startswith('vendor rule: not usable'))]
            fcols={k:next((c for h,c in hdr.items() if fn(h)),None) for k,fn in FIELDS}
            def info(r): return {k:(_num(cv(rv,r,c)) if c else None) for k,c in fcols.items()}
            cB=col('balance'); cDl=col('delivered'); cS=col('scheduled'); cPO=col('po'); cSh=col('shipper'); cRem=col('remarks'); cBk=col('booking approv'); cRd=col('readiness'); cETD=col('etd'); cBL=col('bl'); cVs=col('vessel')
            if cB is None or cS is None: skipped.append(f'{sh} (no Balance / Scheduled Delivery column)'); continue
            codes=set(); uom=''; mname=''
            for r in range(1,hr):
                for c in range(c0,c1+1):
                    x=cv(rv,r,c)
                    if isinstance(x,str):
                        if CODE.match(x.strip()): codes.add(x.strip())
                        if x.strip().startswith('Unit of Measurement'): uom=str(cv(rv,r,c+2) or '').strip()
                        if x.strip().startswith('Supplier & Material'): mname=str(cv(rv,r,c+2) or '').strip()
            rows=[(r,cv(rv,r,c0)) for r in range(hr+1,len(rv)+1) if isinstance(cv(rv,r,c0),dt.datetime)]
            if not rows: skipped.append(f'{sh} (no dated rows)'); continue
            last=None; typed=typed_rows(path,sh,cB)          # balances typed by the planner (not formulas) = real stock dates
            for r,d in rows:
                if r in typed and isinstance(cv(rv,r,cB),(int,float)) and not isinstance(cv(rv,r,cB),bool): last=(r,d)
            if last is None: skipped.append(f'{sh} (no typed balance, so no stock date)'); continue
            asof=last[1]; inbound=[]; last_rec=None
            bal_by_date={d:_num(cv(rv,r,cB)) for r,d in rows if _num(cv(rv,r,cB)) is not None}
            dem_by_date={d:info(r).get('demand') for r,d in rows if isinstance(info(r).get('demand'),(int,float))}
            pdays_by_date={d:(info(r).get('plantdays'),info(r).get('plantport'),info(r).get('ppcport')) for r,d in rows}
            ppcit_by_date={d:info(r).get('ppcit') for r,d in rows if info(r).get('ppcit')}
            vlock_by_date={d:info(r).get('vlock') for r,d in rows if info(r).get('vlock')}     # Settings ALLOWED VENDORS (aligned file, column AK)
            cU_=next((c for h,c in hdr.items() if h.startswith('ppc in transit (supplier')),None)
            ppcsup_names={' '.join(n_.split()) for r,d in rows if cU_ and d>=asof for n_ in str(cv(rv,r,cU_) or '').split(' / ') if n_.strip()}
            for r,d in rows:
                if cDl:
                    qd=_num(cv(rv,r,cDl))
                    if qd and qd>0 and d<=asof: last_rec=(d,qd)
                if d>plan_end: continue
                q=_num(cv(rv,r,cS))
                if q and q>0 and d>=asof-dt.timedelta(days=14):
                    g=lambda c: cv(rv,r,c) if c else None
                    inbound.append(dict(date=d,qty=q,row=r,remarks=str(g(cRem) or '').strip(),po=g(cPO),shipper=g(cSh),booking=g(cBk),readiness=g(cRd),etd=g(cETD),bl=g(cBL),vessel=str(g(cVs) or '').replace('\n',' ').strip(),info=info(r)))
            label=LABELS.get((sh,bi)) or (sh if len(starts)==1 else f'{sh} - {mname[:18]}')
            refs=[(cv(rv,r,cPO) if cPO else None,cv(rv,r,cBL) if cBL else None) for r,d in rows]     # PO + BL# of every row: In_Transit match
            groups.append(dict(label=label,sheet=sh,block=bi,uom=uom,codes=sorted(codes),asof=asof,inbound=inbound,last_receipt=last_rec,asof_info=info(last[0]),bal_by_date=bal_by_date,dem_by_date=dem_by_date,pdays_by_date=pdays_by_date,ppcit_by_date=ppcit_by_date,vlock_by_date=vlock_by_date,ppcsup_names=ppcsup_names,
                               refs=[t for t in refs if any(v not in (None,'') for v in t)]))
    if ARCHIVE_SHEET in wv.sheetnames:          # aligned WSA file: the delivery records of the deleted rows (their PO + BL# count for the match too)
        it_=wv[ARCHIVE_SHEET].iter_rows(values_only=True); h_=[_hnorm(v) for v in next(it_,())]
        iS_,iP_,iB_=(h_.index(k) if k in h_ else None for k in ('planning sheet','po','bl#'))
        byl_={g['label']:g for g in groups}
        v_=lambda r,i: r[i] if (i is not None and i<len(r)) else None
        if iS_ is not None:
            for r in it_:
                g=byl_.get(v_(r,iS_))
                if g and (v_(r,iP_) not in (None,'') or v_(r,iB_) not in (None,'')): g['refs'].append((v_(r,iP_),v_(r,iB_)))
    return groups,skipped

def _sheet_pools(sheets):
    """BW['sheets'] with the sheets that share a material merged (back glass 333 on 'Glass G12R Topcon SB Back' + 'Glass G12R Topcon
       Subanci Back' = one MIS material): balance = total of the sheets (None if one has none), demand = total per date."""
    labs=list(sheets); par={l:l for l in labs}
    def root(l):
        while par[l]!=l: par[l]=par[par[l]]; l=par[l]
        return l
    owner={}
    for l in labs:
        for m in sheets[l].get('members') or []:
            if m in owner: par[root(l)]=root(owner[m])
            else: owner[m]=l
    grp={}
    for l in labs: grp.setdefault(root(l),[]).append(l)
    out={}
    for ls in grp.values():
        if len(ls)==1: out[ls[0]]=sheets[ls[0]]; continue
        shs=[sheets[l] for l in ls]; dem={}
        for sh in shs:
            for d,v in (sh.get('dem') or {}).items(): dem[d]=dem.get(d,0.0)+(v or 0.0)
        out[' + '.join(ls)]=dict(shs[0],members=sorted({m for sh in shs for m in sh.get('members') or []}),dem=dem,n=len(shs),
                                bal=sum(float(sh['bal']) for sh in shs) if all(sh.get('bal') is not None for sh in shs) else None,
                                sheet=' + '.join(sh['sheet'] for sh in shs),asof=max(sh['asof'] for sh in shs))
    return out
def below_zero_periods(g):
    """From the sheet stock date: the periods (from, to) where the sheet's Plant Days < 0 and where its Plant + Port Days < 0.
       A day with no demand ('na') keeps the state of the day before (Plant Days: the sign of the Balance).
       'pi' = Plant + Port Days (PPC In Transit) - the column with the arrivals of your PPC In_Transit sheet.
       Returns {'pd': [...], 'pp': [...], 'pi': [...]} with dict(fr, to, low, open) - open = still below zero on the last date of the sheet."""
    asof=g['asof']; pdd=g.get('pdays_by_date') or {}; bal=g.get('bal_by_date') or {}
    ds=sorted(d for d in (set(pdd)|set(bal)) if d>=asof)
    res={}
    for key,ix in (('pd',0),('pp',1),('pi',2)):
        per=[]; cur=None; prev=False
        for d in ds:
            t_=pdd.get(d) or ()
            v=t_[ix] if ix<len(t_) else None
            neg=((bal.get(d) is not None and bal[d]<0) if ix==0 else prev) if v is None else v<0
            if neg:
                if cur is None: cur=dict(fr=d,to=d,low=v,open=False)
                else:
                    cur['to']=d
                    if v is not None and (cur['low'] is None or v<cur['low']): cur['low']=v
            elif cur is not None: per.append(cur); cur=None
            prev=neg
        if cur is not None: cur['open']=True; per.append(cur)
        res[key]=per
    return res
def periods_text(per, most=4, end='end of sheet'):
    """'08-Nov-26 -> 11-Nov-26; 20-Nov-26 -> 31-Mar-27 (end of sheet)' - or 'none'."""
    if not per: return 'none'
    t_='; '.join(f"{p['fr']:%d-%b-%y} \u2192 {p['to']:%d-%b-%y}"+(f' ({end})' if p['open'] else '') for p in per[:most])
    return t_+(f' ... +{len(per)-most} more (Below zero sheet)' if len(per)>most else '')
def other_materials_calc(AD, covered):
    """The WSA-sheet calculation for every MRP material WITHOUT a WSA planning sheet (nothing is written to the WSA file):
       Balance(day) = stock today - demand of the days before; Plant Days = Balance / Demand;
       Plant + Port Days (PPC In Transit) = (Balance + arrivals of the PPC In_Transit sheet up to that day) / Demand.
       Stock = the stock MRP_Daily uses (MB52 / PPC count, JW at the OEM left out), Demand = module plan x BOM (OEM plan left out)."""
    dates=AD['dates']; nd=len(dates); arr=(AD.get('ppc') or {}).get('arr') or {}; out=[]
    for m,b in AD['blocks'].items():
        if m in covered: continue
        plan=b['plan']; st=float(b['stock'] or 0); a=arr.get(m) or [0.0]*nd
        bal=st; cum=0.0; pdd={}; bal_by={}
        for i,d in enumerate(dates):
            if i: bal-=plan[i-1]
            cum+=a[i]; dem=plan[i]
            pdd[d]=((bal/dem) if dem else None,None,((bal+cum)/dem) if dem else None); bal_by[d]=bal
        per=below_zero_periods(dict(asof=dates[0],pdays_by_date=pdd,bal_by_date=bal_by))
        nxt=next(((dates[i],a[i]) for i in range(nd) if a[i]>0),None)
        out.append(dict(mat=m,common=b.get('common') or m,uom=b.get('uom') or '',kind='Raw material' if (str(m).upper().startswith('RM') and not str(m).upper().startswith('RMLAB')) else 'Packaging / consumable',   # barcode / back label = packaging (user, Nov 2026)
                        stock=st,dem7=sum(plan[:7]),dem30=sum(plan[:30]),pd0=pdd[dates[0]][0],pi0=pdd[dates[0]][2],per=per,it=sum(a),nxt=nxt,
                        ppc_by={dates[i]:a[i] for i in range(nd) if a[i]>0},asof=dates[0]))
    first=lambda x: min([p['fr'] for p in x['per']['pi']]+[p['fr'] for p in x['per']['pd']]+[pd.Timestamp('2100-01-01')])
    return sorted(out,key=lambda x:(first(x),str(x['common'])))

def add_other_materials(path, rows, ppc, wsa_of=None):
    """Inbound: sheet 'Other materials' (one row per material), their periods added to 'Below zero', and sheet 'PPC In Transit'
       (every line of your PPC In_Transit sheet: what it says, the BOM material it was traced to and why, where it is counted)."""
    from openpyxl.styles import Border, Side
    wb=load_workbook(path); AR='Calibri'
    bf=Font(name=AR,size=10); hf=Font(name=AR,bold=True,color='FFFFFF',size=10)
    hfill=PatternFill('solid',fgColor='1F3864'); ctr=Alignment(horizontal='center',vertical='center',wrap_text=True); dfy='dd-mmm-yy'
    if 'Other materials' in wb.sheetnames: del wb['Other materials']
    pos=(wb.sheetnames.index('Below zero')+1) if 'Below zero' in wb.sheetnames else 1
    ws=wb.create_sheet('Other materials',pos); ws.sheet_view.showGridLines=False
    ws['A1']='Other materials - the WSA calculation for the BOM materials that have no WSA planning sheet (calculated by the script, not written to the WSA file)'
    ws['A1'].font=Font(name=AR,size=12,bold=True)
    ws['A2']=('Same rules as the WSA sheets: Balance = stock today (MB52 / PPC count, JW at the OEM left out) minus the demand of the days before; '
              'Demand = module plan x BOM (OEM plan left out); Plant Days = Balance / Demand; Plant + Port Days (PPC In Transit) = (Balance + arrivals of your '
              f"PPC In_Transit sheet up to that day) / Demand{(' - ' + ppc['src']) if ppc and ppc.get('src') else ' - no PPC In_Transit sheet found'}. "
              'Solar cells are not in the MRP, so they are not listed.'); ws['A2'].font=bf
    hd=['Material','Code','Type','UOM','Stock today','Demand next 7 days','Demand next 30 days','Plant Days today','Plant + Port Days (PPC In Transit) today',
        'Plant Days < 0 (from \u2192 to)','Plant + Port Days (PPC In Transit) < 0 (from \u2192 to)','PPC In Transit qty (all)','Next PPC arrival','Next PPC arrival qty']
    for j,h in enumerate(hd,1):
        c=ws.cell(4,j,h); c.font=hf; c.fill=hfill; c.alignment=ctr
    ws.row_dimensions[4].height=44
    for i,x in enumerate(rows):
        r=5+i
        vals=[x['common'],x['mat'],x['kind'],x['uom'],x['stock'],x['dem7'],x['dem30'],x['pd0'] if x['pd0'] is not None else 'no demand today',
              x['pi0'] if x['pi0'] is not None else 'no demand today',periods_text(x['per']['pd'],end='end of plan'),periods_text(x['per']['pi'],end='end of plan'),x['it'] or None,
              x['nxt'][0] if x['nxt'] else None,x['nxt'][1] if x['nxt'] else None]
        for j,v in enumerate(vals,1):
            c=ws.cell(r,j,v); c.font=bf
            if j in (5,6,7,12,14): c.number_format='#,##0'
            elif j in (8,9): c.number_format='0.0'
            elif j==13: c.number_format=dfy
        for j,k_ in ((10,'pd'),(11,'pi')):
            ws.cell(r,j).font=Font(name=AR,size=10,color='C00000' if x['per'][k_] else '008000'); ws.cell(r,j).alignment=Alignment(wrap_text=True,vertical='top')
    for c_,w_ in zip('ABCDEFGHIJKLMN',[40,17,20,7,13,13,13,11,14,40,40,13,12,12]): ws.column_dimensions[c_].width=w_
    ws.freeze_panes='B5'; ws.auto_filter.ref=f'A4:N{4+max(1,len(rows))}'
    if 'Below zero' in wb.sheetnames:                                  # their periods go to the Below zero list too
        bz=wb['Below zero']; old=[]
        for r in bz.iter_rows(min_row=5,max_col=13,values_only=True):
            if r and r[0] and r[4]: old.append(tuple(r))
        new=[]
        for x in rows:
            for key_,meas_ in (('pd','Plant Days < 0'),('pi','Plant + Port Days (PPC In Transit) < 0')):
                for p_ in x['per'][key_]:
                    pa=sorted((d_,q_) for d_,q_ in x['ppc_by'].items() if d_>=p_['fr'])
                    new.append(('(no WSA sheet)',f"{x['common']} ({x['mat']})",x['uom'],x['asof'].to_pydatetime(),meas_,p_['fr'].to_pydatetime(),p_['to'].to_pydatetime(),
                                (p_['to']-p_['fr']).days+1,p_['low'],'Yes' if p_['open'] else '',pa[0][0].to_pydatetime() if pa else None,pa[0][1] if pa else None,'PPC In Transit' if pa else ''))
        allr=sorted(old+new,key=lambda t_:(pd.Timestamp(t_[5]),str(t_[4]),str(t_[0])))
        for row_ in bz.iter_rows(min_row=5,max_row=max(bz.max_row,5)):
            for c in row_: c.value=None
        for i,t_ in enumerate(allr):
            for j,v in enumerate(t_,1):
                c=bz.cell(5+i,j,v); c.font=bf
                if j in (4,6,7,11): c.number_format=dfy
                elif j==9: c.number_format='0.0'
                elif j in (8,12): c.number_format='#,##0'
            if str(t_[4]).startswith('Plant + Port'): bz.cell(5+i,5).font=Font(name=AR,size=10,bold=True,color='C55A11' if 'PPC' in str(t_[4]) else 'C00000')
        bz.auto_filter.ref=f'A4:M{4+max(1,len(allr))}'
    lines=(ppc or {}).get('lines') or []; why_='; '.join((ppc or {}).get('why') or [])
    if lines:                                                          # the trace of every In_Transit line
        if 'PPC In Transit' in wb.sheetnames: del wb['PPC In Transit']
        tr=wb.create_sheet('PPC In Transit',wb.sheetnames.index('Other materials')+1); tr.sheet_view.showGridLines=False
        tr['A1']="PPC In Transit - every line of your PPC In_Transit sheet traced to YOUR BOM material (supplier + description + RM Code + Final RM Code)"
        tr['A1'].font=Font(name=AR,size=12,bold=True)
        tr['A2']=('Your WSA first: the sheet where you typed this PO + BL# (HBLNo. / MBL; else the BL#, else the PO) gives the material when it fits the '
                  'description. Else the Final RM Code is replaced only when the description (size, thickness, HS, GSM, front / back, product type) and the '
                  'supplier clearly fit another code - usually the RM Code. Amber = traced to another code than Final RM Code; red = CHECK (your WSA sheet '
                  'has the shipment as another size / HS) or not a BOM material (not counted). Late lines are counted on today.'); tr['A2'].font=bf
        if why_: tr['A3']='NOTE: '+why_; tr['A3'].font=Font(name=AR,size=10,bold=True,color='C00000')
        hd2=['Supplier','Invoice','PO','Description (In_Transit)','RM Code','Final RM Code','Traced to BOM code','BOM material description','BOM supplier',
             'Description from','How it was traced','Pending qty','MRP code (after code conversion)','Qty in MRP unit','Plant arrival (Final AtA)','Counted on','Shown on',
             'Your WSA sheet with this PO / BL#']
        for j,h in enumerate(hd2,1):
            c=tr.cell(4,j,h); c.font=hf; c.fill=hfill; c.alignment=ctr
        tr.row_dimensions[4].height=44
        amber=PatternFill('solid',fgColor='FFE699')
        wsa_of=wsa_of or {}
        def shown(x):
            if not x.get('code'): return '-'
            ls_=wsa_of.get(x.get('mat'))
            if isinstance(ls_,str): ls_=[ls_]
            if ls_ and len(ls_)>1:                                     # a material on 2+ sheets: the sheet with this PO / BL#, else by supplier
                t_=[l_ for l_ in ls_ if l_ in set(x.get('wsa_sheets') or [])]
                return f"WSA sheet: {t_[0]}" if len(t_)==1 else 'WSA sheet: '+' or '.join(ls_)+' (by supplier - Settings ALLOWED VENDORS)'
            if ls_: return f"WSA sheet: {ls_[0]}"
            if x.get('in_mrp'): return 'Other materials'
            return 'not in the MRP (no plan for it)'
        order_=lambda x:(0 if (x.get('code') and x['code']!=x['final_code']) else (1 if 'CHECK' in str(x.get('how')) else (2 if not x.get('code') else 3)),str(x['supplier']),pd.Timestamp(x['date']) if x.get('date') is not None else pd.Timestamp('2100-01-01'))
        for i,x in enumerate(sorted(lines,key=order_)):
            r=5+i
            vals=[x['supplier'],x['invoice'],x['po'],x['description'],x['rm_code'] or None,x['raw_code'] or None,x.get('code'),x.get('bom_desc') or None,x.get('bom_sup') or None,
                  x.get('src') or None,x.get('how'),x['qty'],x.get('mat'),x.get('qty_mrp'),x['date'].to_pydatetime() if x.get('date') is not None else None,
                  (x['arrive'].to_pydatetime() if x.get('arrive') is not None else None),shown(x),
                  (f"{', '.join(x['wsa_sheets'])} ({x['wsa_how']})" if x.get('wsa_sheets') else '-')]
            for j,v in enumerate(vals,1):
                c=tr.cell(r,j,v); c.font=bf
                if j in (12,14): c.number_format='#,##0.##'
                elif j in (15,16): c.number_format=dfy
            if x.get('code') and x['code']!=x['final_code']:
                for j in (7,11): tr.cell(r,j).fill=amber
            if not x.get('code') or 'CHECK' in str(x.get('how')): tr.cell(r,11).font=Font(name=AR,size=10,color='C00000',bold='CHECK' in str(x.get('how')))
            if x.get('late'): tr.cell(r,16).font=Font(name=AR,size=10,color='C55A11')
        for c_,w_ in zip('ABCDEFGHIJKLMNOPQR',[30,13,12,46,15,17,15,36,26,10,48,11,15,12,12,12,30,34]): tr.column_dimensions[c_].width=w_
        tr.freeze_panes='E5'; tr.auto_filter.ref=f'A4:R{4+len(lines)}'
    else:                                                              # nothing read: say why (also in T3 of every WSA sheet)
        if 'PPC In Transit' in wb.sheetnames: del wb['PPC In Transit']
        tr=wb.create_sheet('PPC In Transit',(wb.sheetnames.index('Other materials')+1) if 'Other materials' in wb.sheetnames else len(wb.sheetnames))
        tr.sheet_view.showGridLines=False
        tr['A1']="PPC In Transit - NOTHING READ: column T 'PPC In Transit (arrivals)' of the WSA sheets is empty"; tr['A1'].font=Font(name=AR,size=12,bold=True,color='C00000')
        tr['A3']='Why: '+(why_ or 'no PPC file with an In_Transit sheet'); tr['A3'].font=bf
    if 'Summary' in wb.sheetnames:
        c=wb['Summary']['D4']; c.value=f"Other BOM materials (no WSA sheet): {len(rows)} on the sheet 'Other materials'"; c.font=Font(name=AR,size=10,bold=True,color='C55A11')
    if 'Read me' in wb.sheetnames:
        rd_=wb['Read me']; r_=rd_.max_row+2
        rd_.cell(r_,1,"Other materials - the WSA calculation (Balance, Plant Days, Plant + Port Days with your PPC In_Transit) for the BOM materials without a WSA planning sheet; their below-zero periods are on the Below zero sheet.").font=bf
        rd_.cell(r_+1,1,"PPC In Transit - every line of your PPC In_Transit sheet, the BOM material it was traced to (your WSA PO + BL#, description, supplier, RM Code, Final RM Code), why, where it is counted, and your WSA sheet with that PO / BL#.").font=bf
    wb.save(path)

def clean_status(rem):
    l=rem.lower()
    if 'canc' in l: return 'Cancelled'
    if 'grn' in l: return 'GRN pending'
    if 'approved' in l: return 'Approved for booking'
    if 'planned' in l: return 'Planned'
    return ''
def port_of(rem):

    if clean_status(rem): return ''
    return rem.title() if rem else ''
def ymd(v): return pd.Timestamp(v).date() if hasattr(v,'year') else None


# ===================== Daily Stock sheet (from the MIS RM Dashboard) =====================
# Plan scope: every MIS Total Plan Req row of the plant (Company = USA). Set PLAN_WSA_LINES_ONLY=True to count only
# production lines USA* with OEM Category WSA (OEM lines Imperial Star, Runergy, Inox and Arizona excluded).
PLAN_OEM='WSA'
def plan_line_ok(line, oem, company, plant='USA'):
    if str(company or '').strip().upper()!=plant.upper(): return False
    if not PLAN_WSA_LINES_ONLY: return True
    ln=str(line or '').strip().upper()
    return str(oem or '').strip().upper()==PLAN_OEM and (ln.startswith('USA') or re.fullmatch(r'L\d+',ln) is not None)
def tp_ok(r, plant='USA'):
    """A row of MIS Total Plan Req counts in the plan (columns: 1 Plant / line, 18 OEM Category, 21 Company)."""
    return len(r)>21 and plan_line_ok(r[1],r[18],r[21],plant)

TAPE_TYPES=('cell fixation','perforated','heating','double side','kapton','insulation')

def tape_info(desc):
    """Tape family and metres per roll read from the description, e.g. 'G2G PERFORATED TAPE 30MM X 500M' -> ('perforated',30,500),
       'CELL FIXATION TAPE-8MM*100M' -> ('cell fixation',8,100), 'CELL FIXATION TAPE 8MM FF-3665 (T50)' -> ('cell fixation',8,50)."""
    t=' '.join(str(desc or '').upper().split())
    typ=next((k for k in TAPE_TYPES if k.upper() in t),None)
    if not typ or 'TAPE' not in t: return None
    w=re.search(r'(\d+(?:\.\d+)?)\s*MM',t)
    m=re.search(r'(?<![\d.])(\d+(?:\.\d+)?)\s*M(?![A-Z])',t) or re.search(r'\bT(\d+)\b',t)
    if not (w and m): return None
    return (typ,float(w.group(1)),float(m.group(1)))


# ===================== MIS without 'Total Plan Req' / 'Daily Dashboard': rebuilt in memory =====================
MIS_REBUILT=[]     # which MIS sheets were rebuilt this run (shown on the console and on Check)
def _prev_links(prev_path):
    """Customer -> (MIS customer, MIS BOM Name) from the MIS LINK table on the Settings sheet of the previous MRP_Daily."""
    out={}
    if not prev_path or not os.path.exists(prev_path): return out
    try:
        on=False
        for r in wb_ro(prev_path)['Settings'].iter_rows(values_only=True):
            a=str((r or [None])[0] or '').strip()
            if a.upper().startswith('MIS LINK'): on=True; continue
            if on and a and a.isupper() and len(a)>8: break
            if on and a and not a.lower().startswith('production plan customer') and len(r)>2:
                bn=(r[4] if len(r)>4 and r[4] not in (None,'') else r[2]); mc=(r[5] if len(r)>5 and r[5] not in (None,'') else r[1])
                if bn: out[_norm(a)]=(str(mc).strip() if mc else a,str(bn).strip())
    except Exception as _e_: _swallow(1015,_e_); pass
    return out

def prepare_mis(mis_path, mrp_path, prev_path, ndays=141):
    """Your MIS file may have no 'Total Plan Req' and / or no 'Daily Dashboard'. Then they are rebuilt in memory (the file is not changed):
       Total Plan Req = your Production Plan (MRP.xlsx) modules x MIS Standard BOM quantity per module (layers listed twice add up);
       Daily Dashboard = one material per code of the Standard BOM of the module types in your plan (cells left out), named from Common Name.
       Customer -> MIS BOM Name: Settings MIS LINK override, else the previous MRP_Daily's MIS LINK table, else the Module Model of a linked customer."""
    MIS_REBUILT.clear()
    w=wb_ro(mis_path); need_tp='Total Plan Req' not in w.sheetnames; need_dd='Daily Dashboard' not in w.sheetnames
    if not (need_tp or need_dd): return
    if not mrp_path: die("The MIS file has no 'Total Plan Req' / 'Daily Dashboard' - then MRP.xlsx (Production Plan) is needed to build the raw-material plan")
    # plant = MB52 Company of the plant rows
    import collections
    co=collections.Counter(str(r[17]).strip() for r in w['MB52'].iter_rows(min_row=2,values_only=True) if r and len(r)>17 and r[17] and 'WAREHOUSE' not in str(r[17]).upper())
    plant=co.most_common(1)[0][0] if co else 'USA'
    # Standard BOM: module type -> [(sr, code, desc, qty, uom, group)]
    sb={}
    for r in w['Standard BOM'].iter_rows(min_row=2,values_only=True):
        if r and len(r)>8 and r[2] and r[3] and isinstance(r[5],(int,float)):
            sb.setdefault(str(r[2]).strip(),[]).append((r[1],str(r[3]).strip(),r[4],float(r[5]),r[6],str(r[8] or '').strip()))
    cn={}
    for r in w['Common Name'].iter_rows(min_row=2,values_only=True):
        if r and r[0]: cn.setdefault(str(r[0]).strip(),r)
    # your Production Plan
    d0=base_d0(); dend=pd.Timestamp(dt.date.today()).normalize()+pd.Timedelta(days=ndays)   # base date earlier: window starts earlier, same end
    pp=[]; models={}
    for r in pp_rows_std(wb_ro(mrp_path)['Production Plan']):
        if not r or not isinstance(r[0],(dt.datetime,dt.date)) or not isinstance(r[4],(int,float)) or not r[4]: continue
        d=pd.Timestamp(r[0]).normalize()
        if d<d0-pd.Timedelta(days=3) or d>=dend: continue
        ln=' '.join(str(r[1] or '').split()); cu=' '.join(str(r[3] or '').split()); mdl=str(r[7] or '').strip() if len(r)>7 else ''
        pp.append((d,ln,cu,float(r[4]),mdl)); models.setdefault(cu,collections.Counter())[mdl]+=r[4]
    prev=_prev_links(prev_path); link={}
    for cu in models:
        ov=link_override(cu)
        if ov and ov[1]: link[cu]=(ov[0] or cu,ov[1]); continue
        if _norm(cu) in prev: link[cu]=prev[_norm(cu)]
    for cu in models:                                # same Module Model as a linked customer
        if cu in link: continue
        mdl=models[cu].most_common(1)[0][0]
        for c2,(mc2,bn2) in list(link.items()):
            if mdl and mdl in models.get(c2,{}): link[cu]=(cu,bn2); break
    missing_=[cu for cu in models if cu not in link or link[cu][1] not in sb]
    if models and len(missing_)==len(models):
        die("MIS has no 'Total Plan Req', so each Production Plan customer needs its MIS BOM Name. None was found (no Settings MIS LINK, no previous MRP_Daily).\n"
            "       Add the MIS LINK table to the Settings sheet of MRP.xlsx (Production Plan customer | MIS customer | MIS BOM Name | How | Override: MIS BOM Name | Override: MIS customer).")
    if need_tp:
        rows=[tuple(['Date','Plant','Customer','Value','BOM Name','JB Type','Watt','Identify Name','Total MW','Standard BOM.Sr. no.','Standard BOM.Material Code',
                     'Standard BOM.Material Description','Standard BOM.BOM','Standard BOM.UOM','Total Requirement','Common Name.Common Name','Month-Year','OA No.',
                     'OEM Category','Match','Dashboard Key','Company','Final Week','Quarter','Match2','Comman Name Cat.1','Comman Name Cat.2'])]
        for d,ln,cu,q,mdl in pp:
            if cu not in link or link[cu][1] not in sb: continue
            mc,bn=link[cu]
            for sr,code,desc,bq,uom,grp in sb[bn]:
                c_=cn.get(code) or [None]*7
                rows.append((d.to_pydatetime(),ln,mc,q,bn,None,None,None,None,sr,code,desc,bq,uom,q*bq,c_[3] if len(c_)>3 else None,None,f'{ln}|{cu}',
                             PLAN_OEM,code,None,plant,None,None,code,c_[4] if len(c_)>4 else None,c_[5] if len(c_)>5 else None))
        w._s['Total Plan Req']=_CachedSheet(rows,'Total Plan Req'); w.sheetnames.append('Total Plan Req')
        MIS_REBUILT.append(f"Total Plan Req ({len(rows)-1} rows = your Production Plan x MIS Standard BOM)")
    if need_dd:
        used=[]; seen_codes=set(); seen_names=set()
        for bn in sorted({link[cu][1] for cu in link if link[cu][1] in sb},key=lambda b_:-sum(q for d,ln,cu,q,m in pp if link.get(cu,(None,None))[1]==b_)):
            for sr,code,desc,bq,uom,grp in sb[bn]:
                if grp.lower() in ('cells','cell') or code in seen_codes: continue
                c_=cn.get(code) or [None]*7; nm=str(c_[3] or desc or code).strip()
                if nm.lower() in seen_names: continue      # same Common Name = alternate code of a material already listed
                seen_codes.add(code); seen_names.add(nm.lower()); used.append((f'{sr}{bn}',sr,grp or (c_[5] if len(c_)>5 else ''),nm,code,desc,bn))
        mts=[]
        for u in used:
            if u[6] not in mts: mts.append(u[6])
        rows=[('Plant',plant),tuple(['Module Type']+mts[:4]),(),tuple(['Key No.2','Sr No.','Short Name','Common Name','Material','Material Description','Headers','Stock Data']+[f'{d0+pd.Timedelta(days=k):%d-%m-%Y}' for k in range(ndays)])]
        for key,sr,short,nm,code,desc,bn in used: rows.append((key,sr,short,nm,code,desc,'Plan'))
        rows.append(())
        w._s['Daily Dashboard']=_CachedSheet(rows,'Daily Dashboard'); w.sheetnames.append('Daily Dashboard')
        MIS_REBUILT.append(f"Daily Dashboard ({len(used)} materials from the MIS Standard BOM of {len(mts)} module types in your plan)")
    print('MIS sheets rebuilt in memory (not in the file): '+'; '.join(MIS_REBUILT))
    if missing_: print('WARNING: no MIS BOM Name for '+', '.join(missing_)+" - their modules get no raw-material demand. Set it on Settings > MIS LINK (Override: MIS BOM Name).")


def read_phys_counts(mrp_path):
    """Stock Count sheet of MRP.xlsx: material -> dict(base=(qty, unit), bom=(qty, unit), date). Counted = 'Physical Count Qty'
       (column C) is a number; a blank C is 'no counting' (MIS is used), whatever the Qty_BOM formula shows."""
    out={}
    if not PPC_SRC[0]: return out                                       # physical counts come ONLY from the PPC workbook ('Count')
    try: rows=[list(r) for r in wb_ro(mrp_path)['Stock Count'].iter_rows(values_only=True)]
    except Exception as _e_: _swallow(1101,_e_); return out
    hi=next((i for i,r in enumerate(rows[:10]) if r and str(r[0] or '').strip().lower()=='material'),None)
    if hi is None: return out
    hl=[' '.join(str(x or '').split()).lower() for x in rows[hi]]
    ix=lambda *ks: next((hl.index(k) for k in ks if k in hl),None)
    ic=ix('physical count qty','today stock qty'); iu=ix('base unit of measure','stock uom'); ib=ix('bom uom')
    iq=ix('physical count qty_bom','today qty_bom'); idt=ix('physical count date','count date'); isc=ix('supplier code','physical count supplier')
    dcols=[]                                                             # date-header columns: one column per count day
    for j_,hv in enumerate(rows[hi]):
        dd_=hv if isinstance(hv,(dt.datetime,dt.date)) else None
        if dd_ is None and isinstance(hv,str) and re.match(r'^\s*\d{1,2}/\d{1,2}/\d{4}\s*$',hv):
            try: dd_=dt.datetime.strptime(hv.strip(),'%m/%d/%Y')
            except ValueError: dd_=None
        if dd_: dcols.append((j_,pd.Timestamp(dd_).normalize()))
    g=lambda r,i: r[i] if i is not None and i<len(r) else None
    for r in rows[hi+1:]:
        if not r or not r[0]: continue
        q=g(r,ic)
        if not isinstance(q,(int,float)) or isinstance(q,bool): continue          # no counting -> MIS
        c=str(r[0]).strip(); o=out.setdefault(c,dict(base=[0.0,str(g(r,iu) or '').strip().upper()],bom=[0.0,str(g(r,ib) or '').strip().upper()],date=None))
        o['base'][0]+=float(q)
        qb=g(r,iq); o['bom'][0]+=float(qb) if isinstance(qb,(int,float)) and not isinstance(qb,bool) else float(q)
        d=g(r,idt)
        if not isinstance(d,(dt.datetime,dt.date)):                       # latest date column with a count in it
            ds_=[dd_ for j_,dd_ in dcols if j_<len(r) and isinstance(r[j_],(int,float)) and not isinstance(r[j_],bool)]
            d=max(ds_) if ds_ else None
        if isinstance(d,(dt.datetime,dt.date)) and (o['date'] is None or pd.Timestamp(d)>pd.Timestamp(o['date'])): o['date']=pd.Timestamp(d).normalize()
        sc=str(g(r,isc) or '').strip().upper()                       # counted supplier (KB / FL ...)
        if sc:
            qb2=float(qb) if isinstance(qb,(int,float)) and not isinstance(qb,bool) else float(q)
            e=o.setdefault('by_sup',{}).setdefault(sc,[0.0,0.0]); e[0]+=float(q); e[1]+=qb2
    return out


PP_COLS=[('date',('date',)),('line',('production line','line')),('wp',('wp','watt')),('cust',('customer',)),('qty',('qty','quantity','modules')),
         ('mw',('mw',)),('plant',('plant',)),('model',('module model','model'))]
def pp_rows_std(ws):
    """Production Plan rows as (Date, Production Line, Wp, Customer, Qty, MW, Plant, Module Model), found by COLUMN HEADER -
       so an added column (e.g. 'Month') does not shift anything. Without a recognisable header: the old fixed layout A..H."""
    it=ws.iter_rows(values_only=True); hdr=next(it,None) or ()
    hl=[' '.join(str(x or '').split()).lower() for x in hdr]
    idx=[]
    for key,names in PP_COLS:
        j=next((hl.index(n) for n in names if n in hl),None)
        idx.append(j)
    if idx[0] is None or idx[3] is None or idx[4] is None: idx=list(range(8))      # old layout
    for r in it:
        if not r: continue
        yield tuple((r[j] if j is not None and j<len(r) else None) for j in idx)

def build_daily_stock(mis_path, start=None, ndays=None):
    """Same logic as the MIS 'Daily Dashboard', computed from the MIS sheets' saved values:
       Plan      = SUMIFS('Total Plan Req'!O, K=material, V=plant, A=date)
       Stock     = MB52 Total Stock for <plant> and <plant> WAREHOUSE + 'RM Status for OEM JW' JW WSA to OEM (column L)
       Intransit = SUMIFS('In Transit'!AI, AN=material, AJ=date)
       Adhoc     = SUMIFS(Adhoc!D, B=material, E=date)"""
    print(f'  [{time.time()-_T0:5.0f} s] reading MIS (stock, plan, in transit) ...',flush=True)
    wb=wb_ro(mis_path)
    def df(name):
        rows=list(wb[name].iter_rows(values_only=True))
        return pd.DataFrame(rows[1:],columns=[str(c).strip() if c is not None else f'_c{i}' for i,c in enumerate(rows[0])])
    def bycol(d,i): return d[d.columns[i]]
    dd=[list(r) for r in wb['Daily Dashboard'].iter_rows(values_only=True)]
    plant=str(dd[0][1] or 'USA').strip()
    modtypes=[x for x in dd[1][1:5] if isinstance(x,str) and x.strip()]
    ncols_dash=sum(1 for x in dd[3][8:] if x not in (None,''))
    blocks=[]
    for r in range(4,len(dd)-1):
        row=dd[r]
        if len(row)>6 and str(row[6] or '').strip()=='Plan' and isinstance(row[4],str) and row[4].strip():
            sr=row[1]; key=str(row[0] or ''); mt=key[len(str(sr)):] if key.startswith(str(sr)) else ''
            blocks.append(dict(key=key,sr=sr,short=row[2],common=row[3],mat=row[4].strip(),desc=row[5],mtype=mt))
    if not blocks: die("No materials found on the MIS 'Daily Dashboard' sheet (rows with Headers = Plan)")
    start=pd.Timestamp(start or base_d0()).normalize(); ndays=(ndays or ncols_dash or 141)+max(0,(pd.Timestamp(dt.date.today()).normalize()-start).days)   # base date earlier: same last day
    dates=[start+pd.Timedelta(days=k) for k in range(ndays)]; idx={d:k for k,d in enumerate(dates)}
    mb=df('MB52'); rs=df('RM Status for OEM JW'); tp=df('Total Plan Req'); it=df('In Transit')
    ad=pd.DataFrame(columns=range(8))       # MIS 'Adhoc' sheet is NOT used (no Adhoc receipts, no Adhoc stock conversions)
    num=lambda s: pd.to_numeric(s,errors='coerce').fillna(0)
    day=lambda s: pd.to_datetime(s,errors='coerce').dt.normalize()
    # MATERIALS MISSING FROM THE MIS DAILY DASHBOARD: every raw-material code that the plant's plan needs (MIS Total Plan Req,
    # today onward) is planned, even when the Daily Dashboard does not list it (e.g. the Sabanci 3.2 frame RMALF00000244 and
    # JB RMJUB00000123 of 500 TC G2G_G12R_3.2). A code whose MIS Common Name equals a dashboard material stays an alternate code of it.
    cn_={}
    if 'Common Name' in wb.sheetnames:
        for r_ in wb['Common Name'].iter_rows(min_row=2,values_only=True):
            if r_ and r_[0]: cn_.setdefault(str(r_[0]).strip(),r_)
    known_={b['mat'] for b in blocks}; dash_names_={str(b.get('common') or '').strip().lower() for b in blocks}; extra_={}
    for d_,ln_,oem_,co_,bn_,c_,ds_,q_ in zip(day(bycol(tp,0)),bycol(tp,1),bycol(tp,18),bycol(tp,21),bycol(tp,4),bycol(tp,10),bycol(tp,11),num(bycol(tp,14))):
        c_=str(c_ or '').strip()
        if not CODE.match(c_) or c_.upper().startswith('JW') or c_ in known_ or q_<=0 or pd.isna(d_) or d_<start or not plan_line_ok(ln_,oem_,co_,plant): continue
        nm_=str((cn_.get(c_) or [None]*4)[3] or '').strip()
        if nm_ and nm_.lower() in dash_names_: continue
        x_=extra_.setdefault(c_,dict(desc=ds_,bn={},name=nm_)); k_=str(bn_ or '').strip(); x_['bn'][k_]=x_['bn'].get(k_,0.0)+q_
    # raw materials of YOUR BOM (MRP.xlsx) that the MIS list does not have (e.g. labels): planned too
    ubm_=load_user_bom_rm(SETTINGS_SRC[0]) if SETTINGS_SRC and SETTINGS_SRC[0] else {}
    for c_,mt_ in (ubm_.items() if UBOM_DEMAND else []):     # (only when your BOM drives demand)
        if not CODE.match(c_) or c_ in known_ or c_ in extra_: continue
        nm_=str((cn_.get(c_) or [None]*4)[3] or '').strip()
        if nm_ and nm_.lower() in dash_names_: continue           # alternate code of a listed material
        ti_=tape_info(mt_.get('desc'))
        if ti_ and any((tape_info(b_.get('desc')) or ())[:2]==ti_[:2] for b_ in blocks): continue   # another roll length of a listed tape
        extra_[c_]=dict(desc=mt_.get('desc'),bn={str(mt_.get('mtype') or 'Your BOM'):1.0},name=nm_ or str(mt_.get('desc') or c_))
    ADDED_MATS.clear()
    for c_,x_ in sorted(extra_.items()):
        mt_=max(x_['bn'],key=x_['bn'].get); r_=cn_.get(c_) or [None]*6
        sr_=1+max([int(b['sr']) for b in blocks if b['mtype']==mt_ and isinstance(b['sr'],(int,float))] or [0])
        blocks.append(dict(key=f'{sr_}{mt_}',sr=sr_,short=(r_[5] if len(r_)>5 and r_[5] else 'Raw material'),common=x_['name'] or str(x_['desc'] or c_),
                           mat=c_,desc=x_['desc'] or (r_[1] if len(r_)>1 else ''),mtype=mt_))
        ADDED_MATS.append((c_,blocks[-1]['common'],mt_))
    if ADDED_MATS: print('Added materials the plan needs but the MIS Daily Dashboard does not list: '+'; '.join(f'{n} ({c})' for c,n,_ in ADDED_MATS))
    uom_all_={str(k).strip():(r_[2] if len(r_)>2 else None) for k,r_ in cn_.items()}
    for c_,u_ in zip(bycol(mb,2),bycol(mb,7)):
        if c_ is not None and u_: uom_all_.setdefault(str(c_).strip(),u_)
    # Stock (SUMIFS text matches ignore case)
    pl=bycol(mb,13).astype(str).str.strip().str.upper()
    mbk=mb[pl.isin([plant.upper(),plant.upper()+' WAREHOUSE'])]
    stock=num(bycol(mbk,14)).groupby(bycol(mbk,2).astype(str).str.strip()).sum().to_dict()
    jw=num(bycol(rs,11)).groupby(bycol(rs,1).astype(str).str.strip()).sum().to_dict()
    raw_stock=dict(stock); raw_jw=dict(jw)          # untouched copies, for the 'raw stock (own unit)' text
    # Alternate codes: other material codes with the same MIS Common Name as a dashboard material (e.g. Cell Fixation Tape T50:
    # dashboard RMTTT00000010 = 50 m rolls, stock and shipments on RMTTT00000018 = 100 m rolls). MIS converts their STOCK through
    # the Adhoc sheet (here 2,343.9 x 2 = 4,687.8), which becomes opening stock below. Their IN-TRANSIT is added to the dashboard
    # material, converted with the same ratio (Adhoc qty / alternate-code stock) when it can be worked out.
    dash_by_name={str(b['common']).strip().lower():b['mat'] for b in blocks if b.get('common')}
    alias={}
    def _al(code,name):
        c_=str(code or '').strip(); n_=str(name or '').strip().lower()
        if c_ and n_ in dash_by_name and c_!=dash_by_name[n_] and not c_.upper().startswith('JW'): alias.setdefault(c_,dash_by_name[n_])
    # ONLY RULE: the MIS 'Common Name' sheet (A = Material Code, D = Common Name). A code is linked to a dashboard material when its
    # Common Name there equals the dashboard material's Common Name. To link another code, add / correct its row on that sheet.
    if 'Common Name' in wb.sheetnames:
        for r_ in list(wb['Common Name'].iter_rows(min_row=2,values_only=True)):
            if r_ and len(r_)>3: _al(r_[0],r_[3])
    # TAPES: each roll length has its own code. Codes of the same tape (type + width) are linked from their DESCRIPTION and converted
    # by roll length: e.g. 1 roll of 200 m = 0.4 roll of 500 m; 1 roll of 100 m = 2 rolls of 50 m. Their stock is counted directly
    # (MIS's own Adhoc conversion for these tapes is then dropped, so nothing is counted twice), in transit is converted the same way.
    desc_of={}
    if 'Common Name' in wb.sheetnames:
        for r_ in wb['Common Name'].iter_rows(min_row=2,values_only=True):
            if r_ and r_[0] and r_[1]: desc_of.setdefault(str(r_[0]).strip(),str(r_[1]))
    for c_,d_ in zip(bycol(mb,2),bycol(mb,5)):
        if c_ is not None and d_: desc_of.setdefault(str(c_).strip(),str(d_))
    for c_,d_ in zip(bycol(it,39),bycol(it,4)):
        if c_ is not None and d_: desc_of.setdefault(str(c_).strip(),str(d_))
    tape_factor={}; tape_dash={}; tape_len={}
    for b in blocks:
        ti=tape_info(b.get('desc') or desc_of.get(b['mat']))
        if not ti: continue
        tape_dash[b['mat']]=ti; tape_len[b['mat']]=ti[2]
        for c_,dsc in desc_of.items():
            tc=tape_info(dsc)
            if c_==b['mat'] or c_.upper().startswith('JW') or not tc or tc[:2]!=ti[:2]: continue
            alias[c_]=b['mat']; tape_factor[c_]=tc[2]/ti[2]; tape_len[c_]=tc[2]
    pack_window={}                                    # dated pack conversions: counted only between 'use from' and 'use until'
    for c_,t_ in PACK_FACTORS.items():               # your own pack conversions (Settings sheet): counted like tape rolls
        d_,f_=t_[0],t_[1]; fr=t_[2] if len(t_)>2 else None; to=t_[3] if len(t_)>3 else None
        if d_ not in {b['mat'] for b in blocks} or not f_: continue
        if fr is None and to is None: alias[c_]=d_; tape_factor[c_]=float(f_); tape_dash.setdefault(d_,('pack',0,1.0)); tape_len.setdefault(d_,1.0); tape_len[c_]=float(f_)
        else: pack_window.setdefault(d_,[]).append(dict(code=c_,factor=float(f_),fr=fr,to=to))
    # YOUR mapping (Material Group sheet, MRP.xlsx): 'MRP material' + 'Factor' - overrides the automatic matching above
    load_user_masters(SETTINGS_SRC[0] if SETTINGS_SRC and SETTINGS_SRC[0] else None)
    bset_={b['mat'] for b in blocks}; USER_ALT_MISS.clear()
    for c_,(d_,f_) in USER_ALT.items():
        if c_==d_:                                        # you say: this code is its own material - undo an automatic link
            if c_ in alias: alias.pop(c_,None); tape_factor.pop(c_,None)
            continue
        if d_ not in bset_: USER_ALT_MISS.append((c_,d_)); continue
        alias[c_]=d_; tape_factor[c_]=f_; tape_dash.setdefault(d_,('pack',0,1.0)); tape_len.setdefault(d_,1.0); tape_len[c_]=tape_len.get(d_,1.0)*f_
    uom_raw={}
    for c_,u_ in zip(bycol(mb,2),bycol(mb,7)):
        if c_ is not None and isinstance(u_,str) and u_.strip(): uom_raw.setdefault(str(c_).strip(),u_.strip())

    # conversion ratio alternate code -> dashboard code, from MIS's own Adhoc conversion
    adq={}
    for c_,q_ in zip(bycol(ad,1),num(bycol(ad,3))):
        if c_ is not None: adq[str(c_).strip()]=adq.get(str(c_).strip(),0.0)+q_
    alias_factor=dict(tape_factor); conv_qty={}       # alternate code -> ratio to the dashboard unit ; dashboard code -> Adhoc qty that is a conversion
    nm_of={}
    for c_,n_ in zip(bycol(mb,2),bycol(mb,18)):
        if c_ is not None and n_: nm_of.setdefault(str(c_).strip(),str(n_))
    first=lambda t: (re.split(r'[^a-z0-9]+',str(t or '').lower()) or [''])[0]
    RATIOS=(0.2,0.25,0.4,0.5,1.0,2.0,2.5,4.0,5.0)
    conv_unlinked={}                    # Adhoc conversion of a code that the Common Name sheet does not link -> Check only
    for d_r,a_r,t_r in zip(bycol(ad,1),num(bycol(ad,3)),day(bycol(ad,4))):
        d_=str(d_r or '').strip()
        if d_ not in {b['mat'] for b in blocks} or a_r<=0 or pd.isna(t_r) or t_r>start or d_ in tape_dash: continue
        dname=next((b['common'] for b in blocks if b['mat']==d_),'')
        hit=False
        for c_,s_ in stock.items():                      # 1) codes linked by the Common Name sheet: ratio from MIS's Adhoc conversion
            if alias.get(c_)!=d_ or s_<=0: continue
            r_=a_r/s_
            if any(abs(r_-x)<=1e-4*x for x in RATIOS):
                alias_factor[c_]=round(r_,6); conv_qty[d_]=conv_qty.get(d_,0.0)+a_r; hit=True; break
        if hit: continue
        for c_,s_ in stock.items():                      # 2) looks like a conversion of an unlinked code: report, do not count
            if c_==d_ or c_ in alias or s_<=0 or c_.upper().startswith('JW') or first(nm_of.get(c_,''))!=first(dname): continue
            r_=a_r/s_
            if any(abs(r_-x)<=1e-4*x for x in RATIOS): conv_unlinked.setdefault(d_,[]).append((c_,nm_of.get(c_,''),round(r_,6),a_r)); break
    # stock by location and unit of measure (for the Inventory sheet)
    loc=pd.DataFrame({'m':bycol(mbk,2).astype(str).str.strip(),'w':pl[pl.isin([plant.upper(),plant.upper()+' WAREHOUSE'])].str.endswith('WAREHOUSE'),'q':num(bycol(mbk,14))})
    st_main=loc[~loc['w']].groupby('m')['q'].sum().to_dict(); st_wh=loc[loc['w']].groupby('m')['q'].sum().to_dict()
    uom={}
    for m_,u_ in zip(bycol(mb,2).astype(str).str.strip(),bycol(mb,7)):
        if isinstance(u_,str) and u_.strip(): uom.setdefault(m_,u_.strip())
    # codes with no MB52 row (no stock yet, e.g. RMTTT00000010, RMGLS00000356): unit from the MIS Common Name sheet, else Total Plan Req
    if 'Common Name' in wb.sheetnames:
        for r_ in wb['Common Name'].iter_rows(min_row=2,values_only=True):
            if r_ and len(r_)>2 and r_[0] and isinstance(r_[2],str) and r_[2].strip(): uom.setdefault(str(r_[0]).strip(),r_[2].strip())
    for m_,u_ in zip(bycol(tp,10).astype(str).str.strip(),bycol(tp,13)):
        if isinstance(u_,str) and u_.strip(): uom.setdefault(m_,u_.strip())
    tpmask=pd.Series([plan_line_ok(a,b,c,plant) for a,b,c in zip(bycol(tp,1),bycol(tp,18),bycol(tp,21))],index=tp.index)   # USA* lines, OEM Category WSA
    tpf=pd.DataFrame({'m':bycol(tp,10).astype(str).str.strip(),'d':day(bycol(tp,0)),'q':num(bycol(tp,14))})[tpmask.values]
    plan_all={}
    for (m_,d_),q_ in tpf.dropna(subset=['d']).groupby(['m','d'])['q'].sum().items(): plan_all.setdefault(m_,{})[d_]=q_
    def daily(frame,qty_i,mat_i,date_i,extra=None,clamp=False,remap=None):
        f=pd.DataFrame({'m':bycol(frame,mat_i).astype(str).str.strip(),'d':day(bycol(frame,date_i)),'q':num(bycol(frame,qty_i))})
        if remap: f['m']=f['m'].map(lambda x: remap.get(x,x))
        if extra is not None: f=f[extra.values]
        if clamp: f.loc[f['d']<start,'d']=start      # arrivals dated before the first day (e.g. MIS saved yesterday) count on day 1
        f=f[f['d'].isin(idx)]
        out={}
        for (m,d),q in f.groupby(['m','d'])['q'].sum().items(): out.setdefault(m,[0.0]*ndays)[idx[d]]+=q
        return out
    plan=daily(tp,14,10,0,extra=tpmask)
    fc=pd.DataFrame({'m':bycol(tp,10).astype(str).str.strip(),'c':bycol(tp,2).map(lambda x:' '.join(str(x).split()) if x is not None and x==x else '(no customer)'),
                     'd':day(bycol(tp,0)),'q':num(bycol(tp,14))})[tpmask.values]
    fc=fc[fc['d'].isin(idx)&(fc['q']!=0)]
    plan_cust={}
    for (m_,c_,d_),q_ in fc.groupby(['m','c','d'])['q'].sum().items(): plan_cust.setdefault(m_,{}).setdefault(c_,[0.0]*ndays)[idx[d_]]+=q_
    if IT_STATUS_ONLY:                         # MIS STATUS 'IN-TRANSIT' only
        it=it[[it_status_group(v_)=='In transit' for v_ in bycol(it,22)]].reset_index(drop=True)
    itq=it.copy(); itq[itq.columns[34]]=[(q_*alias_factor.get(str(c_).strip(),1.0)) if isinstance(q_,(int,float)) else q_ for q_,c_ in zip(bycol(it,34),bycol(it,39))]
    trans=daily(itq,34,39,35,clamp=True,remap=alias); adh=daily(ad,3,1,4,clamp=True)
    # supplier detail: stock by Make (RM Status for OEM JW: USA + USA Warehouse; the JW stock at the OEM is in no calculation) and in transit by Supplier Name
    mk_stock={}
    for mk,cd,a_,b_ in zip(bycol(rs,0),bycol(rs,1),num(bycol(rs,5)),num(bycol(rs,7))):
        if not isinstance(mk,str) or not mk.strip(): continue
        d_=mk_stock.setdefault(str(cd).strip(),{}); d_[mk.strip().upper()]=d_.get(mk.strip().upper(),0.0)+a_+b_
    f=pd.DataFrame({'m':bycol(itq,39).astype(str).str.strip().map(lambda x: alias.get(x,x)),'s':bycol(itq,1).map(lambda x:' '.join(str(x).split()) if x is not None and x==x else '(no supplier)'),'d':day(bycol(itq,35)),'q':num(bycol(itq,34))})
    f.loc[f['d']<start,'d']=start; f=f[f['d'].isin(idx)&(f['q']>0)]
    trans_sup={}
    po_sup={}
    for po_,sn_ in zip(bycol(it,6),bycol(it,1)):
        if po_ is None or not isinstance(sn_,str) or not sn_.strip(): continue
        k_=str(po_).strip(); k_=k_[:-2] if k_.endswith('.0') else k_
        po_sup.setdefault(k_,' '.join(sn_.split()))
    for (m_,s_,d_),q_ in f.groupby(['m','s','d'])['q'].sum().items(): trans_sup.setdefault(m_,{}).setdefault(s_,[0.0]*ndays)[idx[d_]]+=q_
    # Adhoc rows (dated on or before today) for a material whose alternate codes hold stock are MIS stock conversions
    # (e.g. Cell Fixation 100 m rolls -> 50 m rolls, old busbar codes) = stock you have now -> opening Stock. Other Adhoc rows stay as Adhoc.
    adh_open=dict(conv_qty)            # only Adhoc rows proven to be a conversion of another code's stock
    tape_note={}
    for d_ in tape_dash:              # tapes: alternate-roll stock counted directly by length; MIS's Adhoc conversion for them is dropped
        conv=sum(stock.get(c_,0.0)*f_ for c_,f_ in tape_factor.items() if alias.get(c_)==d_)
        convj=sum(jw.get(c_,0.0)*f_ for c_,f_ in tape_factor.items() if alias.get(c_)==d_)
        stock[d_]=stock.get(d_,0.0)+conv; st_main[d_]=st_main.get(d_,0.0)+conv; jw[d_]=jw.get(d_,0.0)+convj
        a0=sum(q_ for c_,q_,t_ in zip(bycol(ad,1),num(bycol(ad,3)),day(bycol(ad,4))) if str(c_ or '').strip()==d_ and pd.notna(t_) and t_<=start)
        dropped=0.0
        if a0>0 and abs(a0-conv)<=max(1.0,0.02*max(a0,conv)) and d_ in adh: adh[d_][0]=max(0.0,adh[d_][0]-a0); dropped=a0
        tape_note[d_]=dict(conv=conv,adhoc=a0,dropped=dropped,lens={c_:tape_len.get(c_) for c_ in tape_factor if alias.get(c_)==d_},len=tape_len.get(d_))
    # alternate codes holding stock that MIS does NOT convert (no matching Adhoc row) - reported on Check, not counted
    alt_uncounted={}
    for c_,d_ in alias.items():
        if stock.get(c_,0)>0 and c_ not in alias_factor and c_ not in tape_factor: alt_uncounted.setdefault(d_,{})[c_]=stock[c_]
    for m_,q_ in adh_open.items():
        stock[m_]=stock.get(m_,0.0)+q_; st_main[m_]=st_main.get(m_,0.0)+q_
        if m_ in adh: adh[m_][0]=max(0.0,adh[m_][0]-q_)
    alt_of={}
    for c_,d_ in alias.items(): alt_of.setdefault(d_,[]).append(c_)
    for b in blocks:
        m=b['mat']; b['stock']=stock.get(m,0.0)+jw.get(m,0.0)
        b['alt_codes']=sorted(alt_of.get(m,[])); b['adhoc_open']=adh_open.get(m,0.0); b['alt_factor']={c_:alias_factor[c_] for c_ in alt_of.get(m,[]) if c_ in alias_factor}
        b['alt_uncounted']=alt_uncounted.get(m,{}); b['adhoc_other']=max(0.0,sum(adh.get(m,[0.0])[:1])) if m in adh else 0.0
        b['conv_unlinked']=conv_unlinked.get(m,[]); b['tape']=tape_note.get(m)
        b['pack_window']=[dict(w_,qty=raw_stock.get(w_['code'],0.0)*w_['factor'])   # own stock only (JW at OEM ignored in calculations)
                        for w_ in pack_window.get(m,[])]
        parts=[]
        if raw_stock.get(m): parts.append(f"{m} {raw_stock[m]:,.1f} {uom_raw.get(m,'')}".strip())
        if raw_jw.get(m): parts.append(f"JW {raw_jw[m]:,.1f}")
        for c_ in sorted(alt_of.get(m,[])):
            f_=alias_factor.get(c_)
            if f_ and raw_stock.get(c_): parts.append(f"{c_} {raw_stock[c_]:,.1f} {uom_raw.get(c_,'')} x{f_:g}".replace('  ',' '))
            if f_ and raw_jw.get(c_) and c_ in tape_factor: parts.append(f"JW {c_} {raw_jw[c_]:,.1f} x{f_:g}")
        b['raw']=' + '.join(parts) or '0'
        b['plan']=plan.get(m,[0.0]*ndays); b['cust_plan']=dict(sorted(plan_cust.get(m,{}).items(),key=lambda kv:-sum(kv[1]))); b['transit']=trans.get(m,[0.0]*ndays); b['adhoc']=adh.get(m,[0.0]*ndays)
    for b in blocks:
        m=b['mat']; b['st_main']=st_main.get(m,0.0); b['st_wh']=st_wh.get(m,0.0); b['st_jw']=jw.get(m,0.0); b['uom']=uom.get(m,'')
        b['plan_all']=plan_all.get(m,{})
    # RAW MATERIALS COUNTED in the PPC 'Count' sheet: the physical count replaces MIS MB52 for that material (JW at OEM stays)
    pc_=read_phys_counts(SETTINGS_SRC[0]) if SETTINGS_SRC and SETTINGS_SRC[0] else {}
    pcb_={}
    for c_,o_ in pc_.items():
        m_=alias.get(c_,c_); f_=tape_factor.get(c_,1.0) if m_!=c_ else 1.0
        mu_=str(uom.get(m_,'') or '').strip().upper()
        q_=o_['bom'][0] if (o_['bom'][1]==mu_ or not mu_) else (o_['base'][0] if o_['base'][1]==mu_ else o_['bom'][0])   # Qty_BOM when its unit is the MRP unit (glass crates -> PC), else the count's own unit (RTV drums)
        d_=o_['date']
        if m_ in {b['mat'] for b in blocks}: e_=pcb_.setdefault(m_,[0.0,[],None]); e_[0]+=q_*f_; e_[1].append((c_,q_,f_)); e_[2]=max([x for x in (e_[2],d_) if x],default=None)
    COUNTED_RM.clear()
    for b in blocks:
        if b['mat'] in pcb_:
            q_,parts_,d_=pcb_[b['mat']]
            COUNTED_RM[b['mat']]=dict(qty=q_,parts=parts_,date=d_,mis=b['st_main']+b['st_wh'],by_sup=(pc_.get(b['mat']) or {}).get('by_sup') or {})
            b['stock']=q_+b['st_jw']; b['st_main']=q_; b['st_wh']=0.0
    # STOCK DATA: every source line behind each material's stock, so the workbook can show stock as a formula (Stock Data sheet)
    stock_lines=[]; bmats={b['mat'] for b in blocks if b['mat'] not in COUNTED_RM}
    for m_,e_ in COUNTED_RM.items():
        for c_,q_,f_ in e_['parts']:
            stock_lines.append(dict(code=m_,type='Stock',source='PPC Count (physical count)'+(f" {e_['date']:%d-%b}" if e_['date'] else ''),
                                    src=c_,batch=None,sloc=None,loc=None,qty=float(q_),uom=uom.get(c_,''),factor=f_))
    for r_ in mbk.itertuples(index=False):
        c_=str(r_[2] or '').strip(); q_=r_[14] if isinstance(r_[14],(int,float)) else 0.0
        if c_ in bmats: tgt_,f_=c_,1.0
        elif alias.get(c_) in bmats and c_ in tape_factor: tgt_,f_=alias[c_],tape_factor[c_]
        else: continue
        if not q_: continue
        stock_lines.append(dict(code=tgt_,type='Stock',source='MIS MB52 (Total Stock)',src=c_,batch=r_[3],sloc=' '.join(str(x) for x in (r_[4],r_[8]) if x),
                                loc=r_[13],qty=float(q_),uom=r_[7],factor=f_))
    for m_,q_ in adh_open.items():
        if m_ in bmats and q_ and False: stock_lines.append(dict(code=m_,type='Stock',source='MIS Adhoc (stock of an alternate code converted by MIS)',src=m_,batch=None,sloc=None,loc=None,qty=float(q_),uom=uom.get(m_,''),factor=1.0))
    bmats_all={b['mat'] for b in blocks}
    for r_ in rs.itertuples(index=False):
        c_=str(r_[1] or '').strip(); q_=r_[11] if len(r_)>11 and isinstance(r_[11],(int,float)) else 0.0
        if not q_: continue
        if c_ in bmats_all: tgt_,f_=c_,1.0
        elif alias.get(c_) in bmats_all and c_ in tape_factor: tgt_,f_=alias[c_],tape_factor[c_]
        else: continue
        stock_lines.append(dict(code=tgt_,type='to OEM',source="MIS RM Status for OEM JW (JW WSA to OEM)",src=c_,batch=None,sloc=None,loc='JW at OEM',qty=float(q_),uom=uom.get(c_,''),factor=f_))
    tot_={}
    for x_ in stock_lines: tot_[x_['code']]=tot_.get(x_['code'],0.0)+x_['qty']*x_['factor']
    bad_=[b['mat'] for b in blocks if abs(tot_.get(b['mat'],0.0)-b['stock'])>max(0.01,1e-6*abs(b['stock']))]
    if bad_: print('NOTE: Stock Data lines do not add up to the stock for: '+', '.join(bad_))
    return dict(plant=plant,modtypes=modtypes,blocks=blocks,dates=dates,src=os.path.basename(mis_path),start=start,alias=alias,alias_factor=alias_factor,
                by_code=dict(transit=trans,adhoc=adh),mk_stock=mk_stock,trans_sup=trans_sup,po_sup=po_sup,
                stock_raw=raw_stock,desc_all=desc_of,cn_all=cn_,uom_all=uom_all_,stock_lines=stock_lines)

SC_COL='C'      # column of the stock quantity on 'Stock Count' used in formulas (set per workbook: 'Today Qty_BOM' when present)
def count_qty_idx(header):
    """Index of the stock quantity to use on a Stock Count header row: 'Today Qty_BOM' (count converted to the BOM unit) first."""
    h=[str(x).strip().lower() if x is not None else '' for x in (header or [])]
    for name in ('today qty_bom','today qty bom','today qty (bom)'):
        if name in h: return h.index(name)
    for name in ('today stock qty','today qty'):
        if name in h: return h.index(name)
    return 2
def count_header_and_rows(ws_rows):
    rows=[list(r) for r in ws_rows]
    hr=next((i for i,r in enumerate(rows) if r and str(r[0] or '').strip()=='Material'),None)
    return (rows[hr],rows[hr+1:]) if hr is not None else (None,rows)

GRID_SD=11; GRID_C0=12   # MRP grid: J = Supplier... K = Stock Data, L.. = days (after the Production Line column)
def set_print(ws, title):
    """Print layout used by every report sheet: landscape, fit to page width, header with sheet title + run date, page numbers."""
    try:
        ws.page_setup.orientation='landscape'; ws.page_setup.paperSize=ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0; ws.sheet_properties.pageSetUpPr.fitToPage=True
        ws.print_options.horizontalCentered=True
        ws.page_margins.left=ws.page_margins.right=0.3; ws.page_margins.top=ws.page_margins.bottom=0.5
        ws.oddHeader.left.text=f'MRP Daily - {title}'; ws.oddHeader.right.text='&D'; ws.oddFooter.center.text='Page &P of &N'
    except Exception as _e_: _swallow(1458,_e_); pass

def make_daily_stock_xlsx(DS, path, title='MRP Daily', table='MRPDaily', weekly_of=None, rowmap=None, cache=None, target_days=None):
    """Build the Daily Stock sheet (or its weekly view) in a temporary workbook, formatted like the MIS 'Daily Dashboard'.
       Columns: A Key, B Sr No., C Short Name, D Common Name, E Material, F Material Description, G Customer, H Supplier, I Headers, J Stock Data, K.. days/weeks.
       Plan is split by customer: the material's Plan row (All customers) is followed by one Plan row per customer.
       Materials with supplier detail (b['sup'], glass) get, after their normal block, one 3-row block per supplier:
       Stock (on hand, in Stock Data only), Intransit, Open PO (no vessel yet). Plan stays on the material block only."""
    from openpyxl.styles import Border, Side
    from openpyxl.styles.colors import Color
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.formatting.rule import ColorScaleRule, Rule
    from openpyxl.styles.differential import DifferentialStyle
    wb=Workbook(); ws=wb.active; ws.title=title
    NAVY='FF002060'; V='Verdana'
    acc='_ * #,##0_ ;_ * \\-#,##0_ ;_ * "-"??_ ;_ @_ '
    thin=Side(style='thin'); dot=Side(style='dotted')
    bthin=Border(left=thin,right=thin,top=thin,bottom=thin); bdot=Border(left=dot,right=dot,top=dot,bottom=dot)
    tfill=lambda th,t: PatternFill('solid',fgColor=Color(theme=th,tint=t))
    ctr=Alignment(horizontal='center',vertical='center',wrap_text=True)
    lab=dict(font=Font(name=V,size=10,bold=True,color=Color(theme=1)),fill=tfill(3,0.8999908444471572),border=bthin,alignment=ctr)
    for r_,t_ in ((1,'Plant'),(2,'Module Type')):
        c=ws.cell(r_,1,t_)
        for k,v in lab.items(): setattr(c,k,v)
    c=ws.cell(1,2,DS['plant']); c.font=Font(name=V,size=10,bold=True,color=Color(theme=1)); c.fill=tfill(5,0.5999938962981048); c.border=bthin; c.alignment=ctr
    for j,mt in enumerate(DS['modtypes'][:4]):
        c=ws.cell(2,2+j,mt); c.font=Font(name=V,size=10,bold=True,color='FFFF0000'); c.border=bthin; c.alignment=ctr
    plan_src=("Plan = module plan (PPC Plan + Revised Plan; without a PPC file the Production Plan of MRP.xlsx) x usage per module (PPC BOM, else MIS Standard BOM), OEM plan left out. ")
    if weekly_of: plan_src=(f"Weekly view of '{weekly_of}' (live formulas). Plan / Intransit / Open PO / Adhoc = week totals; Stock and Days of cover = end of week. "+plan_src)
    c=ws.cell(1,4,f"Updated {dt.date.today():%d-%b-%Y} from {DS['src']}. Line rows (Customer + Production Line): column I = where that line is supplied from, red cells = days that line is not covered. "+plan_src+f"Stock (RM) = MB52 {DS['plant']} + {DS['plant']} Warehouse + MIS code conversions (JW stock at the OEM left out), replaced by the PPC Count (rolled forward to today) when counted; CS / PM = PPC Count. "
              "Stock row = projected stock including in transit and adhoc. Open PO (no vessel yet, WSA) is NOT in the Stock row; "
              f"'Days of cover (stock + in transit + Open PO)' adds it. Days of cover = projected stock / average demand of the material's own target days (Settings); colours: 0 red, half the target yellow, target green. "
              "Glass: supplier rows show each supplier's stock (PPC Count by supplier when counted, else MB52 batches / MIS RM Status by Make; JW at the OEM left out), in transit and Open PO.")
    c.font=Font(name=V,size=8,italic=True,color=NAVY)
    ws.row_dimensions[1].height=34.5; ws.row_dimensions[2].height=36
    dates=DS['dates']; SD=GRID_SD; C0=GRID_C0
    wkn=lambda d: ((d.date()-dt.date(d.year,1,1)).days+dt.date(d.year,1,1).weekday())//7+1   # = WEEKNUM(date,2)
    if weekly_of:                            # one column per Monday week: (week start, first daily index, last daily index)
        per=[]
        for k,d in enumerate(dates):
            mon=d-pd.Timedelta(days=d.weekday())
            if per and per[-1][0]==mon: per[-1][2]=k
            else: per.append([mon,k,k])
    else: per=[[d,k,k] for k,d in enumerate(dates)]
    ND=len(per); CL=C0+ND-1; L_=L; SDL=L_(SD)
    for k,(d,_,_) in enumerate(per):
        c=ws.cell(3,C0+k,f'Week - {wkn(d)}'); c.font=Font(name=V,size=10,bold=True,color=NAVY); c.alignment=ctr
    hdr=['Key No.2','Sr No.','Short Name','Common Name','Material','Material Description','Customer','Production Line','Supplier','Headers','Stock Data']+[f'{d:%d-%m-%Y}' for d,_,_ in per]
    hfill=tfill(5,0.7999816888943144)
    for j,h in enumerate(hdr,1):
        c=ws.cell(4,j,h); c.font=Font(name=V,size=9 if j>=C0 else 10,bold=True,color=NAVY); c.fill=hfill; c.border=bdot; c.alignment=ctr
    ws.row_dimensions[4].height=24
    Dq=("'"+weekly_of.replace("'","''")+"'!") if weekly_of else ''
    DL=lambda k: L_(C0+k)
    r=5; stock_rows=[]; inv_rows=[]; short_rows=[]; inv_td={}      # inv_td: days-of-cover row -> that material's target days (colours)
    body=Font(name=V,size=10,color=NAVY); bold=Font(name=V,size=10,bold=True,color=NAVY); ital=Font(name=V,size=10,italic=True,color='FF595959')
    def ident(rr,b,sup,hd,cust=None,line=None):
        k_=b['key']+(sup if sup and sup!='All suppliers' and sup!=b.get('_sup1') else '')+(f'|{cust.strip()}' if cust and cust!='All customers' else '')+(f'|{line}' if line and not str(line).startswith('All lines') else '')
        for j,v in enumerate([k_,b['sr'],_short_lbl(b),b['common'],b['mat'],b['desc'],cust.strip() if cust else None,line,sup,hd],1): ws.cell(rr,j,v)
        ws.row_dimensions[rr].height=20.1
    # Cached results: formulas are saved together with the value Excel would calculate, so phone / e-mail previews (which do not
    # recalculate) show the numbers, and the colours are also painted as plain fills because previews ignore conditional formats.
    if cache is None: cache={}
    def C(rr,cc,v):
        if rr is not None: cache[(rr,cc)]=v
    def lerp(c1,c2,t): return ''.join(f'{round(int(c1[i:i+2],16)+(int(c2[i:i+2],16)-int(c1[i:i+2],16))*t):02X}' for i in (0,2,4))
    TD=target_days or TARGET_DAYS
    def scale_fill(v,td):                  # 0 red -> half the material's target yellow -> its target green (same as the live rule)
        if not isinstance(v,(int,float)): return None
        lo_,mid_,hi_=0,td/2,td
        if v<=lo_: col='FF0000'
        elif v>=hi_: col='00B050'
        elif v<=mid_: col=lerp('FF0000','FFFF00',(v-lo_)/(mid_-lo_))
        else: col=lerp('FFFF00','00B050',(v-mid_)/(hi_-mid_))
        return PatternFill('solid',fgColor='FF'+col)
    negfill=PatternFill('solid',fgColor='FFFFC7CE'); negfont=Font(name=V,size=10,color='FF9C0006')
    def paint(rr,cc,v,kind):
        if rr is None: return
        c=ws.cell(rr,cc)
        if kind=='inv':
            fl=scale_fill(v,inv_td.get(rr,TD))
            if fl: c.fill=fl
        elif kind in ('stock','short') and isinstance(v,(int,float)) and (v<0 if kind=='stock' else v>0): c.fill=negfill; c.font=negfont
    _sty={}; _ctr=Alignment(horizontal='center',vertical='center')
    from copy import copy as _cp
    def style(rows_,font_for):
        # same look as before, but each font / number-format combination is built once and re-used (much faster on 25,000+ cells)
        for rr in rows_:
            inv_=rr in inv_rows
            for j in range(1,CL+1):
                f=font_for(rr,j); nf=('0.0' if inv_ else acc) if j>=SD else None
                key=(id(f),nf); c=ws.cell(rr,j); sa=_sty.get(key)
                if sa is None:
                    c.border=bdot; c.alignment=_ctr; c.font=f
                    if nf: c.number_format=nf
                    _sty[key]=_cp(c._style)
                else: c._style=_cp(sa)
    for b in DS['blocks']:
        if b.get('cust_line'): custs=list(b['cust_line'].keys())          # (customer, production line)
        else: custs=[(c_,None) for c_ in (b.get('cust_plan') or {}).keys()]
        NC=len(custs)
        CPd=b.get('cust_plan') or {}
        def CPx(key):           # plan array of one Plan row: (customer, line) -> cust_line, else customer -> cust_plan
            if b.get('cust_line') and key in b['cust_line']: return b['cust_line'][key]
            cu_=key[0] if isinstance(key,tuple) else key
            return CPd.get(cu_) or [0.0]*len(dates)
        TDg=TD; TD=b.get('td') or TDg          # this material's target days (ABC / XYZ class); restored after the block
        opo=b.get('openpo') or [0.0]*len(dates)
        has_opo=any(v for v in opo)                 # no Open PO: the Firm receipt row and the '+ Open PO' cover row would only repeat zeros / the row above
        rp=r; rc=list(range(r+1,r+1+NC))
        if has_opo: rs_,rt,ro,ri,rj=range(r+1+NC,r+6+NC)
        else: (rs_,rt,ri),ro,rj=range(r+1+NC,r+4+NC),None,None
        ra=None; rPR=(rj or ri)+1
        inv_lab=('Days of cover (stock + in transit)','Days of cover (stock + in transit + Open PO)')
        used_=[sp_ for sp_ in (b.get('sup') or []) if str(sp_.get('name') or '').upper()!='UNASSIGNED' and
               (float(sp_.get('stock') or 0)>0 or sum(sp_.get('transit') or [])>0 or sum(sp_.get('plan') or [])>0)]
        b['_sup1']=used_[0]['name'] if len(used_)==1 else None             # one real supplier: show its name, not 'All suppliers'
        supname=(b['_sup1'] or 'All suppliers') if b.get('sup') else None
        NN=lambda xs: [x_ for x_ in xs if x_ is not None]
        # standard MRP record (gross requirement -> receipts -> projected available balance -> net requirement -> planned orders)
        ident(rp,b,supname,'Gross requirement','All customers' if NC else None,'All lines' if any(ln for _,ln in custs) else None)
        for rr,(cu,ln) in zip(rc,custs):
            src_=(b.get('line_src') or {}).get((_norm(cu),_line_key(ln))) if ln else None
            if src_: src_=src_.replace(b['mat'],'own stock')
            ident(rr,b,(f'from: {src_}' if (src_ and src_ not in ('-','own stock')) else supname),'Gross requirement',cu,ln)
        for rr,hd in zip((rs_,rt,ro,ri,rj),('Projected available balance','In transit (MIS) - by arrival date','Firm receipt (Open PO, no vessel)')+inv_lab):
            if rr is not None: ident(rr,b,supname,hd)
        ident(rPR,b,supname,'Planned order receipt')
        stock_rows.append(rs_); inv_rows+=NN([ri,rj]); inv_td[ri]=TD
        if rj: inv_td[rj]=TD
        if weekly_of:
            for rr in NN([rp]+rc+[rs_,rt,ro,ri,rj,rPR]): ws.cell(rr,SD,f'={Dq}{SDL}{rr}')
            for w,(d,a,z) in enumerate(per):
                col=C0+w; A_,Z_=DL(a),DL(z); rng=lambda rr: f'{Dq}{A_}{rr}:{Z_}{rr}'
                for rr in NN([rp]+rc+[rt,ro]): ws.cell(rr,col,f'=SUM({rng(rr)})')
                for rr in NN([rs_,ri,rj]): ws.cell(rr,col,f'={Dq}{Z_}{rr}')     # end-of-week values from the daily MRP
                ws.cell(rPR,col,f'=SUM({rng(rPR)})')
        else:
            for rr in [rp]+rc: ws.cell(rr,SD,f'=SUM({L_(C0)}{rr}:{L_(CL)}{rr})')
            if b.get('count_stock'):   # new MRP: CS / PM stock = Today Qty on its own Stock Count sheet (live)
                ws.cell(rs_,SD,f"=SUMIFS('Stock Count'!${SC_COL}:${SC_COL},'Stock Count'!$A:$A,$E{rs_})")
            else: ws.cell(rs_,SD,round(b['stock'],6) if VALUES_ONLY else f"=SUMIFS('Stock Data'!$L$5:$L$3000,'Stock Data'!$A$5:$A$3000,$E{rs_})")   # RM: MB52 + JW lines on Stock Data
            ws.cell(rt,SD,f'=SUM({L_(C0)}{rt}:{L_(CL)}{rt})')
            if ro: ws.cell(ro,SD,f'=SUM({L_(C0)}{ro}:{L_(CL)}{ro})')      # total in transit, like Open PO ws.cell(ri,SD,0); ws.cell(rj,SD,0)
            ws.cell(rPR,SD,f'=SUM({L_(C0)}{rPR}:{L_(CL)}{rPR})')
            for k in range(ND):
                col=C0+k; cl=L_(col); pv=L_(col-1)
                ws.cell(rp,col,round(b['plan'][k],6))
                for rr,cu in zip(rc,custs): ws.cell(rr,col,round(CPx(cu)[k],6))
                ws.cell(rt,col,round(b['transit'][k],6))
                if ro: ws.cell(ro,col,round(opo[k],6))
                ws.cell(rs_,col,f'={pv}{rs_}-{cl}{rp}+{cl}{rt}')
                # forward-looking: stock / average demand of the next TD days (blank when no demand ahead)
                ce=L_(min(col+COVER_AVG_DAYS-1,CL)); win=f'{cl}{rp}:{ce}{rp}'
                ws.cell(ri,col,f'=IF(SUM({win})=0,"",{cl}{rs_}/AVERAGE({win}))')
                if rj: ws.cell(rj,col,f'=IF(SUM({win})=0,"",({cl}{rs_}+SUM(${L_(C0)}{ro}:{cl}{ro}))/AVERAGE({win}))')
                ws.cell(rPR,col,round((b.get('por') or [0.0]*ND)[k],3))
        style(NN([rp]+rc+[rs_,rt,ro,ri,rj,rPR]),lambda rr,j: bold if rr in (ri,rj,rPR) and j>=10 else (ital if rr in rc else body))
        # line rows: red on the days this line is NOT covered (Customer Supplier, substitutes and line ranking applied), next to the receipts
        red_=PatternFill('solid',fgColor='FFC7CE')
        for rr,(cu,ln) in zip(rc,custs):
            k_=(b.get('line_stop') or {}).get(_line_key(ln)) if ln else None
            if k_ is None or k_>=len(dates): continue
            arr_=CPx((cu,ln))
            if weekly_of:
                for w,(d,a,z) in enumerate(per):
                    if z>=k_ and any(arr_[max(a,k_):z+1]): ws.cell(rr,C0+w).fill=red_
            else:
                for k in range(k_,ND):
                    if arr_[k]: ws.cell(rr,C0+k).fill=red_
            ws.cell(rr,8).fill=red_
        if not weekly_of:
            for k in range(min(ND,len(b.get('por') or []))):
                if b['por'][k]: ws.cell(rPR,C0+k).fill=PatternFill('solid',fgColor='FFEB9C')
        # ---- values for the cache (same rules as the formulas) ----
        nd=len(dates); P=list(b.get('plan') or [0.0]*nd); T=b['transit']; A=b['adhoc']; O=opo; S0=float(b.get('stock') or 0.0)
        S=[]; cur=S0
        for k in range(nd): cur=cur-P[k]+T[k]+A[k]; S.append(cur)
        cO=[]; acc_=0.0
        for k in range(nd): acc_+=O[k]; cO.append(acc_)
        win_avg=[]; win_sum=[]
        for k in range(nd):
            w_=P[k:min(k+COVER_AVG_DAYS,nd)]; win_sum.append(sum(w_)); win_avg.append(sum(w_)/len(w_) if w_ else 0.0)
        if weekly_of:
            C(rp,SD,sum(P))
            for rr,cu in zip(rc,custs): C(rr,SD,sum(CPx(cu)))
            C(rs_,SD,S0); C(rt,SD,sum(b.get('transit') or [])); C(ro,SD,sum(O)); C(ri,SD,0); C(rj,SD,0); C(rPR,SD,sum(b.get('por') or []))
            for w,(d,a,z) in enumerate(per):
                col=C0+w; ps=sum(P[a:z+1])
                C(rp,col,ps); C(rt,col,sum(T[a:z+1])); C(ro,col,sum(O[a:z+1]))
                for rr,cu in zip(rc,custs): C(rr,col,sum(CPx(cu)[a:z+1]))
                C(rs_,col,S[z]); paint(rs_,col,S[z],'stock')
                v=(S[z]/win_avg[z]) if win_sum[z] else ''; C(ri,col,v); paint(ri,col,v,'inv')
                v=((S[z]+cO[z])/win_avg[z]) if win_sum[z] else ''; C(rj,col,v); paint(rj,col,v,'inv')
                C(rPR,col,sum((b.get('por') or [0.0]*nd)[a:z+1]))
        else:
            C(rp,SD,sum(P))
            for rr,cu in zip(rc,custs): C(rr,SD,sum(CPx(cu)))
            if b.get('count_stock'): C(rs_,SD,S0)
            C(rt,SD,sum(T)); C(ro,SD,sum(O)); C(rPR,SD,sum(b.get('por') or []))
            for k in range(nd):
                col=C0+k
                C(rs_,col,S[k]); paint(rs_,col,S[k],'stock')
                v=(S[k]/win_avg[k]) if win_sum[k] else ''; C(ri,col,v); paint(ri,col,v,'inv')
                v=((S[k]+cO[k])/win_avg[k]) if win_sum[k] else ''; C(rj,col,v); paint(rj,col,v,'inv')
        if rowmap is not None: rowmap.append(dict(b=b,rp=rp,rs=rs_,rt=rt,ro=ro,ra=ra,ri=ri,rj=rj,rc=list(rc),custs=list(custs),C0=C0,ND=ND))
        r=rPR+1; TD=TDg
        for sp in (b.get('sup') if multi_sup(b) else None) or []:                 # supplier detail (only when 2+ suppliers share the material)
            nd=len(dates); has_plan='plan' in sp; TDs=b.get('td') or TDg      # supplier rows use the material's own target days
            if has_plan:     # customers mapped to this supplier on the 'Customer Supplier' sheet: own plan, projected stock and cover
                sP,s1,s2,s3,sI=r,r+1,r+2,r+3,r+4
                ident(sP,b,sp['name'],sp.get('plan_label','Plan (mapped customers)'),sp.get('cust_text') or None,'All lines' if b.get('cust_line') else None)
                for rr,hd in zip((s1,s2,s3,sI),('Stock','In transit (MIS) - by arrival date','Firm receipt (Open PO, no vessel)','Days of cover')): ident(rr,b,sp['name'],hd)
                style((sP,s1,s2,s3,sI),lambda rr,j: bold if rr==sI and j>=10 else ital)
                Pq=sp['plan']; Ss=[]; cur=sp['stock']
                for k in range(nd): cur=cur-Pq[k]+sp['transit'][k]; Ss.append(cur)
                wsum=[sum(Pq[k:k+COVER_AVG_DAYS]) for k in range(nd)]; wavg=[(sum(Pq[k:k+COVER_AVG_DAYS])/len(Pq[k:k+COVER_AVG_DAYS])) for k in range(nd)]
                stock_rows.append(s1); inv_rows.append(sI); inv_td[sI]=TDs
                if weekly_of:
                    for rr in (sP,s1,s2,s3,sI): ws.cell(rr,SD,f'={Dq}{SDL}{rr}')
                    for w,(d,a,z) in enumerate(per):
                        for rr in (sP,s2,s3): ws.cell(rr,C0+w,f'=SUM({Dq}{DL(a)}{rr}:{DL(z)}{rr})')
                        for rr in (s1,sI): ws.cell(rr,C0+w,f'={Dq}{DL(z)}{rr}')
                        C(sP,C0+w,sum(Pq[a:z+1])); C(s2,C0+w,sum(sp['transit'][a:z+1])); C(s3,C0+w,sum(sp['openpo'][a:z+1]))
                        C(s1,C0+w,Ss[z]); paint(s1,C0+w,Ss[z],'stock')
                        v=(Ss[z]/wavg[z]) if wsum[z] else ''; C(sI,C0+w,v); paint(sI,C0+w,v,'inv')
                    C(sP,SD,sum(Pq)); C(s1,SD,sp['stock']); C(s2,SD,sum(sp['transit'])); C(s3,SD,sum(sp['openpo'])); C(sI,SD,0)
                else:
                    ws.cell(sP,SD,f'=SUM({L_(C0)}{sP}:{L_(CL)}{sP})'); ws.cell(s1,SD,round(sp['stock'],6)); ws.cell(sI,SD,0)
                    for rr,key in ((s2,'transit'),(s3,'openpo')): ws.cell(rr,SD,f'=SUM({L_(C0)}{rr}:{L_(CL)}{rr})')
                    for k in range(ND):
                        col=C0+k; cl=L_(col); pv=L_(col-1); ce=L_(min(col+TDs-1,CL)); win=f'{cl}{sP}:{ce}{sP}'
                        ws.cell(sP,col,round(Pq[k],6)); ws.cell(s2,col,round(sp['transit'][k],6)); ws.cell(s3,col,round(sp['openpo'][k],6))
                        ws.cell(s1,col,f'={pv}{s1}-{cl}{sP}+{cl}{s2}')
                        ws.cell(sI,col,f'=IF(SUM({win})=0,"",{cl}{s1}/AVERAGE({win}))')
                        C(s1,col,Ss[k]); paint(s1,col,Ss[k],'stock')
                        v=(Ss[k]/wavg[k]) if wsum[k] else ''; C(sI,col,v); paint(sI,col,v,'inv')
                    C(sP,SD,sum(Pq)); C(s2,SD,sum(sp['transit'])); C(s3,SD,sum(sp['openpo']))
                r+=5; continue
            s1,s2,s3=r,r+1,r+2
            for rr,hd in zip((s1,s2,s3),('Stock','In transit (MIS) - by arrival date','Firm receipt (Open PO, no vessel)')): ident(rr,b,sp['name'],hd)
            if weekly_of:
                for rr in (s1,s2,s3): ws.cell(rr,SD,f'={Dq}{SDL}{rr}')
                for w,(d,a,z) in enumerate(per):
                    for rr in (s2,s3): ws.cell(rr,C0+w,f'=SUM({Dq}{DL(a)}{rr}:{DL(z)}{rr})')
            else:
                ws.cell(s1,SD,round(sp['stock'],6))
                for rr,key in ((s2,'transit'),(s3,'openpo')):
                    ws.cell(rr,SD,f'=SUM({L_(C0)}{rr}:{L_(CL)}{rr})')
                    for k in range(ND): ws.cell(rr,C0+k,round(sp[key][k],6))
            style((s1,s2,s3),lambda rr,j: ital)
            if weekly_of:
                C(s1,SD,sp['stock']); C(s2,SD,sum(sp['transit'])); C(s3,SD,sum(sp['openpo']))
                for w,(d,a,z) in enumerate(per): C(s2,C0+w,sum(sp['transit'][a:z+1])); C(s3,C0+w,sum(sp['openpo'][a:z+1]))
            else: C(s2,SD,sum(sp['transit'])); C(s3,SD,sum(sp['openpo']))
            r+=3
    last=r-1
    ws.add_table(Table(displayName=table,ref=f'A4:{L_(CL)}{last}',
                       tableStyleInfo=TableStyleInfo(name='TableStyleLight10',showRowStripes=True,showColumnStripes=False)))
    red=DifferentialStyle(font=Font(color='FF9C0006'),fill=PatternFill(bgColor='FFFFC7CE',fill_type='solid'))
    ws.conditional_formatting.add(' '.join(f'{L_(C0)}{x}:{L_(CL)}{x}' for x in stock_rows),Rule(type='cellIs',operator='lessThan',formula=['0'],dxf=red))
    for td_ in sorted(set(inv_td.values())):     # days of cover: 0 red / half target yellow / target green - one rule per target (Settings)
        rows_td=[x for x in inv_rows if inv_td.get(x,TD)==td_]
        ws.conditional_formatting.add(' '.join(f'{L_(C0)}{x}:{L_(CL)}{x}' for x in rows_td),
            ColorScaleRule(start_type='num',start_value=0,start_color='FFFF0000',mid_type='num',mid_value=td_/2,mid_color='FFFFFF00',
                           end_type='num',end_value=td_,end_color='FF00B050'))
    # The painted colours (for phone previews) must never outlive an edit in Excel: every cell in these rows is always covered by a
    # live rule - stock >= 0 and blank Inv. Days get a plain white rule that overrides any painted colour.
    plain=DifferentialStyle(font=Font(color='FF002060'),fill=PatternFill(bgColor='FFFFFFFF',fill_type='solid'))
    sr_=' '.join(f'{L_(C0)}{x}:{L_(CL)}{x}' for x in stock_rows); ir_=' '.join(f'{L_(C0)}{x}:{L_(CL)}{x}' for x in inv_rows)
    if stock_rows: ws.conditional_formatting.add(sr_,Rule(type='cellIs',operator='greaterThanOrEqual',formula=['0'],dxf=plain))
    if inv_rows: ws.conditional_formatting.add(ir_,Rule(type='expression',formula=[f'NOT(ISNUMBER({L_(C0)}{inv_rows[0]}))'],dxf=plain))
    if short_rows:
        xr_=' '.join(f'{L_(C0)}{x}:{L_(CL)}{x}' for x in short_rows)
        ws.conditional_formatting.add(xr_,Rule(type='cellIs',operator='greaterThan',formula=['0'],dxf=red))
        ws.conditional_formatting.add(xr_,Rule(type='cellIs',operator='lessThanOrEqual',formula=['0'],dxf=plain))
    for col,w in zip('ABCDEFGHIJKL',[25.68,8,27.26,29,19.16,51.58,20,16,16,14.58,17,13]): ws.column_dimensions[col].width=w
    for col in range(C0,CL+1): ws.column_dimensions[L_(col)].width=13
    ws.column_dimensions['A'].hidden=True
    ws.freeze_panes=f'{L_(C0)}5'; ws.sheet_view.zoomScale=80; ws.sheet_view.showGridLines=True; set_print(ws,title)
    wb.save(path)
    return len(DS['blocks'])

def graft_sheet(pkg_path, src_xlsx, sheet_name, cache=None, hidden=False, tab_color=None):
    """Put the first sheet of src_xlsx (with its table, conditional formats and styles) into the workbook pkg_path as sheet_name,
       replacing a sheet of that name if there is one. Every other part of pkg_path is copied byte for byte."""
    import zipfile, posixpath
    print(f'  [{time.time()-_T0:5.0f} s] sheet: {sheet_name}',flush=True)     # progress, so a slow step is visible
    from lxml import etree
    M='http://schemas.openxmlformats.org/spreadsheetml/2006/main'; R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    PR='http://schemas.openxmlformats.org/package/2006/relationships'; CT='http://schemas.openxmlformats.org/package/2006/content-types'
    q=lambda t:'{%s}%s'%(M,t)
    WS_T='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'; TB_T='application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'
    with zipfile.ZipFile(pkg_path) as z: infos=z.infolist(); parts={i.filename:z.read(i.filename) for i in infos}
    with zipfile.ZipFile(src_xlsx) as z: S={n:z.read(n) for n in z.namelist()}
    X=lambda b: etree.fromstring(b)
    def rels_path(p): d,f=posixpath.split(p); return f'{d}/_rels/{f}.rels'
    def resolve(base,target): return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(base),target))
    wbx=X(parts['xl/workbook.xml']); wrels=X(parts['xl/_rels/workbook.xml.rels']); ct=X(parts['[Content_Types].xml'])
    def ct_set(part,typ):
        for o in ct.findall('{%s}Override'%CT):
            if o.get('PartName')=='/'+part: ct.remove(o)
        if typ: e=etree.SubElement(ct,'{%s}Override'%CT); e.set('PartName','/'+part); e.set('ContentType',typ)
    # ---- shared strings ----
    if 'xl/sharedStrings.xml' not in parts:
        parts['xl/sharedStrings.xml']=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="{M}" count="0" uniqueCount="0"/>'.encode()
        e=etree.SubElement(wrels,'{%s}Relationship'%PR); e.set('Id','rIdSst1'); e.set('Type',R+'/sharedStrings'); e.set('Target','sharedStrings.xml')
        ct_set('xl/sharedStrings.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml')
    tsst=X(parts['xl/sharedStrings.xml']); tsis=tsst.findall(q('si'))
    plain={}
    for i,si in enumerate(tsis):
        if si.find(q('r')) is None: plain.setdefault(''.join(si.itertext()),i)
    ssst=[''.join(si.itertext()) for si in X(S['xl/sharedStrings.xml']).findall(q('si'))] if 'xl/sharedStrings.xml' in S else []
    smap={}
    for i,t in enumerate(ssst):
        if t not in plain:
            si=etree.SubElement(tsst,q('si')); te=etree.SubElement(si,q('t')); te.text=t
            if t!=t.strip(): te.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
            plain[t]=len(tsst.findall(q('si')))-1
        smap[i]=plain[t]
    tsst.set('uniqueCount',str(len(tsst.findall(q('si')))))
    # ---- styles ----
    tst=X(parts['xl/styles.xml']); sst_=X(S['xl/styles.xml'])
    ORDER=['numFmts','fonts','fills','borders','cellStyleXfs','cellXfs','cellStyles','dxfs','tableStyles','colors','extLst']
    def coll(root,name):
        e=root.find(q(name))
        if e is None:
            e=etree.Element(q(name)); after=None
            for nm in ORDER[:ORDER.index(name)]:
                if root.find(q(nm)) is not None: after=root.find(q(nm))
            if after is None: root.insert(0,e)
            else: after.addnext(e)
        return e
    def canon(e):                       # namespace-free, attribute-order-free fingerprint of a style element
        def num(v):
            try: return repr(round(float(v),6))
            except ValueError: return v
        tag=etree.QName(e).localname
        return (tag,tuple(sorted((etree.QName(k).localname,num(v)) for k,v in e.attrib.items())),tuple(canon(c) for c in e if isinstance(c.tag,str)))
    _idx={}                             # fingerprint -> index, built once per style collection (was re-computed for every lookup)
    def merge(name,elem):
        c=coll(tst,name)
        if name not in _idx:
            d_={}
            for i,k in enumerate(list(c)): d_.setdefault(canon(k),i)
            _idx[name]=d_
        key=canon(elem); d_=_idx[name]
        if key in d_: return d_[key]
        c.append(elem); c.set('count',str(len(c))); d_[key]=len(c)-1; return len(c)-1
    import copy
    # number formats
    tnf=coll(tst,'numFmts'); code2id={n.get('formatCode'):int(n.get('numFmtId')) for n in tnf}
    nfmap={}
    for n in (sst_.find(q('numFmts')) if sst_.find(q('numFmts')) is not None else []):
        fid=int(n.get('numFmtId')); code=n.get('formatCode')
        if code not in code2id:
            new=max([163]+list(code2id.values()))+1; e=etree.SubElement(tnf,q('numFmt')); e.set('numFmtId',str(new)); e.set('formatCode',code); code2id[code]=new
        nfmap[fid]=code2id[code]
    tnf.set('count',str(len(tnf)))
    if len(tnf)==0: tst.remove(tnf)
    sub={nm:[merge(nm,copy.deepcopy(e)) for e in sst_.find(q(nm))] for nm in ('fonts','fills','borders')}
    xfmap=[]
    for xf in sst_.find(q('cellXfs')):
        e=copy.deepcopy(xf)
        for attr,nm in (('fontId','fonts'),('fillId','fills'),('borderId','borders')):
            if e.get(attr) is not None: e.set(attr,str(sub[nm][int(e.get(attr))]))
        if e.get('numFmtId') is not None and int(e.get('numFmtId')) in nfmap: e.set('numFmtId',str(nfmap[int(e.get('numFmtId'))]))
        e.set('xfId','0')
        xfmap.append(merge('cellXfs',e))
    sdx=sst_.find(q('dxfs')); dxmap=[merge('dxfs',copy.deepcopy(d)) for d in (sdx if sdx is not None else [])]
    # ---- the sheet itself ----
    sh=X(S['xl/worksheets/sheet1.xml'])
    def _rc(ref):
        m_=re.match(r'([A-Z]+)(\d+)',ref); n_=0
        for ch_ in m_.group(1): n_=n_*26+ord(ch_)-64
        return int(m_.group(2)),n_
    for c in sh.iter(q('c')):
        if cache and c.find(q('f')) is not None:          # store the calculated value next to the formula
            key=_rc(c.get('r'))
            if key in cache:
                v_=cache[key]
                for old in c.findall(q('v')): c.remove(old)
                if VALUES_ONLY:                            # plain value only (faster file)
                    for f_ in c.findall(q('f')): c.remove(f_)
                    if v_ is None or v_=='': 
                        if 't' in c.attrib: del c.attrib['t']
                        continue
                ve=etree.SubElement(c,q('v'))
                if isinstance(v_,str): c.set('t','str'); ve.text=v_
                elif isinstance(v_,bool): c.set('t','b'); ve.text='1' if v_ else '0'
                else:
                    if 't' in c.attrib: del c.attrib['t']
                    ve.text=repr(round(float(v_),9))
        if c.get('s') is not None: c.set('s',str(xfmap[int(c.get('s'))]))
        if c.get('t')=='s': v=c.find(q('v')); v.text=str(smap[int(v.text)])
    for rw in sh.iter(q('row')):
        if rw.get('s') is not None: rw.set('s',str(xfmap[int(rw.get('s'))]))
    for co in sh.iter(q('col')):
        if co.get('style') is not None: co.set('style',str(xfmap[int(co.get('style'))]))
    for cf in sh.iter(q('cfRule')):
        if cf.get('dxfId') is not None: cf.set('dxfId',str(dxmap[int(cf.get('dxfId'))]))
    if tab_color:                                     # coloured sheet tab (sheetPr must be the first child of the worksheet)
        sp=sh.find(q('sheetPr'))
        if sp is None: sp=etree.Element(q('sheetPr')); sh.insert(0,sp)
        for old in sp.findall(q('tabColor')): sp.remove(old)
        tc=etree.Element(q('tabColor')); tc.set('rgb','FF'+tab_color); sp.insert(0,tc)
    for sv in sh.iter(q('sheetView')):
        if 'tabSelected' in sv.attrib: del sv.attrib['tabSelected']
    # ---- where it goes: replace the existing sheet of that name, or add a new one ----
    sheets=wbx.find(q('sheets')); rid={r_.get('Id'):r_ for r_ in wrels}
    old=next((s for s in sheets if s.get('name')==sheet_name),None)
    if old is not None:
        wpath=resolve('xl/workbook.xml',rid[old.get('{%s}id'%R)].get('Target'))
        rp=rels_path(wpath)
        if rp in parts:                                  # drop its old tables
            for r_ in X(parts[rp]):
                if r_.get('Type','').endswith('/table'):
                    tp_=resolve(wpath,r_.get('Target')); parts.pop(tp_,None); ct_set(tp_,None)
            parts.pop(rp,None)
    else:
        n=1
        while f'xl/worksheets/sheet{n}.xml' in parts: n+=1
        wpath=f'xl/worksheets/sheet{n}.xml'
        ids={r_.get('Id') for r_ in wrels}; k=1
        while f'rId{k}' in ids: k+=1
        e=etree.SubElement(wrels,'{%s}Relationship'%PR); e.set('Id',f'rId{k}'); e.set('Type',R+'/worksheet'); e.set('Target',wpath[3:])
        s=etree.SubElement(sheets,q('sheet')); s.set('name',sheet_name)
        s.set('sheetId',str(max(int(x.get('sheetId')) for x in sheets if x.get('sheetId'))+1)); s.set('{%s}id'%R,f'rId{k}')
        ct_set(wpath,WS_T)
    ent=next(s_ for s_ in sheets if s_.get('name')==sheet_name)
    if hidden: ent.set('state','hidden')
    elif 'state' in ent.attrib: del ent.attrib['state']
    # ---- table ----
    srels=X(S['xl/worksheets/_rels/sheet1.xml.rels']) if 'xl/worksheets/_rels/sheet1.xml.rels' in S else None
    for n_,b_ in S.items():
        if n_.startswith('xl/tables/'):
            names=[c.get('name','').strip().lower() for c in X(b_).iter(q('tableColumn'))]
            dup=sorted({x for x in names if names.count(x)>1})
            if dup: die(f"Internal check: table on sheet '{sheet_name}' has repeated column names {dup} - not saved, so Excel will not need to repair the file.")
    if srels is not None:
        tids=[]; tnames=set()
        for p_,b_ in parts.items():
            if p_.startswith('xl/tables/') and p_.endswith('.xml'):
                t_=X(b_); tids.append(int(t_.get('id'))); tnames.add(t_.get('name'))
        nrels=etree.Element('{%s}Relationships'%PR)
        for r_ in srels:
            if not r_.get('Type','').endswith('/table'): continue
            tb=X(S[resolve('xl/worksheets/sheet1.xml',r_.get('Target'))])
            n=1
            while f'xl/tables/table{n}.xml' in parts: n+=1
            tpath=f'xl/tables/table{n}.xml'; tid=max([0]+tids)+1; tids.append(tid)
            nm=tb.get('name'); base=nm; k=2
            while nm in tnames: nm=f'{base}{k}'; k+=1
            tnames.add(nm); tb.set('id',str(tid)); tb.set('name',nm); tb.set('displayName',nm)
            parts[tpath]=etree.tostring(tb,xml_declaration=True,encoding='UTF-8',standalone=True); ct_set(tpath,TB_T)
            e=etree.SubElement(nrels,'{%s}Relationship'%PR); e.set('Id',r_.get('Id')); e.set('Type',r_.get('Type')); e.set('Target',posixpath.relpath(tpath,posixpath.dirname(wpath)))
        if len(nrels): parts[rels_path(wpath)]=etree.tostring(nrels,xml_declaration=True,encoding='UTF-8',standalone=True)
    parts[wpath]=etree.tostring(sh,xml_declaration=True,encoding='UTF-8',standalone=True)
    # calcChain is rebuilt by Excel; removing it avoids 'repair' messages when formula cells change
    if 'xl/calcChain.xml' in parts:
        parts.pop('xl/calcChain.xml'); ct_set('xl/calcChain.xml',None)
        for r_ in list(wrels):
            if r_.get('Type','').endswith('/calcChain'): wrels.remove(r_)
    calc=wbx.find(q('calcPr'))
    if calc is None: calc=etree.SubElement(wbx,q('calcPr'))
    calc.set('fullCalcOnLoad','1')
    ser=lambda e: etree.tostring(e,xml_declaration=True,encoding='UTF-8',standalone=True)
    parts.update({'xl/workbook.xml':ser(wbx),'xl/_rels/workbook.xml.rels':ser(wrels),'[Content_Types].xml':ser(ct),'xl/styles.xml':ser(tst),'xl/sharedStrings.xml':ser(tsst)})
    tmp=pkg_path+'.tmp'; done=set()
    with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zout:
        zout.writestr('[Content_Types].xml',parts['[Content_Types].xml']); done.add('[Content_Types].xml')
        for i in infos:
            if i.filename in parts and i.filename not in done: zout.writestr(i,parts[i.filename]); done.add(i.filename)
        for n_,b_ in parts.items():
            if n_ not in done: zout.writestr(n_,b_)
    try: os.replace(tmp,pkg_path)
    except PermissionError: os.remove(tmp); die(f'Could not save {pkg_path} - is it open in Excel? Close it and run again.')

def attach_wsa(DS, BW, extra=()):
    """Adds WSA Open PO by day and the glass supplier detail to the MIS-based blocks. Returns (n_open, n_glass)."""
    d0=DS['dates'][0]; nd=len(DS['dates']); od=BW.get('open_daily',{}); n_open=0
    for b in DS['blocks']:
        v=[0.0]*nd
        for d_,q_ in od.get(b['common'],{}).items():
            k=max(0,(d_-d0).days)              # plant date already past -> shown on the first day (overdue)
            if k<nd: v[k]+=q_
        b['openpo']=v; n_open+=any(v)
    def _sup_key(name,makes):          # map a full supplier name to its Make (e.g. 'FLAT (VIETNAM) CO., LIMITED' -> FLAT)
        u=name.upper()
        for mk_ in sorted(makes,key=len,reverse=True):
            if mk_ and mk_ in u: return mk_
        for mk_ in ('FLAT','XINYI','KIBING'):
            if mk_ in u: return mk_
        return name
    # stock by supplier from the MB52 BATCHES (your Batch Master); MIS RM Status 'Make' only when no batch can be named
    bmk={}
    if MIS_PATH_[0] and (BATCH_MASTER[0] or BATCH_MASTER[1]):
        pl_={str(DS.get('plant') or '').upper(),str(DS.get('plant') or '').upper()+' WAREHOUSE'}
        al_=DS.get('alias') or {}; fa_=DS.get('alias_factor') or {}
        try:
            for r_ in list(wb_ro(MIS_PATH_[0])['MB52'].iter_rows(values_only=True))[1:]:
                if not r_ or len(r_)<15 or str(r_[13] or '').strip().upper() not in pl_ or not isinstance(r_[14],(int,float)): continue
                c_=str(r_[2] or '').strip(); m_=al_.get(c_,c_); f_=fa_.get(c_,1.0) if m_!=c_ else 1.0
                mk_=batch_make(r_[3])
                key_=re.split(r'[-\s(]',mk_.strip().upper())[0] if mk_ else None
                bmk.setdefault(m_,{}); bmk[m_][key_]=bmk[m_].get(key_,0.0)+r_[14]*f_
        except Exception as ex_: print(f'NOTE: batch split of glass stock skipped ({type(ex_).__name__})')
    n_gl=0
    for b in DS['blocks']:
        if 'glass' not in str(b['common']).lower() and b['mat'] not in extra: continue
        bm_=bmk.get(b['mat']) or {}
        if bm_ and sum(v for k,v in bm_.items() if k)>0.5*max(1.0,sum(bm_.values())):    # most stock named by batch -> use the batches
            makes={k:v for k,v in bm_.items() if k}
        else: makes=dict(DS['mk_stock'].get(b['mat'],{}))
        sup={}
        def S(k): return sup.setdefault(k,dict(name=k,stock=0.0,transit=[0.0]*nd,openpo=[0.0]*nd))
        for mk_,qv in makes.items(): S(mk_)['stock']+=qv
        gap=own_stock(b)-sum(makes.values())                 # own stock only: JW at the OEM never goes into a supplier row
        if abs(gap)>max(0.5,0.005*abs(b['stock'])): S('Unassigned')['stock']+=gap   # ignore rounding-size gaps
        for sn,v in DS['trans_sup'].get(b['mat'],{}).items():
            t_=S(_sup_key(sn,makes))['transit']
            for k in range(nd): t_[k]+=v[k]
        for sn,dd_ in BW.get('open_sup',{}).get(b['common'],{}).items():
            po_,_,shp_=sn.partition('|')
            nm_=DS['po_sup'].get(po_) or DS['po_sup'].get(shp_) or shp_
            o_=S(_sup_key(nm_,makes))['openpo']
            for d_,q_ in dd_.items():
                k=max(0,(d_-d0).days)
                if k<nd: o_[k]+=q_
        rows_=[x for x in sup.values() if abs(x['stock'])>0.5 or any(x['transit']) or any(x['openpo'])]
        rows_.sort(key=lambda x:(x['name']=='Unassigned',-(x['stock']+sum(x['transit'])+sum(x['openpo']))))
        if rows_: b['sup']=rows_; n_gl+=1
        b['sup_sap']={k_:float(v_) for k_,v_ in makes.items()}          # SAP stock (MB52) by supplier, before any count (Stock Check supplier rows)
    return n_open,n_gl


# ===================== Inventory sheet: daily stock snapshots + analysis =====================
# ===================== In Transit sheet (from the MIS 'In Transit' sheet, By Week layout) =====================
def build_mis_intransit(mis_path, DS, BW):
    """Final Pending Qty (col AI) by Final RM Code (col AN), in Monday weeks of Final AtA (col AJ)."""
    wb=wb_ro(mis_path)
    rows=list(wb['In Transit'].iter_rows(values_only=True))
    weeks=[pd.Timestamp(w).normalize() for w in BW['weeks']]; NW=len(weeks); w0=weeks[0]
    code2name={b['mat']:b['common'] for b in DS['blocks'] if b.get('common')}
    alias_=DS.get('alias') or {}
    code2uom={b['mat']:b['uom'] for b in DS['blocks']}
    q={}; outside={}; nodate={}; lines={}; desc={}
    for r in rows[1:]:
        if len(r)<40: continue
        code=str(r[39] or '').strip(); qty=r[34]; code=alias_.get(code,code)     # alternate code -> dashboard material
        if not code or not isinstance(qty,(int,float)) or isinstance(qty,bool) or qty<=0: continue
        if r[4] and code not in desc: desc[code]=' '.join(str(r[4]).split())[:45]
        d=r[35]
        if not isinstance(d,(dt.datetime,dt.date)): nodate[code]=nodate.get(code,0)+qty; continue
        d=pd.Timestamp(d).normalize(); k=((d-pd.Timedelta(days=d.weekday()))-w0).days//7
        sup=' '.join(str(r[1] or '').split()) or '(no supplier)'
        if 0<=k<NW: q.setdefault(code,{}).setdefault(sup,[0.0]*NW)[k]+=qty; lines[code]=lines.get(code,0)+1
        else: outside[code]=outside.get(code,0)+qty
    # In Transit has no customer (its 'Cust' column is the customs broker), so show who uses the material:
    # customers in MIS Total Plan Req for the plant over the same 26 weeks, with their share of the plan
    tp=list(wb['Total Plan Req'].iter_rows(values_only=True)); w_end=w0+pd.Timedelta(days=7*NW); plant=DS['plant'].upper(); cq={}
    for r in tp[1:]:
        if not tp_ok(r,plant) or not isinstance(r[0],(dt.datetime,dt.date)): continue
        d=pd.Timestamp(r[0]).normalize(); qv=r[14]
        if not (w0<=d<w_end) or not isinstance(qv,(int,float)) or qv<=0: continue
        cu=' '.join(str(r[2] or '').split()) or '(no customer)'; cd=str(r[10] or '').strip()
        cq.setdefault(cd,{}); cq[cd][cu]=cq[cd].get(cu,0.0)+qv
    def cust_text(code):
        d=cq.get(code) or {}; t=sum(d.values())
        if not t: return None
        return ' | '.join(f'{c} {v/t:.0%}' for c,v in sorted(d.items(),key=lambda kv:-kv[1]))
    def split(code,base):
        base=dict(base,customer=cust_text(code) if code else None)
        # one row per supplier (largest first); balances / sheet date only on the material's first row so they are not double counted
        sups=sorted((q.get(code) or {}).items(),key=lambda kv:-sum(kv[1])) if code else []
        if not sups: return [dict(base,supplier=None,w=[0.0]*NW)]
        return [dict(base,supplier=sp,w=w,**({} if i==0 else dict(asof=None,bal1=None,bal2=None))) for i,(sp,w) in enumerate(sups)]
    out=[]; seen=set()
    name2code={v:k for k,v in code2name.items()}
    for n,m in BW['meta'].items():                       # By Week materials first, same order
        code=name2code.get(n); seen.add(code)
        out+=split(code,dict(name=n,code=code,uom=m.get('uom') or code2uom.get(code,''),asof=m.get('asof'),bal1=m.get('bal1'),bal2=m.get('bal2')))
    for b in DS['blocks']:                               # dashboard materials not on By Week
        if b['mat'] in seen: continue
        seen.add(b['mat']); out+=split(b['mat'],dict(name=b['common'] or b['mat'],code=b['mat'],uom=b['uom'],asof=None,bal1=None,bal2=None))
    other=[x for c in sorted(q) if c not in seen for x in split(c,dict(name=(f'{c} - {desc[c]}' if c in desc and c.upper().startswith('RM') else c),code=c,uom='',asof=None,bal1=None,bal2=None))]
    return dict(rows=out,other=other,weeks=BW['weeks'],hdr1=BW['hdr1'],hdr2=BW['hdr2'],outside=outside,nodate=nodate,src=os.path.basename(mis_path),
                wsa=(BW.get('data') or {}).get('transit') or {})

def make_intransit_xlsx(IT, path, cache=None, compact=False):
    AR='Calibri'; hf=Font(name=AR,bold=True,color='FFFFFF',size=10); hfill=PatternFill('solid',fgColor='1F3864'); bf=Font(name=AR,size=10)
    bold=Font(name=AR,size=10,bold=True); blue=Font(name=AR,size=10,color='0000FF'); ttl=Font(name=AR,bold=True,size=12)
    ctr=Alignment(horizontal='center',wrap_text=True); num='#,##0;(#,##0);-'; dfy='dd-mmm-yy'
    wb=Workbook(); ws=wb.active; ws.title='In Transit'; NW=len(IT['weeks'])
    ws['A1']='In transit quantity by material and arrival week (MIS RM Dashboard, In Transit sheet)'; ws['A1'].font=ttl
    ws['A2']=(f"Source {IT['src']}, read {dt.date.today():%d-%b-%Y}. Quantity = Final Pending Qty; week = Monday of Final AtA "
              "(MIS: revised ETA port + 5 days; ETAs already past are set to today). Sheet stock date and balances are from the WSA planning sheet. "
              "Shipments carry no customer, so Customer lists who uses the material: customers in MIS Total Plan Req over the same 26 weeks, with their share of the plan.")
    ws['A2'].font=Font(name=AR,size=9,italic=True)
    def table(top,title,items):
        ws.cell(top,1,title).font=bold; h=top+1
        for j,lab in enumerate(['Material','Supplier Name','Customer (who uses it, share of plan)','UOM','Sheet stock date','Total',IT['hdr1'],IT['hdr2']]):
            c=ws.cell(h,1+j,lab); c.font=hf; c.fill=hfill; c.alignment=ctr
        for k,w in enumerate(IT['weeks']):
            c=ws.cell(h,9+k,w); c.number_format='dd-mmm'; c.font=hf; c.fill=hfill; c.alignment=ctr
        ws.row_dimensions[h].height=32
        for i,x in enumerate(items):
            r=h+1+i; ws.cell(r,1,x['name']).font=bf; ws.cell(r,2,x['supplier']).font=bf
            c=ws.cell(r,3,x.get('customer')); c.font=Font(name=AR,size=9); c.alignment=Alignment(wrap_text=True,vertical='top')
            ws.cell(r,4,x['uom']).font=bf
            if x['asof'] is not None: c=ws.cell(r,5,x['asof']); c.font=blue; c.number_format=dfy
            c=ws.cell(r,6,f'=SUM(I{r}:{L(8+NW)}{r})'); c.font=bold; c.number_format=num
            if cache is not None: cache[(r,6)]=sum(x['w'])
            for j,key in ((7,'bal1'),(8,'bal2')):
                if x[key] is not None: c=ws.cell(r,j,x[key]); c.font=blue; c.number_format=num
            for k,v in enumerate(x['w']):
                c=ws.cell(r,9+k,round(v,6) if v else 0); c.font=bf; c.number_format=num
        tr=h+1+len(items); ws.cell(tr,1,'Lines with quantity (count)').font=bold
        for k in range(NW):
            col=L(9+k); c=ws.cell(tr,9+k,f'=COUNTIF({col}{h+1}:{col}{tr-1},">0")'); c.font=bold
            if cache is not None: cache[(tr,9+k)]=sum(1 for x in items if x['w'][k]>0)
        return tr
    if compact:          # Supply sheet: the week-by-week table is the same data as MRP Daily 'Scheduled receipt' - left out
        ws['A1']='Supply - in transit MIS vs WSA planning file, and supplier coverage'
        ws['A2']=(f"Source {IT['src']}, read {dt.date.today():%d-%b-%Y}. In-transit quantities by week are on MRP Daily ('Scheduled receipt'); "
                  "shipment-by-shipment differences (BL / PO) are on the Check sheet.")
        tr=0
    else:
        tr=table(3,'In transit (MIS) - quantity in each material\'s own unit',IT['rows'])
        if IT['other']:
            tr=table(tr+3,'Other RM codes in transit (not on the Daily Dashboard)',IT['other'])
    notes=[]
    if IT['outside']: notes.append('Outside the 26 weeks: '+', '.join(f'{c} {v:,.0f}' for c,v in IT['outside'].items()))
    if IT['nodate']: notes.append('No Final AtA date: '+', '.join(f'{c} {v:,.0f}' for c,v in IT['nodate'].items()))
    for i,t in enumerate(notes): ws.cell(max(tr,2)+2+i,1,t).font=bf
    # ---------- comparison: MIS in transit vs WSA planning file (RM plan) in transit, by material and week ----------
    from openpyxl.formatting.rule import Rule
    from openpyxl.styles.differential import DifferentialStyle
    top=(tr+3+len(notes)) if not compact else 4+len(notes); NWc=NW
    ws.cell(top,1,'Comparison - in transit MIS vs WSA RM plan (by arrival week)').font=ttl
    ws.cell(top+1,1,('MIS = Final Pending Qty by Final AtA week (as above, all suppliers). WSA = deliveries with a vessel on the WSA planning sheets, by plant-date week '
                     '(pooled glass sheets are split equally between front and back). Difference = MIS - WSA; amber = not equal (rounding under 5 ignored). Shipment-level detail is on the Check sheet.')).font=Font(name=AR,size=9,italic=True)
    h=top+2
    for j,lab in enumerate(['Material','Source','Match?','UOM','','Total']):
        c=ws.cell(h,1+j,lab or None); c.font=hf; c.fill=hfill; c.alignment=ctr
    for k,w in enumerate(IT['weeks']):
        c=ws.cell(h,9+k,w); c.number_format='dd-mmm'; c.font=hf; c.fill=hfill; c.alignment=ctr
    mis_by={}; uom_by={}
    for x in IT['rows']:
        a_=mis_by.setdefault(x['name'],[0.0]*NWc)
        for k,v in enumerate(x['w'][:NWc]): a_[k]+=v
        if x.get('uom'): uom_by.setdefault(x['name'],x['uom'])
    names=[n_ for n_ in dict.fromkeys(list(mis_by)+list(IT.get('wsa') or {})) if any(mis_by.get(n_,[0])) or any((IT.get('wsa') or {}).get(n_,[0]))]
    r=h+1; diff_rows=[]; amber=PatternFill('solid',fgColor='FFEB9C')
    for n_ in names:
        mv=(mis_by.get(n_) or [0.0]*NWc)[:NWc]; wv=list((IT['wsa'].get(n_) or [0.0]*NWc))[:NWc]+[0.0]*max(0,NWc-len(IT['wsa'].get(n_) or []))
        rm,rw,rd=r,r+1,r+2
        for rr,src,vals in ((rm,'MIS',mv),(rw,'WSA',wv)):
            ws.cell(rr,1,n_).font=bf; ws.cell(rr,2,src).font=bf; ws.cell(rr,4,uom_by.get(n_)).font=bf
            c=ws.cell(rr,6,f'=SUM(I{rr}:{L(8+NWc)}{rr})'); c.font=bold; c.number_format=num
            if cache is not None: cache[(rr,6)]=sum(vals)
            for k,v in enumerate(vals): c=ws.cell(rr,9+k,round(v,6) if v else 0); c.font=bf; c.number_format=num
        ws.cell(rd,1,n_).font=bold; ws.cell(rd,2,'Difference (MIS - WSA)').font=bold

        tol=max(5.0,0.0005*max(sum(mv),sum(wv)))          # rounding tolerance
        c=ws.cell(rd,3,f'=IF(SUMPRODUCT(ABS(I{rm}:{L(8+NWc)}{rm}-I{rw}:{L(8+NWc)}{rw}))<=MAX(5,0.0005*MAX(F{rm},F{rw})),"Match","Different")'); c.font=bold
        if cache is not None: cache[(rd,3)]='Match' if sum(abs(a-b) for a,b in zip(mv,wv))<=tol else 'Different'
        c=ws.cell(rd,6,f'=F{rm}-F{rw}'); c.font=bold; c.number_format=num
        if cache is not None: cache[(rd,6)]=sum(mv)-sum(wv)
        for k in range(NWc):
            col=L(9+k); c=ws.cell(rd,9+k,f'={col}{rm}-{col}{rw}'); c.font=bold; c.number_format=num
            d_=mv[k]-wv[k]
            if cache is not None: cache[(rd,9+k)]=d_
            if abs(d_)>5: c.fill=amber
        if sum(abs(a-b) for a,b in zip(mv,wv))>tol: ws.cell(rd,3).fill=amber; ws.cell(rd,6).fill=amber
        diff_rows.append(rd); r+=3
    if diff_rows:
        cells=' '.join(f'I{x}:{L(8+NWc)}{x} F{x}' for x in diff_rows)
        ws.conditional_formatting.add(cells,Rule(type='cellIs',operator='notBetween',formula=['-5','5'],dxf=DifferentialStyle(fill=PatternFill(bgColor='FFFFEB9C',fill_type='solid'))))
        ws.conditional_formatting.add(cells,Rule(type='cellIs',operator='between',formula=['-5','5'],dxf=DifferentialStyle(fill=PatternFill(bgColor='FFFFFFFF',fill_type='solid'))))
        ccol=' '.join(f'C{x}' for x in diff_rows)
        ws.conditional_formatting.add(ccol,Rule(type='containsText',operator='containsText',text='Different',formula=[f'NOT(ISERROR(SEARCH("Different",C{diff_rows[0]})))'],dxf=DifferentialStyle(fill=PatternFill(bgColor='FFFFEB9C',fill_type='solid'))))
        ws.conditional_formatting.add(ccol,Rule(type='containsText',operator='containsText',text='Match',formula=[f'NOT(ISERROR(SEARCH("Match",C{diff_rows[0]})))'],dxf=DifferentialStyle(fill=PatternFill(bgColor='FFC6EFCE',fill_type='solid'))))
    for c_,w_ in zip('ABCDEFGH',[30,34,48,8,13,11,16,16]): ws.column_dimensions[c_].width=w_
    for k in range(NW): ws.column_dimensions[L(9+k)].width=11
    ws.freeze_panes='I5'
    set_print(wb.active,'Supply'); wb.save(path)
    return len({x['name'] for x in IT['rows'] if any(x['w'])})

def own_stock(b):
    """Stock of the material without the JW stock at the OEM (user rule: JW at the OEM is in no calculation).
       After the OEM step b['stock'] is already own stock (jw_out); counted packaging (count_stock) never had JW."""
    s_=float(b.get('stock') or 0.0)
    return s_ if (b.get('jw_out') or b.get('count_stock')) else s_-float(b.get('st_jw') or 0.0)

def multi_sup(b):
    """Supplier detail is shown only when two or more suppliers have stock, in transit or plan for the material."""
    n=0
    for sp in b.get('sup') or []:
        if str(sp.get('name') or '').upper()=='UNASSIGNED': continue
        if float(sp.get('stock') or 0)>0 or sum(sp.get('transit') or [])>0 or sum(sp.get('plan') or [])>0: n+=1
    return n>=2
def cover_days(stock, plan, inflow, openpo):
    """Same rules as the Summary formulas: days until projected stock first goes below zero (len(plan) = no run-out).
       1 existing stock only, 2 + inflow (in transit + adhoc), 3 + inflow + open PO."""
    n=len(plan); out=[]
    for use_in,use_po in ((False,False),(True,False),(True,True)):
        cp=ci=co=0.0; k_=n
        for k in range(n):
            cp+=plan[k]; ci+=inflow[k] if use_in else 0.0; co+=openpo[k] if use_po else 0.0
            if (not use_in and cp>stock) or (use_in and stock+ci+co-cp<0): k_=k; break
        out.append(k_)
    return out


# ===================== days until stock-out (used for Status and the short periods) =====================



# ===================== Substitute Advice: which other material could replace one that runs out =====================
ENCAPS={'EVA','EPE','POE'}
BOM_GROUP_MAP={'aluminium frame':'Frame','aluminum frame':'Frame','silicon sealant':'Sealant','perforated adhesive tap':'Tape','cell fixation tap':'Tape','eva':'Encapsulant',
               'epe':'Encapsulant','poe':'Encapsulant','junction box':'Junction Box','glass':'Glass','busbar':'Busbar','ribbon':'Ribbon','potting a':'Potting A',
               'potting b':'Potting B','backsheet':'Backsheet','cells':'Cell','cell':'Cell'}
def read_bom_groups(mis_path):
    """MIS Standard BOM: the group of every material (column 'Understanding Name' + item 'Sr. no.') and the module types that use it.
       Also the role inside a module: for glass and encapsulant the lower item number of a module type = FRONT, the higher = BACK."""
    out={}
    try: rows=list(wb_ro(mis_path)['Standard BOM'].iter_rows(values_only=True))[1:]
    except Exception as _e_: _swallow(2186,_e_); return out
    per_mt={}
    for r in rows:
        if not r or len(r)<9 or not r[3]: continue
        c=str(r[3]).strip(); grp=BOM_GROUP_MAP.get(str(r[8] or '').strip().lower(),str(r[8] or '').strip())
        x=out.setdefault(c,dict(group=grp,sr=set(),mts=set(),pgr=r[7],role=set()))
        x['sr'].add(r[1]); x['mts'].add(str(r[2] or '').strip())
        per_mt.setdefault((str(r[2] or '').strip(),grp),[]).append((r[1] if isinstance(r[1],(int,float)) else 99,c))
    for (mt,grp),lst in per_mt.items():
        if grp in ('Glass','Encapsulant') and len(lst)>=2:
            lst=sorted(lst)
            out[lst[0][1]]['role'].add('front')
            for _,c in lst[1:]: out[c]['role'].add('back')
    return out



def read_sap_bom(mrp_path):
    """Your SAP BOM (MRP.xlsx 'BOM' sheet): customer, module model, component, description, quantity, unit, valid from, group (RM / PM / CS)."""
    out=[]
    if not mrp_path or not os.path.exists(mrp_path): return out
    try:
        wb=load_workbook(mrp_path,read_only=True,data_only=True)
        if 'BOM' not in wb.sheetnames: return out
        for r in list(wb['BOM'].iter_rows(values_only=True))[1:]:
            if r and len(r)>11 and r[0] and r[4]:
                out.append(dict(cust=' '.join(str(r[0]).split()),model=r[1],code=str(r[4]).strip(),desc=r[5],qty=r[6],uom=r[7],valid=r[8],grp=str(r[11] or '').strip().upper()))
    except Exception as ex_: print(f'NOTE: BOM sheet of {os.path.basename(mrp_path)} not read ({type(ex_).__name__})')
    return out

def read_mis_extra(mis_path, plant):
    """Other MIS data that helps substitution:
       - Total Plan Req 'JB Type' per BOM name (which JB the plan says a module uses)
       - RM Status for OEM JW: 'Make' (approved makers per code), 'Replace Material Code' (MIS's own replacement), stock at other sites
       - Booking Approval Check Summary: pending SO / booking approval per Common Name
       - MB52: unrestricted / quality inspection / blocked per code (plant + warehouse)"""
    import collections
    w=wb_ro(mis_path); X=dict(jbtype={},makes={},replace={},sites={},booking={},mbq={})
    def rows_(n):
        try: return list(w[n].iter_rows(values_only=True))
        except Exception as _e_: _swallow(2226,_e_); return []
    for r in rows_('Total Plan Req')[1:]:
        if r and len(r)>5 and r[4] and r[5] and str(r[21] if len(r)>21 else '').strip().upper() in ('',str(plant).upper()):
            X['jbtype'].setdefault(str(r[4]).strip(),set()).add(str(r[5]).strip())
    rs=rows_('RM Status for OEM JW')
    for r in rs[1:]:
        if not r or len(r)<6 or not r[1]: continue
        c=str(r[1]).strip()
        if c.upper().startswith('JW') or c=='Total': continue
        if r[0] and str(r[0]).strip() not in ('-',''): X['makes'].setdefault(c,[]).append(str(r[0]).strip().title())
        if len(r)>16 and r[16]: X['replace'].setdefault(c,set()).add(str(r[16]).strip()); X['replace'].setdefault(str(r[16]).strip(),set()).add(c)
        st={}
        for j,lab in ((9,'USA Arizona'),):
            if len(r)>j and isinstance(r[j],(int,float)) and r[j]>0: st[lab]=st.get(lab,0)+r[j]
        if st:
            for k,v in st.items(): X['sites'].setdefault(c,{})[k]=X['sites'].get(c,{}).get(k,0)+v
    for r in rs[1:]:                     # job-work stock sits on the JW code
        if r and len(r)>11 and r[1] and str(r[1]).startswith('JW') and isinstance(r[11],(int,float)) and r[11]>0:
            X['sites'].setdefault(str(r[1])[2:].strip(),{})['JW (OEM)']=r[11]
    for r in rows_('Booking Approval Check Summary')[1:]:
        if r and len(r)>32 and r[5] and r[4] and isinstance(r[31],(int,float)):
            X['booking'][str(r[5]).strip().lower()]=(r[31] or 0,r[32] if isinstance(r[32],(int,float)) else 0,(r[33] if len(r)>33 else None))
    pl={str(plant).upper(),str(plant).upper()+' WAREHOUSE'}
    for r in rows_('MB52')[1:]:
        if not r or len(r)<15 or str(r[13] or '').strip().upper() not in pl or not r[2]: continue
        q=X['mbq'].setdefault(str(r[2]).strip(),[0.0,0.0,0.0])
        for i_,j_ in ((0,6),(1,12),(2,9)): q[i_]+=r[j_] if isinstance(r[j_],(int,float)) else 0
    return X


BATCH_MASTER=[{},{}]   # [exact batch -> make, batch code (2-5 letters) -> make] from the 'Batch Master' sheet of MRP.xlsx
USER_ALT_MISS=[]       # your mappings whose main code is not a planned material
USER_ALT={}            # your Material Group 'MRP material' / 'Factor': code -> (main code, factor)
MAT_GROUP={}           # material -> (SAP group, your group name) from the 'Material Group' sheet of MRP.xlsx
MAT_GROUP_A={}         # material -> column A 'Material Group' of that sheet (Glass / Packaging / Label ...), also when it has no column E
MIS_CAT={}             # material -> MIS Standard BOM 'Understanding Name' (column I)
def item_category(m):
    """Category of a material - ONE rule for Material Status and the ALERT (user, Nov 2026: 'Alert sheet, I want item name and category
       from stock management'): your Material Group name (column E) > MIS Standard BOM Understanding Name (also via your BOM alias) >
       your Material Group column A (packaging / labels have no MIS BOM name) > '-'."""
    return ((MAT_GROUP.get(m) or (None,''))[1] or MIS_CAT.get(m) or next((MIS_CAT[c_] for c_,t_ in (BOM_ALIAS[0] or {}).items() if t_==m and c_ in MIS_CAT),'')
            or ' '.join(str(MAT_GROUP_A.get(m) or '').split()) or '-')
def load_mis_cat(mis_path):
    MIS_CAT.clear()
    try:
        for r in list(wb_ro(mis_path)['Standard BOM'].iter_rows(values_only=True))[1:]:
            if r and len(r)>8 and r[3] and r[8]: MIS_CAT.setdefault(str(r[3]).strip(),str(r[8]).strip())
    except Exception as _e_: _swallow(2273,_e_); pass
def load_user_masters(mrp_path):
    """Your Batch Master and Material Group sheets (MRP.xlsx)."""
    BATCH_MASTER[0].clear(); BATCH_MASTER[1].clear(); MAT_GROUP.clear(); MAT_GROUP_A.clear()
    try: w=wb_ro(mrp_path)
    except Exception as _e_: _swallow(2278,_e_); return
    if 'Batch Master' in w.sheetnames:
        for r in list(w['Batch Master'].iter_rows(values_only=True))[1:]:
            if r and r[0] and len(r)>1 and r[1]:
                b=str(r[0]).strip(); m=' '.join(str(r[1]).split())
                (BATCH_MASTER[0] if len(b)>=8 else BATCH_MASTER[1])[b.upper()]=m
    USER_ALT.clear()
    if 'Material Group' in w.sheetnames:
        rows_=list(w['Material Group'].iter_rows(values_only=True))
        hl_=[' '.join(str(x or '').split()).lower() for x in (rows_[0] if rows_ else [])]
        im_=hl_.index('mrp material') if 'mrp material' in hl_ else None; if_=hl_.index('factor') if 'factor' in hl_ else None
        for r in rows_[1:]:
            if r and len(r)>1 and r[1] and r[0]: MAT_GROUP_A.setdefault(str(r[1]).strip(),str(r[0]).strip())
            if r and len(r)>4 and r[1]:
                c_=str(r[1]).strip(); MAT_GROUP[c_]=(r[0],' '.join(str(r[4] or '').split()))
                t_=str(r[im_] or '').strip() if im_ is not None and im_<len(r) else ''
                if t_:
                    f_=r[if_] if if_ is not None and if_<len(r) and isinstance(r[if_],(int,float)) and r[if_] else 1.0
                    USER_ALT[c_]=(t_,float(f_))             # your mapping: this code counts as t_ (x factor); t_==code = main code

def batch_make(batch):
    """Supplier (make) of an SAP batch from your Batch Master: the exact batch, else its code - characters 2-3 (1JU0369004 -> JU),
       or a longer code (TAIC, VINI, ...) at the start or after the first character."""
    b=str(batch or '').strip().upper()
    if not b: return None
    if b in BATCH_MASTER[0]: return BATCH_MASTER[0][b]
    for L_ in (5,4,3,2):
        for st in (1,0):
            k=b[st:st+L_]
            if len(k)==L_ and k in BATCH_MASTER[1]: return BATCH_MASTER[1][k]
    return None

def batch_suppliers(mis_path, plant, makes=None):
    """MB52 stock per material per batch supplier code (2nd-3rd character of the batch), plant + its warehouse only,
       and a guessed supplier name per code from the MIS In Transit suppliers of the same materials (fills BATCH_SEEN)."""
    import collections
    w=wb_ro(mis_path); by={}
    try: mb=list(w['MB52'].iter_rows(values_only=True))[1:]
    except Exception as _e_: _swallow(2316,_e_); return by
    pl={str(plant).upper(),str(plant).upper()+' WAREHOUSE'}
    for r in mb:
        if len(r)<15 or str(r[13] or '').strip().upper() not in pl: continue
        b=str(r[3] or '').strip(); c=str(r[2] or '').strip()
        k=b[1:3].upper() if len(b)>=3 and b[1:3].isalpha() else '?'
        mk_=batch_make(b)
        if mk_: k='='+mk_                         # named by your Batch Master
        q=float(r[14]) if isinstance(r[14],(int,float)) else 0.0
        by.setdefault(c,collections.Counter())[k]+=q
        if k!='?' and not k.startswith('='): BATCH_SEEN.setdefault(k,dict(mats=set(),guess=None))['mats'].add(c)
    sup=collections.defaultdict(collections.Counter)
    try:
        for r in list(w['In Transit'].iter_rows(values_only=True))[1:]:
            c=str(r[39] or r[31] or '').strip(); n=' '.join(str(r[1] or '').split())
            if c and n: sup[c][n]+=1
    except Exception as _e_: _swallow(2332,_e_); pass
    votes=collections.defaultdict(collections.Counter)
    for c,cnt in by.items():
        codes=[k for k in cnt if k!='?' and not k.startswith('=')]; names=sup.get(c)
        if not codes or not names: continue
        for k in codes:
            for n in names: votes[k][n]+=(3 if len(codes)==1 and len(names)==1 else 1)+(2 if any(w_.startswith(k[0]) and k[1] in w_ for w_ in n.upper().split()) else 0)
    for c,cnt in by.items():           # MIS RM Status 'Make' of the same material: a make whose name fits the code letters (JU -> Juren)
        for k in [k for k in cnt if k!='?' and not k.startswith('=')]:
            for m_ in (makes or {}).get(c,[]):
                if m_.upper().startswith(k[0]) and k[1] in m_.upper(): votes[k][m_]+=6
    for k,v in votes.items(): BATCH_SEEN[k]['guess']=v.most_common(1)[0][0]
    return by

def sup_name(k):
    if str(k).startswith('='): return str(k)[1:]                 # make from your Batch Master
    n=BATCH_SUP.get(k) or ((BATCH_SEEN.get(k) or {}).get('guess') and f"{BATCH_SEEN[k]['guess']} (guess)")
    return f"{k} = {n}" if n else k

SAP_GROUP_MAP={'W001':'Glass','W002':'Encapsulant','W003':'Cell','W005':'Ribbon','W007':'Tape','W008':'Frame','W009':'Junction Box','W010':'Sealant',
               'W011':'Busbar','W012':'Potting','W013':'Label','W021':'Tape'}
def sap_group_family(grp,desc):
    """Family from the SAP material group in MB52 (column K). Mixed groups are split by the description:
       W009 = junction boxes + connector caps + jumper cables, W010 / W012 = sealant + potting (base = A, catalyst = B)."""
    f=SAP_GROUP_MAP.get(str(grp or '').strip().upper())
    d=str(desc or '').upper()
    if f=='Junction Box':
        if 'CAP' in d: return 'Connector Cap'
        if 'JUMPER' in d or ('CABLE' in d and not re.search(r'JB|JUNCTION|ZH011|DSJB|KBT|VMSJ',d)): return 'Cable'
    if f in ('Sealant','Potting'):
        if 'POTTING' in d or 'PART A' in d or 'PART B' in d:
            return 'Potting B' if re.search(r'PART B|CATALYST|\bB\b',d) else 'Potting A'
        return 'Sealant'
    return f
USER_GROUP_MAP={'glass':'Glass','glass-g12r':'Glass','eva':'Encapsulant','epe':'Encapsulant','poe':'Encapsulant','jb':'Junction Box',
                'frame':'Frame','frame 1.2':'Frame','busbar':'Busbar','ribbon':'Ribbon','rtv':'Sealant','rtv ml':'Sealant','potting a':'Potting A',
                'potting b':'Potting B','perforated tape':'Tape','cell fixing tape (g-g)':'Tape','label':'Label','connector cap':'Connector Cap',
                'solar cell':'Cell','backsheet':'Backsheet','packaging':'Packaging','frame corner key':'Frame Corner Key'}
def _fam(code,cn_row,bom=None,sapg=None,desc=None):
    """Material family: MIS Standard BOM group ('Understanding Name'), else MIS Common Name Category 2,
       else the SAP material group in MB52, else the code prefix (PMPKM ...)."""
    if bom and bom.get(code) and bom[code]['group']: return bom[code]['group']
    if sapg and not (cn_row and len(cn_row)>5 and cn_row[5]):
        f_=sap_group_family(sapg,desc)
        if f_: return f_
    f=str((cn_row or [None]*6)[5] or '').strip() if cn_row and len(cn_row)>5 else ''
    if not f: m=re.match(r'^(?:JW)?([A-Z]+)',code or ''); f=m.group(1) if m else ''
    return 'Encapsulant' if f.upper() in ENCAPS else f
def _spec(desc):
    """Specs read from a material description."""
    d=' '.join(str(desc or '').upper().replace('–','-').replace('*','X').split())
    sp=dict(d=d)
    m=re.search(r'(\d{3,4}(?:\.\d+)?)\s*X\s*(\d{3,4}(?:\.\d+)?)(?:\s*X\s*(\d+(?:\.\d+)?))?',d)
    if m: sp['dim']=(float(m.group(1)),float(m.group(2))); sp['t']=float(m.group(3)) if m.group(3) else None
    m=re.search(r'(\d{2})\s*HS',d); sp['hs']=int(m.group(1)) if m else None
    m=re.search(r'(\d{2,3})\s*GSM',d); sp['gsm']=int(m.group(1)) if m else None
    w=[int(x) for x in re.findall(r'(?<![\d.])(1[0-4]\d\d)\s*MM',d)] or [int(x) for x in re.findall(r'WIDTH\s*:?\s*(\d{4})',d)]
    sp['width']=w[0] if w else None
    m=re.search(r'(\d{2})\s*A\b',d); sp['amp']=int(m.group(1)) if m else None
    m=re.search(r'(\d{3,4})\s*V\b',d); sp['volt']=int(m.group(1)) if m else None
    sp['excl']=bool(re.search(r'EXC\w*\.?\s*CON|EX\.?\s*CONN|EXCLCON',d)); sp['staubli']='STAUBLI' in d
    sp['model']=next((k for k in ('ZH011C','DSJB','KBT4','VMSJ') if k in d),'JB with connector (no make given)' if 'WITH CONNECTOR' in d else 'make not given')
    sp['kind']=next((k for k,ws_ in (('Edge / corner protector',('PROTECTOR',)),('Wooden pallet',('PALLET',)),('Inner box',('INNERBOX','INNER BOX','CARTON')),
                                      ('Strapping',('STRAP',)),('Stretch film',('WRAPPING','STRETCH')),('Glue sheet',('GLUE',)),('Label',('LABEL',)),('Tape',('TAPE',))) if any(w_ in d for w_ in ws_)),None)
    d2=re.sub(r'\d{3,4}\s*V\b','',d); d2=re.sub(r'ZH011C-5|DSJB12Y|DSJBY|KBT4|VMSJ30','',d2)
    pm=re.search(r'\+\s*(\d{3,4})',d2); nm=re.search(r'-\s*(\d{3,4})(?!\d)',d2)
    nums=[int(x) for x in re.findall(r'(?<![\d.])(\d{3,4})(?![\d.])',d2)]
    if pm and nm: sp['cable']=(int(pm.group(1)),int(nm.group(1)))
    elif nums: sp['cable']=(max(nums),max(nums))
    else: sp['cable']=None
    m=re.search(r'(\d)\s*X\s*(0\.\d+)',d); sp['bb']=(int(m.group(1)),float(m.group(2))) if m else None
    m=re.search(r'(0\.\d+)\s*MM\s*ROUND',d); sp['wire']=float(m.group(1)) if m else None
    m=re.search(r'(\d{1,2})\s*MM\s*X\s*(\d+)\s*M\b',d) or re.search(r'(\d{1,2})\s*MM\s*[X*]\s*(\d+)\s*M\b',d) or re.search(r'TAPE.*?(\d{1,2})\s*MM',d)
    sp['tapew']=int(m.group(1)) if m else None
    sp['front']='FRONT' in d or 'ARC' in d or 'HTAR' in d; sp['back']='BACK' in d or 'HOLE' in d or 'PRINTED' in d
    sp['words']=set(w_ for w_ in re.findall(r'[A-Z]{4,}',d) if w_ not in ('WITH','FROM','SIZE','MATERIAL','PART'))
    return sp

def compare_sub(fam,t,c):
    """Grade candidate c as a replacement for target t (both _spec dicts). Returns (grade, notes) - grade 'A' same spec,
       'B' possible with engineering check, None = not suitable."""
    n=[]; g='A'
    def need(cond,msg):
        nonlocal g
        if not cond: g='B'; n.append(msg)
    f=fam.upper()
    if f=='JUNCTION BOX':
        if t['amp'] and c['amp'] and c['amp']<t['amp']: return None,[]
        if t['amp'] and c['amp'] and c['amp']>t['amp']: n.append(f"higher rating {c['amp']}A vs {t['amp']}A - OK electrically")
        if t['volt'] and c['volt'] and c['volt']<t['volt']: return None,[]
        tc,cc=t.get('cable'),c.get('cable')
        if tc and not cc: g='B'; n.append('cable length not in its description - check it is long enough')
        if tc and cc:
            if cc[0]<tc[0] or cc[1]<tc[1]: return None,[]          # shorter cable cannot reach the next module
            need(cc==tc,f"cable +{cc[0]}/-{cc[1]} mm vs +{tc[0]}/-{tc[1]} mm (longer: check string layout / cable clips)")
        need(c['model']==t['model'],f"different JB model ({c['model']} vs {t['model']}) - must be on the module certificate (UL/IEC)")
        if c['excl'] and not t['excl']: g='B'; n.append('connector NOT included - MC4 connectors must be supplied and crimped')
        if c['staubli']!=t['staubli']: g='B'; n.append('different connector make (Staubli vs other) - check connector compatibility')
    elif f in ('FRAME','GLASS'):
        if not t.get('dim') or not c.get('dim'): return None,[]
        if abs(t['dim'][0]-c['dim'][0])>3 or abs(t['dim'][1]-c['dim'][1])>3: return None,[]
        if f=='GLASS':
            if t['front']!=c['front'] and t['back']!=c['back']: return None,[]
            need(t.get('t')==c.get('t'),f"thickness {c.get('t')} vs {t.get('t')} mm - laminate / frame groove / certificate check")
            if ('HOLE' in t['d'])!=('HOLE' in c['d']): return None,[]
            if ('HTAR' in t['d'])!=('HTAR' in c['d']): n.append('different AR coating (HTAR vs ARC) - power bin check')
        else:
            need(t['hs']==c['hs'],f"frame height {c['hs']} vs {t['hs']} HS - mounting / packing / certificate check")
            gt=re.search(r'(\d+(?:\.\d+)?)\s*MM\s*GROOVE',t['d']); gc=re.search(r'(\d+(?:\.\d+)?)\s*MM\s*GROOVE',c['d'])
            if (gt.group(1) if gt else None)!=(gc.group(1) if gc else None):
                g='B'; n.append(f"glass groove differs ({gc.group(1)+' mm groove' if gc else 'standard G-G groove'} vs {gt.group(1)+' mm groove' if gt else 'standard G-G groove'}) - groove must fit the glass stack thickness")
            if re.search(r'PITCH|\(\d+,\d+',c['d']+t['d']) and c['d']!=t['d']: g='B'; n.append('mounting hole pitch may differ - check drawing')
    elif f=='ENCAPSULANT':
        if t['width'] and c['width'] and abs(t['width']-c['width'])>6: return None,[]
        tt=next((k for k in ENCAPS if k in t['d']),''); ct=next((k for k in ENCAPS if k in c['d']),'')
        need(tt==ct,f"{ct} instead of {tt} - lamination recipe / reliability check")
        tf,cf=('FRONT' in t['d']),('FRONT' in c['d']); tb,cb=('BACK' in t['d']),('BACK' in c['d'])
        if (tf and cb) or (tb and cf): g='B'; n.append(f"{'back' if cb else 'front'} layer material used as {'front' if tf else 'back'} - check")
        need(t['gsm']==c['gsm'],f"{c['gsm']} GSM vs {t['gsm']} GSM - BOM / lamination check")
        if t['width'] and c['width'] and t['width']!=c['width']: g='B'; n.append(f"width {c['width']} vs {t['width']} mm")
    elif f=='BUSBAR':
        if not t['bb'] or not c['bb'] or t['bb'][0]!=c['bb'][0]: return None,[]
        need(t['bb'][1]==c['bb'][1],f"thickness {c['bb'][1]} vs {t['bb'][1]} mm - stringer / soldering setting")
        if ('REFLECTIVE' in t['d'])!=('REFLECTIVE' in c['d']): g='B'; n.append('reflective vs plain busbar - power check')
        if re.search(r'SN\d\dPB\d\d',t['d']+c['d']) and re.search(r'SN\d\dPB\d\d',t['d']) and re.search(r'SN\d\dPB\d\d',c['d']) and re.search(r'SN\d\dPB\d\d',t['d']).group(0)!=re.search(r'SN\d\dPB\d\d',c['d']).group(0):
            g='B'; n.append('different coating alloy - soldering temperature check')
        if re.search(r'(\d+)KG',t['d']+c['d']) and re.search(r'(\d+)KG',t['d'] or '')!=None and re.search(r'(\d+)KG',c['d'] or '')!=None and re.search(r'(\d+)KG',t['d']).group(0)!=re.search(r'(\d+)KG',c['d']).group(0):
            n.append('different spool size - check stringer spool holder')
    elif f=='RIBBON':
        if not t['wire'] or not c['wire'] or abs(t['wire']-c['wire'])>0.021: return None,[]
        need(t['wire']==c['wire'],f"wire {c['wire']} vs {t['wire']} mm - stringer setting / power check")
    elif f=='TAPE':
        if t['tapew'] and c['tapew'] and t['tapew']!=c['tapew']: return None,[]
        if ('PERFORAT' in t['d'])!=('PERFORAT' in c['d']) or ('FIXATION' in t['d'])!=('FIXATION' in c['d']): return None,[]
        if 'HIGH TEMP' in c['d'] and 'HIGH TEMP' not in t['d']: g='B'; n.append('high-temperature tape - check process')
        n.append('roll length may differ - convert quantity (rolls / pieces)')
    elif f in ('POTTING A','POTTING B','SEALANT','POTTING'):
        g='B'; n.append('different maker / formulation - adhesion & cure check; potting A and B must be the SAME brand pair' if 'POTTING' in f else 'different maker / drum size - adhesion check')
    elif f=='CONNECTOR CAP':
        if ('MALE' in t['d'].replace('FEMALE',''))!=('MALE' in c['d'].replace('FEMALE','')): return None,[]
    else:   # packaging, consumables and anything else: same kind of item (shared key words) and same / bigger size
        if not t.get('kind') or t['kind']!=c.get('kind'): return None,[]    # same kind of item only (pallet for pallet, box for box ...)
        if t.get('dim') and c.get('dim'):
            if abs(t['dim'][0]-c['dim'][0])>15 or abs(t['dim'][1]-c['dim'][1])>15:
                if c['dim'][0]<t['dim'][0] or c['dim'][1]<t['dim'][1]: return None,[]
                g='B'; n.append(f"bigger size {c['dim'][0]:.0f}x{c['dim'][1]:.0f} vs {t['dim'][0]:.0f}x{t['dim'][1]:.0f} - check packing / container load")
        elif t.get('dim') or c.get('dim'): g='B'; n.append('size not on both descriptions - check')
        nums_t=set(re.findall(r'\d+',t['d'])); nums_c=set(re.findall(r'\d+',c['d']))
        if nums_t!=nums_c and g=='A':
            sz=lambda d_: (re.search(r'\(([^)]*\d[^)]*)\)',d_) or re.search(r'(\d+(?:[X&]\d+)+)',d_))
            st_,sc_=sz(t['d']),sz(c['d'])
            g='B'; n.append(f"size / spec differs ({sc_.group(1) if sc_ else 'no size given'} vs {st_.group(1) if st_ else 'no size given'}) - check it fits")
    return g,n

def make_substitute_advice_xlsx(A, DS, d0, path, mis_path=None):
    """Substitute Advice: for every planned material, other materials of the same family whose description says they could
       replace it if it runs out - same spec (A) or possible after an engineering check (B) - with their stock, in transit and
       how many extra days they would give at the module plan. Advice only: approve, then add the pair on Settings > Substitutes."""
    from openpyxl.styles import Border, Side
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    cn=DS.get('cn_all') or {}; stk=DS.get('stock_raw') or {}; desc=DS.get('desc_all') or {}; uom=DS.get('uom_all') or {}
    tr=(DS.get('by_code') or {}).get('transit') or {}; alias=DS.get('alias') or {}
    inA={x['b']['mat']:x for x in A}
    ADVICE_ISSUES.clear()
    bom=read_bom_groups(mis_path) if mis_path else {}
    MX=read_mis_extra(mis_path,DS.get('plant')) if mis_path else dict(jbtype={},makes={},replace={},sites={},booking={},mbq={})
    sapg={}; sapt={}
    if mis_path:
        try:
            for r_ in list(wb_ro(mis_path)['MB52'].iter_rows(values_only=True))[1:]:
                if r_ and len(r_)>11 and r_[2]:
                    c_=str(r_[2]).strip(); sapg.setdefault(c_,str(r_[10] or '').strip()); sapt.setdefault(c_,str(r_[11] or '').strip())
                    if r_[5]: desc.setdefault(c_,str(r_[5]))
        except Exception as _e_: _swallow(2505,_e_); pass
    bsup=batch_suppliers(mis_path,DS.get('plant'),MX['makes']) if mis_path else {}
    def sup_txt(c_,name=None):
        # batch suppliers + MIS makes + stock status + other sites + booking, for one code
        out=[]; cnt=bsup.get(c_)
        bt_='; '.join(f"{sup_name(k)}: {q:,.0f}" for k,q in (cnt.most_common() if cnt else []) if q)
        if bt_: out.append('Batches: '+bt_)
        q=MX['mbq'].get(c_)
        if q and (q[1] or q[2]): out.append(f"unrestricted {q[0]:,.0f} / quality insp. {q[1]:,.0f} / blocked {q[2]:,.0f}")
        if MX['makes'].get(c_): out.append('Approved makes (MIS): '+', '.join(dict.fromkeys(MX['makes'][c_])))
        if MX['sites'].get(c_): out.append('Other sites: '+', '.join(f"{k} {v:,.0f}" for k,v in MX['sites'][c_].items()))
        bk=MX['booking'].get(str(name or '').strip().lower())
        if bk and (bk[0] or bk[1]): out.append(f"MIS booking: pending SO {bk[0]:,.0f}, approval given {bk[1]:,.0f}"+(f" ({bk[2]})" if bk[2] else ''))
        return ' | '.join(out)[:400] or None
    # which JB does the MIS plan say each module uses (Total Plan Req 'JB Type', e.g. 500/30A) -> JB codes with that Common Name
    jb_by_type={}
    for c_,r_ in cn.items():
        nm_=str((r_[3] if len(r_)>3 else '') or '').strip().upper()
        m_=re.match(r'^JB\s+(.+)$',nm_)
        if m_: jb_by_type.setdefault(m_.group(1).replace(' ',''),set()).add(c_)
    pool={}
    for c_ in set(cn)|set(stk)|set(inA)|set(bom):
        c_=str(c_).strip()
        if not c_ or c_.upper().startswith('JW') or c_ in alias: continue
        row=cn.get(c_); fam=_fam(c_,row,bom,sapg.get(c_),desc.get(c_))
        if not fam or fam.upper()=='CELL': continue
        x=inA.get(c_); b=x['b'] if x else None
        d_=(b.get('desc') if b else None) or desc.get(c_) or (row[1] if row else '') or ''
        st_=(b['stock'] if b else stk.get(c_,0.0)) or 0.0; it_=(x.get('tr_tot') if x else sum(tr.get(c_) or [])) or 0.0
        sp_=_spec(d_); bi=bom.get(c_)
        if bi and bi['role']:              # front / back from the MIS BOM item order beats the words of the description
            sp_['front']='front' in bi['role']; sp_['back']='back' in bi['role'] and 'front' not in bi['role']
        pool[c_]=dict(code=c_,fam=fam,desc=d_,bom=bi,name=(b['common'] if b else (row[3] if row and row[3] else d_)),uom=(b.get('uom') if b else uom.get(c_)),
                      stock=st_,it=it_,x=x,sp=sp_)
    for p_ in pool.values():                 # what is free for another material: stock + in transit - its own target days of demand
        x=p_['x']
        if x:
            own=(x.get('pl_need') or x.get('avg') or 0)*(x['b'].get('td') or REPORT_DAYS)
            p_['free']=max(0.0,p_['stock']+p_['it']-own); p_['use']=f"Used by plan - {x['status']}, lasts {x['pl_days']:.0f} d at the plan" if x.get('pl_days') is not None else f"Used by plan - {x['status']}"
        else: p_['free']=max(0.0,p_['stock']+p_['it']); p_['use']='Not in the plan - free stock'
    # YOUR SAP BOM (MRP.xlsx) vs MIS Standard BOM: same customer (Settings MIS link -> MIS BOM name), same material group
    mis_bn_fam={}
    try:
        for r_ in list(wb_ro(mis_path)['Standard BOM'].iter_rows(values_only=True))[1:]:
            if r_ and len(r_)>3 and r_[2] and r_[3]:
                c2=str(r_[3]).strip(); f2=(pool.get(c2) or {}).get('fam') or _fam(c2,cn.get(c2),bom,sapg.get(c2),desc.get(c2))
                mis_bn_fam.setdefault((str(r_[2]).strip(),f2),set()).add(c2)
    except Exception as _e_: _swallow(2552,_e_); pass
    sap_alt={}            # MIS code -> [(your code, customer, model, valid from)] where your SAP BOM uses another code of the same group
    lk_={_norm(k):v for k,v in LINKS.items()}
    sap_rows=[x_ for x_ in read_sap_bom(SETTINGS_SRC[0]) if x_['grp']=='RM']
    by_cf={}
    for x_ in sap_rows:
        f2=(pool.get(x_['code']) or {}).get('fam') or _fam(x_['code'],cn.get(x_['code']),bom,sapg.get(x_['code']),x_['desc'])
        by_cf.setdefault((x_['cust'],f2),[]).append(x_)
    for (cu_,f2),xs in sorted(by_cf.items()):
        if not f2 or f2=='Cell': continue
        L_=lk_.get(_norm(cu_)) or {}; bn_=L_.get('bom')
        mis_c=mis_bn_fam.get((str(bn_ or '').strip(),f2),set()); sap_c={x_['code'] for x_ in xs}
        al_=DS.get('alias') or {}
        if not bn_ or not mis_c or {al_.get(c3,c3) for c3 in sap_c}=={al_.get(c3,c3) for c3 in mis_c}: continue   # alternate codes / roll lengths of the same material = same
        for mc in mis_c-sap_c:
            for x_ in xs:
                if x_['code'] not in mis_c: sap_alt.setdefault(mc,[]).append((x_['code'],cu_,x_['model'],x_['valid']))
        ADVICE_ISSUES.append(dict(area='Master data',check='Your BOM (MRP.xlsx) and the MIS BOM use a different raw material',sev='High',material=f"{cu_} - {f2}",code=', '.join(sorted(sap_c)),
            a_lab='Your BOM (MRP.xlsx)',a=', '.join(sorted(sap_c)),b_lab=f'MIS BOM {bn_}',b=', '.join(sorted(mis_c)),diff=None,
            detail=f"{cu_} ({xs[0]['model']}, valid from {xs[0]['valid']}): your BOM has {', '.join(sorted(sap_c))}, MIS Standard BOM {bn_} has {', '.join(sorted(mis_c))}. "
                   "Raw materials are planned with the MIS BOM (your BOM is used for CS / PM), so the MRP plans the MIS code - this row is a cross-check.",
            action='Confirm with MIS / engineering which code the line really uses; correct the MIS BOM (or yours), or add the pair on Settings > Substitutes for this line'))
    wb=Workbook(); ws=wb.active; ws.title='Substitute Advice'; ws.sheet_view.showGridLines=False
    ws['A1']='Substitute Advice - what could replace a material if it runs out (read from material descriptions)'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=('ADVICE ONLY - nothing here changes the MRP. Candidates are materials of the same family (MIS Common Name category, or code group for packaging) whose description '
              'matches: A = same spec (only the maker / pack size / rating-above differs) - confirm it is on the module certificate and BOM; B = possible, but the differences listed '
              'need engineering / quality approval. Not listed = not suitable (e.g. lower amp JB, shorter cable, other glass / frame size). Free qty = stock + in transit minus what '
              'the plan needs of the candidate itself for its target days. Extra days = free qty / need per day of the material at the module plan (blank when the units differ - convert first). '
              'Groups come from the MIS Standard BOM (Understanding Name + item no.; front / back = item order in each module type), then MIS Common Name, then the SAP material group in MB52 (W001 glass, W002 EVA/EPE, W008 frame, W009 JB ...). '
              'Cross-check: YOUR BOM (MRP.xlsx BOM sheet, used for CS / PM) uses the candidate where the MIS BOM has this material (also listed on Check). RM planning itself stays on the MIS BOM. '
              'Supplier = 2nd-3rd character of the SAP batch (MB52) - name the codes on Settings > Batch supplier codes. Also used from MIS: Total Plan Req \'JB Type\' (the JB the plan says a module uses), '
              'RM Status \'Make\' and \'Replace Material Code\', stock at Arizona / JW, Booking Approval pending SO, MB52 unrestricted / QI / blocked. '
              'When approved, add the pair on Settings > Substitutes (MRP material, substitute, lines) so the MRP uses it.')
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:S2'); ws['A2'].alignment=Alignment(wrap_text=True,vertical='top'); ws.row_dimensions[2].height=64
    H=['Material that may run out','Code','Status','Days stock + in transit lasts at line targets','Need / day at line targets','Description',
       'Option','Substitute code','Substitute description','Match','Differences / what to check','Substitute stock','Substitute in transit','Substitute is used by','Free qty for this material','Extra days at line targets',
       'MIS BOM group of substitute (module types using it)','Substitute: suppliers (batch), stock status, MIS makes, other sites, booking','This material: suppliers (batch), stock status, MIS makes, other sites, booking']
    for j,h in enumerate(H,1):
        c=ws.cell(4,j,h); c.font=Font(name=AR,size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=NAVY); c.border=bd; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.row_dimensions[4].height=42
    r=5; fA=PatternFill('solid',fgColor='C6EFCE'); fB=PatternFill('solid',fgColor='FFEB9C')
    order_=sorted(A,key=lambda x:rrank(x['b']['mat']))                  # highest shortage risk first
    nmat=0; nwith=0
    for x in order_:
        b=x['b']; t=pool.get(b['mat']) or dict(fam=_fam(b['mat'],cn.get(b['mat']),bom,sapg.get(b['mat']),b.get('desc')),sp=_spec(b.get('desc') or desc.get(b['mat'])),bom=bom.get(b['mat']))
        need_=x.get('pl_need') or x.get('avg') or 0
        plan_bn=set(); plan_jbt=set(); plan_jb=set()
        if t['fam']=='Junction Box' and t.get('bom'):
            plan_bn={bn_ for bn_ in t['bom']['mts'] if bn_ in MX['jbtype']}
            plan_jbt={jt for bn_ in plan_bn for jt in MX['jbtype'][bn_]}
            plan_jb={c2 for jt in plan_jbt for c2 in jb_by_type.get(jt.upper().replace(' ',''),set())}
            if plan_jb and b['mat'] not in plan_jb:
                ADVICE_ISSUES.append(dict(area='Master data',check="MIS BOM and MIS plan disagree on the junction box",sev='High',material=b['common'],code=b['mat'],
                    a_lab='MIS Standard BOM',a=b['mat'],b_lab="Total Plan Req 'JB Type'",b='/'.join(sorted(plan_jbt)),diff=None,
                    detail=f"Standard BOM gives {b['mat']} for {', '.join(sorted(plan_bn))}, but Total Plan Req 'JB Type' for the same module says {'/'.join(sorted(plan_jbt))} ({', '.join(sorted(plan_jb))})",
                    action='Ask MIS / engineering which JB is right and correct the BOM (or add it on Settings > Substitutes)'))
        cands=[]
        for c_,p_ in pool.items():
            if c_==b['mat'] or (p_['fam']!=t['fam']): continue
            g,notes=compare_sub(t['fam'],t['sp'],p_['sp'])
            if not g: continue
            if (p_['uom'] or '')!=(b.get('uom') or '') and p_['uom'] and b.get('uom'): notes=notes+[f"unit {p_['uom']} vs {b.get('uom')} - convert quantity"]
            tb,cb=t.get('bom'),p_.get('bom')
            if t['fam'] in ('Glass','Encapsulant') and tb and cb and tb['role'] and cb['role'] and not (tb['role']&cb['role']):
                if t['fam']=='Glass': continue                     # front glass never replaces back glass (and the other way)
                g='B'; notes=notes+[f"MIS BOM uses it as {'/'.join(sorted(cb['role']))} layer, this material is {'/'.join(sorted(tb['role']))}"]
            if c_ in MX['replace'].get(b['mat'],set()): g='A'; notes=[f"MIS RM Status 'Replace Material Code' links these two codes"]+notes
            if c_ in (ALT_BOM.get(b['mat']) or {}): g='A'; notes=[f"YOUR ALTERNATIVE BOM (PPC BOM code not in the MIS Standard BOM) for {', '.join(sorted(ALT_BOM[b['mat']][c_]))}"]+notes
            sa_=[z_ for z_ in sap_alt.get(b['mat'],[]) if z_[0]==c_]
            if sa_: g='A'; notes=[f"YOUR SAP BOM (MRP.xlsx, CS/PM BOM) uses this code for {', '.join(dict.fromkeys(z_[1] for z_ in sa_))} ({sa_[0][2]}, valid from {sa_[0][3]}) where the MIS BOM has {b['mat']}"]+notes
            if plan_jb and c_ in plan_jb: g='A'; notes=[f"MIS Total Plan Req 'JB Type' says the module(s) {', '.join(sorted(plan_bn))} use JB {'/'.join(sorted(plan_jbt))} = this code"]+notes
            ts_={k for k,q in (bsup.get(b['mat']) or {}).items() if q and k!='?'}; cs_={k for k,q in (bsup.get(c_) or {}).items() if q and k!='?'}
            if ts_ and cs_ and ts_&cs_: notes=notes+[f"same supplier as this material ({', '.join(sorted(ts_&cs_))})"]
            cands.append((g,p_,notes))
        cands.sort(key=lambda z:(z[0],not any(n_.startswith('YOUR SAP BOM') for n_ in z[2]),not any(n_.startswith('MIS ') for n_ in z[2]),-(z[1]['free']),str(z[1]['code'])))   # code last: same order on every run
        nmat+=1; withstock=[z for z in cands if z[1]['stock']+z[1]['it']>0]
        if withstock: nwith+=1
        head=[b['common'],b['mat'],x['status'],(round(x['pl_days'],1) if x.get('pl_days') is not None else None),(round(need_,1) if need_ else None),b.get('desc') or desc.get(b['mat'])]
        show=withstock[:6] or [None]
        for k,z in enumerate(show):
            if z is None:
                why=(f"{len(cands)} same-family material(s) in master data could fit, but none has stock or in transit: "+', '.join(f"{p_['code']} ({g})" for g,p_,_ in cands[:5])) if cands else 'No suitable substitute found in the master data (same family + matching description)'
                vals=head+[None,None,None,None,why,None,None,None,None,None,None,None,sup_txt(b['mat'],b['common'])]
            else:
                g,p_,notes=z
                vals=(head if k==0 else [None,None,None,None,None,None])+[k+1,p_['code'],p_['desc'],{'A':'A - same spec','B':'B - check'}[g],'; '.join(notes) or 'same spec',
                      round(p_['stock']),round(p_['it']) or None,p_['use'],round(p_['free']),(round(p_['free']/need_,1) if need_ and not (p_['uom'] and b.get('uom') and p_['uom']!=b.get('uom')) else None),
                      ((f"{p_['bom']['group']} (item {'/'.join(str(int(s_)) for s_ in sorted(p_['bom']['sr']) if isinstance(s_,(int,float)))}) - {len(p_['bom']['mts'])} module type(s): "
                        +', '.join(sorted(p_['bom']['mts'])[:3])+(' ...' if len(p_['bom']['mts'])>3 else '')) if p_.get('bom') else 'not in MIS BOM')
                       +(f" | SAP group {sapg[p_['code']]} ({sapt.get(p_['code']) or '-'})" if sapg.get(p_['code']) else ''),
                      sup_txt(p_['code'],p_['name']),(sup_txt(b['mat'],b['common']) if k==0 else None)]
            for j,v in enumerate(vals,1):
                c=ws.cell(r,j,v); c.font=Font(name=AR,size=9,bold=(j==1)); c.border=bd; c.alignment=Alignment(vertical='top',wrap_text=j in (1,6,9,11,14,17,18,19))
                if j in (12,13,15): c.number_format='#,##0'
                if j in (4,5,16): c.number_format='#,##0.0'
            if z is not None: ws.cell(r,10).fill=fA if z[0]=='A' else fB
            if k==0:
                ws.cell(r,3).fill=PatternFill('solid',fgColor=STATUS_FILL[x['status']])
                dv=ws.cell(r,4).value
                if isinstance(dv,(int,float)): ws.cell(r,4).fill=PatternFill('solid',fgColor='FFC7CE' if dv<COVER_CRITICAL else ('FFEB9C' if dv<REPORT_DAYS else 'C6EFCE'))
                for j in range(1,len(H)+1): ws.cell(r,j).border=Border(left=thin,right=thin,bottom=thin,top=Side(style='medium',color=NAVY))
            r+=1
    ws.auto_filter.ref=f'A4:{L(len(H))}{max(5,r-1)}'
    for j,w in enumerate([24,15,10,11,10,30,6,15,34,11,48,11,11,26,11,10,40,48,48],1): ws.column_dimensions[L(j)].width=w
    set_print(ws,'Substitute Advice'); wb.save(path)
    return nmat,nwith

# ===================== In Transit Status: MIS 'In Transit' STATUS column analysed =====================
IT_GROUPS=['In transit','At origin','At port (POD)','Held (customs / issue)','Partly received - rest pending / in transit','Transloaded at warehouse','Delivered','Other','No status']
def it_status_group(t):
    """Group the free-text MIS STATUS into a few statuses (the exact MIS text stays on the detail rows)."""
    u=' '.join(str(t or '').replace('\xa0',' ').split()).lower()
    if not u: return 'No status'
    if any(w in u for w in ('stuck','stucked','inspection','hold')): return 'Held (customs / issue)'
    if u in ('in-transit','in transit','intransit','in - transit'): return 'In transit'
    if 'origin' in u: return 'At origin'
    if re.search(r'\bpod\b',u) or 'at port' in u: return 'At port (POD)'
    if 'pending' in u or 'in transit' in u or 'in-transit' in u: return 'Partly received - rest pending / in transit'
    if 'transload' in u: return 'Transloaded at warehouse'
    if 'deliver' in u or 'completed' in u: return 'Delivered'
    return 'Other'

def make_intransit_status_xlsx(mis_path, DS, path):
    """In Transit Status sheet: how much is in which MIS status - by status, by material x status, and every shipment line."""
    from openpyxl.styles import Border, Side
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    hf=Font(name=AR,size=9,bold=True,color='FFFFFF'); hfill=PatternFill('solid',fgColor=NAVY); bf=Font(name=AR,size=9); bold=Font(name=AR,size=9,bold=True)
    alias=DS.get('alias') or {}; fac=DS.get('alias_factor') or {}
    info={b['mat']:b for b in DS['blocks']}
    rows=list(wb_ro(mis_path)['In Transit'].iter_rows(values_only=True))[1:]
    n_=lambda v: float(v) if isinstance(v,(int,float)) and not isinstance(v,bool) else 0.0
    d_=lambda v: v if isinstance(v,(dt.datetime,dt.date)) else None
    recs=[]
    for r in rows:
        if len(r)<40: continue
        code0=str(r[39] or r[31] or '').strip()
        if not code0 and not r[1]: continue
        if not CODE.match(code0) and not re.match(r'^[A-Z]{2,}[A-Z0-9]*\d{3,}$',code0): code0=''      # e.g. 'No Material Found'
        code=alias.get(code0,code0); b=info.get(code)
        recs.append(dict(grp=it_status_group(r[22]),st=' '.join(str(r[22] or '').replace('\xa0',' ').split()),code=code,code0=code0,
                         mat=(b['common'] if b else ' '.join(str((r[36] if code0 else None) or r[4] or code0 or '(no RM code in MIS)').split())[:60]),uom=(b.get('uom') if b else ''),
                         sup=' '.join(str(r[1] or '').split()),inv=r[2],po=r[6],bl=r[23] or r[24],vessel=' '.join(str(r[12] or '').split()),
                         cont=n_(r[21]),etd=d_(r[26]),atd=d_(r[27]),eta=d_(r[29]) or d_(r[28]),ata=d_(r[35]),
                         qty=n_(r[32]),grn=n_(r[33]),pend=n_(r[34])*fac.get(code0,1.0),rate=n_(r[8]),wh=r[30],planned=b is not None))
    wb=Workbook(); ws=wb.active; ws.title='In Transit Status'; ws.sheet_view.showGridLines=False
    ws['A1']='In Transit Status - how much is in which status (MIS In Transit, STATUS column)'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=("Pending qty = MIS 'Final Pending Qty.' (not yet received at the plant; alternate codes converted to the MRP unit). The free-text STATUS is grouped into a few statuses; "
              "the exact MIS text is on every detail row (column B). Shipped qty = 'Final Qty.', GRN = 'PPC GRN'. Value = unit rate x pending qty (invoice currency).")
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:N2'); ws['A2'].alignment=Alignment(wrap_text=True); ws.row_dimensions[2].height=30
    def head(r,cols,widths=None):
        for j,t in enumerate(cols,1): c=ws.cell(r,j,t); c.font=hf; c.fill=hfill; c.border=bd; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        ws.row_dimensions[r].height=30
    grps=[g for g in IT_GROUPS if any(x['grp']==g for x in recs)]
    r=4
    # ---- 2) pending qty by material x status ----
    ws.cell(r,1,'1. PENDING QTY BY MATERIAL AND STATUS').font=Font(name=AR,size=12,bold=True,color=NAVY); r+=1
    pg=[g for g in grps if any(x['grp']==g and x['pend']>0 for x in recs)]
    head(r,['Material','Code','UOM']+pg+['Total pending','First arrival at plant (pending)']); r+=1
    mats={}
    for x in recs:
        if x['pend']<=0: continue
        m=mats.setdefault(x['code'] or x['mat'],dict(mat=x['mat'],code=x['code'],uom=x['uom'],q={},first=None))
        m['q'][x['grp']]=m['q'].get(x['grp'],0.0)+x['pend']
        if x['ata'] and (m['first'] is None or x['ata']<m['first']): m['first']=x['ata']
    for k,m in sorted(mats.items(),key=lambda kv:(rrank(kv[1]['code']),str(kv[1]['mat']).lower())):     # highest shortage risk first
        vals=[m['mat'],m['code'],m['uom']]+[m['q'].get(g) or None for g in pg]+[sum(m['q'].values()),m['first']]
        for j,v in enumerate(vals,1):
            c=ws.cell(r,j,v); c.font=bold if j==len(vals)-1 else bf; c.border=bd
            if 4<=j<=len(vals)-1: c.number_format='#,##0'
            if j==len(vals): c.number_format='dd-mmm-yy'
        r+=1
    r+=2
    # ---- 3) detail (filter here) ----
    ws.cell(r,1,'2. EVERY SHIPMENT LINE (filter by Status / Material)').font=Font(name=AR,size=12,bold=True,color=NAVY); r+=1
    H=['Status','MIS status (exact text)','Material','Code','UOM','Supplier','Invoice','PO','BL','Vessel / voyage','Containers','ETD','ATD','ETA port (revised)',
       'Arrival at plant (Final AtA)','Shipped qty','GRN qty','Pending qty','Warehouse','In MRP']
    head(r,H); h0=r; r+=1
    order_={g:i for i,g in enumerate(IT_GROUPS)}
    for x in sorted(recs,key=lambda x:(x['pend']<=0,rrank(x['code']),order_.get(x['grp'],99),x['ata'] or dt.datetime(2100,1,1))):
        vals=[x['grp'],x['st'] or None,x['mat'],x['code0'],x['uom'],x['sup'],x['inv'],x['po'],x['bl'],x['vessel'] or None,x['cont'] or None,x['etd'],x['atd'],x['eta'],x['ata'],
              x['qty'] or None,x['grn'] or None,x['pend'] or None,x['wh'],'Yes' if x['planned'] else 'No']
        for j,v in enumerate(vals,1):
            c=ws.cell(r,j,v); c.font=bf if x['pend']>0 else Font(name=AR,size=9,color='808080'); c.border=bd; c.alignment=Alignment(vertical='top',wrap_text=j in (2,3))
            if j in (12,13,14,15): c.number_format='dd-mmm-yy'
            if j in (11,16,17,18): c.number_format='#,##0'
        if x['pend']>0: ws.cell(r,18).font=bold
        r+=1
    if r-1>h0: ws.auto_filter.ref=f'A{h0}:{L(len(H))}{r-1}'
    for j,w in enumerate([24,34,26,16,6,26,13,12,16,20,9,10,10,11,12,11,11,11,16,7],1): ws.column_dimensions[L(j)].width=w
    ws.freeze_panes='B4'; set_print(ws,'In Transit Status'); wb.save(path)
    return len(recs)

# ===================== Check sheet: discrepancies between MIS, the WSA planning file and the MRP workbook =====================
def build_checks(DS, BW, mis_path, mrp_path=None, LIMIT_DAYS=30):
    """Returns a list of issues (dicts). Sources: MIS RM Dashboard (DS + raw sheets), WSA planning file (BW), MRP input workbook."""
    import re as _re
    I=[]; today=pd.Timestamp(DS['dates'][0]).normalize(); nd=len(DS['dates']); H=min(LIMIT_DAYS,nd)
    def add(area,check,sev,material='',code='',a_lab='',a=None,b_lab='',b=None,detail='',action=''):
        diff=(b-a) if isinstance(a,(int,float)) and isinstance(b,(int,float)) else None
        I.append(dict(area=area,check=check,sev=sev,material=material,code=code,a_lab=a_lab,a=a,b_lab=b_lab,b=b,diff=diff,detail=detail,action=action))
    blocks={b['mat']:b for b in DS['blocks']}; c2n={b['mat']:b['common'] for b in DS['blocks']}; n2c={v:k for k,v in c2n.items()}
    for a_,d_ in (DS.get('alias') or {}).items(): c2n.setdefault(a_,c2n.get(d_))
    avgplan={m:(sum(b['plan'][:H])/H if H else 0) for m,b in blocks.items()}
    mis=wb_ro(mis_path)
    num=lambda v: float(v) if isinstance(v,(int,float)) and not isinstance(v,bool) else (float(v) if isinstance(v,str) and _re.fullmatch(r'-?\d+(\.\d+)?',v.strip()) else None)
    fam=lambda n: {'G12R Front Glass':'G12R Glass','G12R Back Glass':'G12R Glass','M10R Glass Front':'M10R Glass','M10R Glass Back':'M10R Glass'}.get(str(n),str(n))

    # ---------- STOCK ----------
    for m,b in blocks.items():
        fac=b.get('alt_factor') or {}
        tp_=b.get('tape')
        if tp_:
            add('Master data','Tape rolls converted by length (from the description)','Low',b['common'],m,f"{m} roll",tp_.get('len'),'Other rolls in stock (converted)',round(tp_['conv'],1),
                f"1 roll of {m} = {tp_.get('len'):g} m; "+', '.join(f"{c} = {l:g} m (x{l/tp_['len']:g})" for c,l in tp_['lens'].items() if l)+
                (f"; MIS Adhoc conversion {tp_['adhoc']:,.1f} dropped (counted directly instead)" if tp_.get('dropped') else (f"; MIS Adhoc {tp_['adhoc']:,.1f} kept (does not match the converted stock)" if tp_.get('adhoc') else '')),
                'OK - check the roll lengths in the descriptions are right')
            fac={}
        if fac:
            add('Master data','Alternate code converted by MIS (counted)','Low',b['common'],m,'',None,'Opening stock from conversion',b.get('adhoc_open') or None,
                'Stock of '+', '.join(f'{c} (x{f:g})' for c,f in fac.items())+f' is in {m} opening stock through the MIS Adhoc conversion; their in-transit is counted with the same ratio',
                'OK - no action, or ask MIS to move the dashboard to the code in use')
        unc=b.get('alt_uncounted') or {}
        if unc:
            tot=sum(unc.values()); ap=avgplan.get(m,0)
            add('Stock','Same-name stock NOT counted (alternate code)','High' if ap and tot>=ap*3 else 'Medium',b['common'],m,'Stock counted',own_stock(b),'Not counted',tot,
                ', '.join(f'{c} {q:,.1f}' for c,q in unc.items())+f' has the same MIS Common Name (alternate code, other unit) - not counted (MIS Adhoc is not used); if usable for {m}, that is'+(f' {tot/ap:.1f} more days' if ap else ' more stock'),
                'Confirm with production: if usable, add it on Settings > Substitutes (MRP material, substitute, factor) so the MRP counts it')
        for c_,nm_,r_,q_ in (b.get('conv_unlinked') or []):
            add('Master data','Code not linked on the MIS Common Name sheet','High',b['common'],m,'',None,'MIS Adhoc today',q_,
                f"MIS converts {c_} ('{nm_}') into {m} on its Adhoc sheet (x{r_:g}), but the Common Name sheet calls it '{nm_}', not '{b['common']}'. "
                f"Its stock shows on the Adhoc row (day 1), not in Stock, and its in-transit is not counted",
                f"Add / correct the row on the MIS Common Name sheet: {c_} -> Common Name '{b['common']}'")
        ao=(b.get('adhoc_other') or 0)-sum(q_ for _,_,_,q_ in (b.get('conv_unlinked') or []))
        if ao>0.5:
            add('Stock','MIS Adhoc quantity counted today (not a code conversion)','High' if ao>own_stock(b) else 'Medium',b['common'],m,'MB52 stock (JW at OEM out)',own_stock(b),'Adhoc today',ao,
                f'MIS Adhoc adds {ao:,.0f} on day 1 ({ao/max(own_stock(b),1):.1f}x the stock) - counted as in MIS; the WSA sheet balance may not include it',
                'Confirm what this Adhoc quantity is (stock elsewhere? receipt?) - if not real, ask MIS to remove it')

    # 1 MIS stock vs WSA sheet balance on the sheet stock date (pooled sheets compared as a whole; sheets sharing a material together)
    for lab,sh in _sheet_pools(BW.get('sheets',{})).items():
        mems=[m for m in sh['members'] if m in n2c]
        if not mems or sh['bal'] is None: continue
        mis_st=sum(own_stock(blocks[n2c[m]]) for m in mems)  # own stock (JW at the OEM in no calculation); pooled sheet (e.g. glass front + back): WSA balance = total of the materials
        wsa=float(sh['bal']); d=mis_st-wsa; ap=sum(avgplan[n2c[m]] for m in mems)
        days=abs(d)/ap if ap else None
        if abs(d)>max(1,0.02*max(abs(wsa),abs(mis_st))):
            sev='High' if (days is not None and days>=2) else 'Medium' if (days is None or days>=0.5) else 'Low'
            add('Stock','MIS stock vs WSA sheet balance',sev,' / '.join(mems),', '.join(n2c[m] for m in mems),
                'WSA balance',wsa,'MIS stock',mis_st,
                f"WSA sheet{'s' if sh.get('n',1)>1 else ''} '{sh['sheet']}' balance on {sh['asof']:%d-%b} vs MIS stock today ({DS['src']})"+(f"; gap = {days:.1f} days of plan" if days is not None else '')
                +("; pooled sheet, MIS = total of its materials" if len(mems)>1 else '')+("; one material on 2+ sheets (by customer and supplier) - WSA = total of the sheets" if sh.get('n',1)>1 else ''),
                'Check which count is current (receipts or issues booked on one side only)')
    # 3 Supplier split gap (RM Status by Make vs MB52) for glass
    for m,b in blocks.items():
        for sp in b.get('sup') or []:
            if sp['name']=='Unassigned':
                add('Stock','Glass stock without supplier (RM Status vs MB52)','Low',b['common'],m,'MB52 total',own_stock(b),'Sum by Make (RM Status)',own_stock(b)-sp['stock'],
                    'Batches in MB52 that MIS RM Status does not assign to a Make','Ask MIS owner to map these batches to a supplier')
    # 5 RM unit: WSA sheet UOM vs MIS MB52 unit
    misu={}
    for r in mis['MB52'].iter_rows(min_row=2,values_only=True):
        c=str(r[2] or '').strip()
        if c and r[7]: misu.setdefault(c,str(r[7]).strip())
    for lab,sh in BW.get('sheets',{}).items():
        for mname in sh['members']:
            c=n2c.get(mname)
            syn=lambda u: {'SQM':'M2','M2':'M2','PCS':'EA','PC':'EA','NOS':'EA','EA':'EA','SET':'EA','SETS':'EA','KG':'KG','KGS':'KG','DR':'DR'}.get(_re.sub(r'[^A-Z0-9]','',str(u).upper()),_re.sub(r'[^A-Z0-9]','',str(u).upper()))
            if c and sh['uom'] and misu.get(c) and syn(sh['uom'])!=syn(misu[c]):
                add('Units','WSA sheet unit vs SAP unit','Low',mname,c,'',None,'',None,f"WSA sheet '{sh['sheet']}' says {sh['uom']}, MB52 uses {misu[c]} (quantities look like {misu[c]})",'Correct the unit label on the planning sheet')

    # ---------- PLAN (next LIMIT_DAYS days) ----------
    # 6 MIS plan vs WSA sheet demand
    for lab,sh in _sheet_pools(BW.get('sheets',{})).items():
        mems=[m for m in sh['members'] if m in n2c]
        if not mems or not sh['dem']: continue
        wsa=sum(v for d,v in sh['dem'].items() if today<=pd.Timestamp(d)<today+pd.Timedelta(days=H))
        mis_p=sum(sum(blocks[n2c[m]]['plan'][:H]) for m in mems)
        if abs(mis_p-wsa)>max(1,0.10*max(mis_p,wsa)):
            add('Plan',f'MIS plan vs WSA sheet demand (next {H} days)','High' if abs(mis_p-wsa)>0.25*max(mis_p,wsa,1) else 'Medium',' / '.join(mems),', '.join(n2c[m] for m in mems),
                'WSA demand',wsa,'MIS plan',mis_p,f"WSA sheet{'s' if sh.get('n',1)>1 else ''} '{sh['sheet']}' Demand column vs MIS Total Plan Req"+(' (pooled sheet, MIS = total of its materials)' if len(mems)>1 else '')
                +(' (one material on 2+ sheets, by customer and supplier: WSA = total of the sheets)' if sh.get('n',1)>1 else '')
                +(' - about 2x: this glass sheet seems to plan front + back glass sets under one material' if mis_p and 1.8<wsa/mis_p<2.2 and 'glass' in sh['sheet'].lower() else ''),
                'Align the production plan used by the planning file and MIS')
    # 9 Customer names
    mis_c=set()
    for r in mis['Total Plan Req'].iter_rows(min_row=2,values_only=True):
        if tp_ok(r,DS['plant']) and r[2]: mis_c.add(' '.join(str(r[2]).split()))
    pp_c=set(); mp_cust=set()
    import difflib
    allc=sorted(mis_c|pp_c|{c for c in mp_cust if c})
    for i,a in enumerate(allc):
        for b in allc[i+1:]:
            if a.lower()!=b.lower() and difflib.SequenceMatcher(None,a.lower(),b.lower()).ratio()>=0.75:
                where=lambda x: '/'.join(n for n,s_ in (('MIS',mis_c),('Production Plan',pp_c),('Material Plan',mp_cust)) if x in s_)
                add('Master data','Customer names that look alike','Low',f'{a}  vs  {b}','','',None,'',None,f"'{a}' in {where(a)}; '{b}' in {where(b)}",'Confirm if same customer and use one spelling')

    # ---------- IN TRANSIT: shipment by shipment (BL number) ----------
    it=list(mis['In Transit'].iter_rows(values_only=True))
    M=[]
    for r in it[1:]:
        if len(r)<40: continue
        q=num(r[34])
        if not q or q<=0: continue
        code=str(r[39] or '').strip()
        M.append(dict(code=code,mat=c2n.get(code),fam=fam(c2n.get(code)) if code in c2n else code,qty=q,po=r[6],sup=r[1],vessel=r[12],
                      date=pd.Timestamp(r[35]).normalize() if isinstance(r[35],(dt.datetime,dt.date)) else None,
                      keys={t for f_ in (r[23],r[24],r[10]) for t in _re.split(r'[^A-Z0-9]+',str(f_ or '').upper()) if len(t)>=8 and any(ch.isdigit() for ch in t)}))
    W=[dict(d,fam=fam(d['name']),keys={t for t in _re.split(r'[^A-Z0-9]+',str(d['bl'] or '').upper()) if len(t)>=8 and any(ch.isdigit() for ch in t)})
       for d in BW.get('deliv',[]) if d['stage']=='In transit']
    idx={}
    for j,m in enumerate(M):
        for t in m['keys']: idx.setdefault(t,set()).add(j)
    parent={}
    def find(x):
        while parent.setdefault(x,x)!=x: x=parent[x]
        return x
    matchedW=set(); matchedM=set()
    for i,w in enumerate(W):
        for t in w['keys']:
            for j in idx.get(t,()):
                parent[find(('W',i))]=find(('M',j)); matchedW.add(i); matchedM.add(j)
    grp={}
    for i in matchedW: grp.setdefault(find(('W',i)),{'W':[],'M':[]})['W'].append(i)
    for j in matchedM: grp.setdefault(find(('M',j)),{'W':[],'M':[]})['M'].append(j)
    for g in grp.values():
        ws_=[W[i] for i in g['W']]; ms_=[M[j] for j in g['M']]
        bl=', '.join(sorted({t for w in ws_ for t in w['keys']})[:2]); ves=str(ws_[0]['vessel'] or '')[:35]
        wq={}; mq={}
        for w in ws_: wq[w['fam']]=wq.get(w['fam'],0)+w['qty']
        for m in ms_: mq[m['fam']]=mq.get(m['fam'],0)+m['qty']
        if set(wq)!=set(mq):
            add('In transit','Same shipment, different material','Low' if abs(sum(wq.values())-sum(mq.values()))<=2 else 'Medium',f"WSA: {', '.join(sorted(wq))} | MIS: {', '.join(sorted(mq))}",'','WSA qty',sum(wq.values()),'MIS qty',sum(mq.values()),
                f'BL {bl} ({ves}) - material coded differently in the two files','Agree one material code for this item')
        else:
            for f_,a in wq.items():
                b=mq.get(f_,0)
                if abs(a-b)>max(2,0.005*max(a,b)):
                    add('In transit','Same shipment, different quantity','High' if abs(a-b)>0.1*max(a,b) else 'Medium',f_,'','WSA qty',a,'MIS qty',b,f'BL {bl} ({ves})','Check the packing list / invoice quantity')
        wd=min(w['date'] for w in ws_); md=min((m['date'] for m in ms_ if m['date'] is not None),default=None)
        if md is not None:
            dd_=(md-pd.Timestamp(wd).normalize()).days
            if abs(dd_)>=3:
                add('In transit','Same shipment, different arrival date','Medium' if abs(dd_)>=7 else 'Low',', '.join(sorted(wq)),'','WSA plant date',None,'MIS Final AtA',None,
                    f"BL {bl} ({ves}): WSA {pd.Timestamp(wd):%d-%b} vs MIS {md:%d-%b} ({dd_:+d} days)",f"MIS is the latest (your rule): move the WSA row to {md:%d-%b}")
    for i,w in enumerate(W):
        if i not in matchedW:
            add('In transit','In WSA, not in MIS','High' if w['bl'] else 'Medium',w['name'],n2c.get(w['name'],''),'WSA qty',w['qty'],'MIS qty',0.0,
                f"Plant date {pd.Timestamp(w['date']):%d-%b}; PO {w['po'] or '-'}; BL {w['bl'] or '-'}; {str(w['vessel'] or '')[:35]}",
                'Ask MIS owner to add the shipment, or confirm it was received' if w['bl'] else 'No BL yet - confirm booking with supplier')
    for j,m in enumerate(M):
        if j not in matchedM:
            ap_=avgplan.get(m['code'],0)
            add('In transit','In MIS, not in WSA','High' if (ap_ and m['qty']>=ap_) else ('Medium' if m['qty']>=100 else 'Low'),m['mat'] or m['code'],m['code'],'WSA qty',0.0,'MIS qty',m['qty'],
                f"{str(m['sup'] or '')[:30]}; PO {m['po'] or '-'}; BL {', '.join(sorted(m['keys'])[:2]) or '-'}; AtA {m['date']:%d-%b}" if m['date'] is not None else f"PO {m['po']}",
                'Ask the planner to add it to the WSA sheet, or confirm it was received')
    # codes in MIS In Transit that MRP cannot place
    oth={}
    for m in M:
        if m['code'] not in c2n: oth[m['code']]=oth.get(m['code'],0)+m['qty']
    for c,q in oth.items():
        add('Master data','MIS In Transit code not on Daily Dashboard','High' if c.lower().startswith('no material') else 'Medium',c,c,'',None,'Pending qty',q,
            'Quantity cannot be counted in Daily Stock' + (" - MIS lookup failed (RM code missing)" if c.lower().startswith('no material') else ''),
            'Fix the RM code in MIS' if c.lower().startswith('no material') else 'Add the code to the dashboard, or map it to the right material')
    # ---------- OPEN PO ----------
    mis_po={}
    for m in M:
        if m['po']: mis_po.setdefault(str(m['po']).strip().removesuffix('.0'),set()).add(m['fam'])
    for d in BW.get('deliv',[]):
        if d['stage']!='Open PO' or not d['first']: continue
        po=str(d['po'] or '').strip().removesuffix('.0')
        if pd.Timestamp(d['date']).normalize()<today:
            add('Open PO','Open PO past its plant date','High',d['name'],n2c.get(d['name'],''),'',None,'Qty',d['qty'],f"Plant date {pd.Timestamp(d['date']):%d-%b}, still no vessel; PO {po or '-'}",'Get a new date from the supplier')
        elif po and po in mis_po and fam(d['name']) in mis_po[po]:
            add('Open PO','WSA Open PO already shipping in MIS?','Low',d['name'],n2c.get(d['name'],''),'',None,'Qty',d['qty'],
                f"PO {po} has shipments in MIS In Transit; this WSA line (plant date {pd.Timestamp(d['date']):%d-%b}) has no vessel yet",'Check whether this line is already booked (avoid double counting)')
    return I

# ===================== MIS layout: columns read by POSITION (checked every run, shown on Check) =====================
# In Transit is read by header name (normalize_in_transit). These sheets are read by column position, so a column MIS inserts
# or removes would silently shift the numbers. Every run: (1) the header in each position is compared with the last run
# (stored on the hidden 'MIS Layout' sheet of MRP_Daily), (2) the data in each column is checked for the right kind of value.
# (column index 0 = A, purpose, kind: date / num / code / text / any)
MIS_POSITIONS={
 'Total Plan Req':[(0,'Date','date'),(1,'Plant = production line','text'),(2,'Customer','text'),(3,'Modules per plan line (Value)','num'),
                   (4,'MIS BOM Name','text'),(10,'Material code','code'),(12,'Number used to select rows for the BOM rate check','any'),
                   (13,'Unit of measure','text'),(14,'Requirement quantity','num'),(17,'OA (plan line key)','any'),(18,'OEM Category','text'),(21,'Company (plant)','text')],
 'MB52':[(2,'Material code','code'),(5,'Material description','text'),(7,'Unit of measure','text'),(13,'Plant / storage (USA, USA WAREHOUSE)','text'),
         (14,'Total stock','num'),(18,'Material name','text')],
 'RM Status for OEM JW':[(0,'Make (supplier)','text'),(1,'Material code','code'),(5,'Stock USA','num'),(7,'Stock USA Warehouse','num'),(11,'JW WSA to OEM','num')],
 'Common Name':[(0,'Material code','code'),(1,'Material description','text'),(2,'Unit of measure','text'),(3,'Common Name','text')],
 'Standard BOM':[(2,'BOM Name','text'),(3,'Material code','code'),(5,'Quantity per module','num')]}
MIS_OPTIONAL={'Standard BOM','Common Name'}
IT_USED=[IT_CANON[i] for i in (1,4,6,7,8,12,23,24,31,34,35,39)]     # In Transit columns the script uses
_MCODE=re.compile(r'^[A-Z]{2,}[A-Z0-9]*\d{3,}$')
def _kind_ok(v,kind):
    if kind=='date': return isinstance(v,(dt.datetime,dt.date))
    if kind=='num': return isinstance(v,(int,float)) and not isinstance(v,bool)
    if kind=='code': return bool(_MCODE.match(str(v).strip().upper()))
    if kind=='text': return isinstance(v,str)
    return True
def mis_layout_check(mis_path, prev_path=None, sample=3000):
    """Returns (Check issues, rows for the Check section and the hidden 'MIS Layout' sheet)."""
    wb=wb_ro(mis_path); prev={}
    if prev_path and os.path.exists(prev_path):
        try:
            pw=wb_ro(prev_path)
            if 'MIS Layout' in pw.sheetnames:
                for r in list(pw['MIS Layout'].iter_rows(values_only=True))[1:]:
                    if r and r[0] and r[1] is not None: prev[(str(r[0]),str(r[1]))]=str(r[2] if r[2] is not None else '')
        except Exception: prev={}
    iss=[]; rows=[]
    def add(sev,check,sheet,col,a_lab,a,b_lab,b,detail,action):
        iss.append(dict(area='MIS layout',check=check,sev=sev,material=sheet,code=col,a_lab=a_lab,a=a,b_lab=b_lab,b=b,diff=None,detail=detail,action=action))
    for sheet,cols in MIS_POSITIONS.items():
        if sheet not in wb.sheetnames:
            if sheet not in MIS_OPTIONAL:
                add('High','MIS sheet missing',sheet,'','',None,'',None,f"Sheet '{sheet}' is not in {os.path.basename(mis_path)} - everything read from it is zero / blank",'Get the full MIS RM Dashboard file')
            continue
        data=wb[sheet]._rows; hdr=list(data[0]) if data else []; body=data[1:1+sample]
        for i,purpose,kind in cols:
            col=L(i+1); cur=str(hdr[i]).strip() if i<len(hdr) and hdr[i] is not None else ''
            old=prev.get((sheet,col)); vals=[r[i] for r in body if i<len(r) and r[i] not in (None,'')]
            ok_n=sum(1 for v in vals if _kind_ok(v,kind)); share=(ok_n/len(vals)) if vals else None
            moved=old is not None and old!=cur
            bad=kind!='any' and vals and len(vals)>=20 and share<0.8
            empty=kind!='any' and not vals and len(body)>=20
            res='OK'
            if moved: res='HEADER CHANGED'; add('High','MIS column header changed since the last run (column moved?)',sheet,col,'Header last run',old,'Header now',cur,
                    f"Column {col} is read as '{purpose}' by position. If MIS inserted or removed a column, every number from this sheet is wrong.",
                    'Check the MIS file layout; if the column moved, tell MIS to restore it (or update the script)')
            if bad: res=(res+' / ' if res!='OK' else '')+'WRONG DATA'; add('High','MIS column does not hold the expected data',sheet,col,'Expected',kind,'Matching values',f'{share:.0%}',
                    f"Column {col} (header '{cur}') should be '{purpose}' ({kind}) but only {share:.0%} of {len(vals)} values look right.",'Check the MIS file layout before using the results')
            if empty: res=(res+' / ' if res!='OK' else '')+'EMPTY'; add('Medium','MIS column read by position is empty',sheet,col,'Expected',purpose,'',None,
                    f"Column {col} (header '{cur}') has no values in {len(body)} rows.",'Check the MIS file')
            rows.append(dict(sheet=sheet,col=col,purpose=purpose,kind=kind,cur=cur,old=old,share=share,res=res))
    if IT_MISSING:
        used=[h for h in IT_MISSING if h in IT_USED]; other=[h for h in IT_MISSING if h not in IT_USED]
        if used: add('High','MIS In Transit: columns used by the script are missing','In Transit',', '.join(used),'',None,'',None,
                     'In Transit is read by header name; these headers were not found, so they are read as blank: '+', '.join(used),'Ask MIS to restore the header names')
        if other: add('Low','MIS In Transit: other columns not found','In Transit',', '.join(other)[:60],'',None,'',None,'Not used for the numbers: '+', '.join(other),'None needed')
    if not prev: add('Low','MIS layout recorded - compared from the next run','(all)','','',None,'',None,
                     'No previous MRP_Daily with a MIS Layout sheet: the headers read by position are stored now and compared on the next run. Data checks were done.','Check once that the headers below are the right columns')
    return iss,rows

def make_mis_layout_xlsx(rows, path, store=False):
    """store=False: the Check section 'MIS columns read by position'; store=True: the hidden 'MIS Layout' sheet kept for the next run."""
    from openpyxl.styles import Border, Side
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    wb=Workbook(); ws=wb.active
    if store:
        ws.title='MIS Layout'; ws.append(['Sheet','Column','Header'])
        for x in rows: ws.append([x['sheet'],x['col'],x['cur']])
        wb.save(path); return
    ws.title='MIS columns'
    ws['A1']='MIS columns read by position - header now vs last run, and whether the data looks right'; ws['A1'].font=Font(name=AR,size=13,bold=True,color=NAVY)
    ws['A2']=('These MIS sheets are read by column position (In Transit is read by header name). A header that changed, or a column whose values are not '
              'the expected kind, means MIS changed the layout and the numbers from that sheet cannot be trusted.')
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:H2'); ws['A2'].alignment=Alignment(wrap_text=True); ws.row_dimensions[2].height=28
    H=['Sheet','Column','Read as','Expected kind','Header now','Header last run','Values of the expected kind','Result']
    for j,h in enumerate(H,1):
        c=ws.cell(4,j,h); c.font=Font(name=AR,size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=NAVY); c.border=bd; c.alignment=Alignment(horizontal='center',wrap_text=True)
    for i,x in enumerate(rows):
        vals=[x['sheet'],x['col'],x['purpose'],x['kind'],x['cur'],x['old'] if x['old'] is not None else '(first run)',x['share'],x['res']]
        for j,v in enumerate(vals,1):
            c=ws.cell(5+i,j,v); c.font=Font(name=AR,size=9,bold=j==8); c.border=bd
            if j==7: c.number_format='0%'
        ws.cell(5+i,8).fill=PatternFill('solid',fgColor='C6EFCE' if x['res']=='OK' else 'FFC7CE')
    for j,w in enumerate([20,8,40,10,28,28,12,22],1): ws.column_dimensions[L(j)].width=w
    wb.save(path)

def make_check_xlsx(issues, DS, path):
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    AR='Calibri'; NAVY='1F3864'
    wb=Workbook(); ws=wb.active; ws.title='Check'
    hf=Font(name=AR,bold=True,color='FFFFFF',size=10); hfill=PatternFill('solid',fgColor=NAVY); bf=Font(name=AR,size=9)
    thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin); ctr=Alignment(horizontal='center',vertical='center',wrap_text=True)
    sevfill={'High':'FFC7CE','Medium':'FFEB9C','Low':'DDEBF7'}
    ws['A1']='Data check - discrepancies between MIS RM Dashboard, WSA planning file and MRP workbook'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=(f"Run {dt.date.today():%d-%b-%Y}. Areas: Stock, Plan (next 30 days), In transit (matched shipment by shipment on BL number), Open PO, Units, Master data. "
              "Severity: High = can change a shortage decision; Medium = should be fixed; Low = tidy-up (collapsed - click + on the left). "
              "MIS layout = MIS columns read by position (header moved / wrong data). Below: MIS columns, Segmentation, Unit check, Pack sizes, Plan compare (collapsed - click + to open).")
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws['A2'].alignment=Alignment(wrap_text=True,vertical='top'); ws.merge_cells('A2:K2'); ws.row_dimensions[2].height=30
    areas=['MIS layout','Stock','Plan','In transit','Open PO','Units','Master data']
    for j,t in enumerate(['Area','High','Medium','Low','Total']): c=ws.cell(4,2+j,t); c.font=hf; c.fill=hfill; c.alignment=ctr; c.border=bd
    for i,a in enumerate(areas+['Total']):
        r=5+i; c=ws.cell(r,2,a); c.font=Font(name=AR,size=10,bold=(a=='Total')); c.border=bd
        sel=[x for x in issues if a=='Total' or x['area']==a]
        for j,sv in enumerate(['High','Medium','Low',None]):
            n=sum(1 for x in sel if sv is None or x['sev']==sv); c=ws.cell(r,3+j,n); c.alignment=ctr; c.border=bd; c.font=Font(name=AR,size=10,bold=True)
            if sv and n: c.fill=PatternFill('solid',fgColor=sevfill[sv])
    top=5+len(areas)+3
    # ---- TOP ISSUES: the few things to fix this week (High first, biggest first), one line each ----
    order0={'High':0,'Medium':1,'Low':2}
    tops=sorted([x for x in issues if x['sev'] in ('High','Medium')],key=lambda x:(order0[x['sev']],rrank(str(x.get('code') or '').strip()),-(abs(x['diff']) if isinstance(x['diff'],(int,float)) else 0)))[:15]
    if tops:
        ws.cell(top-1,1,'TOP ISSUES - fix these first (all issues are listed below)').font=Font(name=AR,size=12,bold=True,color='C00000')
        for j,t in enumerate(['Severity','Area','Check','Material','Code']): c=ws.cell(top,1+j,t); c.font=hf; c.fill=hfill; c.alignment=ctr; c.border=bd
        for i,x in enumerate(tops):
            r=top+1+i
            for j,v in enumerate([x['sev'],x['area'],x['check'],x['material'],x['code']]):              # no Action column (user, Nov 2026: 'no help')
                c=ws.cell(r,1+j,v); c.font=bf; c.border=bd; c.alignment=Alignment(vertical='top',wrap_text=j in (2,3))
            ws.cell(r,1).fill=PatternFill('solid',fgColor=sevfill[x['sev']])
        top=top+len(tops)+4
        ws.cell(top-1,1,'ALL ISSUES').font=Font(name=AR,size=12,bold=True,color=NAVY)
    hdr=['Severity','Area','Check','Material','Code','Source A','Value A','Source B','Value B','Difference (B - A)','Detail']
    for j,t in enumerate(hdr): c=ws.cell(top,1+j,t); c.font=hf; c.fill=hfill; c.alignment=ctr; c.border=bd
    order={'High':0,'Medium':1,'Low':2}
    issues=sorted(issues,key=lambda x:(order[x['sev']],rrank(str(x.get('code') or '').strip()),areas.index(x['area']) if x['area'] in areas else 9,-(abs(x['diff']) if isinstance(x['diff'],(int,float)) else 0)))   # severity, then highest shortage risk
    for i,x in enumerate(issues):
        r=top+1+i
        vals=[x['sev'],x['area'],x['check'],x['material'],x['code'],x['a_lab'] or None,x['a'],x['b_lab'] or None,x['b'],x['diff'],x['detail']]
        for j,v in enumerate(vals,1):
            c=ws.cell(r,j,round(v,3) if isinstance(v,float) else v); c.font=bf; c.border=bd; c.alignment=Alignment(vertical='top',wrap_text=j in (3,4,11))
            if j in (7,9,10): c.number_format='#,##0;-#,##0;0'
        ws.cell(r,1).fill=PatternFill('solid',fgColor=sevfill[x['sev']]); ws.cell(r,1).font=Font(name=AR,size=9,bold=True)
    last=top+max(1,len(issues))
    if not issues: ws.cell(top+1,1,'No discrepancies found')
    ws.add_table(Table(displayName='DataCheck',ref=f'A{top}:K{last}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=True)))
    for col,w in zip('ABCDEFGHIJK',[9,11,30,30,16,16,12,16,12,13,70]): ws.column_dimensions[col].width=w
    ws.freeze_panes=f'A{top+1}'; ws.sheet_view.zoomScale=90
    set_print(wb.active,'Check'); lows=[r_ for r_ in range(1,ws.max_row+1) if ws.cell(r_,1).value=='Low']
    if lows:                                        # Low = tidy-up: collapsed, click + on the left to open
        ws.row_dimensions.group(lows[0],lows[-1],hidden=True,outline_level=1)
        ws.sheet_properties.outlinePr.summaryBelow=False
    wb.save(path)
    return {s_:sum(1 for x in issues if x['sev']==s_) for s_ in ('High','Medium','Low')}

# ===================== CS / PM in the new MRP: user-maintained BOM + Stock Count =====================
BOM_COLS=['Customer','Module Model','Item','Ict','Component','Component desciption','Quantity','UoM','Valid From','Valid to','MIS BOM Name']
def _guess_mis_bom(model, mis_names):
    m=str(model or '').upper()
    for n in mis_names:
        u=n.upper()
        if 'EPEP' in m and 'EPEP' in u: return n
        if ('BIN-21' in m or '625' in m) and u.endswith('G12R_USA'): return n
    return ''

def fill_count_bom(hdr, rows, path):
    """A workbook last saved by a tool other than Excel keeps formulas but no results: Today Qty_BOM then reads as empty.
       Recalculate it as Today Stock Qty x Quantity so packaging stock is never 0 by mistake."""
    hl=[str(h).strip().lower() if h is not None else '' for h in (hdr or [])]
    if 'today qty_bom' not in hl or 'today stock qty' not in hl: return 0
    io,iq=hl.index('today qty_bom'),hl.index('today stock qty'); i_f=hl.index('quantity') if 'quantity' in hl else None
    n=0
    for r in rows:
        if len(r)>max(io,iq) and not isinstance(r[io],(int,float)) and isinstance(r[iq],(int,float)):
            f=r[i_f] if (i_f is not None and len(r)>i_f and isinstance(r[i_f],(int,float))) else 1.0
            r[io]=r[iq]*f; n+=1
    path=PPC_SRC[0] or path
    if n: print(f'NOTE: {n} Count rows in {os.path.basename(path)} had no saved Today Qty_BOM (file last saved outside Excel) - '
                f'calculated as Today Stock Qty x Quantity. Open {os.path.basename(path)} in Excel and save it once to restore its formula results.')
    return n

def read_user_inputs(prev_daily, old_mrp, mis_names, model2bn=None):
    """BOM and Stock Count from YOUR MRP workbook (MRP.xlsx) - it is the master. The previous MRP_Daily is used only when your
       MRP has no such sheet. MIS BOM Name (link to the MIS plan) is not on your BOM: it is taken from the BOM sheet column
       if you add one, else from your Production Plan (customers using that Module Model), else the previous MRP_Daily, else a guess."""
    bom=[]; count_hdr=None; count=[]; src=None; prev_map={}
    model2bn=model2bn or {}
    def load(path):
        if not path or not os.path.exists(path): return None
        try: return wb_ro(path)
        except Exception as _e_: _swallow(3115,_e_); return None
    wp=load(prev_daily)
    if wp is not None and 'BOM' in wp.sheetnames:          # MIS BOM Name chosen earlier for each Module Model
        rows=[list(r) for r in wp['BOM'].iter_rows(values_only=True)]
        hr=next((i for i,r in enumerate(rows) if 'Component' in r and 'Quantity' in r),None)
        if hr is not None and 'MIS BOM Name' in rows[hr]:
            im,ib=rows[hr].index('Module Model'),rows[hr].index('MIS BOM Name')
            for r in rows[hr+1:]:
                if r[im] and r[ib]: prev_map.setdefault(str(r[im]).strip(),r[ib])
    for path in (old_mrp,prev_daily):
        wb=load(path)
        if wb is None: continue
        if 'BOM' in wb.sheetnames and not bom:
            rows=[list(r) for r in wb['BOM'].iter_rows(values_only=True)]
            std_=bom_rows_std(rows)                                  # old or new column layout
            if std_ is not None:
                for d0_ in std_:
                    d={c:d0_.get(c) for c in BOM_COLS}; d['Material group']=d0_.get('Material group')
                    if not d['Component']: continue
                    mdl=str(d['Module Model'] or '').strip()
                    if not d['MIS BOM Name']: d['MIS BOM Name']=model2bn.get(mdl) or prev_map.get(mdl) or _guess_mis_bom(mdl,mis_names)
                    bom.append(d)
                src=src or path
        if 'Stock Count' in wb.sheetnames and count_hdr is None and PPC_SRC[0] and path==old_mrp:   # ONLY the PPC 'Count' sheet
            hdr_,rws_=count_header_and_rows(wb['Stock Count'].iter_rows(values_only=True))
            if hdr_ is not None:
                n_=max(j for j,x in enumerate(hdr_) if x is not None)+1
                count_hdr=hdr_[:n_]; count=[list(r[:n_])+[None]*max(0,n_-len(r)) for r in rws_ if r and r[0]]
                hl_=[str(h_ or '').strip().lower() for h_ in count_hdr]
                # your names: 'Physical Count Qty' (C) = Today Stock Qty, 'Physical Count Qty_BOM' (G) = Today Qty_BOM (same columns, new names)
                ren_={'physical count qty':'Today Stock Qty','physical count qty_bom':'Today Qty_BOM'}
                count_hdr=[ren_.get(h_,count_hdr[i_]) if not (h_=='physical count qty' and 'today stock qty' in hl_) and not (h_=='physical count qty_bom' and 'today qty_bom' in hl_) else count_hdr[i_] for i_,h_ in enumerate(hl_)]
                hl_=[str(h_ or '').strip().lower() for h_ in count_hdr]
                fill_count_bom(count_hdr,count,path)                       # Today Qty_BOM = count x Quantity when Excel's result is missing
                ipc=hl_.index('today qty_bom') if 'today qty_bom' in hl_ else (6 if len(hl_)>6 else 2)      # column G: Physical Count Qty_BOM
                PHYS_COUNT.clear(); PHYS_DATE.clear()
                idt=next((i_ for i_,h_ in enumerate(hl_) if h_ in ('physical count date','count date')),None)    # your 'Physical Count Date' column (any position)
                for r_ in count:
                    c_=str(r_[0] or '').strip(); v_=r_[ipc] if ipc<len(r_) else None
                    if c_ and isinstance(v_,(int,float)) and not isinstance(v_,bool): PHYS_COUNT[c_]=PHYS_COUNT.get(c_,0.0)+float(v_)
                    if c_ and idt is not None and idt<len(r_) and isinstance(r_[idt],(dt.datetime,dt.date)):
                        PHYS_DATE[c_]=max(PHYS_DATE.get(c_,r_[idt]),r_[idt])
                pass
    return bom,count_hdr,count,src

MODPLAN_BY_PLANT={}    # (MIS BOM Name, customer, Plant line) -> modules per day, for the Module Plan sheet
def module_plan(mis_path, plant, dates):
    """Modules per day by MIS BOM Name and customer, from MIS Total Plan Req ('Value', one per line / customer / BOM / day)."""
    wb=wb_ro(mis_path); idx={pd.Timestamp(d).normalize():k for k,d in enumerate(dates)}
    seen=set(); mp={}; MODPLAN_BY_PLANT.clear()
    for r in wb['Total Plan Req'].iter_rows(min_row=2,values_only=True):
        if not tp_ok(r,plant) or not isinstance(r[0],(dt.datetime,dt.date)) or not isinstance(r[3],(int,float)): continue
        k=idx.get(pd.Timestamp(r[0]).normalize())
        if k is None: continue
        key=(r[0],r[1],r[2],r[4],r[17] if len(r)>17 else None)   # one plan line = date + production line + customer + BOM + OA
        if key in seen: continue
        seen.add(key); bn=str(r[4] or '').strip(); cu=' '.join(str(r[2] or '').split())
        mp.setdefault(bn,{}).setdefault(cu,[0.0]*len(dates))[k]+=r[3]
        pl=' '.join(str(r[1] or '').split())                    # MIS 'Plant' = production line (L1..L4, Arizona, OEM lines)
        MODPLAN_BY_PLANT.setdefault((bn,cu,pl),[0.0]*len(dates))[k]+=r[3]
    return mp

def cspm_from_bom(bom, count, modplan, dates, sr0, count_hdr=None, lineplan=None):
    """One block per CS / PM component: plan = average BOM qty per module (per MIS BOM Name) x modules that day."""
    nd=len(dates); per={}
    for d in bom:
        c=str(d['Component'] or '').strip(); bn=str(d['MIS BOM Name'] or '').strip()
        if c[:2].upper() not in ('CS','PM') or not isinstance(d['Quantity'],(int,float)): continue
        x=per.setdefault(c,dict(desc=d['Component desciption'],uom=d['UoM'],q={}))
        if bn:
            mq=x['q'].setdefault(bn,{}); mdl=str(d['Module Model'] or '').strip()
            mq[mdl]=mq.get(mdl,0.0)+float(d['Quantity'])        # rows of one model add up; models are averaged below
    qi=count_qty_idx(count_hdr)
    stock={str(r[0]).strip():(float(r[qi]) if len(r)>qi and isinstance(r[qi],(int,float)) else 0.0) for r in count if r and r[0]}
    hl=[str(h).strip().lower() if h is not None else '' for h in (count_hdr or [])]
    ix=lambda n: hl.index(n) if n in hl else None
    i_q,i_u,i_f=ix('today stock qty'),ix('stock uom'),ix('quantity')
    raw={}
    for r in count:
        if not r or not r[0] or i_q is None: continue
        qv=r[i_q] if len(r)>i_q else None; u=r[i_u] if (i_u is not None and len(r)>i_u) else ''; fv=r[i_f] if (i_f is not None and len(r)>i_f) else None
        if isinstance(qv,(int,float)):
            raw[str(r[0]).strip()]=f"{qv:,.0f} {u}"+(f" x {fv:g}" if isinstance(fv,(int,float)) and fv not in (0,1) else '')
    out=[]; sr=sr0
    for c in sorted(per):
        x=per[c]; plan=[0.0]*nd; cust={}; cl={}; pls=[]
        for bn,mq in x['q'].items():
            q=sum(mq.values())/len(mq)
            for cu,v in (modplan.get(bn) or {}).items():
                arr=cust.setdefault(cu,[0.0]*nd)
                for k in range(nd): plan[k]+=q*v[k]; arr[k]+=q*v[k]
                for ln,lv in (((lineplan or {}).get(bn) or {}).get(cu) or {None:v}).items():
                    a_=cl.setdefault((cu,ln),[0.0]*nd)
                    for k in range(nd): a_[k]+=q*lv[k]
                    pls.append(dict(cu=cu,ln=ln,bn=bn,rate=q,mods=list(lv),src=f'your BOM {bn} (average of {len(mq)} model(s))'))
        sr+=1; desc=' '.join(str(x['desc'] or '').split())
        out.append(dict(key=f'{sr}CS/PM',sr=sr,short='CS - Consumable' if c.upper().startswith('CS') else 'PM - Packaging',common=desc,mat=c,desc=desc,mtype='CS/PM',
                        stock=stock.get(c,0.0),uom=x['uom'] or '',plan=plan,cust_plan=dict(sorted({k:v for k,v in cust.items() if any(v)}.items(),key=lambda kv:-sum(kv[1]))),
                        cust_line=dict(sorted({k:v for k,v in cl.items() if any(v)}.items(),key=lambda kv:(-sum(cust.get(kv[0][0],[0])),kv[0][0],str(kv[0][1])))),
                        transit=[0.0]*nd,adhoc=[0.0]*nd,openpo=[0.0]*nd,count_stock=True,raw=raw.get(c,'not on Stock Count'),plan_lines=pls,
                        st_main=stock.get(c,0.0),st_wh=0.0,st_jw=0.0,plan_all={pd.Timestamp(dates[k]).normalize():plan[k] for k in range(nd) if plan[k]}))
    return out


# ===================== YOUR BOM (MRP.xlsx 'BOM' sheet) - new layout =====================
BOM_HDR_MAP={'customer name':'Customer','customer':'Customer','material code':'Component','component':'Component',
             'material description':'Component desciption','component desciption':'Component desciption','component description':'Component desciption',
             'per module consumption':'Quantity','quantity':'Quantity','uom':'UoM','module model':'Module Model','material group':'Material group',
             'module type':'Module type','valid from':'Valid From','valid to':'Valid to','mis bom name':'MIS BOM Name',
             'fg material':'FG Material','material':'Component','description':'Component desciption','qty per module':'Quantity','line #':'Line #',
             'customers':'Customer group','supplier code':'Supplier Code'}
PP_MODELS={}       # Production Plan customer -> Module Model (to link the BOM's FG Material / customer group to your customers)
def _model_prefix(m):
    m=str(m or '').strip().upper(); i=m.find('-MC4')
    return m[:i+4] if i>0 else m[:22]
def load_pp_models(mrp_path):
    PP_MODELS.clear()
    try:
        for r in pp_rows_std(wb_ro(mrp_path)['Production Plan']):
            if r and r[3] and len(r)>7 and isinstance(r[7],str) and r[7].upper().startswith('BIN'): PP_MODELS.setdefault(' '.join(str(r[3]).split()),r[7])
    except Exception as _e_: _swallow(3235,_e_); pass
def bom_rows_std(rows):
    """Rows of your BOM sheet with standard field names, old layout (Customer / Component / Quantity) or new layout
       (Customer Name / Material Code / Per module Consumption). Returns (list of dicts) or None when no header is found."""
    for i,r in enumerate(rows[:10]):
        hl=[' '.join(str(x or '').split()).lower() for x in r]
        if ('component' in hl and 'quantity' in hl) or ('material code' in hl and 'per module consumption' in hl) or ('material' in hl and 'qty per module' in hl):
            idx={}
            for j,h in enumerate(hl):
                k=BOM_HDR_MAP.get(h)
                if k and k not in idx: idx[k]=j
            out=[]
            for r2 in rows[i+1:]:
                d={k:(r2[j] if j<len(r2) else None) for k,j in idx.items()}
                if not d.get('Component'): continue
                d['Component']=str(d['Component']).strip()
                if isinstance(d.get('Customer'),str): d['Customer']=' '.join(d['Customer'].split())
                if d.get('Customer') or not d.get('FG Material'): out.append(d); continue
                # PPC 'Order' layout: one row per FG Material -> every Production Plan customer with that module model
                if not PP_MODELS and SETTINGS_SRC and SETTINGS_SRC[0]: load_pp_models(SETTINGS_SRC[0])
                pre=_model_prefix(d['FG Material']); hit=False
                for cu,mdl in PP_MODELS.items():
                    if _model_prefix(mdl)==pre: e=dict(d); e['Customer']=cu; e['Module Model']=mdl; out.append(e); hit=True
                if not hit: e=dict(d); e['Customer']=d.get('Customer group') or d['FG Material']; e['Module Model']=d['FG Material']; out.append(e)
            return out
    return None

UBOM_DEMAND=True   # False: raw-material demand = MIS Standard BOM; your BOM (MRP.xlsx) = ALTERNATIVE BOM, used when a standard material is short
UBOM_RM={}         # your BOM, raw materials: normalised customer -> {code: [per-module quantities]} (a code listed twice = two layers, added up)
UBOM_ALT={}        # customer -> {alternative code: its standard code} from your BOM 'Line #' = '<standard>_Alt'
def load_user_bom_rm(mrp_path):
    """Raw-material lines of your BOM (codes RM...), for 'your BOM first, MIS BOM if the customer is not in it'."""
    UBOM_RM.clear(); UBOM_ALT.clear(); META={}
    try:
        rows=[list(r) for r in wb_ro(mrp_path)['BOM'].iter_rows(values_only=True)]
    except Exception as _e_: _swallow(3270,_e_); return META
    for d in bom_rows_std(rows) or []:
        c=d['Component']
        if not c.upper().startswith('RM') or not isinstance(d.get('Quantity'),(int,float)) or not d.get('Customer'): continue
        ln_=str(d.get('Line #') or '').strip()
        if ln_.lower().endswith('_alt'):                                 # your BOM marks an alternative: Line # = '<standard code>_Alt'
            UBOM_ALT.setdefault(_norm(d['Customer']),{})[c]=ln_[:-4].strip()
        UBOM_RM.setdefault(_norm(d['Customer']),{}).setdefault(c,[]).append((float(d['Quantity']),str(d.get('UoM') or '').strip().upper()))
        META.setdefault(c,dict(desc=d.get('Component desciption'),uom=d.get('UoM'),group=d.get('Material group'),mtype=d.get('Module type')))
    return META

PACK_KG={}         # material (common name) -> KG per pack from the WSA sheet unit, e.g. 'DRUM(270KG/DRUM)' -> 270
UBOM_UNIT_ISSUES=[]   # (customer, code, your unit, MRP unit) that could not be converted
def ubom_rate(ucust_norm, block_code, alias, factor, buom=None, bname=None):
    """Your BOM usage per module of a planned material for one customer. Codes of your BOM that are alternates of the same
       material (alternate codes / roll lengths, e.g. Busbar 4 x 0.4 and reflective 4 x 0.36) count ONCE (the largest), a code
       listed twice (front + back layer) counts twice. None = customer not in your BOM."""
    rows=UBOM_RM.get(ucust_norm) if UBOM_DEMAND else None
    if rows is None: return None
    best=0.0; bu=str(buom or '').strip().upper()
    alts_={a for std_,al_ in ALT_BOM.items() for a,cs_ in al_.items()     # an alternative carries no demand ONLY when this customer's BOM
           if any(_norm(x)==ucust_norm for x in cs_) and std_ in rows}         # also lists its standard code; listed alone it IS the demand
    alts_|={a for a,std_ in (UBOM_ALT.get(ucust_norm) or {}).items() if std_ in rows}     # your BOM's '_Alt' rows
    for c,qs in rows.items():
        if alias.get(c,c)!=block_code or c in alts_: continue
        tot=0.0
        for q,u in qs:
            if bu and u and u!=bu and not (c!=block_code and c in factor):     # your unit differs from the MRP unit
                if u=='KG' and PACK_KG.get(bname): q=q/PACK_KG[bname]          # e.g. RTV 0.39 KG per module / 270 KG per drum
                elif (ucust_norm,c,u,bu) not in UBOM_UNIT_ISSUES: UBOM_UNIT_ISSUES.append((ucust_norm,c,u,bu))
            tot+=q
        best=max(best,tot*(factor.get(c,1.0) if c!=block_code else 1.0))
    return best


def ubom_checks(mrp_path, mis_path):
    """Your BOM (raw materials) vs MIS master data: a code whose description has numbers (size, GSM, thickness...) that the MIS
       description of that code does not have = probably the wrong code; units your BOM uses that could not be converted."""
    I=[]; cn={}
    try:
        for r in wb_ro(mis_path)['Common Name'].iter_rows(min_row=2,values_only=True):
            if r and r[0]: cn[str(r[0]).strip()]=(r[1],r[3])
        rows=[list(r) for r in wb_ro(mrp_path)['BOM'].iter_rows(values_only=True)]
    except Exception as _e_: _swallow(3313,_e_); return I
    nums=lambda t: set(re.findall(r'\d+(?:\.\d+)?',str(t or '').upper().replace(',','.')))
    seen=set()
    for d in bom_rows_std(rows) or []:
        c=d['Component']
        if not c.upper().startswith('RM') or c not in cn: continue
        mine=nums(d.get('Component desciption')); mis_=nums(cn[c][0])
        extra=sorted(mine-mis_,key=lambda x:float(x))
        if extra and (c,tuple(extra)) not in seen:
            seen.add((c,tuple(extra))); cus=sorted({' '.join(str(x.get('Customer') or '').split()) for x in (bom_rows_std(rows) or []) if x['Component']==c and nums(x.get('Component desciption'))==mine})
            I.append(dict(area='Your BOM',check='Code and description do not match (MIS master)',sev='High',material=cn[c][1] or c,code=c,
                          a_lab='Your BOM description',a=' '.join(str(d.get('Component desciption') or '').split()),b_lab='MIS description of this code',b=' '.join(str(cn[c][0] or '').split()),diff=None,
                          detail=f"{', '.join(cus)}: your BOM has {', '.join(extra)} in the description, the MIS description of {c} does not - the code may be wrong. The MRP uses the CODE.",
                          action='Correct the code (or the description) in the BOM sheet of MRP.xlsx'))
    for cu,c,u,bu in UBOM_UNIT_ISSUES:
        I.append(dict(area='Your BOM',check='Unit not converted',sev='High',material=c,code=c,a_lab='Your BOM unit',a=u,b_lab='MRP unit',b=bu,diff=None,
                      detail=f'{cu}: your BOM gives {c} per module in {u}, the MRP counts {bu} - used without conversion',
                      action='Put the pack size on the WSA sheet unit (e.g. DRUM(270KG/DRUM)) or give the BOM in the MRP unit'))
    return I


ALT_BOM={}         # standard code -> {alternative code: set(customers)} from your alternative BOM
ALT_EXTRA={}       # code only in your BOM (no standard of that group) -> set(customers)
def build_alt_bom(mis_path, alias):
    """Your BOM = alternative BOM. For each customer: codes of your BOM that are not in the MIS Standard BOM of its module are
       alternatives of the standard code of the same material group (your Material Group, else the MIS group); several standard
       codes in one group -> the one the alternate is linked to (alias / same Common Name). No standard code of that group -> extra."""
    ALT_BOM.clear(); ALT_EXTRA.clear()
    try: rows=list(wb_ro(mis_path)['Standard BOM'].iter_rows(values_only=True))[1:]
    except Exception as _e_: _swallow(3342,_e_); return
    sb={}; mgrp={}
    for r in rows:
        if r and len(r)>8 and r[2] and r[3]:
            sb.setdefault(str(r[2]).strip(),{})[str(r[3]).strip()]=str(r[8] or ''); mgrp.setdefault(str(r[3]).strip(),str(r[8] or ''))
    def grp(c,mis_g=''):
        g_=mis_g or mgrp.get(c)                                 # MIS Standard BOM group of the code (any module) first
        if g_: return BOM_GROUP_MAP.get(str(g_).strip().lower(),str(g_).strip())
        ug=(MAT_GROUP.get(c) or (None,None))[1]
        if ug: return USER_GROUP_MAP.get(ug.lower(),ug.title())
        return BOM_GROUP_MAP.get(str(mis_g).strip().lower(),str(mis_g).strip())
    for cu,codes in UBOM_RM.items():
        L_=next((v for k,v in LINKS.items() if _norm(k)==cu),None); bn=(L_ or {}).get('bom')
        std=sb.get(str(bn or '').strip())
        if not std: continue
        sg={}
        for c,g in std.items(): sg.setdefault(grp(c,g),[]).append(c)
        name=next((k for k in LINKS if _norm(k)==cu),cu)
        for c in codes:
            if c in std or grp(c)=='Cell': continue
            if alias.get(c) in std and c in (BOM_ALIAS[1] or {}): continue    # same material, other roll length / pack (converted) - not an alternative
            cand=sg.get(grp(c)) or []
            if alias.get(c) in std: tgt=[alias[c]]
            elif len(cand)==1: tgt=cand
            elif cand: tgt=[x for x in cand if x not in codes] or cand
            else: ALT_EXTRA.setdefault(c,set()).add(name); continue
            for t in tgt: ALT_BOM.setdefault(t,{}).setdefault(c,set()).add(name)


RM_USAGE_FROM_YOUR_BOM=True   # raw-material usage per module from your BOM (MRP.xlsx) for the customers it lists, else MIS Standard BOM
BOM_ALIAS=[{},{}]  # alias / factor of this run (for your-BOM usage)
STD_RATE={}        # (planned material, MIS BOM name) -> usage per module from the MIS Standard BOM (layers added, alternate codes converted)
CS_RATE={}         # (normalised customer, CS/PM code) -> usage per module from your BOM;  ('model', Module Model, code) -> same by model
def load_rates(mis_path, mrp_path, alias, factor):
    """Usage per module straight from the BOMs (no requirement / module ratios)."""
    STD_RATE.clear(); CS_RATE.clear(); BOM_ALIAS[0]=alias; BOM_ALIAS[1]=factor
    try:
        for r in list(wb_ro(mis_path)['Standard BOM'].iter_rows(values_only=True))[1:]:
            if r and len(r)>5 and r[2] and r[3] and isinstance(r[5],(int,float)):
                c=str(r[3]).strip(); m=alias.get(c,c); f=factor.get(c,1.0) if m!=c else 1.0
                k=(m,str(r[2]).strip()); STD_RATE[k]=STD_RATE.get(k,0.0)+float(r[5])*f
    except Exception as _e_: _swallow(3383,_e_); pass
    try:
        for d in bom_rows_std([list(r) for r in wb_ro(mrp_path)['BOM'].iter_rows(values_only=True)]) or []:
            c=d['Component']
            if c[:2].upper() not in ('CS','PM') or not isinstance(d.get('Quantity'),(int,float)): continue
            k=(_norm(d.get('Customer')),c); CS_RATE[k]=CS_RATE.get(k,0.0)+float(d['Quantity'])
            mdl=str(d.get('Module Model') or '').strip()
            if mdl: CS_RATE.setdefault(('model',mdl,c),0.0); CS_RATE[('model',mdl,c)]=max(CS_RATE[('model',mdl,c)],CS_RATE[k])
    except Exception as _e_: _swallow(3391,_e_); pass

def bom_usage_src(mat, cs, pp_row):
    """(usage per module, source) - source: 'your BOM' (MRP.xlsx), 'MIS Standard BOM', or '' when not used."""
    u=bom_usage(mat,cs,pp_row)
    if cs: return u,('your BOM' if u else '')
    if UBOM_RM.get(_norm(pp_row.get('cust'))) is not None and RM_USAGE_FROM_YOUR_BOM: return u,('your BOM' if u else '')
    return u,('MIS Standard BOM' if u else '')

BLOCK_INFO={}      # material -> (MRP unit, common name) for unit conversion of BOM usage
def bom_usage(mat, cs, pp_row):
    """Usage per module of a material for one Production Plan row (customer / module type), straight from the BOM."""
    if cs:
        return CS_RATE.get((_norm(pp_row.get('cust')),mat)) or CS_RATE.get(('model',str(pp_row.get('model') or '').strip(),mat)) or 0.0
    rows=UBOM_RM.get(_norm(pp_row.get('cust')))
    if rows is not None and RM_USAGE_FROM_YOUR_BOM:            # your BOM lists this customer: use it (PPC way)
        cu_n=_norm(pp_row.get('cust'))
        alts={a for std_,alts_ in ALT_BOM.items() for a,cs_ in alts_.items()     # alternatives for THIS customer, only when its BOM also lists
              if any(_norm(x)==cu_n for x in cs_) and std_ in rows}            # the standard code (listed alone, the alternative carries the demand)
        alts|={a for a,std_ in (UBOM_ALT.get(cu_n) or {}).items() if std_ in rows}         # your BOM's '_Alt' rows
        best=0.0; bu,bname=BLOCK_INFO.get(mat,('',None)); bu=str(bu or '').strip().upper()
        for c,qs in rows.items():
            if c in alts or BOM_ALIAS[0].get(c,c)!=mat: continue
            tot=0.0
            for q,u in qs:                                          # same unit conversion as the plan (ubom_rate): e.g. RTV KG per module -> drums
                if bu and u and u!=bu and not (c!=mat and c in (BOM_ALIAS[1] or {})) and u=='KG' and PACK_KG.get(bname): q=q/PACK_KG[bname]
                tot+=q
            best=max(best,tot*(BOM_ALIAS[1].get(c,1.0) if c!=mat else 1.0))
        return best
    return STD_RATE.get((mat,str(pp_row.get('bom') or '').strip()),0.0)

def rm_plan_from_mis(DS, mis_path, modplan, lineplan=None):
    """Raw-material plan. RM USAGE ALWAYS COMES FROM MIS (your BOM is used for CS / PM only):
         plan = MIS usage per module (Total Plan Req: requirement / modules, all dates) x YOUR modules (Production Plan, via modplan).
       Only for MIS customers of a BOM Name that your Production Plan does not have is the MIS requirement (Total Plan Req) used as it is.
       Returns new blocks (copies of DS['blocks'] with plan, cust_plan, cust_line, plan_all and plan_basis set)."""
    import copy
    dates=DS['dates']; nd=len(dates); idx={pd.Timestamp(d).normalize():k for k,d in enumerate(dates)}
    wb=wb_ro(mis_path); req={}; rate_t={}; seen_=set(); reql={}; lineplan=lineplan or {}
    for r in wb['Total Plan Req'].iter_rows(min_row=2,values_only=True):
        if not tp_ok(r,DS['plant']) or not isinstance(r[0],(dt.datetime,dt.date)) or not isinstance(r[14],(int,float)): continue
        c=str(r[10] or '').strip(); bn=str(r[4] or '').strip(); cu=' '.join(str(r[2] or '').split())
        if isinstance(r[3],(int,float)) and r[3]:          # MIS usage per module = requirement / modules (all dates)
            t_=rate_t.setdefault((c,bn),[0.0,0.0]); t_[0]+=r[14]
            pl_=(c,bn,r[0],r[1],r[2],r[17] if len(r)>17 else None)
            if pl_ not in seen_: seen_.add(pl_); t_[1]+=r[3]
        k=idx.get(pd.Timestamp(r[0]).normalize())
        if k is None: continue
        req.setdefault(c,{}).setdefault(bn,{}).setdefault(cu,[0.0]*nd)[k]+=r[14]
        ln_=' '.join(str(r[1] or '').split()); reql.setdefault((c,bn,cu),{}).setdefault(ln_,[0.0]*nd)[k]+=r[14]
    alias=DS.get('alias') or {}; fac=DS.get('alias_factor') or {}
    load_rates(mis_path,SETTINGS_SRC[0],alias,fac)
    mrate={k_:(a/b_) for k_,(a,b_) in rate_t.items() if b_}
    for k_,v_ in STD_RATE.items(): mrate[k_]=v_                 # usage per module = MIS Standard BOM quantity (not requirement / modules)
    ucu={}                                       # MIS customer -> your customer (normalised) that is in your BOM
    for cu_k,L_ in LINKS.items():
        if _norm(cu_k) in UBOM_RM: ucu.setdefault(_norm(L_.get('mis') or cu_k),_norm(cu_k)); ucu.setdefault(_norm(cu_k),_norm(cu_k))
    for k_ in UBOM_RM: ucu.setdefault(k_,k_)
    out=[]
    for b0 in DS['blocks']:
        b=copy.copy(b0); c=b['mat']
        plan=[0.0]*nd; cust={}; cl={}; used_rate=[]; used_mis=[]; pls=b['plan_lines']=[]     # plan_lines: ALERT Plan lines (formulas sheet)
        def addl(cu,ln,v,f=1.0):
            a_=cl.setdefault((cu,ln or None),[0.0]*nd)
            for k in range(nd): a_[k]+=f*v[k]
        bns_=set((req.get(c) or {}).keys())|{bn_ for bn_,cs_ in modplan.items() if any(ubom_rate(ucu.get(_norm(cu_)),c,alias,fac,b0.get('uom'),b0.get('common')) for cu_ in cs_)}
        for bn in sorted(bns_):                              # MIS BOM Names of this material + those where your BOM uses it
            q=mrate.get((c,bn)); covered=set()
            if modplan.get(bn) and (q is not None or any(ubom_rate(ucu.get(_norm(cu_)),c,alias,fac,b0.get('uom'),b0.get('common')) for cu_ in modplan[bn])):
                for cu,v in modplan[bn].items():
                    qu=ubom_rate(ucu.get(_norm(cu)),c,alias,fac,b0.get('uom'),b0.get('common'))          # YOUR BOM first (None = customer not in your BOM)
                    q_=qu if qu is not None else q
                    covered.add(_norm(cu))
                    if not q_: continue
                    tag_=(f'your BOM ({cu})' if qu is not None else f'MIS {bn}')
                    if tag_ not in used_rate: used_rate.append(tag_)
                    arr=cust.setdefault(cu,[0.0]*nd)
                    for k in range(nd): plan[k]+=q_*v[k]; arr[k]+=q_*v[k]
                    lines_=(lineplan.get(bn) or {}).get(cu)
                    if lines_:
                        for ln,lv in lines_.items(): addl(cu,ln,lv,q_); pls.append(dict(cu=cu,ln=ln,bn=bn,rate=q_,mods=list(lv),src=tag_))
                    else: addl(cu,None,v,q_); pls.append(dict(cu=cu,ln=None,bn=bn,rate=q_,mods=list(v),src=tag_))
            # MIS customers of this BOM Name that your Production Plan does not have: MIS requirement as it is
            for cu,v in ((req.get(c) or {}).get(bn) or {}).items():
                if _norm(cu) in covered or not any(v): continue
                if bn not in used_mis: used_mis.append(bn)
                arr=cust.setdefault(cu,[0.0]*nd)
                for k in range(nd): plan[k]+=v[k]; arr[k]+=v[k]
                for ln,lv in (reql.get((c,bn,cu)) or {None:v}).items(): addl(cu,ln,lv); pls.append(dict(cu=cu,ln=ln,bn=bn,rate=1.0,mods=list(lv),src=f'MIS requirement {bn} (quantity, not modules)'))
        b['plan']=plan; b['cust_plan']=dict(sorted({k:v for k,v in cust.items() if any(v)}.items(),key=lambda kv:-sum(kv[1])))
        b['cust_line']=dict(sorted({k:v for k,v in cl.items() if any(v)}.items(),key=lambda kv:(-sum(cust.get(kv[0][0],[0])),kv[0][0],str(kv[0][1]))))
        b['plan_all']={pd.Timestamp(dates[k]).normalize():plan[k] for k in range(nd) if plan[k]}
        parts=[('Usage x your plan: '+', '.join(used_rate)) if used_rate else '',('MIS plan (not in your plan): '+', '.join(used_mis)) if used_mis else '']
        b['plan_basis']=' | '.join(p_ for p_ in parts if p_)
        out.append(b)
    return out

def bom_rate_checks(mis_path, bom, BW, DS_rm, TD=None):
    TD=TD or TARGET_DAYS
    """Usage per module: MIS standard BOM (Total Plan Req) vs your BOM sheet, for raw materials in both.
       Units are converted with the pack size written on the WSA sheet unit (e.g. 'DRUM(270KG/DRUM)')."""
    I=[]; wb=wb_ro(mis_path)
    mis_rate={}; tot_={}; unit_={}; seen_=set()
    for r in wb['Total Plan Req'].iter_rows(min_row=2,values_only=True):
        if tp_ok(r,DS_rm['plant']) and isinstance(r[12],(int,float)):
            k_=(str(r[10]).strip(),str(r[4]).strip())
            if isinstance(r[3],(int,float)) and r[3] and isinstance(r[14],(int,float)):
                tot_.setdefault(k_,[0.0,0.0]); tot_[k_][0]+=r[14]
                pl_=(k_,r[0],r[1],r[2],r[17] if len(r)>17 else None)     # modules of one plan line counted once per material
                if pl_ not in seen_: seen_.add(pl_); tot_[k_][1]+=r[3]
            unit_.setdefault(k_,str(r[13] or '').strip())
    for k_,(rq,md) in tot_.items(): mis_rate[k_]=(rq/md,unit_.get(k_,''))   # requirement / modules (each plan line counts its modules once per material line)
    pack={}
    for lab,sh in BW.get('sheets',{}).items():
        m=re.search(r'\((\d+(?:\.\d+)?)\s*KG\s*/\s*(\w+)\)',str(sh['uom'] or '').upper())
        if m:
            for nm in sh['members']: pack[nm]=float(m.group(1))
    blocks={b['mat']:b for b in DS_rm['blocks']}
    grp={}; alt_={}
    for d in bom:
        c=str(d['Component'] or '').strip(); bn=str(d['MIS BOM Name'] or '').strip(); f_=1.0; is_alt=False
        if c in (DS_rm.get('alias') or {}): f_=(DS_rm.get('alias_factor') or {}).get(c,1.0); c=DS_rm['alias'][c]; is_alt=True
        if isinstance(d['Quantity'],(int,float)):
            k=(c,bn,str(d['Module Model'] or '').strip(),str(d['UoM'] or '').strip().upper(),str(d.get('Customer') or '').strip())   # per customer: customers sharing one MIS BOM name are NOT added up
            if is_alt: alt_[k]=max(alt_.get(k,0.0),float(d['Quantity'])*f_)   # alternative code (PPC BOM '_Alt' line / other roll length): used INSTEAD of the main code, never on top
            else: grp[k]=grp.get(k,0.0)+float(d['Quantity'])*f_
    for k,q in alt_.items(): grp.setdefault(k,q)            # only when the main code is not listed for that customer / model
    # MIS rate per module for the same material: sum of its standard BOM lines = requirement / modules
    for (c,bn,mdl,u,cu_),q in grp.items():
        if c not in blocks or (c,bn) not in mis_rate: continue
        b=blocks[c]; mr,mu=mis_rate[(c,bn)]
        if u==mu.upper(): conv=q
        elif u=='KG' and b['common'] in pack: conv=q/pack[b['common']]
        else: continue
        if mr and abs(conv/mr-1)>0.05:
            ratio=conv/mr; P=[x*ratio for x in b['plan']]; T=[a+x for a,x in zip(b['transit'],b['adhoc'])]
            cv0=cover_days(own_stock(b),b['plan'],T,b.get('openpo') or [0.0]*len(P))[1]; cv1=cover_days(own_stock(b),P,T,b.get('openpo') or [0.0]*len(P))[1]
            I.append(dict(area='Plan',check='Usage per module: MIS BOM vs your BOM',sev='High' if (cv1<TD<=cv0 or (cv1<TD and cv1<cv0)) else 'Medium',
                          material=b['common'],code=c,a_lab=f'MIS per module ({mu})',a=round(mr,6),b_lab=f'Your BOM per module ({mu})',b=round(conv,6),diff=round(conv-mr,6),
                          detail=(f"{bn}: your BOM {q:.4g} {u}"+(f" ({cu_})" if cu_ else '')+(f" = {conv:.5f} {mu} at {pack.get(b['common']):.0f} KG/{mu}" if u!=mu.upper() else '')+
                                  f" vs MIS {mr:.5g} {mu} ({ratio-1:+.0%}). The MRP uses YOUR BOM for this customer (MIS only when your BOM does not list it)."),
                          action='Information: check which rate is right - if your BOM is right, ask MIS to correct its Standard BOM; if MIS is right, correct your BOM'))
    seen=set(); out=[]
    for x in I:
        k=(x['code'],x['detail'].split(':')[0],x['b'])     # one row per MIS BOM name and rate (customers with the same rate shown once)
        if k not in seen: seen.add(k); out.append(x)
    return out

BOM_SRC_NAME=[None]

def _line_key(x):
    """Common production-line key for your Production Plan and MIS: 'USA Line 3' / 'L3' -> 'L3'; Imperial*, Runergy*, Arizona* -> name."""
    t=' '.join(str(x or '').split()).lower()
    m=re.search(r'(?:usa\s*line|^l)\s*(\d+)',t)
    if m: return f'L{m.group(1)}'
    for k in ('imperial','runergy','arizona','inox'):
        if k in t: return k.capitalize()
    return ' '.join(str(x or '').split()) or '(blank)'

def mis_modules_by(mis_path, plant, dates):
    """MIS modules per day by (customer, line key), one count per plan line (date + line + customer + BOM + OA)."""
    wb=wb_ro(mis_path); idx={pd.Timestamp(d).normalize():k for k,d in enumerate(dates)}
    seen=set(); out={}
    for r in wb['Total Plan Req'].iter_rows(min_row=2,values_only=True):
        if not tp_ok(r,plant) or not isinstance(r[0],(dt.datetime,dt.date)) or not isinstance(r[3],(int,float)): continue
        k=idx.get(pd.Timestamp(r[0]).normalize())
        if k is None: continue
        key=(r[0],r[1],r[2],r[4],r[17] if len(r)>17 else None)
        if key in seen: continue
        seen.add(key); cu=' '.join(str(r[2] or '').split())
        out.setdefault((cu,_line_key(r[1])),[0.0]*len(dates))[k]+=r[3]
    return out

def make_plan_compare_xlsx(pp_rows, mis_by, dates, path):
    """Modules: your Production Plan vs MIS, by week - by customer and by production line."""

    AR='Calibri'; NAVY='1F3864'; hf=Font(name=AR,size=9,bold=True,color='FFFFFF'); hfill=PatternFill('solid',fgColor=NAVY)
    bf=Font(name=AR,size=9); bold=Font(name=AR,size=9,bold=True); amber=PatternFill('solid',fgColor='FFEB9C'); green=PatternFill('solid',fgColor='C6EFCE')
    nd=len(dates); per=[]
    for k,d in enumerate(dates):
        mon=pd.Timestamp(d)-pd.Timedelta(days=pd.Timestamp(d).weekday())
        if per and per[-1][0]==mon: per[-1][2]=k
        else: per.append([mon,k,k])
    wb=Workbook(); ws=wb.active; ws.title='Plan Compare'
    ws['A1']='Module plan compare - your Production Plan vs MIS Total Plan Req (modules per week)'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=(f'From {pd.Timestamp(dates[0]):%d-%b-%Y}, {nd} days. MIS modules counted once per plan line (date + line + customer + BOM + OA). Your customers are matched to MIS customers '
              '(see Module Plan); lines matched as USA Line n = Ln, Imperial*, Runergy*, Arizona*. Difference = yours - MIS; amber = different. The MRP uses YOUR plan.')
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:Z2'); ws['A2'].alignment=Alignment(wrap_text=True); ws.row_dimensions[2].height=28
    mine_c={}; mine_l={}
    for x in pp_rows or []:
        cu=x.get('mis_cust') or x.get('cust'); lk=_line_key(x.get('line'))
        for dct,key in ((mine_c,cu),(mine_l,lk)):
            a_=dct.setdefault(key,[0.0]*nd)
            for k,v in enumerate(x['v'][:nd]): a_[k]+=v
    mis_c={}; mis_l={}
    for (cu,lk),v in mis_by.items():
        for dct,key in ((mis_c,cu),(mis_l,lk)):
            a_=dct.setdefault(key,[0.0]*nd)
            for k in range(nd): a_[k]+=v[k]
    def table(top,title,keyname,mine,mis):
        ws.cell(top,1,title).font=Font(name=AR,size=12,bold=True,color=NAVY); h=top+1
        for j,t in enumerate([keyname,'Source','Match?','Total']): c=ws.cell(h,1+j,t); c.font=hf; c.fill=hfill; c.alignment=Alignment(horizontal='center')
        for w,(mon,a,z) in enumerate(per): c=ws.cell(h,5+w,mon.to_pydatetime()); c.number_format='dd-mmm'; c.font=hf; c.fill=hfill; c.alignment=Alignment(horizontal='center')
        keys=sorted(set(mine)|set(mis),key=lambda k_:-(sum(mine.get(k_,[0]))+sum(mis.get(k_,[0]))))
        keys=[k_ for k_ in keys if sum(mine.get(k_,[0]))+sum(mis.get(k_,[0]))>0]
        r=h+1; tm=[0.0]*len(per); tx=[0.0]*len(per); nd_diff=0
        for k_ in keys+['TOTAL']:
            if k_=='TOTAL': wm,wx=tm,tx
            else:
                a_=mine.get(k_,[0.0]*nd); b_=mis.get(k_,[0.0]*nd)
                wm=[sum(a_[a:z+1]) for _,a,z in per]; wx=[sum(b_[a:z+1]) for _,a,z in per]
                tm=[p+q for p,q in zip(tm,wm)]; tx=[p+q for p,q in zip(tx,wx)]
            diff=[p-q for p,q in zip(wm,wx)]; same=all(abs(d_)<0.5 for d_ in diff)
            for i_,(src,vals) in enumerate((('Your Production Plan',wm),('MIS',wx),('Difference',diff))):
                rr=r+i_; fnt=bold if (k_=='TOTAL' or src=='Difference') else bf
                ws.cell(rr,1,k_).font=fnt; ws.cell(rr,2,src).font=fnt
                c=ws.cell(rr,4,sum(vals)); c.number_format='#,##0;-#,##0;"-"'; c.font=fnt
                for w,v in enumerate(vals):
                    c=ws.cell(rr,5+w,v); c.number_format='#,##0;-#,##0;"-"'; c.font=fnt
                    if src=='Difference' and abs(v)>=0.5: c.fill=amber
                if src=='Difference':
                    c=ws.cell(rr,3,'Match' if same else 'Different'); c.font=bold; c.fill=green if same else amber
                    if not same: ws.cell(rr,4).fill=amber
            if not same and k_!='TOTAL': nd_diff+=1
            r+=3
        return r+1,nd_diff,len(keys)
    nxt,dc,nc=table(4,'By customer','Customer',mine_c,mis_c)
    _,dl,nl=table(nxt+1,'By production line','Production line',mine_l,mis_l)
    ws.cell(3,1,f'Different: {dc} of {nc} customers, {dl} of {nl} production lines.').font=Font(name=AR,size=10,bold=True,color='9C0006')
    ws.column_dimensions['A'].width=28; ws.column_dimensions['B'].width=20; ws.column_dimensions['C'].width=10; ws.column_dimensions['D'].width=11
    for w in range(len(per)): ws.column_dimensions[L(5+w)].width=9
    ws.freeze_panes='E6'; ws.sheet_view.zoomScale=90; wb.save(path)
    return dc,dl

def make_no_demand_xlsx(DS, mis_path, bom, modplan, BW, path):
    """Materials with stock (or supply coming) but NO demand in the horizon:
       A) MRP materials whose plan is 0;  B) other RM codes with stock in MB52 (plant + warehouse) or JW that are neither in the MIS plan
       (Total Plan Req, horizon) nor in your BOM for a model that is planned, nor an alternate code of a planned material."""
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    dates=DS['dates']; d0=pd.Timestamp(dates[0]).normalize(); dN=pd.Timestamp(dates[-1]).normalize()
    mis=wb_ro(mis_path)
    plant=DS['plant'].upper(); stock={}; desc={}; uom={}; wh={}
    for r in mis['MB52'].iter_rows(min_row=2,values_only=True):
        pn=str(r[13] or '').strip().upper()
        if pn not in (plant,plant+' WAREHOUSE'): continue
        c=str(r[2] or '').strip(); q=r[14] if isinstance(r[14],(int,float)) else 0
        stock[c]=stock.get(c,0.0)+q; desc.setdefault(c,r[5]); uom.setdefault(c,r[7])
        if pn.endswith('WAREHOUSE'): wh[c]=wh.get(c,0.0)+q
    jw={}
    for r in mis['RM Status for OEM JW'].iter_rows(min_row=2,values_only=True):
        if r and r[1] and isinstance(r[11],(int,float)) and r[11]: jw[str(r[1]).strip()]=jw.get(str(r[1]).strip(),0.0)+r[11]
    intr={}
    for r in list(mis['In Transit'].iter_rows(values_only=True))[1:]:
        if len(r)>39 and isinstance(r[34],(int,float)) and r[34]>0 and it_ok(r): intr[str(r[39] or '').strip()]=intr.get(str(r[39] or '').strip(),0.0)+r[34]
    planned=set()
    for r in mis['Total Plan Req'].iter_rows(min_row=2,values_only=True):
        if tp_ok(r,DS['plant']) and isinstance(r[0],(dt.datetime,dt.date)) and d0<=pd.Timestamp(r[0]).normalize()<=dN and isinstance(r[14],(int,float)) and r[14]>0:
            planned.add(str(r[10] or '').strip())
    live_bn={bn for bn,cu in (modplan or {}).items() if any(any(v) for v in cu.values())}
    for d in bom or []:
        if str(d.get('MIS BOM Name') or '').strip() in live_bn and d.get('Component'): planned.add(str(d['Component']).strip())
    alias=DS.get('alias') or {}
    blocks={b['mat']:b for b in DS['blocks']}
    for c,dc in alias.items():
        if dc in blocks and sum(blocks[dc]['plan'])>0: planned.add(c)
    rows=[]
    for b in DS['blocks']:
        if sum(b['plan'])>0: continue
        sup=sum(b['transit'])+sum(b.get('openpo') or [0])
        if b['stock']>0 or sup>0:
            rows.append(('A  MRP material, plan = 0','PM / CS' if b.get('count_stock') else 'RM',b['common'],b['mat'],b.get('uom'),b['stock'],0.0,sum(b['transit']),sum(b.get('openpo') or [0]),
                         'In the MRP but no customer plan uses it in the horizon'+(' (check the BOM / MIS BOM Name)' if b.get('count_stock') else '')))
    for c in sorted(set(stock)|set(jw)):
        if not c.upper().startswith('RM') or c in blocks or c in planned: continue
        q=stock.get(c,0.0); j=jw.get(c,0.0)
        if q<=0 and j<=0: continue
        if c.upper().startswith('RMFCM'):     # cells: planned by bin / efficiency, not by one code - kept apart
            rows.append(('C  Cells (not planned by this MRP)','Cell',str(desc.get(c) or ''),c,uom.get(c),q,j,intr.get(c,0.0),0.0,
                         'Cells are not in the MIS plan; your BOM names one cell code per model, so other bins / efficiencies show here - check with production'))
            continue
        why='Alternate code of a material with no plan' if c in alias else 'Not in the MIS plan and not in your BOM for a planned model'
        rows.append(('B  Other RM with stock','RM',str(desc.get(c) or ''),c,uom.get(c),q,j,intr.get(c,0.0),0.0,why))
    wb=Workbook(); ws=wb.active; ws.title='No Demand'
    ws['A1']='Materials with no demand plan - stock or supply but nothing planned to use it'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=(f'Horizon {d0:%d-%b-%Y} to {dN:%d-%b-%Y}. Demand = MIS plan (Total Plan Req) or your BOM for a planned model (your Production Plan). '
              'Stock = MB52 (or the PPC Count when counted); JW = RM Status for OEM JW (shown only, in no calculation); CS / PM stock = PPC Count. Use it to stop new orders, re-use or return slow stock.')
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:J2'); ws['A2'].alignment=Alignment(wrap_text=True); ws.row_dimensions[2].height=28
    H=['Group','Type','Material','Code','UOM','Stock','JW at OEM','In transit (MIS)','Open PO (WSA)','Why no demand']
    for j,h in enumerate(H,1):
        c=ws.cell(4,j,h); c.font=Font(name=AR,size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=NAVY); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=bd
    rows.sort(key=lambda t:(t[0],-(t[5]+t[6]+t[7]+t[8])))
    for i,t in enumerate(rows):
        for j,v in enumerate(t,1):
            c=ws.cell(5+i,j,round(v,3) if isinstance(v,float) else v); c.font=Font(name=AR,size=9); c.border=bd; c.alignment=Alignment(vertical='top',wrap_text=j in (3,10))
            if j in (6,7,8,9): c.number_format='#,##0.##;-#,##0.##;"-"'
        if t[7] or t[8]:
            ws.cell(5+i,8).fill=PatternFill('solid',fgColor='FFEB9C')
            if t[8]: ws.cell(5+i,9).fill=PatternFill('solid',fgColor='FFEB9C')     # (re-assigning a cell's own fill crashes openpyxl)
    if rows: ws.add_table(Table(displayName='NoDemand',ref=f'A4:J{4+len(rows)}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=True)))
    else: ws.cell(5,1,'Every material with stock has a demand plan').font=Font(name=AR,size=10,italic=True)
    ws.cell(3,1,f"{sum(1 for t in rows if t[0].startswith('A'))} MRP materials with plan 0; {sum(1 for t in rows if t[0].startswith('B'))} other RM codes with stock and no demand; "
                 f"{sum(1 for t in rows if t[0].startswith('C'))} cell codes (not planned here). "
                 "Amber = more is still coming (in transit / Open PO).").font=Font(name=AR,size=10,bold=True,color='9C0006')
    for j,w in enumerate([22,9,42,17,7,12,11,12,12,52],1): ws.column_dimensions[L(j)].width=w
    ws.freeze_panes='D5'; set_print(wb.active,'No Demand'); wb.save(path); return len(rows)

MOVE_DAYS=14   # days of expected out / in kept per run in the Movement Log
def read_movement_log(prev_path):
    """Movement Log rows kept in MRP_Daily: one row per run date and material."""
    out=[]
    if not prev_path or not os.path.exists(prev_path): return out
    try: wb=wb_ro(prev_path)
    except Exception as _e_: _swallow(3705,_e_); return out
    if 'Movement Log' not in wb.sheetnames: return out
    rows=list(wb['Movement Log'].iter_rows(values_only=True))
    if not rows: return out
    h=list(rows[0])
    for r in rows[1:]:
        if not r or not r[0]: continue
        d=dict(zip(h,r))
        try:
            out.append(dict(date=pd.Timestamp(d['Date']).normalize(),code=str(d['Code']),material=d['Material'],uom=d['UOM'],stock=float(d['Stock'] or 0),
                            oem=(float(d['to OEM Stock']) if isinstance(d.get('to OEM Stock'),(int,float)) else None),
                            plan=[float(x) for x in str(d['Plan next days'] or '').split(';') if x!=''],
                            inb=[float(x) for x in str(d['In transit next days'] or '').split(';') if x!=''],
                            opo=[float(x) for x in str(d['Open PO next days'] or '').split(';') if x!=''],raw=d.get('Raw stock') or '',
                            sup=[(int(a),float(b),c) for a,b,c in (t.split('~',2) for t in str(d.get('In transit by supplier') or '').split('|') if t.count('~')==2)]))
        except Exception as _e_: _swallow(3720,_e_); continue
    return out

def movement_entries(DS):
    d0=pd.Timestamp(DS['dates'][0]).normalize(); out=[]
    for b in DS['blocks']:
        P=b['plan'][:MOVE_DAYS]; I=[a+c for a,c in zip(b['transit'],b['adhoc'])][:MOVE_DAYS]; O=(b.get('openpo') or [0.0]*MOVE_DAYS)[:MOVE_DAYS]
        ts=(DS.get('trans_sup') or {}).get(b['mat'],{})
        sup=[(k,q,sn) for sn,arr in ts.items() for k,q in enumerate(arr[:MOVE_DAYS]) if q>0]
        out.append(dict(date=d0,code=b['mat'],material=b['common'] if not b.get('count_stock') else b['desc'],uom=b.get('uom') or '',stock=float(b['stock']),
                        oem=0.0 if b.get('count_stock') else float(b.get('st_jw') or 0.0),plan=P,inb=I,opo=O,
                        cs=bool(b.get('count_stock')),raw=b.get('raw') or '',sup=sup))
    return out




NO_OEM_KEYS={'L1','L2','L3','L4'}
ATA_SHOW=4          # Arrival at plant: show the next 4 arrival dates, the rest as one total
PLAN_AVG_DAYS=globals().get('PLAN_AVG_DAYS',49)     # Action List daily plan = average of the next N days that have production (> 0) - Settings
def calc_layout(A, RM_):
    """Row positions on the two calculation sheets, decided before the Action List is written (its formulas point at them)."""
    tg=[]; r=2
    for x in A:
        for lk,u in sorted((x.get('tgt_use') or {}).items()):
            tg.append(dict(code=x['b']['mat'],lk=lk,use=u,row=r)); r+=1
    pl={}; r=3
    for x in A:
        if x['b']['mat'] in RM_: pl[x['b']['mat']]=r; r+=2
    r+=1; pll={}                       # per material and production line (MRP by Line)
    for x in A:
        m=x['b']['mat']
        if m not in RM_: continue
        for lk in dict.fromkeys(_line_key(ln) for (cu,ln) in RM_[m]['custs'] if ln):
            pll[(m,lk)]=r; r+=1
    PLAN_LINE_ROWS.clear(); PLAN_LINE_ROWS.update(pll)
    return tg,pl

def make_calc_target_xlsx(tg, path):
    """Calc Target: per material and line that runs a product using it today: usage per module x the line's daily target."""
    wb=Workbook(); ws=wb.active; ws.title='Calc Target'
    for j,h in enumerate(['Component','Line','Group',f'Usage per module (product run within {int(LOOKAHEAD_DAYS)} days)','Line daily target','Need per day (= D x E)'],1): ws.cell(1,j,h).font=Font(bold=True)
    for t in tg:
        r=t['row']; lab,v=EFF_TARGET.get(t['lk'],(t['lk'],0))
        ws.cell(r,1,t['code']); ws.cell(r,2,lab.split(' Production')[0]); ws.cell(r,3,'Lines 1-4' if t['lk'] in NO_OEM_KEYS else 'OEM line')
        ws.cell(r,4,round(t['use'],8)); ws.cell(r,5,f"={TGT_CELL[t['lk']]}" if t['lk'] in TGT_CELL else v); ws.cell(r,6,f'=D{r}*E{r}')
        if t['lk'] not in TGT_CELL: ws.cell(r,7,'no Settings target: highest planned day')
        ws.cell(r,8,t['lk'])
    for j,w in enumerate([17,30,11,14,12,14,34],1): ws.column_dimensions[L(j)].width=w
    wb.save(path)

def make_calc_plan_xlsx(A, RM_, pl, d0, path):
    """Calc Plan: per material 4 rows - plan with OEM (= MRP Daily Gross requirement), its running total, plan without OEM
       (= MRP Daily line rows of USA Lines 1-4), its running total. Inv. Days at daily plan = days until the running total passes the stock."""
    wb=Workbook(); ws=wb.active; ws.title='Calc Plan'
    ws.cell(1,1,f'Plan per day from MRP Daily (formulas) - Action List daily plan = average of the first {PLAN_AVG_DAYS} days (from today) with production > 0').font=Font(bold=True)
    any_=next(iter(RM_.values()),None)
    if not any_: wb.save(path); return
    C0,ND=any_['C0'],any_['ND']
    ws.cell(2,1,'Component'); ws.cell(2,2,'Row')
    for k in range(ND): ws.cell(2,3+k,(d0+pd.Timedelta(days=k)).to_pydatetime()).number_format='dd-mmm'
    for x in A:
        m=x['b']['mat']
        if m not in pl: continue
        rm=RM_[m]; r=pl[m]
        no_oem=[rr for rr,(cu,ln) in zip(rm['rc'],rm['custs']) if ln and _line_key(ln) in NO_OEM_KEYS]
        for i,lab in enumerate(['Plan with OEM','Plan without OEM (Lines 1-4)']):
            ws.cell(r+i,1,m); ws.cell(r+i,2,lab)
        for k in range(ND):
            src=L(C0+k)
            ws.cell(r,3+k,f"='MRP Daily'!{src}{rm['rp']}")
            ws.cell(r+1,3+k,('='+'+'.join(f"'MRP Daily'!{src}{rr}" for rr in no_oem)) if no_oem else 0)
    for (m,lk),r in PLAN_LINE_ROWS.items():          # one row per material and production line
        rm=RM_[m]; rows_l=[rr for rr,(cu,ln) in zip(rm['rc'],rm['custs']) if ln and _line_key(ln)==lk]
        ws.cell(r,1,m); ws.cell(r,2,f'Plan line {lk}')
        for k in range(ND): ws.cell(r,3+k,('='+'+'.join(f"'MRP Daily'!{L(C0+k)}{rr}" for rr in rows_l)) if rows_l else 0)
    ws.column_dimensions['A'].width=17; ws.column_dimensions['B'].width=28
    wb.save(path)



def al_group(h):
    if 'module plan' in h: return 'mp'
    if 'OEM' in h and ('arget' in h or 'lan' in h):
        return ('tg_' if 'arget' in h else 'pl_')+('wo' if 'without OEM' in h else 'w')
    if 'WSA RM Planning' in h: return 'wsa'
    if h in ('Status','Exception message','Action','Owner','Due','Status update','Comment'): return 'work'
    return 'base'
def al_hdr_light(h): return ('day by day' in h) or ('stock only' in h)
def al_hdr_color(h):
    """Action List header colours: PURPLE = daily target, GOLD = daily plan (light = stock only / day by day);
       the scenario (without / with OEM) is the coloured banner in row 3 and the end of the header text."""
    g=al_group(h)
    if g in ('tg_w','tg_wo','pl_w','pl_wo'):
        tgt=g.startswith('tg_')
        if al_hdr_light(h): return 'D9C3E9' if tgt else 'FFE699'
        return '7030A0' if tgt else 'BF8F00'
    return {'wsa':'833C0B','work':'595959','mp':'BF8F00'}.get(g,'1F3864')



def sup_short(n):
    """Short supplier name: no country in brackets, no Co./Ltd./Glass/Technology..., first 2 words (FLAT (VIETNAM) CO., LTD -> Flat)."""
    t=re.sub(r'\([^)]*\)',' ',str(n or '')).replace(',',' ').replace('.',' ')
    w=[x for x in t.split() if x.upper() not in ('CO','LTD','LIMITED','COMPANY','CORP','CORPORATION','INC','GLASS','TECHNOLOGY','TECH','INDUSTRIAL','INDUSTRY','NEW','MATERIAL','MATERIALS','SOLAR','ENERGY','JSC','PTE','LLC','GROUP')]
    return ' '.join(w[:2]).title() or str(n or '')




SUMMARY_LINES=[]; DEF_SHEETS_SEEN=[set()]
def audit_workbook(path):
    DEF_SHEETS_SEEN[0]={str(r[0]).strip().lower() for r in DEFINITIONS}|{p_.strip().lower() for r in DEFINITIONS for p_ in re.split(r' / |, | \(',str(r[0]))}
    """Self-check after every run (your rules): (1) a material short in the plan is never labelled safe on a visible sheet;
       (2) every visible sheet lists materials in the same shortage-risk order as the ALERT sheet (Action List is checked like the others).
       Prints what breaks the rules AND which visible sheets were NOT checked and why."""
    try:
        wb=load_workbook(path,read_only=True)
        if 'ALERT' not in wb.sheetnames: print('NOTE: self-check skipped (no ALERT sheet)'); return
        order=[]; short=set(); STOCKOUT={ALERT_CLS[0][0],ALERT_CLS[1][0]}        # reference list = the ALERT rows (class, then first stock-zero day)
        for r in wb['ALERT'].iter_rows(min_row=ALERT_HR+1,values_only=True):
            cls_,code_=r[1],r[2]
            if not code_ or not isinstance(code_,str) or not re.match(r'^(JW)?(RM|PM)[A-Z]*\d+$',code_.strip()): continue
            if code_ not in order: order.append(code_)
            if cls_ in STOCKOUT: short.add(code_)                                   # short in the plan (inside or after the transit lead time)
        if not order: print('NOTE: self-check skipped (ALERT sheet has no rows)'); return
        dup_={c for c in order if sum(1 for r in wb['ALERT'].iter_rows(min_row=ALERT_HR+1,min_col=3,max_col=3,values_only=True) if r[0]==c)>1}
        dup_|={c for c in order if POOLG.get(c)}                    # substitute members: own rows on the ALERT, pooled on the Action List / MRP Daily (user, Nov 2026) - no single place   # split in 2 ALERT rows (back glass 333 per WSA sheet): no single place in the order
        rank={c:i for i,c in enumerate([c for c in order if c not in dup_])}; safe={'In stock','In Stock','Excess','Above reorder','OK','no shortage','No action','covered'}
        issues=[]; NOT_CHECKED=[]; CHECKED=[]
        SKIP={'Settings','Stock Count','BOM','Module Plan','MIS Layout','How it works','Definitions','Movement Log','In Transit Status','RM Issues',
              'Check','Stock Data','No Demand','Substitute Advice','ALERT (formulas)','ALERT Calc','ALERT Plan lines','ALERT Arrival lines','ETD alert'}   # these were hidden (never checked) before all sheets became visible; not in shortage order
        FREE={'ALERT',HEAT_NAME,DAILY_NAME}     # Cover Heatmap: grouped RM / PM / CS (user, Nov 2026), risk order inside each group; ALERT: own order (alert class, first stock-zero day; back glass 333 per customer project): only the 'labelled safe' rule is checked
        for ws in wb.worksheets:
            if ws.sheet_state!='visible': continue
            seq=[]; bad=0; SEVL=[None]
            if ws.title in SKIP:                                         # your input copies / logs kept in source or date order
                NOT_CHECKED.append(f'{ws.title} (source / date order or log)'); continue
            for i,r in enumerate(ws.iter_rows(values_only=True),1):
                if not any(v not in (None,'') for v in r):              # blank row: next list starts
                    bad+=sum(1 for a_,b_ in zip(seq,seq[1:]) if rank[b_]<rank[a_]); seq=[]; continue
                sev_=r[0] if r and r[0] in ('High','Medium','Low') else None
                if sev_ and sev_!=SEVL[0]:                                 # Check sheet: each severity group is its own list
                    bad+=sum(1 for a_,b_ in zip(seq,seq[1:]) if rank[b_]<rank[a_]); seq=[]; SEVL[0]=sev_
                codes=[v for v in r[:5] if isinstance(v,str) and v in rank]
                if not codes and any(isinstance(v,str) and v.strip() and not CODE.match(v.strip()) for v in r[:2]) and not any(isinstance(v,(int,float)) for v in r[:6]):
                    bad+=sum(1 for a_,b_ in zip(seq,seq[1:]) if rank[b_]<rank[a_]); seq=[]; continue      # a title / header row: next table starts
                if not codes: continue
                if codes[0] in short and any(isinstance(v,str) and v.strip() in safe for v in r): issues.append(f'{ws.title} row {i}: {codes[0]} short but labelled safe')
                if codes[0] not in seq: seq.append(codes[0])                 # a material split in 2 rows (back glass 333 per WSA sheet) counts once
            bad+=sum(1 for a,b in zip(seq,seq[1:]) if rank[b]<rank[a])
            (CHECKED if ws.title not in FREE else NOT_CHECKED).append(ws.title if ws.title not in FREE else f'{ws.title} (order not checked - only the labelled-safe rule)')
            if bad and ws.title not in FREE: issues.append(f'{ws.title}: {bad} places out of shortage-risk order')
        print(('Self-check (shortage-risk rules): no problems found' if not issues else 'Self-check WARNING:\n  '+'\n  '.join(issues[:10]))
              +f"\n  Order checked on {len(CHECKED)} sheet(s): {', '.join(CHECKED)}\n  NOT checked ({len(NOT_CHECKED)}): {'; '.join(NOT_CHECKED)}")
        SUMMARY_LINES.append(f'Self-check order test NOT applied to {len(NOT_CHECKED)} sheet(s) (listed above in the Self-check lines)')
        miss_=[w_.title for w_ in wb.worksheets if w_.title!='Definitions' and w_.title.lower() not in DEF_SHEETS_SEEN[0]]
        if miss_: print('WARNING: sheets with NO row on the Definitions sheet: '+', '.join(miss_)); SUMMARY_LINES.append('Sheets without a Definitions row: '+', '.join(miss_))
        else: print('Definitions check: every sheet of this workbook has a row on the Definitions sheet')
    except Exception as ex_: print(f'NOTE: self-check skipped ({type(ex_).__name__})')


def run_summary(title, extra=()):
    """What this run did NOT count / skipped / swallowed - printed and appended to Run_Summary_<date>.txt next to the outputs."""
    L_=[f'--- RUN SUMMARY: {title} ---']
    L_+=list(extra)
    L_+=[m_ for m_ in RUN_NOTES if m_.startswith(('WARNING','ERROR'))]
    L_+=[m_ for m_ in RUN_NOTES if m_.startswith('NOTE')]
    L_+=list(SUMMARY_LINES)
    if SWALLOWED:
        L_.append(f'{len(SWALLOWED)} silent error handler(s) caught an error this run (script line: count, error) - mostly file probing, but check any that look like data:')
        L_+=[f'  line {k}: x{v[0]} {v[1]}: {v[2]}' for k,v in sorted(SWALLOWED.items())]
    else: L_.append('No silent error handler caught an error this run.')
    for l_ in L_: _bi.print(l_)
    try:
        fp_=os.path.join(OUTDIR,f'Run_Summary_{dt.date.today():%Y-%m-%d}.txt')
        with open(fp_,'a',encoding='utf-8') as h_: h_.write('\n'.join(L_)+'\n\n')
    except Exception as e_: _bi.print(f'NOTE: Run_Summary file not written ({e_})')

POST={'path':None,'wb':None}
def post_wb(path):
    """The finished workbook is opened ONCE for all finishing steps (dashboards, sheets, style) and saved once (post_flush)."""
    if POST['path']!=path or POST['wb'] is None: POST['wb']=load_workbook(path); POST['path']=path
    return POST['wb']
def post_flush(path):
    if POST['wb'] is not None and POST['path']==path: _safe_save(POST['wb'],path)
    POST['wb']=None; POST['path']=None

def _fix_views(wb):
    """Excel 'Repaired Records: View from /xl/worksheets/sheetN.xml' (user, Nov 2026): a sheet whose frozen pane was removed
       (_tidy_dashboard, no freeze) kept its <selection pane="..."> entries - pointing at panes that no longer exist, and doubled.
       Every sheet: no pane -> one plain selection; a pane -> one selection per pane that exists for that split (last one kept)."""
    from openpyxl.worksheet.views import Selection
    for ws_ in wb.worksheets:
        try:
            sv_=ws_.sheet_view; sel_=list(sv_.selection or [])
            act_=next((x.activeCell for x in reversed(sel_) if x.activeCell),None) or 'A1'
            pn_=sv_.pane
            if pn_ is None or not (pn_.xSplit or pn_.ySplit):
                sv_.pane=None; sv_.selection=[Selection(activeCell=act_,sqref=act_)]; continue
            ok_={'bottomRight','topRight','bottomLeft'} if (pn_.xSplit and pn_.ySplit) else ({'topRight'} if pn_.xSplit else {'bottomLeft'})
            keep_={}
            for x in sel_:
                if x.pane in ok_: keep_[x.pane]=x
            if not keep_: keep_={pn_.activePane or sorted(ok_)[0]:Selection(pane=pn_.activePane or sorted(ok_)[0],activeCell=act_,sqref=act_)}
            sv_.selection=[keep_[k] for k in ('topRight','bottomLeft','bottomRight') if k in keep_]
        except Exception as _e_: _swallow(0,_e_)
def _safe_save(wb, path):
    """Save to a temporary file first and replace the real file only when the save worked - a failing step never damages MRP_Daily."""
    _fix_views(wb)
    tmp_=path+'.saving.xlsx'
    try:
        wb.save(tmp_); os.replace(tmp_,path)
    finally:
        if os.path.exists(tmp_):
            try: os.remove(tmp_)
            except OSError: pass



def strip_code(t):
    """Alternative text without material codes: 'RMJUB00000072 JB 500/30A: 115,961 (Pepper)' -> 'JB 500/30A: 115,961 (Pepper)'."""
    return re.sub(r'\b(RM|PM|CS)[A-Z]*\d{5,}\s*','',str(t or '')).strip() or None
def is_rm(code): return str(code or '').upper().startswith('RM')


def risk_info(x, d0):
    """Shortage risk of one material: first short day (projected stock incl. in transit < 0), lead time, class, short date."""
    w_=x.get('wins') or []; k=w_[0][0] if w_ else None; lt=int(x['b'].get('lt_used') or 0); td=x.get('pl_days')     # days at the module plan
    if x.get('no_demand'): return dict(k=None,lt=lt,cls='No demand',date=None)      # no demand = never Critical / At risk / Watch / Short
    cls=('ALERT' if (k is not None and k<lt) else ('Short later' if k is not None else ('Under 21 d' if (td is not None and td<REPORT_DAYS) else ('OK' if x.get('pl_need') else 'No need'))))
    return dict(k=k,lt=lt,cls=cls,date=(d0+pd.Timedelta(days=k)) if k is not None else None)
POOLG={}   # code -> pooled projection inputs of its substitute group (shortage)
def grp_main(code):
    """Main material of the substitute group of this code (the code itself when it is not in a group)."""
    g=POOLG.get(code); return g['members'][0] if g else code
def grp_follower(code):
    """True for a group member that is NOT the main material: the group is listed / ordered once, on its main material."""
    g=POOLG.get(code); return bool(g) and g['members'][0]!=code
def grp_label(code, name):
    g=POOLG.get(code)
    return (' + '.join(g.get('names') or g['members'])+' (group)') if g else name
def grp_ships(code, ships):
    """Shipments of the material - for a substitute group: all members' shipments, quantities converted to THIS code's unit."""
    g=POOLG.get(code)
    if not g: return list((ships or {}).get(code,[]))
    fx=dict(g.get('fx') or {}); f0=fx.get(code,1.0); out=[]
    for c in g['members']:
        k=fx.get(c,1.0)/f0
        for s_ in (ships or {}).get(c,[]): out.append(dict(s_,qty=s_['qty']*k,shq=(s_.get('shq') or 0)*k,code=c))
    return out
CUST_FIX_=[]; POOL_=[]     # customer spellings matched to the BOM; substitute groups whose demand is pooled
NODEM_=[]            # materials with no demand in this run (mat, name, why)
RISK_RANK={}       # material -> rank (0 = highest shortage risk), same order as the Action List, used by every sheet
def set_risk_rank(A, d0):
    """Rank all materials by shortage risk: ALERT (earliest first) > short later (earliest first) > under 21 days (lowest days) > the rest."""
    G={'ALERT':0,'Short later':1,'Under 21 d':2,'Under 21 d later':2.5,'OK':3,'No need':4,'No demand':5}   # same six groups as the ALERT sheet
    def key(x):
        ri=risk_info(x,d0); td=x.get('pl_days'); cl=ri['cls']; k_=ri['k']
        if cl=='OK':                                                  # ALERT class 'Under 21 d': STORE stock alone covers < REPORT_DAYS (in transit may cover it)
            g_=POOLG.get(x['b']['mat']); st_=(g_['stock'] if g_ else float(x['b'].get('stock') or 0.0)); nd_=x.get('pl_need')
            if nd_ and max(0.0,st_)/nd_<REPORT_DAYS: cl='Under 21 d'
        if cl=='OK':                                                  # ALERT class 'Under 21 d later': cover falls below REPORT_DAYS within 16 weeks
            S_=x.get('S') or []; nd_=x.get('pl_need')
            t_=next((i for i in range(min(len(S_),112)) if nd_ and S_[i]/nd_<REPORT_DAYS),None)
            if t_ is not None: cl='Under 21 d later'; k_=t_
        return (G[cl],k_ if k_ is not None else 0,td if td is not None else 1e9,str(x['b'].get('common') or ''))
    RISK_RANK.clear(); RISK_RANK.update({x['b']['mat']:i for i,x in enumerate(sorted(A,key=key))})
def rrank(code): return RISK_RANK.get(code,10**6)
def avg_prod_days(): return max(1,int(AVG_PROD_WEEKS))*7        # averaging window of the need per day (days)
def short_lt_text(x, d0):
    """All periods with stock below zero that START inside the material transit lead time ('05-Nov → 07-Nov (3 d); 20-Nov → beyond 21-Mar');
       else the first stock-zero day after the transit lead time; else 'not short'. Returns (text, number of periods inside the transit lead time)."""
    if x.get('no_demand'): return 'No demand',0
    w_=x.get('wins') or []; lt_=int(x['b'].get('lt_used') or 0)
    if not w_: return 'not short',0
    D=lambda k: d0+pd.Timedelta(days=k)
    ins=[w for w in w_ if w[0]<lt_]
    if ins:
        return ';\n'.join(f"{D(a):%d-%b} → "+(f"beyond {D(z):%d-%b}" if op else f"{D(z):%d-%b} ({z-a+1} d)") for a,z,op in ins),len(ins)
    return f"{D(w_[0][0]):%d-%b} (after transit lead time {lt_} d)",0

# ===================== ALERT sheet: ONE simple list for both daily outputs =====================
ALERT_ROWS=[]      # built once in build_new_mrp: the first sheet of MRP_Daily AND of the Inbound file (WSA daily) - the same rows
ALERT_D0=[None]    # the date of those rows (MRP day 1 = today)
ALERT_CLS=[('STOCK-OUT IN TRANSIT LEAD TIME','C00000','FFFFFF','EXPEDITE / use an alternative'),
           ('STOCK-OUT LATER','E26B0A','FFFFFF','order by the date'),
           ('UNDER 21 DAYS','FFC000','000000','in transit covers - watch the arrivals'),
           (f'UNDER 21 DAYS LATER','2E75B6','FFFFFF','falls below 21 days within 16 weeks - expedite / order before the date')]
ALERT_CLS+=[('OK','1E8449','FFFFFF','cover above 21 days - no action'),('NO DEMAND','8A949E','FFFFFF','no demand in the plan')]   # every material is listed (user, Nov 2026)
ALERT_SHORT=['Stock-out in LT','Stock-out later','Under 21 d','Under 21 d later','OK','No demand']
def alert_why(S, need, d0, lt, ships_=None, T=None, sup=None):
    """The 'why' columns of the ALERT sheet (were the Risk Charts sheet, merged into ALERT Nov 2026, user request): Below 21 days
       from = first day projected stock covers < REPORT_DAYS days of need (next 16 weeks); Cause = LATE ARRIVAL (what is on the way
       would keep 21+ days, but arrives after that day) / LATE + NOT ENOUGH / NO SUPPLY; Days late = first arrival after that day minus
       that day; Action detail = the Risk Charts action. ships_ = MIS shipments (cover_watch result 'cw' style), else arrivals T per day."""
    TGT=REPORT_DAYS; HZ=min(len(S),112)
    if not need or not S: return dict(t=None,below=None,cause=None,late=None,detail='')
    t=next((k for k in range(HZ) if S[k]/need<TGT),None)
    if t is None: return dict(t=None,below=None,cause=None,late=None,detail='')
    dt_=d0+pd.Timedelta(days=t)
    if ships_ is not None: later=[(s_['raw'],s_['qty'],s_.get('sup')) for s_ in ships_]
    else: later=[(d0+pd.Timedelta(days=k),T[k],sup) for k in range(t+1,min(len(T),len(S))) if T[k]>0]
    later=sorted([l_ for l_ in later if l_[0]>dt_],key=lambda l_:l_[0])
    gap=TGT*need-S[t]; on_way=sum(q for _,q,_ in later); gap_all=max(0.0,TGT*need-min(S[t:HZ]))
    cause='NO SUPPLY' if not later else ('LATE ARRIVAL' if on_way>=gap else 'LATE + NOT ENOUGH')
    nx=later[0] if later else None
    if cause=='LATE ARRIVAL': det=f"Expedite {nx[2] or 'shipment'} to {dt_:%d-%b} (due {nx[0]:%d-%b})"+(f" + order {gap_all:,.0f} for later weeks" if gap_all>0 else '')
    elif cause=='LATE + NOT ENOUGH': det=f"Expedite {nx[2] or ''} + order {gap_all:,.0f} more".replace('  ',' ')
    else: det=f"Order {gap_all:,.0f} (nothing on the way after {dt_:%d-%b})"
    return dict(t=t,below=dt_,cause=cause,late=(nx[0]-dt_).days if nx else None,detail=det)
def alert_eval(name, code, uom, stock, S, T, need, lt, tr_tot, d0, where=None, sup=None, why=None, ata=None):
    """ONE rule for the ALERT sheet of both daily outputs (user request Nov 2026: 'list of less than 21 days inventory, how many days
       stock zero within transit lead time and zero from - to; the other sheets are too complicated - one simple sheet, strong alert'):
       days of stock = stock / need per day (module plan, average per production day, next 9 weeks);
       stock-zero day = that day's need is not covered: stock + in transit arrived up to that day - plan up to that day < 0
       (= the Action List 'Short from'; on a WSA sheet: Plant + Port Days (PPC In Transit) below 1 - the red cell comes a day later);
       stock-zero days inside the transit lead time = such days from today until today + transit lead time - 1.
       Listed when days of stock < REPORT_DAYS or a stock-zero day falls inside the transit lead time; None = not listed."""
    if not need:                                                  # no demand: still listed (class 5), no days of stock
        return dict(cls=5,code=code,name=name,where=where,uom=str(uom or ''),stock=stock,need=0.0,days=None,days_it=None,lt=int(lt or 0),zin=0,
                    first=None,first_in=False,per='none in the plan',nx=None,sup=sup,act='No demand in the plan',k0=None,
                    why=dict(t=None,below=None,cause=None,late=None,detail=''),ata=ata,S=list(S),tr_tot=tr_tot)
    nd=len(S); lt=int(lt or 0); wins=[]; k=0
    while k<nd:
        if S[k]<0:
            k2=k
            while k2+1<nd and S[k2+1]<0: k2+=1
            wins.append((k,k2,k2+1>=nd)); k=k2+1
        else: k+=1
    zin=sum(min(z,lt-1)-a+1 for a,z,_ in wins if a<lt)
    days=max(0.0,stock)/need; days_it=max(0.0,stock+(tr_tot or 0.0))/need
    why=why or dict(t=None,below=None,cause=None,late=None,detail='')
    D=lambda k_: d0+pd.Timedelta(days=k_)
    def ptxt(a,z,op):
        t_=f"{D(a):%d-%b} → beyond {D(z):%d-%b}" if op else (f"{D(a):%d-%b} (1 d)" if a==z else f"{D(a):%d-%b} → {D(z):%d-%b} ({z-a+1} d)")
        return ('⚠ ' if a<lt else '')+t_
    per='\n'.join(ptxt(*w) for w in wins[:6])+(f"\n... +{len(wins)-6} more" if len(wins)>6 else '')
    k0=wins[0][0] if wins else None; cls=0 if zin>0 else (1 if wins else (2 if days<REPORT_DAYS else (3 if why['t'] is not None else 4)))
    if cls==0: act='EXPEDITE / use an alternative'
    elif cls==1: act='ORDER NOW' if k0-lt<=RELEASE_WINDOW else f"ORDER BY {D(k0-lt):%d-%b}"
    elif cls==2: act='Watch the arrivals (in transit covers the plan)'
    elif cls==3: act=f"Keep {REPORT_DAYS} days: act before {why['below']:%d-%b}"
    else: act='OK - no action'
    if why.get('detail'): act+='\n'+why['detail']
    nx=next(((k_,T[k_]) for k_ in range(min(nd,len(T))) if T[k_]>0),None)
    return dict(cls=cls,code=code,name=name,where=where,uom=str(uom or ''),stock=stock,need=need,days=days,days_it=days_it,lt=lt,zin=zin,
                first=D(k0) if k0 is not None else None,first_in=(k0 is not None and k0<lt),per=per or 'none in the plan',
                nx=(D(nx[0]),nx[1]) if nx else None,sup=sup,act=act,k0=k0,why=why,ata=ata,S=list(S),tr_tot=tr_tot)
def build_alert_rows(A, d0, nd):
    """The ALERT list, built ONCE (MRP_Daily first tab = Inbound first tab). One row per material with demand from the Action List
       projection (stock + MIS in transit by arrival date - module plan; a substitute group = one row, pooled). A material that the
       WSA daily splits by customer project and supplier (on 2+ WSA sheets: back glass 333 = SB Back + Subanci Back) gets one row
       per sheet from that sheet's stock, demand and PPC arrivals - the same rule on the numbers of the sheet."""
    on={}
    for g in WSA_REFS[0] or []:
        for c in g.get('codes') or []: on.setdefault(c,[]).append(g)
    dts=[(d0+pd.Timedelta(days=k)).to_pydatetime() for k in range(nd)]; out=[]
    for x in A:
        b=x['b']; m=b['mat']
        if grp_follower(m): continue
        g_=POOLG.get(m); lt=b.get('lt_used')
        shs=[g for g in on.get(m,[]) if g.get('dem_by_date')] if not g_ else []
        if len(shs)>1:                                       # one row per customer project (WSA sheet)
            for g in shs:
                bb=g.get('bal_by_date') or {}; st=bb.get(dts[0])
                if st is None: st=bb.get(g['asof'])
                st=float(st or 0.0)
                P=[float((g.get('dem_by_date') or {}).get(d) or 0.0) for d in dts]; T=[float((g.get('ppcit_by_date') or {}).get(d) or 0.0) for d in dts]
                cur=st; S=[]; VL_=[float((g.get('vlock_by_date') or {}).get(d) or 0.0) for d in dts]   # Settings ALLOWED VENDORS (WSA column AK)
                for k in range(nd): cur=cur-P[k]+T[k]; S.append(cur-VL_[k])
                nz=[v for v in P[:min(nd,avg_prod_days())] if v>0]
                nd_=(sum(nz)/len(nz)) if nz else None
                fut=[(dts[k],T[k]) for k in range(nd) if T[k]>0]
                ata='\n'.join('   |   '.join(f"{d_:%a %d-%b}: {q_:,.0f}" for d_,q_ in fut[i_:i_+2]) for i_ in range(0,min(len(fut),ATA_SHOW),2)) or None
                r=alert_eval(b.get('common') or m,m,b.get('uom'),st,S,T,nd_,lt,sum(T),d0,where=g['sheet'],why=alert_why(S,nd_,d0,lt,T=T),ata=ata)
                if r: r.update(P=list(P),Tp=list(T),vlock=(VL_ if any(VL_) else None)); out.append(r)
            continue
        nx=x.get('nx') or {}
        sh_=[s_ for s_ in grp_ships(m,SHIPS_) if s_.get('raw') is not None]       # = cover_watch (Action List module-plan columns)
        r=alert_eval(grp_label(m,b.get('common') or m),m,b.get('uom'),g_['stock'] if g_ else float(b['stock'] or 0.0),x.get('S') or [],
                     list(g_['T']) if g_ else list(b['transit']),(None if x.get('no_demand') else x.get('pl_need')),lt,g_['tr'] if g_ else float(x.get('tr_tot') or 0.0),d0,sup=nx.get('sup'),
                     why=alert_why(x.get('S') or [],x.get('pl_need'),d0,lt,ships_=sh_),ata=x.get('ata_txt'))
        if r and g_:                                         # Cover Heatmap shows the group per material (user, Nov 2026: 'EPE 480 + EPE 420 (group), i want separate')
            mem_=[]
            for c_ in g_['members']:
                xb_=next((x2['b'] for x2 in A if x2['b']['mat']==c_),None)
                if xb_ is None: continue
                sub_=xb_.get('adhoc_sub') or {}
                T_=[(xb_['transit'][k] if k<len(xb_['transit']) else 0.0)+((xb_.get('adhoc') or [0.0]*nd)[k] if k<len(xb_.get('adhoc') or []) else 0.0)
                    -sum((a_[k] if k<len(a_) else 0.0) for sc_,a_ in sub_.items() if sc_ in g_['members']) for k in range(nd)]   # = _adh: a member's stock is no receipt
                mem_.append(dict(code=c_,name=str(xb_.get('common') or c_),uom=xb_.get('uom'),stock=float(xb_['stock'] or 0.0),P=list(xb_.get('plan') or [0.0]*nd)[:nd],T=T_,lt=xb_.get('lt_used')))
            r['mem']=mem_
        if r and g_:                                         # EPE 480 / EPE 420 SEPARATELY everywhere (user, Nov 2026: "I want epe 480 epe 420 mentioned
            # separately. Group created confusion"; answer = own + pooled date): one ALERT row per member with its OWN stock / plan / in transit /
            # class / dates; the pooled substitute date stays as a note (r['subnote']); the pooled group row itself is not listed
            for m_ in r['mem']:
                x2=next((y_ for y_ in A if y_['b']['mat']==m_['code']),None); P_=list(m_['P']); T_=list(m_['T']); cur=m_['stock']; S_=[]
                for k in range(len(P_)): cur=cur-P_[k]+(T_[k] if k<len(T_) else 0.0); S_.append(cur)
                nz=[v for v in P_[:min(len(P_),avg_prod_days())] if v>0]; need_=(sum(nz)/len(nz)) if nz else None
                lt_=m_['lt'] if m_['lt'] is not None else lt
                sh2=[s_ for s_ in SHIPS_.get(m_['code'],[]) if s_.get('raw') is not None]
                rm=alert_eval(m_['name'],m_['code'],m_['uom'],m_['stock'],S_,T_,need_,lt_,sum(T_),d0,sup=((x2 or {}).get('nx') or {}).get('sup'),
                              why=alert_why(S_,need_,d0,lt_,ships_=sh2),ata=(x2 or {}).get('ata_txt'))
                oth=[y_['name'] for y_ in r['mem'] if y_['code']!=m_['code']]
                rm.update(P=P_,Tp=T_,vlock=None,solo=True,grp=dict(label=r['name'],others=oth,first=r.get('first'),cls=r['cls']),
                          subnote=(f"substitute: with {' / '.join(oth)} together "+(f"runs out {r['first']:%d-%b}" if r.get('first') is not None else 'no stock-out in the plan')))
                out.append(rm)
            continue
        if r:                                                # plan + arrivals of the projection S (ALERT (formulas) / ALERT Calc rebuild S from them)
            vl_=(VP_MRP.get(m) or {}).get('lock') if not g_ else None      # Settings ALLOWED VENDORS (already in x['S'])
            r['vlock']=list(vl_)[:nd] if vl_ and any(v_>0 for v_ in vl_) else None
            r.update(P=([0.0]*nd if x.get('no_demand') else list((g_['P'] if g_ else b.get('plan')) or [])[:nd]),
                     Tp=list(g_['T'])[:nd] if g_ else [a_+c_ for a_,c_ in zip(b['transit'],b.get('adhoc') or [0.0]*nd)][:nd])
            out.append(r)
    for r in out: r.update(alert_stock_fields(r,d0))
    try: alert_stock_src(out,A)                                       # SAP stock (MB52) / physical count next to Estimated stock today
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: ALERT SAP stock / physical count not filled ({type(ex_).__name__}: {ex_})')
    try: alert_lines(out,A,d0,nd)                                     # plan / arrival lines behind P / Tp (ALERT (formulas) SUMIFS)
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: ALERT lines not built ({type(ex_).__name__}: {ex_})')
    for c in {r['code'] for r in out if r.get('where')}:              # 333 per WSA sheet: on-hold stock (blocked + QI) sits with the stock -
        rs=[r for r in out if r['code']==c and r.get('where')]           # the sheet that holds the supplier's stock (by stock share; none: 1st)
        tot=sum(max(0.0,float(r['stock'] or 0)) for r in rs); h=rs[0]['hold']
        for i,r in enumerate(rs): r['hold']=(h*max(0.0,float(r['stock'] or 0))/tot) if tot>0 else (h if i==0 else 0.0)
    out.sort(key=lambda r:(r['cls'],r['k0'] if r['k0'] is not None else (r['why']['t'] or 0),r['days_it'] if r['days_it'] is not None else 1e9,str(r['name'])))   # high to low risk
    return out
def _subset_match(groups, target, tol):
    """Keys of `groups` (key -> list per day) whose sum equals `target` day by day (smallest subset first); None if none / too many."""
    import itertools
    ks=[k for k,v in groups.items() if any(abs(x)>1e-9 for x in v)]
    if len(ks)>16: return None
    tt=sum(target)
    for n_ in range(0,len(ks)+1):
        for sub in itertools.combinations(ks,n_):
            if abs(sum(sum(groups[k]) for k in sub)-tt)>tol*max(1.0,len(target)): continue
            if all(abs(sum(groups[k][d] for k in sub)-target[d])<=tol for d in range(len(target))): return list(sub)
    return None
def alert_lines(rows, A, d0, nd):
    """The lines behind each ALERT row's daily Plan (P) and Arrivals (Tp), for the 'ALERT Plan lines' / 'ALERT Arrival lines' sheets
       (user, Nov 2026: "I want to see formulas in ALERT (formulas): Plan (demand), Arrivals (in transit)"):
       plan line = customer x production line x day: modules x qty per module (your BOM / MIS BOM; MIS requirement = quantity), OEM out;
       arrival line = one MIS In Transit shipment (traced code, alias factor), counted on max(arrival, today); + MIS Adhoc / substitute
       receipts per day. Substitute group = members x factor (group unit). Back glass 333 per WSA sheet: the customers / POs whose sum
       is the sheet's Demand / column T. Every row is checked: the lines must add up to P / Tp day by day; when they do not, the row
       gets one line per day with the projection value (r['pl_ok'] / r['al_ok'] = False, console NOTE)."""
    byb={x['b']['mat']:x['b'] for x in A}; oem=OEM_IS[0] or (lambda cu: False)
    tol=lambda v: max(0.01,abs(v)*1e-6)
    def plines(c,fac):
        out=[]
        for pl in (byb.get(c) or {}).get('plan_lines') or []:
            if oem(pl['cu']): continue
            for k in range(min(nd,len(pl['mods']))):
                if pl['mods'][k]: out.append(dict(code=c,k=k,cu=pl['cu'],ln=pl['ln'],src=pl['src'],mods=float(pl['mods'][k]),rate=float(pl['rate']),f=fac))
        return out
    def alines(c,fac,mem=()):
        out=[]; sub_=(byb.get(c) or {}).get('adhoc_sub') or {}
        for s_ in SHIPS_.get(c,[]):
            if s_.get('raw') is None: continue
            out.append(dict(code=c,sup=s_.get('sup') or '',po=str(s_.get('po') or '').replace('.0',''),bl=str(s_.get('bl') or ''),raw=s_['raw'],qty=float(s_['qty']),f=fac,src='MIS In Transit'))
        for k,v in enumerate(((byb.get(c) or {}).get('adhoc') or [])[:nd]):
            v-=sum((arr_[k] if k<len(arr_) else 0.0) for sc_,arr_ in sub_.items() if sc_ in mem)   # group: a member's own stock is not a receipt (= _adh)
            if abs(v)>1e-9: out.append(dict(code=c,sup='Adhoc / substitute receipt',po='',bl='',raw=d0+pd.Timedelta(days=k),qty=float(v),f=fac,src='MIS Adhoc sheet / Settings substitute (per day)'))
        return out
    def daysum(ls,qf,kf):
        t=[0.0]*nd
        for l_ in ls:
            k=kf(l_)
            if 0<=k<nd: t[k]+=qf(l_)
        return t
    pk=lambda l_: l_['k']; pq=lambda l_: l_['mods']*l_['rate']*l_['f']
    ak=lambda l_: max(0,(l_['raw']-d0).days); aq=lambda l_: l_['qty']*l_['f']
    bad=[]
    for r in rows:
        m=r['code']; g_=None if r.get('solo') else POOLG.get(m); P=list(r.get('P') or [0.0]*nd)[:nd]; T=list(r.get('Tp') or [0.0]*nd)[:nd]
        P+= [0.0]*(nd-len(P)); T+=[0.0]*(nd-len(T))
        if r.get('where'):                                               # 333 per WSA sheet: the customers / POs of this sheet
            pl_all=plines(m,1.0); al_all=alines(m,1.0)
            byc={}
            for l_ in pl_all: byc.setdefault(l_['cu'],[]).append(l_)
            sub=_subset_match({c_:daysum(v_,pq,pk) for c_,v_ in byc.items()},P,0.01)
            pls=[l_ for c_ in (sub or []) for l_ in byc[c_]] if sub is not None else None
            byp={}
            for l_ in al_all: byp.setdefault(l_['po'] or l_['src'],[]).append(l_)
            sub=_subset_match({c_:daysum(v_,aq,ak) for c_,v_ in byp.items()},T,0.01)
            als=[l_ for c_ in (sub or []) for l_ in byp[c_]] if sub is not None else None
        else:
            mem=g_['members'] if g_ else [m]; fx=dict(g_.get('fx') or {}) if g_ else {m:1.0}; f0=fx.get(m,1.0)
            smem=tuple((POOLG.get(m) or {}).get('members') or ()) if r.get('solo') else ()     # solo member: substitute receipts from the other members are not receipts (= _adh)
            pls=[l_ for c in mem for l_ in plines(c,fx.get(c,1.0)/f0)]; als=[l_ for c in mem for l_ in alines(c,fx.get(c,1.0)/f0,mem if g_ else smem)]
        if r.get('cls')==5: pls=[]                                        # no demand: no plan lines
        if pls is None or any(abs(a-b)>tol(b) for a,b in zip(daysum(pls,pq,pk),P)):
            bad.append(f"{m}{' ('+r['where']+')' if r.get('where') else ''} plan")
            pls=[dict(code=m,k=k,cu='(projection value - lines do not add up)',ln=None,src='ALERT projection',mods=P[k],rate=1.0,f=1.0) for k in range(nd) if P[k]]
        if als is None or any(abs(a-b)>tol(b) for a,b in zip(daysum(als,aq,ak),T)):
            bad.append(f"{m}{' ('+r['where']+')' if r.get('where') else ''} arrivals")
            als=[dict(code=m,sup='(projection value - lines do not add up)',po='',bl='',raw=d0+pd.Timedelta(days=k),qty=T[k],f=1.0,src='ALERT projection') for k in range(nd) if T[k]]
        r['plan_l']=pls; r['arr_l']=als
    n_p=sum(len(r['plan_l']) for r in rows); n_a=sum(len(r['arr_l']) for r in rows)
    print(f"ALERT (formulas): {n_p} plan lines, {n_a} arrival lines behind {len(rows)} ALERT rows - "+("every row adds up to the projection" if not bad else "NOT split (lines do not add up): "+'; '.join(bad)),flush=True)
HOLD_={}
def mb52_hold(mis_path):
    """On hold per MRP code = MIS MB52 Blocked (column J) + Quality Inspection (column M) - the old Material Status rule."""
    if HOLD_.get('_read') or not mis_path: return HOLD_
    HOLD_['_read']=True
    try:
        mis=wb_ro(mis_path)
        for r in list(mis['MB52'].iter_rows(values_only=True))[1:]:
            if r and len(r)>14 and r[2]:
                c=str(r[2]).strip(); m=(BOM_ALIAS[0] or {}).get(c,c)
                HOLD_[m]=HOLD_.get(m,0)+(r[12] if isinstance(r[12],(int,float)) else 0)+(r[9] if isinstance(r[9],(int,float)) else 0)
    except Exception as ex_: print(f'NOTE: MB52 on-hold stock not read ({type(ex_).__name__}: {ex_})')
    return HOLD_
def alert_stock_fields(r, d0):
    """The Material Status columns of an ALERT row (user, Nov 2026: 'combine alert and stock management into 1 sheet'):
       in transit (= what the row's projection uses), overdue in transit (MIS arrival date passed, not received), on hold
       (MB52 blocked + QI), reorder level = need/day x (transit lead time + SAFETY_DAYS), status Out of Stock / Below reorder /
       Above reorder ('Short dd-mmm' when above reorder but a stock-zero day is in the plan - never labelled safe when short).
       Substitute group = all members (in this code's unit); 333 per WSA sheet: its PPC arrivals (late ones are on today)."""
    m=r['code']; g_=None if r.get('solo') else POOLG.get(m); hold=mb52_hold(MIS_PATH_[0])      # solo = a substitute member on its own row
    fx=dict(g_.get('fx') or {}) if g_ else {m:1.0}; f0=fx.get(m,1.0)
    hq=sum(float(hold.get(c) or 0)*fx.get(c,1.0)/f0 for c in (g_['members'] if g_ else [m]))
    ovd_m=sum(s_['qty'] for s_ in (grp_ships(m,SHIPS_) if g_ else SHIPS_.get(m,[])) if s_.get('raw') is not None and s_['raw']<d0)     # material (group) total
    ovd=None if r.get('where') else ovd_m                             # per WSA sheet: overdue lines are in today's PPC arrivals
    it=float(r.get('tr_tot') or 0.0)                                  # = 'Days incl. in transit'
    rop=(r['need'] or 0.0)*(int(r['lt'] or 0)+SAFETY_DAYS); st=float(r['stock'] or 0.0)
    status='Out of Stock' if st<=0 else ('Below reorder' if st<rop else 'Above reorder')
    if status=='Above reorder' and r.get('first') is not None: status=f"Short {r['first']:%d-%b}"
    return dict(it=it,ovd=ovd,ovd_m=ovd_m,hold=hq,rop=rop,status=status)
def stock_src(b):
    """SAP stock (MB52), physical count and estimated stock today of a material - ONE rule for Stock Check, Material Status and the
       ALERT (user, Nov 2026: "Estimated stock today in ALERT = same calc logic as Stock Check, then SAP stock (MB52), Physical count").
       sap = MB52 stock (Unrestricted + QI; not counted and no MB52 line: the stock the MRP uses) · cnt = PPC physical count as counted
       (None = not counted) · est = count rolled forward to today (Stock Check rule), else the stock the MRP uses (= SAP) · own = the
       stock every shortage date uses (JW at the OEM out). est = own for every material (same number on every sheet)."""
    m=b['mat']; e=STOCK_EST.get(m) or {}
    own=float(b['stock'])-(0.0 if b.get('jw_out') else float(b.get('st_jw') or 0))
    sap=e.get('sap'); cnt=e.get('count')
    if not isinstance(sap,(int,float)): sap=(own if cnt is None else None)
    est=float(e['est']) if e.get('est') is not None else own
    return dict(sap=(float(sap) if sap is not None else None),cnt=(float(cnt) if cnt is not None else None),est=est,own=own)
def sup_split(b):
    """Stock of a material by supplier (user, Nov 2026: "Stock Check: do not combine suppliers, e.g. glass FLAT and KIBING separated"):
       [dict(name, sap, cnt, est)] when 2+ suppliers hold stock, else None. est = the supplier rows of MRP Daily (b['sup']: MB52 batches,
       or the count by supplier rolled forward); sap = MB52 batches by supplier (b['sup_sap']); cnt = the PPC count by supplier as counted
       (b['sup_cnt']). What MB52 / the count does not name goes to 'Unassigned', so the supplier rows add up to the material."""
    sups=b.get('sup') or []
    if not sups: return None
    ms=stock_src(b); est_by={}
    for sp in sups: est_by[str(sp.get('name'))]=est_by.get(str(sp.get('name')),0.0)+float(sp.get('stock') or 0.0)
    sapd={str(k):float(v) for k,v in (b.get('sup_sap') or {}).items() if k}; cntd={str(k):float(v) for k,v in (b.get('sup_cnt') or {}).items()}
    if sapd and ms['sap'] is not None and abs(ms['sap']-sum(sapd.values()))>0.5: sapd['Unassigned']=sapd.get('Unassigned',0.0)+ms['sap']-sum(sapd.values())
    if ms['cnt'] is not None and cntd and abs(ms['cnt']-sum(cntd.values()))>0.5: cntd['Unassigned']=cntd.get('Unassigned',0.0)+ms['cnt']-sum(cntd.values())
    names=list(dict.fromkeys(list(est_by)+list(sapd)+list(cntd)))
    names=[n for n in names if n!='Unassigned']+[n for n in names if n=='Unassigned']
    out=[dict(name=n,sap=(sapd.get(n,0.0) if sapd else None),cnt=(cntd.get(n,0.0) if ms['cnt'] is not None else None),est=est_by.get(n,0.0)) for n in names]
    out=[o for o in out if any(abs(v)>0.5 for v in (o['sap'] or 0.0,o['cnt'] or 0.0,o['est']))]
    return out if len(out)>=2 else None
def alert_stock_src(rows, A):
    """SAP stock (MB52) / physical count of every ALERT row (r['sap'], r['cnt']): one material = stock_src; substitute group = its members
       (in the row's unit; count = the counted members, None when none is counted); a material on 2+ WSA sheets (back glass 333) = the
       suppliers whose stock is the sheet's stock today (sup_split; r['sups']) - none found: None (shown '-')."""
    byb={x['b']['mat']:x['b'] for x in A}
    for r in rows:
        m=r['code']; g_=None if r.get('solo') else POOLG.get(m); b=byb.get(m); r['sap']=r['cnt']=None; r['sups']=None
        if b is None: continue
        if r.get('where'):
            sp=sup_split(b) or []; st=float(r['stock'] or 0.0)
            sub=_subset_match({o['name']:[o['est']] for o in sp},[st],max(1.0,0.002*abs(st))) if sp else None
            if sub is not None:
                r['sups']=list(sub); hit=[o for o in sp if o['name'] in sub]
                r['sap']=sum(o['sap'] or 0.0 for o in hit) if any(o['sap'] is not None for o in sp) else None
                r['cnt']=sum(o['cnt'] or 0.0 for o in hit) if any(o['cnt'] is not None for o in sp) else None
            continue
        mem=g_['members'] if g_ else [m]; fx=dict(g_.get('fx') or {}) if g_ else {m:1.0}; f0=fx.get(m,1.0)
        ss=[(stock_src(byb[c]),fx.get(c,1.0)/f0) for c in mem if c in byb]
        if not ss: continue
        r['sap']=sum((s_['sap'] or 0.0)*f_ for s_,f_ in ss) if any(s_['sap'] is not None for s_,_ in ss) else None
        r['cnt']=sum((s_['cnt'] or 0.0)*f_ for s_,f_ in ss) if any(s_['cnt'] is not None for s_,_ in ss) else None
ALERT_CHIP=[('FADBD8','C0392B'),('FDEBD0','B9770E'),('FCF3CF','9A7D0A'),('EAF2F8','1F618D'),('E3F4EC','1E8449'),('F2F4F6','8A949E')]   # ALERT status chips (Material Status colours)
def alert_label(code, ri=None):
    """Shortage label of a material = its ALERT class (user rule, Nov 2026: 'ALERT and Stock Check use the same label' - before,
       Stock Check used risk_info (days incl. in transit) and showed 'OK' for 280 / 268 that the ALERT lists as UNDER 21 DAYS).
       Text = ALERT class + first stock-zero day (UNDER 21 DAYS LATER: + below-21-days date); a material on 2+ WSA sheets (333) =
       one line per sheet; a substitute group member = the group's row. Not on the ALERT: OK / No demand / No need (risk_info).
       Returns (text, worst ALERT class index or None)."""
    rows=[r for r in ALERT_ROWS if r['code']==code] or [r for r in ALERT_ROWS if r['code']==grp_main(code)]   # own row first (EPE separately)
    if not rows: return ((ri or {}).get('cls') if (ri or {}).get('cls') in ('No demand','No need') else 'OK'),None
    def one(r):
        t=ALERT_CLS[r['cls']][0]; w_=r.get('why') or {}
        if r.get('first') is not None: t+=f" {r['first']:%d-%b}"
        elif r['cls']==3 and w_.get('below') is not None: t+=f" from {w_['below']:%d-%b}"
        return t+(f" ({r['where']})" if r.get('where') and len(rows)>1 else '')
    return '\n'.join(one(r) for r in rows),min(r['cls'] for r in rows)
def write_alert_sheet(wb, rows, d0, where):
    """The ALERT sheet = ALERT + Material Status in ONE sheet (user, Nov 2026: "combine alert and stock management into 1 sheet,
       style design stock management"; first tab of MRP_Daily and of the Inbound file, the same rows). Material Status design:
       row 1 = group band (STORE / INVENTORY / QUALITY / PRODUCTION / TRANSIT LEAD TIME / BELOW 21 DAYS / WAREHOUSE), row 2 = header
       coloured by group, one row per ALERT row from row 3 (Sr. + light tinted group columns + soft status chips), summary boxes +
       four counter cards and the definitions line below. Rows = the ALERT rows (every old Material Status row is on it; EPE 420 =
       the EPE 480 + 420 group row). Material Status columns added: On hold, Reorder level, Status, In transit, Overdue in transit
       (`alert_stock_fields`); its duplicates (Short from, Days of supply, Shortage risk, Days incl. in transit) = the ALERT columns.
       No Action column (user, Nov 2026)."""
    from openpyxl.styles import Border, Side
    if 'ALERT' in wb.sheetnames: del wb['ALERT']
    if not MIS_CAT and MIS_PATH_[0]: load_mis_cat(MIS_PATH_[0])     # Category (item_category)
    ws=wb.create_sheet('ALERT',0); ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor='C00000'
    F=lambda **k: Font(name='Calibri',**k)
    def fill(c,col): c.fill=PatternFill('solid',fgColor=col)
    NAVY='14213D'; INK='1F2D3D'; MUTED='8A949E'; LINE='E3E8EE'
    STORE,INV,QUAL,PROD,LEAD,BELOW,WH=(('2E86C1','F4F9FD'),('E67E22','FEF8F1'),('17A589','F2FBF8'),('6C3483','F8F4FB'),
                                       ('C0392B','FDF3F2'),('1B2631','F4F6F8'),('229954','F4FBF7'))     # Material Status palette
    thin=Side(style='thin',color=LINE); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    H=['Sr.','Alert','Code','Item name','Category','UOM',
       'Estimated stock today','Days of stock (stock ÷ need)',
       'Reorder level',
       'Status','On hold (blocked + QI)',
       'Need / day (module plan)','Days incl. in transit',
       'Transit lead time (days)','Stock-zero days inside the transit lead time','First stock-zero day','Stock zero from → to (⚠ = starts inside the transit lead time)',
       f'Below {REPORT_DAYS} days from','Why (cause)','Days late',
       'In transit','Next arrival','Overdue in transit','Arrival at plant (Final AtA)']                   # no Action column (user, Nov 2026)
    W=[5,22,16,34,16,7, 12,11, 12, 14,11, 11,10, 9,12,11,38, 11,13,8, 12,24,11,34]
    WR={2,4,10,17,22,23,24}                                                            # wrapped columns
    LEFT={4,5,17,22,24}
    for j,w_ in enumerate(W,1): ws.column_dimensions[L(j)].width=w_
    NCOL=len(H)
    G=[(f'ALERT · STOCK MANAGEMENT · {d0:%d-%b-%Y}',1,6,NAVY),('STORE · stock',7,8,STORE[0]),('INVENTORY',9,9,INV[0]),('QUALITY',10,11,QUAL[0]),
       ('PRODUCTION',12,13,PROD[0]),('TRANSIT LEAD TIME · stock-out',14,17,LEAD[0]),(f'BELOW {REPORT_DAYS} DAYS · why',18,20,BELOW[0]),('WAREHOUSE · supply',21,24,WH[0])]
    # ---------- group band (row GB) ----------
    HR=ALERT_HR; GB=HR-1
    for t,c1,c2,col in G:
        if c2>c1: ws.merge_cells(start_row=GB,start_column=c1,end_row=GB,end_column=c2)
        c=ws.cell(GB,c1,t); c.font=F(size=10,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='center',vertical='center')
        for j in range(c1,c2+1): fill(ws.cell(GB,j),col)
    ws.row_dimensions[GB].height=20; ws.row_dimensions[GB-1].height=8
    # ---------- row 2: header ----------
    hc={}; TINT={}
    for t,c1,c2,col in G:
        for j in range(c1,c2+1):
            hc[j]=col
            if col!=NAVY: TINT[j]=dict((STORE,INV,QUAL,PROD,LEAD,BELOW,WH))[col]
    for j,h in enumerate(H,1):
        c=ws.cell(HR,j,h); c.font=F(size=9,bold=True,color='FFFFFF'); fill(c,hc[j]); c.border=bd
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.row_dimensions[HR].height=44
    # ---------- data from row HR+1 ----------
    CHIP=ALERT_CHIP                                                                   # = Material Status status chips
    CC={'LATE ARRIVAL':('FDEBD0','B9770E'),'LATE + NOT ENOUGH':('FADBD8','C0392B'),'NO SUPPLY':('FADBD8','C0392B')}
    SCH={'Above reorder':('EAF2F8','1F618D'),'Below reorder':('FCF3CF','9A7D0A'),'Out of Stock':('FADBD8','C0392B')}   # Material Status Status chips
    fm=[None,None,None,None,None,None,'#,##0','0.0','#,##0',None,'#,##0','#,##0.0','0.0','0','0','dd-mmm-yy',None,'dd-mmm-yy',None,'0','#,##0',None,'#,##0',None]
    for i,r in enumerate(rows):
        rr=HR+1+i; lab=ALERT_CLS[r['cls']][0]; w_=r.get('why') or {}
        nm=r['name']+(f"\n(WSA sheet: {r['where']})" if r.get('where') else '')+(f"\n({r['subnote']})" if r.get('subnote') else '')
        nx=f"{r['nx'][0]:%d-%b}: {r['nx'][1]:,.0f}"+(f" ({str(r['sup'])[:22]})" if r.get('sup') else '') if r.get('nx') else 'none'
        ovd=r.get('ovd'); hold=r.get('hold') or 0
        vals=[i+1,lab,r['code'],nm,item_category(r['code']),r['uom'],r['stock'],r['days'],r.get('rop'),r.get('status') or '-',(hold if hold else '-'),
              r['need'],r['days_it'],r['lt'],r['zin'],r['first'].to_pydatetime() if r['first'] is not None else 'none',r['per'],
              w_['below'].to_pydatetime() if w_.get('below') is not None else 'not below',w_.get('cause') or '-',w_.get('late'),
              r.get('it') or 0,nx,('in today\'s arrivals' if ovd is None else (ovd if ovd else '-')),r.get('ata') or 'nothing on the way']
        for j,(v,f_) in enumerate(zip(vals,fm),1):
            c=ws.cell(rr,j,v); c.font=F(size=10,color=INK); c.border=Border(bottom=thin)
            c.alignment=Alignment(horizontal=('left' if j in LEFT else 'center'),vertical='center',wrap_text=j in WR)
            if f_: c.number_format=f_
            if j in TINT: fill(c,TINT[j])
        bg,fg=CHIP[r['cls']]; fill(ws.cell(rr,2),bg); ws.cell(rr,2).font=F(size=10,bold=True,color=fg)
        sb,sf=SCH.get(str(vals[9]),('FDEBD0','B9770E')); fill(ws.cell(rr,10),sb); ws.cell(rr,10).font=F(size=10,bold=True,color=sf)
        if hold: ws.cell(rr,11).font=F(size=10,bold=True,color='B9770E')
        if r['days'] is not None and r['days']<REPORT_DAYS: ws.cell(rr,8).font=F(size=10,bold=True,color='C0392B')
        if r['zin']>0: fill(ws.cell(rr,15),'FADBD8'); ws.cell(rr,15).font=F(size=12,bold=True,color='C0392B')
        else: ws.cell(rr,15).font=F(size=10,color=MUTED)
        if r['first_in']: ws.cell(rr,16).font=F(size=10,bold=True,color='C0392B')
        if w_.get('t') is not None and w_['t']<=14: ws.cell(rr,18).font=F(size=10,bold=True,color='C0392B')
        if w_.get('cause'):
            bg2,fg2=CC[w_['cause']]; fill(ws.cell(rr,19),bg2); ws.cell(rr,19).font=F(size=10,bold=True,color=fg2)
        if isinstance(w_.get('late'),int) and w_['late']>0: ws.cell(rr,20).font=F(size=10,bold=True,color='B9770E')
        if isinstance(ovd,(int,float)) and ovd>0: ws.cell(rr,23).font=F(size=10,bold=True,color='C0392B')
        nl_=max(1,str(r['per']).count('\n')+1,nm.count('\n')+1,2 if len(nx)>24 else 1,str(r.get('ata') or '').count('\n')+1,2 if len(lab)>18 else 1)
        ws.row_dimensions[rr].height=max(19,15*nl_)            # every period / arrival line visible in Excel
    if not rows: ws.cell(HR+1,1,f'No material under {REPORT_DAYS} days of stock, none below {REPORT_DAYS} days within 16 weeks and none out of stock inside its transit lead time.').font=F(size=11,bold=True,color='1E8449')
    last=HR+max(1,len(rows))
    ws.auto_filter.ref=f"A{HR}:{L(NCOL)}{last}"
    # ---------- summary boxes (like Material Status) ----------
    top=1                                                                              # boxes on top (rows 1-8)
    def box(c1,c2,title,col,lines):
        ws.merge_cells(start_row=top,start_column=c1,end_row=top,end_column=c2)
        t=ws.cell(top,c1,title); t.font=F(size=11,bold=True,color='FFFFFF'); t.alignment=Alignment(vertical='center',indent=1)
        for j in range(c1,c2+1): fill(ws.cell(top,j),col)
        for i_,(a,v,f_,vc) in enumerate(lines,1):
            ws.merge_cells(start_row=top+i_,start_column=c1,end_row=top+i_,end_column=c2-1)
            x_=ws.cell(top+i_,c1,'•  '+a); x_.font=F(size=10,color=(vc if vc and v is None else INK))
            y_=ws.cell(top+i_,c2,v); y_.font=F(size=11,bold=True,color=vc or NAVY); y_.alignment=Alignment(horizontal='right')
            if f_: y_.number_format=f_
            for j in range(c1,c2+1): fill(ws.cell(top+i_,j),'F7F9FB'); ws.cell(top+i_,j).border=Border(bottom=Side(style='thin',color='ECEFF3'))
    cnt=Counter(r['cls'] for r in rows)
    box(1,6,'⚠️  ALERT SUMMARY',NAVY,[(ALERT_CLS[k][0].capitalize(),cnt.get(k,0),'0',CHIP[k][1]) for k in range(len(ALERT_CLS))]
        +[('Materials listed',len(rows),'0',NAVY)])
    firsts=sorted([r for r in rows if r['first'] is not None],key=lambda r:(r['first'],str(r['name'])))
    box(7,18,'⏱  FIRST STOCK-ZERO DAYS (earliest first)',LEAD[0],          # 7..18 -> G:Q after the column changes (no gap before the next box)
        [(f"{r['name'][:34]}"+(f" ({r['where'][:22]})" if r.get('where') else '')+f" – {r['first']:%d-%b}"+(f" · {r['zin']} d inside transit lead time" if r['zin'] else ''),None,None,('C0392B' if r['first_in'] else 'B9770E')) for r in firsts[:5]]
        or [('No stock-zero day in the plan',None,None,'1E8449')])
    cz=Counter((r.get('why') or {}).get('cause') for r in rows if (r.get('why') or {}).get('cause'))
    hd_={}
    for r in rows: hd_[r['code']]=r.get('hold') or 0                                # on hold per material (333 rows: once)
    if False: box(14,18,'📋  INVENTORY · QUALITY',INV[0],[                          # removed with the INVENTORY / QUALITY columns (Nov 2026)('Out of stock today',sum(1 for r in rows if r.get('status')=='Out of Stock'),'0','C0392B'),
        ('Below reorder level',sum(1 for r in rows if r.get('status')=='Below reorder'),'0','9A7D0A'),
        ('Above reorder, short in the plan',sum(1 for r in rows if str(r.get('status') or '').startswith('Short')),'0','B9770E'),
        ('Materials with stock on hold',sum(1 for v in hd_.values() if v),'0','B9770E'),
        ('Overdue in transit (MIS date passed)',len({r['code'] for r in rows if (r.get('ovd_m') or 0)>0}),'0','C0392B')])
    box(19,24,f'📦  WHY BELOW {REPORT_DAYS} DAYS',BELOW[0],[('Late arrival',cz.get('LATE ARRIVAL',0),'0','B9770E'),('Late + not enough',cz.get('LATE + NOT ENOUGH',0),'0','C0392B'),
        ('No supply',cz.get('NO SUPPLY',0),'0','C0392B'),('Nothing on the way',sum(1 for r in rows if not r.get('ata')),'0',NAVY),
        ('Arrival late (days late > 0)',sum(1 for r in rows if isinstance((r.get('why') or {}).get('late'),int) and r['why']['late']>0),'0','B9770E')])
    # ---------- four cards: one per alert class ----------
    top2=last+2                                                                        # four cards + definitions stay below the table
    spans=[(1,6),(7,12),(13,18),(19,24)]
    for k,(c1,c2) in enumerate(spans):
        lab,dark,fg,txt=ALERT_CLS[k]; light=CHIP[k][0]
        for rr in range(top2,top2+4):
            ws.merge_cells(start_row=rr,start_column=c1,end_row=rr,end_column=c2)
            for j in range(c1,c2+1): fill(ws.cell(rr,j),dark if rr==top2 else light)
        a=ws.cell(top2,c1,lab); a.font=F(size=12,bold=True,color=fg); a.alignment=Alignment(horizontal='center',vertical='center')
        b_=ws.cell(top2+1,c1,txt); b_.font=F(size=10,color=CHIP[k][1]); b_.alignment=Alignment(horizontal='center',wrap_text=True)
        n_=ws.cell(top2+2,c1,cnt.get(k,0)); n_.font=F(size=24,bold=True,color=CHIP[k][1]); n_.alignment=Alignment(horizontal='center',vertical='center'); n_.number_format='0'
    ws.row_dimensions[top2].height=20; ws.row_dimensions[top2+1].height=28; ws.row_dimensions[top2+2].height=34
    dr=top2+5
    ws.merge_cells(start_row=dr,start_column=1,end_row=dr,end_column=NCOL)
    c=ws.cell(dr,1,f"Estimated stock today = the Stock Check rule: physical count (PPC Count) rolled forward to today (count − usage since the count + goods received), "
                  f"not counted: SAP stock (MB52); JW stock at the OEM left out.  SAP stock (MB52) / Physical count = the Stock Check columns (count as counted; "
                  f"substitute group = its members, back glass 333 = the suppliers of the WSA sheet).  Days of stock = estimated stock today ÷ need per day (module plan, average per production day, next 9 weeks, OEM plan left out).  "
                  f"Stock-zero day = that day's need is not covered (stock + in transit arrived − plan < 0).  EVERY material is listed, high to low risk: stock-out in transit lead time, stock-out later, "
                  f"under {REPORT_DAYS} days, under {REPORT_DAYS} days later, OK, no demand; ⚠ = the period starts inside the transit lead time.  "
                  f"Below {REPORT_DAYS} days from = first day projected stock covers < {REPORT_DAYS} days.  Why: LATE ARRIVAL = what is on the way would keep {REPORT_DAYS}+ days but "
                  f"arrives after that day; LATE + NOT ENOUGH; NO SUPPLY = nothing on the way after that day.  Days late = first arrival after that day − that day.  "
                  f"In transit = all pending shipments of the row "
                  f"(MIS In Transit; 333 per WSA sheet: its PPC arrivals); Overdue in transit = MIS arrival date passed, not received.  "
                  f"Back glass 333: one row per customer project (WSA sheet, its own supplier's stock); a substitute group (EPE 480 / EPE 420) = one row per material with its own numbers; the note under the name = the date the stock runs out with the substitute together.  "
                  f"Same sheet in MRP_Daily and in the Inbound file (WSA daily) - {where}; the weekly picture is on {HEAT_NAME} (MRP_Daily).")
    c.font=F(size=8,italic=True,color=MUTED); c.alignment=Alignment(wrap_text=True,vertical='top'); ws.row_dimensions[dr].height=60
    ws.freeze_panes=f'E{HR+1}'; ws.sheet_view.zoomScale=90; ws.print_title_rows=f'{GB}:{HR}'
    ws.page_setup.orientation='landscape'; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0; ws.sheet_properties.pageSetUpPr.fitToPage=True
    # ---------- PRODUCTION (need / day, days incl. in transit) moves next to STORE stock (user, Nov 2026) ----------
    # new column order: STORE G-H | PRODUCTION I-J | INVENTORY K | QUALITY L-M (old order 9..13 = INVENTORY, QUALITY x2, PRODUCTION x2)
    from copy import copy as _cp
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row==GB and mr.max_row==GB and 9<=mr.min_col and mr.max_col<=13: ws.unmerge_cells(str(mr))
    snap={(r_,j_):(ws.cell(r_,j_).value,_cp(ws.cell(r_,j_)._style)) for r_ in range(GB,last+1) for j_ in range(9,14)}
    for r_ in range(GB,last+1):
        for new_,old_ in zip(range(9,14),(12,13,9,10,11)):
            v_,st_=snap[(r_,old_)]; c_=ws.cell(r_,new_); c_.value=v_; c_._style=st_
    for new_,old_ in zip(range(9,14),(12,13,9,10,11)): ws.column_dimensions[L(new_)].width=W[old_-1]
    ws.merge_cells(start_row=GB,start_column=9,end_row=GB,end_column=10); ws.merge_cells(start_row=GB,start_column=12,end_row=GB,end_column=13)
    # ---------- SAP stock (MB52) + Physical count next to Estimated stock today (user, Nov 2026; = Stock Check columns) ----------
    _ins_cols(ws,8,2)
    for j,h in ((8,'SAP stock (MB52)'),(9,'Physical count')):
        c=ws.cell(HR,j,h); c._style=_cp(ws.cell(HR,7)._style)
        ws.column_dimensions[L(j)].width=12
    for i,r in enumerate(rows):
        rr=HR+1+i; one=r.get('where') and r.get('sups') is None          # 333 sheet row whose suppliers were not found: '-'
        for j,v in ((8,(r['sap'] if r.get('sap') is not None else '-')),(9,(r['cnt'] if r.get('cnt') is not None else ('-' if one else 'no counting')))):
            c=ws.cell(rr,j,v); c._style=_cp(ws.cell(rr,7)._style)
            if isinstance(v,str): c.font=F(size=10,color=MUTED)
    # ---------- no INVENTORY / QUALITY on the ALERT (user, Nov 2026: "Alert sheet, delete quality and inventory"): M Reorder level, N Status, O On hold ----------
    _del_cols(ws,[13,14,15])
    ws.auto_filter.ref=f"A{HR}:{L(NCOL+2-3)}{last}"
    return ws
def _add_uom_col(ws, uoms, top=3):
    """UOM column right after Category (column F) on Material Status / Other Materials (user, Nov 2026: "Add UOM into daily material
       status"): everything from column F moves one right; the navy band A1:E1 grows to A1:F1; header F2 'UOM'; data rows from `top`."""
    from copy import copy as _cp
    _ins_cols(ws,6,1)
    for m_ in list(ws.merged_cells.ranges):
        if m_.min_row==1 and m_.max_row==1 and m_.min_col==1 and m_.max_col==5: ws.unmerge_cells(str(m_)); ws.merge_cells('A1:F1')
    ws.cell(1,6)._style=_cp(ws.cell(1,5)._style)
    h=ws.cell(2,6,'UOM'); h._style=_cp(ws.cell(2,5)._style)
    for i,u in enumerate(uoms):
        c=ws.cell(top+i,6,(str(u).strip() if u not in (None,'') else '-')); c._style=_cp(ws.cell(top+i,5)._style)
        c.alignment=Alignment(horizontal='center',vertical='center')
    ws.column_dimensions['F'].width=7
def _del_cols(ws, cols):
    """Delete whole columns (list of indexes): cells right of them move left; merged ranges shrink (a range left with no column is
       dropped; when its first cell is deleted, its value + style go to its first remaining column); widths move left."""
    dl=sorted(set(cols)); ds=set(dl)
    if not dl: return
    nc=lambda c: c-sum(1 for d in dl if d<c)
    rngs=[(m_.min_row,m_.min_col,m_.max_row,m_.max_col) for m_ in ws.merged_cells.ranges]
    for m_ in list(ws.merged_cells.ranges): ws.unmerge_cells(str(m_))
    for r1,c1,r2,c2 in rngs:                                      # anchor deleted: keep its value on the first remaining column
        keep=[c for c in range(c1,c2+1) if c not in ds]
        if c1 in ds and keep:
            a=ws.cell(r1,c1); b=ws.cell(r1,keep[0]); b._value=a._value; b.data_type=a.data_type; b._style=a._style
    maxc=ws.max_column
    wd={j:ws.column_dimensions[L(j)].width for j in range(1,maxc+1) if L(j) in ws.column_dimensions}
    mv=sorted([(k,c) for k,c in ws._cells.items() if k[1]>dl[0]],key=lambda kc:kc[0][1])
    for (r_,c_),cell in mv: del ws._cells[(r_,c_)]
    for (r_,c_),cell in mv:
        if c_ in ds: continue
        n_=ws.cell(r_,nc(c_)); n_._value=cell._value; n_.data_type=cell.data_type; n_._style=cell._style
    for j in range(dl[0],maxc+1):
        if j in ds: continue
        if wd.get(j): ws.column_dimensions[L(nc(j))].width=wd[j]
    for r1,c1,r2,c2 in rngs:
        keep=[c for c in range(c1,c2+1) if c not in ds]
        if not keep: continue
        a_,b_=nc(keep[0]),nc(keep[-1])
        if r2>r1 or b_>a_: ws.merge_cells(start_row=r1,start_column=a_,end_row=r2,end_column=b_)
def _ins_cols(ws, at, n):
    """Insert n empty columns before column `at`: cells (value + style), merged ranges (a range across `at` grows), column widths
       move right. openpyxl's insert_cols does not move merged ranges or widths."""
    rngs=[(m_.min_row,m_.min_col,m_.max_row,m_.max_col) for m_ in ws.merged_cells.ranges]
    for m_ in list(ws.merged_cells.ranges): ws.unmerge_cells(str(m_))
    mv=sorted([(k,c) for k,c in ws._cells.items() if k[1]>=at],key=lambda kc:-kc[0][1])
    for (r_,c_),cell in mv:
        del ws._cells[(r_,c_)]
        nc=ws.cell(r_,c_+n); nc._value=cell._value; nc.data_type=cell.data_type; nc._style=cell._style
    wd={j:ws.column_dimensions[L(j)].width for j in range(at,ws.max_column+1) if L(j) in ws.column_dimensions}
    for j in sorted(wd,reverse=True):
        if wd[j]: ws.column_dimensions[L(j+n)].width=wd[j]
    for r1,c1,r2,c2 in rngs:
        ws.merge_cells(start_row=r1,start_column=c1+(n if c1>=at else 0),end_row=r2,end_column=c2+(n if c2>=at else 0))
ALERT_HR=11      # ALERT sheet: rows 1-8 = summary boxes, row 9 spacer, row 10 group band, row 11 header, data from row 12 (user, Nov 2026)
ALERT_FX="ALERT (formulas)"; ALERT_CALC="ALERT Calc"; ALERT_PL="ALERT Plan lines"; ALERT_AL="ALERT Arrival lines"; OEM_IS=[None]
def write_alert_formulas(wb, rows, d0):
    """'ALERT (formulas)' + 'ALERT Calc' (user, Nov 2026: "New Alert sheet, I want a new same sheet with excel formulas so that I can
       check your work"): the SAME rows and columns as the ALERT, but every calculated column is an Excel formula. Inputs (yellow) =
       Stock today, On hold, Transit lead time, In transit, Overdue + per day the plan and the arrivals (ALERT Calc). ALERT Calc rebuilds the
       projection: Projected stock = stock + arrivals - plan, day by day; 'Stock zero' = projected stock < 0; 'Below 21 days' = projected
       stock / need < 21 in the first 16 weeks. Column Y 'Check vs ALERT' compares each calculated column with the ALERT sheet
       (OK / DIFF + what differs in Z). Both outputs. Excel calculates on open (no cached values)."""
    from openpyxl.styles import Border, Side
    for n_ in (ALERT_FX,ALERT_CALC,ALERT_PL,ALERT_AL):
        if n_ in wb.sheetnames: del wb[n_]
    pos=wb.sheetnames.index('ALERT')+1 if 'ALERT' in wb.sheetnames else 0
    fx=wb.create_sheet(ALERT_FX,pos); cs=wb.create_sheet(ALERT_CALC,pos+1)
    for w_ in (fx,cs): w_.sheet_view.showGridLines=False
    fx.sheet_properties.tabColor='E26B0A'; cs.sheet_properties.tabColor='8A949E'
    F=lambda **k: Font(name='Calibri',**k)
    def fill(c,col): c.fill=PatternFill('solid',fgColor=col)
    NAVY='14213D'; INK='1F2D3D'; MUTED='8A949E'; LINE='E3E8EE'; INP='FFF2CC'
    thin=Side(style='thin',color=LINE); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    TG=REPORT_DAYS; HZ=112
    nd=max([len(r.get('P') or []) for r in rows]+[1])
    LC=L(5+nd)                                                     # last date column of ALERT Calc (dates from column F)
    # ---------------- ALERT Plan lines / ALERT Arrival lines (the data behind Plan and Arrivals) ----------------
    pls_=wb.create_sheet(ALERT_PL,pos+2); als_=wb.create_sheet(ALERT_AL,pos+3)
    for w_ in (pls_,als_): w_.sheet_view.showGridLines=False; w_.sheet_properties.tabColor='8A949E'
    def _lines_sheet(w_,title,note,H_,W_,rows_):
        w_['A1']=title; w_['A1'].font=F(size=16,bold=True,color=NAVY)
        w_['A2']=note; w_['A2'].font=F(size=9,italic=True,color=MUTED); w_.merge_cells(start_row=2,start_column=1,end_row=2,end_column=len(H_))
        w_['A2'].alignment=Alignment(wrap_text=True,vertical='top'); w_.row_dimensions[2].height=40
        for j,(h,wd) in enumerate(zip(H_,W_),1):
            c=w_.cell(4,j,h); fill(c,NAVY); c.font=F(size=9,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
            w_.column_dimensions[L(j)].width=wd
        w_.row_dimensions[4].height=32
        fi_=PatternFill('solid',fgColor=INP)                          # one style object for all cells (16,000 rows: speed)
        for i_,(vals,fms,inp) in enumerate(rows_,5):
            for j,(v_,f_) in enumerate(zip(vals,fms),1):
                c=w_.cell(i_,j,v_)
                if f_: c.number_format=f_
                if j in inp: c.fill=fi_
        w_.freeze_panes='C5'; w_.auto_filter.ref=f"A4:{L(len(H_))}{4+max(1,len(rows_))}"
        return 4+max(1,len(rows_))
    prow=[]; arow=[]
    for i,r in enumerate(rows):
        for l_ in r.get('plan_l') or []:
            n_=5+len(prow)
            prow.append(([i+1,r['code'],l_['code'],(d0+pd.Timedelta(days=l_['k'])).to_pydatetime(),l_['cu'],l_['ln'],l_['src'],l_['mods'],l_['rate'],l_['f'],f'=H{n_}*I{n_}*J{n_}'],
                         [None,None,None,'dd-mmm-yy',None,None,None,'#,##0.##','0.######','0.####','#,##0.##'],{8,9,10}))
        for l_ in r.get('arr_l') or []:
            n_=5+len(arow)
            arow.append(([i+1,r['code'],l_['code'],l_['sup'],(int(l_['po']) if str(l_['po']).isdigit() else (l_['po'] or None)),l_['bl'] or None,l_['raw'].to_pydatetime(),f"=MAX(G{n_},'{ALERT_CALC}'!$F$3)",l_['qty'],l_['f'],f'=I{n_}*J{n_}',l_['src']],
                         [None,None,None,None,None,None,'dd-mmm-yy','dd-mmm-yy','#,##0.##','0.####','#,##0.##',None],{7,9,10}))
    PL_last=_lines_sheet(pls_,"ALERT Plan lines - the plan behind 'Plan (demand)' on ALERT Calc",
        f"One row per ALERT row (Sr.) x material x day x customer x production line, from the module plan (PPC Plan + Revised Plan; OEM customers left out). "
        f"Demand = Modules x Qty per module x Factor (Excel formula). Qty per module = your PPC BOM, else the MIS Standard BOM (column G says which); 'MIS requirement' lines = "
        f"MIS quantity (Modules = quantity, Qty per module = 1). Factor = unit of the ALERT row (substitute group: member -> main material). Yellow = input. "
        f"ALERT Calc Plan = SUMIFS(Demand, Sr., Date).",
        ['Sr. (ALERT row)','ALERT code','Material code','Date','Customer','Production line','Qty per module from','Modules','Qty per module','Factor','Demand (formula)'],
        [8,16,16,11,24,16,30,11,11,8,12],prow)
    AL_last=_lines_sheet(als_,"ALERT Arrival lines - the shipments behind 'Arrivals (in transit)' on ALERT Calc",
        f"One row per ALERT row (Sr.) x shipment: MIS In Transit (pending qty, Final AtA, material traced to your BOM / WSA PO + BL#; alias codes converted), plus MIS Adhoc / "
        f"substitute receipts per day. Counted on = MAX(arrival, today) - late lines are counted today (Excel formula). Qty = Pending qty x Factor (formula). "
        f"Back glass 333 per WSA sheet: the POs of that sheet. ALERT Calc Arrivals = SUMIFS(Qty, Sr., Counted on); In transit = all lines of the row up to the plan end.",
        ['Sr. (ALERT row)','ALERT code','Material code','Supplier','PO','BL / HBL','Arrival at plant (MIS)','Counted on (formula)','Pending qty (MRP unit)','Factor','Qty (formula)','Source'],
        [8,16,16,30,12,18,12,12,13,8,13,26],arow)
    PLK=f"'{ALERT_PL}'!$A$5:$A${PL_last}"; PLD=f"'{ALERT_PL}'!$D$5:$D${PL_last}"; PLQ=f"'{ALERT_PL}'!$K$5:$K${PL_last}"
    ALK=f"'{ALERT_AL}'!$A$5:$A${AL_last}"; ALD=f"'{ALERT_AL}'!$H$5:$H${AL_last}"; ALQ=f"'{ALERT_AL}'!$K$5:$K${AL_last}"
    # ---------------- ALERT Calc ----------------
    cs['A1']='ALERT Calc - the daily projection behind the ALERT (Excel formulas)'; cs['A1'].font=F(size=16,bold=True,color=NAVY)
    cs['A2']=(f"Per ALERT row: Plan = SUMIFS of '{ALERT_PL}' (modules x qty per module per customer / line / day, OEM left out; substitute members on their own rows; "
              f"back glass 333 = the customers of the WSA sheet) and Arrivals = SUMIFS of '{ALERT_AL}' (MIS shipments counted on max(arrival, today); 333 = the POs of the sheet). "
              f"Projected stock = previous day + arrivals - plan (day 1: Estimated stock today on '{ALERT_FX}'). Stock zero = 1 when projected stock < 0. "
              f"Below {TG} days = 1 when projected stock / need per day < {TG} (first 16 weeks = day number < {HZ}).")
    cs['A2'].font=F(size=9,italic=True,color=MUTED); cs.merge_cells('A2:N2'); cs['A2'].alignment=Alignment(wrap_text=True,vertical='top'); cs.row_dimensions[2].height=42
    for j,h in enumerate(['Sr.','Code','Item name','Line','Estimated stock today'],1):
        for rr_,v_ in ((3,h),(4,'Day number' if j==4 else None)):
            c=cs.cell(rr_,j,v_); fill(c,NAVY); c.font=F(size=9,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    for k in range(nd):
        c=cs.cell(3,6+k,(d0+pd.Timedelta(days=k)).to_pydatetime()); c.number_format='dd-mmm'; fill(c,NAVY); c.font=F(size=9,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='center')
        c=cs.cell(4,6+k,k); fill(c,'D6DBDF'); c.font=F(size=8,color=INK); c.alignment=Alignment(horizontal='center')
        cs.column_dimensions[L(6+k)].width=9
    for j,w_ in enumerate([5,16,30,26,11],1): cs.column_dimensions[L(j)].width=w_
    DT=f"'{ALERT_CALC}'!$F$3:${LC}$3"; IDX=f"'{ALERT_CALC}'!$F$4:${LC}$4"
    F8=F(size=8,color=INK); FB8=F(size=8,color=INK,bold=True); FIL_=PatternFill('solid',fgColor='EAF2F8'); BL_=Border(bottom=Side(style='thin',color=LINE)); BL2_=Border(bottom=Side(style='thin',color='8A949E'))
    LINES=['Plan (demand)','Arrivals (in transit)','Projected stock',f'Stock zero (1 = yes)',f'Below {TG} days (1 = yes, first 16 weeks)']
    for i,r in enumerate(rows):
        b0=5+5*i; fr=3+i
        for q,ln in enumerate(LINES):
            rr=b0+q
            ln_=(ln+' - Settings ALLOWED VENDORS: minus stock / arrivals no short customer may use' if (q==2 and r.get('vlock')) else ln)
            for j,v_ in enumerate([i+1,r['code'],r['name']+(f" ({r['where']})" if r.get('where') else ''),ln_],1):
                c=cs.cell(rr,j,v_ if (q==0 or j==4 or (j==1 and q==1)) else None); c.font=F(size=9,bold=(j==4 and q==2),color=INK); c.alignment=Alignment(vertical='center')
            if q==0:
                c=cs.cell(rr,5,f"='{ALERT_FX}'!G{fr}"); c.number_format='#,##0'; c.font=F(size=9,color=INK)
            for k in range(nd):
                col=L(6+k); prv=L(5+k)
                if q==0: v_=f'=SUMIFS({PLQ},{PLK},$A{b0},{PLD},{col}$3)'
                elif q==1: v_=f'=SUMIFS({ALQ},{ALK},$A{b0},{ALD},{col}$3)'
                elif q==2:
                    v_=(f"=$E{b0}+{col}{b0+1}-{col}{b0}" if k==0 else f"={prv}{rr}+{col}{b0+1}-{col}{b0}")
                    vl_=r.get('vlock')                           # Settings ALLOWED VENDORS: stock + arrivals no short customer may use (change vs the day before)
                    dl_=((vl_[k] if k<len(vl_) else 0.0)-((vl_[k-1] if k-1<len(vl_) else 0.0) if k else 0.0)) if vl_ else 0.0
                    if abs(dl_)>1e-9: v_+=f"-({round(dl_,6)})"
                elif q==3: v_=f"=IF({col}{b0+2}<0,1,0)"
                else: v_=f"=IF(AND({col}$4<{HZ},'{ALERT_FX}'!$L{fr}>0),IF({col}{b0+2}/'{ALERT_FX}'!$L{fr}<{TG},1,0),0)"
                c=cs.cell(rr,6+k,v_); c.font=(FB8 if q==2 else F8); c.number_format='#,##0' if q<3 else '0'
                if q<2: c.fill=FIL_
            for j in range(1,6+nd): cs.cell(rr,j).border=(BL2_ if q==4 else BL_)
    cs.freeze_panes='F5'
    # ---------------- ALERT (formulas) ----------------
    H=['Sr.','Alert','Code','Item name','Category','UOM','Estimated stock today','Days of stock (stock ÷ need)','Reorder level','Status','On hold (blocked + QI)',
       'Need / day (module plan)','Days incl. in transit','Transit lead time (days)','Stock-zero days inside the transit lead time','First stock-zero day',
       'Stock-zero days in the plan (count)',f'Below {TG} days from','Why (cause)','Days late','In transit','Next arrival','Overdue in transit',
       'In transit arriving after the plan end','Check vs ALERT','What differs',
       f'Helper: day number below {TG} days','Helper: arrivals after that day',f'Helper: first arrival day after it',f'Helper: gap to {TG} days that day','Helper: next arrival day']
    W=[5,22,16,30,14,7,12,11,12,14,11,11,10,9,12,11,11,11,13,8,12,18,11,12,10,30,10,12,10,12,10]
    for j,w_ in enumerate(W,1): fx.column_dimensions[L(j)].width=w_
    G=[('ALERT (formulas) · same rows as ALERT · yellow = input, white = Excel formula',1,6,NAVY),('STORE · stock',7,8,'2E86C1'),('INVENTORY',9,9,'E67E22'),
       ('QUALITY',10,11,'17A589'),('PRODUCTION',12,13,'6C3483'),('TRANSIT LEAD TIME · stock-out',14,17,'C0392B'),(f'BELOW {TG} DAYS · why',18,20,'1B2631'),
       ('WAREHOUSE · supply',21,24,'229954'),('CHECK',25,26,'C00000'),('HELPER (formulas)',27,31,'8A949E')]
    hc={}
    for t,c1,c2,col in G:
        if c2>c1: fx.merge_cells(start_row=1,start_column=c1,end_row=1,end_column=c2)
        c=fx.cell(1,c1,t); c.font=F(size=10,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='center',vertical='center')
        for j in range(c1,c2+1): fill(fx.cell(1,j),col); hc[j]=col
    for j,h in enumerate(H,1):
        c=fx.cell(2,j,h); c.font=F(size=9,bold=True,color='FFFFFF'); fill(c,hc[j]); c.border=bd; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    fx.row_dimensions[1].height=20; fx.row_dimensions[2].height=44
    INPUTS={7,11,14,23}
    fm={7:'#,##0',8:'0.0',9:'#,##0',11:'#,##0',12:'#,##0.0',13:'0.0',14:'0',15:'0',16:'dd-mmm-yy',17:'0',18:'dd-mmm-yy',20:'0',21:'#,##0',23:'#,##0',24:'#,##0',27:'0',28:'#,##0',29:'0',30:'#,##0',31:'0'}
    AO=ALERT_HR-2; A_="ALERT!"
    for i,r in enumerate(rows):
        n=3+i; b0=5+5*i
        PL=f"'{ALERT_CALC}'!$F${b0}:${LC}${b0}"; AR=f"'{ALERT_CALC}'!$F${b0+1}:${LC}${b0+1}"; SR=f"'{ALERT_CALC}'!$F${b0+2}:${LC}${b0+2}"
        ZR=f"'{ALERT_CALC}'!$F${b0+3}:${LC}${b0+3}"; BR=f"'{ALERT_CALC}'!$F${b0+4}:${LC}${b0+4}"
        ovd=r.get('ovd')
        f_={1:i+1,3:r['code'],4:r['name']+(f"\n(WSA sheet: {r['where']})" if r.get('where') else ''),5:item_category(r['code']),6:r['uom'],
            7:float(r['stock'] or 0.0),11:float(r.get('hold') or 0.0),14:int(r['lt'] or 0),
            21:f'=SUMIFS({ALQ},{ALK},$A{n},{ALD},"<="&\'{ALERT_CALC}\'!${LC}$3)',
            23:("in today's arrivals" if ovd is None else float(ovd or 0.0)),
            2:f'=IF(L{n}=0,\"{ALERT_CLS[5][0]}\",IF(O{n}>0,\"{ALERT_CLS[0][0]}\",IF(ISNUMBER(P{n}),\"{ALERT_CLS[1][0]}\",IF(H{n}<{TG},\"{ALERT_CLS[2][0]}\",IF(ISNUMBER(R{n}),\"{ALERT_CLS[3][0]}\",\"{ALERT_CLS[4][0]}\")))))',
            12:f'=IFERROR(AVERAGEIFS({PL},{PL},">0",{IDX},"<{avg_prod_days()}"),0)',
            8:f'=IF(L{n}>0,MAX(0,G{n})/L{n},"")',
            9:f'=L{n}*(N{n}+{SAFETY_DAYS})',
            10:f'=IF(G{n}<=0,"Out of Stock",IF(G{n}<I{n},"Below reorder",IF(ISNUMBER(P{n}),"Short "&TEXT(P{n},"dd-mmm"),"Above reorder")))',
            13:f'=IF(L{n}>0,MAX(0,G{n}+U{n})/L{n},"")',
            15:f'=IF(N{n}>0,SUMIFS({ZR},{IDX},"<"&N{n}),0)',
            16:f'=IFERROR(INDEX({DT},MATCH(1,{ZR},0)),"none")',
            17:f'=SUM({ZR})',
            27:f'=IFERROR(MATCH(1,{BR},0)-1,"")',
            18:f'=IF(AA{n}="","not below",INDEX({DT},AA{n}+1))',
            28:f'=IF(AA{n}="","",SUMIFS({AR},{IDX},">"&AA{n}))',
            29:f'=IF(AA{n}="","",IF(COUNTIFS({AR},">0",{IDX},">"&AA{n})=0,"",_xlfn.MINIFS({IDX},{AR},">0",{IDX},">"&AA{n})))',
            30:f'=IF(AA{n}="","",{TG}*L{n}-INDEX({SR},AA{n}+1))',
            19:f'=IF(AA{n}="","-",IF(AC{n}="","NO SUPPLY",IF(AB{n}>=AD{n},"LATE ARRIVAL","LATE + NOT ENOUGH")))',
            20:f'=IF(OR(AA{n}="",AC{n}=""),"",AC{n}-AA{n})',
            31:f'=IF(COUNTIF({AR},">0")=0,"",_xlfn.MINIFS({IDX},{AR},">0"))',
            22:f'=IF(AE{n}="","none",TEXT(INDEX({DT},AE{n}+1),"dd-mmm")&": "&TEXT(INDEX({AR},AE{n}+1),"#,##0"))',
            24:f'=SUMIFS({ALQ},{ALK},$A{n},{ALD},">"&\'{ALERT_CALC}\'!${LC}$3)'}
        cl=lambda a,b: f'ABS(N({a})-N({b}))<=MAX(0.01,ABS(N({b}))*0.000001)'
        # ALERT columns: H:I = SAP stock (MB52) + Physical count (not here), no Reorder level / Status / On hold (Nov 2026) -> letters differ from this sheet's
        parts=[(f'B{n}=ALERT!B{n+AO}','Alert'),(cl(f'H{n}',f'ALERT!J{n+AO}'),'Days of stock'),(cl(f'L{n}',f'ALERT!K{n+AO}'),'Need / day'),(cl(f'M{n}',f'ALERT!L{n+AO}'),'Days incl. in transit'),(f'O{n}=ALERT!N{n+AO}','Stock-zero days in transit lead time'),
               (f'IF(ISNUMBER(P{n}),P{n}=ALERT!O{n+AO},ALERT!O{n+AO}="none")','First stock-zero day'),(f'IF(ISNUMBER(R{n}),R{n}=ALERT!Q{n+AO},ALERT!Q{n+AO}="not below")',f'Below {TG} days from'),
               (f'S{n}=ALERT!R{n+AO}','Why'),(cl(f'U{n}',f'ALERT!T{n+AO}'),'In transit'),(f'N(T{n})=N(ALERT!S{n+AO})','Days late'),(f'LEFT(ALERT!U{n+AO},LEN(V{n}))=V{n}','Next arrival')]
        f_[26]='=TRIM('+'&'.join(f'IF({c_},"","{t_}; ")' for c_,t_ in parts)+')'
        f_[25]=f'=IF(Z{n}="","OK","DIFF")'
        for j in range(1,32):
            c=fx.cell(n,j,f_.get(j)); c.font=F(size=10,color=(MUTED if j>=27 else INK)); c.border=Border(bottom=thin)
            c.alignment=Alignment(horizontal=('left' if j in (4,5,22,26) else 'center'),vertical='center',wrap_text=j in (2,4,10,23,26))
            if j in fm: c.number_format=fm[j]
            if j in INPUTS: fill(c,INP)
        fx.row_dimensions[n].height=max(19,15*(f_[4].count('\n')+1),30)
    last=2+max(1,len(rows))
    from openpyxl.formatting.rule import CellIsRule
    for k_,(lab_,_,_,_) in enumerate(ALERT_CLS):                   # Alert chips = the ALERT sheet colours
        fx.conditional_formatting.add(f"B3:B{last}",CellIsRule(operator='equal',formula=[f'"{lab_}"'],fill=PatternFill('solid',fgColor=ALERT_CHIP[k_][0]),font=Font(name='Calibri',bold=True,color=ALERT_CHIP[k_][1])))
    fx.conditional_formatting.add(f"Y3:Y{last}",CellIsRule(operator='equal',formula=['"OK"'],fill=PatternFill('solid',fgColor='D5F5E3'),font=Font(name='Calibri',bold=True,color='1E8449')))
    fx.conditional_formatting.add(f"Y3:Y{last}",CellIsRule(operator='equal',formula=['"DIFF"'],fill=PatternFill('solid',fgColor='FADBD8'),font=Font(name='Calibri',bold=True,color='C0392B')))
    tr_=last+2
    fx.cell(tr_,1,'Rows checked').font=F(size=10,bold=True,color=NAVY); fx.cell(tr_,4,f'=COUNTA(Y3:Y{last})').font=F(size=10,bold=True,color=NAVY)
    fx.cell(tr_+1,1,'OK (same as ALERT)').font=F(size=10,bold=True,color='1E8449'); fx.cell(tr_+1,4,f'=COUNTIF(Y3:Y{last},"OK")').font=F(size=10,bold=True,color='1E8449')
    fx.cell(tr_+2,1,'DIFF').font=F(size=10,bold=True,color='C0392B'); fx.cell(tr_+2,4,f'=COUNTIF(Y3:Y{last},"DIFF")').font=F(size=10,bold=True,color='C0392B')
    nt=fx.cell(tr_+4,1,(f"Yellow = input (Estimated stock today = PPC count rolled forward, else SAP (MB52), JW at the OEM left out - see Stock Check / Stock Data; On hold = MIS MB52 blocked + QI; "
                        f"Transit lead time = Settings; Overdue = MIS arrival date passed). In transit = SUMIFS of '{ALERT_AL}' up to the plan end; plan + arrivals per day on '{ALERT_CALC}' = SUMIFS of '{ALERT_PL}' / '{ALERT_AL}'. "
                        f"White = Excel formula. Need / day = average plan per production day (plan > 0) in the next 9 weeks. Days of stock = stock / need. "
                        f"Reorder level = need × (transit lead time + {SAFETY_DAYS}). Stock-zero day = projected stock < 0; inside the transit lead time = day number < transit lead time. "
                        f"Below {TG} days from = first day projected stock / need < {TG} (16 weeks). Why: NO SUPPLY = no arrival after that day; LATE ARRIVAL = arrivals after "
                        f"that day ≥ the gap to {TG} days; else LATE + NOT ENOUGH. Days late = first arrival after that day − that day. Back glass 333: one row per WSA sheet; "
                        f"a substitute group = one row per member (own numbers). Check vs ALERT = every calculated column compared with the ALERT sheet (same row; the ALERT has "
                        f"SAP stock (MB52) + Physical count in H:I and no Reorder level / Status / On hold, so its column letters differ from this sheet's)."))
    nt.font=F(size=8,italic=True,color=MUTED); nt.alignment=Alignment(wrap_text=True,vertical='top')
    fx.merge_cells(start_row=tr_+4,start_column=1,end_row=tr_+4,end_column=26); fx.row_dimensions[tr_+4].height=64
    fx.freeze_panes='E3'; fx.sheet_view.zoomScale=90; fx.auto_filter.ref=f"A2:{L(31)}{last}"
    try: wb.calculation.fullCalcOnLoad=True                      # Excel calculates every formula when the file opens
    except Exception as _e_: _swallow(4540,_e_); pass
    return fx
ETD_LATE=[]        # WSA delivery rows not on MIS In Transit, ETD passed or to come (built with the Inbound file; Inbound + MRP_Daily + daily email)
ETD_SHEET='ETD alert'
ETD_SOON=7         # 'Ships within 7 days': ETD this close and still not on MIS In Transit
WSA_SHIPPER={}     # material code -> shipper names typed on your WSA delivery rows
def _etd_arow(r):
    """The ALERT row of a WSA delivery row's material (333: the row of that WSA sheet); None when not on the ALERT."""
    mrp_m={str(x_['b']['mat']).strip() for x_ in (AL_ORDER_ or []) if not x_.get('sup_row')}
    for c_ in ([c_ for c_ in r['codes'] if c_ in mrp_m] or list(r['codes'])):
        al=[a_ for a_ in ALERT_ROWS if a_['code']==c_] or [a_ for a_ in ALERT_ROWS if a_['code']==grp_main(c_)]
        if len(al)>1: al=[a_ for a_ in al if a_.get('where')==r['sheet']] or al
        if al: return al[0]
    return None
def _etd_alert(r):
    """ALERT class + first stock-zero day of the planning sheet's material (333: the row of that WSA sheet)."""
    mrp_m={str(x_['b']['mat']).strip() for x_ in (AL_ORDER_ or []) if not x_.get('sup_row')}
    cs=[c_ for c_ in r['codes'] if c_ in mrp_m] or list(r['codes'])
    for c_ in cs:
        al=[a_ for a_ in ALERT_ROWS if a_['code']==c_] or [a_ for a_ in ALERT_ROWS if a_['code']==grp_main(c_)]
        if len(al)>1: al=[a_ for a_ in al if a_.get('where')==r['sheet']] or al
        if al:
            a_=al[0]
            return c_,str(a_['name']).split('\n')[0],a_['cls'],(a_['first'] if a_.get('first') is not None else None),(a_.get('uom') or r.get('uom'))
    return (cs[0] if cs else '-'),'-',None,None,r.get('uom')
ETD_FLAGS=[('ARRIVES AFTER STOCK-OUT','C00000','FFFFFF'),('ETD PASSED','E67E22','FFFFFF'),(f'SHIPS WITHIN {ETD_SOON} DAYS','F4D03F','1F2D3D'),('ON PLAN','1E8449','FFFFFF')]
def _etd_flag(r, first):
    """0 = arrives after the stock-out (expected arrival after the material's first stock-zero day), 1 = ETD passed, 2 = ETD within
       ETD_SOON days, 3 = on plan."""
    if first is not None and r['eta_now']>pd.Timestamp(first).normalize(): return 0
    if r['passed']: return 1
    return 2 if -r['late']<=ETD_SOON else 3
def write_etd_late_sheet(wb, rows, where):
    """Sheet 'ETD alert' (Inbound and MRP_Daily, the same rows): your WSA delivery rows (not cancelled) with an ETD whose BL# is not on
       MIS In Transit, ETD passed AND to come. Alert: ARRIVES AFTER STOCK-OUT (expected arrival after the material's first stock-zero
       day), ETD PASSED, SHIPS WITHIN ETD_SOON DAYS, ON PLAN. ETD vs today: 'Late N d' / 'In N d'. Expected arrival = WSA plant date; ETD
       passed = base date + planned transit (WSA plant date - ETD) if it ships today; delay = that - WSA plant date. With the material's
       ALERT class and first stock-zero day. Not in any projection (the outputs use MIS In Transit) - information to chase."""
    from openpyxl.styles import Border, Side
    for n_ in (ETD_SHEET,'ETD late'):
        if n_ in wb.sheetnames: del wb[n_]
    ix_=max([i_+1 for i_,n_ in enumerate(wb.sheetnames) if n_.startswith('ALERT')]+[0])
    ws=wb.create_sheet(ETD_SHEET,ix_); ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor='C0392B'
    F=lambda **k: Font(name='Calibri',**k); thin=Side(style='thin',color='D0D5DD'); BD=Border(bottom=thin)
    base_=rows[0]['base'] if rows else None
    info=[(r,)+_etd_alert(r) for r in rows]
    info=[t_+(_etd_flag(t_[0],t_[4]),) for t_ in info]
    info.sort(key=lambda t_:(t_[6],-t_[0]['late'] if t_[0]['passed'] else 0,t_[0]['etd'],t_[0]['label']))
    cnt=Counter(t_[6] for t_ in info)
    ws['A1']='ETD alert - WSA deliveries not in transit yet'; ws['A1'].font=F(size=20,bold=True,color='1F3864')
    ws['A2']=(f"Your WSA delivery rows (not cancelled) with an ETD whose BL# is not on MIS In Transit (base date {base_:%d-%b-%Y}). " if base_ is not None else "Your WSA delivery rows (not cancelled) with an ETD whose BL# is not on MIS In Transit. ")+\
             (f"Alert: ARRIVES AFTER STOCK-OUT = expected arrival after the material's first stock-zero day; ETD PASSED; SHIPS WITHIN {ETD_SOON} DAYS; ON PLAN. "
              "Expected arrival = your WSA plant date; ETD passed = base date + planned transit (plant date - ETD) if it ships today. "
              "Not in the projections (they use MIS In Transit). "+where)
    ws['A2'].font=F(size=9,color='5D6D7E')
    ws['A3']='   '.join(f"{lab_}: {cnt.get(k_,0)}" for k_,(lab_,_,_) in enumerate(ETD_FLAGS)); ws['A3'].font=F(size=10,bold=True,color='1F2D3D')
    H=['Sr.','Alert','ETD vs today','Status','Planning sheet','Code','Item name','ALERT','First stock-zero day','PO','Shipper','Quantity','UOM','ETD',
       'WSA plant date','Planned transit (days)','Expected arrival','Delay vs WSA plant date (days)','BL#','Remarks (as written)']
    for j_,h_ in enumerate(H,1):
        c=ws.cell(4,j_,h_); c.font=F(size=10,bold=True,color='1F2D3D'); c.fill=PatternFill('solid',fgColor='F4F6F8')
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=Border(bottom=Side(style='medium',color='8A949E'))
    ws.row_dimensions[4].height=32
    CH={0:('C00000','FFFFFF'),1:('E67E22','FFFFFF'),2:('F4D03F','1F2D3D'),3:('5DADE2','FFFFFF'),4:('1E8449','FFFFFF'),5:('8A949E','FFFFFF')}
    for i_,(r,code,name,cls,first,uom_,fl) in enumerate(info):
        rr=5+i_
        when=(f"Late {r['late']} d" if r['passed'] else ('Today' if r['late']==0 else f"In {-r['late']} d"))
        v=[i_+1,ETD_FLAGS[fl][0],when,r['status'],r['sheet'],code,name,(ALERT_SHORT[cls] if cls is not None and cls<len(ALERT_SHORT) else '-'),
           (first.to_pydatetime() if first is not None else 'none'),r['po'] or '-',r['shipper'] or '-',r['qty'],uom_ or '',r['etd'].to_pydatetime(),
           r['plant'].to_pydatetime(),r['transit'],r['eta_now'].to_pydatetime(),r['delay'],r['bl'] or '-',r['remarks'] or None]
        for j_,x_ in enumerate(v,1):
            c=ws.cell(rr,j_,x_); c.font=F(size=10); c.border=BD; c.alignment=Alignment(vertical='center',horizontal=('left' if j_ in (4,5,7,11,19,20) else 'center'))
            if hasattr(x_,'year'): c.number_format='dd-mmm-yy'
            if j_==12: c.number_format='#,##0'
        c=ws.cell(rr,2); c.fill=PatternFill('solid',fgColor=ETD_FLAGS[fl][1]); c.font=F(size=9,bold=True,color=ETD_FLAGS[fl][2])
        if r['passed']: ws.cell(rr,3).font=F(size=11,bold=True,color=('C0392B' if r['late']>=14 else 'B9770E'))
        if cls is not None:
            bg,fg=CH.get(cls,('8A949E','FFFFFF')); c=ws.cell(rr,8); c.fill=PatternFill('solid',fgColor=bg); c.font=F(size=9,bold=True,color=fg)
        if fl==0: ws.cell(rr,17).font=F(size=10,bold=True,color='C0392B')     # arrives after the stock-out
        if r['status'].startswith('BL#'): ws.cell(rr,4).font=F(size=10,color='9A7D0A')
        ws.row_dimensions[rr].height=18
    if not rows: ws.cell(5,1,'Every WSA delivery row with an ETD is on MIS In Transit (or cancelled).').font=F(size=10)
    for j_,w_ in enumerate([5,24,11,30,30,15,26,14,12,13,24,11,6,10,11,9,11,11,18,30],1): ws.column_dimensions[L(j_)].width=w_
    ws.freeze_panes='D5'; ws.auto_filter.ref=f"A4:T{4+max(1,len(rows))}"
    return len(rows)
def add_alert_inbound(path, rows, d0):
    """The same ALERT sheet as the first tab of the Inbound file (WSA daily)."""
    wb=load_workbook(path)
    write_alert_sheet(wb,rows,d0,'this is the Inbound copy')
    write_alert_formulas(wb,rows,d0)
    try: write_etd_late_sheet(wb,ETD_LATE,'This is the Inbound copy (the same rows as MRP_Daily).')
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: ETD late sheet not added to the Inbound file ({type(ex_).__name__}: {ex_})')
    if 'Read me' in wb.sheetnames:
        rd_=wb['Read me']; r_=rd_.max_row+2
        rd_.cell(r_,1,f"ALERT (first tab) - the same list as the ALERT sheet of MRP_Daily: materials under {REPORT_DAYS} days of stock, below {REPORT_DAYS} days within 16 weeks or out of stock, why (late arrival / not enough / no supply), "
                     "the number of stock-zero days inside the transit lead time, every stock-zero period (from → to), the next arrival, plus the Material Status columns (in transit, overdue).").font=Font(name='Calibri',size=10)
    for w_ in wb.worksheets: w_.sheet_view.tabSelected=False
    wb.active=0; wb.worksheets[0].sheet_view.tabSelected=True
    _safe_save(wb,path)

def cover_watch(x, d0, ships, nd, need=None):
    """ONE calculation for the Risk Charts list AND the Action List columns 'Below 21 Days From' / 'Short From (Stock = 0)' / 'Days Late'.
       Below 21 Days From = first day projected stock (rolled forward + MIS in transit by arrival date - plan) covers fewer than REPORT_DAYS days of need.
       Short From = first day projected stock is below zero. Days Late = arrival of the first shipment that comes AFTER the below-21-days day, minus that day.
       Returns None when the material has no demand (never shown as a shortage) or no need per day."""
    if x.get('no_demand'): return None
    TGT=REPORT_DAYS; HZ=min(nd,112); b=x['b']; m=b['mat']; S=x.get('S') or []
    need=need or x.get('pl_need')                            # default: module plan need per day; the Action List passes the daily-target need
    if not need or not S: return None
    cov=[v/need for v in S[:HZ]]
    t=next((k for k in range(len(cov)) if cov[k]<TGT),None)
    ri=risk_info(x,d0)
    out=dict(need=need,cov=cov,cov0=cov[0],t=t,ri=ri,short=ri['date'],below=None,late=None,nx=None,cause=None,gap_all=0.0,ov=[],ovq=0.0)
    sh=[s_ for s_ in grp_ships(m,ships) if s_.get('raw') is not None]          # substitute group: all members' shipments
    out['ov']=[s_ for s_ in sh if s_['raw']<d0]; out['ovq']=sum(s_['qty'] for s_ in out['ov'])
    if t is None: return out
    dt_=d0+pd.Timedelta(days=t); out['below']=dt_
    later=sorted([s_ for s_ in sh if s_['raw']>dt_],key=lambda s_:s_['raw'])
    gap=TGT*need-S[t]; on_way=sum(s_['qty'] for s_ in later)
    out['gap_all']=max(0.0,TGT*need-min(S[t:HZ]))                     # deepest gap in the 16 weeks (back to 21 days); arrivals already in S
    out['cause']='NO SUPPLY' if not later else ('LATE ARRIVAL' if on_way>=gap else 'LATE + NOT ENOUGH')
    out['nx']=later[0] if later else None
    out['late']=(later[0]['raw']-dt_).days if later else None
    return out

SHIPS_={}
DAILY_NAME='BOM Material Dashboard - Daily'     # one column per day (user, Nov 2026)
HEAT_NAME='BOM Material Dashboard - Weekly'     # the weekly cover heatmap sheet (was 'Cover Heatmap', then 'BOM Material Dashboard'; renamed by the user, Nov 2026)
def bom_day_cover(r):
    """Daily cover of an ALERT / heatmap row: one value per day = stock available that morning (S + P of that day) / that day's demand,
       'na' = no demand that day (= WSA column V rule; below 1 = stock-zero day)."""
    P=list(r.get('P') or []); S=list(r.get('S') or [])
    return [((S[k]+P[k])/P[k]) if (k<len(P) and P[k]>1e-9) else 'na' for k in range(len(S))]
def bom_risk_key(r):
    """Highest risk first - the SAME rule as the RM Material Dashboards (user, Nov 2026: \"all dashboard must sort from highest to lowest
       risk\"): first day the daily cover goes below 1 (stock-zero day), then below 5, then the lowest cover today; rows with no demand last."""
    vs=bom_day_cover(r); P=list(r.get('P') or [])
    nums=[v for v in vs if not isinstance(v,str)]
    if not nums: return (1,0,0,0,str(r.get('name')))
    neg=next((i for i,v in enumerate(vs) if not isinstance(v,str) and v<1),10**6)     # below 1 = the day's need is not covered = stock-zero day (checker finding: <0 missed 1-day stock-outs)
    lo=next((i for i,v in enumerate(vs) if not isinstance(v,str) and v<5),10**6)
    d_=(float(r.get('stock') or 0.0)/P[0]) if (P and P[0]>1e-9) else 1e9
    return (0,neg,lo,d_,str(r.get('name')))
COUNT_ITEM={}      # PPC Count 'Item' per code (Cover Heatmap groups, MRP Daily label) - filled in build_new_mrp
def count_item_of(code):
    """PPC Count 'Item' of a code; a code not on the Count sheet takes the Item of its alias code (other roll length, e.g. Cell
       Fixation tape RMTTT00000010 <- RMTTT00000018)."""
    c=str(code or '').strip()
    if COUNT_ITEM.get(c): return COUNT_ITEM[c]
    return next((COUNT_ITEM[a] for a,m in (BOM_ALIAS[0] or {}).items() if str(m).strip()==c and COUNT_ITEM.get(str(a).strip())),None)
def is_pm_like(code):
    """Packaging per your rule (user answers Nov 2026: tape = PM, barcode / back label = PM): PPC Count Item 'Packaging' (alias too) or RMLAB."""
    return str(count_item_of(code) or '').lower().startswith('packag') or str(code or '').upper().startswith('RMLAB')
def _short_lbl(b):
    """MRP Daily / Weekly column C: 'PM - Packaging' for a material your PPC Count calls Packaging (Kraft labels, barcodes, back label)
       or for RMLAB codes, when its own label says raw material / consumable (user, Nov 2026); other labels unchanged."""
    sh=b.get('short'); m=str(b.get('mat') or '').strip()
    pk=is_pm_like(m)
    if pk and (str(sh or '').lower().startswith(('cs','raw material')) or not sh): return 'PM - Packaging'
    return sh
def count_items(cnt_hdr, cnt_rows):
    """PPC Count column 'Item' (Item Category) per material code: first non-empty, non-#N/A value."""
    hl=[' '.join(str(h or '').split()).lower() for h in (cnt_hdr or [])]
    im=hl.index('material') if 'material' in hl else None; ii=hl.index('item') if 'item' in hl else None; out={}
    if im is None or ii is None: return out
    for r in cnt_rows or []:
        if not r or im>=len(r) or ii>=len(r) or not r[im]: continue
        v=r[ii]
        if v in (None,'') or (isinstance(v,str) and v.startswith('#')): continue
        out.setdefault(' '.join(str(r[im]).split()),' '.join(str(v).split()))
    return out
def add_cover_heatmap(path, rows, d0, nd, items=None):
    """Cover Heatmap: the SAME rows as the ALERT sheet (Risk Charts list merged into ALERT, Nov 2026), weekly projected cover vs 21 days.
       Sorted from high to low risk (user request Nov 2026): alert class (stock-out inside the lead time > stock-out later > under 21 days >
       under 21 days later), then the number of weeks out of stock, then the weeks under 21 days, then the first stock-zero day."""
    from openpyxl.styles import Border, Side
    wb=post_wb(path)
    for n_ in ('Risk Charts','Cover Heatmap','BOM Material Dashboard',HEAT_NAME):
        if n_ in wb.sheetnames: del wb[n_]
    ws=wb.create_sheet(HEAT_NAME); ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor=DASH_TAB
    F=lambda **k: Font(name='Calibri',**k)
    def fill(c,col): c.fill=PatternFill('solid',fgColor=col)
    TGT=REPORT_DAYS; NW=16; r0=1; DARK=DASH_BG; TXT=DASH_TXT; MUT=DASH_MUT
    cL=2; cK=3; cC=4; cM=5; cS=5+len(HEAT_MS_COLS)      # material name in B, Category in C (= Material Status, user Nov 2026), alert class in D,
                                                            # E..J = Material Status columns (user, Nov 2026; filled by _heat_ms_fill), weeks from K
    ws.column_dimensions['A'].width=15; ws.column_dimensions['B'].width=44;     # A = Code (user, Nov 2026: "code and item name in all dashboard")
    ws.column_dimensions['C'].width=18; ws.column_dimensions['D'].width=15
    if not MIS_CAT and MIS_PATH_[0]: load_mis_cat(MIS_PATH_[0])
    for w in range(NW): ws.column_dimensions[L(cS+w)].width=7
    for j_,(h_,w_) in enumerate(HEAT_MS_COLS): ws.column_dimensions[L(cM+j_)].width=w_
    def weeks(r):                                            # the DAILY number on each week's date (user, Nov 2026: "01-Dec in the BOM dashboard
        v_=bom_day_cover(r)                                  # should be the same as 01-Dec in the RM dashboard") = cover that morning / that day's demand
        return [v_[7*w] if 7*w<len(v_) else 'na' for w in range(NW)] if any(not isinstance(x_,str) for x_ in v_) else []
    exp=[]                                                   # substitute group -> one row per material, own stock / plan / in transit (user, Nov 2026)
    for r in rows:
        if not r.get('mem'): exp.append(r); continue
        for m_ in r['mem']:
            P_=m_['P']; T_=m_['T']; cur=m_['stock']; S_=[]
            for k in range(len(P_)): cur=cur-P_[k]+(T_[k] if k<len(T_) else 0.0); S_.append(cur)
            nz=[v for v in P_[:min(len(P_),avg_prod_days())] if v>0]; need_=(sum(nz)/len(nz)) if nz else None
            e_=alert_eval(m_['name'],m_['code'],m_['uom'],m_['stock'],S_,T_,need_,m_['lt'],sum(T_),d0,why=alert_why(S_,need_,d0,m_['lt'],T=T_)) if need_ else None
            if e_ is None:
                k0_=next((k for k,v in enumerate(S_) if v<0),None)
                e_=dict(cls=4,code=m_['code'],name=m_['name'],need=need_,S=S_,k0=k0_)
            e_['name']=f"{m_['name'][:27]} (own)"; e_['P']=list(P_); e_['stock']=m_['stock']
            k0_=next((k for k,v in enumerate(S_) if v<0),None)          # its OWN Material Status values (user answer Nov 2026: option 2)
            e_['own_ms']=dict(stock=float(m_['stock'] or 0.0),need=need_,short=((d0+pd.Timedelta(days=k0_)).to_pydatetime() if k0_ is not None else None))
            exp.append(e_)
    def mgrp(r):                                             # RM first, then PM, then CS (user, Nov 2026: "group rm material, then pm material then cs material")
        c_=str(r.get('code') or '').upper()                # group = your PPC Count 'Item' (user, Nov 2026: "refer to item category from count ppc"):
        it_=str(count_item_of(r.get('code')) or '').strip().lower()      # Packaging -> PM, Consumable -> CS, any other item -> RM (alias code's Item too)
        if it_: return 1 if it_.startswith('packag') else (2 if it_.startswith('consum') else 0)
        if c_.startswith('RMLAB'): return 1                # not on the Count sheet: by code (barcode / back label = packaging)
        return 0 if c_.startswith('RM') else (1 if c_.startswith('PM') else 2)
    GN=['RAW MATERIALS (RM)','PACKAGING MATERIALS (PM)','CONSUMABLES (CS)']
    srt=sorted(exp,key=lambda r:(mgrp(r),)+bom_risk_key(r))      # one risk rule for every dashboard (user, Nov 2026)
    rows_=[]                                                 # band row (str) before each group, then its materials in risk order
    for g_ in range(3):
        mem_=[r for r in srt if mgrp(r)==g_]
        if mem_: rows_.append(f"{GN[g_]} · {len(mem_)}"); rows_.extend(mem_)
    height=4+len(rows_)+5; HEAT_MS_ROWS.clear()
    for rr in range(r0,r0+height):
        for cc in range(1,cS+NW+1): fill(ws.cell(rr,cc),DARK)
    ws.cell(r0,cL,f'Weekly cover of the ALERT materials vs {TGT} days').font=F(size=16,bold=True,color=TXT)
    ws.cell(r0+1,cL,f"Each column = that date's cover (stock that morning ÷ that day's demand, days) - the same number as the daily dashboards and the RM Material Dashboard on that date: red = short (below 1 day) · orange / amber = below {TGT} days · blue = {TGT}+ days · na = no demand that day. "
                    f'Grouped by your PPC Count Item: raw materials (RM), then packaging (PM), then consumables (CS); in each group highest risk on top (same rule as every dashboard): first day the daily cover goes below 1 (stock-zero day), then below 5 days, then the lowest cover today; no demand last. Same materials as the ALERT sheet; a substitute group (EPE 480 + EPE 420) is shown per material (own stock, plan, in transit).').font=F(size=9,color=MUT)
    c=ws.cell(r0+3,cC,'Alert'); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='center')
    c=ws.cell(r0+3,cK,'Category'); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='left',indent=1)
    c=ws.cell(r0+3,1,'Code'); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='left',indent=1)
    for j_,(h_,w_) in enumerate(HEAT_MS_COLS):
        c=ws.cell(r0+3,cM+j_,h_); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='center',vertical='bottom',wrap_text=True)
    ws.row_dimensions[r0+3].height=30
    c=ws.cell(r0+3,cL,'Item name'); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='right',indent=1)
    for w in range(NW):
        c=ws.cell(r0+3,cS+w,f"{(d0+pd.Timedelta(days=7*w)):%d-%b}"); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='center')
    bins=DAY_BINS()                                          # the daily colour rule (short = below 1 day), as the RM / daily dashboards
    def colr(v):
        for a_,b_,c_ in bins:
            if a_<=v<b_: return c_
        return DAY_NA
    bdk=Side(style='thin',color=DARK)
    for i,r in enumerate(rows_):
        rr=r0+4+i
        if isinstance(r,str):                                # group band
            c=ws.cell(rr,cL,r); c.font=F(size=9,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='left',indent=1)
            for cc in range(1,cS+NW): fill(ws.cell(rr,cc),DASH_BAND)      # band from column A (Code column; checker finding)
            ws.row_dimensions[rr].height=16; continue
        nm=str(r['name'])+(f" ({str(r['where']).replace('Glass G12R Topcon ','')})" if r.get('where') else '')    # full item name (= ALERT; checker finding)
        c=ws.cell(rr,1,r.get('code')); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='left',indent=1)
        HEAT_MS_ROWS.append((rr,str(r.get('code') or '').strip(),r.get('where'),r.get('own_ms')))
        c=ws.cell(rr,cL,nm); c.font=F(size=9,color=TXT); c.alignment=Alignment(horizontal='right',indent=1)
        c=ws.cell(rr,cK,item_category(r.get('code'))); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='left',indent=1)
        _,bg,fg,_=ALERT_CLS[r['cls']] if r['cls']<len(ALERT_CLS) else ('','1E8449','FFFFFF','')
        c=ws.cell(rr,cC,ALERT_SHORT[r['cls']] if r['cls']<len(ALERT_SHORT) else 'Not on ALERT'); fill(c,bg); c.font=F(size=8,bold=True,color=fg); c.alignment=Alignment(horizontal='center')
        c.border=Border(left=bdk,right=bdk,top=bdk,bottom=bdk)
        for w,v in enumerate(weeks(r)):
            if isinstance(v,str): cl=ws.cell(rr,cS+w,'na'); fill(cl,DAY_NA); cl.font=F(size=7,color=DAY_NA_TXT)
            else:
                cl=ws.cell(rr,cS+w,(math.floor(v*10)/10 if 0<=v<1 else round(v))); fill(cl,colr(v)); cl.font=F(size=8,color='FFFFFF')
                if 0<=v<1: cl.number_format='0.0'
            cl.alignment=Alignment(horizontal='center')
            cl.border=Border(left=bdk,right=bdk,top=bdk,bottom=bdk)
        ws.row_dimensions[rr].height=15
    rl=r0+4+len(rows_)+1
    ws.cell(rl,cS,"Cover that morning, days (stock ÷ that day's demand):").font=F(size=9,color=TXT)
    labs=DAY_LABS()
    for k,c_ in enumerate([b_[2] for b_ in bins]+[DAY_NA]):
        c=ws.cell(rl+1,cS+2*k,labs[k]); fill(c,c_); c.font=F(size=8,bold=True,color=(DASH_NATXT if c_==DAY_NA else 'FFFFFF')); c.alignment=Alignment(horizontal='center')
    ws.cell(rl+3,cL,"Source: the ALERT projection (stock rolled forward to today + MIS in transit by arrival date − plan demand; back glass 333 per WSA sheet). One column per week = the daily number on that date (BOM Material Dashboard - Daily).").font=F(size=8,color=MUT)
    for rr in range(r0,r0+height): ws.row_dimensions[rr].height=ws.row_dimensions[rr].height or 16
    pos=wb.sheetnames.index('ALERT')+1 if 'ALERT' in wb.sheetnames else 0
    wb.move_sheet(HEAT_NAME,offset=pos-wb.sheetnames.index(HEAT_NAME))
    HEAT_ROWS[0]=(rows_,d0)                                  # the daily version is built after every other sheet is placed (tab order)
HEAT_ROWS=[None]
# + Stock, SAP stock (MB52), Physical count = Material Status 'STORE · stock' (user, Nov 2026: "BOM Material Dashboard, add these")
HEAT_MS_COLS=[('Stock',11),('SAP stock (MB52)',11),('Physical count',11),('In transit',11),('Next arrival',9),('Need / day',9),('Short from',10),('Days of supply',8),('Days incl. in transit',9)]
HEAT_MS_ROWS=[]     # (row, code, WSA sheet) of the weekly dashboard - filled from the Material Status sheet by _heat_ms_fill
def _heat_ms_fill(wb):
    """BOM Material Dashboard - Weekly columns E..J = the Material Status values of the row's material (user, Nov 2026: "Add these from
       material status sheet into bom material dashboard: In transit, Next arrival, Need / day, Short from, Days of supply, Days incl. in
       transit"). Copied cell by cell from the finished Material Status sheet (same value, same number format), matched by code; back glass
       333 by its WSA sheet ('(WSA sheet: ...)' in the Material Status name); substitute members (EPE 480 / 420 own) show the Material Status
       row of the code (pooled group values, as on Material Status). Returns the number of rows filled."""
    if HEAT_NAME not in wb.sheetnames or 'Material Status' not in wb.sheetnames or not HEAT_MS_ROWS: return 0
    ms=wb['Material Status']; ws=wb[HEAT_NAME]
    hd={str(ms.cell(2,j).value or '').strip():j for j in range(1,ms.max_column+1)}
    jc=hd.get('Item code'); jn=hd.get('Item name')
    if not jc: return 0
    src={}
    for r_ in range(3,ms.max_row+1):
        c_=str(ms.cell(r_,jc).value or '').strip()
        if not re.match(r'^(RM|PM|CS)[A-Z]*\d',c_): continue
        m_=re.search(r'\(WSA sheet: ([^)]+)\)',str(ms.cell(r_,jn).value or '') if jn else '')
        src.setdefault((c_,m_.group(1).strip() if m_ else None),r_); src.setdefault((c_,None),r_)
    F=lambda **k: Font(name='Calibri',**k); TXT=DASH_TXT; n_=0
    own_rows=set()
    for rr,code,where,om in HEAT_MS_ROWS:
        r_=src.get((code,str(where).strip() if where else None)) or src.get((code,None))
        if r_ is None: continue
        n_+=1
        for j_,(h_,w_) in enumerate(HEAT_MS_COLS):
            sj=hd.get(h_)
            if not sj: continue
            sc=ms.cell(r_,sj); c=ws.cell(rr,5+j_,sc.value); c.number_format=sc.number_format
            isd=isinstance(sc.value,(dt.datetime,dt.date))
            c.font=F(size=8,color=(DASH_RED if (h_=='Short from' and isd) else TXT),bold=(h_=='Short from' and isd))
            c.alignment=Alignment(horizontal='center',vertical='center')
        if om:                                                        # substitute member '(own)' row (user answer Nov 2026: option 2): Need / day, Short from,
            own_rows.add(str(ws.cell(rr,2).value))                    # Days of supply, Days incl. in transit of the material ALONE (= its week cells);
            jit_=5+[h_ for h_,w_ in HEAT_MS_COLS].index('In transit'); it_=ws.cell(rr,jit_).value if isinstance(ws.cell(rr,jit_).value,(int,float)) else 0.0      # In transit / Next arrival are per code on Material Status already
            nd_=om.get('need'); st_=om.get('stock') or 0.0
            vals={'Need / day':(round(nd_,1) if nd_ else '-'),'Short from':(om.get('short') or 'no shortage'),
                  'Days of supply':(round(max(0.0,st_)/nd_) if nd_ else '-'),'Days incl. in transit':(round(max(0.0,st_+it_)/nd_) if nd_ else '-')}
            for j_,(h_,w_) in enumerate(HEAT_MS_COLS):
                if h_ in vals:
                    c=ws.cell(rr,5+j_,vals[h_]); isd=isinstance(vals[h_],(dt.datetime,dt.date))
                    c.font=F(size=8,color=(DASH_RED if (h_=='Short from' and isd) else TXT),bold=(h_=='Short from' and isd))
    nt=next((r_ for r_ in range(1,4) if isinstance(ws.cell(r_,2).value,str) and ws.cell(r_,2).value.startswith('Each column')),None)
    if nt: ws.cell(nt,2).value=ws.cell(nt,2).value+" Columns Stock .. Days incl. in transit = the Material Status sheet"+(f"; rows {' / '.join(sorted(own_rows))}: Need / day, Short from and the days are the material ALONE (Material Status shows the substitute group together)" if own_rows else '')+"."
    return n_
DAY_NA='474D57'      # 'na' cell (no demand that day); was 1B2335 on the black page
# Dashboard background (user, Nov 2026: "Change all Dashboard background from black to better color"): light theme - soft off-white
# page, navy text, white-ish gaps between the heat cells; the cover colours (DAY_BINS) and the navy group bands stay.
# BLACK AGAIN (user, Nov 2026: "change background back to black as before") - the values before the light theme.
# SOFTER DARK (user, Nov 2026: "Background BOM Material Dashboard too dark"): charcoal 33383F instead of 12131C (light = rejected).
DAY_NA_TXT='C8CED6'; DASH_BAND='2B5A84'   # 'na' text 5.4:1 on DAY_NA; group band steel blue (navy 2C3E50 vanished on charcoal - checker)
DASH_BG='33383F'; DASH_TXT='E5E8EE'; DASH_MUT='AEB6BF'; DASH_TAB='33383F'; DASH_RED='FF8A80'; DASH_NATXT='FFFFFF'
def DAY_BINS():
    """Colour rule of the daily dashboards (legend row of BOM Material Dashboard - Daily): cover in days that morning; short = below 1 day."""
    T=REPORT_DAYS
    return [(-1e9,1,'D40000'),(1,7,'E8590C'),(7,14,'F08C00'),(14,T,'C9A227'),(T,35,'1A5A8E'),(35,60,'1E7BC4'),(60,1e9,'0EA5F5')]
def DAY_LABS():
    T=REPORT_DAYS; return ['short (< 1)','1–7','7–14',f'14–{T}',f'{T}–35','35–60','60+','na']
def _dark_rm_dashboard(ws, title, note, source):
    """RM Material Dashboard (- MIS) in the BOM Material Dashboard design (user, Nov 2026: "apply color rule mentioned in row 55 from
       bom material dashboard to 2 sheets rm material dashboard. I want rm dashboard design style same as bom dashboard"): the sheet built
       by _wsa_dashboard_to (header row 1, data from row 2, risk order) is rewritten as a dark heatmap - title + note, header row 4, data
       from row 5 (same rows, same order, same values), static fills by DAY_BINS, legend, source line. Columns stay the WSA Dashboard's:
       A Code(s) + B Item name(s) of the WSA sheet's material (user, Nov 2026: "code and item name in all dashboard"), C planning sheet,
       D Category, E Today's Morning stock (before today's arrivals), F Short from (first day below 1 = the sort key, user Nov 2026),
       G.. one per day (the WSA lookup column index is left out). The 4 WSA colour rules are dropped."""
    from openpyxl.styles import Border, Side
    from openpyxl.formatting.formatting import ConditionalFormattingList
    mc=ws.max_column
    hdr=[ws.cell(1,j).value for j in range(1,mc+1)]
    rows_=[r for r in range(2,ws.max_row+1) if isinstance(ws.cell(r,1).value,(int,float))]
    data=[[ws.cell(r,j).value for j in range(1,mc+1)] for r in rows_]
    cmap=DASH_CODES.get(ws.title,{})                                # code(s) of each row's WSA sheet -> Code + Item name (user, Nov 2026)
    nm_of={}
    for x_ in (AL_ORDER_ or []):
        b_=x_['b']; nm_of.setdefault(str(b_['mat']).strip(),str(b_.get('common') or '').strip())
    codes=[cmap.get(r,[]) for r in rows_]
    while hdr and hdr[-1] is None: hdr.pop()
    nc=len(hdr)
    for m_ in list(ws.merged_cells.ranges): ws.unmerge_cells(str(m_))
    ws._cells={}; ws.conditional_formatting=ConditionalFormattingList()
    for r_ in list(ws.row_dimensions.keys()): ws.row_dimensions[r_].height=None
    ws.sheet_view.showGridLines=False; ws.sheet_view.zoomScale=None; ws.sheet_properties.tabColor=DASH_TAB
    F=lambda **k: Font(name='Calibri',**k)
    def fill(c,col): c.fill=PatternFill('solid',fgColor=col)
    DARK=DASH_BG; TXT=DASH_TXT; MUT=DASH_MUT; sS=6; cS=7; cF=6; r0=1     # source days from F; output: F 'Short from', days from G (user, Nov 2026)
    oc=lambda j: j+(cS-sS)                                          # source day column -> output column
    nco=oc(nc)
    bins=DAY_BINS(); bdk=Side(style='thin',color=DARK); BD=Border(left=bdk,right=bdk,top=bdk,bottom=bdk)
    def colr(v):
        for a_,b_,c_ in bins:
            if a_<=v<b_: return c_
        return DAY_NA
    for j,w_ in ((1,15),(2,30),(3,26),(4,16),(5,8),(cF,11)): ws.column_dimensions[L(j)].width=w_
    for j in range(cS,nco+1): ws.column_dimensions[L(j)].width=4.6
    height=4+len(data)+5
    for rr in range(r0,r0+height):
        for cc in range(1,max(nco,cS)+1): fill(ws.cell(rr,cc),DARK)
    ws.cell(r0,2,title).font=F(size=16,bold=True,color=TXT)
    ws.cell(r0+1,2,note).font=F(size=9,color=MUT)
    for j,h in ((1,'Code'),(2,'Item name'),(3,'Planning sheet'),(4,'Category'),(5,"Today's Morning stock (before today's arrivals)"),(cF,'Short from')):
        c=ws.cell(r0+3,j,h); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal=('right' if j==2 else ('left' if j in (1,3,4) else 'center')),vertical='bottom',wrap_text=True,indent=(1 if j<=4 else 0))
    for j in range(sS,nc+1):
        v=hdr[j-1]
        c=ws.cell(r0+3,oc(j),v); c.number_format='dd-mmm'; c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='center',vertical='bottom',text_rotation=90)
    ws.row_dimensions[r0+3].height=42
    def val(c,v):
        if isinstance(v,bool) or not isinstance(v,(int,float)):
            fill(c,DAY_NA); c.value=('na' if (isinstance(v,str) and v.strip().lower()=='na') else v); c.font=F(size=7,color=DAY_NA_TXT)
        else:
            fill(c,colr(v)); c.value=(math.floor(v*10)/10 if 0<=v<1 else round(v)); c.font=F(size=8,color='FFFFFF')   # below 1 rounded DOWN: 0.98 must not show 1.0 on a red cell (checker finding)
            if 0<=v<1: c.number_format='0.0'
        c.alignment=Alignment(horizontal='center'); c.border=BD
    for i,row in enumerate(data):
        rr=r0+4+i
        cd_=codes[i]; nms=list(dict.fromkeys(nm_of.get(c_) or c_ for c_ in cd_))
        c=ws.cell(rr,1,' / '.join(cd_) or '-'); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='left',indent=1,wrap_text=True)
        c=ws.cell(rr,2,' / '.join(nms) or '-'); c.font=F(size=9,color=TXT); c.alignment=Alignment(horizontal='right',indent=1,wrap_text=True)
        c=ws.cell(rr,3,row[1]); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='left',indent=1)
        c=ws.cell(rr,4,row[2]); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='left',indent=1)
        nop=isinstance(row[3],str) and row[3].strip().lower()=='no plan'
        if nop:
            c=ws.cell(rr,5,row[3]); fill(c,'5D6D7E'); c.font=F(size=8,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='center'); c.border=BD
        else: val(ws.cell(rr,5),row[3])
        # Short from = first day the cover goes below 1 = the sort key (user, Nov 2026: make the risk order visible, as 'Short from' on the BOM dashboard)
        sf_=next((hdr[j-1] for j in range(sS,nc+1) if j-1<len(row) and isinstance(row[j-1],(int,float)) and not isinstance(row[j-1],bool) and row[j-1]<1 and isinstance(hdr[j-1],(dt.datetime,dt.date))),None)
        lastd=next((hdr[j-1] for j in range(nc,sS-1,-1) if isinstance(hdr[j-1],(dt.datetime,dt.date))),None)
        c=ws.cell(rr,cF)
        if nop: c.value='-'; c.font=F(size=8,color=MUT)
        elif sf_ is not None: c.value=sf_; c.number_format='dd-mmm'; c.font=F(size=9,bold=True,color=DASH_RED)
        else: c.value=(f"none to {pd.Timestamp(lastd):%d-%b}" if lastd is not None else 'no shortage'); c.font=F(size=8,color=MUT)
        c.alignment=Alignment(horizontal='center')
        for j in range(sS,nc+1):
            v=row[j-1] if j-1<len(row) else None
            if nop and v is None: continue                        # no plan: dark row, no cells
            val(ws.cell(rr,oc(j)),v)
        ws.row_dimensions[rr].height=(26 if len(cd_)>1 else 15)
    rl=r0+4+len(data)+1
    ws.cell(rl,cS,"Cover that morning, days:").font=F(size=9,color=TXT)
    labs=DAY_LABS()
    for k,c_ in enumerate([b_[2] for b_ in bins]+[DAY_NA]):
        ws.merge_cells(start_row=rl+1,start_column=cS+3*k,end_row=rl+1,end_column=cS+3*k+1)
        c=ws.cell(rl+1,cS+3*k,labs[k]); fill(c,c_); fill(ws.cell(rl+1,cS+3*k+1),c_); c.font=F(size=8,bold=True,color=(DASH_NATXT if c_==DAY_NA else 'FFFFFF')); c.alignment=Alignment(horizontal='center')
    ws.cell(rl+3,2,source).font=F(size=8,color=MUT)
    for rr in range(r0,r0+height): ws.row_dimensions[rr].height=ws.row_dimensions[rr].height or 16
    ws.freeze_panes=ws.cell(r0+4,cS).coordinate
    return len(data)
def _bom_daily_sheet(wb, rows_, d0):
    """'BOM Material Dashboard - Daily' (user, Nov 2026): the SAME rows / groups / order as the BOM Material Dashboard - Weekly, in the
       SAME format and design (user, Nov 2026: "format and design same as weekly": dark heatmap, title + note, Category + Alert columns,
       coloured cells, legend, source line). One column per day = stock available that morning (stock + MIS arrivals up to that day -
       plan of the days before) / THAT day's demand ('na' = no demand that day) - the rule of WSA column V, so the first value below 1 =
       the ALERT's first stock-zero day. E = Today's Morning stock = stock today / today's demand. Colours = the weekly bins (short = below 1
       day here: the day's need is not covered)."""
    from openpyxl.styles import Border, Side
    if DAILY_NAME in wb.sheetnames: del wb[DAILY_NAME]
    ws=wb.create_sheet(DAILY_NAME); ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor=DASH_TAB
    F=lambda **k: Font(name='Calibri',**k)
    def fill(c,col): c.fill=PatternFill('solid',fgColor=col)
    TGT=REPORT_DAYS; DARK=DASH_BG; TXT=DASH_TXT; MUT=DASH_MUT; r0=1
    cA=1; cL=2; cK=3; cC=4; cT=5; cS=6                     # code, material, Category, Alert, Today, days from F
    mats=[r for r in rows_ if not isinstance(r,str)]
    nd_=max([len(r.get('S') or []) for r in mats]+[1])
    for j,w_ in ((cA,15),(cL,44),(cK,18),(cC,15),(cT,8)): ws.column_dimensions[L(j)].width=w_
    for k in range(nd_): ws.column_dimensions[L(cS+k)].width=4.6
    height=4+len(rows_)+5
    for rr in range(r0,r0+height):
        for cc in range(1,cS+nd_): fill(ws.cell(rr,cc),DARK)
    ws.cell(r0,cL,f'Daily cover of the ALERT materials vs {TGT} days').font=F(size=16,bold=True,color=TXT)
    ws.cell(r0+1,cL,"Each day = stock available that morning ÷ that day's demand (days): red = short (the day's need is not covered) · orange / amber = below "
                    f"{TGT} days · blue = {TGT}+ days · na = no demand that day. Same rows, groups and order as the BOM Material Dashboard - Weekly: "
                    "RM, then PM, then CS; in each group highest risk on top (first day the cover goes below 1 = stock-zero day, then below 5 days, then the lowest cover today).").font=F(size=9,color=MUT)
    for j,h in ((cA,'Code'),(cL,'Item name'),(cK,'Category'),(cC,'Alert'),(cT,"Today's Morning stock")):
        c=ws.cell(r0+3,j,h); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal=('left' if j in (cA,cK) else 'center'),vertical='bottom',wrap_text=True,indent=(1 if j in (cA,cK) else 0))
    for k in range(nd_):
        c=ws.cell(r0+3,cS+k,(d0+pd.Timedelta(days=k)).to_pydatetime()); c.number_format='dd-mmm'; c.font=F(size=8,color=MUT)
        c.alignment=Alignment(horizontal='center',vertical='bottom',text_rotation=90)
    ws.row_dimensions[r0+3].height=42
    bins=DAY_BINS(); NA=DAY_NA
    def colr(v):
        for a_,b_,c_ in bins:
            if a_<=v<b_: return c_
        return NA
    bdk=Side(style='thin',color=DARK); BD=Border(left=bdk,right=bdk,top=bdk,bottom=bdk)
    def val(c,v):
        if isinstance(v,str): fill(c,NA); c.value='na'; c.font=F(size=7,color=DAY_NA_TXT)
        else:
            fill(c,colr(v)); c.value=(math.floor(v*10)/10 if 0<=v<1 else round(v)); c.font=F(size=8,color='FFFFFF')   # below 1 rounded DOWN: 0.98 must not show 1.0 on a red cell (checker finding)
            if 0<=v<1: c.number_format='0.0'                  # 0.5-0.99 would show '1' on a red (short) cell (checker finding)
        c.alignment=Alignment(horizontal='center'); c.border=BD
    for i,r in enumerate(rows_):
        rr=r0+4+i
        if isinstance(r,str):                                   # group band (RM / PM / CS)
            c=ws.cell(rr,cL,r); c.font=F(size=9,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='left',indent=1)
            for cc in range(cA,cS+nd_): fill(ws.cell(rr,cc),DASH_BAND)
            ws.row_dimensions[rr].height=16; continue
        P=list(r.get('P') or [])
        nm=str(r['name'])+(f" ({str(r['where']).replace('Glass G12R Topcon ','')})" if r.get('where') else '')
        c=ws.cell(rr,cA,r.get('code')); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='left',indent=1)
        c=ws.cell(rr,cL,nm); c.font=F(size=9,color=TXT); c.alignment=Alignment(horizontal='right',indent=1)
        c=ws.cell(rr,cK,item_category(r.get('code'))); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='left',indent=1)
        _,bg,fg,_=ALERT_CLS[r['cls']] if r['cls']<len(ALERT_CLS) else ('','1E8449','FFFFFF','')
        c=ws.cell(rr,cC,ALERT_SHORT[r['cls']] if r['cls']<len(ALERT_SHORT) else 'Not on ALERT'); fill(c,bg); c.font=F(size=8,bold=True,color=fg); c.alignment=Alignment(horizontal='center'); c.border=BD
        p0=P[0] if P else 0.0
        val(ws.cell(rr,cT),(float(r.get('stock') or 0.0)/p0) if p0>1e-9 else 'na')
        vs=bom_day_cover(r)
        for k in range(nd_): val(ws.cell(rr,cS+k),vs[k] if k<len(vs) else 'na')
        ws.row_dimensions[rr].height=15
    rl=r0+4+len(rows_)+1
    ws.cell(rl,cS,"Cover that morning, days (stock ÷ that day's demand):").font=F(size=9,color=TXT)
    labs=DAY_LABS()
    for k,c_ in enumerate([b_[2] for b_ in bins]+[NA]):
        ws.merge_cells(start_row=rl+1,start_column=cS+3*k,end_row=rl+1,end_column=cS+3*k+1)
        c=ws.cell(rl+1,cS+3*k,labs[k]); fill(c,c_); fill(ws.cell(rl+1,cS+3*k+1),c_); c.font=F(size=8,bold=True,color=(DASH_NATXT if c_==DAY_NA else 'FFFFFF')); c.alignment=Alignment(horizontal='center')
    ws.cell(rl+3,cL,"Source: the ALERT projection (stock rolled forward to today + MIS in transit by arrival date − plan demand; back glass 333 per WSA sheet). "
                    "One column per day = stock available that morning ÷ that day's demand; first value below 1 = the ALERT's first stock-zero day.").font=F(size=8,color=MUT)
    for rr in range(r0,r0+height): ws.row_dimensions[rr].height=ws.row_dimensions[rr].height or 16
    ws.freeze_panes=ws.cell(r0+4,cS).coordinate
    wb.move_sheet(DAILY_NAME,offset=wb.sheetnames.index(HEAT_NAME)+1-wb.sheetnames.index(DAILY_NAME))


# ===================== RM ISSUES log (one issue per raw material at risk; your entries are carried over each run) =====================
RMI_PREV={}
def load_rmi_prev(paths):
    """Your RM Issues entries from the newest existing MRP_Daily (same-day re-runs included) - read before it is overwritten."""
    RMI_PREV.clear()
    for p_ in [q for q in paths if q and os.path.exists(q)]:
        try:
            pw=load_workbook(p_,read_only=True,data_only=True)
            if 'RM Issues' not in pw.sheetnames: continue
            rows_=list(pw['RM Issues'].iter_rows(values_only=True))
            hi=next((i for i,r in enumerate(rows_[:8]) if r and 'Issue ID' in r),None)
            if hi is None: continue
            h_=list(rows_[hi])
            for r in rows_[hi+1:]:
                if r and r[0]: RMI_PREV[str(r[0])]=dict(zip(h_,r))
            if RMI_PREV: return
        except Exception as _e_: _swallow(4668,_e_); continue
RMI_COLS=['Issue ID','Date Raised','Project','Material','Supplier','PO Reference','Qty Required','Qty Received','Shortfall','Short from (stock = 0)','Required Date (at Factory)',
          'Revised ETA','Shipped qty','Pending qty','Delay (Days)','Production Impact (MW)','Priority','Owner','Status','Date Closed','Aging (Days)']      # no Action / Mitigation (user, Nov 2026)
def add_rm_issues_sheet(path, A, AL_ORDER, d0, nd, ships, prev_path=None):
    """RM Issues: an issue for every raw material that runs short or falls below 21 days in the next 16 weeks.
       Required date = the day cover drops below 21 days (stock 0 if earlier); Qty required = what is missing to keep 21 days at the worst point;
       Qty received = MIS PPC GRN since the issue was raised; Revised ETA / supplier / PO = the next MIS shipment after the required date;
       Production impact (MW) = modules that cannot be built at the worst point x module Wp. Owner, Status, Date Closed and your own
       Action text are kept from the previous MRP_Daily; an issue whose material is no longer at risk is closed automatically."""
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.worksheet.datavalidation import DataValidation
    wb=post_wb(path)
    if 'RM Issues' in wb.sheetnames: del wb['RM Issues']
    prev=dict(RMI_PREV)                                        # your entries, read at the START of the run (before any file is overwritten)
    HZ=min(nd,112); TGT=REPORT_DAYS; dH=d0+pd.Timedelta(days=HZ)
    grn={}                                                     # MIS PPC GRN by code and arrival date (Qty received since the issue was raised)
    try:
        for r in list(wb_ro(MIS_PATH_[0])['In Transit'].iter_rows(values_only=True))[1:]:
            if len(r)>39 and r[39] and isinstance(r[33],(int,float)) and r[33]>0 and isinstance(r[35],(dt.datetime,dt.date)):
                grn.setdefault(str(r[39]).strip(),[]).append((pd.Timestamp(r[35]).normalize(),float(r[33])))
    except Exception as _e_: _swallow(4689,_e_); pass
    # Production Plan inside the horizon, per (line, customer): modules, Wp x modules, MW  -> real Wp and an MW cap
    ppl={}
    try:
        for r in pp_rows_std(wb_ro(USER_MRP or SETTINGS_SRC[0])['Production Plan']):
            dd,ln,wp,cu,q,mw_,_,_=r
            if not isinstance(dd,(dt.datetime,dt.date)) or not isinstance(q,(int,float)) or q<=0: continue
            if not (d0<=pd.Timestamp(dd).normalize()<dH): continue
            cu_=' '.join(str(cu or '').split()); cu_=(LINKS.get(cu_) or {}).get('mis') or cu_   # plan name -> MIS customer (the names cust_line uses)
            e=ppl.setdefault((_line_key(str(ln or '')),_norm(cu_)),[0.0,0.0,0.0])
            e[0]+=q; e[1]+=q*(wp if isinstance(wp,(int,float)) and wp>0 else 0.0)
            e[2]+=(mw_ if isinstance(mw_,(int,float)) and mw_>0 else (q*wp/1e6 if isinstance(wp,(int,float)) and wp>0 else 0.0))
    except Exception as _e_: _swallow(4701,_e_); pass
    out=[]; seen=set()
    base=[x for x in AL_ORDER if not x.get('sup_row') and is_rm(x['b']['mat']) and not is_pm_like(x['b']['mat']) and not x.get('no_demand') and not grp_follower(x['b']['mat'])]   # tapes / barcode / back label = PM (user, Nov 2026)
    byc={x_['b']['mat']:x_ for x_ in A}
    for x in base:
        b=x['b']; m=b['mat']; S=x.get('S') or []
        g=POOLG.get(m); P=(g['P'] if g else b.get('plan')) or []                 # substitute group: pooled demand, like the pooled stock
        mem=(g['members'] if g else [m]); fx=(g.get('fx') if g else {m:1.0}) or {m:1.0}
        if not S or not any(v>0 for v in P[:HZ]): continue
        cw=cover_watch(x,d0,ships,nd)
        t=cw['t'] if cw else None
        ri=risk_info(x,d0)
        if t is None and ri['k'] is None: continue
        t=min([v for v in (t,ri['k']) if v is not None])
        req_d=d0+pd.Timedelta(days=t)
        gap=max([max(0.0,sum(P[k+1:k+1+TGT])-S[k]) for k in range(t,min(HZ,len(S)))] or [0.0])   # missing to keep 21 days of (pooled) demand at the worst point
        sh=sorted([s_ for s_ in grp_ships(m,ships) if s_.get('raw') is not None],key=lambda s_:s_['raw'])
        nx=next((s_ for s_ in sh if s_['raw']>=req_d),None) or (sh[-1] if sh else None)
        iid=f"RMI-{m}"; pv=prev.get(iid,{})
        raised=pd.Timestamp(pv['Date Raised']).normalize() if isinstance(pv.get('Date Raised'),(dt.datetime,dt.date)) else d0
        if str(pv.get('Status') or '').strip().lower()=='closed' and isinstance(pv.get('Date Closed'),(dt.datetime,dt.date)): raised=d0   # re-opened: new date
        rec=sum(q*fx.get(c,1.0)/fx.get(m,1.0) for c in mem for dd,q in grn.get(c,[]) if raised<=dd<=d0)
        # production impact: modules that cannot be built at the worst point x Wp of the plan rows that use it, capped at the MW planned for them
        pairs=set()
        for c in mem:
            for (cu,ln),arr in ((byc.get(c) or {}).get('b',{}).get('cust_line') or {}).items():
                if any((arr[k] if k<len(arr) else 0) for k in range(HZ)): pairs.add((_line_key(str(ln or '')),_norm(cu)))
        rows_pp=[v for k_,v in ppl.items() if k_ in pairs] or [v for k_,v in ppl.items() if k_[0] in {l_ for l_,_ in pairs}]
        mods=sum(v[0] for v in rows_pp); wpq=sum(v[1] for v in rows_pp); mw_plan=sum(v[2] for v in rows_pp)
        wp_avg=(wpq/mods) if mods and wpq else 600.0                   # 600 Wp only when the plan has no Wp
        per_mod=(sum(P[:HZ])/mods) if mods else None
        worst=max(0.0,-min(S[:HZ])) if S else 0.0
        mw=(min(worst/per_mod*wp_avg/1e6,mw_plan) if (per_mod and mw_plan) else (worst/per_mod*wp_avg/1e6 if per_mod else None))
        slt_=short_lt_text(x,d0)
        prio='High' if (ri['cls']=='ALERT' or slt_[1]) else ('Medium' if ri['k'] is not None else 'Low')     # stock zero inside the lead time = High
        late=((nx['raw']-req_d).days if nx else None)
        st=pv.get('Status') if str(pv.get('Status') or '').strip() and str(pv.get('Status')).strip().lower()!='closed' else 'Open'
        proj=str(x.get('proj') or '').replace('\n',', ') or ', '.join(sorted({str(cu) for (cu,ln) in (b.get('cust_line') or {})}))[:60]
        same=[s_ for s_ in sh if nx and s_['raw']==nx['raw'] and s_.get('sup')==nx.get('sup') and s_.get('po')==nx.get('po')]
        out.append([iid,raised.to_pydatetime(),proj,f"{grp_label(m,str(b.get('common') or ''))} ({' + '.join(mem)})",(nx['sup'] if nx else 'no supplier on the way'),
                    (str(nx.get('po')) if nx and nx.get('po') else None),round(gap+rec),round(rec),round(gap),slt_[0],req_d.to_pydatetime(),(nx['raw'].to_pydatetime() if nx else None),
                    (round(sum(s_.get('shq') or 0 for s_ in same)) if nx else None),(round(sum(s_['qty'] for s_ in same)) if nx else None),
                    late,(round(mw,2) if mw is not None else None),prio,pv.get('Owner'),st,None,(d0-raised).days])
        seen.add(iid)
    for iid,pv in prev.items():                                # no longer at risk: closed automatically (kept for the record)
        if iid in seen: continue
        dr=pv.get('Date Raised'); dc=pv.get('Date Closed') if isinstance(pv.get('Date Closed'),(dt.datetime,dt.date)) else d0.to_pydatetime()
        out.append([iid,dr,pv.get('Project'),pv.get('Material'),pv.get('Supplier'),pv.get('PO Reference'),pv.get('Qty Required'),pv.get('Qty Received'),0,pv.get('Short from (stock = 0)'),
                    pv.get('Required Date (at Factory)'),pv.get('Revised ETA'),pv.get('Shipped qty'),pv.get('Pending qty'),pv.get('Delay (Days)'),pv.get('Production Impact (MW)'),pv.get('Priority'),pv.get('Owner'),
                    'Closed',dc,((pd.Timestamp(dc)-pd.Timestamp(dr)).days if isinstance(dr,(dt.datetime,dt.date)) else None)])
    # open issues by Production Impact (MW), highest first; issues without MW after them; closed issues at the bottom
    out.sort(key=lambda r_:(r_[18]=='Closed',r_[15] is None,-(r_[15] if isinstance(r_[15],(int,float)) else 0)))
    # ---- sheet
    ws=wb.create_sheet('RM Issues'); ws.sheet_view.showGridLines=False
    F=lambda **k: Font(name='Calibri',**k)
    ws['A1']='RM Issues — raw materials at risk of shortage'; ws['A1'].font=F(size=20,bold=True,color='14213D')
    ws['A2']=(f"One issue per raw material that runs short or falls below {TGT} days in the next 16 weeks, sorted by Production Impact (MW) high to low (closed issues at the bottom), rebuilt every run. "
              "Required date = first day under 21 days of cover (or stock 0); Qty required = what is missing to keep 21 days at the worst point; Qty received = MIS PPC GRN since the issue was raised; "
              "Revised ETA / supplier / PO = next MIS shipment, Shipped qty / Pending qty = MIS Final Qty. / Final Pending Qty. of that shipment (same supplier, PO, ETA); Production impact = modules that cannot be built at the worst point x the Wp of the plan rows using it (capped at the MW planned for them in 16 weeks); a substitute group = ONE issue (group demand, stock and shipments). "
              "Yellow columns are yours (Owner, Status, Date Closed) - kept from the previous MRP_Daily. Issues no longer at risk are closed automatically.")
    ws['A2'].font=F(size=9,color='8A949E'); ws.merge_cells('A2:U2'); ws['A2'].alignment=Alignment(wrap_text=True,vertical='top'); ws.row_dimensions[2].height=48
    H=list(RMI_COLS)
    for j,h in enumerate(H,1):
        c=ws.cell(4,j,h); c.font=F(size=10,bold=True,color='1F2D3D'); c.fill=PatternFill('solid',fgColor='F4F6F8')
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=Border(bottom=Side(style='medium',color='C9D2DC'))
    ws.row_dimensions[4].height=32
    fm=[None,'dd-mmm-yy',None,None,None,None,'#,##0','#,##0','#,##0',None,'dd-mmm-yy','dd-mmm-yy','#,##0','#,##0','0','0.00',None,None,None,'dd-mmm-yy','0']
    PC={'High':('FCE8E6','C0392B'),'Medium':('FEF3E2','B9770E'),'Low':('FEF9E7','9A7D0A')}
    thin=Side(style='thin',color='E3E8EE')
    for i,row in enumerate(out,5):
        for j,v in enumerate(row,1):
            c=ws.cell(i,j,v); c.font=F(size=10,color='1F2D3D'); c.border=Border(bottom=thin)
            c.alignment=Alignment(vertical='center',wrap_text=(j in (3,10)),horizontal=('left' if isinstance(v,str) else 'right'))
            if fm[j-1]: c.number_format=fm[j-1]
            if j in (18,19,20): c.fill=PatternFill('solid',fgColor='FFFDE7')
        bg,fg=PC.get(row[16],('F2F4F6','5D6D7E')); ws.cell(i,17).fill=PatternFill('solid',fgColor=bg); ws.cell(i,17).font=F(size=10,bold=True,color=fg)
        if isinstance(row[14],(int,float)) and row[14]>0: ws.cell(i,15).font=F(size=10,bold=True,color='C0392B')
        if isinstance(row[9],str) and '→' in row[9]: ws.cell(i,10).font=F(size=10,bold=True,color='C0392B')
        if row[18]=='Closed':
            for j in range(1,22): ws.cell(i,j).font=F(size=10,color='A6ACAF')
        ws.row_dimensions[i].height=30
    for j,w_ in enumerate([22,11,26,30,24,13,11,11,11,30,12,11,11,11,8,11,9,14,10,11,8],1): ws.column_dimensions[L(j)].width=w_
    if out: ws.add_table(Table(displayName='RMIssues',ref=f'A4:U{4+len(out)}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=False)))
    dv=DataValidation(type='list',formula1='"Open,In progress,Expedited,Waiting supplier,Closed"',allow_blank=True); ws.add_data_validation(dv)
    if out: dv.add(f'S5:S{4+len(out)}')
    ws.freeze_panes='E5'
    pos=(wb.sheetnames.index('Requirements')+1) if 'Requirements' in wb.sheetnames else 3
    wb.move_sheet('RM Issues',offset=pos-wb.sheetnames.index('RM Issues'))

def add_stock_check_sheet(path, AL_ORDER, d0, mlog=None):
    """Stock Check: SAP stock (MB52) next to the physical count and its date, and the stock the MRP uses today
       (count rolled forward by usage and receipts when the count is older than today). Highest shortage risk first."""
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    wb=post_wb(path)
    if 'Stock Check' in wb.sheetnames: del wb['Stock Check']
    ws=wb.create_sheet('Stock Check'); ws.sheet_view.showGridLines=False
    F=lambda **k: Font(name='Calibri',**k); INK='1F2D3D'; MUTED='8A949E'; NAVY='14213D'
    rowb=Border(bottom=Side(style='thin',color='E3E8EE')); hb=Border(bottom=Side(style='medium',color='C9D2DC'))
    ws['A1']='Stock Check — SAP stock vs physical count'; ws['A1'].font=F(size=20,bold=True,color=NAVY)
    ws['A2']=('SAP stock = your MB52.xlsx (Unrestricted + Quality Inspection; MIS MB52 only for plant name / final code / categories) · Physical count = PPC Count sheet, count date = the date column holding the count · '
              'when the count is older than today it is rolled forward: Estimated stock = count − usage per day × days since count + received since count (MIS PPC GRN after the count date); '
              'usage per day = average Production Plan demand of the next 7 days (OEM plan left out); not counted: Estimated stock = SAP stock (MB52) · Used in MRP = the stock behind every shortage date in this workbook (= ALERT Estimated stock today) · '
              'a material with stock from 2+ suppliers = one row per supplier (Item name - supplier: MB52 batches / count by supplier; usage and receipts by the count share; Lines / Project / Usage / Module = the whole material).')
    ws['A2'].font=F(size=10,color=MUTED); ws.merge_cells('A2:T2'); ws['A2'].alignment=Alignment(wrap_text=True,vertical='top'); ws.row_dimensions[2].height=58
    H=['Code','Item name','Category','UOM','SAP stock (MB52)','Physical count','Count date','Days since count','Usage / day','Usage since count','Received since count',
       'Estimated stock today','Used in MRP (own stock)','Estimate − SAP','Source','Shortage risk',
       'Lines (share of plan volume)','Project Name (share of plan)','Usage per module by customer (source: your BOM / MIS Standard BOM)',
       'Module (MIS BOM, share of plan)']                   # Q..T moved here from the Action List (user, Nov 2026); Item name + Category = the ALERT (Nov 2026)
    for j,h in enumerate(H,1):
        c=ws.cell(4,j,h); c.font=F(size=10,bold=True,color=INK); c.fill=PatternFill('solid',fgColor='F4F6F8'); c.border=hb
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.row_dimensions[4].height=34
    if not MIS_CAT and MIS_PATH_[0]: load_mis_cat(MIS_PATH_[0])     # Category (item_category = the ALERT rule)
    base=[x for x in AL_ORDER if not x.get('sup_row')]
    supx={(x['b']['mat'],x['sup_row']):x for x in AL_ORDER if x.get('sup_row')}     # Action List supplier rows: customers allowed that vendor
    apos={}                                                        # sorted like its label: the ALERT order first, then the rest (OK / No need / No demand) as before
    for i_,r_ in enumerate(ALERT_ROWS): apos.setdefault(r_['code'],i_)
    base=sorted(base,key=lambda x:apos.get(x['b']['mat'],apos.get(grp_main(x['b']['mat']),10**6)))   # own ALERT row first (EPE 480 / 420 separately), else its group's
    def sup_label(m,nm,ri):                                        # a supplier row of a material on 2+ WSA sheets (333): the ALERT row of the sheet holding its stock
        rs_=[r_ for r_ in ALERT_ROWS if r_['code']==m and r_.get('sups') and nm in r_['sups']]
        if not rs_: return alert_label(m,ri)
        r_=rs_[0]; w_=r_.get('why') or {}; t=ALERT_CLS[r_['cls']][0]
        if r_.get('first') is not None: t+=f" {r_['first']:%d-%b}"
        elif r_['cls']==3 and w_.get('below') is not None: t+=f" from {w_['below']:%d-%b}"
        return t+(f" ({r_['where']})" if r_.get('where') else ''),r_['cls']
    r=5; nsup=0
    for x in base:
        b=x['b']; m=b['mat']; e=STOCK_EST.get(m) or {}; ri=risk_info(x,d0); ss=stock_src(b)
        own=ss['own']; est=ss['est']; sapv=ss['sap']; cnt=ss['cnt']
        nm0=str(b.get('common') or m)                              # = ALERT Item name (MIS Common Name)
        lines=[dict(nm=nm0,sap=sapv,cnt=cnt,est=est,own=own,sh=None,lab=alert_label(m,ri),x=x,src=e.get('src') or 'SAP (MB52)')]
        sp=sup_split(b)
        if sp:                                                     # do not combine suppliers (user, Nov 2026: glass FLAT and KIBING separated)
            ct_=sum(o['cnt'] or 0.0 for o in sp); et_=sum(o['est'] for o in sp)
            lines=[]
            for o in sp:
                sh_=((o['cnt'] or 0.0)/ct_ if (cnt is not None and ct_>0) else ((o['est']/et_) if et_>0 else 0.0))
                lines.append(dict(nm=f"{nm0} - {o['name']}",sap=o['sap'],cnt=o['cnt'],est=o['est'],own=o['est'],sh=sh_,lab=sup_label(m,o['name'],ri),
                                  x=x,src=(e.get('src') or 'SAP (MB52)')+(' · count by supplier' if o['cnt'] is not None else ' · MB52 batches by supplier')
                                      +(f" · customers allowed this supplier: {supx[(m,o['name'])].get('custs')}" if supx.get((m,o['name'])) else '')))
            nsup+=1
        for ln in lines:
            sh_=ln['sh']; k_=(1.0 if sh_ is None else sh_)
            rolled=e.get('est') is not None                         # counted / rolled forward: Estimate − SAP shown (as before)
            vals=[m,ln['nm'],item_category(m),b.get('uom'),
                  (round(ln['sap']) if isinstance(ln['sap'],(int,float)) else None),(round(ln['cnt']) if ln['cnt'] is not None else 'no counting'),
                  (pd.Timestamp(e['date']).to_pydatetime() if e.get('date') is not None else None),e.get('days'),
                  (round(e['use_day']*k_,1) if e.get('use_day') else None),(round(e['use']*k_) if e.get('use') else None),(round(e['rec']*k_) if e.get('count') is not None else None),
                  round(ln['est']),round(ln['own']),
                  (round(ln['est']-ln['sap']) if (rolled and isinstance(ln['sap'],(int,float))) else None),ln['src'],
                  ln['lab'][0],
                  x.get('custs'),x.get('proj'),x.get('usage_src'),x.get('mods')]     # Q..T = the whole material (line / project shares)
            vals=vals[:16]+[(', '.join(map(str,v)) if isinstance(v,(list,tuple)) else v) for v in vals[16:]]
            fm=[None,None,None,None,'#,##0','#,##0','dd-mmm-yy','0','#,##0.0','#,##0','#,##0','#,##0','#,##0','+#,##0;-#,##0;0',None,None,None,None,None,None]
            for j,(v,f_) in enumerate(zip(vals,fm),1):
                c=ws.cell(r,j,v); c.font=F(size=10,color=INK); c.border=rowb
                c.alignment=Alignment(vertical='center',wrap_text=(j>=16),horizontal=('left' if j in (2,3,15,16) or isinstance(v,str) else 'right'))
                if f_: c.number_format=f_
            if isinstance(e.get('days'),int) and e['days']>0:
                for j in (7,8): ws.cell(r,j).fill=PatternFill('solid',fgColor='FEF3E2'); ws.cell(r,j).font=F(size=10,bold=True,color='B9770E')
                ws.cell(r,12).font=F(size=10,bold=True,color=INK)
            ci_=ln['lab'][1]; bg,fg=(ALERT_CHIP[ci_] if ci_ is not None else ('E3F4EC','1E8449'))   # = the ALERT chip colours
            ws.cell(r,16).fill=PatternFill('solid',fgColor=bg); ws.cell(r,16).font=F(size=10,bold=True,color=fg)
            d_=ws.cell(r,14).value
            if isinstance(d_,(int,float)) and isinstance(ln['sap'],(int,float)) and ln['sap'] and abs(d_)>0.1*abs(ln['sap']): ws.cell(r,14).font=F(size=10,bold=True,color='C0392B')
            nl_=max([str(v or '').count('\n')+1 for v in vals[15:]]+[1])
            ws.row_dimensions[r].height=max(19,min(15*nl_,150)); r+=1
    for j,w_ in enumerate([15,30,16,6,13,13,11,9,10,12,12,13,13,12,34,16,18,30,46,18],1): ws.column_dimensions[L(j)].width=w_
    if r>5: ws.add_table(Table(displayName='StockCheck',ref=f'A4:{L(len(H))}{r-1}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=False)))
    ws.freeze_panes='C5'
    pos=wb.sheetnames.index(HEAT_NAME)+1 if HEAT_NAME in wb.sheetnames else 1
    wb.move_sheet('Stock Check',offset=pos-wb.sheetnames.index('Stock Check'))
    pass                                                           # saved once at the end (post_flush)






# ===================== ONE STYLE FOR THE WHOLE WORKBOOK + DEFINITIONS sheet =====================
DEFINITIONS=[
 ('ALERT','Alert','One simple list, strong alert (first tab of MRP_Daily and of the Inbound file - the same rows)','STOCK-OUT IN LEAD TIME (red) = a stock-zero day before a new order can arrive: expedite / use an alternative · STOCK-OUT LATER (orange) = order by the date (first stock-zero day − lead time) · UNDER 21 DAYS (yellow) = under 21 days of stock but in transit covers the plan: watch the arrivals'),
 ('ALERT','Item name / Category','What the material is','Item name = MIS Common Name; a substitute group = one row per member (own numbers; 2nd line = the run-out date with the substitute together), back glass 333 = one row per WSA sheet · Category = your MRP.xlsx Material Group name (column E) > MIS Standard BOM Understanding Name > your Material Group column A (Packaging / Label ...)'),
 ('ALERT','Days of stock / Days incl. in transit','Cover today','Stock today (MB52 / PPC count rolled forward, JW at the OEM left out) ÷ need per day (module plan, average per production day, next 9 weeks); incl. in transit = (stock + all in transit) ÷ need per day. Listed when Days of stock < 21 or a stock-zero day falls inside the lead time'),
 ('ALERT','Stock-zero days inside the lead time / Stock zero from → to','Days without stock','A stock-zero day = that day\'s need is not covered: stock + in transit arrived − plan < 0. Counted from today to today + lead time − 1; every period of the plan is listed, ⚠ = it starts inside the lead time. Back glass 333: one row per WSA sheet (customer project, own supplier stock + PPC arrivals)'),
 ('Substitute Advice','Need / day at line targets','Daily target need','Σ customer daily target (Settings CUSTOMER DAILY TARGET, else line targets) × usage per module from your BOM (\'_Alt\' rows excluded; else MIS Standard BOM), in the MRP unit. Without OEM = customers with OEM = No; with OEM = all. Every other sheet uses the module plan (Need / day (module plan)).'),
 ('All sheets','Usage / day (plan)','Production Plan demand','Modules in the Production Plan on that day × usage per module, per customer / line, OEM customers left out. Drives the projection: stock, shortage dates, planned orders, days of cover, risk lists. Plan usage/day of a count = average of the next 7 days.'),
 ('MRP Daily / Weekly','Days of cover','How long the projected stock lasts','Projected stock at the end of the day (stock + in transit by arrival date − plan usage) ÷ average plan usage/day of the next N days, N = Settings "Days of cover: average plan usage over the next (days)" (21). Blank when no usage ahead.'),
 ('All sheets','OEM','Left out of the calculations','Customers with OEM = Yes (CUSTOMER DAILY TARGET): their plan and the JW stock at the OEM are not part of any projection, shortage date, planned order or risk list. The JW stock stays visible as data (Stock Data "to OEM" lines, Movement Log, No Demand).'),
 ('In Transit','Shipment code check','Which material a shipment is','Description, Supplier Name, RM Code, Final RM Code and the codes on the PO line are compared; the code whose description fits the shipment clearly best (sizes, thickness, busbar / ribbon width, GSM, front / back) is used; blank / unknown codes are matched among the materials the same supplier ships. Every change is listed on Check.'),
 ('MRP Daily / MRP Weekly','Supplier rows','Glass per supplier','Shown only when 2 or more suppliers have stock, in transit or allowed demand for the material - otherwise the row would repeat the material row.'),
 ('RM Issues','Short From (Stock = 0)','Every stock-zero period inside the lead time','All periods where projected stock (counts / MB52 rolled forward + MIS in transit by arrival date − plan) is below zero that START inside the material lead time, each with dates and length ("beyond" = still short at the end of the shortage horizon); none inside: first stock-zero day marked "after lead time"; none at all: not short. Any period inside the lead time = High risk (RM Issues priority High).'),
 ('Inputs','SAP stock (MB52)','Where stock comes from','Your MB52.xlsx (SAP export, same folder, newest MB52*.xlsx): Unrestricted + Quality Inspection = Total Stock; Plant name, Final Material Code, Dashboard Key, Company, Comman Name Cat.1 / Cat.2 are looked up from the MIS MB52 (else MIS Common Name). Stock date = save date of your MB52 file (older than today: rolled forward by usage)'),
 ('RM Issues','Issue log','One issue per raw material at risk','Material short or under 21 days in 16 weeks; Required date = first day under 21 d (or stock 0); Qty required = missing to keep 21 d at the worst point; Qty received = MIS PPC GRN since raised; Short from (stock = 0) = every period stock is below zero that starts inside the lead time (High priority); Revised ETA / Supplier / PO = next MIS shipment; Shipped qty = MIS Final Qty. and Pending qty = MIS Final Pending Qty. of that shipment (same supplier, PO and ETA); Delay = ETA − required date; Production impact = modules not buildable at the worst point × Production Plan Wp, capped at the MW planned on those lines / customers in 16 weeks; substitute group = one issue; Owner / Status / Date Closed / Action kept from the previous run; auto-closed when no longer at risk'),
 ('Inputs','BOM and Stock Count','Where they come from','From the PPC workbook (sheets BOM and Count, newest file with both, name with PPC preferred); MRP.xlsx only when no PPC workbook is in the folder'),

 ('ALERT','Estimated stock today / SAP stock (MB52) / Physical count','Stock of each ALERT row','Estimated stock today = the Stock Check rule: physical count rolled forward to today (count − usage since the count + goods received), not counted: SAP stock (MB52); JW stock at the OEM left out · SAP stock (MB52) / Physical count = the Stock Check columns (count as counted; substitute group = its members; back glass 333 per WSA sheet = the suppliers whose stock is on that sheet)'),
 ('ALERT (Material Status columns)','In transit / Next arrival / Overdue in transit','Supply of each ALERT row','Stock = Estimated stock today, JW stock at the OEM left out · In transit = all pending shipments (MIS In Transit; back glass 333 per WSA sheet: its PPC arrivals) · Next arrival = first arrival from today · Overdue = MIS pending qty whose arrival date has passed (333 per WSA sheet: in today\'s arrivals)'),
 ('Material Status','Rows','Which materials','EVERY material: raw materials, packaging and consumables, also the ones with no demand (brought back Nov 2026). Columns: UOM (after Category; Daily_Material_Status Materials Status other-material rows = PPC Count Base Unit of Measure; that file also has Item Category = PPC Count Item and Consumable Unit = PPC PM_CS_Usage), Stock, SAP stock (MB52), Physical count (= the Stock Check columns), In transit, Next arrival, Overdue in transit, Need / day, Short from, Days of supply, Shortage risk, Days incl. in transit (no INVENTORY / QUALITY columns since Nov 2026)'),
 ('ETD alert','Alert','WSA deliveries not on MIS In Transit',f'Your WSA delivery rows (not cancelled) with an ETD whose BL# is not on MIS In Transit (= not shipped as far as MIS knows) · ARRIVES AFTER STOCK-OUT = expected arrival after the material\'s first stock-zero day (ALERT) · ETD PASSED · SHIPS WITHIN {ETD_SOON} DAYS · ON PLAN; the same rows as the Inbound file'),
 ('ETD alert','ETD vs today / Expected arrival / Delay','How late, when it can arrive','Late N d = base date − ETD; In N d = days until the ETD · expected arrival = your WSA plant date; ETD passed = base date + planned transit (WSA plant date − ETD), i.e. if it ships today · delay = expected arrival − WSA plant date · these rows are in no projection (the outputs use MIS In Transit)'),
 ('Stock Check','SAP stock / Physical count / Count date','Both stock sources side by side','SAP = your MB52.xlsx Unrestricted + Quality Inspection (MIS MB52 only for plant name, final code, key, company, categories) · count = PPC Count sheet; count date = the date column that holds the count'),
 ('ALERT','Below 21 days from / Why (cause) / Days late / Arrival at plant','Which materials drop below 21 days in the next 16 weeks, and why','Below 21 days from = first day projected stock covers < 21 days of need · Why: LATE ARRIVAL = what is on the way would keep 21+ days but arrives after that day, LATE + NOT ENOUGH, NO SUPPLY · Days late = arrival of the first shipment after the below-21-days day minus that day (late = would have kept 21 days if on time) · Arrival at plant (Final AtA) = passed-not-received total + the next arrival dates (MIS Final AtA) · Action second line = expedite / order more. UNDER 21 DAYS LATER (blue) = 21+ days today, below 21 days within 16 weeks, no stock-zero day. Materials with no demand are not listed'),
 ('MRP by Line','Daily plan need / Inv. Days at module plan / Below 21 Days From (module plan) / Days Late (module plan)','MODULE PLAN scenario (gold band)','Need/day = module plan × usage per module, average per production day, next 9 weeks, OEM left out · Inv. days = (own stock + in transit) ÷ that need · Below 21 Days From / Days Late = same as the target columns but with the module-plan need = the ALERT values'),
 ('All sheets','No demand','Material without demand = no shortage status, no action','Needs at least one of: production requirement (plan x BOM usage) in the next shortage-horizon days (Settings, 60), daily target need, average planned demand. Otherwise: not Critical / At risk / Watch / Short, no Expedite / Order now / Release order, not on the ALERT'),
 ('BOM Material Dashboard - Daily','Day columns','Daily cover, same story as the ALERT','Same rows / groups / order and the same design as BOM Material Dashboard - Weekly; each day = stock available that morning (stock today + MIS arrivals up to that day - plan of the days before) / that day\'s demand (na = no demand); E = stock today / today\'s demand; first value below 1 = the ALERT first stock-zero day; colours = the weekly bins (red = below 1 day, orange / amber = below 21 days, blue = 21+ days, grey = na)'),
 ('BOM Material Dashboard - Weekly','Heatmap cells','When it runs short','One column per week = the daily number on that date (stock that morning ÷ that day\'s demand, days) = the same number as BOM Material Dashboard - Daily and the RM Material Dashboard on that date; red = short (below 1 day), orange / amber = below 21 days, blue = 21+ days, na = no demand that day; same rows as the ALERT sheet, grouped RM / PM / CS, highest risk first (same rule as every dashboard): first day the daily cover goes below 1 (stock-zero day), then below 5 days, then the lowest cover today'),
 ('Stock Check','Data dates','How old the inputs are','MIS = save date of the MIS file, WSA = report date of the planning file, Counts = oldest–newest count date; red when older than 2 days. Old MIS: SAP stock of materials that are not counted is rolled forward by usage per day × days since the MIS date.'),
 ('Stock Check','Current month: opening / GR / used / closing','Stock movement of this month','Opening = the first count of the month (PPC Count date columns; a material with several count rows needs every row counted); not counted: the stock of the first MRP run of the month (Movement Log, fills as you run daily) · GR qty = MIS In Transit PPC GRN with GR. Date in this month (alternate codes converted) · Used = opening + GR after the opening date up to the closing date − closing (negative = a GR or count is missing) · Closing = the latest count (not counted: SAP stock today) · Projected = MRP projection (stock + in transit − module plan) at month end'),
 ('Stock Check','Shortage risk','The ALERT label','The same class as the ALERT sheet: STOCK-OUT IN LEAD TIME / STOCK-OUT LATER (+ first stock-zero day) · UNDER 21 DAYS · UNDER 21 DAYS LATER (+ the date it falls below 21 days) · OK = not on the ALERT (No demand / No need = no plan demand) · rows sorted in the ALERT order, then OK · back glass 333: one line per WSA sheet; a substitute group member = its group'),
 ('Stock Check','Estimated stock today','Count rolled forward to today','Count − usage per day × days since count + received since count (MIS PPC GRN with GR. Date after the count date); usage per day = average Production Plan demand of the next 7 days (OEM plan left out); not counted = SAP stock (MB52). This is the stock every shortage date uses (= ALERT Estimated stock today).'),
 ('Stock Check','Item name / Category / supplier rows','What the material is, stock per supplier','Item name / Category = the ALERT (MIS Common Name; item_category) · a material with stock from 2+ suppliers (e.g. glass FLAT / KIBING) = one row per supplier, "Item name - supplier": SAP = MB52 batches of that supplier (Batch Master), count = the PPC count by supplier (Supplier Code), estimated = the MRP Daily supplier row; usage / received since count split by the count share; Unassigned = stock MB52 / the count does not name'),
 ('RM Issues','Priority','Shortage priority','High = a stock-zero day inside the lead time · Medium = stock zero later (time to order) · Low = under 21 days at the module plan, no stock-zero day · OK · No demand. The stock lines behind the Stock column (source, source code, factor, storage location) are on the Stock Data sheet'),
 ('Stock Check','Lines (share of plan volume)','Which lines use the material','Share of the material\'s planned use over the plan horizon by production line (Production Plan × BOM usage)'),
 ('Stock Check','Project Name (share of plan)','Which projects use it','Same share by Project Name (Settings MIS LINK)'),
 ('Stock Check','Usage per module by customer','BOM usage behind the need','Usage per module per customer with its source: your PPC BOM sheet first, else MIS Standard BOM'),
 ('Stock Check','Module (MIS BOM, share of plan)','Module types using it','Short module name from the MIS BOM name with its share'),
 ('ALERT / Material Status','Need / day','Production view','Module plan (PPC Plan + Revised Plan) × BOM usage, average per production day, next 9 weeks, OEM left out (no Reorder level / Status / On hold columns since Nov 2026)'),
 ('ALERT / Material Status','Lead time (days)','Order-to-arrival time','WSA ETD → arrival history, else Settings default (RM 49/50, PM 14)'),
 ('Stock Data / Movement Log','Stock / to OEM Stock','Stock here / at the OEM','MB52 (or physical count) / MIS RM Status "JW WSA to OEM" - the JW stock at the OEM is in no calculation'),
 ('ALERT / MRP by Line','Arrival at plant (Final AtA)','Next arrivals','Passed-not-received total, next 4 arrival dates, rest as one total'),
 ('ALERT / Material Status / Substitute Advice','Substitute group','Material + its substitutes planned as one','Settings SUBSTITUTES (lines blank, valid today): the shortage date, need per day and days of cover use the group stock + group in transit + group demand (each code x its factor); Stock / In transit columns show each code\'s own quantity. An alternative listed ALONE for a customer in your BOM carries that customer\'s demand.'),
 ('MRP by Line','Line can run (days) / Line stops on','How long the line runs','Earliest day a material it needs is short after sharing stock between lines (Line ranking, Vendor priority, Allowed vendors, Settings Substitutes) - and after using your approved ALTERNATIVE (your BOM) for that line\'s customers: free stock of the alternative = what its own plan does not need within its lead time'),
 ('MRP by Line','First needed / Needed now','When this line needs it','First plan day the line uses it; Yes when within Settings look-ahead (49 days)'),
 ('MRP by Line','Plan on this line','Who / when on this line','Usage windows and customers in date order; "Same as today" = unchanged'),
 ('MRP by Line','This line: daily plan need','This line\'s need','This line\'s module plan × usage per module, average per production day over the next 9 weeks (OEM left out)'),
 ('MRP by Line','In transit usable by this line','Supply this line may use','MIS in transit from suppliers allowed for this line (Settings ALLOWED VENDORS / Vendor priority)'),
 ('MRP Daily','Gross requirement / Projected available balance','Day-by-day MRP record','Plan demand per day; balance = previous − requirement + in transit'),
 ('MRP Daily','Firm receipt / Planned order receipt','Open PO / order suggestion','Open PO without vessel; planned receipt keeps stock at target days (lead-time offset)'),
 ('Action List','Purpose','One row per material, highest shortage risk first','Same shortage-risk order as ALERT / RM Issues. Priority: High = stock zero inside the lead time, Medium = short later, Low = under 21 days, OK. Materials with no demand get no shortage status.'),
 ('Action List','Component / Component description / UOM','The material','Material code, MIS Common Name | description, unit of measure (MRP unit)'),
 ('Action List','Lines / Project Name / Usage per module by customer / Module (MIS BOM)','Who uses the material','Share of plan volume by production line and project; usage per module by customer (your PPC BOM first, else MIS Standard BOM); module type (MIS BOM name) and its share of the plan'),
 ('Action List','Short from → to (with in transit) / Below 21 Days From / Days Late','When the material runs short','Same logic as ALERT: projected stock (stock - plan + MIS in transit) below zero; first day cover is under 21 days; days between the shortage and the arrival at plant'),
 ('Action List','Material lead time (days) / Stock / to OEM Stock','Lead time and stock','Lead time = Settings / WSA history; Stock = MB52 or PPC count rolled forward (JW at the OEM left out of calculations); to OEM Stock = JW stock at the OEM, shown only'),
 ('Action List','In transit today / Arrival at plant (Final AtA)','MIS shipments not yet received','MIS In Transit Final Pending Qty., next arrival dates (Final AtA; late lines counted today)'),
 ('Action List','Production Daily Target / Inv. Days at line daily target (without OEM, with OEM)','Cover at the customers\' daily targets','Target need per day = customer daily target x usage per module (Settings CUSTOMER DAILY TARGET); Inv. Days = (stock + in transit) ÷ target need per day, red under 21; "stock only" excludes transit; without OEM = customers with OEM = No, with OEM adds the OEM customer and JW stock at OEM'),
 ('Action List','Daily plan need / Inv. Days at module plan / Below 21 Days From (module plan) / Days Late (module plan)','Cover at the module plan','Average need per production day (module plan x usage per module, next 9 weeks, OEM left out); Inv. Days = (stock + in transit) ÷ that need; below-21 and days-late as on ALERT'),
 ('Action List','In transit / Open PO (WSA RM Planning) / Alternative BOM','Cross-checks','Quantities on the WSA planning sheets (in transit with a vessel, Open PO without vessel); alternative BOM = codes in your BOM that replace the standard code, with their stock and customers'),
 ('ALERT (formulas)','All columns','Same rows as ALERT, as live Excel formulas','Yellow = input, white = formula. Reads ALERT Calc (projection), ALERT Plan lines and ALERT Arrival lines; used to prove every ALERT row adds up. Column "Check vs ALERT" compares each calculated column with the ALERT sheet (OK / DIFF)'),
 ('ALERT Calc','Daily projection','Projection behind each ALERT row','Per ALERT row: Plan = SUMIFS of ALERT Plan lines, Arrivals = SUMIFS of ALERT Arrival lines (counted on max(arrival, today)); projected stock = previous day + arrivals - plan; stock zero = 1 when projected stock < 0; below 21 days = 1 when projected stock ÷ need per day < 21 (first 16 weeks)'),
 ('ALERT Plan lines','Demand (formula)','Plan behind ALERT Calc','One row per ALERT row x material x day x customer x production line: Modules x Qty per module x Factor (OEM customers left out; substitute members on their own rows; qty per module from your PPC BOM else MIS Standard BOM)'),
 ('ALERT Arrival lines','Qty (formula)','Shipments behind ALERT Calc','One row per ALERT row x shipment: MIS In Transit pending qty x Factor, counted on MAX(arrival, today); material traced to your BOM / WSA PO + BL#'),
 ('In Transit Status','Status groups / shipment lines','MIS In Transit by STATUS','Free-text MIS STATUS grouped (In transit, At origin, At port, Held, Partly received, Transloaded, Delivered, Other, No status); pending qty = Final Pending Qty.; every shipment line listed below the summary'),
 ('No Demand','All columns','Stock or supply with nothing planned to use it','A = MRP materials with plan 0; B = other RM codes with MB52 / JW stock not in the MIS plan nor in your BOM for a planned model; C = cell codes (not planned here). Amber = more still coming'),
 ('Settings','All settings','Every planning rule in one place','Target days (A / B / C / packaging, + Z extra), line cover limit, daily targets, MIS LINK, allowed vendors, vendor priority, substitutes, batch supplier codes, line ranking. Read first on every run; copy the sheet into MRP.xlsx to keep your edits'),
 ('Module Plan','All columns','Modules per day by line and customer','From the PPC file: Plan with Revised Plan on top (a date + line in the Revised Plan replaces that date + line in Plan); MIS customer / BOM name = the Settings MIS LINK'),
 ('BOM','All columns','Copy of the BOM in the PPC file','Edit it in the PPC file, not here. MIS BOM Name comes from the module plan (customers using each Module Model). Used for CS / PM and for raw materials of customers it lists'),
 ('Stock Count','All columns','Copy of the PPC Count sheet','Today Stock Qty x Quantity = Today Qty_BOM; a physical count replaces MB52 for that material and is rolled forward to today (receipts - usage since the count date)'),
 ('How it works','Text','Short description of the current logic','Inputs, calculation order, sheet list and daily routine'),
 ('MIS Layout','Sheet / Column / Header','Headers of MIS columns read by position','Stored each run and compared with the next run; a changed header is reported on the Check sheet (MIS layout)'),
 ('MRP Weekly','All columns','Weekly view of MRP Daily','Live formulas on MRP Daily: Plan / In transit / Open PO = week totals; stock and days of cover = end of week'),
 ('Check','All columns','Discrepancies between MIS, the WSA planning file, your PPC file and the MRP','Severity High / Medium / Low; areas MIS layout, Stock, Plan (next 30 days), In transit (matched shipment by shipment on BL number), Open PO, Units, Master data, Your BOM. Difference = value B - value A. Low rows are collapsed (click + on the left)'),
]

def unify_workbook_style(path):
    """Same look on every sheet: Calibri, navy titles and table headers, dark-grey text, no gridlines, navy tabs, no charts."""
    from openpyxl.styles import Border, Side
    from openpyxl.utils import range_boundaries
    wb=post_wb(path)
    NAVY='14213D'; INK='2C3E50'; thin=Side(style='thin',color='DDE3E8')
    # Definitions sheet
    if 'Definitions' in wb.sheetnames: del wb['Definitions']
    d=wb.create_sheet('Definitions')
    d['A1']='Definitions — every header and how it is calculated'; d['A2']='Sources: MIS RM Dashboard (MB52, In Transit, RM Status), WSA RM planning file, PPC file (Plan + Revised Plan = module plan, BOM = quantity per module, Count = physical stock, In_Transit), MB52.xlsx (SAP stock), MRP.xlsx (Settings, Batch Master), Customer_Tracker files.'
    for j,h in enumerate(['Sheet','Header','Definition','Calculation / source'],1): d.cell(4,j,h)
    for i,(a,b,c,e) in enumerate(DEFINITIONS,5):
        for j,v in enumerate((a,b,c,e),1):
            x=d.cell(i,j,v); x.alignment=Alignment(wrap_text=True,vertical='top')
    for col,w_ in zip('ABCD',[22,34,46,90]): d.column_dimensions[col].width=w_
    from openpyxl.worksheet.table import Table, TableStyleInfo
    d.add_table(Table(displayName='Definitions',ref=f'A4:D{4+len(DEFINITIONS)}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=True)))
    hdr_rows={'Action List':[4],'MRP by Line':[4]}
    cut_=wb.sheetnames.index('In Transit Status') if 'In Transit Status' in wb.sheetnames else len(wb.sheetnames)
    hidden_=set(wb.sheetnames[cut_:])                                   # hidden sheets: no cell-by-cell restyle (saves time)
    HEAD='F4F6F8'; INK2='1F2D3D'; hb=Border(bottom=Side(style='medium',color='C9D2DC'),left=Side(style='thin',color='E3E8EE'),right=Side(style='thin',color='E3E8EE'))
    CARD={'green':('E3F4EC','1E8449'),'blue':('E8F1FB','2471A3'),'amber':('FEF3E2','B9770E'),'grey':('F2F4F6','5D6D7E'),'red':('FCE8E6','C0392B'),'teal':('E3F2F0','138D75')}
    def card_of(rgb):                                       # old banner colour -> card tint of the same family
        h=str(rgb or '')[-6:].upper()
        if h in ('548235','70AD47','375623','229954'): return CARD['green']
        if h in ('2E75B6','2F5597','1F4E79','2E86C1','2471A3'): return CARD['blue']
        if h in ('833C0B','C65911','ED7D31','D68910'): return CARD['amber']
        if h in ('C00000','C0392B'): return CARD['red']
        return CARD['grey']
    for ws in wb.worksheets:
        ws._charts=[]; ws._images=[]                                    # no charts anywhere
        ws.sheet_state='visible'
        ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor=('C00000' if ws.title=='ALERT' else NAVY)
        if ws.title in hidden_ and ws.title!='Definitions': continue
        big_=ws.title in ('MRP Daily','MRP Weekly')                     # written in Calibri already: no cell-by-cell font pass
        own=ws.title in ('ALERT',ALERT_FX,ALERT_CALC,ALERT_PL,ALERT_AL,'Material Status',HEAT_NAME,DAILY_NAME,'Stock Check','RM Issues','Action List')   # already in house style
        if not own:
            for row in ([] if big_ else ws.iter_rows()):
                for c in row:
                    if c.value is None: continue
                    f=c.font; col=f.color.rgb if (f.color is not None and isinstance(f.color.rgb,str)) else None
                    c.font=Font(name='Calibri',size=(f.sz or 10),bold=f.b,italic=f.i,color=(col if col and col[-6:] not in ('000000',) else INK))
            t=ws.cell(1,1)
            if isinstance(t.value,str): t.font=Font(name='Calibri',size=18,bold=True,color=NAVY)
            rows=set(hdr_rows.get(ws.title,[]))
            for tb in ws.tables.values():
                c1,r1,c2,r2=range_boundaries(tb.ref); rows.add(r1)
                for j in range(c1,c2+1):
                    c=ws.cell(r1,j); c.fill=PatternFill('solid',fgColor=HEAD); c.font=Font(name='Calibri',size=10,bold=True,color=INK2)
                    c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=hb
            for r in hdr_rows.get(ws.title,[]):
                for c in ws[r]:
                    if c.value is None: continue
                    c.fill=PatternFill('solid',fgColor=HEAD); c.font=Font(name='Calibri',size=10,bold=True,color=INK2)
                    c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=hb
            if ws.title=='Action List':                         # row 3 group banners -> soft cards (same palette as the KPI cards)
                for c in ws[3]:
                    fg=c.fill.fgColor.rgb if c.fill and c.fill.fill_type=='solid' else None
                    if not fg or str(fg)[-6:] in ('FFFFFF','000000'): continue
                    bg,tx=card_of(fg); c.fill=PatternFill('solid',fgColor=bg); c.border=Border(top=Side(style='medium',color=tx))
                    if c.value is not None: c.font=Font(name='Calibri',size=10,bold=True,color=tx)
        code_re=re.compile(r'^(RM|PM|CS)[A-Z]*\d{5,}$')                 # every material-code cell: same format on every sheet
        cfont=Font(name='Calibri',size=10,bold=False,color='2C3E50'); cbord=Border(left=thin,right=thin,top=thin,bottom=thin)
        for row in ([] if ws.title in (ALERT_PL,ALERT_AL,DAILY_NAME,HEAT_NAME) else ws.iter_rows()):      # the line sheets: plain data (speed); the daily dashboard keeps its dark heatmap look
            for c in row:
                if isinstance(c.value,str) and code_re.match(c.value.strip()):
                    c.font=cfont; c.border=cbord; c.alignment=Alignment(horizontal=c.alignment.horizontal,vertical='center',wrap_text=True)
        for tb in ws.tables.values():                           # table safety: header names unique, not blank, same as the table's column list
            c1,r1,c2,r2=range_boundaries(tb.ref); seen_=set()
            for k_,j in enumerate(range(c1,c2+1)):
                c=ws.cell(r1,j); v=str(c.value).strip() if c.value not in (None,'') else f'Column{k_+1}'
                while v in seen_: v=v+' '
                seen_.add(v); c.value=v
                if k_<len(tb.tableColumns): tb.tableColumns[k_].name=v
    d['A1'].font=Font(name='Calibri',size=18,bold=True,color=NAVY); d['A2'].font=Font(name='Calibri',size=10,italic=True,color='8A949E')
    for w_ in wb.worksheets: w_.sheet_state='visible'                # every sheet visible (your choice, Nov 2026)
    wb.active=0
    pass                                                           # saved once at the end (post_flush)







# ===================== RISK CHARTS: treemap (where the shortage is) + heatmap (when it runs short) - drawn with cells =====================

# ===================== STOCK MANAGEMENT dashboard (Store > Warehouse > Inventory > Production > Quality > Aging) =====================
def add_stock_mgmt_dashboard(path, A, AL_ORDER, mlog, ships, d0, nd, mis_path):
    """One row per raw material, grouped like a stock management system:
       STORE opening stock (previous MRP run) · WAREHOUSE stock in (MIS arrivals since then) / stock out (implied) / closing ·
       INVENTORY reorder level · PRODUCTION need per day · QUALITY status + on hold (MB52 blocked + QI) ·
       AGING last arrival / days of supply / movement class / recommended action / stock value. Summary boxes + shortage watch below.
       Brought back Nov 2026 (user: "bring back stock management and include all materials not just raw materials") after it was merged
       into the ALERT: rows = EVERY material (RM + packaging + consumables, also no demand), same columns and rules as before; the ALERT
       keeps its Material Status columns."""
    from openpyxl.styles import Border, Side
    wb=post_wb(path)
    if 'Material Status' in wb.sheetnames: del wb['Material Status']
    ws=wb.create_sheet('Material Status'); ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor='14213D'
    F=lambda **k: Font(name='Calibri',**k)
    NAVY='14213D'; INK='1F2D3D'; MUTED='8A949E'; LINE='E3E8EE'
    STORE,WH,INV,PROD,QUAL,AGE=('2E86C1','D6EAF8'),('229954','D5F5E3'),('E67E22','FDEBD0'),('6C3483','E8DAEF'),('17A589','D1F2EB'),('1B2631','D6DBDF')
    def fill(c,col): c.fill=PatternFill('solid',fgColor=col)
    thin=Side(style='thin',color=LINE); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    widths=[5,16,18,14,17,11,11,11,10,10,11,11,11,11,10,11,9,10,12]
    for j,w_ in enumerate(widths,1): ws.column_dimensions[L(j)].width=w_
    # ---------- data ----------
    base=[x for x in AL_ORDER if not x.get('sup_row')]                 # EVERY material: RM, packaging, consumables (user, Nov 2026: 'include all materials')
    price={}; qi={}; blk={}
    mis=wb_ro(mis_path) if mis_path else None
    if mis:
        for r in list(mis['In Transit'].iter_rows(values_only=True))[1:]:
            if r and len(r)>39 and isinstance(r[8],(int,float)) and r[8]>0: price.setdefault(str(r[39] or '').strip(),[]).append(r[8])
        for r in list(mis['MB52'].iter_rows(values_only=True))[1:]:
            if r and len(r)>14 and r[2]:
                c=str(r[2]).strip(); m=(BOM_ALIAS[0] or {}).get(c,c)
                qi[m]=qi.get(m,0)+(r[12] if isinstance(r[12],(int,float)) else 0); blk[m]=blk.get(m,0)+(r[9] if isinstance(r[9],(int,float)) else 0)
    price={k:sum(v)/len(v) for k,v in price.items()}
    load_mis_cat(mis_path)                                       # category: item_category() (the ALERT uses the same)
    prev={}
    for e in sorted(mlog or [],key=lambda e:e['date']):
        if e['date']<d0: prev[e['code']]=e
    rows=[]
    for i,x in enumerate(base,1):
        b=x['b']; m=b['mat']; need=x.get('pl_need') or 0.0; lt=int(b.get('lt_used') or 0)   # module plan need/day
        close=float(b['stock']); pe=prev.get(m); opening=float(pe['stock']) if pe else close
        since=pe['date'] if pe else d0
        sin=sum(s_['qty'] for s_ in ships.get(m,[]) if s_.get('raw') is not None and since<s_['raw']<=d0)
        sout=max(0.0,opening+sin-close)
        rop=need*lt+need*SAFETY_DAYS
        status='Out of Stock' if close<=0 else ('Below reorder' if close<rop else 'Above reorder')
        if status=='Above reorder' and (x.get('wins') or []): status=f"Short {(d0+pd.Timedelta(days=x['wins'][0][0])):%d-%b}"     # never labelled safe when short
        hold=qi.get(m,0)+blk.get(m,0)
        last=max((s_['raw'] for s_ in ships.get(m,[]) if s_.get('raw') is not None and s_['raw']<=d0),default=None)
        dos=(max(0.0,close)/need) if need else None
        mv='No demand' if dos is None else ('Fast' if dos<=30 else ('Normal' if dos<=60 else ('Slow' if dos<=90 else 'Excess')))
        ri=risk_info(x,d0); k_=ri['k']; mv=ri['cls']                      # shortage-risk class (not stock ageing)
        val=((max(0.0,close)+(x.get('tr_tot') or 0))/need) if need else None      # days incl. in transit (no unit cost)
        sh_=ships.get(m,[])
        nxt_=min((s_['raw'] for s_ in sh_ if s_.get('raw') is not None and s_['raw']>=d0),default=None)
        ovd_=sum(s_['qty'] for s_ in sh_ if s_.get('raw') is not None and s_['raw']<d0)
        rows.append(dict(i=i,code=m,it=float(x.get('tr_tot') or 0),nxt=nxt_,ovd=ovd_,name=str(b.get('common') or m),cat=item_category(m),uom=b.get('uom'),
                         op=opening,sin=sin,sout=sout,close=close,rop=rop,need=need,status=status,hold=hold,last=last,dos=dos,mv=mv,val=val,k=k_,lt=lt,
                         alt=strip_code(str(AL_VAL.get((m,'Alternative BOM (yours): code - stock - customers')) or '').split('\n')[0]) or ''))
    # ---------- ALIGN WITH THE ALERT SHEET (user, Nov 2026: 'stock management sheet and alert sheet are not aligned') ----------
    # ONE source of truth: when a material is on the ALERT sheet, its numbers, class and short-from date come from the ALERT row, so the two
    # sheets can never disagree (before: Under-21-days used another rule, back glass 333 was one row here and one row per WSA sheet there).
    CLS_={0:'ALERT',1:'Short later',2:'Under 21 d',3:'Under 21 d later',4:'OK',5:'No demand'}
    def _from_alert(ar, i_, x_=None):
        nm=str(ar['name'])+(f"\n(WSA sheet: {ar['where']})" if ar.get('where') else '')
        k_=ar.get('k0'); need_=float(ar['need'] or 0.0)
        return dict(i=i_,code=ar['code'],it=float(ar.get('it') or 0.0),nxt=(pd.Timestamp(ar['nx'][0]) if ar.get('nx') else None),
                    ovd=(ar.get('ovd') if ar.get('ovd') is not None else (x_['ovd'] if x_ else 0.0)),ovd_txt=("in today's arrivals" if (ar.get('where') and ar.get('ovd') is None) else None),
                    name=nm,cat=item_category(ar['code']),uom=ar.get('uom'),op=float(ar['stock'] or 0.0),sin=0.0,sout=0.0,close=float(ar['stock'] or 0.0),
                    rop=float(ar['rop']),need=need_,status=ar['status'],hold=float(ar.get('hold') or 0.0),last=(x_['last'] if x_ else None),
                    dos=(float(ar['days']) if ar['days'] is not None else None),mv=CLS_[ar['cls']],val=(float(ar['days_it']) if ar['days_it'] is not None else None),k=k_,lt=int(ar['lt'] or 0),alt=(x_['alt'] if x_ else ''),ar=ar)
    AR={}
    for ar in ALERT_ROWS: AR.setdefault(ar['code'],[]).append(ar)
    new_rows=[]
    for x in rows:
        ars=AR.get(x['code']) or []
        if len(ars)>1 or (ars and ars[0].get('where')):               # split per WSA sheet (back glass 333): same rows as the ALERT
            for ar in ars: new_rows.append(_from_alert(ar,0,x))
        elif ars:                                                      # same material, same numbers as the ALERT row
            y=_from_alert(ars[0],0,x)
            if ars[0].get('mem'):                                      # substitute group shown per member (EPE 480 / EPE 420): own stock, group class + short date
                y.update(uom=x['uom'],name=x['name'],close=x['close'],op=x['op'],it=x['it'],nxt=x['nxt'],ovd=x['ovd'],hold=x['hold'],dos=x['dos'],val=x['val'],rop=x['rop'],status=x['status'])
            new_rows.append(y)
        else:
            gm=AR.get(grp_main(x['code'])) if grp_follower(x['code']) else None     # group follower: the class of its group's ALERT row
            if gm:
                x.update(mv=CLS_[gm[0]['cls']],k=gm[0].get('k0'))
            elif x['mv'] in ('ALERT','Short later','Under 21 d'):      # nothing on the ALERT -> it is not at risk on this sheet either
                x.update(mv='OK' if x['need'] else 'No need',k=None)
            new_rows.append(x)
    on_alert=sorted([y for y in new_rows if y.get('ar') or y['mv'] in CLS_.values()],key=lambda y:(list(CLS_.values()).index(y['mv']) if y['mv'] in CLS_.values() else 9,
                    y['k'] if y['k'] is not None else ((y['ar']['why']['t'] or 0) if y.get('ar') else 0),y['val'] if y['val'] is not None else 1e9,str(y['name'])))
    rest=[y for y in new_rows if not (y.get('ar') or y['mv'] in CLS_.values())]
    rows=on_alert+rest
    for i_,y in enumerate(rows,1): y['i']=i_
    # Stock SAP + Stock Count next to Stock (user, Nov 2026): SAP = MB52 stock of the material; Count = PPC physical count (not rolled forward);
    # Stock = what the MRP uses (count rolled forward to today, else SAP). Same rule as the Stock Check sheet. Rows per WSA sheet (333): not SAP / count -> '-'.
    # SAP stock (MB52) / Physical count = the Stock Check columns (user, Nov 2026; stock_src): ALERT row -> its SAP / count (substitute
    # group member shown with its own stock -> its own; back glass 333 per WSA sheet -> the suppliers of that sheet, else '-').
    byb_={x['b']['mat']:x['b'] for x in base}
    for y in rows:
        ar_=y.get('ar'); y['nosrc']=False
        if ar_ is not None and not ar_.get('mem'):
            y['sap']=ar_.get('sap'); y['cnt']=ar_.get('cnt'); y['nosrc']=bool(ar_.get('where') and ar_.get('sups') is None)
        else:
            ss_=stock_src(byb_[y['code']]) if y['code'] in byb_ else dict(sap=None,cnt=None)
            y['sap']=ss_['sap']; y['cnt']=ss_['cnt']
    # rows 1-7 (title banner + process chevrons) removed (user request Nov 2026): the group band is row 1, the header row 2
    # ---------- table headers ----------
    groups=[('',1,5,NAVY),('STORE · stock',6,8,STORE[0]),('WAREHOUSE · supply',9,11,WH[0]),('INVENTORY',12,12,INV[0]),('PRODUCTION',13,13,PROD[0]),('QUALITY',14,15,QUAL[0]),('AGING ANALYSIS',16,19,NAVY)]
    for t,c1,c2,col in groups:
        if c2>c1: ws.merge_cells(start_row=1,start_column=c1,end_row=1,end_column=c2)
        c=ws.cell(1,c1,t); c.font=F(size=10,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='center',vertical='center')
        for j in range(c1,c2+1): fill(ws.cell(1,j),col)
    ws.row_dimensions[1].height=20
    H=['Sr.','Item code','Item name','','Category','Stock','SAP stock (MB52)','Physical count','In transit','Next arrival','Overdue in transit','Reorder level','Need / day','Status','On hold','Short from','Days of supply','Shortage risk','Days incl. in transit']
    hc=[NAVY]*5+[STORE[0]]*3+[WH[0]]*3+[INV[0]]+[PROD[0]]+[QUAL[0]]*2+[NAVY]*4
    for j,(h,col) in enumerate(zip(H,hc),1):
        c=ws.cell(2,j,h); c.font=F(size=9,bold=True,color='FFFFFF'); fill(c,col); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=bd
    ws.merge_cells('C2:D2'); ws.row_dimensions[2].height=30
    CH={'Above reorder':('EAF2F8','1F618D'),'Below reorder':('FCF3CF','9A7D0A'),'Out of Stock':('FADBD8','C0392B'),
        'ALERT':('FADBD8','C0392B'),'Short later':('FDEBD0','B9770E'),'Under 21 d':('FCF3CF','9A7D0A'),'Under 21 d later':('EAF2F8','1F618D'),'OK':('D5F5E3','1E8449'),'No need':('F2F4F6','8A949E'),'No demand':('F2F4F6','8A949E')}
    r=3
    for x in rows:
        vals=[x['i'],x['code'],x['name'],None,x['cat'],round(x['close']),(round(x['sap']) if x.get('sap') is not None else '-'),(round(x['cnt']) if x.get('cnt') is not None else ('-' if x.get('nosrc') else 'no counting')),round(x['it']),(x['nxt'].to_pydatetime() if x['nxt'] is not None else 'none'),(x.get('ovd_txt') or (round(x['ovd']) if x['ovd'] else '-')),round(x['rop']),round(x['need'],1),
              x['status'],round(x['hold']) if x['hold'] else '-',((d0+pd.Timedelta(days=x['k'])).to_pydatetime() if x['k'] is not None else 'no shortage'),(round(x['dos']) if x['dos'] is not None else '-'),x['mv'],(round(x['val']) if x['val'] is not None else '-')]
        fmts=[None,None,None,None,None,'#,##0','#,##0','#,##0','#,##0','dd-mmm','#,##0','#,##0','#,##0.0',None,'#,##0','dd-mmm-yy','0',None,'0']
        for j,(v,fm) in enumerate(zip(vals,fmts),1):
            c=ws.cell(r,j,v); c.font=F(size=10,color=INK); c.border=Border(bottom=thin)
            c.alignment=Alignment(horizontal=('left' if j in (2,3,5) else 'center'),vertical='center')
            if fm: c.number_format=fm
            if j in (6,7,8):  fill(c,'F4F9FD')
            if j in (9,10,11): fill(c,'F4FBF7')
            if j==12: fill(c,'FEF8F1')
            if j==13: fill(c,'F8F4FB')
            if j in (7,8) and isinstance(v,str): c.font=F(size=10,color=MUTED)
        ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
        for j,key in ((14,x['status']),(18,x['mv'])):
            bg,fg=CH.get(key,('FDEBD0','B9770E')); fill(ws.cell(r,j),bg); ws.cell(r,j).font=F(size=10,bold=True,color=fg)
        if x['k'] is not None: ws.cell(r,16).font=F(size=10,bold=True,color='C0392B')
        if isinstance(vals[14],(int,float)): ws.cell(r,15).font=F(size=10,bold=True,color='B9770E')
        ws.row_dimensions[r].height=19; r+=1
    # ---------- summary boxes ----------
    r+=1; top=r
    def box(c1,c2,title,col,lines):
        ws.merge_cells(start_row=top,start_column=c1,end_row=top,end_column=c2)
        t=ws.cell(top,c1,title); t.font=F(size=11,bold=True,color='FFFFFF'); t.alignment=Alignment(vertical='center',indent=1)
        for j in range(c1,c2+1): fill(ws.cell(top,j),col)
        for i_,(a,v,fm,vc) in enumerate(lines,1):
            ws.merge_cells(start_row=top+i_,start_column=c1,end_row=top+i_,end_column=c2-1)
            x_=ws.cell(top+i_,c1,'•  '+a); x_.font=F(size=10,color=(vc if vc and fm is None and v is None else INK))
            y_=ws.cell(top+i_,c2,v); y_.font=F(size=11,bold=True,color=vc or NAVY); y_.alignment=Alignment(horizontal='right')
            if fm: y_.number_format=fm
            for j in range(c1,c2+1): fill(ws.cell(top+i_,j),'F7F9FB'); ws.cell(top+i_,j).border=Border(bottom=Side(style='thin',color='ECEFF3'))
    LRr=line_run_days([x for x in A if is_rm(x['b']['mat'])],d0,nd,ALLOC_STOP_)
    stop14=[l for l in LRr if l['t'][0]<COVER_CRITICAL]; first_stop=min(LRr,key=lambda l:l['t'][0]) if LRr else None
    shorts=[x for x in rows if x['k'] is not None]; f0=min(shorts,key=lambda x:x['k']) if shorts else None
    box(1,4,'⚠️  KEY RISK HIGHLIGHTS',NAVY,[('Short inside transit lead time (ALERT)',sum(1 for x in rows if x['mv']=='ALERT'),'0','C0392B'),('Short later (order in time)',sum(1 for x in rows if x['mv']=='Short later'),'0','B9770E'),('Under 21 days (+ later)',sum(1 for x in rows if x['mv'] in ('Under 21 d','Under 21 d later')),'0','9A7D0A'),
        ('Earliest shortage: '+(f"{f0['name'][:16]} {(d0+pd.Timedelta(days=f0['k'])):%d-%b}" if f0 else 'none'),None,None,'C0392B' if f0 else '1E8449'),
        ('Lines stopping ≤ 14 days',len(stop14),'0','C0392B' if stop14 else '1E8449'),
        ('First line stop: '+(f"{first_stop['name']} {(d0+pd.Timedelta(days=first_stop['t'][0])):%d-%b}" if first_stop and first_stop['t'][0]<nd else 'none'),None,None,'C0392B' if (first_stop and first_stop['t'][0]<nd) else '1E8449')])
    b_=[0]*6                                                      # shortage timeline from the SAME rows as the table / ALERT sheet
    for x in rows:
        k=x['k']; b_[5 if k is None else (0 if k<=7 else (1 if k<=14 else (2 if k<=30 else (3 if k<=60 else 4))))]+=1
    tl=list(zip(['≤ 7 DAYS','8 – 14 DAYS','15 – 30 DAYS','31 – 60 DAYS','> 60 DAYS','NO SHORTAGE'],b_))
    box(6,9,'⏱  SHORTAGE TIMELINE (first short day)',NAVY,[(lab.title().replace('Days','days'),v,'0',('C0392B' if i_<2 else ('B9770E' if i_<4 else '9A7D0A'))) for i_,(lab,v) in enumerate(tl[:5])])
    al=[x for x in rows if x['k'] is not None and x['k']<x['lt']]
    box(11,19,'⚠️  SHORTAGE WATCH (short inside transit lead time)','C0392B',
        [(f"{x['name'][:26]} – short {(d0+pd.Timedelta(days=x['k'])):%d-%b}"+(f" · alt {x['alt'].split(':')[0][:24]}" if x['alt'] else ''),None,None,'C0392B') for x in al[:5]] or [('No material short inside its transit lead time',None,None,'1E8449')])
    # five cards
    top2=top+7
    cards=[('🏪 STORE','Short within 7 days',tl[0][1],'0',STORE),('🏭 WAREHOUSE','Lines stopping ≤ 14 days',len(stop14),'0',WH),
           ('⚙️ PRODUCTION','Materials short in the plan',len(shorts),'0',PROD)]      # no INVENTORY / QUALITY card (user, Nov 2026)
    spans=[(1,5),(6,10),(11,19)]                                  # 11..19 -> 11..16 after the 3 columns below are deleted
    for (t,lab,v,fm,(dark,light)),(c1,c2) in zip(cards,spans):
        for rr in range(top2,top2+4):
            ws.merge_cells(start_row=rr,start_column=c1,end_row=rr,end_column=c2)
            for j in range(c1,c2+1): fill(ws.cell(rr,j),dark if rr==top2 else light)
        a=ws.cell(top2,c1,t); a.font=F(size=12,bold=True,color='FFFFFF'); a.alignment=Alignment(horizontal='center',vertical='center')
        b_=ws.cell(top2+1,c1,lab); b_.font=F(size=10,color=dark); b_.alignment=Alignment(horizontal='center')
        n_=ws.cell(top2+2,c1,v); n_.font=F(size=24,bold=True,color=dark); n_.alignment=Alignment(horizontal='center',vertical='center'); n_.number_format=fm
        ws.row_dimensions[top2+2].height=34
    ws.cell(top2+5,1,'Every material (raw materials, packaging, consumables) in the risk order of the MRP · In transit = MIS pending · Overdue = MIS arrival date passed, not received · '
                     'Shortage risk: ALERT = short inside the transit lead time · Short later · Under 21 d (no shortage in plan) · Days incl. in transit = (stock + MIS in transit) ÷ need per day.').font=F(size=8,italic=True,color=MUTED)
    ws.cell(top2+6,1,'Stock = the stock the MRP uses (physical count rolled forward to today, else SAP) = ALERT Estimated stock today · SAP stock (MB52) / Physical count = the Stock Check columns (count as counted; no counting = not counted; back glass 333 per WSA sheet = the suppliers of that sheet) · Why Stock differs: a counted material = Count − plan usage per day × days since the count + goods received since the count (never below 0); an uncounted one = SAP · count dates, usage and receipts per material: Stock Check sheet.').font=F(size=8,italic=True,color=MUTED)
    _del_cols(ws,[12,14,15])                                      # no INVENTORY (Reorder level) / QUALITY (Status, On hold) (user, Nov 2026: "delete quality and inventory")
    _add_uom_col(ws,[x.get('uom') for x in rows])                # UOM after Category (user, Nov 2026)
    wb.move_sheet('Material Status',offset=-wb.sheetnames.index('Material Status'))
    pass                                                           # saved once at the end (post_flush)

def _copy_sheet_to(src, dst):
    """Copy one finished sheet (values, styles, merges, widths, heights) into a sheet of ANOTHER workbook."""
    from copy import copy
    for row in src.iter_rows():
        for c in row:
            n=dst.cell(c.row,c.column,c.value)
            if c.has_style:
                n.font=copy(c.font); n.fill=copy(c.fill); n.border=copy(c.border); n.alignment=copy(c.alignment); n.number_format=c.number_format
    for rng in src.merged_cells.ranges: dst.merge_cells(str(rng))
    for k,d in src.column_dimensions.items():
        if d.width: dst.column_dimensions[k].width=d.width
        if d.hidden: dst.column_dimensions[k].hidden=True
    for k,d in src.row_dimensions.items():
        if d.height: dst.row_dimensions[k].height=d.height
    dst.sheet_view.showGridLines=src.sheet_view.showGridLines; dst.sheet_view.zoomScale=src.sheet_view.zoomScale
    if src.sheet_properties.tabColor is not None: dst.sheet_properties.tabColor=copy(src.sheet_properties.tabColor)

DASH_CODES={}   # sheet title -> {row: the WSA sheet's material code(s)} (RM dashboards: Code / Item name columns)
DASH_MAP={}     # old row -> new row of the last Dashboard sorted by _wsa_dashboard_to (risk order)
def sort_wsa_dashboard(path, d0):
    """The aligned WSA file's own 'Dashboard' sorted highest risk first (user answer Nov 2026: 'Yes' - the same order as the RM Material
       Dashboard of Daily_Material_Status). XML level: the rows move with their formulas (each formula's own-row references, e.g. $A5,
       follow the row) and every unquoted 'Dashboard!B5' reference in other sheets (hidden Summary Dashboard) is moved too. Cached values
       stay (Excel recalculates on open). Returns the number of rows moved."""
    import zipfile, shutil, tempfile
    from openpyxl import Workbook as _WB
    _wsa_dashboard_to(path,_WB().active,d0)                     # works out today's values + the risk order (DASH_MAP)
    mp=dict(DASH_MAP)
    z=zipfile.ZipFile(path); wbx=z.read('xl/workbook.xml').decode(); rels=z.read('xl/_rels/workbook.xml.rels').decode()
    rid=re.search(r'<sheet[^>]*name="Dashboard"[^>]*r:id="([^"]+)"',wbx).group(1)
    tg=(re.search(r'Id="%s"[^>]*Target="([^"]+)"'%rid,rels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"'%rid,rels)).group(1)
    dpath='xl/'+tg.replace('/xl/','').lstrip('/')
    ref_rx=re.compile(r'(?<![!\w$:])(\$?[A-Z]{1,3}\$?)(\d+)(?![\d(])')
    def _fix_f(m_):
        segs=re.split(r"('[^']*'|\"[^\"]*\"|&quot;.*?&quot;)",m_.group(2))      # sheet names / text in quotes are never cell references
        f=''.join(sg if k_%2 else ref_rx.sub(lambda q: q.group(1)+str(mp.get(int(q.group(2)),int(q.group(2)))),sg) for k_,sg in enumerate(segs))
        return m_.group(1)+f+m_.group(3)
    out={}
    x=z.read(dpath).decode('utf8')
    sd=re.search(r'<sheetData>(.*)</sheetData>',x,re.S)
    parts=re.findall(r'(<row [^>]*?r="(\d+)"(?:[^>]*/>|.*?</row>))',sd.group(1),re.S)
    newp=[]
    for txt,rn in parts:
        rn=int(rn); nr=mp.get(rn,rn)
        if nr!=rn:
            txt=re.sub(r'(<row [^>]*?r=")%d(")'%rn,r'\g<1>%d\2'%nr,txt,count=1)
            txt=re.sub(r'(<c r="[A-Z]{1,3})%d(")'%rn,r'\g<1>%d\2'%nr,txt)
            txt=re.sub(r'(<f[^>]*>)([^<]*)(</f>)',_fix_f,txt)
        newp.append((nr,txt))
    newp.sort(key=lambda t:t[0])
    out[dpath]=x[:sd.start(1)]+''.join(t for _,t in newp)+x[sd.end(1):]
    def _f1(m_):                                                 # Dashboard F1 = day 1 of the date row (user rule: base date = MB52 date)
        c_=m_.group(0); b_=pd.Timestamp(d0).normalize()
        if b_<pd.Timestamp(dt.date.today()).normalize():
            return re.sub(r'<f[^>]*>(?:TODAY\(\)|DATE\(\d+,\d+,\d+\))</f>(?:<v>[^<]*</v>)?',f'<f>DATE({b_.year},{b_.month},{b_.day})</f><v>{(b_-pd.Timestamp("1899-12-30")).days}</v>',c_)
        return re.sub(r'<f>DATE\(\d+,\d+,\d+\)</f>(?:<v>[^<]*</v>)?','<f ca="1">TODAY()</f>',c_)   # an earlier base-date run: live again
    if not mp:                                                   # rows already in risk order: only F1 may need the base date / TODAY() again
        x1=re.sub(r'<c r="F1"[^>]*>.*?</c>',_f1,x,count=1)
        if x1==x: z.close(); return 0
        out[dpath]=x1
    else: out[dpath]=re.sub(r'<c r="F1"[^>]*>.*?</c>',_f1,out[dpath],count=1)
    ext_rx=re.compile(r"(?<![\w'])(Dashboard!\$?[A-Z]{1,3}\$?)(\d+)(?!\d)")
    for n in z.namelist():                                       # other sheets pointing at Dashboard cells
        if n.startswith('xl/worksheets/') and n.endswith('.xml') and n!=dpath:
            t=z.read(n).decode('utf8')
            if 'Dashboard!' not in t: continue
            t2=ext_rx.sub(lambda q: q.group(1)+str(mp.get(int(q.group(2)),int(q.group(2)))),t)
            if t2!=t: out[n]=t2
    tmp=path[:-5]+'.sorting.xlsx'
    with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zo:
        for it in z.infolist(): zo.writestr(it,out[it.filename].encode('utf8') if it.filename in out else z.read(it.filename))
    z.close()
    try:
        zipfile.ZipFile(tmp).testzip(); load_workbook(tmp,read_only=True)['Dashboard']   # read-back check before it replaces the file
        err_=_put_file(tmp,path)                                     # OneDrive / Excel lock: retry, write in place (checker finding Nov 2026)
        if err_ is not None: raise err_
    finally:
        if os.path.exists(tmp):                                      # never leave '<stem>.sorting.xlsx' behind (the next run could read it)
            try: os.remove(tmp)
            except OSError as e_: _swallow(0,e_)
    return len(mp)
CAT_COL=[False]
DASH_MIS='RM Material Dashboard - MIS'      # Daily_Material_Status: the WSA Dashboard on column V (MIS arrival dates); sheet names max 31 characters
def _wsa_dashboard_to(wsa_path, dst, d0, use_v=False):
    """'RM Material Dashboard' = the Dashboard sheet of the daily WSA file (user, Nov 2026: "copy dashboard from daily wsa to daily
       material status"). Its formulas VLOOKUP the planning sheets of the WSA file, which another workbook cannot see, and its saved results
       are from the last time Excel calculated it - so every formula is worked out here the same way: date row 1 = TODAY() + 1 per column,
       IFERROR(VLOOKUP(date, '<sheet>'!A:.., column, 0), fallback) on the planning sheet's numbers (the script's own values for today on).
       Same layout: fonts / fills / borders / number formats, column widths, frozen pane, zoom, the 4 colour rules (NA grey, < 5 red,
       < 15 green, > 15 gold). Returns the number of materials, None when the file has no Dashboard."""
    import zipfile
    from copy import copy as _cp
    from openpyxl.formatting.rule import CellIsRule, FormulaRule
    def _styles_light(archive, wb):                              # read-only load WITH fonts / fills / borders (the speed patch drops them)
        from openpyxl.utils.indexed_list import IndexedList
        node=_oxl_fromstring(archive.read('xl/styles.xml'))
        for ch in list(node):
            if ch.tag.rsplit('}',1)[-1] not in ('numFmts','cellXfs','fonts','fills','borders'): node.remove(ch)
        ss=_OxlStylesheet.from_tree(node)
        wb._borders=IndexedList(ss.borders); wb._fonts=IndexedList(ss.fonts); wb._fills=IndexedList(ss.fills)
        wb._cell_styles=ss.cell_styles; wb._number_formats=ss.number_formats; wb._date_formats=ss.date_formats
        wb._timedelta_formats=ss.timedelta_formats; wb._alignments=ss.alignments; wb._protections=ss.protections
        return wb
    _prev=_oxl_reader.apply_stylesheet; _oxl_reader.apply_stylesheet=_styles_light
    try: wf=load_workbook(wsa_path,read_only=True)
    finally: _oxl_reader.apply_stylesheet=_prev
    if 'Dashboard' not in wf.sheetnames: return None
    wv=load_workbook(wsa_path,read_only=True,data_only=True)
    src=list(wf['Dashboard'].iter_rows())
    hdr={}; tab={}
    rx=re.compile(r"^=IFERROR\(VLOOKUP\(([A-Z]+)\$1,'?(.+?)'?!\$A:\$[A-Z]+,\$A(\d+)(-1)?,0\),\"(.*?)\"\)$")
    vcol={}
    def table(nm):
        if nm not in tab:
            t_={}
            if nm in wv.sheetnames:
                for r_ in wv[nm].iter_rows(values_only=True):
                    if r_ and isinstance(r_[0],(dt.datetime,dt.date)): t_.setdefault(pd.Timestamp(r_[0]).normalize(),r_)
                    elif r_ and nm not in vcol and str(r_[0] or '').strip().startswith('Required Plant Date'):   # header row: where column V is
                        vcol[nm]=next((j+1 for j,h in enumerate(r_) if _hnorm(h).startswith('plant + port days (ppc')),None)
            tab[nm]=t_
        return tab[nm]
    nmat=0; buf={}; rsheet={}
    for row in src:
        for c in row:
            if not hasattr(c,'column') or c.column is None: continue
            v=c.value; out=v
            if c.row==1 and isinstance(v,str) and v.startswith('='):
                if v.upper()=='=TODAY()' or re.match(r'^=DATE\(\d+,\d+,\d+\)$',v.upper()): out=pd.Timestamp(d0).to_pydatetime()   # DATE() = a base-date run (F1)
                else:
                    m_=re.match(r'^=([A-Z]+)1\+(\d+)$',v)
                    out=(pd.Timestamp(hdr[m_.group(1)])+pd.Timedelta(days=int(m_.group(2)))).to_pydatetime() if m_ and m_.group(1) in hdr else None
                hdr[c.column_letter]=out
            elif isinstance(v,str) and v.startswith('='):
                m_=rx.match(v); out=None
                if m_:
                    rsheet.setdefault(c.row,m_.group(2).replace('&amp;','&'))
                    a_=src[c.row-1][0].value
                    idx=(int(a_) if isinstance(a_,(int,float)) else None)
                    if idx is not None and m_.group(4): idx-=1
                    fb=m_.group(5) or None; dd_=hdr.get(m_.group(1)); out=fb
                    if use_v and c.column>=6:                        # MIS-dates version (user, Nov 2026): every daily cell = column V 'Plant + Port Days
                        table(m_.group(2)); idx=vcol.get(m_.group(2)); dd_=hdr.get(c.column_letter); fb=fb or ''   # by header (user, Nov 2026: 9 customer columns moved V right - the sheet showed supplier names)   # (PPC In Transit)' of the cell's OWN date (the planner's row 'Frame G12R 35HS' reads R, 2 days shifted)
                        out=fb or None
                        # a sheet not in the rebuilt layout has no column V: blank, never a fixed column (user screenshot, Nov 2026: PO numbers / supplier names in the day cells)
                    if idx and dd_ is not None:
                        r_=table(m_.group(2)).get(pd.Timestamp(dd_).normalize())
                        if r_ is not None and idx-1<len(r_):
                            x_=r_[idx-1]
                            out=fb if (isinstance(x_,str) and x_.startswith('#')) else (0 if x_ is None else x_)
            buf.setdefault(c.row,[]).append((c,out))
        if row and getattr(row[0],'row',1)>1 and isinstance(getattr(row[0],'value',None),(int,float)): nmat+=1
    if use_v and not any(vcol.values()):
        raise ValueError("the WSA file read by this run was NOT rebuilt (no column 'Plant + Port Days (PPC In Transit)' on its planning sheets) - close the WSA file in Excel and run again without --no-align")
    # highest risk on top (user, Nov 2026: "All dashboard, sort from highest risk to lowest"): first day the cover (days) goes below 1 (stock-zero day),
    # then below 5, then the lowest 'Today's Morning stock'; rows with no plan last; row 1 (dates) and rows without a material stay put
    def _risk(rn):
        vs=[o for c,o in buf[rn] if c.column>=6]; nums=[v for v in vs if isinstance(v,(int,float)) and not isinstance(v,bool)]
        dv=next((o for c,o in buf[rn] if c.column==4),None)
        if not nums and not isinstance(dv,(int,float)): return (1,0,0,0)
        neg=next((i for i,v in enumerate(vs) if isinstance(v,(int,float)) and v<1),10**6)     # stock-zero day (below 1), as the BOM dashboards
        lo=next((i for i,v in enumerate(vs) if isinstance(v,(int,float)) and v<5),10**6)
        return (0,neg,lo,dv if isinstance(dv,(int,float)) else 1e9)
    data=[rn for rn in sorted(buf) if rn>1 and isinstance(next((o for c,o in buf[rn] if c.column==1),None),(int,float))]
    to={rn:rn for rn in buf}
    for new_,old_ in zip(data,sorted(data,key=_risk)): to[old_]=new_
    DASH_MAP.clear(); DASH_MAP.update({o_:n_ for o_,n_ in to.items() if o_!=n_})
    for rn,cells in buf.items():
        for c,out in cells:
            if use_v and c.column==1 and rn>1 and isinstance(out,(int,float)): out=22      # lookup column = V on the MIS sheet
            n_=dst.cell(to[rn],c.column,out)
            if c.has_style:
                n_.font=_cp(c.font); n_.fill=_cp(c.fill); n_.border=_cp(c.border); n_.alignment=_cp(c.alignment); n_.number_format=c.number_format
    # Category in column C (user, Nov 2026: "rm material dashboard add category") = item_category of the WSA sheet's material code(s)
    if not MIS_CAT and MIS_PATH_[0]: load_mis_cat(MIS_PATH_[0])
    sh_codes={}
    for g_ in (WSA_REFS[0] or []): sh_codes.setdefault(str(g_.get('sheet')),[]).extend(g_.get('codes') or [])
    h_=dst.cell(1,3,'Category'); h_._style=_cp(dst.cell(1,2)._style)
    mrp_m={str(x_['b']['mat']).strip() for x_ in (AL_ORDER_ or []) if not x_.get('sup_row')}      # the sheet's MRP material(s) only (not old / alias codes; checker finding)
    for rn in data:
        cl_=sh_codes.get(rsheet.get(rn,''),[]); cm_=[c_ for c_ in cl_ if c_ in mrp_m] or cl_
        cs_=list({str(k_).strip().lower():str(k_).strip() for k_ in reversed([item_category(c_) for c_ in cm_]) if k_ not in (None,'','-')}.values())[::-1]   # one per name, any case
        k_=dst.cell(to[rn],3,' / '.join(cs_) or '-'); k_._style=_cp(dst.cell(to[rn],2)._style); k_.alignment=Alignment(horizontal='left',vertical='center')
        DASH_CODES.setdefault(dst.title,{})[to[rn]]=list(dict.fromkeys(str(c_).strip() for c_ in cm_))     # code + item name columns of the dark design
    CAT_COL[0]=True
    try:                                                            # column widths, frozen pane, zoom, colour rules from the sheet's own XML
        z_=zipfile.ZipFile(wsa_path); wbx=z_.read('xl/workbook.xml').decode(); rels=z_.read('xl/_rels/workbook.xml.rels').decode()
        rid=re.search(r'<sheet[^>]*name="Dashboard"[^>]*r:id="([^"]+)"',wbx).group(1)
        tg=(re.search(r'Id="%s"[^>]*Target="([^"]+)"'%rid,rels) or re.search(r'Target="([^"]+)"[^>]*Id="%s"'%rid,rels)).group(1)
        x_=z_.read('xl/'+tg.replace('/xl/','').lstrip('/')).decode()
        for a_,b_,w_ in re.findall(r'<col min="(\d+)" max="(\d+)" width="([\d.]+)"',x_):
            for j_ in range(int(a_),min(int(b_),dst.max_column)+1): dst.column_dimensions[L(j_)].width=float(w_)
        for r_,h_ in re.findall(r'<row r="(\d+)"[^>]*?ht="([\d.]+)"[^>]*customHeight="1"',x_): dst.row_dimensions[int(r_)].height=float(h_)
        dst.column_dimensions['C'].width=18                       # Category column (was a thin spacer)   # row heights (rotated dates in row 1)
        if re.search(r'<sheetView[^>]*showGridLines="0"',x_): dst.sheet_view.showGridLines=False                # gridlines as in the source
        pn=re.search(r'<pane xSplit="(\d+)" ySplit="(\d+)"',x_)
        if pn: dst.freeze_panes=f"{L(int(pn.group(1))+1)}{int(pn.group(2))+1}"
        zm=re.search(r'<sheetView[^>]*zoomScale="(\d+)"',x_)
        if zm: dst.sheet_view.zoomScale=int(zm.group(1))
        sq=re.search(r'<conditionalFormatting sqref="([^"]+)"',x_)
        rng=sq.group(1) if sq else f"D2:{L(dst.max_column)}{dst.max_row}"
    except Exception as ex_:
        print(f'NOTE: Dashboard layout not fully copied ({type(ex_).__name__}: {ex_})'); rng=f"D2:{L(dst.max_column)}{dst.max_row}"
    tl=rng.split()[0].split(':')[0]
    dst.conditional_formatting.add(rng,FormulaRule(formula=[f'NOT(ISERROR(SEARCH("NA",{tl})))'],fill=PatternFill('solid',bgColor='BFBFBF'),font=Font(color='BFBFBF'),stopIfTrue=True))
    dst.conditional_formatting.add(rng,CellIsRule(operator='lessThan',formula=['5'],fill=PatternFill('solid',bgColor='FF0000'),font=Font(color='000000'),stopIfTrue=True))
    dst.conditional_formatting.add(rng,CellIsRule(operator='lessThan',formula=['15'],fill=PatternFill('solid',bgColor='00B050'),font=Font(color='006100'),stopIfTrue=True))
    dst.conditional_formatting.add(rng,CellIsRule(operator='greaterThan',formula=['15'],fill=PatternFill('solid',bgColor='BF8F00'),font=Font(color='FFFFFF')))
    return nmat
def _drop_card_rows(w, cr, n=4):
    """Remove the n-row card block that starts at row cr (+ the blank row after it): its merges and cells go, the rows below move up."""
    for m_ in list(w.merged_cells.ranges):
        if cr<=m_.min_row<=cr+n-1: w.unmerge_cells(str(m_))
    for k_ in [k_ for k_ in w._cells if cr<=k_[0]<=cr+n]: del w._cells[k_]
    mr_=w.max_row
    if mr_>cr+n:
        for m_ in [m_ for m_ in list(w.merged_cells.ranges) if m_.min_row>cr+n]:
            w.unmerge_cells(str(m_)); w.merge_cells(start_row=m_.min_row-n,start_column=m_.min_col,end_row=m_.max_row-n,end_column=m_.max_col)
        hs={r_:w.row_dimensions[r_].height for r_ in range(cr+n+1,mr_+1)}
        w.move_range(f"A{cr+n+1}:{L(w.max_column)}{mr_}",rows=-n)
        for r_ in range(cr,mr_+1): w.row_dimensions[r_].height=None
        for r_,h_ in hs.items(): w.row_dimensions[r_-n].height=h_
    else:
        for r_ in range(cr,mr_+1): w.row_dimensions[r_].height=None
def _remap_rows(w, f):
    """Move / delete whole rows of a finished sheet: f(old row) -> new row, or None = delete. Cells (value + style), merges (all rows of a
       merge must land on consecutive rows, else the merge is dropped), row heights, freeze panes, autofilter and print title rows follow."""
    import re as _re
    cells=[(k_,c_) for k_,c_ in w._cells.items()]
    merges=[(m_.min_row,m_.min_col,m_.max_row,m_.max_col) for m_ in w.merged_cells.ranges]
    hs={r_:d_.height for r_,d_ in w.row_dimensions.items() if d_.height}
    for m_ in list(w.merged_cells.ranges): w.unmerge_cells(str(m_))
    w._cells={}
    for r_ in list(w.row_dimensions.keys()): w.row_dimensions[r_].height=None
    for (r_,c_),cell in cells:
        n_=f(r_)
        if n_ is None: continue
        cell.row=n_; w._cells[(n_,c_)]=cell
    for r_,h_ in hs.items():
        n_=f(r_)
        if n_ is not None: w.row_dimensions[n_].height=h_
    for a,b,c,d in merges:
        na,nc=f(a),f(c)
        if na is not None and nc is not None and nc-na==c-a: w.merge_cells(start_row=na,start_column=b,end_row=nc,end_column=d)
    def _cell(ref):
        m_=_re.match(r'^([A-Z]+)(\d+)$',ref or '')
        if not m_: return ref
        n_=f(int(m_.group(2))); return f"{m_.group(1)}{n_}" if n_ else None
    if w.freeze_panes: w.freeze_panes=_cell(w.freeze_panes)
    if w.auto_filter.ref and ':' in w.auto_filter.ref:
        a_,b_=w.auto_filter.ref.split(':'); a2,b2=_cell(a_),_cell(b_)
        w.auto_filter.ref=f"{a2}:{b2}" if a2 and b2 else None
    pt=w.print_title_rows
    if pt:
        r1,r2=[int(x) for x in _re.findall(r'\d+',pt)][:2]
        keep=[f(x) for x in range(r1,r2+1) if f(x) is not None]
        w.print_title_rows=f"{min(keep)}:{max(keep)}" if keep else None
IT_SUP_HEAD='In transit supplier name (qty)'; IT_CODE=re.compile(r'^(JW)?[A-Z]{1,8}\d{4,}$')     # any material code (RM / PM / CS)
def _sup_short(n):
    """Supplier name without the legal form ('FLAT (HONG KONG) CO., LIMITED' -> 'FLAT (HONG KONG)')."""
    t=' '.join(str(n or '').split())
    t=re.sub(r'[,.]?\s*(CO\.?,?\s*(LTD|LIMITED)\.?|COMPANY LIMITED|LIMITED|LTD\.?|SDN\.?\s*BHD\.?|INC\.?|LLC|CORP(ORATION)?\.?)\s*$','',t,flags=re.I)
    return t.strip(' ,.') or '?'
def it_sup_text(codes, sheet=None, own=False, names=False):
    """In transit by supplier of a Daily_Material_Status row (user, Nov 2026: "Add in transit supplier name in all sheet of daily material
       status"): 'SUPPLIER qty; ...' (largest first). A material on 2+ WSA sheets (333) with its sheet = the PPC In_Transit lines of the
       suppliers shown in that sheet's column U (= the sheet's 'In transit'); else the MIS In Transit shipments of the code (substitute
       group row: every member, in this code's unit; any other row: the code only) = the 'In transit' column."""
    by={}
    if sheet:
        g=next((g_ for g_ in (WSA_REFS[0] or []) if sheet in (g_.get('sheet'),g_.get('label'))),None)
        nm=(g or {}).get('ppcsup_names') or set()
        for x in ((CAP[0] or {}).get('ppc') or {}).get('rows') or []:
            sn=' '.join(str(x.get('supplier') or '').split())
            if x.get('mat') in codes and sn in nm and x.get('qty_mrp'): by[sn]=by.get(sn,0.0)+float(x['qty_mrp'])
    else:
        for c in codes:
            for s_ in (SHIPS_.get(c,[]) if own else grp_ships(c,SHIPS_)):
                if s_.get('qty') and s_['qty']>0: by[s_.get('sup') or '?']=by.get(s_.get('sup') or '?',0.0)+float(s_['qty'])
    agg={}
    for n,q in by.items(): agg[_sup_short(n)]=agg.get(_sup_short(n),0.0)+q
    if names: return [n for n,q in sorted(agg.items(),key=lambda t:-t[1]) if q>0.5]      # supplier names only (daily email)
    return '; '.join(f"{n} {q:,.0f}" for n,q in sorted(agg.items(),key=lambda t:-t[1]) if q>0.5) or 'none'
def _email_suppliers(r, bom_sup):
    """Supplier name(s) of an ALERT row for the daily email (user, Nov 2026: "put supplier names in the critical table"): the suppliers
       of its shipments in transit (333 per WSA sheet = that sheet's suppliers; substitute group = every member), else the supplier on
       your PPC BOM (the one you order from), else the stock suppliers of the MRP supplier rows (glass / materials with 2+ suppliers),
       else the Shipper typed on your WSA delivery rows; 'no supplier on file' when none is known."""
    codes=[m_['code'] for m_ in (r.get('mem') or [])] or [r['code']]
    try: n_=it_sup_text([r['code']],sheet=r['where'],names=True) if r.get('where') else it_sup_text(codes,own=not r.get('mem'),names=True)
    except Exception: n_=[]
    if not n_:
        al_=[str(a_).strip() for a_,m_ in (BOM_ALIAS[0] or {}).items() if str(m_).strip() in codes]     # other roll length on your BOM (tape 018 -> 010)
        n_=[_sup_short(v_) for c_ in codes+al_ for v_ in sorted(bom_sup.get(str(c_).upper(),()))]
    if not n_:
        for x_ in (AL_ORDER_ or []):
            b_=x_['b']
            if x_.get('sup_row') or str(b_.get('mat')) not in codes: continue
            n_+=[_sup_short(sp_.get('name')) for sp_ in (b_.get('sup') or []) if str(sp_.get('name') or '').upper()!='UNASSIGNED' and float(sp_.get('stock') or 0)>0]
    if not n_: n_=[f"{v_} (WSA shipper)" for c_ in codes for v_ in dict.fromkeys(WSA_SHIPPER.get(str(c_).upper(),[]))]
    return ' / '.join(dict.fromkeys(n_)) or 'no supplier on file'
def write_daily_email(rows, d0, out_dir, attach=None):
    """Daily email to top management (user, Nov 2026: "make this email come out of the script every day"), written from the ALERT rows
       (the same rows as the ALERT of MRP_Daily / Inbound / Daily_Material_Status): subject with the counts, summary per ALERT class,
       CRITICAL table (stock-out inside the transit lead time) with supplier, ETD alert for FUTURE ETDs only (arrives after stock-out /
       ships within ETD_SOON days, per material), Needs attention, Notes (substitutes, the attachment). No Order this week / Overdue /
       passed-ETD tables (user, Nov 2026: "I dont want 3 last tables"). Two files next to MRP_Daily:
       Daily_Email_<date>.eml = an Outlook draft (double-click, add the To line, check, Send; Daily_Material_Status attached) and
       Daily_Email_<date>.html = the same body (open in a browser, copy, paste). Facts only - no advice text."""
    import html as _h
    from email.message import EmailMessage
    from email.utils import formatdate
    if not rows: return None
    E=lambda v: _h.escape(str(v))
    D=lambda d_: f"{pd.Timestamp(d_):%d-%b}"
    def cover(v):
        if v is None: return '-'
        if v<0.05: return '0 days'
        return (f"{v:.1f} days" if v<1 else (f"{v:,.0f} day" if round(v)==1 else f"{v:,.0f} days"))
    def runs_out(r):
        out_=[]
        for ln_ in str(r.get('per') or '').split('\n')[:3]:
            ln_=ln_.replace('⚠','').strip()
            if not ln_ or ln_.startswith('...'): continue
            m_=re.match(r'^(\d{2}-\w{3}) → beyond',ln_)
            out_.append(f"from {m_.group(1)}" if m_ else re.sub(r'\s*\(\d+ d\)$','',ln_).replace(' → ','–'))
        return ', '.join(out_) or '-'
    bom_sup={}
    try:
        for x_ in wb_ro(PPC_SRC[0])['BOM'].iter_rows(min_row=2,values_only=True):
            if x_ and len(x_)>6 and x_[4] and x_[6]: bom_sup.setdefault(str(x_[4]).strip().upper(),set()).add(' '.join(str(x_[6]).split()))
    except Exception as _e_: pass                                 # no PPC file: in-transit / stock suppliers only
    def mat(r): return str(r['name'])+(f" ({str(r['where']).replace('Glass G12R Topcon ','')})" if r.get('where') else '')
    def nxt(r):
        nx_=r.get('nx'); t_=f"{D(nx_[0])}: {nx_[1]:,.0f} {r.get('uom') or ''}".strip() if nx_ else 'nothing in transit'
        return t_
    def note(r):
        w_=r.get('why') or {}; n_=[]
        if not (r.get('tr_tot') or 0): n_.append(f"nothing in transit (transit lead time {r.get('lt') or 0} days)")
        elif r.get('first') is not None and r.get('nx') and pd.Timestamp(r['nx'][0])>pd.Timestamp(r['first']): n_.append('next arrival is after the stock-out')
        elif r.get('first') is not None: n_.append(f"{r['tr_tot']:,.0f} {r.get('uom') or ''} in transit is not enough or arrives too late".replace('  ',' '))
        if isinstance(r.get('ovd'),(int,float)) and r['ovd']>0: n_.append(f"{r['ovd']:,.0f} overdue (arrival date passed)")
        if r.get('subnote'): n_.append(r['subnote'])
        return '; '.join(n_) or '-'
    cnt=Counter(r['cls'] for r in rows); crit=[r for r in rows if r['cls']==0]
    sup_={id(r):_email_suppliers(r,bom_sup) for r in crit}
    subs=[]
    for r in rows:                                           # substitute members are separate rows (user, Nov 2026); the pooled date as a note
        g_=r.get('grp')
        if g_ and float(r.get('stock') or 0)<=0:
            subs.append(f"{r['name']} has no stock today; with its approved substitute {' / '.join(g_['others'])} together the stock "
                        +(f"runs out {D(g_['first'])}." if g_.get('first') is not None else 'lasts through the end of the plan.'))
    n_tot=len(rows)
    subj=f"Daily Material Status – {pd.Timestamp(d0):%d-%b-%Y} | {cnt.get(0,0)} critical, {cnt.get(1,0)} at risk"
    LBL=[('🔴',f"{cnt.get(0,0)} critical",'stock runs out before a new order could arrive (inside the transit lead time)'),
         ('🟠',f"{cnt.get(1,0)} at risk",'stock runs out later in the plan; an order is needed in time'),
         ('🟡',f"{cnt.get(2,0)+cnt.get(3,0)} watch",f"below {REPORT_DAYS} days of cover now or later, no stock-out in the plan"),
         ('🟢',f"{cnt.get(4,0)+cnt.get(5,0)} OK / no demand",'')]
    TH='style="background:#1F3864;color:#fff;padding:5px 8px;text-align:left;font-size:12px;border:1px solid #1F3864"'
    TD='style="padding:5px 8px;border:1px solid #D0D5DD;font-size:12px;vertical-align:top"'
    TDR='style="padding:5px 8px;border:1px solid #D0D5DD;font-size:12px;vertical-align:top;color:#C0392B;font-weight:bold"'
    def table(hd,body):
        return ('<table style="border-collapse:collapse;margin:6px 0 14px 0">'+'<tr>'+''.join(f'<th {TH}>{E(h_)}</th>' for h_ in hd)+'</tr>'
                +''.join('<tr>'+''.join(f'<td {TDR if j_==hl else TD}>{E(v_)}</td>' for j_,v_ in enumerate(vals))+'</tr>' for vals,hl in body)+'</table>')   # hl = the column in red
    H=['<div style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#1F2D3D">','<p>Dear all,</p>',
       f"<p>Here is today's raw and packaging material status, based on SAP MB52 + the physical count, the module plan, and MIS in-transit arrival dates (base date {pd.Timestamp(d0):%d-%b-%Y}).</p>",
       f'<p style="margin-bottom:4px"><b>Summary ({n_tot} materials)</b></p><ul style="margin-top:0">'
       +''.join(f"<li>{i_} <b>{E(a_)}</b>"+(f": {E(t_)}" if t_ else '')+'</li>' for i_,a_,t_ in LBL)+'</ul>']
    T=[subj,'','Dear all,','',f"Here is today's raw and packaging material status (base date {pd.Timestamp(d0):%d-%b-%Y}).",'',f"Summary ({n_tot} materials)"]
    T+=[f"- {a_}"+(f": {t_}" if t_ else '') for i_,a_,t_ in LBL]
    if crit:
        hd=['Material','Supplier','Cover today','Runs out','Next arrival','Situation']
        body=[([mat(r),sup_[id(r)],cover(r.get('days')),runs_out(r),nxt(r),note(r)],3) for r in crit]
        H.append('<p style="margin-bottom:2px"><b>🔴 Critical – action needed now</b></p>'+table(hd,body))
        T+=['','CRITICAL - action needed now']+[f"- {b_[0][0]} | {b_[0][1]} | cover {b_[0][2]} | runs out {b_[0][3]} | next {b_[0][4]} | {b_[0][5]}" for b_ in body]
    # no Order this week / Overdue / passed-ETD tables (user, Nov 2026: "I dont want 3 last tables" - the ETD alert sheet keeps that detail)
    # ETD alert, FUTURE ETDs only (user, Nov 2026: "I want etd alert in email also, just focus on future etd not passed etd"): WSA deliveries
    # not on MIS In Transit with an ETD still to come and an alert - ARRIVES AFTER STOCK-OUT or SHIPS WITHIN ETD_SOON DAYS (ON PLAN = no
    # alert); one line per material and alert, the same rule / rows as the 'ETD alert' sheet
    fut_={}
    for e_ in ETD_LATE:
        if e_['passed']: continue
        code_,name_,cls_,first_,uom_=_etd_alert(e_); fl_=_etd_flag(e_,first_)
        if fl_ not in (0,2): continue
        k_=(fl_,code_,e_['sheet'] if code_=='RMGLS00000333' else '')
        t_=fut_.setdefault(k_,dict(fl=fl_,name=name_+(f" ({e_['sheet'].replace('Glass G12R Topcon ','')})" if k_[2] else ''),sup=set(),n=0,qty=0.0,
                                   uom=uom_ or '',etd0=e_['etd'],pl0=e_['plant'],first=first_))
        t_['n']+=1; t_['qty']+=float(e_['qty'] or 0); t_['etd0']=min(t_['etd0'],e_['etd']); t_['pl0']=min(t_['pl0'],e_['plant'])
        ar_=_etd_arow(e_)
        if ar_ is not None: t_['sup'].add(_email_suppliers(ar_,bom_sup))
        elif e_.get('shipper') and not str(e_['shipper']).replace('.0','').isdigit(): t_['sup'].add(str(e_['shipper']).strip())
    futl=sorted(fut_.values(),key=lambda t_:(t_['fl'],t_['first'] if (t_['fl']==0 and t_['first'] is not None) else t_['etd0'],t_['etd0'],t_['name']))
    if futl:
        FL_={0:'🔴 Arrives after stock-out',2:f'🟡 Ships within {ETD_SOON} days'}
        body=[([FL_[t_['fl']],t_['name'],' / '.join(sorted(t_['sup'])) or '-',str(t_['n']),f"{t_['qty']:,.0f} {t_['uom']}".strip(),D(t_['etd0']),D(t_['pl0']),
                (D(t_['first']) if t_['first'] is not None else 'no stock-out')],(7 if t_['fl']==0 else 5)) for t_ in futl]
        H.append(f'<p style="margin-bottom:2px"><b>⏱ ETD alert – next shipments not on MIS In Transit yet ({sum(t_["n"] for t_ in futl)} WSA deliveries, ETD still to come)</b></p>'
                 +table(['Alert','Material','Supplier','Deliveries','Quantity','Next ETD','WSA plant date','Runs out'],body))
        T+=['',f"ETD ALERT - NEXT SHIPMENTS NOT ON MIS IN TRANSIT YET ({sum(t_['n'] for t_ in futl)} WSA deliveries, ETD still to come)"]
        T+=[f"- {b_[0][0][2:]}: {b_[0][1]} ({b_[0][2]}): {b_[0][3]} deliver{'y' if b_[0][3]=='1' else 'ies'}, {b_[0][4]}, next ETD {b_[0][5]}, plant date {b_[0][6]}, runs out {b_[0][7]}" for b_ in body]
    nopo=[mat(r) for r in crit if not (r.get('tr_tot') or 0)]
    late_=[mat(r) for r in crit if (r.get('tr_tot') or 0)]
    att=[]
    if nopo: att.append(f"Nothing in transit: {', '.join(nopo)}")
    if late_: att.append(f"In transit, but the stock runs out first or the quantity is not enough: {', '.join(late_)}")
    if att:
        H.append('<p style="margin-bottom:4px"><b>Needs attention</b></p><ul style="margin-top:0">'+''.join(f'<li>{E(a_)}</li>' for a_ in att)+'</ul>')
        T+=['','NEEDS ATTENTION']+[f"- {a_}" for a_ in att]
    notes=subs+([f"Full details: the attached {os.path.basename(attach)} (BOM Material Dashboard – Weekly = every material in risk order)."] if attach else [])
    if notes:
        H.append('<p style="margin-bottom:4px"><b>Notes</b></p><ul style="margin-top:0">'+''.join(f'<li>{E(n_)}</li>' for n_ in notes)+'</ul>')
        T+=['','Notes']+[f"- {n_}" for n_ in notes]
    H.append('<p>Best regards,<br>[Your name]<br>Production Planning – Raw Materials</p></div>')
    T+=['','Best regards,','[Your name]','Production Planning – Raw Materials']
    html_='\n'.join(H)
    stem=os.path.join(out_dir,f"Daily_Email_{dt.date.today():%Y-%m-%d}")
    with open(stem+'.html','w',encoding='utf-8') as f_:
        f_.write(f'<!doctype html><html><head><meta charset="utf-8"><title>{E(subj)}</title></head><body>{html_}</body></html>')
    msg=EmailMessage(); msg['Subject']=subj; msg['To']=''; msg['Date']=formatdate(localtime=True); msg['X-Unsent']='1'   # X-Unsent: Outlook opens it as a draft
    msg.set_content('\n'.join(T[2:])); msg.add_alternative(f'<html><body>{html_}</body></html>',subtype='html')
    if attach and os.path.exists(attach):
        with open(attach,'rb') as f_: msg.add_attachment(f_.read(),maintype='application',subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',filename=os.path.basename(attach))
    with open(stem+'.eml','wb') as f_: f_.write(bytes(msg))
    for n_ in ([] if globals().get('KEEP_OLD') else sorted(os.listdir(out_dir))):     # older daily emails go, like MRP_Daily
        if re.fullmatch(r'Daily_Email_\d{4}-\d{2}-\d{2}\.(eml|html)',n_) and not n_.startswith(os.path.basename(stem)):
            try: os.remove(os.path.join(out_dir,n_))
            except OSError: pass
    print(f"DAILY EMAIL: {stem}.eml (double-click = Outlook draft, Daily_Material_Status attached; add the To line and send) + .html - "
          f"subject '{subj}'; {len(crit)} critical; ETD alert (future ETDs) {sum(t_['n'] for t_ in futl)} deliveries on {len(futl)} line(s)",flush=True)
    return stem
def _add_it_sup_col(w, hr, at, rows_, code_col, text_col, sheet_col=None, plain=False, dark=False):
    """Insert the 'In transit supplier name (qty)' column before column `at` (header row hr, data rows rows_); freeze pane and
       autofilter follow. Codes: the code cell (several joined by ' / ' or ', '); the WSA sheet: sheet_col, else '(WSA sheet: X)' or a
       '(... Back)' label of the text cell - only for a material on 2+ WSA sheets."""
    from copy import copy as _cp
    from openpyxl.utils import column_index_from_string as _CI
    shared={}
    for g_ in WSA_REFS[0] or []:
        for c_ in g_.get('codes') or []: shared.setdefault(c_,[]).append(g_.get('sheet'))
    shared={c_:v_ for c_,v_ in shared.items() if len(v_)>1}
    fz=w.freeze_panes; flt=w.auto_filter.ref
    _ins_cols(w,at,1)
    h=w.cell(hr,at,IT_SUP_HEAD); h._style=_cp(w.cell(hr,at-1)._style); h.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    if hr>1 and w.cell(hr-1,at-1).has_style and not w.cell(hr-1,at).has_style: w.cell(hr-1,at)._style=_cp(w.cell(hr-1,at-1)._style)
    n_=0
    for r_ in rows_:
        codes=[c_ for c_ in re.split(r'\s*(?:/|,|\+|\n)\s*',str(w.cell(r_,code_col).value or '')) if IT_CODE.match(c_.strip())]
        if not codes:                                              # a Count row without a material code
            c=w.cell(r_,at,'-'); c._style=_cp(w.cell(r_,text_col if plain else at-1)._style); c.number_format='General'; continue
        tx=str(w.cell(r_,text_col).value or ''); sh=None
        if any(c_ in shared for c_ in codes):
            cand=[s_ for c_ in codes for s_ in shared.get(c_,[])]
            if sheet_col: sh=next((s_ for s_ in cand if s_==' '.join(str(w.cell(r_,sheet_col).value or '').split())),None)
            m_=re.search(r'\(WSA sheet:\s*([^)]+)\)',tx)
            if sh is None and m_: sh=next((s_ for s_ in cand if s_==m_.group(1).strip()),None)
            m_=re.search(r'\(([^()]+)\)\s*$',tx)
            if sh is None and m_: sh=next((s_ for s_ in cand if s_.endswith(' '+m_.group(1).strip())),None)
        v_=it_sup_text(codes,sheet=sh,own='group' not in tx.lower())     # a substitute group row (ALERT) = every member; any other row = its own code
        c=w.cell(r_,at,v_); c._style=_cp(w.cell(r_,text_col if plain else at-1)._style); c.number_format='General'   # dashboards: plain text style (the cell left of it is a heat cell)
        if plain: c.font=Font(name='Calibri',size=8,color=(c.font.color.rgb if (c.font and c.font.color is not None and isinstance(c.font.color.rgb,str)) else DASH_TXT))
        c.alignment=Alignment(horizontal='left',vertical='center',wrap_text=True); n_+=1
        ln_=max(1,-(-len(v_)//38))
        if ln_>1: w.row_dimensions[r_].height=max(w.row_dimensions[r_].height or 15,min(13*ln_+4,80))
    if plain or dark:                                              # dashboards: no white hole on the black page (title / band / No Plan / legend rows; checker finding)
        for r_ in range(1,w.max_row+1):
            c=w.cell(r_,at); o=w.cell(r_,text_col)                 # the item-name cell: page colour, or the band colour on a band row
            if c.value is None and not c.fill.fill_type and o.fill.fill_type: c.fill=_cp(o.fill)
    w.column_dimensions[L(at)].width=40
    if fz:
        cf_,rf_=re.match(r'([A-Z]+)(\d+)',str(fz)).groups()
        if _CI(cf_)>=at: w.freeze_panes=f"{L(_CI(cf_)+1)}{rf_}"
    if flt:
        a_,b_=flt.split(':'); cb_,rb_=re.match(r'([A-Z]+)(\d+)',b_).groups()
        if _CI(cb_)>=at-1: w.auto_filter.ref=f"{a_}:{L(_CI(cb_)+1)}{rb_}"
    return n_
def make_daily_material_status(out_path, src_wb, AL_ORDER, bom, cnt_hdr, cnt_rows, d0, mis_path, ships=None):
    """Daily_Material_Status_<date>.xlsx (user, Nov 2026) - 2 sheets: 'Materials Status' = (1) + (2) in ONE table (user, Nov 2026: "combine
       direct materials and other material then rename to Materials Status") + 'ALERT' (copy of the MRP_Daily ALERT, without the 4 class cards).
       Columns F 'Item Category' (PPC Count column I 'Item') and H 'Consumable Unit (Production Confirmed)' (PPC PM_CS_Usage) added.
       1) Direct Material   = the 'Material Status' sheet of MRP_Daily (materials of your BOM), copied as it is
       2) Other Materials (was 'Indirect Material', renamed Nov 2026) = materials on the PPC 'Count' sheet that are NOT in the BOM: SAP stock (MB52),
          physical count, in transit; Need / day = PPC Count 'Material Req Qty per Day' (user, Nov 2026; rows of one material added up) ->
          days of supply, days incl. in transit, short from (stock + MIS arrivals - need per day, day by day), shortage risk.
       Returns the file path (None when there is nothing to write)."""
    from openpyxl import Workbook
    from openpyxl.styles import Border, Side
    if 'Material Status' not in src_wb.sheetnames: return None
    F=lambda **k: Font(name='Calibri',**k); NAVY='14213D'; INK='1F2D3D'; MUTED='8A949E'; LINE='E3E8EE'; STORE=('2E86C1','D6EAF8'); QUAL=('17A589','D1F2EB')
    thin=Side(style='thin',color=LINE); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    out_file=os.path.join(os.path.dirname(os.path.abspath(out_path)),f'Daily_Material_Status_{dt.date.today():%Y-%m-%d}.xlsx')
    # ---- which codes are direct (BOM) ----
    direct={str(d.get('Component') or '').strip() for d in (bom or []) if isinstance(d,dict)}
    direct|={str(x['b']['mat']).strip() for x in AL_ORDER if not x.get('sup_row')}
    al_=BOM_ALIAS[0] or {}; also={c for k,v in al_.items() for c in (str(k).strip(),str(v).strip())}      # alias / substitute codes of a BOM material are inside its direct row
    direct|=also
    # ---- the PPC Count sheet: one entry per material code, in the order of the sheet ----
    hl=[' '.join(str(h or '').split()).lower() for h in (cnt_hdr or [])]
    ix=lambda *ks: next((hl.index(k) for k in ks if k in hl),None)
    i_m,i_d,i_u,i_s,i_i=ix('material'),ix('material description'),ix('base unit of measure','stock uom'),ix('storage location'),ix('item')
    i_q=ix('material req qty per day','material req qty / day','req qty per day')          # need per day of a material that is not in the BOM
    g=lambda r,i: (r[i] if (i is not None and i<len(r)) else None)
    ent={}
    for r in cnt_rows or []:
        if not r or not g(r,i_m): continue
        c=' '.join(str(r[i_m]).split()).strip(); e=ent.setdefault(c,dict(desc=None,uom=None,loc=None,item=None,need=None))
        q_=g(r,i_q)
        if isinstance(q_,(int,float)) and not isinstance(q_,bool): e['need']=(e['need'] or 0.0)+float(q_)   # several count rows: added up
        for k,i in (('desc',i_d),('uom',i_u),('loc',i_s),('item',i_i)):
            v=g(r,i)
            if e[k] is None and v not in (None,'') and not (isinstance(v,str) and v.startswith('#')): e[k]=' '.join(str(v).split())
    pc=read_phys_counts(SETTINGS_SRC[0]) if SETTINGS_SRC and SETTINGS_SRC[0] else {}      # counted materials: qty, unit, count date
    sap={}
    try:
        mis=wb_ro(mis_path)
        for r in list(mis['MB52'].iter_rows(values_only=True))[1:]:
            if r and len(r)>14 and r[2] and isinstance(r[14],(int,float)): sap[str(r[2]).strip()]=sap.get(str(r[2]).strip(),0.0)+r[14]
    except Exception as ex_: print(f'NOTE: MB52 not read for the Other Materials sheet ({type(ex_).__name__}: {ex_})')
    ind=[(c,e) for c,e in ent.items() if c not in direct]
    skipped=[c for c in ent if c in also and c not in {str(d.get('Component') or '').strip() for d in (bom or []) if isinstance(d,dict)} and c not in {str(x['b']['mat']).strip() for x in AL_ORDER}]
    # ---- workbook ----
    wb=Workbook(); ws1=wb.active; ws1.title='Direct Material'
    _copy_sheet_to(src_wb['Material Status'],ws1)
    # no STORE / WAREHOUSE / PRODUCTION cards in the daily file (user, Nov 2026: "No need from daily material status"; MRP_Daily keeps them)
    cr=next((c.row for c in ws1['A'] if isinstance(c.value,str) and c.value.startswith('🏪')),None)
    if cr: _drop_card_rows(ws1,cr)
    ws=wb.create_sheet('Other Materials'); ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor='14213D'
    WH=('229954','D5F5E3'); INV=('E67E22','FDEBD0'); PROD=('6C3483','E8DAEF')
    # SAME 19 columns, widths, bands, header colours and cell styles as the Direct Material sheet (Material Status)
    for j,w_ in enumerate([5,16,18,14,17,11,11,11,10,10,11,11,11,11,10,11,9,10,12],1): ws.column_dimensions[L(j)].width=w_
    for t,c1,c2,col in (('',1,5,NAVY),('STORE · stock',6,8,STORE[0]),('WAREHOUSE · supply',9,11,WH[0]),('INVENTORY',12,12,INV[0]),('PRODUCTION',13,13,PROD[0]),('QUALITY',14,15,QUAL[0]),('AGING ANALYSIS',16,19,NAVY)):
        if c2>c1: ws.merge_cells(start_row=1,start_column=c1,end_row=1,end_column=c2)
        c=ws.cell(1,c1,t); c.font=F(size=10,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='center',vertical='center')
        for j in range(c1,c2+1): ws.cell(1,j).fill=PatternFill('solid',fgColor=col)
    ws.row_dimensions[1].height=20
    H=['Sr.','Item code','Item name','','Category','Stock','SAP stock (MB52)','Physical count','In transit','Next arrival','Overdue in transit','Reorder level','Need / day','Status','On hold','Short from','Days of supply','Shortage risk','Days incl. in transit']
    hc=[NAVY]*5+[STORE[0]]*3+[WH[0]]*3+[INV[0]]+[PROD[0]]+[QUAL[0]]*2+[NAVY]*4
    for j,(h,col) in enumerate(zip(H,hc),1):
        c=ws.cell(2,j,h); c.font=F(size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=col); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=bd
    ws.merge_cells('C2:D2'); ws.row_dimensions[2].height=30
    CH={'Out of Stock':('FADBD8','C0392B'),'No demand':('F2F4F6','8A949E'),'Short':('FDEBD0','B9770E'),'Under 21 d':('FCF3CF','9A7D0A'),'OK':('D5F5E3','1E8449')}
    HZ_=180; nneed=0
    hold=mb52_hold(mis_path); fmts=[None,None,None,None,None,'#,##0','#,##0','#,##0','#,##0','dd-mmm','#,##0','#,##0','#,##0.0',None,'#,##0','dd-mmm-yy','0',None,'0']
    r=3; ncount=0; nsap=0
    for i_,(c,e) in enumerate(ind,1):
        p=pc.get(c); cq=p['base'][0] if p else None; sv=sap.get(c)
        if cq is not None: ncount+=1
        if sv is not None: nsap+=1
        stock=cq if cq is not None else sv                                     # a physical count replaces the SAP stock (no demand: nothing to roll forward)
        sh_=(ships or {}).get(c,[]); it=sum(s_['qty'] for s_ in sh_)
        nxt=min((s_['raw'] for s_ in sh_ if s_.get('raw') is not None and s_['raw']>=d0),default=None)
        ovd=sum(s_['qty'] for s_ in sh_ if s_.get('raw') is not None and s_['raw']<d0)
        hq=float(hold.get(c) or 0)
        status='Out of Stock' if (stock is not None and stock<=0) else 'No demand'
        need=e.get('need') if (e.get('need') or 0)>0 else None              # PPC Count 'Material Req Qty per Day'
        if need and cq is not None and p.get('date') is not None:          # counted + a need: rolled forward like Stock Check (checker finding)
            dc_=max(0,(d0-pd.Timestamp(p['date']).normalize()).days)
            if dc_>0: stock=max(0.0,cq-need*dc_+gr_between(c,pd.Timestamp(p['date']).normalize(),d0))
        kz=None; dos=dit=None; risk='No demand'
        if need:
            nneed+=1; st0=float(stock or 0.0); dos=max(0.0,st0)/need; dit=max(0.0,st0+it)/need
            arr=[0.0]*HZ_
            for s_ in sh_:
                if s_.get('raw') is not None:
                    k_=max(0,(s_['raw']-d0).days)
                    if k_<HZ_: arr[k_]+=s_['qty']
            cur=st0
            for k_ in range(HZ_):                                           # stock + arrivals - need per day: first day below zero
                cur+=arr[k_]-need
                if cur<0: kz=k_; break
            risk=('Short' if kz is not None else ('Under 21 d' if dit<REPORT_DAYS else 'OK'))
        nm=e['desc'] or ''
        vals=[i_,c,nm,None,((item_category(c) if re.search(r'\d',c) else None) or '-') if (item_category(c) if re.search(r'\d',c) else None) not in (None,'','-') else (e['item'] or '-'),(round(stock) if stock is not None else '-'),(round(sv) if sv is not None else '-'),(round(cq) if cq is not None else 'no counting'),
              round(it),(nxt.to_pydatetime() if nxt is not None else 'none'),(round(ovd) if ovd else '-'),'-',(round(need,1) if need else '-'),status,(round(hq) if hq else '-'),
              ((d0+pd.Timedelta(days=kz)).to_pydatetime() if kz is not None else 'no shortage'),(round(dos) if dos is not None else '-'),
              (f"Short {(d0+pd.Timedelta(days=kz)):%d-%b}" if kz is not None else risk),(round(dit) if dit is not None else '-')]
        for j,(v,f_) in enumerate(zip(vals,fmts),1):
            cc=ws.cell(r,j,v); cc.font=F(size=10,color=INK); cc.border=Border(bottom=thin)
            cc.alignment=Alignment(horizontal=('left' if j in (2,3,5) else 'center'),vertical='center')
            if f_: cc.number_format=f_
            if j in (6,7,8): cc.fill=PatternFill('solid',fgColor='F4F9FD')
            if j in (9,10,11): cc.fill=PatternFill('solid',fgColor='F4FBF7')
            if j==12: cc.fill=PatternFill('solid',fgColor='FEF8F1')
            if j==13: cc.fill=PatternFill('solid',fgColor='F8F4FB')
        ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
        if kz is not None: ws.cell(r,16).font=F(size=10,bold=True,color='C0392B')
        for j,key in ((14,status),(18,risk)):
            bg,fg=CH[key]; ws.cell(r,j).fill=PatternFill('solid',fgColor=bg); ws.cell(r,j).font=F(size=10,bold=True,color=fg)
        if hq: ws.cell(r,15).font=F(size=10,bold=True,color='B9770E')
        ws.row_dimensions[r].height=19; r+=1
    _del_cols(ws,[12,14,15])                                      # same columns as Direct Material (no INVENTORY / QUALITY, Nov 2026)
    _add_uom_col(ws,[e.get('uom') for _,e in ind])                # UOM after Category = the PPC Count Base Unit of Measure (user, Nov 2026)
    ws.cell(r+1,1,f'Other materials = on the PPC Count sheet but not in the BOM ({len(ind)} materials; {ncount} counted, {nsap} with MB52 stock, {nneed} with a need per day) · same columns as the BOM materials above · Stock = physical count if counted (with a need per day: rolled forward to today = count − need × days since the count + received since the count, like Stock Check), else SAP (MB52) stock · '
                  'SAP stock (MB52) = MB52 stock · Physical count = PPC physical count (no counting = not counted) · In transit / Overdue = MIS pending shipments · '
                  f'Need / day = PPC Count column Material Req Qty per Day (rows of one material added up) · Days of supply = stock ÷ need · Days incl. in transit = (stock + in transit) ÷ need · '
                  f'Short from = first day stock + MIS arrivals − need per day < 0 (next {HZ_} days) · Shortage risk: Short dd-mmm / Under 21 d (incl. in transit) / OK / No demand (no need per day on the Count sheet).').font=F(size=8,italic=True,color=MUTED)
    # ---- Item Category (PPC Count column I 'Item') after Category + Consumable Unit (PPC PM_CS_Usage) after UOM (user, Nov 2026) ----
    from copy import copy as _cp
    cons={}
    try:
        _r=list(wb_ro(PPC_SRC[0])['PM_CS_Usage'].iter_rows(values_only=True)) if PPC_SRC and PPC_SRC[0] else []
        _h=next((i for i,x in enumerate(_r[:10]) if x and 'code' in [' '.join(str(v or '').split()).lower() for v in x]),None)
        if _h is not None:
            _hl=[' '.join(str(v or '').split()).lower() for v in _r[_h]]
            _ic=_hl.index('code'); _iu=next((j for j,h in enumerate(_hl) if h.startswith('consumable unit')),None)
            for x in _r[_h+1:]:
                if not x or _ic>=len(x) or not x[_ic] or _iu is None or _iu>=len(x): continue
                v_=x[_iu]
                if v_ not in (None,'','-',0) and not (isinstance(v_,str) and v_.startswith('#')): cons.setdefault(' '.join(str(x[_ic]).split()),' '.join(str(v_).split()))
    except Exception as ex_: print(f'NOTE: PPC PM_CS_Usage (Consumable Unit) not read ({type(ex_).__name__}: {ex_})')
    item_of={c_:e_.get('item') for c_,e_ in ent.items()}
    def _data_rows(w):
        out=[]; r_=3
        while isinstance(w.cell(r_,1).value,int) and w.cell(r_,2).value: out.append(r_); r_+=1
        return out
    def _add_col(w, at, head, vals_by_row, band_to):
        _ins_cols(w,at,1)
        for m_ in list(w.merged_cells.ranges):
            if m_.min_row==1 and m_.max_row==1 and m_.min_col==1 and m_.max_col<band_to: w.unmerge_cells(str(m_)); w.merge_cells(start_row=1,start_column=1,end_row=1,end_column=band_to)
        w.cell(1,at)._style=_cp(w.cell(1,1)._style)
        h=w.cell(2,at,head); h._style=_cp(w.cell(2,at-1)._style)
        for r_,v_ in vals_by_row.items():
            c=w.cell(r_,at,v_ if v_ not in (None,'') else '-'); c._style=_cp(w.cell(r_,at-1)._style); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        w.column_dimensions[L(at)].width=14
    for w in (ws1,ws):
        drs=_data_rows(w)
        _add_col(w,6,'Item Category',{r_:item_of.get(' '.join(str(w.cell(r_,2).value).split())) for r_ in drs},6)
        _add_col(w,8,'Consumable Unit (Production Confirmed)',{r_:cons.get(' '.join(str(w.cell(r_,2).value).split())) for r_ in drs},8)
    # ---- ONE sheet 'Materials Status' = BOM materials, then the other Count materials (user, Nov 2026) ----
    ms=wb.create_sheet('Materials Status',0); ms.sheet_view.showGridLines=False; ms.sheet_properties.tabColor='14213D'
    def _copy_rows(src,r1,r2,off):
        for rr in range(r1,r2+1):
            for c in src[rr]:
                if c.value is None and not c.has_style: continue
                n_=ms.cell(rr+off,c.column); n_._value=c._value; n_.data_type=c.data_type; n_._style=_cp(c._style)
            if src.row_dimensions[rr].height: ms.row_dimensions[rr+off].height=src.row_dimensions[rr].height
        for m_ in src.merged_cells.ranges:
            if r1<=m_.min_row and m_.max_row<=r2: ms.merge_cells(start_row=m_.min_row+off,start_column=m_.min_col,end_row=m_.max_row+off,end_column=m_.max_col)
    d1=_data_rows(ws1); d2=_data_rows(ws); ncol=ws1.max_column
    lastD=d1[-1] if d1 else 2
    _copy_rows(ws1,1,lastD,0)
    sep=lastD+1                                                    # band row: where the other materials start
    ms.merge_cells(start_row=sep,start_column=1,end_row=sep,end_column=ncol)
    sc=ms.cell(sep,1,f'OTHER MATERIALS · on the PPC Count sheet, not in the BOM ({len(d2)}) · Need / day = PPC Count Material Req Qty per Day')
    sc.font=F(size=10,bold=True,color='FFFFFF'); sc.fill=PatternFill('solid',fgColor=NAVY); sc.alignment=Alignment(horizontal='left',vertical='center',indent=1)
    ms.row_dimensions[sep].height=20
    if d2:                                                         # highest risk first (user, Nov 2026: "sort from highest risk to lowest"):
        hd_=[ws.cell(2,j).value for j in range(1,ws.max_column+1)]  # Short dd-mmm (earliest first), Under 21 d, OK, No demand
        jr_=hd_.index('Shortage risk')+1 if 'Shortage risk' in hd_ else None; jf_=hd_.index('Short from')+1 if 'Short from' in hd_ else None
        ji_=hd_.index('Days incl. in transit')+1 if 'Days incl. in transit' in hd_ else None
        def _ok(r_):
            rk=str(ws.cell(r_,jr_).value or '') if jr_ else ''; sf=ws.cell(r_,jf_).value if jf_ else None; di=ws.cell(r_,ji_).value if ji_ else None
            g_=0 if rk.startswith('Short') else (1 if rk.startswith('Under') else (2 if rk=='OK' else 3))
            return (g_,sf if isinstance(sf,dt.datetime) else dt.datetime(2999,1,1),di if isinstance(di,(int,float)) else 1e12,r_)
        for k_,r_ in enumerate(sorted(d2,key=_ok)):
            _copy_rows(ws,r_,r_,sep+1+k_-r_)
        for k_,r_ in enumerate(range(sep+1,sep+1+len(d2)),1): ms.cell(r_,1).value=len(d1)+k_      # Sr. continues
    endO=sep+len(d2)
    _copy_rows(ws1,lastD+1,ws1.max_row,endO-lastD)               # summary boxes + notes of the BOM materials
    _copy_rows(ws,(d2[-1] if d2 else 2)+1,ws.max_row,ms.max_row+1-((d2[-1] if d2 else 2)+1))   # other materials note
    for k_,d_ in ws1.column_dimensions.items():
        if d_.width: ms.column_dimensions[k_].width=d_.width
    ms.freeze_panes='C3'
    for row_ in ms.iter_rows(min_row=endO+1):                     # the summary boxes count the BOM materials only (checker finding)
        for c in row_:
            if isinstance(c.value,str) and c.value.startswith(('⚠️  KEY RISK','⏱  SHORTAGE TIMELINE','⚠️  SHORTAGE WATCH')): c.value=c.value+' · BOM materials'
    del wb['Direct Material']; del wb['Other Materials']
    # ---- ALERT copy (user, Nov 2026: "Copy Alert this workbook but remove STOCK-OUT IN TRANSIT LEAD TIME / STOCK-OUT LATER / UNDER 21 DAYS / UNDER 21 DAYS LATER" = the 4 class cards) ----
    if 'ALERT' in src_wb.sheetnames:
        wa=wb.create_sheet('ALERT',1); sa=src_wb['ALERT']
        _copy_sheet_to(sa,wa); wa.freeze_panes=sa.freeze_panes; wa.sheet_view.zoomScale=sa.sheet_view.zoomScale
        if sa.auto_filter.ref: wa.auto_filter.ref=sa.auto_filter.ref
        wa.print_title_rows=sa.print_title_rows; wa.page_setup.orientation=sa.page_setup.orientation       # print like the MRP_Daily ALERT (checker finding)
        wa.page_setup.fitToWidth=sa.page_setup.fitToWidth; wa.page_setup.fitToHeight=sa.page_setup.fitToHeight; wa.sheet_properties.pageSetUpPr.fitToPage=True
        cr=next((c.row for c in wa['A'] if c.row>ALERT_HR and c.value==ALERT_CLS[0][0]),None)
        if cr: _drop_card_rows(wa,cr)
    # ---- 'Usage per module by customer' (= Stock Check / old Action List column) as the last column of both sheets (user, Nov 2026) ----
    UH='Usage per module by customer (source: your BOM / MIS Standard BOM)'
    use_of={}
    for x_ in AL_ORDER:
        if x_.get('sup_row'): continue
        v_=x_.get('usage_src'); v_=', '.join(map(str,v_)) if isinstance(v_,(list,tuple)) else v_
        if v_: use_of.setdefault(str(x_['b']['mat']).strip(),v_)
    def _usage_col(w, hr, rows_, code_col, band_r=None, last_r=None):
        nc=w.max_column if last_r is None else last_r
        j=nc+1
        h=w.cell(hr,j,UH); h._style=_cp(w.cell(hr,nc)._style); h.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        if band_r:
            b_=w.cell(band_r,j,'BOM · usage'); b_._style=_cp(w.cell(band_r,1)._style); b_.alignment=Alignment(horizontal='center',vertical='center')
        for r_ in rows_:
            cd_=' '.join(str(w.cell(r_,code_col).value or '').split())
            v_=use_of.get(cd_) or '-'
            c=w.cell(r_,j,v_); c._style=_cp(w.cell(r_,nc)._style); c.alignment=Alignment(horizontal='left',vertical='center',wrap_text=True)
            n_=sum(max(1,-(-len(t_)//44)) for t_ in str(v_).split('\n'))      # wrapped lines too (width 46 ~ 44 characters; checker finding)
            w.row_dimensions[r_].height=max(w.row_dimensions[r_].height or 15,min(15*n_,150))
        w.column_dimensions[L(j)].width=46
        return j
    jm=_usage_col(ms,2,list(range(3,lastD+1))+list(range(sep+1,endO+1)),2,band_r=1,last_r=ncol)
    for m_ in list(ms.merged_cells.ranges):                       # band row of the other materials spans the new column too
        if m_.min_row==sep and m_.max_row==sep: ms.unmerge_cells(str(m_)); ms.merge_cells(start_row=sep,start_column=1,end_row=sep,end_column=jm)
    if 'ALERT' in wb.sheetnames:
        wa=wb['ALERT']; ar_=[r_ for r_ in range(ALERT_HR+1,wa.max_row+1) if isinstance(wa.cell(r_,1).value,int) and wa.cell(r_,3).value]
        lastA=max([c.column for c in wa[ALERT_HR] if c.value is not None]+[1])
        ja=_usage_col(wa,ALERT_HR,ar_,3,band_r=ALERT_HR-1,last_r=lastA)
        if wa.auto_filter.ref: wa.auto_filter.ref=f"A{ALERT_HR}:{L(ja)}{ar_[-1] if ar_ else ALERT_HR+1}"
    # ---- Cover Heatmap (MRP_Daily, grouped RM / PM / CS) + RM Material Dashboard (daily WSA 'Dashboard') (user, Nov 2026) ----
    for n_ in (HEAT_NAME,DAILY_NAME):
        if n_ in src_wb.sheetnames:
            wc=wb.create_sheet(n_); sc_=src_wb[n_]; _copy_sheet_to(sc_,wc); wc.freeze_panes=sc_.freeze_panes
            for cf_ in sc_.conditional_formatting:                  # the daily sheet's colour rules
                for rl_ in cf_.rules: wc.conditional_formatting.add(str(cf_.sqref),rl_)
    try:
        wd=wb.create_sheet('RM Material Dashboard'); nd_=_wsa_dashboard_to(SRC,wd,d0)
        if nd_ is None: del wb['RM Material Dashboard']; print('NOTE: the WSA file has no Dashboard sheet - no RM Material Dashboard in Daily_Material_Status')
        else:                                                      # same dashboard on the MIS arrival dates (user, Nov 2026: "I want same story ... in new sheet")
            try:
                wm=wb.create_sheet(DASH_MIS); _wsa_dashboard_to(SRC,wm,d0,use_v=True)     # same layout; daily cells = column V (MIS arrival dates)
            except Exception as ex2_:
                if DASH_MIS in wb.sheetnames: del wb[DASH_MIS]
                print(f'NOTE: {DASH_MIS} not built ({type(ex2_).__name__}: {ex2_})')
    except Exception as ex_:
        if 'RM Material Dashboard' in wb.sheetnames: del wb['RM Material Dashboard']
        print(f'NOTE: RM Material Dashboard not built ({type(ex_).__name__}: {ex_})')
    # RM dashboards in the BOM dashboard design + its colour rule (user, Nov 2026)
    RM_NOTE=("Each day = cover that morning, days: red = short (below 1 day: the day's need is not covered) · orange / amber = below "
             f"{REPORT_DAYS} days · blue = {REPORT_DAYS}+ days · na = no demand that day. Highest risk on top (first day below 1, then below 5 days, "
             "then the lowest cover today); No Plan rows last. Same colour rule as the BOM Material Dashboard.")
    for n_,t_,src_ in (('RM Material Dashboard','Daily cover of your WSA planning sheets (your Scheduled Delivery)',
                        "Source: your WSA Dashboard, worked out on the aligned WSA planning sheets - column S Plant + Port Days (balance + your typed Scheduled Delivery) ÷ demand, one column per day."),
                       (DASH_MIS,'Daily cover of your WSA planning sheets (MIS arrival dates)',
                        "Source: your WSA Dashboard on column V Plant + Port Days (PPC In Transit) = balance + MIS / PPC In Transit arrivals by arrival date ÷ demand, one column per day.")):
        if n_ in wb.sheetnames:
            try: _dark_rm_dashboard(wb[n_],t_,RM_NOTE,src_)
            except Exception as ex_: print(f'NOTE: {n_} kept in the WSA design ({type(ex_).__name__}: {ex_})')
    # In transit supplier name (qty) on EVERY sheet (user, Nov 2026: "Add in transit supplier name in all sheet of daily material status"):
    # right after 'In transit' (Materials Status, ALERT, BOM Weekly), before the day columns (BOM Daily, RM dashboards)
    def _col_of(w,hr,t_): return next((c.column for c in w[hr] if isinstance(c.value,str) and c.value.strip().startswith(t_)),None)
    def _code_rows(w,r0,cc):
        return [r_ for r_ in range(r0,w.max_row+1) if any(IT_CODE.match(x_.strip()) for x_ in re.split(r'\s*(?:/|,|\+|\n)\s*',str(w.cell(r_,cc).value or '')))]
    nsup={}
    try:
        if 'Materials Status' in wb.sheetnames:
            w=wb['Materials Status']; j=_col_of(w,2,'In transit')
            if j: nsup['Materials Status']=_add_it_sup_col(w,2,j+1,list(range(3,lastD+1))+list(range(sep+1,endO+1)),2,3)
        if 'ALERT' in wb.sheetnames:
            w=wb['ALERT']; j=_col_of(w,ALERT_HR,'In transit')
            if j: nsup['ALERT']=_add_it_sup_col(w,ALERT_HR,j+1,[r_ for r_ in range(ALERT_HR+1,w.max_row+1) if isinstance(w.cell(r_,1).value,int) and w.cell(r_,3).value],3,4)
        if HEAT_NAME in wb.sheetnames:
            w=wb[HEAT_NAME]; j=_col_of(w,4,'In transit')
            if j: nsup[HEAT_NAME]=_add_it_sup_col(w,4,j+1,_code_rows(w,5,1),1,2,dark=True)
        for n_ in (DAILY_NAME,'RM Material Dashboard',DASH_MIS):
            if n_ in wb.sheetnames:
                w=wb[n_]; j=_col_of(w,4,"Today's Morning stock")
                if j: nsup[n_]=_add_it_sup_col(w,4,j+1,_code_rows(w,5,1),1,2,sheet_col=(3 if n_!=DAILY_NAME else None),plain=True)
    except Exception as ex_:
        import traceback; traceback.print_exc(); print(f'NOTE: In transit supplier column not added everywhere ({type(ex_).__name__}: {ex_})')
    # user, Nov 2026: "Materials Status, delete row 1, then hide this sheet; ALERT, delete row 10 then move row 1 to 9 to bottom"
    if 'Materials Status' in wb.sheetnames:                        # the group band row goes: header = row 1, data from row 2
        _remap_rows(wb['Materials Status'],lambda r_: None if r_==1 else r_-1)
    if 'ALERT' in wb.sheetnames:                                   # group band (row ALERT_HR-1) goes; summary boxes (rows 1 .. ALERT_HR-2) below the table
        wa=wb['ALERT']; top=ALERT_HR-2; end_=wa.max_row-(top+1)    # rows after the band move up by top + 1
        _remap_rows(wa,lambda r_: (end_+1+r_ if r_<=top else (None if r_==top+1 else r_-(top+1))))
    for n_ in (HEAT_NAME,DAILY_NAME,'RM Material Dashboard',DASH_MIS):          # user, Nov 2026: "Dont freeze or wrap text ... all dashboards look organized"
        if n_ in wb.sheetnames: _tidy_dashboard(wb[n_])
    if HEAT_NAME in wb.sheetnames: _heat_final(wb[HEAT_NAME])             # rows 1-2 out, 9 pt, Material Status columns grouped (user, Nov 2026)
    for n_ in ('Materials Status','ALERT','RM Material Dashboard',DAILY_NAME):   # hidden in the daily file (user, Nov 2026: "Hide sheet RM Material Dashboard ... Hide sheet alert"; Materials Status, BOM Material Dashboard - Daily too)
        if n_ in wb.sheetnames: wb[n_].sheet_state='hidden'
    vis=[i_ for i_,w_ in enumerate(wb.worksheets) if w_.sheet_state=='visible']
    wb.active=vis[0] if vis else 0                                 # the first VISIBLE sheet opens (a hidden active sheet = Excel repair)
    for i_,w_ in enumerate(wb.worksheets): w_.sheet_view.tabSelected=(i_==wb.index(wb.active))
    if wb.views: wb.views[0].firstSheet=wb.index(wb.active)   # first tab shown = a visible sheet (checker note)
    _safe_save(wb,out_file)
    print(f"Daily Material Status saved: {out_file} - Materials Status ({len(d1)} BOM materials + other materials), ALERT, {HEAT_NAME}, {DAILY_NAME}, RM Material Dashboard, {DASH_MIS}; other materials ({len(ind)} Count materials not in the BOM: {ncount} counted, {nsap} with MB52 stock, {nneed} with a need per day)"
          +(f"; left out of Other Materials because they are alias / substitute codes of a BOM material: {', '.join(skipped)}" if skipped else ''),flush=True)
    if nsup: print(f"  '{IT_SUP_HEAD}' added: "+', '.join(f'{k} {v} rows' for k,v in nsup.items()),flush=True)
    return out_file

def _heat_final(w):
    """BOM Material Dashboard - Weekly, last step in both outputs (user, Nov 2026: "Bom material dashboard: group column E to N. Font size
       9. Delete row 1 & 2"): title + note rows deleted (header row 4 -> 2, data from row 3), every cell font 9 pt, the Material Status
       columns (E .. the column before the first week = E..N in Daily_Material_Status with the supplier column, E..M in MRP_Daily)
       grouped as an Excel outline (expanded; the - / + button above the columns collapses them). Values unchanged."""
    import copy as _cp
    from openpyxl.utils import get_column_letter as _gl
    _remap_rows(w,lambda r_: None if r_<=2 else r_-2)
    for row_ in w.iter_rows():
        for c_ in row_:
            f_=c_.font
            if f_ is not None and f_.sz!=9:
                f2_=_cp.copy(f_); f2_.sz=9; c_.font=f2_
    last=next((c_.column for c_ in w[2] if str(c_.value or '').strip().lower()=='days incl. in transit'),None)   # last Material Status column
    first=last+1 if last else None
    if first and first>5:
        for j_ in range(5,first):
            w.column_dimensions[_gl(j_)].outline_level=1
        w.column_dimensions.max_outline=1                             # -> sheetFormatPr outlineLevelCol=1 (openpyxl writes it from this; Excel draws the outline bar from it)
    return first
def _tidy_dashboard(w, hr=4):
    """Daily_Material_Status dashboards (user, Nov 2026: "Dont freeze or wrap text, I want all dashboards ... look organized" - no cell
       under the dates with another width / height): no frozen pane, no wrapped text, every row below the header the same height,
       every day / week column the same width, text columns as wide as their longest text (one line). Values untouched."""
    from copy import copy as _cp
    w.freeze_panes=None
    isday=lambda v: isinstance(v,(dt.datetime,dt.date)) or bool(re.fullmatch(r'\d{1,2}-[A-Za-z]{3}(-\d{2,4})?',str(v or '').strip()))
    last=max([c.column for c in w[hr] if c.value is not None]+[1])
    d1=next((c for c in range(1,last+1) if isday(w.cell(hr,c).value)),last+1)
    for row in w.iter_rows():
        for c in row:
            if c.alignment is not None and c.alignment.wrap_text:
                al=_cp(c.alignment); al.wrap_text=False; c.alignment=al
    rows_=[r_ for r_ in range(hr+1,w.max_row+1) if IT_CODE.match(str(w.cell(r_,1).value or '').split(' ')[0].strip()) or str(w.cell(r_,1).value or '').strip()=='-']
    ln=lambda v: 0 if v is None else (9 if isinstance(v,(dt.datetime,dt.date)) else (len(f'{v:,.1f}') if isinstance(v,float) else max(len(x_) for x_ in str(v).split('\n'))))
    for c in range(1,d1):
        m_=max([ln(w.cell(r_,c).value) for r_ in [hr]+rows_]+[4])
        w.column_dimensions[L(c)].width=min(round(m_*0.9+3,1),120)
        if str(w.cell(hr,c).value or '')==IT_SUP_HEAD:               # 2-3 very long supplier lists made it 113 wide = the day columns off the
            w.column_dimensions[L(c)].width=min(w.column_dimensions[L(c)].width,48)   # screen (checker finding): 48 fits most rows; longer = cut on screen, full text in the cell
            al=_cp(w.cell(hr,c).alignment); al.horizontal='left'; w.cell(hr,c).alignment=al
        if str(w.cell(hr,c).value or '').startswith("Today's Morning stock ("):    # long header over 1-3 digit numbers: old width (checker finding)
            w.column_dimensions[L(c)].width=min(w.column_dimensions[L(c)].width,22)
    dw=max([w.column_dimensions[L(c)].width or 0 for c in range(d1,last+1)]+[4.6]) if d1<=last else 0
    for c in range(d1,last+1): w.column_dimensions[L(c)].width=dw
    for r_ in range(1,w.max_row+1):
        if r_!=hr: w.row_dimensions[r_].height=15

def make_stock_data_xlsx(DS, A, path):
    """Stock Data: every line behind the stock of each material. Action List, MRP Daily and Stock Movement add these lines up with SUMIFS,
       so every stock number can be traced back here (and from here to MB52 / RM Status / Stock Count)."""
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    wb=Workbook(); ws=wb.active; ws.title='Stock Data'; ws.sheet_view.showGridLines=False
    ws['A1']='Stock Data - every line behind the stock of each material'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=("Type Stock = MIS MB52 Total Stock (plant + warehouse, blocked stock is not in it) - a counted material uses the PPC Count sheet, rolled forward to today when the count is older (usage / receipts lines below). "
              "Type to OEM = MIS RM Status for OEM JW, column 'JW WSA to OEM'. Qty (MRP unit) = Qty x Factor (alternate tape roll lengths converted to the material's roll). "
              "MRP Daily (stock today) = the Type Stock lines of the material; the Type to OEM lines are shown for information only - the JW stock at the OEM is in no calculation.")
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:M2'); ws['A2'].alignment=Alignment(wrap_text=True,vertical='top'); ws.row_dimensions[2].height=40
    H=['Component','Material','Type','Source','Source code','Batch','Storage location','Location','Qty (source unit)','Source UOM','Factor','Qty (MRP unit)','MRP UOM']
    for j,h in enumerate(H,1):
        c=ws.cell(4,j,h); c.font=Font(name=AR,size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=NAVY); c.border=bd; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    info={x['b']['mat']:x['b'] for x in A}
    lines=[l_ for l_ in (DS.get('stock_lines') or []) if l_['code'] in info]
    for x in A:
        if x['b'].get('count_stock'): lines.append(dict(code=x['b']['mat'],type='Stock',source="PPC Count (physical count)",src=x['b']['mat'],batch=None,sloc=None,loc=None,
                                                        qty=None,uom=x['b'].get('uom'),factor=1.0,cs=True))
    lines.sort(key=lambda l_:(rrank(l_['code']),l_['type']!='Stock',str(l_['src']),str(l_['batch'] or '')))     # highest shortage risk first
    r=5
    for l_ in lines:
        b=info[l_['code']]
        vals=[l_['code'],b.get('desc') if b.get('count_stock') else b['common'],l_['type'],l_['source'],l_['src'],l_['batch'],l_['sloc'],l_['loc'],
              (round(b['stock'],6) if (l_.get('cs') and VALUES_ONLY) else (f"=SUMIFS('Stock Count'!${SC_COL}:${SC_COL},'Stock Count'!$A:$A,$E{r})" if l_.get('cs') else round(l_['qty'],6))),
              l_['uom'],l_['factor'],(round((b['stock'] if l_.get('cs') else l_['qty'])*l_['factor'],6) if VALUES_ONLY else f'=I{r}*K{r}'),b.get('uom')]
        for j,v in enumerate(vals,1):
            c=ws.cell(r,j,v); c.font=Font(name=AR,size=9,bold=(j==12),color=('9C0006' if l_['type']=='to OEM' and j==3 else '000000')); c.border=bd
            if j in (9,12): c.number_format='#,##0.###;-#,##0.###;"-"'
            if j==11: c.number_format='0.####'
        r+=1
    if r>5: ws.add_table(Table(displayName='StockData',ref=f'A4:M{r-1}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=True)))
    for j,w in enumerate([17,30,9,40,17,14,22,18,14,8,7,14,8],1): ws.column_dimensions[L(j)].width=w
    set_print(ws,'Stock Data'); wb.save(path)
    return len(lines)

def make_movement_log_xlsx(log, path):
    wb=Workbook(); lg=wb.active; lg.title='Movement Log'
    lg.append(['Date','Code','Material','UOM','Stock','Plan next days','In transit next days','Open PO next days','Raw stock','In transit by supplier','to OEM Stock'])
    for e in sorted(log,key=lambda e_:(-e_['date'].value,rrank(e_['code']))):        # newest day first; within a day highest shortage risk first
        lg.append([e['date'].to_pydatetime(),e['code'],e['material'],e['uom'],e['stock'],';'.join(f'{x:.6g}' for x in e['plan']),';'.join(f'{x:.6g}' for x in e['inb']),';'.join(f'{x:.6g}' for x in e['opo']),e.get('raw',''),
                   '|'.join(f"{k}~{q:.6g}~{sn}" for k,q,sn in (e.get('sup') or [])),e.get('oem')])
    wb.save(path)

UNIT_SAME=[{'PC','PCS','NOS','NO','EA','EACH'},{'KG','KGS'},{'M2','SQM'},{'ROL','ROLL','ROLLS'},{'SET','SETS'},{'DR','DRUM'}]
def _u(x): return re.sub(r'[^A-Z0-9]','',str(x or '').upper())
def same_unit(a,b):
    a,b=_u(a),_u(b)
    return a==b or any(a in g and b in g for g in UNIT_SAME)

def make_unit_check_xlsx(DS, bom, cnt_hdr, cnt_rows, modplan, mis_path, path, H=30):
    """A) stock unit vs usage unit for every material and the conversion used; B) plan = modules x usage per module, by MIS BOM Name,
       reconciled to the MRP Daily plan (next H days). Returns the unit issues for the Check sheet."""
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    fills={'OK':'C6EFCE','Converted':'DDEBF7','Missing conversion':'FFC7CE','Unit unknown':'FFEB9C','Lookup error':'FFEB9C','Mismatch':'FFC7CE'}
    issues=[]
    mis=wb_ro(mis_path)
    plant=DS['plant']
    mbu={}; tpu={}; rate_t={}; seen=set()
    for r in mis['MB52'].iter_rows(min_row=2,values_only=True):
        if r[2] and r[7]: mbu.setdefault(str(r[2]).strip(),str(r[7]).strip())
    for r in mis['Total Plan Req'].iter_rows(min_row=2,values_only=True):
        if not tp_ok(r,plant): continue
        c=str(r[10] or '').strip(); bn=str(r[4] or '').strip()
        if r[13]: tpu.setdefault(c,str(r[13]).strip())
        if isinstance(r[3],(int,float)) and r[3] and isinstance(r[14],(int,float)):
            t_=rate_t.setdefault((c,bn),[0.0,0.0]); t_[0]+=r[14]
            k_=(c,bn,r[0],r[1],r[2],r[17] if len(r)>17 else None)
            if k_ not in seen: seen.add(k_); t_[1]+=r[3]
    mrate={k:(a/b) for k,(a,b) in rate_t.items() if b}
    hl=[str(h).strip().lower() if h is not None else '' for h in (cnt_hdr or [])]
    ix=lambda n: hl.index(n) if n in hl else None
    i_su,i_bu,i_q=ix('stock uom'),ix('bom uom'),ix('quantity')
    cnt={str(r[0]).strip():r for r in (cnt_rows or []) if r and r[0]}
    bomu={}; bomq={}
    for d in bom or []:
        c=str(d.get('Component') or '').strip(); bn=str(d.get('MIS BOM Name') or '').strip()
        if not c: continue
        if d.get('UoM'): bomu.setdefault(c,str(d['UoM']).strip())
        if bn and isinstance(d.get('Quantity'),(int,float)):
            mq=bomq.setdefault((c,bn),{}); mdl=str(d.get('Module Model') or '').strip(); mq[mdl]=mq.get(mdl,0.0)+float(d['Quantity'])
    wb=Workbook(); ws=wb.active; ws.title='Unit Check'
    ws['A1']='Unit check - stock unit vs usage unit, and plan = modules x usage'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=('A) Every material: the unit its stock is counted in and the unit its usage per module is in, and the conversion the MRP uses. '
              'RM: stock MB52 or PPC Count / usage PPC BOM, else MIS Standard BOM (alternate codes converted with MIS\'s Adhoc ratio). CS / PM: stock = PPC Count (Today Qty_BOM = Today Stock Qty x Quantity) / usage = PPC BOM. '
              f'B) Plan check: modules in the next {H} days (your Production Plan, by MIS BOM Name) x usage per module = plan; the total must match the MRP Daily plan.')
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:K2'); ws['A2'].alignment=Alignment(wrap_text=True); ws.row_dimensions[2].height=40
    H1=['Group','Material','Code','Stock unit','Usage unit','Conversion used','Factor','Status','Note']
    for j,h in enumerate(H1,1):
        c=ws.cell(4,j,h); c.font=Font(name=AR,size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=NAVY); c.border=bd; c.alignment=Alignment(horizontal='center',wrap_text=True)
    r=5; alias=DS.get('alias') or {}; fac=DS.get('alias_factor') or {}
    for b in sorted(DS['blocks'],key=lambda b_:rrank(b_['mat'])):
        c=b['mat']
        if b.get('count_stock'):
            row=cnt.get(c); su=row[i_su] if (row and i_su is not None) else None; bu=bomu.get(c); q=row[i_q] if (row and i_q is not None) else None
            luk=row[i_bu] if (row and i_bu is not None) else None
            if row is None: st,conv,note='Unit unknown','-','Not on your Stock Count - stock is 0'
            elif su in (None,'','#N/A'): st,conv,note,su='Unit unknown','-','Stock UOM not found on the PPC Count (code not in MB52)','(unknown)'
            elif same_unit(su,bu): st,conv,note='Units match','none (same unit)',''
            elif isinstance(q,(int,float)) and q not in (0,1): st,conv,note='Converted',f'1 {su} = {q:g} {bu}',''
            else: st,conv,note='Missing conversion','none',f'Stock in {su}, BOM usage in {bu} - put {bu} per {su} in Quantity on your Stock Count'
            if luk in ('#N/A',None,'') and bu and st!='Unit unknown': note=(note+'; ' if note else '')+'BOM UOM lookup fails on your Stock Count (BOM range too short) - unit taken from the BOM sheet'
            vals=['CS / PM',b['desc'],c,su,bu,conv,(q if st=='Converted' else None),st,note]
            if st=='Missing conversion': issues.append(dict(area='Units',check='Stock unit differs from BOM unit, no conversion',sev='High',material=b['desc'],code=c,a_lab='Stock unit',a=None,b_lab='BOM unit',b=None,diff=None,detail=note,action=f'Fill Quantity ({bu} per {su}) on your Stock Count'))
        else:
            su=mbu.get(c) or b.get('uom'); pu=tpu.get(c) or b.get('uom')
            alts=[(a_,fac.get(a_)) for a_,d_ in alias.items() if d_==c]
            st='Units match' if same_unit(su,pu) else 'Mismatch'
            conv='none (same unit)' if st=='Units match' else 'none'
            note=''
            tp_=b.get('tape')
            if tp_:
                conv=f"tape rolls by length: 1 {pu} of {c} = {tp_.get('len'):g} m; "+', '.join(f"{a_} {l:g} m = x{l/tp_['len']:g}" for a_,l in tp_['lens'].items() if l)
                alts=[]; st='Converted' if tp_['lens'] else st
            if alts:
                conv+='; alternate codes: '+', '.join(f"{a_} (x{f_:g} {mbu.get(a_,'')}->{pu})" if f_ else f"{a_} (not converted, not counted)" for a_,f_ in alts)
                if any(f_ for _,f_ in alts): st='Converted' if st=='Units match' else st
            if st=='Mismatch':
                note=f'MB52 unit {su} vs MIS plan unit {pu}'
                issues.append(dict(area='Units',check='RM stock unit differs from MIS plan unit',sev='High',material=b['common'],code=c,a_lab='',a=None,b_lab='',b=None,diff=None,detail=note,action='Ask MIS to align the units'))
            vals=['RM',b['common'],c,su,pu,conv,None,st,note]
        for j,v in enumerate(vals,1):
            cc=ws.cell(r,j,v); cc.font=Font(name=AR,size=9); cc.border=bd; cc.alignment=Alignment(vertical='top',wrap_text=j in (2,6,9))
        ws.cell(r,8).fill=PatternFill('solid',fgColor=fills.get(vals[7],'FFFFFF')); r+=1
    ws.add_table(Table(displayName='UnitCheck',ref=f'A4:I{r-1}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=True)))
    # ---------- B) plan = modules x usage ----------
    top=r+2; ws.cell(top,1,f'B) Plan check - modules x usage per module (next {H} days) vs MRP Daily plan').font=Font(name=AR,size=12,bold=True,color=NAVY)
    H2=['Material','Code','MIS BOM Name','Modules (next %d d)'%H,'Usage / module','Usage unit','Plan = modules x usage','Usage from','MRP Daily plan','Difference','Match?']
    for j,h in enumerate(H2,1):
        c=ws.cell(top+1,j,h); c.font=Font(name=AR,size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=NAVY); c.border=bd; c.alignment=Alignment(horizontal='center',wrap_text=True)
    r=top+2; nbad=0
    for b in DS['blocks']:
        c=b['mat']; cs=b.get('count_stock'); tot=0.0
        for bn,cu in sorted((modplan or {}).items()):
            mods=sum(sum(v[:H]) for v in cu.values())
            if not mods: continue
            if cs:
                mq=bomq.get((c,bn)); q=(sum(mq.values())/len(mq)) if mq else None; src='Your BOM'
            else:
                q=mrate.get((c,bn)); src='MIS standard BOM'
            if not q: continue
            pl=mods*q; tot+=pl
            for j,v in enumerate([b['desc'] if cs else b['common'],c,bn,round(mods),round(q,6),bomu.get(c) if cs else (tpu.get(c) or b.get('uom')),round(pl,2),src,None,None,None],1):
                cc=ws.cell(r,j,v); cc.font=Font(name=AR,size=9); cc.border=bd
                if j in (4,7): cc.number_format='#,##0'
            r+=1
        mrp=sum(b['plan'][:H])
        if tot==0 and mrp==0:
            continue
        other=mrp-tot
        if abs(other)>max(1,0.005*mrp):     # MIS plan for customers not in your Production Plan, or rounding
            for j,v in enumerate([b['desc'] if cs else b['common'],c,'(MIS plan - customers not in your Production Plan)',None,None,None,round(other,2),'MIS plan',None,None,None],1):
                cc=ws.cell(r,j,v); cc.font=Font(name=AR,size=9,italic=True); cc.border=bd
                if j==7: cc.number_format='#,##0'
            r+=1; tot+=other
        diff=mrp-tot; ok=abs(diff)<=max(1,0.001*mrp)
        for j,v in enumerate([b['desc'] if cs else b['common'],c,'TOTAL',None,None,None,round(tot,2),None,round(mrp,2),round(diff,2),'Match' if ok else 'Different'],1):
            cc=ws.cell(r,j,v); cc.font=Font(name=AR,size=9,bold=True); cc.border=bd
            if j in (7,9,10): cc.number_format='#,##0'
        ws.cell(r,11).fill=PatternFill('solid',fgColor='C6EFCE' if ok else 'FFC7CE'); nbad+=0 if ok else 1
        r+=2
    nunit=sum(1 for x in issues)
    ws.cell(3,1,f'Unit issues: {nunit}. Plan check: {nbad} materials where plan does not equal modules x usage.').font=Font(name=AR,size=10,bold=True,color='9C0006' if (nunit or nbad) else '006100')
    for j,w in enumerate([30,18,30,12,12,10,40,16,12,10,10],1): ws.column_dimensions[L(j)].width=w
    ws.freeze_panes='C5'; wb.save(path)
    return issues

def line_substitutes(DS, mis_path):
    """Substitutes for SOME lines only (Settings): {MRP code: [dict(code, factor, base, arr, only=set of lines, win=(from,to), order)]}."""
    alt={}
    if not SUBS_LINES: return alt
    mis=wb_ro(mis_path); stock={}; arr={}; d0=pd.Timestamp(DS['dates'][0]).normalize(); nd=len(DS['dates'])
    for r in mis['MB52'].iter_rows(min_row=2,values_only=True):
        c=str(r[2] or '').strip()
        if c and str(r[13] or '').strip().upper() in (DS['plant'].upper(),DS['plant'].upper()+' WAREHOUSE') and isinstance(r[14],(int,float)): stock[c]=stock.get(c,0.0)+r[14]
    for r in mis['RM Status for OEM JW'].iter_rows(min_row=2,values_only=True):
        if r and r[1] and isinstance(r[11],(int,float)): stock[str(r[1]).strip()]=stock.get(str(r[1]).strip(),0.0)+r[11]
    for r in list(mis['In Transit'].iter_rows(values_only=True))[1:]:
        if len(r)>39 and r[39] and isinstance(r[34],(int,float)) and r[34]>0 and it_ok(r):
            try: k=max(0,(pd.Timestamp(r[35]).normalize()-d0).days)
            except Exception as _e_: _swallow(5373,_e_); continue
            if k<nd: arr.setdefault(str(r[39]).strip(),[0.0]*nd)[k]+=r[34]
    for x in SUBS_LINES:
        kf=0 if x['fr'] is None else max(0,(x['fr']-d0).days); kt=nd-1 if x['to'] is None else (x['to']-d0).days
        alt.setdefault(x['mrp_code'],[]).append(dict(code=x['sub_code'],factor=x['factor'],base=stock.get(x['sub_code'],0.0),arr=arr.get(x['sub_code'],[0.0]*nd),
                                                    only=set(x['lines']),win=(kf,kt),order=x['order'],pri={lk:x['order'] for lk in x['lines']}))
    return alt




DATA_DATES={}      # source -> data date (MIS, WSA, counts) - shown on the dashboards
def file_data_date(path):
    """Date the file was last saved by Excel (docProps/core.xml 'modified'), else the file date."""
    import zipfile
    try:
        x=zipfile.ZipFile(path).read('docProps/core.xml').decode('utf-8','ignore')
        m_=re.search(r'<dcterms:modified[^>]*>([^<]+)<',x)
        if m_: return pd.Timestamp(m_.group(1)[:10]).normalize()
    except Exception as _e_: _swallow(5392,_e_); pass
    try: return pd.Timestamp(dt.datetime.fromtimestamp(os.path.getmtime(path))).normalize()
    except Exception as _e_: _swallow(5394,_e_); return None

GR_LINES={}       # MRP code -> [(GR date, qty in the MRP unit)] from MIS In Transit 'PPC GRN' + 'GR. Date'
GR_NODATE={}      # MRP code -> PPC GRN qty with no GR. Date (cannot be placed in a month)
COUNT_HIST={}     # MRP code -> {count date: stock} (PPC Count date columns; every count row of the material counted on / before that date)
COUNT_HIST_BOM={} # the same in the BOM unit (Physical Count Qty_BOM / Physical Count Qty, e.g. cartons x 50 pcs) - what CS / PM materials use
def _any_date(v):
    """A date cell: datetime, an Excel serial number (MIS 'GR. Date' is saved as a number, e.g. 46330 = 04-Nov-2026) or text."""
    if isinstance(v,(dt.datetime,dt.date)): return pd.Timestamp(v).normalize()
    if isinstance(v,(int,float)) and not isinstance(v,bool) and 30000<v<80000: return (pd.Timestamp('1899-12-30')+pd.Timedelta(days=int(v))).normalize()
    if isinstance(v,str) and v.strip():
        try: return pd.Timestamp(pd.to_datetime(v.strip(),dayfirst=False)).normalize()
        except Exception as _e_: _swallow(5406,_e_); return None
    return None
def read_gr_lines(mis_path, alias=None, fac=None):
    """Goods received (GR) per material and day: MIS In Transit 'PPC GRN' on 'GR. Date' (the shipment's traced material,
       alternate codes converted to the MRP unit like the in-transit quantities)."""
    GR_LINES.clear(); GR_NODATE.clear(); alias=alias or {}; fac=fac or {}
    try: rows=list(wb_ro(mis_path)['In Transit'].iter_rows(values_only=True))
    except Exception as _e_: _swallow(5413,_e_); return
    I_=IT_CANON.index; iF,iG,iD=I_('Final RM Code'),I_('PPC GRN'),I_('GR. Date')
    for r in rows[1:]:
        if not r or len(r)<=iD: continue
        q=r[iG]; c=str(r[iF] or '').strip()
        if not c or not isinstance(q,(int,float)) or isinstance(q,bool) or q<=0: continue
        if c in alias: q=q*fac.get(c,1.0); c=alias[c]
        d=_any_date(r[iD])
        if d is None: GR_NODATE[c]=GR_NODATE.get(c,0.0)+float(q); continue
        GR_LINES.setdefault(c,[]).append((d,float(q)))
def gr_between(code, a, b):
    """GR quantity of the material with a GR date after a and up to b (a None = no lower limit)."""
    return sum(q for d,q in GR_LINES.get(code,[]) if (a is None or d>pd.Timestamp(a)) and d<=pd.Timestamp(b))
def read_count_history(mrp_path, alias=None, fac=None):
    """PPC Count date columns -> COUNT_HIST: the material's counted stock on each count day. A material with several count rows
       (e.g. glass per supplier / pallet) has a stock on a day only when every row has a count on or before that day
       (each row's latest count is used). Same unit as the count the MRP uses (Physical Count Qty)."""
    COUNT_HIST.clear(); COUNT_HIST_BOM.clear(); alias=alias or {}; fac=fac or {}
    if not PPC_SRC[0] or not mrp_path: return
    try: rows=[list(r) for r in wb_ro(mrp_path)['Stock Count'].iter_rows(values_only=True)]
    except Exception as _e_: _swallow(5433,_e_); return
    hi=next((i for i,r in enumerate(rows[:10]) if r and str(r[0] or '').strip().lower()=='material'),None)
    if hi is None: return
    hl=[' '.join(str(x or '').split()).lower() for x in rows[hi]]
    ic=next((hl.index(k) for k in ('physical count qty','today stock qty') if k in hl),None)
    iq=next((hl.index(k) for k in ('physical count qty_bom','today qty_bom') if k in hl),None); iqf=hl.index('quantity') if 'quantity' in hl else None
    dcols=[(j,pd.Timestamp(h).normalize()) for j,h in enumerate(rows[hi]) if isinstance(h,(dt.datetime,dt.date))]
    per={}
    for r in rows[hi+1:]:
        if not r or not r[0] or ic is None or ic>=len(r) or not isinstance(r[ic],(int,float)) or isinstance(r[ic],bool): continue   # counted rows only (= the MRP)
        c=str(r[0]).strip(); f=1.0
        if c in alias: f=fac.get(c,1.0); c=alias[c]
        h={d:float(r[j])*f for j,d in dcols if j<len(r) and isinstance(r[j],(int,float)) and not isinstance(r[j],bool)}
        qb=r[iq] if iq is not None and iq<len(r) else None; qf=r[iqf] if iqf is not None and iqf<len(r) else None
        bf=(float(qb)/float(r[ic])) if (isinstance(qb,(int,float)) and not isinstance(qb,bool) and r[ic]) else (float(qf) if isinstance(qf,(int,float)) and not isinstance(qf,bool) and qf else 1.0)
        if h: per.setdefault(c,[]).append((h,bf))
    for c,hs in per.items():
        days=sorted({d for h,_ in hs for d in h}); out={}; outb={}
        for d in days:
            if all(any(x<=d for x in h) for h,_ in hs):
                out[d]=sum(h[max(x for x in h if x<=d)] for h,_ in hs); outb[d]=sum(h[max(x for x in h if x<=d)]*bf for h,bf in hs)
        if out: COUNT_HIST[c]=out; COUNT_HIST_BOM[c]=outb

def apply_count_by_supplier(blocks):
    """Counted material with supplier detail (b['sup']): the PPC Count is the stock, ALSO by supplier (user rule, Nov 2026:
       "if have count then use count"). Each supplier row = its counted quantity (Count 'Supplier Code'), scaled to today's
       rolled-forward plant stock (st_main, JW at the OEM not included) (count - usage + GR since the count; same share for every counted supplier). A supplier with MB52
       batches but no count row = 0; no 'Unassigned' row. Runs after estimate_counted_stock, so MRP Daily, the Action List supplier
       rows and the WSA sheets (align_capture 'sup') all read the same split. Returns the number of materials changed."""
    n_=0
    for b in blocks:
        cnt_=(COUNTED_RM.get(b['mat']) or {}).get('by_sup') or {}
        if not b.get('sup') or not cnt_: continue
        tot_=sum(float(q_[1] or 0) for q_ in cnt_.values())
        own_=float(b['st_main']) if b.get('st_main') is not None else float(b['stock'])-(0.0 if b.get('jw_out') else float(b.get('st_jw') or 0))
        f_=(max(0.0,own_)/tot_) if tot_>0 else 0.0                       # plant stock only (JW at the OEM is not in the count)
        rows_=[dict(sp) for sp in b['sup']]            # new rows: earlier copies of the block (Check 'Glass stock without supplier') keep the MB52 split
        def _hit(nm_,sc_):
            u_=str(nm_ or '').upper(); return u_!='UNASSIGNED' and (SUP_CODE.get(sc_,sc_) in u_ or u_.startswith(sc_))
        for sp in rows_: sp['stock']=0.0
        b['sup_cnt']={}
        for sc_,q_ in cnt_.items():
            r_=next((sp for sp in rows_ if _hit(sp.get('name'),sc_)),None)
            if r_ is None:
                nd_=len(rows_[0]['transit']) if rows_ else 0
                r_=dict(name=SUP_CODE.get(sc_,sc_),stock=0.0,transit=[0.0]*nd_,openpo=[0.0]*nd_); rows_.append(r_)
            r_['stock']+=float(q_[1] or 0)*f_
            b.setdefault('sup_cnt',{})[r_['name']]=b.setdefault('sup_cnt',{}).get(r_['name'],0.0)+float(q_[1] or 0)   # physical count by supplier, as counted (Stock Check)
        b['sup']=[sp for sp in rows_ if abs(sp['stock'])>0.5 or any(sp['transit']) or any(sp['openpo'])]
        b['sup'].sort(key=lambda x:(x['name']=='Unassigned',-(x['stock']+sum(x['transit'])+sum(x['openpo']))))
        n_+=1
    return n_

def _plan_without_oem(b):
    """Daily plan of a material WITHOUT the OEM customers (OEM = Yes on Settings): the OEM builds from its own JW stock,
       so its plan does not use plant stock. Used by the count roll-forward, which runs before the OEM step of build_new_mrp."""
    P=list(b.get('plan') or [])
    if b.get('plan_oem') is not None: return P                     # OEM step already done: plan has no OEM
    oem_=[_norm(cu) for cu,ln,tg,o in CUST_TARGET if str(o).strip().lower()=='yes']
    if not oem_: return P
    def _is_oem(cu):
        n_=_norm(cu); return bool(n_) and any(n_==o_ or n_.startswith(o_) or o_.startswith(n_) for o_ in oem_)
    cl_=b.get('cust_line') or {}
    ok_=[k_ for k_ in cl_ if _is_oem(k_[0] if isinstance(k_,tuple) else k_)]
    for k_ in ok_:
        v_=cl_[k_]; P=[max(0.0,p_-(v_[k] if k<len(v_) else 0.0)) for k,p_ in enumerate(P)]
    return P

def estimate_counted_stock(blocks, mis_path, d0):
    """A physical count from an earlier day is rolled forward to today:
       estimated stock = count - usage per day x days since the count + goods received (MIS PPC GRN) after the count date.
       Usage per day = average Production Plan demand of the next 7 days, OEM plan left out. Fills STOCK_EST for the Stock Check sheet."""
    STOCK_EST.clear()
    pc=read_phys_counts(SETTINGS_SRC[0]) if SETTINGS_SRC and SETTINGS_SRC[0] else {}
    mis_d=file_data_date(USER_MB52[0]) if USER_MB52[0] else file_data_date(mis_path)        # stock date: your MB52 file (else the MIS file)
    DATA_DATES['MB52' if USER_MB52[0] else 'MIS']=mis_d
    if globals().get('STOCK_DATE') is not None: DATA_DATES['WSA']=pd.Timestamp(globals()['STOCK_DATE']).normalize()
    mis_age=max(0,(d0-mis_d).days) if mis_d is not None else 0
    if mis_age>=1: print(f"NOTE: {'your MB52' if USER_MB52[0] else 'MIS'} file is dated {mis_d:%d-%b} ({mis_age} days old) - its stock is used as TODAY's stock (your rule: an old MB52 = the newest), not rolled forward.")
    mis_age=0          # user, Nov 2026: "when MB52 is old, just take it as newest" - no roll-forward of SAP stock by the file's age
    sap={}
    try:
        mis=wb_ro(mis_path)
        for r in list(mis['MB52'].iter_rows(values_only=True))[1:]:
            if r and len(r)>14 and r[2] and isinstance(r[14],(int,float)): sap[str(r[2]).strip()]=sap.get(str(r[2]).strip(),0.0)+r[14]
    except Exception as _e_: _swallow(5515,_e_); pass
    for b in blocks:
        m=b['mat']; cnt=None; date=None
        if m in COUNTED_RM: cnt=COUNTED_RM[m]['qty']; date=COUNTED_RM[m].get('date')
        elif b.get('count_stock') and m in pc: cnt=float(b.get('st_main') if b.get('st_main') is not None else b['stock']); date=pc[m].get('date')
        e=dict(sap=sap.get(m),count=cnt,date=date,days=None,use_day=None,use=0.0,rec=0.0,est=None,src='SAP (MB52) - not counted' if cnt is None else 'Physical count')
        if cnt is not None and date is not None:
            days=max(0,(d0-pd.Timestamp(date).normalize()).days)
            P=_plan_without_oem(b); ud=(sum(P[:7])/min(7,len(P))) if P else 0.0
            rc=gr_between(m,pd.Timestamp(date).normalize(),d0)   # GR after the count day (GR. Date, Excel number read as a date - was Final AtA)
            e.update(days=days,use_day=ud,use=ud*days,rec=rc)
            if days>0:
                est=max(0.0,cnt-ud*days+rc); delta=est-cnt
                b['stock']=float(b['stock'])+delta
                if b.get('st_main') is not None: b['st_main']=float(b['st_main'])+delta
                e.update(est=est,src=f'Physical count {pd.Timestamp(date):%d-%b}, rolled forward {days} d')
            else: e.update(est=cnt,src='Physical count today')
        if cnt is None and mis_age>=1 and not b.get('count_stock'):   # SAP stock from an old MIS file: roll forward like a count
            own=float(b['stock'])-float(b.get('st_jw') or 0)
            P=_plan_without_oem(b); ud=(sum(P[:7])/min(7,len(P))) if P else 0.0
            if own>0 and ud>0:
                est=max(0.0,own-ud*mis_age); delta=est-own
                b['stock']=float(b['stock'])+delta
                if b.get('st_main') is not None: b['st_main']=float(b['st_main'])+delta
                e.update(date=mis_d,days=mis_age,use_day=ud,use=ud*mis_age,est=est,src=f'SAP (MB52) {mis_d:%d-%b}, rolled forward {mis_age} d')
        STOCK_EST[m]=e

LT_OF={}
def alt_bom_substitutes(A, DS, mis_path, pp_rows, nd):
    """Your alternative BOM as line substitutes for 'Line can run': for each standard material, its approved alternative is used
       (after the standard one) on the lines that run the customers it is approved for.
       Alternative that is itself a planned material -> only its FREE stock (lowest projected balance of its own plan, never needed by it);
       otherwise -> its MB52 + JW stock and MIS arrivals. Codes already counted inside the standard material (roll lengths, your mapping) are skipped."""
    out={}
    if not ALT_BOM: return out
    S_={x['b']['mat']:x.get('S') or [] for x in A}
    LT_OF.clear(); LT_OF.update({x['b']['mat']:x['b'].get('lt_used') for x in A})
    fac=DS.get('alias_factor') or {}
    cl={}
    for x_ in pp_rows or []:
        if any(x_['v']): cl.setdefault(_norm(x_['cust']),set()).add(_line_key(x_['line']))
    stock={}; arr={}; d0=pd.Timestamp(DS['dates'][0]).normalize()
    mis=wb_ro(mis_path)
    for r in mis['MB52'].iter_rows(min_row=2,values_only=True):
        c=str(r[2] or '').strip()
        if c and str(r[13] or '').strip().upper() in (DS['plant'].upper(),DS['plant'].upper()+' WAREHOUSE') and isinstance(r[14],(int,float)): stock[c]=stock.get(c,0.0)+r[14]
    for r in mis['RM Status for OEM JW'].iter_rows(min_row=2,values_only=True):
        if r and r[1] and isinstance(r[11],(int,float)): stock[str(r[1]).strip()]=stock.get(str(r[1]).strip(),0.0)+r[11]
    for r in list(mis['In Transit'].iter_rows(values_only=True))[1:]:
        if len(r)>39 and r[39] and isinstance(r[34],(int,float)) and r[34]>0 and it_ok(r):
            try: k=max(0,(pd.Timestamp(r[35]).normalize()-d0).days)
            except Exception as _e_: _swallow(5566,_e_); continue
            if k<nd: arr.setdefault(str(r[39]).strip(),[0.0]*nd)[k]+=r[34]
    for m,alts in ALT_BOM.items():
        for c,custs in alts.items():
            if c in fac: continue                                  # already counted in the standard material
            lines=set()
            for cu in custs: lines|=cl.get(_norm(cu),set())
            if not lines: continue
            if c in S_ and S_[c]:                                  # planned itself: only what its own plan does not need before a new order could arrive
                lt_=int(LT_OF.get(c) or 50); base=max(0.0,min(S_[c][:max(1,lt_)+1])); a_=[0.0]*nd
            else: base=max(0.0,stock.get(c,0.0)); a_=arr.get(c,[0.0]*nd)
            if base<=0 and not any(a_): continue
            out.setdefault(m,[]).append(dict(code=c,factor=1.0,base=base,arr=a_,only=lines,win=(0,nd-1),order=9,pri={lk:9 for lk in lines},alt_bom=True))
    return out

def allocate_lines(DS, pri, alt, cmap, line_rank, nd, mat_rank=None):
    """Allocates every material's stock + in-transit to the lines that need it, day by day.
       Customer Supplier = which KNOWN vendors a customer is ALLOWED to use (only those; blank = any vendor). Stock of unknown vendor
       ('Unassigned' / other stock) is not covered by Customer Supplier: Vendor Priority decides (0 = never, blank = after the rest). Within what is allowed: Vendor Priority by vendor (1, 2, 3), then the rest, then alternative materials
       by description (1, 2, 3). 0 on Vendor Priority = that line never uses it.
       Lines are served in order of POTENTIAL STOPPAGE: the line that would stop first (pooled view) is served first, and so on.
       Returns {(material, line): first day the line cannot be served} and {(material, line): supply order text}."""
    stop={}; used={}; mat_rank=mat_rank or {}; ALLOC_IT.clear()
    for b in DS['blocks']:
        need={}
        for (cu,ln),arr in (b.get('cust_line') or {}).items():
            if ln and any(arr): need[(cu,_line_key(ln))]=arr
        if not need: continue
        lines_=sorted({lk for _,lk in need})
        if len(lines_)<2 and not b.get('sup') and not (alt or {}).get(b['mat']): continue   # one line, one source: pooled answer is exact
        P_=pri.get(b['mat']) or {}
        src={}; vend=[]
        inflow=[x+y for x,y in zip(b['transit'],b['adhoc'])]
        if b.get('sup'):
            for sp in b['sup']:
                nm=sp['name'].upper(); src[nm]=dict(base=sp['stock'],arr=list(sp['transit']),sp=sp); vend.append(nm)
            rest_b=b['stock']-sum(v['base'] for v in src.values()); rest_a=[a-sum(v['arr'][k] for v in src.values()) for k,a in enumerate(inflow)]
            if rest_b>0.5 or any(x>0.5 for x in rest_a): src['OTHER STOCK']=dict(base=max(0.0,rest_b),arr=[max(0.0,x) for x in rest_a])
        else: src[b['mat']]=dict(base=b['stock'],arr=inflow)
        alts=[]
        for e in (alt or {}).get(b['mat'],[]):
            if e['code']==b['mat']: alts.append((e,b['mat'])); continue
            src[e['code']]=dict(base=e['base']*e['factor'],arr=[x*e['factor'] for x in e['arr']],win=e.get('win')); alts.append((e,e['code']))
        def order(cu,lk):
            excl={v for v,lns in P_.items() if lns.get(lk)==0}|{c for e,c in alts if e['pri'].get(lk)==0 or (e.get('only') and lk not in e['only'])}
            allowed=cmap.get((_norm(cu),b['mat']))       # Customer Supplier: this customer may ONLY use these vendors (blank = any vendor)
            if allowed: excl|={v for v in vend if v not in allowed and v not in ('UNASSIGNED',)}   # stock of unknown vendor: Vendor Priority decides
            o=[]
            for _,v in sorted((n,v) for v,lns in P_.items() if v in src and isinstance(lns.get(lk),int) and lns.get(lk)>0):   # Vendor Priority 1,2,3
                if v not in excl and v not in o: o.append(v)
            for v in (allowed or [])+vend+['OTHER STOCK',b['mat']]:                        # then the remaining allowed stock
                if v in src and v not in excl and v not in o and v not in {c for _,c in alts if c!=b['mat']}: o.append(v)
            for _,c in sorted(((e['pri'].get(lk) or 9),c) for e,c in alts):               # 4 alternative materials by description
                if c in src and c not in excl and c not in o: o.append(c)
            return o
        ords={key:order(*key) for key in need}
        altc_={c for _,c in alts if c!=b['mat']}
        for lk_ in {lk for _,lk in need}:                    # in transit this line may use (allowed suppliers only, alternatives not included)
            ok_=set()
            for (cu_,l2_),o_ in ords.items():
                if l2_==lk_: ok_|={v for v in o_ if v in src and v not in altc_}
            ALLOC_IT[(b['mat'],lk_)]=sum(sum(src[v]['arr']) for v in ok_)
        bal={v:s_['base'] for v,s_ in src.items()}; draw={v:[0.0]*nd for v in src}
        rk_=mat_rank.get(b['mat']) or {}
        keys=sorted(need,key=lambda key:(rk_.get(key[1],line_rank.get(key[1],999)),key[1],key[0]))   # your ranking first, then most urgent line
        for k in range(nd):
            for v,s_ in src.items(): bal[v]+=s_['arr'][k]
            for key in keys:
                q=need[key][k]
                for v in ords[key]:
                    if q<=1e-9: break
                    w_=src[v].get('win')
                    if w_ and not (w_[0]<=k<=w_[1]): continue                 # substitute only inside its dates
                    t_=min(q,max(0.0,bal[v])); bal[v]-=t_; draw[v][k]+=t_; q-=t_
                if q>1e-6 and (b['mat'],key[1]) not in stop: stop[(b['mat'],key[1])]=k
        for key in need:
            stop.setdefault((b['mat'],key[1]),nd)
            used.setdefault((b['mat'],key[1]),' > '.join(ords[key]) or '-')
            used[(b['mat'],_norm(key[0]),key[1])]=' > '.join(ords[key]) or '-'      # per customer and line (MRP Daily line rows)
        for v,s_ in src.items():
            if s_.get('sp') is not None:
                s_['sp']['plan']=draw[v]; s_['sp']['plan_label']='Plan (allocated to lines)'
                s_['sp']['cust_text']='Lines: '+', '.join(sorted({lk for (cu,lk) in need if v in ords[(cu,lk)]}))
    return stop,used

def line_run_days(A, d0, nd, alloc_stop=None):
    """How many days each production line can keep running with the material it needs.
       For every material the line uses (plan rows with that line), the line stops on the first day the material's projected stock
       (all lines together) is below zero AND the line needs that material that day. Line run days = the earliest such stop."""
    lines={}
    for x in A:
        b=x['b']
        for (cu,ln),arr in (b.get('cust_line') or {}).items():
            if not ln or not any(arr): continue
            key=_line_key(ln); L_=lines.setdefault(key,dict(name=ln,mats=[],cust=set(),mod=0.0)); L_['cust'].add(cu)
            for scen,stk in (('t',x['S']),('o',[a+c for a,c in zip(x['S'],x['cO'])])):
                k=next((k for k in range(nd) if stk[k]<0 and arr[k]>0),None)
                if alloc_stop and (b['mat'],key) in alloc_stop and scen=='t':   # Vendor Priority: the line stops when ITS vendors run out
                    k=alloc_stop[(b['mat'],key)]; k=None if k>=nd else k
                L_['mats'].append((scen,k if k is not None else nd,b['common'] if not b.get('count_stock') else b['desc'],b['mat']))
                if not b.get('count_stock'): L_['mats'].append(('r'+scen,k if k is not None else nd,b['common'],b['mat']))   # raw materials only
    out=[]
    for key,L_ in lines.items():
        res={}
        for scen in ('t','o','rt'):
            ms=sorted([(k,n,c) for s_,k,n,c in L_['mats'] if s_==scen])
            k=ms[0][0] if ms else nd
            lim=list(dict.fromkeys(n for kk,n,c in ms if kk==k and k<nd))
            nxt=list(dict.fromkeys(f'{n} ({kk} d)' for kk,n,c in ms if kk>k and kk<nd and n not in lim))[:3]
            res[scen]=(k,lim,nxt)
        out.append(dict(key=key,name=L_['name'],cust=sorted(L_['cust']),t=res['t'],o=res['o'],rt=res['rt']))
    out.sort(key=lambda r:(r['t'][0],r['key']))
    return out


def pack_size_scan(DS, mis_path, user_mrp=None):
    """Reads every material description (MIS Common Name, MB52, In Transit, your BOM / Stock Count) and finds codes that look like the
       SAME material in another pack size or unit (roll length, drum / cartridge, spool...). Returns rows for the Pack Size Check sheet."""
    import itertools
    mis=wb_ro(mis_path); desc={}; unit={}; src={}; stock={}; intr={}; cn={}
    def put(c,d,u,tag):
        c=str(c or '').strip()
        if not c or not d: return
        desc.setdefault(c,str(d)); src.setdefault(c,set()).add(tag)
        if u: unit.setdefault(c,str(u))
    for r in mis['Common Name'].iter_rows(min_row=2,values_only=True):
        if r and r[0]: put(r[0],r[1],r[2],'MIS'); cn.setdefault(str(r[0]).strip(),r[3])
    for r in mis['MB52'].iter_rows(min_row=2,values_only=True):
        put(r[2],r[5],r[7],'MB52')
        if str(r[13] or '').strip().upper() in (DS['plant'].upper(),DS['plant'].upper()+' WAREHOUSE') and isinstance(r[14],(int,float)):
            stock[str(r[2]).strip()]=stock.get(str(r[2]).strip(),0.0)+r[14]
    for r in list(mis['In Transit'].iter_rows(values_only=True))[1:]:
        if len(r)>39 and r[39]:
            put(r[39],r[4],None,'In transit')
            if isinstance(r[34],(int,float)) and r[34]>0: intr[str(r[39]).strip()]=intr.get(str(r[39]).strip(),0.0)+r[34]
    if user_mrp and os.path.exists(user_mrp):
        mw=wb_ro(user_mrp)
        if 'BOM' in mw.sheetnames:
            for r in mw['BOM'].iter_rows(min_row=2,values_only=True):
                if r and len(r)>7: put(r[4],r[5],r[7],'Your BOM')
    PACK=r'(?<![\d.])(\d+(?:\.\d+)?)\s*(M|MTR|MTRS|METER|METRE|ML|LTR|L|KG|KGS|PCS|NOS)(?![A-Z])'
    def parse(d):
        t=' '.join(str(d).upper().split())
        packs=[f'{a}{b}' for a,b in re.findall(PACK,t)]
        if re.search(r'\bDRUM\b',t): packs.append('DRUM')
        if re.search(r'\bSPOOL\b',t): packs.append('SPOOL')
        rest=re.sub(PACK,' ',t); rest=re.sub(r'\b(T\d+|DRUM|SPOOL)\b',' ',rest)
        return {w for w in re.split(r'[^A-Z0-9]+',rest) if w and w not in ('X','FOR','THE','OF','AND','WITH','G2G','MODULE','GG','SOLAR')},packs
    info={c:parse(d) for c,d in desc.items() if c.upper()[:2] in ('RM','PM','CS') and not c.upper().startswith('JW')}
    pref=lambda c: re.match(r'[A-Z]+',c).group(0)[:5]
    par={}
    def f(x):
        while par.setdefault(x,x)!=x: x=par[x]
        return x
    for a,b in itertools.combinations(sorted(info),2):
        if pref(a)!=pref(b): continue
        (ta,pa),(tb,pb)=info[a],info[b]
        if ta and tb and len(ta&tb)/len(ta|tb)>=0.5 and pa!=pb and (pa or pb): par[f(a)]=f(b)
    cl={}
    for c in par: cl.setdefault(f(c),[]).append(c)
    blocks={b['mat']:b for b in DS['blocks']}; alias=DS.get('alias') or {}; fac=DS.get('alias_factor') or {}
    rows=[]
    for fam in cl.values():
        if len(fam)<2: continue
        mrp=[c for c in fam if c in blocks]
        if not mrp and not any(stock.get(c) or intr.get(c) for c in fam): continue
        jac=lambda a,b: len(info[a][0]&info[b][0])/max(1,len(info[a][0]|info[b][0]))
        keep=[c for c in fam if c in blocks or stock.get(c) or intr.get(c) or 'Your BOM' in src.get(c,())]   # leave out empty codes
        for c in sorted(keep,key=lambda c:(min((-jac(c,m),m) for m in mrp)[1] if mrp else '',c not in blocks,c)):
            target=c if c in blocks else (alias.get(c) if alias.get(c) in mrp else (max(mrp,key=lambda m:jac(c,m)) if mrp else None))   # nearest MRP material
            ti=tape_info(desc[c]); u=unit.get(c)
            if c==target: how,act='MRP material',''
            elif target and alias.get(c)==target and c in fac: how,act=f'Counted for {target} (x{fac[c]:g})',''
            elif target is None: how,act='Not an MRP material','Information only'
            else:
                same_u=same_unit(u,unit.get(target)); spec_diff=(info[c][0]-info[target][0])|(info[target][0]-info[c][0])
                if same_u and not spec_diff: how,act='Same material, same unit - NOT counted',f'If usable for {target}: link it on the MIS Common Name sheet, or add {c} to PACK_FACTORS with factor 1'
                elif same_u: how,act='Same unit, different spec - NOT counted',f'Spec differs ({", ".join(sorted(spec_diff))[:40]}) - confirm with production if it can replace {target}'
                else: how,act=f'Different unit ({u} vs {unit.get(target)}) - NOT counted',f'Give the conversion (e.g. {u} per {unit.get(target)}) and add {c} to PACK_FACTORS'
            rows.append(dict(family=(blocks[target]['common'] if target in blocks else desc[target]) if target else desc[fam[0]],code=c,desc=desc[c],pack=', '.join(info[c][1]) or (f'{ti[2]:g} m' if ti else '-'),unit=u,
                             stock=stock.get(c,0.0),intr=intr.get(c,0.0),src=', '.join(sorted(src.get(c,()))),cn=cn.get(c),how=how,act=act,target=target))
    return rows

def make_pack_check_xlsx(rows, path):
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    wb=Workbook(); ws=wb.active; ws.title='Pack Size Check'
    ws['A1']='Pack size check - same material under several codes (roll length, drum / cartridge, spool...)'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=('Found by reading every description (MIS Common Name, MB52, In Transit, your BOM). Tapes are converted automatically by roll length. '
              'Other codes are only counted when linked on the MIS Common Name sheet (with an MIS conversion) or listed in PACK_FACTORS at the top of the script.')
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:K2'); ws['A2'].alignment=Alignment(wrap_text=True); ws.row_dimensions[2].height=28
    H=['Material (family)','Code','Description','Pack / size','Unit','Stock (MB52)','In transit (MIS)','Found in','MIS Common Name','How the MRP treats it','What to do']
    for j,h in enumerate(H,1):
        c=ws.cell(4,j,h); c.font=Font(name=AR,size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=NAVY); c.border=bd; c.alignment=Alignment(horizontal='center',wrap_text=True)
    rows=sorted(rows,key=lambda r_:rrank(str((r_.get('mrp') or r_.get('code') or r_.get('mat') or '') if isinstance(r_,dict) else '')))
    for i,r_ in enumerate(rows):
        vals=[r_['family'],r_['code'],r_['desc'],r_['pack'],r_['unit'],r_['stock'],r_['intr'],r_['src'],r_['cn'],r_['how'],r_['act']]
        for j,v in enumerate(vals,1):
            c=ws.cell(5+i,j,round(v,3) if isinstance(v,float) else v); c.font=Font(name=AR,size=9,bold=(r_['how']=='MRP material' and j<=3)); c.border=bd
            c.alignment=Alignment(vertical='top',wrap_text=j in (1,3,10,11))
            if j in (6,7): c.number_format='#,##0.#;-#,##0.#;"-"'
        col={'MRP material':'DDEBF7'}.get(r_['how'],'C6EFCE' if r_['how'].startswith('Counted') else ('FFC7CE' if ('NOT counted' in r_['how'] and r_['stock']>0) else 'FFF2CC'))
        ws.cell(5+i,10).fill=PatternFill('solid',fgColor=col)
    if rows: ws.add_table(Table(displayName='PackSizeCheck',ref=f'A4:K{4+len(rows)}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=True)))
    for j,w in enumerate([30,16,42,12,6,12,12,18,22,32,60],1): ws.column_dimensions[L(j)].width=w
    ws.freeze_panes='C5'; wb.save(path)

def mis_usage_check(mis_path, DS):
    """MIS usage per module used by the MRP (Total Plan Req: requirement / modules) must equal the MIS 'Standard BOM' sheet
       (sum of the material's lines for that BOM Name, e.g. EPE 480 = 2 x 2.55 = 5.1). Returns Check rows for any difference."""
    mis=wb_ro(mis_path); I=[]
    if 'Standard BOM' not in mis.sheetnames: return I
    sb={}
    for r in mis['Standard BOM'].iter_rows(min_row=2,values_only=True):
        if r and len(r)>5 and r[2] and r[3] and isinstance(r[5],(int,float)): k=(str(r[2]).strip(),str(r[3]).strip()); sb[k]=sb.get(k,0.0)+r[5]
    req={}; mods={}; seen=set()
    for r in mis['Total Plan Req'].iter_rows(min_row=2,values_only=True):
        if not tp_ok(r,DS['plant']) or not r[10] or not isinstance(r[14],(int,float)) or not isinstance(r[3],(int,float)) or not r[3]: continue
        k=(str(r[4]).strip(),str(r[10]).strip()); req[k]=req.get(k,0.0)+r[14]
        pl=(k,r[0],r[1],r[2],r[17] if len(r)>17 else None)
        if pl not in seen: seen.add(pl); mods[k]=mods.get(k,0.0)+r[3]
    names={b['mat']:b['common'] for b in DS['blocks']}
    for k,q in req.items():
        if k[1] not in names or not mods.get(k): continue
        rate=q/mods[k]; ref=sb.get(k)
        if ref is None:
            I.append(dict(area='Plan',check='MIS usage not found on the MIS Standard BOM sheet',sev='Medium',material=names[k[1]],code=k[1],a_lab='Total Plan Req usage',a=round(rate,6),b_lab='Standard BOM',b=None,diff=None,
                          detail=f'{k[0]}: Total Plan Req uses {rate:.6g} per module but the Standard BOM sheet has no line',action='Ask MIS to check the Standard BOM'))
        elif abs(rate-ref)>1e-6*max(1.0,ref):
            I.append(dict(area='Plan',check='MIS Total Plan Req usage differs from MIS Standard BOM',sev='High',material=names[k[1]],code=k[1],a_lab='Standard BOM',a=round(ref,6),b_lab='Total Plan Req',b=round(rate,6),diff=round(rate-ref,6),
                          detail=f'{k[0]}: requirement / modules = {rate:.6g} vs Standard BOM {ref:.6g} - MIS plan is not on its own BOM',action='Ask MIS to refresh Total Plan Req'))
    return I

def append_sheet_into(dst, src_path, top, cache=None):
    """Copies the first sheet of an openpyxl-built workbook into worksheet dst from row 'top' (values, formats, merges, tables).
       Returns the next free row. Used to put several sections on one sheet (Check, Supply)."""
    from copy import copy
    from openpyxl.worksheet.table import Table, TableStyleInfo
    sw=load_workbook(src_path).active; off=top-1
    for row in sw.iter_rows():
        for c in row:
            if c.value is None and not c.has_style: continue
            if type(c).__name__=='MergedCell': continue
            v=c.value
            if cache and isinstance(v,str) and v.startswith('=') and (c.row,c.column) in cache: v=cache[(c.row,c.column)]   # value, not formula
            d=dst.cell(c.row+off,c.column,v)
            if c.has_style:
                d.font=copy(c.font); d.fill=copy(c.fill); d.border=copy(c.border); d.alignment=copy(c.alignment); d.number_format=c.number_format
    for mr in sw.merged_cells.ranges:
        dst.merge_cells(start_row=mr.min_row+off,start_column=mr.min_col,end_row=mr.max_row+off,end_column=mr.max_col)
    for r,dim in sw.row_dimensions.items():
        if dim.height: dst.row_dimensions[r+off].height=dim.height
    for k,dim in sw.column_dimensions.items():
        if dim.width and (dst.column_dimensions[k].width or 0)<dim.width: dst.column_dimensions[k].width=dim.width
    for t in sw.tables.values():
        a,b=t.ref.split(':'); ma=re.match(r'([A-Z]+)(\d+)',a); mb=re.match(r'([A-Z]+)(\d+)',b)
        dst.add_table(Table(displayName=t.displayName,ref=f'{ma.group(1)}{int(ma.group(2))+off}:{mb.group(1)}{int(mb.group(2))+off}',
                            tableStyleInfo=TableStyleInfo(name=t.tableStyleInfo.name,showRowStripes=True)))
    return sw.max_row+off+3

def combine_sheets(main_path, extra_paths, out_path, collapse=False, caches=None):
    """main sheet + extra sections below it, saved as out_path (all openpyxl-built temp files)."""
    wb=load_workbook(main_path); ws=wb.active; top=ws.max_row+3
    for p_ in extra_paths:
        if p_ and os.path.exists(p_):
            ws.cell(top-1,1,'_'*60).font=Font(color='BFBFBF')
            start=top; top=append_sheet_into(ws,p_,top,cache=(caches or {}).get(p_))
            if collapse and top-3>start:              # section title stays visible, the rest opens with +
                ws.row_dimensions.group(start+1,top-3,hidden=True,outline_level=1)
                ws.cell(start,1).value=f"{ws.cell(start,1).value}   (+ to open)"
    if collapse: ws.sheet_properties.outlinePr.summaryBelow=False
    wb.save(out_path)

# ===================== Settings sheet (simple config in Excel) =====================
BASIC_ROWS=[('Target days - A items (top value)','TARGET_A','Lean, watched closely'),
            ('Target days - B items','TARGET_B',''),
            ('Target days - C items (low value / no price)','TARGET_C','Cheap to hold'),
            ('Target days - packaging / consumables','TARGET_CSPM','Local supply'),
            ('Extra days for erratic demand (Z)','EXTRA_Z_DAYS','Added to Z items'),
            ('Target days - anything not classified','TARGET_DAYS','Default target'),
            ('Line CRITICAL if a production line can run fewer than (days)','COVER_CRITICAL','Red - escalate (MRP by Line, dashboards)'),
            ('Lead time raw materials (days) - if no ETD history','LEAD_TIME_RM','Dates planned orders'),
            ('Lead time packaging / consumables (days)','LEAD_TIME_CSPM',''),
            ('Daily report: list materials lasting fewer than (days) at the module plan','REPORT_DAYS','Stock + in transit / module plan need per day'),
            ('Look-ahead (days): a line counts for a material if it runs a product using it within','LOOKAHEAD_DAYS','Daily target need + MRP by Line "Needed now"'),
            ('Safety stock (days of need at the module plan)','SAFETY_DAYS','ALERT (formulas) (reorder level)'),
            ('Days of cover: average plan usage over the next (days)','COVER_AVG_DAYS','MRP Daily / Weekly "Days of cover" = projected stock ÷ average usage/day of the next N days'),
            ('Averaged production (need per day): average per production day over the next (weeks)','AVG_PROD_WEEKS','ALERT, MRP by Line, Action list: need per day, days of stock')]
ADV_ROWS=[('ABC: A = top % of value','ABC_A_PCT','80'),('ABC: B = up to % of value','ABC_B_PCT','95 (rest = C)'),
          ('ABC: days of requirement used for the value','ABC_HORIZON','90'),('XYZ: X if weekly variation below','XYZ_X','0.25'),
          ('XYZ: Y if weekly variation below (above = Z)','XYZ_Y','0.5'),('Shortage measured over the next (days)','SHORT_HORIZON','60'),
          ('RELEASE ORDER when an order is due within (days)','RELEASE_WINDOW','7'),('RESCHEDULE OUT when stock is more than x times the target','RESCHED_OUT_X','3'),
          ('Count only WSA lines (Yes / No)','PLAN_WSA_LINES_ONLY','No'),
          ('Inbound: short lead if ETD to plant date is under (days)','SHORT_LEAD','20'),('Inbound: long lead if ETD to plant date is over (days)','LONG_LEAD','75')]
SETTING_ROWS=BASIC_ROWS+ADV_ROWS
SETTING_OLD_LABELS={'escalate if a line / material can run fewer than (days)':'COVER_CRITICAL','expedite if fewer than (days)':'COVER_RISK',   # labels of older
                    'watch if fewer than (days)':'COVER_WATCH','show on today if short of target by more than (days)':'SHORT_WATCH'}   # Settings sheets still read
SUBS_X=[]          # substitutes of this run, one per substitute code ('A or B' split), resolved to material codes
SUBS=[]            # substitutes from the Settings sheet: dict(mrp, sub, factor, lines(set or None), fr, to, order) - text as typed
def _date(x):
    try: return pd.Timestamp(x).normalize() if x not in (None,'') else None
    except Exception as _e_: _swallow(5868,_e_); return None
def read_settings(paths):
    """Settings sheet: from your MRP.xlsx if it has one, else from the previous MRP_Daily. Sets the rules for this run.
       Sections: basic numbers, unlimited materials, allowed vendors (by customer), substitutes, line ranking, advanced numbers.
       Old sections (Customer Supplier, Pack conversions, Vendor Priority) are still read and turned into the new ones."""
    g=globals(); src=None
    for pth in paths:
        if not pth or not os.path.exists(pth): continue
        try: w=wb_ro(pth)
        except Exception as _e_: _swallow(5877,_e_); continue
        if 'Settings' not in w.sheetnames: continue
        rows=[list(r)+[None]*8 for r in w['Settings'].iter_rows(values_only=True)]
        key={a.lower():b for a,b,_ in SETTING_ROWS}; key.update(SETTING_OLD_LABELS)
        unl=set(); words=[]; rank={}; rmat={}; lhdr=None; cs_=[]; subs=[]; sec=None; vhdr=None; lo={}; ltg={}; bsp={}; ct_=[]
        for r in rows:
            a=str(r[0] or '').strip(); al=a.lower()
            if al.startswith('line ranking'): lhdr=None
            for t,sc in (('mis link','k'),('unlimited materials','u'),('allowed vendors','c'),('customer supplier','c'),('substitutes','s'),('pack conversions','p'),
                         ('line ranking','l'),('vendor priority','v'),('batch supplier codes','b'),('customer daily target','t'),('advanced','x'),('basic','x')):
                if al.startswith(t): sec=sc; vhdr=None; break
            else:
                if 'daily target' in al and isinstance(r[1],(int,float)) and not isinstance(r[1],bool):   # 'USA Line 1 Production Plan Daily Target'
                    ltg[_line_key(re.split(r'production plan|daily target',a,flags=re.I)[0])]=(a,float(r[1])); continue
                if al in key and r[1] not in (None,''):
                    nm=key[al]
                    if nm=='PLAN_WSA_LINES_ONLY': g[nm]=str(r[1]).strip().lower() in ('yes','y','true','1')
                    else:
                        try: g[nm]=(float(r[1]) if nm in ('XYZ_X','XYZ_Y') else int(float(r[1])))
                        except (TypeError,ValueError): pass
                    continue
                if not a and not any(r[1:6]): continue
                if sec=='t':
                    if a and not al.startswith('customer') and isinstance(r[2] if len(r)>2 else None,(int,float)):
                        ct_.append((a,str(r[1] or '').strip(),float(r[2]),'Yes' if str(r[3] if len(r)>3 else '').strip().lower() in ('yes','y','oem','x') else 'No'))
                    continue
                if sec=='b':
                    if a and not al.startswith('batch code') and r[1] not in (None,''): bsp[a.upper()[:2]]=str(r[1]).strip()
                    continue
                if sec=='u' and a and not al.startswith(('code','material code')):
                    (unl.add(a) if re.match(r'^[A-Z]{2}[A-Z0-9]{6,}$',a.upper()) else words.append(al))
                elif sec=='c' and al!='customer' and a:
                    cs_.append({'Customer':a,'Material':r[1],'Code':r[2],'Supplier':r[3],'Share %':None})
                elif sec=='s' and not al.startswith('mrp material') and (a and r[1]):
                    lines=None if r[3] in (None,'') else {_line_key(x) for x in re.split(r'[,;/]',str(r[3])) if x.strip()}
                    try: f=float(r[2]) if r[2] not in (None,'') else 1.0
                    except (TypeError,ValueError): f=1.0
                    subs.append(dict(mrp=a,sub=str(r[1]),factor=f,lines=lines,fr=_date(r[4]),to=_date(r[5]),order=r[6] if isinstance(r[6],(int,float)) else 1))
                elif sec=='p' and a and not al.startswith('code') and r[1] and r[2] not in (None,''):     # old: pack conversions
                    try: subs.append(dict(mrp=str(r[1]),sub=a,factor=float(r[2]),lines=None,fr=_date(r[3]),to=_date(r[4]),order=1))
                    except (TypeError,ValueError): pass
                elif sec=='k' and al.startswith('production plan customer'):
                    hk_=[' '.join(str(x or '').split()).lower() for x in r]; PROJ_COL[0]=hk_.index('project name') if 'project name' in hk_ else None
                elif sec=='k' and a and PROJ_COL[0] is not None and PROJ_COL[0]<len(r) and r[PROJ_COL[0]] and (PROJECT_OF.__setitem__(_norm(a),' '.join(str(r[PROJ_COL[0]]).split())) or r[4] in (None,'')): pass
                elif sec=='k' and a and not al.startswith('production plan customer') and r[4] not in (None,''):
                    lo[_norm(a)]=(str(r[5]).strip() if r[5] not in (None,'') else None,str(r[4]).strip())
                elif sec=='l':
                    if al.startswith('material'): lhdr=r; continue
                    if lhdr is not None:                         # matrix: one row per material, one column per line
                        rk={}
                        for j in range(3,len(lhdr)):
                            if lhdr[j] and isinstance(r[j],(int,float)): rk[_line_key(lhdr[j])]=float(r[j])
                        if rk:
                            if al.startswith('all materials'): rank.update(rk)
                            else: rmat[str(r[2] or a or r[1]).strip()]=dict(rk,_name=a,_desc=r[1])
                    elif a and not al.startswith('production line') and r[1] not in (None,''):   # old format: line | rank
                        try: rank[_line_key(a)]=float(r[1])
                        except (TypeError,ValueError): pass
                elif sec=='v':                                                                         # old: vendor priority matrix
                    if al=='material': vhdr=r; continue
                    if vhdr is not None and (a or r[2]) and not r[1]:      # rows without vendor -> substitutes per line
                        for j in range(3,len(vhdr)):
                            if vhdr[j] and isinstance(r[j],(int,float)) and r[j]>0:
                                subs.append(dict(mrp=a,sub=str(r[2] or a),factor=1.0,lines={_line_key(vhdr[j])},fr=None,to=None,order=r[j]))
        g['UNLIMITED_CODES']=unl; g['UNLIMITED_WORDS']=tuple(words); g['LINE_RANK']=rank; g['LINE_RANK_MAT']=rmat
        LINE_TARGET.clear(); LINE_TARGET.update(ltg); BATCH_SUP.clear(); BATCH_SUP.update(bsp)
        # no CUSTOMER DAILY TARGET section: each customer takes its LINE's daily target from Settings (e.g. USA Line 1 = 4,000)
        CUST_TARGET[:]=ct_ or [(c_,l_,((ltg.get(_line_key(l_)) or (None,None))[1] or t_),o_) for c_,l_,t_,o_ in CUST_TARGET_DEFAULT]
        g['SET_CUSTSUP']=[x for x in cs_ if x.get('Customer')]; g['SUBS']=subs; LINK_OVERRIDE.clear(); LINK_OVERRIDE.update(lo); src=pth; break
    return src

def resolve_substitutes(mis_path):
    """Turns the Substitutes table (code OR description, as you typed it) into material codes, using the MIS descriptions.
       Substitutes for ALL lines become pack conversions (counted in the material's stock, with dates); substitutes for
       some lines only are used when sharing stock out to those lines. Returns Check rows for anything not found."""
    g=globals(); iss=[]; mis=wb_ro(mis_path); desc_of={}; dash={}; names={}
    for r in mis['Daily Dashboard'].iter_rows(min_row=5,values_only=True):
        if len(r)>6 and str(r[6] or '').strip()=='Plan' and r[4]: dash[str(r[4]).strip()]=f"{r[3]} {r[5] or ''}"; names[str(r[4]).strip()]=str(r[3] or '').strip()
    for r in mis['Common Name'].iter_rows(min_row=2,values_only=True):
        if r and r[0] and r[1]: desc_of.setdefault(str(r[0]).strip(),f"{r[1]} {r[3] or ''}")
    for r in mis['MB52'].iter_rows(min_row=2,values_only=True):
        if r[2] and r[5]: desc_of.setdefault(str(r[2]).strip(),str(r[5]))
    tok=lambda t: {w_ for w_ in re.split(r'[^A-Z0-9.]+',str(t or '').upper()) if w_ and w_ not in ('MM','X','FOR','THE','OF','GG','G2G','WITH')}
    def find(text,pool):
        t=str(text or '').strip()
        if t.upper() in {c.upper() for c in pool}: return next(c for c in pool if c.upper()==t.upper()),1.0       # a code
        for c,nm in names.items():                                                                              # an MRP material name
            if c in pool and nm.upper()==t.upper(): return c,1.0
        tt=tok(t)
        def sc_(kv):          # best of: words of the description, or how much of what you typed is found in it
            tk=tok(kv[1]); return max(len(tt&tk)/max(1,len(tt|tk)), len(tt&tk)/max(1,len(tt))*0.9)
        best=max(pool.items(),key=sc_,default=(None,''))
        sc=sc_(best) if best[0] else 0
        return (best[0],sc) if sc>=0.5 else (None,sc)
    for c_,n_,*_r in ADDED_MATS: dash.setdefault(str(c_).strip(),str(n_ or ''))      # materials added because the plan needs them are MRP materials too
    pf={}; lines=[]; SUBS_X[:]=[]
    exp_=[]                                                                      # 'A or B' / 'A; B' = several substitutes in one cell
    for x in SUBS:
        parts_=[p_.strip() for p_ in re.split(r'\s+or\s+|\s*;\s*|\s*\|\s*',str(x['sub'] or ''),flags=re.I) if p_.strip()]
        for p_ in (parts_ or [x['sub']]): exp_.append(dict(x,sub=p_))
    for x in exp_:
        m,sm=find(x['mrp'],dash); c,sc=find(x['sub'],{**desc_of,**dash})
        if not m or not c or c==m:
            iss.append(dict(area='Master data',check='Substitute not understood (Settings)',sev='Medium',material=str(x['mrp']),code=str(x['sub']),a_lab='',a=None,b_lab='',b=None,diff=None,
                            detail=f"MRP material '{x['mrp']}' -> {m or 'not found'}; substitute '{x['sub']}' -> {c or 'not found'}",action='Use the code, or more words from the description')); continue
        x['mrp_code'],x['sub_code']=m,c; x['sub_desc']=desc_of.get(c,dash.get(c,'')); SUBS_X.append(x)
        iss.append(dict(area='Master data',check='Substitute read from Settings (check it is the right material)',sev='Low',material=f"{x['mrp']} -> {m}",code=c,a_lab='',a=None,b_lab='Factor',b=x['factor'],diff=None,
                        detail=f"'{x['sub']}' found as {c} {x['sub_desc'][:50]}; "+('all lines' if x['lines'] is None else 'lines: '+', '.join(sorted(x['lines'])))+
                               (f"; from {x['fr']:%d-%b}" if x['fr'] is not None else '')+(f" until {x['to']:%d-%b}" if x['to'] is not None else ''),action='If wrong, type the code instead'))
        if x['lines'] is None: pf[c]=(m,x['factor'],x['fr'],x['to'])
        else: lines.append(x)
    g['PACK_FACTORS']={**_PACK_FACTORS_SCRIPT,**pf}; g['SUBS_LINES']=lines      # script entries kept (review finding Nov 2026)
    return iss

def make_settings_sheet(path, src=None):
    from openpyxl.styles import Border, Side
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    wb=Workbook(); ws=wb.active; ws.title='Settings'; ws.sheet_view.showGridLines=False
    ws['A1']='Settings - all planning rules in one place. Type in the yellow cells, save, run the script'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=('Tip: copy this sheet into your MRP.xlsx once (right-click the tab > Move or Copy > Create a copy) and keep editing it there. Leave a rule blank when there is no real constraint.'
              +(f'  Read this run from: {os.path.basename(src)}' if src else '  (no Settings sheet found - defaults)'))
    ws['A2'].font=Font(name=AR,size=9,italic=True)
    yel=PatternFill('solid',fgColor='FFF2CC'); hf=Font(name=AR,size=10,bold=True,color='FFFFFF'); hfill=PatternFill('solid',fgColor=NAVY)
    def title(r,t,sub=''):
        ws.cell(r,1,t).font=Font(name=AR,size=12,bold=True,color=NAVY)
        if sub: ws.cell(r+1,1,sub).font=Font(name=AR,size=9,italic=True,color='595959'); return r+2
        return r+1
    def head(r,cols):
        for j,t in enumerate(cols,1): c=ws.cell(r,j,t); c.font=hf; c.fill=hfill; c.border=bd; c.alignment=Alignment(wrap_text=True,vertical='center')
        return r+1
    def numrows(r,rows_):
        r=head(r,['Setting','Value','Note'])
        for lab,nm,expl in rows_:
            v=globals()[nm]; v=('Yes' if v else 'No') if isinstance(v,bool) else v
            ws.cell(r,1,lab).border=bd; c=ws.cell(r,2,v); c.fill=yel; c.border=bd; c.alignment=Alignment(horizontal='center'); ws.cell(r,3,expl).border=bd; r+=1
        return r
    def table(r,cols,data,yellow_cols,n_blank=5,date_cols=()):
        r=head(r,cols)
        for x in list(data)+[[None]*len(cols)]*n_blank:
            for j in range(len(cols)):
                v=x[j] if j<len(x) else None
                if isinstance(v,pd.Timestamp): v=v.to_pydatetime()
                c=ws.cell(r,j+1,v); c.border=bd
                if j in yellow_cols or not any(x): c.fill=yel
                if j in date_cols: c.number_format='dd-mmm-yy'
            r+=1
        return r
    r=4
    r=title(r,'BASIC SETTINGS'); r=numrows(r,BASIC_ROWS)
    tg_=sorted(LINE_TARGET.items()) or [(f'L{i}',(f'USA Line {i} Production Plan Daily Target',None)) for i in range(1,5)]
    for _,(lab,v) in tg_:                  # modules per day each line should make - used for 'days stock + in transit lasts' and the daily report
        ws.cell(r,1,lab).border=bd; c=ws.cell(r,2,v); c.fill=yel; c.border=bd; c.alignment=Alignment(horizontal='center')
        ws.cell(r,3,'Modules per day').border=bd; r+=1
    r+=1
    r=title(r,'UNLIMITED MATERIALS','Never short - one material code, or a word from the description, per row')
    r=table(r,['Code or description word'],[[v] for v in sorted(UNLIMITED_CODES)+list(UNLIMITED_WORDS)],[0],3)+1
    r=title(r,'ALLOWED VENDORS (by customer)','A customer may use ONLY the vendors listed for a material (several rows = several vendors, top row preferred). '
            'No row / blank Supplier = any vendor. Stock of unknown vendor is always allowed.')
    r=table(r,['Customer','Material','Code','Supplier'],[[x.get('Customer'),x.get('Material'),x.get('Code'),x.get('Supplier')] for x in CUSTSUP_ROWS],[3])+1
    r=title(r,'SUBSTITUTES','Another material that can replace an MRP material. Type the CODE or words from the DESCRIPTION (the script finds it). '
            'Lines blank = all lines (counted in the material stock); otherwise only for the lines listed (e.g. USA Line 1, USA Line 2). Dates blank = always.')
    r=table(r,['MRP material (code or description)','Substitute (code or description)','Factor (MRP units per 1 unit)','Lines (blank = all)','Use from','Use until','Order (1 = first)'],
            [[x['mrp'],x['sub'],x['factor'],', '.join(sorted(x['lines'])) if x['lines'] else None,x['fr'],x['to'],x['order']] for x in SUBS],[1,2,3,4,5,6],5,(4,5))+1
    r=title(r,'CUSTOMER DAILY TARGET','Daily need of a material = sum over these customers of: daily target x usage per module (MIS Standard BOM / your BOM for CS-PM), '
            'like your PPC Usage sheet. OEM = Yes counts only in the "with OEM" columns.')
    r=table(r,['Customer','Line','Daily target (modules)','OEM (Yes / No)'],[[c_,l_,v_,o_] for c_,l_,v_,o_ in (CUST_TARGET or CUST_TARGET_DEFAULT)],[3],3)+1
    r=title(r,'BATCH SUPPLIER CODES','The 2nd and 3rd characters of an SAP batch tell the supplier (1JU0369004 -> JU). Type / correct the supplier name in column B. '
            '"(guess)" = the script guessed it from MIS In Transit suppliers of the same materials - overwrite it with the right name')
    bc_=sorted(set(BATCH_SUP)|set(BATCH_SEEN))
    r=table(r,['Batch code (2 letters)','Supplier name','Seen on materials (MB52 batches)'],
            [[k,BATCH_SUP.get(k) or ((BATCH_SEEN.get(k) or {}).get('guess') and f"{BATCH_SEEN[k]['guess']} (guess)"),
              ', '.join(sorted((BATCH_SEEN.get(k) or {}).get('mats') or []))[:120] or None] for k in bc_],[1],2)+1
    r=title(r,'LINE RANKING','Which line gets a scarce material first: 1 = first, 2 = second ... per material row (only materials used by 2 or more lines are listed). '
            'Row "All materials" = default for every material. Blank = closest to stopping first')
    lines_=list(LINES_SEEN) or sorted({_line_key(k) for v in LINE_RANK_MAT.values() for k in v if not k.startswith('_')})
    lk_=[_line_key(n) for n in lines_]
    data=[['All materials (default)','','']+[LINE_RANK.get(k) for k in lk_]]
    seen_=set()
    for nm,ds,cd,lns in SHARED_MATS:
        rk=LINE_RANK_MAT.get(cd) or next((v for k,v in LINE_RANK_MAT.items() if k.upper() in (str(nm).upper(),str(ds).upper())),{})
        data.append([nm,ds,cd]+[(rk.get(k) if k in lns else '-') for k in lk_]); seen_.add(cd)
    for k,v in LINE_RANK_MAT.items():                    # rows you typed for materials not shared this run - kept
        if k not in seen_ and v.get('_name') is not None and k not in {d[2] for d in data}: data.append([v.get('_name'),v.get('_desc'),k]+[v.get(x) for x in lk_])
    r=table(r,['Material (common name)','Component description','Code']+lines_,data,list(range(3,3+len(lines_))),2)+1
    r=title(r,'MIS LINK (Production Plan customer -> MIS plan / BOM)','How each customer of your Production Plan is linked to the MIS plan (for raw-material usage). '
            'Found by the script: same day + same production line in both plans (no names needed). Only if it is wrong: type the MIS BOM Name (and MIS customer) in the yellow cells')
    lk_rows=[[cu,L_.get('mis'),L_.get('bom'),(f"{L_['how']}"+(f" ({L_['conf']:.0%})" if L_.get('conf') is not None else '')),
              (link_override(cu,quiet=True) or (None,None))[1],(link_override(cu,quiet=True) or (None,None))[0]] for cu,L_ in sorted(LINKS.items())]
    r=table(r,['Production Plan customer','MIS customer (found)','MIS BOM Name (found)','How it was found','Override: MIS BOM Name','Override: MIS customer'],lk_rows,[4,5],2)+1
    r=title(r,'ADVANCED SETTINGS','Leave as they are unless you have a reason'); r=numrows(r,ADV_ROWS)
    for col,w in zip('ABCDEFG',[46,40,16,22,12,12,12]): ws.column_dimensions[col].width=w
    for j in range(8,16): ws.column_dimensions[L(j)].width=12
    set_print(ws,'Settings'); wb.save(path)

def segment(DS, mis_path, H=None):
    """ABC by value (next H days of gross requirement x unit price from MIS In Transit 'Unit Rate'), XYZ by demand variability
       (coefficient of variation of weekly requirement over the next 13 weeks). Sets b['abc'], b['xyz'], b['value'], b['price'], b['td']."""
    H=H or ABC_HORIZON
    mis=wb_ro(mis_path); alias=DS.get('alias') or {}; fac=DS.get('alias_factor') or {}
    num={}; den={}
    for r in list(mis['In Transit'].iter_rows(values_only=True))[1:]:
        if len(r)<40 or not isinstance(r[8],(int,float)) or r[8]<=0: continue
        c=str(r[39] or '').strip(); q=r[34] if isinstance(r[34],(int,float)) and r[34]>0 else (r[7] if isinstance(r[7],(int,float)) and r[7]>0 else 1.0)
        f=1.0
        if c in alias: f=fac.get(c,1.0) or 1.0; c=alias[c]
        num[c]=num.get(c,0.0)+(r[8]/f)*q*f; den[c]=den.get(c,0.0)+q*f      # price per MRP unit, weighted by quantity
    rows=[]
    for b in DS['blocks']:
        b['price']=(num[b['mat']]/den[b['mat']]) if den.get(b['mat']) else None
        P=b['plan']; b['value']=(sum(P[:H])*b['price']) if b['price'] else None
        wk=[sum(P[i:i+7]) for i in range(0,min(len(P),91),7)]
        m_=sum(wk)/len(wk) if wk else 0
        cv=((sum((x-m_)**2 for x in wk)/len(wk))**0.5/m_) if m_ else None
        b['cv']=cv; b['xyz']='-' if cv is None else ('X' if cv<XYZ_X else 'Y' if cv<XYZ_Y else 'Z')
    rm=[b for b in DS['blocks'] if not b.get('count_stock') and b['value']]
    tot=sum(b['value'] for b in rm) or 1; acc=0
    for b in sorted(rm,key=lambda b:-b['value']):
        acc+=b['value']; b['abc']='A' if acc-b['value']<ABC_A_PCT/100*tot else ('B' if acc-b['value']<ABC_B_PCT/100*tot else 'C')
    for b in DS['blocks']:
        if b.get('count_stock'): b['abc']='PM'; base=TARGET_CSPM
        else:
            b.setdefault('abc','C'); base={'A':TARGET_A,'B':TARGET_B}.get(b['abc'],TARGET_C)
        b['td']=int(base+(EXTRA_Z_DAYS if b['xyz']=='Z' else 0))
    return rows

def make_segment_xlsx(DS, path, H=90):
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    wb=Workbook(); ws=wb.active; ws.title='Segmentation'
    ws['A1']='Segmentation - ABC by value, XYZ by demand variability, and the target days each class gets'; ws['A1'].font=Font(name=AR,size=13,bold=True,color=NAVY)
    ws['A2']=(f'ABC: value = next {H} days of requirement x unit price (MIS In Transit Unit Rate); A = top 80% of value, B = next 15%, C = rest or no price. '
              'XYZ: weekly requirement variability over 13 weeks (X stable < 0.25, Y 0.25-0.5, Z erratic > 0.5). PM = packaging / consumables. Target days are set on the Settings sheet.')
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:J2'); ws['A2'].alignment=Alignment(wrap_text=True); ws.row_dimensions[2].height=30
    H_=['Material','Code','UOM','Class (ABC)','Variability (XYZ)','CV weekly','Unit price','Value next %d d'%H,'Share of value','Target days']
    for j,h in enumerate(H_,1):
        c=ws.cell(4,j,h); c.font=Font(name=AR,size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=NAVY); c.border=bd; c.alignment=Alignment(horizontal='center',wrap_text=True)
    tot=sum(b['value'] for b in DS['blocks'] if b.get('value')) or 1
    rows=sorted(DS['blocks'],key=lambda b:rrank(b['mat']))          # highest shortage risk first
    for i,b in enumerate(rows):
        vals=[b['desc'] if b.get('count_stock') else b['common'],b['mat'],b.get('uom'),b['abc'],b['xyz'],round(b['cv'],2) if b.get('cv') is not None else None,
              round(b['price'],4) if b.get('price') else None,round(b['value']) if b.get('value') else None,(b['value']/tot) if b.get('value') else None,b['td']]
        for j,v in enumerate(vals,1):
            c=ws.cell(5+i,j,v); c.font=Font(name=AR,size=9,bold=j in (4,10)); c.border=bd
            if j==8: c.number_format='#,##0'
            if j==9: c.number_format='0.0%'
        ws.cell(5+i,4).fill=PatternFill('solid',fgColor={'A':'FFC7CE','B':'FFEB9C','C':'C6EFCE','PM':'DDEBF7'}.get(b['abc'],'FFFFFF'))
    ws.add_table(Table(displayName='Segmentation',ref=f'A4:J{4+len(rows)}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=True)))
    for j,w in enumerate([32,18,6,9,9,9,10,14,9,9],1): ws.column_dimensions[L(j)].width=w
    wb.save(path)


LINE_MODS={}       # production line key -> modules per day (your Production Plan) - for 'modules/day affected' in the emails
def material_report(A, d0, fname=''):
    """ONE short daily report, easy to read as a text message: materials whose stock + in transit lasts fewer than
       REPORT_DAYS days at the module plan, shortest first."""
    tg=([(c_,v_) for c_,l_,v_,o_ in CUST_TARGET] if CUSTOMER_TARGETS else [(lab.split(' Production')[0].replace('USA ','').replace(' (plan peak)','*'),v) for _,(lab,v) in sorted(EFF_TARGET.items()) if v])
    low=sorted([x for x in A if x.get('pl_days') is not None and x['pl_days']<REPORT_DAYS],key=lambda x:((x.get('tr_tot') or 0)>0,x['pl_days']))
    n=len(low)
    subj=f"Material report {d0:%d %b}: {n} under {REPORT_DAYS} days" if n else f"Material report {d0:%d %b}: all OK ({REPORT_DAYS}+ days)"
    body=[]
    if not tg: body.append('No line daily targets on Settings - no report.')
    else:
        body.append(f"Under {REPORT_DAYS} days at the module plan - nothing in transit first")
        body.append('('+' / '.join(f'{a} {v:,.0f}' for a,v in tg)+')')
        if any(a.endswith('*') for a,_ in tg): body.append('* no target on Settings: highest planned day')
        body.append('')
        for i,x in enumerate(low,1):
            b=x['b']; u=b.get('uom') or ''
            body.append(f"{i}) {b['common']}")
            body.append(f"   {x['pl_days']:.1f} days | stock {b['stock']:,.0f} | in transit {x.get('tr_tot') or 0:,.0f} {u}".rstrip())
            ln_=[EFF_TARGET[k_][0].split(' Production')[0].replace(' (plan peak)','').replace('USA ','') for k_ in (x.get('tgt_lines') or []) if k_ in EFF_TARGET]
            if ln_: body.append(f"   Lines: {', '.join(ln_)}")
        if not low: body.append(f"All materials {REPORT_DAYS}+ days.")
    return subj,body

def unfreeze_all(path, keep=()):
    """Remove every frozen pane from the output workbook (edits the sheet XML directly, so formulas, cached values and formats stay),
       except on the sheets named in keep."""
    import zipfile, shutil
    tmp_=path+'.unfreeze'
    with zipfile.ZipFile(path) as z0:
        wbx=z0.read('xl/workbook.xml').decode('utf-8'); rels=z0.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    keep_parts=set()
    for n_ in keep:
        m_=re.search(r'<sheet\b[^>]*name="'+re.escape(n_)+r'"[^>]*r:id="([^"]+)"',wbx)
        if m_:
            t_=re.search(r'<Relationship\b[^>]*Id="'+m_.group(1)+r'"[^>]*Target="([^"]+)"',rels) or re.search(r'<Relationship\b[^>]*Target="([^"]+)"[^>]*Id="'+m_.group(1)+'"',rels)
            if t_: keep_parts.add('xl/'+t_.group(1).lstrip('/').replace('xl/','',1))
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp_,'w',zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data=zin.read(it.filename)
            if it.filename.startswith('xl/worksheets/sheet') and it.filename.endswith('.xml') and it.filename not in keep_parts:
                t=data.decode('utf-8')
                t=re.sub(r'<pane\b[^>]*/>','',t)                                   # frozen / split pane
                t=re.sub(r'(<selection\b[^>]*?)\s+pane="[^"]*"',r'\1',t)            # selections that pointed at a pane
                data=t.encode('utf-8')
            zout.writestr(it,data)
    shutil.move(tmp_,path)

def reorder_sheets(pkg_path, first):
    """Puts the sheets named in 'first' at the front in that order (others keep their order after them) by editing the
       workbook XML only - charts and saved formula results stay. Sheet-scoped names (print titles) follow their sheet."""
    import zipfile, shutil
    from lxml import etree
    M='http://schemas.openxmlformats.org/spreadsheetml/2006/main'; q=lambda t:'{%s}%s'%(M,t)
    z=zipfile.ZipFile(pkg_path); parts={n:z.read(n) for n in z.namelist()}; z.close()
    wb=etree.fromstring(parts['xl/workbook.xml']); sh=wb.find(q('sheets')); old=list(sh)
    new=[e for n in first for e in old if e.get('name')==n]+[e for e in old if e.get('name') not in first]
    pos={id(e):i for i,e in enumerate(old)}; newidx={pos[id(e)]:i for i,e in enumerate(new)}
    for e in old: sh.remove(e)
    for e in new: sh.append(e)
    dn=wb.find(q('definedNames'))
    if dn is not None:
        for d in dn:
            if d.get('localSheetId') is not None: d.set('localSheetId',str(newidx.get(int(d.get('localSheetId')),int(d.get('localSheetId')))))
    bv=wb.find(q('bookViews'))
    if bv is not None:
        for v in bv: v.set('activeTab','0'); v.set('firstSheet','0')
    parts['xl/workbook.xml']=etree.tostring(wb,xml_declaration=True,encoding='UTF-8',standalone=True)
    # only the first sheet is 'selected'
    rels=etree.fromstring(parts['xl/_rels/workbook.xml.rels']); tg={r.get('Id'):r.get('Target') for r in rels}
    RID='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    for i,e in enumerate(new):
        pth='xl/'+tg[e.get(RID)].lstrip('/').replace('xl/','')
        x=etree.fromstring(parts[pth])
        for sv in x.iter(q('sheetView')):
            if i==0: sv.set('tabSelected','1')
            elif 'tabSelected' in sv.attrib: del sv.attrib['tabSelected']
        parts[pth]=etree.tostring(x,xml_declaration=True,encoding='UTF-8',standalone=True)
    tmpz=pkg_path+'.reorder'
    with zipfile.ZipFile(tmpz,'w',zipfile.ZIP_DEFLATED) as zo:
        for n,b in parts.items(): zo.writestr(n,b)
    shutil.move(tmpz,pkg_path)


ALLOC_STOP_={}
def add_merges_keep_values(path, ranges):
    """Merge cell ranges in sheet 1 of path but KEEP the value in every cell of the range (openpyxl would clear them):
       then a filter that hides the first row of a merged block still shows the number in the rows left visible."""
    import zipfile, shutil
    if not ranges: return
    z=zipfile.ZipFile(path); tmp_=path+'.mg'
    with zipfile.ZipFile(tmp_,'w',zipfile.ZIP_DEFLATED) as o:
        for it in z.infolist():
            d=z.read(it.filename)
            if it.filename=='xl/worksheets/sheet1.xml':
                t=d.decode('utf-8'); mc=''.join(f'<mergeCell ref="{r_}"/>' for r_ in ranges)
                if '<mergeCells' in t: t=re.sub(r'(<mergeCells[^>]*>)',lambda m_: m_.group(1)+mc,t,count=1); t=re.sub(r'<mergeCells count="(\d+)"',lambda m_: f'<mergeCells count="{int(m_.group(1))+len(ranges)}"',t,count=1)
                else:
                    blk=f'<mergeCells count="{len(ranges)}">{mc}</mergeCells>'
                    pos=[t.find(k_) for k_ in ('<phoneticPr','<conditionalFormatting','<dataValidations','<hyperlinks','<printOptions','<pageMargins','<pageSetup','<headerFooter','<drawing','<tableParts') if t.find(k_)>0]
                    p_=min(pos) if pos else t.rfind('</worksheet>'); t=t[:p_]+blk+t[p_:]
                d=t.encode('utf-8')
            o.writestr(it,d)
    z.close(); shutil.move(tmp_,path)



def module_short_names(mis_path, bns):
    """Short module name from the MIS BOM name: drop '500 TC G2G_' and 'USA', add the cell size (G12R / M10R) when the name
       does not say it (read from the BOM's glass / frame), e.g. '500 TC G2G_G12R_3.2' -> 'G12R 3.2', '500 TC G2G_EPEP_USA' -> 'M10R EPEP'."""
    out={}; size={}; hs={}
    try:
        cn={}
        for r in wb_ro(mis_path)['Common Name'].iter_rows(min_row=2,values_only=True):
            if r and r[0]: cn[str(r[0]).strip()]=str(r[3] or '')
        for r in wb_ro(mis_path)['Standard BOM'].iter_rows(min_row=2,values_only=True):
            if r and len(r)>4 and r[2] and r[3]:
                t=(cn.get(str(r[3]).strip(),'')+' '+str(r[4] or '')).upper()
                mh=re.search(r'(\d\d)\s*HS',t)
                if mh and str(r[8] or '').lower().startswith('alum'): hs.setdefault(str(r[2]).strip(),mh.group(1)+'HS')
                for z in ('G12R','M10R','G12','M10'):
                    if z in t: size.setdefault(str(r[2]).strip(),{}); size[str(r[2]).strip()][z]=size[str(r[2]).strip()].get(z,0)+1; break
    except Exception as _e_: _swallow(6251,_e_); pass
    for bn in bns:
        if not bn: continue
        t=re.sub(r'^\d+\s*(TC|MP)\s*G2G_?','',str(bn).strip())
        parts=[p_ for p_ in re.split(r'[_\s]+',t) if p_ and p_.upper()!='USA']
        nm=' '.join(parts) or str(bn)
        if not re.search(r'G12|M10',nm,re.I) and size.get(str(bn).strip()):
            nm=f"{max(size[str(bn).strip()],key=size[str(bn).strip()].get)} {nm}"
        if re.fullmatch(r'G12R?',nm) and hs.get(str(bn).strip()): nm=f"{nm} {hs[str(bn).strip()]}"     # plain G12R: add the frame height
        out[bn]=nm
    return out

def plan_windows(arr, d0, gap=3):
    """Runs of days with demand (gaps of up to 'gap' days without production = same run): [(start_k, end_k)]."""
    out=[]; st=None; last=None
    for k,v in enumerate(arr):
        if v>0:
            if st is None: st=k
            elif k-last>gap+1: out.append((st,last)); st=k
            last=k
    if st is not None: out.append((st,last))
    return out

def future_plan_text(b, lk, d0, nd, k_today):
    """How this line uses the material over the plan, in one short line:
       'Same as today' = used from the line's first production day to the end of the plan by one customer;
       otherwise when it is used (starts later / stops / only a few days) | the customers in date order."""
    per={}
    for (cu,ln),arr in (b.get('cust_line') or {}).items():
        if not ln or _line_key(ln)!=lk or not any(arr): continue
        us=UCMAP.get((_norm(cu),lk)) or {}
        nm=max(us,key=us.get) if us else cu
        a_=per.setdefault(nm,[0.0]*nd)
        for k in range(min(nd,len(arr))): a_[k]+=arr[k]
    if not per: return None,False
    tot=[sum(a_[k] for a_ in per.values()) for k in range(nd)]
    k0=k_today or 0; dd=lambda k: f"{(d0+pd.Timedelta(days=k)):%d-%b}"
    w=plan_windows(tot,d0)
    whole=len(w)==1 and w[0][0]<=k0+3 and w[0][1]>=nd-4
    seq=[]                                            # customer runs in date order (gaps without production ignored)
    for k in range(nd):
        if tot[k]<=0: continue
        nm=max(per,key=lambda n_: per[n_][k])
        if not seq or seq[-1][0]!=nm: seq.append((nm,k))
    if whole and len(seq)==1: return 'Same as today',False
    if whole: use='Used all plan'
    else:
        txt=[]
        for a,z in w: txt.append(f"{dd(a)} → "+("end of plan" if z>=nd-4 else dd(z)))
        use=('Starts later: ' if w[0][0]>k0+3 else 'Used ')+'; '.join(txt)+('' if w[-1][1]>=nd-4 else ' (then not used)')
    cust=' → '.join(f"{nm} ({dd(k)})" for nm,k in seq)
    return f"{use} | {cust}",True

def line_plan_need(b, lk, nd):
    """Module plan need per day of one material on one production line: average per production day (> 0), next 9 weeks."""
    arr=[0.0]*min(nd,avg_prod_days())
    for (cu_,ln_),a_ in (b.get('cust_line') or {}).items():
        if ln_ and _line_key(ln_)==lk:
            for k in range(len(arr)): arr[k]+=a_[k] if k<len(a_) else 0.0
    nz=[v for v in arr if v>0]
    return round(sum(nz)/len(nz),1) if nz else None
def make_line_view_xlsx(A, LR, d0, nd, path, TD=None, AL_COLS=None):
    """MRP by Line: one row per material and production line, sorted like the Action List (material, then line).
       Line columns: how long the line can run, this line's need at its daily target and at its daily plan (average of the next
       21 days with production), when this line is short (allocated) and where it is supplied from.
       Then the Action List columns Stock ... Daily plan with OEM for the material (formulas = the Action List row),
       merged over the material's lines but with the value kept in every cell, so filtering a line still shows them."""
    from openpyxl.styles import Border, Side
    AR='Calibri'; NAVY='1F3864'; TEAL='1F6E6E'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    wb=Workbook(); ws=wb.active; ws.title='MRP by Line'; ws.sheet_view.showGridLines=False
    ws['A1']='MRP by Line - each material per production line, with the module-plan numbers of the material'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=('Filter column A (line) and / or column D (component). Teal columns = this line only. The other columns are the material totals (module plan: need per day, '
              'days of cover, below 21 days from, days late - the same numbers as the ALERT); they are merged over the material\'s lines, and each cell still holds the number, so a filter keeps it visible. '
              '')
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:N2'); ws['A2'].alignment=Alignment(wrap_text=True,vertical='top'); ws.row_dimensions[2].height=40
    LCOL=['Production line','Line can run (days)','Line stops on','Component','Component description','UOM','First needed (this line)',f'Needed now (next {int(LOOKAHEAD_DAYS)} days)',
          'Plan on this line (who / when) - differs from today?',
          'This line: daily plan need','This line: short from (allocated)','This line: supplied from (in order)']
    AL_COLS=[(h_,c_) for h_,c_ in (AL_COLS or []) if 'oem' not in h_.lower() and 'daily target' not in h_.lower()]   # targets / OEM: Action List only
    H=LCOL+[('In transit usable by this line (MIS, allowed suppliers)' if h_=='In transit today (total qty, MIS)' else h_) for h_,_ in AL_COLS]
    for j_,h in enumerate(H,1):
        c=ws.cell(4,j_,h); c.border=bd; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        if j_<=len(LCOL): col_,txt_=(TEAL if j_>=7 else NAVY),'FFFFFF'
        else: col_=al_hdr_color(h); txt_='1F1F1F' if al_hdr_light(h) else 'FFFFFF'
        c.fill=PatternFill('solid',fgColor=col_); c.font=Font(name=AR,size=9,bold=True,color=txt_)
    ws.row_dimensions[4].height=48
    info={r_['key']:r_ for r_ in (LR or [])}; lorder={r_['key']:i_ for i_,r_ in enumerate(LR or [])}
    order=[x for x in (AL_ORDER_ or A) if not x.get('sup_row')] or A          # same order as the Action List: highest shortage risk first
    r=5; merges=[]
    for x in order:
        b=x['b']; lines={}
        for (cu,ln),arr in (b.get('cust_line') or {}).items():
            if not ln or not any(arr): continue
            a_=lines.setdefault(_line_key(ln),[0.0]*nd)
            for k in range(min(nd,len(arr))): a_[k]+=arr[k]
        if not lines: continue
        r0=r
        for lk in sorted(lines,key=lambda k_:(lorder.get(k_,99),k_)):
            a_=lines[lk]; L_=info.get(lk); kt,limt,_=L_['t'] if L_ else (nd,[],[])
            f_=next((k for k in range(nd) if a_[k]),None)
            fut_txt,fut_diff=future_plan_text(b,lk,d0,nd,LINE_TODAY.get(lk))
            vals=[(L_['name'] if L_ else lk),(kt if kt<nd else f'{nd}+'),((d0+pd.Timedelta(days=kt)).to_pydatetime() if kt<nd else None),b['mat'],
                  (b['desc'] if b.get('count_stock') else b['common']),b.get('uom'),((d0+pd.Timedelta(days=f_)).to_pydatetime() if f_ is not None else None),
                  (('Yes' if (f_ is not None and f_<int(LOOKAHEAD_DAYS)) else 'No') if VALUES_ONLY else f'=IF(AND(ISNUMBER(G{r}),G{r}<DATE({d0.year},{d0.month},{d0.day})+{TGT_CELL.get("_look") or int(LOOKAHEAD_DAYS)}),"Yes","No")'),
                  fut_txt,
                  line_plan_need(b,lk,nd),                                    # this line's module plan x usage, average per production day (9 weeks)
                  ((d0+pd.Timedelta(days=ALLOC_STOP_[(b['mat'],lk)])).to_pydatetime() if ALLOC_STOP_.get((b['mat'],lk),nd)<nd else
                   ('No demand' if x.get('no_demand') else 'covered here – short on other lines' if (x.get('wins') or []) else 'covered')),
                  (x.get('used') or {}).get(lk) or '']
            if VALUES_ONLY:
                vals+=[(round(ALLOC_IT[(b['mat'],lk)]) if (h_=='In transit today (total qty, MIS)' and (b['mat'],lk) in ALLOC_IT) else AL_VAL.get((b['mat'],h_))) for h_,_ in AL_COLS]
            else: vals+=[f"=INDEX('Action List'!${cl}$5:${cl}$1000,MATCH($D{r},'Action List'!$A$5:$A$1000,0))" for _,cl in AL_COLS]
            for j_,v in enumerate(vals,1):
                c=ws.cell(r,j_,v); c.font=Font(name=AR,size=9); c.border=bd
                c.alignment=Alignment(vertical=('top' if (j_<=len(LCOL) or 'Arrival' in H[j_-1]) else 'center'),wrap_text=j_ in (5,9,12) or (j_>len(LCOL) and 'Arrival' in H[j_-1]))
                if j_ in (3,7): c.number_format='dd-mmm-yy'
                if j_==10: c.number_format='#,##0.0'
                if j_==8: c.alignment=Alignment(horizontal='center',vertical='top')
                if j_==11 and isinstance(v,dt.datetime): c.number_format='dd-mmm'; c.font=Font(name='Calibri',size=10,bold=True,color='9C0006')
                if j_>len(LCOL):
                    hh=H[j_-1]; c.number_format=('0.0' if hh.startswith('Inv. Days') else ('#,##0.0' if 'Daily' in hh else '#,##0;-#,##0;"-"'))
                    if hh.startswith('Inv. Days') and isinstance(v,(int,float)) and v<REPORT_DAYS:     # under 21 days: red highlight
                        c.fill=PatternFill('solid',fgColor='FFC7CE'); c.font=Font(name='Calibri',size=10,bold=True,color='9C0006')
            if L_ and kt<nd: ws.cell(r,2).font=Font(name=AR,size=9,bold=True,color='9C0006')
            if fut_diff: ws.cell(r,9).fill=PatternFill('solid',fgColor='FFEB9C')      # the plan on this line is not the same as today
            r+=1
        n_=str(x.get('ata_txt') or '').count('\n')+1; need_h=12.5*n_; cnt_=r-r0      # room for the arrival lines
        if cnt_ and need_h>15*cnt_:
            for rr_ in range(r0,r): ws.row_dimensions[rr_].height=need_h/cnt_
        if r-1>r0:
            for j_ in range(len(LCOL)+1,len(H)+1):
                if H[j_-1]!='In transit usable by this line (MIS, allowed suppliers)': merges.append(f'{L(j_)}{r0}:{L(j_)}{r-1}')
            for j_ in (4,5,6): merges.append(f'{L(j_)}{r0}:{L(j_)}{r-1}')     # same component on every line row
        for j_ in range(1,len(H)+1): ws.cell(r0,j_).border=Border(left=thin,right=thin,bottom=thin,top=Side(style='medium',color=NAVY))
    if r>5: ws.auto_filter.ref=f'A4:{L(len(H))}{r-1}'
    widths=[16,9,10,16,30,6,10,9,40,11,10,30]+[(42 if 'Arrival' in h_ else 11) for h_,_ in AL_COLS]
    for j_,w in enumerate(widths,1): ws.column_dimensions[L(j_)].width=w
    set_print(ws,'MRP by Line'); wb.save(path)
    add_merges_keep_values(path,merges)

def cspm_from_bom_pp(bom, count, pp_rows, dates, sr0, count_hdr=None):
    """CS / PM plan straight from YOUR BOM and YOUR Production Plan, row by row:
       the Production Plan customer is found in the BOM by customer name; if not, by its Module Model (the BOM of that model).
       Returns (blocks, Check rows for Production Plan customers that have no BOM)."""
    nd=len(dates); iss=[]
    bycm={}                                            # (customer, model) -> {component: qty per module} (rows of one model add up)
    meta={}
    for d in bom:
        c=str(d.get('Component') or '').strip()
        if c[:2].upper() not in ('CS','PM') or not isinstance(d.get('Quantity'),(int,float)): continue
        cu=_norm(d.get('Customer')); mdl=str(d.get('Module Model') or '').strip()
        bycm.setdefault((cu,mdl),{}); bycm[(cu,mdl)][c]=bycm[(cu,mdl)].get(c,0.0)+float(d['Quantity'])
        meta.setdefault(c,(d.get('Component desciption'),d.get('UoM')))
    bycust={}; bymodel={}
    for (cu,mdl),comp in bycm.items():
        bycust.setdefault(cu,[]).append(comp); bymodel.setdefault(mdl,[]).append(comp)
    def avg(lst):
        out={}
        for comp in lst:
            for c,q in comp.items(): out[c]=out.get(c,0.0)+q/len(lst)
        return out
    plan={}; cust={}; cl={}; how={}; pls={}
    for x in pp_rows or []:
        key=_norm(x['cust']); mdl=str(x.get('model') or '').strip(); cu=(LINKS.get(x['cust']) or {}).get('mis') or x['cust']   # MIS name shown
        if key in bycust: comp=avg(bycust[key]); how[cu]='customer'
        elif mdl and mdl in bymodel: comp=avg(bymodel[mdl]); how[cu]=f'model {mdl[:14]}'
        else:
            if sum(x['v']) and cu not in how:
                how[cu]=None
                iss.append(dict(area='Plan',check='Production Plan customer has no packaging BOM (CS / PM not planned)',sev='Medium',material=x['cust'],code=x['line'],
                                a_lab='',a=None,b_lab='Modules in horizon',b=round(sum(x['v'])),diff=None,
                                detail=f"{x['cust']} ({mdl or 'no model'}): not in your BOM by customer or by model",action='Add its BOM rows, or the Module Model on your Production Plan / Index'))
            continue
        for c,q in comp.items():
            p_=plan.setdefault(c,[0.0]*nd); a_=cust.setdefault(c,{}).setdefault(cu,[0.0]*nd); l_=cl.setdefault(c,{}).setdefault((cu,x['line']),[0.0]*nd)
            for k in range(nd):
                v=q*x['v'][k]; p_[k]+=v; a_[k]+=v; l_[k]+=v
            pls.setdefault(c,[]).append(dict(cu=cu,ln=x['line'],bn=None,rate=q,mods=list(x['v']),src=f"your BOM (packaging, by {how.get(cu) or 'customer'})"))
    qi=count_qty_idx(count_hdr)
    stock={str(r[0]).strip():(float(r[qi]) if len(r)>qi and isinstance(r[qi],(int,float)) else 0.0) for r in count if r and r[0]}
    hl=[str(h).strip().lower() if h is not None else '' for h in (count_hdr or [])]
    ix=lambda n: hl.index(n) if n in hl else None
    i_q,i_u,i_f=ix('today stock qty'),ix('stock uom'),ix('quantity')
    raw={}
    for r in count:
        if not r or not r[0] or i_q is None: continue
        qv=r[i_q] if len(r)>i_q else None; u=r[i_u] if (i_u is not None and len(r)>i_u) else ''; fv=r[i_f] if (i_f is not None and len(r)>i_f) else None
        if isinstance(qv,(int,float)): raw[str(r[0]).strip()]=f"{qv:,.0f} {u}"+(f" x {fv:g}" if isinstance(fv,(int,float)) and fv not in (0,1) else '')
    out=[]; sr=sr0
    for c in sorted(set(meta)|set(plan)):
        sr+=1; desc=' '.join(str((meta.get(c) or ('',''))[0] or '').split()); P=plan.get(c,[0.0]*nd)
        cp=dict(sorted({k:v for k,v in (cust.get(c) or {}).items() if any(v)}.items(),key=lambda kv:-sum(kv[1])))
        out.append(dict(key=f'{sr}CS/PM',sr=sr,short='CS - Consumable' if c.upper().startswith('CS') else 'PM - Packaging',common=desc,mat=c,desc=desc,mtype='CS/PM',
                        stock=stock.get(c,0.0),uom=(meta.get(c) or ('',''))[1] or '',plan=P,cust_plan=cp,
                        cust_line=dict(sorted({k:v for k,v in (cl.get(c) or {}).items() if any(v)}.items(),key=lambda kv:(-sum(cp.get(kv[0][0],[0])),kv[0][0],str(kv[0][1])))),
                        transit=[0.0]*nd,adhoc=[0.0]*nd,openpo=[0.0]*nd,count_stock=True,raw=raw.get(c,'not on Stock Count'),plan_lines=pls.get(c,[]),
                        st_main=stock.get(c,0.0),st_wh=0.0,st_jw=0.0,plan_all={pd.Timestamp(dates[k]).normalize():P[k] for k in range(nd) if P[k]}))
    return out,iss

def make_bom_sheet(bom, mis_names, path):
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.table import Table, TableStyleInfo
    wb=Workbook(); ws=wb.active; ws.title='BOM'
    src_=os.path.basename(PPC_SRC[0]) if PPC_SRC[0] and 'BOM' in _sheetnames(PPC_SRC[0]) else BOM_SRC_NAME[0]
    ws['A1']=(f'Copy of the BOM in {src_} - edit it THERE. MIS BOM Name (link to the MIS plan) comes from the module plan (customers using each Module Model).'
              if BOM_SRC_NAME and BOM_SRC_NAME[0] and not BOM_SRC_NAME[0].startswith('MRP_Daily') else
              'BOM - your input. Edit, add or delete rows here; the next run uses this sheet. MIS BOM Name links each model to the MIS production plan.')
    ws['A1'].font=Font(name='Arial',size=10,bold=True,color='C00000')
    for j,h in enumerate(BOM_COLS,1):
        c=ws.cell(3,j,h); c.font=Font(name='Arial',size=10,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='1F3864')
    for i,d in enumerate(bom):
        for j,h in enumerate(BOM_COLS,1): ws.cell(4+i,j,d[h])
    last=3+max(1,len(bom))
    ws.add_table(Table(displayName='BOMInput',ref=f'A3:{L(len(BOM_COLS))}{last}',tableStyleInfo=TableStyleInfo(name='TableStyleLight9',showRowStripes=True)))
    if mis_names:
        lst=wb.create_sheet('lists'); [lst.cell(i+1,1,n) for i,n in enumerate(mis_names)]
        ws.cell(2,14,'MIS BOM Names: '+' | '.join(mis_names)).font=Font(name='Arial',size=9,italic=True)
        dv=DataValidation(type='list',formula1='"'+','.join(mis_names)[:250]+'"',allow_blank=True); ws.add_data_validation(dv); dv.add(f'K4:K{last+200}')
        del wb['lists']
    for j,w in enumerate([24,42,7,5,22,40,10,7,11,11,26],1): ws.column_dimensions[L(j)].width=w
    ws.freeze_panes='A4'; wb.save(path)

def make_count_sheet(hdr, rows, path, bom=None, cache=None):
    """Stock Count (your input). With the columns Today Stock Qty / Stock UOM / BOM UOM / Quantity / Today Qty_BOM, BOM UOM and
       Today Qty_BOM stay formulas (Today Qty_BOM = Today Stock Qty x Quantity), and the MRP uses Today Qty_BOM."""
    wb=Workbook(); ws=wb.active; ws.title='Stock Count'
    src_=(os.path.basename(PPC_SRC[0])+' (sheet Count)') if PPC_SRC[0] else BOM_SRC_NAME[0]
    ws['A1']=(f'Copy of the stock count in {src_} - update it THERE (Today Stock Qty x Quantity = Today Qty_BOM, used by the MRP).'
              if BOM_SRC_NAME and BOM_SRC_NAME[0] and not BOM_SRC_NAME[0].startswith('MRP_Daily') else
              'Stock Count - your input (CS / PM physical count). Enter Today Stock Qty in the stock unit and Quantity = BOM units per stock unit '
              '(e.g. 21.8 KG per roll); the MRP uses Today Qty_BOM.')
    ws['A1'].font=Font(name='Arial',size=10,bold=True,color='C00000')
    hdr=list(hdr or ['Material','Material Description','Today Qty'])
    hl=[str(h).strip().lower() if h is not None else '' for h in hdr]
    ix=lambda n: hl.index(n) if n in hl else None
    iq,ibu,ifac,iout=ix('today stock qty'),ix('bom uom'),ix('quantity'),ix('today qty_bom')
    bomu={str(d['Component']).strip():d['UoM'] for d in (bom or []) if d.get('Component') and d.get('UoM')}
    for j,h in enumerate(hdr,1):
        c=ws.cell(3,j,h); c.font=Font(name='Arial',size=10,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='1F3864')
        if isinstance(h,(dt.datetime,dt.date)): c.number_format='dd-mmm'
    for i,r in enumerate(rows):
        rr=4+i
        for j,v in enumerate(r,1):
            if (j-1)==ibu:          # BOM UOM looked up from this workbook's BOM sheet (Component = E, UoM = H)
                ws.cell(rr,j,f'=IFERROR(INDEX(BOM!$H:$H,MATCH($A{rr},BOM!$E:$E,0)),"")')
                if cache is not None: cache[(rr,j)]=str(bomu.get(str(r[0]).strip(),'') or '')
            elif (j-1)==iout and iq is not None and ifac is not None:
                ws.cell(rr,j,f'={L(iq+1)}{rr}*IF({L(ifac+1)}{rr}="",1,{L(ifac+1)}{rr})')
                qv=r[iq] if isinstance(r[iq],(int,float)) else 0; fv=r[ifac] if isinstance(r[ifac],(int,float)) else 1
                if cache is not None: cache[(rr,j)]=float(qv)*float(fv)
                ws.cell(rr,j).font=Font(bold=True)
            else: ws.cell(rr,j,None if isinstance(v,str) and v.strip() in ('#N/A','#REF!','#VALUE!','#DIV/0!','#NAME?') else v)
    for j,w in enumerate([18,42,14,10,10,10,15],1): ws.column_dimensions[L(j)].width=w
    if iout is not None:
        for rr in range(4,4+len(rows)): ws.cell(rr,iout+1).fill=PatternFill('solid',fgColor='E2EFDA')
    ws.freeze_panes='C4'; wb.save(path)

def cspm_checks(bom, count, modplan, mis_names, today, H=30):
    I=[]
    def add(area,check,sev,material='',code='',a_lab='',a=None,b_lab='',b=None,detail='',action=''):
        I.append(dict(area=area,check=check,sev=sev,material=material,code=code,a_lab=a_lab,a=a,b_lab=b_lab,b=b,
                      diff=(b-a) if isinstance(a,(int,float)) and isinstance(b,(int,float)) else None,detail=detail,action=action))
    used={str(d['MIS BOM Name'] or '').strip() for d in bom}
    for bn,cu in modplan.items():
        q=sum(sum(v[:H]) for v in cu.values())
        if q and bn not in used:
            add('Plan','MIS BOM Name has no BOM rows','High',bn,'','',None,f'Modules next {H} days',q,
                'Customers: '+', '.join(sorted(cu))+' - no CS / PM plan for these modules','Add BOM rows (or set MIS BOM Name on existing rows) in the BOM sheet')
    for d in bom:
        bn=str(d['MIS BOM Name'] or '').strip()
        if not bn: add('Master data','BOM row without MIS BOM Name','Medium',str(d['Module Model'] or ''),str(d['Component'] or ''),'',None,'',None,'This row is not used in the plan','Fill MIS BOM Name in the BOM sheet')
        elif mis_names and bn not in mis_names: add('Master data','BOM row with unknown MIS BOM Name','Medium',bn,str(d['Component'] or ''),'',None,'',None,'Name not found in MIS Total Plan Req','Pick a name from the list')
    pm_bom={str(d['Component']).strip() for d in bom if str(d['Component'] or '')[:2].upper() in ('CS','PM')}
    cnt={str(r[0]).strip() for r in count if r and r[0]}
    for c in sorted(pm_bom-cnt): add('Stock','CS / PM component not on Stock Count','Medium','',c,'',None,'',None,'Stock is taken as 0','Add the material to the Stock Count sheet')
    return I

# ===================== Customer -> supplier map (your input) =====================
MAP_COLS=['Customer','Material','Code','Supplier','Share %','Note','Suggested supplier']
def _norm(x): return ' '.join(str(x or '').split()).lower()
_LO_SAID=set()
def link_override(cu, quiet=False):
    """Settings MIS LINK override for a plan customer: exact name first, else the same name spelled slightly differently
       (e.g. 'Clenera Crimson' in the PPC plan vs 'Cleanera Crimson' on Settings, similarity >= 0.9). Prints a note once."""
    n=_norm(cu); ov=LINK_OVERRIDE.get(n)
    if ov or not n or not LINK_OVERRIDE: return ov
    import difflib
    k_=max(LINK_OVERRIDE,key=lambda k2:difflib.SequenceMatcher(None,n,k2).ratio())
    if difflib.SequenceMatcher(None,n,k_).ratio()<0.9: return None
    if not quiet and n not in _LO_SAID:
        _LO_SAID.add(n); print(f"Settings MIS LINK '{k_}' used for plan customer '{cu}' (spelling differs - make them the same to remove this note)",flush=True)
    return LINK_OVERRIDE[k_]

def read_supplier_map(prev_path):
    rows=[]
    if not prev_path or not os.path.exists(prev_path): return rows
    try: wb=wb_ro(prev_path)
    except Exception as _e_: _swallow(6548,_e_); return rows
    if 'Customer Supplier' not in wb.sheetnames: return rows
    data=[list(r) for r in wb['Customer Supplier'].iter_rows(values_only=True)]
    hr=next((i for i,r in enumerate(data) if 'Customer' in r and 'Supplier' in r),None)
    if hr is None: return rows
    h=[str(x).strip() if x is not None else '' for x in data[hr]]
    for r in data[hr+1:]:
        d={c:(r[h.index(c)] if c in h and h.index(c)<len(r) else None) for c in MAP_COLS}
        if d['Customer'] and (d['Code'] or d['Material']): rows.append(d)
    return rows

def apply_supplier_map(DS, maprows, add_new=False):
    """Puts each supplier's share of the plan on its supplier block. The sheet is yours: rows are only added when add_new
       (first creation); customers missing from it are reported on Check instead. Returns (rows, issues for Check)."""
    nd=len(DS['dates']); issues=[]
    def add(sev,check,material,code,detail,action):
        issues.append(dict(area='Plan',check=check,sev=sev,material=material,code=code,a_lab='',a=None,b_lab='',b=None,diff=None,detail=detail,action=action))
    rows=list(maprows); have={(_norm(r['Customer']),str(r['Code'] or '').strip()) for r in rows}
    for b in (DS['blocks'] if add_new else []):
        if not b.get('sup') or b.get('count_stock'): continue
        for cu,v in (b.get('cust_plan') or {}).items():
            if sum(v)>0 and (_norm(cu),b['mat']) not in have:
                rows.append({'Customer':cu,'Material':b['common'],'Code':b['mat'],'Supplier':None,'Share %':None})
                have.add((_norm(cu),b['mat']))
    bycode={}
    for r in rows: bycode.setdefault(str(r['Code'] or '').strip(),[]).append(r)
    for b in DS['blocks']:
        mr=bycode.get(b['mat'])
        if not mr or b.get('count_stock'): continue
        cp={_norm(k):(k,v) for k,v in (b.get('cust_plan') or {}).items()}
        sup=b.setdefault('sup',[]); byname={s_['name'].upper():s_ for s_ in sup}
        share={}; mapped=set()
        for r in mr:
            sname=str(r['Supplier'] or '').strip().upper(); cu=_norm(r['Customer'])
            if not sname: continue
            try: pct=float(r.get('Share %')) if r.get('Share %') not in (None,'') else 100.0
            except (TypeError,ValueError): pct=100.0
            share.setdefault(cu,{}); share[cu][sname]=share[cu].get(sname,0)+pct; mapped.add(cu)
        for cu,ss in share.items():
            if abs(sum(ss.values())-100)>0.5:
                add('Medium','Customer Supplier shares do not add up to 100%',b['common'],b['mat'],f"{cp.get(cu,(cu,))[0]}: "+', '.join(f'{k} {v:g}%' for k,v in ss.items()),'Correct Share % on the Customer Supplier sheet')
        for cu,(name,v) in cp.items():
            if sum(v[:30])>0 and cu not in mapped:
                add('Medium','Customer has plan but no supplier assigned',b['common'],b['mat'],f'{name}: {sum(v[:30]):,.0f} in the next 30 days not linked to a supplier',
                    'Add / fill this customer on the Customer Supplier sheet of your MRP workbook')
        for cu,ss in share.items():
            if cu not in cp: continue
            name,v=cp[cu]; tot=sum(ss.values()) or 100
            for sname,pct in ss.items():
                s_=byname.get(sname)
                if s_ is None:
                    s_=dict(name=sname,stock=0.0,transit=[0.0]*nd,openpo=[0.0]*nd); sup.append(s_); byname[sname]=s_
                    add('Low','Supplier on the map has no stock or shipment',b['common'],b['mat'],f'{sname} is assigned to {name} but has no stock, in-transit or open PO for this material','Check the supplier name, or plan a PO')
                s_.setdefault('plan',[0.0]*nd); s_.setdefault('custs',[])
                f=pct/tot
                for k in range(nd): s_['plan'][k]+=v[k]*f
                s_['custs'].append(f'{name}' if pct>=99.5 else f'{name} {pct:g}%')
        for s_ in sup:
            if 'plan' in s_: s_['cust_text']=' | '.join(s_['custs'])
    return rows,issues


def make_supplier_cover_sheet(DS, path, TD=None):
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    nd=len(DS['dates']); d0=pd.Timestamp(DS['dates'][0]).normalize(); AR='Calibri'; NAVY='1F3864'
    thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    wb=Workbook(); ws=wb.active; ws.title='Supplier Coverage'
    ws['A1']=f'Supplier coverage - each supplier\'s stock and shipments against the plan of the customers mapped to it'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=('From Settings ALLOWED VENDORS (which supplier each customer may use). Run-out date = first day that supplier\'s projected stock (stock + in transit) goes below zero. '
              'Status = CRITICAL / AT RISK / WATCH / OK against the material\'s own target days (Settings). Red run-out date = inside that target. Stock by supplier = the PPC Count by supplier when counted, else MB52 batches (Batch Master) / MIS RM Status by Make (JW at the OEM left out); "Unassigned" stock is not counted for any supplier.')
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:M2'); ws['A2'].alignment=Alignment(wrap_text=True); ws.row_dimensions[2].height=28
    H=['Material','Supplier','Customers mapped','Stock today','Avg plan / day (next target days)','Run-out date (with in transit)','In transit (total)','Open PO (total)','Next arrival','Status']
    for j,h in enumerate(H,1):
        c=ws.cell(4,j,h); c.font=Font(name=AR,size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=NAVY); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=bd
    ws.row_dimensions[4].height=36; r=5; n=0
    for b in sorted(DS['blocks'],key=lambda b_:rrank(b_['mat'])):
        for sp in b.get('sup') or []:
            if 'plan' not in sp: continue
            TDm=b.get('td') or TD or TARGET_DAYS
            cov=cover_days(sp['stock'],sp['plan'],sp['transit'],sp['openpo']); avg=sum(sp['plan'][:TDm])/TDm
            k=next((i for i in range(nd) if sp['transit'][i]>0),None)
            st=material_status(cov,TDm)
            vals=[b['common'],sp['name'],sp.get('cust_text'),round(sp['stock']),round(avg),
                  (d0+pd.Timedelta(days=cov[1])).to_pydatetime() if cov[1]<nd else None,round(sum(sp['transit'])),round(sum(sp['openpo'])),
                  (f"{(d0+pd.Timedelta(days=k)):%d-%b}: {sp['transit'][k]:,.0f}" if k is not None else None),st]
            for j,v in enumerate(vals,1):
                c=ws.cell(r,j,v); c.font=Font(name=AR,size=9); c.border=bd; c.alignment=Alignment(vertical='top',wrap_text=j==3)
                if j in (4,5,7,8): c.number_format='#,##0'
                if j==6: c.number_format='dd-mmm-yy'
            if cov[1]<TDm: ws.cell(r,6).fill=PatternFill('solid',fgColor='FFC7CE')
            ws.cell(r,10).fill=PatternFill('solid',fgColor=STATUS_FILL[st])
            r+=1; n+=1
    if n:
        ws.add_table(Table(displayName='SupplierCoverage',ref=f'A4:J{r-1}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=True)))
    else: ws.cell(5,1,'No supplier mapping yet - fill the Customer Supplier sheet').font=Font(name=AR,size=10,italic=True)
    for j,w in enumerate([22,14,44,12,12,13,12,12,18,11],1): ws.column_dimensions[L(j)].width=w
    ws.freeze_panes='C5'; wb.save(path); return n

# ===================== Module Plan from YOUR Production Plan (columns A-F) =====================
_STOP={'solar','farm','llc','inc','the','encap','ltd','co','project','g12r','m10r','energy'}
def _toks(s): return {t for t in re.split(r'[^a-z0-9]+',str(s or '').lower()) if t and t not in _STOP}
MODEL2BN={}
LINKS={}; LINK_OVERRIDE={}
def module_plan_from_pp(mrp_path, mis_path, plant, dates):
    """Modules per day from your Production Plan (A Date, B Production Line, C Wp, D Customer, E Qty, F MW).
       Each customer is linked to the MIS customer / MIS BOM Name it matches (same name, or shared words such as 'Lucky 7'),
       so the BOM sheet and the Customer Supplier sheet keep working. Returns (modplan, sheet rows, issues) or None."""
    try: wb=wb_ro(mrp_path)
    except Exception as _e_: _swallow(6657,_e_); return None
    if 'Production Plan' not in wb.sheetnames: return None
    idx={pd.Timestamp(d).normalize():k for k,d in enumerate(dates)}; nd=len(dates)
    # MIS customer -> BOM Name (most modules)
    mw=wb_ro(mis_path); cb={}; mis_day={}; seen_=set()
    for r in mw['Total Plan Req'].iter_rows(min_row=2,values_only=True):
        if len(r)>21 and str(r[21] or '').strip().upper()==plant.upper() and r[2] and r[4] and isinstance(r[3],(int,float)):
            cu=' '.join(str(r[2]).split()); cb.setdefault(cu,{}); cb[cu][str(r[4]).strip()]=cb[cu].get(str(r[4]).strip(),0)+r[3]
            if isinstance(r[0],(dt.datetime,dt.date)):
                kk=(r[0],r[1],r[2],r[4],r[17] if len(r)>17 else None)
                if kk not in seen_:                         # one count per MIS plan line
                    seen_.add(kk); mis_day.setdefault((pd.Timestamp(r[0]).normalize(),_line_key(r[1])),[]).append((cu,str(r[4]).strip(),r[3]))
    mis_bom={cu:max(d,key=d.get) for cu,d in cb.items()}
    def link(cu):
        n=' '.join(str(cu or '').split())
        for m in mis_bom:
            if m.lower()==n.lower(): return m
        t=_toks(n); best=None; sc=0
        for m in mis_bom:
            o=len(t & _toks(m))
            if o>sc: best,sc=m,o
        return best if sc>0 else None
    lines={}; mws={}; modplan={}; issues=[]
    models={}
    for r in pp_rows_std(wb['Production Plan']):
        if not r or not isinstance(r[0],(dt.datetime,dt.date)) or not isinstance(r[4],(int,float)) or not r[4]: continue
        k=idx.get(pd.Timestamp(r[0]).normalize())
        if k is None: continue
        ln=' '.join(str(r[1] or '').split()); cu=' '.join(str(r[3] or '').split())
        mdl=str(r[7] or '').strip() if len(r)>7 else ''
        if mdl and mdl not in ('0','#N/A'): models.setdefault(cu,{}); models[cu][mdl]=models[cu].get(mdl,0)+r[4]
        if PLAN_WSA_LINES_ONLY and not (ln.upper().startswith('USA') or re.fullmatch(r'L\d+',ln.upper())): continue
        key=(ln,cu); lines.setdefault(key,[0.0]*nd)[k]+=r[4]
        mw_=r[5] if len(r)>5 and isinstance(r[5],(int,float)) and not isinstance(r[5],bool) else (r[4]*r[2]/1e6 if isinstance(r[2],(int,float)) else 0.0)
        mws.setdefault(key,[0.0]*nd)[k]+=mw_                  # MW of the plan line (for the WSA 'MW per day' column)
    # LINK (most reliable first):  1) your override on the Settings sheet;  2) the two plans themselves - the MIS plan line
    # on the SAME DAY and the SAME PRODUCTION LINE (no names needed);  3) customer name;  4) Module Model of another customer
    d_of={k:d for d,k in idx.items()}; votes={}
    for (ln,cu),v in lines.items():
        lk=_line_key(ln)
        for k,q in enumerate(v):
            if not q: continue
            for mc_,bn_,mq in mis_day.get((d_of[k],lk),[]):
                w=min(q,mq)*(1.5 if (_toks(cu)&_toks(mc_)) else 1.0)
                votes.setdefault(cu,{}); votes[cu][(mc_,bn_)]=votes[cu].get((mc_,bn_),0)+w
    LINKS.clear()
    for cu in {c for (_,c) in lines}:
        ov=link_override(cu)
        if ov and ov[1]: LINKS[cu]=dict(mis=ov[0] or link(cu),bom=ov[1],how='your override (Settings)',conf=1.0); continue
        vt=votes.get(cu)
        if vt:
            (mc_,bn_),w=max(vt.items(),key=lambda kv:kv[1]); tot=sum(vt.values())
            LINKS[cu]=dict(mis=mc_,bom=bn_,how='same day + same line in the MIS plan',conf=w/tot); continue
        mc_=link(cu)
        if mc_: LINKS[cu]=dict(mis=mc_,bom=mis_bom.get(mc_),how='customer name',conf=None); continue
        LINKS[cu]=dict(mis=None,bom=None,how='not found',conf=None)
    for cu,L_ in LINKS.items():                           # 4) still nothing: a customer with the same Module Model
        if L_['bom'] is None and models.get(cu):
            mdl=max(models[cu].items(),key=lambda kv:kv[1])[0]
            for c2,L2 in LINKS.items():
                if c2!=cu and L2['bom'] and mdl in (models.get(c2) or {}): LINKS[cu]=dict(mis=L2['mis'],bom=L2['bom'],how=f'same Module Model as {c2}',conf=None); break
    rows=[]
    for (ln,cu),v in lines.items():
        mc=LINKS[cu]['mis']; bn=LINKS[cu]['bom']
        mdl_=max((models.get(cu) or {}).items(),key=lambda kv:kv[1])[0] if models.get(cu) else ''
        rows.append(dict(line=ln,cust=cu,mis_cust=mc,bom=bn,v=v,model=mdl_,mw=mws.get((ln,cu))))
        if bn: arr=modplan.setdefault(bn,{}).setdefault(mc,[0.0]*nd); [arr.__setitem__(k,arr[k]+v[k]) for k in range(nd)]
        else:
            issues.append(dict(area='Plan',check='Production Plan customer not linked to a MIS BOM Name',sev='High',material=cu,code=ln,a_lab='',a=None,
                               b_lab='Modules in horizon',b=sum(v),diff=None,detail=f'{ln} / {cu}: not found in the MIS plan (same day / line), by name or by Module Model - these modules get no raw-material plan',
                               action='Set its MIS BOM Name in Settings - MIS LINK'))
    for cu,L_ in LINKS.items():
        if L_['bom'] and str(L_['how']).startswith('same day') and not (_toks(cu)&_toks(L_['mis'] or '')) and not (L_['conf'] is not None and L_['conf']<0.6):
            issues.append(dict(area='Master data',check='Production Plan customer name differs from MIS - MIS name used',sev='Low',material=cu,code=L_['bom'],a_lab='Confidence',a=round(L_['conf'] or 0,2),
                               b_lab='',b=None,diff=None,detail=f"Your Production Plan says '{cu}', MIS says '{L_['mis']}' (same days / lines) - the MRP shows the MIS name",
                               action="Use the MIS name on your Production Plan when convenient"))
        if L_['bom'] and L_['conf'] is not None and L_['conf']<0.6:
            issues.append(dict(area='Plan',check='Customer link to MIS is uncertain',sev='Medium' if L_['conf']<0.4 else 'Low',material=cu,code=L_['bom'],a_lab='Confidence',a=round(L_['conf'],2),b_lab='',b=None,diff=None,
                               detail=f"{cu} -> {L_['mis']} / {L_['bom']} ({L_['how']}, {L_['conf']:.0%} of matching days)",action='Check it, or set it in Settings - MIS LINK'))
    rows.sort(key=lambda x:(str(x['bom']),x['line'],x['cust']))
    m2b={}                                                  # Module Model -> MIS BOM Name (most modules)
    for x in rows:
        for mdl,q in (models.get(x['cust']) or {}).items():
            if x['bom']: m2b.setdefault(mdl,{}); m2b[mdl][x['bom']]=m2b[mdl].get(x['bom'],0)+q
    MODEL2BN.clear(); MODEL2BN.update({m:max(d,key=d.get) for m,d in m2b.items()})
    return modplan,rows,issues

def line_plan(pp_rows):
    """{MIS BOM Name: {MIS customer: {production line: modules per day}}} from the Module Plan rows."""
    lp={}
    for x in pp_rows or []:
        if not x.get('bom') or not x.get('mis_cust'): continue
        d=lp.setdefault(x['bom'],{}).setdefault(x['mis_cust'],{})
        arr=d.setdefault(x['line'],[0.0]*len(x['v']))
        for k,v in enumerate(x['v']): arr[k]+=v
    return lp

def module_plan_sheet(modplan, dates, path, pp_rows=None, src=None):
    wb=Workbook(); ws=wb.active; ws.title='Module Plan'
    if pp_rows is not None:
        ppc_=os.path.basename(PPC_SRC[0]) if PPC_SRC[0] and 'Plan' in _sheetnames(PPC_SRC[0]) else None
        ws['A1']=((f"Modules per day from your PPC file ({ppc_}: sheet 'Plan', with sheet 'Revised Plan' on top) by Production Line and Customer. " if ppc_
                   else f'Modules per day from your Production Plan ({src}, columns A-F) by Production Line and Customer. ')
                  +'MIS Customer / MIS BOM Name = the MIS customer and module BOM the plan customer is linked to (Settings MIS LINK, else the closest name). '
                  +(f"Quantity per module = your PPC 'BOM' sheet; customers it does not list use the MIS Standard BOM. " if ppc_ else 'Quantity per module = your BOM sheet, else the MIS Standard BOM. ')
                  +'Which supplier a customer may use = Settings ALLOWED VENDORS.')
        hdr=['Production Line','Customer','MIS Customer','MIS BOM Name','Total']
        data=[[x['line'],x['cust'],x['mis_cust'],x['bom'],x['v']] for x in pp_rows]
    else:
        ws['A1']='Modules per day by MIS BOM Name and customer (from MIS Total Plan Req) - used with the BOM sheet for the CS / PM plan.'
        hdr=['MIS BOM Name','Customer','Total']
        data=[[bn,cu,v] for bn in sorted(modplan) for cu,v in sorted(modplan[bn].items(),key=lambda kv:-sum(kv[1]))]
    ws['A1'].font=Font(name='Arial',size=10,bold=True,color='1F3864')
    nfix=len(hdr); hdr=hdr+[pd.Timestamp(d).to_pydatetime() for d in dates]
    for j,h in enumerate(hdr,1):
        c=ws.cell(3,j,h); c.font=Font(name='Arial',size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='1F3864')
        if j>nfix: c.number_format='dd-mmm'
    for i,row in enumerate(data):
        v=row[-1]; r=4+i
        for j,x in enumerate(row[:-1],1):
            c=ws.cell(r,j,x)
            if x is None and j>=nfix-1: c.fill=PatternFill('solid',fgColor='FFC7CE')
        ws.cell(r,nfix,sum(v)).number_format='#,##0'
        for k,x in enumerate(v): ws.cell(r,nfix+1+k,x or None)
    for j,w in enumerate([18,26,24,26,11][:nfix],1): ws.column_dimensions[L(j)].width=w
    ws.freeze_panes=ws.cell(4,nfix+1).coordinate; wb.save(path)

# ===================== NEW MRP workbook (built from MIS + WSA only) =====================
CAP_DOC=999
def is_unlimited(b):
    t=f"{b.get('desc') or ''} {b.get('common') or ''}".lower()
    return b.get('mat') in UNLIMITED_CODES or any(w in t for w in UNLIMITED_WORDS)

# =================================================================
#  Sheet builders: Today and Planner
# =================================================================




def build_new_mrp(DS, BW, mis_path, out_path, prev_path=None, TD=None, old_mrp=None):
    """MRP_Daily_<date>.xlsx: ALERT, Material Status, MRP Daily, MRP by Line, MRP Weekly, Check, No Demand, Settings (see the sheet list printed at the end)
       (+ hidden: Line Run Days, Module Plan, BOM, Stock Count, How it works, Movement Log, MIS Layout).
       Target: every material covered for its own target days (Settings: A / B / C / packaging, + Z extra); TD = default for anything unclassified."""
    load_rmi_prev([out_path,prev_path])
    BLOCK_INFO.clear(); BLOCK_INFO.update({b_['mat']:(b_.get('uom'),b_.get('common')) for b_ in DS['blocks']})                        # your RM Issues entries (Owner / Status / Date Closed / Action) before anything is overwritten
    print(f'  [{time.time()-_T0:5.0f} s] building the MRP workbook ...',flush=True)
    TD=TD or TARGET_DAYS
    import statistics
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    nd=len(DS['dates']); d0=pd.Timestamp(DS['dates'][0]).normalize(); dates=DS['dates']
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF')
    # CS / PM from the user's BOM + Stock Count (kept in this workbook, seeded once from the old MRP)
    modplan=module_plan(mis_path,DS['plant'],dates); mis_names=sorted(modplan)
    ppm_=module_plan_from_pp(old_mrp,mis_path,DS['plant'],dates) if old_mrp else None   # your Production Plan drives the modules when available
    pp_rows=None; pp_issues=[]
    if ppm_ and ppm_[1]: modplan,pp_rows,pp_issues=ppm_
    lineplan=line_plan(pp_rows)                                   # modules by production line (for the Production Line column)
    MOD_SHORT.clear(); MOD_SHORT.update(module_short_names(mis_path,{(L2_ or {}).get('bom') for L2_ in LINKS.values()}|{x_.get('bom') for x_ in pp_rows or []}))
    LINE_DISP.clear(); LINE_ORD.clear(); _cnt={}
    for x_ in pp_rows or []:
        k_=_line_key(x_['line']); _cnt.setdefault(k_,{}); _cnt[k_][x_['line']]=_cnt[k_].get(x_['line'],0)+sum(x_['v'])
    for k_,d_ in _cnt.items(): LINE_DISP[k_]=' '.join(max(d_,key=d_.get).replace('USA ','').split())
    for i_,k_ in enumerate(sorted(LINE_DISP,key=lambda k_:(not re.match(r'^L\d',str(k_)),str(k_)))): LINE_ORD[k_]=i_
    UCMAP.clear()
    for x_ in pp_rows or []:
        k_=(_norm(x_.get('mis_cust')),_line_key(x_['line']))
        UCMAP.setdefault(k_,{})[x_['cust']]=UCMAP.setdefault(k_,{}).get(x_['cust'],0.0)+sum(x_['v'])
    LINE_MODS.clear()
    for x_ in pp_rows or []:
        a_=LINE_MODS.setdefault(_line_key(x_['line']),[0.0]*len(dates))
        for k,v in enumerate(x_['v'][:len(dates)]): a_[k]+=v
    bom,cnt_hdr,cnt_rows,bom_src=read_user_inputs(prev_path,old_mrp,mis_names,MODEL2BN)
    COUNT_ITEM.clear(); COUNT_ITEM.update(count_items(cnt_hdr,cnt_rows))
    sr0_=max([b['sr'] for b in DS['blocks'] if isinstance(b['sr'],(int,float))]+[0])
    if pp_rows:      # your Production Plan: customer -> your BOM (by customer name, else by Module Model)
        cs,cs_issues=cspm_from_bom_pp(bom,cnt_rows,pp_rows,dates,sr0_,cnt_hdr); pp_issues=list(pp_issues)+cs_issues
    else: cs=cspm_from_bom(bom,cnt_rows,modplan,dates,sr0_,cnt_hdr,lineplan)
    # unlimited materials (e.g. plastic bags): never short, so they are left out of the shortage planning
    unlimited=[b for b in cs if is_unlimited(b)]; cs=[b for b in cs if not is_unlimited(b)]
    # Customer Supplier rule: the Settings sheet first; else the Customer Supplier sheet of your MRP.xlsx; else the last MRP_Daily
    map_master=old_mrp if (old_mrp and read_supplier_map(old_mrp)) else None
    if SET_CUSTSUP: maprows=[dict(r_) for r_ in SET_CUSTSUP]; map_master=SETTINGS_SRC[0]
    else: maprows=read_supplier_map(map_master) if map_master else read_supplier_map(prev_path)
    mapped_codes={str(r['Code'] or '').strip() for r in maprows if r.get('Supplier')}
    if mapped_codes: attach_wsa(DS,BW,extra=mapped_codes)       # supplier detail also for mapped non-glass materials
    global SC_COL
    SC_COL=L(count_qty_idx(cnt_hdr)+1) if cnt_hdr else 'C'     # MRP Daily reads this column of its Stock Count sheet
    DS_mis=DS                                                     # MIS plan kept for the Check comparisons
    # RM plan = your Production Plan modules x MIS usage per module (RM usage always from MIS; your BOM is for CS / PM only).
    # MIS plan itself only for MIS customers your Production Plan does not have. CS / PM = your BOM x your modules (cspm_from_bom_pp).
    PACK_KG.clear(); UBOM_UNIT_ISSUES.clear()                        # pack sizes from the WSA sheet units (your BOM in KG -> drums ...)
    for sh_ in (BW.get('sheets') or {}).values():
        m_=re.search(r'\((\d+(?:\.\d+)?)\s*KG\s*/\s*(\w+)\)',str(sh_.get('uom') or '').upper())
        if m_:
            for nm_ in sh_.get('members') or []: PACK_KG[nm_]=float(m_.group(1))
    BOM_ALIAS[0]=DS.get('alias') or {}; BOM_ALIAS[1]=DS.get('alias_factor') or {}
    load_user_masters(SETTINGS_SRC[0]); build_alt_bom(mis_path,DS.get('alias') or {})       # alternatives per customer, before the plan
    DS_rm=dict(DS,blocks=rm_plan_from_mis(DS,mis_path,modplan,lineplan))
    load_user_masters(SETTINGS_SRC[0]); build_alt_bom(mis_path,DS.get('alias') or {})       # your alternative BOM: standard -> alternatives
    maprows,map_issues=apply_supplier_map(DS_rm,maprows,add_new=not maprows)   # rows are added only when the sheet does not exist yet
    CUSTSUP_ROWS[:]=maprows                                                     # shown on the Settings sheet
    vpri={}; valt=line_substitutes(DS_rm,mis_path); vp_issues=list(SUBS_ISSUES)   # substitutes for some lines (Settings)
    if valt: print(f"Substitutes for some lines: {sum(len(v) for v in valt.values())}")
    map_issues=map_issues+vp_issues
    for b in DS_rm['blocks']:
        for w_ in b.get('pack_window') or []:
            if w_['qty']<=0: continue
            kf=0 if w_['fr'] is None else max(0,(w_['fr']-d0).days); kt=len(dates)-1 if w_['to'] is None else (w_['to']-d0).days
            if kf>=len(dates) or kt<kf: continue
            b['adhoc']=list(b['adhoc']); b['adhoc'][kf]+=w_['qty']                     # other receipt: substitute available from 'use from'
            sub_=b.setdefault('adhoc_sub',{}).setdefault(str(w_.get('code') or ''),[0.0]*len(dates)); sub_[kf]+=w_['qty']   # (kept apart: a substitute GROUP counts the member's stock itself)
            used=min(w_['qty'],sum(b['plan'][kf:kt+1]))
            if w_['qty']-used>0.5 and kt+1<len(dates): b['adhoc'][kt+1]-=w_['qty']-used; sub_[kt+1]-=w_['qty']-used   # not usable any more after 'use until'
            w_['kf'],w_['kt'],w_['used']=kf,kt,used
            map_issues.append(dict(area='Stock',check='Pack conversion with dates (Settings)',sev='Low',material=b['common'],code=w_['code'],a_lab='Counted',a=round(w_['qty'],1),
                b_lab='Used inside the window',b=round(used,1),diff=None,detail=f"{w_['code']} counted for {b['mat']} from {dates[kf]:%d-%b}"+(f" until {dates[min(kt,len(dates)-1)]:%d-%b}" if w_['to'] is not None else '')+
                ' (shown on the Other receipt row; any unused part is taken out the day after)',action='Check the dates on the Settings sheet'))
    DS=dict(DS_rm,blocks=DS_rm['blocks']+cs)
    try:
        read_gr_lines(mis_path,DS_rm.get('alias'),DS_rm.get('alias_factor'))        # goods received per day (PPC GRN on GR. Date)
        read_count_history(SETTINGS_SRC[0] if SETTINGS_SRC else None,DS_rm.get('alias'),DS_rm.get('alias_factor'))   # stock on every count day
        estimate_counted_stock(DS['blocks'],mis_path,d0)                  # older physical counts rolled forward to today
        apply_count_by_supplier(DS['blocks'])                              # counted: stock by supplier = the count by supplier, rolled forward
        for m_,e_ in STOCK_EST.items():
            if e_.get('days') and e_['days']>0 and e_.get('est') is not None:
                u_={b_['mat']:b_.get('uom') for b_ in DS['blocks']}.get(m_,'')
                if e_.get('use'): DS.setdefault('stock_lines',[]).append(dict(code=m_,type='Stock',source=f"Usage since {'count' if e_.get('count') is not None else 'MIS date'} {pd.Timestamp(e_['date']):%d-%b} (estimated, {e_['days']} d × {e_['use_day']:,.1f}/d)",src=m_,batch=None,sloc=None,loc=None,qty=-float(min(e_['use'],(e_['count'] if e_.get('count') is not None else (e_['est']+e_['use']))+e_['rec'])),uom=u_,factor=1.0))
                if e_.get('rec'): DS.setdefault('stock_lines',[]).append(dict(code=m_,type='Stock',source=f"Received since count {pd.Timestamp(e_['date']):%d-%b} (MIS PPC GRN)",src=m_,batch=None,sloc=None,loc=None,qty=float(e_['rec']),uom=u_,factor=1.0))
        cds=[pd.Timestamp(e_['date']) for e_ in STOCK_EST.values() if e_.get('count') is not None and e_.get('date') is not None]
        if cds: DATA_DATES['Counts']=(min(cds),max(cds))
        n_=sum(1 for e_ in STOCK_EST.values() if e_.get('days') and e_.get('count') is not None)
        n2_=sum(1 for e_ in STOCK_EST.values() if e_.get('days') and e_.get('count') is None)
        if n_ or n2_: print(f"Stock rolled forward to today: {n_} physical counts (count date older than today), {n2_} SAP stocks (old MIS file)")
    except Exception as ex_: print(f'NOTE: count roll-forward skipped ({type(ex_).__name__}: {ex_})')
    segment(DS,mis_path)                                          # ABC / XYZ class and target days per material
    # ---------- per-material analysis ----------
    mis=wb_ro(mis_path)
    ships={}
    for r in list(mis['In Transit'].iter_rows(values_only=True))[1:]:
        if len(r)<40 or not isinstance(r[34],(int,float)) or r[34]<=0 or not it_ok(r): continue
        d=pd.Timestamp(r[35]).normalize() if isinstance(r[35],(dt.datetime,dt.date)) else None
        ships.setdefault((DS.get('alias') or {}).get(str(r[39] or '').strip(),str(r[39] or '').strip()),[]).append(dict(date=max(d,d0) if d is not None else None,raw=d,qty=r[34]*(DS.get('alias_factor') or {}).get(str(r[39] or '').strip(),1.0),sup=' '.join(str(r[1] or '').split()),po=r[6],shq=(r[32] if isinstance(r[32],(int,float)) else 0.0)*(DS.get('alias_factor') or {}).get(str(r[39] or '').strip(),1.0),bl=r[23] or r[24],vessel=r[12],
            grp=it_status_group(r[22]),
            shipped=bool(isinstance(r[27],(dt.datetime,dt.date)) or (isinstance(r[26],(dt.datetime,dt.date)) and pd.Timestamp(r[26]).normalize()<=d0)
                         or (r[23] or r[24]) or it_status_group(r[22]) in ('In transit','At port (POD)','Held (customs / issue)','Partly received - rest pending / in transit','Transloaded at warehouse'))))
    lead={}
    for x in BW.get('deliv',[]):
        etd=x.get('etd')
        if hasattr(etd,'year') and hasattr(x['date'],'year'): lead.setdefault(x['name'],[]).append((x['date']-etd).days)
    # ---- OEM: its production plan (customers with OEM = Yes) and the JW stock at the OEM are NOT part of any projection,
    #      shortage date, planned order, dashboard or risk list. The Action List still shows the with-OEM daily target and Inv. days.
    oem_c_=[_norm(cu) for cu,ln,tg,o in CUST_TARGET if str(o).strip().lower()=='yes']
    def _is_oem(cu):
        n_=_norm(cu); return bool(n_) and any(n_==o_ or n_.startswith(o_) or o_.startswith(n_) for o_ in oem_c_)
    OEM_IS[0]=_is_oem                                                    # ALERT Plan lines leave the OEM lines out (same rule)
    n_oem=0
    for b_ in DS['blocks']:
        cl_=b_.get('cust_line') or {}; ok_=[k_ for k_ in cl_ if _is_oem(k_[0] if isinstance(k_,tuple) else k_)]
        if ok_:
            arr_=[sum((cl_[k_][k] if k<len(cl_[k_]) else 0.0) for k_ in ok_) for k in range(nd)]
            b_['plan_oem']=arr_; b_['plan']=[max(0.0,p_-o_) for p_,o_ in zip(b_['plan'],arr_)]
            b_['cust_line_oem']={k_:cl_.pop(k_) for k_ in ok_}; n_oem+=1
        cp_=b_.get('cust_plan')
        if isinstance(cp_,dict):
            for k_ in [k_ for k_ in cp_ if _is_oem(k_)]: cp_.pop(k_)
        if not b_.get('count_stock') and abs(float(b_.get('st_jw') or 0))>0 and not b_.get('jw_out'):   # negative JW too: it was added to b['stock']
            b_['stock']=float(b_['stock'])-float(b_['st_jw']); b_['jw_out']=True
            if b_.get('st_main') is not None: pass
    try:
        CAP[0]=align_capture(DS,pp_rows,dates)                 # numbers for the WSA planning sheets and the Inbound 'Other materials' sheet
        if ALIGN_MODE[0]: ALIGN[0]=CAP[0]
    except Exception as ex_: print(f'NOTE: WSA numbers not captured ({type(ex_).__name__}: {ex_})')
    if ALIGN_MODE[0] and ALIGN[0] is not None: raise _CaptureDone()   # step 1 of the alignment needs only these numbers: the MRP_Daily it
                                                                     # would write goes to a temp folder and is thrown away (speed, Nov 2026)
    if oem_c_: print(f"OEM left out of the calculations: plan of {', '.join(cu for cu,ln,tg,o in CUST_TARGET if str(o).strip().lower()=='yes')} ({n_oem} materials) and JW stock at the OEM - Action List still shows the with-OEM target",flush=True)
    # ---- SUBSTITUTE GROUPS for the SHORTAGE projection: a material and its substitutes (Settings SUBSTITUTES, all lines, valid today)
    #      are projected as ONE: group stock, group in transit, group plan demand (each code x its factor = MRP units of the main material).
    #      Every member gets the same shortage date (shown in its own unit). Stock / In transit columns still show each code's own quantity.
    POOLG.clear(); byb_={b_['mat']:b_ for b_ in DS['blocks']}; grpb_={}
    for s_ in SUBS_X:
        m_,c_=s_.get('mrp_code'),s_.get('sub_code')
        if not m_ or not c_ or s_.get('lines') is not None or m_ not in byb_ or c_ not in byb_: continue
        if (s_.get('fr') is not None and s_['fr']>d0) or (s_.get('to') is not None and s_['to']<d0): continue
        if c_ not in [x_ for x_,_ in grpb_.get(m_,[])]: grpb_.setdefault(m_,[]).append((c_,float(s_.get('factor') or 1.0)))
    for m_,subs_ in grpb_.items():
        mem_=[(m_,1.0)]+subs_
        Pg=[sum((byb_[c_]['plan'][k] if k<len(byb_[c_]['plan']) else 0.0)*f_ for c_,f_ in mem_) for k in range(nd)]
        def _adh(c_,k):                                                 # ad-hoc receipts WITHOUT the stock of a member of this same group (counted below)
            v_=byb_[c_]['adhoc'][k]
            for sc_,arr_ in (byb_[c_].get('adhoc_sub') or {}).items():
                if sc_ in [m2 for m2,_ in mem_] and k<len(arr_): v_-=arr_[k]
            return v_
        Tg=[sum(((byb_[c_]['transit'][k]+_adh(c_,k)) if k<len(byb_[c_]['transit']) else 0.0)*f_ for c_,f_ in mem_) for k in range(nd)]
        Og=[sum(((byb_[c_].get('openpo') or [0.0]*nd)[k])*f_ for c_,f_ in mem_) for k in range(nd)]
        stg=sum(float(byb_[c_]['stock'])*f_ for c_,f_ in mem_)
        owng=sum((float(byb_[c_]['stock'])-(0.0 if (byb_[c_].get('count_stock') or byb_[c_].get('jw_out')) else float(byb_[c_].get('st_jw') or 0)))*f_ for c_,f_ in mem_)   # JW already out after the OEM step
        for c_,f_ in mem_:
            POOLG[c_]=dict(P=[v/f_ for v in Pg],T=[v/f_ for v in Tg],O=[v/f_ for v in Og],stock=stg/f_,own=owng/f_,tr=sum(Tg)/f_,members=[c2 for c2,_ in mem_],f=f_,
                           fx={c2:f2 for c2,f2 in mem_},names=[str(byb_[c2].get('common') or c2) for c2,_ in mem_])
    if POOLG: print('Substitute groups - shortage projected as one (group stock + in transit vs group demand): '+'; '.join(sorted({' + '.join(v['members']) for v in POOLG.values()})),flush=True)
    A=[]; VR_=_vendor_rules({'alias':DS.get('alias') or {}}); VP_MRP.clear()
    for b in DS['blocks']:
        TDm=b.get('td') or TD          # target days of this material (ABC / XYZ class)
        P=b['plan']; T=[a+c for a,c in zip(b['transit'],b['adhoc'])]; O=b.get('openpo') or [0.0]*nd
        stk_=b['stock']
        if b['mat'] in POOLG: g_=POOLG[b['mat']]; P,T,O,stk_=g_['P'],g_['T'],g_['O'],g_['stock']      # substitute group: one projection
        S=[]; cur=stk_
        for k in range(nd): cur=cur-P[k]+T[k]; S.append(cur)
        vp_=None; S_pool=list(S)
        if b['mat'] not in POOLG and any(k_[1]==b['mat'] for k_ in VR_):     # Settings ALLOWED VENDORS: same rule as the WSA planning sheet
            try:
                sh_=_wsa_sheet_series(b['mat'],d0,nd)
                # on 2+ WSA sheets (333): the sheets ARE the split by customer + supplier (Settings ALLOWED VENDORS, _wsa_shared) -
                # take it from them (checker finding: the MIS shipments alone gave Encap's FLAT HK glass to Subanci)
                vp_=_vp_from_sheets(S,sh_) if len(sh_)>1 else _mrp_vendor_pools(b,P,T,stk_,VR_,ships,d0,nd)
                if vp_: VP_MRP[b['mat']]=vp_; S=[S[k]-vp_['lock'][k] for k in range(nd)]
            except Exception as ex_: print(f"NOTE: Settings ALLOWED VENDORS not applied to {b['mat']} ({type(ex_).__name__}: {ex_})")
        cO=[]; acc=0.0
        for k in range(nd): acc+=O[k]; cO.append(acc)
        win=[sum(P[k:k+TDm]) for k in range(nd)]
        cov=cover_days(stk_,P,T,O)
        below=next((k for k in range(nd) if win[k]>0 and S[k]<win[k]),None)
        cover_from_first = None
        if below is not None:
            cur = S[below]
            for j in range(below, nd):
                if j > below:
                    cur = cur - P[j] + T[j]
                if cur < 0:
                    cover_from_first = j - below
                    break
            else:
                cover_from_first = nd - below
        below3=next((k for k in range(nd) if win[k]>0 and S[k]+cO[k]<win[k]),None)
        SW=min(nd,SHORT_HORIZON)    # shortage measured over the next N days (Settings) (orders beyond that are not placed yet, so far-future gaps would mislead)
        short=max([max(0.0,win[k]-S[k]) for k in range(SW)]+[0.0]); short3=max([max(0.0,win[k]-S[k]-cO[k]) for k in range(SW)]+[0.0])
        # shortage in DAYS = shortage qty / average demand per day of the next TDm days (the 'Avg demand / day' column),
        # i.e. how many days of demand must be added (sooner than now planned) to keep the TDm-day target for the next SHORT_HORIZON days
        _avg=sum(P[:TDm])/TDm
        short_d=(short/_avg) if _avg else 0.0; short3_d=(short3/_avg) if _avg else 0.0
        avg=sum(P[:TDm])/TDm
        nxt=sorted([s_ for s_ in ships.get(b['mat'],[]) if s_['date'] is not None],key=lambda s_:s_['date'])
        nx=None
        if nxt:                          # first arrival day: total of every shipment arriving that day
            same=[s_ for s_ in nxt if s_['date']==nxt[0]['date']]
            nx=dict(date=nxt[0]['date'],qty=sum(s_['qty'] for s_ in same),n=len(same),sup=', '.join(dict.fromkeys(s_['sup'] for s_ in same)))
        nopo=next(((d0+pd.Timedelta(days=k)),O[k]) for k in range(nd) if O[k]>0) if any(O) else None
        lt=int(statistics.median(lead[b['common']])) if lead.get(b['common']) else None
        cust=b.get('cust_plan') or {}
        clines={}                                 # customer -> {line: qty}  (which line makes it for which customer)
        if b.get('cust_line') and UCMAP:          # your customer names: MIS customer + line -> the customer(s) on that line in your Production Plan
            cust={}
            for (cu_,ln_),arr_ in b['cust_line'].items():
                us_=UCMAP.get((_norm(cu_),_line_key(ln_) if ln_ else None)) if ln_ else None
                tot_=sum(arr_); lnm_=_line_key(ln_) if ln_ else '?'
                if us_ and sum(us_.values())>0:
                    for u_,m_ in us_.items():
                        q_=tot_*m_/sum(us_.values()); cust.setdefault(u_,[0.0])[0]+=q_
                        clines.setdefault(u_,{}); clines[u_][lnm_]=clines[u_].get(lnm_,0.0)+q_
                else:
                    cust.setdefault(cu_,[0.0])[0]+=tot_; clines.setdefault(cu_,{}); clines[cu_][lnm_]=clines[cu_].get(lnm_,0.0)+tot_
        ct=sum(sum(v) for v in cust.values())
        psh_={}                                   # project (MIS LINK 'Project Name', else your customer) -> share of the plan
        for c_,v_ in cust.items():
            pn_=project_of(c_); psh_[pn_]=psh_.get(pn_,0.0)+sum(v_)
        pt_=sum(psh_.values())
        proj_txt='\n'.join(f'{k_} {q_/pt_:.0%}' for k_,q_ in sorted(psh_.items(),key=lambda kv:-kv[1]) if q_>0) if pt_ else ''
        lsh_={}                                   # line -> share of the plan (customer names left out)
        for c_,d_ in clines.items():
            for k_,q_ in d_.items(): lsh_[k_]=lsh_.get(k_,0.0)+q_
        lt_=sum(lsh_.values())
        custs='\n'.join(f"{LINE_DISP.get(k_,k_)} {q_/lt_:.0%}" for k_,q_ in sorted(lsh_.items(),key=lambda kv:-kv[1]) if q_>0) if lt_ else ''
        mods_={}                                  # module (MIS BOM name, short) -> share of the plan
        for c_,v_ in cust.items():
            bn_=(LINKS.get(c_) or {}).get('bom') or next((L2_.get('bom') for k2_,L2_ in LINKS.items() if _norm(k2_)==_norm(c_) or _norm(L2_.get('mis'))==_norm(c_)),None)
            nm_=MOD_SHORT.get(bn_) or bn_ or '(no MIS BOM link)'
            mods_[nm_]=mods_.get(nm_,0.0)+sum(v_)
        mods_txt='\n'.join(f'{m_} {q_/ct:.0%}' for m_,q_ in sorted(mods_.items(),key=lambda kv:-kv[1]) if q_>0) if ct else ''
        dt_=lambda k: (d0+pd.Timedelta(days=k)) if k is not None and k<nd else None
        LT=lt or (LEAD_TIME_CSPM if b.get('count_stock') else LEAD_TIME_RM)     # lead time: WSA ETD history, else Settings
        status=material_status(cov,TDm,below,LT)                               # the one status rule (top of the script)
        if status=='CRITICAL':
            need=dt_(below3) or d0; etd_=(need-pd.Timedelta(days=lt)) if lt else None
            action=(f"New supply needed: {short3:,.0f} {b['uom'] or ''} within {SHORT_HORIZON} days, first by {need:%d-%b}, to keep {TDm} days"+
                   (f"; to arrive in time it must ship by {etd_:%d-%b} (typical lead {lt} d)" if etd_ is not None else '')+
                   ('  - already too late by sea, consider air / transfer' if etd_ is not None and etd_<d0 else ''))
        elif status=='AT RISK':
            action=("Confirm vessel for the Open PO"+(f" ({nopo[1]:,.0f} {b['uom']}, plant date {nopo[0]:%d-%b})" if nopo else '')+
                    (f"; without it stock runs out {dt_(cov[1]):%d-%b}" if dt_(cov[1]) else ''))
        elif status=='WATCH' and cov[0]<TDm:
            action="Covered only if in-transit arrives on time"+(f" - next: {nx['qty']:,.0f} on {nx['date']:%d-%b} ({nx['sup'][:25]})" if nx else '')
        elif status=='WATCH':                                  # below target within the supply lead time: order now
            action=f"Falls below {TDm}-day target on {dt_(below):%d-%b} - order now (lead about {LT} d)"
        else:
            action=(f"Covered; next order needed to arrive by {dt_(below):%d-%b}" if below is not None else f"Covered for {TDm}+ days across the horizon")
        # shortage periods (stock + in transit): from the first day stock is below zero to the last day before it recovers
        wins=[]; k=0
        while k<nd:
            if S[k]<0:
                k2=k
                while k2+1<nd and S[k2+1]<0: k2+=1
                wins.append((k,k2,k2+1>=nd)); k=k2+1
            else: k+=1
        def wtxt(w):
            a_,z_,open_=w; f_=d0+pd.Timedelta(days=a_); t_=d0+pd.Timedelta(days=z_)
            return f"{f_:%d-%b} → beyond {t_:%d-%b} (no supply)" if open_ else (f"{f_:%d-%b} → {t_:%d-%b} ({z_-a_+1} d)")
        short_txt=(wtxt(wins[0])+(f"; +{len(wins)-1} more" if len(wins)>1 else '')) if wins else ''
        # ---- planned orders (lot-for-lot): receive whatever keeps projected stock (incl. Open PO) at the target, release lead time earlier ----
        por=[0.0]*nd; cum=0.0
        S_por=[(S_pool[k]-min(max(0.0,-S[k]),S_pool[k]-S[k])) if S[k]!=S_pool[k] else S[k] for k in range(nd)]   # vendor rule: order the gap of the short customers, not the stock they may not use (checker finding)
        for k in range(nd):
            if win[k]>0:
                g_=win[k]-(S_por[k]+cO[k]+cum)
                if g_>0.5: por[k]=g_; cum+=g_
        if grp_follower(b['mat']): por=[0.0]*nd                     # substitute group: the group's planned orders are on its main material only
        rel=[0.0]*nd; past=0.0
        for k in range(nd):
            if por[k]:
                j=k-LT
                if j<0: past+=por[k]; rel[0]+=por[k]
                else: rel[j]+=por[k]
        b['por']=por; b['rel']=rel; b['lt_used']=LT
        # first order that can still be placed in time (need date beyond the lead time) and the gap inside the lead time
        kp=next((k for k in range(LT,nd) if por[k]),None)
        nxt_po=dict(qty=sum(por[kp:kp+7]) if kp is not None else 0,need=dt_(kp),release=(dt_(kp-LT) if kp is not None else None),past=past)
        kr=(kp-LT) if kp is not None else None
        # ---- exception message (SAP MD06 style) ----
        st_out=wins[0][0] if wins else None
        late=[s_ for s_ in nxt_list if st_out is not None and s_['date'] is not None and (s_['date']-d0).days>st_out] if (nxt_list:=sorted([s_ for s_ in ships.get(b['mat'],[]) if s_['date'] is not None],key=lambda s_:s_['date'])) else []
        S30=S[min(29,nd-1)]+cO[min(29,nd-1)]; tgt30=win[min(29,nd-1)]
        k_in=next((k for k in range(min(LT,nd)) if por[k]),None)          # first gap inside the lead time (a new order is too late)
        if late and st_out is not None:
            s0_=late[0]; gap_=(s0_['date']-d0).days-st_out; u_=b.get('uom') or ''
            if s0_.get('shipped'):        # already on board / at port: the supplier cannot ship it earlier - speed up the last leg or bridge the gap
                where_={'At port (POD)':'at port','Held (customs / issue)':'held at customs','Transloaded at warehouse':'at warehouse'}.get(s0_.get('grp'),'on water')
                ves_=' '.join(str(s0_.get('vessel') or '').split())[:22] or s0_['sup'][:22]
                exc=(f"FAST-TRACK: out of stock {dt_(st_out):%d-%b}, {s0_['qty']:,.0f} {u_} {where_} arrives {s0_['date']:%d-%b} ({gap_} d late, {ves_}) - "
                     f"rush customs / trucking or cover {gap_} d (re-sequence, transfer)").replace('  ',' ')
            else:                         # not shipped yet: the supplier can still ship earlier
                exc=f"RESCHEDULE IN: ask {s0_['sup'][:22]} to ship {s0_['qty']:,.0f} {u_} earlier - arrives {s0_['date']:%d-%b}, out of stock {dt_(st_out):%d-%b}".replace('  ',' ')
            exg=1
        elif st_out is not None and st_out<LT:
            tr_=sum(b['transit'])
            exc=(f"EXPEDITE: short {wtxt(wins[0])}; "+(f"in transit {tr_:,.0f} {b.get('uom') or ''}, next {nx['date']:%d-%b} {nx['qty']:,.0f}" if tr_>0 and nx else 'nothing in transit')).replace('  ',' '); exg=1
        elif past>0.5: exc=f"EXPEDITE: {past:,.0f} {b['uom'] or ''} below target from {dt_(k_in):%d-%b} inside the {LT}-day lead time - pull receipts forward"; exg=2
        elif kr is not None and kr<=RELEASE_WINDOW: exc=f"RELEASE ORDER by {dt_(kr):%d-%b}: {sum(por[kp:kp+7]):,.0f} {b['uom'] or ''} for {dt_(kp):%d-%b}"; exg=3
        elif tgt30>0 and S30>RESCHED_OUT_X*tgt30 and (any(b['transit'][:30]) or any(O[:30])): exc=f"RESCHEDULE OUT: stock {S30/max(tgt30/TDm,1e-9):,.0f} days of demand on {dt_(29):%d-%b} with more arriving - push receipts"; exg=4
        elif kr is not None: exc=f"Next order release {dt_(kr):%d-%b}"; exg=5
        else: exc='No action'; exg=6
        A.append(dict(b=b,cov=cov,below=dt_(below),short_txt=short_txt,wins=wins,exc=exc,exg=exg,nxt_po=nxt_po,mods=mods_txt,proj=proj_txt,short=short,short3=short3,avg=avg,nx=nx,lt=lt,custs=custs,status=status,action=action,
                      runout=dt_(cov[1]) if cov[1]<nd else None,tr_tot=sum(b['transit']),tr_today=sum(s_['qty'] for s_ in ships.get(b['mat'],[]) if s_.get('raw') is not None and s_['raw']==d0),tr_last7=sum(s_['qty'] for s_ in ships.get(b['mat'],[]) if s_.get('raw') is not None and d0-pd.Timedelta(days=7)<=s_['raw']<d0),tr_next14=sum(s_['qty'] for s_ in ships.get(b['mat'],[]) if s_.get('raw') is not None and d0<s_['raw']<=d0+pd.Timedelta(days=14)),op_tot=sum(O),short_d=short_d,short3_d=short3_d,S=S,cO=cO,
                      cover_from_first=cover_from_first,
                      below3=dt_(below3),
                      doc=(S[0]/(sum(P[:TDm])/len(P[:TDm])) if sum(P[:TDm]) else float(CAP_DOC)),S_pool=S_pool))   # MRP Daily 'Days of cover' on day 1 (own target window)
    A.sort(key=lambda x:(STATUS_ORDER[x['status']],x.get('exg',9),x['cov'][1],x['cov'][2],bool(x['b'].get('count_stock'))))   # most urgent first (status, exception, days)
    if VP_MRP: print('Settings ALLOWED VENDORS - projection by allowed vendor (same rule as the WSA sheets): '+vp_console(VP_MRP,d0),flush=True)
    # ---------- base workbook: Action List (the old Dashboard was removed: headline / top risks -> ALERT) ----------
    wb=Workbook(); al=wb.active; al.title='Action List'
    al.sheet_properties.tabColor='C00000'; set_print(al,'Action List')
    n=len(A);
    # allocation to lines: Customer Supplier > Vendor Priority (vendor) > no preference > Vendor Priority (material description);
    # lines served in order of potential stoppage (pooled view first, most urgent line first)
    LR0=line_run_days(A,d0,nd)
    LINES_SEEN[:]=[r_['name'] for r_ in LR0]
    # fixed line ranking from the Settings sheet first (1 = served first); unranked lines after, closest to stopping first
    line_rank={r_['key']:(LINE_RANK[r_['key']] if r_['key'] in LINE_RANK else 1000+i) for i,r_ in enumerate(LR0)}
    cmap={}
    for r_ in maprows:
        if r_.get('Supplier'): cmap.setdefault((_norm(r_['Customer']),str(r_['Code'] or '').strip()),[]).append(str(r_['Supplier']).strip().upper())
    SHARED_MATS[:]=[]
    for x in A:
        lns={_line_key(ln) for (cu,ln),arr in (x['b'].get('cust_line') or {}).items() if ln and any(arr)}
        if len(lns)>=2: SHARED_MATS.append((x['b']['desc'] if x['b'].get('count_stock') else x['b']['common'],x['b']['desc'],x['b']['mat'],lns))
    SHARED_MATS.sort(key=lambda t:(t[2][:2]!='RM',str(t[0])))
    mat_rank={}
    for k,v in LINE_RANK_MAT.items():                  # material typed as code, common name or description
        kk=str(k).strip().upper()
        m=next((b['mat'] for b in DS['blocks'] if b['mat'].upper()==kk),None) or next((b['mat'] for b in DS['blocks'] if str(b.get('common') or '').upper()==kk or str(b.get('desc') or '').upper()==kk),None)
        if m: mat_rank[m]={a:r for a,r in v.items() if not a.startswith('_')}
    try:                                                            # your alternative BOM counts for 'Line can run'
        for k_,v_ in alt_bom_substitutes(A,DS,mis_path,pp_rows,nd).items(): valt.setdefault(k_,[]).extend(v_)
    except Exception as ex_: print(f'NOTE: alternatives not used for line run days ({type(ex_).__name__}: {ex_})')
    alloc_stop,alloc_used=allocate_lines(DS,vpri,valt,cmap,line_rank,nd,mat_rank)
    LR=line_run_days(A,d0,nd,alloc_stop)
    for x in A:
        lks={_line_key(ln) for (cu,ln) in (x['b'].get('cust_line') or {}) if ln}
        x['used']={lk:alloc_used.get((x['b']['mat'],lk)) for lk in lks}
        x['b']['line_src']={(_norm(cu),_line_key(ln)):alloc_used.get((x['b']['mat'],_norm(cu),_line_key(ln))) for (cu,ln) in (x['b'].get('cust_line') or {}) if ln}
        x['b']['line_stop']={lk:alloc_stop.get((x['b']['mat'],lk)) for lk in lks}      # for the line rows of MRP Daily
    # ---- how many days stock + in transit lasts when the lines run at their daily targets (Settings) ----
    # Only lines that run a product using the material TODAY count (today = the line's first production day from the run date).
    # Need / day = sum over those lines of: daily target x usage per module of that product (material / modules of that product
    # on that line that day) - i.e. the whole daily target is assumed to be that product (peak), no averaging over days or mix.
    ppm_={}
    for x_ in pp_rows or []:
        for nm_ in {_norm(x_.get('mis_cust')),_norm(x_.get('cust'))}:
            if not nm_: continue
            a_=ppm_.setdefault((nm_,_line_key(x_['line'])),[0.0]*len(dates))
            for k,v in enumerate(x_['v'][:len(dates)]): a_[k]+=v
    EFF_TARGET.clear(); EFF_TARGET.update(LINE_TARGET)
    lname_={_line_key(x_['line']):x_['line'] for x_ in pp_rows or []}
    for lk,arr_ in LINE_MODS.items():          # no Settings target: the line's highest planned day in your Production Plan
        if lk not in EFF_TARGET and arr_ and max(arr_)>0: EFF_TARGET[lk]=(f"{lname_.get(lk,lk)} (plan peak)",float(max(arr_)))
    LINE_TODAY.clear()
    for lk,(lab,tg) in EFF_TARGET.items():
        lm_=LINE_MODS.get(lk) or []
        LINE_TODAY[lk]=next((k for k in range(len(lm_)) if lm_[k]>0),None)
    if not CUST_TARGET: CUST_TARGET[:]=list(CUST_TARGET_DEFAULT)
    pp_by_cust={}
    for x_ in pp_rows or []: pp_by_cust.setdefault(_norm(x_['cust']),x_)
    def canon_cust(cu_):                       # Settings spelling -> the spelling used in your BOM / MIS LINK ('Clenera Crimson' = 'Cleanera Crimson')
        import difflib
        k_=_norm(cu_)
        if k_ in UBOM_RM or any(_norm(n_)==k_ for n_ in LINKS) or any(k_==_norm(r_.get('cust')) for r_ in pp_rows or []): return cu_
        pool_={_norm(n_):n_ for n_ in LINKS}; pool_.update({k2_:k2_ for k2_ in UBOM_RM})
        m_=difflib.get_close_matches(k_,list(pool_),n=1,cutoff=0.85)
        if m_ and (cu_,pool_[m_[0]]) not in CUST_FIX_: CUST_FIX_.append((cu_,pool_[m_[0]]))
        return pool_[m_[0]] if m_ else cu_
    def cust_row(cu_):                         # customer -> its MIS BOM name / Module Model (for the BOM usage)
        cu_=canon_cust(cu_)
        L_=next((v for k,v in LINKS.items() if _norm(k)==_norm(cu_)),None) or {}
        r_=pp_by_cust.get(_norm(cu_)) or {}
        return dict(cust=cu_,bom=L_.get('bom') or r_.get('bom'),model=r_.get('model'))
    for x in A:                                # YOUR PPC WAY: need = sum of customer daily target x usage per module
        b=x['b']; nw=0.0; nwo=0.0; parts_={}; use_={}; cl_=[]
        src_=[]
        for cu_,ln_,tg_,oem_ in CUST_TARGET:
            u_,s_=bom_usage_src(b['mat'],bool(b.get('count_stock')),cust_row(cu_))
            if not u_: continue
            src_.append(f"{cu_}: {u_:g} {b.get('uom') or ''} ({s_})".replace('  ',' '))
            q_=u_*tg_; nw+=q_; lk_=_line_key(ln_) if ln_ else ln_
            if oem_!='Yes': nwo+=q_
            parts_[lk_]=parts_.get(lk_,0.0)+q_; use_[lk_]=max(use_.get(lk_,0.0),u_); cl_.append(lk_)
        x['usage_src']='\n'.join(src_) or None
        x['tgt_need_c']=nw; x['tgt_need_wo_c']=nwo; x['tgt_parts_c']=parts_; x['tgt_use_c']=use_; x['tgt_lines_c']=list(dict.fromkeys(cl_))
    for x in A:
        b=x['b']; need=0.0; used_=[]; parts_={}; use_={}
        for lk,(lab,tg) in sorted(EFF_TARGET.items()):
            k=LINE_TODAY.get(lk)
            if not tg or k is None: continue
            best=0.0; kend=min(len(dates),k+int(LOOKAHEAD_DAYS))    # from the line's first production day, LOOKAHEAD_DAYS ahead (Settings)
            for x_ in pp_rows or []:                    # every customer / module type this line runs in the window
                if _line_key(x_['line'])!=lk or not any(x_['v'][j]>0 for j in range(k,min(kend,len(x_['v'])))): continue
                best=max(best,bom_usage(b['mat'],bool(b.get('count_stock')),x_))   # usage per module from the BOM
            if best>0: need+=best*tg; used_.append(lk); parts_[lk]=best*tg; use_[lk]=best
        x['tgt_need']=need; x['tgt_lines']=used_; x['tgt_parts']=parts_; x['tgt_use']=use_
        if CUSTOMER_TARGETS:                  # customer targets (PPC way) replace the line-target need
            x['tgt_need']=x['tgt_need_c']; x['tgt_lines']=x['tgt_lines_c']; x['tgt_parts']=x['tgt_parts_c']; x['tgt_use']=x['tgt_use_c']
        x['tgt_days']=(max(0.0,b['stock']+(x.get('tr_tot') or 0))/x['tgt_need']) if x.get('tgt_need') else None   # with OEM target; JW stock at the OEM ignored (user rule, Nov 2026)
    # ---- days at your DAILY PRODUCTION PLAN (not the targets): need = plan requirement on the first production day from today;
    #      days = how long stock + all in transit lasts when consumed day by day by the plan (calendar days, max = plan horizon)
    for x in A:
        b=x['b']; p_=b.get('plan') or []; k0=next((k for k in range(len(p_)) if p_[k]>0),None)
        x['plan_need']=p_[k0] if k0 is not None else None
        left=max(0.0,b['stock']+(x.get('tr_tot') or 0)); dd_=None
        if k0 is not None:
            for k in range(len(p_)):
                if p_[k]>left: dd_=k+(left/p_[k] if p_[k] else 0); break
                left-=p_[k]
            if dd_ is None: dd_=float(len(p_))
        x['plan_days']=dd_
    if CUST_FIX_: print('Customer name matched to your BOM / MIS LINK: '+'; '.join(f"'{a_}' -> '{b_}'" for a_,b_ in CUST_FIX_)+'  (fix the spelling on the Settings sheet to remove this note)',flush=True)
    # ---- SUBSTITUTE GROUPS: a material and its substitutes (Settings SUBSTITUTES, all lines, valid today) share one stock pool, so each
    #      member shows the demand of the WHOLE group (own daily target + the target of the others, converted with the factor).
    #      Codes that are not planned materials themselves (e.g. RMBUS00000005 / 33 -> RMBUS00000048) are already folded into the material (stock, plan, usage).
    byc_={x['b']['mat']:x for x in A}; POOL_[:]=[]; grp_={}
    for s_ in SUBS_X:
        m_,c_=s_.get('mrp_code'),s_.get('sub_code')
        if not m_ or not c_ or s_.get('lines') is not None or m_ not in byc_ or c_ not in byc_: continue
        if (s_.get('fr') is not None and s_['fr']>d0) or (s_.get('to') is not None and s_['to']<d0): continue
        grp_.setdefault(m_,[]).append((c_,float(s_.get('factor') or 1.0)))
    for m_,subs_ in grp_.items():
        mem_=[(m_,1.0)]+subs_                                   # (code, MRP units per 1 unit of this code)
        old_={c_:{k_:(byc_[c_].get(k_) or 0.0) for k_ in ('tgt_need','tgt_need_c','tgt_need_wo_c')} for c_,_f in mem_}
        oldp_={c_:{k_:dict(byc_[c_].get(k_) or {}) for k_ in ('tgt_parts','tgt_parts_c')} for c_,_f in mem_}
        for c_,f_ in mem_:
            x_=byc_[c_]; own_=old_[c_]['tgt_need']
            for k_ in ('tgt_need','tgt_need_c','tgt_need_wo_c'): x_[k_]=sum(old_[o_][k_]*fo_ for o_,fo_ in mem_)/f_
            for k_ in ('tgt_parts','tgt_parts_c'):
                tot_={}
                for o_,fo_ in mem_:
                    for ln_,q_ in oldp_[o_][k_].items(): tot_[ln_]=tot_.get(ln_,0.0)+q_*fo_
                x_[k_]={ln_:q_/f_ for ln_,q_ in tot_.items()}
            x_['tgt_lines']=list(dict.fromkeys(list(x_.get('tgt_lines') or [])+[l_ for o_,_f in mem_ for l_ in (byc_[o_].get('tgt_lines') or [])]))
            g_=POOLG.get(c_);
            st_=((g_['stock']+g_['tr']) if g_ else (x_['b']['stock']+(x_.get('tr_tot') or 0)))          # JW stock at the OEM ignored (user rule, Nov 2026)
            x_['tgt_days']=(max(0.0,st_)/x_['tgt_need']) if x_.get('tgt_need') else None          # group stock + in transit ÷ group need
            oth_=[f"+ {o_} {str(byc_[o_]['b'].get('common') or '')}: {old_[o_]['tgt_need']*fo_/f_:,.0f} {x_['b'].get('uom') or ''} per day (substitute group)" for o_,fo_ in mem_ if o_!=c_ and old_[o_]['tgt_need']]
            if oth_: x_['usage_src']='\n'.join([t_ for t_ in [x_.get('usage_src')] if t_]+oth_)
            POOL_.append((c_,own_,x_['tgt_need']))
    if POOL_: print('Substitute groups - daily target pooled: '+'; '.join(f"{c_} {a_:,.0f} -> {b_:,.0f}" for c_,a_,b_ in POOL_),flush=True)
    # ---- NEED PER DAY FROM THE MODULE PLAN (every sheet except the Action List, which keeps the daily targets / OEM):
    #      pl_need = average plan per production day (> 0) over the next 9 weeks (OEM plan already left out; substitute group pooled);
    #      pl_days = (own stock + in transit) / pl_need  (no JW stock at the OEM)
    for x in A:
        b=x['b']; g_=POOLG.get(b['mat'])
        P_=(g_['P'] if g_ else b.get('plan')) or []; nz_=[v for v in P_[:min(nd,avg_prod_days())] if v>0]
        x['pl_need']=(sum(nz_)/len(nz_)) if nz_ else None
        st_=(g_['stock']+g_['tr']) if g_ else (float(b['stock'])+(x.get('tr_tot') or 0))
        x['pl_days']=(max(0.0,st_)/x['pl_need']) if x['pl_need'] else None
    # ---- NO DEMAND rule: a material is only Critical / At risk / Watch / Short when it has demand. Demand = at least one of
    #      production requirement / BOM demand (plan x usage) inside the demand horizon, planned consumption in the daily-plan window, daily target need.
    #      Otherwise 'No demand': no shortage status, no planned order / release, no Expedite / Order now, not on any risk list.
    DH_=max(1,min(nd,int(DEMAND_HORIZON or SHORT_HORIZON)))
    NODEM_[:]=[]
    for x in A:
        b=x['b']; P_=b.get('plan') or []
        x['demand_h']=sum(P_[:DH_])
        x['no_demand']=not (x['demand_h']>0 or sum(P_[:int(PLAN_AVG_DAYS)])>0)          # module plan only (daily targets = Action List only)
        if x['no_demand']:
            b['por']=[0.0]*nd; b['rel']=[0.0]*nd
            x['nxt_po']=dict(qty=0,need=None,release=None,past=0)
            x['status']='NO DEMAND'; x['exc']='No demand - no action'; x['exg']=7; x['action']='No demand - no shortage action'; x['short_txt']=''
            f_=next((k for k,v in enumerate(P_) if v>0),None)
            NODEM_.append((b['mat'],str(b.get('common') or b.get('desc') or ''),f"first plan demand in {f_} d ({sum(P_):,.0f} {b.get('uom') or ''} in the plan)" if f_ is not None else 'no plan demand at all'))
    if NODEM_: print(f"No demand (no shortage status / action): {len(NODEM_)} materials - "+'; '.join(f"{c_} {n_[:18]} ({w_})" for c_,n_,w_ in NODEM_[:12]),flush=True)
    for x in A:
        x['cw']=cover_watch(x,d0,ships,nd,need=x.get('tgt_need') or x.get('pl_need'))   # Action List: daily TARGET scenario (with OEM target)
        x['cw_p']=cover_watch(x,d0,ships,nd)                                            # Action List: MODULE PLAN scenario (= ALERT 'Below 21 days from')
    # ---- SCENARIOS for the Action List
    #  with OEM    = all lines (Settings targets / highest planned day for lines without one), stock = MB52 + JW stock at the OEM
    #  without OEM = USA Lines 1-4 only, stock = MB52 only (CS / PM: Stock Count 'Today Qty_BOM' in both)
    NO_OEM_LINES={'L1','L2','L3','L4'}
    for x in A:
        b=x['b']; tr_=x.get('tr_tot') or 0
        x['oem_stock']=0.0 if b.get('count_stock') else (b.get('st_jw') or 0.0)
        x['own_stock']=b['stock'] if b.get('jw_out') else b['stock']-x['oem_stock']      # JW stock is already out of b['stock'] (OEM left out)
        x['tgt_need_wo']=x['tgt_need_wo_c'] if CUSTOMER_TARGETS else sum(v for k,v in (x.get('tgt_parts') or {}).items() if k in NO_OEM_LINES)
        g_=POOLG.get(b['mat'])
        x['tgt_days_wo']=(max(0.0,(g_['own']+g_['tr']) if g_ else (x['own_stock']+tr_))/x['tgt_need_wo']) if x['tgt_need_wo'] else None
        nd_=len(b.get('plan') or [])
        p_wo=[0.0]*nd_
        for (cu_,ln_),arr_ in (b.get('cust_line') or {}).items():
            if ln_ and _line_key(ln_) in NO_OEM_LINES:
                for k in range(min(nd_,len(arr_))): p_wo[k]+=arr_[k]
        avg21=lambda p__: (lambda v__: sum(v__)/len(v__) if v__ else None)([q for q in p__[:PLAN_AVG_DAYS] if q>0])
        x['plan_need_wo']=avg21(p_wo); x['plan_need']=avg21(b.get('plan') or [])
        x['plan_days_wo']=(max(0.0,x['own_stock']+tr_)/x['plan_need_wo']) if x['plan_need_wo'] else None
        x['plan_days']=(max(0.0,b['stock']+tr_)/x['plan_need']) if x.get('plan_need') else None
        # every MIS arrival date (Final AtA) with its pending quantity; dates already passed = not yet received
        by_d={}
        for s_ in ships.get(b['mat'],[]):
            if s_.get('raw') is not None: by_d[s_['raw']]=by_d.get(s_['raw'],0.0)+s_['qty']
        # Arrival at plant, one line each: dates already passed (not yet received) as one total, the next arrivals, the rest as one total
        past_=[(d_,q_) for d_,q_ in sorted(by_d.items()) if q_ and d_<d0]; fut_=[(d_,q_) for d_,q_ in sorted(by_d.items()) if q_ and d_>=d0]
        ln_=[]
        if past_: ln_.append(f"Passed, not received: {sum(q for _,q in past_):,.0f}"+(f" ({past_[0][0]:%d-%b})" if len(past_)==1 else f" ({past_[0][0]:%d-%b} → {past_[-1][0]:%d-%b})"))
        nx_=[f"{d_:%a %d-%b}: {q_:,.0f}" for d_,q_ in fut_[:ATA_SHOW]]
        for i_ in range(0,len(nx_),2): ln_.append('   |   '.join(nx_[i_:i_+2]))      # two arrival dates per line
        if len(fut_)>ATA_SHOW: rest_=fut_[ATA_SHOW:]; ln_.append(f"+{len(rest_)} more:  {sum(q for _,q in rest_):,.0f}  ({rest_[0][0]:%d-%b} → {rest_[-1][0]:%d-%b})")
        x['ata_txt']='\n'.join(ln_) or None
        # same, with the supplier of each arrival (RM / PM Status shows this for glass)
        sup_d={}
        for s_ in ships.get(b['mat'],[]):
            if s_.get('raw') is not None: sup_d.setdefault(s_['raw'],[]).append(sup_short(s_.get('sup')))
        sp_=lambda d_: f" ({', '.join(dict.fromkeys(sup_d.get(d_,[])))})" if sup_d.get(d_) else ''
        ls_=[]
        if past_: ls_.append(f"Passed, not received: {sum(q for _,q in past_):,.0f}"+(f" ({past_[0][0]:%d-%b})" if len(past_)==1 else f" ({past_[0][0]:%d-%b} → {past_[-1][0]:%d-%b})")
                             +f" - {', '.join(dict.fromkeys(n for d_,_ in past_ for n in sup_d.get(d_,[])))}")
        nxs_=[f"{d_:%a %d-%b}: {q_:,.0f}{sp_(d_)}" for d_,q_ in fut_[:ATA_SHOW]]
        for i_ in range(0,len(nxs_),2): ls_.append('   |   '.join(nxs_[i_:i_+2]))
        if len(fut_)>ATA_SHOW: rest_=fut_[ATA_SHOW:]; ls_.append(f"+{len(rest_)} more:  {sum(q for _,q in rest_):,.0f}  ({rest_[0][0]:%d-%b} → {rest_[-1][0]:%d-%b})")
        x['ata_txt_sup']='\n'.join(ls_) or None
    # ---- WSA planning file: Plant Days / Plant + Port Days on the sheet's stock date (pooled sheets: same for all its materials)
    wsa_by_name={}
    for lab_,sh_ in (BW.get('sheets') or {}).items():
        for m_ in sh_.get('members') or []: wsa_by_name.setdefault(str(m_).strip().lower(),sh_)
    for x in A:
        sh_=wsa_by_name.get(str(x['b']['common']).strip().lower())
        x['wsa_pd']=sh_.get('plantdays') if sh_ else None; x['wsa_ppd']=sh_.get('plantport') if sh_ else None
        x['wsa_sheet']=f"{sh_['sheet']} ({sh_['asof']:%d-%b})" if sh_ and sh_.get('asof') else None
        for key_,i_ in (('wsa_pd7',0),('wsa_ppd7',1)):     # first Required Plant Date from today with Plant Days (or Plant + Port Days) < 7
            x[key_]=None
            for dd_,v_ in sorted((sh_ or {}).get('pdays',{}).items()):
                if dd_>=d0.to_pydatetime() and isinstance(v_[i_],(int,float)) and v_[i_]<7: x[key_]=dd_; break
    # ---- WSA planning file: quantity in transit today (deliveries at stage 'In transit')
    wsa_it={}
    for e_ in BW.get('deliv') or []:
        if e_.get('stage')=='In transit': wsa_it[str(e_['name']).strip().lower()]=wsa_it.get(str(e_['name']).strip().lower(),0.0)+(e_.get('qty') or 0)
    for x in A: x['wsa_it']=wsa_it.get(str(x['b']['common']).strip().lower())
    # Action List
    al['A1']='Action List — every material by shortage risk'; al['A1'].font=Font(name='Calibri',size=20,bold=True,color='14213D')
    al['A2']=('Highest shortage risk first. Priority (= RM Issues): High = stock zero inside the lead time, Medium = short later, Low = under 21 days, OK. Stock lines (source, factor, storage location): Stock Data sheet. Short from = first day stock (count rolled forward, or SAP) − plan demand + MIS in transit < 0; ALERT = inside the lead time. '
              'Daily target need = customer daily targets × usage per module (your BOM first, else MIS Standard BOM); without OEM (green banner) = customers with OEM = No, with OEM (blue banner) adds the OEM and JW stock at OEM. '
              'Inv. Days = (stock + MIS in transit) ÷ daily target need, red under 21. '
              f'Below {REPORT_DAYS} Days From / Days Late = the ALERT logic (Short from → to holds every stock-zero period). No demand = no production requirement, planned consumption or daily target in the next {int(DEMAND_HORIZON or SHORT_HORIZON)} days: no shortage status and no order action. '
              'Details per column: Definitions sheet.')
    al['A2'].font=Font(name='Calibri',size=9,color='8A949E'); al.merge_cells('A2:S2'); al['A2'].alignment=Alignment(wrap_text=True,vertical='top'); al.row_dimensions[2].height=54
    wtxt_=lambda x: ('; '.join(f"{(d0+pd.Timedelta(days=w[0])):%d-%b} → "+(f"beyond {(d0+pd.Timedelta(days=w[1])):%d-%b}" if w[2] else f"{(d0+pd.Timedelta(days=w[1])):%d-%b} ({w[1]-w[0]+1} d)") for w in x.get('wins') or []) or 'no stock-out')
    def short_lt(x):                     # short window + the material's lead time; ALERT when the shortage starts inside the lead time
        if x.get('no_demand'): return 'No demand'
        t_=wtxt_(x); lt_=x['b'].get('lt_used'); w_=x.get('wins') or []
        if not w_: return t_
        k_=w_[0][0]
        if lt_ is not None and k_<lt_: return f"ALERT inside lead time: {t_}  |  lead {lt_} d - short in {k_} d"
        return f"{t_}  |  lead {lt_} d - order in time ({k_-lt_} d spare)" if lt_ is not None else t_
    stk_raw_=DS.get('stock_raw') or {}; nm_={x_['b']['mat']:x_['b']['common'] for x_ in A}; cnr_=DS.get('cn_all') or {}
    def alt_txt(x):                       # approved alternatives of this standard material (your BOM) with their MB52 stock
        al_=ALT_BOM.get(x['b']['mat']) or {}
        out_=[]
        for c_,cus_ in sorted(al_.items()):
            nm=nm_.get(c_) or str((cnr_.get(c_) or [None]*4)[3] or c_)
            out_.append(f"{c_} {nm}: {stk_raw_.get(c_,0.0):,.0f} ({', '.join(sorted(cus_))})")
        return '\n'.join(out_) or None
    def _desc(x):
        b=x['b']; c_=str(b.get('common') or '').strip(); d_=' '.join(str(b.get('desc') or '').split())
        return d_ if (not c_ or c_.lower() in d_.lower()) else (f"{c_} | {d_}" if d_ else c_)
    rnd=lambda k: (lambda x: round(x[k],1) if isinstance(x.get(k),(int,float)) and x.get(k) else None)
    rndz=lambda k: (lambda x: round(x[k],1) if isinstance(x.get(k),(int,float)) else None)
    # stock source lines per material (same as the Stock Data sheet), one line per source code
    sdg_=SDG; sdg_.clear()
    for l_ in DS.get('stock_lines') or []:
        g_=sdg_.setdefault(l_['code'],{}); k_=(l_['type'],l_['source'],l_['src'])
        e_=g_.setdefault(k_,dict(qty=0.0,uom=l_.get('uom'),factor=l_.get('factor') or 1.0,sl={}))
        e_['qty']+=l_['qty']; s__=' '.join(str(l_.get('sloc') or '').split()[1:]) or str(l_.get('loc') or '') or ''
        if s__: e_['sl'][s__]=e_['sl'].get(s__,0.0)+l_['qty']*(l_.get('factor') or 1.0)
    for x_ in A:
        if x_['b'].get('count_stock'):
            sdg_[x_['b']['mat']]={('Stock',"PPC Count (physical count)",x_['b']['mat']):dict(qty=x_['b']['stock'],uom=x_['b'].get('uom'),factor=1.0,sl={})}
    def _cwv(x,k,key='cw'):              # cover_watch: date / days / '-' / 'No demand'  (cw = daily target, cw_p = module plan)
        if x.get('no_demand'): return 'No demand'
        cw=x.get(key)
        if x.get('sup_row') or cw is None: return None
        if k=='below': return cw['below'].to_pydatetime() if cw['below'] is not None else '-'
        if k=='short': return short_lt_text(x,d0)[0]                        # all stock-zero periods inside the lead time
        return cw['late'] if cw['late'] is not None else '-'
    def prio_(x):                         # = RM Issues Priority: High = stock zero inside the lead time, Medium = short later, Low = under 21 days
        ri_=risk_info(x,d0)
        if ri_['cls']=='No demand': return 'No demand'
        if ri_['cls']=='ALERT' or short_lt_text(x,d0)[1]: return 'High'
        return {'Short later':'Medium','Under 21 d':'Low'}.get(ri_['cls'],'OK')
    # Nov 2026 redesign (user: 'Action List ... look like sheet RM Issues'): RM Issues style, no banner row; the stock-source detail
    # columns (Source, Source code, Qty, UOM, Factor, Storage Loc.) are on Stock Data and were dropped; Priority added
    COLS=[('Component',lambda x:x['b']['mat'],None,16,False),
          ('Component description',_desc,None,34,True),
          ('UOM',lambda x:x['b']['uom'],None,6,False),
          ('Priority',prio_,None,10,False),
          ('Lines (share of plan volume)',lambda x:x['custs'],None,16,True),
          ('Project Name (share of plan)',lambda x:x.get('proj'),None,30,True),
          ('Usage per module by customer (source: your BOM / MIS Standard BOM)',lambda x:x.get('usage_src'),None,46,True),
          ('Module (MIS BOM, share of plan)',lambda x:x.get('mods'),None,16,True),
          ('Short from → to (with in transit)',short_lt,None,40,True),
          (f'Below {REPORT_DAYS} Days From',lambda x:_cwv(x,'below'),'dd-mmm-yy',12,False),
          ('Days Late',lambda x:_cwv(x,'late'),'0',9,False),
          ('Material lead time (days)',lambda x:x['b'].get('lt_used'),'0',9,False),
          ('Stock',lambda x:round(x['own_stock']),'#,##0',11,False),
          ('to OEM Stock',lambda x:round(x['oem_stock']) if x.get('oem_stock') else None,'#,##0;-#,##0;"-"',11,False),
          ('In transit today (total qty, MIS)',lambda x:round(x.get('tr_tot') or 0),'#,##0;-#,##0;"-"',12,False),
          ('Arrival at plant (Final AtA)',lambda x:x.get('ata_txt'),None,42,True),
          ('Production Daily Target_without OEM',rnd('tgt_need_wo'),'#,##0.0',12,False),
          ('Inv. Days at line daily target_without OEM',rndz('tgt_days_wo'),'0.0',12,False),
          ('Inv. Days at line daily target_without OEM (stock only, no transit)',lambda x:(round(max(0.0,(POOLG[x['b']['mat']]['own'] if x['b']['mat'] in POOLG else x['own_stock']))/x['tgt_need_wo'],1) if x.get('tgt_need_wo') else None),'0.0',12,False),
          ('Production Daily Target_with OEM',rnd('tgt_need'),'#,##0.0',12,False),
          ('Inv. Days at line daily target_with OEM',rndz('tgt_days'),'0.0',12,False),
          ('Daily plan need (module plan)',lambda x:(round(x['pl_need'],1) if x.get('pl_need') else None),'#,##0.0',12,False),
          ('Inv. Days at module plan (stock + in transit)',lambda x:(round(x['pl_days'],1) if x.get('pl_days') is not None else None),'0.0',12,False),
          (f'Below {REPORT_DAYS} Days From (module plan)',lambda x:_cwv(x,'below','cw_p'),'dd-mmm-yy',12,False),
          ('Days Late (module plan)',lambda x:_cwv(x,'late','cw_p'),'0',9,False),
          ('In transit (total qty, WSA RM Planning)',lambda x:round(x['wsa_it']) if x.get('wsa_it') else None,'#,##0',12,False),
          ('Open PO qty (WSA RM Planning)',lambda x:round(x.get('op_tot') or 0),'#,##0;-#,##0;"-"',12,False),
          ('Alternative BOM (yours): code - stock - customers',alt_txt,None,46,True)]
    from openpyxl.styles import Side as _Sd
    AL_INK='1F2D3D'; al_hb=Border(bottom=_Sd(style='medium',color='C9D2DC')); al_rb=Border(bottom=_Sd(style='thin',color='E3E8EE'))
    for j,(h,_,_,w,_) in enumerate(COLS,1):            # header row 4 = RM Issues: light grey, dark bold text, medium line below
        c=al.cell(4,j,h); c.font=Font(name='Calibri',size=10,bold=True,color=AL_INK); c.fill=PatternFill('solid',fgColor='F4F6F8')
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=al_hb
        al.column_dimensions[L(j)].width=w
    al.row_dimensions[4].height=44
    jst_=[h_ for h_,*_ in COLS].index('Stock')+1; joe_=[h_ for h_,*_ in COLS].index('to OEM Stock')+1
    # (MRP Daily is built here: after the daily targets, so its order is the same single risk ranking as every other sheet)
    rows_=[]; chd={}; part=out_path+'.part.xlsx'
    def _daily(p_): make_daily_stock_xlsx(DS,p_,title='MRP Daily',table='MRPDaily',rowmap=rows_,cache=chd,target_days=TD)
    set_risk_rank(A,d0)                                                # highest shortage risk first - in MRP Daily
    for x_ in A:                                                       # one risk definition: status follows the shortage risk
        x_['status']={'ALERT':'CRITICAL','Short later':'AT RISK','Under 21 d':'WATCH','No demand':'NO DEMAND'}.get(risk_info(x_,d0)['cls'],'OK')
    DS['blocks']=sorted(DS['blocks'],key=lambda b_:rrank(b_['mat']))
    _daily(part); daily_part=out_path+'.daily.xlsx'; os.replace(part,daily_part)
    RM_={x['b']['mat']:x for x in rows_}                       # material -> its rows on MRP Daily
    # Settings cells of the line daily targets (the Settings sheet is written the same way at the end)
    TGT_CELL.clear()
    try:
        make_settings_sheet(part,SETTINGS_SRC[0]); ws_=load_workbook(part)['Settings']
        for r_ in range(1,ws_.max_row+1):
            v_=ws_.cell(r_,1).value
            if isinstance(v_,str) and v_.startswith('Look-ahead (days)'): TGT_CELL['_look']=f"Settings!$B${r_}"; continue
            if isinstance(v_,str) and v_.startswith('Daily plan = average'): TGT_CELL['_avg']=f"Settings!$B${r_}"; continue
            if isinstance(v_,str) and 'daily target' in v_.lower(): TGT_CELL[_line_key(re.split(r'production plan|daily target',v_,flags=re.I)[0])]=f"Settings!$B${r_}"
        os.remove(part)
    except Exception as ex_: print(f'NOTE: Settings target cells not found ({type(ex_).__name__}) - targets written as numbers')
    AL_ORDER=sorted(A,key=lambda x_:rrank(x_['b']['mat']))   # ONE ranking for every sheet: highest shortage risk first
    # GLASS PER SUPPLIER (PPC Usage way): each customer's need goes to its first allowed vendor (Settings ALLOWED VENDORS);
    # stock = that vendor's batches, in transit = that vendor's shipments. Rows follow the material's own row.
    allow_={}
    for r_ in SET_CUSTSUP or []:
        if r_.get('Code') and r_.get('Supplier'): allow_.setdefault((_norm(r_['Customer']),str(r_['Code']).strip()),[]).append(str(r_['Supplier']).strip().upper())
    def vend_of(cu_,mat_):
        L_=next((v for k,v in LINKS.items() if _norm(k)==_norm(cu_)),None) or {}
        for k_ in (_norm(cu_),_norm(L_.get('mis'))):
            if allow_.get((k_,mat_)): return allow_[(k_,mat_)][0]
        return None
    def sup_key(name_,keys_):
        n_=str(name_ or '').upper()
        return next((k_ for k_ in keys_ if k_ and k_.upper() in n_),None)
    out_=[]
    for x in AL_ORDER:
        out_.append(x); b=x['b']
        if 'glass' not in str(b.get('common') or '').lower() or not b.get('sup'): continue
        need_={}; need0_={}; cus_={}
        for cu_,ln_,tg_,oem_ in CUST_TARGET:
            u_=bom_usage(b['mat'],False,cust_row(cu_))
            if not u_: continue
            v_=vend_of(cu_,b['mat']) or 'ANY VENDOR'
            need_[v_]=need_.get(v_,0.0)+u_*tg_; cus_.setdefault(v_,[]).append(cu_)
            if oem_!='Yes': need0_[v_]=need0_.get(v_,0.0)+u_*tg_
        supd_={str(e_.get('name') or ''):e_ for e_ in b['sup']}
        sups_=[k_ for k_ in supd_ if k_ and k_.upper()!='UNASSIGNED']
        act_=[v_ for v_ in set(sups_)|{k_ for k_ in need_ if k_!='ANY VENDOR'}
              if float((supd_.get(sup_key(v_,sups_) or v_) or {}).get('stock') or 0)>0 or sum((supd_.get(sup_key(v_,sups_) or v_) or {}).get('transit') or [])>0 or need_.get(v_,0)>0]
        if len(act_)<2: continue                                        # one supplier only: its row would repeat the material row
        for v_ in sorted(set(sups_)|{k_ for k_ in need_ if k_!='ANY VENDOR'}):
            k_=sup_key(v_,sups_) or v_
            sp_=supd_.get(k_) or {}
            st_=float(sp_.get('stock') or 0.0); it_=float(sum(sp_.get('transit') or []))
            # counted material: b['sup'] already holds the count by supplier rolled forward (apply_count_by_supplier)
            nw_=need_.get(v_,0.0); nwo_=need0_.get(v_,0.0)
            px=dict(x); px['b']=dict(b,common=f"{b['common']} - {v_}",desc=f"{b.get('desc') or ''} | supplier {v_}"); px['sup_row']=v_
            vpx_=VP_MRP.get(b['mat'])
            if vpx_:                                                     # vendor rule: the supplier row is not the pooled (locked) projection (checker finding)
                Sv_=x.get('S_pool') or x.get('S') or []
                if vpx_.get('sheets'):                                   # 2+ WSA sheets: the sheet holding this supplier's stock
                    hit_=[sv[1] for sv in vpx_['sheets'].values() if abs(sv[0]-st_)<0.5]
                    if len(hit_)==1: Sv_=hit_[0]
                wv_=[]; k_=0
                while k_<len(Sv_):
                    if Sv_[k_]<0:
                        k2_=k_
                        while k2_+1<len(Sv_) and Sv_[k2_+1]<0: k2_+=1
                        wv_.append((k_,k2_,k2_+1>=len(Sv_))); k_=k2_+1
                    else: k_+=1
                px['S']=list(Sv_); px['wins']=wv_
            px.update(cw=None,own_stock=st_,oem_stock=0.0,tr_tot=it_,tgt_need=nw_,tgt_need_wo=nwo_,
                      tgt_days=(max(0.0,st_+it_)/nw_ if nw_ else None),tgt_days_wo=(max(0.0,st_+it_)/nwo_ if nwo_ else None),
                      custs=(', '.join(cus_.get(v_,[])) or 'no customer allowed this vendor'),proj=None,mods=None)
            out_.append(px)
    AL_ORDER=out_
    AL_ORDER_[:]=AL_ORDER; COLS_[:]=COLS; AL_VAL.clear()          # for the status sheets (same order as the Action List)
    for i,x in enumerate(AL_ORDER):
        rr=5+i
        for j,(h,fn,nf,w,wrap) in enumerate(COLS,1):
            v_=fn(x)
            if not x.get('sup_row'): AL_VAL[(x['b']['mat'],h)]=v_
            c=al.cell(rr,j,v_); c.font=Font(name='Calibri',size=10,color=AL_INK); c.border=al_rb; c.alignment=Alignment(vertical='center',wrap_text=wrap)
            if nf: c.number_format=nf
            if h=='Priority':
                bg_,fg_={'High':('FCE8E6','C0392B'),'Medium':('FEF3E2','B9770E'),'Low':('FEF9E7','9A7D0A'),'OK':('E3F4EC','1E8449')}.get(v_,('F2F4F6','5D6D7E'))
                c.fill=PatternFill('solid',fgColor=bg_); c.font=Font(name='Calibri',size=10,bold=True,color=fg_); c.alignment=Alignment(horizontal='center',vertical='center')
            if h.startswith('Short from') and c.value not in ('no stock-out','No demand'): c.font=Font(name='Calibri',size=10,bold=True,color='9C0006')
            if h in ('Short From (Stock = 0)',) and isinstance(c.value,str) and '→' in c.value: c.font=Font(name='Calibri',size=10,bold=True,color='9C0006')     # stock zero inside the lead time
            elif h in ('Short From (Stock = 0)',) and isinstance(c.value,str) and 'after lead time' in c.value: c.font=Font(name='Calibri',size=10,bold=True,color='B9770E')
            if h.startswith('Days Late') and isinstance(c.value,(int,float)) and c.value>0: c.font=Font(name='Calibri',size=10,bold=True,color='B9770E')
            if x.get('no_demand') and h in ('Component description','Short from → to (with in transit)'): c.font=Font(name='Calibri',size=10,italic=True,color='7F7F7F')
            if h.startswith('Inv. Days') and isinstance(c.value,(int,float)) and c.value<REPORT_DAYS:     # under 21 days: red highlight
                c.fill=PatternFill('solid',fgColor='FFC7CE'); c.font=Font(name='Calibri',size=10,bold=True,color='9C0006')
            if h.startswith('Short from') and str(c.value or '').startswith('ALERT'): c.fill=PatternFill('solid',fgColor='FFC7CE')
    TG_,PL_=calc_layout(A,RM_); CALC_LAYOUT[:]=[TG_,PL_]
    _hs=[h_ for h_,*_ in COLS]; _j0=_hs.index('Stock'); _j1=max(i_ for i_,h_ in enumerate(_hs) if ('OEM' in h_ and ('arget' in h_ or 'lan' in h_)) or 'module plan' in h_)
    AL_COPY[:]=[(h_,L(i_+1)) for i_,h_ in enumerate(_hs) if _j0<=i_<=_j1]          # Action List F .. last daily-plan column -> MRP by Line
    hidx={h_:i_+1 for i_,(h_,*_) in enumerate(COLS)}; cl_=lambda h_: L(hidx[h_]) if h_ in hidx else 'A'
    F_,G_,H_,J_,K_,Lc,M_,N_,O_,P_,Q_=(cl_(h_) for h_ in ('Stock','to OEM Stock','In transit today (total qty, MIS)','Production Daily Target_with OEM',
        'Inv. Days at line daily target_with OEM','Production Daily Plan_with OEM','Inv. Days at line daily plan_with OEM','Production Daily Target_without OEM',
        'Inv. Days at line daily target_without OEM','Production Daily Plan_without OEM','Inv. Days at line daily plan_without OEM'))
    d0x=f'DATE({d0.year},{d0.month},{d0.day})'
    for rr,x in (zip(range(5,5+len(AL_ORDER)),AL_ORDER) if not VALUES_ONLY else []):   # every number from stock onward is a formula (change stock on Stock Data -> all follow)
        m=x['b']['mat']; rm=RM_.get(m)
        tgr=f"SUMIFS('Calc Target'!$F$2:$F$3000,'Calc Target'!$A$2:$A$3000,$A{rr})"; tgr0=f"SUMIFS('Calc Target'!$F$2:$F$3000,'Calc Target'!$A$2:$A$3000,$A{rr},'Calc Target'!$C$2:$C$3000,\"Lines 1-4\")"
        al[f'{J_}{rr}']=f'=IF({tgr}=0,"",{tgr})'; al[f'{N_}{rr}']=f'=IF({tgr0}=0,"",{tgr0})'
        al[f'{K_}{rr}']=f'=IF(N({J_}{rr})>0,MAX(0,{F_}{rr}+{H_}{rr})/{J_}{rr},"")'
        al[f'{O_}{rr}']=f'=IF(N({N_}{rr})>0,MAX(0,{F_}{rr}+{H_}{rr})/{N_}{rr},"")'
        al[f"{cl_('Inv. Days at line daily target_without OEM (stock only, no transit)')}{rr}"]=f'=IF(N({N_}{rr})>0,MAX(0,{F_}{rr})/{N_}{rr},"")'
        al[f"{cl_('Inv. Days at line daily plan_without OEM (stock only, no transit)')}{rr}"]=f'=IF(N({P_}{rr})>0,MAX(0,{F_}{rr})/{P_}{rr},"")'
        if rm:
            c0,ndx=rm['C0'],rm['ND']; a_,z_=L(c0),L(c0+ndx-1)
            al[f'{H_}{rr}']=f"=SUM('MRP Daily'!${a_}${rm['rt']}:${z_}${rm['rt']})"
            al[f"{cl_('Open PO qty (WSA RM Planning)')}{rr}"]=(f"=SUM('MRP Daily'!${a_}${rm['ro']}:${z_}${rm['ro']})" if rm.get('ro') else 0)
            bal=f"'MRP Daily'!${a_}${rm['rs']}:${z_}${rm['rs']}"; f0=f'MATCH(TRUE,INDEX({bal}<0,0),0)'
            g0=f'MATCH(TRUE,INDEX(INDEX({bal},{f0}):INDEX({bal},{ndx})>=0,0),0)'                       # (INDEX:INDEX ranges, not volatile OFFSET)
            more=f'IF({f0}+{g0}-1>{ndx},"",IF(COUNTIF(INDEX({bal},MIN({ndx},{f0}+{g0}-1)):INDEX({bal},{ndx}),"<0")>0,"; more later",""))'
            al[f"{cl_('Short from → to (with in transit)')}{rr}"]=(f'=IFERROR(TEXT({d0x}+{f0}-1,"dd-mmm")&" → "&IFERROR(TEXT({d0x}+{f0}+{g0}-3,"dd-mmm")&" ("&({g0}-1)&" d)"&{more},'
                                                                     f'"beyond "&TEXT({d0x}+{ndx}-1,"dd-mmm")),"no stock-out")')
            pr=PL_.get(m)
            if pr:
                ce=L(3+ndx-1)
                navg=TGT_CELL.get('_avg') or str(int(PLAN_AVG_DAYS))
                for need_c,days_c,prow,stk in ((Lc,M_,pr,f'{F_}{rr}+{G_}{rr}+{H_}{rr}'),(P_,Q_,pr+1,f'{F_}{rr}+{H_}{rr}')):
                    al[f'{need_c}{rr}']=f"=IFERROR(AVERAGEIF('Calc Plan'!$C${prow}:INDEX('Calc Plan'!$C${prow}:${ce}${prow},{navg}),\">0\"),\"\")"   # average of the next N days with production (Settings) - no OFFSET
                    al[f'{days_c}{rr}']=f'=IF(N({need_c}{rr})>0,MAX(0,{stk})/{need_c}{rr},"")'
    jat_=[h_ for h_,*_ in COLS].index('Arrival at plant (Final AtA)')+1; jcu_=[h_ for h_,*_ in COLS].index('Lines (share of plan volume)')+1
    for rr,x in zip(range(5,5+len(AL_ORDER)),AL_ORDER):   # row height fits the arrival lines
        n_=max(str(x.get('ata_txt') or '').count('\n'),str(x.get('custs') or '').count('\n'),str(x.get('mods') or '').count('\n'),
               str(AL_VAL.get((x['b']['mat'],'Alternative BOM (yours): code - stock - customers')) or '').count('\n'),
               str(x.get('proj') or '').count('\n'),str(x.get('usage_src') or '').count('\n'))+1
        for j__ in (jat_,jcu_,jcu_+1): al.cell(rr,j__).alignment=Alignment(wrap_text=True,vertical='center')
        al.row_dimensions[rr].height=max(30,13.5*n_)                    # RM Issues row height (30), more for multi-line cells
    for rr in (range(5,5+len(AL_ORDER)) if not VALUES_ONLY else []):     # stock = the lines on Stock Data (double-check there)
        al.cell(rr,jst_,f"=SUMIFS('Stock Data'!$L$5:$L$3000,'Stock Data'!$A$5:$A$3000,$A{rr},'Stock Data'!$C$5:$C$3000,\"Stock\")")
        al.cell(rr,joe_,f"=SUMIFS('Stock Data'!$L$5:$L$3000,'Stock Data'!$A$5:$A$3000,$A{rr},'Stock Data'!$C$5:$C$3000,\"to OEM\")")
    al.add_table(Table(displayName='ActionList',ref=f'A4:{L(len(COLS))}{4+len(AL_ORDER)}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=False)))   # = RM Issues
    al.freeze_panes='C5'
    # ---------- TODAY + PLANNER sheets ----------
    pass                                                           # Today sheet removed (your request)
    pass                                                           # cover chart was on the Today sheet (removed)
    pass                                                           # Planner sheet removed (your request)
    tmp=out_path+'.build.xlsx'; wb.save(tmp)
    # ---------- grafted sheets (same builders as the old MRP) ----------
    part=out_path+'.part.xlsx'
    def add(title,builder,cache=None):
        builder(part); graft_sheet(tmp,part,title,cache=cache); os.remove(part)
    # (MRP Daily and its row map were built before the Action List)
    for x in rows_:
        b=x['b']; x['cov']=cover_days(b['stock'],b['plan'],[a+c for a,c in zip(b['transit'],b['adhoc'])],b.get('openpo') or [0.0]*nd)
        cp=b.get('cust_plan') or {}; x['cust']={c:sum(v) for c,v in cp.items()}
    # ---------- layout: 8 visible sheets (red = daily, blue = weekly, grey = data); details hidden ----------
    RED,BLUE,GREY='C00000','2F5597','7F7F7F'
    sec=lambda n: out_path+f'.{n}.xlsx'
    graft_sheet(tmp,daily_part,'MRP Daily',cache=chd,tab_color=RED); os.remove(daily_part)
    ALLOC_STOP_.clear(); ALLOC_STOP_.update(alloc_stop)
    make_line_view_xlsx(A,LR,d0,nd,part,TD,AL_COPY); graft_sheet(tmp,part,'MRP by Line',tab_color=RED); os.remove(part)
    # (RM Status / PM Status sheets removed - your request)

    # MRP by Line: per production line, the materials it needs and which are at risk (compact, no duplicate grid)

    chw={}; make_daily_stock_xlsx(DS,part,title='MRP Weekly',table='MRPWeekly',weekly_of='MRP Daily',cache=chw,target_days=TD)
    graft_sheet(tmp,part,'MRP Weekly',cache=chw,tab_color=BLUE); os.remove(part)
    make_intransit_status_xlsx(mis_path,DS,part); graft_sheet(tmp,part,'In Transit Status',tab_color=BLUE); os.remove(part)   # MIS STATUS analysed
    load_user_masters(SETTINGS_SRC[0])                               # your Batch Master + Material Group (MRP.xlsx)
    make_substitute_advice_xlsx(A,DS,d0,part,mis_path); graft_sheet(tmp,part,'Substitute Advice',tab_color=BLUE); os.remove(part)       # replacements from descriptions
    # In transit MIS vs WSA + Supplier Coverage: collapsed sections at the bottom of Check (data reconciliation)
    IT=build_mis_intransit(mis_path,DS_mis,BW); chi={}; make_intransit_xlsx(IT,sec('it'),cache=chi,compact=True)
    make_supplier_cover_sheet(DS_rm,sec('sc'),TD)
    # Stock Movement (replaces Stock History) + its hidden log
    mlog=[e for e in read_movement_log(prev_path) if e['date']!=d0 and e['date']>=d0-pd.Timedelta(days=60)]+movement_entries(DS)
    make_stock_data_xlsx(DS,A,part); graft_sheet(tmp,part,'Stock Data',tab_color=BLUE); os.remove(part)   # every stock line (formulas point here)
    if not VALUES_ONLY:
        make_calc_target_xlsx(CALC_LAYOUT[0],part); graft_sheet(tmp,part,'Calc Target',tab_color=GREY); os.remove(part)
        make_calc_plan_xlsx(A,RM_,CALC_LAYOUT[1],d0,part); graft_sheet(tmp,part,'Calc Plan',tab_color=GREY); os.remove(part)
    make_movement_log_xlsx(mlog,part); graft_sheet(tmp,part,'Movement Log',hidden=True); os.remove(part)   # data kept for the next run
    # Check = issues + Unit Check + Pack Size Check + Plan Compare (sections on one sheet)
    unit_issues=[]; unit_issues.extend(make_unit_check_xlsx(DS,bom,cnt_hdr,cnt_rows,modplan,mis_path,sec('uc')))
    prow=pack_size_scan(DS_mis,mis_path,old_mrp); make_pack_check_xlsx(prow,sec('pk'))
    for r_ in prow:                        # stock of the same material under another code that the MRP does not count
        if 'NOT counted' in r_['how'] and r_['stock']>0 and r_['target']:
            unit_issues.append(dict(area='Stock',check='Same material under another code / pack - not counted',sev='Medium',material=r_['desc'],code=r_['code'],
                a_lab='',a=None,b_lab='Stock',b=round(r_['stock'],1),diff=None,detail=f"{r_['how']} for {r_['target']}",action=r_['act']))
    if pp_rows: make_plan_compare_xlsx(pp_rows,mis_modules_by(mis_path,DS['plant'],dates),dates,sec('pc'))
    try: ml_iss,ml_rows=mis_layout_check(mis_path,prev_path); make_mis_layout_xlsx(ml_rows,sec('ml'))
    except Exception as ex_:
        print(f'WARNING: MIS layout check skipped ({type(ex_).__name__}: {ex_})'); ml_iss,ml_rows=[],[]
    make_segment_xlsx(DS,sec('sg'))
    try:
        iss=build_checks(DS_mis,BW,mis_path,None,LIMIT_DAYS=30)+cspm_checks(bom,cnt_rows,modplan,mis_names,d0)+bom_rate_checks(mis_path,bom,BW,DS_mis,TD)+map_issues+pp_issues+unit_issues+mis_usage_check(mis_path,DS_mis)+ml_iss+list(ADVICE_ISSUES)+ubom_checks(SETTINGS_SRC[0],mis_path)
        iss+=[dict(area='In transit',check='MIS Final RM Code does not match the shipment - traced code used',sev='High' if (f['qty'] or 0)>0 else 'Low',
                   material=f['desc'][:60],code=f"{f['shipped']} (MIS: {f['mis']})",a_lab='Pending qty',a=round(f['qty'] or 0),b_lab='',b=None,diff=None,
                   detail=f"Shipment '{f['desc'][:60]}' matches {f['shipped']} ({f['cd_s'][:40]}) not {f['mis']} ({f['cd_m'][:40]}); PO {f['po']}, {str(f['sup'] or '')[:25]}; how: {f.get('how','')[:90]}",
                   action='Ask MIS to correct the Final RM Code (or the PO material)') for f in IT_FIX]
        ucodes={b['mat'] for b in unlimited}
        iss=[x for x in iss if x.get('code') not in ucodes]+[dict(area='Master data',check='Treated as unlimited stock',sev='Low',material=b['desc'],code=b['mat'],a_lab='',a=None,
                 b_lab='Plan next 30 d',b=round(sum(b['plan'][:30])),diff=None,detail='Not planned for shortages (unlimited stock)',action='Remove from UNLIMITED_CODES at the top of the script if this changes') for b in unlimited]
        cnt=make_check_xlsx(iss,DS,sec('ck'))
        combine_sheets(sec('ck'),[sec('ml'),sec('sg'),sec('uc'),sec('pk'),sec('pc'),sec('it'),sec('sc')],part,collapse=True,caches={sec('it'):chi})
        graft_sheet(tmp,part,'Check',tab_color=GREY); os.remove(part)
    except Exception as ex_:
        print(f'WARNING: Check sheet skipped ({type(ex_).__name__}: {ex_})'); cnt={'High':0,'Medium':0,'Low':0}
    for n_ in ('it','sc','uc','pk','pc','ck','sg','ml'):
        if os.path.exists(sec(n_)): os.remove(sec(n_))
    make_no_demand_xlsx(DS,mis_path,bom,modplan,BW,part); graft_sheet(tmp,part,'No Demand',tab_color=GREY); os.remove(part)
    def _how(p_):
        w=Workbook(); h=w.active; h.title='How it works'
        lines=[('How this MRP works (current logic - details per header on the Definitions sheet)',True),('',False),
            ('Inputs (put them in the folder each day, then run the script)',True),
            ('- MIS RM Dashboard: SAP stock (MB52), JW stock at OEM, in transit (Final Pending Qty, Final AtA, PPC GRN), Standard BOM.',False),
            ('- WSA planning file: Open PO lines without vessel, supplier / PO / ETD history (lead times).',False),
            ('- Your MB52.xlsx (SAP export, same folder): the stock quantities (Unrestricted + Quality Inspection); MIS MB52 only adds plant name, final material code, dashboard key, company and categories.',False),
            ('- Your MRP.xlsx (read only): Settings, Material Group, Batch Master. Module plan (Plan + Revised Plan), BOM and Stock Count come from the PPC workbook (sheets Plan, Revised Plan, BOM, Count) when it is in the folder; MRP.xlsx Production Plan / BOM only without a PPC file.',False),
            ('- RM Issues: one issue per raw material at risk (required date, qty, supplier / PO / ETA, delay, MW impact); your Owner / Status / Date Closed / Action are kept run to run.',False),
            ('- Customer_Tracker_*.xlsx: no longer used (SCM Dashboard removed).',False),
            ('',False),('Stock',True),
            ('- Counted materials: the physical count is used; a count older than today is rolled forward: count - usage per day x days since count + received since count (MIS PPC GRN). Usage per day = Production Plan demand, next 7 days.',False),
            ('- Not counted: SAP MB52 Total Stock (+ alternate codes / roll lengths converted); when the MIS file is older than today (its save date) this SAP stock is rolled forward the same way (− usage per day × days since the MIS date) and the console prints a WARNING. JW stock at OEM is added in the with-OEM figures.',False),
            ('- OEM (customers with OEM = Yes): their plan and the JW stock at the OEM are left out of every calculation; the JW stock stays visible as data (Stock Data, Movement Log, No Demand).',False),
            ('- In transit: each shipment is matched to its material by Description + Supplier + RM Code + Final RM Code + PO-line codes (busbar / ribbon width included); changes are listed on Check.',False),
            ('- Days of cover (MRP Daily / Weekly) = projected stock ÷ average plan usage of the next N days (Settings, 21). Need/day = target-based; Usage/day = plan-based - see Definitions.',False),
            ('- Data dates (MIS, WSA, counts) are shown on the dashboards, red when older than 2 days. Stock Check sheet shows SAP vs count vs estimate per material.',False),
            ('',False),('Demand',True),
            ('- Plan demand per day = Production Plan modules x usage per module: your BOM first (an alternative listed next to its standard code for the same customer carries no demand; listed alone it carries the demand), else MIS Standard BOM.',False),
            ('- Substitute groups (Settings SUBSTITUTES, lines blank, valid today): the material and its substitutes are planned as ONE - group stock + group in transit vs group demand (each code x its factor) for the daily target, Inv. days and the shortage date of every member.',False),
            ('- Daily target need = customer daily targets (Settings CUSTOMER DAILY TARGET, else line targets) x usage; without OEM = customers with OEM = No.',False),
            ('',False),('Shortage risk - ONE rule on every sheet',True),
            ('- Day-by-day projection: stock - plan demand + MIS in transit by arrival date. Short from = first day below zero.',False),
            ('- ALERT = short inside the material lead time · Short later = short after it · Under 21 d = no shortage but under 21 days at the targets · OK.',False),
            ('- Priority: EXPEDITE (short inside lead time) · ORDER NOW (last order date within the release window) · ORDER BY dd-mmm · Planned only when NOT short (below reorder point).',False),
            ('- NO DEMAND: a material is only Critical / At risk / Watch / Short when it has demand (production requirement, planned consumption, daily target or BOM demand) inside the shortage horizon (Settings, 60 days). Without it: status NO DEMAND, no planned order / release, no Expedite / Order now, not on the ALERT.',False),
            ('- Every list is sorted by this risk (highest first); a short material is never labelled safe; the script checks itself after every run (console: Self-check).',False),
            ('',False),('Lines',True),
            ('- Each material is shared out to the lines day by day (line ranking, allowed vendors, vendor priority, Settings substitutes) and then your approved ALTERNATIVE (your BOM) is used before a line stops; free stock of an alternative = what its own plan does not need within its lead time.',False),
            ('',False),('Sheets (visible)',True),
            ('ALERT (first tab) = ALERT + Material Status in one sheet: the ALERT rows (RM and packaging) with estimated stock today, SAP stock (MB52), physical count, need per day, stock-zero days, below 21 days, in transit and overdue. Materials with NO DEMAND are never listed as a risk.',False),
            ('ALERT (first tab): every material under 21 days, falling below 21 days in the next 16 weeks or out of stock, and why (Below 21 days from, stock zero from → to, Why, Days late, Arrival at plant (Final AtA), action); BOM Material Dashboard - Weekly (was Cover Heatmap) and BOM Material Dashboard - Daily - the same materials, weekly projected cover around the 21-day line, highest risk on top.',False),
            ('Action List - every material by shortage risk with daily-target and module-plan cover; Stock Check (SAP vs count) · MRP Daily · RM Issues · MRP by Line · MRP Weekly. ALERT (formulas), ALERT Calc, ALERT Plan lines, ALERT Arrival lines = the formulas behind ALERT.',False),
            ('Detail sheets (right of In Transit Status, all visible): In Transit Status, Substitute Advice, Stock Data, Check, No Demand, Settings, Movement Log, Module Plan, BOM, Stock Count, How it works, MIS Layout, Definitions.',False),
            ('',False),('Daily routine',True),
            ('1 Run the script and read the console Self-check line. 2 ALERT: stock-out in lead time, overdue in transit. 3 ALERT: Below 21 days from, Why, Days late. 4 Stock Check: counts older than today and big SAP differences.',False)]
        for i,(t,bo) in enumerate(lines,1):
            c=h.cell(i,1,t); c.font=Font(name=AR,size=12 if i==1 else 10,bold=bo,color=NAVY if bo else '000000')
        h.column_dimensions['A'].width=150; w.save(p_)
    # hidden details / inputs (one click to unhide); Customer Supplier stays visible when it is edited here
    def addh(title,builder,cache=None,hidden=True):
        builder(part); graft_sheet(tmp,part,title,cache=cache,hidden=hidden,tab_color=None if hidden else GREY); os.remove(part)
    addh('Module Plan',lambda p_: module_plan_sheet(modplan,dates,p_,pp_rows=pp_rows,src=os.path.basename(old_mrp) if old_mrp else None))
    BOM_SRC_NAME[0]=os.path.basename(bom_src) if bom_src else None
    addh('BOM',lambda p_: make_bom_sheet(bom,mis_names,p_))
    chc={}; addh('Stock Count',lambda p_: make_count_sheet(cnt_hdr,cnt_rows,p_,bom=bom,cache=chc),chc)
    addh('How it works',_how)
    addh('MIS Layout',lambda p_: make_mis_layout_xlsx(ml_rows,p_,store=True))      # headers read by position, compared on the next run
    make_settings_sheet(part,SETTINGS_SRC[0]); graft_sheet(tmp,part,'Settings',tab_color='7F7F7F'); os.remove(part)
    subj,body=material_report(A,d0,os.path.basename(out_path))          # ONE daily report (the 3 old e-mails are gone)
    try:
        for old_ in ('_email.txt','_email_RM.txt','_email_CSPM.txt'):          # remove today's old-style e-mail files if a previous version left them
            if os.path.exists(out_path.replace('.xlsx',old_)): os.remove(out_path.replace('.xlsx',old_))
    except OSError: pass
    # Start Here sheet removed (your request)
    reorder_sheets(tmp,['Action List','MRP Daily','MRP by Line','MRP Weekly','In Transit Status','Substitute Advice','Stock Movement','Stock Data','Check','No Demand','Settings'])
    unfreeze_all(tmp)                                              # no frozen panes on any sheet of MRP_Daily
    # all sheets visible (your choice); the main sheets stay first in the tab order                                # only the Action List is visible; every other sheet is hidden (formulas still use them)
    os.replace(tmp,out_path)
    SHIPS_.clear(); SHIPS_.update(ships)
    try: ALERT_ROWS[:]=build_alert_rows(A,d0,nd); ALERT_D0[0]=d0         # ALERT list (incl. the old Risk Charts rows), built once
    except Exception as ex_: import traceback; traceback.print_exc(); ALERT_ROWS[:]=[]; print(f'NOTE: ALERT list not built ({type(ex_).__name__}: {ex_})')
    try: add_cover_heatmap(out_path,ALERT_ROWS,d0,nd,count_items(cnt_hdr,cnt_rows))                    # weekly cover of the ALERT rows, high to low risk (cells, no chart objects)
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: BOM Material Dashboard not added ({type(ex_).__name__}: {ex_})')
    try: add_rm_issues_sheet(out_path,A,AL_ORDER,d0,nd,ships,prev_path)      # RM Issues log (your entries carried over)
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: RM Issues not added ({type(ex_).__name__}: {ex_})')
    try: add_stock_check_sheet(out_path,AL_ORDER,d0,mlog)               # SAP stock vs physical count vs estimate today (after Inventory)
    except Exception as ex_: print(f'NOTE: Stock Check sheet not added ({type(ex_).__name__}: {ex_})')
    try: add_stock_mgmt_dashboard(out_path,A,AL_ORDER,mlog,ships,d0,nd,MIS_PATH_[0])     # Material Status (every material; brought back Nov 2026)
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: Material Status dashboard not added ({type(ex_).__name__}: {ex_})')
    try:                                                         # ALERT = first tab: under 21 days / stock zero inside the lead time (same rows in the Inbound file)
        write_alert_sheet(post_wb(out_path),ALERT_ROWS,d0,'this is the MRP_Daily copy')
        write_alert_formulas(post_wb(out_path),ALERT_ROWS,d0)       # the same sheet with Excel formulas (check the work)
        wb__=post_wb(out_path)
        if 'Material Status' in wb__.sheetnames: wb__.move_sheet('Material Status',offset=1-wb__.sheetnames.index('Material Status'))   # right after the ALERT
        _c=Counter(r_['cls'] for r_ in ALERT_ROWS)
        print(f"ALERT sheet (first tab of MRP_Daily and of the Inbound file; Risk Charts merged in): {_c.get(0,0)} out of stock inside the transit lead time, {_c.get(1,0)} out of stock later, "
              f"{_c.get(2,0)} under {REPORT_DAYS} days of stock (in transit covers), {_c.get(3,0)} under {REPORT_DAYS} days later (within 16 weeks), {_c.get(4,0)} OK, {_c.get(5,0)} no demand - all {len(ALERT_ROWS)} materials listed",flush=True)
    except Exception as ex_: import traceback; traceback.print_exc(); ALERT_ROWS[:]=[]; print(f'NOTE: ALERT sheet not added ({type(ex_).__name__}: {ex_})')
    try:                                                         # BOM Material Dashboard - Daily, right after the weekly one (built last: no other tab moves)
        _heat_ms_fill(post_wb(out_path))                         # weekly dashboard E..J = Material Status columns (built after Material Status)
        if HEAT_ROWS[0]: _bom_daily_sheet(post_wb(out_path),*HEAT_ROWS[0])
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: {DAILY_NAME} not added ({type(ex_).__name__}: {ex_})')
    try: unify_workbook_style(out_path)                          # one style for all sheets, no charts, Definitions sheet
    except Exception as ex_: print(f'NOTE: style / definitions not applied ({type(ex_).__name__}: {ex_})')
    try:
        wb_=post_wb(out_path)
        # Action List stays in the file (user, Nov 2026: 'bring action list back')
        for w_ in wb_.worksheets: w_.sheet_view.tabSelected=False
        wb_.active=0; wb_.worksheets[0].sheet_view.tabSelected=True
    except Exception as ex_: print(f'NOTE: active sheet not set ({type(ex_).__name__}: {ex_})')
    try: make_daily_material_status(out_path,post_wb(out_path),AL_ORDER,bom,cnt_hdr,cnt_rows,d0,MIS_PATH_[0],ships)     # Daily_Material_Status_<date>.xlsx: Materials Status (BOM + other materials) + ALERT
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: Daily Material Status workbook not written ({type(ex_).__name__}: {ex_})')
    try:                                                           # after the Daily_Material_Status copy (it reads the weekly sheet with header row 4)
        wb_=post_wb(out_path)
        if HEAT_NAME in wb_.sheetnames: _heat_final(wb_[HEAT_NAME])  # rows 1-2 out, 9 pt, Material Status columns grouped (user, Nov 2026)
    except Exception as ex_: print(f'NOTE: {HEAT_NAME} layout not finished ({type(ex_).__name__}: {ex_})')
    try: write_etd_late_sheet(post_wb(out_path),ETD_LATE,'This is the MRP_Daily copy (the same rows as the Inbound file).')   # ETD passed, not in transit (user, Nov 2026)
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: ETD late sheet not added to MRP_Daily ({type(ex_).__name__}: {ex_})')
    try: post_flush(out_path)                                      # ONE save for all finishing steps
    except Exception as ex_: print(f'NOTE: finishing save failed ({type(ex_).__name__}: {ex_})')
    audit_workbook(out_path)                                       # self-check: risk order + no short material labelled safe
    run_summary('MRP_Daily + Inbound',extra=[f"Materials with no demand (no status / action): {', '.join(c_+' '+n_[:20] for c_,n_,w_ in NODEM_)}"] if NODEM_ else [])
    return dict(n=n,crit=sum(1 for x in A if x['status']=='CRITICAL'),risk=sum(1 for x in A if x['status']=='AT RISK'),
                watch=sum(1 for x in A if x['status']=='WATCH'),ok=sum(1 for x in A if x['status']=='OK'),nodem=sum(1 for x in A if x['status']=='NO DEMAND'),check=cnt,cs=len(cs),bom_rows=len(bom),bom_src=bom_src)

# ===================== WSA alignment: rebuild the planning sheets of the WSA file =====================
# Runs on every run of the script (python update_inbound.py); --no-align skips it.
# Output = your WSA file, updated in place (no backup folder; older '(aligned <date>)' files deleted); then the Inbound and
# MRP_Daily files are built from it. Every planning sheet gets ONE standard layout - the same header row and the same columns:
#   rows 1-6 material info (title, stock, supplier & material, unit, codes, totals), row 7 headers, one row per day from the
#   1st of the current month (older rows are deleted - their delivery records go to the Archive sheet) to the last plan date.
#   Columns: A Date, B Day, C Balance, D Planned modules, E Demand, F.. customer columns (the same customers in the same order
#   on every sheet), Actual modules, Actual, MW, Actual MW, Plant Days, Plant + Port Days, Delivered, Scheduled Delivery, PO,
#   Shipper, Remarks, Booking Approval Date, Readiness, ETD, BL#, Vessel ETA, Status, Invoice, Reference, Remarks 2.
# From today the planning columns are LIVE formulas on the PPC_Plan_Req sheet (module plan x BOM); Balance = stock typed on
# today's row, then previous Balance - Actual + Delivered. Rows before today keep the numbers they had (values).
# Delivered and the columns after it are copied as they are (the planner's own formulas such as =10000*4 or =A747-40 are kept).
# A sheet with two materials side by side (Potting Material, Tapes) becomes two sheets. The Dashboard's lookup column index is
# set to the new Plant + Port Days column. Edited at XML level: every other sheet, pivot, query and format stays as it is.
ALIGN_MODE=[False]
ALIGN=[None]
WSA_OEM_IN_BALANCE=False     # False = like MRP_Daily: OEM plan + JW stock at the OEM left out of Balance / Demand (OEM columns still show the modules)
ALIGN_SHEET='PPC_Plan_Req'   # the plan table the rebuilt columns read (added to the aligned WSA file)
ARCHIVE_SHEET='Archive'      # delivery records of the deleted rows (before the 1st of the current month)
WSA_HDR=7                    # header row of every planning sheet
WSA_SLOTS=8                  # customer columns on every planning sheet (at least)
WSA_SPLIT_NAMES={'Potting Material':'Potting Material B','Tapes':'Tapes Perforated'}   # 2nd material of a two-material sheet
LABELS.update({('Potting Material B',0):'Potting Material - Part B',('Tapes Perforated',0):'Tapes - Perforated 500m'})
ALIGN_COLS=['Date','Production line','Plan customer','MIS customer','WSA planning sheet','WSA column','Counted in Balance',
            'Modules','Qty per module','Requirement','MW','Unit','Material code(s)','Qty source']
ARCH_COLS=['Planning sheet','Required Plant Date','Delivered','Scheduled Delivery','PO','Shipper','Remarks','Booking Approval Date',
           'Readiness','ETD','BL#','Vessel ETA','Invoice','Reference','State']
WSA_FIX1=[('date','Required Plant Date',12),('day','Day',6),('bal','Balance',12),('plan','Planned modules per day',11),('dem','Demand',12)]
WSA_FIX2=[('actmod','Actual modules per day',10),('act','Actual',11),('mw','MW per day',9),('actmw','Actual MW per day',9),
          ('pdays','Plant Days',9),('pport','Plant + Port Days ETA +5 Days',11),('ppcit','PPC In Transit (arrivals)',11),('ppcsup','PPC In Transit (Supplier Name)',30),
          ('ppcpp','Plant + Port Days (PPC In Transit)',12),('dlv','Delivered',11),('sch','Scheduled Delivery',12),
          ('po','PO',13),('ship','Shipper',16),('rem','Remarks',22),('book','Booking Approval Date',12),('ready','Readiness',11),
          ('etd','ETD',11),('bl','BL#',18),('vessel','Vessel ETA',26),('status','Status',11),('inv','Invoice',13),('ref','Reference',13),
          ('rem2','Remarks 2',18),('vlock','Vendor rule: not usable (Settings ALLOWED VENDORS)',14),('vpool','Vendor rule: stock + arrivals - demand by allowed vendor (up to that day)',46)]
_ALIGN_SFX=re.compile(r'\s*\(aligned \d{4}-\d{2}-\d{2}\)$')

WSA_REFS=[None]              # the WSA planning blocks of this run (PO + BL# of every row): PPC In_Transit lines are matched to them
ALIGN_STAT={}                # align_wsa: ppc_sheets = [planning sheets with PPC arrivals, planning sheets]
CAP=[None]                   # align_capture of this run (every run with MRP_Daily): Inbound 'Other materials' + PPC In Transit
# ---- PPC In_Transit -> YOUR BOM material: supplier + description + RM Code / Final RM Code (Final RM Code is sometimes wrong) ----
_KIND_KW=[('CELL FIXATION','TAPE'),('FIXATION TAPE','TAPE'),('PERFORATED TAPE','TAPE'),('TAPE','TAPE'),('EPE','EPE'),('EVA','EVA'),('POE','POE'),
          ('BUSBAR','BUSBAR'),('BUS BAR','BUSBAR'),('ROUND WIRE','RIBBON'),('RIBBON HOLDER','OTHER'),('RIBBON','RIBBON'),('GLASS','GLASS'),('FRAME','FRAME'),
          ('JUNCTION BOX','JB'),('J-BOX','JB'),('CABLE','JB'),('SEALANT','SEALANT'),('RTV','SEALANT'),('POTTING','POTTING'),('LABEL','LABEL'),('BARCODE','LABEL'),
          ('PALLET','PALLET'),('INNERBOX','BOX'),('BOX','BOX'),('N-TYPE','CELL'),('SOLAR CELL','CELL')]
_SUP_PLACES={'SHANGHAI','VIETNAM','CHINA','HONG','KONG','ZHEJIANG','JIANGSU','NINGBO','HANGZHOU','TAICANG','NINGHAI','SHANDONG','THAILAND','INDIA',
             'USA','HK','SUZHOU','CHANGZHOU','WUXI','GUANGDONG','SHENZHEN','JIAXING'}
_SUP_STOP={'CO','LTD','LIMITED','INC','LLC','COMPANY','COMPA','CORP','CORPORATION','GROUP','THE','AND','OF','PVT','PRIVATE','TRADE','TRADING','INTERNATIONAL',
           'TECHNOLOGY','MATERIAL','MATERIALS','NEW','INDUSTRY','INDUSTRIAL','INVESTMENT','SCIENCE','APPLIED','ENERGY','SOLAR','PV','POWER','ELECTRIC',
           'ELECTRONIC','ELECTRONICS','MANUFACTURING','IMPORT','EXPORT','HOLDINGS','SYSTEMS','DATA','FILM'}
SUPPLIER_ALIAS={'KB':'KIBING'}      # short supplier names on your BOM sheet -> the name on the shipping documents
def _kind_of(t):
    """Product type of a description = the FIRST product word in it ('EPE FILM - EPE (EVA, ...)' is EPE, not EVA)."""
    u=' '.join(str(t or '').upper().split()); best=None
    for kw,k in _KIND_KW:
        i=u.find(kw)
        if i>=0 and (best is None or i<best[0] or (i==best[0] and len(kw)>best[2])): best=(i,k,len(kw))
    return best[1] if best else None
def _sup_tokens(s):
    """Distinctive words of a supplier name (no place names, no 'CO LTD ...'): 'TAICANG JUREN INTERNATIONAL TRADE CO., LTD' -> {'JUREN'}."""
    return {SUPPLIER_ALIAS.get(w,w) for w in re.split(r'[^A-Z0-9]+',str(s or '').upper()) if len(w)>1 and w not in _SUP_STOP and w not in _SUP_PLACES}
def _code_family(c):
    c=str(c).upper()
    for g in (('RMBUS','RMRIB'),('RMEVA','RMEPE'),('RMTAP','RMTTT'),('RMGLS',),('RMALF',),('RMRTV',),('RMJUB',),('RMLAB',),('PMPK',),('CST',),('RMFC',)):
        if c.startswith(g): return g
    return (c[:5],)
def in_transit_catalog(ppc_path, mis_path, DS=None):
    """The materials an In_Transit line can be mapped to: YOUR PPC BOM (code, material description, supplier) first, then the
       other MRP materials (description from MIS Common Name / MB52)."""
    cat={}
    try:
        for r in wb_ro(ppc_path)['BOM'].iter_rows(min_row=2,values_only=True):
            if r and len(r)>6 and r[4]:
                c=str(r[4]).strip().upper(); e=cat.setdefault(c,dict(desc=str(r[5] or '').strip(),sup=set(),src='PPC BOM'))
                if r[6]: e['sup'].add(' '.join(str(r[6]).split()))
    except Exception as _e_: _swallow(7775,_e_); pass
    if mis_path:
        try:
            w=wb_ro(mis_path)
            for r in w['Common Name'].iter_rows(min_row=2,values_only=True):
                if r and r[0] and r[1] and str(r[0]).strip().upper().startswith(('RM','PM','CS')): cat.setdefault(str(r[0]).strip().upper(),dict(desc=str(r[1]),sup=set(),src='MIS'))
            for r in w['MB52'].iter_rows(min_row=2,values_only=True):
                if r and len(r)>5 and r[2] and r[5] and str(r[2]).strip().upper().startswith(('RM','PM','CS')): cat.setdefault(str(r[2]).strip().upper(),dict(desc=str(r[5]),sup=set(),src='MB52'))
        except Exception as _e_: _swallow(7783,_e_); pass
    for b in (DS or {}).get('blocks') or []:
        cat.setdefault(str(b['mat']).upper(),dict(desc=str(b.get('common') or b['mat']),sup=set(),src='MRP'))
    return cat
def _po_set(v):
    """PO numbers in a cell (SAP numbers of 8+ digits; one cell can hold several)."""
    return set(re.findall(r'\d{8,}',str(v).replace('﻿',''))) if v not in (None,'') else set()
def _bl_set(v):
    """BL numbers in a cell (HBL / MBL / BL#; one cell can hold several): words of 6+ letters / digits with 4+ digits."""
    if v in (None,''): return set()
    if isinstance(v,float) and v.is_integer(): v=int(v)
    return {t for t in re.split(r'[^A-Z0-9]+',str(v).upper().replace('﻿','')) if len(t)>=6 and sum(ch.isdigit() for ch in t)>=4}
def wsa_ref_index(groups):
    """PO and BL# of your WSA planning sheets (every dated row, and the Archive sheet of an aligned file) -> lookup tables."""
    sheets=[]; by_po={}; by_bl={}; by_tail={}
    for g in groups or []:
        k=len(sheets); sheets.append(dict(label=g['label'],sheet=g['sheet'],codes=set(g['codes'])))
        for po,bl in g.get('refs') or []:
            ps=_po_set(po); bs=_bl_set(bl)
            for p in ps: by_po.setdefault(p,set()).add(k)
            for b in bs:
                by_bl.setdefault(b,[]).append((k,ps))
                if len(b)>=8: by_tail.setdefault(b[-8:],[]).append((k,ps,b))
    return dict(sheets=sheets,po=by_po,bl=by_bl,tail=by_tail)
def wsa_find(idx, po, bls):
    """The WSA sheets that have this shipment: rows with the PO AND one of its BL numbers (HBL / MBL; 'CMDUCHN4068035' = 'CHN4068035'),
       else rows with one of its BL numbers, else rows with the PO. Returns (how matched, [sheet numbers])."""
    if not idx: return None,[]
    ps=_po_set(po); pb=set(); b_=set()
    for b in bls:
        hits=list(idx['bl'].get(b,[]))
        if len(b)>=8: hits+=[(k,p_) for k,p_,b2 in idx['tail'].get(b[-8:],[]) if b2!=b and (b2.endswith(b) or b.endswith(b2))]
        for k,p_ in hits:
            b_.add(k)
            if ps&p_: pb.add(k)
    if pb: return 'PO + BL#',sorted(pb)
    if b_: return 'BL#',sorted(b_)
    p_=set().union(*(idx['po'].get(p,set()) for p in ps)) if ps else set()
    return ('PO',sorted(p_)) if p_ else (None,[])

def map_in_transit(lines, cat, wsa=None):
    """Each line -> the best material of the catalog.
       1) YOUR WSA sheets first (user rule, Nov 2026): the sheet(s) where you typed this shipment's PO + BL# (HBLNo. / MBL; else the
          BL#, else the PO) - the material of that sheet that fits the description (same product type, no contradicting size /
          thickness / HS / GSM / front-back) is taken; several fit -> the best description match. A WSA sheet whose material
          contradicts the description is reported ('check') and not used.
       2) Else the description (sizes, thickness, HS, GSM, front / back, product type - a contradicting size counts against),
          the supplier (your BOM supplier, or the codes this supplier ships on the same sheet), RM Code (shipped) > Final RM Code
          (PO) > Match / PO-wise codes. The Final RM Code is replaced only when another code is clearly better (2+ points); a line
          with no code at all needs a strong, specific description match, else it stays unmapped."""
    CODE_=re.compile(r'^(RM|PM|CS)[A-Z0-9\-]{6,}$')
    al=(wsa or {}).get('alias') or {}; A=lambda c: al.get(c,c); idx=(wsa or {}).get('idx')
    sheet_sup={}
    for x in lines:
        for c in (x['rm_code'],x['final_code']):
            if CODE_.match(c) and c in cat: sheet_sup.setdefault(c,set()).update(_sup_tokens(x['supplier']))
    kd_cache={}
    def keys(t):
        if t not in kd_cache: kd_cache[t]=(_desc_keys(t),_kind_of(t))
        return kd_cache[t]
    for x in lines:
        desc=x['description']; st=_sup_tokens(x['supplier']); ka,k1=keys(desc)
        ev={}
        for c,w in ((x['rm_code'],1.5),(x['final_code'],1.0),(x['match'],0.5)):
            if CODE_.match(c) and c in cat: ev[c]=ev.get(c,0)+w
        for c in x['po_codes']:
            if c in cat: ev[c]=ev.get(c,0)+0.5
        sup_ok=lambda c: bool(st) and (any(_sup_tokens(s_)&st for s_ in cat[c]['sup']) or bool(sheet_sup.get(c,set())&st))
        lvl,ks=wsa_find(idx,x.get('po'),x.get('bls') or set())
        wsh=[idx['sheets'][k] for k in ks]; wcodes={A(c) for s in wsh for c in s['codes']}
        x['wsa_how']=lvl; x['wsa_sheets']=[s['label'] for s in wsh]
        fams={_code_family(c) for c in ev}
        cands=set(ev)|{c for c in cat if (fams and _code_family(c) in fams) or (st and sup_ok(c)) or A(c) in wcodes}
        sc={}
        for c in cands:
            kb,k2=keys(cat[c]['desc']); s_=_desc_score(desc,cat[c]['desc'])
            for f in ('dim','th','hs','wd','gsm','len'):
                if ka[f] and kb[f] and (ka[f]-kb[f]) and (kb[f]-ka[f]): s_-=3          # a size / thickness / HS / GSM that contradicts
            if k1 and k2: s_+=4 if k1==k2 else -6
            if sup_ok(c): s_+=1.5
            sc[c]=s_+ev.get(c,0)
        def fits(c):                       # same product type and nothing in the description contradicts it
            kb,k2=keys(cat[c]['desc'])
            return bool(k1 and k2 and k1==k2) and not any(ka[f] and kb[f] and (ka[f]-kb[f]) and (kb[f]-ka[f]) for f in ('dim','th','hs','wd','gsm','len','side'))
        def same_part(c):                  # same product type and side as the line, but another size / thickness / HS / GSM
            kb,k2=keys(cat[c]['desc'])
            if not (k1 and k2 and k1==k2) or (ka['side'] and kb['side'] and ka['side']!=kb['side']): return False
            return any(ka[f] and kb[f] and (ka[f]-kb[f]) and (kb[f]-ka[f]) for f in ('dim','th','hs','wd','gsm','len'))
        x['score']=None; x['final_score']=sc.get(x['final_code']); x['bom_desc']=''; x['bom_sup']=''; x['src']=''
        fc=x['final_code']; where=', '.join(x['wsa_sheets'])
        okw=sorted((c for c in sc if A(c) in wcodes and fits(c)),key=lambda c:-sc[c])
        if okw:                                                        # 1) your WSA sheet decides
            best=okw[0]; near=[c for c in okw if sc[best]-sc[c]<1.0]
            for pref in (x['rm_code'],fc,x['match']):
                if pref in near: best=pref; break
            how=f"WSA {lvl} on sheet {where}: "+('Final RM Code confirmed' if best==fc else f"Final RM Code {fc or '-'} replaced")
        else:                                                          # 2) description + supplier + codes
            if not sc: x['code']=None; x['how']='not a BOM material (no code, no matching supplier / description)'; continue
            order=sorted(sc,key=lambda c:-sc[c]); best=order[0]
            near=[c for c in order if sc[best]-sc[c]<1.0]
            for pref in (x['rm_code'],fc,x['match']):
                if pref in near: best=pref; break
            if fc in cat and best!=fc and sc[best]-sc[fc]<2.0: best=fc; how='Final RM Code kept (nothing clearly better)'
            elif fc in cat and best==fc: how='Final RM Code confirmed by description' + (' + supplier' if sup_ok(fc) else '')
            elif not ev:
                spec=any(ka[f] for f in ('dim','th','hs','wd','gsm','len','side'))
                if not (sc[best]>=6 and k1 and k1==keys(cat[best]['desc'])[1] and spec):
                    x['code']=None; x['how']='not a BOM material (description matches nothing clearly)'; continue
                how='mapped by description + supplier (no code on the line)'
            elif best==x['rm_code']: how=f"RM Code used - its description fits; Final RM Code {fc or '-'} does not"
            else: how=f"mapped by description + supplier (Final RM Code {fc or '-'}, RM Code {x['rm_code'] or '-'})"
            if wsh and A(best) not in wcodes:                         # your WSA sheet has the same part in another size / HS = check
                cl_=sorted(c for c in cat if A(c) in wcodes and same_part(c))
                if cl_: how+=f" - CHECK: your WSA sheet {where} has this {lvl} as {cl_[0]} {cat[cl_[0]]['desc']}"
        x['code']=best; x['how']=how; x['score']=sc[best]
        x['bom_desc']=cat[best]['desc']; x['bom_sup']=' / '.join(sorted(cat[best]['sup'])); x['src']=cat[best]['src']
    return lines

IT_HDR_KEYS=('final pending qty.','final pending qty','final rm code','rm code','final qty.','final ata2','final ata','revised eta port','supplier name','ppc grn')
def _it_sheet(names):
    """Your In_Transit sheet: 'In_Transit', else the same name written another way ('In Transit', 'IN-TRANSIT', 'PPC In Transit')."""
    k_=lambda s: re.sub(r'[^a-z0-9]','',str(s).lower())
    return (next((s for s in names if s=='In_Transit'),None) or next((s for s in names if k_(s)=='intransit'),None)
            or next((s for s in names if 'intransit' in k_(s)),None))
def _num_cell(v):
    """A quantity cell -> float; a number typed as text ('1,200') counts too. None when it is not a number."""
    if v is None or isinstance(v,bool): return None
    if isinstance(v,(int,float)): return float(v)
    s_=str(v).strip().replace(',','')
    return float(s_) if re.fullmatch(r'-?\d+(?:\.\d+)?',s_) else None
def _date_cell(v):
    """A date cell -> Timestamp (day): a real date, an Excel date number, or a date typed as text (2026-11-10, 10-Nov-2026, 10 Nov 26)."""
    if isinstance(v,(dt.datetime,dt.date)): return pd.Timestamp(v).normalize()
    if isinstance(v,(int,float)) and not isinstance(v,bool): return (pd.Timestamp('1899-12-30')+pd.Timedelta(days=int(v))) if 20000<v<80000 else None
    if isinstance(v,str):
        for f_ in ('%Y-%m-%d','%Y-%m-%d %H:%M:%S','%d-%b-%Y','%d-%b-%y','%d %b %Y','%d %b %y','%b %d, %Y','%d-%B-%Y'):
            try: return pd.Timestamp(dt.datetime.strptime(v.strip(),f_))
            except ValueError: pass
    return None
def _it_formulas(path, sn, rows_, cols_):
    """How many of these cells (Excel row numbers x column indexes) hold a formula: such a cell reads as empty when the file
       was last saved outside Excel (no saved result)."""
    try: wb=load_workbook(path,read_only=True)
    except Exception as _e_: _swallow(7926,_e_); return 0
    want=set(rows_); n=0
    for i,r in enumerate(wb[sn].iter_rows(max_row=max(want),values_only=True),1):
        if i in want: n+=sum(1 for j in cols_ if j is not None and j<len(r) and isinstance(r[j],str) and r[j].startswith('='))
    wb.close(); return n
def _ppc_files():
    """Workbooks with 'PPC' in the name in the folders the PPC file is searched in (newest first)."""
    fs=[os.path.join(fo,n) for fo in (globals().get('_PPC_DIRS') or []) if os.path.isdir(fo) for n in os.listdir(fo)
        if n.lower().endswith('.xlsx') and 'ppc' in n.lower() and not n.startswith(('~$','MRP_Daily_','Inbound_'))]
    return sorted(dict.fromkeys(fs),key=os.path.getmtime,reverse=True)

def read_ppc_in_transit(path, cat=None, why=None, info=None, wsa=None):
    """Your PPC 'In_Transit' sheet (the MIS In Transit columns + your own): one dict per shipment line with a quantity still
       pending - qty = Final Pending Qty. (else Final Qty. - PPC GRN), date = Final ATA2 / Final AtA (the plant arrival = port
       ETA + 5 days; else Revised ETA Port + 5). The material is traced by map_in_transit (supplier + description + RM Code +
       Final RM Code) against YOUR BOM. Returns the lines ([] = none pending), or None when there is no such sheet.
       Also reads: the sheet named another way ('In Transit'), title rows above the header, numbers / dates typed as text.
       why (list) gets the reason when nothing is read, and a NOTE for quantity formulas with no saved result; info (dict) the sheet."""
    why=[] if why is None else why; info={} if info is None else info; bn=os.path.basename(path)
    try: wb=load_workbook(path,read_only=True,data_only=True)
    except Exception as ex_: why.append(f'{bn} could not be opened ({type(ex_).__name__}: {ex_})'); return None
    sn=_it_sheet(wb.sheetnames)
    if not sn: why.append(f"{bn} has no In_Transit sheet (its sheets: {', '.join(wb.sheetnames)})"); wb.close(); return None
    info['sheet']=sn; raw=[]; empty=0
    for i,r in enumerate(wb[sn].iter_rows(values_only=True),1):
        if not any(v not in (None,'') for v in r):
            empty+=1
            if empty>(500 if raw else 2000): break       # the sheet is formatted down to row 1,048,000
            continue
        empty=0; raw.append((i,tuple(r)))
    wb.close()
    if not raw: why.append(f'{bn} sheet {sn} is empty'); return []
    is_hdr=lambda r: sum(k_ in {_hnorm(v) for v in r} for k_ in IT_HDR_KEYS)>=2
    hi=next((k for k,(i,r) in enumerate(raw[:30]) if is_hdr(r)),0); info['header_row']=raw[hi][0]
    hdr=[_hnorm(h) for h in raw[hi][1]]
    ix=lambda *ns: next((hdr.index(n) for n in ns if n in hdr),None)
    ixf=lambda f: next((j for j,h in enumerate(hdr) if h and f(h)),None)
    g=lambda r,i: r[i] if (i is not None and i<len(r)) else None
    iC=ix('final rm code'); iC2=ix('rm code'); iC3=ix('match'); iPW=ix('po wise material desc.'); iQ=ix('final pending qty.','final pending qty')
    if iQ is None: iQ=ixf(lambda h:'pending' in h and 'qty' in h)
    iFQ=ix('final qty.','qty'); iG=ix('ppc grn'); iD=ix('final ata2','final ata'); iE=ix('revised eta port'); iPO=ix('po'); iB=ix('hblno.','mbl'); iHB=ix('hblno.','hbl no.','hbl'); iMB=ix('mbl')
    iV=ix('vessel / voy'); iS=ix('status'); iSup=ix('supplier name'); iInv=ix('invoice number'); iDs=ix('description')
    up=lambda v: str(v or '').strip().upper()
    out=[]; noq=[]
    for xr,r in raw[hi+1:]:
        q=_num_cell(g(r,iQ))
        if q is None:
            fq=_num_cell(g(r,iFQ)); gr=_num_cell(g(r,iG))
            if fq is None and any(g(r,j_) not in (None,'') for j_ in (iC,iC2,iSup)): noq.append(xr)
            q=(fq or 0.0)-(gr or 0.0)
        if not q or q<=0: continue
        d=_date_cell(g(r,iD))
        if d is None:
            e=_date_cell(g(r,iE)); d=(e+pd.Timedelta(days=5)) if e is not None else None
        out.append(dict(final_code=up(g(r,iC)),rm_code=up(g(r,iC2)),match=up(g(r,iC3)),po_codes=re.findall(r'\b(?:RM|PM|CS)[A-Z0-9]{6,}\b',up(g(r,iPW))),
                        description=' '.join(str(g(r,iDs) or '').split()),qty=float(q),date=d,
                        po=g(r,iPO),bl=g(r,iB),bls=_bl_set(g(r,iHB))|_bl_set(g(r,iMB)),vessel=' '.join(str(g(r,iV) or '').split()),status=' '.join(str(g(r,iS) or '').split()),
                        supplier=' '.join(str(g(r,iSup) or '').split()),invoice=g(r,iInv),raw_code=str(g(r,iC) or '').strip()))
    if noq and _it_formulas(path,sn,noq,[iQ,iFQ]):
        why.append(f"{len(noq)} rows of {bn} sheet {sn} have a quantity formula with no saved result (the file was last saved outside Excel)"
                   f" - open {bn} in Excel and save it once")
    if not out and not why:
        why.append(f"{bn} sheet {sn}: no 'Final Pending Qty.' column (header row {raw[hi][0]})" if iQ is None and iFQ is None
                   else f"{bn} sheet {sn}: {len(raw)-hi-1} rows, none with a pending quantity above 0")
    return map_in_transit(out,cat or {},wsa)

def ppc_arrivals(DS, dates):
    """PPC In_Transit -> arrivals per MRP material and day (MRP units: the same code conversions as MIS In Transit, e.g.
       RMTTT00000018 = 2 x RMTTT00000010). A late shipment (arrival before today, still pending) arrives today.
       why = why there are no arrivals (shown in T3 of the WSA sheets, the Inbound sheet 'PPC In Transit', the end of the console)."""
    why=[]; info={}
    none_=lambda lines_=None: dict(arr={},rows=[],skipped=[],src=None,fixes=[],lines=lines_ or [],why=why,sheet=info.get('sheet'))
    if not PPC_SRC[0]:
        why.append('no PPC file found (a workbook with Count + BOM sheets next to MRP.xlsx or the script)'); return none_()
    cat=in_transit_catalog(PPC_SRC[0],MIS_PATH_[0] or globals().get('MIS'),DS)
    wsa_=dict(idx=wsa_ref_index(WSA_REFS[0]),alias=DS.get('alias') or {}) if WSA_REFS[0] else None
    rows=read_ppc_in_transit(PPC_SRC[0],cat,why,info,wsa_)
    if rows is None:
        oth=[os.path.basename(p_) for p_ in _ppc_files() if os.path.abspath(p_)!=os.path.abspath(PPC_SRC[0]) and _it_sheet(_sheetnames(p_))]
        if oth: why.append(f"an In_Transit sheet is in {', '.join(oth)}, but the PPC file used is {os.path.basename(PPC_SRC[0])} "
                           "(the newest workbook with Count + BOM sheets) - keep one PPC file in the folder")
        return none_()
    if not rows: return none_()
    fixes=[x for x in rows if x.get('code') and x['code']!=x['final_code']]
    d0=pd.Timestamp(dates[0]).normalize(); nd=len(dates); idx={pd.Timestamp(d).normalize():k for k,d in enumerate(dates)}
    alias=DS.get('alias') or {}; fac=DS.get('alias_factor') or {}; mats={b['mat'] for b in DS['blocks']}
    arr={}; used=[]; skipped=[]; sup={}
    for x in rows:
        c=x['code']
        x['mat']=None; x['qty_mrp']=None; x['arrive']=None; x['late']=False; x['in_mrp']=False
        if not c: skipped.append(f"{x['description'][:28]} - {x['supplier'][:24]} PO {x['po']} qty {x['qty']:,.0f}"); continue
        m=alias.get(c,c); q=x['qty']*fac.get(c,1.0)
        if x['date'] is None: skipped.append(f"{c} PO {x['po']} qty {x['qty']:,.0f} (no arrival date)"); continue
        x.update(mat=m,qty_mrp=q,arrive=max(x['date'],d0),late=x['date']<d0,in_mrp=m in mats)
        k=idx.get(max(x['date'],d0))
        if k is None: continue                                        # after the planning horizon
        arr.setdefault(m,[0.0]*nd)[k]+=q; used.append(x); sup.setdefault(m,{}).setdefault(k,set()).add(x['supplier'] or '?')
    if not arr:
        nd_=sum(1 for x in rows if x.get('code') and x['date'] is None); nm_=sum(1 for x in rows if not x.get('code'))
        why.append(f"{len(rows)} lines read from {os.path.basename(PPC_SRC[0])}, none arrives in the planning period ({nd_} with no arrival date, {nm_} not BOM materials)")
    return dict(arr=arr,rows=used,skipped=skipped,src=os.path.basename(PPC_SRC[0]),fixes=fixes,lines=rows,why=why,sheet=info.get('sheet'),sup=sup)
def _ppc_arrivals_safe(DS, dates):
    """ppc_arrivals; an unexpected cell in the In_Transit sheet must not stop the WSA alignment - it is reported instead."""
    try: return ppc_arrivals(DS,dates)
    except Exception as ex_:
        import traceback
        tb_=traceback.extract_tb(ex_.__traceback__)[-1]
        return dict(arr={},rows=[],skipped=[],src=None,fixes=[],lines=[],sheet=None,
                    why=[f'error reading the PPC In_Transit sheet ({type(ex_).__name__}: {ex_}, line {tb_.lineno})'])

def align_capture(DS, pp_rows, dates):
    """After the OEM step of build_new_mrp: per material its non-OEM plan, OEM plan, own stock and per (MIS customer, line) use;
       the module plan (and MW) per (MIS customer, line)."""
    n_=len(dates); mods={}
    for x in pp_rows or []:
        k=(_norm(x.get('mis_cust') or x.get('cust')),' '.join(str(x.get('line') or '').split()))
        a=mods.setdefault(k,dict(cust=[],mis=' '.join(str(x.get('mis_cust') or x.get('cust') or '').split()),line=k[1],v=[0.0]*n_,mw=[0.0]*n_))
        cu=' '.join(str(x.get('cust') or '').split())
        if cu and cu not in a['cust']: a['cust'].append(cu)
        for i,v in enumerate((x.get('v') or [])[:n_]): a['v'][i]+=float(v or 0)
        for i,v in enumerate((x.get('mw') or [])[:n_]): a['mw'][i]+=float(v or 0)
    bl={}
    for b in DS['blocks']:
        cl={}; oem=set(); disp={}
        for src,is_oem in ((b.get('cust_line') or {},False),(b.get('cust_line_oem') or {},True)):
            for k,v in src.items():
                cu,ln=(k if isinstance(k,tuple) else (k,None))
                key=(_norm(cu),' '.join(str(ln or '').split()))
                disp.setdefault(key,(' '.join(str(cu or '').split()),' '.join(str(ln or '').split())))
                a=cl.setdefault(key,[0.0]*n_)
                for i,q in enumerate(list(v)[:n_]): a[i]+=float(q or 0)
                if is_oem: oem.add(key)
        sup=[]                                      # stock by supplier = the MRP_Daily supplier rows: MB52 batches, or the count by supplier
        for sp in b.get('sup') or []:
            nm=str(sp.get('name') or ''); st=float(sp.get('stock') or 0.0)
            sup.append((nm,st))                     # counted: already the count by supplier, rolled forward (apply_count_by_supplier)
        bl[b['mat']]=dict(plan=list(b.get('plan') or [0.0]*n_),plan_oem=list(b.get('plan_oem') or [0.0]*n_),cl=cl,oem=oem,disp=disp,
                          basis=str(b.get('plan_basis') or ''),stock=float(b.get('stock') or 0),jw=float(b.get('st_jw') or 0) if b.get('jw_out') else 0.0,
                          uom=b.get('uom'),common=b.get('common'),desc=b.get('desc'),sup=sup)
    return dict(dates=[pd.Timestamp(d).normalize() for d in dates],blocks=bl,mods=mods,alias=dict(DS.get('alias') or {}),ppc=_ppc_arrivals_safe(DS,dates))

def _wsa_lookup(src):
    """WSA 'Lookup' sheet: (MIS plant, MIS customer) -> customer column label; customer -> labels (any plant)."""
    by={}; bycu={}
    try:
        for r in wb_ro(src)['Lookup'].iter_rows(min_row=2,values_only=True):
            if not r or len(r)<4 or not r[1] or not r[2] or not r[3]: continue
            pl,cu,lab=_norm(r[1]),_norm(r[2]),' '.join(str(r[3]).split())
            by[(pl,cu)]=lab; bycu.setdefault(cu,Counter())[lab]+=1
    except Exception as ex_: print(f"NOTE: WSA 'Lookup' sheet not read ({type(ex_).__name__}: {ex_}) - customer columns named after the MIS customer")
    return by,bycu

def _plant_keys(line):
    """Plan production line -> the plant names the WSA Lookup uses ('USA Line 3' -> 'l3', 'Runergy' -> 'runergy l1' ...)."""
    n=_norm(line); m=re.match(r'^(?:usa\s*)?line\s*(\d+)$',n) or re.match(r'^l\s*(\d+)$',n)
    if m: return [f'l{m.group(1)}']
    if n.startswith('runergy') or n.startswith('runenergy'): return [n,'runergy l1','runergy l2','runergy']
    if n.startswith('imperial') or n.startswith('ic star'): return [n,'imperial star']
    if n.startswith('arizona'): return [n,'arizona']
    if n.startswith('inox'): return [n,'inox 1','inox 2']
    return [n]

def _hnorm(x): return ' '.join(str(x if x is not None else '').replace('\n',' ').split()).lower()
def _hdr_role(h):
    """Planning-sheet header (normalised) -> field."""
    if not h: return None
    if h.startswith('required plant date'): return 'date'
    if h in ('d','day'): return 'day'
    if h.startswith('balance'): return 'bal'
    if h.startswith('planned modules') or h.startswith('modules per day'): return 'plan'
    if h.startswith('demand'): return 'dem'
    if h.startswith('actual modules'): return 'actmod'
    if h.startswith('actual mw'): return 'actmw'
    if h=='actual' or h.startswith('actual consum'): return 'act'
    if h.startswith('mw per day'): return 'mw'
    if h.startswith('plant days'): return 'pdays'
    if h.startswith('plant + port days (ppc'): return 'ppcpp'
    if h.startswith('ppc in transit (supplier'): return 'ppcsup'
    if h.startswith('ppc in transit'): return 'ppcit'
    if h.startswith('plant + port'): return 'pport'
    if h.startswith('delivered'): return 'dlv'
    if h.startswith('scheduled'): return 'sch'
    if h=='po' or h.startswith('po '): return 'po'
    if h.startswith('shipper'): return 'ship'
    if h.startswith('remarks'): return 'rem'
    if h.startswith('booking'): return 'book'
    if h.startswith('readiness'): return 'ready'
    if h=='etd' or h.startswith('etd '): return 'etd'
    if h.startswith('bl'): return 'bl'
    if h.startswith('vessel'): return 'vessel'
    if h.startswith('invoice'): return 'inv'
    if h.startswith('reference'): return 'ref'
    if h.startswith('status'): return 'status'
    if h.startswith('vendor rule: not usable'): return 'vlock'
    if h.startswith('vendor rule'): return 'vpool'
    return 'other'

def _wsa_read(wv, wf):
    """Planning blocks of a WSA file in any layout: the column of each field by its header, the customer columns (between
       Balance and Delivered), codes, unit, material name and the row of each date."""
    out=[]
    for sh in wv.sheetnames:
        if sh in SKIP or sh in (ALIGN_SHEET,ARCHIVE_SHEET): continue
        ws=wv[sh]
        if ws.sheet_state!='visible': continue
        rv=[list(r) for r in ws.iter_rows(values_only=True)]
        starts=sorted([(r+1,c+1) for r,row in enumerate(rv[:12]) for c,x in enumerate(row[:200]) if str(x).strip().startswith('Required Plant Date')],key=lambda x:x[1])
        if not starts: continue
        rf=[list(r) for r in wf[sh].iter_rows(values_only=True)]
        maxc=max(len(r) for r in rv)
        for bi,(hr,c0) in enumerate(starts):
            c1=(starts[bi+1][1]-1) if bi+1<len(starts) else maxc
            role={'date':c0}; cust=[]; dropped=[]
            for c in range(c0+1,c1+1):
                raw=cv(rv,hr,c); ro=_hdr_role(_hnorm(raw))
                if ro is None: continue
                if ro=='day' and c==c0+1: role['day']=c; continue
                if ro in ('other','day'):
                    if 'bal' in role and 'dlv' not in role: cust.append((c,' '.join(str(raw).split())))
                    elif 'dlv' in role: dropped.append(' '.join(str(raw).split()))
                    continue
                if ro=='rem' and 'rem' in role: ro='rem2'
                if ro=='act': role['act']=c; continue              # 'Actual consumption' twice (EPE 480): the last one is the one Balance uses
                role.setdefault(ro,c)
            if 'bal' not in role or 'sch' not in role or 'dem' not in role: continue
            if 'book' not in role and 'rem' in role and 'ready' in role:          # a Booking column without a header (EPE 420 GSM)
                for c in range(role['rem']+1,role['ready']):
                    if not _hnorm(cv(rv,hr,c)): role['book']=c; break
            codes=[]; uom=''; mname=''
            for r in range(1,hr):
                for c in range(c0,c1+1):
                    x=cv(rv,r,c)
                    if isinstance(x,str):
                        if CODE.match(x.strip()) and x.strip() not in codes: codes.append(x.strip())
                        if x.strip().startswith('Unit of Measurement'): uom=str(cv(rv,r,c+2) or '').strip()
                        if x.strip().startswith('Supplier & Material'): mname=str(cv(rv,r,c+2) or '').strip()
            dates={}
            for r in range(hr+1,len(rv)+1):
                x=cv(rv,r,c0)
                if isinstance(x,(dt.datetime,dt.date)): dates.setdefault(pd.Timestamp(x).normalize(),r)
            out.append(dict(sheet=sh,block=bi,nblocks=len(starts),label=LABELS.get((sh,bi)) or (sh if len(starts)==1 else f'{sh} #{bi+1}'),
                            hr=hr,c0=c0,c1=c1,role=role,cust=cust,dropped=dropped,codes=codes,uom=uom,mname=mname,dates=dates,rv=rv,rf=rf))
    return out

class _Styles:
    """styles.xml of the WSA file + the few cell formats the rebuilt sheets use (appended once, reused on the next run)."""
    def __init__(self, sty):
        import copy
        from lxml import etree
        self.E=etree; self.cp=copy; self.M='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
        self.st=etree.fromstring(sty); M=self.M
        self.nfs=self.st.find(M+'numFmts')
        if self.nfs is None: self.nfs=etree.Element(M+'numFmts',count='0'); self.st.insert(0,self.nfs)
        self.fonts=self.st.find(M+'fonts'); self.fills=self.st.find(M+'fills'); self.borders=self.st.find(M+'borders'); self.xfs=self.st.find(M+'cellXfs')
        self.dxfs=self.st.find(M+'dxfs')
        if self.dxfs is None:
            self.dxfs=etree.Element(M+'dxfs',count='0'); cs=self.st.find(M+'cellStyles')
            (cs if cs is not None else self.xfs).addnext(self.dxfs)
        self.f0=self.fonts.findall(M+'font')[0]
    @staticmethod
    def _sig(e): return (e.tag,tuple(sorted(e.attrib.items())),tuple(_Styles._sig(c) for c in e))
    def _add(self, parent, el, tag):
        M=self.M; kids=parent.findall(M+tag); sg=self._sig(el)
        for i,k in enumerate(kids):
            if self._sig(k)==sg: return i
        parent.append(el); parent.set('count',str(len(kids)+1)); return len(kids)
    def numfmt(self, code):
        M=self.M
        for e in self.nfs.findall(M+'numFmt'):
            if e.get('formatCode')==code: return int(e.get('numFmtId'))
        nid=max([163]+[int(e.get('numFmtId')) for e in self.nfs.findall(M+'numFmt')])+1
        self.E.SubElement(self.nfs,M+'numFmt',numFmtId=str(nid),formatCode=code); self.nfs.set('count',str(len(self.nfs.findall(M+'numFmt'))))
        return nid
    def font(self, bold=False, size=None, rgb=None):
        M=self.M; f=self.cp.deepcopy(self.f0)
        for ch in list(f):
            if ch.tag in (M+'b',M+'sz',M+'color') or (ch.tag==M+'scheme' and rgb): f.remove(ch)
        if bold: f.insert(0,self.E.Element(M+'b'))
        sz=self.E.Element(M+'sz'); sz.set('val',str(size or self._size0())); f.append(sz)
        if rgb: c=self.E.Element(M+'color'); c.set('rgb',rgb); f.append(c)
        else:
            c0=self.f0.find(M+'color')
            if c0 is not None: f.append(self.cp.deepcopy(c0))
        return self._add(self.fonts,f,'font')
    def _size0(self):
        s=self.f0.find(self.M+'sz'); return s.get('val') if s is not None else '11'
    def fill(self, rgb):
        M=self.M; f=self.E.Element(M+'fill'); p=self.E.SubElement(f,M+'patternFill',patternType='solid')
        self.E.SubElement(p,M+'fgColor',rgb=rgb); self.E.SubElement(p,M+'bgColor',indexed='64')
        return self._add(self.fills,f,'fill')
    def border(self):
        M=self.M; b=self.E.Element(M+'border')
        for side in ('left','right','top','bottom'):
            s=self.E.SubElement(b,M+side,style='thin'); self.E.SubElement(s,M+'color',rgb='FFBFBFBF')
        self.E.SubElement(b,M+'diagonal')
        return self._add(self.borders,b,'border')
    def xf(self, nf=0, font=0, fill=0, border=0, h=None, v=None, wrap=False):
        M=self.M; a={'numFmtId':str(nf),'fontId':str(font),'fillId':str(fill),'borderId':str(border),'xfId':'0'}
        if nf: a['applyNumberFormat']='1'
        if font: a['applyFont']='1'
        if fill: a['applyFill']='1'
        if border: a['applyBorder']='1'
        e=self.E.Element(M+'xf',**a)
        if h or v or wrap:
            e.set('applyAlignment','1'); al=self.E.SubElement(e,M+'alignment')
            if h: al.set('horizontal',h)
            if v: al.set('vertical',v)
            if wrap: al.set('wrapText','1')
        return self._add(self.xfs,e,'xf')
    def dxf(self, font_rgb, fill_rgb, bold=False):
        M=self.M; d=self.E.Element(M+'dxf'); f=self.E.SubElement(d,M+'font')
        if bold: self.E.SubElement(f,M+'b')
        self.E.SubElement(f,M+'color',rgb=font_rgb)
        p=self.E.SubElement(self.E.SubElement(d,M+'fill'),M+'patternFill'); self.E.SubElement(p,M+'bgColor',rgb=fill_rgb)
        return self._add(self.dxfs,d,'dxf')
    def yellow_xfs(self):
        """cellXfs indexes with a yellow fill (the planner's typed balances)."""
        M=self.M; fills=self.fills.findall(M+'fill'); out=set()
        for i,x in enumerate(self.xfs.findall(M+'xf')):
            try:
                pf=fills[int(x.get('fillId') or 0)].find(M+'patternFill'); fg=pf.find(M+'fgColor') if pf is not None else None
                if fg is not None and str(fg.get('rgb') or '').upper() in ('FFFFFF00','FFFF00'): out.add(str(i))
            except (IndexError,ValueError): pass
        return out
    def bytes(self): return self.E.tostring(self.st,xml_declaration=True,encoding='UTF-8',standalone=True)

def _xcell(ref, v=None, s=0, f=None):
    """One <c> element as text: formula + cached value, a number, a date (as serial), text (inline) or an empty styled cell."""
    from xml.sax.saxutils import escape
    sa=f' s="{s}"' if s else ''
    if isinstance(v,(dt.datetime,dt.date)): v=(pd.Timestamp(v).normalize()-pd.Timestamp('1899-12-30')).days
    if isinstance(v,float) and v!=v: v=None                        # NaN
    def num(x): return str(int(x)) if float(x)==int(x) and abs(x)<1e15 else repr(float(x))
    if f:
        fx='<f>'+escape(f)+'</f>'
        if isinstance(v,bool): return f'<c r="{ref}"{sa} t="b">{fx}<v>{int(v)}</v></c>'
        if isinstance(v,(int,float)): return f'<c r="{ref}"{sa}>{fx}<v>{num(v)}</v></c>'
        if isinstance(v,str): return f'<c r="{ref}"{sa} t="str">{fx}<v>{escape(v)}</v></c>'
        return f'<c r="{ref}"{sa}>{fx}</c>'
    if isinstance(v,bool): return f'<c r="{ref}"{sa} t="b"><v>{int(v)}</v></c>'
    if isinstance(v,(int,float)): return f'<c r="{ref}"{sa}><v>{num(v)}</v></c>'
    if isinstance(v,str) and v!='':
        sp=' xml:space="preserve"' if v!=v.strip() or '\n' in v else ''
        return f'<c r="{ref}"{sa} t="inlineStr"><is><t{sp}>{escape(v)}</t></is></c>'
    return f'<c r="{ref}"{sa}/>' if s else ''

def _carry_formula(fx, r_old, colmap, r_new):
    """The planner's formula in a delivery column, moved to the new layout: only constants (=10000*4) or references to cells of
       the same row (=A747-40) whose column exists in the new layout; anything else -> None (the value is copied instead)."""
    s=str(fx)[1:] if str(fx).startswith('=') else str(fx)
    if '!' in s or ':' in s or '"' in s or '[' in s: return None
    from openpyxl.utils import column_index_from_string as CI
    out=[]; pos=0
    for m in re.finditer(r'(\$?)([A-Z]{1,3})(\$?)(\d+)',s):
        if m.start()>0 and (s[m.start()-1].isalpha() or s[m.start()-1]=='_'): return None
        if int(m.group(4))!=r_old: return None
        nc=colmap.get(CI(m.group(2)))
        if nc is None: return None
        out.append(s[pos:m.start()]+m.group(1)+nc+m.group(3)+str(r_new)); pos=m.end()
    out.append(s[pos:])
    if re.search(r'[A-Za-z]{2,}\(',s): return None                     # functions: not moved
    return ''.join(out)

# ---------- front / back glass: one sheet per glass (user rule, Nov 2026) ----------
WSA_DROP_SHEETS=('Total_Plan_Req',)   # left out of the aligned file. 'Stock' is KEPT: it is plain values, and 'Summary Dashboard' (column B) and the
                                      # Frame Topcon 30HS / 35HS sheets read it - dropping it turned those formulas into =SUMIFS(#REF!,#REF!,...)
_GQ_PLAN=('bal','dem','act','ppcit'); _GQ_DLV=('sch','dlv')
_GT=('po','ship','rem','book','ready','etd','bl','vessel','inv','ref','rem2')
def _gcell(rows,r,c): return rows[r-1][c-1] if (r and c and 0<r<=len(rows) and 0<c<=len(rows[r-1])) else None
def _gset(rows,r,c,v):
    while len(rows)<r: rows.append([])
    if len(rows[r-1])<c: rows[r-1].extend([None]*(c-len(rows[r-1])))
    rows[r-1][c-1]=v
def _glass_side(m, AD):
    """'F' (front glass: Front / HTAR / ARC) or 'B' (back glass: Back / Printed) by the MIS Common Name, else None."""
    t=str(((AD.get('blocks') or {}).get(m) or {}).get('common') or '')
    if _kind_of(t)!='GLASS': return None
    s=_desc_keys(t)['side']
    return next(iter(s)) if len(s)==1 else None
def _it_bl_index():
    """MIS In Transit lines (all of them, material traced like MRP_Daily) by BL number: (line, material, Final Qty.)."""
    ix={}; tail={}
    try: rows=list(wb_ro(MIS_PATH_[0])['In Transit'].iter_rows(values_only=True)) if MIS_PATH_[0] else []
    except Exception: rows=[]
    if len(rows)<2: return ix,tail
    I_=IT_CANON.index; iF,iQ,iQ0,iHB,iMB=I_('Final RM Code'),I_('Final Qty.'),I_('Qty'),I_('HBLNo.'),I_('MBL')
    for k,r in enumerate(rows[1:]):
        q=_num_cell(r[iQ]) if iQ<len(r) else None
        if q is None and iQ0<len(r): q=_num_cell(r[iQ0])
        c=str(r[iF] or '').strip().upper()
        if not c or not q: continue
        for b in _bl_set(r[iHB] if iHB<len(r) else None)|_bl_set(r[iMB] if iMB<len(r) else None):
            ix.setdefault(b,[]).append((k,c,q))
            if len(b)>=8: tail.setdefault(b[-8:],[]).append((k,c,q,b))
    return ix,tail
def _mis_bl_dates(mis_path):
    """MIS In Transit: BL number (HBLNo. / MBL) -> the plant arrival dates (Final AtA) of its pending lines.
       User rule, Nov 2026: 'MIS is always the most updated' - a WSA delivery row with another date is to be moved to the MIS date."""
    ix={}; tail={}
    try: rows=list(wb_ro(mis_path)['In Transit'].iter_rows(values_only=True)) if mis_path else []
    except Exception: rows=[]
    if len(rows)<2: return ix,tail
    I_=IT_CANON.index; iA,iHB,iMB=I_('Final AtA'),I_('HBLNo.'),I_('MBL')
    for r in rows[1:]:
        d=_date_cell(r[iA]) if iA<len(r) else None
        if d is None: continue
        d=pd.Timestamp(d).normalize()
        for b in _bl_set(r[iHB] if iHB<len(r) else None)|_bl_set(r[iMB] if iMB<len(r) else None):
            ix.setdefault(b,[]).append(d)
            if len(b)>=8: tail.setdefault(b[-8:],[]).append((b,d))
    return ix,tail
def _mis_date_of(bl_cell, idx):
    """The MIS arrival date of a WSA row: the earliest Final AtA of the MIS lines with one of its BL numbers (suffix match like the
       PPC In_Transit match: MBL 'CMDUCHN4068035' = 'CHN4068035'); None when MIS has none of them."""
    ix,tail=idx; ds=[]
    for b in _bl_set(bl_cell):
        ds+=ix.get(b,[])
        if len(b)>=8: ds+=[d for b2,d in tail.get(b[-8:],[]) if b2!=b and (b2.endswith(b) or b.endswith(b2))]
    return min(ds) if ds else None
def _glass_shares(bl_cell, sides, ix, tail, A):
    """Front / back share of one planner row from the In_Transit lines of its BL numbers (None: no line of these glasses)."""
    per={}; seen=set()
    for b in _bl_set(bl_cell):
        hits=list(ix.get(b,[]))
        if len(b)>=8: hits+=[(k,c,q) for k,c,q,b2 in tail.get(b[-8:],[]) if b2!=b and (b2.endswith(b) or b.endswith(b2))]
        for k,c,q in hits:
            s=sides.get(A(c))
            if s and k not in seen: seen.add(k); per[s]=per.get(s,0.0)+q
    tot=sum(per.values())
    return {s:per.get(s,0.0)/tot for s in ('F','B')} if tot>0 else None
def _glass_part(B, side, codes, name, shares, note, block_no, dash, oname):
    """One glass of a planning block that counts front + back: every quantity of the planner's rows split - deliveries
       (Scheduled, Delivered) by the In_Transit lines of the row's BL# (shares), else 1:1; Balance / Demand / Actual / PPC
       arrivals of the kept rows 1:1 (one front + one back glass per module). Delivery text stays only on the rows whose
       shipment has this glass."""
    rv=[list(x) for x in B['rv']]; rf=[list(x) for x in B['rf']]; ro=B['role']
    for d,r in B['dates'].items():
        sd=(shares.get(r) or {}).get(side,0.5)
        has_q=any((_num(_gcell(B['rv'],r,ro[k])) or 0)!=0 for k in _GQ_DLV if k in ro)
        for k in _GQ_PLAN+_GQ_DLV:
            if k not in ro: continue
            v=_num(_gcell(B['rv'],r,ro[k]))
            if v is None: continue
            nv=round(v*(sd if k in _GQ_DLV else 0.5),6)
            _gset(rv,r,ro[k],None if (k in _GQ_DLV and v and not nv) else nv); _gset(rf,r,ro[k],None)     # none of this glass: blank
        if has_q and sd<=1e-9:
            for k in _GT:
                if k in ro: _gset(rv,r,ro[k],None); _gset(rf,r,ro[k],None)
    return dict(B,rv=rv,rf=rf,codes=list(codes),label=name,oname=oname,block=block_no,note=note,dash=dash,src_label=B['label'])
def _glass_split(blocks, AD):
    """Front and back glass on their own sheets (user rule, Nov 2026: 'I want front glass and back glass separated in different
       sheets' - one Balance per glass; the BOM uses both codes). A planning block with front + back glass becomes
       '<sheet> Front' + '<sheet> Back'. A block that lists one glass but counts its partner too (BL# lines of the other glass:
       Glass G12R Topcon Subanci = 356 front + 333 back) is split the same way: its own glass keeps the sheet ('<sheet> Front'),
       the partner gets '<sheet> Back' - the customer project keeps its own sheets (user rule, Nov 2026: 'follow the MRP daily
       settings ... which supplier goes with which customer'; a glass on 2+ sheets is split by customer and supplier in
       _wsa_shared). Delivery rows are split by the In_Transit lines of their BL#, else 1 front : 1 back. The new sheets carry
       'glass only' in A1, so the next run (on the aligned file) does not split them again. Returns (blocks, info); info['arch']
       splits the Archive records of the split blocks the same way."""
    al=AD.get('alias') or {}; A=lambda c: al.get(c,c)
    ix,tail=_it_bl_index()
    res=[]; info=dict(renames={},lines=[],src={},n_bl=0,n_11=0)
    word={'F':'front','B':'back'}
    def shares_of(B,sides):
        sh={}
        if 'bl' not in B['role']: return sh
        for d,r in B['dates'].items():
            has_q=any((_num(_gcell(B['rv'],r,B['role'][k])) or 0)!=0 for k in _GQ_DLV if k in B['role'])
            s_=_glass_shares(_gcell(B['rv'],r,B['role']['bl']),sides,ix,tail,A)
            if s_ is not None: sh[r]=s_
            if has_q: info['n_bl' if s_ is not None else 'n_11']+=1
        return sh
    def split(o,B,mats,sides,extra=''):
        sh_=B['sheet']; nm={'F':f'{sh_[:25]} Front','B':f'{sh_[:26]} Back'}
        shares=shares_of(B,sides)
        own=sides[A(mats[0])] if extra else 'F'; oth='B' if own=='F' else 'F'
        P=_glass_part(B,own,[c for c in mats if sides[A(c)]==own],nm[own],shares,f"({word[own]} glass only - {word[oth]} glass: sheet '{nm[oth]}')",0,True,nm[own])
        Q=_glass_part(B,oth,[c for c in sides if sides[c]==oth and (not extra or c not in mats)] if extra else [c for c in mats if sides[A(c)]==oth],
                      nm[oth],shares,f"({word[oth]} glass only - {word[own]} glass: sheet '{nm[own]}')",100,False,nm[oth])
        res.extend([(o,0,P),(o,1,Q)]); info['renames'][sh_]=nm[own]; info['src'][B['label']]=dict(F=nm['F'],B=nm['B'],sides=sides)
        info['lines'].append(f"{sh_} -> {nm[own]} ({', '.join(P['codes'])}) + {nm[oth]} ({', '.join(Q['codes'])}{extra})")
    for o,B in enumerate(blocks):
        mats=[c for c in B['codes'] if c in AD['blocks']]
        sides={A(c):_glass_side(c,AD) for c in mats}
        mk_=next((v for row in B['rv'][:max(B['hr']-1,0)] for v in row if isinstance(v,str) and 'glass only' in v.lower()),None)
        marked=mk_ is not None
        if marked:                                 # split by an earlier run: keep its note (and so the marker) in A1
            i_=mk_.lower().find('glass only'); j_=mk_.rfind('(',0,i_); B=dict(B,note=mk_[j_ if j_>=0 else i_:].strip())
            if re.search(r"its (front|back) glass RM|includes the (front|back) glass of sheet",mk_):
                info['lines'].append(f"WARNING: '{B['sheet']}' comes from an aligned file of an older version (a customer's glass added to "
                                     "another customer's sheet) - run the script on your own WSA file")
        if mats and all(sides.values()) and set(sides.values())=={'F','B'}:
            split(o,B,mats,sides)
        elif mats and not marked and all(sides.values()) and len(set(sides.values()))==1 and 'bl' in B['role']:
            own=next(iter(sides.values())); oth='B' if own=='F' else 'F'; cnt=Counter()
            for d,r in B['dates'].items():
                for b in _bl_set(_gcell(B['rv'],r,B['role']['bl'])):
                    for k,c,q in ix.get(b,[]):
                        if _glass_side(c,AD)==oth: cnt[A(c)]+=q
            if cnt:                                # its BL# lines carry the partner glass: this sheet counts front + back
                partner=cnt.most_common(1)[0][0]; s2=dict(sides); s2[partner]=oth
                split(o,B,mats,s2,extra=': its BL# lines')
            else: res.append((o,0,B))
        else: res.append((o,0,B))
    res.sort(key=lambda t:(t[0],t[1]))
    def arch(old):
        """Archive records of a block split in this run -> one record per glass (BL# lines, else 1:1)."""
        out=[]
        for r in old:
            sp=info['src'].get(str(r[0]))
            if not sp: out.append(r); continue
            s_=_glass_shares(r[10],sp['sides'],ix,tail,A) or {'F':0.5,'B':0.5}
            q_d,q_s=_num(r[2]),_num(r[3])
            for side in ('F','B'):
                f=s_.get(side,0.5)
                if (q_d or q_s) and f<=1e-9: continue
                n=list(r); n[0]=sp[side]
                if q_d: n[2]=round(q_d*f,6)
                if q_s: n[3]=round(q_s*f,6)
                out.append(n)
        return out
    info['arch']=arch
    return [B for _,_,B in res],info
def _ref_fix(text, renames, gone):
    """Formulas / defined names: references to a renamed sheet follow the new name; references to a removed sheet become
       #REF! (what Excel does when a sheet is deleted)."""
    for o,n in renames.items():
        text=text.replace(f"'{o}'!",f"'{n}'!").replace(f"&apos;{o}&apos;!",f"&apos;{n}&apos;!")
    cell=r"(?:\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+)"
    for g in gone:
        q=re.escape(g); alt=[f"'{q}'",f"&apos;{q}&apos;"]
        if re.fullmatch(r'[A-Za-z_][A-Za-z0-9_.]*',g): alt.append(r"(?<![A-Za-z0-9_.'\s;&])"+q)
        text=re.sub('(?:'+'|'.join(alt)+')!'+cell,'#REF!',text)
    return text

def _out_sheet_name(B):
    """Name of a planning block's sheet in the aligned WSA file."""
    if B.get('oname'): return B['oname'][:31]
    if B['block']==0: return B['sheet']
    o=WSA_SPLIT_NAMES.get(B['sheet']) if B['block']==1 else None
    return (o or f"{B['sheet'][:27]} {B['block']+1}")[:31]
def _vendor_rules(AD):
    """Settings ALLOWED VENDORS: {(customer, MRP material): [vendors, top row first]} - the customer under its Settings name and
       its MIS name (LINKS). Like MRP_Daily, a customer's need goes to its FIRST allowed vendor; no row = any vendor."""
    al=AD.get('alias') or {}; out={}
    for r_ in SET_CUSTSUP or []:
        cu=r_.get('Customer'); c=str(r_.get('Code') or '').strip(); v=' '.join(str(r_.get('Supplier') or '').split()).upper()
        if not (cu and c and v): continue
        L_=next((x for k,x in LINKS.items() if _norm(k)==_norm(cu)),None) or {}
        for n in {_norm(cu),_norm(L_.get('mis') or '')}-{''}:
            lst=out.setdefault((n,al.get(c,c)),[])
            if v not in lst: lst.append(v)
    return out
def _vend_hit(a, b):
    """Same vendor? Settings 'FLAT' = batch Make 'FLAT' = 'FLAT (HONG KONG) CO., LIMITED'; KB = KIBING."""
    a=str(a or '').upper().strip(); b=str(b or '').upper().strip()
    a=SUPPLIER_ALIAS.get(a,a); b=SUPPLIER_ALIAS.get(b,b)
    return bool(a and b) and 'UNASSIGNED' not in (a,b) and (a in b or b in a)
VP_UNKNOWN=('','?','UNASSIGNED','OTHER / UNASSIGNED'); VP_WSA={}; VP_MRP={}
def vendor_pools(dem, stock, arr, nd, vend):
    """Settings ALLOWED VENDORS on one projection - ONE rule for the WSA daily (planning sheet) and MRP_Daily (user rule, Nov 2026:
       "WSA output must follow settings, e.g. allowed vendors by customer ... all output workbooks must follow settings").
       dem   = {customer name or None: [demand per day]} (OEM left out; None = not split by customer = any vendor)
       stock = {supplier name: qty today}; arr = {supplier name: [arrivals per day]}; vend(customer) = its FIRST allowed vendor or None.
       Pools: one per vendor a customer with demand is tied to (that vendor's stock + arrivals), FREE = stock / arrivals of unknown
       vendor (Settings: 'always allowed'), ANY = other vendors (only customers without a vendor rule may use them).
       A tied customer may use its vendor + FREE; a customer without a rule may use everything. Stock-zero day = the first day the
       customers tied to a vendor need more (up to that day) than that vendor + FREE supplied (up to that day), or the pooled
       projection is below zero - the 'up to that day' rule of every projection here (no fixed order of use).
       Returns None when no customer with demand has a rule, else dict(lock=[per day], eff=[per day], pools={label: [own balance:
       supply - demand of that vendor's customers, end of day]}, who={label: [customers]}, first=first day the rule makes short or
       None). eff = projection that follows the rule (< 0 = short by that much); lock = pooled projection - eff = stock + arrivals
       no customer short that day may use (0 = the rule changes nothing that day). Sum of the pools = the pooled projection."""
    tied=sorted({vend(c) for c,v in dem.items() if c is not None and vend(c) and any(x>1e-9 for x in v)})
    if not tied: return None
    def pool_of(name):
        n=' '.join(str(name or '').split()).upper()
        if n in VP_UNKNOWN: return 'FREE'
        return next((V for V in tied if _vend_hit(V,n)),'ANY')
    sup={p:[0.0]*nd for p in tied+['FREE','ANY']}; need={p:[0.0]*nd for p in tied+['ANY']}
    for n,q in stock.items(): sup[pool_of(n)][0]+=float(q or 0.0)
    for n,v in arr.items():
        p=pool_of(n)
        for k in range(min(nd,len(v))): sup[p][k]+=float(v[k] or 0.0)
    who={p:[] for p in tied+['ANY']}
    for c,v in dem.items():
        p=(vend(c) if c is not None else None) or 'ANY'
        if c is not None and c not in who[p]: who[p].append(c)
        for k in range(min(nd,len(v))): need[p][k]+=float(v[k] or 0.0)
    def cum(x):
        o=[]; t_=0.0
        for v in x: t_+=v; o.append(t_)
        return o
    Vd={V:[a_-b_ for a_,b_ in zip(cum(need[V]),cum(sup[V]))] for V in tied}     # vendor V: demand - supply up to day k
    Ad=[a_-b_ for a_,b_ in zip(cum(need['ANY']),cum(sup['ANY']))]; Fc=cum(sup['FREE'])
    # Short on day K when the customers tied to vendor(s) need more (up to K) than those vendors + unknown-vendor supply (up to K), or
    # when everything together is short (= the pooled projection) - the same 'up to that day' rule as every projection of the script
    # (Hall's condition: customers without a rule may use any supply, so they never add a condition of their own).
    lock=[0.0]*nd; eff=[0.0]*nd
    for K in range(nd):
        S_=Fc[K]-sum(Vd[V][K] for V in tied)-Ad[K]
        viol=sum(max(0.0,Vd[V][K]) for V in tied)-Fc[K]
        eff[K]=min(S_,-viol) if viol>1e-6 else S_
        lock[K]=max(0.0,S_-eff[K])
    pools={**{V:[-x for x in Vd[V]] for V in tied},'any vendor':[-x for x in Ad],'unknown vendor':Fc}
    first=next((k for k in range(nd) if lock[k]>0.5),None)
    return dict(lock=lock,eff=eff,pools=pools,who={**{V:who[V] for V in tied},'any vendor':who['ANY'],'unknown vendor':[]},first=first,
                sup={lab:sum(x) for lab,x in (('any vendor' if p=='ANY' else 'unknown vendor' if p=='FREE' else p,v) for p,v in sup.items())})
def _mrp_vendor_pools(b, P, T, stk, rules, ships, d0, nd):
    """vendor_pools on the MRP_Daily projection of one material: demand per customer (b['cust_line'], OEM already out), stock by
       supplier (b['sup']: MB52 batches / count by supplier), in transit by supplier (b['sup'] transit, else the MIS shipments);
       what cannot be named goes to 'unknown vendor' (stock) / 'not split by customer' (demand), so the pools add up to S."""
    m=b['mat']; dem={}; nv={}
    for k_,v_ in (b.get('cust_line') or {}).items():
        cu=(k_[0] if isinstance(k_,tuple) else k_); nm=' '.join(str(cu or '').split())
        a_=dem.setdefault(nm,[0.0]*nd)
        for k in range(min(nd,len(v_))): a_[k]+=float(v_[k] or 0.0)
        nv[nm]=(rules.get((_norm(nm),m)) or [None])[0]
    r_=[P[k]-sum(v[k] for v in dem.values()) for k in range(nd)]
    if any(abs(x)>1e-6 for x in r_): dem[None]=r_
    st={}; av={}
    if b.get('sup'):
        for sp in b['sup']:
            n=str(sp.get('name') or ''); st[n]=st.get(n,0.0)+float(sp.get('stock') or 0.0)
            av.setdefault(n,[0.0]*nd); av[n]=[av[n][k]+float((sp.get('transit') or [0.0]*nd)[k] if k<len(sp.get('transit') or []) else 0.0) for k in range(nd)]
    else:
        for s_ in ships.get(m,[]):
            if s_.get('date') is None: continue
            k=(s_['date']-d0).days
            if 0<=k<nd: av.setdefault(str(s_.get('sup') or ''),[0.0]*nd)[k]+=float(s_['qty'])
    rs=stk-sum(st.values())
    if abs(rs)>1e-6: st['']=st.get('',0.0)+rs
    ra=[T[k]-sum(v[k] for v in av.values()) for k in range(nd)]
    if any(abs(x)>1e-6 for x in ra): av['']=[av.get('',[0.0]*nd)[k]+ra[k] for k in range(nd)]
    return vendor_pools(dem,st,av,nd,lambda c: nv.get(c))
def _wsa_sheet_series(m, d0, nd):
    """Per WSA planning sheet of material m (step 2, aligned file): {sheet: (stock today, end-of-day projection)} = stock + PPC arrivals
       (column T) - Demand, minus the sheet's own vendor lock (column AK) - the build_alert_rows rule for a material on 2+ sheets."""
    dts=[(d0+pd.Timedelta(days=k)).to_pydatetime() for k in range(nd)]; out={}
    for g in WSA_REFS[0] or []:
        if m not in (g.get('codes') or []) or not g.get('dem_by_date'): continue
        bb=g.get('bal_by_date') or {}; st=bb.get(dts[0])
        if st is None: st=bb.get(g['asof'])
        st=float(st or 0.0); cur=st; S=[]
        for d in dts:
            cur=cur-float((g.get('dem_by_date') or {}).get(d) or 0.0)+float((g.get('ppcit_by_date') or {}).get(d) or 0.0)
            S.append(cur-float((g.get('vlock_by_date') or {}).get(d) or 0.0))
        out[g['sheet']]=(st,S)
    return out
def _vp_from_sheets(S, sheets):
    """vendor_pools-like result of a material split over 2+ WSA sheets: short when any sheet is short (its customers may not use the
       other sheet's suppliers) - eff = the short sheets together (never above the pooled projection S); lock = S - eff."""
    nd=len(S); eff=[0.0]*nd; lock=[0.0]*nd
    for k in range(nd):
        v=[sv[1][k] for sv in sheets.values() if k<len(sv[1])]
        neg=sum(x for x in v if x<0)
        eff[k]=min(S[k],neg) if neg<-1e-6 else S[k]; lock[k]=max(0.0,S[k]-eff[k])
    return dict(lock=lock,eff=eff,first=next((k for k in range(nd) if lock[k]>0.5),None),sheets=sheets,pools={sh:sv[1] for sh,sv in sheets.items()},who={},sup={})
def vp_console(VP, d0):
    """One console text for vendor_pools results {name: vp}: the rule and whether it moved a stock-zero day."""
    out=[]
    for nm,vp in VP.items():
        lk=vp['first']
        rt_=('split by WSA sheet as the WSA daily: '+', '.join(f"'{sh}' stock {sv[0]:,.0f}" for sh,sv in vp['sheets'].items())) if vp.get('sheets') else vp_rules_txt(vp)
        out.append(f"{nm} [{rt_}] -> "+(f"the rule makes {d0+pd.Timedelta(days=lk):%d-%b} a stock-zero day (the allowed vendors cannot cover every customer; pooled stock still positive)" if lk is not None else 'no stock-zero day changes'))
    return '; '.join(out)
def vp_rules_txt(vp):
    """Short text of a vendor_pools result: who uses which vendor (Settings) and the supply of each."""
    t=[]
    for lab,cs in vp['who'].items():
        if lab=='unknown vendor' and not vp['sup'].get(lab): continue
        if lab=='any vendor' and not cs and not vp['sup'].get(lab): continue
        t.append(f"{lab}: {', '.join(cs) or ('every customer' if lab=='unknown vendor' else 'no customer')} (stock + arrivals {vp['sup'].get(lab,0):,.0f})")
    return '; '.join(t)
def _wsa_shared(blocks, AD, label_of, since=None):
    """A material on 2+ planning sheets (back glass 333: 'Glass G12R Topcon SB Back' + 'Glass G12R Topcon Subanci Back').
       User rule, Nov 2026: 'you have the demand of each customer - follow the MRP daily settings and definitions to calculate the
       WSA daily, so you know which supplier goes with which customer':
       - customers: each (customer, line) of the material goes to ONE sheet - the sheet with numbers in its customer column on the
         rows from the 1st of the month (`since`; the planner's sheets keep old customers on 2025 rows), else the sheet whose
         customers have the same first allowed vendor (Settings ALLOWED VENDORS), else the first sheet; demand not split by
         customer stays on the first sheet;
       - stock by supplier, as the MRP_Daily supplier rows (MB52 batches / count by supplier): a supplier's stock goes to the sheet
         whose customers have it as FIRST allowed vendor (2+ sheets: by their demand); a supplier no customer is tied to and
         'Unassigned' stock go to the sheets with 'any vendor' customers (by demand), else the first sheet;
       - PPC In Transit arrivals: the sheet where you typed the PO + BL# when its customers may use that supplier (tied to it, or
         any vendor), else by supplier like the stock.
       Returns {material: dict(sheets, first, of={key: block}, stock={block: qty}, by_v={block: {vendor: qty}},
       cust={block: {vendor or '': [customers]}}, arr={block: [qty per day]}, sup={block: {day: {supplier names}}})}."""
    where={}
    for j,B in enumerate(blocks):
        for m in B['codes']:
            if m in AD['blocks'] and j not in where.setdefault(m,[]): where[m].append(j)
    shared={m:js for m,js in where.items() if len(js)>1}
    if not shared: return {}
    rules=_vendor_rules(AD); nd=len(AD['dates']); dix={d:i for i,d in enumerate(AD['dates'])}; PPC=AD.get('ppc') or {}
    def labs_of(B):
        out={}
        for d,r in B['dates'].items():
            for c,lab_ in B['cust']:
                if _num(_gcell(B['rv'],r,c)): out.setdefault(since is None or d>=since,set()).add(_norm(lab_))
        return out.get(True) or out.get(False) or set()          # the current customers; a sheet with none from `since`: all rows
    labs={j:labs_of(blocks[j]) for js in shared.values() for j in js}
    res={}
    for m,js in shared.items():
        a=AD['blocks'][m]; first=js[0]
        vend=lambda key: (rules.get((key[0],m)) or [None])[0]
        dem={key:sum(v) for key,v in a['cl'].items() if key not in a['oem']}
        of={}
        for key in a['cl']:
            mo=AD['mods'].get(key); dn=a['disp'].get(key,key)
            mis_=mo['mis'] if mo else dn[0]; lab=label_of(mis_,mo['line'] if mo else dn[1]) or mis_
            hit=[j for j in js if _norm(lab) in labs[j]]
            if len(hit)==1: of[key]=hit[0]
        for key in a['cl']:
            if key in of: continue
            v=vend(key); hit=sorted({of[k] for k in of if v and vend(k)==v})
            of[key]=hit[0] if len(hit)==1 else first
        claim={}; anyw={}; cust={j:{} for j in js}
        for key,j in of.items():
            v=vend(key); w_=dem.get(key,0.0)+1e-9
            if key not in a['oem']:
                if v: claim.setdefault(v,{})[j]=claim.get(v,{}).get(j,0.0)+w_
                else: anyw[j]=anyw.get(j,0.0)+w_
            nm_=(AD['mods'].get(key) or {}).get('mis') or a['disp'].get(key,key)[0]
            lst=cust[j].setdefault(v or '',[])
            if nm_ and nm_ not in lst: lst.append(nm_)
        def tied(name):
            out={}
            for v,wd in claim.items():
                if _vend_hit(v,name):
                    for j,w_ in wd.items(): out[j]=out.get(j,0.0)+w_
            return out
        st={j:0.0 for j in js}; by_v={j:{} for j in js}
        def put(q,name,wts):
            s_=sum(wts.values())
            for j,w_ in wts.items(): st[j]+=q*w_/s_; by_v[j][name]=by_v[j].get(name,0.0)+q*w_/s_
        named=[(n,q) for n,q in (a.get('sup') or []) if n and n.upper()!='UNASSIGNED' and abs(q)>1e-9]
        tn=sum(q for _,q in named); stock=a['stock']
        if tn>stock+0.5 and tn>0: named=[(n,q*stock/tn) for n,q in named]; tn=stock     # JW at the OEM left out / count: the MRP stock is the total
        for n,q in named: put(q,n,tied(n) or anyw or {first:1.0})
        if abs(stock-tn)>1e-6: put(stock-tn,'other / unassigned',anyw or {first:1.0})
        arr={j:[0.0]*nd for j in js}; sup={j:{} for j in js}; arr_v={j:{} for j in js}
        for x in PPC.get('rows') or []:
            if x.get('mat')!=m: continue
            i=dix.get(x.get('arrive'))
            if i is None: continue
            sn=str(x.get('supplier') or ''); t_=tied(sn); ws_=set(x.get('wsa_sheets') or [])
            tr=[j for j in js if blocks[j]['label'] in ws_ or blocks[j].get('src_label') in ws_]
            ok=[j for j in tr if j in t_ or j in anyw]
            j=ok[0] if ok else (max(t_,key=t_.get) if t_ else (max(anyw,key=anyw.get) if anyw else (tr[0] if tr else first)))
            arr[j][i]+=x['qty_mrp']; sup[j].setdefault(i,set()).add(sn or '?')
            arr_v[j].setdefault(sn,[0.0]*nd)[i]+=x['qty_mrp']
        res[m]=dict(sheets=js,first=first,of=of,stock=st,by_v=by_v,cust=cust,arr=arr,sup=sup,arr_v=arr_v)
    return res
def _shared_txt(S_, jb, blocks):
    """Row 2 note of a sheet that shares its material: whose stock this is (Settings ALLOWED VENDORS) and what the other sheets hold."""
    fmt=lambda d: ', '.join(f'{n} {q:,.0f}' for n,q in sorted(d.items(),key=lambda t:-t[1]) if abs(q)>0.5) or '0'
    who='; '.join(f"{', '.join(v)} -> {k or 'any vendor'}" for k,v in sorted(S_['cust'][jb].items(),key=lambda t:(t[0]=='',t[0])))
    oth='; '.join(f"'{_out_sheet_name(blocks[j])}' {fmt(S_['by_v'][j])}" for j in S_['sheets'] if j!=jb)
    return f"stock by supplier (Settings ALLOWED VENDORS, as MRP_Daily): here {fmt(S_['by_v'][jb])} for {who or 'no customer'}; other sheet: {oth}"

WSA_OLD_DIR='old WSA files'; WSA_BACKUP=[None]; WSA_MOVED=[]
WSA_DAMAGED=[None]  # (file, safety copy) when an in-place write failed half-way AND could not be undone - repeated at the end of the run
WSA_LOCKED=[None]   # set when your WSA file could not be replaced (open in Excel / OneDrive): the WSA daily went to WSA_LOCKED[0]
def _put_file(tmp, out):
    """Put tmp in place of out when Windows may lock out (OneDrive sync / Excel). Returns None when done, else the last error (tmp kept).
       1) replace, 5 tries over ~15 s (OneDrive holds a file a few seconds); 2) write the content INTO out (a rename-only lock - OneDrive -
       still allows it; Excel's lock refuses the open, so nothing is touched). Before writing in place, out is copied to '<tmp>.keep'; a
       write that fails half-way is undone from it, and if even that fails the console names the copy (it is kept). Removing tmp after a
       good save is cleanup only - its failure never undoes the save (checker findings Nov 2026)."""
    import shutil
    last=None
    for w_ in (0,1,2,4,8):
        if w_: time.sleep(w_)
        try: os.replace(tmp,out); return None
        except PermissionError as e_: last=e_
    if not os.path.exists(out): return last
    try: f_=open(out,'r+b')                                           # refused while Excel has it open -> the file is not touched
    except OSError as e_: return e_
    keep=tmp+'.keep'
    try: shutil.copyfile(out,keep)
    except OSError as e_: f_.close(); return e_                       # no safety copy -> no in-place write
    try:
        with f_, open(tmp,'rb') as s_:
            shutil.copyfileobj(s_,f_); f_.truncate()
    except Exception as e_:                                           # half written: put the file back as it was
        try: shutil.copyfile(keep,out); os.remove(keep)
        except OSError as e2_:
            WSA_DAMAGED[0]=(out,keep)
            print('#'*100+f"\nWARNING: {os.path.basename(out)} WAS HALF WRITTEN AND COULD NOT BE PUT BACK ({e2_}).\n"
                  f"  Your file as it was before this run: {keep} - close Excel / pause OneDrive and copy it over {os.path.basename(out)}.\n"+'#'*100,flush=True)
        return e_
    for f2_ in (tmp,keep):                                            # cleanup only
        try: os.remove(f2_)
        except OSError as e_: _swallow(0,e_)
    return None
def _wsa_save(tmp, out, stem, d0):
    """Save the rebuilt WSA file as today's WSA_Daily (your WSA file is never written - user answer '1', 11-Nov-2026). History: (user, Nov 2026: "I don't see the WSA daily anymore" - the file is on OneDrive and Windows said
       'Access is denied' when the rename replaced it, so no WSA daily was made and step 2 ran on the old layout). _put_file (retry, write
       in place); when the file stays locked (open in Excel): save as '<stem> (aligned <date>).xlsx' next to it - today's WSA daily exists,
       the next run with the file closed reads the newest of the two and puts the result back in '<stem>.xlsx' (the aligned one goes to
       deleted by _wsa_tidy). Returns the path written."""
    WSA_LOCKED[0]=None
    last=_put_file(tmp,out)
    if last is None: return out
    alt=os.path.join(os.path.dirname(os.path.abspath(out)),f'{stem} (2).xlsx')     # today's WSA daily is open in Excel: a second name
    os.replace(tmp,alt)                                               # raises when this fails too (then nothing was saved)
    WSA_LOCKED[0]=(alt,f'{type(last).__name__}: {last}'); return alt
def _same_file(a, b):
    import filecmp
    try: return os.path.getsize(a)==os.path.getsize(b) and filecmp.cmp(a,b,shallow=False)
    except OSError: return False
def _wsa_backup(out, ran_on, d0):
    """No backup copies any more (user, Nov 2026: "old WSA files folder? I don't want this, too complicated"): the WSA file lives on
       OneDrive, whose version history keeps the earlier versions; the save itself is safe (atomic replace, or an in-place write with a
       temporary safety copy - _put_file). Kept as a stub so callers stay unchanged."""
    return None
def _wsa_tidy(out, stem, ran_on=None, ran_mtime=None):
    """Your WSA file is never changed (user, Nov 2026, answer '1'); the WSA daily is 'WSA_Daily_<date>.xlsx' next to it. Like the older
       MRP_Daily files, WSA_Daily files of EARLIER dates are deleted (today's file and its '(2)' copy stay). Old '<your file> (aligned
       <date>).xlsx' files of earlier versions are only named in a NOTE (never read when your own file is there). The 'old WSA files'
       folder of earlier versions is removed - only the names the script made there (stem = your file's name)."""
    d_=os.path.dirname(os.path.abspath(out)); gone=[]
    today_=f'{dt.date.today():%Y-%m-%d}'
    for n_ in sorted(os.listdir(d_)):
        p_=os.path.join(d_,n_); m_=re.fullmatch(r'WSA_Daily_(\d{4}-\d{2}-\d{2})( \(2\))?\.xlsx',n_)
        if m_ and m_.group(1)<today_ and p_!=os.path.abspath(out) and not globals().get('KEEP_OLD'):   # --keep-old keeps them, like MRP_Daily
            try: os.remove(p_); gone.append(n_)
            except OSError as e_: print(f'NOTE: {n_} could not be deleted ({e_}) - close it in Excel; it can be deleted by hand')
        elif n_.lower().endswith('.xlsx') and _ALIGN_SFX.search(os.path.splitext(n_)[0]) and _ALIGN_SFX.sub('',os.path.splitext(n_)[0])==stem and p_!=(os.path.abspath(ran_on) if ran_on else None):
            print(f"NOTE: {n_} = an older WSA output (not read, not changed) - the WSA daily is now WSA_Daily_<date>.xlsx; delete it when you no longer need it")
        elif n_==f'{stem}.xlsx.tmp.keep':                             # safety copy of an earlier version's in-place write
            try: os.remove(p_); gone.append(n_)
            except OSError as e_: _swallow(0,e_)
    od=os.path.join(d_,WSA_OLD_DIR)
    if os.path.isdir(od):
        mine_=re.compile(re.escape(stem)+r'( \(aligned \d{4}-\d{2}-\d{2}\))?( \(before \d{4}-\d{2}-\d{2} \d{6}\))?\.xlsx',re.I)
        for n_ in os.listdir(od):
            if mine_.fullmatch(n_) and ('(before ' in n_ or '(aligned ' in n_):
                try: os.remove(os.path.join(od,n_))
                except OSError as e_: _swallow(0,e_)
        try: os.rmdir(od); gone.append(f"folder '{WSA_OLD_DIR}'")
        except OSError: print(f"NOTE: folder '{WSA_OLD_DIR}' kept - it holds files the script did not make")
    return gone
def align_wsa(src, AD, today=None, name_from=None):
    """Write the rebuilt WSA file as 'WSA_Daily_<date>.xlsx' next to your WSA file (yours is only read). Returns (path, report rows)."""
    ALIGN_STAT['ppc_sheets']=[0,0]
    import zipfile, posixpath
    from collections import defaultdict
    from lxml import etree
    from openpyxl.formula.translate import Translator
    M_='http://schemas.openxmlformats.org/spreadsheetml/2006/main'; M='{%s}'%M_
    RNS='http://schemas.openxmlformats.org/officeDocument/2006/relationships'; PNS='http://schemas.openxmlformats.org/package/2006/relationships'
    CTN='http://schemas.openxmlformats.org/package/2006/content-types'
    from openpyxl.utils import column_index_from_string as _CI
    def CX_order(s__):
        m_=re.match(r'<c r="([A-Z]+)',s__); return _CI(m_.group(1)) if m_ else 0
    _ug=lambda u: next((g for g,ws in (('pc',('PC','PCS','NOS','SET','EA')),('m2',('M2','SQM','SQ')),('roll',('ROL','ROLL'))) if any(str(u).upper().startswith(w) for w in ws)),str(u).upper()[:2])
    d0=pd.Timestamp(today or base_d0()).normalize(); keep_from=d0.replace(day=1)
    dix={d:i for i,d in enumerate(AD['dates'])}; nd=len(AD['dates'])
    serial=lambda d: (pd.Timestamp(d).normalize()-pd.Timestamp('1899-12-30')).days
    nf_=name_from or src; stem=_ALIGN_SFX.sub('',os.path.splitext(os.path.basename(nf_))[0])
    own_stem=stem; stem=f'WSA_Daily_{dt.date.today():%Y-%m-%d}'               # user, Nov 2026 (answer '1'): YOUR WSA FILE IS NEVER CHANGED -
    out=os.path.join(os.path.dirname(os.path.abspath(nf_)),f'{stem}.xlsx')     # the WSA daily is a new file each day (was: saved over your file, option 2):
    wv=load_workbook(src,data_only=True,read_only=True); wf=load_workbook(src,read_only=True)
    blocks,GS=_glass_split(_wsa_read(wv,wf),AD)
    look,lookcu=_wsa_lookup(src)
    lab_cache={}
    def label_of(mis,line):
        k=(_norm(mis),line)
        if k in lab_cache: return lab_cache[k]
        lab=next((look[(p,_norm(mis))] for p in _plant_keys(line) if (p,_norm(mis)) in look),None)
        if lab is None and _norm(mis) in lookcu: lab=lookcu[_norm(mis)].most_common(1)[0][0]
        lab_cache[k]=lab; return lab
    T=ALIGN_SHEET; q=lambda s_: str(s_).replace('"','""')
    z=zipfile.ZipFile(src); names_=set(z.namelist())
    wbx=etree.fromstring(z.read('xl/workbook.xml')); rels=etree.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    tg={r.get('Id'):r.get('Target') for r in rels}
    shs=wbx.find(M+'sheets'); sheet_el={s.get('name'):s for s in shs}
    part={s.get('name'):('xl/'+tg[s.get('{%s}id'%RNS)].lstrip('/').replace('xl/','',1)) for s in shs}
    for o_,n_ in GS['renames'].items():          # front + back glass sheet: the front glass keeps the sheet, under its new name
        if o_ in sheet_el and n_ not in sheet_el: sheet_el[o_].set('name',n_); sheet_el[n_]=sheet_el[o_]; part[n_]=part[o_]
    STY=_Styles(z.read('xl/styles.xml')); yellow=STY.yellow_xfs()
    report=[]; skipped=[]; unmatched=Counter(); misonly=Counter(); other=Counter(); unit_notes=[]; dropped=set(); arch_new=[]
    # ---------- 1) the plan table rows of every block (module plan x BOM) and the customer columns ----------
    work=[]; tab=[]; lab_tot=Counter(); lab_oem=set(); hist_lab={}
    SHR=_wsa_shared(blocks,AD,label_of,keep_from)   # a material on 2+ sheets: customers, stock and arrivals by Settings ALLOWED VENDORS
    VRULES=_vendor_rules(AD); VP_WSA.clear()
    for jb,B in enumerate(blocks):
        mats=[c for c in B['codes'] if c in AD['blocks']]
        mine=lambda m,key: m not in SHR or SHR[m]['of'].get(key,SHR[m]['first'])==jb
        if not mats: skipped.append(f"{B['label']} (codes {', '.join(B['codes']) or 'none'} not in the MRP - sheet left as it is)"); continue
        U_=str(AD['blocks'][mats[0]].get('uom') or '').upper()
        if B['uom'] and U_ and _ug(B['uom'])!=_ug(U_) and not ({_ug(B['uom']),_ug(U_)}=={'pc','roll'}):
            unit_notes.append(f"{B['label']}: unit '{B['uom']}' -> '{U_}'")
        bid=B['label']; dropped.update(f"{B['sheet']}: {h}" for h in B['dropped'])
        stock=sum((SHR[m]['stock'][jb] if m in SHR else AD['blocks'][m]['stock'])
                  +(AD['blocks'][m]['jw'] if WSA_OEM_IN_BALANCE and (m not in SHR or SHR[m]['first']==jb) else 0.0) for m in mats)
        def want_m(m,i):                                      # MRP plan of the material; shared: its customers on this sheet (+ the unsplit rest on the first sheet)
            a_=AD['blocks'][m]; tot_=a_['plan'][i]+(a_['plan_oem'][i] if WSA_OEM_IN_BALANCE else 0.0)
            if m not in SHR: return tot_
            ks_=[k_ for k_ in a_['cl'] if WSA_OEM_IN_BALANCE or k_ not in a_['oem']]
            return (sum(a_['cl'][k_][i] for k_ in ks_ if mine(m,k_))
                    +((tot_-sum(a_['cl'][k_][i] for k_ in ks_)) if SHR[m]['first']==jb else 0.0))
        want=[sum(want_m(m,i) for m in mats) for i in range(nd)]
        oemk=set().union(*[AD['blocks'][m]['oem'] for m in mats])
        keys=sorted(set().union(*[{k_ for k_ in AD['blocks'][m]['cl'] if mine(m,k_)} for m in mats]),key=lambda k:(k in oemk,k))
        brow=[]; got=[0.0]*nd
        for key in keys:
            per={m:AD['blocks'][m]['cl'].get(key) for m in mats if mine(m,key) and any(AD['blocks'][m]['cl'].get(key) or [])}
            use=[sum(v[i] for v in per.values()) for i in range(nd)]
            if not any(u>1e-12 for u in use): continue
            is_oem=key in oemk; inbal='No' if (is_oem and not WSA_OEM_IN_BALANCE) else 'Yes'
            mo=AD['mods'].get(key)
            dn=next((AD['blocks'][m]['disp'][key] for m in per if key in AD['blocks'][m]['disp']),key)
            mis_d=mo['mis'] if mo else dn[0]; line_d=mo['line'] if mo else dn[1]
            lab=label_of(mis_d,line_d)
            if lab is None: unmatched[f"{mis_d} / {line_d}"]+=1; lab=mis_d          # not on the Lookup sheet: its own column
            if mo is None: misonly[f"{bid}: {mis_d} ({line_d or 'no line'})"]+=1
            srcs=sorted({('PPC BOM' if f'your BOM ({dn[0]})' in AD['blocks'][m]['basis'] else 'MIS BOM') for m in per})
            for i in range(nd):
                u=use[i]; d=AD['dates'][i]
                if u<=1e-12 or d<d0: continue
                m_=mo['v'][i] if mo else 0.0; ok=m_>1e-12
                brow.append(dict(date=d,serial=serial(d),line=line_d,cust=', '.join(mo['cust']) if mo else '',mis=mis_d,block=bid,col=lab,inbal=inbal,
                                 mods=m_ if ok else None,qty=(u/m_) if ok else None,req=u,mw=((mo.get('mw') or [0.0]*nd)[i] if ok else None),
                                 unit=U_ or B['uom'],codes='+'.join(m for m in per if per[m][i]>1e-12),src=(' + '.join(srcs) if ok else 'MIS plan (no modules)')))
                if inbal=='Yes': got[i]+=u
        for i in range(nd):                                   # anything the customer split does not explain (keeps Demand = MRP plan)
            if abs(want[i]-got[i])>1e-6 and AD['dates'][i]>=d0:
                d=AD['dates'][i]; other[bid]+=1
                brow.append(dict(date=d,serial=serial(d),line='',cust='',mis='',block=bid,col='(not split by customer)',inbal='Yes',mods=None,qty=None,
                                 req=want[i]-got[i],mw=None,unit=U_ or B['uom'],codes='+'.join(mats),src='MRP plan'))
        for x in brow:
            lab_tot[x['col']]+=x['mods'] or 0.0
            if x['inbal']=='No': lab_oem.add(x['col'])
        for d,r in B['dates'].items():                        # customers with numbers on the rows kept before today
            if keep_from<=d<d0:
                for c,lab_ in B['cust']:
                    if _num(cv(B['rv'],r,c)): hist_lab.setdefault(_norm(lab_),lab_)
        ufix=bool(B['uom'] and U_ and _ug(B['uom'])!=_ug(U_) and not ({_ug(B['uom']),_ug(U_)}=={'pc','roll'}))   # your label is another unit: the numbers are in the MRP (SAP) unit
        kd_,km_=_kind_of(B['mname'] or ''),_kind_of(' '.join(str(AD['blocks'][m].get('common') or '') for m in mats))
        mfix=(AD['blocks'][mats[0]].get('desc') or AD['blocks'][mats[0]].get('common') or mats[0]) if (kd_ and km_ and kd_!=km_) else None   # e.g. 'Frame 72 cells' on the EVA sheet
        if mfix: unit_notes.append(f"{B['label']}: material '{str(B['mname']).strip()}' -> '{mfix}'")
        vp=None                                               # Settings ALLOWED VENDORS on this sheet (one material per sheet)
        if len(mats)==1 and any(k_[1]==mats[0] for k_ in VRULES):
            m=mats[0]; a_=AD['blocks'][m]; dem_={}; nvend={}
            for k_ in a_['cl']:
                if not mine(m,k_) or (k_ in a_['oem'] and not WSA_OEM_IN_BALANCE): continue
                mo=AD['mods'].get(k_); nm_=(mo['mis'] if mo else a_['disp'].get(k_,k_)[0]) or k_[0]
                v_=dem_.setdefault(nm_,[0.0]*nd)
                for i in range(nd): v_[i]+=a_['cl'][k_][i]
                nvend[nm_]=(VRULES.get((k_[0],m)) or VRULES.get((_norm(nm_),m)) or [None])[0]
            rest_=[want[i]-sum(v_[i] for v_ in dem_.values()) for i in range(nd)]
            if any(abs(x)>1e-6 for x in rest_): dem_[None]=rest_
            if m in SHR: st_={n_:q_ for n_,q_ in SHR[m]['by_v'][jb].items()}; av_={n_:list(v_) for n_,v_ in SHR[m]['arr_v'][jb].items()}
            else:
                named=[(n_,q_) for n_,q_ in (a_.get('sup') or []) if n_ and n_.upper()!='UNASSIGNED' and abs(q_)>1e-9]
                tn=sum(q_ for _,q_ in named)
                if tn>stock+0.5 and tn>0: named=[(n_,q_*stock/tn) for n_,q_ in named]; tn=stock
                st_={}
                for n_,q_ in named: st_[n_]=st_.get(n_,0.0)+q_
                if abs(stock-tn)>1e-6: st_['']=st_.get('',0.0)+stock-tn
                av_={}
                for x in (AD.get('ppc') or {}).get('rows') or []:
                    i=dix.get(x.get('arrive'))
                    if x.get('mat')!=m or i is None: continue
                    av_.setdefault(str(x.get('supplier') or ''),[0.0]*nd)[i]+=x['qty_mrp']
                pa_=((AD.get('ppc') or {}).get('arr') or {}).get(m) or [0.0]*nd
                r_=[pa_[i]-sum(v_[i] for v_ in av_.values()) for i in range(nd)]
                if any(abs(x)>1e-6 for x in r_): av_.setdefault('',[0.0]*nd); av_['']=[av_[''][i]+r_[i] for i in range(nd)]
            vp=vendor_pools(dem_,st_,av_,nd,lambda c: nvend.get(c))
            if vp: VP_WSA[bid]=vp
        work.append(dict(B=B,mats=mats,bid=bid,stock=stock,brow=brow,U=U_ or str(B['uom']).upper(),jb=jb,ufix=ufix,mfix=mfix,vp=vp,
                         shr=' | '.join(f"{m} {_shared_txt(SHR[m],jb,blocks)}" for m in mats if m in SHR)))
    seen=set(); slots=[]
    for lab in (sorted([l for l in lab_tot if l not in lab_oem and l!='(not split by customer)'],key=lambda l:(-lab_tot[l],l))
                +sorted([l for l in lab_tot if l in lab_oem],key=lambda l:(-lab_tot[l],l))+list(hist_lab.values())):
        if _norm(lab) not in seen: seen.add(_norm(lab)); slots.append(lab)
    while len(slots)<WSA_SLOTS: slots.append('')
    canon={_norm(s_):s_ for s_ in slots if s_}
    for w in work:
        for x in w['brow']: x['col']=canon.get(_norm(x['col']),x['col'])
        tab+=w['brow']
    cols=WSA_FIX1+[(f'c{i}',lab,11) for i,lab in enumerate(slots)]+WSA_FIX2
    CX={k:j+1 for j,(k,_,_) in enumerate(cols)}; LC=L(len(cols)); XL=lambda k: L(CX[k])
    end_d=max([AD['dates'][-1]]+[max(w['B']['dates']) for w in work if w['B']['dates']])
    days=[keep_from+pd.Timedelta(days=k) for k in range((end_d-keep_from).days+1)]
    R0=WSA_HDR+1; rowof=lambda d: R0+(d-keep_from).days; t=rowof(d0); last=rowof(end_d)
    # ---------- 2) styles ----------
    fb=STY.font(bold=True); f14=STY.font(bold=True,size=14); fw=STY.font(bold=True,rgb='FF1F3864'); bd=STY.border()
    yel=STY.fill('FFFFFF00')
    HDRF={'g':STY.fill('FFD8E4BC'),'b':STY.fill('FFB7DEE8'),'s':STY.fill('FFD9D9D9'),'d':STY.fill('FFEBF1DE'),'p':STY.fill('FFFCE4D6')}
    SX=dict(title=STY.xf(font=f14),label=STY.xf(font=fb),info=STY.xf(),note=STY.xf(font=fw),
            date=STY.xf(nf=STY.numfmt('dd-mmm-yy'),border=bd,h='center'),text=STY.xf(border=bd),day=STY.xf(border=bd,h='center'),
            mod=STY.xf(nf=3,border=bd),days=STY.xf(nf=STY.numfmt('0.0'),border=bd,h='right'),mw=STY.xf(nf=STY.numfmt('0.000'),border=bd),
            tmod=STY.xf(nf=3,font=fb),tmw=STY.xf(nf=STY.numfmt('0.000'),font=fb),
            t_date=STY.xf(nf=STY.numfmt('dd-mmm-yy')),t_mod=STY.xf(nf=3),t_qty=STY.xf(nf=STY.numfmt('#,##0.0#####')),
            t_req=STY.xf(nf=STY.numfmt('#,##0.###')),t_mw=STY.xf(nf=STY.numfmt('0.000')),t_head=STY.xf(font=fb,wrap=True,v='center'))
    for k_,fl in HDRF.items(): SX['h'+k_]=STY.xf(font=fb,fill=fl,border=bd,h='center',v='center',wrap=True)
    def qfmt(u): return '#,##0' if _ug(u) in ('pc','roll') or str(u).upper().startswith(('SET','NOS','PCS')) else '#,##0.0'
    for code_ in ('#,##0','#,##0.0'):
        nfid=3 if code_=='#,##0' else STY.numfmt(code_)
        SX['q'+code_]=STY.xf(nf=nfid,border=bd); SX['bal'+code_]=STY.xf(nf=nfid,border=bd)
        SX['typed'+code_]=STY.xf(nf=nfid,font=fb,fill=yel,border=bd); SX['tot'+code_]=STY.xf(nf=nfid,font=fb)
    DX_NEG=STY.dxf('FF9C0006','FFFFC7CE'); DX_TODAY=STY.dxf('FF000000','FFFFF2CC',bold=True)
    hgroup=lambda k: 'b' if k.startswith('c') and k[1:].isdigit() else ('g' if k in ('date','day','bal','plan','dem') else ('s' if k in ('actmod','act','mw','actmw','pdays','pport') else ('p' if k in ('ppcit','ppcsup','ppcpp') else 'd')))
    PPC=AD.get('ppc') or {}; PARR=PPC.get('arr') or {}
    # ---------- 3) one standard sheet per block ----------
    def keyrow(B,d): return B['dates'].get(d)
    def ival(B,role,r): return cv(B['rv'],r,B['role'][role]) if (r and role in B['role']) else None
    def ifx(B,role,r): return cv(B['rf'],r,B['role'][role]) if (r and role in B['role']) else None
    def bal_styles(B):
        """style index of the planner's Balance cells on the rows kept before today (to keep their yellow 'typed' look)"""
        want_={f"{L(B['role']['bal'])}{r}" for d,r in B['dates'].items() if keep_from<=d<d0}
        if not want_: return {}
        res={}
        for _,c in etree.iterparse(z.open(part[B['sheet']]),tag=M+'c'):
            if c.get('r') in want_: res[c.get('r')]=c.get('s')
            c.clear()
        return res
    old_arch=[]
    if ARCHIVE_SHEET in wv.sheetnames:
        for r in wv[ARCHIVE_SHEET].iter_rows(min_row=2,values_only=True):
            r=list(r)+[None]*len(ARCH_COLS)
            if r[0] and isinstance(r[1],(dt.datetime,dt.date)): old_arch.append(r[:len(ARCH_COLS)])
    old_arch=GS['arch'](old_arch)                 # records of a block split into front / back glass: one per glass
    akey=lambda r: (str(r[0]),pd.Timestamp(r[1]).normalize(),_num(r[3]),str(r[4] or ''),str(r[10] or ''))
    arch_have={akey(r) for r in old_arch}; arch_old_by={}
    for r in old_arch: arch_old_by.setdefault(str(r[0]),[]).append(r)
    plan_sheets=[]          # (output sheet name, xml, last row)
    out_name={}
    for w in work:
        B=w['B']; bid=w['bid']; brow=w['brow']; qf=qfmt(w['U']); SQ=SX['q'+qf]; STYP=SX['typed'+qf]; SBAL=SX['bal'+qf]; STOT=SX['tot'+qf]
        oname=_out_sheet_name(B)
        out_name[(B['sheet'],B['block'])]=oname
        S_lab=defaultdict(float); S_mod=defaultdict(float); S_req=defaultdict(float); S_mw=defaultdict(float)
        for x in brow:
            s_=x['serial']
            if x['mods'] is not None: S_lab[(s_,str(x['col']).lower())]+=x['mods']
            if x['inbal']=='Yes':
                if x['mods'] is not None: S_mod[s_]+=x['mods']; S_mw[s_]+=x['mw'] or 0.0
                S_req[s_]+=(x['mods']*x['qty'] if x['mods'] is not None and x['qty'] is not None else x['req'])
        colmap={B['role'][k]:XL(k) for k in B['role'] if k in CX}
        for j,(c,lab_) in enumerate(B['cust']):
            k_=next((f'c{i}' for i,s_ in enumerate(slots) if s_ and _norm(s_)==_norm(lab_)),None)
            if k_: colmap[c]=XL(k_)
        bst=bal_styles(B)
        for d,ri in sorted(B['dates'].items()):                   # deleted rows (before the 1st of the month) with a delivery record -> Archive
            if d>=keep_from: continue
            vals={k_:ival(B,k_,ri) for k_ in ('dlv','sch','po','ship','rem','book','ready','etd','bl','vessel','inv','ref')}
            dl_=_num(vals['dlv']) or 0; sch_=_num(vals['sch']) or 0
            if not (dl_>0 or sch_>0 or vals['po'] or vals['bl'] or vals['vessel']): continue
            rem_=str(vals['rem'] or '')
            state='Cancelled' if 'canc' in rem_.lower() else ('Delivered' if dl_>0 else ('Not delivered' if sch_>0 else ''))
            arch_new.append([bid,d.to_pydatetime()]+[vals[k_] for k_ in ('dlv','sch','po','ship','rem','book','ready','etd','bl','vessel','inv','ref')]+[state])
        carried=sum((_num(r_[3]) or 0) for r_ in (arch_old_by.get(bid,[])+[a_ for a_ in arch_new if a_[0]==bid and akey(a_) not in arch_have]) if r_[14]=='Not delivered')
        # old numbers for the console report (today and the next 7 days of the sheet as it was)
        r_t=keyrow(B,d0); old_bal=_num(ival(B,'bal',r_t)); old_d7=sum(_num(ival(B,'dem',keyrow(B,d0+pd.Timedelta(days=k)))) or 0 for k in range(7))
        old_neg=next((d for d in sorted(B['dates']) if d>=d0 and (_num(ival(B,'bal',B['dates'][d])) or 0)<0),None)
        rowsx=[]
        # rows 1-6: material info
        title=' + '.join(dict.fromkeys(str(AD['blocks'][m].get('common') or m) for m in w['mats']))
        r1=_xcell('A1',f'{bid}  -  {title}'+(f"  {B['note']}" if B.get('note') else ''),SX['title'])
        SCH=XL('sch'); ARC=f'SUMIFS({ARCHIVE_SHEET}!$D:$D,{ARCHIVE_SHEET}!$A:$A,"{q(bid)}",{ARCHIVE_SHEET}!$O:$O,"Not delivered")'
        r2=(_xcell('A2',f'Stock on {d0:%d-%b-%y}',SX['label'])+_xcell('C2',round(w['stock'],3),STOT,f'C{t}')+
            _xcell('E2','MB52 + PPC count'+('' if WSA_OEM_IN_BALANCE else ', JW at OEM excluded')+(f" - {w['shr']}" if w['shr'] else ''),SX['note'])+
            _xcell(f"{XL('actmod')}2",f'Open deliveries dated before {keep_from:%d-%b-%y} (Archive, State = Not delivered):',SX['note'])+
            _xcell(f'{SCH}2',carried,STOT,ARC))
        r2=''.join(sorted(re.findall(r'<c [^>]*?(?:/>|>.*?</c>)',r2),key=CX_order))
        r3=(_xcell('A3','Supplier & Material',SX['label'])+_xcell('C3',w['mfix'] or B['mname'] or title,SX['info'])+
            _xcell(f"{XL('ppcit')}3",(f"PPC In Transit = {PPC['src']} sheet {PPC.get('sheet') or 'In_Transit'}: Final Pending Qty. on Final AtA (late = today); material traced by your WSA PO + BL#, description, supplier, RM Code (Inbound sheet 'PPC In Transit')" if PARR else 'PPC In Transit: EMPTY - '+'; '.join(PPC.get('why') or ['no PPC file / In_Transit sheet'])),SX['note']))
        r3+=_xcell(f"{XL('vlock')}3",('Settings ALLOWED VENDORS: '+vp_rules_txt(w['vp'])+'. Customers without a vendor rule may use the other vendors\' surplus; a vendor\'s customers may use only that vendor (+ stock of unknown vendor). '
                                       f"Column V (Plant + Port Days, PPC In Transit) subtracts column {XL('vlock')} = stock + arrivals no customer short that day may use (0 = the rule changes nothing that day)."
                                       if w.get('vp') else 'Settings ALLOWED VENDORS: no vendor rule for the customers of this sheet (any vendor).'),SX['note'])
        u4_=w['U'] if w['ufix'] else (B['uom'] or w['U'])          # user, Nov 2026: keep M2 (SAP) - a wrong label ('Pcs' on EVA / EPE 420) shows the MRP unit
        r4=_xcell('A4','Unit of Measurement',SX['label'])+_xcell('C4',u4_,SX['info'])+(_xcell('E4',f"MRP unit: {w['U']}",SX['info']) if w['U'] and _ug(u4_)!=_ug(w['U']) else '')
        r5=_xcell('A5','Material code(s)',SX['label'])+''.join(_xcell(f'{L(3+i)}5',c_,SX['info']) for i,c_ in enumerate(B['codes']))
        rowsx+= [f'<row r="1" ht="21" customHeight="1">{r1}</row>',f'<row r="2">{r2}</row>',f'<row r="3">{r3}</row>',f'<row r="4">{r4}</row>',f'<row r="5">{r5}</row>']
        # rows 8..: one per day
        drows=[]; bal=None; prevr=None; cum=0.0; cum_p=0.0; tot=defaultdict(float); new_vals={}
        PA_=lambda m: (SHR[m]['arr'][w['jb']] if m in SHR else PARR.get(m)) or [0.0]*nd      # shared material: the arrivals of this sheet
        parr=[sum(PA_(m)[i] for m in w['mats']) for i in range(nd)]
        ALIGN_STAT['ppc_sheets'][1]+=1; ALIGN_STAT['ppc_sheets'][0]+=int(any(parr))
        PSUP=PPC.get('sup') or {}       # supplier names of the PPC In_Transit lines arriving that day
        PS_=lambda m: (SHR[m]['sup'][w['jb']] if m in SHR else PSUP.get(m)) or {}
        psup=[' / '.join(sorted(set().union(*[PS_(m).get(i,set()) for m in w['mats']]))) for i in range(nd)]
        for d in days:
            r=rowof(d); ri=keyrow(B,d); cells=[]; s_=serial(d)
            cells.append(_xcell(f'A{r}',d.to_pydatetime(),SX['date'])); cells.append(_xcell(f'B{r}',d.strftime('%a'),SX['day']))
            _s=_num(ival(B,'sch',ri)) or 0.0; _rm=str(ival(B,'rem',ri) or ''); _dl=ival(B,'dlv',ri)
            if 'canc' not in _rm.lower() and (_dl is None or _dl==''): cum+=_s     # = SUMPRODUCT(Scheduled, Remarks without 'canc', Delivered blank) from row 8
            if d<d0:                                                    # rows kept before today: the numbers as they were
                st_=bst.get(f"{L(B['role']['bal'])}{ri}") if ri else None
                cells.append(_xcell(f'C{r}',ival(B,'bal',ri),STYP if st_ in yellow else SBAL))
                for k_,sty_ in (('plan',SX['mod']),('dem',SQ)): cells.append(_xcell(f'{XL(k_)}{r}',ival(B,k_,ri),sty_))
                hv={}
                for c,lab_ in B['cust']:
                    k_=next((f'c{i}' for i,s2 in enumerate(slots) if s2 and _norm(s2)==_norm(lab_)),None)
                    v_=_num(cv(B['rv'],ri,c)) if (k_ and ri) else None
                    if v_ is not None: hv[k_]=(hv.get(k_) or 0)+v_          # blank stays blank
                for i,s2 in enumerate(slots): cells.append(_xcell(f"{XL(f'c{i}')}{r}",hv.get(f'c{i}'),SX['mod']))
                for k_,sty_ in (('actmod',SX['mod']),('act',SQ),('mw',SX['mw']),('actmw',SX['mw']),('pdays',SX['days']),('pport',SX['days']),('ppcit',SQ),('ppcsup',SX['text']),('ppcpp',SX['days']),('vlock',SQ),('vpool',SX['text'])):
                    cells.append(_xcell(f'{XL(k_)}{r}',ival(B,k_,ri),sty_))
            else:                                                       # today and later: live formulas on PPC_Plan_Req
                ID=q(bid); dref=f'$A{r}'
                if d==d0: bal=round(w['stock'],3); cells.append(_xcell(f'C{r}',bal,STYP))
                else:
                    bal=bal-new_vals[(prevr,'act')]+(_num(ival(B,'dlv',keyrow(B,d-pd.Timedelta(days=1)))) or 0.0)
                    cells.append(_xcell(f'C{r}',bal,SBAL,f"C{prevr}-{XL('act')}{prevr}+{XL('dlv')}{prevr}"))
                pv=S_mod.get(s_,0.0); dv=S_req.get(s_,0.0); mv=S_mw.get(s_,0.0)
                cells.append(_xcell(f"{XL('plan')}{r}",pv,SX['mod'],f'SUMIFS({T}!$H:$H,{T}!$E:$E,"{ID}",{T}!$G:$G,"Yes",{T}!$A:$A,{dref})'))
                cells.append(_xcell(f"{XL('dem')}{r}",dv,SQ,f'SUMIFS({T}!$J:$J,{T}!$E:$E,"{ID}",{T}!$G:$G,"Yes",{T}!$A:$A,{dref})'))
                for i,s2 in enumerate(slots):
                    a_=f"{XL(f'c{i}')}"
                    if s2: cells.append(_xcell(f'{a_}{r}',S_lab.get((s_,s2.lower()),0.0),SX['mod'],f'SUMIFS({T}!$H:$H,{T}!$E:$E,"{ID}",{T}!$F:$F,{a_}${WSA_HDR},{T}!$A:$A,{dref})'))
                    else: cells.append(_xcell(f'{a_}{r}',None,SX['mod']))
                    if s2: tot[f'c{i}']+=S_lab.get((s_,s2.lower()),0.0)
                cells.append(_xcell(f"{XL('actmod')}{r}",pv,SX['mod'],f"{XL('plan')}{r}"))
                cells.append(_xcell(f"{XL('act')}{r}",dv,SQ,f"{XL('dem')}{r}"))
                cells.append(_xcell(f"{XL('mw')}{r}",mv,SX['mw'],f'SUMIFS({T}!$K:$K,{T}!$E:$E,"{ID}",{T}!$G:$G,"Yes",{T}!$A:$A,{dref})'))
                cells.append(_xcell(f"{XL('actmw')}{r}",mv,SX['mw'],f"{XL('mw')}{r}"))
                cells.append(_xcell(f"{XL('pdays')}{r}",(bal/dv) if dv else 'na',SX['days'],f'IFERROR(C{r}/{XL("dem")}{r},"na")'))
                U1,X1,T1=XL('sch'),XL('rem'),XL('dlv')
                cells.append(_xcell(f"{XL('pport')}{r}",((bal+carried+cum)/dv) if dv else 'na',SX['days'],
                                    f'IFERROR((C{r}+${U1}$2+SUMPRODUCT({U1}${R0}:{U1}{r},--ISERROR(SEARCH("canc",{X1}${R0}:{X1}{r})),--({T1}${R0}:{T1}{r}="")))/{XL("dem")}{r},"na")'))
                ii=dix.get(d); av=parr[ii] if ii is not None else 0.0; cum_p+=av; P1=XL('ppcit')
                cells.append(_xcell(f'{P1}{r}',av if av else None,SQ))
                cells.append(_xcell(f"{XL('ppcsup')}{r}",(psup[ii] or None) if ii is not None else None,SX['text']))
                VP_=w.get('vp'); lk_=(VP_['lock'][ii] if (VP_ and ii is not None) else 0.0)
                if VP_:                                          # Settings ALLOWED VENDORS: column V leaves out what no short customer may use
                    VL=XL('vlock')
                    cells.append(_xcell(f"{XL('ppcpp')}{r}",((bal+cum_p-lk_)/dv) if dv else 'na',SX['days'],
                                        f'IFERROR((C{r}+SUM({P1}${t}:{P1}{r})-{VL}{r})/{XL("dem")}{r},"na")'))
                    cells.append(_xcell(f'{VL}{r}',round(lk_,3) if ii is not None else None,SQ))
                    if ii is not None:
                        cells.append(_xcell(f"{XL('vpool')}{r}",' | '.join(f"{lab_} {p_[ii]:,.0f}" for lab_,p_ in VP_['pools'].items()
                                                                        if VP_['who'].get(lab_) or VP_['sup'].get(lab_)),SX['text']))
                else:
                    cells.append(_xcell(f"{XL('ppcpp')}{r}",((bal+cum_p)/dv) if dv else 'na',SX['days'],
                                        f'IFERROR((C{r}+SUM({P1}${t}:{P1}{r}))/{XL("dem")}{r},"na")'))
                new_vals[(r,'act')]=dv; tot['plan']+=pv; tot['dem']+=dv; tot['mw']+=mv; tot['ppcit']+=av
            # delivery columns: as they are (the planner's constants and same-row formulas kept)
            for k_ in ('dlv','sch','po','ship','rem','book','ready','etd','bl','vessel','inv','ref','rem2'):
                v_=ival(B,k_,ri); f_=ifx(B,k_,ri); fx=None
                if isinstance(f_,str) and f_.startswith('='): fx=_carry_formula(f_,ri,colmap,r)
                sty_=SX['date'] if (isinstance(v_,(dt.datetime,dt.date)) or k_ in ('book','ready','etd')) else (SQ if k_ in ('dlv','sch') else SX['text'])
                cells.append(_xcell(f'{XL(k_)}{r}',v_,sty_,fx))
                if k_ in ('dlv','sch') and d>=d0: tot[k_]+=_num(v_) or 0.0
            sch_=_num(ival(B,'sch',ri)) or 0; rem_=str(ival(B,'rem',ri) or ''); dl_=_num(ival(B,'dlv',ri)) or 0; po_=ival(B,'po',ri); ve_=ival(B,'vessel',ri)
            st_v='' if not sch_>0 else ('Cancelled' if 'canc' in rem_.lower() else ('Delivered' if dl_>0 else ('PO Pending' if str(po_ or '').lower()=='new po' else ('Open PO' if ve_ in (None,'') else 'In transit'))))
            U1,X1,T1,V1,AC1=XL('sch'),XL('rem'),XL('dlv'),XL('po'),XL('vessel')
            cells.append(_xcell(f"{XL('status')}{r}",st_v,SX['text'],
                                f'IF({U1}{r}>0,IF(ISNUMBER(SEARCH("canc",{X1}{r})),"Cancelled",IF({T1}{r}>0,"Delivered",IF({V1}{r}="NEW PO","PO Pending",IF({AC1}{r}="","Open PO","In transit")))),"")'))
            cells=sorted((c_ for c_ in cells if c_),key=CX_order)
            drows.append(f'<row r="{r}">'+''.join(cells)+'</row>'); prevr=r
        # row 6 totals (today and later) and row 7 headers
        tc=[_xcell('A6','Total from today',SX['label'])]
        for k_ in ('plan','dem','mw','ppcit','dlv','sch')+tuple(f'c{i}' for i,s2 in enumerate(slots) if s2):
            sty_=SX['tmw'] if k_=='mw' else (SX['tmod'] if k_=='plan' or k_.startswith('c') else STOT)
            tc.append(_xcell(f'{XL(k_)}6',tot.get(k_,0.0),sty_,f'SUM({XL(k_)}{t}:{XL(k_)}{last})'))
        tc=sorted((c_ for c_ in tc if c_),key=CX_order)
        rowsx.append('<row r="6">'+''.join(tc)+'</row>')
        rowsx.append(f'<row r="{WSA_HDR}" ht="45" customHeight="1">'+''.join(_xcell(f'{L(j+1)}{WSA_HDR}',h_ if h_ else None,SX['h'+hgroup(k_)]) for j,(k_,h_,_) in enumerate(cols))+'</row>')
        rowsx+=drows
        # sheet view: same zoom / tab selection as before, frozen at the header and column C, opens at today's row
        old_sv={}
        try:
            for _,e in etree.iterparse(z.open(part[B['sheet']]),tag=M+'sheetView'):
                old_sv={k:e.get(k) for k in ('tabSelected','zoomScale','zoomScaleNormal')}; break
        except Exception as _e_: _swallow(8850,_e_); pass
        if B['block']>0: old_sv.pop('tabSelected',None)
        sva=' '.join(f'{k}="{v}"' for k,v in old_sv.items() if v)
        top=max(R0,t-3)
        TODAY_F=("TODAY()" if d0>=pd.Timestamp(dt.date.today()).normalize() else f"DATE({d0.year},{d0.month},{d0.day})")   # base-date run: the stock row is "today"
        widths=''.join(f'<col min="{j+1}" max="{j+1}" width="{wd}" customWidth="1"/>' for j,(_,_,wd) in enumerate(cols))
        xml=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
             f'<worksheet xmlns="{M_}" xmlns:r="{RNS}"><dimension ref="A1:{LC}{last}"/>'
             f'<sheetViews><sheetView {sva} workbookViewId="0"><pane xSplit="3" ySplit="{WSA_HDR}" topLeftCell="D{top}" activePane="bottomRight" state="frozen"/>'
             f'<selection pane="topRight"/><selection pane="bottomLeft"/><selection pane="bottomRight" activeCell="D{t}" sqref="D{t}"/></sheetView></sheetViews>'
             f'<sheetFormatPr defaultRowHeight="15"/><cols>{widths}</cols><sheetData>'+''.join(rowsx)+'</sheetData>'
             f'<autoFilter ref="A{WSA_HDR}:{LC}{last}"/>'
             f'<conditionalFormatting sqref="C{R0}:C{last} {XL("pdays")}{R0}:{XL("pport")}{last} {XL("ppcpp")}{R0}:{XL("ppcpp")}{last}"><cfRule type="cellIs" dxfId="{DX_NEG}" priority="1" operator="lessThan"><formula>0</formula></cfRule></conditionalFormatting>'
             f'<conditionalFormatting sqref="A{R0}:B{last}"><cfRule type="expression" dxfId="{DX_TODAY}" priority="2"><formula>$A{R0}={TODAY_F}</formula></cfRule></conditionalFormatting>'
             '<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/></worksheet>').encode('utf-8')
        plan_sheets.append((oname,xml,last,B))
        new_neg=None; b_=round(w['stock'],3)
        for d in days:
            if d<d0: continue
            if d>d0: b_=b_-S_req.get(serial(d-pd.Timedelta(days=1)),0.0)+(_num(ival(B,'dlv',keyrow(B,d-pd.Timedelta(days=1)))) or 0.0)
            if b_<0: new_neg=d; break
        report.append(dict(label=bid,sheet=oname,codes='+'.join(w['mats']),old_bal=old_bal,new_bal=w['stock'],old_d7=old_d7,
                           new_d7=sum(S_req.get(serial(d0+pd.Timedelta(days=k)),0.0) for k in range(7)),old_neg=old_neg,new_neg=new_neg))
    # ---------- 4) PPC_Plan_Req and Archive ----------
    tab.sort(key=lambda x:(x['serial'],str(x['block']),x['inbal']!='Yes',str(x['col']),str(x['line'])))
    def plan_xml(rows):
        hdr='<row r="1">'+''.join(_xcell(f'{L(j+1)}1',h_,SX['t_head']) for j,h_ in enumerate(ALIGN_COLS))+_xcell('P1',
            f"Built by update_inbound.py on {dt.date.today():%d-%b-%Y}"+(f" (base date {d0:%d-%b-%Y} = your MB52 date)" if d0<pd.Timestamp(dt.date.today()).normalize() else '')+": module plan (PPC Plan + Revised Plan) x quantity per module (PPC BOM, else MIS BOM). "
            "The planning sheets read this table with SUMIFS (customer columns, Planned modules, Demand, MW), so a change here updates them at once. "
            "Counted in Balance = No: OEM plan (made at the OEM from JW stock, not from our warehouse). Rebuilt from today on every run.",SX['info'])+'</row>'
        body=[]
        for i,x in enumerate(rows,2):
            c=[_xcell(f'A{i}',x['date'].to_pydatetime(),SX['t_date']),_xcell(f'B{i}',x['line']),_xcell(f'C{i}',x['cust']),_xcell(f'D{i}',x['mis']),
               _xcell(f'E{i}',x['block']),_xcell(f'F{i}',x['col']),_xcell(f'G{i}',x['inbal']),_xcell(f'H{i}',x['mods'],SX['t_mod']),_xcell(f'I{i}',x['qty'],SX['t_qty'])]
            if x['mods'] is not None and x['qty'] is not None: c.append(_xcell(f'J{i}',x['mods']*x['qty'],SX['t_req'],f'H{i}*I{i}'))
            else: c.append(_xcell(f'J{i}',x['req'],SX['t_req']))
            c+=[_xcell(f'K{i}',x['mw'],SX['t_mw']),_xcell(f'L{i}',x['unit']),_xcell(f'M{i}',x['codes']),_xcell(f'N{i}',x['src'])]
            body.append(f'<row r="{i}">'+''.join(c)+'</row>')
        n=len(rows)+1; widths=[11,16,26,24,26,26,10,10,13,13,9,7,30,16,3,90]
        return (n,('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<worksheet xmlns="{M_}" xmlns:r="{RNS}"><dimension ref="A1:P{n}"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
                '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
                '<cols>'+''.join(f'<col min="{j}" max="{j}" width="{w_}" customWidth="1"/>' for j,w_ in enumerate(widths,1))+'</cols><sheetData>'
                +hdr+''.join(body)+f'</sheetData><autoFilter ref="A1:N{n}"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>').encode('utf-8'))
    arch_add=[r for r in arch_new if akey(r) not in arch_have]
    arch=sorted(old_arch+arch_add,key=lambda r:(pd.Timestamp(r[1]),str(r[0])))
    def arch_xml(rows):
        hdr='<row r="1">'+''.join(_xcell(f'{L(j+1)}1',h_,SX['t_head']) for j,h_ in enumerate(ARCH_COLS))+_xcell('Q1',
            'Delivery records of the planning-sheet rows deleted by update_inbound.py (rows before the 1st of the month). Kept here so nothing is lost; not read by the script.',SX['info'])+'</row>'
        body=[]
        for i,r in enumerate(rows,2):
            c=[]
            for j,v in enumerate(r):
                sty_=SX['t_date'] if isinstance(v,(dt.datetime,dt.date)) else (SX['t_mod'] if j in (2,3) else 0)
                c.append(_xcell(f'{L(j+1)}{i}',v,sty_))
            body.append(f'<row r="{i}">'+''.join(c)+'</row>')
        n=len(rows)+1; widths=[26,12,11,12,13,16,22,12,11,11,18,26,13,13,13,3,90]
        return (n,('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<worksheet xmlns="{M_}" xmlns:r="{RNS}"><dimension ref="A1:Q{max(n,1)}"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
                '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
                '<cols>'+''.join(f'<col min="{j}" max="{j}" width="{w_}" customWidth="1"/>' for j,w_ in enumerate(widths,1))+'</cols><sheetData>'
                +hdr+''.join(body)+f'</sheetData><autoFilter ref="A1:O{n}"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>').encode('utf-8'))
    # ---------- 5) workbook: sheet list, parts, names ----------
    old_order=[s.get('name') for s in shs]; patched={}; drop=set(['xl/calcChain.xml'])
    ct=etree.fromstring(z.read('[Content_Types].xml'))
    gone=[]
    for nm_ in WSA_DROP_SHEETS:                    # Stock / Total_Plan_Req: not needed in the WSA daily file (user, Nov 2026)
        if nm_ not in sheet_el or nm_ not in part or sheet_el[nm_].getparent() is None: continue
        el_=sheet_el[nm_]; rid_=el_.get('{%s}id'%RNS); shs.remove(el_)
        for r_ in list(rels):
            if r_.get('Id')==rid_: rels.remove(r_)
        pth_=part[nm_]; drop.add(pth_); rp_=posixpath.join(posixpath.dirname(pth_),'_rels',posixpath.basename(pth_)+'.rels')
        if rp_ in names_:
            for rel in etree.fromstring(z.read(rp_)):
                if rel.get('TargetMode')!='External': drop.add(posixpath.normpath(posixpath.join(posixpath.dirname(pth_),rel.get('Target'))))
            drop.add(rp_)
        gone.append(nm_)
    def free_part():
        k_=1
        while f'xl/worksheets/sheet{k_}.xml' in names_ or f'xl/worksheets/sheet{k_}.xml' in patched: k_+=1
        return f'xl/worksheets/sheet{k_}.xml'
    def add_sheet(name, after_el=None):
        pth=free_part(); ids_=set(r.get('Id') for r in rels); n_=1
        while f'rId{n_}' in ids_: n_+=1
        etree.SubElement(rels,'{%s}Relationship'%PNS,Id=f'rId{n_}',Type=RNS+'/worksheet',Target=pth.replace('xl/','',1))
        e_=etree.Element(M+'sheet',name=name,sheetId=str(max(int(s.get('sheetId')) for s in shs)+1)); e_.set('{%s}id'%RNS,f'rId{n_}')
        if after_el is not None: after_el.addnext(e_)
        else: shs.append(e_)
        etree.SubElement(ct,'{%s}Override'%CTN,PartName='/'+pth,ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
        part[name]=pth; patched[pth]=b''; return e_
    redesigned=set()
    for oname,xml,last_,B in plan_sheets:
        if oname in part and oname==B['sheet']: patched[part[oname]]=xml; redesigned.add(oname)
        elif oname in part: patched[part[oname]]=xml; redesigned.add(oname)         # an earlier run's split sheet (input already redesigned)
        else:
            prev_el=sheet_el[B['sheet']]
            while prev_el.getnext() is not None and prev_el.getnext().get('name') in [p[0] for p in plan_sheets if p[3]['sheet']==B['sheet']]: prev_el=prev_el.getnext()
            add_sheet(oname,prev_el); patched[part[oname]]=xml; redesigned.add(oname)
    # replaced sheets lose their old comments / drawings / print settings (all on deleted rows)
    for oname in redesigned:
        rp=posixpath.join(posixpath.dirname(part[oname]),'_rels',posixpath.basename(part[oname])+'.rels')
        if rp in names_:
            for rel in etree.fromstring(z.read(rp)):
                if rel.get('TargetMode')=='External': continue
                tp=posixpath.normpath(posixpath.join(posixpath.dirname(part[oname]),rel.get('Target')))
                drop.add(tp)
            drop.add(rp)
    n_plan,pxml=plan_xml(tab)
    if T not in part: add_sheet(T)
    patched[part[T]]=pxml
    n_arch,axml=arch_xml(arch)
    if ARCHIVE_SHEET not in part: add_sheet(ARCHIVE_SHEET)
    patched[part[ARCHIVE_SHEET]]=axml
    for o in list(ct):
        pn=str(o.get('PartName') or '').lstrip('/')
        if pn in drop: ct.remove(o)
    new_order=[s.get('name') for s in shs]; idx_new={n:i for i,n in enumerate(new_order)}
    imap={i:idx_new[n] for i,n in enumerate(old_order) if n in idx_new}
    dns=wbx.find(M+'definedNames')
    if dns is None:
        dns=etree.Element(M+'definedNames'); cp_=wbx.find(M+'calcPr')
        if cp_ is not None: cp_.addprevious(dns)
        else: wbx.append(dns)
    red_old={i for i,n in enumerate(old_order) if n in redesigned or n in gone}
    for e in list(dns):
        ls=e.get('localSheetId')
        if ls is None: continue
        if int(ls) in red_old or e.get('name')=='_xlnm._FilterDatabase' and new_order[imap.get(int(ls),0)] in (T,ARCHIVE_SHEET): dns.remove(e); continue
        if int(ls) in imap: e.set('localSheetId',str(imap[int(ls)]))
    def qn(n): return "'"+n.replace("'","''")+"'"
    for oname,xml,last_,B in plan_sheets:
        etree.SubElement(dns,M+'definedName',name='_xlnm._FilterDatabase',localSheetId=str(idx_new[oname]),hidden='1').text=f"{qn(oname)}!$A${WSA_HDR}:${LC}${last_}"
    etree.SubElement(dns,M+'definedName',name='_xlnm._FilterDatabase',localSheetId=str(idx_new[T]),hidden='1').text=f"{T}!$A$1:$N${n_plan}"
    etree.SubElement(dns,M+'definedName',name='_xlnm._FilterDatabase',localSheetId=str(idx_new[ARCHIVE_SHEET]),hidden='1').text=f"{ARCHIVE_SHEET}!$A$1:$O${n_arch}"
    bv=wbx.find(f'{M}bookViews/{M}workbookView')
    if bv is not None:
        for a_ in ('activeTab','firstSheet'):
            if bv.get(a_) is not None: bv.set(a_,str(imap.get(int(bv.get(a_)),0)))
    cp_=wbx.find(M+'calcPr')
    if cp_ is None: cp_=etree.SubElement(wbx,M+'calcPr')
    cp_.set('fullCalcOnLoad','1')
    for r in list(rels):
        if str(r.get('Target') or '').endswith('calcChain.xml'): rels.remove(r)
    # ---------- 6) Dashboard: one lookup column for every planning sheet ----------
    dash_note=''
    if 'Dashboard' in part and part['Dashboard'] not in patched:
        try:
            droot=etree.fromstring(z.read(part['Dashboard'])); sd=droot.find(M+'sheetData'); masters={}
            for c in sd.iter(M+'c'):
                f=c.find(M+'f')
                if f is not None and f.get('t')=='shared' and f.get('ref') is not None: masters[f.get('si')]=(c.get('r'),f.text or '')
            for c in sd.iter(M+'c'):
                f=c.find(M+'f')
                if f is None or f.get('t')!='shared': continue
                mr,mt=masters.get(f.get('si'),(None,None))
                if mr is None: continue
                f.text=mt if c.get('r')==mr else Translator('='+mt,origin=mr).translate_formula(c.get('r'))[1:]
                for a_ in ('t','ref','si'):
                    if a_ in f.attrib: del f.attrib[a_]
            info_={}
            for w in work:
                B=w['B']
                if B.get('dash') is False: continue        # back glass of a split sheet: no Dashboard row of its own
                info_.setdefault(B['sheet'],[]).append((B['c0'],out_name[(B['sheet'],B['block'])]))
            for oname,xml,last_,B in plan_sheets: info_.setdefault(oname,[(1,oname)])
            sst_=[]; dash_fixed=[]
            if 'xl/sharedStrings.xml' in names_:
                sst_=[''.join(si.itertext()) for si in etree.fromstring(z.read('xl/sharedStrings.xml')).findall(M+'si')]
            pat=re.compile(r"('(?:[^']|'')+'|[A-Za-z0-9_.]+)!\$A:\$[A-Z]{1,3}")
            nfix=0; PP=CX['pport']
            for rowel in sd.iter(M+'row'):
                rr=int(rowel.get('r'))
                if rr<2: continue
                cells=rowel.findall(M+'c'); fx=[c.find(M+'f') for c in cells if c.find(M+'f') is not None and c.find(M+'f').text]
                hit={m.group(1).strip("'").replace("''","'") for f in fx for m in pat.finditer(f.text)}
                hit=[h_ for h_ in hit if h_ in info_]
                if not hit: continue
                acell=next((c for c in cells if c.get('r')==f'A{rr}'),None); av=None
                if acell is not None and acell.find(M+'v') is not None:
                    try: av=float(acell.find(M+'v').text)
                    except (TypeError,ValueError): av=None
                tgt={}
                for h_ in hit:                                   # the block the old column index pointed into (Potting B = column AW ...)
                    blks=sorted(info_[h_]); tgt[h_]=blks[0][1]
                    for c0_,nm_ in blks:
                        if av is not None and av>=c0_: tgt[h_]=nm_
                if len(set(tgt.values()))>1:                     # one row reading two sheets (EPE 420 GSM row read EVA 420 GSM in column D): use the sheet named in column B
                    bcell=next((c for c in cells if c.get('r')==f'B{rr}'),None); blab=''
                    if bcell is not None:
                        if bcell.get('t')=='s' and bcell.find(M+'v') is not None:
                            try: blab=sst_[int(bcell.find(M+'v').text)]
                            except (ValueError,IndexError): blab=''
                        elif bcell.get('t')=='inlineStr': blab=''.join(bcell.itertext())
                        elif bcell.find(M+'v') is not None: blab=bcell.find(M+'v').text or ''
                    own=next((v_ for v_ in tgt.values() if _norm(v_)==_norm(blab)),None)
                    if own:
                        for h_ in list(tgt):
                            if tgt[h_]!=own: dash_fixed.append(f"row {rr} '{blab}' read '{tgt[h_]}' -> '{own}'"); tgt[h_]=own
                def _sub(m):
                    h_=m.group(1).strip("'").replace("''","'")
                    return (qn(tgt[h_])+f'!$A:${LC}') if h_ in tgt else m.group(0)
                for f in fx: f.text=pat.sub(_sub,f.text)
                if acell is not None:
                    for ch in list(acell): acell.remove(ch)
                    if 't' in acell.attrib: del acell.attrib['t']
                    etree.SubElement(acell,M+'v').text=str(PP)
                nfix+=1
            # back-glass rows (user, Nov 2026: 'Dashboard rows for the back-glass sheets?' - 'Yes'): one row per '<x> Back' sheet that no
            # row reads yet = a copy of its front row (same formulas / format, reading the Back sheet), added below the table; a rerun on
            # the aligned file finds the rows and adds nothing
            import copy as _cp
            rows_=sorted(sd.findall(M+'row'),key=lambda e:int(e.get('r'))); read_={}; last_=1
            for rowel in rows_:
                rr=int(rowel.get('r')); used=False
                for c in rowel.findall(M+'c'):
                    f=c.find(M+'f')
                    if f is not None and f.text:
                        used=True
                        for m in pat.finditer(f.text): read_.setdefault(m.group(1).strip("'").replace("''","'"),rowel)
                    elif c.find(M+'v') is not None or c.find(M+'is') is not None: used=True
                if used and rr>1: last_=max(last_,rr)
            old_last=last_; back_add=[]
            for fname,rowel in sorted(read_.items(),key=lambda kv:int(kv[1].get('r'))):
                if not fname.endswith(' Front'): continue
                bk=[n_ for n_ in part if n_.endswith(' Back') and n_[:-5].startswith(fname[:-6])]
                if len(bk)!=1 or bk[0] in read_: continue
                bname=bk[0]; src_r=int(rowel.get('r')); last_+=1; nr=last_
                new=_cp.deepcopy(rowel); new.set('r',str(nr))
                for c in new.findall(M+'c'):
                    col_=re.match(r'[A-Z]+',c.get('r')).group(0); old_ref=c.get('r'); c.set('r',f'{col_}{nr}')
                    f=c.find(M+'f')
                    if f is not None and f.text:
                        for a_ in ('t','ref','si'):
                            if a_ in f.attrib: del f.attrib[a_]
                        t_=Translator('='+f.text,origin=old_ref).translate_formula(c.get('r'))[1:]
                        f.text=re.sub(r"('(?:[^']|'')+'|[A-Za-z0-9_.]+)!",lambda m:(qn(bname)+'!') if m.group(1).strip("'").replace("''","'")==fname else m.group(0),t_)
                        v_=c.find(M+'v')
                        if v_ is not None: c.remove(v_)                 # Excel calculates on open (fullCalcOnLoad)
                        if c.get('t') in ('str','e','n','b'): del c.attrib['t']
                    elif col_=='B':                                     # label = the old label + ' Back'
                        lab=''
                        if c.get('t')=='s' and c.find(M+'v') is not None:
                            try: lab=sst_[int(c.find(M+'v').text)]
                            except (ValueError,IndexError): lab=''
                        elif c.get('t')=='inlineStr': lab=''.join(c.itertext())
                        for ch in list(c): c.remove(ch)
                        c.set('t','inlineStr'); is_=etree.SubElement(c,M+'is'); etree.SubElement(is_,M+'t').text=(lab or fname[:-6])+' Back'
                pos_=next((i for i,e in enumerate(list(sd)) if int(e.get('r'))>=nr),None)
                if pos_ is not None and int(list(sd)[pos_].get('r'))==nr: sd.remove(list(sd)[pos_])   # an empty formatted row there
                pos_=next((i for i,e in enumerate(list(sd)) if int(e.get('r'))>nr),None)
                if pos_ is None: sd.append(new)
                else: sd.insert(pos_,new)
                back_add.append(f"row {nr} '{bname}' (copy of row {src_r})")
            if back_add:                                                # filter / conditional formats / size: cover the new rows
                def _ext(ref):
                    return ' '.join(re.sub(rf'^([A-Z]+\$?)(\d+):([A-Z]+\$?){old_last}$',lambda m:f"{m.group(1)}{m.group(2)}:{m.group(3)}{last_}",x_) for x_ in ref.split())
                for e in droot.iter(M+'autoFilter',M+'conditionalFormatting'):
                    k_='ref' if e.tag==M+'autoFilter' else 'sqref'
                    if e.get(k_): e.set(k_,_ext(e.get(k_)))
                for e in (wbx.find(M+'definedNames') if wbx.find(M+'definedNames') is not None else []):   # the hidden filter name follows the filter
                    if e.get('name')=='_xlnm._FilterDatabase' and re.match(r"'?Dashboard'?!",e.text or ''):
                        e.text=re.sub(rf'(\$[A-Z]+\$){old_last}$',lambda m:f"{m.group(1)}{last_}",e.text)
                dm=droot.find(M+'dimension')
                if dm is not None and dm.get('ref'):
                    m=re.match(r'([A-Z]+\d+):([A-Z]+)(\d+)$',dm.get('ref'))
                    if m and int(m.group(3))<last_: dm.set('ref',f"{m.group(1)}:{m.group(2)}{last_}")
            patched[part['Dashboard']]=etree.tostring(droot,xml_declaration=True,encoding='UTF-8',standalone=True)
            dash_note=f'Dashboard: {nfix} rows now read column {PP} (Plant + Port Days) of the new layout'+(('; fixed: '+'; '.join(dash_fixed)) if dash_fixed else '')+(('; back-glass rows added: '+'; '.join(back_add)) if back_add else '')
        except Exception as ex_: dash_note=f'NOTE: Dashboard not updated ({type(ex_).__name__}: {ex_})'
    if GS['renames'] or gone:
        own_=set(part[o_] for o_ in redesigned)|{part[T],part[ARCHIVE_SHEET]}; nref=0
        for nm_ in z.namelist():
            if not nm_.startswith(('xl/worksheets/sheet','xl/charts/chart')) or nm_ in drop or nm_ in own_: continue
            raw_=patched.get(nm_) or z.read(nm_)
            t_=raw_.decode('utf-8')
            if not any(k_ in t_ for k_ in list(GS['renames'])+gone): continue
            t2_=_ref_fix(t_,GS['renames'],gone)
            if t2_!=t_: patched[nm_]=t2_.encode('utf-8'); nref+=1
        for e in dns:
            if e.text: e.text=_ref_fix(e.text,GS['renames'],gone)
    patched['xl/styles.xml']=STY.bytes()
    patched['xl/workbook.xml']=etree.tostring(wbx,xml_declaration=True,encoding='UTF-8',standalone=True)
    patched['xl/_rels/workbook.xml.rels']=etree.tostring(rels,xml_declaration=True,encoding='UTF-8',standalone=True)
    patched['[Content_Types].xml']=etree.tostring(ct,xml_declaration=True,encoding='UTF-8',standalone=True)
    tmp=out+'.tmp'
    with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zo:
        seen_=set()
        for info in z.infolist():
            if info.filename in drop: continue
            data=patched.get(info.filename); seen_.add(info.filename)
            zo.writestr(info,data if data is not None else z.read(info.filename),compress_type=zipfile.ZIP_DEFLATED)
        for nm_,data in patched.items():
            if nm_ not in seen_: zo.writestr(nm_,data,compress_type=zipfile.ZIP_DEFLATED)
    z.close(); wv.close(); wf.close()
    try:                                                              # never replace anything with a broken file: every part must read back
        with zipfile.ZipFile(tmp) as zc:
            bad_=zc.testzip()
            if bad_: raise ValueError(f'damaged part {bad_}')
            for nm_ in zc.namelist():
                if nm_.endswith(('.xml','.rels')): etree.fromstring(zc.read(nm_))
    except Exception as ex_:
        try: os.remove(tmp)
        except OSError: pass
        raise RuntimeError(f'the aligned file did not pass its own check ({type(ex_).__name__}: {ex_}) - nothing was replaced')
    try:
        ran_mt_=os.path.getmtime(nf_) if os.path.exists(nf_) else None   # the file this run read: second files newer than it are kept (_wsa_tidy)
        WSA_BACKUP[0]=_wsa_backup(out,nf_,d0)                        # no backup folder (user, Nov 2026); OneDrive version history keeps old versions
        out=_wsa_save(tmp,out,stem,d0)                                # OneDrive / Excel lock: retry, write in place, else '<stem> (aligned <date>).xlsx'
    except Exception as ex_:                                          # every way failed: your file stays as it is, the run goes on
        try: os.remove(tmp)
        except OSError: pass
        raise RuntimeError(f"could not save {os.path.basename(out)} ({type(ex_).__name__}: {ex_}) - is it open in Excel? Close it and run again; "
                           "your WSA file was NOT changed") from ex_
    WSA_MOVED[:]=_wsa_tidy(out,own_stem,nf_)           # WSA_Daily files of earlier dates: deleted
    # ---------- console ----------
    oct_open=[r for r in arch_add if r[-1]=='Not delivered' and pd.Timestamp(r[1])>=keep_from-pd.Timedelta(days=31)]
    print(f"WSA aligned: {os.path.basename(out)} - {len(plan_sheets)} planning sheets in one layout (header row {WSA_HDR}, columns A to {LC}), rows {keep_from:%d-%b-%y} to {end_d:%d-%b-%y}")
    print(f"  Customer columns (same on every sheet): {', '.join(s_ for s_ in slots if s_)}")
    if PPC.get('src'):
        rows_=PPC.get('rows') or []; late_=sum(1 for x in rows_ if x['late']); notm_=sorted({x['mat'] for x in rows_ if not x['in_mrp']})
        print(f"  PPC In Transit ({PPC['src']}, sheet {PPC.get('sheet') or 'In_Transit'}): {len(rows_)} shipment lines with a pending quantity ({late_} late - counted on today)"
              +(f"; {len(PPC['skipped'])} lines not mapped to a BOM material: "+'; '.join(PPC['skipped'][:4]) if PPC.get('skipped') else '')
              +(f"; codes not in the MRP: {', '.join(notm_)}" if notm_ else ''))
        if PPC.get('fixes'):
            print(f"  PPC In Transit: {len(PPC['fixes'])} lines traced to a different BOM code than their Final RM Code (your WSA PO + BL#, description, supplier, RM Code), e.g. "
                  +'; '.join(f"{x['final_code'] or '-'} -> {x['code']} {x['qty']:,.0f}" for x in PPC['fixes'][:3])+" - every line is on the Inbound sheet 'PPC In Transit'")
        chk_=[x for x in (PPC.get('lines') or []) if 'CHECK' in str(x.get('how'))]
        if chk_: print(f"  PPC In Transit: {len(chk_)} lines to CHECK - your WSA sheet has the shipment as another size / HS than the description says: "
                       +'; '.join(f"{x['description'][:40]} {x['qty']:,.0f} on {', '.join(x['wsa_sheets'])}" for x in chk_[:3]))
        for w_ in PPC.get('why') or []: print('  NOTE (PPC In Transit): '+w_)
    else: print("  PPC In Transit: NOT READ - "+'; '.join(PPC.get('why') or ['no PPC file with an In_Transit sheet'])+" - the PPC columns stay empty")
    print(f"  From today: live formulas on the sheet {T} ({len(tab):,} rows = module plan x BOM); Balance = stock typed on today's row (yellow), then Balance - Actual + Delivered")
    print("  Balance on today's row = stock MRP_Daily uses (MB52, PPC count rolled forward)"+('' if WSA_OEM_IN_BALANCE else ', JW stock at the OEM left out')+
          "; Demand = module plan x BOM"+('' if WSA_OEM_IN_BALANCE else ' without the OEM plan (OEM columns still show their modules)'))
    print(f"  Rows before {keep_from:%d-%b-%y} deleted; {len(arch_add)} delivery records from them moved to the sheet {ARCHIVE_SHEET}"
          +(f" ({len(oct_open)} of {keep_from-pd.Timedelta(days=1):%b-%y} with no Delivered quantity and not cancelled)" if oct_open else ''))
    splits=[f"{B['sheet']} -> {o}" for o,_,_,B in plan_sheets if o!=B['sheet'] and not B.get('src_label')]
    for l_ in GS['lines']:
        if l_.startswith('WARNING'): print('  '+l_)
    if [l_ for l_ in GS['lines'] if not l_.startswith('WARNING')]:
        print('  Front and back glass on their own sheets (one Balance per glass): '+'; '.join(l_ for l_ in GS['lines'] if not l_.startswith('WARNING')))
        print(f"    delivery rows split by the In_Transit lines of their BL#: {GS['n_bl']}; by 1 front : 1 back (no BL# line found): {GS['n_11']}")
    for m_,S_ in SHR.items():
        print(f"  {m_} on {len(S_['sheets'])} sheets - each customer on its own sheet, stock and PPC arrivals by supplier (Settings ALLOWED VENDORS, as MRP_Daily): "
              +' | '.join(f"'{_out_sheet_name(blocks[j])}' stock "+(', '.join(f'{n} {q:,.0f}' for n,q in sorted(S_['by_v'][j].items(),key=lambda t:-t[1]) if abs(q)>0.5) or '0')
                          +f", PPC arrivals {sum(S_['arr'][j]):,.0f} ("+'; '.join(f"{', '.join(v)} -> {k or 'any vendor'}" for k,v in sorted(S_['cust'][j].items(),key=lambda t:(t[0]=='',t[0])))+')'
                          for j in S_['sheets']))
    if gone: print('  Sheets left out of the WSA daily file (not needed): '+', '.join(gone))
    if splits: print('  Two materials on one sheet -> second material on its own sheet: '+'; '.join(splits))
    if dash_note: print('  '+dash_note)
    w_=max(len(x['label']) for x in report) if report else 10
    fmt=lambda v: '-' if v is None else f'{v:,.0f}'
    print(f"  {'Block':<{w_}}  {'Balance today old -> new':>28}  {'Demand next 7 d old -> new':>30}  {'Balance < 0 from old -> new':>28}")
    for x in report:
        print(f"  {x['label']:<{w_}}  {fmt(x['old_bal']):>13} -> {fmt(x['new_bal']):<11}  {fmt(x['old_d7']):>14} -> {fmt(x['new_d7']):<12}  "
              f"{(x['old_neg'].strftime('%d-%b-%y') if x['old_neg'] is not None else 'never'):>12} -> {(x['new_neg'].strftime('%d-%b-%y') if x['new_neg'] is not None else 'never')}")
    if skipped: print('  Not rebuilt: '+'; '.join(skipped))
    if unmatched: print("  NOTE: plan customers not on the WSA 'Lookup' sheet (own column, named after the MIS customer): "+'; '.join(sorted(unmatched)))
    if misonly: print('  NOTE: MIS demand of customers not in the module plan (in Demand, no modules): '+'; '.join(sorted(misonly)))
    if other: print(f"  NOTE: demand not split by customer (row '(not split by customer)' on {T}): "+', '.join(f'{k} ({v} d)' for k,v in other.items()))
    if dropped: print('  NOTE: columns not in the standard layout (left out): '+'; '.join(sorted(dropped)))
    if unit_notes: print('  Labels fixed (your text did not fit the material; the numbers are in the SAP / MRP unit): '+'; '.join(unit_notes))
    if VP_WSA: print(f"  Settings ALLOWED VENDORS on {len(VP_WSA)} planning sheets (columns {XL('vlock')} / {XL('vpool')}; column V subtracts {XL('vlock')}): "+vp_console(VP_WSA,d0))
    return out,report

# ===================== Material Plan sheet: one column per day =====================

if __name__=='__main__':
    t0=time.time()
    # ---------- Settings FIRST: every rule (targets, line thresholds, lead limits ...) is fixed before anything is calculated,
    #            so the Inbound file and MRP_Daily use the same values. Your MRP.xlsx first, else the last MRP_Daily. ----------
    MRP_DAILY_OUT=os.path.join(OUTDIR,f'MRP_Daily_{dt.date.today():%Y-%m-%d}.xlsx')
    if ALIGN_WSA and (not MIS or NO_MRP):
        print('NOTE: the WSA planning columns were not rebuilt - that needs the MIS RM Dashboard in the folder'+(' and no --no-mrp' if NO_MRP else ''))
        ALIGN_WSA=False
    if ALIGN_WSA:      # step 1 only calculates (outputs to a temporary folder), step 2 = normal run on the aligned WSA file
        import tempfile, shutil, subprocess
        ALIGN_MODE[0]=True; _ALIGN_TMP=tempfile.mkdtemp(prefix='align_wsa_')
        ROWS_CACHE_DIR=tempfile.mkdtemp(prefix='wsa_rows_'); os.environ['WSA_ROWS_CACHE']=ROWS_CACHE_DIR
        SRC_USER=SRC; SRC=os.path.join(_ALIGN_TMP,os.path.basename(SRC_USER)); shutil.copy2(SRC_USER,SRC)   # a copy: the aligned file may replace the file you ran it on
        MRP_DAILY_OUT=os.path.join(_ALIGN_TMP,os.path.basename(MRP_DAILY_OUT)); OUT=os.path.join(_ALIGN_TMP,os.path.basename(OUT))
        print(f'ALIGN WSA - step 1 of 2: rebuilding the planning columns of {os.path.basename(SRC_USER)} from the module plan x BOM and the stock (MB52 / PPC count) - about 20 s ...',flush=True)
        _QUIET[0]=True
    _prevs=sorted([os.path.join(OUTDIR,n_) for n_ in os.listdir(OUTDIR) if n_.startswith('MRP_Daily_') and n_.endswith('.xlsx') and not n_.startswith('~$')],key=os.path.getmtime)
    PREV_MRP=MRP_DAILY_OUT if os.path.exists(MRP_DAILY_OUT) else (_prevs[-1] if _prevs else None)
    SETTINGS_SRC[0]=read_settings([USER_MRP,PREV_MRP])
    print(f'Settings read from {os.path.basename(SETTINGS_SRC[0])}' if SETTINGS_SRC[0] else 'NOTE: no Settings sheet found (MRP.xlsx or last MRP_Daily) - script defaults used',
          f'- targets A {TARGET_A} / B {TARGET_B} / C {TARGET_C} / packaging {TARGET_CSPM} / other {TARGET_DAYS} days; lines under {COVER_CRITICAL} days')
    print('Reading the plan (materials and codes)...',flush=True)
    _tpc=('Date','Common Name.Common Name','Standard BOM.Material Code','Standard BOM.UOM')
    try: d=pd.read_excel(SRC,sheet_name='Total_Plan_Req',usecols=lambda c: str(c).strip() in _tpc)
    except ValueError:                     # the aligned WSA file has no Total_Plan_Req (it was a copy of the MIS plan): the MIS one
        try: d=pd.read_excel(MIS,sheet_name='Total Plan Req',usecols=lambda c: str(c).strip() in _tpc) if MIS else None
        except ValueError: d=None
        if d is None: die("No plan found: no 'Total_Plan_Req' sheet in the WSA file and no 'Total Plan Req' sheet in the MIS RM Dashboard")
    d.columns=[str(c).strip() for c in d.columns]
    d['Date']=pd.to_datetime(d['Date'],errors='coerce'); d=d.dropna(subset=['Date','Common Name.Common Name'])
    plan_end=d['Date'].max().to_pydatetime()
    codes_by_name=d.groupby('Common Name.Common Name')['Standard BOM.Material Code'].first().astype(str).to_dict()
    uom_name=d.groupby('Common Name.Common Name')['Standard BOM.UOM'].first().to_dict()
    print('Reading the planning sheets...',flush=True)
    G,skipped=extract_fast(SRC,plan_end); WSA_REFS[0]=G
    if not G: die('No planning sheets found (visible sheets with a "Required Plant Date" header)')
    for s in skipped: print('WARNING: planning sheet skipped -',s)
    print(f'Planning sheets read: {len(G)} blocks; with no scheduled deliveries: '+(', '.join(g['label'] for g in G if not g['inbound']) or 'none'))
    name2grp={}; name2grps={}; grp_mats={}     # a material can be on 2+ sheets (back glass 333: SB Back + Subanci Back - by customer and supplier)
    for g in G:
        for n,c in codes_by_name.items():
            if c in g['codes']:
                name2grp.setdefault(n,g['label']); name2grps.setdefault(n,[]).append(g['label']); grp_mats.setdefault(g['label'],[]).append(n)
    ginfo={g['label']:g for g in G}
    def pool_g(n):
        """The planning sheet of a material; on 2+ sheets: one pooled view like the MRP_Daily material row (balance = total of the
           sheets, days = total balance / total demand on the stock date); 'multi' = the sheets (their own periods)."""
        gs=[ginfo[l] for l in name2grps.get(n) or [] if l in ginfo]
        if len(gs)<2: return ginfo.get(name2grp.get(n,n))
        both=lambda key: {d:sum(g_[key][d] for g_ in gs) for d in set.intersection(*[set(g_.get(key) or {}) for g_ in gs])}
        ais=[g_.get('asof_info') or {} for g_ in gs]; dems=[(g_.get('dem_by_date') or {}).get(g_['asof']) for g_ in gs]
        ai={}
        if all(a_.get('balance') is not None for a_ in ais):
            ai['balance']=sum(a_['balance'] for a_ in ais)
            if all(dm_ for dm_ in dems):
                ai['plantdays']=ai['balance']/sum(dems)
                for k_ in ('plantport','ppcport'):
                    if all(a_.get(k_) is not None for a_ in ais): ai[k_]=sum(a_[k_]*dm_ for a_,dm_ in zip(ais,dems))/sum(dems)
        lr_=[g_['last_receipt'] for g_ in gs if g_.get('last_receipt')]; d_lr=max((x_[0] for x_ in lr_),default=None)
        return dict(gs[0],label=' + '.join(g_['label'] for g_ in gs),sheet=' + '.join(g_['sheet'] for g_ in gs),asof=max(g_['asof'] for g_ in gs),
                    asof_info=ai,bal_by_date=both('bal_by_date'),dem_by_date=both('dem_by_date'),multi=gs,
                    last_receipt=(d_lr,sum(x_[1] for x_ in lr_ if x_[0]==d_lr)) if lr_ else None)
    STOCK_DATE=Counter(g['asof'] for g in G).most_common(1)[0][0]
    # ---- one record per delivery (before splitting between materials) ----
    deliveries=[]
    for g in G:
        nn=sorted(grp_mats.get(g['label']) or [])
        if not nn: nn=[g['label']]
        for x in g['inbound']:
            deliveries.append(dict(g=g['label'],sheet=g['sheet'],mats=nn,x=x,asof=g['asof']))
    leads=[(e['x']['date']-e['x']['etd']).days for e in deliveries if hasattr(e['x']['etd'],'year')]
    MISBL=_mis_bl_dates(MIS) if MIS else ({},{})                     # MIS In Transit BL -> plant arrival date (MIS is the latest)
    for e in deliveries:
        x=e['x']; rem=x['remarks']
        e['status']=clean_status(rem); e['port']=port_of(rem)
        e['cancelled']='canc' in rem.lower()
        e['past']=(not e['cancelled']) and x['date']<e['asof']
        e['stage']='Cancelled' if e['cancelled'] else ('Past due' if x['date']<e['asof'] else ('In transit' if x['vessel'] else 'Open PO'))
        e['lead']=(x['date']-x['etd']).days if hasattr(x['etd'],'year') else None
        e['datecheck']='' if e['lead'] is None else ('Short lead' if e['lead']<SHORT_LEAD else ('Long lead' if e['lead']>LONG_LEAD else 'OK'))
        e['etdpassed']=(not e['cancelled']) and hasattr(x['etd'],'year') and x['etd']<pd.Timestamp(STOCK_DATE).to_pydatetime() and not x['vessel'] and not x['bl']
        miss=[]
        if not e['cancelled']:
            if not x['po']: miss.append('PO')
            if not x['shipper']: miss.append('Shipper')
            if not hasattr(x['etd'],'year'): miss.append('ETD')
        e['missing']=' '.join(miss)
        md_=None if e['cancelled'] else _mis_date_of(x['bl'],MISBL)          # MIS is always the most updated (user rule): 3+ days off -> move the WSA row
        dd_=(md_-pd.Timestamp(x['date']).normalize()).days if md_ is not None else None
        e['mis_date'],e['misdiff']=(md_,dd_) if (dd_ is not None and abs(dd_)>=3) else (None,None)
        e['chase']=0 if e['cancelled'] else int(e['stage']=='Past due' or bool(e['missing']) or e['etdpassed'] or e['datecheck']=='Short lead' or e['misdiff'] is not None)
    # ETD alert (user, Nov 2026: "My WSA have ETD that not in transit yet. I want alert how late those etd" + "ETD in the future also needed
    # alert"): every open WSA delivery row (not cancelled) with an ETD whose BL# is not on MIS In Transit = not shipped yet as far as MIS knows
    ETD_LATE.clear(); WSA_SHIPPER.clear(); gby={g_['label']:g_ for g_ in G}; base_=pd.Timestamp(STOCK_DATE).normalize()
    for e in deliveries:                                                         # shipper names typed on your WSA rows, per material code (email supplier fallback)
        sh_=' '.join(str(e['x']['shipper'] or '').split())
        if e['cancelled'] or not sh_ or re.fullmatch(r'[\d.]+',sh_): continue
        for c_ in (gby.get(e['g']) or {}).get('codes') or []: WSA_SHIPPER.setdefault(str(c_).upper(),[]).append(sh_)
    arch_=[]                                                                     # open deliveries of rows the redesign moved to the Archive sheet
    try:                                                                         # (State 'Not delivered' - the sheet still counts them in X2; checker finding)
        if ARCHIVE_SHEET in wb_ro(SRC).sheetnames:
            bys_={g_['sheet']:g_ for g_ in G}; AR_=list(wb_ro(SRC)[ARCHIVE_SHEET].iter_rows(values_only=True))
            hh_=[' '.join(str(v_ or '').split()).lower() for v_ in (AR_[0] if AR_ else [])]
            col_=lambda n_: hh_.index(n_) if n_ in hh_ else None
            iS,iD,iQ,iPO,iSh,iR,iE,iB,iSt=(col_(n_) for n_ in ('planning sheet','required plant date','scheduled delivery','po','shipper','remarks','etd','bl#','state'))
            for r_ in AR_[1:]:
                if not r_ or None in (iS,iD,iQ,iE,iSt) or str(r_[iSt] or '').strip()!='Not delivered': continue
                g_=bys_.get(str(r_[iS] or '').strip())
                if not g_ or not isinstance(r_[iQ],(int,float)) or r_[iQ]<=0 or not hasattr(r_[iD],'year'): continue
                rm_=str((r_[iR] if iR is not None else '') or '').strip()
                arch_.append(dict(g=g_['label'],sheet=g_['sheet'],cancelled='canc' in rm_.lower(),arch=True,
                                  x=dict(date=r_[iD],qty=r_[iQ],po=(r_[iPO] if iPO is not None else None),shipper=(r_[iSh] if iSh is not None else None),
                                         remarks=rm_,etd=r_[iE],bl=(r_[iB] if iB is not None else None))))
    except Exception as ex_: print(f'NOTE: Archive sheet not read for the ETD alert ({type(ex_).__name__}: {ex_})')
    def _bl_txt(v_):                                                             # a date in the BL# column (planner formula =A95-42) is no BL# (checker finding)
        if v_ is None or hasattr(v_,'year') or isinstance(v_,(int,float)): return ''
        t_=' '.join(str(v_).split())
        return '' if re.fullmatch(r'\d{4}-\d{2}-\d{2}( \d{2}:\d{2}:\d{2})?',t_) else t_
    for e in deliveries+arch_:
        x=e['x']
        if e['cancelled'] or not hasattr(x['etd'],'year'): continue
        if _mis_date_of(x['bl'],MISBL) is not None: continue                     # shipped: its BL# is on MIS In Transit
        g_=gby.get(e['g']) or {}; etd_=pd.Timestamp(x['etd']).normalize(); pd_=pd.Timestamp(x['date']).normalize()
        tr_=max(0,(pd_-etd_).days); passed=etd_<base_; blt_=_bl_txt(x['bl']); hasbl=bool(blt_)
        eta_=(base_+pd.Timedelta(days=tr_)) if passed else pd_                   # ETD passed: earliest arrival if it ships today; else the WSA plant date
        ETD_LATE.append(dict(sheet=e['sheet'],label=e['g'],codes=list(g_.get('codes') or []),uom=g_.get('uom'),po=(int(x['po']) if isinstance(x['po'],float) and x['po']==int(x['po']) else x['po']),
                             shipper=x['shipper'],qty=x['qty'],etd=etd_,late=(base_-etd_).days,passed=passed,plant=pd_,transit=tr_,eta_now=eta_,
                             delay=(eta_-pd_).days,bl=blt_,remarks=(('[Archive] ' if e.get('arch') else '')+str(x['remarks'] or '')).strip() or None,
                             status=(('BL# not on MIS In Transit (received? BL# typo? check)' if hasbl else 'Not shipped (no BL#)') if passed else
                                     ('BL# on WSA, not on MIS In Transit yet' if hasbl else 'Not shipped yet (no BL#)')),base=base_))
    if ETD_LATE:
        np_=sum(1 for r_ in ETD_LATE if r_['passed']); nf_=len(ETD_LATE)-np_; old_=max((r_['late'] for r_ in ETD_LATE if r_['passed']),default=None)
        print(f"ETD alert (WSA deliveries not on MIS In Transit): {len(ETD_LATE)} row(s) - ETD passed {np_}"+(f" (oldest {old_} days ago)" if old_ is not None else '')
              +f", ETD to come {nf_} ({sum(1 for r_ in ETD_LATE if not r_['passed'] and -r_['late']<=ETD_SOON)} within {ETD_SOON} days) - sheet '{ETD_SHEET}' in Inbound and MRP_Daily",flush=True)
    rows=[]
    for k,e in enumerate(deliveries):
        for n in e['mats']: rows.append((n,e['g'],e,1.0/len(e['mats']),n==e['mats'][0]))
    empties=[g for g in G if not g['inbound']]
    for g in empties:
        nn=sorted(grp_mats.get(g['label']) or []) or [g['label']]
        for k,n in enumerate(nn): rows.append((n,g['label'],dict(placeholder=True,g=g,asof=g['asof']),1.0/len(nn),k==0))
    rows.sort(key=lambda t:(t[0],dt.datetime.max if t[2].get('placeholder') else t[2]['x']['date']))
    def lastrec(g): return (f". Last actual receipt {g['last_receipt'][0]:%d-%b-%y}: {g['last_receipt'][1]:,.0f}" if g.get('last_receipt') else '')
    rem=Counter((e['x']['remarks'] or '(blank)') for e in deliveries)
    print('Status wording found on the sheets (deliveries): '+'; '.join(f'{k} x{v}' for k,v in rem.most_common()))
    sus=[k for k in rem if re.search(r'hold|resched|delay|postpon|defer|stop|no longer',k,re.I) and 'canc' not in k.lower()]
    if sus: print('CHECK: these remarks may mean the delivery is not coming but are still counted:',', '.join(sus))
    # ---------- previous file for the Changes sheet ----------
    if PREV is None:
        cands=sorted(p for p in glob.glob(os.path.join(OUTDIR,'Inbound_*.xlsx')) if os.path.abspath(p)!=os.path.abspath(OUT))
        PREV=cands[-1] if cands else None
    prev_del={}
    if PREV and os.path.exists(PREV):
        try:
            pw=wb_ro(PREV)['Inbound']
            hdr_prev=[str(h) for h in next(pw.iter_rows(min_row=5,max_row=5,values_only=True))]
            pidx=hdr_prev.index('Sheet stock date') if 'Sheet stock date' in hdr_prev else None
            for r in pw.iter_rows(min_row=6,values_only=True):
                if r[0] and r[2] and r[4] is not None:
                    k=(str(r[18]),ymd(r[2]),str(r[7] or ''),round(float(r[4]),3))
                    e=prev_del.setdefault(k,dict(mats=[],date=r[2],qty=r[4],remarks=r[6],po=r[7],etd=r[11],vessel=r[13],sheet=r[18],asof=(r[pidx] if pidx is not None and pidx<len(r) else None)))
                    if r[0] not in e['mats']: e['mats'].append(r[0])
        except Exception as ex: print('Could not read the previous file, so no Changes sheet:',ex); prev_del={}
    # ---------- styles ----------
    AR='Calibri'; hf=Font(name=AR,bold=True,color='FFFFFF',size=10); hfill=PatternFill('solid',fgColor='1F3864'); bf=Font(name=AR,size=10); bold=Font(name=AR,size=10,bold=True)
    blue=Font(name=AR,size=10,color='0000FF'); green=Font(name=AR,size=10,color='008000'); yel=PatternFill('solid',fgColor='FFFF00'); ctr=Alignment(horizontal='center',wrap_text=True)
    ttl=Font(name=AR,bold=True,size=12); num='#,##0;(#,##0);-'; dfy='dd-mmm-yy'
    amb=PatternFill('solid',bgColor='FFE699',fgColor='FFE699'); redf=PatternFill('solid',bgColor='F8CBAD',fgColor='F8CBAD')
    def hrow(ws,r,labels,start=1):
        for j,h in enumerate(labels):
            c=ws.cell(r,start+j,h); c.font=hf; c.fill=hfill; c.alignment=ctr
    def cell(v):
        if isinstance(v,float) and v==int(v): return int(v)
        return v
    wb=Workbook(); sm=wb.active; sm.title='Summary'
    ws=wb.create_sheet('Inbound'); tc=wb.create_sheet('To chase'); wk=wb.create_sheet('By Week'); ch=wb.create_sheet('Changes'); rd=wb.create_sheet('Read me')
    # ================= Inbound =================
    ws['A1']='Inbound - scheduled deliveries from the planning sheets (PO, shipper, ETD, ETA)'; ws['A1'].font=ttl
    ws['A2']='Plant date = Required Plant Date row of the planning sheet. A sheet that pools several materials (front + back glass, tapes) is split equally, one row per material. Counts = 1 unless cancelled or dated before that sheet\'s stock date; type 1 or 0 in the yellow column to force a row in or out.'; ws['A2'].font=bf
    ws['A3']='Report date'; ws['A3'].font=bold; ws['B3']=STOCK_DATE; ws['B3'].number_format=dfy; ws['B3'].font=blue; ws['B3'].fill=yel
    ws['C3']='Used for "Days to arrival", the weekly view and "next arrival". Default: the date most planning sheets were last updated. Each delivery is counted from its own sheet\'s stock date (column T).'; ws['C3'].font=bf
    heads=['Material','UOM','Plant date (arrival)','Quantity (this material)','Sheet quantity','Share of sheet qty','Remarks (as written)','PO','Shipper','Booking approval','Readiness','ETD','BL#','Vessel / ETA note','Stage','Include override (1/0)','Counts','Arrival week (Mon)','Planning sheet','Sheet stock date',
           'Status (clean)','Port','Lead days (plant date - ETD)','Date check','ETD passed, no vessel/BL','Missing info','Days to arrival','Source row','Unique delivery (1 = first row)','Chase?','BOM code','Sheet: Balance','Sheet: Planned modules per day','Sheet: Demand','Sheet: Actual modules per day','Sheet: Actual','Sheet: MW per day','Sheet: Actual MW per day','Sheet: Plant Days','Sheet: Plant + Port Days (ETA +5 days)','Sheet: Delivered','Sheet: Scheduled Delivery','MIS arrival (Final AtA) - the latest; shown when 3+ days from the plant date']
    HR=5; R0=HR+1
    def write_info(r,inf,sched):
        fm=['#,##0;(#,##0);-','#,##0','#,##0.0','#,##0','#,##0.0','#,##0.0','#,##0.0','0.0','0.0','#,##0']
        for k,key in enumerate(['balance','planned','demand','actmod','actual','mw','actmw','plantdays','plantport','delivered']):
            v=inf.get(key) if inf else None
            if v is not None:
                c=ws.cell(r,32+k,v); c.font=blue; c.number_format=fm[k]
        if sched is not None:
            c=ws.cell(r,42,sched); c.font=blue; c.number_format=num
    hrow(ws,HR,heads); ws.row_dimensions[HR].height=44
    for i,(n,g,e,sh,first) in enumerate(rows):
        r=R0+i
        if e.get('placeholder'):
            g_=e['g']; ws.cell(r,1,n).font=blue; ws.cell(r,2,uom_name.get(n,g_['uom'])).font=bf
            ws.cell(r,7,'No scheduled delivery on the planning sheet'+lastrec(g_)).font=blue
            ws.cell(r,15,'No scheduled delivery').font=bf; ws.cell(r,16).fill=yel; ws.cell(r,17,0).font=bold; ws.cell(r,19,g).font=blue
            c=ws.cell(r,27,g_['asof']); c.font=blue; c.number_format=dfy
            ws.cell(r,29,0).font=blue; ws.cell(r,30,0).font=bold; ws.cell(r,31,codes_by_name.get(n,'')).font=blue
            write_info(r,g_.get('asof_info'),None)
            continue
        x=e['x']
        vals=[n,uom_name.get(n,ginfo[g]['uom']),x['date'],x['qty']*sh,x['qty'],sh,x['remarks'] or None,cell(x['po']),x['shipper'],x['booking'],x['readiness'],x['etd'],x['bl'],x['vessel'] or None]
        for j,v in enumerate(vals,1):
            c=ws.cell(r,j,v); c.font=blue if j!=2 else bf
            if j in (3,10,11,12) and hasattr(v,'year'): c.number_format=dfy
            if j in (4,5): c.number_format=num
            if j==6: c.number_format='0%'
        ws.cell(r,15,f'=IF(ISNUMBER(SEARCH("canc",G{r})),"Cancelled",IF(C{r}<T{r},"Past due",IF(N{r}<>"","In transit","Open PO")))').font=bf
        ws.cell(r,16).fill=yel
        ws.cell(r,17,f'=IF(C{r}<T{r},0,IF(P{r}<>"",P{r},IF(O{r}="Cancelled",0,1)))').font=bold
        c=ws.cell(r,18,f'=C{r}-WEEKDAY(C{r},3)'); c.font=bf; c.number_format=dfy
        ws.cell(r,19,g).font=blue
        c=ws.cell(r,20,e['asof']); c.font=blue; c.number_format=dfy
        ws.cell(r,21,e['status'] or None).font=blue; ws.cell(r,22,e['port'] or None).font=blue
        ws.cell(r,23,f'=IF(ISNUMBER(L{r}),C{r}-L{r},"")').font=bf
        ws.cell(r,24,f'=IF(W{r}="","",IF(W{r}<Summary!$B$7,"Short lead",IF(W{r}>Summary!$B$8,"Long lead","OK")))').font=bf
        ws.cell(r,25,f'=IF(AND(O{r}<>"Cancelled",ISNUMBER(L{r}),N{r}="",M{r}=""),IF(L{r}<$B$3,"ETD passed - no vessel/BL",""),"")').font=bf
        ws.cell(r,26,f'=IF(O{r}="Cancelled","",TRIM(IF(H{r}="","PO ","")&IF(I{r}="","Shipper ","")&IF(NOT(ISNUMBER(L{r})),"ETD","")))').font=bf
        c=ws.cell(r,27,f'=C{r}-$B$3'); c.font=bf; c.number_format='0;-0;0'
        ws.cell(r,28,x['row']).font=blue
        ws.cell(r,29,1 if first else 0).font=blue
        ws.cell(r,30,f'=IF(O{r}="Cancelled",0,IF(OR(O{r}="Past due",Y{r}<>"",Z{r}<>"",X{r}="Short lead",AQ{r}<>""),1,0))').font=bold
        if e.get('misdiff') is not None:                          # MIS is the latest: the WSA row is to be moved to this date
            c=ws.cell(r,43,e['mis_date'].to_pydatetime()); c.number_format=dfy; c.font=Font(name=AR,size=10,bold=True,color='C00000')
        ws.cell(r,31,codes_by_name.get(n,'')).font=blue
        write_info(r,x.get('info'),x['qty'])
    R1=R0+len(rows)-1
    rng=lambda col: f'Inbound!${col}${R0}:${col}${R1}'
    if rows:
        ws.auto_filter.ref=f'A{HR}:AQ{R1}'
        ws.conditional_formatting.add(f'A{R0}:AE{R1}',FormulaRule(formula=[f'$O{R0}="Past due"'],fill=amb))
        ws.conditional_formatting.add(f'A{R0}:AE{R1}',FormulaRule(formula=[f'$O{R0}="No scheduled delivery"'],fill=PatternFill('solid',bgColor='DDEBF7',fgColor='DDEBF7')))
        ws.conditional_formatting.add(f'A{R0}:AE{R1}',FormulaRule(formula=[f'AND($Q{R0}=0,$O{R0}<>"No scheduled delivery")'],font=Font(name=AR,color='999999',strike=True)))
        ws.conditional_formatting.add(f'X{R0}:Z{R1}',FormulaRule(formula=[f'AND(X{R0}<>"",X{R0}<>"OK")'],fill=redf))
    for c_,w_ in zip(['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE'],[28,8,13,14,13,9,22,12,14,14,14,13,16,30,12,12,9,13,26,13,20,14,12,12,24,18,10,9,12,8,16]): ws.column_dimensions[c_].width=w_
    for c_ in ['AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP']: ws.column_dimensions[c_].width=15
    ws.column_dimensions['AQ'].width=18
    ws['AF4']='From the planning sheet, on the plant-date row (rows with no scheduled delivery: the stock-date row)'; ws['AF4'].font=bold
    ws.freeze_panes=f'B{R0}'
    # ================= Summary =================
    sm['A1']='Inbound summary'; sm['A1'].font=ttl
    sm['A2']='Counts and quantities below follow the Counts column on the Inbound sheet (cancelled and past-due deliveries are excluded). Quantities are in each material\'s own unit.'; sm['A2'].font=bf
    sm['A4']='Source file'; sm['B4']=os.path.basename(SRC); sm['A5']='Generated'; sm['B5']=dt.datetime.now().replace(microsecond=0); sm['B5'].number_format='dd-mmm-yy hh:mm'
    sm['A6']='Report date'; sm['B6']='=Inbound!B3'; sm['B6'].number_format=dfy
    sm['A7']='Short lead limit (days, ETD to plant date)'; sm['B7']=SHORT_LEAD; sm['A8']='Long lead limit (days)'; sm['B8']=LONG_LEAD
    for a_ in ('A4','A5','A6','A7','A8'): sm[a_].font=bold
    for b_ in ('B4','B5','B6'): sm[b_].font=bf
    for b_ in ('B7','B8'): sm[b_].font=blue; sm[b_].fill=yel
    hrow(sm,10,['Deliveries','Count'])
    K=[('Counted (will arrive)',f'=SUMIFS({rng("AC")},{rng("Q")},1)'),('Cancelled',f'=SUMIFS({rng("AC")},{rng("O")},"Cancelled")'),('Past due (not counted)',f'=SUMIFS({rng("AC")},{rng("O")},"Past due")'),
       ('To chase (see To chase sheet)',f'=SUMIFS({rng("AC")},{rng("AD")},1)'),('Arriving in the next 14 days (counted)',f'=SUMIFS({rng("AC")},{rng("Q")},1,{rng("C")},">="&$B$6,{rng("C")},"<="&$B$6+14)'),
       ('In transit (counted)',f'=SUMIFS({rng("AC")},{rng("Q")},1,{rng("O")},"In transit")'),('Open PO, not shipped (counted)',f'=SUMIFS({rng("AC")},{rng("Q")},1,{rng("O")},"Open PO")'),('Materials with no scheduled delivery',f'=COUNTIFS({rng("O")},"No scheduled delivery")')]
    for i,(lab,f_) in enumerate(K):
        sm.cell(11+i,1,lab).font=bf; c=sm.cell(11+i,2,f_); c.font=bold; c.number_format=num
    MR=20
    hrow(sm,MR,['Material','UOM','Counted deliveries','Counted quantity','In transit qty','Open PO qty','Next arrival','Next arrival qty','Past due qty (not counted)','Cancelled qty','Deliveries to chase','Sheet stock date','Last actual receipt (info)','Last receipt qty (info)','Sheet balance at stock date','Sheet Plant Days at stock date','Sheet Plant + Port Days at stock date',
                 'Plant Days < 0 (from \u2192 to)','Plant + Port Days < 0 (from \u2192 to)',
                 'Sheet Plant + Port Days (PPC In Transit) at stock date','Plant + Port Days (PPC In Transit) < 0 (from \u2192 to)']); sm.row_dimensions[MR].height=44
    mats=sorted({n for n,_,_,_,_ in rows})
    mat_date={}
    for n_,g_,e_,sh_,f_ in rows: mat_date.setdefault(n_,e_['asof'])
    for i,n in enumerate(mats):
        r=MR+1+i; A=f'$A{r}'
        sm.cell(r,1,n).font=bf
        sm.cell(r,2,f'=IFERROR(INDEX({rng("B")},MATCH({A},{rng("A")},0)),"")').font=green
        f=lambda s: sm.cell(r,s)
        f(3).value=f'=COUNTIFS({rng("A")},{A},{rng("Q")},1)'
        f(4).value=f'=SUMIFS({rng("D")},{rng("A")},{A},{rng("Q")},1)'
        f(5).value=f'=SUMIFS({rng("D")},{rng("A")},{A},{rng("Q")},1,{rng("O")},"In transit")'
        f(6).value=f'=SUMIFS({rng("D")},{rng("A")},{A},{rng("Q")},1,{rng("O")},"Open PO")'
        f(7).value=f'=IF(COUNTIFS({rng("A")},{A},{rng("Q")},1,{rng("C")},">="&$B$6)=0,"",_xlfn.MINIFS({rng("C")},{rng("A")},{A},{rng("Q")},1,{rng("C")},">="&$B$6))'
        f(8).value=f'=IF(G{r}="",0,SUMIFS({rng("D")},{rng("A")},{A},{rng("Q")},1,{rng("C")},G{r}))'
        f(9).value=f'=SUMIFS({rng("D")},{rng("A")},{A},{rng("O")},"Past due")'
        f(10).value=f'=SUMIFS({rng("D")},{rng("A")},{A},{rng("O")},"Cancelled")'
        f(11).value=f'=COUNTIFS({rng("A")},{A},{rng("AD")},1)'
        for s in range(3,12):
            f(s).font=bf; f(s).number_format=dfy if s==7 else num
        f(12).value=mat_date[n]; f(12).font=blue; f(12).number_format=dfy
        gl=pool_g(n); lr=gl.get('last_receipt') if gl else None
        if lr:
            c=f(13); c.value=lr[0]; c.font=blue; c.number_format=dfy
            c=f(14); c.value=lr[1]; c.font=blue; c.number_format=num
        ai=(gl.get('asof_info') or {}) if gl else {}
        for k_,key_,fm_ in ((15,'balance',num),(16,'plantdays','0.0'),(17,'plantport','0.0')):
            if ai.get(key_) is not None:
                c=f(k_); c.value=ai[key_]; c.font=blue; c.number_format=fm_
        if gl:                                                    # from -> to dates where the two columns above are below zero
            gl1=gl.get('multi') or [gl]                           # a material on 2+ sheets: the periods of each sheet
            for k_,key_ in ((18,'pd'),(19,'pp'),(21,'pi')):
                t_=[]; neg_=False
                for g1_ in gl1:
                    bzp=below_zero_periods(g1_)
                    if key_=='pi' and (g1_.get('asof_info') or {}).get('ppcport') is None and not bzp[key_]: continue      # no PPC column on this sheet
                    t_.append(periods_text(bzp[key_]) if len(gl1)==1 else f"{g1_['label']}: {periods_text(bzp[key_])}"); neg_|=bool(bzp[key_])
                if not t_: continue
                c=f(k_); c.value='\n'.join(t_); c.font=Font(name=AR,size=10,color='C00000' if neg_ else '008000')
                c.alignment=Alignment(wrap_text=True,vertical='top')
            if ai.get('ppcport') is not None:
                c=f(20); c.value=ai['ppcport']; c.font=blue; c.number_format='0.0'
    MEND=MR+len(mats)
    SR=MEND+3
    sm.cell(SR-1,1,'Every planning sheet read (including those with no scheduled deliveries)').font=bold
    hrow(sm,SR,['Planning sheet','Deliveries (all)','Total scheduled qty on the sheet','Counted deliveries','Cancelled','Past due','To chase','Stock date']); sm.row_dimensions[SR].height=32
    sheets=sorted(ginfo)
    for i,g in enumerate(sheets):
        r=SR+1+i; A=f'$A{r}'; sm.cell(r,1,g).font=bf
        sm.cell(r,2,f'=SUMIFS({rng("AC")},{rng("S")},{A})'); sm.cell(r,3,f'=SUMIFS({rng("E")},{rng("S")},{A},{rng("AC")},1)')
        sm.cell(r,4,f'=SUMIFS({rng("AC")},{rng("S")},{A},{rng("Q")},1)'); sm.cell(r,5,f'=SUMIFS({rng("AC")},{rng("S")},{A},{rng("O")},"Cancelled")')
        sm.cell(r,6,f'=SUMIFS({rng("AC")},{rng("S")},{A},{rng("O")},"Past due")'); sm.cell(r,7,f'=SUMIFS({rng("AC")},{rng("S")},{A},{rng("AD")},1)')
        for s in range(2,8): sm.cell(r,s).font=bf; sm.cell(r,s).number_format=num
        c=sm.cell(r,8,ginfo[g]['asof']); c.font=blue; c.number_format=dfy
    for c_,w_ in zip('ABCDEFGHIJKLMNOPQRSTU',[44,16,16,16,14,14,14,14,16,14,14,14,16,16,16,16,18,40,40,18,40]): sm.column_dimensions[c_].width=w_
    sm.cell(SR+len(sheets)+2,1,'A sheet with 0 deliveries was read but has nothing in its Scheduled Delivery column (for example only actual receipts in the Delivered column).').font=bf
    sm.freeze_panes='A4'
    # ================= Below zero: from / to dates where Plant Days or Plant + Port Days < 0 =================
    bz=wb.create_sheet('Below zero',1)
    bz['A1']='Below zero - the dates (from \u2192 to) where a planning sheet\'s Plant Days or Plant + Port Days are less than 0'; bz['A1'].font=ttl
    bz['A2']=('Plant Days = Balance / Demand (stock only, no delivery counted). Plant + Port Days = (Balance + scheduled deliveries at the plant by that day) / Demand. '
              'Plant + Port Days (PPC In Transit) = (Balance + arrivals of your PPC In_Transit sheet by that day) / Demand. '
              'From the sheet stock date to the last date of the sheet; a day with no demand keeps the state of the day before. '
              'Until end of sheet = Yes: still below zero on the last date of the sheet.'); bz['A2'].font=bf
    hrow(bz,4,['Planning sheet','Materials','UOM','Sheet stock date','Below zero','From','To','Days','Lowest days value','Until end of sheet',
               'Next delivery on / after From','Next delivery qty','Its stage']); bz.row_dimensions[4].height=32
    bzr=[]
    for lab_,g_ in ginfo.items():
        bzp=below_zero_periods(g_); mats_=sorted(grp_mats.get(lab_) or []) or [lab_]
        for key_,meas_ in (('pd','Plant Days < 0'),('pp','Plant + Port Days < 0'),('pi','Plant + Port Days (PPC In Transit) < 0')):
            for p_ in bzp[key_]:
                if key_=='pi':
                    pa=sorted((d_,q_) for d_,q_ in (g_.get('ppcit_by_date') or {}).items() if d_>=p_['fr'])
                    nx=(pa[0][0],pa[0][1],'PPC In Transit') if pa else (None,None,'')
                else:
                    nxt=min((e for e in deliveries if e['g']==lab_ and not e['cancelled'] and e['x']['date']>=p_['fr']),key=lambda e:e['x']['date'],default=None)
                    nx=(nxt['x']['date'],nxt['x']['qty'],nxt['stage']) if nxt else (None,None,'')
                bzr.append((lab_,', '.join(mats_),g_['uom'],g_['asof'],meas_,p_['fr'],p_['to'],(p_['to']-p_['fr']).days+1,p_['low'],
                            'Yes' if p_['open'] else '')+nx)
    bzr.sort(key=lambda t_:(t_[5],t_[4],t_[0]))
    for i,t_ in enumerate(bzr):
        for j,v in enumerate(t_,1):
            c=bz.cell(5+i,j,v); c.font=bf
            if j in (4,6,7,11): c.number_format=dfy
            elif j==9: c.number_format='0.0'
            elif j in (8,12): c.number_format=num
        if t_[4].startswith('Plant + Port'): bz.cell(5+i,5).font=Font(name=AR,size=10,bold=True,color='C55A11' if 'PPC' in t_[4] else 'C00000')
    if not bzr: bz.cell(5,1,'No planning sheet goes below zero from its stock date.').font=bf
    for c_,w_ in zip('ABCDEFGHIJKLM',[30,34,8,13,22,12,12,7,12,11,16,14,14]): bz.column_dimensions[c_].width=w_
    bz.freeze_panes='A5'; bz.auto_filter.ref=f'A4:M{4+max(1,len(bzr))}'
    # ================= To chase =================
    tc['A1']='To chase - deliveries that need follow-up (snapshot at the time of the run)'; tc['A1'].font=ttl
    tc['A2']='Reasons: past due = dated before the sheet stock date and not cancelled (not counted); missing PO / shipper / ETD; ETD passed with no vessel or BL; short lead = plant date under the short-lead limit after ETD; MIS arrival 3+ days from the plant date (MIS is always the latest: move the WSA row).'; tc['A2'].font=bf
    hrow(tc,4,['Planning sheet','Sheet stock date','Materials','Plant date','Days to arrival (vs report date)','Sheet quantity','Stage','Remarks (as written)','PO','Shipper','ETD','BL#','Vessel / ETA note','Reasons']); tc.row_dimensions[4].height=32
    chase=[e for e in deliveries if e['chase']]
    for i,e in enumerate(sorted(chase,key=lambda e:e['x']['date'])):
        x=e['x']; why=[]
        if e['stage']=='Past due': why.append(f"Past due ({(pd.Timestamp(e['asof'])-pd.Timestamp(x['date'])).days} d before stock date)")
        if e['missing']: why.append('Missing '+e['missing'].replace(' ',', '))
        if e['etdpassed']: why.append('ETD passed, no vessel/BL')
        if e['datecheck']=='Short lead': why.append(f"Short lead ({e['lead']} d after ETD)")
        if e.get('misdiff') is not None: why.append(f"MIS arrival {e['mis_date']:%d-%b} ({e['misdiff']:+d} d) - MIS is the latest: move the WSA row")
        vals=[e['g'],e['asof'],', '.join(e['mats']),x['date'],(pd.Timestamp(x['date'])-pd.Timestamp(STOCK_DATE)).days,x['qty'],e['stage'],x['remarks'] or None,cell(x['po']),x['shipper'],x['etd'],x['bl'],x['vessel'] or None,'; '.join(why)]
        for j,v in enumerate(vals,1):
            c=tc.cell(5+i,j,v); c.font=bf
            if j in (2,4,11) and hasattr(v,'year'): c.number_format=dfy
            if j==6: c.number_format=num
    base=5+len(chase)
    for j,g_ in enumerate(empties):
        nn=[n for n,gg,e,sh,fi in rows if gg==g_['label']]
        vals=[g_['label'],g_['asof'],', '.join(nn),None,None,None,'No scheduled delivery',None,None,None,None,None,None,'No scheduled delivery on the planning sheet'+lastrec(g_)]
        for k,v in enumerate(vals,1):
            c=tc.cell(base+j,k,v); c.font=bf
            if k==2: c.number_format=dfy
    for c_,w_ in zip('ABCDEFGHIJKLMN',[28,13,34,13,14,14,22,22,12,14,13,16,30,60]): tc.column_dimensions[c_].width=w_
    tc.freeze_panes='A5'
    if chase or empties: tc.auto_filter.ref=f'A4:N{4+len(chase)+len(empties)}'
    # ================= By Week =================
    wk['A1']='Counted deliveries by material and arrival week (Counts = 1 on the Inbound sheet)'; wk['A1'].font=ttl
    w0=pd.Timestamp(STOCK_DATE); w0=(w0-pd.Timedelta(days=w0.weekday())).to_pydatetime(); NW=26
    _mons={(g_['asof']-dt.timedelta(days=g_['asof'].weekday())).date() for g_ in G}
    BALHDR=('Balance on Monday %s (sheet)'%min(_mons).strftime('%d-%b-%y')) if len(_mons)==1 else 'Balance on the latest Monday (sheet)'
    _mons2={m_+dt.timedelta(days=7) for m_ in _mons}
    BALHDR2=('Balance on Monday %s (sheet)'%min(_mons2).strftime('%d-%b-%y')) if len(_mons2)==1 else 'Balance on the next Monday (sheet)'
    def block(top,title,stage):
        wk.cell(top,1,title).font=bold
        hrow(wk,top+1,['Material','UOM','Sheet stock date','Total',BALHDR,BALHDR2]); wk.row_dimensions[top+1].height=32
        for k in range(NW):
            c=wk.cell(top+1,7+k,w0+dt.timedelta(days=7*k)); c.number_format='dd-mmm'; c.font=hf; c.fill=hfill; c.alignment=ctr
        for i,n in enumerate(mats):
            r=top+2+i; wk.cell(r,1,n).font=bf
            wk.cell(r,2,f'=IFERROR(INDEX({rng("B")},MATCH($A{r},{rng("A")},0)),"")').font=bf
            c=wk.cell(r,3,mat_date[n]); c.font=blue; c.number_format=dfy
            c=wk.cell(r,4,f'=SUM(G{r}:{L(6+NW)}{r})'); c.font=bold; c.number_format=num
            gl_=pool_g(n); bal_=None
            if gl_:
                mon_=gl_['asof']-dt.timedelta(days=gl_['asof'].weekday())
                for back_ in range(0,7):
                    bal_=gl_['bal_by_date'].get(mon_-dt.timedelta(days=back_))
                    if bal_ is not None: break
            if bal_ is not None:
                c=wk.cell(r,5,bal_); c.font=blue; c.number_format=num
            if gl_:
                b2_=gl_['bal_by_date'].get(mon_+dt.timedelta(days=7))
                if b2_ is not None:
                    c=wk.cell(r,6,b2_); c.font=blue; c.number_format=num
            for k in range(NW):
                col=L(7+k); extra=f',{rng("O")},"{stage}"' if stage else ''
                c=wk.cell(r,7+k,f'=SUMIFS({rng("D")},{rng("A")},$A{r},{rng("Q")},1,{rng("R")},{col}${top+1}{extra})'); c.font=bf; c.number_format=num
        tr=top+2+len(mats)
        wk.cell(tr,1,'Total deliveries (count)').font=bold
        for k in range(NW):
            col=L(7+k); extra=f',{rng("O")},"{stage}"' if stage else ''
            c=wk.cell(tr,7+k,f'=SUMIFS({rng("AC")},{rng("Q")},1,{rng("R")},{col}${top+1}{extra})'); c.font=bold
        return tr
    t1=block(3,'All counted deliveries (quantity in each material\'s own unit)',None)
    t2=block(t1+3,'In transit (has a vessel / ETA note)','In transit'); t3=block(t2+3,'Open PO (no vessel yet)','Open PO')
    wk.column_dimensions['A'].width=30; wk.column_dimensions['B'].width=8
    wk.column_dimensions['C'].width=13
    for k in range(NW+3): wk.column_dimensions[L(4+k)].width=11
    wk.column_dimensions['E'].width=16; wk.column_dimensions['F'].width=16
    wk.freeze_panes='G4'
    wk.cell(t3+2,1,'Balances are the planning sheet\'s own start-of-day figures: the first is the latest Monday on or before the stock date (typed by the planners); the second is the following Monday, which is the sheet\'s projection (demand taken off, scheduled deliveries not added). For sheets that pool two materials (front + back glass, perforated tapes) it is the pooled balance, shown on both materials.').font=bf
    # chart: deliveries per week by stage (count)
    ch1=BarChart(); ch1.type='col'; ch1.grouping='stacked'; ch1.overlap=100; ch1.title='Counted deliveries per arrival week (number of deliveries)'
    for tr,name in ((t2,'In transit'),(t3,'Open PO')):
        ref=Reference(wk,min_col=7,max_col=6+NW,min_row=tr); ch1.add_data(ref,from_rows=True,titles_from_data=False)
    from openpyxl.chart.series import SeriesLabel
    ch1.series[0].tx=SeriesLabel(v='In transit'); ch1.series[1].tx=SeriesLabel(v='Open PO')
    ch1.set_categories(Reference(wk,min_col=7,max_col=6+NW,min_row=4)); ch1.height=8; ch1.width=26
    wk.add_chart(ch1,f'A{t3+4}')
    # ---------- By Week numbers as plain values (for the MRP workbook) ----------
    def compute_byweek():
        data={'all':{},'transit':{},'open':{}}; tot={k:[0]*NW for k in data}; meta={}; open_daily={}; open_sup={}
        for n,g,e,sh,first in rows:
            meta.setdefault(n,dict(uom=uom_name.get(n,ginfo[g]['uom']),asof=e['asof'],g=g))
            for k in data: data[k].setdefault(n,[0.0]*NW)
            if e.get('placeholder') or e['stage'] not in ('In transit','Open PO'): continue
            if e['stage']=='Open PO':        # by plant date, for the Daily Stock sheet
                d_=pd.Timestamp(e['x']['date']).normalize(); open_daily.setdefault(n,{}); open_daily[n][d_]=open_daily[n].get(d_,0.0)+e['x']['qty']*sh
                sp_=' '.join(str(e['x'].get('shipper') or '').split()) or '(no supplier)'
                po_=str(e['x'].get('po') or '').strip()
                if po_.endswith('.0'): po_=po_[:-2]
                sp_=f'{po_}|{sp_}'                  # PO is used to find the supplier name (shipper may be a vendor code)
                od_=open_sup.setdefault(n,{}).setdefault(sp_,{}); od_[d_]=od_.get(d_,0.0)+e['x']['qty']*sh
            x=e['x']; dd=pd.Timestamp(x['date']); mon=(dd-pd.Timedelta(days=dd.weekday())).to_pydatetime(); k=(mon-w0).days//7
            if not (0<=k<NW): continue
            q=x['qty']*sh; key='transit' if e['stage']=='In transit' else 'open'
            data['all'][n][k]+=q; data[key][n][k]+=q
            if first: tot['all'][k]+=1; tot[key][k]+=1
        for n,m in meta.items():
            gl_=pool_g(n); b1=b2=None
            if gl_:
                mon_=gl_['asof']-dt.timedelta(days=gl_['asof'].weekday())
                for back_ in range(0,7):
                    b1=gl_['bal_by_date'].get(mon_-dt.timedelta(days=back_))
                    if b1 is not None: break
                b2=gl_['bal_by_date'].get(mon_+dt.timedelta(days=7))
            m['bal1']=b1; m['bal2']=b2
        deliv=[]
        for n,g,e,sh,first in rows:
            if e.get('placeholder'): continue
            x=e['x']; deliv.append(dict(name=n,grp=g,sheet=e['sheet'],stage=e['stage'],qty=x['qty']*sh,date=x['date'],po=x.get('po'),bl=x.get('bl'),
                                        vessel=x.get('vessel'),shipper=x.get('shipper'),remarks=x.get('remarks'),etd=x.get('etd'),first=first))
        sheets={lab:dict(asof=g_['asof'],bal=g_['bal_by_date'].get(g_['asof']),dem=g_.get('dem_by_date',{}),uom=g_['uom'],
                         plantdays=(g_.get('asof_info') or {}).get('plantdays'),plantport=(g_.get('asof_info') or {}).get('plantport'),pdays=g_.get('pdays_by_date') or {},
                         members=sorted(grp_mats.get(lab) or []) or [lab],sheet=g_['sheet']) for lab,g_ in ginfo.items()}
        return dict(data=data,tot=tot,meta=meta,open_daily=open_daily,open_sup=open_sup,deliv=deliv,sheets=sheets,weeks=[w0+dt.timedelta(days=7*k) for k in range(NW)],hdr1=BALHDR,hdr2=BALHDR2)
    BW=compute_byweek()
    # ================= Changes =================
    ch['A1']='Changes since the previous Inbound file'; ch['A1'].font=ttl
    ch['A2']=(f'Compared with {os.path.basename(PREV)}' if prev_del else 'No previous Inbound file found (put the previous Inbound_*.xlsx in the same folder, or use --prev). Nothing to compare, so this sheet is empty.'); ch['A2'].font=bf
    cur={}
    for e in deliveries:
        x=e['x']; cur[(e['g'],ymd(x['date']),str(x['po'] or ''),round(x['qty'],3))]=e
    out=[]
    if prev_del:
        left_cur={k:e for k,e in cur.items() if k not in prev_del}; left_prev={k:p for k,p in prev_del.items() if k not in cur}
        for k,e in cur.items():
            if k in prev_del:
                p=prev_del[k]; x=e['x']; diff=[]
                if str(p['remarks'] or '')!=str(x['remarks'] or ''): diff.append(f"status: {p['remarks'] or '-'} -> {x['remarks'] or '-'}")
                if str(p['etd'] or '')[:10]!=str(x['etd'] or '')[:10]: diff.append(f"ETD: {str(p['etd'] or '-')[:10]} -> {str(x['etd'] or '-')[:10]}")
                if str(p['vessel'] or '')!=str(x['vessel'] or ''): diff.append('vessel/ETA note changed')
                if diff: out.append(('Changed',', '.join(e['mats']),e['g'],x['date'],x['qty'],x['remarks'],x['po'],x['etd'],'; '.join(diff)))
        for k,e in list(left_cur.items()):          # same PO and quantity on another date -> date moved
            if not k[2]: continue
            m=[kp for kp in left_prev if kp[0]==k[0] and kp[2]==k[2] and kp[3]==k[3]]
            if m:
                p=left_prev.pop(m[0]); del left_cur[k]; x=e['x']
                out.append(('Date moved',', '.join(e['mats']),e['g'],x['date'],x['qty'],x['remarks'],x['po'],x['etd'],f"{p['date']:%d-%b-%y} -> {x['date']:%d-%b-%y}"))
        for k,e in list(left_cur.items()):          # same PO and date, different quantity
            if not k[2]: continue
            m=[kp for kp in left_prev if kp[0]==k[0] and kp[1]==k[1] and kp[2]==k[2]]
            if m:
                p=left_prev.pop(m[0]); del left_cur[k]; x=e['x']
                out.append(('Quantity changed',', '.join(e['mats']),e['g'],x['date'],x['qty'],x['remarks'],x['po'],x['etd'],f"{p['qty']:,.0f} -> {x['qty']:,.0f}"))
        for k,e in left_cur.items():
            x=e['x']; out.append(('New',', '.join(e['mats']),e['g'],x['date'],x['qty'],x['remarks'],x['po'],x['etd'],''))
        for k,p in left_prev.items(): out.append(('Removed',', '.join(p['mats']),p['sheet'],p['date'],p['qty'],p['remarks'],p['po'],p['etd'],''))
        prev_sheets={k[0] for k in prev_del}; now_sheets={e['g'] for e in deliveries}
        gone=sorted(prev_sheets-now_sheets)
        if gone: print('NOTE: planning sheets that had deliveries in the previous file but have none now:',', '.join(gone))
        for g_ in gone: out.append(('Sheet emptied','',g_,None,None,None,None,None,'had scheduled deliveries in the previous file, none now'))
        for g_ in sorted(now_sheets-prev_sheets): out.append(('Sheet restarted','',g_,None,None,None,None,None,'had no scheduled deliveries in the previous file, has some now'))
    hrow(ch,4,['Change','Material(s)','Planning sheet','Sheet stock date (now)','Sheet stock date (previous file)','Plant date','Sheet quantity','Remarks (status)','PO','ETD','What changed']); ch.row_dimensions[4].height=32
    order={'New':0,'Date moved':1,'Quantity changed':2,'Changed':3,'Removed':4,'Sheet emptied':5,'Sheet restarted':6}
    for i,t in enumerate(sorted(out,key=lambda t:(order[t[0]],t[1],str(t[3])))):
        for j,v in enumerate(t,1):
            c=ch.cell(5+i,j,v); c.font=bf
            if j in (4,8) and hasattr(v,'year'): c.number_format=dfy
            if j==5: c.number_format=num
    for c_,w_ in zip('ABCDEFGHI',[16,30,26,13,14,22,12,13,50]): ch.column_dimensions[c_].width=w_
    ch.freeze_panes='A5'
    # ================= Read me =================
    lines=['Inbound workbook - what each sheet and column means','',
     'Summary   - counts and quantities by material and by planning sheet; short/long lead limits (yellow) drive the Date check column.',
     'Below zero - one row per period (from \u2192 to) where a planning sheet\'s Plant Days or Plant + Port Days are below 0, with the next delivery; the same periods are in the last two Summary columns.',
     'Inbound   - one row per delivery and material. Yellow column = type 1 or 0 to force a delivery in or out of the counts.',
     'To chase  - deliveries needing follow-up, with the reason (snapshot at run time).',
     'By Week   - counted deliveries per arrival week: all, in transit, open PO; chart at the bottom.',
     'Changes   - new / date moved / quantity changed / changed (status, ETD, vessel) / removed since the previous file.','',
     'Stage: Cancelled (status contains "canc"), Past due (plant date before that sheet\'s stock date), In transit (vessel / ETA note present), Open PO (no vessel yet).',
     'Counts: 1 = the delivery is expected to arrive and is counted; 0 = cancelled or past due (unless overridden).',
     'Date check: Short lead / Long lead = days from ETD to plant date outside the limits on Summary.',
     'ETD passed, no vessel/BL: ETD is before the report date but neither a vessel / ETA note nor a BL# is filled in.',
     'Missing info: PO, Shipper or ETD blank (not for cancelled deliveries).',
     'Status (clean) and Port: the Remarks column mixes statuses (Approved for booking, Planned, GRN pending, Cancelled) with ports (Houston, Los Angeles); they are separated here. Remarks (as written) is unchanged.',
     'Source row: the row on the planning sheet the delivery was read from.',
     'Sheet: ... columns (AF to AP on Inbound): Balance, Planned modules per day, Demand, Actual modules per day, Actual, MW per day, Actual MW per day, Plant Days, Plant + Port Days (ETA +5 days), Delivered and Scheduled Delivery, copied from the planning sheet row of the plant date (or of the stock date when nothing is scheduled). On the Summary: balance and plant days at the stock date.',
     'A planning sheet with no scheduled deliveries still has a row (Stage = No scheduled delivery, blue) on Inbound, in By Week and on the Summary, and a line on To chase, with its last actual receipt as information.',
     'Unique delivery: 1 on the first row of a delivery, so a delivery split between two materials is counted once in the Summary.']
    for i,t in enumerate(lines):
        c=rd.cell(1+i,1,t); c.font=ttl if i==0 else bf
    rd.column_dimensions['A'].width=160
    wb.calculation.fullCalcOnLoad=True
    wb.save(OUT)
    nc=sum(1 for e in deliveries if e['cancelled']); npast=sum(1 for e in deliveries if e['stage']=='Past due')
    print(f'Saved: {OUT}')
    if MIS:      # ---------- the new MRP (MIS + WSA only) ----------
        try:
            outn=MRP_DAILY_OUT; prev=PREV_MRP                         # Settings were already read at the start of the run
            prepare_mis(MIS,USER_MRP,prev)                             # MIS without Total Plan Req / Daily Dashboard: rebuilt from your plan + Standard BOM
            SUBS_ISSUES[:]=resolve_substitutes(MIS) if SUBS else []
            MIS_PATH_[0]=MIS; DSn=build_daily_stock(MIS); attach_wsa(DSn,BW)
            rn=build_new_mrp(DSn,BW,MIS,outn,prev_path=prev,TD=TARGET_DAYS,old_mrp=USER_MRP)
            if CAP[0]:                                         # the WSA calculation for the BOM materials without a WSA sheet -> Inbound
                try:
                    al_=CAP[0].get('alias') or {}
                    cov_={al_.get(c_,c_) for g_ in G for c_ in g_['codes']}
                    om_=other_materials_calc(CAP[0],cov_)
                    wof_={}
                    for g_ in G:
                        for c_ in g_['codes']:
                            if g_['label'] not in wof_.setdefault(al_.get(c_,c_),[]): wof_[al_.get(c_,c_)].append(g_['label'])
                    add_other_materials(OUT,om_,CAP[0].get('ppc'),wof_)
                    nb_=sum(1 for x in om_ if x['per']['pi'] or x['per']['pd'])
                    print(f"Other materials (no WSA sheet): {len(om_)} calculated the WSA way - Inbound sheet 'Other materials' ({nb_} go below zero; periods also on 'Below zero')")
                except Exception as ex_: print(f'NOTE: Other materials not added to the Inbound file ({type(ex_).__name__}: {ex_})')
            if ALERT_ROWS or ALERT_D0[0] is not None:            # the same ALERT sheet as MRP_Daily: first tab of the Inbound file (WSA daily)
                try: add_alert_inbound(OUT,ALERT_ROWS,ALERT_D0[0])
                except Exception as ex_: print(f'NOTE: ALERT sheet not added to the Inbound file ({type(ex_).__name__}: {ex_})')
            print(f"NEW MRP saved: {outn} - {rn['crit']} critical, {rn['risk']} at risk, {rn['watch']} watch, {rn['ok']} OK, {rn['nodem']} no demand; "
                  f"Check: {rn['check'].get('High',0)} high issues; CS / PM: {rn['cs']} components from {rn['bom_rows']} BOM rows"
                  +(f" (BOM taken from {os.path.basename(PPC_SRC[0]) if PPC_SRC[0] else os.path.basename(rn['bom_src'])})" if rn['bom_src'] else ' - no BOM yet: fill the BOM sheet'))
            # the new MRP_Daily is saved: delete the older MRP_Daily files (and their _report.txt) in this folder
            gone_=[]; kept_=[]
            for n_ in ([] if KEEP_OLD else sorted(os.listdir(os.path.dirname(os.path.abspath(outn))))):
                fp_=os.path.join(os.path.dirname(os.path.abspath(outn)),n_)
                if os.path.abspath(fp_)==os.path.abspath(outn) or n_.startswith('~$'): continue      # old _report.txt files go too
                if re.fullmatch(r'MRP_Daily_\d{4}-\d{2}-\d{2}(_report\.txt|\.xlsx)',n_):
                    try: os.remove(fp_); gone_.append(n_)
                    except OSError: kept_.append(n_)
            if gone_: print('Old MRP_Daily deleted: '+', '.join(gone_))
            if kept_: print('NOTE: could not delete (open in Excel?): '+', '.join(kept_))
            try:                                                     # the daily email to top management (user, Nov 2026)
                dir_=os.path.dirname(os.path.abspath(outn)); dms_=os.path.join(dir_,f'Daily_Material_Status_{dt.date.today():%Y-%m-%d}.xlsx')
                write_daily_email(ALERT_ROWS,ALERT_D0[0],dir_,attach=dms_ if os.path.exists(dms_) else None)
            except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: daily email not written ({type(ex_).__name__}: {ex_})')
        except _CaptureDone: pass                                    # step 1 of the alignment: numbers captured, no MRP_Daily needed
        except PermissionError:
            print('WARNING: MRP_Daily file is open in Excel - close it and run again.')
        except SystemExit: raise
        except Exception as ex_:                     # never leave half-built files; the Inbound file is already saved
            import traceback
            tb_=traceback.extract_tb(ex_.__traceback__)[-1]
            print(f'ERROR building MRP_Daily: {type(ex_).__name__}: {ex_} (line {tb_.lineno}, {tb_.name}). The Inbound file was saved; the previous MRP_Daily is unchanged.')
        finally:
            for f_ in glob.glob(os.path.join(OUTDIR,'MRP_Daily_*.xlsx.*.xlsx')):
                try: os.remove(f_)
                except OSError: pass
    print(f'{len(deliveries)} deliveries -> {len(rows)} rows ({nc} cancelled, {npast} past due, {sum(1 for e in deliveries if e["stage"] in ("In transit","Open PO"))} counted), {len(mats)} materials, report date {STOCK_DATE:%d-%b-%Y}; took {time.time()-t0:.0f} s')
    print('Materials with no scheduled delivery: '+(', '.join(sorted({n for n,gg,e,sh,fi in rows if e.get('placeholder')})) or 'none'))
    print(f'To chase: {len(chase)} deliveries (past due {npast}, missing PO/shipper/ETD {sum(1 for e in chase if e["missing"])}, ETD passed with no vessel/BL {sum(1 for e in chase if e["etdpassed"])}, short lead {sum(1 for e in chase if e["datecheck"]=="Short lead")}, MIS arrival 3+ days from the plant date {sum(1 for e in chase if e.get("misdiff") is not None)} - MIS is the latest)')
    # (the old By Week / MRP_<date> update was retired: your MRP.xlsx is read only, MRP_Daily is the output)
    if prev_del:
        cnt=Counter(t[0] for t in out); print(f'Changes vs {os.path.basename(PREV)}: '+', '.join(f'{v} {k.lower()}' for k,v in cnt.items()) if cnt else f'Changes vs {os.path.basename(PREV)}: none')
    if ALIGN_WSA:
        _QUIET[0]=False; _aligned=None; _t1=time.time()
        if ALIGN[0] is None: print('ERROR: the WSA file was not aligned - MRP_Daily could not be calculated (see the ERROR above)')
        else:
            try:
                _aligned,_rep=align_wsa(SRC,ALIGN[0],name_from=SRC_USER)
                if _aligned:
                    try:
                        _ns=sort_wsa_dashboard(_aligned,base_d0())
                        if _ns: print(f'WSA Dashboard sorted highest risk first ({_ns} rows moved; the same order as the RM Material Dashboard)')
                    except Exception as _e2: print(f'NOTE: WSA Dashboard not sorted ({type(_e2).__name__}: {_e2})')
            except SystemExit: raise
            except Exception as ex_:
                import traceback
                tb_=traceback.extract_tb(ex_.__traceback__)[-1]
                print(f'ERROR aligning the WSA file: {type(ex_).__name__}: {ex_} (line {tb_.lineno}, {tb_.name}) - Inbound and MRP_Daily are built from your file as it is')
        try: _b4=_ref_formula_counts(SRC)                             # your file as it was (a copy in the temp folder; your own file is replaced now)
        except Exception as _e_: _b4={}; _swallow(0,_e_)
        _release(SRC); shutil.rmtree(_ALIGN_TMP,ignore_errors=True)
        print('ALIGN WSA - step 2 of 2: Inbound and MRP_Daily from the aligned WSA file ...' if _aligned else 'Inbound and MRP_Daily from your WSA file ...',flush=True)
        if not _aligned and not os.path.exists(SRC_USER) and os.path.exists(os.path.join(os.path.dirname(os.path.abspath(SRC_USER)),WSA_OLD_DIR,os.path.basename(SRC_USER))):
            SRC_USER=os.path.join(os.path.dirname(os.path.abspath(SRC_USER)),WSA_OLD_DIR,os.path.basename(SRC_USER))   # moved by _wsa_tidy
        _cmd=[sys.executable,os.path.abspath(__file__),_aligned or SRC_USER,'--no-align','--mis',MIS]+(['--mrp',USER_MRP] if USER_MRP else [])
        _t2=time.time(); sys.stdout.flush(); _rc=subprocess.call(_cmd); _t3=time.time()
        shutil.rmtree(ROWS_CACHE_DIR,ignore_errors=True)
        if _aligned:
            try:
                _af=_ref_formula_counts(_aligned)
                _new={k_:v_-_b4.get(k_,0) for k_,v_ in _af.items() if v_>_b4.get(k_,0)}
                if _new: print('WARNING: the alignment created #REF! formulas on: '+'; '.join(f'{k_} (+{v_})' for k_,v_ in sorted(_new.items()))+' - a kept sheet points at a sheet that was left out')
            except Exception as _e_: _swallow(0,_e_)
            if WSA_LOCKED[0]:                                          # today's WSA_Daily was open in Excel: saved under a second name
                print(f"WSA DAILY SAVED: {_aligned} (today's WSA_Daily was open in Excel - {WSA_LOCKED[0][1]}; close it to get one file next run). Your WSA file was only read.")
            else: print(f"WSA DAILY SAVED: {_aligned} - your WSA file {os.path.basename(SRC_USER)} was only read, not changed")
            if WSA_MOVED: print("  Deleted (older WSA_Daily files): "+', '.join(WSA_MOVED))
            print(f"  Each planning sheet opens at today's row ({base_d0():%d-%b}; Balance in yellow = stock typed by the script). Rows before today are as you typed them.")
            print(f"  The module plan x BOM behind the formulas is the last sheet, {ALIGN_SHEET}. "+f"Keep typing in your own file {os.path.basename(SRC_USER)}; WSA_Daily is rebuilt from it on every run.")
            P_=(ALIGN[0] or {}).get('ppc') or {}; k_,n_=ALIGN_STAT.get('ppc_sheets') or (0,0)
            if P_.get('arr'): print(f"  Column T 'PPC In Transit (arrivals)': {len(P_.get('rows') or [])} lines from {P_.get('src')} (sheet {P_.get('sheet') or 'In_Transit'}) - arrivals on {k_} of {n_} planning sheets")
            else: print("WARNING: column T 'PPC In Transit (arrivals)' is EMPTY on every planning sheet - "+'; '.join(P_.get('why') or ['no PPC In_Transit data']))
        if not _aligned:                                              # user, Nov 2026: the messy dashboards came from a run like this one - say it where it is read
            print('#'*100+"\nWARNING: NO WSA_Daily FILE WAS WRITTEN IN THIS RUN (see the ERROR line above).\n"
                  "  The outputs of this run were built from your WSA file as it is; 'RM Material Dashboard - MIS' is left out of Daily_Material_Status.\n"
                  "  Close today's WSA_Daily file in Excel and run the script again.\n"+'#'*100,flush=True)
        print(f"Timing: step 1 (calculation) {_t1-t0:.0f} s, aligned WSA file {_t2-_t1:.0f} s, step 2 (Inbound + MRP_Daily) {_t3-_t2:.0f} s - total {_t3-t0:.0f} s",flush=True)
        _x=[]
        if (ALIGN[0] or {}).get('ppc'):
            _p=(ALIGN[0] or {}).get('ppc') or {}
            if _p.get('skipped'): _x.append(f"{len(_p['skipped'])} PPC In_Transit line(s) not mapped to a BOM material, NOT counted: "+'; '.join(_p['skipped'][:6]))
            _k,_n=ALIGN_STAT.get('ppc_sheets') or (0,0)
            if _n and _k<_n: _x.append(f'PPC arrivals reach only {_k} of {_n} planning sheets ({_n-_k} sheet(s) get no arrivals)')
            _chk=[l_ for l_ in (_p.get('rows') or []) if 'CHECK' in str(l_.get('how'))]
            if _chk: _x.append(f'{len(_chk)} In_Transit line(s) marked CHECK (your WSA sheet has another size / HS) - see Inbound sheet PPC In Transit')
        if WSA_DAMAGED[0]:
            _dm=f"YOUR WSA FILE {os.path.basename(WSA_DAMAGED[0][0])} IS DAMAGED (a write failed half-way and could not be undone): your file as it was = {WSA_DAMAGED[0][1]} - copy it over your WSA file (Excel closed), or use OneDrive version history"
            print('#'*100+'\nWARNING: '+_dm+'\n'+'#'*100,flush=True)
        if _aligned: _x.insert(0,f"WSA DAILY SAVED: {os.path.basename(_aligned)} (your WSA file {os.path.basename(SRC_USER)} only read)")
        else: _x.insert(0,"NO WSA_Daily file written in this run (see the ERROR line) - close today's WSA_Daily in Excel and run again")
        if WSA_DAMAGED[0]: _x.insert(0,'WARNING: '+_dm)
        _x.append('WSA file: only the planning sheets were rebuilt; other sheets of your WSA file (Summary, Dashboard, solar cell, JB ...) are carried over, not recalculated')
        run_summary('step 1 - aligned WSA file',extra=_x)
        sys.exit(_rc)
