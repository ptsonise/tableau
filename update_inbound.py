"""
=================================================================
  MRP Daily + Inbound  —  one script for the planner
=================================================================

WHAT IT DOES
  Reads your daily WSA file, your MRP.xlsx, and the MIS RM Dashboard.
  Writes two files:
    - MRP_Daily_<today>.xlsx   (Start Here, Action List, MRP Daily, Planner, Today, Check, ...)
    - Inbound_<today>.xlsx     (deliveries, to-chase, changes)

HOW TO RUN
  Put these three files in the same folder:
    1. WSA - Raw Material planning per day.xlsx
    2. MRP.xlsx
    3. MIS RM Dashboard - WSA.xlsx

  Then double-click or run:
    python update_inbound.py

  Or with the file names:
    python update_inbound.py "WSA - Raw Material planning per day.xlsx" MRP.xlsx "MIS RM Dashboard - WSA.xlsx"

WHAT YOU CAN CHANGE
  Everything is on the 'Settings' sheet (copy it into your MRP.xlsx once and edit it there). It is read FIRST, before
  anything is calculated, so the Inbound file and MRP_Daily always use the same rules. The CONFIG values below are only
  the defaults used when there is no Settings sheet yet.

STATUS RULES (one rule each, used on every sheet - see material_status() and line_status())
  - Material status: compares days of cover with that material's OWN target days (A / B / C / packaging, + Z extra):
      CRITICAL = runs out within its target even with Open PO
      AT RISK  = runs out within its target unless the Open PO ships
      WATCH    = covered only if in transit arrives on time, or falls below target within its lead time
      OK       = everything else
  - Line status: days a production line can run vs COVER_CRITICAL / COVER_RISK / COVER_WATCH (Settings).

Needs: Python 3.9+, pandas, openpyxl, lxml
"""
import sys, os, re, glob, time, datetime as dt, warnings
_T0=time.time()     # start of the run - progress lines show seconds since start
warnings.filterwarnings('ignore')
import pandas as pd
# SPEED: when a workbook is opened read-only (to read values), only number formats are needed (to recognise dates).
# openpyxl normally rebuilds every font / fill / named style first - the WSA file has ~36,000 named styles, which took
# about 70% of the run time. Read-only loads now parse just the number formats; normal loads (copying formats) are unchanged.
import openpyxl.reader.excel as _oxl_reader
from openpyxl.styles.stylesheet import Stylesheet as _OxlStylesheet
from openpyxl.xml.functions import fromstring as _oxl_fromstring
_oxl_apply_full=_oxl_reader.apply_stylesheet
def _oxl_apply_fast(archive, wb):
    if not getattr(wb,'_read_only',False): return _oxl_apply_full(archive, wb)
    try: node=_oxl_fromstring(archive.read('xl/styles.xml'))
    except KeyError: return wb
    for ch in list(node):
        if ch.tag.rsplit('}',1)[-1] not in ('numFmts','cellXfs'): node.remove(ch)
    ss=_OxlStylesheet.from_tree(node)
    wb._cell_styles=ss.cell_styles; wb._number_formats=ss.number_formats; wb._date_formats=ss.date_formats
    wb._timedelta_formats=ss.timedelta_formats; wb._alignments=ss.alignments; wb._protections=ss.protections
    return wb
_oxl_reader.apply_stylesheet=_oxl_apply_fast
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
COVER_RISK      = 21     # PRODUCTION LINES: AT RISK (expedite) if fewer days than this
COVER_WATCH     = 28     # PRODUCTION LINES: WATCH if fewer days than this
SHORT_WATCH     = 14     # Today sheet: also show a material that is OK / WATCH but more than this many days short of its target
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
SAFETY_DAYS=7      # safety stock = days of need at the daily targets (Stock Management reorder level) - Settings
LOOKAHEAD_DAYS=49  # a line counts for a material (daily target need, 'Needed now') if it uses it within these days - Settings
REPORT_DAYS=21     # daily report: materials whose stock + in transit lasts fewer days than this at the line daily targets
SHARED_MATS=[]     # materials used by 2+ lines this run: (name, description, code, [lines]) - listed on the Settings sheet
LINES_SEEN=[]      # production line names found in this run (to list them on the Settings sheet)
# =================================================================

# =================================================================
#  STATUS RULES  -  the ONLY place statuses are decided. Every sheet (Action List, Today, Planner, MRP by Line,
#  Line Run Days, Supplier Coverage, Start Here, e-mails) calls these, so a material / line has one status everywhere.
# =================================================================
STATUS_ORDER  = {'CRITICAL':0,'AT RISK':1,'WATCH':2,'OK':3,'NO DEMAND':4}
STATUS_FILL   = {'CRITICAL':'FFC7CE','AT RISK':'FFEB9C','WATCH':'FFF2CC','OK':'C6EFCE','NO DEMAND':'EDEDED'}
STATUS_FONT   = {'CRITICAL':'9C0006','AT RISK':'9C5700','WATCH':'7F6000','OK':'006100','NO DEMAND':'595959'}
STATUS_ACTION = {'CRITICAL':'Escalate today - re-sequence or air freight','AT RISK':'Expedite - confirm vessel / pull shipment',
                 'WATCH':'Watch - place or advance order','OK':'Monitor','NO DEMAND':'No demand - no action'}
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
def line_status(days):
    """days a production line can run (stock + in transit) vs the Settings line thresholds."""
    if days<COVER_CRITICAL: return 'CRITICAL'
    if days<COVER_RISK: return 'AT RISK'
    if days<COVER_WATCH: return 'WATCH'
    return 'OK'

# ---------- read-only workbook cache: each input workbook / sheet is parsed once per run (big speed-up) ----------
_WB_CACHE={}
class _CachedSheet:
    def __init__(self,rows,title): self._rows=rows; self.title=title; self.max_row=len(rows); self.max_column=max((len(r) for r in rows),default=0)
    def iter_rows(self,min_row=1,max_row=None,min_col=1,max_col=None,values_only=True):
        for r in self._rows[(min_row or 1)-1:max_row]:
            yield r[(min_col or 1)-1:max_col] if (max_col or (min_col or 1)>1) else r

USER_MB52=[None]   # your SAP MB52 export (same folder): the stock quantities; MIS MB52 only gives the extra columns (plant name, final code, key, company, categories)
def _is_mis_file(p):
    try: return {'MB52','Standard BOM'}<=set(_sheetnames(p))
    except Exception: return False
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
            except Exception: continue
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
    except Exception: pass
    try:
        for r in wbc['Plant Master'].iter_rows(min_row=2,values_only=True):
            if r and r[0] and len(r)>1 and r[1]: pl_name.setdefault(str(r[0]).strip().split('.')[0],r[1])
    except Exception: pass
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

class _CachedWB:
    def __init__(self,path):
        self._wb=load_workbook(path,read_only=True,data_only=True); self.sheetnames=list(self._wb.sheetnames); self._s={}; self._path=path
    def __getitem__(self,name):
        if name not in self._s:
            if name not in self.sheetnames:          # a sheet missing from an input file: warn once and carry on with no rows
                print(f"NOTE: sheet '{name}' not found in {os.path.basename(self._path)} - that part is skipped")
                self._s[name]=_CachedSheet([],name); return self._s[name]
            rows=[tuple(r) for r in self._wb[name].iter_rows(values_only=True)]
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
IT_FIX=[]          # shipments whose MIS Final RM Code was replaced by the shipped RM Code (listed on Check)
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
       2) MIS 'Final RM Code' comes from the PO; when the SHIPMENT description clearly matches the shipped 'RM Code' better
          (e.g. 2272 glass booked on a 2376 PO, 35HS frame on a 30HS PO), the shipped RM Code is used instead."""
    hdr=[str(h).strip() if h is not None else '' for h in rows[0]]
    pos={h:i for i,h in enumerate(hdr) if h}
    IT_MISSING[:]=[h for h in IT_CANON if h not in pos]
    extra=[i for i,h in enumerate(hdr) if h and h not in IT_CANON]
    out=[tuple(IT_CANON+[hdr[i] for i in extra])]
    for r in rows[1:]:
        out.append(tuple([r[pos[h]] if (h in pos and pos[h]<len(r)) else None for h in IT_CANON]+[r[i] if i<len(r) else None for i in extra]))
    desc={}
    try:
        for r in wbc['Common Name'].iter_rows(min_row=2,values_only=True):
            if r and r[0] and r[1]: desc[str(r[0]).strip()]=str(r[1])
    except Exception: pass
    try:                                                                  # descriptions of codes that are not on Common Name (MB52)
        for r in wbc['MB52'].iter_rows(min_row=2,values_only=True):
            if r and len(r)>5 and r[2] and r[5]: desc.setdefault(str(r[2]).strip(),str(r[5]))
    except Exception: pass
    iR,iF,iD,iQ=IT_CANON.index('RM Code'),IT_CANON.index('Final RM Code'),IT_CANON.index('Description'),IT_CANON.index('Final Pending Qty.')
    iS,iPW=IT_CANON.index('Supplier Name'),IT_CANON.index('PO Wise Material Desc.')
    supn=lambda v: ' '.join(str(v or '').upper().split())[:18]
    by_sup={}                                                             # supplier -> codes it ships (from all lines)
    for r in out[1:]:
        for c_ in (r[iR],r[iF]):
            if c_ and str(c_).strip() in desc: by_sup.setdefault(supn(r[iS]),set()).add(str(c_).strip())
    IT_FIX.clear(); fixed=[out[0]]
    for r in out[1:]:
        rc=str(r[iR] or '').strip(); fc=str(r[iF] or '').strip(); dsc=r[iD]
        if not dsc: fixed.append(r); continue
        cands={c_ for c_ in (rc,fc) if c_ in desc}|{c_ for c_ in re.findall(r'\b(?:RM|PM|CS)[A-Z0-9]{6,}\b',str(r[iPW] or '').upper()) if c_ in desc}
        if not (fc in desc): cands|=by_sup.get(supn(r[iS]),set())     # blank / unknown code: materials this supplier ships
        if not cands: fixed.append(r); continue
        sc={c_:_desc_score(dsc,desc[c_]) for c_ in cands}
        best=max(sc,key=sc.get); cur=fc if fc in desc else None
        second=max([v for c_,v in sc.items() if c_!=best] or [-99])
        if best!=cur and (cur is None and sc[best]>0 and sc[best]>second or cur is not None and sc[best]>sc[cur]+1):
            r=list(r); r[iF]=best; r=tuple(r)                             # the description clearly fits another code of this line / PO / supplier
            IT_FIX.append(dict(desc=str(dsc).replace('\n',' ')[:80],shipped=best,mis=fc or '(blank)',qty=r[iQ] if isinstance(r[iQ],(int,float)) else 0,
                               po=r[IT_CANON.index('PO')],sup=r[iS],cd_s=desc[best],cd_m=desc.get(fc,'')))
        fixed.append(r)
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
    except Exception: pass
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
            try: n=set(_sheetnames(f))
            except Exception: continue
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
import builtins as _bi
def print(*a, **k):                     # every NOTE / WARNING / ERROR line is also kept for the 'Start Here' sheet
    m=' '.join(str(x) for x in a)
    if m.startswith(('NOTE','WARNING','ERROR')): RUN_NOTES.append(m)
    _bi.print(*a, **k)
def die(m): print('ERROR:',m); sys.exit(1)
ARGS=sys.argv[1:]; PREV=None; MRP=None; MRP_OUT=None; MIS=None
NO_MRP='--no-mrp' in ARGS
if NO_MRP: ARGS.remove('--no-mrp')
for flag in ('--prev','--mrp','--mrp-out','--mis'):
    if flag in ARGS:
        i=ARGS.index(flag); val=ARGS[i+1] if i+1<len(ARGS) else die(f'{flag} needs a file'); del ARGS[i:i+2]
        if flag=='--prev': PREV=val
        elif flag=='--mrp': MRP=val
        elif flag=='--mis': MIS=val
        else: MRP_OUT=val
def _sheetnames(p):
    try: return wb_ro(p).sheetnames
    except Exception: return []
if not ARGS:      # double-click / no file names: newest WSA daily file (has a Total_Plan_Req sheet) in this folder or the script's folder
    _cands=[]
    for _f in dict.fromkeys([os.getcwd(),os.path.dirname(os.path.abspath(sys.argv[0]))]):
        for _n in os.listdir(_f):
            _nl=_n.lower()
            if _nl.endswith('.xlsx') and not _n.startswith('~$') and not _nl.startswith(('mrp','inbound_')) and 'mis' not in _nl and 'dashboard' not in _nl:
                _cands.append(os.path.join(_f,_n))
    _cands=[p for p in sorted(_cands,key=os.path.getmtime,reverse=True) if 'Total_Plan_Req' in _sheetnames(p)]
    if not _cands: die('No WSA daily file found here. Usage: python update_inbound.py <WSA daily file.xlsx> [MRP.xlsx] [MIS RM Dashboard.xlsx]')
    ARGS=[_cands[0]]; print(f'WSA daily file: {_cands[0]}')
def _is_mrp(p):
    n=_sheetnames(p)            # the MRP workbook has a By Week sheet; an Inbound output also has one, but it has an Inbound sheet too
    return 'By Week' in n and 'Inbound' not in n and 'Total_Plan_Req' not in n
def _is_mis(p):
    n=_sheetnames(p); return 'MB52' in n and ('Daily Dashboard' in n or 'Standard BOM' in n)
for a_ in list(ARGS):                     # an MIS RM Dashboard given in any position
    if MIS is None and a_.lower().endswith('.xlsx') and os.path.exists(a_) and _is_mis(a_): MIS=a_; ARGS.remove(a_)
if not ARGS: die('Usage: python update_inbound.py <daily_file.xlsx> [MRP.xlsx] [MIS_RM_Dashboard.xlsx] [output.xlsx] [--prev ...]')
if len(ARGS)>=2 and os.path.exists(ARGS[0]) and os.path.exists(ARGS[1]):      # accept the daily file and the MRP workbook in either order
    n0,n1=_sheetnames(ARGS[0]),_sheetnames(ARGS[1])
    if 'Total_Plan_Req' not in n0 and _is_mrp(ARGS[0]) and 'Total_Plan_Req' in n1: ARGS[0],ARGS[1]=ARGS[1],ARGS[0]
SRC=ARGS[0]
if not os.path.exists(SRC): die(f'File not found: {SRC}')
OUT=None
def _is_user_mrp(p):
    n=_sheetnames(p); return 'Production Plan' in n and ('BOM' in n or 'Settings' in n) and not os.path.basename(p).startswith('MRP_Daily_')
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
PPC_SRC[0]=find_ppc(list(dict.fromkeys([os.path.dirname(os.path.abspath(USER_MRP)) if USER_MRP else os.getcwd(),OUTDIR,os.path.abspath(os.getcwd()),os.path.dirname(os.path.abspath(__file__))])))
if PPC_SRC[0]: print(f'BOM and Stock Count read from: {os.path.basename(PPC_SRC[0])} (sheets BOM, Count'+(", module plan from Plan + Revised Plan)" if 'Plan' in _sheetnames(PPC_SRC[0]) else ')'))
else: print('WARNING: no PPC workbook (sheets Count + BOM) in the folder - physical stock counts come ONLY from the PPC Count sheet, so this run uses MB52 stock for every material; BOM is read from MRP.xlsx')
USER_MB52[0]=find_user_mb52(list(dict.fromkeys([os.path.dirname(os.path.abspath(USER_MRP)) if USER_MRP else os.getcwd(),OUTDIR,os.path.abspath(os.getcwd()),os.path.dirname(os.path.abspath(__file__))])))
if not USER_MB52[0]: print('NOTE: no MB52.xlsx (your SAP export) found in the folder - stock is read from the MIS MB52 sheet')
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
        if sh in SKIP: continue
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
                    ('actmw',lambda h:h.startswith('actual mw per day')),('plantdays',lambda h:h.startswith('plant days')),('plantport',lambda h:h.startswith('plant + port')),('delivered',lambda h:h.startswith('delivered'))]
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
            pdays_by_date={d:(info(r).get('plantdays'),info(r).get('plantport')) for r,d in rows}
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
            groups.append(dict(label=label,sheet=sh,block=bi,uom=uom,codes=sorted(codes),asof=asof,inbound=inbound,last_receipt=last_rec,asof_info=info(last[0]),bal_by_date=bal_by_date,dem_by_date=dem_by_date,pdays_by_date=pdays_by_date))
    return groups,skipped

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
    except Exception: pass
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
    d0=pd.Timestamp(dt.date.today()).normalize(); dend=d0+pd.Timedelta(days=ndays)
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
    except Exception: return out
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
    start=pd.Timestamp(start or dt.date.today()).normalize(); ndays=ndays or ncols_dash or 141
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
    # supplier detail: stock by Make (RM Status for OEM JW: USA + USA Warehouse + JW) and in transit by Supplier Name
    mk_stock={}
    for mk,cd,a_,b_,c_ in zip(bycol(rs,0),bycol(rs,1),num(bycol(rs,5)),num(bycol(rs,7)),num(bycol(rs,11))):
        if not isinstance(mk,str) or not mk.strip(): continue
        d_=mk_stock.setdefault(str(cd).strip(),{}); d_[mk.strip().upper()]=d_.get(mk.strip().upper(),0.0)+a_+b_+c_
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
        b['pack_window']=[dict(w_,qty=(raw_stock.get(w_['code'],0.0)+raw_jw.get(w_['code'],0.0))*w_['factor']) for w_ in pack_window.get(m,[])]
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
    except Exception: pass

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
    plan_src=("Plan = your Production Plan modules x usage (RM: MIS, CS / PM: your BOM). ")
    if weekly_of: plan_src=(f"Weekly view of '{weekly_of}' (live formulas). Plan / Intransit / Open PO / Adhoc = week totals; Stock and Days of cover = end of week. "+plan_src)
    c=ws.cell(1,4,f"Updated {dt.date.today():%d-%b-%Y} from {DS['src']}. Line rows (Customer + Production Line): column I = where that line is supplied from, red cells = days that line is not covered. "+plan_src+f"Stock (RM) = MB52 {DS['plant']} + {DS['plant']} Warehouse + JW + MIS code conversions; CS / PM = your Stock Count. "
              "Stock row = projected stock including in transit and adhoc. Open PO (no vessel yet, WSA) is NOT in the Stock row; "
              f"'Days of cover (stock + in transit + Open PO)' adds it. Days of cover = projected stock / average demand of the material's own target days (Settings); colours: 0 red, half the target yellow, target green. "
              "Glass: supplier rows show each supplier's stock (MIS RM Status by Make), in transit and Open PO.")
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
        k_=b['key']+(sup if sup and sup!='All suppliers' else '')+(f'|{cust.strip()}' if cust and cust!='All customers' else '')+(f'|{line}' if line and not str(line).startswith('All lines') else '')
        for j,v in enumerate([k_,b['sr'],b['short'],b['common'],b['mat'],b['desc'],cust.strip() if cust else None,line,sup,hd],1): ws.cell(rr,j,v)
        ws.row_dimensions[rr].height=20.1
    # Cached results: formulas are saved together with the value Excel would calculate, so phone / e-mail previews (which do not
    # recalculate) show the numbers, and the colours are also painted as plain fills because previews ignore conditional formats.
    if cache is None: cache={}
    def C(rr,cc,v): cache[(rr,cc)]=v
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
        rp=r; rc=list(range(r+1,r+1+NC)); rs_,rt,ro,ri,rj=range(r+1+NC,r+6+NC); ra=None
        rPR=rj+1
        opo=b.get('openpo') or [0.0]*len(dates)
        inv_lab=('Days of cover (stock + in transit)','Days of cover (stock + in transit + Open PO)')
        supname='All suppliers' if b.get('sup') else None
        # standard MRP record (gross requirement -> receipts -> projected available balance -> net requirement -> planned orders)
        ident(rp,b,supname,'Gross requirement','All customers' if NC else None,'All lines' if any(ln for _,ln in custs) else None)
        for rr,(cu,ln) in zip(rc,custs):
            src_=(b.get('line_src') or {}).get((_norm(cu),_line_key(ln))) if ln else None
            if src_: src_=src_.replace(b['mat'],'own stock')
            ident(rr,b,(f'from: {src_}' if (src_ and src_ not in ('-','own stock')) else supname),'Gross requirement',cu,ln)
        for rr,hd in zip((rs_,rt,ro,ri,rj),('Projected available balance','In transit (MIS) - by arrival date','Firm receipt (Open PO, no vessel)')+inv_lab): ident(rr,b,supname,hd)
        ident(rPR,b,supname,'Planned order receipt')
        stock_rows.append(rs_); inv_rows+=[ri,rj]; inv_td[ri]=inv_td[rj]=TD
        if weekly_of:
            for rr in [rp]+rc+[rs_,rt,ro,ri,rj,rPR]: ws.cell(rr,SD,f'={Dq}{SDL}{rr}')
            for w,(d,a,z) in enumerate(per):
                col=C0+w; A_,Z_=DL(a),DL(z); rng=lambda rr: f'{Dq}{A_}{rr}:{Z_}{rr}'
                for rr in [rp]+rc+[rt,ro]: ws.cell(rr,col,f'=SUM({rng(rr)})')
                for rr in (rs_,ri,rj): ws.cell(rr,col,f'={Dq}{Z_}{rr}')     # end-of-week values from the daily MRP
                ws.cell(rPR,col,f'=SUM({rng(rPR)})')
        else:
            for rr in [rp]+rc: ws.cell(rr,SD,f'=SUM({L_(C0)}{rr}:{L_(CL)}{rr})')
            if b.get('count_stock'):   # new MRP: CS / PM stock = Today Qty on its own Stock Count sheet (live)
                ws.cell(rs_,SD,f"=SUMIFS('Stock Count'!${SC_COL}:${SC_COL},'Stock Count'!$A:$A,$E{rs_})")
            else: ws.cell(rs_,SD,round(b['stock'],6) if VALUES_ONLY else f"=SUMIFS('Stock Data'!$L$5:$L$3000,'Stock Data'!$A$5:$A$3000,$E{rs_})")   # RM: MB52 + JW lines on Stock Data
            ws.cell(rt,SD,f'=SUM({L_(C0)}{rt}:{L_(CL)}{rt})'); ws.cell(ro,SD,f'=SUM({L_(C0)}{ro}:{L_(CL)}{ro})');      # total in transit, like Open PO ws.cell(ri,SD,0); ws.cell(rj,SD,0)
            ws.cell(rPR,SD,f'=SUM({L_(C0)}{rPR}:{L_(CL)}{rPR})')
            for k in range(ND):
                col=C0+k; cl=L_(col); pv=L_(col-1)
                ws.cell(rp,col,round(b['plan'][k],6))
                for rr,cu in zip(rc,custs): ws.cell(rr,col,round(CPx(cu)[k],6))
                ws.cell(rt,col,round(b['transit'][k],6)); ws.cell(ro,col,round(opo[k],6))
                ws.cell(rs_,col,f'={pv}{rs_}-{cl}{rp}+{cl}{rt}')
                # forward-looking: stock / average demand of the next TD days (blank when no demand ahead)
                ce=L_(min(col+COVER_AVG_DAYS-1,CL)); win=f'{cl}{rp}:{ce}{rp}'
                ws.cell(ri,col,f'=IF(SUM({win})=0,"",{cl}{rs_}/AVERAGE({win}))')
                ws.cell(rj,col,f'=IF(SUM({win})=0,"",({cl}{rs_}+SUM(${L_(C0)}{ro}:{cl}{ro}))/AVERAGE({win}))')
                ws.cell(rPR,col,round((b.get('por') or [0.0]*ND)[k],3))
        style([rp]+rc+[rs_,rt,ro,ri,rj,rPR],lambda rr,j: bold if rr in (ri,rj,rPR) and j>=10 else (ital if rr in rc else body))
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
        gap=b['stock']-sum(makes.values())
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
    except Exception: return out
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
        except Exception: return []
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
def load_user_masters(mrp_path):
    """Your Batch Master and Material Group sheets (MRP.xlsx)."""
    BATCH_MASTER[0].clear(); BATCH_MASTER[1].clear(); MAT_GROUP.clear()
    try: w=wb_ro(mrp_path)
    except Exception: return
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
    except Exception: return by
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
    except Exception: pass
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
       how many extra days they would give at the line daily targets. Advice only: approve, then add the pair on Settings > Substitutes."""
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
        except Exception: pass
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
            own=(x.get('tgt_need') or x.get('avg') or 0)*(x['b'].get('td') or REPORT_DAYS)
            p_['free']=max(0.0,p_['stock']+p_['it']-own); p_['use']=f"Used by plan - {x['status']}, lasts {x['tgt_days']:.0f} d at targets" if x.get('tgt_days') is not None else f"Used by plan - {x['status']}"
        else: p_['free']=max(0.0,p_['stock']+p_['it']); p_['use']='Not in the plan - free stock'
    # YOUR SAP BOM (MRP.xlsx) vs MIS Standard BOM: same customer (Settings MIS link -> MIS BOM name), same material group
    mis_bn_fam={}
    try:
        for r_ in list(wb_ro(mis_path)['Standard BOM'].iter_rows(values_only=True))[1:]:
            if r_ and len(r_)>3 and r_[2] and r_[3]:
                c2=str(r_[3]).strip(); f2=(pool.get(c2) or {}).get('fam') or _fam(c2,cn.get(c2),bom,sapg.get(c2),desc.get(c2))
                mis_bn_fam.setdefault((str(r_[2]).strip(),f2),set()).add(c2)
    except Exception: pass
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
              'the plan needs of the candidate itself for its target days. Extra days = free qty / need per day of the material at the line daily targets (blank when the units differ - convert first). '
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
        need_=x.get('tgt_need') or x.get('avg') or 0
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
            if c_ in (ALT_BOM.get(b['mat']) or {}): g='A'; notes=[f"MIS YOUR ALTERNATIVE BOM (MRP.xlsx) for {', '.join(sorted(ALT_BOM[b['mat']][c_]))}"]+notes
            sa_=[z_ for z_ in sap_alt.get(b['mat'],[]) if z_[0]==c_]
            if sa_: g='A'; notes=[f"YOUR SAP BOM (MRP.xlsx, CS/PM BOM) uses this code for {', '.join(dict.fromkeys(z_[1] for z_ in sa_))} ({sa_[0][2]}, valid from {sa_[0][3]}) where the MIS BOM has {b['mat']}"]+notes
            if plan_jb and c_ in plan_jb: g='A'; notes=[f"MIS Total Plan Req 'JB Type' says the module(s) {', '.join(sorted(plan_bn))} use JB {'/'.join(sorted(plan_jbt))} = this code"]+notes
            ts_={k for k,q in (bsup.get(b['mat']) or {}).items() if q and k!='?'}; cs_={k for k,q in (bsup.get(c_) or {}).items() if q and k!='?'}
            if ts_ and cs_ and ts_&cs_: notes=notes+[f"same supplier as this material ({', '.join(sorted(ts_&cs_))})"]
            cands.append((g,p_,notes))
        cands.sort(key=lambda z:(z[0],not any(n_.startswith('YOUR SAP BOM') for n_ in z[2]),not any(n_.startswith('MIS ') for n_ in z[2]),-(z[1]['free'])))
        nmat+=1; withstock=[z for z in cands if z[1]['stock']+z[1]['it']>0]
        if withstock: nwith+=1
        head=[b['common'],b['mat'],x['status'],(round(x['tgt_days'],1) if x.get('tgt_days') is not None else None),(round(need_,1) if need_ else None),b.get('desc') or desc.get(b['mat'])]
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
            add('Stock','Same-name stock NOT counted (alternate code)','High' if ap and tot>=ap*3 else 'Medium',b['common'],m,'Stock counted',b['stock'],'Not counted',tot,
                ', '.join(f'{c} {q:,.1f}' for c,q in unc.items())+f' has the same MIS Common Name (alternate code, other unit) - not counted (MIS Adhoc is not used); if usable for {m}, that is'+(f' {tot/ap:.1f} more days' if ap else ' more stock'),
                'Confirm with production: if usable, add it on Settings > Substitutes (MRP material, substitute, factor) so the MRP counts it')
        for c_,nm_,r_,q_ in (b.get('conv_unlinked') or []):
            add('Master data','Code not linked on the MIS Common Name sheet','High',b['common'],m,'',None,'MIS Adhoc today',q_,
                f"MIS converts {c_} ('{nm_}') into {m} on its Adhoc sheet (x{r_:g}), but the Common Name sheet calls it '{nm_}', not '{b['common']}'. "
                f"Its stock shows on the Adhoc row (day 1), not in Stock, and its in-transit is not counted",
                f"Add / correct the row on the MIS Common Name sheet: {c_} -> Common Name '{b['common']}'")
        ao=(b.get('adhoc_other') or 0)-sum(q_ for _,_,_,q_ in (b.get('conv_unlinked') or []))
        if ao>0.5:
            add('Stock','MIS Adhoc quantity counted today (not a code conversion)','High' if ao>b['stock'] else 'Medium',b['common'],m,'MB52 + JW stock',b['stock'],'Adhoc today',ao,
                f'MIS Adhoc adds {ao:,.0f} on day 1 ({ao/max(b["stock"],1):.1f}x the stock) - counted as in MIS; the WSA sheet balance may not include it',
                'Confirm what this Adhoc quantity is (stock elsewhere? receipt?) - if not real, ask MIS to remove it')

    # 1 MIS stock vs WSA sheet balance on the sheet stock date (pooled sheets compared as a whole)
    for lab,sh in BW.get('sheets',{}).items():
        mems=[m for m in sh['members'] if m in n2c]
        if not mems or sh['bal'] is None: continue
        mis_st=sum(blocks[n2c[m]]['stock'] for m in mems)  # pooled sheet (e.g. glass front + back): WSA balance = total of the materials
        wsa=float(sh['bal']); d=mis_st-wsa; ap=sum(avgplan[n2c[m]] for m in mems)
        days=abs(d)/ap if ap else None
        if abs(d)>max(1,0.02*max(abs(wsa),abs(mis_st))):
            sev='High' if (days is not None and days>=2) else 'Medium' if (days is None or days>=0.5) else 'Low'
            add('Stock','MIS stock vs WSA sheet balance',sev,' / '.join(mems),', '.join(n2c[m] for m in mems),
                'WSA balance',wsa,'MIS stock',mis_st,
                f"WSA sheet '{sh['sheet']}' balance on {sh['asof']:%d-%b} vs MIS stock today ({DS['src']})"+(f"; gap = {days:.1f} days of plan" if days is not None else '')
                +("; pooled sheet, MIS = total of its materials" if len(mems)>1 else ''),
                'Check which count is current (receipts or issues booked on one side only)')
    # 3 Supplier split gap (RM Status by Make vs MB52) for glass
    for m,b in blocks.items():
        for sp in b.get('sup') or []:
            if sp['name']=='Unassigned':
                add('Stock','Glass stock without supplier (RM Status vs MB52)','Low',b['common'],m,'MB52 total',b['stock'],'Sum by Make (RM Status)',b['stock']-sp['stock'],
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
    for lab,sh in BW.get('sheets',{}).items():
        mems=[m for m in sh['members'] if m in n2c]
        if not mems or not sh['dem']: continue
        wsa=sum(v for d,v in sh['dem'].items() if today<=pd.Timestamp(d)<today+pd.Timedelta(days=H))
        mis_p=sum(sum(blocks[n2c[m]]['plan'][:H]) for m in mems)
        if abs(mis_p-wsa)>max(1,0.10*max(mis_p,wsa)):
            add('Plan',f'MIS plan vs WSA sheet demand (next {H} days)','High' if abs(mis_p-wsa)>0.25*max(mis_p,wsa,1) else 'Medium',' / '.join(mems),', '.join(n2c[m] for m in mems),
                'WSA demand',wsa,'MIS plan',mis_p,f"WSA sheet '{sh['sheet']}' Demand column vs MIS Total Plan Req"+(' (pooled sheet, MIS = total of its materials)' if len(mems)>1 else '')
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
                    f"BL {bl} ({ves}): WSA {pd.Timestamp(wd):%d-%b} vs MIS {md:%d-%b} ({dd_:+d} days)",'Agree the arrival date (MIS = ETA port + 5 days)')
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
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws['A2'].alignment=Alignment(wrap_text=True,vertical='top'); ws.merge_cells('A2:L2'); ws.row_dimensions[2].height=30
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
        for j,t in enumerate(['Severity','Area','Check','Material','Code','Action']): c=ws.cell(top,1+j,t); c.font=hf; c.fill=hfill; c.alignment=ctr; c.border=bd
        for i,x in enumerate(tops):
            r=top+1+i
            for j,v in enumerate([x['sev'],x['area'],x['check'],x['material'],x['code'],x['action']]):
                c=ws.cell(r,1+j,v); c.font=bf; c.border=bd; c.alignment=Alignment(vertical='top',wrap_text=j in (2,3,5))
            ws.cell(r,1).fill=PatternFill('solid',fgColor=sevfill[x['sev']])
        top=top+len(tops)+4
        ws.cell(top-1,1,'ALL ISSUES').font=Font(name=AR,size=12,bold=True,color=NAVY)
    hdr=['Severity','Area','Check','Material','Code','Source A','Value A','Source B','Value B','Difference (B - A)','Detail','Action']
    for j,t in enumerate(hdr): c=ws.cell(top,1+j,t); c.font=hf; c.fill=hfill; c.alignment=ctr; c.border=bd
    order={'High':0,'Medium':1,'Low':2}
    issues=sorted(issues,key=lambda x:(order[x['sev']],rrank(str(x.get('code') or '').strip()),areas.index(x['area']) if x['area'] in areas else 9,-(abs(x['diff']) if isinstance(x['diff'],(int,float)) else 0)))   # severity, then highest shortage risk
    for i,x in enumerate(issues):
        r=top+1+i
        vals=[x['sev'],x['area'],x['check'],x['material'],x['code'],x['a_lab'] or None,x['a'],x['b_lab'] or None,x['b'],x['diff'],x['detail'],x['action']]
        for j,v in enumerate(vals,1):
            c=ws.cell(r,j,round(v,3) if isinstance(v,float) else v); c.font=bf; c.border=bd; c.alignment=Alignment(vertical='top',wrap_text=j in (3,4,11,12))
            if j in (7,9,10): c.number_format='#,##0;-#,##0;0'
        ws.cell(r,1).fill=PatternFill('solid',fgColor=sevfill[x['sev']]); ws.cell(r,1).font=Font(name=AR,size=9,bold=True)
    last=top+max(1,len(issues))
    if not issues: ws.cell(top+1,1,'No discrepancies found')
    ws.add_table(Table(displayName='DataCheck',ref=f'A{top}:L{last}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=True)))
    for col,w in zip('ABCDEFGHIJKL',[9,11,30,30,16,16,12,16,12,13,60,42]): ws.column_dimensions[col].width=w
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
        except Exception: return None
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
        x=per[c]; plan=[0.0]*nd; cust={}; cl={}
        for bn,mq in x['q'].items():
            q=sum(mq.values())/len(mq)
            for cu,v in (modplan.get(bn) or {}).items():
                arr=cust.setdefault(cu,[0.0]*nd)
                for k in range(nd): plan[k]+=q*v[k]; arr[k]+=q*v[k]
                for ln,lv in (((lineplan or {}).get(bn) or {}).get(cu) or {None:v}).items():
                    a_=cl.setdefault((cu,ln),[0.0]*nd)
                    for k in range(nd): a_[k]+=q*lv[k]
        sr+=1; desc=' '.join(str(x['desc'] or '').split())
        out.append(dict(key=f'{sr}CS/PM',sr=sr,short='CS - Consumable' if c.upper().startswith('CS') else 'PM - Packaging',common=desc,mat=c,desc=desc,mtype='CS/PM',
                        stock=stock.get(c,0.0),uom=x['uom'] or '',plan=plan,cust_plan=dict(sorted({k:v for k,v in cust.items() if any(v)}.items(),key=lambda kv:-sum(kv[1]))),
                        cust_line=dict(sorted({k:v for k,v in cl.items() if any(v)}.items(),key=lambda kv:(-sum(cust.get(kv[0][0],[0])),kv[0][0],str(kv[0][1])))),
                        transit=[0.0]*nd,adhoc=[0.0]*nd,openpo=[0.0]*nd,count_stock=True,raw=raw.get(c,'not on Stock Count'),
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
    except Exception: pass
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
    except Exception: return META
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
    except Exception: return I
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
    except Exception: return
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
    except Exception: pass
    try:
        for d in bom_rows_std([list(r) for r in wb_ro(mrp_path)['BOM'].iter_rows(values_only=True)]) or []:
            c=d['Component']
            if c[:2].upper() not in ('CS','PM') or not isinstance(d.get('Quantity'),(int,float)): continue
            k=(_norm(d.get('Customer')),c); CS_RATE[k]=CS_RATE.get(k,0.0)+float(d['Quantity'])
            mdl=str(d.get('Module Model') or '').strip()
            if mdl: CS_RATE.setdefault(('model',mdl,c),0.0); CS_RATE[('model',mdl,c)]=max(CS_RATE[('model',mdl,c)],CS_RATE[k])
    except Exception: pass

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
        plan=[0.0]*nd; cust={}; cl={}; used_rate=[]; used_mis=[]
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
                        for ln,lv in lines_.items(): addl(cu,ln,lv,q_)
                    else: addl(cu,None,v,q_)
            # MIS customers of this BOM Name that your Production Plan does not have: MIS requirement as it is
            for cu,v in ((req.get(c) or {}).get(bn) or {}).items():
                if _norm(cu) in covered or not any(v): continue
                if bn not in used_mis: used_mis.append(bn)
                arr=cust.setdefault(cu,[0.0]*nd)
                for k in range(nd): plan[k]+=v[k]; arr[k]+=v[k]
                for ln,lv in (reql.get((c,bn,cu)) or {None:v}).items(): addl(cu,ln,lv)
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
    grp={}
    for d in bom:
        c=str(d['Component'] or '').strip(); bn=str(d['MIS BOM Name'] or '').strip(); f_=1.0
        if c in (DS_rm.get('alias') or {}): f_=(DS_rm.get('alias_factor') or {}).get(c,1.0); c=DS_rm['alias'][c]
        if isinstance(d['Quantity'],(int,float)):
            k=(c,bn,str(d['Module Model'] or '').strip(),str(d['UoM'] or '').strip().upper()); grp[k]=grp.get(k,0.0)+float(d['Quantity'])*f_
    # MIS rate per module for the same material: sum of its standard BOM lines = requirement / modules
    for (c,bn,mdl,u),q in grp.items():
        if c not in blocks or (c,bn) not in mis_rate: continue
        b=blocks[c]; mr,mu=mis_rate[(c,bn)]
        if u==mu.upper(): conv=q
        elif u=='KG' and b['common'] in pack: conv=q/pack[b['common']]
        else: continue
        if mr and abs(conv/mr-1)>0.05:
            ratio=conv/mr; P=[x*ratio for x in b['plan']]; T=[a+x for a,x in zip(b['transit'],b['adhoc'])]
            cv0=cover_days(b['stock'],b['plan'],T,b.get('openpo') or [0.0]*len(P))[1]; cv1=cover_days(b['stock'],P,T,b.get('openpo') or [0.0]*len(P))[1]
            I.append(dict(area='Plan',check='Usage per module: MIS BOM vs your BOM',sev='High' if (cv1<TD<=cv0 or (cv1<TD and cv1<cv0)) else 'Medium',
                          material=b['common'],code=c,a_lab=f'MIS per module ({mu})',a=round(mr,6),b_lab=f'Your BOM per module ({mu})',b=round(conv,6),diff=round(conv-mr,6),
                          detail=(f"{bn}: your BOM {q:.4g} {u}"+(f" = {conv:.5f} {mu} at {pack.get(b['common']):.0f} KG/{mu}" if u!=mu.upper() else '')+
                                  f" vs MIS {mr:.5g} {mu} ({ratio-1:+.0%}). The MRP uses YOUR BOM for this customer (MIS only when your BOM does not list it)."),
                          action='Information: check which rate is right - if your BOM is right, ask MIS to correct its Standard BOM; if MIS is right, correct your BOM'))
    seen=set(); out=[]
    for x in I:
        k=(x['code'],x['detail'].split(':')[0])
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
              'Stock = MIS MB52 plant + warehouse; JW = RM Status for OEM JW; CS / PM stock = your Stock Count. Use it to stop new orders, re-use or return slow stock.')
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
    except Exception: return out
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
        except Exception: continue
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

def make_movement_xlsx(log, path):
    """Stock Movement: for each material, run to run: opening, expected out (plan) and in (MIS in transit / adhoc due),
       expected closing vs actual closing, and the implied receipts - to verify stock in and out. Plus the hidden Movement Log."""
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    wb=Workbook(); ws=wb.active; ws.title='Stock Movement'
    ws['A1']='Daily stock movement - expected vs actual, run to run'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=('Each run saves stock and the expected movement for the next days. Next run: Expected closing = opening - plan used + in transit / adhoc due (MIS); '
              'Actual closing = today\'s stock = Stock + to OEM Stock (like the Action List; today\'s row adds up the lines on the Stock Data sheet). Implied receipts = actual change + plan used. '
              'Check = difference above 2% of opening and above half a day of plan: a receipt, issue or count not as expected. Open PO (WSA) is shown separately, not in the expected closing.')
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:T2'); ws['A2'].alignment=Alignment(wrap_text=True); ws.row_dimensions[2].height=40
    H=['Date','Material','Code','UOM','Days since previous','Opening (previous run)','Plan used (expected out)','In transit / adhoc due (expected in)','Open PO due (WSA, info)',
       'Expected closing','Actual closing','Actual closing - Stock','Actual closing - to OEM Stock','Raw stock (own unit, as counted / in MB52)','Difference (actual - expected)','Actual change',
       'Implied receipts (change + plan)','Implied - expected in','Result','Supplier (in transit due in the period / next arrivals)']
    for j,h in enumerate(H,1):
        c=ws.cell(4,j,h); c.font=Font(name=AR,size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=NAVY); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=bd
    ws.row_dimensions[4].height=42
    by={}
    for e in log: by.setdefault(e['code'],{})[e['date']]=e
    rows=[]
    for code,dd in by.items():
        ds=sorted(dd)
        for i,d in enumerate(ds):
            e=dd[d]
            if i==0:
                nxt=sorted(e.get('sup') or [])[:3]
                oe_=e.get('oem'); rows.append([d.to_pydatetime(),e['material'],code,e['uom'],None,None,None,None,None,None,e['stock'],(e['stock']-oe_) if oe_ is not None else None,oe_,e.get('raw') or '',None,None,None,None,'First snapshot',
                             ('next: '+'; '.join(f"{(d+pd.Timedelta(days=k)):%d-%b} {q:,.0f} {sn[:28]}" for k,q,sn in nxt)) if nxt else '']); continue
            p=dd[ds[i-1]]; n=(d-ds[i-1]).days
            out=sum(p['plan'][:n]); inn=sum(p['inb'][:n]); opo=sum(p['opo'][:n])
            exp=p['stock']-out+inn; diff=e['stock']-exp; chg=e['stock']-p['stock']; imp=chg+out
            avg=(sum(p['plan'][:7])/7) if p['plan'] else 0
            ok=abs(diff)<=max(0.02*abs(p['stock']),0.5*avg,1)
            res='OK' if ok else ('Check: more stock than expected' if diff>0 else 'Check: less stock than expected')
            if n>len(p['plan']): res+=f' (gap of {n} days)'
            due={}
            for k,q,sn in (p.get('sup') or []):
                if k<n: due[sn]=due.get(sn,0.0)+q
            oe_=e.get('oem'); rows.append([d.to_pydatetime(),e['material'],code,e['uom'],n,p['stock'],out,inn,opo,exp,e['stock'],(e['stock']-oe_) if oe_ is not None else None,oe_,e.get('raw') or '',diff,chg,imp,imp-inn,res,
                         '; '.join(f"{sn[:28]} {q:,.0f}" for sn,q in sorted(due.items(),key=lambda kv:-kv[1]))])
    rows.sort(key=lambda r_:(-r_[0].toordinal(),str(r_[1])))
    last_=max((r_[0] for r_ in rows),default=None)
    for i,r_ in enumerate(rows):
        rr=5+i
        for j,v in enumerate(r_,1):
            c=ws.cell(rr,j,round(v,3) if isinstance(v,float) else v); c.font=Font(name=AR,size=9,italic=(j==14)); c.border=bd; c.alignment=Alignment(vertical='top',wrap_text=j in (2,14,19,20))
            if j==1: c.number_format='dd-mmm-yy'
            if 6<=j<=13 or 15<=j<=18: c.number_format='#,##0;-#,##0;"-"'
        if r_[0]==last_ and not VALUES_ONLY:   # today: actual closing = the lines on Stock Data (double-check there)
            ws.cell(rr,12,f"=SUMIFS('Stock Data'!$L$5:$L$3000,'Stock Data'!$A$5:$A$3000,$C{rr},'Stock Data'!$C$5:$C$3000,\"Stock\")")
            ws.cell(rr,13,f"=SUMIFS('Stock Data'!$L$5:$L$3000,'Stock Data'!$A$5:$A$3000,$C{rr},'Stock Data'!$C$5:$C$3000,\"to OEM\")")
            ws.cell(rr,11,f'=L{rr}+M{rr}')
            if r_[9] is not None: ws.cell(rr,15,f'=K{rr}-J{rr}')
            if r_[5] is not None: ws.cell(rr,16,f'=K{rr}-F{rr}'); ws.cell(rr,17,f'=P{rr}+G{rr}'); ws.cell(rr,18,f'=Q{rr}-H{rr}')
        res=str(r_[18]); f_={'OK':'C6EFCE','First snapshot':'DDEBF7'}.get(res,'FFEB9C' if res.startswith('Check') else None)
        if f_: ws.cell(rr,19).fill=PatternFill('solid',fgColor=f_)
        if res.startswith('Check'): ws.cell(rr,15).fill=PatternFill('solid',fgColor='FFEB9C')
    if rows: ws.add_table(Table(displayName='StockMovement',ref=f'A4:T{4+len(rows)}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=True)))
    nchk=sum(1 for r_ in rows if str(r_[18]).startswith('Check')); ndays=len({r_[0] for r_ in rows})
    ws.cell(3,1,(f'{ndays} run date(s) logged; {nchk} movements to check.' if ndays>1 else
                 'First run: today\'s stock is saved - from the next run the sheet shows the movement since the previous run.')).font=Font(name=AR,size=10,bold=True,color='9C0006')
    for j,w in enumerate([10,34,18,6,8,12,12,13,12,12,12,12,12,40,13,12,13,12,28,44],1): ws.column_dimensions[L(j)].width=w
    ws.freeze_panes='D5'; ws.sheet_view.zoomScale=90
    set_print(wb.active,'Stock Movement'); wb.save(path); return nchk



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

def hide_all_but(pkg_path, keep='Action List', also=()):
    """Only one sheet visible (the others hidden, still there for the formulas); it opens on that sheet."""
    import zipfile, shutil
    z=zipfile.ZipFile(pkg_path); wbx=z.read('xl/workbook.xml').decode('utf-8')
    names=re.findall(r'<sheet\b[^>]*name="([^"]+)"',wbx); idx=names.index(keep) if keep in names else 0
    def fix(m_):
        t=m_.group(0); t=re.sub(r'\s+state="[^"]*"','',t)
        return t if any(re.search(r'name="'+re.escape(n_)+'"',t) for n_ in (keep,)+tuple(also)) else t.replace('<sheet ','<sheet state="hidden" ',1)
    wbx=re.sub(r'<sheet\b[^>]*/>',fix,wbx)
    wbx=re.sub(r'(<workbookView\b[^>]*?)\s+activeTab="\d+"',r'\1',wbx); wbx=re.sub(r'(<workbookView\b[^>]*?)\s+firstSheet="\d+"',r'\1',wbx)
    wbx=re.sub(r'<workbookView\b',f'<workbookView activeTab="{idx}" firstSheet="{idx}"',wbx,count=1)
    rels=z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    rid={n_:re.search(r'name="'+re.escape(n_)+r'"[^>]*r:id="([^"]+)"',wbx).group(1) for n_ in names}
    tgt={}
    for m_ in re.finditer(r'<Relationship\b[^>]*>',rels):
        i_=re.search(r'Id="([^"]+)"',m_.group(0)); t_=re.search(r'Target="([^"]+)"',m_.group(0))
        if i_ and t_: tgt[i_.group(1)]='xl/'+t_.group(1).lstrip('/').replace('xl/','',1)
    keep_part=tgt.get(rid.get(keep))
    tmp_=pkg_path+'.hide'
    with zipfile.ZipFile(tmp_,'w',zipfile.ZIP_DEFLATED) as o:
        for it in z.infolist():
            d=z.read(it.filename)
            if it.filename=='xl/workbook.xml': d=wbx.encode('utf-8')
            elif it.filename.startswith('xl/worksheets/sheet') and it.filename.endswith('.xml'):
                t=d.decode('utf-8'); t=re.sub(r'\s+tabSelected="[^"]*"','',t)
                if it.filename==keep_part: t=re.sub(r'<sheetView\b','<sheetView tabSelected="1"',t,count=1)
                d=t.encode('utf-8')
            o.writestr(it,d)
    z.close(); shutil.move(tmp_,pkg_path)


def al_group(h):
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
    return {'wsa':'833C0B','work':'595959'}.get(g,'1F3864')
AL_BANNERS={'base':('MATERIAL / STOCK / IN TRANSIT','1F3864'),
            'tg_wo':('DAILY TARGET - WITHOUT OEM (Lines 1-4, Stock + in transit)','548235'),'tg_w':('DAILY TARGET - WITH OEM (all lines, + to OEM Stock)','2E75B6'),
            'pl_wo':('DAILY PLAN - WITHOUT OEM (Lines 1-4, Stock + in transit)','548235'),'pl_w':('DAILY PLAN - WITH OEM (all lines, + to OEM Stock)','2E75B6'),
            'wsa':('WSA RM PLANNING','833C0B'),'work':('STATUS / ACTION (script)','595959')}



def sup_short(n):
    """Short supplier name: no country in brackets, no Co./Ltd./Glass/Technology..., first 2 words (FLAT (VIETNAM) CO., LTD -> Flat)."""
    t=re.sub(r'\([^)]*\)',' ',str(n or '')).replace(',',' ').replace('.',' ')
    w=[x for x in t.split() if x.upper() not in ('CO','LTD','LIMITED','COMPANY','CORP','CORPORATION','INC','GLASS','TECHNOLOGY','TECH','INDUSTRIAL','INDUSTRY','NEW','MATERIAL','MATERIALS','SOLAR','ENERGY','JSC','PTE','LLC','GROUP')]
    return ' '.join(w[:2]).title() or str(n or '')




def audit_workbook(path):
    """Self-check after every run (your rules): (1) a material short in the plan is never labelled safe on a visible sheet;
       (2) every visible sheet lists materials in the same shortage-risk order as the Action List. Prints what breaks the rules."""
    try:
        wb=load_workbook(path,read_only=True)
        al=list(wb['Action List'].iter_rows(min_row=4,values_only=True)); h=list(al[0]); S='Short from → to (with in transit)'
        order=[]; short=set()
        for r in al[1:]:
            d=dict(zip(h,r))
            if not d.get('Component') or ' - ' in str(d.get('Component description'))[:40]: continue
            if d['Component'] not in order: order.append(d['Component'])
            if d.get(S) and 'no stock-out' not in str(d[S]) and str(d[S])!='No demand': short.add(d['Component'])
        rank={c:i for i,c in enumerate(order)}; safe={'In stock','In Stock','Excess','Above reorder','OK','no shortage','No action','covered'}
        issues=[]
        SKIP={'Settings','Stock Count','BOM','Module Plan','MIS Layout','How it works','Definitions','Movement Log','In Transit Status','RM Issues'}
        for ws in wb.worksheets:
            if ws.sheet_state!='visible': continue
            seq=[]; bad=0; SEVL=[None]
            if ws.title in SKIP:                                         # your input copies / logs kept in source or date order
                continue
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
                if not seq or seq[-1]!=codes[0]: seq.append(codes[0])
            bad+=sum(1 for a,b in zip(seq,seq[1:]) if rank[b]<rank[a])
            if bad: issues.append(f'{ws.title}: {bad} places out of shortage-risk order')
        print('Self-check (shortage-risk rules): OK' if not issues else 'Self-check WARNING:\n  '+'\n  '.join(issues[:10]))
    except Exception as ex_: print(f'NOTE: self-check skipped ({type(ex_).__name__})')


POST={'path':None,'wb':None}
def post_wb(path):
    """The finished workbook is opened ONCE for all finishing steps (dashboards, sheets, style) and saved once (post_flush)."""
    if POST['path']!=path or POST['wb'] is None: POST['wb']=load_workbook(path); POST['path']=path
    return POST['wb']
def post_flush(path):
    if POST['wb'] is not None and POST['path']==path: _safe_save(POST['wb'],path)
    POST['wb']=None; POST['path']=None

def _safe_save(wb, path):
    """Save to a temporary file first and replace the real file only when the save worked - a failing step never damages MRP_Daily."""
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
def rm_lines(LR, nd):
    """Line can run / limiting material counting RAW MATERIALS only (stop day per material and line from the allocation)."""
    out=[]
    for l in (LR or []):
        ks=[(k,m) for (m,lk),k in ALLOC_STOP_.items() if lk==l['key'] and is_rm(m)]
        kt=min([k for k,_ in ks],default=nd)
        lim=[]
        if kt<nd:
            for k,m in ks:
                if k==kt:
                    nm=next((x['b'].get('common') for x in (A_ALL or []) if x['b']['mat']==m),m)
                    if nm not in lim: lim.append(nm)
        out.append(dict(l,t=(kt,lim)+tuple(l['t'][2:])))
    return out
A_ALL=[]


def risk_info(x, d0):
    """Shortage risk of one material: first short day (projected stock incl. in transit < 0), lead time, class, short date."""
    w_=x.get('wins') or []; k=w_[0][0] if w_ else None; lt=int(x['b'].get('lt_used') or 0); td=x.get('tgt_days')
    if x.get('no_demand'): return dict(k=None,lt=lt,cls='No demand',date=None)      # no demand = never Critical / At risk / Watch / Short
    cls=('ALERT' if (k is not None and k<lt) else ('Short later' if k is not None else ('Under 21 d' if (td is not None and td<REPORT_DAYS) else ('OK' if x.get('tgt_need') else 'No need'))))
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
    G={'ALERT':0,'Short later':1,'Under 21 d':2,'OK':3,'No need':4,'No demand':5}
    def key(x):
        ri=risk_info(x,d0); td=x.get('tgt_days')
        return (G[ri['cls']],ri['k'] if ri['k'] is not None else 0,td if td is not None else 1e9,str(x['b'].get('common') or ''))
    RISK_RANK.clear(); RISK_RANK.update({x['b']['mat']:i for i,x in enumerate(sorted(A,key=key))})
def rrank(code): return RISK_RANK.get(code,10**6)
def risk_priority(x, d0, below_rop=False):
    """Order priority from the SHORTAGE risk (day-by-day plan), not from reorder-point totals:
       EXPEDITE = short inside the lead time · ORDER NOW = short later, last order date within the release window ·
       ORDER BY dd-mmm = short later, latest order date · Planned = not short, only below the reorder point."""
    ri=risk_info(x,d0)
    if ri['cls']=='No demand': return None                      # no demand: no Expedite / Order now / Release
    if ri['k'] is not None and ri['k']<ri['lt']: return 'EXPEDITE'
    if ri['k'] is not None:
        last=ri['k']-ri['lt']
        return 'ORDER NOW' if last<=RELEASE_WINDOW else f"ORDER BY {(d0+pd.Timedelta(days=last)):%d-%b}"
    return 'Planned' if below_rop else None
def at_risk(x, d0):
    """Dashboards show a material only when it is at risk: short in the plan (ALERT / short later) or under 21 days at the daily target."""
    return risk_info(x,d0)['cls'] in ('ALERT','Short later','Under 21 d')
def short_lt_text(x, d0):
    """All periods with stock below zero that START inside the material lead time ('05-Nov → 07-Nov (3 d); 20-Nov → beyond 21-Mar');
       else the first stock-zero day after the lead time; else 'not short'. Returns (text, number of periods inside the lead time)."""
    if x.get('no_demand'): return 'No demand',0
    w_=x.get('wins') or []; lt_=int(x['b'].get('lt_used') or 0)
    if not w_: return 'not short',0
    D=lambda k: d0+pd.Timedelta(days=k)
    ins=[w for w in w_ if w[0]<lt_]
    if ins:
        return ';\n'.join(f"{D(a):%d-%b} → "+(f"beyond {D(z):%d-%b}" if op else f"{D(z):%d-%b} ({z-a+1} d)") for a,z,op in ins),len(ins)
    return f"{D(w_[0][0]):%d-%b} (after lead time {lt_} d)",0

def cover_watch(x, d0, ships, nd):
    """ONE calculation for the Risk Charts list AND the Action List columns 'Below 21 Days From' / 'Short From (Stock = 0)' / 'Days Late'.
       Below 21 Days From = first day projected stock (rolled forward + MIS in transit by arrival date - plan) covers fewer than REPORT_DAYS days of need.
       Short From = first day projected stock is below zero. Days Late = arrival of the first shipment that comes AFTER the below-21-days day, minus that day.
       Returns None when the material has no demand (never shown as a shortage) or no need per day."""
    if x.get('no_demand'): return None
    TGT=REPORT_DAYS; HZ=min(nd,112); b=x['b']; m=b['mat']; S=x.get('S') or []
    P=b.get('plan') or []; nz=[v for v in P[:HZ] if v>0]
    need=(x.get('tgt_need') or (sum(nz)/len(nz) if nz else 0)) or None
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
def short_buckets(xs, d0):
    """How many materials start running short in: <= 7 d, 8-14, 15-30, 31-60, > 60 days, none."""
    b=[0,0,0,0,0,0]
    for x in xs:
        k=risk_info(x,d0)['k']
        b[5 if k is None else (0 if k<=7 else (1 if k<=14 else (2 if k<=30 else (3 if k<=60 else 4))))]+=1
    return list(zip(['≤ 7 DAYS','8 – 14 DAYS','15 – 30 DAYS','31 – 60 DAYS','> 60 DAYS','NO SHORTAGE'],b))





# ===================== RISK CHARTS: treemap (where the shortage is) + heatmap (when it runs short) - drawn with cells =====================
def _squarify(vals, x, y, w, h):
    """Squarified treemap: list of (x, y, w, h) for values (largest first) inside the rectangle."""
    out=[]; vals=[v for v in vals]
    def worst(row, side):
        s=sum(row); return max(max(side*side*r/(s*s), (s*s)/(side*side*r)) for r in row) if s>0 else 1e9
    rects=[]; items=list(vals)
    tot=sum(items) or 1.0; scale=w*h/tot; items=[v*scale for v in items]
    while items:
        side=min(w,h); row=[items.pop(0)]
        while items and worst(row+[items[0]],side)<=worst(row,side): row.append(items.pop(0))
        s=sum(row)
        if w>=h:                                   # lay the row vertically on the left
            cw=s/h if h else 0; cy=y
            for r in row: rh=r/cw if cw else 0; rects.append((x,cy,cw,rh)); cy+=rh
            x+=cw; w-=cw
        else:                                      # lay the row horizontally on top
            rh=s/w if w else 0; cx=x
            for r in row: cw=r/rh if rh else 0; rects.append((cx,y,cw,rh)); cx+=cw
            y+=rh; h-=rh
    return rects

SHIPS_={}
def add_risk_charts(path, A, AL_ORDER, d0, nd, ships=None):
    """Risk Charts: (1) Cover watch - raw materials that fall below 21 days in the next 16 weeks, when, and WHY
       (late arrival / late + not enough / no supply) with the shipment, days late and the action; (2) heatmap of weekly cover vs 21 days."""
    from openpyxl.styles import Border, Side
    wb=post_wb(path)
    if 'Risk Charts' in wb.sheetnames: del wb['Risk Charts']
    ws=wb.create_sheet('Risk Charts'); ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor='C0392B'
    F=lambda **k: Font(name='Calibri',**k)
    def fill(c,col): c.fill=PatternFill('solid',fgColor=col)
    thin=Side(style='thin',color='E3E8EE')
    base=[x for x in AL_ORDER if not x.get('sup_row') and is_rm(x['b']['mat']) and not x.get('no_demand') and not grp_follower(x['b']['mat'])]     # no demand = not on any risk list; a substitute group once
    HZ=min(nd,112); TGT=REPORT_DAYS
    def need_of(x):
        P=x['b'].get('plan') or []; nz=[v for v in P[:HZ] if v>0]
        return (x.get('tgt_need') or (sum(nz)/len(nz) if nz else 0)) or None
    # ---------------- 1) COVER WATCH ----------------
    rows=[]
    for x in base:
        cw=cover_watch(x,d0,ships,nd)                                   # same calculation as the Action List columns
        if not cw or cw['t'] is None: continue
        b=x['b']; m=b['mat']; nx=cw['nx']; dt_=cw['below']; gap_all=cw['gap_all']; cause=cw['cause']
        if cause=='LATE ARRIVAL': act=f"Expedite {nx.get('sup') or 'shipment'} to {dt_:%d-%b} (due {nx['raw']:%d-%b})"+(f" + order {gap_all:,.0f} for later weeks" if gap_all>0 else '')
        elif cause=='LATE + NOT ENOUGH': act=f"Expedite {nx.get('sup') or ''} + order {gap_all:,.0f} more"
        else: act=f"Order {gap_all:,.0f} now"+(" – inside lead time: alternative / expedite" if cw['t']<int(b.get('lt_used') or 0) else '')
        if cw['ov']: act+=f" · chase {len(cw['ov'])} overdue"
        rows.append(dict(name=grp_label(m,str(b.get('common') or m)),code=m,uom=b.get('uom'),cov0=cw['cov0'],t=cw['t'],ri=cw['ri'],cause=cause,nx=nx,late=cw['late'],act=act,ata=x.get('ata_txt'),slt=short_lt_text(x,d0)))
    rows.sort(key=lambda r_:rrank(r_['code']))
    ws.column_dimensions['A'].width=2
    W=[15,28,6,9,12,30,17,8,40,24,11,46]
    for j_,w_ in enumerate(W,2): ws.column_dimensions[L(j_)].width=w_
    ws.cell(1,2,'COVER WATCH · RAW MATERIALS · NEXT 16 WEEKS').font=F(size=11,bold=True,color='E67E22')
    ws.cell(2,2,f'MATERIALS THAT FALL BELOW {TGT} DAYS – AND WHY').font=F(size=24,bold=True,color='14213D')
    nl=sum(1 for r_ in rows if r_['cause']=='LATE ARRIVAL'); nn=sum(1 for r_ in rows if r_['cause']=='LATE + NOT ENOUGH'); ns=sum(1 for r_ in rows if r_['cause']=='NO SUPPLY')
    ws.cell(3,2,f"{len(rows)} raw materials fall below {TGT} days within 16 weeks:   {nl} because shipments arrive late   ·   {nn} late and not enough   ·   {ns} with nothing on the way").font=F(size=11,bold=True,color='C0392B')
    ws.cell(4,2,f"Cover = projected stock (rolled forward to today + MIS in transit by arrival date − plan) ÷ need per day.  LATE ARRIVAL = what is already on the way would keep {TGT}+ days, but it arrives after that date.  Materials with no demand are not listed.").font=F(size=9,color='5D6D7E')
    H=['Code','Material','UOM','Cover today (d)',f'Below {TGT} Days From','Short From (Stock = 0)','Cause','Days Late','Arrival at Plant (Final AtA)','Supplier','Qty','Action']
    for j_,h in enumerate(H,2):
        c=ws.cell(6,j_,h); c.font=F(size=10,bold=True,color='1F2D3D'); fill(c,'F4F6F8'); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        c.border=Border(bottom=Side(style='medium',color='C9D2DC'))
    ws.row_dimensions[6].height=32
    CC={'LATE ARRIVAL':('FEF3E2','B9770E'),'LATE + NOT ENOUGH':('FCE8E6','C0392B'),'NO SUPPLY':('FCE8E6','C0392B')}
    r=7
    for r_ in rows:
        vals=[r_['code'],r_['name'][:44],r_['uom'],round(r_['cov0'],1),(d0+pd.Timedelta(days=r_['t'])).to_pydatetime(),
              r_['slt'][0],r_['cause'],r_['late'],
              r_['ata'] or 'nothing on the way',(r_['nx'].get('sup') or '') if r_['nx'] else '-',(round(r_['nx']['qty']) if r_['nx'] else None),r_['act']]
        fm=[None,None,None,'0.0','dd-mmm-yy',None,None,'0',None,None,'#,##0',None]
        for j_,(v,f_) in enumerate(zip(vals,fm),2):
            c=ws.cell(r,j_,v); c.font=F(size=10,color='1F2D3D'); c.border=Border(bottom=thin)
            c.alignment=Alignment(vertical='top',horizontal=('left' if isinstance(v,str) else 'right'),wrap_text=(j_ in (7,10,11,13)))
            if f_: c.number_format=f_
        bg,fg=CC[r_['cause']]; fill(ws.cell(r,8),bg); ws.cell(r,8).font=F(size=10,bold=True,color=fg); ws.cell(r,8).alignment=Alignment(horizontal='center',vertical='center')
        if r_['cov0']<TGT: fill(ws.cell(r,5),'FCE8E6'); ws.cell(r,5).font=F(size=10,bold=True,color='C0392B')
        if r_['t']<=14: ws.cell(r,6).font=F(size=10,bold=True,color='C0392B')
        if r_['slt'][1]: ws.cell(r,7).font=F(size=10,bold=True,color='C0392B')                 # stock zero inside the lead time
        elif 'after lead time' in r_['slt'][0]: ws.cell(r,7).font=F(size=10,bold=True,color='B9770E')
        if isinstance(r_['late'],int) and r_['late']>0: ws.cell(r,9).font=F(size=10,bold=True,color='B9770E')
        nl_=max(str(r_['ata'] or '').count('\n')+1,-(-len(r_['act'])//50),-(-len(str((r_['nx'] or {}).get('sup') or ''))//24),r_['slt'][0].count('\n')+1,-(-len(r_['slt'][0])//32),1)
        ws.row_dimensions[r].height=max(19,13*nl_); r+=1
    if not rows: ws.cell(7,2,f'No raw material falls below {TGT} days in the next 16 weeks.').font=F(size=11,color='1E8449'); r=8
    # ---------------- 2) HEATMAP: weekly cover vs 21 days ----------------
    if 'Cover Heatmap' in wb.sheetnames: del wb['Cover Heatmap']
    ws=wb.create_sheet('Cover Heatmap'); ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor='12131C'
    r0=1; DARK='12131C'; TXT='E5E8EE'; MUT='8A949E'
    NW=16; cL=2; cS=4                                    # material name in B:C, weeks from D (one column per week)
    ws.column_dimensions['A'].width=2; ws.column_dimensions['B'].width=26; ws.column_dimensions['C'].width=8
    for w in range(NW): ws.column_dimensions[L(cS+w)].width=7
    rows_=[x for x in base if at_risk(x,d0)]
    height=4+len(rows_)+5
    for rr in range(r0,r0+height):
        for cc in range(1,cS+NW+1): fill(ws.cell(rr,cc),DARK)
    ws.cell(r0,cL,f'Weekly cover of at-risk raw materials vs {TGT} days').font=F(size=16,bold=True,color=TXT)
    ws.cell(r0+1,cL,f'Days of cover at the end of each week: red = short · orange / amber = below {TGT} days · blue = {TGT}+ days. Highest shortage risk on top.').font=F(size=9,color=MUT)
    for w in range(NW):
        c=ws.cell(r0+3,cS+w,f"{(d0+pd.Timedelta(days=7*w)):%d-%b}"); c.font=F(size=8,color=MUT); c.alignment=Alignment(horizontal='center')
    bins=[(-1e9,0,'D40000'),(0,7,'E8590C'),(7,14,'F08C00'),(14,TGT,'C9A227'),(TGT,35,'1A5A8E'),(35,60,'1E7BC4'),(60,1e9,'0EA5F5')]
    def colr(v):
        for a_,b_,c_ in bins:
            if a_<=v<b_: return c_
        return '1B2335'
    for i,x in enumerate(rows_):
        rr=r0+4+i; b=x['b']; S=x.get('S') or []; need=need_of(x)
        ws.merge_cells(start_row=rr,start_column=cL,end_row=rr,end_column=cS-1)
        c=ws.cell(rr,cL,str(b.get('common') or b['mat'])[:34]); c.font=F(size=9,color=TXT); c.alignment=Alignment(horizontal='right',indent=1)
        for w in range(NW):
            z=min(7*w+6,len(S)-1)
            if z<0 or not need: continue
            v=S[z]/need; cl=ws.cell(rr,cS+w,round(v)); fill(cl,colr(v)); cl.font=F(size=8,color='FFFFFF'); cl.alignment=Alignment(horizontal='center')
            cl.border=Border(left=Side(style='thin',color=DARK),right=Side(style='thin',color=DARK),top=Side(style='thin',color=DARK),bottom=Side(style='thin',color=DARK))
        ws.row_dimensions[rr].height=15
    rl=r0+4+len(rows_)+1
    ws.cell(rl,cS,'Cover at week end, days:').font=F(size=9,color=TXT)
    labs=['short','0–7','7–14',f'14–{TGT}',f'{TGT}–35','35–60','60+']
    for k,(a_,b_,c_) in enumerate(bins):
        c=ws.cell(rl+1,cS+2*k,labs[k]); fill(c,c_); c.font=F(size=8,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='center')
    ws.cell(rl+3,cL,'Source: MRP Daily projection (stock rolled forward to today + MIS in transit by arrival date − plan demand). Need per day = daily target need, else average planned demand.').font=F(size=8,color=MUT)
    for rr in range(r0,r0+height): ws.row_dimensions[rr].height=ws.row_dimensions[rr].height or 16
    pos=wb.sheetnames.index('Stock Management')+1 if 'Stock Management' in wb.sheetnames else 0
    wb.move_sheet('Risk Charts',offset=pos-wb.sheetnames.index('Risk Charts'))
    wb.move_sheet('Cover Heatmap',offset=pos+1-wb.sheetnames.index('Cover Heatmap'))


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
        except Exception: continue
RMI_COLS=['Issue ID','Date Raised','Project','Material','Supplier','PO Reference','Qty Required','Qty Received','Shortfall','Short from (stock = 0)','Required Date (at Factory)',
          'Revised ETA','Shipped qty','Pending qty','Delay (Days)','Production Impact (MW)','Priority','Owner','Status','Date Closed','Aging (Days)','Action / Mitigation']
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
    except Exception: pass
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
    except Exception: pass
    out=[]; seen=set()
    base=[x for x in AL_ORDER if not x.get('sup_row') and is_rm(x['b']['mat']) and not x.get('no_demand') and not grp_follower(x['b']['mat'])]
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
        auto_act=((f"Expedite {nx['sup']} PO {nx.get('po') or '-'} to {req_d:%d-%b} (due {nx['raw']:%d-%b})" if (nx and late and late>0) else
                   ((f"Order {gap:,.0f} {b.get('uom') or ''} now" if gap>=0.5 else f"Watch: under {TGT} days, no gap in the plan") if not nx else f"Follow up {nx['sup']} PO {nx.get('po') or '-'} due {nx['raw']:%d-%b}"))
                  +(" - inside lead time: use alternative" if ri['cls']=='ALERT' else ''))
        st=pv.get('Status') if str(pv.get('Status') or '').strip() and str(pv.get('Status')).strip().lower()!='closed' else 'Open'
        act=pv.get('Action / Mitigation') if (pv.get('Action / Mitigation') and pv.get('_auto') and pv.get('Action / Mitigation')!=pv.get('_auto')) else auto_act
        proj=str(x.get('proj') or '').replace('\n',', ') or ', '.join(sorted({str(cu) for (cu,ln) in (b.get('cust_line') or {})}))[:60]
        same=[s_ for s_ in sh if nx and s_['raw']==nx['raw'] and s_.get('sup')==nx.get('sup') and s_.get('po')==nx.get('po')]
        out.append([iid,raised.to_pydatetime(),proj,f"{grp_label(m,str(b.get('common') or ''))} ({' + '.join(mem)})",(nx['sup'] if nx else 'no supplier on the way'),
                    (str(nx.get('po')) if nx and nx.get('po') else None),round(gap+rec),round(rec),round(gap),slt_[0],req_d.to_pydatetime(),(nx['raw'].to_pydatetime() if nx else None),
                    (round(sum(s_.get('shq') or 0 for s_ in same)) if nx else None),(round(sum(s_['qty'] for s_ in same)) if nx else None),
                    late,(round(mw,2) if mw is not None else None),prio,pv.get('Owner'),st,None,(d0-raised).days,act,auto_act])
        seen.add(iid)
    for iid,pv in prev.items():                                # no longer at risk: closed automatically (kept for the record)
        if iid in seen: continue
        dr=pv.get('Date Raised'); dc=pv.get('Date Closed') if isinstance(pv.get('Date Closed'),(dt.datetime,dt.date)) else d0.to_pydatetime()
        out.append([iid,dr,pv.get('Project'),pv.get('Material'),pv.get('Supplier'),pv.get('PO Reference'),pv.get('Qty Required'),pv.get('Qty Received'),0,pv.get('Short from (stock = 0)'),
                    pv.get('Required Date (at Factory)'),pv.get('Revised ETA'),pv.get('Shipped qty'),pv.get('Pending qty'),pv.get('Delay (Days)'),pv.get('Production Impact (MW)'),pv.get('Priority'),pv.get('Owner'),
                    'Closed',dc,((pd.Timestamp(dc)-pd.Timestamp(dr)).days if isinstance(dr,(dt.datetime,dt.date)) else None),(f"Merged into RMI-{grp_main(iid[4:])} (substitute group)" if grp_follower(iid[4:]) else (pv.get('Action / Mitigation') or 'No longer at risk')),pv.get('_auto')])
    # open issues by Production Impact (MW), highest first; issues without MW after them; closed issues at the bottom
    out.sort(key=lambda r_:(r_[18]=='Closed',r_[15] is None,-(r_[15] if isinstance(r_[15],(int,float)) else 0)))
    # ---- sheet
    ws=wb.create_sheet('RM Issues'); ws.sheet_view.showGridLines=False
    F=lambda **k: Font(name='Calibri',**k)
    ws['A1']='RM Issues — raw materials at risk of shortage'; ws['A1'].font=F(size=20,bold=True,color='14213D')
    ws['A2']=(f"One issue per raw material that runs short or falls below {TGT} days in the next 16 weeks, sorted by Production Impact (MW) high to low (closed issues at the bottom), rebuilt every run. "
              "Required date = first day under 21 days of cover (or stock 0); Qty required = what is missing to keep 21 days at the worst point; Qty received = MIS PPC GRN since the issue was raised; "
              "Revised ETA / supplier / PO = next MIS shipment, Shipped qty / Pending qty = MIS Final Qty. / Final Pending Qty. of that shipment (same supplier, PO, ETA); Production impact = modules that cannot be built at the worst point x the Wp of the plan rows using it (capped at the MW planned for them in 16 weeks); a substitute group = ONE issue (group demand, stock and shipments). "
              "Yellow columns are yours (Owner, Status, Date Closed, Action) - kept from the previous MRP_Daily. Issues no longer at risk are closed automatically.")
    ws['A2'].font=F(size=9,color='8A949E'); ws.merge_cells('A2:V2'); ws['A2'].alignment=Alignment(wrap_text=True,vertical='top'); ws.row_dimensions[2].height=48
    H=RMI_COLS+['_auto']
    for j,h in enumerate(H,1):
        c=ws.cell(4,j,h); c.font=F(size=10,bold=True,color='1F2D3D'); c.fill=PatternFill('solid',fgColor='F4F6F8')
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=Border(bottom=Side(style='medium',color='C9D2DC'))
    ws.row_dimensions[4].height=32
    fm=[None,'dd-mmm-yy',None,None,None,None,'#,##0','#,##0','#,##0',None,'dd-mmm-yy','dd-mmm-yy','#,##0','#,##0','0','0.00',None,None,None,'dd-mmm-yy','0',None,None]
    PC={'High':('FCE8E6','C0392B'),'Medium':('FEF3E2','B9770E'),'Low':('FEF9E7','9A7D0A')}
    thin=Side(style='thin',color='E3E8EE')
    for i,row in enumerate(out,5):
        for j,v in enumerate(row,1):
            c=ws.cell(i,j,v); c.font=F(size=10,color='1F2D3D'); c.border=Border(bottom=thin)
            c.alignment=Alignment(vertical='center',wrap_text=(j in (3,10,22)),horizontal=('left' if isinstance(v,str) else 'right'))
            if fm[j-1]: c.number_format=fm[j-1]
            if j in (18,19,20,22): c.fill=PatternFill('solid',fgColor='FFFDE7')
        bg,fg=PC.get(row[16],('F2F4F6','5D6D7E')); ws.cell(i,17).fill=PatternFill('solid',fgColor=bg); ws.cell(i,17).font=F(size=10,bold=True,color=fg)
        if isinstance(row[14],(int,float)) and row[14]>0: ws.cell(i,15).font=F(size=10,bold=True,color='C0392B')
        if isinstance(row[9],str) and '→' in row[9]: ws.cell(i,10).font=F(size=10,bold=True,color='C0392B')
        if row[18]=='Closed':
            for j in range(1,23): ws.cell(i,j).font=F(size=10,color='A6ACAF')
        ws.row_dimensions[i].height=30
    for j,w_ in enumerate([22,11,26,30,24,13,11,11,11,30,12,11,11,11,8,11,9,14,10,11,8,46,4],1): ws.column_dimensions[L(j)].width=w_
    ws.column_dimensions['W'].hidden=True                       # auto action of this run (to keep your own Action text next time)
    if out: ws.add_table(Table(displayName='RMIssues',ref=f'A4:W{4+len(out)}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=False)))
    dv=DataValidation(type='list',formula1='"Open,In progress,Expedited,Waiting supplier,Closed"',allow_blank=True); ws.add_data_validation(dv)
    if out: dv.add(f'S5:S{4+len(out)}')
    ws.freeze_panes='E5'
    pos=(wb.sheetnames.index('Requirements')+1) if 'Requirements' in wb.sheetnames else 3
    wb.move_sheet('RM Issues',offset=pos-wb.sheetnames.index('RM Issues'))

def add_stock_check_sheet(path, AL_ORDER, d0):
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
              'usage per day = average Production Plan demand of the next 7 days · Used in MRP = the stock behind every shortage date in this workbook.')
    ws['A2'].font=F(size=10,color=MUTED); ws.merge_cells('A2:O2'); ws['A2'].alignment=Alignment(wrap_text=True,vertical='top'); ws.row_dimensions[2].height=46
    H=['Code','Material','UOM','SAP stock (MB52)','Physical count','Count date','Days since count','Usage / day','Usage since count','Received since count',
       'Estimated stock today','Used in MRP (own stock)','Estimate − SAP','Source','Shortage risk']
    for j,h in enumerate(H,1):
        c=ws.cell(4,j,h); c.font=F(size=10,bold=True,color=INK); c.fill=PatternFill('solid',fgColor='F4F6F8'); c.border=hb
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.row_dimensions[4].height=34
    base=[x for x in AL_ORDER if not x.get('sup_row')]
    r=5
    for x in base:
        b=x['b']; m=b['mat']; e=STOCK_EST.get(m) or {}; ri=risk_info(x,d0)
        own=float(b['stock'])-(0.0 if b.get('jw_out') else float(b.get('st_jw') or 0))   # JW already out of b['stock'] after the OEM step - do not subtract it twice
        est=e.get('est'); sapv=e.get('sap')
        if e.get('count') is None and not isinstance(sapv,(int,float)): sapv=own          # not counted: the MRP uses the SAP stock (incl. alternate codes)
        vals=[m,b.get('common') if not b.get('count_stock') else (b.get('desc') or b.get('common')),b.get('uom'),
              (round(sapv) if isinstance(sapv,(int,float)) else None),(round(e['count']) if e.get('count') is not None else 'no counting'),
              (pd.Timestamp(e['date']).to_pydatetime() if e.get('date') is not None else None),e.get('days'),
              (round(e['use_day'],1) if e.get('use_day') else None),(round(e['use']) if e.get('use') else None),(round(e['rec']) if e.get('rec') else None),
              (round(est) if est is not None else None),round(own),
              (round(est-sapv) if (est is not None and isinstance(sapv,(int,float))) else None),e.get('src') or 'SAP (MB52)',
              ri['cls']+(f" {ri['date']:%d-%b}" if ri['date'] is not None else '')]
        fm=[None,None,None,'#,##0','#,##0','dd-mmm-yy','0','#,##0.0','#,##0','#,##0','#,##0','#,##0','+#,##0;-#,##0;0',None,None]
        for j,(v,f_) in enumerate(zip(vals,fm),1):
            c=ws.cell(r,j,v); c.font=F(size=10,color=INK); c.border=rowb
            c.alignment=Alignment(vertical='center',horizontal=('left' if j in (2,14,15) or isinstance(v,str) else 'right'))
            if f_: c.number_format=f_
        if isinstance(e.get('days'),int) and e['days']>0:
            for j in (6,7): ws.cell(r,j).fill=PatternFill('solid',fgColor='FEF3E2'); ws.cell(r,j).font=F(size=10,bold=True,color='B9770E')
            ws.cell(r,11).font=F(size=10,bold=True,color=INK)
        cls=ri['cls']; bg,fg={'ALERT':('FCE8E6','C0392B'),'Short later':('FEF3E2','B9770E'),'Under 21 d':('FEF9E7','9A7D0A')}.get(cls,('E3F4EC','1E8449'))
        ws.cell(r,15).fill=PatternFill('solid',fgColor=bg); ws.cell(r,15).font=F(size=10,bold=True,color=fg)
        d_=ws.cell(r,13).value
        if isinstance(d_,(int,float)) and isinstance(sapv,(int,float)) and sapv and abs(d_)>0.1*abs(sapv): ws.cell(r,13).font=F(size=10,bold=True,color='C0392B')
        ws.row_dimensions[r].height=19; r+=1
    for j,w_ in enumerate([15,30,6,13,13,11,9,10,12,12,13,13,12,34,16],1): ws.column_dimensions[L(j)].width=w_
    if r>5: ws.add_table(Table(displayName='StockCheck',ref=f'A4:{L(len(H))}{r-1}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=False)))
    ws.freeze_panes='C5'
    pos=wb.sheetnames.index('Cover Heatmap')+1 if 'Cover Heatmap' in wb.sheetnames else 1
    wb.move_sheet('Stock Check',offset=pos-wb.sheetnames.index('Stock Check'))
    pass                                                           # saved once at the end (post_flush)






# ===================== ONE STYLE FOR THE WHOLE WORKBOOK + DEFINITIONS sheet =====================
DEFINITIONS=[
 ('All sheets','Need / day (target)','Daily target need','Σ customer daily target (Settings CUSTOMER DAILY TARGET, else line targets) × usage per module from your BOM (\'_Alt\' rows excluded; else MIS Standard BOM), in the MRP unit. Without OEM = customers with OEM = No; with OEM = all. Used for Inv. Days and the 21-day watch in the Action List.'),
 ('All sheets','Usage / day (plan)','Production Plan demand','Modules in the Production Plan on that day × usage per module, per customer / line, OEM customers left out. Drives the projection: stock, shortage dates, planned orders, days of cover, risk lists. Plan usage/day of a count = average of the next 7 days.'),
 ('MRP Daily / Weekly','Days of cover','How long the projected stock lasts','Projected stock at the end of the day (stock + in transit by arrival date − plan usage) ÷ average plan usage/day of the next N days, N = Settings "Days of cover: average plan usage over the next (days)" (21). Blank when no usage ahead.'),
 ('Action List','Inv. Days at line daily target','Cover at the daily target','(Stock + ALL in transit, whatever its arrival date) ÷ need/day (target). Without OEM: own stock; with OEM: + JW stock at the OEM. A target-based measure - it does not look at WHEN shipments arrive; the shortage dates do.'),
 ('All sheets','OEM','Left out of the calculations','Customers with OEM = Yes (CUSTOMER DAILY TARGET): their plan and the JW stock at the OEM are not part of any projection, shortage date, planned order or risk list. Only the Action List shows the with-OEM target and Inv. days.'),
 ('In Transit','Shipment code check','Which material a shipment is','Description, Supplier Name, RM Code, Final RM Code and the codes on the PO line are compared; the code whose description fits the shipment clearly best (sizes, thickness, busbar / ribbon width, GSM, front / back) is used; blank / unknown codes are matched among the materials the same supplier ships. Every change is listed on Check.'),
 ('Action List / MRP Daily','Supplier rows','Glass per supplier','Shown only when 2 or more suppliers have stock, in transit or allowed demand for the material - otherwise the row would repeat the material row.'),
 ('Action List','Daily target need (unit)','In the MRP unit of the material','Usage per module from your BOM converted to the MRP unit the same way as the plan (e.g. RTV KG per module ÷ 270 KG per drum), so need, stock and Inv. days are in the same unit'),
 ('Action List / Risk Charts / RM Issues','Short From (Stock = 0)','Every stock-zero period inside the lead time','All periods where projected stock (counts / MB52 rolled forward + MIS in transit by arrival date − plan) is below zero that START inside the material lead time, each with dates and length ("beyond" = still short at the end of the shortage horizon); none inside: first stock-zero day marked "after lead time"; none at all: not short. Any period inside the lead time = High risk (RM Issues priority High).'),
 ('Inputs','SAP stock (MB52)','Where stock comes from','Your MB52.xlsx (SAP export, same folder, newest MB52*.xlsx): Unrestricted + Quality Inspection = Total Stock; Plant name, Final Material Code, Dashboard Key, Company, Comman Name Cat.1 / Cat.2 are looked up from the MIS MB52 (else MIS Common Name). Stock date = save date of your MB52 file (older than today: rolled forward by usage)'),
 ('RM Issues','Issue log','One issue per raw material at risk','Material short or under 21 days in 16 weeks; Required date = first day under 21 d (or stock 0); Qty required = missing to keep 21 d at the worst point; Qty received = MIS PPC GRN since raised; Short from (stock = 0) = every period stock is below zero that starts inside the lead time (High priority); Revised ETA / Supplier / PO = next MIS shipment; Shipped qty = MIS Final Qty. and Pending qty = MIS Final Pending Qty. of that shipment (same supplier, PO and ETA); Delay = ETA − required date; Production impact = modules not buildable at the worst point × Production Plan Wp, capped at the MW planned on those lines / customers in 16 weeks; substitute group = one issue; Owner / Status / Date Closed / Action kept from the previous run; auto-closed when no longer at risk'),
 ('Inputs','BOM and Stock Count','Where they come from','From the PPC workbook (sheets BOM and Count, newest file with both, name with PPC preferred); MRP.xlsx only when no PPC workbook is in the folder'),

 ('Stock Management','Stock / In transit / Next arrival / Overdue in transit','Supply of each at-risk material','Stock = MB52 (or physical count) + JW · In transit = MIS pending qty · Next arrival = earliest MIS arrival date from today · Overdue = pending qty whose arrival date has passed'),
 ('Stock Management','Reorder level / Need per day','Inventory and production view','Reorder level = need/day × lead time + safety stock · Need/day = customer daily targets × BOM usage'),
 ('Stock Management','Status / On hold','Stock and quality status','Out of Stock (≤ 0) · Below reorder (below reorder level) · Above reorder · On hold = MB52 Blocked + Quality Inspection. Only at-risk materials are listed (ALERT / short later / under 21 days)'),
 ('Stock Management','Shortage risk','Shortage class of the at-risk material','ALERT = short inside the lead time · Short later · Under 21 d (no shortage in the plan). Materials with no demand are not listed'),
 ('Stock Check','SAP stock / Physical count / Count date','Both stock sources side by side','SAP = your MB52.xlsx Unrestricted + Quality Inspection (MIS MB52 only for plant name, final code, key, company, categories) · count = PPC Count sheet; count date = the date column that holds the count'),
 ('Risk Charts','Below 21 Days From / Short From (Stock = 0) / Days Late','Which raw materials drop below 21 days in the next 16 weeks, and why','Below 21 Days From = first day projected stock covers < 21 days of need · Short From = first day projected stock < 0 · Days Late = arrival of the first shipment after the below-21-days day minus that day (late = would have kept 21 days if on time) · Arrival at Plant (Final AtA) = same text as the Action List · Supplier / Qty = that next shipment · Action = Expedite / order more. Materials with no demand are not listed'),
 ('Action List','Below 21 Days From / Short From (Stock = 0) / Days Late','Same three columns as the Risk Charts','Same calculation (cover_watch) for every material, RM and PM; "No demand" for materials without demand'),
 ('Action List / all sheets','No demand','Material without demand = no shortage status, no action','Needs at least one of: production requirement (plan x BOM usage) in the next shortage-horizon days (Settings, 60), daily target need, average planned demand. Otherwise: not Critical / At risk / Watch / Short, no Expedite / Order now / Release order, not on Risk Charts or Stock Management'),
 ('Cover Heatmap','Heatmap cells','When it runs short','Projected stock at the end of each week ÷ need per day (daily target need, else average planned demand) = days of cover; red = short, orange / amber = below 21 days, blue = 21+ days; rows in shortage-risk order'),
 ('Stock Check / Stock Management','Data dates','How old the inputs are','MIS = save date of the MIS file, WSA = report date of the planning file, Counts = oldest–newest count date; red when older than 2 days. Old MIS: SAP stock of materials that are not counted is rolled forward by usage per day × days since the MIS date.'),
 ('Stock Check','Estimated stock today','Count rolled forward to today','Count − usage per day × days since count + received since count (MIS PPC GRN after the count date); usage per day = average Production Plan demand of the next 7 days. This is the stock every shortage date uses.'),
 ('Action List','Source / Source code / Qty / Factor','Where the stock comes from','Each stock line (MB52, physical count, JW at OEM); Qty (MRP unit) = Qty (source unit) × Factor (other roll length / pack)'),
 ('Action List','Descr. of Storage Loc.','Storage locations','MB52 storage-location description with quantity'),
 ('Action List','Lines (share of plan volume)','Which lines use the material','Share of the material\'s planned use over the plan horizon by production line (Production Plan × BOM usage)'),
 ('Action List','Project Name (share of plan)','Which projects use it','Same share by Project Name (Settings MIS LINK)'),
 ('Action List','Usage per module by customer','BOM usage behind the need','Usage per module per customer with its source: your BOM (MRP.xlsx) first, else MIS Standard BOM'),
 ('Action List','Module (MIS BOM, share of plan)','Module types using it','Short module name from the MIS BOM name with its share'),
 ('Action List','Short from → to (with in transit)','When stock runs out','Day-by-day projection: stock − plan + MIS in transit; ALERT when inside the lead time'),
 ('Action List','Material lead time (days)','Order-to-arrival time','WSA ETD → arrival history, else Settings default (RM 49/50, PM 14)'),
 ('Action List','Stock / to OEM Stock','Stock here / at the OEM','MB52 (or physical count) / MIS RM Status "JW WSA to OEM"'),
 ('Action List','In transit today (total qty, MIS)','Everything on the way','Σ MIS In Transit pending qty (all statuses)'),
 ('Action List','Arrival at plant (Final AtA)','Next arrivals','Passed-not-received total, next 4 arrival dates, rest as one total'),
 ('Action List','Production Daily Target_without / with OEM','Need per day at the targets','Σ customer daily target × usage per module; without OEM = customers with OEM = No (Settings CUSTOMER DAILY TARGET)'),
 ('Action List','Inv. Days … target','Days of cover at the targets','(Stock + in transit) ÷ daily target need; with OEM adds to-OEM stock; "stock only" excludes in transit; red under 21. Substitute group (Settings SUBSTITUTES, all lines): stock, in transit and need of the WHOLE group'),
 ('Action List','Substitute group','Material + its substitutes planned as one','Settings SUBSTITUTES (lines blank, valid today): daily target, Inv. days and the shortage date use the group stock + group in transit + group demand (each code x its factor); Stock / In transit columns show each code\'s own quantity. An alternative listed ALONE for a customer in your BOM carries that customer\'s demand.'),
 ('Action List','In transit / Open PO (WSA RM Planning)','WSA planning file view','Shipments at stage In transit / without vessel in the WSA RM planning file'),
 ('Action List','Alternative BOM (yours)','Approved alternatives','Codes in your BOM next to the MIS standard code of the same group, with MB52 stock and customers'),
 ('MRP by Line','Line can run (days) / Line stops on','How long the line runs','Earliest day a material it needs is short after sharing stock between lines (Line ranking, Vendor priority, Allowed vendors, Settings Substitutes) - and after using your approved ALTERNATIVE (your BOM) for that line\'s customers: free stock of the alternative = what its own plan does not need within its lead time'),
 ('MRP by Line','First needed / Needed now','When this line needs it','First plan day the line uses it; Yes when within Settings look-ahead (49 days)'),
 ('MRP by Line','Plan on this line','Who / when on this line','Usage windows and customers in date order; "Same as today" = unchanged'),
 ('MRP by Line','This line: daily target need','This line\'s need','Line\'s customer daily target × usage per module'),
 ('MRP by Line','In transit usable by this line','Supply this line may use','MIS in transit from suppliers allowed for this line (Settings ALLOWED VENDORS / Vendor priority)'),
 ('MRP Daily (hidden)','Gross requirement / Projected available balance','Day-by-day MRP record','Plan demand per day; balance = previous − requirement + in transit'),
 ('MRP Daily (hidden)','Firm receipt / Planned order receipt','Open PO / order suggestion','Open PO without vessel; planned receipt keeps stock at target days (lead-time offset)'),
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
    d['A1']='Definitions — every header and how it is calculated'; d['A2']='Sources: MIS RM Dashboard (MB52, In Transit, RM Status), WSA RM planning file, MRP.xlsx (BOM, Production Plan, Stock Count, Settings), Customer_Tracker files.'
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
        ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor=NAVY
        if ws.title in hidden_ and ws.title!='Definitions': continue
        big_=ws.title in ('MRP Daily','MRP Weekly')                     # written in Calibri already: no cell-by-cell font pass
        own=ws.title in ('Stock Management','Risk Charts','Cover Heatmap','Stock Check','RM Issues')   # already in house style
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
        for row in ws.iter_rows():
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
    if 'In Transit Status' in wb.sheetnames:                          # In Transit Status and every sheet to its right: hidden (your choice)
        cut=wb.sheetnames.index('In Transit Status')
        for k_,w_ in enumerate(wb.worksheets):
            w_.sheet_state='hidden' if k_>=cut else 'visible'
        wb.active=0
    pass                                                           # saved once at the end (post_flush)


# ===================== STOCK MANAGEMENT dashboard (Store > Warehouse > Inventory > Production > Quality > Aging) =====================
def add_stock_mgmt_dashboard(path, A, AL_ORDER, mlog, ships, d0, nd, mis_path):
    """One row per raw material, grouped like a stock management system:
       STORE opening stock (previous MRP run) · WAREHOUSE stock in (MIS arrivals since then) / stock out (implied) / closing ·
       INVENTORY reorder level · PRODUCTION need per day · QUALITY status + on hold (MB52 blocked + QI) ·
       AGING last arrival / days of supply / movement class / recommended action / stock value. Summary boxes + shortage watch below."""
    from openpyxl.styles import Border, Side
    wb=post_wb(path)
    if 'Stock Management' in wb.sheetnames: del wb['Stock Management']
    ws=wb.create_sheet('Stock Management'); ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor='14213D'
    F=lambda **k: Font(name='Calibri',**k)
    NAVY='14213D'; INK='1F2D3D'; MUTED='8A949E'; LINE='E3E8EE'
    STORE,WH,INV,PROD,QUAL,AGE=('2E86C1','D6EAF8'),('229954','D5F5E3'),('E67E22','FDEBD0'),('6C3483','E8DAEF'),('17A589','D1F2EB'),('1B2631','D6DBDF')
    def fill(c,col): c.fill=PatternFill('solid',fgColor=col)
    thin=Side(style='thin',color=LINE); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    widths=[5,16,18,14,15,11,10,10,11,11,11,11,10,11,9,10,12]
    for j,w_ in enumerate(widths,1): ws.column_dimensions[L(j)].width=w_
    NC=len(widths)
    # ---------- data ----------
    base=[x for x in AL_ORDER if not x.get('sup_row') and is_rm(x['b']['mat']) and at_risk(x,d0)]      # at-risk raw materials only
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
    grp={}
    if mis:
        for r in list(mis['Standard BOM'].iter_rows(values_only=True))[1:]:
            if r and len(r)>8 and r[3] and r[8]: grp.setdefault(str(r[3]).strip(),str(r[8]).strip())
    prev={}
    for e in sorted(mlog or [],key=lambda e:e['date']):
        if e['date']<d0: prev[e['code']]=e
    rows=[]; agg=dict(op=0.0,inn=0.0,out=0.0,cl=0.0)
    for i,x in enumerate(base,1):
        b=x['b']; m=b['mat']; pr=price.get(m,0.0); need=x.get('tgt_need') or 0.0; lt=int(b.get('lt_used') or 0)
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
        rows.append(dict(i=i,code=m,it=float(x.get('tr_tot') or 0),nxt=nxt_,ovd=ovd_,name=str(b.get('common') or m),cat=(MAT_GROUP.get(m) or (None,''))[1] or grp.get(m) or next((grp[c_] for c_,t_ in (BOM_ALIAS[0] or {}).items() if t_==m and c_ in grp),'') or '-',uom=b.get('uom'),
                         op=opening,sin=sin,sout=sout,close=close,rop=rop,need=need,status=status,hold=hold,last=last,dos=dos,mv=mv,val=val,k=k_,lt=lt,
                         alt=strip_code(str(AL_VAL.get((m,'Alternative BOM (yours): code - stock - customers')) or '').split('\n')[0]) or ''))
    # ---------- header banner ----------
    for rr in (1,2,3):
        for j in range(1,NC+1): fill(ws.cell(rr,j),NAVY)
    ws.merge_cells('A1:K1'); c=ws['A1']; c.value='  📦  Stock Management System  ·  WSA Raw Materials'; c.font=F(size=24,bold=True,color='FFFFFF'); c.alignment=Alignment(vertical='center')
    ws.merge_cells('A2:K2'); c=ws['A2']; c.value='      Store   |   Warehouse   |   Inventory   |   Production   |   Quality'; c.font=F(size=11,bold=True,color='FFFFFF')
    ws.merge_cells('A3:K3'); c=ws['A3']; c.value='      Right stock · Right time · Right place'; c.font=F(size=10,italic=True,color='D6DBDF')
    ws.merge_cells('O1:Q1'); c=ws['O1']; c.value='📅  Last updated'; c.font=F(size=11,color='D6DBDF'); c.alignment=Alignment(horizontal='right',vertical='bottom')
    ws.merge_cells('O2:Q2'); c=ws['O2']; c.value=d0.to_pydatetime(); c.number_format='dd-mm-yyyy'; c.font=F(size=16,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='right')
    ws.merge_cells('O3:Q3'); c=ws['O3']; c.value=f"{len(rows)} materials at risk of shortage"; c.font=F(size=10,bold=True,color='F5B7B1'); c.alignment=Alignment(horizontal='right')
    dd_=data_dates_text(d0)
    if dd_:
        ws.merge_cells('L2:N3'); c=ws['L2']; c.value='DATA: '+' · '.join(t_ for t_,_ in dd_); c.alignment=Alignment(horizontal='right',vertical='center',wrap_text=True)
        c.font=F(size=9,bold=True,color=('F5B7B1' if any(a_>DATA_STALE_DAYS for _,a_ in dd_) else 'D6DBDF'))
    ws.row_dimensions[1].height=36; ws.row_dimensions[2].height=18; ws.row_dimensions[3].height=16
    # ---------- process chevrons ----------
    chev=[('🏪  STORE','Receive | Issue | Track',STORE,1,4),('🏭  WAREHOUSE','Store | Manage | Dispatch',WH,5,8),('📋  INVENTORY','Monitor | Reorder | Balance',INV,9,11),
          ('⚙️  PRODUCTION','Plan | Produce | Track',PROD,12,14),('🛡️  QUALITY','Check | Accept | Hold',QUAL,15,17)]
    for t1,t2,(dark,light),c1,c2 in chev:
        ws.merge_cells(start_row=5,start_column=c1,end_row=5,end_column=c2); ws.merge_cells(start_row=6,start_column=c1,end_row=6,end_column=c2)
        a=ws.cell(5,c1,t1); a.font=F(size=13,bold=True,color='FFFFFF'); a.alignment=Alignment(horizontal='center',vertical='center')
        b_=ws.cell(6,c1,t2+'   ▶'); b_.font=F(size=9,color='FFFFFF'); b_.alignment=Alignment(horizontal='center',vertical='center')
        for j in range(c1,c2+1): fill(ws.cell(5,j),dark); fill(ws.cell(6,j),dark)
    ws.row_dimensions[5].height=24; ws.row_dimensions[6].height=16
    # ---------- table headers ----------
    groups=[('',1,5,NAVY),('STORE · stock',6,6,STORE[0]),('WAREHOUSE · supply',7,9,WH[0]),('INVENTORY',10,10,INV[0]),('PRODUCTION',11,11,PROD[0]),('QUALITY',12,13,QUAL[0]),('AGING ANALYSIS',14,17,NAVY)]
    for t,c1,c2,col in groups:
        if c2>c1: ws.merge_cells(start_row=8,start_column=c1,end_row=8,end_column=c2)
        c=ws.cell(8,c1,t); c.font=F(size=10,bold=True,color='FFFFFF'); c.alignment=Alignment(horizontal='center',vertical='center')
        for j in range(c1,c2+1): fill(ws.cell(8,j),col)
    H=['Sr.','Item code','Item name','','Category','Stock','In transit','Next arrival','Overdue in transit','Reorder level','Need / day','Status','On hold','Short from','Days of supply','Shortage risk','Days incl. in transit']
    hc=[NAVY]*5+[STORE[0]]+[WH[0]]*3+[INV[0]]+[PROD[0]]+[QUAL[0]]*2+[NAVY]*4
    for j,(h,col) in enumerate(zip(H,hc),1):
        c=ws.cell(9,j,h); c.font=F(size=9,bold=True,color='FFFFFF'); fill(c,col); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=bd
    ws.merge_cells('C9:D9'); ws.row_dimensions[9].height=30
    CH={'Above reorder':('EAF2F8','1F618D'),'Below reorder':('FCF3CF','9A7D0A'),'Out of Stock':('FADBD8','C0392B'),
        'ALERT':('FADBD8','C0392B'),'Short later':('FDEBD0','B9770E'),'Under 21 d':('FCF3CF','9A7D0A'),'OK':('D5F5E3','1E8449'),'No need':('F2F4F6','8A949E')}
    r=10
    for x in rows:
        vals=[x['i'],x['code'],x['name'][:30],None,x['cat'][:16],round(x['close']),round(x['it']),(x['nxt'].to_pydatetime() if x['nxt'] is not None else 'none'),round(x['ovd']) if x['ovd'] else '-',round(x['rop']),round(x['need'],1),
              x['status'],round(x['hold']) if x['hold'] else '-',((d0+pd.Timedelta(days=x['k'])).to_pydatetime() if x['k'] is not None else 'no shortage'),(round(x['dos']) if x['dos'] is not None else '-'),x['mv'],(round(x['val']) if x['val'] is not None else '-')]
        fmts=[None,None,None,None,None,'#,##0','#,##0','dd-mmm','#,##0','#,##0','#,##0.0',None,'#,##0','dd-mmm-yy','0',None,'0']
        for j,(v,fm) in enumerate(zip(vals,fmts),1):
            c=ws.cell(r,j,v); c.font=F(size=10,color=INK); c.border=Border(bottom=thin)
            c.alignment=Alignment(horizontal=('left' if j in (2,3,5) else 'center'),vertical='center')
            if fm: c.number_format=fm
            if j in (6,):  fill(c,'F4F9FD')
            if j in (7,8,9): fill(c,'F4FBF7')
            if j==10: fill(c,'FEF8F1')
            if j==11: fill(c,'F8F4FB')
        ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=4)
        for j,key in ((12,x['status']),(16,x['mv'])):
            bg,fg=CH.get(key,('FDEBD0','B9770E')); fill(ws.cell(r,j),bg); ws.cell(r,j).font=F(size=10,bold=True,color=fg)
        if x['k'] is not None: ws.cell(r,14).font=F(size=10,bold=True,color='C0392B')
        if isinstance(vals[12],(int,float)): ws.cell(r,13).font=F(size=10,bold=True,color='B9770E')
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
    box(1,4,'⚠️  KEY RISK HIGHLIGHTS',NAVY,[('Short inside lead time (ALERT)',sum(1 for x in rows if x['mv']=='ALERT'),'0','C0392B'),('Short later (order in time)',sum(1 for x in rows if x['mv']=='Short later'),'0','B9770E'),
        ('Earliest shortage: '+(f"{f0['name'][:16]} {(d0+pd.Timedelta(days=f0['k'])):%d-%b}" if f0 else 'none'),None,None,'C0392B' if f0 else '1E8449'),
        ('Lines stopping ≤ 14 days',len(stop14),'0','C0392B' if stop14 else '1E8449'),
        ('First line stop: '+(f"{first_stop['name']} {(d0+pd.Timedelta(days=first_stop['t'][0])):%d-%b}" if first_stop and first_stop['t'][0]<nd else 'none'),None,None,'C0392B' if (first_stop and first_stop['t'][0]<nd) else '1E8449')])
    tl=short_buckets(base,d0)
    box(6,9,'⏱  SHORTAGE TIMELINE (first short day)',NAVY,[(lab.title().replace('Days','days'),v,'0',('C0392B' if i_<2 else ('B9770E' if i_<4 else '9A7D0A'))) for i_,(lab,v) in enumerate(tl[:5])])
    al=[x for x in rows if x['k'] is not None and x['k']<x['lt']]
    box(11,17,'⚠️  SHORTAGE WATCH (short inside lead time)','C0392B',
        [(f"{x['name'][:26]} – short {(d0+pd.Timedelta(days=x['k'])):%d-%b}"+(f" · alt {x['alt'].split(':')[0][:24]}" if x['alt'] else ''),None,None,'C0392B') for x in al[:5]] or [('No material short inside its lead time',None,None,'1E8449')])
    # five cards
    top2=top+7
    cards=[('🏪 STORE','Short within 7 days',tl[0][1],'0',STORE),('🏭 WAREHOUSE','Lines stopping ≤ 14 days',len(stop14),'0',WH),
           ('📋 INVENTORY','Below reorder level',sum(1 for x in rows if x['status']!='Above reorder'),'0',INV),('⚙️ PRODUCTION','Materials short in the plan',len(shorts),'0',PROD),
           ('🛡️ QUALITY','Short, with stock on hold',sum(1 for x in rows if x['hold'] and x['k'] is not None),'0',QUAL)]
    spans=[(1,4),(5,8),(9,11),(12,14),(15,17)]
    for (t,lab,v,fm,(dark,light)),(c1,c2) in zip(cards,spans):
        for rr in range(top2,top2+4):
            ws.merge_cells(start_row=rr,start_column=c1,end_row=rr,end_column=c2)
            for j in range(c1,c2+1): fill(ws.cell(rr,j),dark if rr==top2 else light)
        a=ws.cell(top2,c1,t); a.font=F(size=12,bold=True,color='FFFFFF'); a.alignment=Alignment(horizontal='center',vertical='center')
        b_=ws.cell(top2+1,c1,lab); b_.font=F(size=10,color=dark); b_.alignment=Alignment(horizontal='center')
        n_=ws.cell(top2+2,c1,v); n_.font=F(size=24,bold=True,color=dark); n_.alignment=Alignment(horizontal='center',vertical='center'); n_.number_format=fm
        ws.row_dimensions[top2+2].height=34
    ws.cell(top2+5,1,'Only materials at risk of shortage (ALERT / short later / under 21 days) · In transit = MIS pending · Overdue = MIS arrival date passed, not received · Reorder level = need/day × lead time + safety stock · '
                     'Shortage risk: ALERT = short inside the lead time · Short later · Under 21 d (no shortage in plan) · Days incl. in transit = (stock + MIS in transit) ÷ need per day.').font=F(size=8,italic=True,color=MUTED)
    wb.move_sheet('Stock Management',offset=-wb.sheetnames.index('Stock Management'))
    pass                                                           # saved once at the end (post_flush)

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
              "Action List (Stock, to OEM Stock) and MRP Daily (stock today) add up column L by Component + Type.")
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
              'RM: stock MB52 / usage MIS standard BOM (alternate codes converted with MIS\'s Adhoc ratio). CS / PM: stock = your Stock Count (Today Qty_BOM = Today Stock Qty x Quantity) / usage = your BOM. '
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
            elif su in (None,'','#N/A'): st,conv,note,su='Unit unknown','-','Stock UOM not found on your Stock Count (code not in MB52)','(unknown)'
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
            except Exception: continue
            if k<nd: arr.setdefault(str(r[39]).strip(),[0.0]*nd)[k]+=r[34]
    for x in SUBS_LINES:
        kf=0 if x['fr'] is None else max(0,(x['fr']-d0).days); kt=nd-1 if x['to'] is None else (x['to']-d0).days
        alt.setdefault(x['mrp_code'],[]).append(dict(code=x['sub_code'],factor=x['factor'],base=stock.get(x['sub_code'],0.0),arr=arr.get(x['sub_code'],[0.0]*nd),
                                                    only=set(x['lines']),win=(kf,kt),order=x['order'],pri={lk:x['order'] for lk in x['lines']}))
    return alt




DATA_DATES={}      # source -> data date (MIS, WSA, counts) - shown on the dashboards, red when older than DATA_STALE_DAYS
DATA_STALE_DAYS=2
def file_data_date(path):
    """Date the file was last saved by Excel (docProps/core.xml 'modified'), else the file date."""
    import zipfile
    try:
        x=zipfile.ZipFile(path).read('docProps/core.xml').decode('utf-8','ignore')
        m_=re.search(r'<dcterms:modified[^>]*>([^<]+)<',x)
        if m_: return pd.Timestamp(m_.group(1)[:10]).normalize()
    except Exception: pass
    try: return pd.Timestamp(dt.datetime.fromtimestamp(os.path.getmtime(path))).normalize()
    except Exception: return None
def data_dates_text(d0):
    out=[]
    for k in ('MB52','MIS','WSA','Counts'):
        v=DATA_DATES.get(k)
        if not v: continue
        if isinstance(v,tuple):
            a_,b_=v; age=(d0-a_).days
            lab_=f"Counts {a_:%d-%b}–{b_:%d-%b}" if a_!=b_ else f"Counts {a_:%d-%b}"
            out.append((lab_+f" ({age} d old)",age))
        else: age=(d0-v).days; out.append((f"{k} {v:%d-%b} ({age} d old)",age))
    return out

def estimate_counted_stock(blocks, mis_path, d0):
    """A physical count from an earlier day is rolled forward to today:
       estimated stock = count - usage per day x days since the count + goods received (MIS PPC GRN) after the count date.
       Usage per day = average Production Plan demand of the next 7 days. Fills STOCK_EST for the Stock Check sheet."""
    STOCK_EST.clear()
    pc=read_phys_counts(SETTINGS_SRC[0]) if SETTINGS_SRC and SETTINGS_SRC[0] else {}
    mis_d=file_data_date(USER_MB52[0]) if USER_MB52[0] else file_data_date(mis_path)        # stock date: your MB52 file (else the MIS file)
    DATA_DATES['MB52' if USER_MB52[0] else 'MIS']=mis_d
    if globals().get('STOCK_DATE') is not None: DATA_DATES['WSA']=pd.Timestamp(globals()['STOCK_DATE']).normalize()
    mis_age=max(0,(d0-mis_d).days) if mis_d is not None else 0
    if mis_age>=1: print(f"WARNING: {'your MB52' if USER_MB52[0] else 'MIS'} stock is {mis_age} days old (saved {mis_d:%d-%b}) - SAP stock of materials that are not counted is rolled forward by usage; refresh it for exact figures.")
    sap={}; rec={}
    try:
        mis=wb_ro(mis_path)
        for r in list(mis['MB52'].iter_rows(values_only=True))[1:]:
            if r and len(r)>14 and r[2] and isinstance(r[14],(int,float)): sap[str(r[2]).strip()]=sap.get(str(r[2]).strip(),0.0)+r[14]
        its=list(mis['In Transit'].iter_rows(values_only=True))[1:]
    except Exception: its=[]
    for b in blocks:
        m=b['mat']; cnt=None; date=None
        if m in COUNTED_RM: cnt=COUNTED_RM[m]['qty']; date=COUNTED_RM[m].get('date')
        elif b.get('count_stock') and m in pc: cnt=float(b.get('st_main') if b.get('st_main') is not None else b['stock']); date=pc[m].get('date')
        e=dict(sap=sap.get(m),count=cnt,date=date,days=None,use_day=None,use=0.0,rec=0.0,est=None,src='SAP (MB52) - not counted' if cnt is None else 'Physical count')
        if cnt is not None and date is not None:
            days=max(0,(d0-pd.Timestamp(date).normalize()).days)
            P=b.get('plan') or []; ud=(sum(P[:7])/min(7,len(P))) if P else 0.0
            rc=0.0
            for r in its:
                if len(r)>39 and str(r[39] or '').strip()==m and isinstance(r[33],(int,float)) and r[33]>0 and isinstance(r[35],(dt.datetime,dt.date)):
                    if pd.Timestamp(date)<pd.Timestamp(r[35]).normalize()<=d0: rc+=r[33]
            e.update(days=days,use_day=ud,use=ud*days,rec=rc)
            if days>0:
                est=max(0.0,cnt-ud*days+rc); delta=est-cnt
                b['stock']=float(b['stock'])+delta
                if b.get('st_main') is not None: b['st_main']=float(b['st_main'])+delta
                e.update(est=est,src=f'Physical count {pd.Timestamp(date):%d-%b}, rolled forward {days} d')
            else: e.update(est=cnt,src='Physical count today')
        if cnt is None and mis_age>=1 and not b.get('count_stock'):   # SAP stock from an old MIS file: roll forward like a count
            own=float(b['stock'])-float(b.get('st_jw') or 0)
            P=b.get('plan') or []; ud=(sum(P[:7])/min(7,len(P))) if P else 0.0
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
            except Exception: continue
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

def make_line_days_xlsx(rows, d0, nd, path, TD=None):
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    wb=Workbook(); ws=wb.active; ws.title='Line Run Days'
    ws['A1']='How many days each production line can run with the material it needs'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=('Per line: the first day a material that line needs is out of stock (projected stock of all lines together below zero on a day the line needs it). '
              'Run days count from today; "N+" = no stop within the horizon. Limiting material = what stops the line first; next = the following stops. '
              'Lines come from your Production Plan (MIS lines for customers not in it). '
              f'Status (Settings): CRITICAL under {COVER_CRITICAL} days, AT RISK under {COVER_RISK}, WATCH under {COVER_WATCH}.')
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:J2'); ws['A2'].alignment=Alignment(wrap_text=True); ws.row_dimensions[2].height=30
    H=['Production line','Customers on the line','Can run (days) - stock + in transit','Stops on','Limiting material','Next stops',
       'Can run (days) - + Open PO','Stops on (+ Open PO)','Limiting material (+ Open PO)','Status','Can run (days) - RM only','Limiting RM']
    for j,h in enumerate(H,1):
        c=ws.cell(4,j,h); c.font=Font(name=AR,size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=NAVY); c.border=bd; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.row_dimensions[4].height=36
    fmt=f'[>={nd}]"{nd}+";0'
    for i,r_ in enumerate(rows):
        kt,limt,nxt=r_['t']; ko,limo,_=r_['o']; kr,limr,_=r_['rt']
        st=line_status(kt)                                     # same rule as Planner / MRP by Line (Settings thresholds)
        if st in ('CRITICAL','AT RISK') and line_status(ko) not in ('CRITICAL','AT RISK'): st+=' (OK if Open PO ships)'
        vals=[r_['name'],', '.join(r_['cust']),kt,(d0+pd.Timedelta(days=kt)).to_pydatetime() if kt<nd else None,', '.join(limt) or '-','; '.join(nxt) or '-',
              ko,(d0+pd.Timedelta(days=ko)).to_pydatetime() if ko<nd else None,', '.join(limo) or '-',st,kr,', '.join(limr) or '-']
        for j,v in enumerate(vals,1):
            c=ws.cell(5+i,j,v); c.font=Font(name=AR,size=9,bold=j in (1,3,7,11)); c.border=bd; c.alignment=Alignment(vertical='top',wrap_text=j in (2,5,6,9,12))
            if j in (3,7,11): c.number_format=fmt; c.alignment=Alignment(horizontal='center',vertical='top')
            if j in (4,8): c.number_format='dd-mmm-yy'
        for j,k in ((3,kt),(7,ko),(11,kr)):
            ws.cell(5+i,j).fill=PatternFill('solid',fgColor=STATUS_FILL[line_status(k)])
        ws.cell(5+i,10).fill=PatternFill('solid',fgColor=STATUS_FILL[st.split(' (')[0]])
    if rows: ws.add_table(Table(displayName='LineRunDays',ref=f'A4:L{4+len(rows)}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=True)))
    for j,w in enumerate([18,34,12,11,30,44,12,11,30,16,12,30],1): ws.column_dimensions[L(j)].width=w
    ws.freeze_panes='B5'; wb.save(path)

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
            ('Line AT RISK if it can run fewer than (days)','COVER_RISK','Amber - expedite'),
            ('Line WATCH if it can run fewer than (days)','COVER_WATCH','Yellow'),
            ('Lead time raw materials (days) - if no ETD history','LEAD_TIME_RM','Dates planned orders'),
            ('Lead time packaging / consumables (days)','LEAD_TIME_CSPM',''),
            ('Daily report: list materials lasting fewer than (days) at the line daily targets','REPORT_DAYS','Stock + in transit / need at the daily targets below'),
            ('Look-ahead (days): a line counts for a material if it runs a product using it within','LOOKAHEAD_DAYS','Daily target need + MRP by Line "Needed now"'),
            ('Safety stock (days of need at the daily targets)','SAFETY_DAYS','Stock Management (reorder level)'),
            ('Days of cover: average plan usage over the next (days)','COVER_AVG_DAYS','MRP Daily / Weekly "Days of cover" = projected stock ÷ average usage/day of the next N days')]
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
    except Exception: return None
def read_settings(paths):
    """Settings sheet: from your MRP.xlsx if it has one, else from the previous MRP_Daily. Sets the rules for this run.
       Sections: basic numbers, unlimited materials, allowed vendors (by customer), substitutes, line ranking, advanced numbers.
       Old sections (Customer Supplier, Pack conversions, Vendor Priority) are still read and turned into the new ones."""
    g=globals(); src=None
    for pth in paths:
        if not pth or not os.path.exists(pth): continue
        try: w=wb_ro(pth)
        except Exception: continue
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
    g['PACK_FACTORS']=pf; g['SUBS_LINES']=lines
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
       REPORT_DAYS days at the line daily targets (Settings), shortest first."""
    tg=([(c_,v_) for c_,l_,v_,o_ in CUST_TARGET] if CUSTOMER_TARGETS else [(lab.split(' Production')[0].replace('USA ','').replace(' (plan peak)','*'),v) for _,(lab,v) in sorted(EFF_TARGET.items()) if v])
    low=sorted([x for x in A if x.get('tgt_days') is not None and x['tgt_days']<REPORT_DAYS],key=lambda x:((x.get('tr_tot') or 0)>0,x['tgt_days']))
    n=len(low)
    subj=f"Material report {d0:%d %b}: {n} under {REPORT_DAYS} days" if n else f"Material report {d0:%d %b}: all OK ({REPORT_DAYS}+ days)"
    body=[]
    if not tg: body.append('No line daily targets on Settings - no report.')
    else:
        body.append(f"Under {REPORT_DAYS} days at daily target - nothing in transit first")
        body.append('('+' / '.join(f'{a} {v:,.0f}' for a,v in tg)+')')
        if any(a.endswith('*') for a,_ in tg): body.append('* no target on Settings: highest planned day')
        body.append('')
        for i,x in enumerate(low,1):
            b=x['b']; u=b.get('uom') or ''
            body.append(f"{i}) {b['common']}")
            body.append(f"   {x['tgt_days']:.1f} days | stock {b['stock']:,.0f} | in transit {x.get('tr_tot') or 0:,.0f} {u}".rstrip())
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

def make_start_here_xlsx(path, info):
    from openpyxl.worksheet.hyperlink import Hyperlink   # internal links (no relationship part - safe to graft)
    """Start Here: 1) can I trust today's data?  2) the daily workflow with links, in the order planners in large companies use."""
    from openpyxl.styles import Border, Side
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    G,Y,R='C6EFCE','FFEB9C','FFC7CE'
    wb=Workbook(); ws=wb.active; ws.title='Start Here'; ws.sheet_view.showGridLines=False
    d0=info['d0']; A=info['A']; LR=info['LR'] or []
    ws['A1']=f'MRP Daily - {d0:%A %d %b %Y}'; ws['A1'].font=Font(name=AR,size=18,bold=True,color=NAVY)
    ws['A2']='Start here every morning: 1) check the data can be trusted, 2) work the steps below in order, 3) take Today to the daily meeting.'
    ws['A2'].font=Font(name=AR,size=10,italic=True,color='595959')
    # ---- headline ----
    crit=sum(1 for x in A if x['status']=='CRITICAL'); risk=sum(1 for x in A if x['status']=='AT RISK')
    exp=sum(1 for x in A if x.get('exg')==1); rel=sum(1 for x in A if x.get('exg')==3)
    stop=[r_ for r_ in LR if line_status(r_['t'][0]) in ('CRITICAL','AT RISK')]
    hl=[('Critical materials',crit,R if crit else G),('At risk',risk,Y if risk else G),('Expedite / reschedule in',exp,R if exp else G),
        ('Orders to release this week',rel,Y if rel else G),(f'Lines that can run under {COVER_RISK} days',len(stop),R if stop else G)]
    r=4
    for j,(lab,v,col) in enumerate(hl):
        c=ws.cell(r,1+j*2,v); c.font=Font(name=AR,size=20,bold=True,color=NAVY); c.alignment=Alignment(horizontal='center'); c.fill=PatternFill('solid',fgColor=col)
        ws.merge_cells(start_row=r,start_column=1+j*2,end_row=r,end_column=2+j*2)
        c=ws.cell(r+1,1+j*2,lab); c.font=Font(name=AR,size=9,bold=True); c.alignment=Alignment(horizontal='center',wrap_text=True); c.fill=PatternFill('solid',fgColor=col)
        ws.merge_cells(start_row=r+1,start_column=1+j*2,end_row=r+1,end_column=2+j*2); ws.row_dimensions[r+1].height=28
    if stop:
        ws.cell(r+2,1,'Lines at risk: '+';  '.join(f"{x['name']} {x['t'][0]} d ({', '.join(x['t'][1])[:40]})" for x in stop[:6])).font=Font(name=AR,size=10,bold=True,color='9C0006')
    # ---- top risks: what runs out, from when to when ----
    r=8; ws.cell(r,1,'TOP RISKS - what runs out, from when to when (all items: Action List)').font=Font(name=AR,size=13,bold=True,color=NAVY); r+=1
    cols=[('Status',1,1),('Material',2,4),('Short from -> to (with in transit)',5,7),('In transit (MIS)',8,8),('Open PO (WSA)',9,9),('Exception message',10,13)]
    for t_,a_,z_ in cols:
        c=ws.cell(r,a_,t_); c.font=Font(name=AR,size=9,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=NAVY); c.alignment=Alignment(horizontal='center',wrap_text=True)
        if z_>a_: ws.merge_cells(start_row=r,start_column=a_,end_row=r,end_column=z_)
    r+=1; stf=STATUS_FILL
    for x in [x for x in A if x['status'] in ('CRITICAL','AT RISK')][:12] or A[:5]:
        vals=[(x['status'],1,1),(x['b']['common'] if not x['b'].get('count_stock') else x['b']['desc'],2,4),(x.get('short_txt') or 'no stock-out',5,7),
              (round(x.get('tr_tot') or 0),8,8),(round(x.get('op_tot') or 0),9,9),(x.get('exc') or '',10,13)]
        for v,a_,z_ in vals:
            c=ws.cell(r,a_,v); c.font=Font(name=AR,size=9,bold=(a_ in (1,5)),color=('9C0006' if a_==5 and v!='no stock-out' else '000000')); c.border=bd
            c.alignment=Alignment(vertical='center',wrap_text=True,horizontal='center' if a_ in (1,8,9) else 'left')
            if a_ in (8,9): c.number_format='#,##0;-#,##0;"-"'
            if z_>a_: ws.merge_cells(start_row=r,start_column=a_,end_row=r,end_column=z_)
        ws.cell(r,1).fill=PatternFill('solid',fgColor=stf.get(x['status'],'FFFFFF')); ws.row_dimensions[r].height=26; r+=1
    # ---- step 0: data trust ----
    r+=1; ws.cell(r,1,'STEP 0 - CAN I TRUST TODAY\'S DATA?  (2 minutes)').font=Font(name=AR,size=13,bold=True,color=NAVY); r+=1
    for j,t in enumerate(['Input','File','Last saved','Data as of','Check']):
        c=ws.cell(r,1+j*2 if j<2 else 4+j,t) if False else ws.cell(r,[1,3,7,8,9][j],t); c.font=Font(name=AR,size=10,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=NAVY); c.border=bd
    ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=6); r+=1
    today=pd.Timestamp(dt.date.today())
    for lab,pth,asof in info['files']:
        ok=True; msg='OK'
        if not pth or not os.path.exists(str(pth)):
            ok,msg=(False,'no Settings sheet yet - defaults used') if lab.startswith('Settings') else (None,'not found')
        else:
            mt=pd.Timestamp(dt.datetime.fromtimestamp(os.path.getmtime(pth)))
            age=(today-mt.normalize()).days
            if age>1: ok=False; msg=f'saved {age} days ago - is it the latest?'
            if asof is not None and (today-pd.Timestamp(asof).normalize()).days>2: ok=False; msg=f'data as of {pd.Timestamp(asof):%d-%b} - old?'
        vals=[lab,os.path.basename(str(pth)) if pth else '-',(dt.datetime.fromtimestamp(os.path.getmtime(pth)) if pth and os.path.exists(str(pth)) else None),
              (pd.Timestamp(asof).to_pydatetime() if asof is not None else None),msg]
        for v,cn in zip(vals,[1,3,7,8,9]):
            c=ws.cell(r,cn,v); c.font=Font(name=AR,size=10); c.border=bd
            if cn==7: c.number_format='dd-mmm hh:mm'
            if cn==8: c.number_format='dd-mmm-yy'
        ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=6)
        ws.cell(r,9).fill=PatternFill('solid',fgColor=G if ok else (Y if ok is False else R)); r+=1
    chk=info.get('check') or {}
    for lab,v,warn,link in [
                            ('High data issues',chk.get('High',0),'see Check - TOP ISSUES','Check'),
                            ('Notes from this run',len(info['notes']),'listed below',None)]:
        ws.cell(r,1,lab).font=Font(name=AR,size=10); ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=6)
        c=ws.cell(r,7,v); c.font=Font(name=AR,size=10,bold=True); c.alignment=Alignment(horizontal='center')
        c=ws.cell(r,9,warn if v else 'OK'); c.fill=PatternFill('solid',fgColor=Y if v else G); c.font=Font(name=AR,size=10,underline='single' if (v and link) else None,color='0563C1' if (v and link) else '000000')
        if v and link: c.hyperlink=Hyperlink(ref=c.coordinate,location=f"'{link}'!A1")
        r+=1
    for n in info['notes'][:8]:
        ws.cell(r,1,'• '+n[:220]).font=Font(name=AR,size=9,italic=True,color='9C5700'); ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=10); r+=1
    # ---- daily workflow ----
    r+=1; ws.cell(r,1,'DAILY WORKFLOW - work down the list').font=Font(name=AR,size=13,bold=True,color=NAVY); r+=1
    steps=[('1','Action List','Work the exceptions top-down: EXPEDITE / FAST-TRACK / RESCHEDULE IN first, then RELEASE ORDER. Fill Owner, Due, Status update.','SAP MD06 / MD07'),
           ('2','MRP Daily','Before you act on an item, check its full picture: requirement, receipts, projected balance, planned orders.','SAP MD04'),
           ('3','Planner','Translate shortages into line impact: which line stops, when, on what.','Line / capacity view'),
           ('4','MRP by Line','Per line: the materials it needs, what is short and where it is supplied from. Day by day: MRP Daily, filter column H on the line - red cells = days that line is not covered, next to the receipts.','Line detail'),
           ('5','Today','The 3-5 decisions for the daily meeting with production, purchasing and logistics.','Daily stand-up')]
    for j,t in enumerate(['Step','Open','What to do','Like in SAP / big companies']):
        c=ws.cell(r,[1,2,3,9][j],t); c.font=Font(name=AR,size=10,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=NAVY); c.border=bd
    ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=8); r+=1
    for n_,sh,what,like in steps:
        ws.cell(r,1,n_).font=Font(name=AR,size=12,bold=True,color=NAVY); ws.cell(r,1).alignment=Alignment(horizontal='center',vertical='center')
        c=ws.cell(r,2,sh); c.hyperlink=Hyperlink(ref=c.coordinate,location=f"'{sh}'!A1"); c.font=Font(name=AR,size=10,bold=True,underline='single',color='0563C1')
        c=ws.cell(r,3,what); c.font=Font(name=AR,size=10); c.alignment=Alignment(wrap_text=True,vertical='center'); ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=8)
        ws.cell(r,9,like).font=Font(name=AR,size=9,italic=True,color='595959')
        for cc in (1,2,3,9): ws.cell(r,cc).border=bd
        ws.row_dimensions[r].height=32; r+=1
    r+=1; ws.cell(r,1,'WEEKLY').font=Font(name=AR,size=13,bold=True,color=NAVY); r+=1
    for sh,what in [('Check','TOP ISSUES: data fixes for MIS / Engineering / your MRP.xlsx; at the bottom (+ to open): in transit MIS vs WSA, supplier coverage'),
                    ('MRP Weekly','Week view for the S&OP / weekly planning meeting'),('No Demand','Slow stock to use up, transfer or return (monthly)'),('Settings','Rules: targets, allowed vendors, substitutes, line ranking')]:
        c=ws.cell(r,2,sh); c.hyperlink=Hyperlink(ref=c.coordinate,location=f"'{sh}'!A1"); c.font=Font(name=AR,size=10,bold=True,underline='single',color='0563C1')
        ws.cell(r,3,what).font=Font(name=AR,size=10); ws.merge_cells(start_row=r,start_column=3,end_row=r,end_column=10); r+=1
    for col,w in zip('ABCDEFGHIJKLM',[22,16,12,12,12,12,16,12,34,12,12,12,12]): ws.column_dimensions[col].width=w
    set_print(ws,'Start Here'); wb.save(path)

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
    except Exception: pass
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

def make_line_view_xlsx(A, LR, d0, nd, path, TD=None, AL_COLS=None):
    """MRP by Line: one row per material and production line, sorted like the Action List (material, then line).
       Line columns: how long the line can run, this line's need at its daily target and at its daily plan (average of the next
       21 days with production), when this line is short (allocated) and where it is supplied from.
       Then the Action List columns Stock ... Daily plan with OEM for the material (formulas = the Action List row),
       merged over the material's lines but with the value kept in every cell, so filtering a line still shows them."""
    from openpyxl.styles import Border, Side
    AR='Calibri'; NAVY='1F3864'; TEAL='1F6E6E'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    wb=Workbook(); ws=wb.active; ws.title='MRP by Line'; ws.sheet_view.showGridLines=False
    ws['A1']='MRP by Line - each material per production line, with the Action List numbers of the material'; ws['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    ws['A2']=('Filter column A (line) and / or column D (component). Teal columns = this line only. The other columns are the material totals from the Action List '
              '(same formulas: change stock on Stock Data and they follow); they are merged over the material\'s lines, and each cell still holds the number, so a filter keeps it visible. '
              '')
    ws['A2'].font=Font(name=AR,size=9,italic=True); ws.merge_cells('A2:N2'); ws['A2'].alignment=Alignment(wrap_text=True,vertical='top'); ws.row_dimensions[2].height=40
    LCOL=['Production line','Line can run (days)','Line stops on','Component','Component description','UOM','First needed (this line)',f'Needed now (next {int(LOOKAHEAD_DAYS)} days)',
          'Plan on this line (who / when) - differs from today?',
          'This line: daily target need','This line: short from (allocated)','This line: supplied from (in order)']
    AL_COLS=AL_COLS or []
    H=LCOL+[('In transit usable by this line (MIS, allowed suppliers)' if h_=='In transit today (total qty, MIS)' else h_) for h_,_ in AL_COLS]
    for j_,h in enumerate(H,1):
        c=ws.cell(4,j_,h); c.border=bd; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        if j_<=len(LCOL): col_,txt_=(TEAL if j_>=7 else NAVY),'FFFFFF'
        else: col_=al_hdr_color(h); txt_='1F1F1F' if al_hdr_light(h) else 'FFFFFF'
        c.fill=PatternFill('solid',fgColor=col_); c.font=Font(name=AR,size=9,bold=True,color=txt_)
    ws.row_dimensions[4].height=48
    info={r_['key']:r_ for r_ in (LR or [])}; lorder={r_['key']:i_ for i_,r_ in enumerate(LR or [])}
    tg_row={(t['code'],t['lk']):t['row'] for t in (CALC_LAYOUT[0] or [])}
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
            prow=PLAN_LINE_ROWS.get((b['mat'],lk)); c21=L(3+min(PLAN_AVG_DAYS,nd)-1)
            fut_txt,fut_diff=future_plan_text(b,lk,d0,nd,LINE_TODAY.get(lk))
            vals=[(L_['name'] if L_ else lk),(kt if kt<nd else f'{nd}+'),((d0+pd.Timedelta(days=kt)).to_pydatetime() if kt<nd else None),b['mat'],
                  (b['desc'] if b.get('count_stock') else b['common']),b.get('uom'),((d0+pd.Timedelta(days=f_)).to_pydatetime() if f_ is not None else None),
                  (('Yes' if (f_ is not None and f_<int(LOOKAHEAD_DAYS)) else 'No') if VALUES_ONLY else f'=IF(AND(ISNUMBER(G{r}),G{r}<DATE({d0.year},{d0.month},{d0.day})+{TGT_CELL.get("_look") or int(LOOKAHEAD_DAYS)}),"Yes","No")'),
                  fut_txt,
                  ((round((x.get('tgt_parts') or {}).get(lk),1) if (x.get('tgt_parts') or {}).get(lk) else None) if VALUES_ONLY else (f"='Calc Target'!F{tg_row[(b['mat'],lk)]}" if (b['mat'],lk) in tg_row else None)),
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
                if j_==11 and isinstance(v,dt.datetime): c.number_format='dd-mmm'; c.font=Font(name=AR,size=9,bold=True,color='9C0006')
                if j_>len(LCOL):
                    hh=H[j_-1]; c.number_format=('0.0' if hh.startswith('Inv. Days') else ('#,##0.0' if 'Daily' in hh else '#,##0;-#,##0;"-"'))
                    if hh.startswith('Inv. Days') and isinstance(v,(int,float)) and v<REPORT_DAYS:     # under 21 days: red highlight
                        c.fill=PatternFill('solid',fgColor='FFC7CE'); c.font=Font(name=AR,size=9,bold=True,color='9C0006')
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
    plan={}; cust={}; cl={}; how={}
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
                        transit=[0.0]*nd,adhoc=[0.0]*nd,openpo=[0.0]*nd,count_stock=True,raw=raw.get(c,'not on Stock Count'),
                        st_main=stock.get(c,0.0),st_wh=0.0,st_jw=0.0,plan_all={pd.Timestamp(dates[k]).normalize():P[k] for k in range(nd) if P[k]}))
    return out,iss

def make_bom_sheet(bom, mis_names, path):
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.table import Table, TableStyleInfo
    wb=Workbook(); ws=wb.active; ws.title='BOM'
    ws['A1']=(f'Copy of the BOM in {BOM_SRC_NAME[0]} - edit it THERE. MIS BOM Name (link to the MIS plan) comes from your Production Plan (customers using each Module Model).'
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
    ws['A1']=(f'Copy of the Stock Count in {BOM_SRC_NAME[0]} - update it THERE (Today Stock Qty x Quantity = Today Qty_BOM, used by the MRP).'
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
    except Exception: return rows
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
    ws['A2']=('From the Customer Supplier sheet. Run-out date = first day that supplier\'s projected stock (stock + in transit) goes below zero. '
              'Status = same rule as the Action List, against the material\'s own target days (Settings). Red run-out date = inside that target. Stock by supplier = MIS RM Status by Make; "Unassigned" stock is not counted for any supplier.')
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
    except Exception: return None
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
    lines={}; modplan={}; issues=[]
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
        rows.append(dict(line=ln,cust=cu,mis_cust=mc,bom=bn,v=v,model=mdl_))
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
        ws['A1']=(f'Modules per day from your Production Plan ({src}, columns A-F) by Production Line and Customer. '
                  'MIS Customer / MIS BOM Name = the MIS name it is linked to (used with the BOM sheet and the Customer Supplier sheet).')
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
def _style_header_cell(c, font_name='Arial', bg='1F4E79'):
    """Style a header cell (white text, dark background, centered, wrapped)."""
    from openpyxl.styles import Border, Side
    c.font = Font(name=font_name, size=9, bold=True, color='FFFFFF')
    c.fill = PatternFill('solid', fgColor=bg)
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = Border(left=Side(style='thin', color='D0D0D0'), right=Side(style='thin', color='D0D0D0'),
                      top=Side(style='thin', color='D0D0D0'), bottom=Side(style='thin', color='D0D0D0'))

def add_cover_chart(ws, A, TD, top, AR='Calibri', NAVY='1F3864'):
    """All-materials picture: days of cover today (stock + in transit, as MRP Daily), lowest at the top, red below the target."""
    from openpyxl.chart import BarChart, Reference
    from openpyxl.chart.series import DataPoint
    from openpyxl.chart.label import DataLabelList
    ws.cell(top-1,1,f'Days of cover today by material (stock + in transit, as MRP Daily) - red bars are below that material\'s own target days (Settings)').font=Font(name=AR,size=12,bold=True,color=NAVY)
    dc=40; CAP=60
    ws.cell(1,dc,'Material'); ws.cell(1,dc+1,'Days of cover')
    srt=sorted(A,key=lambda x:x['doc'])
    for i,x in enumerate(srt):
        nm=str(x['b']['common'] if not x['b'].get('count_stock') else x['b']['desc'] or x['b']['mat']); nm=nm if len(nm)<=34 else nm[:32]+'..'
        ws.cell(2+i,dc,nm); ws.cell(2+i,dc+1,round(min(CAP,max(0,x['doc'])),1))
    for cc in range(dc,dc+2): ws.column_dimensions[L(cc)].hidden=True
    bar=BarChart(); bar.type='bar'; bar.style=10; bar.grouping='clustered'
    bar.title=f'Days of cover today (stock + in transit), capped at {CAP}'
    bar.x_axis.title='Material'; bar.y_axis.title='Days of cover'
    bar.add_data(Reference(ws,min_col=dc+1,min_row=1,max_row=1+len(srt)),titles_from_data=True)
    bar.set_categories(Reference(ws,min_col=dc,min_row=2,max_row=1+len(srt)))
    se=bar.series[0]; se.graphicalProperties.solidFill='70AD47'; se.graphicalProperties.line.noFill=True
    for i,x in enumerate(srt):
        if x['doc']<(x['b'].get('td') or TD):
            pt=DataPoint(idx=i); pt.graphicalProperties.solidFill='C00000'; pt.graphicalProperties.line.noFill=True; se.dPt.append(pt)
    se.dLbls=DataLabelList(); se.dLbls.showVal=True; se.dLbls.showSerName=False; se.dLbls.showCatName=False; se.dLbls.showLegendKey=False
    bar.x_axis.scaling.orientation='maxMin'; bar.y_axis.crosses='max'
    bar.y_axis.scaling.min=0; bar.y_axis.scaling.max=CAP; bar.y_axis.majorUnit=10
    bar.x_axis.delete=False; bar.y_axis.delete=False; bar.legend=None; bar.gapWidth=40
    bar.height=max(9,0.55*len(srt)+3); bar.width=26; bar.visible_cells_only=False
    ws.add_chart(bar,f'A{top}')

def build_today_sheet(wb, A, AR='Calibri'):
    """Create the Today sheet: materials that need a decision today."""
    from openpyxl.chart import BarChart, Reference
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    thin = Side(style='thin', color='BFBFBF')
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    today = wb.create_sheet('Today', 0)
    today.sheet_view.showGridLines = False
    today.page_setup.orientation = 'landscape'
    today.page_setup.fitToPage = True
    today.page_setup.fitToWidth = 1
    today.page_setup.paperSize = today.PAPERSIZE_TABLOID
    today.sheet_properties.pageSetUpPr.fitToPage = True
    today.sheet_properties.tabColor = 'C00000'
    today['A1'] = 'TODAY — What needs a decision'
    today['A1'].font = Font(name=AR, size=20, bold=True, color='FFFFFF')
    today['A1'].fill = PatternFill('solid', fgColor='1F4E79')
    today['A1'].alignment = Alignment(horizontal='left', vertical='center')
    today['A2'] = (f"Materials that are CRITICAL or AT RISK (same status as the Action List: cover vs each material's own target days on the Settings sheet), "
                   f"plus any material more than {SHORT_WATCH} days short of its target.   Run {dt.date.today():%d-%b-%Y}")
    today['A2'].font = Font(name=AR, size=9, italic=True, color='595959')
    today['A2'].fill = PatternFill('solid', fgColor='F8F8F8')
    today.merge_cells('A1:L1')
    today.merge_cells('A2:L2')
    today.row_dimensions[1].height = 32
    today.row_dimensions[2].height = 20
    today.oddHeader.left.text = 'MRP Daily'
    today.oddHeader.right.text = '&D'
    today.oddFooter.center.text = 'Page &P of &N'
    today.print_title_rows = '1:4'
    items = [x for x in A if x['status'] in ('CRITICAL', 'AT RISK') or x.get('short_d', 0) > SHORT_WATCH]
    items.sort(key=lambda x: (STATUS_ORDER[x['status']], x['cov'][1], -x.get('short_d', 0)))
    headers = ['Material', 'Status', 'Class', 'Target days', 'Cover (days, stock + in transit)', 'Days short of target', 'First short of target on',
               'Can still run from first short date (days)', 'Next arrival', 'Recommended action', 'Owner', 'Due']
    for j, h in enumerate(headers, 1):
        _style_header_cell(today.cell(4, j, h), AR)
    today.row_dimensions[4].height = 36
    for i, x in enumerate(items):
        rr = 5 + i
        st = x['status']
        cover = x['cov'][1]
        short_d = x.get('short_d', 0)
        first_miss = x.get('below')
        cover_first = x.get('cover_from_first')
        next_arr = (f"{x['nx']['date']:%d-%b}: {x['nx']['qty']:,.0f}" if x.get('nx') else 'None')
        abc = x['b'].get('abc', '')
        xyz = x['b'].get('xyz', '')
        klass = f"{abc}{('-' + xyz) if xyz and xyz != '-' else ''}"
        action, fill_c = STATUS_ACTION[st], STATUS_FILL[st]          # one rule: the material's status
        if x.get('exg', 9) <= 4:
            action = x['exc']
        vals = [x['b']['common'] if not x['b'].get('count_stock') else x['b']['desc'], st, klass, x['b'].get('td'), round(cover, 1), round(short_d, 1),
                first_miss.to_pydatetime() if first_miss is not None else None,
                cover_first, next_arr, action, '', '']
        for j, v in enumerate(vals, 1):
            c = today.cell(rr, j, v)
            c.font = Font(name=AR, size=9)
            c.alignment = Alignment(vertical='center', wrap_text=(j in (9, 10)))
            c.border = bd
            if j == 2:
                c.fill = PatternFill('solid', fgColor=fill_c); c.font = Font(name=AR, size=9, bold=True, color=STATUS_FONT[st])
                c.alignment = Alignment(horizontal='center', vertical='center')
            if j in (5, 6):
                c.number_format = '0.0'
                c.fill = PatternFill('solid', fgColor=fill_c)
                c.alignment = Alignment(horizontal='center', vertical='center')
            if j == 8 and isinstance(v, (int, float)):
                c.number_format = '0'
                c.alignment = Alignment(horizontal='center', vertical='center')
            if j == 7 and v:
                c.number_format = 'dd-mmm-yy'
                c.alignment = Alignment(horizontal='center', vertical='center')
            if j in (3, 4):
                c.alignment = Alignment(horizontal='center', vertical='center')
        today.row_dimensions[rr].height = 22
    if items:
        today.add_table(Table(displayName='TodayList', ref=f'A4:L{4+len(items)}',
                              tableStyleInfo=TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)))
        chart = BarChart()
        chart.type = 'bar'
        chart.style = 10
        chart.title = 'Days of cover (stock + in transit)'
        chart.legend = None
        chart.gapWidth = 40
        data = Reference(today, min_col=5, min_row=4, max_row=4+len(items))
        cats = Reference(today, min_col=1, min_row=5, max_row=4+len(items))
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.height = max(10, 0.55 * len(items) + 2)
        chart.width = 16
        today.add_chart(chart, 'N3')
    for j, w in enumerate([30, 11, 8, 9, 12, 12, 14, 18, 22, 44, 14, 12], 1):
        today.column_dimensions[L(j)].width = w
    today.freeze_panes = 'A5'

def build_planner_sheet(wb, LR, d0, nd, AR='Calibri'):
    """Create the Planner sheet: one row per production line."""
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.table import Table, TableStyleInfo
    thin = Side(style='thin', color='BFBFBF')
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    planner = wb.create_sheet('Planner', 0)
    planner.sheet_view.showGridLines = False
    planner.page_setup.orientation = 'landscape'
    planner.page_setup.fitToPage = True
    planner.page_setup.fitToWidth = 1
    planner.page_setup.paperSize = planner.PAPERSIZE_TABLOID
    planner.sheet_properties.pageSetUpPr.fitToPage = True
    planner.sheet_properties.tabColor = 'C00000'
    planner['A1'] = 'PLANNER — Production line by production line'
    planner['A1'].font = Font(name=AR, size=20, bold=True, color='FFFFFF')
    planner['A1'].fill = PatternFill('solid', fgColor='1F4E79')
    planner['A1'].alignment = Alignment(horizontal='left', vertical='center')
    planner['A2'] = (f"Which line stops, when, and because of which material. Status (Settings): CRITICAL under {COVER_CRITICAL} days, "
                     f"AT RISK under {COVER_RISK}, WATCH under {COVER_WATCH}.   Run {dt.date.today():%d-%b-%Y}")
    planner['A2'].font = Font(name=AR, size=9, italic=True, color='595959')
    planner['A2'].fill = PatternFill('solid', fgColor='F8F8F8')
    planner.merge_cells('A1:H1')
    planner.merge_cells('A2:H2')
    planner.row_dimensions[1].height = 32
    planner.row_dimensions[2].height = 20
    planner.oddHeader.left.text = 'MRP Daily'
    planner.oddHeader.right.text = '&D'
    planner.oddFooter.center.text = 'Page &P of &N'
    planner.print_title_rows = '1:4'
    headers = ['Production line', 'Customers', 'Can run (days)', 'Stops on', 'Limiting material', 'Next material that stops', 'What to do', 'Status']
    for j, h in enumerate(headers, 1):
        _style_header_cell(planner.cell(4, j, h), AR)
    planner.row_dimensions[4].height = 32
    if LR:
        for i, r_ in enumerate(LR):
            rr = 5 + i
            kt, limt, nxt = r_['t']
            stop_date = (d0 + pd.Timedelta(days=kt)).to_pydatetime() if kt < nd else None
            next_stop = '; '.join(nxt) if nxt else '-'
            st = line_status(kt)                                   # one rule for every line view (Settings thresholds)
            what, fill_c = STATUS_ACTION[st], STATUS_FILL[st]
            lim_txt = ', '.join(f"{m_} ({kt if kt < nd else str(nd) + '+'} d)" for m_ in limt) or '-'   # same style as 'Next material that stops'
            vals = [r_['name'], ', '.join(r_['cust']), kt, stop_date, lim_txt, next_stop, what, st]
            for j, v in enumerate(vals, 1):
                c = planner.cell(rr, j, v)
                c.font = Font(name=AR, size=9, bold=(j in (1, 3)))
                c.alignment = Alignment(vertical='center', wrap_text=(j in (2, 5, 6, 7)))
                c.border = bd
                if j == 3:
                    c.number_format = '0'
                    c.fill = PatternFill('solid', fgColor=fill_c)
                    c.alignment = Alignment(horizontal='center', vertical='center')
                if j == 4 and v:
                    c.number_format = 'dd-mmm-yy'
                    c.alignment = Alignment(horizontal='center', vertical='center')
            planner.row_dimensions[rr].height = 22
        planner.add_table(Table(displayName='PlannerLines', ref=f'A4:H{4+len(LR)}',
                                tableStyleInfo=TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)))
    for j, w in enumerate([22, 36, 14, 14, 34, 36, 44, 12], 1):
        planner.column_dimensions[L(j)].width = w
    planner.freeze_panes = 'A5'

def build_new_mrp(DS, BW, mis_path, out_path, prev_path=None, TD=None, old_mrp=None):
    """MRP_Daily_<date>.xlsx: Start Here, Action List, MRP Daily, MRP by Line, Planner, Today, MRP Weekly, Stock Movement, Check, No Demand, Settings
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
    AR='Calibri'; NAVY='1F3864'; thin=Side(style='thin',color='BFBFBF'); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
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
        estimate_counted_stock(DS['blocks'],mis_path,d0)                  # older physical counts rolled forward to today
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
    ctr=Alignment(horizontal='center',vertical='center',wrap_text=True)
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
        if not b_.get('count_stock') and float(b_.get('st_jw') or 0)>0 and not b_.get('jw_out'):
            b_['stock']=float(b_['stock'])-float(b_['st_jw']); b_['jw_out']=True
            if b_.get('st_main') is not None: pass
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
    A=[]
    for b in DS['blocks']:
        TDm=b.get('td') or TD          # target days of this material (ABC / XYZ class)
        P=b['plan']; T=[a+c for a,c in zip(b['transit'],b['adhoc'])]; O=b.get('openpo') or [0.0]*nd
        stk_=b['stock']
        if b['mat'] in POOLG: g_=POOLG[b['mat']]; P,T,O,stk_=g_['P'],g_['T'],g_['O'],g_['stock']      # substitute group: one projection
        S=[]; cur=stk_
        for k in range(nd): cur=cur-P[k]+T[k]; S.append(cur)
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
        def _ln(c_):                          # one name per line (as in your Production Plan), in line order
            l_=[LINE_DISP.get(k_,k_) for k_,q_ in sorted((clines.get(c_) or {}).items(),key=lambda kv:(LINE_ORD.get(kv[0],99),str(kv[0]))) if q_>0]
            return f" ({', '.join(l_)})" if l_ else ''
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
        for k in range(nd):
            if win[k]>0:
                g_=win[k]-(S[k]+cO[k]+cum)
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
                      doc=(S[0]/(sum(P[:TDm])/len(P[:TDm])) if sum(P[:TDm]) else float(CAP_DOC))))   # MRP Daily 'Days of cover' on day 1 (own target window)
    A.sort(key=lambda x:(STATUS_ORDER[x['status']],x.get('exg',9),x['cov'][1],x['cov'][2],bool(x['b'].get('count_stock'))))   # most urgent first (status, exception, days)
    # ---------- base workbook: Action List (the old Dashboard was removed: headline -> Start Here, top risks -> Start Here / Action List,
    #            lines -> Planner, cover chart -> Today) ----------
    wb=Workbook(); al=wb.active; al.title='Action List'
    al.sheet_properties.tabColor='C00000'; set_print(al,'Action List')
    n=len(A); nok=sum(1 for x in A if x['status']=='OK'); stc={k:(STATUS_FILL[k],STATUS_FONT[k]) for k in STATUS_FILL}
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
        x['tgt_days']=(max(0.0,b['stock']+(float(b.get('st_jw') or 0) if b.get('jw_out') else 0.0)+(x.get('tr_tot') or 0))/x['tgt_need']) if x.get('tgt_need') else None   # with OEM: + JW stock
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
            g_=POOLG.get(c_); jw_=sum(float(byc_[m2]['b'].get('st_jw') or 0) for m2 in (g_['members'] if g_ else [c_]) if m2 in byc_ and byc_[m2]['b'].get('jw_out'))
            st_=((g_['stock']+g_['tr']) if g_ else (x_['b']['stock']+(x_.get('tr_tot') or 0)))+jw_      # with OEM: + JW stock
            x_['tgt_days']=(max(0.0,st_)/x_['tgt_need']) if x_.get('tgt_need') else None          # group stock + in transit ÷ group need
            oth_=[f"+ {o_} {str(byc_[o_]['b'].get('common') or '')}: {old_[o_]['tgt_need']*fo_/f_:,.0f} {x_['b'].get('uom') or ''} per day (substitute group)" for o_,fo_ in mem_ if o_!=c_ and old_[o_]['tgt_need']]
            if oth_: x_['usage_src']='\n'.join([t_ for t_ in [x_.get('usage_src')] if t_]+oth_)
            POOL_.append((c_,own_,x_['tgt_need']))
    if POOL_: print('Substitute groups - daily target pooled: '+'; '.join(f"{c_} {a_:,.0f} -> {b_:,.0f}" for c_,a_,b_ in POOL_),flush=True)
    # ---- NO DEMAND rule: a material is only Critical / At risk / Watch / Short when it has demand. Demand = at least one of
    #      production requirement / BOM demand (plan x usage) inside the demand horizon, planned consumption in the daily-plan window, daily target need.
    #      Otherwise 'No demand': no shortage status, no planned order / release, no Expedite / Order now, not on any risk list.
    DH_=max(1,min(nd,int(DEMAND_HORIZON or SHORT_HORIZON)))
    NODEM_[:]=[]
    for x in A:
        b=x['b']; P_=b.get('plan') or []
        x['demand_h']=sum(P_[:DH_])
        x['no_demand']=not (x['demand_h']>0 or sum(P_[:int(PLAN_AVG_DAYS)])>0 or (x.get('tgt_need') or 0)>0 or (x.get('tgt_need_c') or 0)>0)
        if x['no_demand']:
            b['por']=[0.0]*nd; b['rel']=[0.0]*nd
            x['nxt_po']=dict(qty=0,need=None,release=None,past=0)
            x['status']='NO DEMAND'; x['exc']='No demand - no action'; x['exg']=7; x['action']='No demand - no shortage action'; x['short_txt']=''
            f_=next((k for k,v in enumerate(P_) if v>0),None)
            NODEM_.append((b['mat'],str(b.get('common') or b.get('desc') or ''),f"first plan demand in {f_} d ({sum(P_):,.0f} {b.get('uom') or ''} in the plan)" if f_ is not None else 'no plan demand at all'))
    if NODEM_: print(f"No demand (no shortage status / action): {len(NODEM_)} materials - "+'; '.join(f"{c_} {n_[:18]} ({w_})" for c_,n_,w_ in NODEM_[:12]),flush=True)
    for x in A: x['cw']=cover_watch(x,d0,ships,nd)                       # Below 21 Days From / Short From / Days Late (same as Risk Charts)
    # ---- SCENARIOS for the Action List
    #  with OEM    = all lines (Settings targets / highest planned day for lines without one), stock = MB52 + JW stock at the OEM
    #  without OEM = USA Lines 1-4 only, stock = MB52 only (CS / PM: Stock Count 'Today Qty_BOM' in both)
    NO_OEM_LINES={'L1','L2','L3','L4'}
    def run_days(stock_,plan_):           # need on the first production day, and calendar days until stock is used up by the plan
        k0=next((k for k in range(len(plan_)) if plan_[k]>0),None)
        if k0 is None: return None,None
        left=max(0.0,stock_)
        for k in range(len(plan_)):
            if plan_[k]>left: return plan_[k0],k+(left/plan_[k] if plan_[k] else 0)
            left-=plan_[k]
        return plan_[k0],float(len(plan_))
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
    tg_txt=' + '.join(f"{lab.split(' Production')[0]} {v:,.0f}" for _,(lab,v) in sorted(EFF_TARGET.items()) if v) or 'no daily targets'
    al['A1']='Action List - what to do to keep every material covered for its own target days (Settings)'; al['A1'].font=Font(name=AR,size=14,bold=True,color=NAVY)
    al['A2']=('Highest shortage risk first (ALERT > short later > under 21 d). Short from = first day stock (count rolled forward, or SAP) − plan demand + MIS in transit < 0; ALERT = inside the lead time. '
              'Daily target need = customer daily targets × usage per module (your BOM first, else MIS Standard BOM); without OEM (green banner) = customers with OEM = No, with OEM (blue banner) adds the OEM and JW stock at OEM. '
              'Inv. Days = (stock + MIS in transit) ÷ daily target need, red under 21. '
              f'Below {REPORT_DAYS} Days From / Short From (Stock = 0) / Days Late = the Risk Charts logic. No demand = no production requirement, planned consumption or daily target in the next {int(DEMAND_HORIZON or SHORT_HORIZON)} days: no shortage status and no order action. '
              'Details per column: Definitions sheet.')
    al['A2'].font=Font(name=AR,size=9,italic=True); al.merge_cells('A2:S2'); al['A2'].alignment=Alignment(wrap_text=True,vertical='top'); al.row_dimensions[2].height=54
    wtxt_=lambda x: ('; '.join(f"{(d0+pd.Timedelta(days=w[0])):%d-%b} → "+(f"beyond {(d0+pd.Timedelta(days=w[1])):%d-%b}" if w[2] else f"{(d0+pd.Timedelta(days=w[1])):%d-%b} ({w[1]-w[0]+1} d)") for w in x.get('wins') or []) or 'no stock-out')
    def short_lt(x):                     # short window + the material's lead time; ALERT when the shortage starts inside the lead time
        if x.get('no_demand'): return 'No demand'
        t_=wtxt_(x); lt_=x['b'].get('lt_used'); w_=x.get('wins') or []
        if not w_: return t_
        k_=w_[0][0]
        if lt_ is not None and k_<lt_: return f"ALERT inside lead time: {t_}  |  lead {lt_} d - short in {k_} d"
        return f"{t_}  |  lead {lt_} d - order in time ({k_-lt_} d spare)" if lt_ is not None else t_
    def nx_txt(x):                       # next in-transit arrival day: TOTAL of all shipments arriving that day
        n_=x.get('nx')
        if not n_: return None
        return f"{n_['date']:%d-%b}: {n_['qty']:,.0f} {x['b'].get('uom') or ''} ({n_['n']} shipment{'s' if n_['n']>1 else ''}: {n_['sup'][:40]})".replace('  ',' ')
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
    def sd_(x,what):
        g_=sdg_.get(x['b']['mat']) or {}
        ks=sorted(g_,key=lambda k_:(k_[0]!='Stock',k_[2]!=x['b']['mat'],k_[2]))
        f={'src':lambda k_,e_: ('to OEM: ' if k_[0]=='to OEM' else '')+k_[1].replace(' (Total Stock)','').replace(' (JW WSA to OEM)',''),
           'code':lambda k_,e_: k_[2],'q':lambda k_,e_: f"{e_['qty']:,.0f}",'u':lambda k_,e_: str(e_['uom'] or ''),
           'f':lambda k_,e_: f"{e_['factor']:g}",'qm':lambda k_,e_: f"{e_['qty']*e_['factor']:,.0f}",'mu':lambda k_,e_: str(x['b'].get('uom') or '')}[what]
        return '\n'.join(f(k_,g_[k_]) for k_ in ks) or None
    def sloc_(x):
        tot={}
        for e_ in (sdg_.get(x['b']['mat']) or {}).values():
            for s__,q__ in e_['sl'].items(): tot[s__]=tot.get(s__,0.0)+q__
        return '\n'.join(f"{s__}: {q__:,.0f}" for s__,q__ in sorted(tot.items(),key=lambda kv:-kv[1])) or None
    def _cwv(x,k):                       # Risk Charts logic (cover_watch): date / days / '-' / 'No demand'
        if x.get('no_demand'): return 'No demand'
        cw=x.get('cw')
        if x.get('sup_row') or cw is None: return None
        if k=='below': return cw['below'].to_pydatetime() if cw['below'] is not None else '-'
        if k=='short': return short_lt_text(x,d0)[0]                        # all stock-zero periods inside the lead time
        return cw['late'] if cw['late'] is not None else '-'
    COLS=[('Component',lambda x:x['b']['mat'],None,16,False),
          ('Source',lambda x:sd_(x,'src'),None,22,True),('Source code',lambda x:sd_(x,'code'),None,16,True),
          ('Qty (source unit)',lambda x:sd_(x,'q'),None,11,True),('Source UOM',lambda x:sd_(x,'u'),None,7,True),('Factor',lambda x:sd_(x,'f'),None,6,True),
          ('Qty (MRP unit)',lambda x:sd_(x,'qm'),None,11,True),('MRP UOM',lambda x:sd_(x,'mu'),None,7,True),
          ('Descr. of Storage Loc.',sloc_,None,24,True),
          ('Component description',_desc,None,34,True),
          ('UOM',lambda x:x['b']['uom'],None,6,False),
          ('Lines (share of plan volume)',lambda x:x['custs'],None,16,True),
          ('Project Name (share of plan)',lambda x:x.get('proj'),None,30,True),
          ('Usage per module by customer (source: your BOM / MIS Standard BOM)',lambda x:x.get('usage_src'),None,46,True),
          ('Module (MIS BOM, share of plan)',lambda x:x.get('mods'),None,16,True),
          ('Short from → to (with in transit)',short_lt,None,40,True),
          (f'Below {REPORT_DAYS} Days From',lambda x:_cwv(x,'below'),'dd-mmm-yy',12,False),
          ('Short From (Stock = 0)',lambda x:_cwv(x,'short'),None,30,True),
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
          ('In transit (total qty, WSA RM Planning)',lambda x:round(x['wsa_it']) if x.get('wsa_it') else None,'#,##0',12,False),
          ('Open PO qty (WSA RM Planning)',lambda x:round(x.get('op_tot') or 0),'#,##0;-#,##0;"-"',12,False),
          ('Alternative BOM (yours): code - stock - customers',alt_txt,None,46,True)]
    for j,(h,_,_,w,_) in enumerate(COLS,1):
        c=al.cell(4,j,h); c.font=Font(name=AR,size=9,bold=True,color='FFFFFF'); c.alignment=ctr; c.border=bd
        c.fill=PatternFill('solid',fgColor=al_hdr_color(h)); c.font=Font(name=AR,size=9,bold=True,color=('1F1F1F' if al_hdr_light(h) else 'FFFFFF'))
        al.column_dimensions[L(j)].width=w
    al.row_dimensions[4].height=42
    grp_=[al_group(h_) for h_,*_ in COLS]; j0=0
    while j0<len(grp_):                       # row 3: one coloured banner per scenario group
        j1=j0
        while j1+1<len(grp_) and grp_[j1+1]==grp_[j0]: j1+=1
        lab_,col_=AL_BANNERS[grp_[j0]]
        c=al.cell(3,j0+1,lab_); c.font=Font(name=AR,size=10,bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor=col_); c.alignment=Alignment(horizontal='center',vertical='center')
        for jj in range(j0+2,j1+2): al.cell(3,jj).fill=PatternFill('solid',fgColor=col_)
        if j1>j0: al.merge_cells(start_row=3,start_column=j0+1,end_row=3,end_column=j1+1)
        j0=j1+1
    al.row_dimensions[3].height=20
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
    def risk_key(x):                       # highest shortage risk first
        w_=x.get('wins') or []; lt_=x['b'].get('lt_used') or 0
        k_=w_[0][0] if w_ else None                          # day the shortage (with in transit) starts
        td_=x.get('tgt_days')
        if k_ is not None and k_<lt_: grp_=0                 # 1) short INSIDE the lead time (ALERT) - earliest first
        elif k_ is not None: grp_=1                          # 2) short later, still time to order - earliest first
        elif td_ is not None: grp_=2                         # 3) no shortage in the plan - lowest days at the daily target first
        else: grp_=3                                         # 4) no need at the daily targets
        return (grp_,k_ if k_ is not None else 0,td_ if td_ is not None else 1e9,x['b']['common'])
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
            cnt_=(COUNTED_RM.get(b['mat']) or {}).get('by_sup') or {}
            hit_=[qq_[1] for sc_,qq_ in cnt_.items() if SUP_CODE.get(sc_,sc_) in v_.upper() or v_.upper().startswith(sc_)]
            if hit_: st_=sum(hit_)                # this supplier was counted: its physical count (Stock Count 'Supplier Code'); else its MB52 batches
            nw_=need_.get(v_,0.0); nwo_=need0_.get(v_,0.0)
            px=dict(x); px['b']=dict(b,common=f"{b['common']} - {v_}",desc=f"{b.get('desc') or ''} | supplier {v_}"); px['sup_row']=v_
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
            c=al.cell(rr,j,v_); c.font=Font(name=AR,size=9); c.border=bd; c.alignment=Alignment(vertical='top',wrap_text=wrap)
            if nf: c.number_format=nf
            if h.startswith('Short from') and c.value not in ('no stock-out','No demand'): c.font=Font(name=AR,size=9,bold=True,color='9C0006')
            if h in ('Short From (Stock = 0)',) and isinstance(c.value,str) and '→' in c.value: c.font=Font(name=AR,size=9,bold=True,color='9C0006')     # stock zero inside the lead time
            elif h in ('Short From (Stock = 0)',) and isinstance(c.value,str) and 'after lead time' in c.value: c.font=Font(name=AR,size=9,bold=True,color='B9770E')
            if h=='Days Late' and isinstance(c.value,(int,float)) and c.value>0: c.font=Font(name=AR,size=9,bold=True,color='B9770E')
            if x.get('no_demand') and h in ('Component description','Short from → to (with in transit)'): c.font=Font(name=AR,size=9,italic=True,color='7F7F7F')
            if h.startswith('Inv. Days') and isinstance(c.value,(int,float)) and c.value<REPORT_DAYS:     # under 21 days: red highlight
                c.fill=PatternFill('solid',fgColor='FFC7CE'); c.font=Font(name=AR,size=9,bold=True,color='9C0006')
            if h.startswith('Short from') and str(c.value or '').startswith('ALERT'): c.fill=PatternFill('solid',fgColor='FFC7CE')
    TG_,PL_=calc_layout(A,RM_); CALC_LAYOUT[:]=[TG_,PL_]
    _hs=[h_ for h_,*_ in COLS]; _j0=_hs.index('Stock'); _j1=max(i_ for i_,h_ in enumerate(_hs) if 'OEM' in h_ and ('arget' in h_ or 'lan' in h_))
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
        al[f'{K_}{rr}']=f'=IF(N({J_}{rr})>0,MAX(0,{F_}{rr}+{G_}{rr}+{H_}{rr})/{J_}{rr},"")'
        al[f'{O_}{rr}']=f'=IF(N({N_}{rr})>0,MAX(0,{F_}{rr}+{H_}{rr})/{N_}{rr},"")'
        al[f"{cl_('Inv. Days at line daily target_without OEM (stock only, no transit)')}{rr}"]=f'=IF(N({N_}{rr})>0,MAX(0,{F_}{rr})/{N_}{rr},"")'
        al[f"{cl_('Inv. Days at line daily plan_without OEM (stock only, no transit)')}{rr}"]=f'=IF(N({P_}{rr})>0,MAX(0,{F_}{rr})/{P_}{rr},"")'
        if rm:
            c0,ndx=rm['C0'],rm['ND']; a_,z_=L(c0),L(c0+ndx-1)
            al[f'{H_}{rr}']=f"=SUM('MRP Daily'!${a_}${rm['rt']}:${z_}${rm['rt']})"
            al[f"{cl_('Open PO qty (WSA RM Planning)')}{rr}"]=f"=SUM('MRP Daily'!${a_}${rm['ro']}:${z_}${rm['ro']})"
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
               str(AL_VAL.get((x['b']['mat'],'Source code')) or '').count('\n'),str(x.get('proj') or '').count('\n'),str(x.get('usage_src') or '').count('\n'),str(AL_VAL.get((x['b']['mat'],'Descr. of Storage Loc.')) or '').count('\n'))+1
        for j__ in (jat_,jcu_,jcu_+1): al.cell(rr,j__).alignment=Alignment(wrap_text=True,vertical='top')
        if n_>1: al.row_dimensions[rr].height=max(15,12.5*n_)
    for rr in (range(5,5+len(AL_ORDER)) if not VALUES_ONLY else []):     # stock = the lines on Stock Data (double-check there)
        al.cell(rr,jst_,f"=SUMIFS('Stock Data'!$L$5:$L$3000,'Stock Data'!$A$5:$A$3000,$A{rr},'Stock Data'!$C$5:$C$3000,\"Stock\")")
        al.cell(rr,joe_,f"=SUMIFS('Stock Data'!$L$5:$L$3000,'Stock Data'!$A$5:$A$3000,$A{rr},'Stock Data'!$C$5:$C$3000,\"to OEM\")")
    al.add_table(Table(displayName='ActionList',ref=f'A4:{L(len(COLS))}{4+len(AL_ORDER)}',tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=True)))
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
    nmove=0                                                        # Stock Movement sheet removed (your request); the Movement Log is still kept
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
        iss+=[dict(area='In transit',check='MIS Final RM Code does not match the shipment - shipped code used',sev='High' if (f['qty'] or 0)>0 else 'Low',
                   material=f['desc'][:60],code=f"{f['shipped']} (MIS: {f['mis']})",a_lab='Pending qty',a=round(f['qty'] or 0),b_lab='',b=None,diff=None,
                   detail=f"Shipment '{f['desc'][:60]}' matches {f['shipped']} ({f['cd_s'][:40]}) not {f['mis']} ({f['cd_m'][:40]}); PO {f['po']}, {str(f['sup'] or '')[:25]}",
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
            ('- Your MRP.xlsx (read only): Settings, Production Plan (columns found by header), Material Group. BOM and Stock Count come from the PPC workbook (sheets BOM, Count) when it is in the folder.',False),
            ('- RM Issues: one issue per raw material at risk (required date, qty, supplier / PO / ETA, delay, MW impact); your Owner / Status / Date Closed / Action are kept run to run.',False),
            ('- Customer_Tracker_*.xlsx: no longer used (SCM Dashboard removed).',False),
            ('',False),('Stock',True),
            ('- Counted materials: the physical count is used; a count older than today is rolled forward: count - usage per day x days since count + received since count (MIS PPC GRN). Usage per day = Production Plan demand, next 7 days.',False),
            ('- Not counted: SAP MB52 Total Stock (+ alternate codes / roll lengths converted); when the MIS file is older than today (its save date) this SAP stock is rolled forward the same way (− usage per day × days since the MIS date) and the console prints a WARNING. JW stock at OEM is added in the with-OEM figures.',False),
            ('- OEM (customers with OEM = Yes): their plan and the JW stock at the OEM are left out of every calculation; only the Action List shows the with-OEM target and Inv. days.',False),
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
            ('- NO DEMAND: a material is only Critical / At risk / Watch / Short when it has demand (production requirement, planned consumption, daily target or BOM demand) inside the shortage horizon (Settings, 60 days). Without it: status NO DEMAND, no planned order / release, no Expedite / Order now, not on Risk Charts or Stock Management.',False),
            ('- Every list is sorted by this risk (highest first); a short material is never labelled safe; the script checks itself after every run (console: Self-check).',False),
            ('',False),('Lines',True),
            ('- Each material is shared out to the lines day by day (line ranking, allowed vendors, vendor priority, Settings substitutes) and then your approved ALTERNATIVE (your BOM) is used before a line stops; free stock of an alternative = what its own plan does not need within its lead time.',False),
            ('',False),('Sheets (visible)',True),
            ('Stock Management: raw materials AT RISK only (ALERT / short later / under 21 d), lines that stop, shipments and suppliers of short materials. Materials with NO DEMAND are never listed as a risk.',False),
            ('Risk Charts: list of raw materials dropping below 21 days in the next 16 weeks and why (Below 21 Days From, Short From, Days Late, Arrival at Plant (Final AtA), supplier, qty, action); Cover Heatmap - weekly projected cover in days around the 21-day line.',False),
            ('Stock Check (SAP vs count) · Action List · MRP Daily · MRP by Line · MRP Weekly.',False),
            ('Hidden (right of In Transit Status): In Transit Status, Substitute Advice, Stock Data, Check, No Demand, Settings, Movement Log, Module Plan, BOM, Stock Count, How it works, MIS Layout, Definitions.',False),
            ('',False),('Daily routine',True),
            ('1 Run the script and read the console Self-check line. 2 Stock Management + Risk Charts: ALERT / short ≤ 7 days / lines stopping. 3 Action List: Below 21 Days From, Short From, Days Late. 4 Stock Check: counts older than today and big SAP differences.',False)]
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
    # ---------- Start Here: data trust check + the daily workflow, then tabs in workflow order ----------
    files=[('WSA planning file',globals().get('SRC'),globals().get('STOCK_DATE')),('MIS RM Dashboard',mis_path,None),('Your MRP.xlsx',old_mrp,None),('Settings from',SETTINGS_SRC[0],None)]
    subj,body=material_report(A,d0,os.path.basename(out_path))          # ONE daily report (the 3 old e-mails are gone)
    try:
        pass                                                       # no _report.txt any more (your request)
        for old_ in ('_email.txt','_email_RM.txt','_email_CSPM.txt'):          # remove today's old-style e-mail files if a previous version left them
            if os.path.exists(out_path.replace('.xlsx',old_)): os.remove(out_path.replace('.xlsx',old_))
    except OSError: pass
    # Start Here sheet removed (your request)
    reorder_sheets(tmp,['Start Here','Action List','MRP Daily','MRP by Line','Planner','Today','MRP Weekly','In Transit Status','Substitute Advice','Stock Movement','Stock Data','Check','No Demand','Settings'])
    unfreeze_all(tmp)                                              # no frozen panes on any sheet of MRP_Daily
    # all sheets visible (your choice); the main sheets stay first in the tab order                                # only the Action List is visible; every other sheet is hidden (formulas still use them)
    os.replace(tmp,out_path)
    SHIPS_.clear(); SHIPS_.update(ships)
    try: add_risk_charts(out_path,A,AL_ORDER,d0,nd,ships)                 # treemap + heatmap of the shortage risk (cells, no chart objects)
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: Risk Charts not added ({type(ex_).__name__}: {ex_})')
    try: add_rm_issues_sheet(out_path,A,AL_ORDER,d0,nd,ships,prev_path)      # RM Issues log (your entries carried over)
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: RM Issues not added ({type(ex_).__name__}: {ex_})')
    try: add_stock_check_sheet(out_path,AL_ORDER,d0)               # SAP stock vs physical count vs estimate today (after Inventory)
    except Exception as ex_: print(f'NOTE: Stock Check sheet not added ({type(ex_).__name__}: {ex_})')
    try: add_stock_mgmt_dashboard(out_path,A,AL_ORDER,mlog,ships,d0,nd,MIS_PATH_[0])     # Stock Management System dashboard (RM)
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: Stock Management dashboard not added ({type(ex_).__name__}: {ex_})')
    try: unify_workbook_style(out_path)                          # one style for all sheets, no charts, Definitions sheet
    except Exception as ex_: print(f'NOTE: style / definitions not applied ({type(ex_).__name__}: {ex_})')
    try:
        wb_=post_wb(out_path)
        for w_ in wb_.worksheets: w_.sheet_view.tabSelected=False
        wb_.active=0; wb_.worksheets[0].sheet_view.tabSelected=True
    except Exception as ex_: print(f'NOTE: active sheet not set ({type(ex_).__name__}: {ex_})')
    try: post_flush(out_path)                                      # ONE save for all finishing steps
    except Exception as ex_: print(f'NOTE: finishing save failed ({type(ex_).__name__}: {ex_})')
    audit_workbook(out_path)                                       # self-check: risk order + no short material labelled safe
    return dict(n=n,crit=sum(1 for x in A if x['status']=='CRITICAL'),risk=sum(1 for x in A if x['status']=='AT RISK'),
                watch=sum(1 for x in A if x['status']=='WATCH'),ok=sum(1 for x in A if x['status']=='OK'),nodem=sum(1 for x in A if x['status']=='NO DEMAND'),check=cnt,cs=len(cs),bom_rows=len(bom),bom_src=bom_src)

# ===================== Material Plan sheet: one column per day =====================
if __name__=='__main__':
    t0=time.time()
    # ---------- Settings FIRST: every rule (targets, line thresholds, lead limits ...) is fixed before anything is calculated,
    #            so the Inbound file and MRP_Daily use the same values. Your MRP.xlsx first, else the last MRP_Daily. ----------
    MRP_DAILY_OUT=os.path.join(OUTDIR,f'MRP_Daily_{dt.date.today():%Y-%m-%d}.xlsx')
    _prevs=sorted([os.path.join(OUTDIR,n_) for n_ in os.listdir(OUTDIR) if n_.startswith('MRP_Daily_') and n_.endswith('.xlsx') and not n_.startswith('~$')],key=os.path.getmtime)
    PREV_MRP=MRP_DAILY_OUT if os.path.exists(MRP_DAILY_OUT) else (_prevs[-1] if _prevs else None)
    SETTINGS_SRC[0]=read_settings([USER_MRP,PREV_MRP])
    print(f'Settings read from {os.path.basename(SETTINGS_SRC[0])}' if SETTINGS_SRC[0] else 'NOTE: no Settings sheet found (MRP.xlsx or last MRP_Daily) - script defaults used',
          f'- targets A {TARGET_A} / B {TARGET_B} / C {TARGET_C} / packaging {TARGET_CSPM} / other {TARGET_DAYS} days; lines {COVER_CRITICAL} / {COVER_RISK} / {COVER_WATCH} days')
    print('Reading the plan (materials and codes)...',flush=True)
    try: d=pd.read_excel(SRC,sheet_name='Total_Plan_Req',usecols=lambda c: str(c).strip() in ('Date','Common Name.Common Name','Standard BOM.Material Code','Standard BOM.UOM'))
    except ValueError: die("Sheet 'Total_Plan_Req' not found")
    d.columns=[str(c).strip() for c in d.columns]
    d['Date']=pd.to_datetime(d['Date'],errors='coerce'); d=d.dropna(subset=['Date','Common Name.Common Name'])
    plan_end=d['Date'].max().to_pydatetime()
    codes_by_name=d.groupby('Common Name.Common Name')['Standard BOM.Material Code'].first().astype(str).to_dict()
    uom_name=d.groupby('Common Name.Common Name')['Standard BOM.UOM'].first().to_dict()
    print('Reading the planning sheets...',flush=True)
    G,skipped=extract_fast(SRC,plan_end)
    if not G: die('No planning sheets found (visible sheets with a "Required Plant Date" header)')
    for s in skipped: print('WARNING: planning sheet skipped -',s)
    print(f'Planning sheets read: {len(G)} blocks; with no scheduled deliveries: '+(', '.join(g['label'] for g in G if not g['inbound']) or 'none'))
    name2grp={}
    for g in G:
        for n,c in codes_by_name.items():
            if c in g['codes']: name2grp.setdefault(n,g['label'])
    ginfo={g['label']:g for g in G}
    STOCK_DATE=Counter(g['asof'] for g in G).most_common(1)[0][0]
    # ---- one record per delivery (before splitting between materials) ----
    deliveries=[]
    for g in G:
        nn=sorted(n for n in name2grp if name2grp[n]==g['label'])
        if not nn: nn=[g['label']]
        for x in g['inbound']:
            deliveries.append(dict(g=g['label'],sheet=g['sheet'],mats=nn,x=x,asof=g['asof']))
    leads=[(e['x']['date']-e['x']['etd']).days for e in deliveries if hasattr(e['x']['etd'],'year')]
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
        e['chase']=0 if e['cancelled'] else int(e['stage']=='Past due' or bool(e['missing']) or e['etdpassed'] or e['datecheck']=='Short lead')
    rows=[]
    for k,e in enumerate(deliveries):
        for n in e['mats']: rows.append((n,e['g'],e,1.0/len(e['mats']),n==e['mats'][0]))
    empties=[g for g in G if not g['inbound']]
    for g in empties:
        nn=sorted(n for n in name2grp if name2grp[n]==g['label']) or [g['label']]
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
           'Status (clean)','Port','Lead days (plant date - ETD)','Date check','ETD passed, no vessel/BL','Missing info','Days to arrival','Source row','Unique delivery (1 = first row)','Chase?','BOM code','Sheet: Balance','Sheet: Planned modules per day','Sheet: Demand','Sheet: Actual modules per day','Sheet: Actual','Sheet: MW per day','Sheet: Actual MW per day','Sheet: Plant Days','Sheet: Plant + Port Days (ETA +5 days)','Sheet: Delivered','Sheet: Scheduled Delivery']
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
        ws.cell(r,30,f'=IF(O{r}="Cancelled",0,IF(OR(O{r}="Past due",Y{r}<>"",Z{r}<>"",X{r}="Short lead"),1,0))').font=bold
        ws.cell(r,31,codes_by_name.get(n,'')).font=blue
        write_info(r,x.get('info'),x['qty'])
    R1=R0+len(rows)-1
    rng=lambda col: f'Inbound!${col}${R0}:${col}${R1}'
    if rows:
        ws.auto_filter.ref=f'A{HR}:AP{R1}'
        ws.conditional_formatting.add(f'A{R0}:AE{R1}',FormulaRule(formula=[f'$O{R0}="Past due"'],fill=amb))
        ws.conditional_formatting.add(f'A{R0}:AE{R1}',FormulaRule(formula=[f'$O{R0}="No scheduled delivery"'],fill=PatternFill('solid',bgColor='DDEBF7',fgColor='DDEBF7')))
        ws.conditional_formatting.add(f'A{R0}:AE{R1}',FormulaRule(formula=[f'AND($Q{R0}=0,$O{R0}<>"No scheduled delivery")'],font=Font(name=AR,color='999999',strike=True)))
        ws.conditional_formatting.add(f'X{R0}:Z{R1}',FormulaRule(formula=[f'AND(X{R0}<>"",X{R0}<>"OK")'],fill=redf))
    for c_,w_ in zip(['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE'],[28,8,13,14,13,9,22,12,14,14,14,13,16,30,12,12,9,13,26,13,20,14,12,12,24,18,10,9,12,8,16]): ws.column_dimensions[c_].width=w_
    for c_ in ['AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP']: ws.column_dimensions[c_].width=15
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
    hrow(sm,MR,['Material','UOM','Counted deliveries','Counted quantity','In transit qty','Open PO qty','Next arrival','Next arrival qty','Past due qty (not counted)','Cancelled qty','Deliveries to chase','Sheet stock date','Last actual receipt (info)','Last receipt qty (info)','Sheet balance at stock date','Sheet Plant Days at stock date','Sheet Plant + Port Days at stock date']); sm.row_dimensions[MR].height=32
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
        gl=ginfo.get(name2grp.get(n,n)); lr=gl.get('last_receipt') if gl else None
        if lr:
            c=f(13); c.value=lr[0]; c.font=blue; c.number_format=dfy
            c=f(14); c.value=lr[1]; c.font=blue; c.number_format=num
        ai=(gl.get('asof_info') or {}) if gl else {}
        for k_,key_,fm_ in ((15,'balance',num),(16,'plantdays','0.0'),(17,'plantport','0.0')):
            if ai.get(key_) is not None:
                c=f(k_); c.value=ai[key_]; c.font=blue; c.number_format=fm_
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
    for c_,w_ in zip('ABCDEFGHIJKLMNOPQ',[44,16,16,16,14,14,14,14,16,14,14,14,16,16,16,16,18]): sm.column_dimensions[c_].width=w_
    sm.cell(SR+len(sheets)+2,1,'A sheet with 0 deliveries was read but has nothing in its Scheduled Delivery column (for example only actual receipts in the Delivered column).').font=bf
    sm.freeze_panes='A4'
    # ================= To chase =================
    tc['A1']='To chase - deliveries that need follow-up (snapshot at the time of the run)'; tc['A1'].font=ttl
    tc['A2']='Reasons: past due = dated before the sheet stock date and not cancelled (not counted); missing PO / shipper / ETD; ETD passed with no vessel or BL; short lead = plant date under the short-lead limit after ETD.'; tc['A2'].font=bf
    hrow(tc,4,['Planning sheet','Sheet stock date','Materials','Plant date','Days to arrival (vs report date)','Sheet quantity','Stage','Remarks (as written)','PO','Shipper','ETD','BL#','Vessel / ETA note','Reasons']); tc.row_dimensions[4].height=32
    chase=[e for e in deliveries if e['chase']]
    for i,e in enumerate(sorted(chase,key=lambda e:e['x']['date'])):
        x=e['x']; why=[]
        if e['stage']=='Past due': why.append(f"Past due ({(pd.Timestamp(e['asof'])-pd.Timestamp(x['date'])).days} d before stock date)")
        if e['missing']: why.append('Missing '+e['missing'].replace(' ',', '))
        if e['etdpassed']: why.append('ETD passed, no vessel/BL')
        if e['datecheck']=='Short lead': why.append(f"Short lead ({e['lead']} d after ETD)")
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
            gl_=ginfo.get(name2grp.get(n,n)); bal_=None
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
            gl_=ginfo.get(name2grp.get(n,n)); b1=b2=None
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
                         members=sorted(n for n in name2grp if name2grp[n]==lab) or [lab],sheet=g_['sheet']) for lab,g_ in ginfo.items()}
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
            print(f"NEW MRP saved: {outn} - {rn['crit']} critical, {rn['risk']} at risk, {rn['watch']} watch, {rn['ok']} OK, {rn['nodem']} no demand; "
                  f"Check: {rn['check'].get('High',0)} high issues; CS / PM: {rn['cs']} components from {rn['bom_rows']} BOM rows"
                  +(f" (BOM taken from {os.path.basename(PPC_SRC[0]) if PPC_SRC[0] else os.path.basename(rn['bom_src'])})" if rn['bom_src'] else ' - no BOM yet: fill the BOM sheet'))
            # the new MRP_Daily is saved: delete the older MRP_Daily files (and their _report.txt) in this folder
            gone_=[]; kept_=[]
            for n_ in sorted(os.listdir(os.path.dirname(os.path.abspath(outn)))):
                fp_=os.path.join(os.path.dirname(os.path.abspath(outn)),n_)
                if os.path.abspath(fp_)==os.path.abspath(outn) or n_.startswith('~$'): continue      # old _report.txt files go too
                if re.fullmatch(r'MRP_Daily_\d{4}-\d{2}-\d{2}(_report\.txt|\.xlsx)',n_):
                    try: os.remove(fp_); gone_.append(n_)
                    except OSError: kept_.append(n_)
            if gone_: print('Old MRP_Daily deleted: '+', '.join(gone_))
            if kept_: print('NOTE: could not delete (open in Excel?): '+', '.join(kept_))
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
    print(f'To chase: {len(chase)} deliveries (past due {npast}, missing PO/shipper/ETD {sum(1 for e in chase if e["missing"])}, ETD passed with no vessel/BL {sum(1 for e in chase if e["etdpassed"])}, short lead {sum(1 for e in chase if e["datecheck"]=="Short lead")})')
    # (the old By Week / MRP_<date> update was retired: your MRP.xlsx is read only, MRP_Daily is the output)
    if prev_del:
        cnt=Counter(t[0] for t in out); print(f'Changes vs {os.path.basename(PREV)}: '+', '.join(f'{v} {k.lower()}' for k,v in cnt.items()) if cnt else f'Changes vs {os.path.basename(PREV)}: none')
