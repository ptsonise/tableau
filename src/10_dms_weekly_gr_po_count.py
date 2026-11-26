BOMD_HEAD='Component Description from BOM'
BOM_DESC={}      # material code -> Description of the PPC BOM (filled in make_daily_material_status; used by the web page)
DMS_NAME2='Item name'      # second Item name column, right before 'Stock zero from -> to' (user, 23-Nov-2026)
SUB_HEAD='Covered by substitute'      # Settings SUBSTITUTES (sub_cover, user 25-Nov-2026); only when the run has a substitute
DMS_WEEKLY_KEEP=('Item name',BOMD_HEAD,'Item Category','UOM','Estimated stock today','SAP stock (MB52)','Physical count','Days of cover (stock only)',SUB_HEAD,
                 'Usage/Day (module plan)','First stock-zero day',DMS_NAME2,'Stock zero from \u2192 to (\u26a0 = starts inside the transit lead time)',
                 'Arrival at plant (Final AtA)',PO_HEAD,PR_HEAD,PGRN_HEAD,POD_HEAD,STS_HEAD,GRR_HEAD,GRL_HEAD)
DMS_ORDER=[]     # (code, item name) in the row order of the DMS BOM Weekly sheet (web page list order, user 25-Nov-2026)
DMS_INFO={}      # (code, item name) -> [(header, text)] of every info column of the DMS BOM Weekly sheet (filled by _dms_info_capture; read by the web page)
def _dms_info_capture(w, ck, hr=2):
    """User 25-Nov-2026: "all info I want to see in BOM Material Dashboard must be seen in the web page always": after `_bom_weekly_keep`, every
       info column of the final DMS 'BOM Material Dashboard - Weekly' row (DMS_WEEKLY_KEEP, as written - so a column added there shows up on
       the web page without another change) is stored per (code, item name) for the 5_Web Details. Empty cell = '-'.
       The week columns are the Weeks tab of the web page (same `bom_day_cover`). Presentation only."""
    DMS_INFO.clear(); DMS_ORDER.clear()
    hd=[(j,str(w.cell(hr,j).value or '').strip()) for j in range(1,w.max_column+1)]
    jn=next((j for j,h in hd if h=='Item name'),None); seen=set(); cols=[]
    for j,h in hd:
        if h in DMS_WEEKLY_KEEP and h not in seen and h!='Item name': seen.add(h); cols.append((j,h))
    def txt(v,h):
        if v is None or (isinstance(v,str) and not v.strip()): return '-'
        if isinstance(v,(dt.datetime,dt.date)): return f"{v:%d-%b-%y}"
        if isinstance(v,bool): return str(v)
        if isinstance(v,(int,float)): return f"{v:,.1f}" if (h.startswith('Days of cover') or h.startswith('Usage/Day')) else f"{v:,.0f}"
        return ' '.join(str(v).split())
    from openpyxl.comments import Comment
    wkc=[j for j,h in hd if re.fullmatch(r'\d{1,2}-[A-Za-z]{3}',h)]
    for r_,c_ in ck.items():
        nm=str(w.cell(r_,jn).value or '').strip() if jn else ''
        DMS_INFO[(c_.upper(),nm)]=[(h,txt(w.cell(r_,j).value,h)) for j,h in cols]; DMS_ORDER.append((c_.upper(),nm))
        for i_,j in enumerate(wkc):                                   # week cell that includes a substitute's cover: note with the own-stock number (user 25-Nov-2026)
            t_=HEAT_SUBNOTE.get((c_.upper(),nm,i_))
            if t_: cm=Comment(t_,'MRP'); cm.width=260; cm.height=60; w.cell(r_,j).comment=cm
def _bom_weekly_keep(w, hr=2):
    """4_Status 'BOM Material Dashboard - Weekly' (user, 18-Nov-2026: "keep these headers only"; layout = the user's own
       edited file of 23-Nov-2026 'format output same as it'): Item name, Component Description from BOM, Item Category, UOM, Estimated
       stock today, SAP stock (MB52), Physical count, Days of cover, Usage/Day, First stock-zero day, Item name (repeated), Stock zero
       from -> to, Arrival at plant, Pending GRN Qty, SAP GR since count, Last SAP GR, then the week columns (header = a date text like
       '18-Nov'); every other column goes. DMS only."""
    import re as _re
    from copy import copy as _cp2
    hd={j:str(w.cell(hr,j).value or '').strip() for j in range(1,w.max_column+1)}
    wk=lambda h: bool(_re.fullmatch(r'\d{1,2}-[A-Za-z]{3}',h))
    jc_=next((j for j,h in hd.items() if h=='Code'),None); jn_=next((j for j,h in hd.items() if h=='Item name'),None)
    if jc_ and jn_ and BOM_DESC:                                   # Component Description from BOM (PPC BOM 'Description') right after Item name (user, 19-Nov-2026)
        _ins_cols(w,jn_+1,1); w.column_dimensions[L(jn_+1)].width=40
        for r_ in range(1,w.max_row+1):
            src_=w.cell(r_,jn_); c_=w.cell(r_,jn_+1); c_._style=_cp2(src_._style)
            if r_==hr: c_.value=BOMD_HEAD
            elif r_>hr:
                ds_=[]
                for cd_ in re.split(r'\s*(?:/|,|\+|\n)\s*',str(w.cell(r_,jc_).value or '')):
                    d_=BOM_DESC.get(cd_.strip().upper())
                    if d_ and d_ not in ds_: ds_.append(d_)
                if ds_: c_.value=' / '.join(ds_); c_.alignment=Alignment(horizontal='left',vertical='center',wrap_text=False)
        hd={j:str(w.cell(hr,j).value or '').strip() for j in range(1,w.max_column+1)}
    sc_={str(r_['code']).strip().upper():r_['subcov'] for r_ in ALERT_ROWS if r_.get('subcov') and not r_.get('where')}
    jd_=next((j for j,h in hd.items() if h=='Days of cover (stock only)'),None)
    if jc_ and jd_ and sc_:                                        # Settings SUBSTITUTES: which code covers when, how much (sub_cover), after Days of cover
        _ins_cols(w,jd_+1,1)
        for r_ in range(1,w.max_row+1):
            c_=w.cell(r_,jd_+1); c_._style=_cp2(w.cell(r_,jn_ if jn_ and jn_<=jd_ else jd_)._style)
            if r_==hr: c_.value=SUB_HEAD
            elif r_>hr:
                v_=sc_.get(str(w.cell(r_,jc_).value or '').strip().upper())
                if v_: c_.value=v_; c_.alignment=Alignment(horizontal='left',vertical='center',wrap_text=False)
        hd={j:str(w.cell(hr,j).value or '').strip() for j in range(1,w.max_column+1)}
    miss=[h for h in DMS_WEEKLY_KEEP if h not in hd.values() and h not in (DMS_NAME2,SUB_HEAD)]
    if miss: print(f"NOTE: {HEAT_NAME} (4_Status): column(s) not found, left out: {', '.join(miss)}",flush=True)
    _del_cols(w,[j for j,h in hd.items() if h and h not in DMS_WEEKLY_KEEP and not wk(h)])
    hd={j:str(w.cell(hr,j).value or '').strip() for j in range(1,w.max_column+1)}
    jn_=next((j for j,h in hd.items() if h=='Item name'),None); jz_=next((j for j,h in hd.items() if h.startswith('Stock zero from')),None)
    if jn_ and jz_:                                                # repeat the item name next to the stock-zero text
        _ins_cols(w,jz_,1); w.column_dimensions[L(jz_)].width=w.column_dimensions[L(jn_)].width
        for r_ in range(1,w.max_row+1):
            c_=w.cell(r_,jz_); c_._style=_cp2(w.cell(r_,jn_)._style)
            c_.value=w.cell(r_,jn_).value                      # the repeat also on the group band rows (user's file, 23-Nov-2026)
    for j in range(1,w.max_column+3): w.column_dimensions[L(j)].outline_level=0
    w.column_dimensions.max_outline=0
    _dms_weekly_header(w,hr)
    return len(DMS_WEEKLY_KEEP)-len(miss)-(0 if sc_ else 1)
def _dms_weekly_header(w, hr=2):
    """DMS BOM Weekly header row (user, 18-Nov-2026 screenshot): every header centered (both ways) and wrapped, row tall enough for
       3 lines, filter buttons on the header row. Only the header row wraps; data rows stay one line (_tidy_dashboard)."""
    from copy import copy as _cp
    last=max([c.column for c in w[hr] if c.value is not None]+[1])
    for c in range(1,last+1):
        al=_cp(w.cell(hr,c).alignment); al.horizontal='center'; al.vertical='center'; al.wrap_text=True; w.cell(hr,c).alignment=al
    w.row_dimensions[hr].height=57.6                                  # user's own edited file, 23-Nov-2026 ('format output same as it'): header Calibri 11 bold black, 57.6 high, zoom 90
    for c in range(1,last+1):
        f_=_cp(w.cell(hr,c).font); f_.sz=11; f_.b=True; f_.color=None; w.cell(hr,c).font=f_
    for r_ in range(w.max_row,hr,-1):                                 # the 'Source: ...' note under the table is not in the user's file
        if str(w.cell(r_,1).value or '').startswith('Source:'): w.delete_rows(r_)
    import re as _re3
    for c in range(1,last+1):
        if _re3.fullmatch(r'\d{1,2}-[A-Za-z]{3}',str(w.cell(hr,c).value or '')): w.column_dimensions[L(c)].width=13
    wid_={'Item name':39,BOMD_HEAD:40,SUB_HEAD:60,'Item Category':18,'UOM':5.6,'Estimated stock today':11.6,'SAP stock (MB52)':13,'Physical count':13,'Days of cover (stock only)':13,
          'Usage/Day (module plan)':13,'First stock-zero day':13,PGRN_HEAD:11.6,POD_HEAD:14,STS_HEAD:40,GRR_HEAD:13,GRL_HEAD:13}      # widths of the user's own file
    for c in range(1,last+1):
        h_=str(w.cell(hr,c).value or '')
        if c>1 and h_=='Item name': w.column_dimensions[L(c)].width=37.2
        elif h_ in wid_: w.column_dimensions[L(c)].width=wid_[h_]
        elif h_.startswith('Stock zero from') or h_.startswith('Arrival at plant'):      # as wide as the longest text (one line)
            w.column_dimensions[L(c)].width=min(round(max(len(str(w.cell(r_,c).value or '')) for r_ in range(hr+1,w.max_row+1))*0.9+3,1),120)
    w.sheet_view.zoomScale=90
    w.auto_filter.ref=f'A{hr}:{L(last)}{max(w.max_row,hr+1)}'
    _dms_weekly_style(w,hr,last)
def _dms_weekly_style(w, hr, last):
    """DMS BOM Weekly look = the 1_MRP 'Material Status' sheet (user 25-Nov-2026: "organize it like the sheet Material Status ... I don't
       want rows with different heights"; replaces the wrapped A..S table of the same day): group band row 1 (MATERIAL / STORE / COVER /
       SHORTAGE / WAREHOUSE / WEEKS, merged over each group), header row in the group colour (white bold 9 pt, centered, wrapped), data
       rows Calibri 10, ONE line per cell (no wrap), EVERY data row 19 high, thin grey borders, light tint per column group, text left /
       numbers centered. Week cell colours and band-row fills stay. Long text is cut at the column width (full text in the cell / web page).
       Presentation only."""
    from copy import copy as _cp
    from openpyxl.styles import Border, Side
    wk=lambda h: bool(re.fullmatch(r'\d{1,2}-[A-Za-z]{3}',h))
    heads={c:str(w.cell(hr,c).value or '').strip() for c in range(1,last+1)}
    grp_={'Item name':'M',BOMD_HEAD:'M','Item Category':'M','UOM':'M','Estimated stock today':'S','SAP stock (MB52)':'S','Physical count':'S',
          'Days of cover (stock only)':'C',SUB_HEAD:'C','Usage/Day (module plan)':'C','First stock-zero day':'C',
          'Arrival at plant (Final AtA)':'W',PO_HEAD:'W',PR_HEAD:'W',PGRN_HEAD:'W',POD_HEAD:'W',STS_HEAD:'W',GRR_HEAD:'W',GRL_HEAD:'W'}
    look={'M':('MATERIAL','14213D',None),'S':('STORE · stock','2E86C1','F4F9FD'),'C':('COVER · production','6C3483','F8F4FB'),
          'Z':('SHORTAGE','922B21','FDF5F4'),'W':('WAREHOUSE · supply','229954','F4FBF7'),
          'K':('DAYS OF COVER AT THE START OF EACH WEEK (red < 7 · amber < 21 · green covered)','14213D',None)}
    g={}; seen_name=False
    for c,h in heads.items():
        if wk(h): g[c]='K'
        elif h=='Item name' and seen_name: g[c]='Z'
        elif h.startswith('Stock zero from'): g[c]='Z'
        else: g[c]=grp_.get(h,'M' if c==1 else g.get(c-1,'M'))
        if h=='Item name': seen_name=True
    wid={'Item name':34,BOMD_HEAD:34,'Item Category':18,'UOM':6,'Estimated stock today':11,'SAP stock (MB52)':11,'Physical count':11,
         'Days of cover (stock only)':10,SUB_HEAD:44,'Usage/Day (module plan)':11,'First stock-zero day':11,'Arrival at plant (Final AtA)':48,
         PO_HEAD:42,PR_HEAD:42,PGRN_HEAD:11,POD_HEAD:14,STS_HEAD:40,GRR_HEAD:11,GRL_HEAD:16}
    for c,h in heads.items():
        if wk(h): w.column_dimensions[L(c)].width=9
        elif h.startswith('Stock zero from'): w.column_dimensions[L(c)].width=40
        elif h in wid: w.column_dimensions[L(c)].width=wid[h]
    for rng in list(w.merged_cells.ranges):
        if rng.min_row<=hr: w.unmerge_cells(str(rng))
    th=Side(style='thin',color='D5DBE1'); bd=Border(left=th,right=th,top=th,bottom=th)
    c0=1
    while c0<=last:                                               # band row 1: one merged title per group run
        c1=c0
        while c1+1<=last and g[c1+1]==g[c0]: c1+=1
        t,col,_=look[g[c0]]
        for c in range(c0,c1+1):
            x=w.cell(1,c); x.value=t if c==c0 else None
            x.font=Font(name='Calibri',size=10,bold=True,color='FFFFFF'); x.fill=PatternFill('solid',fgColor=col)
            x.alignment=Alignment(horizontal='center',vertical='center'); x.border=bd
        if c1>c0: w.merge_cells(start_row=1,start_column=c0,end_row=1,end_column=c1)
        c0=c1+1
    w.row_dimensions[1].height=20
    for c in range(1,last+1):                                     # header row: group colour, white bold 9 pt, wrapped (the only wrapped row)
        x=w.cell(hr,c); x.font=Font(name='Calibri',size=9,bold=True,color='FFFFFF'); x.fill=PatternFill('solid',fgColor=look[g[c]][1])
        x.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); x.border=bd
    w.row_dimensions[hr].height=45
    for r_ in range(hr+1,w.max_row+1):
        if all(w.cell(r_,c).value is None for c in range(1,last+1)): continue
        band=str(w.cell(r_,1).fill.fgColor.rgb or '').upper().endswith('E8ECF1')
        for c in range(1,last+1):
            x=w.cell(r_,c); v=x.value
            f_=_cp(x.font); f_.name='Calibri'; f_.sz=10
            if band: f_.b=True
            x.font=f_; x.border=bd
            if band: x.fill=PatternFill('solid',fgColor='E8ECF1')
            elif g[c]!='K':
                fg=str(x.fill.fgColor.rgb or '').upper() if x.fill and x.fill.fill_type else ''
                if (not fg or fg.endswith('FFFFFF') or fg=='00000000') and look[g[c]][2]: x.fill=PatternFill('solid',fgColor=look[g[c]][2])
            num=isinstance(v,(int,float,dt.datetime,dt.date)) and not isinstance(v,bool)
            x.alignment=Alignment(horizontal='left' if band or not (num or g[c]=='K' or v in ('-','na') or heads[c]=='UOM') else 'center',vertical='center',wrap_text=False)
        w.row_dimensions[r_].height=19
    w.sheet_view.showGridLines=False
    for c in range(5,min(19,last)+1): w.column_dimensions[L(c)].width=11     # user 25-Nov-2026: "column width 11 from column E to S"
    _dms_cover_filter(w,hr,last)
def _dms_cover_filter(w, hr, last):
    """Print / view filter (user 25-Nov-2026: "pdf, I want to print Days of cover (stock only) less than 22 days only"): the sheet's
       autofilter is set on 'Days of cover (stock only)' < PRINT_MAX_COVER and the other material rows are hidden, so Ctrl+P / Save as
       PDF prints only those. A group band row stays visible when one of its materials is shown; rows under the table stay. Clear the
       filter in Excel to see every material. Values unchanged."""
    jc=next((c for c in range(1,last+1) if str(w.cell(hr,c).value or '').strip()=='Days of cover (stock only)'),None)
    if not jc: return
    from openpyxl.worksheet.filters import FilterColumn, CustomFilters, CustomFilter
    rows=[r_ for r_ in range(hr+1,w.max_row+1) if any(w.cell(r_,c).value not in (None,'') for c in range(1,last+1))]
    band=lambda r_: str(w.cell(r_,1).fill.fgColor.rgb or '').upper().endswith('E8ECF1')
    end=max((r_ for r_ in rows if not band(r_) and w.cell(r_,jc).value is not None),default=hr)
    show={}; cur=None
    for r_ in rows:
        if r_>end: break
        if band(r_): cur=r_; show[r_]=False; continue
        v=w.cell(r_,jc).value
        ok=isinstance(v,(int,float)) and not isinstance(v,bool) and v<PRINT_MAX_COVER
        show[r_]=ok
        if ok and cur: show[cur]=True
    for r_,ok in show.items(): w.row_dimensions[r_].hidden=not ok
    w.auto_filter.ref=f'A{hr}:{L(last)}{end}'
    w.auto_filter.filterColumn=[FilterColumn(colId=jc-1,customFilters=CustomFilters(customFilter=[CustomFilter(operator='lessThan',val=str(PRINT_MAX_COVER))]))]
def _drop_alert_col(w, hr):
    """Delete the 'Alert' class column (column D) of a BOM Material Dashboard (user, 15-Nov-2026: "Column alert is not helpful", both outputs).
       The risk stays visible in the week / day colours, Short from and Inv. Days; the ALERT tab keeps its classes. Cells, widths, merges move left."""
    if str(w.cell(hr,4).value or '').strip()!='Alert': return False
    _del_cols(w,[4])
    return True
def _heat_final(w):
    """BOM Material Dashboard - Weekly, last step in both outputs (user, Nov 2026: "Bom material dashboard: group column E to N. Font size
       9. Delete row 1 & 2"): title + note rows deleted (header row 4 -> 2, data from row 3), every cell font 9 pt, the Material Status
       columns (E .. the column before the first week = E..N in 4_Status with the supplier column, E..M in 1_MRP)
       grouped as an Excel outline (expanded; the - / + button above the columns collapses them). Values unchanged."""
    import copy as _cp
    from openpyxl.utils import get_column_letter as _gl
    _remap_rows(w,lambda r_: None if r_<=2 else r_-2)
    _drop_alert_col(w,2)
    for row_ in w.iter_rows():
        for c_ in row_:
            f_=c_.font
            if f_ is not None and f_.sz!=9:
                f2_=_cp.copy(f_); f2_.sz=9; c_.font=f2_
    last=next((c_.column for c_ in w[2] if str(c_.value or '').strip().lower()=='days incl. in transit'),None)   # last Material Status column
    first=last+1 if last else None
    if first and first>4:
        for j_ in range(4,first):
            w.column_dimensions[_gl(j_)].outline_level=1
        w.column_dimensions.max_outline=1                             # -> sheetFormatPr outlineLevelCol=1 (openpyxl writes it from this; Excel draws the outline bar from it)
    return first
def _tidy_dashboard(w, hr=4):
    """4_Status dashboards (user, Nov 2026: "Dont freeze or wrap text, I want all dashboards ... look organized" - no cell
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

WATCH_NAME='Shortage Watch'      # 4_Status first sheet (user, 12-Nov-2026: dashboards "design style like" their WSA Material Coverage page)
COV_RED_DAYS=7      # user, 12-Nov-2026: 'Red is under 7 days'
COV_RED='E5484D'; COV_AMB='F5B041'; COV_GRN='CFEBD8'; COV_NA='F0F2F5'; COV_BAND='E8ECF1'; COV_INK='1F2933'; COV_MUT='6B7280'; COV_LINE='E3E7EC'
def _coverage_light(w):
    """4_Status dashboards in the look of the user's 'WSA Material Coverage' page (user, 12-Nov-2026: "can daily material
       status dashboard design style like this" -> "Follow your decision but no need line chart"): white page, dark text, THREE cover
       colours instead of seven - red = under COV_RED_DAYS (7) days, amber = under REPORT_DAYS, green = REPORT_DAYS+, blank = no demand /
       no plan (the user's own wording, 12-Nov-2026). A pure colour translation of the cells the dashboard builders wrote (fills, fonts, cell borders, legend, the
       colour words of the note): no value, row or column changes. 4_Status only - 1_MRP keeps the dark dashboards."""
    from openpyxl.styles import Border, Side
    bins=DAY_BINS(); T=bins[3][1]                          # bin positions are the same for every row (amber = 14 - lead time)
    FILL={DASH_BG:'FFFFFF',DAY_NA:COV_NA,DASH_BAND:COV_BAND,'5D6D7E':COV_BAND}
    for a_,b_,c_ in bins: FILL[c_]=(COV_RED if b_<=COV_RED_DAYS else (COV_AMB if b_<=T else COV_GRN))      # user, 12-Nov-2026: red = under 7 days
    FILL[DAY_NA]='FFFFFF'                                          # user, 12-Nov-2026: 'blank means no plan' (no demand that day / No Plan)
    INK={COV_RED:'FFFFFF',COV_AMB:'5C3B00',COV_GRN:'1E6B3A',COV_NA:'9AA3AE',COV_BAND:COV_INK}      # text on a translated fill
    FONT={DASH_TXT:COV_INK,DASH_MUT:COV_MUT,DASH_RED:'C0392B',DAY_NA_TXT:'9AA3AE',DASH_NATXT:COV_INK}   # text on the page itself
    rgb=lambda col: (col.rgb[-6:].upper() if (col is not None and isinstance(getattr(col,'rgb',None),str)) else None)
    fills={}; fonts={}; sd=Side(style='thin',color='FFFFFF'); BD=Border(left=sd,right=sd,top=sd,bottom=sd); NOBD=Border(); n=0
    labs=DAY_LABS(); keep={0:f'under {COV_RED_DAYS}',2:f'{COV_RED_DAYS} – lead time',4:'lead time +'}
    for row in w.iter_rows():
        for c in row:
            if not c.has_style: continue
            old=rgb(c.fill.fgColor) if (c.fill is not None and c.fill.fill_type=='solid') else None
            new=FILL.get(old)
            if c.value=='na' and old==DAY_NA: c.value=None        # blank = no demand that day (data cells and the legend's 'na' chip)
            elif isinstance(c.value,str) and c.value in labs and old in FILL and old!=DASH_BG:      # legend chips: one per colour
                k=labs.index(c.value)
                if k in keep: c.value=keep[k]
                else:
                    c.value=None; new='FFFFFF'
                    nb=w.cell(c.row,c.column+1)
                    if nb.fill is not None and nb.fill.fill_type=='solid' and rgb(nb.fill.fgColor)==old: nb.fill=PatternFill('solid',fgColor='FFFFFF')
            if new:
                c.fill=fills.setdefault(new,PatternFill('solid',fgColor=new)); n+=1
                if c.border is not None and c.border.left is not None and c.border.left.style: c.border=(NOBD if (new=='FFFFFF' and old==DAY_NA) else BD)   # a blank cell carries no border at all
            elif c.border is not None and c.border.left is not None and c.border.left.style and rgb(c.border.left.color)==DASH_BG: c.border=BD   # Alert chips: the dark page-colour outline
            f_=c.font; fc=rgb(f_.color) if f_ is not None else None
            tgt=INK.get(new) if new in INK else (FONT.get(fc) if (new=='FFFFFF' or (new is None and old is None)) else None)
            if tgt and fc!=tgt:
                key=(f_.name,f_.sz,f_.b,f_.i,tgt)
                if key not in fonts: fonts[key]=Font(name=f_.name,size=f_.sz,bold=f_.b,italic=f_.i,color=tgt)
                c.font=fonts[key]
            if isinstance(c.value,str) and ' blue = ' in c.value:
                c.value=re.sub(r'red = short.*?na = no demand that day',f'red = under {COV_RED_DAYS} Inv. Days · amber = under its lead time · green = lead time or more (covered) · blank = no demand that day / no plan',c.value)
    w.sheet_properties.tabColor=COV_MUT
    return n
def _organize_bom_weekly(wb):
    """'BOM Material Dashboard - Weekly' of 4_Status in the column order of the 'Shortage Watch' sheet (user, 14-Nov-2026:
       "organize the bom materials dashboard output": column order and grouping + match the Shortage Watch layout): Code, Item name,
       Category, Item Category, UOM | the Shortage Watch columns from 'Estimated stock today' to 'Arrival at plant (Final AtA)' (one
       Excel outline group, expanded) | the week columns | 'Usage per module ...' at the end. The values are the Shortage Watch cells of the
       same code (= the ALERT); rows, bands, colours and the week columns are untouched. 4_Status only."""
    import copy as _cp
    if HEAT_NAME not in wb.sheetnames or WATCH_NAME not in wb.sheetnames: return None
    w=wb[HEAT_NAME]; sw=wb[WATCH_NAME]; hr=2
    last=next((c_.column for c_ in w[hr] if str(c_.value or '').strip()=='Days of cover (+ all in transit)'),None)
    if not last or last<4 or str(w.cell(hr,4).value or '').strip()=='Alert': return None     # block = Stock .. Days of cover (+ all in transit) from column D
    wh={str(c_.value).strip():c_.column for c_ in sw[7] if isinstance(c_.value,str)}
    SKIP=('Code','Item name','Category','Item Category','UOM','Usage per module by customer (source: your BOM / MIS Standard BOM)')
    info=[h_ for h_ in sorted(wh,key=wh.get) if h_ not in SKIP]
    um=next((h_ for h_ in wh if h_.startswith('Usage per module')),None)
    wrows={}
    for r_ in range(8,sw.max_row+1):
        if sw.cell(r_,wh['Code']).value: wrows.setdefault(str(sw.cell(r_,wh['Code']).value).strip(),[]).append(r_)
    oit=next((c_.column for c_ in w[hr] if str(c_.value or '').strip()=='In transit'),None)      # old In transit: tells apart a code on 2 rows (333 per WSA sheet)
    wk_last=max(c_.column for c_ in w[hr] if c_.value is not None)
    drows=[r_ for r_ in range(hr+1,w.max_row+1) if IT_CODE.match(str(w.cell(r_,1).value or '').split(' ')[0].strip())]
    bands=[r_ for r_ in range(hr+1,w.max_row+1) if w.cell(r_,1).value is None and '·' in str(w.cell(r_,2).value or '')]
    old_it={r_:w.cell(r_,oit).value for r_ in drows} if oit else {}
    hst=_cp.copy(w.cell(hr,4)._style); dst={r_:_cp.copy(w.cell(r_,4)._style) for r_ in drows}; bst={r_:_cp.copy(w.cell(r_,1)._style) for r_ in bands}
    nweek=wk_last-last
    _del_cols(w,list(range(4,last+1)))                               # old Material Status block (Stock .. Days of cover (+ all in transit))
    _ins_cols(w,4,2)                                                 # Item Category, UOM
    _ins_cols(w,6,len(info))                                         # Shortage Watch columns
    wk0=6+len(info); um_col=wk0+nweek
    def put(r_,c_,v_,h_,st_):
        c=w.cell(r_,c_,v_); c._style=_cp.copy(st_)
        txt=h_ in TXT_; c.alignment=Alignment(horizontal=('left' if txt else 'center'),vertical='center',wrap_text=False)
        return c
    TXT_={'Item Category','UOM','Why (cause)','In transit supplier name (qty)','Next arrival','Arrival at plant (Final AtA)','Stock zero from → to (⚠ = starts inside the transit lead time)'}
    heads=[(4,'Item Category'),(5,'UOM')]+[(6+k,h_) for k,h_ in enumerate(info)]
    if um: heads.append((um_col,um))
    for c_,h_ in heads: put(hr,c_,h_,h_,hst)
    for r_ in drows:
        code=str(w.cell(r_,1).value or '').strip(); cand=wrows.get(code) or []; sr=cand[0] if cand else None
        if len(cand)>1:
            ov=old_it.get(r_); sr=next((x_ for x_ in cand if isinstance(ov,(int,float)) and isinstance(sw.cell(x_,wh['In transit']).value,(int,float))
                                         and abs(sw.cell(x_,wh['In transit']).value-ov)<1),cand[0])
            cand.remove(sr); cand.append(sr)
        for c_,h_ in heads:
            v=sw.cell(sr,wh[h_]).value if sr else '-'
            if v is None or v=='': v='-'
            c=put(r_,c_,v,h_,dst[r_])
            if isinstance(v,(dt.datetime,dt.date)):
                c.number_format='dd-mmm'
                if h_ in ('First stock-zero day','Below lead time from'): c.font=Font(name=c.font.name,size=c.font.sz,bold=True,color='C0392B')
            elif isinstance(v,float): c.number_format='0.0' if h_.startswith(('Inv. Days','Days of cover')) else '#,##0'
            elif isinstance(v,int) and not isinstance(v,bool): c.number_format='#,##0'
    for r_ in bands:
        for c_ in range(4,wk0): w.cell(r_,c_)._style=_cp.copy(bst[r_])
    WID={'Item Category':18,'UOM':7,'Why (cause)':20,'In transit supplier name (qty)':38,'Next arrival':34,'Arrival at plant (Final AtA)':44,
         'Stock zero from → to (⚠ = starts inside the transit lead time)':36}
    for c_,h_ in heads: w.column_dimensions[L(c_)].width=(60 if h_.startswith('Usage per module') else WID.get(h_,16 if len(h_)<26 else 22))
    w.column_dimensions['D'].width=18; w.column_dimensions['E'].width=7
    for j_ in range(1,um_col+2): w.column_dimensions[L(j_)].outline_level=0
    for j_ in range(6,wk0): w.column_dimensions[L(j_)].outline_level=1         # the Shortage Watch block = one group (- / + above the columns)
    w.column_dimensions.max_outline=1
    return len(drows),len(info)

def _shortage_watch(wb, src_name=None):
    """Sheet 'Shortage Watch' (first tab of 4_Status) = the ALERT rows of the workbook, same order, no new calculation (user,
       14-Nov-2026: the sheet "restructured like this": Code, Item name, Category, Item Category, UOM, stock columns, Inv. Days, Usage/Day,
       transit lead time, stock-zero days / first day / from -> to, Below lead time from, Why, Days late, In transit ..., Arrival at plant,
       Usage per module). = the ALERT columns without 'Sr.' / 'Alert', plus 'Item Category' (Materials Status column F) after 'Category'.
       Text with line breaks is shown on one line (no wrapped text on the daily dashboards)."""
    from openpyxl.styles import Border, Side
    if 'ALERT' not in wb.sheetnames: return None
    s=wb['ALERT']
    hr=next((r_ for r_ in range(1,4) if any(str(c.value or '').strip()=='Code' for c in s[r_])),None)
    if hr is None: return None
    hdr={str(c.value).strip():c.column for c in s[hr] if isinstance(c.value,str)}
    names=[h_ for h_,_ in sorted(hdr.items(),key=lambda kv:kv[1]) if h_ not in ('Sr.','Alert')]
    icat={}                                                         # Item Category by code (Materials Status col 'Item Category')
    if 'Materials Status' in wb.sheetnames:
        m=wb['Materials Status']; mh=next((r_ for r_ in range(1,4) if any(str(c.value or '').strip()=='Item Category' for c in m[r_])),None)
        if mh:
            mc={str(c.value).strip():c.column for c in m[mh] if isinstance(c.value,str)}
            for r_ in range(mh+1,m.max_row+1):
                k_=str(m.cell(r_,mc.get('Item code',2)).value or '').strip()
                if k_ and k_ not in icat: icat[k_]=m.cell(r_,mc['Item Category']).value
    if 'Category' in names: names.insert(names.index('Category')+1,'Item Category')
    else: names.insert(3,'Item Category')
    rows=[]
    for r_ in range(hr+1,s.max_row+1):
        code=str(s.cell(r_,hdr['Code']).value or '').strip()
        if not IT_CODE.match(code.split(' ')[0]): continue
        rows.append({h_:(icat.get(code,'-') if h_=='Item Category' else s.cell(r_,hdr[h_]).value) for h_ in names})
    d0=ALERT_D0[0]; base=pd.Timestamp(d0).normalize() if d0 is not None else None
    isd=lambda v: isinstance(v,(dt.datetime,dt.date))
    zin='Stock-zero days inside the transit lead time'; fz='First stock-zero day'
    n0=sum(1 for x in rows if base is not None and isd(x.get(fz)) and pd.Timestamp(x[fz]).normalize()<=base)
    n1=sum(1 for x in rows if isinstance(x.get(zin),(int,float)) and x[zin]>0)-n0
    if WATCH_NAME in wb.sheetnames: del wb[WATCH_NAME]
    ws=wb.create_sheet(WATCH_NAME,0); ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor=COV_RED
    F=lambda **k: Font(name='Calibri',**k)
    ln=Side(style='thin',color=COV_LINE); BB=Border(bottom=ln)
    ws['B1']='Material coverage'; ws['B1'].font=F(size=18,color=COV_INK)
    ws['B2']=(f"Base date {base:%d-%b-%Y}: " if base is not None else '')+f"{n0} material(s) are short today, and {max(n1,0)} more go short within their transit lead time."; ws['B2'].font=F(size=11,color=COV_MUT)
    ws['B4']='Shortage watch'; ws['B4'].font=F(size=12,bold=True,color=COV_INK)
    ws['B5']=("The ALERT list (same rows, same order, highest risk first): stock rolled forward + in transit - module plan. "
              "Stock-zero day = the stock does not cover the day's plan. Transit lead time = the days a new order needs to arrive; Below lead time from = first day the cover is under it."); ws['B5'].font=F(size=9,color=COV_MUT)
    WID={'Code':16,'Item name':32,'Category':18,'Item Category':18,'UOM':7,'Why (cause)':20,'In transit supplier name (qty)':38,'Next arrival':34,
         'Arrival at plant (Final AtA)':44,'Stock zero from → to (⚠ = starts inside the transit lead time)':36,'Usage per module by customer (source: your BOM / MIS Standard BOM)':60}
    TXT={'Code','Item name','Category','Item Category','UOM','Why (cause)','In transit supplier name (qty)','Next arrival','Arrival at plant (Final AtA)',
         'Stock zero from → to (⚠ = starts inside the transit lead time)','Usage per module by customer (source: your BOM / MIS Standard BOM)'}
    for j,h in enumerate(names,2):
        c=ws.cell(7,j,h); c.font=F(size=9,bold=True,color=COV_MUT); c.alignment=Alignment(horizontal=('left' if h in TXT else 'right'),vertical='center'); c.border=BB
        ws.column_dimensions[L(j)].width=WID.get(h,15 if len(h)<24 else 20)
    ws.column_dimensions['A'].width=2
    RED_B={fz,'Below lead time from'}
    for i,x in enumerate(rows):
        r=8+i
        for j,h in enumerate(names,2):
            v=x[h]
            if v is None or v=='': v='-'
            if isinstance(v,str) and '\n' in v: v=' | '.join(t_.strip() for t_ in v.split('\n') if t_.strip())
            c=ws.cell(r,j,v); c.border=BB; c.font=F(size=10,color=(COV_MUT if h in ('Item name','Category','Item Category','UOM') else COV_INK),bold=(h in RED_B and isd(v)))
            if h in RED_B and isd(v): c.font=F(size=10,bold=True,color='C0392B')
            c.alignment=Alignment(horizontal=('left' if h in TXT else 'right'),vertical='center')
            if isd(v): c.number_format='dd-mmm'
            elif isinstance(v,float): c.number_format='0.0' if h.startswith(('Inv. Days','Days of cover')) else '#,##0'
            elif isinstance(v,int) and h in ('Estimated stock today','SAP stock (MB52)','Physical count','In transit','Overdue in transit'): c.number_format='#,##0'
        ws.row_dimensions[r].height=15
    ws.row_dimensions[7].height=20
    return len(rows),n0,max(n1,0)

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
        if x['b'].get('count_stock'):
            alts_=[a_ for a_,s_ in CSPM_ALT.items() if s_==x['b']['mat'] and a_ in CSPM_ALT_QTY]      # CS / PM '_Alt' code: its own source line
            for c_ in [x['b']['mat']]+alts_:
                lines.append(dict(code=x['b']['mat'],type='Stock',source="PPC Count (physical count)",src=c_,batch=None,sloc=None,loc=None,
                                  qty=None,uom=x['b'].get('uom'),factor=1.0,cs=True,
                                  csq=(CSPM_ALT_QTY[c_] if c_ in alts_ else (float(x['b']['stock'])-sum(CSPM_ALT_QTY[a_] for a_ in alts_) if alts_ else None))))
    lines.sort(key=lambda l_:(rrank(l_['code']),l_['type']!='Stock',str(l_['src']),str(l_['batch'] or '')))     # highest shortage risk first
    r=5
    for l_ in lines:
        b=info[l_['code']]
        vals=[l_['code'],b.get('desc') if b.get('count_stock') else b['common'],l_['type'],l_['source'],l_['src'],l_['batch'],l_['sloc'],l_['loc'],
              (round(l_['csq'] if l_.get('csq') is not None else b['stock'],6) if (l_.get('cs') and VALUES_ONLY) else (f"=SUMIFS('Stock Count'!${SC_COL}:${SC_COL},'Stock Count'!$A:$A,$E{r})" if l_.get('cs') else round(l_['qty'],6))),
              l_['uom'],l_['factor'],(round(((l_['csq'] if l_.get('csq') is not None else b['stock']) if l_.get('cs') else l_['qty'])*l_['factor'],6) if VALUES_ONLY else f'=I{r}*K{r}'),b.get('uom')]
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
        c=str(r[10] or '').strip(); bn=str(r[4] or '').strip(); q14_=r[14]
        if c in _SWP_INV() and isinstance(q14_,(int,float)): c,f_=_SWP_INV()[c]; q14_=q14_*f_     # MIS tape code -> your BOM code (CODE_SWAP)
        if r[13]: tpu.setdefault(c,str(r[13]).strip())
        if isinstance(r[3],(int,float)) and r[3] and isinstance(q14_,(int,float)):
            t_=rate_t.setdefault((c,bn),[0.0,0.0]); t_[0]+=q14_
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

GR_LINES={}       # MRP code -> [(GR date, qty in the MRP unit)] from PPC In_Transit 'PPC GRN' + 'GR. Date'
PRPO_LINES={}     # MRP code -> [dict(date, qty, kind 'PO'/'PR', po, sup)] open PR / PO lines of PPC sheet 'PR_PO' - information only (user 24-Nov-2026)
GR_REF={}         # MRP code -> [(Posting Date, Qty in unit of entry)] from PPC sheet 'GR' - REFERENCE ONLY (user 20-Nov-2026), in no calculation
GR_SUP={}         # MRP code -> [(GR date, qty in the MRP unit, supplier name)] - the same lines as GR_LINES with their supplier
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
    """Goods received (GR) per material and day: PPC In_Transit 'PPC GRN' on 'GR. Date' (the shipment's traced material,
       alternate codes converted to the MRP unit like the in-transit quantities)."""
    GR_LINES.clear(); GR_SUP.clear(); GR_NODATE.clear(); alias=alias or {}; fac=fac or {}
    read_ppc_gr(mis_path,alias,fac)
    try: rows=list(wb_ro(mis_path)['In Transit'].iter_rows(values_only=True))
    except Exception as _e_: _swallow(5413,_e_); return
    I_=IT_CANON.index; iF,iG,iD,iS=I_('Final RM Code'),I_('PPC GRN'),I_('GR. Date'),I_('Supplier Name')
    for r in rows[1:]:
        if not r or len(r)<=iD: continue
        q=r[iG]; c=str(r[iF] or '').strip()
        if not c or not isinstance(q,(int,float)) or isinstance(q,bool) or q<=0: continue
        if c in alias: q=q*fac.get(c,1.0); c=alias[c]
        d=_any_date(r[iD])
        if d is None: GR_NODATE[c]=GR_NODATE.get(c,0.0)+float(q); continue
        GR_LINES.setdefault(c,[]).append((d,float(q))); GR_SUP.setdefault(c,[]).append((d,float(q),' '.join(str(r[iS] or '').split())))
def _mb52_final(mis_path):
    """{material: Final Material Code} of the MIS MB52 sheet (col C -> col P)."""
    fin={}
    try:
        for r in list(wb_ro(mis_path)['MB52'].iter_rows(values_only=True))[1:]:
            if r and len(r)>15 and r[2] and r[15]: fin.setdefault(str(r[2]).strip(),str(r[15]).strip())
    except Exception as _e_: _swallow(0,_e_)
    return fin
def read_pr_po(mis_path, alias, fac):
    """PPC sheet 'PR_PO' (SAP purchase requisitions + PO schedule lines; user 24-Nov-2026 'Deep dive sheet pr_po ... insert ETA next to AtA')
       -> PRPO_LINES, the supply that is NOT yet on PPC In_Transit: per line (deleted ones skipped)
         PO part = 'Quantity ordered' minus what SAP already received on that PO + material (PPC GR sheet, oldest delivery date first)
                   minus what PPC In_Transit still shows pending on that PO + material (that one is under 'Arrival at plant');
         PR part = 'Quantity requested' - 'Quantity ordered' (no PO yet), same 'Delivery Date'.
       Quantities in the MRP unit (alternate-code factor like the GR sheet). Information only - in no calculation."""
    PRPO_LINES.clear()
    if not PPC_SRC[0]: return False
    nh=lambda v: re.sub(r'[^a-z]','',str(v or '').lower())
    try:
        w=wb_ro(PPC_SRC[0]); n_=next((n for n in w.sheetnames if nh(n)=='prpo'),None) or next((n for n in w.sheetnames if nh(n)=='pr'),None)
        if not n_: return False
        rows=list(w[n_].iter_rows(values_only=True))
        npo_=next((n for n in w.sheetnames if nh(n)=='po'),None)      # PPC of 25-Nov-2026: separate 'PO' sheet (open PO lines, no delivery date) + 'PR' sheet (Delivery Date + Purchase order)
        porows=list(w[npo_].iter_rows(values_only=True)) if npo_ and n_!=npo_ else []
        g_=next((n for n in w.sheetnames if n.strip().lower()=='gr'),None); grr=list(w[g_].iter_rows(values_only=True)) if g_ else []
    except Exception as _e_: _swallow(0,_e_); return False
    if not rows: return False
    hl=[nh(x) for x in rows[0]]
    need=('material','quantityrequested','quantityordered','purchaseorder','deliverydate')
    if not set(need)<=set(hl): print(f"NOTE: PPC sheet '{n_}' has no column(s) {', '.join(k for k in need if k not in hl)} - ETA columns left out",flush=True); return False
    ix=lambda k: hl.index(k) if k in hl else None
    iM,iR,iO,iP,iD,iX,iS=ix('material'),ix('quantityrequested'),ix('quantityordered'),ix('purchaseorder'),ix('deliverydate'),ix('deletionindicator'),ix('nameofsupplier')
    iI=ix('purchaseorderitem')
    num=lambda v: (float(v) if isinstance(v,(int,float)) and not isinstance(v,bool) else float(str(v).replace(',','')) if str(v or '').replace(',','').replace('.','',1).replace('-','',1).isdigit() else 0.0)
    po_=lambda v: re.sub(r'\.0$','',str(v if v is not None else '').strip())
    gv=lambda r,i: r[i] if (i is not None and i<len(r)) else None
    # SAP GR per PO + material (raw SAP code), PPC GR sheet
    gr_={}
    if grr:
        gh=[nh(x) for x in grr[0]]
        if {'material','purchaseorder','qtyinunitofentry'}<=set(gh):
            a,b,c_=gh.index('material'),gh.index('purchaseorder'),gh.index('qtyinunitofentry'); h_=gh.index('debitcreditind') if 'debitcreditind' in gh else None
            for r in grr[1:]:
                if not r or not gv(r,b) or not gv(r,a): continue
                q=abs(num(gv(r,c_)))*(-1.0 if h_ is not None and str(gv(r,h_) or '').strip().upper()=='H' else 1.0)
                k_=(po_(gv(r,b)),str(gv(r,a)).strip()); gr_[k_]=gr_.get(k_,0.0)+q
    # PPC In_Transit pending per PO + traced material (MRP unit) -> converted back to the raw SAP unit is not needed: compared in the MRP unit below
    it_={}
    for x in ((CAP[0] or {}).get('ppc') or {}).get('lines') or []:
        c_=x.get('mat') or x.get('code') or x.get('final_code')
        if x.get('po') in (None,'') or not c_: continue
        q=float(x['qty_mrp'] if x.get('qty_mrp') is not None else x.get('qty') or 0.0)
        k_=(po_(x.get('po')),str(c_).strip()); it_[k_]=it_.get(k_,0.0)+q
    fin=_mb52_final(mis_path); lines={}
    po_open={}; po_mat={}                                              # PO sheet: (PO, item) -> 'Still to be delivered (qty)' (already net of SAP GR); raw material -> its PO items
    if porows:
        ph=[nh(x) for x in porows[0]]
        if {'purchasingdocument','item','material','stilltobedeliveredqty'}<=set(ph):
            a_,b_,c2,d_=ph.index('purchasingdocument'),ph.index('item'),ph.index('material'),ph.index('stilltobedeliveredqty')
            for r in porows[1:]:                                       # ('Deletion Indicator' of this sheet is an age in days, not a flag - not used)
                if not r or not gv(r,a_) or not gv(r,c2): continue
                q=num(gv(r,d_)); k_=(po_(gv(r,a_)),po_(gv(r,b_)))
                if q>0: po_open[k_]=po_open.get(k_,0.0)+q; po_mat.setdefault(str(gv(r,c2)).strip(),[]).append(k_)
        else: porows=[]
    for raw_ in po_mat:
        c0=fin.get(raw_,raw_); lines.setdefault((raw_,c0,fac.get(c0,1.0) if c0 in alias else 1.0),[])
    for r in rows[1:]:
        if not r or not gv(r,iM): continue
        if iX is not None and str(gv(r,iX) or '').strip().lower() in ('true','x','1','l','s'): continue        # deletion indicator
        d=gv(r,iD)
        d=_any_date(dt.datetime.strptime(str(d).strip()[:8],'%Y%m%d')) if re.fullmatch(r'\d{8}(\.0)?',str(d or '').strip()) else _any_date(d)
        if d is None: continue
        raw=str(gv(r,iM)).strip(); c=fin.get(raw,raw); f_=1.0
        if c in alias: f_=fac.get(c,1.0); c=alias[c]
        lines.setdefault((raw,c,f_),[]).append(dict(date=d,req=num(gv(r,iR)),ordq=num(gv(r,iO)),po=po_(gv(r,iP)),item=po_(gv(r,iI)),sup=' '.join(str(gv(r,iS) or '').split())))
    n=0
    for (raw,c,f_),L_ in lines.items():
        L_.sort(key=lambda z:z['date']); used={}
        for z in L_:                                                   # PO part: oldest delivery date takes the received / in-transit quantity first
            if z['po'] and z['ordq']>0:
                k_=(z['po'],raw)
                if porows:                                             # PO sheet: open qty of this PO item (a PO not on the sheet = delivered / closed)
                    ki_=(z['po'],z['item']); pool=po_open.get(ki_,0.0)-used.get(('p',ki_),0.0)
                    left=min(max(pool,0.0),z['ordq']); used[('p',ki_)]=used.get(('p',ki_),0.0)+left
                else:
                    pool=gr_.get(k_,0.0)-used.get(('g',k_),0.0)
                    take=min(max(pool,0.0),z['ordq']); used[('g',k_)]=used.get(('g',k_),0.0)+take; left=z['ordq']-take
                pool=it_.get((z['po'],c),0.0)-used.get(('i',k_),0.0)
                take=min(max(pool/f_ if f_ else 0.0,0.0),left); used[('i',k_)]=used.get(('i',k_),0.0)+take*f_; left-=take
                if left>1e-9: PRPO_LINES.setdefault(c,[]).append(dict(date=z['date'],qty=left*f_,kind='PO',po=z['po'],sup=z['sup'])); n+=1
            rem=z['req']-z['ordq']                                     # PR part: requested, no PO yet
            if rem>1e-9: PRPO_LINES.setdefault(c,[]).append(dict(date=z['date'],qty=rem*f_,kind='PR',po='',sup=z['sup'])); n+=1
        for ki_ in po_mat.get(raw,[]):                                 # open PO quantity without a delivery date (PO sheet has no date; PR sheet dates only some PO items) = 'No date'
            left=po_open.get(ki_,0.0)-used.get(('p',ki_),0.0); used[('p',ki_)]=po_open.get(ki_,0.0)
            k_=(ki_[0],raw); pool=it_.get((ki_[0],c),0.0)-used.get(('i',k_),0.0)
            take=min(max(pool/f_ if f_ else 0.0,0.0),left); used[('i',k_)]=used.get(('i',k_),0.0)+take*f_; left-=take
            if left>1e-9: PRPO_LINES.setdefault(c,[]).append(dict(date=None,qty=left*f_,kind='PO',po=ki_[0],sup=''))
    print(f"PR_PO (open PR / PO lines, information only, in no calculation) read from {os.path.basename(PPC_SRC[0])} sheet '{n_}': {len(rows)-1} lines -> {n} open lines, {len(PRPO_LINES)} materials",flush=True)
    return True
def _ata_from_po(w, hr, at, rows_, code_col):
    """BOM Material Dashboard - Weekly, 'Arrival at plant (Final AtA)' (user 25-Nov-2026: "leave blank if nothing on the way. For material that
       no info in intransit sheet, update info from eta - po (po delivery date)"): a row with PPC In_Transit arrivals keeps them; a row with none
       takes the dated open PO lines of the PPC PR / PO sheets ('PO ETA: dd-Mmm qty ...', `pr_po_text` kind 'PO', undated quantity left out),
       then the dated open PR lines ('PR ETA: ...', user 25-Nov-2026 "need PR ETA also");
       nothing on either = blank. 4_Status only, information - the ALERT / projections are unchanged. Returns (filled, blank)."""
    nf=nb=0; jn=next((c_.column for c_ in w[hr] if str(c_.value or '').strip()=='Item name'),None)
    for r_ in rows_:
        c=w.cell(r_,at); v=str(c.value or '').strip()
        if v and v!='nothing on the way':                     # PPC In_Transit arrivals: quantity -> days of cover (user 25-Nov-2026)
            cd_=str(w.cell(r_,code_col).value or '').strip().upper(); nm_=str(w.cell(r_,jn).value or '').strip() if jn else ''
            nd_=HEAT_NEED.get((cd_,nm_),HEAT_NEED.get((cd_,None)))
            c.value=ata_days(c.value,nd_); continue
        codes=[c_.strip() for c_ in re.split(r'\s*(?:/|,|\+|\n)\s*',str(w.cell(r_,code_col).value or '')) if IT_CODE.match(c_.strip())]
        cd_=str(w.cell(r_,code_col).value or '').strip().upper(); nm_=str(w.cell(r_,jn).value or '').strip() if jn else ''
        nd_=HEAT_NEED.get((cd_,nm_),HEAT_NEED.get((cd_,None)))
        t=[(k_+' ETA: '+eta_days(x_,nd_)) for k_ in ('PO','PR') for x_ in [pr_po_text(codes,kind=k_,dated_only=True) if codes and PRPO_LINES else '-'] if x_ and x_!='-']
        if t: c.value='  ;  '.join(t); nf+=1           # user 25-Nov-2026: "need PR ETA also" - PO first, then PR (one line in Excel, user 25-Nov-2026 'no different row heights'; the web page breaks the line at ';'); qty = days of cover
        else: c.value=None; nb+=1
    return nf,nb
def _eta_cols_days(w, hr, rows_, code_col):
    """'ETA - PO' / 'ETA - PR' columns of the BOM Weekly: quantity -> days of cover of that quantity (`eta_days`, user 25-Nov-2026)."""
    jn=next((c_.column for c_ in w[hr] if str(c_.value or '').strip()=='Item name'),None); n=0
    for j in [c_.column for c_ in w[hr] if str(c_.value or '').strip() in (PO_HEAD,PR_HEAD)]:
        for r_ in rows_:
            c=w.cell(r_,j)
            if not isinstance(c.value,str) or c.value.strip() in ('','-'): continue
            cd_=str(w.cell(r_,code_col).value or '').strip().upper(); nm_=str(w.cell(r_,jn).value or '').strip() if jn else ''
            nd_=HEAT_NEED.get((cd_,nm_),HEAT_NEED.get((cd_,None)))
            v_=eta_days(c.value,nd_)
            if v_!=c.value: c.value=v_; n+=1
    return n
def _prpo_by(codes, kind=None):
    """Open PR / PO lines of a row's codes (alias codes too, supplier rule applied): ({(date, kind): qty}, undated qty). One source for
       the DMS 'ETA - PO' / 'ETA - PR' text and the web page Arrival / Weeks tabs (user 25-Nov-2026)."""
    by={}; seen=set(); undated=0.0
    for c_ in _gr_ref_codes(codes):
        if c_ in seen: continue
        seen.add(c_); _,_,is_dead=vp_dead(c_)
        for z in PRPO_LINES.get(c_,[]):
            if (kind and z['kind']!=kind) or (is_dead and z['sup'] and is_dead(z['sup'])): continue
            if z['date'] is None: undated+=z['qty']; continue
            k_=(z['date'],z['kind']); by[k_]=by.get(k_,0.0)+z['qty']
    return by,undated
def pr_po_text(codes, sheet=None, own=False, kind=None, dated_only=False):
    """'ETA - PO' / 'ETA - PR' (kind) of a 4_Status row: open lines of the row's codes (PPC PR_PO, `read_pr_po`), one line:
       passed dates as one total, the next ATA_SHOW dates ('dd-Mmm qty'), the rest as one total. '-' = nothing open.
       Supplier rule (customer goes with supplier): a line of a supplier no customer of the material may use is left out."""
    d0=base_d0(); by,undated=_prpo_by(codes,kind)
    if dated_only: undated=0.0
    if not by and undated<=1e-9: return '-'
    past=[(k,q) for k,q in sorted(by.items()) if k[0]<d0]; fut=[(k,q) for k,q in sorted(by.items()) if k[0]>=d0]
    ln=[]
    if past: ln.append(f"Date passed: {sum(q for _,q in past):,.0f} ({past[0][0][0]:%d-%b}"+('' if len(past)==1 or past[0][0][0]==past[-1][0][0] else f" → {past[-1][0][0]:%d-%b}")+')')
    ln+= [f"{k[0]:%d-%b} {q:,.0f}" for k,q in fut[:ATA_SHOW]]
    if len(fut)>ATA_SHOW: rest=fut[ATA_SHOW:]; ln.append(f"+{len(rest)} more: {sum(q for _,q in rest):,.0f} ({rest[0][0][0]:%d-%b}"+('' if rest[0][0][0]==rest[-1][0][0] else f" → {rest[-1][0][0]:%d-%b}")+')')
    if undated>1e-9: ln.append(f"No date: {undated:,.0f}")
    return '  |  '.join(ln)
def read_ppc_gr(mis_path, alias, fac):
    """PPC sheet 'GR' (SAP material documents) -> GR_REF: 'Qty in unit of entry' on 'Posting Date' (Debit/Credit 'H' = minus).
       Material -> MIS MB52 Final Material Code, then the alternate-code alias. User 20-Nov-2026: "GR sheet is for reference only,
       don't count it in any calculation yet" -> only the 4_Status reference columns read GR_REF (`gr_ref_since`, `gr_ref_last`)."""
    GR_REF.clear()
    if not PPC_SRC[0]: return False
    try:
        w=wb_ro(PPC_SRC[0]); n_=next((n for n in w.sheetnames if n.strip().lower()=='gr'),None)
        if not n_: return False
        rows=list(w[n_].iter_rows(values_only=True))
    except Exception as _e_: _swallow(0,_e_); return False
    if not rows: return False
    hl=[' '.join(str(x or '').split()).lower() for x in rows[0]]
    if not {'material','posting date','qty in unit of entry'}<=set(hl): return False
    iM,iD,iQ=hl.index('material'),hl.index('posting date'),hl.index('qty in unit of entry'); iH=hl.index('debit/credit ind') if 'debit/credit ind' in hl else None
    fin=_mb52_final(mis_path)
    n=0
    for r in rows[1:]:
        if not r or len(r)<=max(iM,iD,iQ): continue
        c=str(r[iM] or '').strip(); q=r[iQ]
        if isinstance(q,str):
            try: q=float(q.replace(',',''))
            except ValueError: q=None
        if not c or not isinstance(q,(int,float)) or isinstance(q,bool) or not q: continue
        q=abs(float(q))*(-1.0 if iH is not None and str(r[iH] or '').strip().upper()=='H' else 1.0)
        c=fin.get(c,c)
        if c in alias: q=q*fac.get(c,1.0); c=alias[c]
        d=_any_date(r[iD]); n+=1
        if d is not None: GR_REF.setdefault(c,[]).append((d,q))
    print(f"SAP GR (reference only, in no calculation) read from {os.path.basename(PPC_SRC[0])} sheet '{n_}': {n} lines",flush=True)
    return True
def _gr_ref_codes(codes):
    """{GR code: the MRP code whose count date applies} - the row's codes, their BOM alias and CS / PM '_Alt' codes (cspm_alt_fold)."""
    al_=BOM_ALIAS[0] or {}; cs={c_:al_.get(c_,c_) for c_ in codes}; cs.update({al_.get(c_,c_):al_.get(c_,c_) for c_ in codes})
    cs.update({a_:s_ for a_,s_ in CSPM_ALT.items() if s_ in cs.values()})
    return cs
def gr_ref_since(codes, sheet=None, own=False):
    """4_Status reference column: PPC GR 'Qty in unit of entry' with Posting Date after the material's count date, up to the
       base date. Not counted = '-' (SAP MB52 already holds those receipts). Information only - no calculation uses it."""
    d0=base_d0(); t=0.0; any_=False
    for c_,m_ in _gr_ref_codes(codes).items():
        e_=STOCK_EST.get(m_) or STOCK_EST.get(c_) or {}
        if e_.get('date') is None or e_.get('count') is None: continue
        any_=True; a_=pd.Timestamp(e_['date']).normalize()
        t+=sum(q for d,q in GR_REF.get(c_,[]) if a_<d<=d0)
    return round(t) if any_ else '-'
def gr_ref_last(codes, sheet=None, own=False):
    """4_Status reference column: latest PPC GR Posting Date (up to the base date) and its quantity, e.g. '20-Nov (52,560)'."""
    d0=base_d0(); L=[(d,q) for c_ in _gr_ref_codes(codes) for d,q in GR_REF.get(c_,[]) if d<=d0]
    if not L: return '-'
    m_=max(d for d,q in L); return f"{m_:%d-%b} ({sum(q for d,q in L if d==m_):,.0f})"
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
    idt=next((hl.index(k) for k in ('physical count date','count date') if k in hl),None)   # one 'Count date' column (Count sheet since 19-Nov-2026)
    per={}
    for r in rows[hi+1:]:
        if not r or not r[0] or ic is None or ic>=len(r) or not isinstance(r[ic],(int,float)) or isinstance(r[ic],bool): continue   # counted rows only (= the MRP)
        c=str(r[0]).strip(); f=1.0
        if c in alias: f=fac.get(c,1.0); c=alias[c]
        h={d:float(r[j])*f for j,d in dcols if j<len(r) and isinstance(r[j],(int,float)) and not isinstance(r[j],bool)}
        if not dcols and idt is not None and idt<len(r) and isinstance(r[idt],(dt.datetime,dt.date)): h={pd.Timestamp(r[idt]).normalize():float(r[ic])*f}
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
       rolled-forward plant stock (st_main, JW at the OEM not included): count - usage by the count share, and every GR since the count
       on the row of ITS OWN supplier (PPC In_Transit 'Supplier Name' of the GRN line; user 21-Nov-2026 "receipt always go to its own supplier"). A supplier with MB52
       batches but no count row = 0; no 'Unassigned' row. Runs after estimate_counted_stock, so MRP Daily, the Action List supplier
       rows and the WSA sheets (align_capture 'sup') all read the same split. Returns the number of materials changed."""
    n_=0
    for b in blocks:
        cnt_=(COUNTED_RM.get(b['mat']) or {}).get('by_sup') or {}
        if not b.get('sup') or not cnt_: continue
        tot_=sum(float(q_[1] or 0) for q_ in cnt_.values())
        own_=float(b['st_main']) if b.get('st_main') is not None else float(b['stock'])-(0.0 if b.get('jw_out') else float(b.get('st_jw') or 0))
        e_=STOCK_EST.get(b['mat']) or {}
        rs_={k_:v_ for k_,v_ in (e_.get('rec_sup') or {}).items() if v_} if e_.get('est') is not None and e_.get('days') else {}   # receipts in the roll-forward
        f_=(max(0.0,own_-sum(rs_.values()))/tot_) if tot_>0 else 0.0     # count - usage by the count share (plant stock only, JW at the OEM is not in the count);
                                                                          # each receipt since the count goes to ITS OWN supplier (user, 21-Nov-2026)
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
        for gs_,gq_ in rs_.items():                                       # receipt -> the supplier row of that supplier (first word / Count supplier code)
            gu_=gs_.upper(); w1_=(re.findall(r'[A-Z0-9]+',gu_) or [''])[0]
            r_=next((sp for sp in rows_ if str(sp.get('name') or '').upper()!='UNASSIGNED' and w1_ and
                     (w1_ in str(sp.get('name') or '').upper() or (re.findall(r'[A-Z0-9]+',str(sp.get('name') or '').upper()) or ['#'])[0] in gu_)),None)
            if r_ is None:
                nd_=len(rows_[0]['transit']) if rows_ else 0
                r_=dict(name=w1_ or gs_ or 'Unassigned',stock=0.0,transit=[0.0]*nd_,openpo=[0.0]*nd_); rows_.append(r_)
            r_['stock']+=gq_
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
       estimated stock = count - usage per day x days since the count + goods received (PPC GRN) after the count date.
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
        alts_=[a_ for a_,s_ in CSPM_ALT.items() if s_==m]                # CS / PM '_Alt' codes = the same material (cspm_alt_fold)
        sv_=[sap[c_] for c_ in [m]+alts_ if c_ in sap]
        e=dict(sap=(sum(sv_) if sv_ else None),count=cnt,date=date,days=None,use_day=None,use=0.0,rec=0.0,est=None,src='SAP (MB52) - not counted' if cnt is None else 'Physical count')
        if cnt is not None and date is not None:
            days=max(0,(d0-pd.Timestamp(date).normalize()).days)
            P=_plan_without_oem(b); ud=(sum(P[:7])/min(7,len(P))) if P else 0.0
            rc=sum(gr_between(c_,pd.Timestamp(date).normalize(),d0) for c_ in [m]+alts_)   # GR after the count day (GR. Date, Excel number read as a date - was Final AtA)
            rs_={}                                                # the same receipts by their supplier (apply_count_by_supplier: a receipt goes to its own supplier)
            for c_ in [m]+alts_:
                for d_,q_,sp_ in GR_SUP.get(c_,[]):
                    if pd.Timestamp(date).normalize()<d_<=d0: rs_[sp_]=rs_.get(sp_,0.0)+q_
            e.update(days=days,use_day=ud,use=ud*days,rec=rc,rec_sup=rs_)
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
       otherwise -> its MB52 + JW stock and PPC arrivals. Codes already counted inside the standard material (roll lengths, your mapping) are skipped."""
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

LT_OVERRIDE={}     # material code -> transit lead time (days) typed on Settings 'TRANSIT LEAD TIME' (column F 'Your lead time'); wins over the WSA history
LT_BY_CODE={}      # material code (upper) -> transit lead time used this run (dashboards colour each row by its own lead time)
LT_TABLE=[]        # this run's lead time per material (for the Settings table): code, name, RM/PM, history median, deliveries, yours, used, source
LINE_LT={}         # production line key -> longest transit lead time of the materials it uses (line CRITICAL / AT RISK / WATCH)
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
            LINE_LT[key]=max(LINE_LT.get(key,0),tgt_days(b.get('lt_used')))   # line threshold = longest lead time of its materials
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
            if scen=='t': stop_by=list(dict.fromkeys(f'{c} {n}' for kk,n,c in ms if kk==k and k<nd))   # MRP by Line 'Line stops because of'
        out.append(dict(key=key,name=L_['name'],cust=sorted(L_['cust']),t=res['t'],o=res['o'],rt=res['rt'],stop_by=stop_by))
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
    H=['Material (family)','Code','Description','Pack / size','Unit','Stock (MB52)','In transit (PPC)','Found in','MIS Common Name','How the MRP treats it','What to do']
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

