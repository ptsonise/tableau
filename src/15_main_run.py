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
    ETD_FLAGS[2]=(f'SHIPS WITHIN {ETD_SOON} DAYS',)+ETD_FLAGS[2][1:]; ALERT_CLS[3]=(ALERT_CLS[3][0],ALERT_CLS[3][1],ALERT_CLS[3][2],f'falls below its lead time within {HZW()} - expedite / order before the date')   # texts follow the Settings numbers
    print(f'Settings read from {os.path.basename(SETTINGS_SRC[0])}' if SETTINGS_SRC[0] else 'NOTE: no Settings sheet found (MRP.xlsx or last MRP_Daily) - script defaults used',
          f'- every day threshold = the material transit lead time (Settings TRANSIT LEAD TIME, else WSA history, else RM {LEAD_TIME_RM} / packaging {LEAD_TIME_CSPM} days)')
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
    codes_by_name={k_:_bom_code_of(v_,MIS) for k_,v_ in codes_by_name.items()}   # MIS tape code -> your BOM code (same rule as CODE_SWAP)
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
    MISBL=_mis_bl_dates(MIS) if MIS else ({},{})                     # PPC In_Transit BL -> plant arrival date (PPC In_Transit is the latest)
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
    # alert"): every open WSA delivery row (not cancelled) with an ETD whose BL# is not on PPC In_Transit = not shipped yet as far as PPC In_Transit knows
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
        if _mis_date_of(x['bl'],MISBL) is not None: continue                     # shipped: its BL# is on PPC In_Transit
        g_=gby.get(e['g']) or {}; etd_=pd.Timestamp(x['etd']).normalize(); pd_=pd.Timestamp(x['date']).normalize()
        tr_=max(0,(pd_-etd_).days); passed=etd_<base_; blt_=_bl_txt(x['bl']); hasbl=bool(blt_)
        eta_=(base_+pd.Timedelta(days=tr_)) if passed else pd_                   # ETD passed: earliest arrival if it ships today; else the WSA plant date
        ETD_LATE.append(dict(sheet=e['sheet'],label=e['g'],codes=list(g_.get('codes') or []),uom=g_.get('uom'),po=(int(x['po']) if isinstance(x['po'],float) and x['po']==int(x['po']) else x['po']),
                             shipper=x['shipper'],qty=x['qty'],etd=etd_,late=(base_-etd_).days,passed=passed,plant=pd_,transit=tr_,eta_now=eta_,
                             delay=(eta_-pd_).days,bl=blt_,remarks=(('[Archive] ' if e.get('arch') else '')+str(x['remarks'] or '')).strip() or None,
                             status=(('BL# not on PPC In_Transit (received? BL# typo? check)' if hasbl else 'Not shipped (no BL#)') if passed else
                                     ('BL# on WSA, not on PPC In_Transit yet' if hasbl else 'Not shipped yet (no BL#)')),base=base_))
    if ETD_LATE:
        np_=sum(1 for r_ in ETD_LATE if r_['passed']); nf_=len(ETD_LATE)-np_; old_=max((r_['late'] for r_ in ETD_LATE if r_['passed']),default=None)
        print(f"ETD alert (WSA deliveries not on PPC In_Transit): {len(ETD_LATE)} row(s) - ETD passed {np_}"+(f" (oldest {old_} days ago)" if old_ is not None else '')
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
           'Status (clean)','Port','Lead days (plant date - ETD)','Date check','ETD passed, no vessel/BL','Missing info','Days to arrival','Source row','Unique delivery (1 = first row)','Chase?','BOM code','Sheet: Balance','Sheet: Planned modules per day','Sheet: Demand','Sheet: Actual modules per day','Sheet: Actual','Sheet: MW per day','Sheet: Actual MW per day','Sheet: Plant Days','Sheet: Plant + Port Days (ETA +5 days)','Sheet: Delivered','Sheet: Scheduled Delivery','PPC arrival (Final AtA) - the latest; shown when 3+ days from the plant date']
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
        if e.get('misdiff') is not None:                          # PPC In_Transit is the latest: the WSA row is to be moved to this date
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
    tc['A2']='Reasons: past due = dated before the sheet stock date and not cancelled (not counted); missing PO / shipper / ETD; ETD passed with no vessel or BL; short lead = plant date under the short-lead limit after ETD; PPC arrival 3+ days from the plant date (MIS is always the latest: move the WSA row).'; tc['A2'].font=bf
    hrow(tc,4,['Planning sheet','Sheet stock date','Materials','Plant date','Days to arrival (vs report date)','Sheet quantity','Stage','Remarks (as written)','PO','Shipper','ETD','BL#','Vessel / ETA note','Reasons']); tc.row_dimensions[4].height=32
    chase=[e for e in deliveries if e['chase']]
    for i,e in enumerate(sorted(chase,key=lambda e:e['x']['date'])):
        x=e['x']; why=[]
        if e['stage']=='Past due': why.append(f"Past due ({(pd.Timestamp(e['asof'])-pd.Timestamp(x['date'])).days} d before stock date)")
        if e['missing']: why.append('Missing '+e['missing'].replace(' ',', '))
        if e['etdpassed']: why.append('ETD passed, no vessel/BL')
        if e['datecheck']=='Short lead': why.append(f"Short lead ({e['lead']} d after ETD)")
        if e.get('misdiff') is not None: why.append(f"PPC arrival {e['mis_date']:%d-%b} ({e['misdiff']:+d} d) - PPC In_Transit is the latest: move the WSA row")
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
            rn=build_new_mrp(DSn,BW,MIS,outn,prev_path=prev,TD=LEAD_TIME_RM,old_mrp=USER_MRP)
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
            g2_,k2_=delete_old_outputs(os.path.dirname(os.path.abspath(outn)),OUT)    # user, Nov 2026: old Inbound / Daily_Material_Status / Run_Summary go too
            kept_+=k2_
            if g2_: print('Old output files deleted: '+', '.join(g2_))
            if kept_: print('NOTE: could not delete (open in Excel?): '+', '.join(kept_))
            try:                                                     # the daily email to top management (user, Nov 2026)
                dir_=os.path.dirname(os.path.abspath(outn)); dms_=os.path.join(dir_,f'Daily_Material_Status_{dt.date.today():%Y-%m-%d}.xlsx')
                write_daily_email(ALERT_ROWS,ALERT_D0[0],dir_,attach=dms_ if os.path.exists(dms_) else None)
            except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: daily email not written ({type(ex_).__name__}: {ex_})')
            _wsh=[]                                                  # the WSA daily web page (user, 21-Nov-2026 'Do all': the two web pages line up)
            try:                                                     # - only when this run reads a WSA_Daily file (step 2)
                _wsh=wsa_alert_match(wsa_open_deliveries(SRC,ALERT_D0[0]),_heat_expand(ALERT_ROWS,ALERT_D0[0])) if ALERT_D0[0] is not None else []
            except Exception as ex_: _wsh=[]; print(f'NOTE: WSA web page not built ({type(ex_).__name__}: {ex_})')
            WSA_DLV.clear()
            for _s in _wsh:                                          # open WSA deliveries for the detail panel of the other two web pages (user 22-Nov-2026)
                _a=_s.get('_a')
                if _a: WSA_DLV.setdefault(str(_a.get('where') or _a.get('code') or '').strip(),(_s.get('dlv') or [],_s.get('nd') or 0))
            try: write_material_daily_update(ALERT_ROWS,ALERT_D0[0],os.path.dirname(os.path.abspath(outn)))     # match-page style dashboard (user, 22-Nov-2026)
            except Exception as ex_: print(f'NOTE: Material Daily Update not written ({type(ex_).__name__}: {ex_})')
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
    print(f'To chase: {len(chase)} deliveries (past due {npast}, missing PO/shipper/ETD {sum(1 for e in chase if e["missing"])}, ETD passed with no vessel/BL {sum(1 for e in chase if e["etdpassed"])}, short lead {sum(1 for e in chase if e["datecheck"]=="Short lead")}, PPC arrival 3+ days from the plant date {sum(1 for e in chase if e.get("misdiff") is not None)} - PPC In_Transit is the latest)')
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
                        if _ns: print(f'WSA Dashboard sorted highest risk first ({_ns} rows moved; the same risk order as the BOM Material Dashboard)')
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
        _cmd=[sys.executable,os.path.abspath(__file__),_aligned or SRC_USER,'--no-align','--mis',MIS]+(['--mrp',USER_MRP] if USER_MRP else [])+(['--keep-old'] if KEEP_OLD else [])   # step 2 deletes the old outputs: it must know (checker finding)
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
                  "  The outputs of this run were built from your WSA file as it is; \n"
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
