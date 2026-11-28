# ===================== NEW MRP workbook (built from MIS + WSA only) =====================
CAP_DOC=999
def is_unlimited(b):
    t=f"{b.get('desc') or ''} {b.get('common') or ''}".lower()
    return b.get('mat') in UNLIMITED_CODES or any(w in t for w in UNLIMITED_WORDS)

# =================================================================
#  Sheet builders: Today and Planner
# =================================================================




def build_new_mrp(DS, BW, mis_path, out_path, prev_path=None, TD=None, old_mrp=None):
    """1_MRP_<date>.xlsx: ALERT, Material Status, MRP Daily, MRP by Line, MRP Weekly, Check, No Demand, Settings (see the sheet list printed at the end)
       (+ hidden: Line Run Days, Module Plan, BOM, Stock Count, How it works, Movement Log, MIS Layout).
       Target: every material covered for its own target days (Settings: A / B / C / packaging, + Z extra); TD = default for anything unclassified."""
    load_rmi_prev([out_path,prev_path])                 # your RM Issues entries (Owner / Status / Date Closed / Action) before anything is overwritten
    print(f'  [{time.time()-_T0:5.0f} s] building the MRP workbook ...',flush=True)
    TD=TD or LEAD_TIME_RM
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
    # Customer Supplier rule: the Settings sheet first; else the Customer Supplier sheet of your MRP.xlsx; else the last 1_MRP
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
    UBOM_UNIT_ISSUES.clear()
    BOM_ALIAS[0]=DS.get('alias') or {}; BOM_ALIAS[1]=DS.get('alias_factor') or {}
    load_user_masters(SETTINGS_SRC[0]); build_alt_bom(mis_path,DS.get('alias') or {})       # alternatives per customer, before the plan
    DS_rm=dict(DS,blocks=rm_plan_from_mis(DS,mis_path,modplan,lineplan))
    load_user_masters(SETTINGS_SRC[0]); build_alt_bom(mis_path,DS.get('alias') or {})       # your alternative BOM: standard -> alternatives
    maprows,map_issues=apply_supplier_map(DS_rm,maprows,add_new=not maprows)   # rows are added only when the sheet does not exist yet
    CUSTSUP_ROWS[:]=maprows                                                     # shown on the Settings sheet
    valt=line_substitutes(DS_rm,mis_path); vp_issues=list(SUBS_ISSUES)   # substitutes for some lines (Settings)
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
        apply_count_by_supplier(DS['blocks'],DS_rm.get('alias'))                              # counted: stock by supplier = the count by supplier, rolled forward
        for m_,e_ in STOCK_EST.items():
            if e_.get('days') and e_['days']>0 and e_.get('est') is not None:
                u_={b_['mat']:b_.get('uom') for b_ in DS['blocks']}.get(m_,'')
                if e_.get('use'): DS.setdefault('stock_lines',[]).append(dict(code=m_,type='Stock',source=f"Usage since {'count' if e_.get('count') is not None else 'MIS date'} {pd.Timestamp(e_['date']):%d-%b} (estimated, {e_['days']} d × {e_['use_day']:,.1f}/d)",src=m_,batch=None,sloc=None,loc=None,qty=-float(min(e_['use'],(e_['count'] if e_.get('count') is not None else (e_['est']+e_['use']))+e_['rec'])),uom=u_,factor=1.0))
                if e_.get('rec'): DS.setdefault('stock_lines',[]).append(dict(code=m_,type='Stock',source=f"Received since count {pd.Timestamp(e_['date']):%d-%b} (PPC GRN)",src=m_,batch=None,sloc=None,loc=None,qty=float(e_['rec']),uom=u_,factor=1.0))
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
    LT_TABLE[:]=[]; SS_STATS.clear()
    for b_ in DS['blocks']:                # transit lead time = target days of the material (user rule, Nov 2026: Settings rows 6-19 -> lead time)
        h_=lead.get(b_['common']) or []; lt_h=int(statistics.median(h_)) if h_ else None
        yo_=LT_OVERRIDE.get(str(b_['mat']).upper())          # Settings TRANSIT LEAD TIME 'Your lead time' > WSA ETD history > fallback
        fb_=LEAD_TIME_CSPM if b_.get('count_stock') else LEAD_TIME_RM
        b_['lt_used']=yo_ or lt_h or fb_; b_['td']=tgt_days(b_['lt_used']); b_['lt_hist']=lt_h; b_['lt_sd']=(lambda s_: s_ if s_>=1 else None)(statistics.pstdev(h_)) if len(h_)>=2 else None; b_['lt_n']=len(h_); b_['lt_your']=yo_; LT_BY_CODE[str(b_['mat']).upper()]=b_['lt_used']
        LT_TABLE.append((b_['mat'],(b_.get('desc') if b_.get('count_stock') else b_.get('common')) or b_.get('desc'),'Packaging / CS' if b_.get('count_stock') else 'Raw material',
                         lt_h,len(h_) or None,yo_,b_['lt_used'],'Your lead time' if yo_ else ('WSA history' if lt_h else f"Fallback ({'packaging' if b_.get('count_stock') else 'RM'})"),
                         round(b_['lt_sd'],1) if b_.get('lt_sd') is not None else None))
    # ---- OEM: its production plan (customers with OEM = Yes) and the JW stock at the OEM are NOT part of any projection,
    #      shortage date, planned order, dashboard or risk list. The Action List still shows the with-OEM daily target and Inv. days.
    oem_c_=oem_names()                                                   # Settings OEM = Yes + every PPC 'OEM Plan' customer (user 28-Nov-2026)
    _is_oem=lambda cu: is_oem_cust(cu,oem_c_)
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
    for b in DS['blocks']:            # safety stock inputs (`ss_rop`) - before the step-1 capture, so 2_WSA gets them too (OEM plan already out)
        SS_STATS[str(b['mat']).upper()]=dict(cv=demand_cv(b.get('plan') or [],lt_of(b)),sd_lt=b.get('lt_sd'),n_lt=b.get('lt_n') or 0,
                                             cls='PM' if b.get('abc')=='PM' else f"{b.get('abc') or 'C'}{b.get('xyz') if b.get('xyz') in ('X','Y','Z') else 'Z'}")
    try:
        CAP[0]=align_capture(DS,pp_rows,dates)                 # numbers for the WSA planning sheets and the Inbound 'Other materials' sheet
        if ALIGN_MODE[0]: ALIGN[0]=CAP[0]
    except Exception as ex_: print(f'NOTE: WSA numbers not captured ({type(ex_).__name__}: {ex_})')
    if ALIGN_MODE[0] and ALIGN[0] is not None: raise _CaptureDone()   # step 1 of the alignment needs only these numbers: the 1_MRP it
                                                                     # would write goes to a temp folder and is thrown away (speed, Nov 2026)
    read_pr_po(mis_path,BOM_ALIAS[0],BOM_ALIAS[1])                   # open PO / PR lines: before the first 'Arrival at plant' text (ata_show); needs the traced PPC lines of CAP[0]
    for b_ in DS['blocks']:                                           # item name of a swapped tape = YOUR BOM code's roll (after the WSA capture: names link sheets there)
        n_=swap_name(b_)
        if n_: b_['common']=n_
    if oem_c_: print(f"OEM left out of the calculations: plan of {', '.join(sorted({cu for cu,ln,tg,o in CUST_TARGET if str(o).strip().lower()=='yes'}|OEM_PLAN_CUST))} ({n_oem} materials) and JW stock at the OEM - Action List still shows the with-OEM target",flush=True)
    # ---- SUBSTITUTES between two MRP materials (Settings SUBSTITUTES, all lines): NOT pooled (user 25-Nov-2026: "EPE 480 and EPE 420 SEPARATELY on
    #      every sheet and dashboard: each with its own stock, plan, in transit and arrivals"). Every code keeps its own projection; `sub_cover`
    #      (ALERT rows) tells when the substitute covers the main material and with how much. 'Use from / Use until' limit those days.
    SUBG.clear(); byb_={b_['mat'] for b_ in DS['blocks']}
    for s_ in SUBS_X:
        m_,c_=s_.get('mrp_code'),s_.get('sub_code')
        if not m_ or not c_ or s_.get('lines') is not None or m_ not in byb_ or c_ not in byb_: continue
        if c_ not in [x_[0] for x_ in SUBG.get(m_,[])]: SUBG.setdefault(m_,[]).append((c_,float(s_.get('factor') or 1.0),s_.get('fr'),s_.get('to')))
    if SUBG: print('Substitutes (Settings SUBSTITUTES) - each code planned on its OWN stock / plan / in transit; substitute cover shown on the ALERT: '
                   +'; '.join(f"{m_} <- {', '.join(c_ for c_,*_ in v_)}" for m_,v_ in SUBG.items()),flush=True)
    A=[]; VR_=_vendor_rules({'alias':DS.get('alias') or {}}); VP_MRP.clear()
    for b in DS['blocks']:
        TDm=b.get('td') or TD          # target days of this material (ABC / XYZ class)
        P=b['plan']; T=[a+c for a,c in zip(b['transit'],b['adhoc'])]; O=b.get('openpo') or [0.0]*nd
        stk_=b['stock']
        S=[]; cur=stk_
        for k in range(nd): cur=cur-P[k]+T[k]; S.append(cur)
        vp_=None; S_pool=list(S)
        if any(k_[1]==b['mat'] for k_ in VR_):     # Settings ALLOWED VENDORS: same rule as the WSA planning sheet
            try:
                sh_=_wsa_sheet_series(b['mat'],d0,nd)
                # on 2+ WSA sheets (333): the sheets ARE the split by customer + supplier (Settings ALLOWED VENDORS, _wsa_shared) -
                # take it from them (checker finding: the PPC In_Transit shipments alone gave Encap's FLAT HK glass to Subanci)
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
        lt=b.get('lt_your') or b.get('lt_hist')           # your Settings lead time, else the WSA ETD history (None = fallback below)
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
        wins=floor_wins(S,stk_)                      # stock never below 0 (user 23-Nov-2026), same as the ALERT
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
        A.append(dict(stk0=stk_,b=b,cov=cov,below=dt_(below),short_txt=short_txt,wins=wins,exc=exc,exg=exg,nxt_po=nxt_po,mods=mods_txt,proj=proj_txt,short=short,short3=short3,avg=avg,nx=nx,lt=lt,custs=custs,status=status,action=action,
                      runout=dt_(cov[1]) if cov[1]<nd else None,tr_tot=sum(b['transit']),tr_today=sum(s_['qty'] for s_ in ships.get(b['mat'],[]) if s_.get('raw') is not None and s_['raw']==d0),tr_last7=sum(s_['qty'] for s_ in ships.get(b['mat'],[]) if s_.get('raw') is not None and d0-pd.Timedelta(days=7)<=s_['raw']<d0),tr_next14=sum(s_['qty'] for s_ in ships.get(b['mat'],[]) if s_.get('raw') is not None and d0<s_['raw']<=d0+pd.Timedelta(days=14)),op_tot=sum(O),short_d=short_d,short3_d=short3_d,S=S,cO=cO,
                      cover_from_first=cover_from_first,
                      below3=dt_(below3),
                      doc=(S[0]/(sum(P[:TDm])/len(P[:TDm])) if sum(P[:TDm]) else float(CAP_DOC)),S_pool=S_pool))   # MRP Daily 'Inv. Days' on day 1 (own target window)
    A.sort(key=lambda x:(STATUS_ORDER[x['status']],x.get('exg',9),x['cov'][1],x['cov'][2],bool(x['b'].get('count_stock'))))   # most urgent first (status, exception, days)
    if VP_MRP: print('Settings ALLOWED VENDORS - projection by allowed vendor (same rule as the WSA sheets): '+vp_console(VP_MRP,d0),flush=True)
    # ---------- base workbook: Action List (the old Dashboard was removed: headline / top risks -> ALERT) ----------
    wb=Workbook(); al=wb.active; al.title='Action List'
    al.sheet_properties.tabColor='C00000'; set_print(al,'Action List')
    n=len(A);
    # allocation to lines: Customer Supplier (allowed vendors) > the rest > alternative materials by description;
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
    alloc_stop,alloc_used=allocate_lines(DS,valt,cmap,line_rank,nd,mat_rank)
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
        for cx_ in dict.fromkeys(canon_cust(' '.join(str(y_.get('cust') or '').split())) for y_ in pp_rows or []):   # customers with a plan but no Settings daily target (Pepper, user 22-Nov-2026 "pepper and lucky same model, same bom"): usage line only, no target
            if not cx_ or any(_norm(cx_)==_norm(t_[0]) for t_ in CUST_TARGET): continue
            u_,s_=bom_usage_src(b['mat'],bool(b.get('count_stock')),cust_row(cx_))
            if u_ and s_=='your BOM': src_.append(f"{cx_}: {u_:g} {b.get('uom') or ''} ({s_})".replace('  ',' '))   # your PPC BOM only (MIS-BOM-only project customers stay out)
        x['usage_src']='\n'.join(src_) or None
        x['tgt_need_c']=nw; x['tgt_need_wo_c']=nwo; x['tgt_parts_c']=parts_; x['tgt_use_c']=use_; x['tgt_lines_c']=list(dict.fromkeys(cl_))
    for x in A:
        b=x['b']; need=0.0; used_=[]; parts_={}; use_={}
        for lk,(lab,tg) in sorted(EFF_TARGET.items()):
            k=LINE_TODAY.get(lk)
            if not tg or k is None: continue
            best=0.0; kend=min(len(dates),k+tgt_days(b.get('lt_used')))    # from the line's first production day, the material's lead time ahead
            for x_ in pp_rows or []:                    # every customer / module type this line runs in the window
                if _line_key(x_['line'])!=lk or not any(x_['v'][j]>0 for j in range(k,min(kend,len(x_['v'])))): continue
                best=max(best,bom_usage(b['mat'],bool(b.get('count_stock')),x_))   # usage per module from the BOM
            if best>0: need+=best*tg; used_.append(lk); parts_[lk]=best*tg; use_[lk]=best
        x['tgt_need']=need; x['tgt_lines']=used_; x['tgt_parts']=parts_; x['tgt_use']=use_
        if CUSTOMER_TARGETS:                  # customer targets (PPC way) replace the line-target need
            x['tgt_need']=x['tgt_need_c']; x['tgt_lines']=x['tgt_lines_c']; x['tgt_parts']=x['tgt_parts_c']; x['tgt_use']=x['tgt_use_c']
        dst_,dT_,_=vp_dead(b['mat'])                                    # Settings ALLOWED VENDORS: stock / arrivals no customer may use are out
        x['tgt_days']=(max(0.0,b['stock']-dst_+(x.get('tr_tot') or 0)-sum(dT_ or []))/x['tgt_need']) if x.get('tgt_need') else None   # with OEM target; JW stock at the OEM ignored (user rule, Nov 2026)
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
    # ---- NEED PER DAY FROM THE MODULE PLAN (every sheet except the Action List, which keeps the daily targets / OEM):
    #      pl_need = average plan per production day (> 0) over the material's transit lead time (OEM plan already left out);
    #      pl_days = (own stock + in transit) / pl_need  (no JW stock at the OEM)
    BYC_={x_['b']['mat']:x_['b'] for x_ in A}
    for x in A:
        b=x['b']; P_=b.get('plan') or []
        x['pl_need']=usage_day(P_,lt_of(b))                               # Usage/Day over the transit lead time
        dst_,dT_,_=vp_dead(b['mat'])                                      # Settings ALLOWED VENDORS: stock / arrivals no customer may use are out
        st_=float(b['stock'])-dst_+(x.get('tr_tot') or 0)-sum(dT_ or [])
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
        x['cw_p']=cover_watch(x,d0,ships,nd)                                            # Action List: MODULE PLAN scenario (= ALERT 'Below lead time from')
    # ---- SCENARIOS for the Action List
    #  with OEM    = all lines (Settings targets / highest planned day for lines without one), stock = MB52 + JW stock at the OEM
    #  without OEM = USA Lines 1-4 only, stock = MB52 only (CS / PM: Stock Count 'Today Qty_BOM' in both)
    NO_OEM_LINES={'L1','L2','L3','L4'}
    for x in A:
        b=x['b']; tr_=x.get('tr_tot') or 0
        x['oem_stock']=0.0 if b.get('count_stock') else (b.get('st_jw') or 0.0)
        x['own_stock']=b['stock'] if b.get('jw_out') else b['stock']-x['oem_stock']      # JW stock is already out of b['stock'] (OEM left out)
        dst_,dT_,_=vp_dead(b['mat']); x['own_stock']-=dst_; tr_-=sum(dT_ or [])      # Settings ALLOWED VENDORS: stock / arrivals no customer may use are out (= ALERT)
        x['tr_al']=tr_
        x['tgt_need_wo']=x['tgt_need_wo_c'] if CUSTOMER_TARGETS else sum(v for k,v in (x.get('tgt_parts') or {}).items() if k in NO_OEM_LINES)
        x['tgt_days_wo']=(max(0.0,x['own_stock']+tr_)/x['tgt_need_wo']) if x['tgt_need_wo'] else None
        nd_=len(b.get('plan') or [])
        p_wo=[0.0]*nd_
        for (cu_,ln_),arr_ in (b.get('cust_line') or {}).items():
            if ln_ and _line_key(ln_) in NO_OEM_LINES:
                for k in range(min(nd_,len(arr_))): p_wo[k]+=arr_[k]
        avg21=lambda p__: (lambda v__: sum(v__)/len(v__) if v__ else None)([q for q in p__[:PLAN_AVG_DAYS] if q>0])
        x['plan_need_wo']=avg21(p_wo); x['plan_need']=avg21(b.get('plan') or [])
        x['plan_days_wo']=(max(0.0,x['own_stock']+tr_)/x['plan_need_wo']) if x['plan_need_wo'] else None
        x['plan_days']=(max(0.0,b['stock']-(b['stock']-x['own_stock']-(0.0 if b.get('jw_out') else x['oem_stock']))+tr_)/x['plan_need']) if x.get('plan_need') else None   # dead (vendor rule) stock out
        # every PPC arrival date (Final AtA) with its pending quantity; dates already passed = not yet received
        dn_=vp_dead(b['mat'])[2]                                          # Settings ALLOWED VENDORS: shipments no customer may use are left out
        shu_=[s_ for s_ in ships.get(b['mat'],[]) if not (dn_ and dn_(s_.get('sup')))]
        x['ata_txt']=_ata_text(shu_,d0)
        nc_=vp_nc(b['mat'],ships,d0)                                      # shown, not counted (user 28-Nov-2026): the no-demand supplier's arrivals as one line
        x['nc']=nc_
        if nc_ and nc_.get('txt_it'): x['ata_txt']=((x['ata_txt']+'\n') if x['ata_txt'] else '')+nc_['txt_it']
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
    al['A2']=('Highest shortage risk first. Priority (= RM Issues): High = stock zero inside the lead time, Medium = short later, Low = under the lead time, OK. Stock lines (source, factor, storage location): Stock Data sheet. Short from = first day stock (count rolled forward, or SAP) − plan demand + PPC in transit < 0; ALERT = inside the lead time. '
              'Daily target need = customer daily targets × usage per module (your BOM first, else MIS Standard BOM); without OEM (green banner) = customers with OEM = No, with OEM (blue banner) adds the OEM and JW stock at OEM. '
              'Inv. Days = (stock + PPC in transit) ÷ daily target need, red under the lead time. '
              f'Below Lead Time From / Days Late = the ALERT logic (Short from → to holds every stock-zero period). No demand = no production requirement, planned consumption or daily target in the next {int(DEMAND_HORIZON or SHORT_HORIZON)} days: no shortage status and no order action. '
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
    def prio_(x):                         # = RM Issues Priority: High = stock zero inside the lead time, Medium = short later, Low = under the lead time
        if x.get('sup_row') and x.get('no_cust'): return 'No demand'    # supplier row no customer may use (Settings ALLOWED VENDORS, e.g. 312 FLAT)
        ri_=risk_info(x,d0)
        if ri_['cls']=='No demand': return 'No demand'
        if ri_['cls']=='ALERT' or short_lt_text(x,d0)[1]: return 'High'
        return {'Short later':'Medium','Under lead time':'Low'}.get(ri_['cls'],'OK')
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
          (f'Below Lead Time From',lambda x:_cwv(x,'below'),'dd-mmm-yy',12,False),
          ('Days Late',lambda x:_cwv(x,'late'),'0',9,False),
          ('Material lead time (days)',lambda x:x['b'].get('lt_used'),'0',9,False),
          ('Stock',lambda x:round(x['own_stock']+((x.get('nc') or {}).get('st') or 0.0)),'#,##0',11,False),        # shown = physical (user 28-Nov-2026), days below use own_stock
          ('to OEM Stock',lambda x:round(x['oem_stock']) if x.get('oem_stock') else None,'#,##0;-#,##0;"-"',11,False),
          (IT_TODAY_HEAD,lambda x:round((x.get('tr_al',x.get('tr_tot')) or 0)+((x.get('nc') or {}).get('it') or 0.0)),'#,##0;-#,##0;"-"',12,False),
          ('Arrival at plant (Final AtA)',lambda x:sup_ata(x,ships,d0),None,42,True),   # ONE rule = ALERT / 4_Status / web / email (user 26-Nov-2026)
          ('Production Daily Target_without OEM',rnd('tgt_need_wo'),'#,##0.0',12,False),
          ('Inv. Days at line daily target_without OEM',rndz('tgt_days_wo'),'0.0',12,False),
          ('Inv. Days at line daily target_without OEM (stock only, no transit)',lambda x:(round(max(0.0,x['own_stock'])/x['tgt_need_wo'],1) if x.get('tgt_need_wo') else None),'0.0',12,False),
          ('Production Daily Target_with OEM',rnd('tgt_need'),'#,##0.0',12,False),
          ('Inv. Days at line daily target_with OEM',rndz('tgt_days'),'0.0',12,False),
          ('Usage/Day (module plan)',lambda x:(round(x['pl_need'],1) if x.get('pl_need') else None),'#,##0.0',12,False),
          ('Days of cover (+ all in transit) (module plan)',lambda x:(round(x['pl_days'],1) if x.get('pl_days') is not None else None),'0.0',12,False),
          (f'Below Lead Time From (module plan)',lambda x:_cwv(x,'below','cw_p'),'dd-mmm-yy',12,False),
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
        x_['status']={'ALERT':'CRITICAL','Short later':'AT RISK','Under lead time':'WATCH','No demand':'NO DEMAND'}.get(risk_info(x_,d0)['cls'],'OK')
    DS['blocks']=sorted(DS['blocks'],key=lambda b_:rrank(b_['mat']))
    _daily(part); daily_part=out_path+'.daily.xlsx'; os.replace(part,daily_part)
    RM_={x['b']['mat']:x for x in rows_}                       # material -> its rows on MRP Daily
    # Settings cells of the line daily targets (the Settings sheet is written the same way at the end)
    TGT_CELL.clear()
    try:
        make_settings_sheet(part,SETTINGS_SRC[0]); ws_=load_workbook(part)['Settings']
        for r_ in range(1,ws_.max_row+1):
            v_=ws_.cell(r_,1).value
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
    def vends_of(cu_,mat_):                                            # every allowed vendor of a customer, top row first ([] = no rule)
        L_=next((v for k,v in LINKS.items() if _norm(k)==_norm(cu_)),None) or {}
        for k_ in (_norm(cu_),_norm(L_.get('mis'))):
            if allow_.get((k_,mat_)): return allow_[(k_,mat_)]
        return []
    def vend_of(cu_,mat_): return (vends_of(cu_,mat_) or [None])[0]
    def sup_key(name_,keys_):
        return next((k_ for k_ in keys_ if k_ and _vend_hit(k_,name_)),None)      # Settings 'HU' = batch Make 'SHANGHAI HUITIAN ...' (same rule as vendor_pools)
    out_=[]
    for x in AL_ORDER:
        out_.append(x); b=x['b']
        if not b.get('sup') or not ('glass' in str(b.get('common') or '').lower() or b['mat'] in VP_MRP): continue   # glass + every material with a Settings ALLOWED VENDORS rule (user 28-Nov-2026: RTV DOW / HU)
        need_={}; need0_={}; cus_={}
        for cu_,ln_,tg_,oem_ in CUST_TARGET:
            u_=bom_usage(b['mat'],False,cust_row(cu_))
            if not u_: continue
            v_=vend_of(cu_,b['mat']) or 'ANY VENDOR'
            need_[v_]=need_.get(v_,0.0)+u_*tg_
            for vv_ in (vends_of(cu_,b['mat']) or ['ANY VENDOR']): cus_.setdefault(vv_,[]).append(cu_)   # customers allowed this vendor (every Settings row; plan names)
            if oem_!='Yes': need0_[v_]=need0_.get(v_,0.0)+u_*tg_
        supd_={str(e_.get('name') or ''):e_ for e_ in b['sup']}
        sups_=[k_ for k_ in supd_ if k_ and k_.upper()!='UNASSIGNED']
        vall_=sorted(set(sups_)|{k_ for k_ in need_ if k_!='ANY VENDOR' and not sup_key(k_,sups_)})   # one row per supplier (Settings name = its batch Make)
        nby_=lambda v_,d_: sum(q_ for k_,q_ in d_.items() if k_!='ANY VENDOR' and (k_==v_ or _vend_hit(k_,v_)))
        act_=[v_ for v_ in vall_
              if float((supd_.get(sup_key(v_,sups_) or v_) or {}).get('stock') or 0)>0 or sum((supd_.get(sup_key(v_,sups_) or v_) or {}).get('transit') or [])>0 or nby_(v_,need_)>0]
        if len(act_)<2: continue                                        # one supplier only: its row would repeat the material row
        for v_ in vall_:
            k_=sup_key(v_,sups_) or v_
            sp_=supd_.get(k_) or {}
            st_=float(sp_.get('stock') or 0.0); it_=float(sum(sp_.get('transit') or []))
            # counted material: b['sup'] already holds the count by supplier rolled forward (apply_count_by_supplier)
            nw_=nby_(v_,need_); nwo_=nby_(v_,need0_)
            px=dict(x); px['b']=dict(b,common=f"{b['common']} - {v_}",desc=f"{b.get('desc') or ''} | supplier {v_}"); px['sup_row']=v_; px['nc']=None   # a supplier row = that supplier's own stock / arrivals (already physical)
            ul_=[l_ for l_ in str(x.get('usage_src') or '').split('\n') if l_.strip() and not l_.strip().startswith('+')
                 and (lambda w_: (not w_ and v_=='ANY VENDOR') or any(_vend_hit(z_,v_) for z_ in w_))(vends_of(l_.split(':')[0].strip(),b['mat']))]   # a customer with 2 vendor rows is on both (RTV DOW then HU)
            px['usage_src']='\n'.join(ul_) or 'no customer may use this vendor (Settings ALLOWED VENDORS)'; px['no_cust']=not ul_   # customer goes with supplier (user 22-Nov-2026)
            vpx_=VP_MRP.get(b['mat']); pv_=None
            if vpx_:                                                     # vendor rule: the supplier row is not the pooled (locked) projection (checker finding)
                Sv_=x.get('S_pool') or x.get('S') or []
                if vpx_.get('sheets'):                                   # 2+ WSA sheets: the sheet holding this supplier's stock
                    hit_=[sv[1] for sv in vpx_['sheets'].values() if abs(sv[0]-st_)<0.5]
                    if len(hit_)==1: Sv_=hit_[0]
                if vpx_.get('dead') and not (vpx_.get('sheets') and len(hit_)==1) and not vp_dead(b['mat'])[2](v_):   # pooled minus the supply no customer may use (= ALERT)
                    Sv_=[s_-(vpx_['dead_cum'][k] if k<len(vpx_.get('dead_cum') or []) else 0.0) for k,s_ in enumerate(Sv_)]
                if not vpx_.get('sheets') and not vpx_.get('dead'):      # vendor_pools own balance of this supplier (RTV DOW used first, then HU; user 23-Nov-2026)
                    pv_=next((V_ for V_ in vpx_.get('tied') or [] if _vend_hit(V_,v_)),None)
                    if pv_ is not None: Sv_=list(vpx_['pools'][pv_]); px['pool']=pv_
                wv_=floor_wins(Sv_,(st_ if ((vpx_.get('sheets') and len(hit_)==1) or pv_ is not None) else x.get('stk0')))
                px['S']=list(Sv_); px['wins']=wv_
            px['ata_txt']=_ata_text([s_ for s_ in ships.get(b['mat'],[]) if _vend_hit(v_,s_.get('sup'))],d0)   # each supplier row = its OWN arrivals (user 26-Nov-2026 'Own')
            if vpx_ and vpx_.get('dead'):                               # vendor rule with a supplier no customer may use
                px['pl_days']=None if px.get('no_cust') else ((max(0.0,st_+it_)/x['pl_need']) if x.get('pl_need') else None)   # no customer may use it: no days of cover (show it, don't count it)
            px.update(cw=None,own_stock=st_,oem_stock=0.0,tr_tot=it_,tr_al=it_,tgt_need=nw_,tgt_need_wo=nwo_,
                      tgt_days=(max(0.0,st_+it_)/nw_ if nw_ else None),tgt_days_wo=(max(0.0,st_+it_)/nwo_ if nwo_ else None),
                      custs=(', '.join(dict.fromkeys(c_ for k_,cs_ in cus_.items() if k_==v_ or _vend_hit(k_,v_) for c_ in cs_)) or 'no customer allowed this vendor'),proj=None,mods=None)
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
            if h.startswith(('Inv. Days','Days of cover')) and isinstance(c.value,(int,float)) and c.value<tgt_days(x['b'].get('lt_used')):     # under the lead time: red highlight
                c.fill=PatternFill('solid',fgColor='FFC7CE'); c.font=Font(name='Calibri',size=10,bold=True,color='9C0006')
            if h.startswith('Short from') and str(c.value or '').startswith('ALERT'): c.fill=PatternFill('solid',fgColor='FFC7CE')
    TG_,PL_=calc_layout(A,RM_); CALC_LAYOUT[:]=[TG_,PL_]
    _hs=[h_ for h_,*_ in COLS]; _j0=_hs.index('Stock'); _j1=max(i_ for i_,h_ in enumerate(_hs) if ('OEM' in h_ and ('arget' in h_ or 'lan' in h_)) or 'module plan' in h_)
    AL_COPY[:]=[(h_,L(i_+1)) for i_,h_ in enumerate(_hs) if _j0<=i_<=_j1]          # Action List F .. last daily-plan column -> MRP by Line
    hidx={h_:i_+1 for i_,(h_,*_) in enumerate(COLS)}; cl_=lambda h_: L(hidx[h_]) if h_ in hidx else 'A'
    F_,G_,H_,J_,K_,Lc,M_,N_,O_,P_,Q_=(cl_(h_) for h_ in ('Stock','to OEM Stock',IT_TODAY_HEAD,'Production Daily Target_with OEM',
        'Inv. Days at line daily target_with OEM','Production Daily Plan_with OEM','Inv. Days at line daily plan_with OEM','Production Daily Target_without OEM',
        'Inv. Days at line daily target_without OEM','Production Daily Plan_without OEM','Inv. Days at line daily plan_without OEM'))
    d0x=f'DATE({d0.year},{d0.month},{d0.day})'
    for rr,x in (zip(range(5,5+len(AL_ORDER)),AL_ORDER) if not VALUES_ONLY else []):   # every number from stock onward is a formula (change stock on Stock Data -> all follow)
        m=x['b']['mat']; rm=RM_.get(m)
        nc_=x.get('nc') or {}; nS=(f"-{round(nc_.get('st') or 0.0,6)}" if nc_.get('st') else ''); nT=(f"-{round((nc_.get('st') or 0.0)+(nc_.get('it') or 0.0),6)}" if nc_ else '')   # not counted (no customer demand yet)
        tgr=f"SUMIFS('Calc Target'!$F$2:$F$3000,'Calc Target'!$A$2:$A$3000,$A{rr})"; tgr0=f"SUMIFS('Calc Target'!$F$2:$F$3000,'Calc Target'!$A$2:$A$3000,$A{rr},'Calc Target'!$C$2:$C$3000,\"Lines 1-4\")"
        al[f'{J_}{rr}']=f'=IF({tgr}=0,"",{tgr})'; al[f'{N_}{rr}']=f'=IF({tgr0}=0,"",{tgr0})'
        al[f'{K_}{rr}']=f'=IF(N({J_}{rr})>0,MAX(0,{F_}{rr}+{H_}{rr}{nT})/{J_}{rr},"")'
        al[f'{O_}{rr}']=f'=IF(N({N_}{rr})>0,MAX(0,{F_}{rr}+{H_}{rr}{nT})/{N_}{rr},"")'
        al[f"{cl_('Inv. Days at line daily target_without OEM (stock only, no transit)')}{rr}"]=f'=IF(N({N_}{rr})>0,MAX(0,{F_}{rr}{nS})/{N_}{rr},"")'
        al[f"{cl_('Inv. Days at line daily plan_without OEM (stock only, no transit)')}{rr}"]=f'=IF(N({P_}{rr})>0,MAX(0,{F_}{rr}{nS})/{P_}{rr},"")'
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
                for need_c,days_c,prow,stk in ((Lc,M_,pr,f'{F_}{rr}+{G_}{rr}+{H_}{rr}{nT}'),(P_,Q_,pr+1,f'{F_}{rr}+{H_}{rr}{nT}')):
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
        iss=build_checks(DS_mis,BW,mis_path,None,LIMIT_DAYS=30)+cspm_checks(bom,cnt_rows,modplan,mis_names,d0)+bom_rate_checks(mis_path,bom,DS_mis,TD)+map_issues+pp_issues+unit_issues+mis_usage_check(mis_path,DS_mis)+ml_iss+list(ADVICE_ISSUES)+ubom_checks(SETTINGS_SRC[0],mis_path)+unit_bom_checks()
        iss+=[dict(area='In transit',check='PPC Final RM Code does not match the shipment - traced code used',sev='High' if (f['qty'] or 0)>0 else 'Low',
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
            ('- WSA planning file: Open PO lines without vessel, supplier / PO / ETD history (lead times, unless you typed one on Settings TRANSIT LEAD TIME).',False),
            ('- PPC file sheet MB52 (else your MB52.xlsx): the stock quantities (Unrestricted + Quality Inspection); MIS MB52 only adds plant name, final material code, dashboard key, company and categories.',False),
            ('- Your MRP.xlsx (read only): Settings, Material Group, Batch Master. Module plan (Plan + Revised Plan), BOM and Stock Count come from the PPC workbook (sheets Plan, Revised Plan, BOM, Count) when it is in the folder; MRP.xlsx Production Plan / BOM only without a PPC file.',False),
            ('- RM Issues (now in 4_Status): one issue per raw material at risk (required date, qty, supplier / PO / ETA, delay, MW impact); your Owner / Status / Date Closed / Action are kept run to run.',False),
            ('- Customer_Tracker_*.xlsx: no longer used (SCM Dashboard removed).',False),
            ('',False),('Stock',True),
            ('- Counted materials: the physical count is used; a count older than today is rolled forward: count - usage per day x days since count + received since count (PPC GRN). Usage per day = Production Plan demand, next 7 days.',False),
            ('- Not counted: SAP MB52 Total Stock (+ alternate codes / roll lengths converted); when the MIS file is older than today (its save date) this SAP stock is rolled forward the same way (− usage per day × days since the MIS date) and the console prints a WARNING. JW stock at OEM is added in the with-OEM figures.',False),
            ('- OEM (customers with OEM = Yes): their plan and the JW stock at the OEM are left out of every calculation; the JW stock stays visible as data (Stock Data, Movement Log, No Demand).',False),
            ('- In transit: each shipment is matched to its material by Description + Supplier + RM Code + Final RM Code + PO-line codes (busbar / ribbon width included); changes are listed on Check.',False),
            ('- Usage/Day = average plan per production day over the material\'s transit lead time; Inv. Days = stock ÷ Usage/Day (MRP Daily / Weekly: projected stock ÷ Usage/Day from that day) - see Definitions.',False),
            ('- Data dates (MIS, WSA, counts) are shown on the dashboards, red when older than 2 days. Stock Check sheet shows SAP vs count vs estimate per material.',False),
            ('',False),('Demand',True),
            ('- Plan demand per day = Production Plan modules x usage per module: your BOM first (an alternative listed next to its standard code for the same customer carries no demand; listed alone it carries the demand), else MIS Standard BOM.',False),
            ('- Substitutes (Settings SUBSTITUTES) between two MRP materials (EPE 420 <- EPE 480): each code is planned SEPARATELY on every sheet - own stock, plan, in transit and arrivals (a substitute code is never folded into the other code). The ALERT note / 4_Status column \'Covered by substitute\' / email / web page tell when and how much the substitute covers.',False),
            ('- Daily target need = customer daily targets (Settings CUSTOMER DAILY TARGET, else line targets) x usage; without OEM = customers with OEM = No.',False),
            ('',False),('Shortage risk - ONE rule on every sheet',True),
            ('- Day-by-day projection: stock - plan demand + PPC in transit by arrival date. Short from = first day below zero.',False),
            ('- ALERT = short inside the material lead time · Short later = short after it · Under lead time = no shortage but under the lead time at the targets · OK.',False),
            ('- Priority: EXPEDITE (short inside lead time) · ORDER NOW (last order date within the release window) · ORDER BY dd-mmm · Planned only when NOT short (below reorder point).',False),
            ('- NO DEMAND: a material is only Critical / At risk / Watch / Short when it has demand (production requirement, planned consumption, daily target or BOM demand) inside the shortage horizon (Settings, 60 days). Without it: status NO DEMAND, no planned order / release, no Expedite / Order now, not on the ALERT.',False),
            ('- Every list is sorted by this risk (highest first); a short material is never labelled safe; the script checks itself after every run (console: Self-check).',False),
            ('',False),('Lines',True),
            ('- Each material is shared out to the lines day by day (line ranking, allowed vendors, Settings substitutes) and then your approved ALTERNATIVE (your BOM) is used before a line stops; free stock of an alternative = what its own plan does not need within its lead time.',False),
            ('',False),('Sheets (visible)',True),
            ('ALERT (first tab) = ALERT + Material Status in one sheet: the ALERT rows (RM and packaging) with estimated stock today, SAP stock (MB52), physical count, Usage/Day, stock-zero days, below the lead time, in transit and overdue. Materials with NO DEMAND are never listed as a risk.',False),
            (f'ALERT (first tab): every material under the lead time, falling below the lead time in the next {HZW()} or out of stock, and why (Below lead time from, stock zero from → to, Why, Days late, Arrival at plant (Final AtA), action). The BOM Material Dashboards (weekly / daily cover) are in 4_Status only.',False),
            ('Action List - every material by shortage risk with daily-target and module-plan cover; Stock Check (SAP vs count) · MRP Daily · MRP by Line · MRP Weekly. ALERT (formulas), ALERT Calc, ALERT Plan lines, ALERT Arrival lines = the formulas behind ALERT.',False),
            ('Detail sheets (right of In Transit Status, all visible): In Transit Status, Substitute Advice, Stock Data, Check, No Demand, Settings, Movement Log, Module Plan, BOM, Stock Count, How it works, MIS Layout, Definitions.',False),
            ('',False),('Daily routine',True),
            ('1 Run the script and read the console Self-check line. 2 ALERT: stock-out in lead time, overdue in transit. 3 ALERT: Below lead time from, Why, Days late. 4 Stock Check: counts older than today and big SAP differences.',False)]
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
    unfreeze_all(tmp)                                              # no frozen panes on any sheet of 1_MRP
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
    try:                                                         # ALERT = first tab: under the lead time / stock zero inside the lead time (same rows in the Inbound file)
        write_alert_sheet(post_wb(out_path),ALERT_ROWS,d0,'this is the 1_MRP copy')
        write_alert_formulas(post_wb(out_path),ALERT_ROWS,d0)       # the same sheet with Excel formulas (check the work)
        wb__=post_wb(out_path)
        if 'Material Status' in wb__.sheetnames: wb__.move_sheet('Material Status',offset=1-wb__.sheetnames.index('Material Status'))   # right after the ALERT
        _c=Counter(r_['cls'] for r_ in ALERT_ROWS)
        print(f"ALERT sheet (first tab of 1_MRP and of the Inbound file; Risk Charts merged in): {_c.get(0,0)} out of stock inside the transit lead time, {_c.get(1,0)} out of stock later, "
              f"{_c.get(2,0)} under lead time (in transit covers), {_c.get(3,0)} under lead time later (within {HZW()}), {_c.get(4,0)} OK, {_c.get(5,0)} no demand - all {len(ALERT_ROWS)} materials listed",flush=True)
    except Exception as ex_: import traceback; traceback.print_exc(); ALERT_ROWS[:]=[]; print(f'NOTE: ALERT sheet not added ({type(ex_).__name__}: {ex_})')
    try: _heat_ms_fill(post_wb(out_path))                        # weekly dashboard E..J = Material Status columns (built after Material Status)
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: weekly dashboard Material Status columns not added ({type(ex_).__name__}: {ex_})')
    try: unify_workbook_style(out_path)                          # one style for all sheets, no charts, Definitions sheet
    except Exception as ex_: print(f'NOTE: style / definitions not applied ({type(ex_).__name__}: {ex_})')
    try:
        wb_=post_wb(out_path)
        # Action List stays in the file (user, Nov 2026: 'bring action list back')
        for w_ in wb_.worksheets: w_.sheet_view.tabSelected=False
        wb_.active=0; wb_.worksheets[0].sheet_view.tabSelected=True
    except Exception as ex_: print(f'NOTE: active sheet not set ({type(ex_).__name__}: {ex_})')
    try: make_daily_material_status(out_path,post_wb(out_path),AL_ORDER,bom,cnt_hdr,cnt_rows,d0,MIS_PATH_[0],ships)     # 4_Status_<date>.xlsx: Materials Status (BOM + other materials) + ALERT
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: Daily Material Status workbook not written ({type(ex_).__name__}: {ex_})')
    try:                                                           # RM Issues now lives in 4_Status (user, 19-Nov-2026) - out of 1_MRP
        wb_=post_wb(out_path)
        if 'RM Issues' in wb_.sheetnames: del wb_['RM Issues']
    except Exception as ex_: print(f'NOTE: RM Issues not removed from 1_MRP ({type(ex_).__name__}: {ex_})')
    try:                                                           # user 21-Nov-2026: "Daily mrp, delete all dashboard sheet" - built for 4_Status only, then out of 1_MRP
        wb_=post_wb(out_path)
        for n_ in (HEAT_NAME,):
            if n_ in wb_.sheetnames: del wb_[n_]
    except Exception as ex_: print(f'NOTE: dashboard sheets not removed from 1_MRP ({type(ex_).__name__}: {ex_})')
    try: write_etd_late_sheet(post_wb(out_path),ETD_LATE,'This is the 1_MRP copy (the same rows as the Inbound file).')   # ETD passed, not in transit (user, Nov 2026)
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: ETD late sheet not added to 1_MRP ({type(ex_).__name__}: {ex_})')
    try: web_style_mrp(out_path,{'ALERT':(ALERT_HR,ALERT_HR-1,()),'Material Status':(2,1,()),'ETD alert':(4,None,(1,2,3)),'Action List':(4,None,(1,2)),ALERT_FX:(2,1,())})   # sheets 1-5 = 5_Web Summary look (user 27-Nov-2026)
    except Exception as ex_: import traceback; traceback.print_exc(); print(f'NOTE: web look not applied to sheets 1-5 ({type(ex_).__name__}: {ex_})')
    try: post_flush(out_path)                                      # ONE save for all finishing steps
    except Exception as ex_: print(f'NOTE: finishing save failed ({type(ex_).__name__}: {ex_})')
    audit_workbook(out_path)                                       # self-check: risk order + no short material labelled safe
    run_summary('1_MRP + Inbound',extra=[f"Materials with no demand (no status / action): {', '.join(c_+' '+n_[:20] for c_,n_,w_ in NODEM_)}"] if NODEM_ else [])
    return dict(n=n,crit=sum(1 for x in A if x['status']=='CRITICAL'),risk=sum(1 for x in A if x['status']=='AT RISK'),
                watch=sum(1 for x in A if x['status']=='WATCH'),ok=sum(1 for x in A if x['status']=='OK'),nodem=sum(1 for x in A if x['status']=='NO DEMAND'),check=cnt,cs=len(cs),bom_rows=len(bom),bom_src=bom_src)

# ===================== WSA alignment: rebuild the planning sheets of the WSA file =====================
# Runs on every run of the script (python update_inbound.py); --no-align skips it.
# Output = your WSA file, updated in place (no backup folder; older '(aligned <date>)' files deleted); then the Inbound and
# 1_MRP files are built from it. Every planning sheet gets ONE standard layout - the same header row and the same columns:
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
WSA_OEM_IN_BALANCE=False     # False = like 1_MRP: OEM plan + JW stock at the OEM left out of Balance / Demand (OEM columns still show the modules)
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
CAP=[None]                   # align_capture of this run (every run with 1_MRP): Inbound 'Other materials' + PPC In Transit
